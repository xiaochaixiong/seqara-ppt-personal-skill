"""Native OOXML primitives. Never flatten a slide or rewrite a source file."""
from __future__ import annotations
from copy import deepcopy
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import hashlib
import json
import os
import posixpath as pp
import re
import uuid
from lxml import etree as E

NS = {
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}
EMU = 12700
ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / 'assets/reference.pptx'
FAMILIES = ['cover','overview','metrics','monthly-bars','ranking','segmentation',
            'donut','judgments','ranking','project-signals','photo-2','photo-2',
            'photo-3','photo-3','photo-2','photo-3','section','path-overview',
            'process-6','matrix','conversion-path','specialty','judgment',
            'employment-evidence','section','priorities','schedule','metrics-detail',
            'proposal','diagnosis','comparison','policy','experience','platform',
            'process-12','architecture','timeline','closing']
DECORATIVE_PARTS = {'ppt/media/image1.png','ppt/media/image2.png','ppt/media/image3.png',
                    'ppt/media/image19.png','ppt/media/image23.png'}

def q(name):
    prefix, local = name.split(':')
    return '{' + NS[prefix] + '}' + local

def xml(data):
    return E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True))

def serialize(root):
    return E.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)

def sha(data):
    return hashlib.sha256(data).hexdigest()

def file_sha(path):
    return sha(Path(path).read_bytes())

def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))

def write_json(path, data):
    p = Path(path); p.parent.mkdir(parents=True, exist_ok=True)
    staged=p.with_name('.'+p.name+'.'+uuid.uuid4().hex+'.tmp')
    try:
        staged.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
        os.replace(staged,p)
    finally:
        if staged.exists():staged.unlink()

def package(path):
    with ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist() if not n.endswith('/')}

def save_package(parts, path):
    with ZipFile(path, 'w', ZIP_DEFLATED) as z:
        for n in sorted(parts):
            z.writestr(n, parts[n])

def relpath(part):
    return pp.join(pp.dirname(part), '_rels', pp.basename(part) + '.rels') if part else '_rels/.rels'

def target(part, relative):
    return pp.normpath(pp.join(pp.dirname(part), relative)) if not relative.startswith('/') else relative[1:]

def relationships(parts, part):
    return xml(parts[relpath(part)]) if relpath(part) in parts else E.Element(q('rel:Relationships'), nsmap={None:NS['rel']})

def add_rel(root, rid, typ, dest):
    return E.SubElement(root, q('rel:Relationship'), Id=rid, Type=NS['r']+'/'+typ, Target=dest)

def ordered_slides(parts):
    r = xml(parts['ppt/presentation.xml'])
    rels = {x.get('Id'):target('ppt/presentation.xml',x.get('Target')) for x in relationships(parts,'ppt/presentation.xml')}
    return [rels[x.get(q('r:id'))] for x in r.findall('p:sldIdLst/p:sldId', NS)]

def text_of(node):
    return '\n'.join(''.join(p.xpath('.//a:t/text()', namespaces=NS)) for p in node.findall('.//a:p',NS))

def shapes(slide):
    """Stable ID ancestry supports duplicate names and nested groups."""
    out = {}
    def walk(parent, prefix):
        for el in parent:
            if E.QName(el).localname not in ('sp','pic','graphicFrame','grpSp','cxnSp'):
                continue
            nv = el.find('.//p:cNvPr',NS)
            if nv is None: continue
            key = '/'.join(prefix + [nv.get('id')])
            out[key] = el
            if el.tag == q('p:grpSp'): walk(el, prefix + [nv.get('id')])
    walk(slide.find('p:cSld/p:spTree',NS), [])
    return out

def bounds(node):
    t = node.find('p:spPr/a:xfrm',NS)
    if t is None: t = node.find('p:xfrm',NS)
    if t is None: t = node.find('p:grpSpPr/a:xfrm',NS)
    if t is None: return None
    off, ext = t.find('a:off',NS), t.find('a:ext',NS)
    return [int(off.get('x'))/EMU,int(off.get('y'))/EMU,int(ext.get('cx'))/EMU,int(ext.get('cy'))/EMU] if off is not None and ext is not None else None

def set_bounds(node, box):
    t = node.find('p:spPr/a:xfrm',NS)
    if t is None: t = node.find('p:xfrm',NS)
    if t is None: raise ValueError('Shape has no direct transform')
    off, ext = t.find('a:off',NS), t.find('a:ext',NS)
    off.set('x',str(round(box[0]*EMU))); off.set('y',str(round(box[1]*EMU)))
    ext.set('cx',str(round(box[2]*EMU))); ext.set('cy',str(round(box[3]*EMU)))

def style_signature(node):
    c = deepcopy(node)
    for el in c.findall('.//a:t',NS): el.text = ''
    return sha(E.tostring(c,method='c14n'))

def run_signature(run):
    pr = run.find('a:rPr',NS)
    return E.tostring(pr) if pr is not None else b''

def replace_text(node, value):
    body = node.find('p:txBody',NS)
    if body is None: raise ValueError('Target has no text body')
    old = body.findall('a:p',NS)
    runs = node.findall('.//a:r',NS)
    if isinstance(value, str):
        if len({run_signature(r) for r in runs}) > 1:
            raise ValueError('Mixed formatting requires paragraphs/runs with style_run indices')
        value = [{'runs':[{'text':line,'style_run':0}]} for line in value.split('\n')]
    for p in old: body.remove(p)
    for i, item in enumerate(value):
        base = old[min(item.get('style_paragraph',i),len(old)-1)] if old else E.Element(q('a:p'))
        p = deepcopy(base)
        for child in list(p):
            if child.tag not in (q('a:pPr'),q('a:endParaRPr')): p.remove(child)
        end = p.find('a:endParaRPr',NS)
        for spec in item['runs']:
            idx = spec.get('style_run',0)
            if runs and not 0 <= idx < len(runs): raise ValueError('Invalid style_run index')
            r = deepcopy(runs[idx]) if runs else E.Element(q('a:r'))
            for x in list(r):
                if x.tag != q('a:rPr'): r.remove(x)
            t = E.SubElement(r,q('a:t')); t.text = str(spec['text'])
            t.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
            p.insert(list(p).index(end),r) if end is not None else p.append(r)
        body.append(p)

def inventory(path=TEMPLATE):
    parts=package(path); slides=[]
    for i,part in enumerate(ordered_slides(parts),1):
        root=xml(parts[part]); rr={r.get('Id'):r for r in relationships(parts,part)}; slots=[]
        for key,s in shapes(root).items():
            nv=s.find('.//p:cNvPr',NS); b=bounds(s)
            body=s.find('p:txBody',NS)
            rec={'key':key,'name':nv.get('name'),'type':E.QName(s).localname,'bounds_pt':b,'text':text_of(s),'xml_sha256':sha(serialize(s))}
            if body is not None:
                rec['text_body_xml']=E.tostring(body,encoding='unicode')
                rec['runs']=[{'index':j,'text':''.join(r.xpath('a:t/text()',namespaces=NS)),
                              'properties_xml':E.tostring(r.find('a:rPr',NS),encoding='unicode') if r.find('a:rPr',NS) is not None else None}
                             for j,r in enumerate(s.findall('.//a:r',NS))]
                rec['capacity']={'reference_characters':len(rec['text']),'reference_paragraphs':len(body.findall('a:p',NS)),
                                 'policy':'Reference occupancy is a hint; PowerPoint measurement is authoritative. Split before shrinking.'}
            blip=s.find('.//a:blip',NS)
            if blip is not None:
                r=rr.get(blip.get(q('r:embed')))
                if r is not None:
                    dest=target(part,r.get('Target'));rec.update(media=dest,media_sha256=sha(parts[dest]),role='decoration' if dest in DECORATIVE_PARTS else 'evidence')
                crop=s.find('.//a:srcRect',NS);rec['crop']=dict(crop.attrib) if crop is not None else {}
            slots.append(rec)
        slides.append({'number':i,'part':part,'family':FAMILIES[i-1] if i<=38 else 'unknown','slots':slots})
    pres=xml(parts['ppt/presentation.xml']);sz=pres.find('p:sldSz',NS)
    return {'schema_version':1,'template_sha256':file_sha(path),'canvas_pt':[int(sz.get('cx'))/EMU,int(sz.get('cy'))/EMU],
            'inheritance':'Native theme/master/layout parts retained; text_body_xml stores explicit overrides without inventing defaults.',
            'slides':slides}

def content_type(parts, part, mime):
    root=xml(parts['[Content_Types].xml'])
    for el in list(root):
        if el.get('PartName')=='/'+part: root.remove(el)
    E.SubElement(root,q('ct:Override'),PartName='/'+part,ContentType=mime)
    parts['[Content_Types].xml']=serialize(root)

def repair_chart(parts, source_part, new_part, data=None):
    """Rebuild the workbook and caches, retaining chart drawing/style XML."""
    from openpyxl import Workbook
    chart=xml(parts[source_part]);series=chart.findall('.//c:ser',NS)
    if len(series)!=1: raise ValueError('Personal chart currently requires one doughnut series')
    ser=series[0]
    old_cats=ser.xpath('c:cat//c:pt/c:v/text()',namespaces=NS)
    old_vals=ser.xpath('c:val//c:pt/c:v/text()',namespaces=NS)
    cats=data['categories'] if data is not None else old_cats
    vals=data['values'] if data is not None else [float(x) for x in old_vals]
    if not cats or len(cats)!=len(vals) or any(isinstance(x,bool) or not isinstance(x,(int,float)) or not __import__('math').isfinite(x) or x<0 for x in vals) or sum(vals)<=0:
        raise ValueError('Invalid doughnut categories/values')
    if len(cats)>len(old_cats): raise ValueError('More chart categories need a reviewed palette variant; split/reclassify instead')
    name=data.get('name','数量') if data else '数量'
    for tag in ['tx','cat','val']:
        old=ser.find('c:'+tag,NS)
        idx=list(ser).index(old) if old is not None else len(ser)
        if old is not None: ser.remove(old)
        holder=E.Element(q('c:'+tag));ser.insert(idx,holder)
        numeric=tag=='val';ref=E.SubElement(holder,q('c:numRef' if numeric else 'c:strRef'))
        E.SubElement(ref,q('c:f')).text="'Data'!$B$1" if tag=='tx' else f"'Data'!${'B' if numeric else 'A'}$2:${'B' if numeric else 'A'}${len(cats)+1}"
        cache=E.SubElement(ref,q('c:numCache' if numeric else 'c:strCache'))
        if numeric: E.SubElement(cache,q('c:formatCode')).text='General'
        values=vals if numeric else ([name] if tag=='tx' else cats)
        E.SubElement(cache,q('c:ptCount'),val=str(len(values)))
        for i,v in enumerate(values): E.SubElement(E.SubElement(cache,q('c:pt'),idx=str(i)),q('c:v')).text=str(v)
    for dp in list(ser.findall('c:dPt',NS)):
        if int(dp.find('c:idx',NS).get('val'))>=len(cats): ser.remove(dp)
    wb=Workbook();ws=wb.active;ws.title='Data';ws.append(['类别',name])
    for c,v in zip(cats,vals): ws.append([c,v])
    stream=BytesIO();wb.save(stream)
    workbook='ppt/embeddings/'+pp.basename(new_part).replace('.xml','.xlsx');parts[workbook]=stream.getvalue()
    rels=relationships(parts,source_part)
    for rel in list(rels):
        if rel.get('Type').endswith(('/oleObject','/package')): rels.remove(rel)
    rid='rIdPersonalData';add_rel(rels,rid,'package',pp.relpath(workbook,pp.dirname(new_part)))
    external=chart.find('c:externalData',NS)
    if external is None: external=E.SubElement(chart,q('c:externalData'))
    external.set(q('r:id'),rid)
    auto=external.find('c:autoUpdate',NS)
    if auto is None: auto=E.SubElement(external,q('c:autoUpdate'))
    auto.set('val','0')
    parts[new_part]=serialize(chart);parts[relpath(new_part)]=serialize(rels)
    content_type(parts,new_part,'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
    content_type(parts,workbook,'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def notes_part(parts, slide_part, index, text):
    name=f'ppt/notesSlides/personal-notes-{index}.xml'
    root=E.Element(q('p:notes'),nsmap={k:NS[k] for k in ('a','r','p')})
    tree=E.SubElement(E.SubElement(root,q('p:cSld')),q('p:spTree'))
    nv=E.SubElement(tree,q('p:nvGrpSpPr'));E.SubElement(nv,q('p:cNvPr'),id='1',name='');E.SubElement(nv,q('p:cNvGrpSpPr'));E.SubElement(nv,q('p:nvPr'))
    E.SubElement(tree,q('p:grpSpPr'))
    sp=E.SubElement(tree,q('p:sp'));nv=E.SubElement(sp,q('p:nvSpPr'));E.SubElement(nv,q('p:cNvPr'),id='2',name='讲述备注');E.SubElement(nv,q('p:cNvSpPr'))
    E.SubElement(E.SubElement(nv,q('p:nvPr')),q('p:ph'),type='body',idx='1')
    E.SubElement(sp,q('p:spPr'));body=E.SubElement(sp,q('p:txBody'));E.SubElement(body,q('a:bodyPr'));E.SubElement(body,q('a:lstStyle'))
    for line in text.split('\n'):
        p=E.SubElement(body,q('a:p'));r=E.SubElement(p,q('a:r'));E.SubElement(r,q('a:t')).text=line
    E.SubElement(E.SubElement(root,q('p:clrMapOvr')),q('a:masterClrMapping'))
    rels=relationships({},name);add_rel(rels,'rId1','slide',pp.relpath(slide_part,pp.dirname(name)))
    masters=[n for n in parts if re.fullmatch(r'ppt/notesMasters/[^/]+\.xml',n)]
    if masters:add_rel(rels,'rId2','notesMaster',pp.relpath(masters[0],pp.dirname(name)))
    parts[name]=serialize(root);parts[relpath(name)]=serialize(rels)
    content_type(parts,name,'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml')
    return name

def prune(parts):
    """Remove unreachable historical pages, notes, media, thumbnails and chart caches."""
    seen=set();queue=['']
    while queue:
        part=queue.pop()
        if part in seen: continue
        if part and part not in parts: raise ValueError('Missing OPC part: '+part)
        seen.add(part)
        rp=relpath(part)
        if rp in parts:
            seen.add(rp)
            for rel in relationships(parts,part):
                if rel.get('TargetMode')!='External': queue.append(target(part,rel.get('Target')))
    out={n:v for n,v in parts.items() if n in seen or n=='[Content_Types].xml'}
    ct=xml(out['[Content_Types].xml'])
    for el in list(ct):
        if el.get('PartName') and el.get('PartName')[1:] not in out:ct.remove(el)
    out['[Content_Types].xml']=serialize(ct)
    return out
