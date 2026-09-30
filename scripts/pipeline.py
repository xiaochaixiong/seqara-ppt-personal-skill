"""Source contracts, native slide assembly, provenance and immutable revisions."""
from __future__ import annotations
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
import json
import math
import os
import re
import uuid
from native_pptx import *
from data_sources import load_data_sources, evaluate_metrics, metric_report
from native_data_charts import validate_data_chart, add_data_chart
from media_assets import validate_asset_bindings
from transitions import validate_transitions

KINDS={'source','derived','inference','proposal','gap','label'}

def normalized(s):
    return re.sub(r'\s+','',str(s)).replace('，',',').replace('％','%')

def extract_source(path, source_id, role='primary'):
    p=Path(path).resolve(); suffix=p.suffix.lower();fragments=[]
    if suffix=='.docx':
        root=xml(package(p)['word/document.xml'])
        # Include table paragraphs in document order; never execute text as instructions.
        for i,para in enumerate(root.findall('.//w:body//w:p',NS),1):
            txt=''.join(para.xpath('.//w:t/text()',namespaces=NS))
            if txt.strip():fragments.append({'id':f'p{i:03d}','text':txt})
    elif suffix=='.pptx':
        parts=package(p)
        for i,part in enumerate(ordered_slides(parts),1):
            fragments.append({'id':f's{i:03d}','text':text_of(xml(parts[part]))})
    elif suffix in ('.txt','.md'):
        for i,line in enumerate(p.read_text(encoding='utf-8-sig').splitlines(),1):
            if line.strip():fragments.append({'id':f'p{i:03d}','text':line})
    else: raise ValueError('Supported sources: DOCX, PPTX, UTF-8 TXT/MD. Convert spreadsheet facts into a cited source ledger first.')
    return {'schema_version':1,'id':source_id,'role':role,'path':str(p),'sha256':file_sha(p),'fragments':fragments}

def load_sources(plan, base):
    sources={};refs={};inputs=[]
    for f in plan.get('sources',[]):
        path=(base/f).resolve();s=read_json(path);sid=s['id']
        if sid in sources:raise ValueError('Duplicate source id: '+sid)
        original=(path.parent/s['path']).resolve()
        if file_sha(original)!=s['sha256']:raise ValueError('Source changed: '+str(original))
        # Re-extract so tampering with a source ledger cannot invent supporting evidence.
        actual=extract_source(original,sid,s['role'])
        if actual['fragments']!=s['fragments']:raise ValueError('Source ledger differs from source: '+sid)
        sources[sid]=s
        for fragment in s['fragments']:refs[sid+':'+fragment['id']]=fragment['text']
        inputs.extend([{'path':str(original),'sha256':s['sha256']},{'path':str(path),'sha256':file_sha(path)}])
    data,cell_refs,cells,data_inputs=load_data_sources(plan.get('data_sources',[]),base,sources)
    sources.update(data);refs.update(cell_refs);inputs.extend(data_inputs)
    return sources,refs,inputs

def evaluate_plan_facts(plan, refs, sources):
    """Merge optional table metrics into the legacy numeric fact namespace."""
    facts=evaluate_facts(plan.get('facts',[]),refs)
    cells={f'{sid}:{cell["id"]}':cell for sid,s in sources.items()
           if s.get('kind')=='data_source' for cell in s['cells']}
    metrics=evaluate_metrics(plan.get('metrics',[]),cells)
    duplicate=set(facts)&set(metrics)
    if duplicate:raise ValueError('Fact and metric IDs overlap: '+', '.join(sorted(duplicate)))
    facts.update(metrics)
    return facts

def data_summary(plan, base):
    sources,refs,inputs=load_sources(plan,base)
    metrics=evaluate_metrics(plan.get('metrics',[]),
        {f'{sid}:{cell["id"]}':cell for sid,s in sources.items()
         if s.get('kind')=='data_source' for cell in s['cells']})
    return {'data_sources':sum(s.get('kind')=='data_source' for s in sources.values()),
            'cells':sum(len(s['cells']) for s in sources.values() if s.get('kind')=='data_source'),
            'metrics':metric_report(metrics), 'inputs':inputs}

def evaluate_facts(facts, refs):
    by_id={f['id']:f for f in facts};done={};active=set()
    if len(by_id)!=len(facts):raise ValueError('Duplicate fact id')
    def evaluate(fid):
        if fid in done:return done[fid]
        if fid in active:raise ValueError('Cyclic derived facts')
        if fid not in by_id:raise ValueError('Unknown fact '+fid)
        active.add(fid);f=by_id[fid];value=f['value']
        if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value):raise ValueError('Fact must be finite numeric')
        if not f.get('unit') or not f.get('period'):raise ValueError('Fact requires unit and period')
        if f['kind']=='source':
            if f['source'] not in refs:raise ValueError('Unknown fact source')
            quote=f.get('quote','')
            if not quote or normalized(quote) not in normalized(refs[f['source']]):raise ValueError('Fact quote not found')
            nums=re.findall(r'(?<![\d.])-?\d+(?:\.\d+)?',quote.replace(',',''))
            if not any(math.isclose(float(x),value,abs_tol=1e-9) for x in nums):raise ValueError('Fact value not in quote')
        elif f['kind']=='derived':
            ids=f['operands'];values=[evaluate(x) for x in ids]
            if not values:raise ValueError('Empty derived operands')
            units={by_id[x]['unit'] for x in ids};periods={by_id[x]['period'] for x in ids}
            if len(units)!=1 or periods!={f['period']}:raise ValueError('Incompatible fact units/periods')
            op=f['operation']
            if op in ('difference','ratio_percent') and len(values)!=2:raise ValueError('Operation needs two operands')
            if op=='sum':result=sum(values)
            elif op=='mean':result=sum(values)/len(values)
            elif op=='difference':result=values[0]-values[1]
            elif op=='ratio_percent':
                if values[1]==0:raise ValueError('Zero denominator')
                result=values[0]/values[1]*100
            else:raise ValueError('Unsupported calculation')
            expected_unit='%' if op=='ratio_percent' else next(iter(units))
            if f['unit']!=expected_unit:raise ValueError('Invalid derived unit')
            if not math.isclose(round(result,f.get('decimals',1)),value,abs_tol=1e-8):raise ValueError('Incorrect derived value: '+fid)
        else:raise ValueError('Facts must be source or derived')
        done[fid]=value;active.remove(fid);return value
    for fid in by_id:evaluate(fid)
    return by_id

def output_text(binding, original):
    if binding['action']=='keep':return original
    if binding['action']=='remove':return ''
    if 'text' in binding:return binding['text']
    return '\n'.join(''.join(str(r['text']) for r in p['runs']) for p in binding['paragraphs'])

def validate_binding(binding, original, refs, facts, sources, template_sha):
    kind=binding.get('kind')
    if kind not in KINDS:raise ValueError('Content binding requires a valid kind')
    action=binding.get('action')
    if action not in ('keep','replace','remove'):raise ValueError('Invalid binding action')
    if action=='remove':
        if not binding.get('reason'):raise ValueError('Removal requires reason')
        return
    rs=binding.get('sources',[])
    for r in rs:
        if r not in refs:raise ValueError('Unknown source fragment: '+r)
    if kind in ('source','derived','inference') and not rs:raise ValueError('Factual/analytical content needs sources')
    if action=='keep':
        # Exact-current-report reuse is opt-in per slot, never a future-period default.
        valid=any(sources[r.split(':')[0]]['role']=='evidence' and sources[r.split(':')[0]]['sha256']==template_sha
                  and normalized(original) in normalized(refs[r]) for r in rs)
        if not valid:raise ValueError('Keeping historical content needs matching template evidence source')
    txt=output_text(binding,original)
    def has_value(number,values):
        return any(math.isclose(float(number),float(v),rel_tol=0,abs_tol=1e-9) for v in values)
    if kind=='source':
        available=' '.join(refs[r] for r in rs)
        number_pattern=r'(?<![\d.])[-−]?\d+(?:\.\d+)?'
        available_numbers=set(re.findall(number_pattern,available.replace(',','')))
        for n in re.findall(number_pattern,txt.replace(',','')):
            bound_values=[facts[f]['value'] for f in binding.get('facts',[]) if f in facts and facts[f]['value'] is not None]
            number=n.replace('−','-')
            if not has_value(number,[x.replace('−','-') for x in available_numbers]) and not has_value(number,bound_values):
                raise ValueError('Unbound visible number '+n)
    if kind=='derived' and not binding.get('facts'):raise ValueError('Derived content needs evaluated facts')
    for fid in binding.get('facts',[]):
        if fid not in facts:raise ValueError('Unknown bound fact '+fid)
        f=facts[fid]
        if f['value'] is None:raise ValueError('Missing metric cannot bind a numeric slot: '+fid)
        visible=re.findall(r'(?<![\d.])-?\d+(?:\.\d+)?',txt.replace(',',''))
        if not has_value(f['value'],visible):
            raise ValueError('Fact value missing in bound slot: '+fid)
        for term in f.get('visible_qualifiers',[]):
            if term not in txt:raise ValueError('Required fact qualifier missing: '+term)

def validate_bar_updates(page, ss):
    """Require deterministic geometry updates when known native chart values change."""
    number=page['template_slide'];pairs=[]
    if number==4:pairs=[(f'v{i}',[f'b{i}']) for i in range(1,8)]
    elif number==5 and not page.get('ranking'):pairs=[(f'sc{i}-v',[f'sc{i}-b']) for i in range(17)]
    elif number==7:pairs=[(f'ot-{name}-v',[f'ot-{name}-b']) for name in ['实训营','赛事','品牌活动','人培研讨','课程']]
    elif number==9:pairs=[(f'r{i}-v',[f'r{i}-b1',f'r{i}-b2']) for i in range(9)]
    elif number==10:pairs=[(f't-{name}-v',[f't-{name}-b']) for name in ['走出去','师资培训','请进来','赛事','人培研讨','实训营','品牌活动','课程']]
    names={s.find('.//p:cNvPr',NS).get('name'):key for key,s in ss.items() if s.tag!=q('p:grpSp')}
    bar_specs=page.get('bars',[])
    if len({b['shape'] for b in bar_specs})!=len(bar_specs):raise ValueError('Duplicate native bar update')
    updated={(b['value_slot'],b['shape']) for b in bar_specs}
    removed=set(page.get('remove_shapes',[]))|set(page.get('data_chart',{}).get('replace_shapes',[]))
    def numbers(txt):return re.findall(r'[-+]?\d+(?:\.\d+)?',txt.replace('−','-').replace(',',''))
    for label_name,bar_names in pairs:
        label=names[label_name];binding=page['slots'][label]
        original=text_of(ss[label]);current=output_text(binding,original)
        if numbers(original)==numbers(current):continue
        for name in bar_names:
            key=names[name]
            if binding['action']=='remove':
                if key not in removed:raise ValueError(f'Removed numeric label needs bar removal: {label}/{key}')
            elif (label,key) not in updated:raise ValueError(f'Changed numeric label needs synchronized bar update: {label}/{key}')

def validate_storyboard(plan, require=False):
    """Check a declared beat-to-page map and its speaking-time accounting.

    This is intentionally mechanical: it cannot decide whether the argument,
    pacing or imagery is persuasive.  A beat may span pages, but its assigned
    seconds may only be counted once in the declared talk duration.
    """
    if 'storyboard' not in plan:
        if require:raise ValueError('Plan has no storyboard metadata')
        return None
    story=plan['storyboard']
    if not isinstance(story,dict):raise ValueError('Storyboard must be an object')
    def seconds(value,label):
        if isinstance(value,bool) or not isinstance(value,(int,float)) or not math.isfinite(value) or value<=0:
            raise ValueError(f'Storyboard {label} must be positive finite seconds')
        return float(value)
    duration=seconds(story.get('duration_seconds'),'duration_seconds')
    pages=plan.get('pages')
    if not isinstance(pages,list) or not pages:raise ValueError('Storyboard needs deck pages')
    page_ids=[]
    for page in pages:
        pid=page.get('id') if isinstance(page,dict) else None
        if not isinstance(pid,str) or not pid.strip():raise ValueError('Storyboard page needs an id')
        page_ids.append(pid)
    if len(page_ids)!=len(set(page_ids)):raise ValueError('Storyboard deck has duplicate page ids')
    beats=story.get('beats')
    if not isinstance(beats,list) or not beats:raise ValueError('Storyboard needs beats')
    expected={}
    for beat in beats:
        bid=beat.get('id') if isinstance(beat,dict) else None
        if not isinstance(bid,str) or not bid.strip():raise ValueError('Storyboard beat needs an id')
        if bid in expected:raise ValueError('Storyboard duplicate beat id: '+bid)
        expected[bid]=seconds(beat.get('expected_seconds'),'expected_seconds for '+bid)
    assignments=story.get('page_assignments')
    if not isinstance(assignments,list) or not assignments:raise ValueError('Storyboard needs page_assignments')
    beat_totals={bid:0.0 for bid in expected};page_totals={pid:0.0 for pid in page_ids};seen=set()
    for item in assignments:
        if not isinstance(item,dict):raise ValueError('Storyboard assignment must be an object')
        pid=item.get('page_id');bid=item.get('beat_id')
        if pid not in page_totals:raise ValueError('Storyboard assignment has unknown page: '+str(pid))
        if bid not in beat_totals:raise ValueError('Storyboard assignment has unknown beat: '+str(bid))
        pair=(pid,bid)
        if pair in seen:raise ValueError(f'Storyboard duplicate beat assignment: {bid} on {pid}')
        seen.add(pair)
        value=seconds(item.get('seconds'),f'assignment {bid} on {pid}')
        beat_totals[bid]+=value;page_totals[pid]+=value
    uncovered=[bid for bid,total in beat_totals.items() if total==0]
    if uncovered:raise ValueError('Storyboard unassigned beats: '+', '.join(uncovered))
    unassigned=[pid for pid,total in page_totals.items() if total==0]
    if unassigned:raise ValueError('Storyboard unassigned deck pages: '+', '.join(unassigned))
    for bid,target in expected.items():
        if not math.isclose(beat_totals[bid],target,rel_tol=0,abs_tol=1e-6):
            raise ValueError(f'Storyboard beat {bid} assigned {beat_totals[bid]:g}s, expected {target:g}s')
    beat_sum=sum(expected.values());assigned_sum=sum(page_totals.values())
    if not math.isclose(beat_sum,duration,rel_tol=0,abs_tol=1e-6):
        raise ValueError(f'Storyboard duration {duration:g}s differs from beat sum {beat_sum:g}s')
    if not math.isclose(assigned_sum,duration,rel_tol=0,abs_tol=1e-6):
        raise ValueError(f'Storyboard duration {duration:g}s differs from page assignment sum {assigned_sum:g}s')
    return {'duration_seconds':duration,'beat_count':len(expected),'page_count':len(page_ids),
            'assignment_count':len(assignments),
            'beats':[{'id':bid,'expected_seconds':target,'assigned_seconds':beat_totals[bid]}
                     for bid,target in expected.items()],
            'pages':[{'id':pid,'assigned_seconds':page_totals[pid]} for pid in page_ids]}

def validate_plan(plan, base):
    if plan.get('schema_version')!=1:raise ValueError('Unsupported plan version')
    if plan.get('template_sha256')!=file_sha(TEMPLATE):raise ValueError('Template fingerprint mismatch')
    sources,refs,inputs=load_sources(plan,base);facts=evaluate_plan_facts(plan,refs,sources)
    media=validate_asset_bindings(plan,base);inputs.extend(media['inputs'])
    deck=package(TEMPLATE);slideparts=ordered_slides(deck);ids=[];coverage=set();segments={};gaps=list(plan.get('gaps',[]))+media.get('gaps',[])
    if not 1<=len(plan['pages'])<=100:raise ValueError('Plan needs 1–100 pages')
    for page in plan['pages']:
        pid=page['id'];ids.append(pid)
        number=page['template_slide']
        if not isinstance(number,int) or not 1<=number<=len(slideparts):raise ValueError('Invalid template slide')
        if not page.get('purpose'):raise ValueError('Page needs purpose')
        part=slideparts[number-1];root=xml(deck[part]);ss=shapes(root)
        for key,s in ss.items():
            name=s.find('.//p:cNvPr',NS).get('name');txt=text_of(s)
            if s.tag==q('p:grpSp'):continue
            if txt and name!='page-num':
                if key not in page.get('slots',{}):raise ValueError(f'{pid}: unbound historical text {key} ({name})')
                binding=page['slots'][key];validate_binding(binding,txt,refs,facts,sources,plan['template_sha256'])
                if binding['kind']=='gap':gaps.append({'page':pid,'blocking':True,'message':f'Unresolved text slot {key}'})
        for key,b in page.get('slots',{}).items():
            if key not in ss or ss[key].find('p:txBody',NS) is None:raise ValueError(f'Unknown text slot {pid}/{key}')
            if b.get('action')=='replace':
                copy=deepcopy(ss[key]);replace_text(copy,b.get('text',b.get('paragraphs')))
        # In reference pages 1–37 the native key 6 is the visible main title.
        # A nonempty purpose in the JSON cannot compensate for a removed title.
        # Page 38 is a closing composition whose key 6 is only a column label.
        if number<=37:
            title=page['slots'].get('6')
            visible='' if title is None else re.sub(r'[\s\u200b-\u200d\ufeff]','',output_text(title,text_of(ss['6'])))
            if not visible:
                raise ValueError(f'{pid}: main title slot 6 must contain visible text')
        for key in page.get('remove_shapes',[]):
            if key not in ss or text_of(ss[key]) or ss[key].tag in (q('p:pic'),q('p:grpSp'),q('p:graphicFrame')):
                raise ValueError('remove_shapes only accepts empty non-media leaf shapes')
        for key,spec in page.get('geometry',{}).items():
            box=spec['bounds_pt']
            if key not in ss or len(box)!=4 or any(isinstance(v,bool) or not isinstance(v,(int,float)) or not math.isfinite(v) for v in box) or box[0]<0 or box[1]<0 or box[2]<=0 or box[3]<=0 or box[0]+box[2]>960 or box[1]+box[3]>540 or not spec.get('reason'):
                raise ValueError('Geometry requires a valid in-canvas box and adaptation reason')
        for bar in page.get('bars',[]):
            key=bar['shape'];fid=bar['fact'];label=bar['value_slot'];maximum=bar['maximum']
            if key not in ss or text_of(ss[key]) or fid not in facts or label not in page.get('slots',{}) or fid not in page['slots'][label].get('facts',[]):
                raise ValueError('Bar requires native shape, evaluated fact and bound numeric label')
            if bar['axis'] not in ('width','height') or not isinstance(maximum,(int,float)) or not math.isfinite(maximum) or maximum<=0 or not 0<=facts[fid]['value']<=maximum:
                raise ValueError('Invalid bar axis or scale')
            extent=bar['extent_pt'];box=bounds(ss[key])
            if not isinstance(extent,(int,float)) or not math.isfinite(extent) or extent<=0 or (bar['axis']=='width' and box[0]+extent>960) or (bar['axis']=='height' and extent>box[1]+box[3]):raise ValueError('Bar extent outside canvas')
        validate_bar_updates(page,ss)
        rr={r.get('Id'):r for r in relationships(deck,part)}
        for key,s in ss.items():
            blip=s.find('.//a:blip',NS)
            if s.tag!=q('p:pic') or blip is None:continue
            media=target(part,rr[blip.get(q('r:embed'))].get('Target'))
            if media in DECORATIVE_PARTS and key not in page.get('pictures',{}):continue
            pic=page.get('pictures',{}).get(key)
            if not pic:raise ValueError(f'{pid}: unbound historical photo {key}')
            if pic['action']=='keep':
                if not any(sources.get(r.split(':')[0],{}).get('sha256')==plan['template_sha256'] and sources[r.split(':')[0]]['role']=='evidence' and r in refs for r in pic.get('sources',[])):
                    raise ValueError('Historical evidence image requires matching evidence source')
            elif pic['action']=='replace':
                image=(base/pic['path']).resolve()
                if file_sha(image)!=pic['sha256']:raise ValueError('Image fingerprint mismatch')
                if pic.get('role')!='evidence' or not pic.get('sources') or any(r not in refs for r in pic['sources']):raise ValueError('Replacement evidence photo needs provenance')
                if pic.get('generated'):raise ValueError('Generated image cannot be evidence')
                crop=pic.get('crop',{})
                if set(crop)-{'l','t','r','b'} or any(not isinstance(v,int) or v<0 or v>=100000 for v in crop.values()) or crop.get('l',0)+crop.get('r',0)>=100000 or crop.get('t',0)+crop.get('b',0)>=100000:
                    raise ValueError('Invalid picture crop')
                from PIL import Image
                with Image.open(image) as im:im.verify()
                inputs.append({'path':str(image),'sha256':pic['sha256']})
            elif pic['action']=='remove':
                if not pic.get('reason'):raise ValueError('Image removal needs reason')
                gaps.append({'page':pid,'blocking':True,'message':'Missing evidence image '+key+': '+pic['reason']})
            else:raise ValueError('Invalid image action')
        charts=root.findall('.//c:chart',NS)
        if page.get('data_chart'):
            if charts:raise ValueError('data_chart needs a reference page without an existing native chart')
            validate_data_chart(page['data_chart'],page,ss,facts)
            visible='\n'.join(output_text(b,text_of(ss[key])) for key,b in page.get('slots',{}).items())
            if page['data_chart']['source_label'] not in visible:
                raise ValueError('data_chart source_label must be visible in a bound slot')
            if page['data_chart']['interpretation'] not in visible+'\n'+page.get('notes_extra',''):
                raise ValueError('data_chart interpretation must appear on screen or in notes')
        if charts:
            binding=page.get('chart')
            if not binding:raise ValueError('Chart requires explicit data/evidence binding')
            if binding['action']=='keep':
                if not any(r in refs and sources[r.split(':')[0]]['sha256']==plan['template_sha256'] and sources[r.split(':')[0]]['role']=='evidence' for r in binding.get('sources',[])):
                    raise ValueError('Historical chart requires matching evidence source')
            elif binding['action']=='replace':
                if not 1<=len(binding['categories'])==len(binding['fact_ids'])==len(binding['values'])<=4:raise ValueError('Chart needs 1–4 categories, values and bound facts')
                for val,fid in zip(binding['values'],binding['fact_ids']):
                    if fid not in facts or facts[fid]['value']!=val:raise ValueError('Chart value differs from bound fact')
                if len({facts[f]['unit'] for f in binding['fact_ids']})!=1 or len({facts[f]['period'] for f in binding['fact_ids']})!=1:raise ValueError('Chart facts need common units/period')
            else:raise ValueError('Unsupported chart action')
        for r in page.get('notes_refs',[]):
            if r not in refs:raise ValueError('Unknown notes fragment')
            coverage.add(r)
        for seg in page.get('notes_segments',[]):
            ref=seg['source'];quote=seg['quote']
            if ref not in refs or not quote or quote not in refs[ref]:raise ValueError('Notes segment is not an exact source excerpt')
            source=normalized(refs[ref]);quote=normalized(quote);start=source.index(quote)
            segments.setdefault(ref,set()).update(range(start,start+len(quote)))
        for item in page.get('coverage',[]):
            if item['source'] not in refs or item.get('destination') not in ('slide','notes','omit'):raise ValueError('Invalid coverage record')
            if not item.get('reason'):raise ValueError('Coverage needs an explanation')
            if item['destination']=='notes' and item['source'] not in page.get('notes_refs',[]):raise ValueError('Claimed notes coverage not present')
            if item['destination']=='slide' and not any(item['source'] in b.get('sources',[]) for b in page.get('slots',{}).values()):raise ValueError('Claimed slide coverage not bound')
            coverage.add(item['source'])
    if len(ids)!=len(set(ids)):raise ValueError('Duplicate page ids')
    for ref,covered in segments.items():
        if len(covered)==len(normalized(refs[ref])):coverage.add(ref)
    for s in sources.values():
        if s['role']=='primary':
            missing=[s['id']+':'+f['id'] for f in s['fragments'] if s['id']+':'+f['id'] not in coverage]
            if missing:raise ValueError('Uncovered manuscript paragraphs: '+', '.join(missing))
    for p in plan['pages']:
        for b in p.get('slots',{}).values():
            if 'target_page' in b and b['target_page'] not in ids:raise ValueError('Unknown slide reference')
    validate_storyboard(plan)
    validate_transitions(plan,require_applied=True)
    return sources,refs,inputs,gaps

def apply_ranking(root, page, refs, facts):
    """Reference 5 ranks: retain row geometry/colors; unused rows removed, capacity 17."""
    data=page.get('ranking')
    if not data:return
    if page['template_slide']!=5:raise ValueError('Ranking component applies to reference slide 5; use bars for paired comparisons on slide 9')
    rows=data['rows']
    if not 1<=len(rows)<=17:raise ValueError('Ranking supports 1–17 rows per page; split larger sets')
    names={s.find('.//p:cNvPr',NS).get('name'):s for s in shapes(root).values()}
    groups=[]
    for i in range(17):
        group=[names.get(f'sc{i}-{suffix}') for suffix in ('n','b','v')]
        if any(x is None for x in group):raise ValueError('Reference ranking topology changed')
        groups.append(group)
    max_width=max(bounds(g[1])[2] for g in groups)
    values=[]
    for row in rows:
        fid=row['fact'];f=facts.get(fid)
        if not f or f['value']<0:raise ValueError('Ranking requires nonnegative bound facts')
        if not row.get('sources') or any(r not in refs for r in row['sources']):raise ValueError('Ranking label needs source')
        values.append(f['value'])
    if values!=sorted(values,reverse=True):raise ValueError('Ranking rows must be descending')
    for i,g in enumerate(groups):
        if i>=len(rows):
            for s in g:s.getparent().remove(s)
        else:
            replace_text(g[0],rows[i]['label']);replace_text(g[2],format(values[i],'g'))
            b=bounds(g[1]);b[2]=max(.01,max_width*values[i]/max(max(values),1));set_bounds(g[1],b)

def build(plan_path, out):
    plan_path=Path(plan_path).resolve();plan=read_json(plan_path);base=plan_path.parent
    sources,refs,inputs,gaps=validate_plan(plan,base)
    out=Path(out).resolve();out.mkdir(parents=True,exist_ok=True)
    lock=out/'.build.lock';fd=os.open(lock,os.O_CREAT|os.O_EXCL|os.O_WRONLY)
    try:
        revision=datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')+'-'+uuid.uuid4().hex[:8]
        dest=out/'revisions'/revision;dest.mkdir(parents=True)
        parts=package(TEMPLATE);original=dict(parts);order=ordered_slides(parts)
        presentation=xml(parts['ppt/presentation.xml']);sldlist=presentation.find('p:sldIdLst',NS)
        for child in list(sldlist):sldlist.remove(child)
        for tag in ['p:custShowLst','p:extLst']:
            for child in presentation.findall(tag,NS):presentation.remove(child)
        presrels=relationships(parts,'ppt/presentation.xml')
        for r in list(presrels):
            if r.get('Type').endswith('/slide'):presrels.remove(r)
        expected=[];notes_md=[];chart_manifest=[];page_index={p['id']:i for i,p in enumerate(plan['pages'],1)}
        media_manifest=validate_asset_bindings(plan,base)['bindings']
        facts=evaluate_plan_facts(plan,refs,sources)
        for n,page in enumerate(plan['pages'],1):
            source=order[page['template_slide']-1];new=f'ppt/slides/personal-slide-{n}.xml'
            root=xml(original[source]);ss=shapes(root);rels=relationships(original,source)
            for key,b in page.get('slots',{}).items():
                node=ss[key]
                if b['action']=='remove':node.getparent().remove(node)
                elif b['action']=='replace':
                    value=b.get('text',b.get('paragraphs'))
                    if b.get('target_page'):
                        if not isinstance(value,str):raise ValueError('Dynamic references require a plain text slot')
                        value=value.replace('{page}',str(page_index[b['target_page']]))
                    replace_text(node,value)
            for key in page.get('remove_shapes',[]):ss[key].getparent().remove(ss[key])
            for key in page.get('data_chart',{}).get('replace_shapes',[]):
                if ss[key].getparent() is not None:ss[key].getparent().remove(ss[key])
            for s in ss.values():
                if s.find('.//p:cNvPr',NS).get('name')=='page-num':replace_text(s,f'{n:02d} / {len(plan["pages"]):02d}')
            for key,b in page.get('pictures',{}).items():
                node=ss[key]
                if b['action']=='remove':node.getparent().remove(node)
                elif b['action']=='replace':
                    path=(base/b['path']).resolve();name=f'ppt/media/personal-{n}-{key.replace("/","-")}{path.suffix.lower()}'
                    parts[name]=path.read_bytes();rid='rIdPersonalImage'+key.replace('/','_')
                    add_rel(rels,rid,'image',pp.relpath(name,pp.dirname(new)))
                    node.find('.//a:blip',NS).set(q('r:embed'),rid)
                    if 'crop' in b:
                        bf=node.find('p:blipFill',NS);crop=bf.find('a:srcRect',NS)
                        if crop is None:crop=E.Element(q('a:srcRect'));bf.insert(1,crop)
                        crop.attrib.clear()
                        for k,v in b['crop'].items():crop.set(k,str(v))
                    mime={'.png':'image/png','.jpg':'image/jpeg','.jpeg':'image/jpeg'}.get(path.suffix.lower())
                    if not mime:raise ValueError('Use PNG/JPEG evidence images')
                    content_type(parts,name,mime)
            apply_ranking(root,page,refs,facts)
            remaining=shapes(root)
            for key,spec in page.get('geometry',{}).items():set_bounds(remaining[key],spec['bounds_pt'])
            for bar in page.get('bars',[]):
                s=remaining[bar['shape']];box=bounds(s);extent=max(.01,bar['extent_pt']*facts[bar['fact']]['value']/bar['maximum'])
                if bar['axis']=='width':box[2]=extent
                else:box[1]+=box[3]-extent;box[3]=extent
                set_bounds(s,box)
            for r in list(rels):
                typ=r.get('Type').split('/')[-1]
                if typ in ('notesSlide','comments','commentAuthors','slide'):rels.remove(r)
                elif typ=='chart':
                    old=target(source,r.get('Target'));chart=f'ppt/charts/personal-{n}.xml'
                    binding=page['chart'];repair_chart(parts,old,chart,binding if binding['action']=='replace' else None)
                    r.set('Target',pp.relpath(chart,pp.dirname(new)))
            if page.get('data_chart'):
                chart_manifest.append(dict(page_index=n,page_id=page['id'],**add_data_chart(
                    parts,root,rels,new,n,page['data_chart'],facts)))
            used_images={b.get(q('r:embed')) for b in root.findall('.//a:blip',NS)}
            for r in list(rels):
                if r.get('Type').endswith('/image') and r.get('Id') not in used_images:rels.remove(r)
            # Old slide hyperlinks cannot silently resolve to a historical slide.
            valid_rids={r.get('Id') for r in rels}
            for h in root.findall('.//a:hlinkClick',NS):
                if h.get(q('r:id')) and h.get(q('r:id')) not in valid_rids:h.getparent().remove(h)
            note='\n\n'.join(refs[r] for r in page.get('notes_refs',[]))
            segments='\n\n'.join(s['quote'] for s in page.get('notes_segments',[]))
            if segments:note=(note+'\n\n'+segments).strip()
            if page.get('notes_extra'):note+='\n\n'+page['notes_extra']
            np=notes_part(parts,new,n,note);add_rel(rels,'rIdPersonalNotes','notesSlide',pp.relpath(np,pp.dirname(new)))
            parts[new]=serialize(root);parts[relpath(new)]=serialize(rels)
            content_type(parts,new,'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
            rid=f'rIdPersonalSlide{n}';add_rel(presrels,rid,'slide',pp.relpath(new,'ppt'))
            E.SubElement(sldlist,q('p:sldId'),id=str(255+n),attrib={q('r:id'):rid})
            expected.append({'index':n,'id':page['id'],'template_slide':page['template_slide'],'part':new,
                             'purpose':page['purpose'],'notes_sha256':sha(note.encode()),
                             'objects':{key:{'signature':style_signature(s),'text':text_of(s),'bounds_pt':bounds(s)} for key,s in shapes(root).items()},
                             'slide_xml_sha256':sha(parts[new])})
            notes_md.append(f'## 第{n}页 · {page["purpose"]}\n\n{note}\n')
        parts['ppt/presentation.xml']=serialize(presentation);parts[relpath('ppt/presentation.xml')]=serialize(presrels)
        rr=relationships(parts,'')
        for r in list(rr):
            if r.get('Type').endswith('/thumbnail'):rr.remove(r)
        parts[relpath('')]=serialize(rr)
        # Generic metadata, no stale source title or last author's name in new output.
        if 'docProps/core.xml' in parts:
            core=xml(parts['docProps/core.xml'])
            for el in core:
                if E.QName(el).localname in ('title','subject','creator','lastModifiedBy','description'):
                    el.text=plan.get('title','演示汇报') if E.QName(el).localname=='title' else ''
            parts['docProps/core.xml']=serialize(core)
        if 'docProps/app.xml' in parts:
            app=xml(parts['docProps/app.xml'])
            for el in list(app):
                if E.QName(el).localname in ('HeadingPairs','TitlesOfParts'):app.remove(el)
                elif E.QName(el).localname in ('Slides','Notes'):el.text=str(len(plan['pages']))
            parts['docProps/app.xml']=serialize(app)
        parts=prune(parts);pptx=dest/'presentation.pptx';save_package(parts,pptx)
        write_json(dest/'plan.json',plan)
        inputs += [{'path':str(plan_path),'sha256':file_sha(plan_path)}, {'path':str(TEMPLATE),'sha256':file_sha(TEMPLATE)}]
        manifest={'schema_version':1,'revision':revision,'created_utc':datetime.now(timezone.utc).isoformat(),
                  'pptx_sha256':file_sha(pptx),'inputs':inputs,'pages':expected,'gaps':gaps,
                  'data_charts':chart_manifest,'media_bindings':media_manifest,
                  'engine':'native-ooxml','status':'built_unverified'}
        write_json(dest/'manifest.json',manifest);write_json(dest/'gaps.json',gaps)
        (dest/'speaking-notes.md').write_text('# 逐页讲述备注\n\n'+'\n'.join(notes_md),encoding='utf-8')
        write_json(out/'current.json',{'revision':revision,'directory':str(dest)})
        return {'revision':revision,'directory':str(dest),'pptx':str(pptx),'status':'built_unverified','gaps':len(gaps)}
    finally:
        os.close(fd);lock.unlink()
