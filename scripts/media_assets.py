"""Trace DOCX image occurrences to source paragraphs, projects and PPT photo slots.

The matcher proposes project attribution; neither alt text nor image recognition
is treated as proof of what an image depicts.  No source document is modified.
"""
from __future__ import annotations

from html import escape
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile
import hashlib
import json
import math
import os
import posixpath
import re
import unicodedata

from lxml import etree as ET
from PIL import Image, UnidentifiedImageError

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
V = 'urn:schemas-microsoft-com:vml'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W, 'r': R, 'a': A, 'wp': WP, 'v': V, 'rel': REL}
SUPPORTED = {'.jpg', '.jpeg', '.png'}
CAPTION_RE = re.compile(r'^(?:图|照片|图片|Fig\.?|Figure)\s*[\d一二三四五六七八九十.-]*[：:.\s-]', re.I)


def _sha(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _file_sha(path: Path) -> str:
    return _sha(path.read_bytes())


def _xml(data: bytes):
    return ET.fromstring(data, ET.XMLParser(resolve_entities=False, no_network=True))


def _write_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    staged = path.with_name('.' + path.name + '.tmp')
    staged.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    os.replace(staged, path)


def _relpath(part: str) -> str:
    return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')


def _text(para) -> str:
    return ''.join(para.xpath('.//w:t/text()', namespaces=NS)).strip()


def _style(para) -> str:
    tag = para.find('w:pPr/w:pStyle', NS)
    return tag.get('{' + W + '}val', '') if tag is not None else ''


def _heading_level(para) -> int | None:
    style = _style(para).lower()
    match = re.search(r'(?:heading|标题)\s*([1-9])', style)
    if match:
        return int(match.group(1))
    level = para.find('w:pPr/w:outlineLvl', NS)
    if level is not None:
        return int(level.get('{' + W + '}val', 9)) + 1
    return None


def _table_context(para) -> dict | None:
    cell = next((x for x in para.iterancestors('{' + W + '}tc')), None)
    if cell is None:
        return None
    row = cell.getparent()
    table = next((x for x in row.iterancestors('{' + W + '}tbl')), None)
    if table is None:
        return None
    tables = list(table.getparent().findall('{' + W + '}tbl'))
    rows = row.getparent().findall('{' + W + '}tr')
    cells = row.findall('{' + W + '}tc')
    cell_text = ' '.join(t for p in cell.findall('.//w:p', NS) if (t := _text(p)))
    row_text = ' '.join(t for c in cells for p in c.findall('.//w:p', NS) if (t := _text(p)))
    return {'table_index': tables.index(table) + 1, 'row_index': rows.index(row) + 1,
            'cell_index': cells.index(cell) + 1, 'cell_text': cell_text[:500],
            'row_text': row_text[:800]}


def _dimensions(raw: bytes) -> tuple[list[int] | None, str | None]:
    try:
        with Image.open(BytesIO(raw)) as image:
            image.verify()
        with Image.open(BytesIO(raw)) as image:
            width, height = image.size
            orientation = image.getexif().get(274, 1)
            if orientation in (5, 6, 7, 8):
                width, height = height, width
            return [width, height], None
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        return None, type(exc).__name__


def _relationships(parts: dict[str, bytes], part: str) -> dict:
    rp = _relpath(part)
    if rp not in parts:
        return {}
    root = _xml(parts[rp])
    return {x.get('Id'): {'target': x.get('Target'), 'external': x.get('TargetMode') == 'External',
                          'type': x.get('Type', '')} for x in root}


def _image_nodes(para):
    for wrapper in para.xpath('.//wp:inline | .//wp:anchor | .//w:pict', namespaces=NS):
        kind = ET.QName(wrapper).localname
        if kind == 'pict':
            images = wrapper.xpath('.//v:imagedata', namespaces=NS)
            for image in images:
                shape = next((x for x in image.iterancestors('{' + V + '}shape')), None)
                alt = shape.get('alt', '') if shape is not None else ''
                yield kind, image.get('{' + R + '}id') or image.get('{' + R + '}href'), alt
        else:
            metadata = wrapper.find('wp:docPr', NS)
            alt = ' '.join(x for x in ((metadata.get('title', ''), metadata.get('descr', '')) if metadata is not None else ()) if x)
            for image in wrapper.xpath('.//a:blip', namespaces=NS):
                yield kind, image.get('{' + R + '}embed') or image.get('{' + R + '}link'), alt


def scan_docx(path: str | Path, source_id: str) -> dict:
    """Inventory every image occurrence, including image-only paragraphs.

    Stable IDs are structural anchors within the immutable document fingerprint.
    The image bytes are not exported until write_inventory is called.
    """
    doc = Path(path).resolve()
    if doc.suffix.lower() != '.docx' or not re.fullmatch(r'[\w-]+', source_id):
        raise ValueError('Expected a DOCX and a short source id of letters, digits, _ or -')
    with ZipFile(doc) as archive:
        parts = {n: archive.read(n) for n in archive.namelist() if not n.endswith('/')}
    stories = ['word/document.xml'] + sorted(p for p in parts if re.fullmatch(r'word/(?:header|footer)\d+\.xml', p))
    entries = []
    for part in stories:
        if part not in parts:
            continue
        root = _xml(parts[part])
        paras = root.findall('.//w:body//w:p', NS) if part == 'word/document.xml' else root.findall('.//w:p', NS)
        rels = _relationships(parts, part)
        texts = [_text(p) for p in paras]
        occurrences = [list(_image_nodes(p)) for p in paras]
        heading_path: list[tuple[int, str]] = []
        for pidx, para in enumerate(paras, 1):
            ptext = texts[pidx - 1]
            level = _heading_level(para)
            if level is not None and ptext:
                heading_path = [(n, t) for n, t in heading_path if n < level] + [(level, ptext)]
            pictures = occurrences[pidx - 1]
            if not pictures:
                continue
            before = next(({'paragraph_index': j + 1, 'text': texts[j], 'style': _style(paras[j])}
                           for j in range(pidx - 2, max(-1, pidx - 5), -1) if texts[j]), None)
            after = next(({'paragraph_index': j + 1, 'text': texts[j], 'style': _style(paras[j])}
                          for j in range(pidx, min(len(paras), pidx + 3)) if texts[j]), None)
            nearby = [x for x in (before, after) if x]
            if ptext:
                nearby.insert(0, {'paragraph_index': pidx, 'text': ptext, 'style': _style(para)})
            captions = []
            for context in nearby:
                if not (CAPTION_RE.search(context['text']) or 'caption' in context['style'].lower() or '题注' in context['style']):
                    continue
                start, end = sorted((pidx, context['paragraph_index']))
                # A caption separated by another picture is not safely tied to
                # this occurrence, even when it is the nearest text paragraph.
                if any(occurrences[j - 1] for j in range(start + 1, end)):
                    continue
                captions.append(context)
            for ordinal, (kind, rid, alt) in enumerate(pictures, 1):
                occurrence = f'{posixpath.basename(part).replace(".xml", "")}-p{pidx:04d}-i{ordinal:02d}'
                rel = rels.get(rid)
                rec = {'asset_id': f'{source_id}:{occurrence}', 'occurrence_id': occurrence,
                       'story_part': part, 'paragraph_index': pidx,
                       'paragraph_anchor': f'{source_id}:p{pidx:03d}' if part == 'word/document.xml' else None,
                       'inline_or_anchor': kind, 'relationship_id': rid, 'alt_text': alt,
                       'paragraph_text': ptext, 'heading_path': [t for _, t in heading_path],
                       'paragraph_before': before, 'paragraph_after': after,
                       'caption_candidates': captions, 'table': _table_context(para)}
                if rel is None:
                    rec.update(status='unresolved_relationship', media_part=None)
                elif rel['external']:
                    rec.update(status='external_link', media_part=None, external_target=rel['target'])
                else:
                    media_part = (rel['target'].lstrip('/') if rel['target'].startswith('/') else
                                  posixpath.normpath(posixpath.join(posixpath.dirname(part), rel['target'])))
                    raw = parts.get(media_part)
                    if raw is None:
                        rec.update(status='missing_media', media_part=media_part)
                    else:
                        extension = Path(media_part).suffix.lower()
                        dimensions, decode_error = _dimensions(raw)
                        rec.update(status='ready' if extension in SUPPORTED and dimensions else 'unsupported_format',
                                   media_part=media_part, media_sha256=_sha(raw),
                                   media_filename=posixpath.basename(media_part), extension=extension,
                                   pixel_size=dimensions, decode_error=decode_error)
                entries.append(rec)
    return {'schema_version': 1, 'source_id': source_id, 'document_path': str(doc),
            'document_sha256': _file_sha(doc), 'assets': entries}


def write_inventory(docx: str | Path, source_id: str, out: str | Path) -> dict:
    out = Path(out).resolve()
    ledger = scan_docx(docx, source_id)
    doc = Path(docx).resolve()
    with ZipFile(doc) as archive:
        for rec in ledger['assets']:
            if not rec.get('media_part') or rec.get('media_sha256') is None:
                continue
            media_name = rec['media_sha256'][:20] + rec['extension']
            extracted = out.parent / (out.stem + '-media') / media_name
            extracted.parent.mkdir(parents=True, exist_ok=True)
            raw = archive.read(rec['media_part'])
            if not extracted.exists() or _sha(extracted.read_bytes()) != rec['media_sha256']:
                extracted.write_bytes(raw)
            rec['extracted_path'] = os.path.relpath(extracted, out.parent)
    ledger['document_path'] = os.path.relpath(doc, out.parent)
    _write_json(out, ledger)
    return ledger


def verify_inventory(path: str | Path) -> dict:
    ledger_path = Path(path).resolve()
    ledger = json.loads(ledger_path.read_text(encoding='utf-8-sig'))
    source = (ledger_path.parent / ledger['document_path']).resolve()
    if _file_sha(source) != ledger['document_sha256']:
        raise ValueError('DOCX fingerprint changed: ' + str(source))
    actual = scan_docx(source, ledger['source_id'])
    old_assets = [{k: v for k, v in asset.items() if k != 'extracted_path'} for asset in ledger['assets']]
    if actual['assets'] != old_assets:
        raise ValueError('Media ledger differs from source DOCX')
    for asset in ledger['assets']:
        if asset.get('extracted_path'):
            extracted = (ledger_path.parent / asset['extracted_path']).resolve()
            if _file_sha(extracted) != asset['media_sha256']:
                raise ValueError('Extracted media fingerprint mismatch: ' + asset['asset_id'])
    return ledger


def _norm(value: str) -> str:
    return ''.join(c for c in unicodedata.normalize('NFKC', value).casefold() if c.isalnum())


def draft_project_catalog(inventory: dict) -> dict:
    """List heading candidates near photos; a reviewer must confirm identities."""
    names = []
    seen = set()
    for asset in inventory['assets']:
        for heading in reversed(asset['heading_path']):
            norm = _norm(heading)
            if norm and norm not in seen:
                seen.add(norm)
                names.append(heading)
                break
    return {'source_doc_sha256': inventory['document_sha256'],
            'reviewed': False,
            'projects': [{'id': f'project-{i:02d}', 'name': name, 'aliases': []}
                         for i, name in enumerate(names, 1)]}


def match_projects(ledger: dict, projects: list[dict]) -> list[dict]:
    """Conservative project suggestions based on document context, not image pixels."""
    ids = [p['id'] for p in projects]
    if len(ids) != len(set(ids)) or any(not p.get('name') for p in projects):
        raise ValueError('Project ids must be unique and names nonempty')
    results = []
    for asset in ledger['assets']:
        fields = []
        fields += [('caption', x['text'], .98) for x in asset['caption_candidates']]
        if asset['paragraph_text']:
            fields.append(('same_paragraph', asset['paragraph_text'], .96))
        if asset['table']:
            fields.append(('same_cell', asset['table']['cell_text'], .94))
            fields.append(('same_row', asset['table']['row_text'], .76))
        fields += [('heading', x, .93) for x in asset['heading_path']]
        fields += [('neighbor', x['text'], .68) for x in (asset['paragraph_before'], asset['paragraph_after']) if x]
        fields += [('alt_text', asset['alt_text'], .72), ('filename', asset.get('media_filename', ''), .4)]
        candidates = []
        for project in projects:
            aliases = [_norm(x) for x in [project['name']] + project.get('aliases', []) if _norm(x)]
            hits = [{'field': field, 'evidence': value[:160], 'weight': weight} for field, value, weight in fields
                    if any(alias in _norm(value) for alias in aliases)]
            if hits:
                score = min(1.0, max(h['weight'] for h in hits) + .02 * (len({h['field'] for h in hits}) - 1))
                candidates.append({'project_id': project['id'], 'score': round(score, 2), 'evidence': hits})
        candidates.sort(key=lambda c: (-c['score'], c['project_id']))
        top = candidates[0] if candidates else None
        runner = candidates[1] if len(candidates) > 1 else None
        strong = bool(top and any(h['field'] in ('caption', 'same_paragraph', 'same_cell', 'heading') for h in top['evidence']))
        conflicted = bool(runner and runner['score'] >= .75)
        chosen = top['project_id'] if asset['status'] == 'ready' and strong and not conflicted and top['score'] >= .9 else None
        results.append({'asset_id': asset['asset_id'], 'status': 'auto' if chosen else 'review',
                        'project_id': chosen, 'candidates': candidates,
                        'reason': ('unique_explicit_context' if chosen else
                                   'unsupported_or_missing_media' if asset['status'] != 'ready' else
                                   'conflicting_project_context' if conflicted else 'insufficient_project_evidence')})
    # The same pixels can be reused in different chapters. Do not silently
    # treat surrounding text as proof that one photo depicts both projects.
    by_sha = {}
    for asset, decision in zip(ledger['assets'], results):
        if asset.get('media_sha256'):
            by_sha.setdefault(asset['media_sha256'], []).append(decision)
    for decisions in by_sha.values():
        named = {d['project_id'] for d in decisions if d['project_id']}
        if len(named) > 1:
            for decision in decisions:
                decision.update(status='review', project_id=None, reason='same_media_conflicting_project_context')
    return results


def write_project_matches(inventory_path: str | Path, project_catalog_path: str | Path,
                          out: str | Path) -> dict:
    inventory_path = Path(inventory_path).resolve()
    project_catalog_path = Path(project_catalog_path).resolve()
    out = Path(out).resolve()
    ledger = verify_inventory(inventory_path)
    catalog = json.loads(project_catalog_path.read_text(encoding='utf-8-sig'))
    if catalog.get('reviewed') is not True:
        raise ValueError('Project catalog needs reviewed:true after checking candidate identities')
    if catalog.get('source_doc_sha256') != ledger['document_sha256']:
        raise ValueError('Project catalog refers to another DOCX revision')
    decisions = match_projects(ledger, catalog['projects'])
    report = {'schema_version': 1, 'inventory_path': os.path.relpath(inventory_path, out.parent),
              'inventory_sha256': _file_sha(inventory_path),
              'project_catalog_path': os.path.relpath(project_catalog_path, out.parent),
              'project_catalog_sha256': _file_sha(project_catalog_path), 'decisions': decisions}
    _write_json(out, report)
    by_id = {a['asset_id']: a for a in ledger['assets']}
    rows = []
    for decision in decisions:
        if decision['status'] != 'review':
            continue
        asset = by_id[decision['asset_id']]
        image = (inventory_path.parent / asset['extracted_path']).resolve().as_uri() if asset.get('extracted_path') and asset['status'] == 'ready' else ''
        candidate = ', '.join(c['project_id'] + ' (' + str(c['score']) + ')' for c in decision['candidates']) or '无'
        context = ' / '.join(asset['heading_path']) + ' · ' + (asset['paragraph_text'] or
                   (asset['paragraph_after'] or {}).get('text', ''))
        rows.append('<article><img src="' + escape(image) + '"><div><b>' + escape(asset['asset_id']) +
                    '</b><p>候选：' + escape(candidate) + '</p><p>上下文：' + escape(context[:250]) +
                    '</p><p>原因：' + escape(decision['reason']) + '</p></div></article>')
    html = '<!doctype html><meta charset="utf-8"><title>待确认项目图片</title><style>' + \
           'body{font:16px sans-serif;max-width:1100px;margin:40px auto;color:#233;}article{display:flex;gap:24px;padding:18px;border-bottom:1px solid #ddd}img{width:260px;height:160px;object-fit:contain;background:#eee}p{margin:6px 0}</style>' + \
           '<h1>待确认项目图片</h1><p>候选仅用于核对；图像本身不构成项目事实证明。</p>' + ''.join(rows)
    out.with_suffix('.review.html').write_text(html, encoding='utf-8')
    return report


def verify_project_matches(path: str | Path) -> tuple[dict, dict]:
    report_path = Path(path).resolve()
    report = json.loads(report_path.read_text(encoding='utf-8-sig'))
    inventory_path = (report_path.parent / report['inventory_path']).resolve()
    catalog_path = (report_path.parent / report['project_catalog_path']).resolve()
    if _file_sha(inventory_path) != report['inventory_sha256'] or _file_sha(catalog_path) != report['project_catalog_sha256']:
        raise ValueError('Project match inputs changed')
    inventory = verify_inventory(inventory_path)
    catalog = json.loads(catalog_path.read_text(encoding='utf-8-sig'))
    if catalog.get('reviewed') is not True:
        raise ValueError('Project catalog is not marked reviewed')
    if catalog.get('source_doc_sha256') != inventory['document_sha256']:
        raise ValueError('Project catalog refers to another DOCX revision')
    if match_projects(inventory, catalog['projects']) != report['decisions']:
        raise ValueError('Project match report differs from inputs')
    return report, inventory


def crop_for_frame(pixel_size: list[int], bounds_pt: list[float]) -> dict[str, int]:
    """Centered geometry-only crop suggestion; visual review must protect the subject."""
    if len(pixel_size) != 2 or min(pixel_size) <= 0 or len(bounds_pt) != 4 or min(bounds_pt[2:]) <= 0:
        raise ValueError('Invalid image size or picture frame')
    source_ratio = pixel_size[0] / pixel_size[1]
    frame_ratio = bounds_pt[2] / bounds_pt[3]
    if source_ratio > frame_ratio:
        side = round((1 - frame_ratio / source_ratio) * 50000)
        return {'l': side, 't': 0, 'r': side, 'b': 0}
    top = round((1 - source_ratio / frame_ratio) * 50000)
    return {'l': 0, 't': top, 'r': 0, 'b': top}


def suggest_photo_layout(asset_ids: list[str], inventory: dict) -> dict:
    """Select native 11/13 photo slots; single-image span is a reviewed variant."""
    if not 1 <= len(asset_ids) <= 3 or len(asset_ids) != len(set(asset_ids)):
        raise ValueError('A native photo page needs 1–3 distinct occurrences')
    assets = {a['asset_id']: a for a in inventory['assets']}
    if any(a not in assets or assets[a]['status'] != 'ready' for a in asset_ids):
        raise ValueError('Photo layout requires extractable JPEG/PNG assets')
    count = len(asset_ids)
    slide = 13 if count == 3 else 11
    boxes = ([[30, 150, 876, 240]] if count == 1 else
             [[30, 150, 428, 240], [478, 150, 428, 240]] if count == 2 else
             [[30, 150, 290, 218], [332, 150, 290, 218], [634, 150, 290, 218]])
    keys = ['34', '37'] if slide == 11 else ['38', '41', '44']
    pictures = {}
    for asset_id, key, box in zip(asset_ids, keys, boxes):
        pictures[key] = {'asset_id': asset_id, 'crop': crop_for_frame(assets[asset_id]['pixel_size'], box),
                         'frame_pt': box, 'crop_review_required': True}
    if count == 1:
        pictures['37'] = {'action': 'remove', 'reason': '单图布局'}
    return {'template_slide': slide, 'pictures': pictures,
            'geometry': {'34': {'bounds_pt': boxes[0], 'reason': '单图横向展开'}} if count == 1 else {},
            'caption_slots': ['10'] if count == 1 else ['10', '13'] if count == 2 else ['10', '13', '16'],
            'unused_caption_slots': ['13'] if count == 1 else [],
            'note': '仅为几何建议；必须检查主体裁切、图注与项目归属，且完整填充其余原生文本槽位。'}


def suggest_project_photo_pages(inventory: dict, match_report: dict) -> dict:
    """Group clear project matches in source order; leave ambiguous images out."""
    assets = {a['asset_id']: a for a in inventory['assets']}
    groups = {}
    review = []
    duplicates = []
    used_sha = {}
    for decision in match_report['decisions']:
        aid = decision['asset_id']
        if aid not in assets:
            raise ValueError('Match report contains unknown media asset')
        if decision['status'] != 'auto':
            review.append(aid)
            continue
        project = decision['project_id']
        if not project:
            raise ValueError('Automatic match has no project')
        key = (project, assets[aid]['media_sha256'])
        if key in used_sha:
            duplicates.append({'asset_id': aid, 'same_media_as': used_sha[key], 'project_id': project})
            continue
        used_sha[key] = aid
        groups.setdefault(project, []).append(aid)
    projects = []
    for project, selected in groups.items():
        pages = [suggest_photo_layout(selected[i:i + 3], inventory) for i in range(0, len(selected), 3)]
        projects.append({'project_id': project, 'asset_ids': selected, 'pages': pages})
    return {'projects': projects, 'review_assets': review, 'duplicate_occurrences': duplicates,
            'note': '仅依据DOCX上下文和图片尺寸生成候选；讲述顺序、图片身份、图注与裁切须实际审阅。'}


def validate_asset_bindings(plan: dict, base: str | Path) -> dict:
    """Optional plan supplement. Legacy plans without media_inventory remain valid."""
    if 'media_inventory' not in plan:
        if any(pic.get('asset_id') for page in plan.get('pages', []) for pic in page.get('pictures', {}).values()):
            raise ValueError('asset_id picture binding needs media_inventory')
        return {'bindings': [], 'inputs': []}
    base = Path(base).resolve()
    inventory_path = (base / plan['media_inventory']).resolve()
    inventory = verify_inventory(inventory_path)
    assets = {a['asset_id']: a for a in inventory['assets']}
    matches = None
    project_ids = set()
    inputs = [{'path': str(inventory_path), 'sha256': _file_sha(inventory_path)},
              {'path': str((inventory_path.parent / inventory['document_path']).resolve()),
               'sha256': inventory['document_sha256']}]
    if plan.get('project_matches'):
        match_path = (base / plan['project_matches']).resolve()
        matches_report, source_inventory = verify_project_matches(match_path)
        if source_inventory['document_sha256'] != inventory['document_sha256']:
            raise ValueError('Project matches use another DOCX')
        matches = {x['asset_id']: x for x in matches_report['decisions']}
        catalog_path = (match_path.parent / matches_report['project_catalog_path']).resolve()
        project_ids = {x['id'] for x in json.loads(catalog_path.read_text(encoding='utf-8-sig'))['projects']}
        inputs.append({'path': str(match_path), 'sha256': _file_sha(match_path)})
    bindings = []
    for page in plan.get('pages', []):
        for key, picture in page.get('pictures', {}).items():
            aid = picture.get('asset_id')
            if not aid:
                continue
            if matches is None:
                raise ValueError('asset_id picture binding needs project_matches for project identity')
            if picture.get('action') != 'replace':
                raise ValueError('asset_id only applies to replacement evidence images')
            asset = assets.get(aid)
            if not asset or asset['status'] != 'ready':
                raise ValueError('Unknown or unsupported asset_id: ' + aid)
            image = (base / picture['path']).resolve()
            extracted = (inventory_path.parent / asset['extracted_path']).resolve()
            if _file_sha(image) != asset['media_sha256'] or _file_sha(extracted) != asset['media_sha256'] or picture.get('sha256') != asset['media_sha256']:
                raise ValueError('Picture does not match DOCX media SHA: ' + aid)
            decision = matches.get(aid) if matches else None
            if picture.get('project_id') not in project_ids:
                raise ValueError('Picture project not found in project catalog: ' + aid)
            if decision and decision['status'] == 'auto':
                if picture.get('project_id') != decision['project_id']:
                    raise ValueError('Picture project differs from supported automatic attribution: ' + aid)
                attribution = 'auto'
            else:
                if not picture.get('project_id') or not picture.get('confirmation_note'):
                    raise ValueError('Ambiguous picture needs manual project confirmation: ' + aid)
                attribution = 'manual'
            bindings.append({'page_id': page['id'], 'shape_key': key, 'asset_id': aid,
                             'project_id': picture['project_id'], 'attribution': attribution,
                             'occurrence_id': asset['occurrence_id'], 'paragraph_anchor': asset['paragraph_anchor'],
                             'story_part': asset['story_part'], 'media_sha256': asset['media_sha256'],
                             'document_sha256': inventory['document_sha256'], 'crop': picture.get('crop')})
    used = {b['asset_id'] for b in bindings}
    dispositions = plan.get('asset_dispositions', {})
    if set(dispositions) != set(assets):
        raise ValueError('Every DOCX image occurrence needs an asset_dispositions record')
    gaps = []
    for aid, decision in dispositions.items():
        if not isinstance(decision, dict) or decision.get('status') not in ('used', 'duplicate', 'omit', 'pending'):
            raise ValueError('Invalid asset disposition: ' + aid)
        if (decision['status'] == 'used') != (aid in used):
            raise ValueError('Asset use/disposition mismatch: ' + aid)
        if decision['status'] != 'used' and not decision.get('reason'):
            raise ValueError('Unused asset disposition needs reason: ' + aid)
        if decision['status'] == 'pending':
            gaps.append({'asset_id': aid, 'blocking': True, 'message': decision['reason']})
    return {'bindings': bindings, 'inputs': inputs, 'gaps': gaps}


def main() -> int:
    import argparse
    parser = argparse.ArgumentParser(description='DOCX project-photo inventory and conservative binding checks')
    sub = parser.add_subparsers(dest='command', required=True)
    cmd = sub.add_parser('inventory', help='Extract all inline, floating and VML image occurrences')
    cmd.add_argument('docx'); cmd.add_argument('--id', required=True); cmd.add_argument('--out', required=True)
    cmd = sub.add_parser('match', help='Match image occurrences to a reviewed project catalog')
    cmd.add_argument('inventory'); cmd.add_argument('--projects', required=True); cmd.add_argument('--out', required=True)
    cmd = sub.add_parser('project-candidates', help='Draft nearby heading candidates, requiring manual review')
    cmd.add_argument('inventory'); cmd.add_argument('--out', required=True)
    cmd = sub.add_parser('layout', help='Suggest 1–3 native photo slots and centered crop')
    cmd.add_argument('inventory'); cmd.add_argument('--assets', required=True, help='Comma-separated asset IDs')
    cmd.add_argument('--out', required=True)
    cmd = sub.add_parser('layout-projects', help='Group confident project photos, excluding ambiguous occurrences')
    cmd.add_argument('matches'); cmd.add_argument('--out', required=True)
    cmd = sub.add_parser('check-bindings', help='Check optional plan media bindings and occurrence dispositions')
    cmd.add_argument('plan')
    args = parser.parse_args()
    if args.command == 'inventory':
        ledger = write_inventory(args.docx, args.id, args.out)
        result = {'inventory': str(Path(args.out).resolve()), 'occurrences': len(ledger['assets']),
                  'ready': sum(a['status'] == 'ready' for a in ledger['assets'])}
    elif args.command == 'match':
        report = write_project_matches(args.inventory, args.projects, args.out)
        result = {'matches': str(Path(args.out).resolve()),
                  'review_html': str(Path(args.out).resolve().with_suffix('.review.html')),
                  'automatic': sum(x['status'] == 'auto' for x in report['decisions']),
                  'review': sum(x['status'] == 'review' for x in report['decisions'])}
    elif args.command == 'project-candidates':
        inventory = verify_inventory(args.inventory)
        result = draft_project_catalog(inventory)
        _write_json(Path(args.out).resolve(), result)
    elif args.command == 'layout':
        inventory = verify_inventory(args.inventory)
        result = suggest_photo_layout([x.strip() for x in args.assets.split(',')], inventory)
        _write_json(Path(args.out).resolve(), result)
    elif args.command == 'layout-projects':
        report, inventory = verify_project_matches(args.matches)
        result = suggest_project_photo_pages(inventory, report)
        _write_json(Path(args.out).resolve(), result)
    else:
        path = Path(args.plan).resolve()
        result = validate_asset_bindings(json.loads(path.read_text(encoding='utf-8-sig')), path.parent)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except Exception as error:
        import sys
        print(json.dumps({'status': 'error', 'error': str(error)}, ensure_ascii=False), file=sys.stderr)
        raise SystemExit(2)
