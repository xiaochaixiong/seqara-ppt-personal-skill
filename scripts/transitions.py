"""Evidence-aware beat transitions for native personal-PPT plans.

This module does not infer a rhetorical turn from page count or visual sameness.
It exposes beat boundaries and applies only decisions explicitly made in a plan.
"""
from __future__ import annotations

from copy import deepcopy
import re


RELATIONS = {
    'question_answer', 'evidence_judgment', 'diagnosis_action',
    'case_rule', 'known_unknown', 'speaker_shift', 'section_shift', 'other',
}
TIERS = {'micro', 'content', 'pause'}
CONTENT_TEMPLATES = {2, 18, 23, 26, 29}
PAUSE_TEMPLATES = {17, 25}


def _words(value):
    return isinstance(value, str) and bool(value.strip())


def _binding_text(binding):
    if binding.get('action') == 'remove':
        return ''
    if 'text' in binding:
        return str(binding['text'])
    return '\n'.join(''.join(str(run.get('text', '')) for run in para.get('runs', []))
                     for para in binding.get('paragraphs', []))


def beat_boundaries(plan):
    """Expose adjacent narrative beats and their actual page positions.

    The storyboard, when present, is the source of beat order. A boundary
    inside one page is included but cannot receive an inserted page until the
    author splits that page. No semantic tier is guessed here.
    """
    story = plan.get('storyboard')
    if not story:
        return []
    beats = story.get('beats', [])
    assignments = story.get('page_assignments', [])
    page_positions = {page['id']: index for index, page in enumerate(plan.get('pages', []))}
    locations = {beat['id']: [] for beat in beats}
    for assignment in assignments:
        bid, pid = assignment.get('beat_id'), assignment.get('page_id')
        if bid in locations and pid in page_positions:
            locations[bid].append((page_positions[pid], pid))
    result = []
    for previous, following in zip(beats, beats[1:]):
        left = locations.get(previous['id'], [])
        right = locations.get(following['id'], [])
        if not left or not right:
            continue
        left_index, left_page = max(left)
        right_index, right_page = min(right)
        result.append({
            'from_beat': previous['id'], 'to_beat': following['id'],
            'from_page_id': left_page, 'to_page_id': right_page,
            'same_page': left_index == right_index,
            'insertable': right_index == left_index + 1,
        })
    return result


def suggest_transitions(plan):
    """Return review prompts, preserving any authored decisions.

    `candidate` means that a human/model still has to state the relationship,
    audience question and necessity. The default is deliberately *no* new page.
    """
    authored = {(item.get('from_beat'), item.get('to_beat')): item
                for item in plan.get('transitions', [])}
    result = []
    for boundary in beat_boundaries(plan):
        key = boundary['from_beat'], boundary['to_beat']
        item = dict(boundary)
        item.update(authored.get(key, {}))
        item['status'] = 'authored' if key in authored else 'candidate'
        if key not in authored:
            item['tier'] = 'micro'
            item['review_prompt'] = (
                '先写听众此刻的问题和前后论点的关系；只在新判断必须独立呈现时考虑内容页，'
                '重大转折且删去会损失必要判断时才考虑停顿页。')
        result.append(item)
    return result


def _section_bindings(spec):
    number = spec.get('number')
    if number is not None and (isinstance(number, bool) or not isinstance(number, int)
                               or not 1 <= number <= 99):
        raise ValueError('Section number must be an integer from 1 to 99')
    label = spec.get('label')
    if not _words(label):
        raise ValueError('Section label is required')
    if number is None:
        ghost = {'action': 'remove', 'kind': 'label', 'reason': 'No chapter number in this transition'}
        tag = label.strip()
    else:
        ghost = {'action': 'replace', 'kind': 'label', 'text': f'{number:02d}', 'sources': []}
        tag = f'PART {number:02d} · {label.strip()}'
    return ghost, {'action': 'replace', 'kind': 'label', 'text': tag, 'sources': []}


def _explanation_paragraphs(text, highlight, slide):
    if highlight is not None and (not _words(highlight) or text.count(highlight) != 1):
        raise ValueError('Pause explanation_highlight must appear exactly once')
    paragraphs = []
    for line in text.split('\n'):
        if highlight and highlight in line:
            before, after = line.split(highlight, 1)
            runs = []
            if before:
                runs.append({'text': before, 'style_run': 0})
            runs.append({'text': highlight, 'style_run': 2 if slide == 17 else 1})
            if after:
                runs.append({'text': after, 'style_run': 0})
        else:
            runs = [{'text': line, 'style_run': 0}]
        paragraphs.append({'runs': runs})
    return paragraphs


def make_pause_page(edge):
    """Create a native T17/T25 pause, replacing every historical text slot.

    Page 17 has a page-number object, updated by the regular build pipeline;
    page 25 does not. Their ghost-number keys differ, hence explicit mapping.
    """
    slide = edge.get('template_slide', 17)
    if slide not in PAUSE_TEMPLATES:
        raise ValueError('Pause must use native reference slide 17 or 25')
    content = edge.get('content')
    if not isinstance(content, dict):
        raise ValueError('Pause requires authored content')
    refs = content.get('sources', [])
    if not isinstance(refs, list) or not refs:
        raise ValueError('Pause claim needs source references')
    for field in ('headline', 'explanation'):
        if not _words(content.get(field)):
            raise ValueError('Pause needs '+field)
    kind = content.get('kind', 'inference')
    if kind not in ('source', 'inference', 'proposal'):
        raise ValueError('Pause content kind must be source, inference or proposal')
    section = edge.get('section')
    if not isinstance(section, dict):
        raise ValueError('Pause requires section label/number decision')
    ghost, tag = _section_bindings(section)
    ghost_key = '3' if slide == 17 else '4'
    slots = {
        ghost_key: ghost,
        '5': tag,
        '6': {'action': 'replace', 'kind': kind,
              'text': content['headline'], 'sources': refs},
        '8': {'action': 'replace', 'kind': kind,
              'paragraphs': _explanation_paragraphs(
                  content['explanation'], content.get('explanation_highlight'), slide),
              'sources': refs},
    }
    if _words(content.get('evidence_line')):
        slots['10'] = {'action': 'replace', 'kind': kind,
                       'text': content['evidence_line'], 'sources': refs}
    else:
        slots['10'] = {'action': 'remove', 'kind': 'label',
                       'reason': 'No distinct evidence line for this pause'}
    return {
        'id': edge['page_id'], 'template_slide': slide,
        'purpose': edge['bridge'], 'transition_id': edge['id'],
        'section': deepcopy(section), 'slots': slots, 'pictures': {},
        'notes_refs': [], 'notes_extra': (
            f"转场：{edge['bridge']}\n听众问题：{edge['audience_question']}\n"
            f"独立成页的必要性：{edge['if_removed_loss']}")
    }


def _find_boundary(plan, edge):
    for boundary in beat_boundaries(plan):
        if (boundary['from_beat'], boundary['to_beat']) == (edge['from_beat'], edge['to_beat']):
            return boundary
    raise ValueError(f"Transition {edge['id']} does not join adjacent storyboard beats")


def validate_transitions(plan, require_applied=False):
    """Check authored transition choices without changing older plans.

    This validates necessary metadata and native section labels, not whether
    the asserted rhetorical need is actually convincing. That remains a deck
    review decision after rendering.
    """
    edges = plan.get('transitions')
    if edges is None:
        return {'transition_count': 0, 'added_pages': 0}
    if not isinstance(edges, list):
        raise ValueError('Transitions must be a list')
    ids, pairs = set(), set()
    page_by_id = {page['id']: page for page in plan.get('pages', [])}
    if len(page_by_id) != len(plan.get('pages', [])):
        raise ValueError('Duplicate page ids')
    added = 0
    page_positions = {page['id']: index for index, page in enumerate(plan.get('pages', []))}
    for edge in edges:
        if not isinstance(edge, dict) or not _words(edge.get('id')):
            raise ValueError('Transition needs an id')
        if edge['id'] in ids:
            raise ValueError('Duplicate transition id')
        ids.add(edge['id'])
        pair = edge.get('from_beat'), edge.get('to_beat')
        if pair in pairs:
            raise ValueError('Duplicate transition boundary')
        pairs.add(pair)
        boundary = _find_boundary(plan, edge)
        if edge.get('relation') not in RELATIONS:
            raise ValueError('Transition needs a known rhetorical relation')
        for field in ('audience_question', 'bridge', 'rationale'):
            if not _words(edge.get(field)):
                raise ValueError('Transition needs '+field)
        tier = edge.get('tier')
        if tier not in TIERS:
            raise ValueError('Transition tier must be micro, content or pause')
        if tier == 'micro':
            if edge.get('page_id'):
                raise ValueError('Micro transition cannot create a page')
            if edge.get('applied'):
                anchor = page_by_id[boundary['from_page_id']]
                if anchor.get('transition_to_next') != edge['bridge'] or ('转场：'+edge['bridge']) not in anchor.get('notes_extra', ''):
                    raise ValueError('Applied micro transition bridge is missing from the anchor page')
            elif require_applied:
                raise ValueError('Transition decision has not been applied: '+edge['id'])
            continue
        if not _words(edge.get('if_removed_loss')):
            raise ValueError('Added transition page needs if_removed_loss')
        if boundary['same_page'] or not boundary['insertable']:
            raise ValueError('Added transition requires adjacent separate beat pages')
        seconds = edge.get('seconds')
        if isinstance(seconds, bool) or not isinstance(seconds, (int, float)) or not 0 < seconds < 120:
            raise ValueError('Added transition needs 0 < seconds < 120')
        pid = edge.get('page_id')
        if not _words(pid):
            raise ValueError('Added transition needs page_id')
        if tier == 'pause' and edge.get('significance') != 'major_pivot':
            raise ValueError('Pause requires a declared major_pivot')
        if tier == 'content' and edge.get('significance') not in ('new_judgment', 'major_pivot'):
            raise ValueError('Content transition requires an independent new judgment')
        if edge.get('applied'):
            page = page_by_id.get(pid)
            if page is None or page.get('transition_id') != edge['id']:
                raise ValueError('Applied transition page is missing or mismatched')
            if (boundary['to_page_id'] != pid or
                    page_positions[pid] != page_positions[boundary['from_page_id']] + 1):
                raise ValueError('Applied transition is not at its beat boundary')
            assignments = plan['storyboard']['page_assignments']
            if not any(item['page_id'] == pid and item['beat_id'] == edge['to_beat']
                       and item['seconds'] == seconds for item in assignments):
                raise ValueError('Applied transition time allocation is missing')
            if tier == 'pause':
                _validate_pause_section(page)
            elif page.get('template_slide') not in CONTENT_TEMPLATES:
                raise ValueError('Content transition uses unsupported reference slide')
            added += 1
        elif pid in page_by_id:
            raise ValueError('Unapplied transition page_id already exists')
        elif require_applied:
            raise ValueError('Transition decision has not been applied: '+edge['id'])
    return {'transition_count': len(edges), 'added_pages': added}


def _validate_pause_section(page):
    slide = page.get('template_slide')
    if slide not in PAUSE_TEMPLATES:
        raise ValueError('Pause transition must use reference 17/25')
    section = page.get('section', {})
    ghost, tag = _section_bindings(section)
    ghost_key = '3' if slide == 17 else '4'
    for key, expected in ((ghost_key, ghost), ('5', tag)):
        actual = page.get('slots', {}).get(key)
        if actual is None or actual.get('action') != expected['action']:
            raise ValueError('Pause ghost/section label was not updated')
        if expected['action'] == 'replace' and _binding_text(actual) != expected['text']:
            raise ValueError('Pause ghost/section label is inconsistent')
    if not _binding_text(page.get('slots', {}).get('6', {})).strip():
        raise ValueError('Pause headline is empty')


def apply_transitions(plan):
    """Apply approved choices to a copy of the plan; preserve speaking time.

    Content transitions carry an authored native page in `edge.page`. Pause
    transitions generate reference 17/25 from explicit sourced content.
    Existing beat seconds are reallocated from the next beat's first page,
    never silently added to the talk.
    """
    result = deepcopy(plan)
    validate_transitions(result)
    story = result.get('storyboard')
    if not story and result.get('transitions'):
        raise ValueError('Transition application needs storyboard')
    for edge in result.get('transitions', []):
        if edge.get('applied'):
            continue
        boundary = _find_boundary(result, edge)
        if edge['tier'] == 'micro':
            anchor = next(page for page in result['pages'] if page['id'] == boundary['from_page_id'])
            bridge = edge['bridge'].strip()
            anchor['transition_to_next'] = bridge
            note = anchor.get('notes_extra', '')
            addition = '转场：'+bridge
            if addition not in note:
                anchor['notes_extra'] = (note+'\n\n'+addition).strip()
            edge['applied'] = True
            continue
        if edge['tier'] == 'pause':
            page = make_pause_page(edge)
        else:
            page = deepcopy(edge.get('page'))
            if not isinstance(page, dict) or page.get('id') != edge['page_id']:
                raise ValueError('Content transition needs authored page with matching id')
            if page.get('template_slide') not in CONTENT_TEMPLATES:
                raise ValueError('Content transition must use a reviewed native content layout')
            page['transition_id'] = edge['id']
        next_id = boundary['to_page_id']
        donor = next((assignment for assignment in story['page_assignments']
                      if assignment['page_id'] == next_id and assignment['beat_id'] == edge['to_beat']), None)
        if donor is None or donor['seconds'] <= edge['seconds']:
            raise ValueError('Next beat page has insufficient time to fund transition')
        donor['seconds'] -= edge['seconds']
        index = next(i for i, item in enumerate(result['pages']) if item['id'] == next_id)
        result['pages'].insert(index, page)
        story['page_assignments'].insert(
            next(i for i, assignment in enumerate(story['page_assignments']) if assignment is donor),
            {'page_id': page['id'], 'beat_id': edge['to_beat'], 'seconds': edge['seconds']})
        edge.pop('page', None)
        edge['applied'] = True
    validate_transitions(result)
    return result
