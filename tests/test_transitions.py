import copy
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'scripts'))
from transitions import (apply_transitions, beat_boundaries, make_pause_page,
                         suggest_transitions, validate_transitions)
from pipeline import validate_storyboard


def sample_plan():
    return {
        'pages': [{'id': 'evidence', 'template_slide': 8, 'purpose': '现状',
                   'notes_extra': '先说明依据'},
                  {'id': 'action', 'template_slide': 26, 'purpose': '下一步'}],
        'storyboard': {
            'duration_seconds': 90,
            'beats': [{'id': 'B1', 'expected_seconds': 40},
                      {'id': 'B2', 'expected_seconds': 50}],
            'page_assignments': [
                {'page_id': 'evidence', 'beat_id': 'B1', 'seconds': 40},
                {'page_id': 'action', 'beat_id': 'B2', 'seconds': 50}],
        },
    }


def pause_edge():
    return {
        'id': 'turn-1', 'from_beat': 'B1', 'to_beat': 'B2',
        'relation': 'diagnosis_action',
        'audience_question': '已有证据之后，下一步做什么？',
        'bridge': '证据提示下一步应优先解决执行问题',
        'rationale': '从诊断转向行动，听众需要先理解行动原则。',
        'tier': 'pause', 'significance': 'major_pivot',
        'if_removed_loss': '失去从诊断推导行动原则的独立判断。',
        'seconds': 12, 'page_id': 'decision-turn',
        'section': {'number': 2, 'label': '下一步行动'},
        'content': {'headline': '先解决执行条件，\n再安排项目扩展',
                    'explanation': '现有证据只支持先明确职责与边界，后续项目量仍需核对。',
                    'evidence_line': '依据：制度条款与项目清单',
                    'kind': 'inference', 'sources': ['rules:p001']},
    }


class TransitionsTest(unittest.TestCase):
    def test_candidates_do_not_invent_pause_pages(self):
        plan = sample_plan()
        self.assertEqual(beat_boundaries(plan)[0]['from_page_id'], 'evidence')
        suggestion = suggest_transitions(plan)[0]
        self.assertEqual(suggestion['status'], 'candidate')
        self.assertEqual(suggestion['tier'], 'micro')
        self.assertNotIn('transitions', plan)

    def test_micro_adds_only_speaker_bridge(self):
        plan = sample_plan()
        edge = pause_edge()
        edge['tier'] = 'micro'
        for key in ('page_id', 'seconds', 'section', 'content', 'significance', 'if_removed_loss'):
            edge.pop(key, None)
        plan['transitions'] = [edge]
        result = apply_transitions(plan)
        self.assertEqual(len(result['pages']), 2)
        self.assertEqual(result['pages'][0]['transition_to_next'], edge['bridge'])
        self.assertIn('转场：'+edge['bridge'], result['pages'][0]['notes_extra'])
        self.assertEqual(result['storyboard'], plan['storyboard'])
        self.assertTrue(result['transitions'][0]['applied'])
        self.assertEqual(apply_transitions(result), result)

    def test_native_pause_reallocates_time_and_updates_section(self):
        plan = sample_plan()
        plan['transitions'] = [pause_edge()]
        with self.assertRaisesRegex(ValueError, 'has not been applied'):
            validate_transitions(plan, require_applied=True)
        result = apply_transitions(plan)
        self.assertEqual([page['id'] for page in result['pages']],
                         ['evidence', 'decision-turn', 'action'])
        page = result['pages'][1]
        self.assertEqual(page['template_slide'], 17)
        self.assertEqual(page['slots']['3']['text'], '02')
        self.assertEqual(page['slots']['5']['text'], 'PART 02 · 下一步行动')
        self.assertNotIn('上半年', repr(page))
        self.assertEqual([item['seconds'] for item in result['storyboard']['page_assignments']],
                         [40, 12, 38])
        self.assertEqual(validate_storyboard(result)['duration_seconds'], 90)
        self.assertEqual(validate_transitions(result)['added_pages'], 1)
        self.assertEqual(validate_transitions(result, require_applied=True)['added_pages'], 1)

    def test_t25_ghost_and_unnumbered_section(self):
        edge = pause_edge()
        edge['template_slide'] = 25
        page = make_pause_page(edge)
        self.assertEqual(page['slots']['4']['text'], '02')
        self.assertEqual(page['slots']['5']['text'], 'PART 02 · 下一步行动')
        edge['section'] = {'label': '参与前请确认'}
        page = make_pause_page(edge)
        self.assertEqual(page['slots']['4']['action'], 'remove')
        self.assertEqual(page['slots']['5']['text'], '参与前请确认')

    def test_content_transition_requires_authored_native_page(self):
        plan = sample_plan()
        edge = pause_edge()
        edge['tier'] = 'content'
        edge['significance'] = 'new_judgment'
        edge['page'] = {'id': 'decision-turn', 'template_slide': 23,
                        'purpose': edge['bridge'], 'slots': {}, 'pictures': {}}
        plan['transitions'] = [edge]
        result = apply_transitions(plan)
        self.assertEqual(result['pages'][1]['template_slide'], 23)
        self.assertEqual(result['pages'][1]['transition_id'], edge['id'])

    def test_no_arbitrary_or_unfunded_pause(self):
        plan = sample_plan()
        edge = pause_edge()
        edge['significance'] = 'ordinary'
        plan['transitions'] = [edge]
        with self.assertRaisesRegex(ValueError, 'major_pivot'):
            validate_transitions(plan)
        edge['significance'] = 'major_pivot'
        edge['seconds'] = 50
        with self.assertRaisesRegex(ValueError, 'insufficient time'):
            apply_transitions(plan)
        edge['seconds'] = 12
        plan['storyboard']['page_assignments'][1]['page_id'] = 'evidence'
        with self.assertRaisesRegex(ValueError, 'adjacent separate'):
            validate_transitions(plan)

    def test_inconsistent_native_chapter_label_rejected(self):
        plan = sample_plan()
        plan['transitions'] = [pause_edge()]
        result = apply_transitions(plan)
        result['pages'][1]['slots']['3']['text'] = '03'
        with self.assertRaisesRegex(ValueError, 'ghost/section label'):
            validate_transitions(result)
        result = apply_transitions(plan)
        result['pages'][1]['slots']['5']['text'] = 'PART 01 · 旧时期内容'
        with self.assertRaisesRegex(ValueError, 'ghost/section label'):
            validate_transitions(result)

    def test_applied_transition_must_stay_at_boundary(self):
        plan = sample_plan()
        plan['transitions'] = [pause_edge()]
        result = apply_transitions(plan)
        result['pages'].append(result['pages'].pop(1))
        with self.assertRaisesRegex(ValueError, 'not at its beat boundary'):
            validate_transitions(result)


if __name__ == '__main__':
    unittest.main()
