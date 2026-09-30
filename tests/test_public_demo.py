"""Regression checks for the distributable, synthetic template and example."""
from pathlib import Path
from tempfile import TemporaryDirectory
import json
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))

from native_pptx import TEMPLATE, file_sha, inventory, package
from pipeline import build, validate_plan
from quality import structure_check


class PublicDemoTests(unittest.TestCase):
    def test_all_38_template_pages_have_a_matching_profile(self):
        extracted = inventory()
        saved = json.loads((ROOT / 'assets/profile.json').read_text(encoding='utf-8'))
        self.assertEqual(len(extracted['slides']), 38)
        self.assertEqual(saved['template_sha256'], file_sha(TEMPLATE))
        self.assertEqual(
            [[slot['key'] for slot in page['slots']] for page in extracted['slides']],
            [[slot['key'] for slot in page['slots']] for page in saved['slides']],
        )

    def test_template_omits_embedded_font_and_local_paths(self):
        parts = package(TEMPLATE)
        self.assertFalse(any(name.endswith('.fntdata') for name in parts))
        text = '\n'.join(value.decode('utf-8', errors='ignore') for value in parts.values())
        for forbidden in ('C:\\Users\\', 'D:\\Coedx\\', 'PRIVATE_TEMPLATE', 'CONFIDENTIAL_SOURCE'):
            with self.subTest(term=forbidden):
                self.assertNotIn(forbidden, text)

    def test_synthetic_plan_builds_an_editable_deck(self):
        plan_path = ROOT / 'examples/synthetic-data/plan.json'
        plan = json.loads(plan_path.read_text(encoding='utf-8'))
        validate_plan(plan, plan_path.parent)
        with TemporaryDirectory() as temp:
            result = build(str(plan_path), temp)
            report = structure_check(result['directory'])
        self.assertEqual(report['errors'], [])
        self.assertEqual(report['pages'], 2)
        self.assertEqual(len(report['charts']), 2)
        self.assertTrue(all(chart['embedded_workbook'] for chart in report['charts']))


if __name__ == '__main__':
    unittest.main()
