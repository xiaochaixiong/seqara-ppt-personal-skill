"""Regression coverage for native page adaptations used by personal reports."""

from pathlib import Path
from tempfile import TemporaryDirectory
import sys
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))

from layout_variants import three_status_metrics
from native_pptx import (NS, TEMPLATE, bounds, file_sha,
                         inventory, ordered_slides, package, serialize,
                         shapes, text_of, write_json, xml)
from pipeline import build, extract_source, validate_plan


class NativeLayoutVariantTests(unittest.TestCase):
    def test_three_status_removes_weak_fourth_metric_without_font_changes(self):
        with TemporaryDirectory() as temp:
            base = Path(temp)
            write_json(base / "source.json", extract_source(TEMPLATE, "reference", "evidence"))
            ref = "reference:s003"
            slots = {}
            for shape in inventory()["slides"][2]["slots"]:
                if shape["text"] and shape["name"] != "page-num" and shape["type"] != "grpSp":
                    slots[shape["key"]] = {"action": "keep", "kind": "source", "sources": [ref]}
            page = {
                "id": "three-status", "template_slide": 3,
                "purpose": "Three meaningful states use three native metric rows",
                "slots": slots, "pictures": {}, "notes_refs": [ref],
            }
            plan = {
                "schema_version": 1, "template_sha256": file_sha(TEMPLATE),
                "sources": ["source.json"], "facts": [], "gaps": [], "pages": [page],
            }
            three_status_metrics(page)
            self.assertEqual(page["remove_shapes"], ["21"])
            self.assertEqual([page["slots"][key]["action"] for key in ("22", "23")],
                             ["remove", "remove"])
            self.assertEqual(page["layout_variant"], "three-status-native-metrics")
            validate_plan(plan, base)
            plan_path = base / "plan.json"
            write_json(plan_path, plan)
            out = Path(build(plan_path, base / "out")["directory"])

            native_parts = package(TEMPLATE)
            native = shapes(xml(native_parts[ordered_slides(native_parts)[2]]))
            built_parts = package(out / "presentation.pptx")
            built = shapes(xml(built_parts[ordered_slides(built_parts)[0]]))
            for key in ("21", "22", "23"):
                self.assertNotIn(key, built)
            for key in ("13", "16", "19"):
                original_runs = native[key].findall(".//a:r", NS)
                output_runs = built[key].findall(".//a:r", NS)
                self.assertEqual(text_of(built[key]), text_of(native[key]))
                self.assertEqual([serialize(r.find("a:rPr", NS)) for r in output_runs],
                                 [serialize(r.find("a:rPr", NS)) for r in original_runs])
                self.assertEqual([r.find("a:rPr", NS).get("sz") for r in output_runs],
                                 ["2200", "1100"])
            self.assertEqual(bounds(built["13"]), [434, 154, 110, 34])
            self.assertEqual(bounds(built["16"]), [434, 207, 110, 34])
            self.assertEqual(bounds(built["19"]), [434, 260, 110, 34])

            authored = dict(page, slots=dict(page["slots"]))
            authored["slots"]["22"] = {"action": "replace", "kind": "inference", "text": "extra"}
            with self.assertRaisesRegex(ValueError, "Fourth metric"):
                three_status_metrics(authored)


if __name__ == "__main__":
    unittest.main()
