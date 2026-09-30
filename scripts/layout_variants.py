"""Small plan-only adaptations of native reference-page objects.

These helpers change object geometry or removal bindings. They leave the
reference slide's font runs, fills, shapes, and template resources untouched.
"""

from __future__ import annotations


def three_status_metrics(page: dict) -> None:
    """Show three states in reference slide 3's original metric area.

    The first three icon/value/label trios become three rows. The unused
    fourth trio is removed rather than populated with an unrelated KPI.
    Fill value slots 13, 16, and 19 before applying this variant.
    """
    if page.get("template_slide") != 3:
        raise ValueError("Three-status metrics apply only to reference slide 3")
    slots = page.get("slots")
    if not isinstance(slots, dict) or any(key not in slots for key in ("13", "16", "19", "22", "23")):
        raise ValueError("Three-status metrics require the native metric slots")
    if any(slots[key].get("action") == "replace" for key in ("22", "23")):
        raise ValueError("Fourth metric already contains authored content")

    reason = "原生四指标区按三种建设状态改为纵向三行；保持原图标、字体和数字单位样式"
    geometry = page.setdefault("geometry", {})
    for row, (icon, value, label) in enumerate(((12, 13, 14), (15, 16, 17), (18, 19, 20))):
        y = 154 + 53 * row
        for key, box in ((icon, [410, y + 8, 16, 16]),
                         (value, [434, y, 110, 34]),
                         (label, [572, y + 10, 300, 18])):
            geometry[str(key)] = {"bounds_pt": box, "reason": reason}
    removed = page.setdefault("remove_shapes", [])
    if "21" not in removed:
        removed.append("21")
    for key in ("22", "23"):
        slots[key] = {
            "action": "remove", "kind": "label",
            "reason": "仅存在三种建设状态；不为第四个指标槽创造弱相关数字",
        }
        geometry.pop(key, None)
    page["layout_variant"] = "three-status-native-metrics"
