#!/usr/bin/env python3
"""Fixture tests for the sheet-kind-guard check.

Run from the repo root: `python3 -m tools.dev.binding_drift.sheet_kind_guard_test`.
"""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from tools.dev.binding_drift.sheet_kind_guard import check_sheet_kind_guard

ENTRIES = {
    "kind.get": {"kind": "op", "scope": "sheet", "c": "fm_sheet_get_kind", "sheet_kinds": "all"},
    "value.get": {"kind": "op", "scope": "sheet", "c": "fm_sheet_get_value", "sheet_kinds": "worksheet"},
    "merge.get_count": {
        "kind": "op",
        "scope": "sheet",
        "c": "fm_sheet_get_merge_count",
        "sheet_kinds": "worksheet_or_chartsheet",
    },
    "gone": {"kind": "op", "scope": "sheet", "c": "fm_sheet_gone", "drop": {"c": "fm_sheet_gone"}},
}

GUARD = "  if (auto rc = check_sheet_use(wb, sheet_index, SheetUse::%s); rc != 0) {\n    return rc;\n  }\n"


def source(kind_guard: str = "", value_use: str = "kWorksheet", count_use: str = "kWorksheetOrChartsheet") -> str:
    return (
        'extern "C" fm_status_t fm_sheet_get_kind(const fm_workbook_t* wb, size_t sheet_index, int32_t* out) {\n'
        + kind_guard
        + "  return 0;\n}\n\n"
        + 'extern "C" fm_status_t fm_sheet_get_value(const fm_workbook_t* wb, size_t sheet_index, fm_value_t* out) {\n'
        + (GUARD % value_use if value_use else "")
        + "  return 0;\n}\n\n"
        + 'extern "C" fm_status_t fm_sheet_get_merge_count(const fm_workbook_t* wb, size_t sheet_index, size_t* n) {\n'
        + (GUARD % count_use if count_use else "")
        + "  return 0;\n}\n"
    )


def run(text: str) -> list:
    with TemporaryDirectory() as tmp:
        parts = Path(tmp) / "parts"
        parts.mkdir()
        (Path(tmp) / "generated").mkdir()
        (parts / "a.cpp").write_text(text, encoding="utf-8")
        return check_sheet_kind_guard(Path(tmp), ENTRIES)


class SheetKindGuardTest(unittest.TestCase):
    def test_matching_bodies_pass(self) -> None:
        self.assertEqual(run(source()), [])

    def test_all_entry_with_a_guard_fails(self) -> None:
        problems = run(source(kind_guard=GUARD % "kWorksheet"))
        self.assertEqual(len(problems), 1)
        self.assertIn("fm_sheet_get_kind", problems[0])

    def test_missing_guard_fails(self) -> None:
        problems = run(source(value_use=""))
        self.assertEqual(len(problems), 1)
        self.assertIn("fm_sheet_get_value", problems[0])

    def test_wrong_enumerator_fails(self) -> None:
        problems = run(source(count_use="kWorksheet"))
        self.assertEqual(len(problems), 1)
        self.assertIn("kWorksheetOrChartsheet", problems[0])

    def test_missing_body_fails(self) -> None:
        problems = run(source().replace("fm_sheet_get_merge_count", "fm_sheet_get_other"))
        self.assertEqual(len(problems), 1)
        self.assertIn("fm_sheet_get_merge_count", problems[0])


if __name__ == "__main__":
    unittest.main()
