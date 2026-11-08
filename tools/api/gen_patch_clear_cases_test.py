"""Tests for `gen_patch_clear_cases.py`.

Run: `python3 -m unittest tools/api/gen_patch_clear_cases_test.py`
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import gen_patch_clear_cases as gen  # noqa: E402

DTS = """
export type Patch<T> = { [K in keyof T]?: T[K] | null };
export interface StyleOptions { styleName: string | null; showRowStripes: boolean; }
export type StyleOptionsPatch = Patch<StyleOptions>;
export interface PlainPatch { name?: string; }
"""


def _data(**entries: dict) -> dict:
    return {"schema": 1, "entries": entries}


class ClearCasesTest(unittest.TestCase):
    def test_alias_and_interface_patch_types(self) -> None:
        data = _data(
            a={"js": "StyleOptionsPatch", "clearable": ["style_name"]},
            b={"js": "PlainPatch", "clearable": []},
        )
        self.assertEqual(
            gen.cases(data, DTS),
            {
                "PlainPatch": {"clearable": [], "non_clearable": ["name"]},
                "StyleOptionsPatch": {"clearable": ["style_name"], "non_clearable": ["show_row_stripes"]},
            },
        )

    def test_empty_value_is_an_empty_list(self) -> None:
        got = gen.cases(_data(a={"js": "PlainPatch", "clearable": None}), DTS)
        self.assertEqual(got["PlainPatch"]["clearable"], [])

    def test_entries_without_clearable_are_skipped(self) -> None:
        self.assertEqual(gen.cases(_data(a={"js": "PlainPatch"}), DTS), {})

    def test_unknown_field_is_refused(self) -> None:
        with self.assertRaises(SystemExit):
            gen.cases(_data(a={"js": "PlainPatch", "clearable": ["nope"]}), DTS)


if __name__ == "__main__":
    unittest.main()
