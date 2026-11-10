from __future__ import annotations

import unittest
from xml.etree import ElementTree as ET

from tools.oracle import pivot_refresh_golden as prg


class CanonTest(unittest.TestCase):
    def test_root_save_stamps_and_ext_lst_are_dropped(self) -> None:
        elem = ET.fromstring(
            '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'refreshedDate="46336.1" refreshedBy="me" recordCount="5" missingItemsLimit="0">'
            '<cacheFields count="1"><cacheField name="A" numFmtId="0"/></cacheFields>'
            "<extLst><ext/></extLst></pivotCacheDefinition>"
        )
        self.assertEqual(
            prg._canon(elem, root=True),
            [
                "pivotCacheDefinition",
                {"recordCount": "5", "missingItemsLimit": "0"},
                [["cacheFields", {"count": "1"}, [["cacheField", {"name": "A", "numFmtId": "0"}, []]]]],
            ],
        )

    def test_child_attributes_named_like_root_stamps_are_kept(self) -> None:
        elem = ET.fromstring('<root><item id="3" x="1"/></root>')
        self.assertEqual(prg._canon(elem, root=True)[2][0][1], {"id": "3", "x": "1"})


class DiffTest(unittest.TestCase):
    def test_equal_projections_have_no_diff(self) -> None:
        doc = {"caches": [{"records": [[["x", "0"]]]}], "pivots": [], "formula_cells": {}}
        self.assertEqual(prg.diff_projection(doc, doc), [])

    def test_value_length_and_key_differences_are_reported(self) -> None:
        golden = {"a": [1, 2], "b": {"c": "x"}}
        actual = {"a": [1], "b": {"c": "y", "d": 1}}
        diffs = prg.diff_projection(golden, actual)
        self.assertIn(".a: length 2 != 1", diffs)
        self.assertIn('.b.c: "x" != "y"', diffs)
        self.assertIn(".b.d: present only in actual", diffs)


class EngineGapTest(unittest.TestCase):
    GAPS = {"g1": "reason"}

    def test_listed_case_that_still_differs_is_skipped(self) -> None:
        code, lines = prg.verdict("g1", [".a: 1 != 2"], self.GAPS)
        self.assertEqual(code, prg.SKIP)
        self.assertEqual(lines[0], "ENGINE-GAP g1: reason")

    def test_listed_case_that_matches_fails(self) -> None:
        code, lines = prg.verdict("g1", [], self.GAPS)
        self.assertEqual(code, 1)
        self.assertIn("remove the entry", lines[0])

    def test_unlisted_case_passes_or_fails_on_its_diff(self) -> None:
        self.assertEqual(prg.verdict("c1", [], self.GAPS)[0], 0)
        self.assertEqual(prg.verdict("c1", ["x"], self.GAPS)[0], 1)

    def test_gap_list_must_name_cases_and_carry_reasons(self) -> None:
        problems = prg.engine_gap_problems({"gone": "r", "c1": " "}, ["c1"])
        self.assertEqual(problems, ["engine gap 'gone' names no case", "engine gap 'c1' has no reason"])

    def test_committed_gap_list_names_existing_cases(self) -> None:
        names = [d.name for d in prg._case_dirs([])]
        self.assertEqual(prg.engine_gap_problems(prg.load_engine_gaps(), names), [])


class RefTest(unittest.TestCase):
    def test_parse_ref_handles_single_cells_and_absolute_markers(self) -> None:
        self.assertEqual(prg._parse_ref("A3"), (3, 1, 3, 1))
        self.assertEqual(prg._parse_ref("$B$2:$AA$10"), (2, 2, 10, 27))


if __name__ == "__main__":
    unittest.main()
