from __future__ import annotations

import unittest

from tools.oracle import golden_diff


def _doc(cases, tolerance=None):
    return {"tolerance": tolerance or {"abs": 0.0, "rel": 0.0}, "cases": cases}


class GoldenDiffTest(unittest.TestCase):
    def test_numbers_within_the_new_suite_tolerance_agree(self) -> None:
        old = {"s": _doc([{"id": "a", "expect": {"kind": "number", "value": 1.0}}])}
        new = {"s": _doc([{"id": "a", "expect": {"kind": "number", "value": 1.0 + 1e-13}}], {"abs": 1e-12})}
        self.assertEqual(golden_diff.diff_sides(old, new), [])

    def test_kinds_and_skips_are_classified(self) -> None:
        old = {
            "s": _doc(
                [
                    {"id": "a", "expect": {"kind": "text", "value": "x"}},
                    {"id": "b", "expect": {"kind": "number", "value": 2.0}},
                    {"id": "c", "skipped": "reason"},
                    {"id": "d", "expect": {"kind": "bool", "value": True}},
                ]
            )
        }
        new = {
            "s": _doc(
                [
                    {"id": "a", "expect": {"kind": "text", "value": "y"}},
                    {"id": "b", "skipped": "reason"},
                    {"id": "c", "expect": {"kind": "number", "value": 1.0}},
                    {"id": "e", "expect": {"kind": "number", "value": 1.0}},
                ]
            )
        }
        changes = {(r[1], r[2]) for r in golden_diff.diff_sides(old, new)}
        self.assertEqual(
            changes, {("a", "changed"), ("b", "skipped"), ("c", "unskipped"), ("d", "removed"), ("e", "added")}
        )

    def test_arrays_compare_shape_and_elements(self) -> None:
        old = {"s": _doc([{"id": "a", "expect": {"kind": "array", "value": [1.0, None], "shape": [2, 1]}}])}
        new = {"s": _doc([{"id": "a", "expect": {"kind": "array", "value": [1.0, ""], "shape": [2, 1]}}])}
        self.assertEqual([r[2] for r in golden_diff.diff_sides(old, new)], ["changed"])

    def test_one_sided_suites_are_skipped_by_default(self) -> None:
        old = {"s": _doc([{"id": "a", "expect": {"kind": "number", "value": 1.0}}])}
        self.assertEqual(golden_diff.diff_sides(old, {}), [])
        self.assertEqual([r[2] for r in golden_diff.diff_sides(old, {}, one_sided=True)], ["removed"])


if __name__ == "__main__":
    unittest.main()
