#!/usr/bin/env python3
"""Unit tests for the DataBar / IconSet documented-semantics computation."""

from __future__ import annotations

import unittest

from tools.oracle import cf_visual_semantics as sem

SQREF = [{"first_row": 1, "first_col": 1, "last_row": 5, "last_col": 1}]


def _case(values: list[object]) -> dict:
    """A one-column sheet; `None` leaves the row absent, strings become text cells."""

    sheet = []
    for i, v in enumerate(values, start=1):
        if v is None:
            continue
        if isinstance(v, str):
            sheet.append({"row": i, "col": 1, "kind": "text", "value": v})
        elif isinstance(v, bool):
            sheet.append({"row": i, "col": 1, "kind": "bool", "value": v})
        else:
            sheet.append({"row": i, "col": 1, "kind": "number", "value": v})
    return {
        "sheet": sheet,
        "range": {"first_row": 1, "first_col": 1, "last_row": len(values), "last_col": 1},
    }


def _data_bar(min_type="Min", min_value=None, max_type="Max", max_value=None, **extra) -> dict:
    return {
        "type": "DataBar",
        "priority": extra.pop("priority", 1),
        "sqref": extra.pop("sqref", SQREF),
        "data_bar": {
            "min_type": min_type,
            "min_value": min_value,
            "max_type": max_type,
            "max_value": max_value,
            "min_length_pct": extra.pop("min_length_pct", 10),
            "max_length_pct": extra.pop("max_length_pct", 90),
            "fill": extra.pop("fill", {"r": 1, "g": 2, "b": 3}),
        },
    }


def _icon_set(thresholds, name_ooxml="3Arrows", reverse=False, priority=1) -> dict:
    return {
        "type": "IconSet",
        "priority": priority,
        "sqref": SQREF,
        "icon_set": {
            "name_ooxml": name_ooxml,
            "name_formulon": "Three_Arrows",
            "reverse": reverse,
            "thresholds": [{"type": t, "value": v, "gte": g} for t, v, g in thresholds],
        },
    }


class CellInBlockTest(unittest.TestCase):
    def test_bounds_are_inclusive(self):
        self.assertTrue(sem._cell_in_block(1, 1, SQREF))
        self.assertTrue(sem._cell_in_block(5, 1, SQREF))
        self.assertFalse(sem._cell_in_block(6, 1, SQREF))
        self.assertFalse(sem._cell_in_block(1, 2, SQREF))

    def test_any_block_matches(self):
        blocks = SQREF + [{"first_row": 10, "first_col": 3, "last_row": 10, "last_col": 3}]
        self.assertTrue(sem._cell_in_block(10, 3, blocks))
        self.assertFalse(sem._cell_in_block(10, 1, blocks))


class PopulationTest(unittest.TestCase):
    def test_sorted_numbers_only(self):
        case = _case([5, "x", 1.5, True, 3])
        self.assertEqual(sem._numeric_population(case, SQREF), [1.5, 3.0, 5.0])

    def test_cells_outside_block_ignored(self):
        case = _case([1, 2, 3, 4, 5])
        block = [{"first_row": 2, "first_col": 1, "last_row": 3, "last_col": 1}]
        self.assertEqual(sem._numeric_population(case, block), [2.0, 3.0])

    def test_cell_value_for(self):
        case = _case([7, "x", None])
        self.assertEqual(sem._cell_value_for(case, 1, 1), 7.0)
        self.assertIsNone(sem._cell_value_for(case, 2, 1))
        self.assertIsNone(sem._cell_value_for(case, 3, 1))
        self.assertIsNone(sem._cell_value_for(case, 9, 9))


class ThresholdTest(unittest.TestCase):
    POP = [0.0, 10.0, 20.0, 40.0]

    def test_number_ignores_population(self):
        self.assertEqual(sem._resolve_threshold_value("Number", 7, []), 7.0)
        self.assertIsNone(sem._resolve_threshold_value("Number", None, self.POP))

    def test_empty_population_needs_none(self):
        for kind in ("Min", "Max", "Percent", "Percentile"):
            self.assertIsNone(sem._resolve_threshold_value(kind, 50, []))

    def test_min_max(self):
        self.assertEqual(sem._resolve_threshold_value("Min", None, self.POP), 0.0)
        self.assertEqual(sem._resolve_threshold_value("Max", None, self.POP), 40.0)

    def test_percent_is_linear_between_extremes(self):
        self.assertEqual(sem._resolve_threshold_value("Percent", 25, self.POP), 10.0)
        self.assertIsNone(sem._resolve_threshold_value("Percent", None, self.POP))

    def test_percentile_interpolates_between_ranks(self):
        # position = 0.5 * 3 = 1.5 -> halfway between 10 and 20
        self.assertEqual(sem._resolve_threshold_value("Percentile", 50, self.POP), 15.0)
        self.assertEqual(sem._resolve_threshold_value("Percentile", 100, self.POP), 40.0)
        self.assertEqual(sem._resolve_threshold_value("Percentile", 0, self.POP), 0.0)
        self.assertEqual(sem._resolve_threshold_value("Percentile", 50, [3.0]), 3.0)

    def test_unknown_type_is_none(self):
        self.assertIsNone(sem._resolve_threshold_value("Formula", "A1", self.POP))


class DataBarTest(unittest.TestCase):
    POP = [0.0, 50.0, 100.0]

    def test_length_scales_into_configured_range(self):
        match = sem._data_bar_match_for_cell(50.0, _data_bar(), self.POP)
        self.assertEqual(match["kind"], "DataBar")
        self.assertEqual(match["bar_length_pct"], 50.0)
        self.assertEqual(match["bar_axis_position_pct"], 0.0)
        self.assertFalse(match["is_negative"])
        self.assertEqual(match["fill"], {"r": 1, "g": 2, "b": 3, "a": 255})

    def test_length_clamps_to_unit_fraction(self):
        desc = _data_bar(min_type="Number", min_value=20, max_type="Number", max_value=80)
        self.assertEqual(sem._data_bar_match_for_cell(0.0, desc, self.POP)["bar_length_pct"], 10.0)
        self.assertEqual(sem._data_bar_match_for_cell(100.0, desc, self.POP)["bar_length_pct"], 90.0)

    def test_explicit_alpha_kept(self):
        desc = _data_bar(fill={"r": 1, "g": 2, "b": 3, "a": 40})
        self.assertEqual(sem._data_bar_match_for_cell(0.0, desc, self.POP)["fill"]["a"], 40)

    def test_no_match_for_blank_degenerate_or_empty(self):
        self.assertIsNone(sem._data_bar_match_for_cell(None, _data_bar(), self.POP))
        self.assertIsNone(sem._data_bar_match_for_cell(1.0, _data_bar(), [4.0]))
        self.assertIsNone(sem._data_bar_match_for_cell(1.0, _data_bar(), []))


class IconSetTest(unittest.TestCase):
    POP = [0.0, 50.0, 100.0]
    THRESHOLDS = [("Percent", 33, True), ("Percent", 67, True)]

    def _index(self, value, **kw):
        match = sem._icon_set_match_for_cell(value, _icon_set(self.THRESHOLDS, **kw), self.POP)
        return match["icon_index"]

    def test_buckets_walk_thresholds(self):
        self.assertEqual(self._index(0.0), 0)
        self.assertEqual(self._index(40.0), 1)
        self.assertEqual(self._index(100.0), 2)

    def test_gte_flag_controls_boundary(self):
        strict = [("Number", 50, False)]
        inclusive = [("Number", 50, True)]
        at = lambda th: sem._icon_set_match_for_cell(50.0, _icon_set(th), self.POP)["icon_index"]  # noqa: E731
        self.assertEqual(at(strict), 0)
        self.assertEqual(at(inclusive), 1)

    def test_reverse_mirrors_index_across_buckets(self):
        self.assertEqual(self._index(100.0, reverse=True), 0)
        self.assertEqual(self._index(0.0, reverse=True), 2)

    def test_bucket_count_follows_set_name(self):
        self.assertEqual(sem._ICON_SET_BUCKET_COUNT["3Arrows"], 3)
        self.assertEqual(sem._ICON_SET_BUCKET_COUNT["5Quarters"], 5)
        thresholds = [("Number", 0, True)] * 4
        match = sem._icon_set_match_for_cell(1.0, _icon_set(thresholds, name_ooxml="5Rating", reverse=True), self.POP)
        self.assertEqual(match["icon_index"], 0)

    def test_payload_shape_and_absent_cases(self):
        match = sem._icon_set_match_for_cell(100.0, _icon_set(self.THRESHOLDS, priority=4), self.POP)
        self.assertEqual(match["kind"], "IconSet")
        self.assertEqual(match["priority"], 4)
        self.assertEqual(match["icon_set_name"], "Three_Arrows")
        self.assertIsNone(sem._icon_set_match_for_cell(None, _icon_set(self.THRESHOLDS), self.POP))
        self.assertIsNone(sem._icon_set_match_for_cell(1.0, _icon_set(self.THRESHOLDS), []))

    def test_name_tables_are_inverse(self):
        for ooxml, formulon in sem._ICON_SET_OOXML_TO_FORMULON.items():
            self.assertEqual(sem._ICON_SET_FORMULON_TO_OOXML[formulon], ooxml)


class ComputeVisualMatchesTest(unittest.TestCase):
    def test_every_numeric_cell_in_range_gets_a_match(self):
        case = _case([0, 50, 100, "x", None])
        out = sem._compute_visual_matches(case, [_data_bar()])
        self.assertEqual(sorted(out), [(1, 1), (2, 1), (3, 1)])
        self.assertEqual([out[(r, 1)][0]["bar_length_pct"] for r in (1, 2, 3)], [10.0, 50.0, 90.0])

    def test_priority_order_and_rule_type_filter(self):
        case = _case([0, 50, 100])
        thresholds = [("Percent", 33, True), ("Percent", 67, True)]
        descriptors = [
            _icon_set(thresholds, priority=2),
            {"type": "CellIs", "priority": 0, "sqref": SQREF},
            _data_bar(priority=1),
        ]
        out = sem._compute_visual_matches(case, descriptors)
        self.assertEqual([m["kind"] for m in out[(1, 1)]], ["DataBar", "IconSet"])

    def test_sqref_limits_cells_and_population(self):
        case = _case([0, 50, 100])
        block = [{"first_row": 2, "first_col": 1, "last_row": 3, "last_col": 1}]
        out = sem._compute_visual_matches(case, [_data_bar(sqref=block)])
        self.assertEqual(sorted(out), [(2, 1), (3, 1)])
        self.assertEqual(out[(2, 1)][0]["bar_length_pct"], 10.0)

    def test_degenerate_population_yields_no_matches(self):
        self.assertEqual(sem._compute_visual_matches(_case([5, "x"]), [_data_bar()]), {})


if __name__ == "__main__":
    unittest.main()
