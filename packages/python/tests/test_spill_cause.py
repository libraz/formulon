"""Spill.cause: what blocks a formula's spill or a PivotTable report, on Sheet.get_spill and Pivot.get_spill."""

from __future__ import annotations

import enum
import unittest
from typing import Callable, Tuple

import formulon
from formulon import (
    CellAddress,
    CellRect,
    Pivot,
    PivotAggregation,
    PivotAxis,
    PivotDataFieldInput,
    PivotInput,
    PivotWorksheetSourceInput,
    Spill,
    SpillCause,
    Workbook,
)

LAST_ROW = 1048575
LAST_COL = 16383


def _build_pivot(wb: Workbook) -> Tuple[Pivot, Callable[[], None]]:
    """A pivot summing Amount by Region at E1, drawn over E1:F4; the callable makes it one row longer, into E5."""
    ws = wb.sheet(0)
    ws.set_text("A1", "Region")
    ws.set_text("B1", "Amount")
    rows = (("East", 10.0), ("West", 30.0), ("East", 5.0))
    for i, (region, amount) in enumerate(rows):
        ws.set_text(i + 1, 0, region)
        ws.set_number(i + 1, 1, amount)
    cache = wb.add_pivot_cache()
    cache.set_worksheet_source(PivotWorksheetSourceInput(range="A1:B4", sheet="Sheet1"))
    for name in ("Region", "Amount"):
        cache.add_field(name)
    for item in ("East", "West"):
        cache.add_field_shared_item("Region", item)
    for region, amount in rows:
        record = cache.add_record()
        cache.set_record_value(record, 0, 0 if region == "East" else 1)
        cache.set_record_value(record, 1, amount)
    pivot = ws.add_pivot(PivotInput(name="Pivot1", cache_id=cache.cache_id, anchor=CellAddress(0, 4)))
    pivot.set_field_axis("Region", PivotAxis.ROW)
    pivot.add_data_field(PivotDataFieldInput(field_index="Amount", aggregation=PivotAggregation.SUM))
    pivot.project()

    def grow() -> None:
        ws.set_text("A4", "North")
        cache.refresh()

    return pivot, grow


class FormulaSpillCauseTests(unittest.TestCase):
    def test_a_drawn_spill_reads_no_cause_at_its_anchor_and_phantoms(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_formula("A1", "=SEQUENCE(3)")
            drawn = Spill(anchor=CellAddress(0, 0), blocked=False, cause=None, range=CellRect(0, 0, 2, 0))
            self.assertEqual(ws.get_spill("A1"), drawn)
            self.assertEqual(ws.get_spill("A3"), drawn)

    def test_a_value_is_occupied_and_a_merge_is_merged(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_formula("C1", "=SEQUENCE(3)")
            ws.set_text("C3", "x")
            ws.add_merge("D2:E2")
            ws.set_formula("D1", "=SEQUENCE(3)")
            self.assertEqual(
                ws.get_spill("C1"),
                Spill(anchor=CellAddress(0, 2), blocked=True, cause=SpillCause.OCCUPIED, range=CellRect(0, 2, 2, 2)),
            )
            self.assertEqual(
                ws.get_spill("D1"),
                Spill(anchor=CellAddress(0, 3), blocked=True, cause=SpillCause.MERGED, range=CellRect(0, 3, 2, 3)),
            )
            self.assertIsNone(ws.get_spill("C2"))

    def test_past_the_edge_is_beyond_edge_clipped_to_the_grid(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_formula(LAST_ROW - 1, 0, "=SEQUENCE(5)")
            ws.set_formula(0, LAST_COL - 1, "=SEQUENCE(1,5)")
            self.assertEqual(
                ws.get_spill(LAST_ROW - 1, 0),
                Spill(
                    anchor=CellAddress(LAST_ROW - 1, 0),
                    blocked=True,
                    cause=SpillCause.BEYOND_EDGE,
                    range=CellRect(LAST_ROW - 1, 0, LAST_ROW, 0),
                ),
            )
            self.assertEqual(
                ws.get_spill(0, LAST_COL - 1),
                Spill(
                    anchor=CellAddress(0, LAST_COL - 1),
                    blocked=True,
                    cause=SpillCause.BEYOND_EDGE,
                    range=CellRect(0, LAST_COL - 1, 0, LAST_COL),
                ),
            )

    def test_a_cell_in_no_spill_reads_none(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number("Z1", 1.0)
            self.assertIsNone(wb.sheet(0).get_spill("Z1"))


class PivotSpillCauseTests(unittest.TestCase):
    def test_a_drawn_report_reads_its_extent_at_its_anchor_only(self) -> None:
        with Workbook.create_default() as wb:
            pivot, _ = _build_pivot(wb)
            ws = wb.sheet(0)
            drawn = Spill(anchor=CellAddress(0, 4), blocked=False, cause=None, range=CellRect(0, 4, 3, 5))
            self.assertEqual(pivot.get_spill(), drawn)
            self.assertEqual(ws.get_spill("E1"), drawn)
            self.assertIsNone(ws.get_spill("E2"))
            self.assertIsNone(ws.get_spill("F4"))

    def test_a_value_in_the_way_is_occupied(self) -> None:
        with Workbook.create_default() as wb:
            pivot, grow = _build_pivot(wb)
            wb.sheet(0).set_text("E5", "x")
            grow()
            blocked = Spill(
                anchor=CellAddress(0, 4), blocked=True, cause=SpillCause.OCCUPIED, range=CellRect(0, 4, 4, 5)
            )
            self.assertEqual(pivot.get_spill(), blocked)
            self.assertEqual(wb.sheet(0).get_spill("E1"), blocked)

    def test_a_merge_in_the_way_is_merged(self) -> None:
        with Workbook.create_default() as wb:
            pivot, grow = _build_pivot(wb)
            wb.sheet(0).add_merge("E5:F5")
            grow()
            self.assertEqual(
                pivot.get_spill(),
                Spill(anchor=CellAddress(0, 4), blocked=True, cause=SpillCause.MERGED, range=CellRect(0, 4, 4, 5)),
            )

    def test_past_the_edge_is_beyond_edge_clipped_to_the_grid(self) -> None:
        with Workbook.create_default() as wb:
            pivot, grow = _build_pivot(wb)
            pivot.set_anchor(CellAddress(LAST_ROW - 3, 4))
            grow()
            blocked = Spill(
                anchor=CellAddress(LAST_ROW - 3, 4),
                blocked=True,
                cause=SpillCause.BEYOND_EDGE,
                range=CellRect(LAST_ROW - 3, 4, LAST_ROW, 5),
            )
            self.assertEqual(pivot.get_spill(), blocked)
            self.assertEqual(wb.sheet(0).get_spill(LAST_ROW - 3, 4), blocked)

    def test_a_drawn_report_read_from_a_file_reads_its_location(self) -> None:
        with Workbook.create_default() as wb:
            _build_pivot(wb)
            with Workbook.load(wb.save()) as loaded:
                drawn = Spill(anchor=CellAddress(0, 4), blocked=False, cause=None, range=CellRect(0, 4, 3, 5))
                self.assertEqual(loaded.sheet(0).pivot(0).get_spill(), drawn)
                self.assertEqual(loaded.sheet(0).get_spill("E1"), drawn)
                self.assertIsNone(loaded.sheet(0).get_spill("E2"))


class SpillCauseEnumTests(unittest.TestCase):
    def test_members_and_no_member_for_absence(self) -> None:
        self.assertTrue(issubclass(SpillCause, enum.IntEnum))
        self.assertEqual({m.name: m.value for m in SpillCause}, {"OCCUPIED": 0, "MERGED": 1, "BEYOND_EDGE": 2})
        self.assertIs(formulon.SpillCause, SpillCause)


if __name__ == "__main__":
    unittest.main()
