"""PivotCache.refresh rebuilds a cache from its worksheet source; PivotCacheField states the sharedItems flags."""

from __future__ import annotations

import dataclasses
import unittest
from typing import List, Tuple

from formulon import (
    CellAddress,
    CellRect,
    FormulonError,
    NotFoundError,
    Pivot,
    PivotAggregation,
    PivotAxis,
    PivotCache,
    PivotCacheField,
    PivotDataFieldInput,
    PivotInput,
    PivotNumericGroup,
    PivotNumericGroupInput,
    PivotWorksheetSourceInput,
    Spill,
    StatusCode,
    Workbook,
)

ROWS: List[Tuple[str, float]] = [("East", 10.0), ("West", 30.0), ("East", 5.0)]


def _build(wb: Workbook, rows: List[Tuple[str, float]] = ROWS) -> Tuple[PivotCache, Pivot]:
    """Source cells Region / Amount, a cache over them with records, and a pivot summing Amount by Region at E1."""
    ws = wb.sheet(0)
    ws.set_text("A1", "Region")
    ws.set_text("B1", "Amount")
    for i, (region, amount) in enumerate(rows):
        ws.set_text(i + 1, 0, region)
        ws.set_number(i + 1, 1, amount)
    cache = wb.add_pivot_cache()
    cache.set_worksheet_source(PivotWorksheetSourceInput(range=f"A1:B{len(rows) + 1}", sheet="Sheet1"))
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
    return cache, pivot


def _shared(cache: PivotCache, field: str) -> List[object]:
    return [item.value for item in cache.list_field_shared_items(field)]


class PivotCacheRefreshTests(unittest.TestCase):
    def test_re_reads_the_source_into_the_cache_and_the_report(self) -> None:
        with Workbook.create_default() as wb:
            cache, _ = _build(wb)
            ws = wb.sheet(0)
            self.assertEqual(_shared(cache, "Region"), ["East", "West"])
            ws.set_text("A4", "North")
            ws.set_number("B2", 1000.0)
            self.assertIsNone(cache.refresh())
            self.assertEqual(_shared(cache, "Region"), ["East", "West", "North"])
            self.assertEqual(cache.get_record_value(0, "Amount").value, 1000.0)
            shown = [ws.get_value(row, col) for row in range(12) for col in range(4, 7)]
            self.assertIn("North", shown)

    def test_unresolved_source_fails_and_changes_nothing(self) -> None:
        with Workbook.create_default() as wb:
            cache, _ = _build(wb)
            cache.set_worksheet_source(PivotWorksheetSourceInput(range="A1:B4", sheet="Missing"))
            wb.sheet(0).set_text("A4", "North")
            with self.assertRaises(FormulonError) as raised:
                cache.refresh()
            self.assertEqual(raised.exception.code, StatusCode.PRECONDITION_FAILED)
            self.assertEqual(_shared(cache, "Region"), ["East", "West"])
            cache.set_worksheet_source(None)
            with self.assertRaises(FormulonError) as raised:
                cache.refresh()
            self.assertEqual(raised.exception.code, StatusCode.PRECONDITION_FAILED)

    def test_report_that_cannot_grow_is_blocked_until_its_range_is_free(self) -> None:
        with Workbook.create_default() as wb:
            cache, pivot = _build(wb)
            ws = wb.sheet(0)
            self.assertIsNone(pivot.get_spill())
            # The report fills E1:F4; North makes it one row longer, into E5.
            for row in range(4, 9):
                ws.set_text(row, 4, "keep")
            ws.set_text("A4", "North")
            self.assertIsNone(cache.refresh())
            self.assertEqual(_shared(cache, "Region"), ["East", "West", "North"])
            self.assertFalse(pivot.is_projection_pending())
            needed = Spill(blocked=True, range=CellRect(0, 4, 4, 5))
            self.assertEqual(pivot.get_spill(), needed)
            self.assertEqual(ws.get_spill("E1"), needed)
            self.assertIsNone(ws.get_spill("F2"))
            self.assertEqual(ws.get_value(4, 4), "keep")
            for row in range(4, 9):
                ws.set_blank(row, 4)
            self.assertIsNone(pivot.get_spill())
            self.assertIsNone(ws.get_spill("E1"))
            self.assertIn("North", [ws.get_value(row, 4) for row in range(6)])

    def test_get_spill_of_a_removed_pivot_is_not_found(self) -> None:
        with Workbook.create_default() as wb:
            _, pivot = _build(wb)
            wb.sheet(0).remove_pivot(pivot)
            with self.assertRaises(NotFoundError):
                pivot.get_spill()

    def test_source_over_the_record_cell_limit_fails(self) -> None:
        with Workbook.create_default() as wb:
            cache, _ = _build(wb)
            cache.set_worksheet_source(PivotWorksheetSourceInput(range="A1:AG1048576", sheet="Sheet1"))
            with self.assertRaises(FormulonError) as raised:
                cache.refresh()
            self.assertEqual(raised.exception.code, StatusCode.SECURITY_RESOURCE_LIMIT)

    def test_removed_cache_is_not_found(self) -> None:
        with Workbook.create_default() as wb:
            cache = wb.add_pivot_cache()
            wb.remove_pivot_cache(cache)
            with self.assertRaises(NotFoundError):
                cache.refresh()


class PivotCacheFieldTypeTests(unittest.TestCase):
    def test_text_and_number_fields(self) -> None:
        with Workbook.create_default() as wb:
            cache, _ = _build(wb)
            region, amount = cache.list_fields()
            self.assertTrue(region.contains_string)
            self.assertFalse(region.contains_number)
            self.assertTrue(region.contains_semi_mixed_types)
            self.assertTrue(region.contains_non_date)
            self.assertTrue(amount.contains_number)
            self.assertTrue(amount.contains_integer)
            self.assertFalse(amount.contains_string)
            self.assertFalse(amount.contains_blank)
            self.assertFalse(amount.contains_date)
            self.assertFalse(amount.contains_mixed_types)

    def test_flags_follow_a_refresh(self) -> None:
        with Workbook.create_default() as wb:
            cache, _ = _build(wb)
            ws = wb.sheet(0)
            ws.set_number("B2", 1.5)
            ws.set_blank("B3")
            cache.refresh()
            amount = cache.list_fields()[1]
            self.assertTrue(amount.contains_number)
            self.assertFalse(amount.contains_integer)
            self.assertTrue(amount.contains_blank)
            self.assertTrue(amount.contains_semi_mixed_types)

    def test_flags_are_read_only_and_default_false(self) -> None:
        bare = PivotCacheField(name="X", formula=None)
        for field in dataclasses.fields(PivotCacheField):
            if field.name.startswith("contains_"):
                self.assertIs(getattr(bare, field.name), False)
        with self.assertRaises(dataclasses.FrozenInstanceError):
            bare.contains_number = True  # type: ignore[misc]


NUMBERS: List[Tuple[str, float]] = [("East", 20.0), ("West", 150.0), ("East", 250.0), ("West", 300.0)]


class PivotNumericGroupTests(unittest.TestCase):
    def test_groups_reads_back_and_removes(self) -> None:
        with Workbook.create_default() as wb:
            _, pivot = _build(wb, NUMBERS)
            self.assertIsNone(pivot.get_field_numeric_group("Amount"))
            self.assertIsNone(pivot.get_field_date_group("Amount"))

            pivot.set_field_numeric_group("Amount", PivotNumericGroupInput(interval=100))
            self.assertEqual(
                pivot.get_field_numeric_group("Amount"),
                PivotNumericGroup(interval=100.0, start=20.0, end=300.0, auto_start=True, auto_end=True),
            )
            self.assertEqual(len(pivot.list_fields()), 3)
            # The group field's items: <20, 20-119, 120-219, 220-319, >320.
            pivot.set_field_axis(2, PivotAxis.ROW)
            pivot.project()
            self.assertEqual(len(pivot.list_field_items(2)), 5)
            shown = [wb.sheet(0).get_value(row, col) for row in range(12) for col in range(4, 7)]
            for label in ("20-119", "120-219", "220-319"):
                self.assertIn(label, shown)
            self.assertEqual(pivot.get_field_numeric_group(2), pivot.get_field_numeric_group("Amount"))

            pivot.set_field_numeric_group("Amount", PivotNumericGroupInput(interval=100, start=0, end=400))
            self.assertEqual(
                pivot.get_field_numeric_group("Amount"),
                PivotNumericGroup(interval=100.0, start=0.0, end=400.0, auto_start=False, auto_end=False),
            )

            pivot.set_field_numeric_group("Amount", None)
            self.assertIsNone(pivot.get_field_numeric_group("Amount"))
            self.assertEqual(len(pivot.list_fields()), 2)

    def test_non_integer_interval_is_not_implemented(self) -> None:
        with Workbook.create_default() as wb:
            _, pivot = _build(wb, NUMBERS)
            with self.assertRaises(FormulonError) as raised:
                pivot.set_field_numeric_group("Amount", PivotNumericGroupInput(interval=1.5))
            self.assertEqual(raised.exception.code, StatusCode.NOT_IMPLEMENTED)
            with self.assertRaises(FormulonError) as raised:
                pivot.set_field_numeric_group("Amount", PivotNumericGroupInput(interval=0))
            self.assertEqual(raised.exception.code, StatusCode.INVALID_ARGUMENT)
            self.assertIsNone(pivot.get_field_numeric_group("Amount"))
            self.assertEqual(len(pivot.list_fields()), 2)

    def test_rejects_malformed_groups(self) -> None:
        with Workbook.create_default() as wb:
            _, pivot = _build(wb, NUMBERS)
            with self.assertRaises(TypeError):
                pivot.set_field_numeric_group("Amount", 5)  # type: ignore[arg-type]
            with self.assertRaises(TypeError):
                pivot.set_field_numeric_group("Amount", PivotNumericGroupInput(interval="x"))  # type: ignore[arg-type]


if __name__ == "__main__":
    unittest.main()
