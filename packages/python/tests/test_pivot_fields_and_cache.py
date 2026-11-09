"""Pivot fields mirroring the cache, cache values and field-name references."""

from __future__ import annotations

import dataclasses
import datetime
import unittest

import formulon
from formulon import (
    PIVOT_VALUES_FIELD,
    CellAddress,
    InvalidArgumentError,
    NotFoundError,
    Pivot,
    PivotAggregation,
    PivotAxis,
    PivotCache,
    PivotCacheValue,
    PivotDataFieldInput,
    PivotFilterInput,
    PivotFilterKind,
    PivotInput,
    PivotWorksheetSourceInput,
    Workbook,
)

FIELDS = ("Region", "Product", "Amount")


def _cache(wb: Workbook) -> PivotCache:
    cache = wb.add_pivot_cache()
    cache.set_worksheet_source(PivotWorksheetSourceInput(range="A1:C5", sheet="Sheet1"))
    for name in FIELDS:
        cache.add_field(name)
    for item in ("East", "West"):
        cache.add_field_shared_item("Region", item)
    for values in (("East", "Apple", 10.0), ("West", "Pear", 30.0)):
        record = cache.add_record()
        for column, value in enumerate(values):
            cache.set_record_value(record, column, value)
    return cache


def _pivot(wb: Workbook, cache: PivotCache, name: str = "Pivot1", col: int = 6) -> Pivot:
    return wb.sheet(0).add_pivot(PivotInput(name=name, cache_id=cache.cache_id, anchor=CellAddress(0, col)))


class PivotFieldTests(unittest.TestCase):
    def test_fields_from_creation(self) -> None:
        with Workbook.create_default() as wb:
            cache = _cache(wb)
            pivot = _pivot(wb, cache)
            fields = pivot.list_fields()
            self.assertEqual([f.source_name for f in fields], list(FIELDS))
            self.assertEqual([f.cache_field_index for f in fields], [0, 1, 2])
            self.assertTrue(all(f.axis is None for f in fields))
            # A field added to the cache reaches every pivot on it.
            cache.add_field("Extra")
            self.assertEqual(len(pivot.list_fields()), 4)
            # Fields are placed by name on a new pivot, with no add step.
            pivot.set_field_axis("Region", PivotAxis.ROW)
            self.assertEqual(pivot.list_fields()[0].axis, PivotAxis.ROW)

    def test_pivot_has_no_add_field_and_no_field_input(self) -> None:
        self.assertFalse(hasattr(Pivot, "add_field"))
        self.assertFalse(hasattr(formulon, "PivotFieldInput"))

    def test_axis_order_names(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb, _cache(wb))
            pivot.set_field_axis("Region", PivotAxis.ROW)
            pivot.set_field_axis("Product", PivotAxis.ROW)
            pivot.set_field_axis("Amount", PivotAxis.VALUE)
            # The Values field takes a place in an order once there are two data fields.
            pivot.add_data_field(PivotDataFieldInput(field_index="Amount", name="Sum of Amount"))
            pivot.add_data_field(
                PivotDataFieldInput(field_index="Amount", name="Count of Amount", aggregation=PivotAggregation.COUNT)
            )
            pivot.set_row_field_order(["Product", "Region"])
            self.assertEqual(pivot.get_row_field_order(), [1, 0])
            pivot.set_row_field_order([PIVOT_VALUES_FIELD, "Region", "Product"])
            self.assertEqual(pivot.get_row_field_order()[0], PIVOT_VALUES_FIELD)
            with self.assertRaises(NotFoundError):
                pivot.set_row_field_order(["Nope"])
            with self.assertRaises(NotFoundError):
                pivot.set_column_field_order(["Nope"])
            with self.assertRaises(NotFoundError):
                pivot.set_page_field_order(["Nope"])

    def test_rename_empty(self) -> None:
        with Workbook.create_default() as wb:
            cache = _cache(wb)
            pivot = _pivot(wb, cache)
            with self.assertRaises(InvalidArgumentError):
                pivot.rename("")
            self.assertEqual(pivot.name, "Pivot1")
            _pivot(wb, cache, "Other", col=12)
            with self.assertRaises(formulon.AlreadyExistsError):
                pivot.rename("other")
            self.assertEqual(pivot.name, "Pivot1")


class PivotCacheTests(unittest.TestCase):
    def test_get_record_value_is_date(self) -> None:
        with Workbook.create_default() as wb:
            cache = wb.add_pivot_cache()
            cache.add_field("When")
            cache.add_field("Amount")
            record = cache.add_record()
            cache.set_record_value(record, 0, datetime.date(2024, 1, 1))
            cache.set_record_value(record, 1, 7.0)
            when = cache.get_record_value(record, 0)
            amount = cache.get_record_value(record, "Amount")
            self.assertIsInstance(when, PivotCacheValue)
            self.assertEqual((when.value, when.is_date), (45292.0, True))
            self.assertEqual((amount.value, amount.is_date), (7.0, False))
            # A read value feeds back into the setter and keeps the date.
            other = cache.add_record()
            cache.set_record_value(other, 0, when)
            self.assertEqual(cache.get_record_value(other, 0), when)
            self.assertFalse(hasattr(formulon, "PivotCacheSharedItem"))
            self.assertEqual({f.name for f in dataclasses.fields(PivotCacheValue)}, {"value", "is_date"})
            shared = cache.list_field_shared_items(0)
            self.assertTrue(all(isinstance(item, PivotCacheValue) for item in shared))

    def test_field_name_refs(self) -> None:
        with Workbook.create_default() as wb:
            cache = _cache(wb)
            self.assertEqual([i.value for i in cache.list_field_shared_items("region")], ["East", "West"])
            cache.add_field_shared_item("Region", "North")
            self.assertEqual(len(cache.list_field_shared_items(0)), 3)
            cache.clear_field_shared_items("Region")
            self.assertEqual(cache.list_field_shared_items(0), [])
            cache.set_record_value(0, "Amount", 99.0)
            self.assertEqual(cache.get_record_value(0, 2).value, 99.0)
            cache.add_calculated_field("Double", "Amount*2")
            cache.set_calculated_field_formula("Double", "Amount*3")
            # A calculated item needs its field on an axis of a pivot over the cache and the items it names.
            _pivot(wb, cache).set_field_axis("Region", PivotAxis.ROW)
            for name in ("East", "West"):
                cache.add_field_shared_item("Region", name)
            item = cache.add_calculated_item("Region", "Both", "East+West")
            self.assertGreaterEqual(item, 0)
            cache.remove_calculated_field("Double")
            for call in (
                lambda: cache.list_field_shared_items("Nope"),
                lambda: cache.add_field_shared_item("Nope", 1.0),
                lambda: cache.get_record_value(0, "Nope"),
                lambda: cache.set_calculated_field_formula("Nope", "1"),
                lambda: cache.add_calculated_item("Nope", "x", "1"),
            ):
                with self.assertRaises(NotFoundError):
                    call()

    def test_pivot_cache_method(self) -> None:
        with Workbook.create_default() as wb:
            cache = _cache(wb)
            pivot = _pivot(wb, cache)
            self.assertIs(pivot.cache(), wb.pivot_cache(cache.cache_id))
            self.assertFalse(hasattr(Pivot, "get_cache"))


class PivotFilterTests(unittest.TestCase):
    def test_kind_field(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb, _cache(wb))
            pivot.set_field_axis("Region", PivotAxis.ROW)
            index = pivot.add_filter(
                PivotFilterInput(field_index="Region", kind=PivotFilterKind.LABEL_CONTAINS, value="East")
            )
            read = pivot.get_filter(index)
            self.assertEqual(read.kind, PivotFilterKind.LABEL_CONTAINS)
            self.assertFalse(hasattr(read, "type"))
            names = {f.name for f in dataclasses.fields(PivotFilterInput)}
            self.assertIn("kind", names)
            self.assertNotIn("type", names)
            with self.assertRaises(TypeError):
                PivotFilterInput(field_index="Region", type=PivotFilterKind.LABEL_CONTAINS)  # type: ignore[call-arg]


if __name__ == "__main__":
    unittest.main()
