"""Empty strings in the PIVOT setters.

Where an empty caption is a stored state (item captions, data field names) it is a value and ``None`` clears or
derives; where the file holds one state for both (field names and number formats, option captions) they mean the
same; a worksheet source attribute may not be empty.
"""

from __future__ import annotations

import unittest
from typing import Callable

from formulon import (
    CellAddress,
    InvalidArgumentError,
    Pivot,
    PivotAggregation,
    PivotAxis,
    PivotCache,
    PivotDataFieldInput,
    PivotInput,
    PivotWorksheetSourceInput,
    Workbook,
    WorkbookFormat,
)

_ROWS = (("East", 10.0), ("West", 30.0), ("East", 5.0))
_OPTION_TEXTS = ("error_caption", "missing_caption", "alt_text", "alt_text_summary", "grand_total_caption")


def _build(wb: Workbook) -> "tuple[Pivot, PivotCache]":
    """A refreshed, projected pivot over Region / Amount with one data field named ``Total``."""
    sheet = wb.sheet(0)
    sheet.set_text(0, 0, "Region")
    sheet.set_text(0, 1, "Amount")
    for r, (region, amount) in enumerate(_ROWS, start=1):
        sheet.set_text(r, 0, region)
        sheet.set_number(r, 1, amount)
    cache = wb.add_pivot_cache()
    cache.set_worksheet_source(PivotWorksheetSourceInput(range="A1:B4", sheet="Sheet1"))
    for name in ("Region", "Amount"):
        cache.add_field(name)
    for item in ("East", "West"):
        cache.add_field_shared_item(0, item)
    for region, amount in _ROWS:
        record = cache.add_record()
        cache.set_record_value(record, 0, 0 if region == "East" else 1)
        cache.set_record_value(record, 1, amount)
    pivot = sheet.add_pivot(PivotInput(name="PT", cache_id=cache.cache_id, anchor=CellAddress(0, 5)))
    pivot.set_field_axis("Region", PivotAxis.ROW)
    pivot.set_field_axis("Amount", PivotAxis.VALUE)
    pivot.add_data_field(PivotDataFieldInput(name="Total", field_index="Amount", aggregation=PivotAggregation.SUM))
    cache.refresh()
    pivot.project()
    return pivot, cache


def _reloaded(wb: Workbook, fmt: WorkbookFormat, check: Callable[[Pivot], None]) -> None:
    """Save ``wb`` as ``fmt``, reload it and pass the first pivot to ``check``."""
    with Workbook.load(wb.save(format=fmt)) as back:
        check(back.sheet(0).pivot(0))


class FieldNamesAndFormatsTests(unittest.TestCase):
    def test_custom_name_empty_and_none_restore_the_source_name(self) -> None:
        for cleared in ("", None):
            with self.subTest(cleared=cleared), Workbook.create_default() as wb:
                pivot, _ = _build(wb)
                pivot.update_field("Region", custom_name="Area")
                self.assertEqual(pivot.get_field("Region").custom_name, "Area")
                pivot.update_field("Region", custom_name=cleared)
                field = pivot.get_field("Region")
                self.assertIsNone(field.custom_name)
                self.assertEqual(field.source_name, "Region")

    def test_field_number_format_empty_and_none_mean_none(self) -> None:
        for cleared in ("", None):
            with self.subTest(cleared=cleared), Workbook.create_default() as wb:
                pivot, _ = _build(wb)
                pivot.update_field("Amount", number_format="0.00")
                self.assertEqual(pivot.get_field("Amount").number_format, "0.00")
                pivot.update_field("Amount", number_format=cleared)
                self.assertIsNone(pivot.get_field("Amount").number_format)


class ItemCaptionTests(unittest.TestCase):
    def test_empty_caption_is_a_label_and_none_unrenames(self) -> None:
        with Workbook.create_default() as wb:
            pivot, _ = _build(wb)
            self.assertIsNone(pivot.get_field_item("Region", 0).caption)
            pivot.update_field_item("Region", 0, caption="")
            self.assertEqual(pivot.get_field_item("Region", 0).caption, "")
            pivot.update_field_item("Region", 0, caption=None)
            self.assertIsNone(pivot.get_field_item("Region", 0).caption)
            pivot.update_field_item("Region", 1, caption="Ouest")
            pivot.update_field_item("Region", 1, caption="")
            self.assertEqual(pivot.get_field_item("Region", 1).caption, "")

    def test_empty_label_survives_a_save_and_reload(self) -> None:
        for fmt in (WorkbookFormat.XLSX, WorkbookFormat.XLSB):
            with self.subTest(format=fmt), Workbook.create_default() as wb:
                pivot, _ = _build(wb)
                pivot.update_field_item("Region", 0, caption="")
                pivot.update_field_item("Region", 1, caption="Ouest")

                def check(back: Pivot) -> None:
                    self.assertEqual(back.get_field_item("Region", 0).caption, "")
                    self.assertEqual(back.get_field_item("Region", 1).caption, "Ouest")

                _reloaded(wb, fmt, check)
                pivot.update_field_item("Region", 0, caption=None)
                _reloaded(wb, fmt, lambda back: self.assertIsNone(back.get_field_item("Region", 0).caption))


class DataFieldTests(unittest.TestCase):
    def test_empty_name_is_kept_and_none_derives(self) -> None:
        with Workbook.create_default() as wb:
            pivot, _ = _build(wb)
            derived = pivot.get_data_field_default_name("Amount", PivotAggregation.SUM, 0)
            self.assertNotEqual(derived, "")
            pivot.set_data_field(0, PivotDataFieldInput(name="", field_index="Amount"))
            self.assertEqual(pivot.get_data_field(0).name, "")
            pivot.set_data_field(0, PivotDataFieldInput(name=None, field_index="Amount"))
            self.assertEqual(pivot.get_data_field(0).name, derived)
            added = pivot.add_data_field(
                PivotDataFieldInput(name="", field_index="Amount", aggregation=PivotAggregation.COUNT)
            )
            self.assertEqual(pivot.get_data_field(added).name, "")

    def test_empty_name_survives_a_save_and_reload(self) -> None:
        for fmt in (WorkbookFormat.XLSX, WorkbookFormat.XLSB):
            with self.subTest(format=fmt), Workbook.create_default() as wb:
                pivot, _ = _build(wb)
                pivot.set_data_field(0, PivotDataFieldInput(name="", field_index="Amount"))

                def check(back: Pivot) -> None:
                    self.assertEqual(len(back.list_data_fields()), 1)
                    self.assertEqual(back.get_data_field(0).name, "")

                _reloaded(wb, fmt, check)

    def test_number_format_empty_and_none_mean_none(self) -> None:
        for none in ("", None):
            with self.subTest(none=none), Workbook.create_default() as wb:
                pivot, _ = _build(wb)
                pivot.set_data_field(0, PivotDataFieldInput(name="Total", field_index="Amount", number_format="0.00"))
                self.assertEqual(pivot.get_data_field(0).number_format, "0.00")
                pivot.set_data_field(0, PivotDataFieldInput(name="Total", field_index="Amount", number_format=none))
                self.assertIsNone(pivot.get_data_field(0).number_format)


class OptionCaptionTests(unittest.TestCase):
    def test_empty_and_none_unset_every_caption_and_alt_text(self) -> None:
        for unset in ("", None):
            with self.subTest(unset=unset), Workbook.create_default() as wb:
                pivot, _ = _build(wb)
                pivot.update_options(**{name: "x" for name in _OPTION_TEXTS})
                for name in _OPTION_TEXTS:
                    self.assertEqual(getattr(pivot.get_options(), name), "x", name)
                pivot.update_options(**{name: unset for name in _OPTION_TEXTS})
                for name in _OPTION_TEXTS:
                    self.assertIsNone(getattr(pivot.get_options(), name), name)


class WorksheetSourceTests(unittest.TestCase):
    def test_empty_attribute_is_refused_and_none_is_absent(self) -> None:
        with Workbook.create_default() as wb:
            _, cache = _build(wb)
            for source in (
                PivotWorksheetSourceInput(range="A1:B4", sheet=""),
                PivotWorksheetSourceInput(range="A1:B4", sheet="Sheet1", name=""),
                PivotWorksheetSourceInput(name=""),
            ):
                with self.subTest(source=source), self.assertRaises(InvalidArgumentError):
                    cache.set_worksheet_source(source)
            current = cache.get_worksheet_source()
            self.assertIsNotNone(current)
            self.assertIsNone(current.name)
            cache.set_worksheet_source(PivotWorksheetSourceInput(range="A1:B4", sheet="Sheet1", name=None))
            self.assertIsNone(cache.get_worksheet_source().name)


if __name__ == "__main__":
    unittest.main()
