"""Pivot surface and marshalling tests."""

from __future__ import annotations

import datetime
import inspect
import io
import unittest
import zipfile
from dataclasses import replace
from typing import Any, Dict, List, Tuple

from formulon import (
    PIVOT_VALUES_FIELD,
    CellAddress,
    CellRect,
    ErrorCode,
    ExcelError,
    FormulonError,
    IndexOutOfRangeError,
    InvalidArgumentError,
    NotFoundError,
    Pivot,
    PivotAggregation,
    PivotAxis,
    PivotCache,
    PivotCalendar,
    PivotDataFieldInput,
    PivotDateGroupInput,
    PivotDatePeriod,
    PivotFieldInput,
    PivotFilterInput,
    PivotFilterKind,
    PivotGrandTotals,
    PivotInput,
    PivotReportLayout,
    PivotSortOrder,
    PivotWorksheetSourceInput,
    Sheet,
    SheetRect,
    StatusCode,
    Workbook,
)


class PivotTests(unittest.TestCase):
    def test_build_and_project_pivot(self) -> None:
        with Workbook.create_default() as wb:
            cache_id = wb.add_pivot_cache().cache_id
            wb.pivot_cache(cache_id).add_field("Region")
            wb.pivot_cache(cache_id).add_field("Amount")
            for region, amount in [("East", 10.0), ("East", 20.0), ("West", 30.0)]:
                rec = wb.pivot_cache(cache_id).add_record()
                wb.pivot_cache(cache_id).set_record_value(rec, 0, region)
                wb.pivot_cache(cache_id).set_record_value(rec, 1, amount)
            self.assertEqual(wb.pivot_cache(cache_id).get_record_count(), 3)

            pivot = wb.sheet(0).add_pivot(PivotInput(name="Pivot1", cache_id=cache_id, anchor=CellAddress(0, 4)))
            region_field = wb.sheet(0).pivot(pivot).add_field(PivotFieldInput(field_index="Region", axis=PivotAxis.ROW))
            amount_field = (
                wb.sheet(0).pivot(pivot).add_field(PivotFieldInput(field_index="Amount", axis=PivotAxis.VALUE))
            )
            self.assertEqual(region_field, 0)
            wb.sheet(0).pivot(pivot).add_data_field(
                PivotDataFieldInput(
                    name="Sum of Amount",
                    field_index=amount_field,
                    aggregation=PivotAggregation.SUM,
                )
            )

            layout = wb.sheet(0).pivot(pivot).get_layout()
            self.assertGreater(len(layout.cells), 0)
            numbers = [c.value for c in layout.cells if isinstance(c.value, float)]
            # The single grand total of 10 + 20 + 30 must appear.
            self.assertIn(60.0, numbers)

    def test_cache_shared_item_value_reads_each_kind(self) -> None:
        with Workbook.create_default() as wb:
            cache_id = wb.add_pivot_cache().cache_id
            wb.pivot_cache(cache_id).add_field("Region")
            wb.pivot_cache(cache_id).add_field_shared_item(0, "East")
            wb.pivot_cache(cache_id).add_field_shared_item(0, 45292.0)
            wb.pivot_cache(cache_id).add_field_shared_item(0, True)
            wb.pivot_cache(cache_id).add_field_shared_item(0, None)
            wb.pivot_cache(cache_id).add_field_shared_item(0, ExcelError(ErrorCode(1)))
            items = wb.pivot_cache(cache_id).list_field_shared_items(0)
            self.assertEqual(
                [item.value for item in items],
                ["East", 45292.0, True, None, ExcelError(ErrorCode(1))],
            )
            self.assertIs(type(items[1].value), float)
            self.assertIs(items[2].value, True)
            self.assertFalse(any(item.is_date for item in items))
            with self.assertRaises(FormulonError):
                wb.pivot_cache(cache_id).list_field_shared_items(1)

    def test_cache_date_shared_items(self) -> None:
        with Workbook.create_default() as wb:
            cache_id = wb.add_pivot_cache().cache_id
            wb.pivot_cache(cache_id).add_field("When")
            wb.pivot_cache(cache_id).add_field_shared_item(0, datetime.date(2024, 1, 1))
            wb.pivot_cache(cache_id).add_field_shared_item(0, 45292.0)
            wb.pivot_cache(cache_id).add_field_shared_item(0, datetime.datetime(2024, 1, 1, 12, 0, 0))
            items = wb.pivot_cache(cache_id).list_field_shared_items(0)
            self.assertEqual([item.is_date for item in items], [True, False, True])
            self.assertEqual([item.value for item in items], [45292.0, 45292.0, 45292.5])
            with self.assertRaises(FormulonError) as caught:
                wb.pivot_cache(cache_id).add_field_shared_item(0, datetime.date(1800, 1, 1))
            self.assertEqual(caught.exception.code, 2)
            wb.pivot_cache(cache_id).add_record()
            wb.pivot_cache(cache_id).set_record_value(0, 0, datetime.date(2024, 1, 1))
            with self.assertRaises(FormulonError):
                wb.pivot_cache(cache_id).set_record_value(0, 0, datetime.date(1800, 1, 1))

    def test_cache_source_and_report_layout_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            cache = wb.add_pivot_cache().cache_id
            wb.pivot_cache(cache).set_worksheet_source(PivotWorksheetSourceInput(range="A1:B9", sheet="Sheet1"))
            self.assertEqual(wb.pivot_cache(cache).get_worksheet_source().range, SheetRect(0, CellRect(0, 0, 8, 1)))
            pivot = wb.sheet(0).add_pivot(PivotInput(name="Layout", cache_id=cache, anchor=CellAddress(0, 0)))
            wb.sheet(0).pivot(pivot).set_report_layout(PivotReportLayout.TABULAR)
            self.assertEqual(wb.sheet(0).pivot(pivot).get_report_layout(), PivotReportLayout.TABULAR)
            wb.pivot_cache(cache).set_worksheet_source(None)
            self.assertIsNone(wb.pivot_cache(cache).get_worksheet_source())


class _PivotFixture:
    """Region / Product carry shared items their records index; Amount and
    Date are plain numbers. Only Region is laid out."""

    def __init__(self, wb: Workbook) -> None:
        self.wb = wb
        wb.set_profile_id("win-365-ja_JP")
        self.cache = wb.add_pivot_cache().cache_id
        wb.pivot_cache(self.cache).set_worksheet_source(PivotWorksheetSourceInput(range="A1:D5", sheet="Sheet1"))
        for name in ("Region", "Product", "Amount", "Date"):
            wb.pivot_cache(self.cache).add_field(name)
        for item in ("East", "West"):
            wb.pivot_cache(self.cache).add_field_shared_item(0, item)
        for item in ("Apple", "Pear"):
            wb.pivot_cache(self.cache).add_field_shared_item(1, item)
        for values in ((0, 0, 10, 45292), (1, 0, 30, 45323), (0, 1, 5, 45658), (1, 1, 7, 45689)):
            rec = wb.pivot_cache(self.cache).add_record()
            for field, value in enumerate(values):
                wb.pivot_cache(self.cache).set_record_value(rec, field, value)
        self.pivot = wb.sheet(0).add_pivot(PivotInput(name="Pivot1", cache_id=self.cache, anchor=CellAddress(0, 6)))
        self.region = wb.sheet(0).pivot(self.pivot).add_field(PivotFieldInput(field_index="Region", axis=PivotAxis.ROW))
        wb.sheet(0).pivot(self.pivot).add_field_item(self.region, "East", visible=True)
        wb.sheet(0).pivot(self.pivot).add_field_item(self.region, "West", visible=True)
        self.product = (
            wb.sheet(0).pivot(self.pivot).add_field(PivotFieldInput(field_index="Product", axis=PivotAxis.VALUE))
        )
        wb.sheet(0).pivot(self.pivot).add_field_item(self.product, "Apple", visible=True)
        wb.sheet(0).pivot(self.pivot).add_field_item(self.product, "Pear", visible=True)
        self.amount = (
            wb.sheet(0).pivot(self.pivot).add_field(PivotFieldInput(field_index="Amount", axis=PivotAxis.VALUE))
        )
        self.date = wb.sheet(0).pivot(self.pivot).add_field(PivotFieldInput(field_index="Date", axis=PivotAxis.VALUE))
        wb.sheet(0).pivot(self.pivot).set_row_field_order([self.region])
        wb.sheet(0).pivot(self.pivot).add_data_field(
            PivotDataFieldInput(name="Sum of Amount", field_index=self.amount, aggregation=0)
        )

    def grid(self, pivot: int = 0) -> List[str]:
        """The projected grid as one ``|``-joined string per sheet row."""
        return _grid(self.wb, pivot)


def _grid(wb: Workbook, pivot: int = 0) -> List[str]:
    """The projected grid of ``pivot`` on sheet 0, one ``|``-joined string per row."""
    layout = wb.sheet(0).pivot(pivot).get_layout()
    rows: Dict[int, List[str]] = {}
    for cell in layout.cells:
        v = cell.value
        if isinstance(v, bool):
            text = ""
        elif isinstance(v, float):
            text = "%.15g" % v
        elif isinstance(v, str):
            text = v
        elif isinstance(v, ExcelError):
            text = f"#{int(v.code)}"
        else:
            text = ""
        assert layout.range is not None
        rows.setdefault(cell.row, [""] * (layout.range.last_col - layout.range.first_col + 1))[
            cell.col - layout.range.first_col
        ] = text
    return ["|".join(rows[r]) for r in sorted(rows)]


_BASE_GRID = ["行ラベル|Sum of Amount", "East|15", "West|37", "総計|52"]


_TWO_LEVEL_GRID = [
    "行ラベル|Sum of Amount",
    "East|15",
    "Apple|10",
    "Pear|5",
    "West|37",
    "Apple|30",
    "Pear|7",
    "総計|52",
]


_WEST_ONLY = ["行ラベル|Sum of Amount", "West|37", "総計|37"]


_GREATER_THAN_20 = PivotFilterInput(field_index="Region", type=PivotFilterKind.VALUE_GREATER_THAN, value=20.0)


_BETWEEN_10_20 = PivotFilterInput(
    field_index="Region",
    type=PivotFilterKind.VALUE_BETWEEN,
    value=10.0,
    value_high=20.0,
)


_PIVOT_READERS = {
    "Pivot.get_cache",
    "Pivot.get_column_field_order",
    "Pivot.get_data_field",
    "Pivot.get_data_field_default_name",
    "Pivot.get_field",
    "Pivot.get_field_item",
    "Pivot.get_filter",
    "Pivot.get_grand_totals",
    "Pivot.get_layout",
    "Pivot.get_options",
    "Pivot.get_page_field_order",
    "Pivot.get_properties",
    "Pivot.get_report_layout",
    "Pivot.get_row_field_order",
    "Pivot.get_style_options",
    "Pivot.is_projection_pending",
    "Pivot.list_data_fields",
    "Pivot.list_field_items",
    "Pivot.list_fields",
    "Pivot.list_filters",
    "PivotCache.get_options",
    "PivotCache.get_record_count",
    "PivotCache.get_record_value",
    "PivotCache.get_worksheet_source",
    "PivotCache.list_calculated_items",
    "PivotCache.list_field_shared_items",
    "PivotCache.list_fields",
    "Sheet.list_pivots",
    "Sheet.pivot",
    "Workbook.pivot_cache",
}


def _first_number_cell(f: _PivotFixture) -> Tuple[int, int]:
    """Position of a data cell of the fixture pivot's layout."""
    for cell in f.wb.sheet(0).pivot(f.pivot).get_layout().cells:
        if isinstance(cell.value, float):
            return cell.row, cell.col
    raise AssertionError("pivot layout has no number cell")


def _two_levels(f: _PivotFixture) -> None:
    f.wb.sheet(0).pivot(f.pivot).set_row_field_order([f.region, f.product])


def _date_rows(f: _PivotFixture) -> None:
    f.wb.sheet(0).pivot(f.pivot).set_row_field_order([f.date])


def _hide_east(f: _PivotFixture) -> None:
    f.wb.sheet(0).pivot(f.pivot).update_field_item(f.region, 0, visible=False)


def _point_at_new_item(f: _PivotFixture) -> None:
    f.wb.pivot_cache(f.cache).set_record_value(0, 0, 2)


def _package_part(data: bytes, name: str) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        return package.read(name).decode("utf-8")


class PivotMutatorTableTests(unittest.TestCase):
    # name -> (setup, act, after, expect). `act` is the one call under test
    # and returns its result; `expect(test, fixture, result)` reads it back.
    def _rows(self) -> Dict[str, Tuple[Any, Any, Any, Any]]:
        eq = self.assertEqual
        IDX = StatusCode.INDEX_OUT_OF_RANGE

        def grid_is(expected: List[str]) -> Any:
            return lambda t, f, r: eq(f.grid(), expected)

        def rejects(call: Any, code: StatusCode = StatusCode.INVALID_ARGUMENT) -> None:
            with self.assertRaises(FormulonError) as caught:
                call()
            eq(caught.exception.code, code)

        def cache_create(t: Any, f: _PivotFixture, r: PivotCache) -> None:
            self.assertIsInstance(r, PivotCache)
            eq(r.cache_id, 77)
            self.assertIs(f.wb.pivot_cache(77), r)
            eq(len(f.wb.list_pivot_caches()), 2)

        def cache_remove_setup(f: _PivotFixture) -> None:
            f.spare = f.wb.add_pivot_cache().cache_id  # type: ignore[attr-defined]

        def cache_remove(t: Any, f: _PivotFixture, r: None) -> None:
            eq([c.cache_id for c in f.wb.list_pivot_caches()], [f.cache])

        def cache_source(t: Any, f: _PivotFixture, r: None) -> None:
            source = f.wb.pivot_cache(f.cache).get_worksheet_source()
            eq(source.range, SheetRect(0, CellRect(1, 1, 5, 4)))
            eq(f.grid(), _BASE_GRID)

        def cache_field_add(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 4)
            eq(f.wb.pivot_cache(f.cache).list_fields()[4].name, "Extra")

        def cache_field_clear(t: Any, f: _PivotFixture, r: None) -> None:
            eq(len(f.wb.pivot_cache(f.cache).list_fields()), 0)
            # The data field now names a cache field that no longer exists.
            with self.assertRaises(FormulonError):
                f.wb.sheet(0).pivot(f.pivot).get_layout()

        def record_add(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 4)
            eq(f.grid(), ["行ラベル|Sum of Amount", "East|15", "West|37", "(空白)|", "総計|52"])

        def pivot_create(t: Any, f: _PivotFixture, r: Pivot) -> None:
            self.assertIsInstance(r, Pivot)
            eq(r.index, 1)
            self.assertIs(f.wb.sheet(0).pivot(1), r)
            layout = f.wb.sheet(0).pivot(1).get_layout()
            eq((layout.range.first_row, layout.range.first_col), (20, 1))
            eq(f.grid(0), _BASE_GRID)

        def pivot_remove(t: Any, f: _PivotFixture, r: None) -> None:
            eq(len(f.wb.sheet(0).list_pivots()), 0)
            with self.assertRaises(FormulonError):
                f.wb.sheet(0).pivot(0).get_layout()

        def set_name(t: Any, f: _PivotFixture, r: None) -> None:
            # A pivot's name is only observable in the part the writer emits.
            part = _package_part(f.wb.save(), "xl/pivotTables/pivotTable1.xml")
            self.assertRegex(part, r'<pivotTableDefinition [^>]*name="Renamed"')

        def update_style_options(t: Any, f: _PivotFixture, r: None) -> None:
            style = f.wb.sheet(0).pivot(f.pivot).get_style_options()
            eq(style.style_name, "PivotStyleMedium9")
            eq(
                (style.show_row_headers, style.show_column_headers, style.show_row_stripes, style.show_column_stripes),
                (False, True, True, False),
            )
            self.assertTrue(style.show_last_column)
            eq(f.grid(), _BASE_GRID)

        def set_anchor(t: Any, f: _PivotFixture, r: None) -> None:
            layout = f.wb.sheet(0).pivot(f.pivot).get_layout()
            eq((layout.range.first_row, layout.range.first_col), (10, 2))
            eq(f.grid(), _BASE_GRID)

        def set_layout(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.sheet(0).pivot(f.pivot).get_report_layout(), PivotReportLayout.TABULAR)
            eq(f.grid(), ["Region|Sum of Amount", "East|15", "West|37", "総計|52"])

        def field_add(t: Any, f: _PivotFixture, r: int) -> None:
            # Product is a cache column the fixture already added, so it is updated in place.
            eq(r, f.product)
            eq(f.grid(), ["Fruit|(すべて)", "|", *_BASE_GRID])

        def field_clear(t: Any, f: _PivotFixture, r: None) -> None:
            eq(len(f.wb.sheet(0).pivot(f.pivot).list_fields()), 0)
            # The row order empties with the fields; East's hidden flag goes too.
            eq(f.grid(), ["行ラベル|Sum of Amount", "|52"])

        def add_item_at_setup(f: _PivotFixture) -> None:
            # Only the index-addressed form can name the blank item.
            f.wb.pivot_cache(f.cache).add_field_shared_item(0, None)
            f.wb.pivot_cache(f.cache).set_record_value(0, 0, None)

        def clear_subtotal_setup(f: _PivotFixture) -> None:
            _two_levels(f)
            f.wb.sheet(0).pivot(f.pivot).set_field_subtotal_functions(f.region, [PivotAggregation.MAX])

        def clear_date_group_setup(f: _PivotFixture) -> None:
            _date_rows(f)
            f.wb.sheet(0).pivot(f.pivot).set_field_date_group(
                f.date, PivotDateGroupInput(periods=PivotDatePeriod.YEARS, calendar=PivotCalendar.GREGORIAN)
            )

        def number_format(t: Any, f: _PivotFixture, r: None) -> None:
            rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_field(99, number_format="#,##0.00"), IDX)
            # A format code, not an id: a number is no format code.
            type_rejected(lambda: f.wb.sheet(0).pivot(f.pivot).update_field(f.amount, number_format=4))
            # A field's numFmtId is only observable in the part the writer emits.
            part = _package_part(f.wb.save(), "xl/pivotTables/pivotTable1.xml")
            self.assertRegex(part, r'<pivotField [^>]*numFmtId="4"')
            eq(f.wb.sheet(0).pivot(f.pivot).list_fields()[f.amount].number_format, "#,##0.00")
            eq(f.grid(), _BASE_GRID)

        def data_field_add(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 1)
            eq(
                f.grid(),
                ["行ラベル|Sum of Amount|Max of Amount", "East|15|10", "West|37|30", "総計|52|30"],
            )

        def data_field_set(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.grid(), ["行ラベル|Count of Amount", "East|2", "West|2", "総計|4"])
            formats = [
                c.number_format for c in f.wb.sheet(0).pivot(f.pivot).get_layout().cells if isinstance(c.value, float)
            ]
            eq(formats, ["2", "2", "2"])

        def filter_remove_setup(f: _PivotFixture) -> None:
            # Each filter alone keeps a different region, so removing the
            # wrong index leaves the wrong one.
            f.wb.sheet(0).pivot(f.pivot).add_filter(_GREATER_THAN_20)
            f.wb.sheet(0).pivot(f.pivot).add_filter(_BETWEEN_10_20)

        def type_rejected(call: Any) -> None:
            with self.assertRaises(TypeError):
                call()

        def calc_field_setup(f: _PivotFixture) -> None:
            f.wb.pivot_cache(f.cache).add_calculated_field("Double", "Amount*2")

        def calc_item_setup(f: _PivotFixture) -> None:
            f.wb.pivot_cache(f.cache).add_calculated_item(f.region, "Both", "East+West")

        def calc_two_items_setup(f: _PivotFixture) -> None:
            calc_item_setup(f)
            f.wb.pivot_cache(f.cache).add_calculated_item(f.region, "Diff", "East-West")

        def pending_setup(f: _PivotFixture) -> None:
            f.wb.pivot_cache(f.cache).add_field_shared_item(0, "North")

        def options_set(t: Any, f: _PivotFixture, r: None) -> None:
            o = f.wb.sheet(0).pivot(f.pivot).get_options()
            eq((o.show_error, o.error_caption), (True, "N/A"))
            eq((o.show_missing, o.missing_caption, o.indent, o.page_wrap), (True, "-", 3, 2))
            eq((o.alt_text, o.alt_text_summary, o.grand_total_caption), ("Title", None, "Sum"))
            rejects(lambda: f.wb.sheet(0).pivot(9).update_options(indent=3), IDX)
            rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_options(indent=127))
            rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_options(page_wrap=256))
            rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_options(error_caption="x" * 256))
            type_rejected(lambda: f.wb.sheet(0).pivot(f.pivot).update_options(indent="3"))
            type_rejected(lambda: f.wb.sheet(0).pivot(f.pivot).update_options(error_caption=5))
            type_rejected(lambda: f.wb.sheet(0).pivot(f.pivot).update_options({}))  # a record is no longer accepted
            eq(f.wb.sheet(0).pivot(f.pivot).get_options(), o)

        def cache_options_set(t: Any, f: _PivotFixture, r: None) -> None:
            o = f.wb.pivot_cache(f.cache).get_options()
            eq((o.save_data, o.missing_items_limit, o.refresh_on_load), (False, 32500, True))
            rejects(lambda: f.wb.pivot_cache(9999).update_options(save_data=False), StatusCode.NOT_FOUND)
            rejects(lambda: f.wb.pivot_cache(f.cache).update_options(missing_items_limit=-2))
            rejects(lambda: f.wb.pivot_cache(f.cache).update_options(missing_items_limit=1048577))
            type_rejected(lambda: f.wb.pivot_cache(f.cache).update_options(missing_items_limit="1"))
            eq(f.wb.pivot_cache(f.cache).get_options(), o)

        def calc_field_add(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 4)
            eq(f.wb.pivot_cache(f.cache).list_fields()[4].name, "Double")
            eq(f.wb.pivot_cache(f.cache).list_fields()[4].formula, "=Amount*2")
            eq(f.wb.pivot_cache(f.cache).list_fields()[0].formula, None)
            rejects(
                lambda: f.wb.pivot_cache(f.cache).add_calculated_field("Amount", "Amount*2"), StatusCode.ALREADY_EXISTS
            )
            rejects(lambda: f.wb.pivot_cache(f.cache).add_calculated_field("Other", ""))
            rejects(lambda: f.wb.pivot_cache(f.cache).add_calculated_field("Other", "NoSuchField+1"))
            rejects(lambda: f.wb.pivot_cache(9999).add_calculated_field("Other", "Amount*2"), StatusCode.NOT_FOUND)
            type_rejected(lambda: f.wb.pivot_cache(f.cache).add_calculated_field(5, "Amount*2"))
            type_rejected(lambda: f.wb.pivot_cache(f.cache).add_calculated_field("Other", None))
            eq(len(f.wb.pivot_cache(f.cache).list_fields()), 5)

        def calc_field_set_formula(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.pivot_cache(f.cache).list_fields()[4].formula, "=Amount*3")
            rejects(lambda: f.wb.pivot_cache(f.cache).set_calculated_field_formula(0, "Amount*3"))
            rejects(lambda: f.wb.pivot_cache(f.cache).set_calculated_field_formula(4, ""))
            type_rejected(lambda: f.wb.pivot_cache(f.cache).set_calculated_field_formula(4, 3))
            eq(f.wb.pivot_cache(f.cache).list_fields()[4].formula, "=Amount*3")

        def calc_field_remove(t: Any, f: _PivotFixture, r: None) -> None:
            eq(len(f.wb.pivot_cache(f.cache).list_fields()), 4)
            rejects(lambda: f.wb.pivot_cache(f.cache).remove_calculated_field(0))

        def calc_item_add(t: Any, f: _PivotFixture, r: None) -> None:
            eq(len(f.wb.pivot_cache(f.cache).list_calculated_items()), 1)
            item = f.wb.pivot_cache(f.cache).list_calculated_items()[0]
            eq((item.field_index, item.name, item.scoped), (f.region, "Both", False))
            eq(item.formula, "=Region[East]+Region[West]")
            with self.assertRaises(IndexError):
                f.wb.pivot_cache(f.cache).list_calculated_items()[5]
            rejects(
                lambda: f.wb.pivot_cache(f.cache).add_calculated_item(f.region, "East", "East+West"),
                StatusCode.ALREADY_EXISTS,
            )
            rejects(lambda: f.wb.pivot_cache(f.cache).add_calculated_item(f.region, "Other", ""))
            rejects(lambda: f.wb.pivot_cache(f.cache).add_calculated_item(99, "Other", "East+West"), IDX)
            type_rejected(lambda: f.wb.pivot_cache(f.cache).add_calculated_item(f.region, None, "East+West"))
            eq(len(f.wb.pivot_cache(f.cache).list_calculated_items()), 1)

        def calc_item_set_formula(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.pivot_cache(f.cache).list_calculated_items()[0].formula, "=Region[East]*2")
            rejects(lambda: f.wb.pivot_cache(f.cache).set_calculated_item_formula(5, "East*2"), IDX)
            type_rejected(lambda: f.wb.pivot_cache(f.cache).set_calculated_item_formula(0, 2))

        def calc_item_move(t: Any, f: _PivotFixture, r: None) -> None:
            eq([f.wb.pivot_cache(f.cache).list_calculated_items()[i].name for i in (0, 1)], ["Diff", "Both"])
            rejects(lambda: f.wb.pivot_cache(f.cache).move_calculated_item(0, 5), IDX)
            type_rejected(lambda: f.wb.pivot_cache(f.cache).move_calculated_item(0, "1"))

        def calc_item_remove(t: Any, f: _PivotFixture, r: None) -> None:
            eq(len(f.wb.pivot_cache(f.cache).list_calculated_items()), 0)
            rejects(lambda: f.wb.pivot_cache(f.cache).remove_calculated_item(0), IDX)

        def custom_name(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.sheet(0).pivot(f.pivot).list_fields()[f.region].custom_name, "Area")
            rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_field(99, custom_name="x"), IDX)
            rejects(
                lambda: f.wb.sheet(0).pivot(f.pivot).update_field(f.region, custom_name="Product"),
                StatusCode.ALREADY_EXISTS,
            )
            rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_field(f.region, custom_name="x" * 256))
            type_rejected(lambda: f.wb.sheet(0).pivot(f.pivot).update_field(f.region, custom_name=5))
            eq(f.wb.sheet(0).pivot(f.pivot).list_fields()[f.region].custom_name, "Area")
            f.wb.sheet(0).pivot(f.pivot).update_field(f.region, custom_name=None)
            eq(f.wb.sheet(0).pivot(f.pivot).list_fields()[f.region].custom_name, None)
            f.wb.sheet(0).pivot(f.pivot).update_field(f.region, custom_name="Region")
            f.wb.sheet(0).pivot(f.pivot).update_field(f.region, custom_name="")
            eq(f.wb.sheet(0).pivot(f.pivot).list_fields()[f.region].custom_name, None)

        def project(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.sheet(0).pivot(f.pivot).is_projection_pending(), False)
            rejects(lambda: f.wb.sheet(0).pivot(9).project(), IDX)
            rejects(lambda: f.wb.sheet(0).pivot(9).is_projection_pending(), IDX)

        W = PivotAggregation
        return {
            "Pivot.update_options": (
                None,
                lambda f: (
                    f.wb.sheet(0)
                    .pivot(f.pivot)
                    .update_options(
                        show_error=True,
                        error_caption="N/A",
                        show_missing=True,
                        missing_caption="-",
                        indent=3,
                        page_wrap=2,
                        alt_text="Title",
                        alt_text_summary="",
                        grand_total_caption="Sum",
                    )
                ),
                None,
                options_set,
            ),
            "PivotCache.update_options": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).update_options(
                    save_data=False, missing_items_limit=32500, refresh_on_load=True
                ),
                None,
                cache_options_set,
            ),
            "PivotCache.add_calculated_field": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).add_calculated_field("Double", "Amount*2"),
                None,
                calc_field_add,
            ),
            "PivotCache.set_calculated_field_formula": (
                calc_field_setup,
                lambda f: f.wb.pivot_cache(f.cache).set_calculated_field_formula(4, "Amount*3"),
                None,
                calc_field_set_formula,
            ),
            "PivotCache.remove_calculated_field": (
                calc_field_setup,
                lambda f: f.wb.pivot_cache(f.cache).remove_calculated_field(4),
                None,
                calc_field_remove,
            ),
            "PivotCache.add_calculated_item": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).add_calculated_item(f.region, "Both", "East+West"),
                None,
                calc_item_add,
            ),
            "PivotCache.set_calculated_item_formula": (
                calc_item_setup,
                lambda f: f.wb.pivot_cache(f.cache).set_calculated_item_formula(0, "East*2"),
                None,
                calc_item_set_formula,
            ),
            "PivotCache.move_calculated_item": (
                calc_two_items_setup,
                lambda f: f.wb.pivot_cache(f.cache).move_calculated_item(0, 1),
                None,
                calc_item_move,
            ),
            "PivotCache.remove_calculated_item": (
                calc_item_setup,
                lambda f: f.wb.pivot_cache(f.cache).remove_calculated_item(0),
                None,
                calc_item_remove,
            ),
            "Pivot.update_field#custom_name": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_field(f.region, custom_name="Area"),
                None,
                custom_name,
            ),
            "Pivot.update_field#show_all": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_field(f.region, show_all=True),
                None,
                lambda t, f, r: (
                    eq(f.wb.sheet(0).pivot(f.pivot).list_fields()[f.region].show_all, True),
                    eq(f.wb.sheet(0).pivot(f.pivot).list_fields()[f.product].show_all, False),
                    rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_field(99, show_all=True), IDX),
                    rejects(
                        lambda: f.wb.sheet(0).pivot(f.pivot).update_field("0", show_all=False), StatusCode.NOT_FOUND
                    ),
                    eq(f.wb.sheet(0).pivot(f.pivot).list_fields()[f.region].show_all, True),
                ),
            ),
            "Pivot.project": (
                pending_setup,
                lambda f: f.wb.sheet(0).pivot(f.pivot).project(),
                None,
                project,
            ),
            "Workbook.add_pivot_cache": (None, lambda f: f.wb.add_pivot_cache(77), None, cache_create),
            "Workbook.list_pivot_caches": (
                None,
                lambda f: f.wb.list_pivot_caches(),
                None,
                lambda t, f, r: eq([c.cache_id for c in r], [f.cache]),
            ),
            "Workbook.remove_pivot_cache": (
                cache_remove_setup,
                lambda f: f.wb.remove_pivot_cache(f.spare),
                None,
                cache_remove,
            ),
            "PivotCache.set_worksheet_source": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).set_worksheet_source(
                    PivotWorksheetSourceInput(range="B2:E6", sheet="Sheet1")
                ),
                None,
                cache_source,
            ),
            "PivotCache.add_field": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).add_field("Extra"),
                None,
                cache_field_add,
            ),
            "PivotCache.clear_fields": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).clear_fields(),
                None,
                cache_field_clear,
            ),
            "PivotCache.add_field_shared_item#number": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).add_field_shared_item(0, 42.5),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "42.5|10", "East|5", "West|37", "総計|52"]),
            ),
            "PivotCache.add_field_shared_item#date": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).add_field_shared_item(0, datetime.date(2024, 1, 1)),
                None,
                lambda t, f, r: eq(f.wb.pivot_cache(f.cache).list_field_shared_items(0)[2].is_date, True),
            ),
            "PivotCache.add_field_shared_item#text": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).add_field_shared_item(0, "North"),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "North|10", "West|37", "総計|52"]),
            ),
            "PivotCache.add_field_shared_item#bool": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).add_field_shared_item(0, True),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "TRUE|10", "総計|52"]),
            ),
            "PivotCache.add_field_shared_item#blank": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).add_field_shared_item(0, None),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "(空白)|10", "総計|52"]),
            ),
            "PivotCache.add_field_shared_item#error": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).add_field_shared_item(0, ExcelError(ErrorCode(6))),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "#N/A|10", "総計|52"]),
            ),
            "PivotCache.clear_field_shared_items": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).clear_field_shared_items(0),
                None,
                # With no shared items the records' indices render as themselves.
                grid_is(["行ラベル|Sum of Amount", "0|15", "1|37", "総計|52"]),
            ),
            "PivotCache.add_record": (None, lambda f: f.wb.pivot_cache(f.cache).add_record(), None, record_add),
            "PivotCache.clear_records": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).clear_records(),
                None,
                grid_is(["行ラベル|Sum of Amount", "総計|"]),
            ),
            "PivotCache.set_record_value#number": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).set_record_value(0, 2, 1000),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|1005", "West|37", "総計|1042"]),
            ),
            "PivotCache.set_record_value#date": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).set_record_value(0, 2, datetime.date(2024, 1, 1)),
                None,
                lambda t, f, r: eq(r, None),
            ),
            "PivotCache.set_record_value#text": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).set_record_value(0, 0, "West"),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|47", "総計|52"]),
            ),
            "PivotCache.set_record_value#bool": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).set_record_value(0, 2, True),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "総計|42"]),
            ),
            "PivotCache.set_record_value#blank": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).set_record_value(0, 2, None),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "総計|42"]),
            ),
            "PivotCache.set_record_value#error": (
                None,
                lambda f: f.wb.pivot_cache(f.cache).set_record_value(0, 2, ExcelError(ErrorCode(1))),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|#1", "West|37", "総計|#1"]),
            ),
            "Sheet.add_pivot": (
                None,
                lambda f: f.wb.sheet(0).add_pivot(
                    PivotInput(name="Pivot2", cache_id=f.cache, anchor=CellAddress(20, 1))
                ),
                None,
                pivot_create,
            ),
            "Sheet.remove_pivot": (None, lambda f: f.wb.sheet(0).remove_pivot(f.pivot), None, pivot_remove),
            "Pivot.rename": (None, lambda f: f.wb.sheet(0).pivot(f.pivot).rename("Renamed"), None, set_name),
            "Pivot.set_anchor": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).set_anchor(CellAddress(10, 2)),
                None,
                set_anchor,
            ),
            "Pivot.update_grand_totals": (
                lambda f: f.wb.sheet(0).pivot(f.pivot).set_column_field_order([f.product]),
                # Rows off drops the per-row total column; columns on keeps
                # the bottom total row, so a swapped pair shows the opposite.
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_grand_totals(rows=False, columns=True),
                None,
                grid_is(["Sum of Amount|列ラベル|", "行ラベル|Apple|Pear", "East|10|5", "West|30|7", "総計|40|12"]),
            ),
            "Pivot.update_style_options": (
                None,
                lambda f: (
                    f.wb.sheet(0)
                    .pivot(f.pivot)
                    .update_style_options(
                        style_name="PivotStyleMedium9",
                        show_row_headers=False,
                        show_column_headers=True,
                        show_row_stripes=True,
                        show_column_stripes=False,
                        show_last_column=True,
                    )
                ),
                None,
                update_style_options,
            ),
            "Pivot.set_report_layout": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).set_report_layout(PivotReportLayout.TABULAR),
                None,
                set_layout,
            ),
            "Pivot.add_field": (
                None,
                lambda f: (
                    f.wb.sheet(0)
                    .pivot(f.pivot)
                    .add_field(PivotFieldInput(field_index="Product", custom_name="Fruit", axis=PivotAxis.PAGE))
                ),
                None,
                field_add,
            ),
            "Pivot.clear_fields": (
                _hide_east,
                lambda f: f.wb.sheet(0).pivot(f.pivot).clear_fields(),
                None,
                field_clear,
            ),
            "Pivot.set_field_axis": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).set_field_axis(f.product, PivotAxis.PAGE),
                None,
                grid_is(["Product|(すべて)", "|", *_BASE_GRID]),
            ),
            "Pivot.update_field#sort_order": (
                # East now outsums West, so sorting ascending by the data
                # field differs from both label orders.
                lambda f: f.wb.pivot_cache(f.cache).set_record_value(0, 2, 40),
                lambda f: (
                    f.wb.sheet(0)
                    .pivot(f.pivot)
                    .update_field(f.region, sort_order=PivotSortOrder.ASCENDING, sort_by_field="Sum of Amount")
                ),
                None,
                grid_is(["行ラベル|Sum of Amount", "West|37", "East|45", "総計|82"]),
            ),
            "Pivot.update_field#subtotal_top": (
                _two_levels,
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_field(f.region, subtotal_top=False),
                None,
                grid_is(
                    [
                        "行ラベル|Sum of Amount",
                        "East|",
                        "Apple|10",
                        "Pear|5",
                        "East 集計|15",
                        "West|",
                        "Apple|30",
                        "Pear|7",
                        "West 集計|37",
                        "総計|52",
                    ]
                ),
            ),
            "Pivot.add_field_item": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).add_field_item(f.region, "East", visible=False),
                None,
                grid_is(_WEST_ONLY),
            ),
            "Pivot.add_field_item#cache_index": (
                add_item_at_setup,
                lambda f: f.wb.sheet(0).pivot(f.pivot).add_field_item(f.region, cache_index=2, visible=False),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "総計|42"]),
            ),
            "Pivot.clear_field_items": (
                _hide_east,
                lambda f: f.wb.sheet(0).pivot(f.pivot).clear_field_items(f.region),
                None,
                grid_is(_BASE_GRID),
            ),
            "Pivot.update_field_item#visible": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_field_item(f.region, 1, visible=False),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|15", "総計|15"]),
            ),
            "Pivot.set_field_subtotal_functions": (
                _two_levels,
                lambda f: f.wb.sheet(0).pivot(f.pivot).set_field_subtotal_functions(f.region, [W.MAX]),
                None,
                grid_is(
                    [
                        "行ラベル|Sum of Amount",
                        "East|10",
                        "Apple|10",
                        "Pear|5",
                        "West|30",
                        "Apple|30",
                        "Pear|7",
                        "総計|52",
                    ]
                ),
            ),
            "Pivot.set_field_subtotal_functions#empty": (
                clear_subtotal_setup,
                lambda f: f.wb.sheet(0).pivot(f.pivot).set_field_subtotal_functions(f.region, []),
                None,
                grid_is(_TWO_LEVEL_GRID),
            ),
            "Pivot.set_field_date_group": (
                _date_rows,
                lambda f: (
                    f.wb.sheet(0)
                    .pivot(f.pivot)
                    .set_field_date_group(
                        f.date, PivotDateGroupInput(periods=PivotDatePeriod.YEARS, calendar=PivotCalendar.JAPANESE)
                    )
                ),
                None,
                grid_is(["行ラベル|Sum of Amount", "令和6年|40", "令和7年|12", "総計|52"]),
            ),
            "Pivot.set_field_date_group#clears": (
                clear_date_group_setup,
                lambda f: f.wb.sheet(0).pivot(f.pivot).set_field_date_group(f.date, None),
                None,
                grid_is(["行ラベル|Sum of Amount", "45292|10", "45323|30", "45658|5", "45689|7", "総計|52"]),
            ),
            "Pivot.update_field#number_format": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_field(f.amount, number_format="#,##0.00"),
                None,
                number_format,
            ),
            "Pivot.set_row_field_order": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).set_row_field_order([f.product]),
                None,
                grid_is(["行ラベル|Sum of Amount", "Apple|40", "Pear|12", "総計|52"]),
            ),
            "Pivot.set_page_field_order": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).set_page_field_order([f.product]),
                None,
                lambda t, f, r: (
                    eq(f.wb.sheet(0).pivot(f.pivot).get_page_field_order(), [f.product]),
                    eq(f.wb.sheet(0).pivot(f.pivot).get_row_field_order(), [f.region]),
                ),
            ),
            "Pivot.set_column_field_order": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).set_column_field_order([f.product]),
                None,
                grid_is(
                    [
                        "Sum of Amount|列ラベル||",
                        "行ラベル|Apple|Pear|総計",
                        "East|10|5|15",
                        "West|30|7|37",
                        "総計|40|12|52",
                    ]
                ),
            ),
            "Pivot.add_data_field": (
                None,
                lambda f: (
                    f.wb.sheet(0)
                    .pivot(f.pivot)
                    .add_data_field(PivotDataFieldInput(name="Max of Amount", field_index=f.amount, aggregation=W.MAX))
                ),
                None,
                data_field_add,
            ),
            "Pivot.clear_data_fields": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).clear_data_fields(),
                None,
                grid_is(["行ラベル", "East", "West", "総計"]),
            ),
            "Pivot.set_data_field": (
                None,
                lambda f: (
                    f.wb.sheet(0)
                    .pivot(f.pivot)
                    .set_data_field(
                        0,
                        PivotDataFieldInput(
                            name="Count of Amount", field_index=f.amount, aggregation=W.COUNT, number_format="2"
                        ),
                    )
                ),
                None,
                data_field_set,
            ),
            "Pivot.add_filter": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).add_filter(_GREATER_THAN_20),
                None,
                grid_is(_WEST_ONLY),
            ),
            "Pivot.clear_filters": (
                lambda f: f.wb.sheet(0).pivot(f.pivot).add_filter(_GREATER_THAN_20),
                lambda f: f.wb.sheet(0).pivot(f.pivot).clear_filters(),
                None,
                grid_is(_BASE_GRID),
            ),
            "Pivot.remove_filter": (
                filter_remove_setup,
                lambda f: f.wb.sheet(0).pivot(f.pivot).remove_filter(1),
                None,
                grid_is(_WEST_ONLY),
            ),
            "Pivot.update_field#default_subtotal": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_field(f.region, default_subtotal=False),
                None,
                lambda t, f, r: (
                    eq(f.wb.sheet(0).pivot(f.pivot).list_fields()[f.region].default_subtotal, False),
                    rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_field(99, default_subtotal=True), IDX),
                ),
            ),
            "Pivot.update_field#insert_blank_row": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_field(f.region, insert_blank_row=True),
                None,
                lambda t, f, r: (
                    eq(f.wb.sheet(0).pivot(f.pivot).list_fields()[f.region].insert_blank_row, True),
                    rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_field(99, insert_blank_row=True), IDX),
                ),
            ),
            "Pivot.update_field#fill_down_labels": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_field(f.region, fill_down_labels=True),
                None,
                lambda t, f, r: (
                    eq(f.wb.sheet(0).pivot(f.pivot).list_fields()[f.region].fill_down_labels, True),
                    rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_field(99, fill_down_labels=True), IDX),
                ),
            ),
            "Pivot.update_field#show_detail_all": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_field(f.region, show_detail=False),
                None,
                lambda t, f, r: (
                    eq(
                        [f.wb.sheet(0).pivot(f.pivot).get_field_item(f.region, i).show_detail for i in (0, 1)],
                        [False, False],
                    ),
                    rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_field(99, show_detail=False), IDX),
                ),
            ),
            "Pivot.update_field_item#show_detail": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_field_item(f.region, 1, show_detail=False),
                None,
                lambda t, f, r: (
                    eq(
                        [f.wb.sheet(0).pivot(f.pivot).get_field_item(f.region, i).show_detail for i in (0, 1)],
                        [True, False],
                    ),
                    rejects(
                        lambda: f.wb.sheet(0).pivot(f.pivot).update_field_item(f.region, 9, show_detail=False), IDX
                    ),
                ),
            ),
            "Pivot.update_field_item#caption": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).update_field_item(f.region, 0, caption="Eastern"),
                None,
                lambda t, f, r: (
                    eq(f.wb.sheet(0).pivot(f.pivot).get_field_item(f.region, 0).caption, "Eastern"),
                    eq(f.wb.sheet(0).pivot(f.pivot).get_field_item(f.region, 1).caption, None),
                    f.wb.sheet(0).pivot(f.pivot).update_field_item(f.region, 0, caption=None),
                    eq(f.wb.sheet(0).pivot(f.pivot).get_field_item(f.region, 0).caption, None),
                    rejects(lambda: f.wb.sheet(0).pivot(f.pivot).update_field_item(f.region, 9, caption="x"), IDX),
                    rejects(
                        lambda: f.wb.sheet(0).pivot(f.pivot).update_field_item(f.region, 0, caption="West"),
                        StatusCode.ALREADY_EXISTS,
                    ),
                ),
            ),
            "Pivot.show_details": (
                None,
                lambda f: f.wb.sheet(0).pivot(f.pivot).show_details(*_first_number_cell(f)),
                None,
                lambda t, f, r: (
                    eq(r.index, 0),
                    eq(len(f.wb), 2),
                    rejects(lambda: f.wb.sheet(1).pivot(f.pivot).show_details(0, 0)),
                ),
            ),
        }

    def test_table_covers_every_pivot_mutator(self) -> None:
        declared = {
            f"{cls.__name__}.{name}"
            for cls in (Workbook, Sheet, Pivot, PivotCache)
            for name, member in inspect.getmembers(cls, inspect.isfunction)
            if (cls in (Pivot, PivotCache) or "pivot" in name)
            and not name.startswith("_")
            and f"{cls.__name__}.{name}" not in _PIVOT_READERS
        }
        self.assertGreater(len(declared), 40)
        rows = {name.split("#")[0] for name in self._rows()}
        self.assertEqual(sorted(declared - rows), [], "pivot mutators without a row")
        self.assertEqual(sorted(rows - declared), [], "rows for undeclared pivot mutators")

    def test_manual_sort_order_keeps_item_order(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            wb.sheet(0).pivot(f.pivot).clear_field_items(f.region)
            wb.sheet(0).pivot(f.pivot).add_field_item(f.region, "West", visible=True)
            wb.sheet(0).pivot(f.pivot).add_field_item(f.region, "East", visible=True)
            wb.sheet(0).pivot(f.pivot).update_field(f.region, sort_order=PivotSortOrder.MANUAL)
            self.assertEqual(f.grid(), ["行ラベル|Sum of Amount", "West|37", "East|15", "総計|52"])
            wb.sheet(0).pivot(f.pivot).update_field(f.region, sort_order=PivotSortOrder.ASCENDING)
            self.assertEqual(f.grid(), _BASE_GRID)
            with self.assertRaises(FormulonError):
                wb.sheet(0).pivot(f.pivot).update_field(f.region, sort_order=3)
            with self.assertRaises(FormulonError):
                wb.sheet(0).pivot(f.pivot).update_field(
                    f.region, sort_order=PivotSortOrder.MANUAL, sort_by_field="Sum of Amount"
                )

    def test_pivot_mutators_forward_their_arguments(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(_PivotFixture(wb).grid(), _BASE_GRID)
        for name, (setup, act, after, expect) in self._rows().items():
            with self.subTest(name), Workbook.create_default() as wb:
                fixture = _PivotFixture(wb)
                if setup is not None:
                    setup(fixture)
                result = act(fixture)
                if after is not None:
                    after(fixture)
                expect(self, fixture, result)


class PivotDefinitionReadBackTests(unittest.TestCase):
    def _assert_fixture_definition(self, wb: Workbook, f: _PivotFixture, pivot: int) -> None:
        eq = self.assertEqual
        eq(wb.sheet(0).pivot(pivot).name, "Pivot1")
        eq(wb.sheet(0).pivot(pivot).get_grand_totals(), PivotGrandTotals(rows=True, columns=True))
        eq(wb.sheet(0).pivot(pivot).get_row_field_order(), [f.region])
        eq(wb.sheet(0).pivot(pivot).get_column_field_order(), [])

        region = wb.sheet(0).pivot(pivot).list_fields()[f.region]
        eq(region.source_name, "Region")
        eq(region.axis, PivotAxis.ROW)
        eq(region.cache_field_index, 0)
        eq(region.sort_order, PivotSortOrder.ASCENDING)
        eq(region.date_periods, PivotDatePeriod(0))
        eq(region.group_base_index, None)
        eq((region.insert_blank_row, region.fill_down_labels), (False, False))

        eq(len(wb.sheet(0).pivot(pivot).list_field_items(f.region)), 2)
        for index, label in enumerate(("East", "West")):
            item = wb.sheet(0).pivot(pivot).get_field_item(f.region, index)
            eq((item.name, item.cache_index, item.visible), (label, index, True))

        eq(len(wb.sheet(0).pivot(pivot).list_data_fields()), 1)
        data_field = wb.sheet(0).pivot(pivot).list_data_fields()[0]
        eq(data_field.name, "Sum of Amount")
        eq(data_field.field_index, f.amount)
        eq(data_field.aggregation, PivotAggregation.SUM)
        eq((data_field.show_as_base_field, data_field.show_as_base_item), (None, None))

    def test_reads_back_what_the_setters_wrote(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            self.assertEqual(wb.sheet(0).pivot(f.pivot).get_cache().cache_id, f.cache)
            self._assert_fixture_definition(wb, f, f.pivot)

            wb.sheet(0).pivot(f.pivot).update_grand_totals(rows=False, columns=True)
            wb.sheet(0).pivot(f.pivot).update_field(
                f.region, sort_order=PivotSortOrder.DESCENDING, sort_by_field="Sum of Amount"
            )
            wb.sheet(0).pivot(f.pivot).update_field_item(f.region, 1, visible=False)
            wb.sheet(0).pivot(f.pivot).set_field_date_group(
                f.date, PivotDateGroupInput(periods=PivotDatePeriod.YEARS, calendar=PivotCalendar.GREGORIAN)
            )
            self.assertEqual(wb.sheet(0).pivot(f.pivot).get_grand_totals(), PivotGrandTotals(rows=False, columns=True))
            region = wb.sheet(0).pivot(f.pivot).list_fields()[f.region]
            self.assertEqual((region.sort_order, region.sort_by_field), (PivotSortOrder.DESCENDING, "Sum of Amount"))
            self.assertFalse(wb.sheet(0).pivot(f.pivot).get_field_item(f.region, 1).visible)
            date = wb.sheet(0).pivot(f.pivot).list_fields()[f.date]
            self.assertEqual(date.date_periods, PivotDatePeriod.YEARS)
            self.assertEqual(date.date_calendar, PivotCalendar.GREGORIAN)
            self.assertEqual(date.axis, PivotAxis.VALUE)

            wb.sheet(0).pivot(f.pivot).set_field_axis(f.product, PivotAxis.COLUMN)
            self.assertEqual(wb.sheet(0).pivot(f.pivot).get_column_field_order(), [f.product])
            wb.sheet(0).pivot(f.pivot).set_row_field_order([f.product, f.region])
            self.assertEqual(wb.sheet(0).pivot(f.pivot).get_row_field_order(), [f.product, f.region])
            self.assertEqual(wb.sheet(0).pivot(f.pivot).get_column_field_order(), [])

    def test_the_column_axis_member_is_spelled_out(self) -> None:
        self.assertEqual(PivotAxis.COLUMN, 1)
        self.assertFalse(hasattr(PivotAxis, "COL"))
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            pivot = wb.sheet(0).pivot(f.pivot)
            added = pivot.add_field(PivotFieldInput(field_index="Date", axis=PivotAxis.COLUMN))
            self.assertIn(added, pivot.get_column_field_order())
            self.assertEqual(pivot.list_fields()[added].axis, PivotAxis.COLUMN)

    def test_unknown_option_names_are_refused(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            with self.assertRaises(TypeError):
                wb.sheet(0).pivot(f.pivot).add_field_item(f.region, "East", visble=False)  # type: ignore[call-arg]

    def test_page_field_order_is_the_whole_axis_membership(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            pivot = wb.sheet(0).pivot(f.pivot)
            self.assertEqual(pivot.get_page_field_order(), [])
            pivot.set_page_field_order([f.product, f.date])
            self.assertEqual(pivot.get_page_field_order(), [f.product, f.date])
            pivot.set_page_field_order([f.date, f.product])
            self.assertEqual(pivot.get_page_field_order(), [f.date, f.product])
            self.assertEqual(pivot.list_fields()[f.product].axis, PivotAxis.PAGE)
            # A field left out leaves the axis.
            pivot.set_page_field_order([f.date])
            self.assertEqual(pivot.get_page_field_order(), [f.date])
            self.assertNotEqual(pivot.list_fields()[f.product].axis, PivotAxis.PAGE)
            # A listed row field moves onto the page axis.
            pivot.set_page_field_order([f.region, f.date])
            self.assertEqual(pivot.get_page_field_order(), [f.region, f.date])
            self.assertEqual(pivot.get_row_field_order(), [])
            pivot.set_page_field_order([])
            self.assertEqual(pivot.get_page_field_order(), [])

    def test_page_field_order_refuses_bad_lists_and_changes_nothing(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            pivot = wb.sheet(0).pivot(f.pivot)
            pivot.set_page_field_order([f.product])
            for bad, error in (
                ([f.date, f.date], InvalidArgumentError),
                ([PIVOT_VALUES_FIELD], InvalidArgumentError),
                ([99], IndexOutOfRangeError),
            ):
                with self.subTest(bad=bad), self.assertRaises(error):
                    pivot.set_page_field_order(bad)
                self.assertEqual(pivot.get_page_field_order(), [f.product])

    def test_values_level_marker_round_trips_in_field_order(self) -> None:
        self.assertEqual(PIVOT_VALUES_FIELD, -2)
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            wb.sheet(0).pivot(f.pivot).add_data_field(
                PivotDataFieldInput(name="Count of Amount", field_index=f.amount, aggregation=PivotAggregation.COUNT)
            )
            wb.sheet(0).pivot(f.pivot).set_column_field_order([PIVOT_VALUES_FIELD])
            self.assertEqual(wb.sheet(0).pivot(f.pivot).get_column_field_order(), [PIVOT_VALUES_FIELD])
            with self.assertRaises(FormulonError):
                wb.sheet(0).pivot(f.pivot).set_column_field_order([-3])

    def test_style_options_defaults_on_create_and_round_trips(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            style = wb.sheet(0).pivot(f.pivot).get_style_options()
            self.assertIsNotNone(style)
            self.assertEqual(style.style_name, "PivotStyleLight16")
            self.assertEqual(
                (
                    style.show_row_headers,
                    style.show_column_headers,
                    style.show_row_stripes,
                    style.show_column_stripes,
                    style.show_last_column,
                ),
                (True, True, False, False, True),
            )
            wb.sheet(0).pivot(f.pivot).update_style_options(style_name="PivotStyleDark3", show_row_stripes=True)
            data = wb.save()
            # A None name keeps the element without a style name.
            wb.sheet(0).pivot(f.pivot).update_style_options(style_name=None)
            bare = wb.sheet(0).pivot(f.pivot).get_style_options()
            self.assertIsNotNone(bare)
            self.assertIsNone(bare.style_name)
        with Workbook.load(data) as loaded:
            style = loaded.sheet(0).pivot(0).get_style_options()
            self.assertEqual(style.style_name, "PivotStyleDark3")
            self.assertEqual(
                (style.show_row_headers, style.show_row_stripes, style.show_last_column),
                (True, True, True),
            )
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            with self.assertRaises(FormulonError):
                wb.sheet(0).pivot(9).get_style_options()
            with self.assertRaises(FormulonError):
                wb.sheet(0).pivot(9).update_style_options(style_name="X")

    def test_reads_back_a_saved_and_reloaded_pivot(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            data = wb.save()
        with Workbook.load(data) as loaded:
            self.assertEqual(len(loaded.sheet(0).list_pivots()), 1)
            self.assertEqual(loaded.sheet(0).pivot(0).get_cache().cache_id, f.cache)
            self._assert_fixture_definition(loaded, f, 0)

    def test_filters_and_subtotal_functions_survive_save_and_reload(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            wb.sheet(0).pivot(f.pivot).add_filter(_GREATER_THAN_20)
            wb.sheet(0).pivot(f.pivot).set_field_subtotal_functions(
                f.region, [PivotAggregation.AVERAGE, PivotAggregation.SUM]
            )
            region = wb.sheet(0).pivot(f.pivot).list_fields()[f.region]
            self.assertFalse(region.default_subtotal)
            self.assertEqual(region.subtotal_functions, [PivotAggregation.SUM, PivotAggregation.AVERAGE])
            amount = wb.sheet(0).pivot(f.pivot).list_fields()[f.amount]
            self.assertEqual((amount.default_subtotal, amount.subtotal_functions), (True, []))
            data = wb.save()
        self.assertIn(
            '<filters count="1"><filter fld="0" type="valueGreaterThan"',
            _package_part(data, "xl/pivotTables/pivotTable1.xml"),
        )
        with Workbook.load(data) as loaded:
            loaded.set_profile_id("win-365-ja_JP")
            self.assertEqual(len(loaded.sheet(0).pivot(0).list_filters()), 1)
            spec = loaded.sheet(0).pivot(0).list_filters()[0]
            self.assertEqual((spec.type, spec.field_index, spec.value), (_GREATER_THAN_20.type, f.region, 20.0))
            self.assertEqual(_grid(loaded), _WEST_ONLY)
            region = loaded.sheet(0).pivot(0).list_fields()[f.region]
            self.assertFalse(region.default_subtotal)
            self.assertEqual(region.subtotal_functions, [PivotAggregation.SUM, PivotAggregation.AVERAGE])
            loaded.sheet(0).pivot(0).clear_filters()
            self.assertNotIn("<filters", _package_part(loaded.save(), "xl/pivotTables/pivotTable1.xml"))

    def test_out_of_range_indices_raise(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            invalid, bad_sheet = StatusCode.INDEX_OUT_OF_RANGE, StatusCode.SHEET_INDEX_OUT_OF_RANGE
            for call, code in (
                (lambda: wb.sheet(0).pivot(9).get_cache(), invalid),
                (lambda: wb.sheet(0).pivot(9).name, invalid),
                (lambda: wb.sheet(9).pivot(f.pivot).get_grand_totals(), bad_sheet),
                (lambda: wb.sheet(0).pivot(9).get_row_field_order(), invalid),
                (lambda: wb.sheet(0).pivot(9).get_column_field_order(), invalid),
                (lambda: len(wb.sheet(0).pivot(f.pivot).list_field_items(99)), invalid),
                (lambda: wb.sheet(0).pivot(f.pivot).get_field_item(f.region, 99), invalid),
            ):
                with self.assertRaises(FormulonError) as caught:
                    call()
                self.assertEqual(caught.exception.code, code)

    def test_empty_data_field_name_is_derived_from_the_profile(self) -> None:
        for profile, expected in (("win-365-ja_JP", "合計 / Amount"), ("win-365-en_US", "Sum of Amount")):
            for name in ("", None):
                with Workbook.create_default() as wb:
                    f = _PivotFixture(wb)
                    wb.set_profile_id(profile)
                    wb.sheet(0).pivot(f.pivot).clear_data_fields()
                    self.assertEqual(
                        wb.sheet(0).pivot(f.pivot).get_data_field_default_name(f.amount, PivotAggregation.SUM), expected
                    )
                    idx = (
                        wb.sheet(0).pivot(f.pivot).add_data_field(PivotDataFieldInput(name=name, field_index=f.amount))
                    )
                    self.assertEqual(wb.sheet(0).pivot(f.pivot).list_data_fields()[idx].name, expected)
                    # The same source twice: the second name is numbered.
                    self.assertEqual(
                        wb.sheet(0).pivot(f.pivot).get_data_field_default_name(f.amount, PivotAggregation.SUM),
                        expected + "2",
                    )
                    idx = (
                        wb.sheet(0).pivot(f.pivot).add_data_field(PivotDataFieldInput(name=name, field_index=f.amount))
                    )
                    self.assertEqual(wb.sheet(0).pivot(f.pivot).list_data_fields()[idx].name, expected + "2")
                    # Replacing an entry does not count its own name as taken.
                    self.assertEqual(
                        wb.sheet(0).pivot(f.pivot).get_data_field_default_name(f.amount, PivotAggregation.SUM, 0),
                        expected,
                    )
                    wb.sheet(0).pivot(f.pivot).set_data_field(
                        0, PivotDataFieldInput(name=name, field_index=f.amount, aggregation=PivotAggregation.COUNT)
                    )
                    self.assertTrue(wb.sheet(0).pivot(f.pivot).list_data_fields()[0].name.endswith(" Amount"))

    def test_default_name_rejects_bad_arguments(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            for args, code in (
                ((0, 9, f.amount, 0), StatusCode.INDEX_OUT_OF_RANGE),
                ((0, f.pivot, 99, 0), StatusCode.INDEX_OUT_OF_RANGE),
                ((0, f.pivot, f.amount, 99), StatusCode.INVALID_ARGUMENT),
                ((0, f.pivot, f.amount, 0, 5), StatusCode.INDEX_OUT_OF_RANGE),
            ):
                with self.assertRaises(FormulonError) as caught:
                    wb.sheet(args[0]).pivot(args[1]).get_data_field_default_name(*args[2:])
                self.assertEqual(caught.exception.code, code)
            with self.assertRaises(NotFoundError):
                wb.sheet(0).pivot(f.pivot).get_data_field_default_name("0", 0)

    def test_projection_is_pending_after_a_cache_change(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            self.assertFalse(wb.sheet(0).pivot(f.pivot).is_projection_pending())
            wb.pivot_cache(f.cache).add_field_shared_item(0, "North")
            self.assertTrue(wb.sheet(0).pivot(f.pivot).is_projection_pending())
            wb.sheet(0).pivot(f.pivot).project()
            self.assertFalse(wb.sheet(0).pivot(f.pivot).is_projection_pending())

    def test_options_round_trip_through_save(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            changes = {
                "show_error": True,
                "error_caption": "N/A",
                "grid_drop_zones": True,
                "indent": 2,
                "alt_text": "Sales",
                "grand_total_caption": "All",
            }
            options = replace(wb.sheet(0).pivot(f.pivot).get_options(), **changes)
            wb.sheet(0).pivot(f.pivot).update_options(**changes)
            self.assertEqual(wb.sheet(0).pivot(f.pivot).get_options(), options)
            # Members left out keep their value; "" clears a string option, which then reads None.
            wb.sheet(0).pivot(f.pivot).update_options(indent=5, error_caption="")
            kept = wb.sheet(0).pivot(f.pivot).get_options()
            self.assertEqual((kept.indent, kept.error_caption, kept.alt_text), (5, None, "Sales"))
            self.assertEqual((kept.show_error, kept.grid_drop_zones, kept.grand_total_caption), (True, True, "All"))
            options = kept
            data = wb.save()
        with Workbook.load(data) as loaded:
            self.assertEqual(loaded.sheet(0).pivot(0).get_options(), options)

    def test_new_filter_types_and_high_bound_round_trip(self) -> None:
        self.assertEqual(len(PivotFilterKind), 68)
        self.assertEqual(PivotFilterKind.M12, 67)
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            wb.sheet(0).pivot(f.pivot).add_filter(
                PivotFilterInput(
                    field_index="Region",
                    type=PivotFilterKind.LABEL_BETWEEN,
                    value="A",
                    value_high="Z",
                )
            )
            spec = wb.sheet(0).pivot(f.pivot).list_filters()[0]
            self.assertEqual((spec.type, spec.value, spec.value_high), (PivotFilterKind.LABEL_BETWEEN, "A", "Z"))

    def test_derived_name_needs_an_in_range_field(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            with self.assertRaises(FormulonError):
                wb.sheet(0).pivot(f.pivot).add_data_field(PivotDataFieldInput(name="", field_index=99))
            self.assertEqual(len(wb.sheet(0).pivot(f.pivot).list_data_fields()), 1)


class PivotFieldBindingTests(unittest.TestCase):
    def _wide_pivot(self, wb: Workbook) -> int:
        cache = wb.add_pivot_cache().cache_id
        wb.pivot_cache(cache).set_worksheet_source(PivotWorksheetSourceInput(range="A1:D3", sheet="Sheet1"))
        for name in ("Region", "Product", "Sales", "Qty"):
            wb.pivot_cache(cache).add_field(name)
        for sales in (100, 200):
            rec = wb.pivot_cache(cache).add_record()
            for col, value in enumerate((0, 0, sales, 1)):
                wb.pivot_cache(cache).set_record_value(rec, col, value)
        return wb.sheet(0).add_pivot(PivotInput(name="PT", cache_id=cache, anchor=CellAddress(0, 6)))

    def test_field_add_binds_to_the_cache_column_by_name(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._wide_pivot(wb)
            self.assertEqual(
                wb.sheet(0).pivot(pivot).add_field(PivotFieldInput(field_index="Region", axis=PivotAxis.ROW)), 0
            )
            sales = wb.sheet(0).pivot(pivot).add_field(PivotFieldInput(field_index="Sales", axis=PivotAxis.VALUE))
            self.assertEqual(sales, 2)
            self.assertEqual(len(wb.sheet(0).pivot(pivot).list_fields()), 4)
            self.assertIsNone(wb.sheet(0).pivot(pivot).list_fields()[1].axis)
            # Adding a field again updates it instead of appending another.
            self.assertEqual(
                wb.sheet(0).pivot(pivot).add_field(PivotFieldInput(field_index="Sales", axis=PivotAxis.VALUE)), 2
            )
            self.assertEqual(len(wb.sheet(0).pivot(pivot).list_fields()), 4)

    def test_axis_none_takes_a_field_off_the_report(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._wide_pivot(wb)
            wb.sheet(0).pivot(pivot).add_field(PivotFieldInput(field_index="Region", axis=PivotAxis.ROW))
            sales = wb.sheet(0).pivot(pivot).add_field(PivotFieldInput(field_index="Sales", axis=PivotAxis.VALUE))
            wb.sheet(0).pivot(pivot).add_data_field(PivotDataFieldInput(name="Total", field_index=sales))
            wb.sheet(0).pivot(pivot).set_field_axis(sales, None)
            self.assertIsNone(wb.sheet(0).pivot(pivot).list_fields()[sales].axis)
            self.assertEqual(len(wb.sheet(0).pivot(pivot).list_data_fields()), 0)
            wb.sheet(0).pivot(pivot).set_field_axis(0, None)
            self.assertEqual(wb.sheet(0).pivot(pivot).get_row_field_order(), [])
            with self.assertRaises(FormulonError):
                wb.sheet(0).pivot(pivot).set_field_axis(0, 4)

    def test_save_drops_a_cache_no_pivot_uses(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._wide_pivot(wb)
            self.assertIn(b"pivotCacheDefinition1.xml", wb.save())
            wb.sheet(0).remove_pivot(pivot)
            self.assertEqual(len(wb.list_pivot_caches()), 1)
            self.assertNotIn(b"pivotCacheDefinition1.xml", wb.save())


if __name__ == "__main__":
    unittest.main()
