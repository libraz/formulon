"""Pivot surface and marshalling tests."""

from __future__ import annotations

import inspect
import io
import unittest
import zipfile
from dataclasses import replace
from typing import Any, Dict, List, Tuple

from formulon import (
    PIVOT_VALUES_FIELD,
    CellRect,
    FormulonError,
    PivotAggregation,
    PivotAxis,
    PivotCalendar,
    PivotDataFieldInput,
    PivotDatePeriod,
    PivotFieldInput,
    PivotFilterInput,
    PivotFilterType,
    PivotGrandTotals,
    PivotOptions,
    PivotReportLayout,
    PivotSortOrder,
    PivotStyleInput,
    PivotWorksheetSource,
    Sheet,
    StatusCode,
    ValueKind,
    Workbook,
)


class PivotTests(unittest.TestCase):
    def test_build_and_project_pivot(self) -> None:
        with Workbook.create_default() as wb:
            cache_id = wb.add_pivot_cache()
            wb.add_pivot_cache_field(cache_id, "Region")
            wb.add_pivot_cache_field(cache_id, "Amount")
            for region, amount in [("East", 10.0), ("East", 20.0), ("West", 30.0)]:
                rec = wb.add_pivot_cache_record(cache_id)
                wb.set_pivot_cache_record_text(cache_id, rec, 0, region)
                wb.set_pivot_cache_record_number(cache_id, rec, 1, amount)
            self.assertEqual(wb.get_pivot_cache_record_count(cache_id), 3)

            pivot = wb.sheet(0).add_pivot("Pivot1", cache_id, 0, 4)
            region_field = wb.sheet(0).add_pivot_field(pivot, PivotFieldInput(source_name="Region", axis=PivotAxis.ROW))
            amount_field = wb.sheet(0).add_pivot_field(
                pivot, PivotFieldInput(source_name="Amount", axis=PivotAxis.VALUE)
            )
            self.assertEqual(region_field, 0)
            wb.sheet(0).add_pivot_data_field(
                pivot,
                PivotDataFieldInput(
                    name="Sum of Amount",
                    field_index=amount_field,
                    aggregation=PivotAggregation.SUM,
                ),
            )

            layout = wb.sheet(0).get_pivot_layout(pivot)
            self.assertGreater(len(layout.cells), 0)
            numbers = [c.value.to_python() for c in layout.cells if c.value.kind == ValueKind.NUMBER]
            # The single grand total of 10 + 20 + 30 must appear.
            self.assertIn(60.0, numbers)

    def test_cache_shared_item_value_reads_each_kind(self) -> None:
        with Workbook.create_default() as wb:
            cache_id = wb.add_pivot_cache()
            wb.add_pivot_cache_field(cache_id, "Region")
            wb.add_pivot_cache_field_shared_item_text(cache_id, 0, "East")
            wb.add_pivot_cache_field_shared_item_number(cache_id, 0, 45292.0)
            wb.add_pivot_cache_field_shared_item_bool(cache_id, 0, value=True)
            wb.add_pivot_cache_field_shared_item_blank(cache_id, 0)
            wb.add_pivot_cache_field_shared_item_error(cache_id, 0, 1)
            kinds = [wb.get_pivot_cache_field_shared_item_value(cache_id, 0, i).kind for i in range(5)]
            self.assertEqual(
                kinds, [ValueKind.TEXT, ValueKind.NUMBER, ValueKind.BOOL, ValueKind.BLANK, ValueKind.ERROR]
            )
            self.assertEqual(wb.get_pivot_cache_field_shared_item_value(cache_id, 0, 0).text, "East")
            self.assertEqual(wb.get_pivot_cache_field_shared_item_value(cache_id, 0, 1).number, 45292.0)
            self.assertIs(wb.get_pivot_cache_field_shared_item_value(cache_id, 0, 2).boolean, True)
            self.assertEqual(wb.get_pivot_cache_field_shared_item_value(cache_id, 0, 4).error_code, 1)
            with self.assertRaises(FormulonError):
                wb.get_pivot_cache_field_shared_item_value(cache_id, 0, 5)

    def test_cache_date_shared_items(self) -> None:
        with Workbook.create_default() as wb:
            cache_id = wb.add_pivot_cache()
            wb.add_pivot_cache_field(cache_id, "When")
            wb.add_pivot_cache_field_shared_item_date(cache_id, 0, 45292.0)
            wb.add_pivot_cache_field_shared_item_number(cache_id, 0, 45292.0)
            wb.add_pivot_cache_field_shared_item_date(cache_id, 0, 45292.5)
            self.assertIs(wb.is_pivot_cache_field_shared_item_date(cache_id, 0, 0), True)
            self.assertIs(wb.is_pivot_cache_field_shared_item_date(cache_id, 0, 1), False)
            self.assertIs(wb.is_pivot_cache_field_shared_item_date(cache_id, 0, 2), True)
            value = wb.get_pivot_cache_field_shared_item_value(cache_id, 0, 0)
            self.assertEqual(value.kind, ValueKind.NUMBER)
            self.assertEqual(value.number, 45292.0)
            with self.assertRaises(FormulonError):
                wb.is_pivot_cache_field_shared_item_date(cache_id, 0, 3)
            for bad in (-1.0, float("nan")):
                with self.assertRaises(FormulonError) as caught:
                    wb.add_pivot_cache_field_shared_item_date(cache_id, 0, bad)
                self.assertEqual(caught.exception.code, 2)
            wb.add_pivot_cache_record(cache_id)
            wb.set_pivot_cache_record_date(cache_id, 0, 0, 45292.0)
            with self.assertRaises(FormulonError):
                wb.set_pivot_cache_record_date(cache_id, 0, 0, float("nan"))

    def test_cache_source_and_report_layout_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            cache = wb.add_pivot_cache()
            wb.set_pivot_cache_worksheet_source(cache, PivotWorksheetSource(ref="A1:B9", sheet="Sheet1"))
            self.assertEqual(wb.get_pivot_cache_worksheet_source(cache).sheet, "Sheet1")
            pivot = wb.sheet(0).add_pivot("Layout", cache, 0, 0)
            wb.sheet(0).set_pivot_report_layout(pivot, PivotReportLayout.TABULAR)
            self.assertEqual(wb.sheet(0).get_pivot_report_layout(pivot), PivotReportLayout.TABULAR)
            wb.set_pivot_cache_worksheet_source(cache, None)
            self.assertIsNone(wb.get_pivot_cache_worksheet_source(cache))


class _PivotFixture:
    """Region / Product carry shared items their records index; Amount and
    Date are plain numbers. Only Region is laid out."""

    def __init__(self, wb: Workbook) -> None:
        self.wb = wb
        wb.set_excel_profile_id("win-365-ja_JP")
        self.cache = wb.add_pivot_cache()
        wb.set_pivot_cache_worksheet_source(self.cache, PivotWorksheetSource(ref="A1:D5", sheet="Sheet1"))
        for name in ("Region", "Product", "Amount", "Date"):
            wb.add_pivot_cache_field(self.cache, name)
        for item in ("East", "West"):
            wb.add_pivot_cache_field_shared_item_text(self.cache, 0, item)
        for item in ("Apple", "Pear"):
            wb.add_pivot_cache_field_shared_item_text(self.cache, 1, item)
        for values in ((0, 0, 10, 45292), (1, 0, 30, 45323), (0, 1, 5, 45658), (1, 1, 7, 45689)):
            rec = wb.add_pivot_cache_record(self.cache)
            for field, value in enumerate(values):
                wb.set_pivot_cache_record_number(self.cache, rec, field, value)
        self.pivot = wb.sheet(0).add_pivot("Pivot1", self.cache, 0, 6)
        self.region = wb.sheet(0).add_pivot_field(self.pivot, PivotFieldInput(source_name="Region", axis=PivotAxis.ROW))
        wb.sheet(0).add_pivot_field_item(self.pivot, self.region, "East", visible=True)
        wb.sheet(0).add_pivot_field_item(self.pivot, self.region, "West", visible=True)
        self.product = wb.sheet(0).add_pivot_field(
            self.pivot, PivotFieldInput(source_name="Product", axis=PivotAxis.VALUE)
        )
        wb.sheet(0).add_pivot_field_item(self.pivot, self.product, "Apple", visible=True)
        wb.sheet(0).add_pivot_field_item(self.pivot, self.product, "Pear", visible=True)
        self.amount = wb.sheet(0).add_pivot_field(
            self.pivot, PivotFieldInput(source_name="Amount", axis=PivotAxis.VALUE)
        )
        self.date = wb.sheet(0).add_pivot_field(self.pivot, PivotFieldInput(source_name="Date", axis=PivotAxis.VALUE))
        wb.sheet(0).set_pivot_row_field_order(self.pivot, [self.region])
        wb.sheet(0).add_pivot_data_field(
            self.pivot, PivotDataFieldInput(name="Sum of Amount", field_index=self.amount, aggregation=0)
        )

    def grid(self, pivot: int = 0) -> List[str]:
        """The projected grid as one ``|``-joined string per sheet row."""
        return _grid(self.wb, pivot)


def _grid(wb: Workbook, pivot: int = 0) -> List[str]:
    """The projected grid of ``pivot`` on sheet 0, one ``|``-joined string per row."""
    layout = wb.sheet(0).get_pivot_layout(pivot)
    rows: Dict[int, List[str]] = {}
    for cell in layout.cells:
        v = cell.value
        if v.kind == ValueKind.NUMBER:
            text = "%.15g" % v.number
        elif v.kind == ValueKind.TEXT:
            text = v.text
        elif v.kind == ValueKind.ERROR:
            text = f"#{v.error_code}"
        else:
            text = ""
        assert layout.rect is not None
        rows.setdefault(cell.row, [""] * (layout.rect.last_col - layout.rect.first_col + 1))[
            cell.col - layout.rect.first_col
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


_GREATER_THAN_20 = PivotFilterInput(
    axis=PivotAxis.ROW, field_name="Region", type=PivotFilterType.VALUE_GREATER_THAN, value_kind=1, value_double=20.0
)


_BETWEEN_10_20 = PivotFilterInput(
    axis=PivotAxis.ROW,
    field_name="Region",
    type=PivotFilterType.VALUE_BETWEEN,
    value_kind=1,
    value_double=10.0,
    value_high_kind=1,
    value_high_double=20.0,
)


_PIVOT_READERS = {
    "get_pivot_count",
    "get_pivot_layout",
    "get_pivot_cache_field_count",
    "get_pivot_cache_field_name",
    "get_pivot_cache_field_shared_item_count",
    "get_pivot_cache_field_shared_item_value",
    "is_pivot_cache_field_shared_item_date",
    "get_pivot_cache_record_count",
    "get_pivot_field_count",
    "get_pivot_data_field_count",
    "get_pivot_filter_count",
    "get_pivot_filter",
    "get_pivot_cache_id",
    "get_pivot_name",
    "get_pivot_grand_totals",
    "get_pivot_style_info",
    "get_pivot_row_field_order",
    "get_pivot_column_field_order",
    "get_pivot_field",
    "get_pivot_field_item_count",
    "get_pivot_field_item",
    "get_pivot_options",
    "is_pivot_projection_pending",
    "get_pivot_data_field_default_name",
    "get_pivot_cache_options",
    "get_pivot_cache_field_formula",
    "get_pivot_cache_calculated_item_count",
    "get_pivot_cache_calculated_item",
    "get_pivot_data_field",
    "get_pivot_cache_worksheet_source",
    "get_pivot_report_layout",
    "list_pivot_cache_calculated_items",
    "list_pivot_data_fields",
    "list_pivot_fields",
    "list_pivot_filters",
}


def _first_number_cell(f: _PivotFixture) -> Tuple[int, int]:
    """Position of a data cell of the fixture pivot's layout."""
    for cell in f.wb.sheet(0).get_pivot_layout(f.pivot).cells:
        if cell.value.kind == ValueKind.NUMBER:
            return cell.row, cell.col
    raise AssertionError("pivot layout has no number cell")


def _two_levels(f: _PivotFixture) -> None:
    f.wb.sheet(0).set_pivot_row_field_order(f.pivot, [f.region, f.product])


def _date_rows(f: _PivotFixture) -> None:
    f.wb.sheet(0).set_pivot_row_field_order(f.pivot, [f.date])


def _hide_east(f: _PivotFixture) -> None:
    f.wb.sheet(0).set_pivot_field_item_visible(f.pivot, f.region, 0, False)


def _point_at_new_item(f: _PivotFixture) -> None:
    f.wb.set_pivot_cache_record_number(f.cache, 0, 0, 2)


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

        def cache_create(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 77)
            eq(len(f.wb.list_pivot_caches()), 2)

        def cache_remove_setup(f: _PivotFixture) -> None:
            f.spare = f.wb.add_pivot_cache()  # type: ignore[attr-defined]

        def cache_remove(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.list_pivot_caches(), [f.cache])

        def cache_source(t: Any, f: _PivotFixture, r: None) -> None:
            source = f.wb.get_pivot_cache_worksheet_source(f.cache)
            eq((source.ref, source.sheet), ("B2:E6", "Data"))
            eq(f.grid(), _BASE_GRID)

        def cache_field_add(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 4)
            eq(f.wb.get_pivot_cache_field_name(f.cache, 4), "Extra")

        def cache_field_clear(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.get_pivot_cache_field_count(f.cache), 0)
            # The data field now names a cache field that no longer exists.
            with self.assertRaises(FormulonError):
                f.wb.sheet(0).get_pivot_layout(f.pivot)

        def record_add(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 4)
            eq(f.grid(), ["行ラベル|Sum of Amount", "East|15", "West|37", "(空白)|", "総計|52"])

        def pivot_create(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 1)
            layout = f.wb.sheet(0).get_pivot_layout(1)
            eq((layout.rect.first_row, layout.rect.first_col), (20, 1))
            eq(f.grid(0), _BASE_GRID)

        def pivot_remove(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.sheet(0).get_pivot_count(), 0)
            with self.assertRaises(FormulonError):
                f.wb.sheet(0).get_pivot_layout(0)

        def set_name(t: Any, f: _PivotFixture, r: None) -> None:
            # A pivot's name is only observable in the part the writer emits.
            part = _package_part(f.wb.save(), "xl/pivotTables/pivotTable1.xml")
            self.assertRegex(part, r'<pivotTableDefinition [^>]*name="Renamed"')

        def set_style_info(t: Any, f: _PivotFixture, r: None) -> None:
            style = f.wb.sheet(0).get_pivot_style_info(f.pivot)
            eq(style.name, "PivotStyleMedium9")
            eq(
                (style.show_row_headers, style.show_col_headers, style.show_row_stripes, style.show_col_stripes),
                (False, True, True, False),
            )
            self.assertTrue(style.show_last_column)
            eq(f.grid(), _BASE_GRID)

        def set_anchor(t: Any, f: _PivotFixture, r: None) -> None:
            layout = f.wb.sheet(0).get_pivot_layout(f.pivot)
            eq((layout.rect.first_row, layout.rect.first_col), (10, 2))
            eq(f.grid(), _BASE_GRID)

        def set_layout(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.sheet(0).get_pivot_report_layout(f.pivot), PivotReportLayout.TABULAR)
            eq(f.grid(), ["Region|Sum of Amount", "East|15", "West|37", "総計|52"])

        def field_add(t: Any, f: _PivotFixture, r: int) -> None:
            # Product is a cache column the fixture already added, so it is updated in place.
            eq(r, f.product)
            eq(f.grid(), ["Fruit|(すべて)", "|", *_BASE_GRID])

        def field_clear(t: Any, f: _PivotFixture, r: None) -> None:
            eq(len(f.wb.sheet(0).list_pivot_fields(f.pivot)), 0)
            # The row order empties with the fields; East's hidden flag goes too.
            eq(f.grid(), ["行ラベル|Sum of Amount", "|52"])

        def add_item_at_setup(f: _PivotFixture) -> None:
            # Only the index-addressed form can name the blank item.
            f.wb.add_pivot_cache_field_shared_item_blank(f.cache, 0)
            f.wb.set_pivot_cache_record_blank(f.cache, 0, 0)

        def clear_subtotal_setup(f: _PivotFixture) -> None:
            _two_levels(f)
            f.wb.sheet(0).add_pivot_field_subtotal_fn(f.pivot, f.region, PivotAggregation.MAX)

        def clear_date_group_setup(f: _PivotFixture) -> None:
            _date_rows(f)
            f.wb.sheet(0).set_pivot_field_date_group(f.pivot, f.date, PivotDatePeriod.YEARS, PivotCalendar.GREGORIAN)

        def number_format(t: Any, f: _PivotFixture, r: None) -> None:
            rejects(lambda: f.wb.sheet(0).set_pivot_field_number_format(f.pivot, 99, "4"), IDX)
            rejects(lambda: f.wb.sheet(0).set_pivot_field_number_format(f.pivot, f.amount, "0.00"))
            # A field's numFmtId is only observable in the part the writer emits.
            part = _package_part(f.wb.save(), "xl/pivotTables/pivotTable1.xml")
            self.assertRegex(part, r'<pivotField [^>]*numFmtId="4"')
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
                c.number_format
                for c in f.wb.sheet(0).get_pivot_layout(f.pivot).cells
                if c.value.kind == ValueKind.NUMBER
            ]
            eq(formats, ["2", "2", "2"])

        def filter_remove_setup(f: _PivotFixture) -> None:
            # Each filter alone keeps a different region, so removing the
            # wrong index leaves the wrong one.
            f.wb.sheet(0).add_pivot_filter(f.pivot, _GREATER_THAN_20)
            f.wb.sheet(0).add_pivot_filter(f.pivot, _BETWEEN_10_20)

        def type_rejected(call: Any) -> None:
            with self.assertRaises(TypeError):
                call()

        def calc_field_setup(f: _PivotFixture) -> None:
            f.wb.add_pivot_cache_calculated_field(f.cache, "Double", "Amount*2")

        def calc_item_setup(f: _PivotFixture) -> None:
            f.wb.add_pivot_cache_calculated_item(f.cache, f.region, "Both", "East+West")

        def calc_two_items_setup(f: _PivotFixture) -> None:
            calc_item_setup(f)
            f.wb.add_pivot_cache_calculated_item(f.cache, f.region, "Diff", "East-West")

        def pending_setup(f: _PivotFixture) -> None:
            f.wb.add_pivot_cache_field_shared_item_text(f.cache, 0, "North")

        def options_set(t: Any, f: _PivotFixture, r: None) -> None:
            o = f.wb.sheet(0).get_pivot_options(f.pivot)
            eq((o.show_error, o.error_caption), (True, "N/A"))
            eq((o.show_missing, o.missing_caption, o.indent, o.page_wrap), (True, "-", 3, 2))
            eq((o.alt_text, o.alt_text_summary, o.grand_total_caption), ("Title", "", "Sum"))
            rejects(lambda: f.wb.sheet(0).set_pivot_options(9, o), IDX)
            rejects(lambda: f.wb.sheet(0).set_pivot_options(f.pivot, replace(o, indent=127)))
            rejects(lambda: f.wb.sheet(0).set_pivot_options(f.pivot, replace(o, page_wrap=256)))
            rejects(lambda: f.wb.sheet(0).set_pivot_options(f.pivot, replace(o, error_caption="x" * 256)))
            type_rejected(lambda: f.wb.sheet(0).set_pivot_options(f.pivot, replace(o, indent="3")))
            type_rejected(lambda: f.wb.sheet(0).set_pivot_options(f.pivot, replace(o, error_caption=5)))
            type_rejected(lambda: f.wb.sheet(0).set_pivot_options(f.pivot, {}))
            eq(f.wb.sheet(0).get_pivot_options(f.pivot), o)

        def cache_options_set(t: Any, f: _PivotFixture, r: None) -> None:
            o = f.wb.get_pivot_cache_options(f.cache)
            eq((o.save_data, o.missing_items_limit, o.refresh_on_load), (False, 32500, True))
            rejects(lambda: f.wb.set_pivot_cache_options(9999, o))
            rejects(lambda: f.wb.set_pivot_cache_options(f.cache, replace(o, missing_items_limit=-2)))
            rejects(lambda: f.wb.set_pivot_cache_options(f.cache, replace(o, missing_items_limit=1048577)))
            type_rejected(lambda: f.wb.set_pivot_cache_options(f.cache, replace(o, missing_items_limit="1")))
            eq(f.wb.get_pivot_cache_options(f.cache), o)

        def calc_field_add(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 4)
            eq(f.wb.get_pivot_cache_field_name(f.cache, 4), "Double")
            eq(f.wb.get_pivot_cache_field_formula(f.cache, 4), "Amount*2")
            eq(f.wb.get_pivot_cache_field_formula(f.cache, 0), None)
            rejects(lambda: f.wb.add_pivot_cache_calculated_field(f.cache, "Amount", "Amount*2"))
            rejects(lambda: f.wb.add_pivot_cache_calculated_field(f.cache, "Other", ""))
            rejects(lambda: f.wb.add_pivot_cache_calculated_field(f.cache, "Other", "NoSuchField+1"))
            rejects(lambda: f.wb.add_pivot_cache_calculated_field(9999, "Other", "Amount*2"))
            type_rejected(lambda: f.wb.add_pivot_cache_calculated_field(f.cache, 5, "Amount*2"))
            type_rejected(lambda: f.wb.add_pivot_cache_calculated_field(f.cache, "Other", None))
            eq(f.wb.get_pivot_cache_field_count(f.cache), 5)

        def calc_field_set_formula(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.get_pivot_cache_field_formula(f.cache, 4), "Amount*3")
            rejects(lambda: f.wb.set_pivot_cache_calculated_field_formula(f.cache, 0, "Amount*3"))
            rejects(lambda: f.wb.set_pivot_cache_calculated_field_formula(f.cache, 4, ""))
            type_rejected(lambda: f.wb.set_pivot_cache_calculated_field_formula(f.cache, 4, 3))
            eq(f.wb.get_pivot_cache_field_formula(f.cache, 4), "Amount*3")

        def calc_field_remove(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.get_pivot_cache_field_count(f.cache), 4)
            rejects(lambda: f.wb.remove_pivot_cache_calculated_field(f.cache, 0))

        def calc_item_add(t: Any, f: _PivotFixture, r: None) -> None:
            eq(len(f.wb.list_pivot_cache_calculated_items(f.cache)), 1)
            item = f.wb.list_pivot_cache_calculated_items(f.cache)[0]
            eq((item.field_index, item.name, item.scoped), (f.region, "Both", False))
            eq(item.formula, "Region[East]+Region[West]")
            with self.assertRaises(IndexError):
                f.wb.list_pivot_cache_calculated_items(f.cache)[5]
            rejects(lambda: f.wb.add_pivot_cache_calculated_item(f.cache, f.region, "East", "East+West"))
            rejects(lambda: f.wb.add_pivot_cache_calculated_item(f.cache, f.region, "Other", ""))
            rejects(lambda: f.wb.add_pivot_cache_calculated_item(f.cache, 99, "Other", "East+West"), IDX)
            type_rejected(lambda: f.wb.add_pivot_cache_calculated_item(f.cache, f.region, None, "East+West"))
            eq(len(f.wb.list_pivot_cache_calculated_items(f.cache)), 1)

        def calc_item_set_formula(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.list_pivot_cache_calculated_items(f.cache)[0].formula, "Region[East]*2")
            rejects(lambda: f.wb.set_pivot_cache_calculated_item_formula(f.cache, 5, "East*2"), IDX)
            type_rejected(lambda: f.wb.set_pivot_cache_calculated_item_formula(f.cache, 0, 2))

        def calc_item_move(t: Any, f: _PivotFixture, r: None) -> None:
            eq([f.wb.list_pivot_cache_calculated_items(f.cache)[i].name for i in (0, 1)], ["Diff", "Both"])
            rejects(lambda: f.wb.move_pivot_cache_calculated_item(f.cache, 0, 5), IDX)
            type_rejected(lambda: f.wb.move_pivot_cache_calculated_item(f.cache, 0, "1"))

        def calc_item_remove(t: Any, f: _PivotFixture, r: None) -> None:
            eq(len(f.wb.list_pivot_cache_calculated_items(f.cache)), 0)
            rejects(lambda: f.wb.remove_pivot_cache_calculated_item(f.cache, 0), IDX)

        def custom_name(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.sheet(0).list_pivot_fields(f.pivot)[f.region].custom_name, "Area")
            rejects(lambda: f.wb.sheet(0).set_pivot_field_custom_name(f.pivot, 99, "x"), IDX)
            rejects(lambda: f.wb.sheet(0).set_pivot_field_custom_name(f.pivot, f.region, "Product"))
            rejects(lambda: f.wb.sheet(0).set_pivot_field_custom_name(f.pivot, f.region, "x" * 256))
            type_rejected(lambda: f.wb.sheet(0).set_pivot_field_custom_name(f.pivot, f.region, 5))
            eq(f.wb.sheet(0).list_pivot_fields(f.pivot)[f.region].custom_name, "Area")
            f.wb.sheet(0).set_pivot_field_custom_name(f.pivot, f.region, None)
            eq(f.wb.sheet(0).list_pivot_fields(f.pivot)[f.region].custom_name, None)
            f.wb.sheet(0).set_pivot_field_custom_name(f.pivot, f.region, "Region")
            f.wb.sheet(0).set_pivot_field_custom_name(f.pivot, f.region, "")
            eq(f.wb.sheet(0).list_pivot_fields(f.pivot)[f.region].custom_name, None)

        def project(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.sheet(0).is_pivot_projection_pending(f.pivot), False)
            rejects(lambda: f.wb.sheet(0).project_pivot(9), IDX)
            rejects(lambda: f.wb.sheet(0).is_pivot_projection_pending(9), IDX)

        W = PivotAggregation
        return {
            "set_pivot_options": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_options(
                    f.pivot,
                    replace(
                        f.wb.sheet(0).get_pivot_options(f.pivot),
                        show_error=True,
                        error_caption="N/A",
                        show_missing=True,
                        missing_caption="-",
                        indent=3,
                        page_wrap=2,
                        alt_text="Title",
                        alt_text_summary="",
                        grand_total_caption="Sum",
                    ),
                ),
                None,
                options_set,
            ),
            "set_pivot_cache_options": (
                None,
                lambda f: f.wb.set_pivot_cache_options(
                    f.cache,
                    replace(
                        f.wb.get_pivot_cache_options(f.cache),
                        save_data=False,
                        missing_items_limit=32500,
                        refresh_on_load=True,
                    ),
                ),
                None,
                cache_options_set,
            ),
            "add_pivot_cache_calculated_field": (
                None,
                lambda f: f.wb.add_pivot_cache_calculated_field(f.cache, "Double", "Amount*2"),
                None,
                calc_field_add,
            ),
            "set_pivot_cache_calculated_field_formula": (
                calc_field_setup,
                lambda f: f.wb.set_pivot_cache_calculated_field_formula(f.cache, 4, "Amount*3"),
                None,
                calc_field_set_formula,
            ),
            "remove_pivot_cache_calculated_field": (
                calc_field_setup,
                lambda f: f.wb.remove_pivot_cache_calculated_field(f.cache, 4),
                None,
                calc_field_remove,
            ),
            "add_pivot_cache_calculated_item": (
                None,
                lambda f: f.wb.add_pivot_cache_calculated_item(f.cache, f.region, "Both", "East+West"),
                None,
                calc_item_add,
            ),
            "set_pivot_cache_calculated_item_formula": (
                calc_item_setup,
                lambda f: f.wb.set_pivot_cache_calculated_item_formula(f.cache, 0, "East*2"),
                None,
                calc_item_set_formula,
            ),
            "move_pivot_cache_calculated_item": (
                calc_two_items_setup,
                lambda f: f.wb.move_pivot_cache_calculated_item(f.cache, 0, 1),
                None,
                calc_item_move,
            ),
            "remove_pivot_cache_calculated_item": (
                calc_item_setup,
                lambda f: f.wb.remove_pivot_cache_calculated_item(f.cache, 0),
                None,
                calc_item_remove,
            ),
            "set_pivot_field_custom_name": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_field_custom_name(f.pivot, f.region, "Area"),
                None,
                custom_name,
            ),
            "set_pivot_field_show_all": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_field_show_all(f.pivot, f.region, True),
                None,
                lambda t, f, r: (
                    eq(f.wb.sheet(0).list_pivot_fields(f.pivot)[f.region].show_all, True),
                    eq(f.wb.sheet(0).list_pivot_fields(f.pivot)[f.product].show_all, False),
                    rejects(lambda: f.wb.sheet(0).set_pivot_field_show_all(f.pivot, 99, True), IDX),
                    type_rejected(lambda: f.wb.sheet(0).set_pivot_field_show_all(f.pivot, "0", False)),
                    eq(f.wb.sheet(0).list_pivot_fields(f.pivot)[f.region].show_all, True),
                ),
            ),
            "project_pivot": (
                pending_setup,
                lambda f: f.wb.sheet(0).project_pivot(f.pivot),
                None,
                project,
            ),
            "add_pivot_cache": (None, lambda f: f.wb.add_pivot_cache(77), None, cache_create),
            "list_pivot_caches": (None, lambda f: f.wb.list_pivot_caches(), None, lambda t, f, r: eq(r, [f.cache])),
            "remove_pivot_cache": (cache_remove_setup, lambda f: f.wb.remove_pivot_cache(f.spare), None, cache_remove),
            "set_pivot_cache_worksheet_source": (
                None,
                lambda f: f.wb.set_pivot_cache_worksheet_source(
                    f.cache, PivotWorksheetSource(ref="B2:E6", sheet="Data")
                ),
                None,
                cache_source,
            ),
            "add_pivot_cache_field": (
                None,
                lambda f: f.wb.add_pivot_cache_field(f.cache, "Extra"),
                None,
                cache_field_add,
            ),
            "clear_pivot_cache_fields": (
                None,
                lambda f: f.wb.clear_pivot_cache_fields(f.cache),
                None,
                cache_field_clear,
            ),
            "add_pivot_cache_field_shared_item_number": (
                None,
                lambda f: f.wb.add_pivot_cache_field_shared_item_number(f.cache, 0, 42.5),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "42.5|10", "East|5", "West|37", "総計|52"]),
            ),
            "add_pivot_cache_field_shared_item_date": (
                None,
                lambda f: f.wb.add_pivot_cache_field_shared_item_date(f.cache, 0, 45292.0),
                None,
                lambda t, f, r: eq(f.wb.is_pivot_cache_field_shared_item_date(f.cache, 0, 2), True),
            ),
            "add_pivot_cache_field_shared_item_text": (
                None,
                lambda f: f.wb.add_pivot_cache_field_shared_item_text(f.cache, 0, "North"),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "North|10", "West|37", "総計|52"]),
            ),
            "add_pivot_cache_field_shared_item_bool": (
                None,
                lambda f: f.wb.add_pivot_cache_field_shared_item_bool(f.cache, 0, value=True),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "TRUE|10", "総計|52"]),
            ),
            "add_pivot_cache_field_shared_item_blank": (
                None,
                lambda f: f.wb.add_pivot_cache_field_shared_item_blank(f.cache, 0),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "(空白)|10", "総計|52"]),
            ),
            "add_pivot_cache_field_shared_item_error": (
                None,
                lambda f: f.wb.add_pivot_cache_field_shared_item_error(f.cache, 0, 6),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "#N/A|10", "総計|52"]),
            ),
            "clear_pivot_cache_field_shared_items": (
                None,
                lambda f: f.wb.clear_pivot_cache_field_shared_items(f.cache, 0),
                None,
                # With no shared items the records' indices render as themselves.
                grid_is(["行ラベル|Sum of Amount", "0|15", "1|37", "総計|52"]),
            ),
            "add_pivot_cache_record": (None, lambda f: f.wb.add_pivot_cache_record(f.cache), None, record_add),
            "clear_pivot_cache_records": (
                None,
                lambda f: f.wb.clear_pivot_cache_records(f.cache),
                None,
                grid_is(["行ラベル|Sum of Amount", "総計|"]),
            ),
            "set_pivot_cache_record_number": (
                None,
                lambda f: f.wb.set_pivot_cache_record_number(f.cache, 0, 2, 1000),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|1005", "West|37", "総計|1042"]),
            ),
            "set_pivot_cache_record_date": (
                None,
                lambda f: f.wb.set_pivot_cache_record_date(f.cache, 0, 2, 45292.0),
                None,
                lambda t, f, r: eq(r, None),
            ),
            "set_pivot_cache_record_text": (
                None,
                lambda f: f.wb.set_pivot_cache_record_text(f.cache, 0, 0, "West"),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|47", "総計|52"]),
            ),
            "set_pivot_cache_record_bool": (
                None,
                lambda f: f.wb.set_pivot_cache_record_bool(f.cache, 0, 2, True),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "総計|42"]),
            ),
            "set_pivot_cache_record_blank": (
                None,
                lambda f: f.wb.set_pivot_cache_record_blank(f.cache, 0, 2),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "総計|42"]),
            ),
            "set_pivot_cache_record_error": (
                None,
                lambda f: f.wb.set_pivot_cache_record_error(f.cache, 0, 2, 1),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|#1", "West|37", "総計|#1"]),
            ),
            "add_pivot": (None, lambda f: f.wb.sheet(0).add_pivot("Pivot2", f.cache, 20, 1), None, pivot_create),
            "remove_pivot": (None, lambda f: f.wb.sheet(0).remove_pivot(f.pivot), None, pivot_remove),
            "set_pivot_name": (None, lambda f: f.wb.sheet(0).set_pivot_name(f.pivot, "Renamed"), None, set_name),
            "set_pivot_anchor": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_anchor(f.pivot, CellRect(10, 2, 13, 3)),
                None,
                set_anchor,
            ),
            "set_pivot_grand_totals": (
                lambda f: f.wb.sheet(0).set_pivot_column_field_order(f.pivot, [f.product]),
                # Rows off drops the per-row total column; columns on keeps
                # the bottom total row, so a swapped pair shows the opposite.
                lambda f: f.wb.sheet(0).set_pivot_grand_totals(f.pivot, rows=False, columns=True),
                None,
                grid_is(["Sum of Amount|列ラベル|", "行ラベル|Apple|Pear", "East|10|5", "West|30|7", "総計|40|12"]),
            ),
            "set_pivot_style_info": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_style_info(
                    f.pivot,
                    PivotStyleInput(
                        "PivotStyleMedium9",
                        show_col_headers=True,
                        show_row_stripes=True,
                        show_last_column=True,
                    ),
                ),
                None,
                set_style_info,
            ),
            "set_pivot_report_layout": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_report_layout(f.pivot, PivotReportLayout.TABULAR),
                None,
                set_layout,
            ),
            "add_pivot_field": (
                None,
                lambda f: f.wb.sheet(0).add_pivot_field(
                    f.pivot, PivotFieldInput(source_name="Product", custom_name="Fruit", axis=PivotAxis.PAGE)
                ),
                None,
                field_add,
            ),
            "clear_pivot_fields": (_hide_east, lambda f: f.wb.sheet(0).clear_pivot_fields(f.pivot), None, field_clear),
            "set_pivot_field_axis": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_field_axis(f.pivot, f.product, PivotAxis.PAGE),
                None,
                grid_is(["Product|(すべて)", "|", *_BASE_GRID]),
            ),
            "set_pivot_field_sort": (
                # East now outsums West, so sorting ascending by the data
                # field differs from both label orders.
                lambda f: f.wb.set_pivot_cache_record_number(f.cache, 0, 2, 40),
                lambda f: f.wb.sheet(0).set_pivot_field_sort(
                    f.pivot, f.region, PivotSortOrder.ASCENDING, "Sum of Amount"
                ),
                None,
                grid_is(["行ラベル|Sum of Amount", "West|37", "East|45", "総計|82"]),
            ),
            "set_pivot_field_subtotal_top": (
                _two_levels,
                lambda f: f.wb.sheet(0).set_pivot_field_subtotal_top(f.pivot, f.region, False),
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
            "add_pivot_field_item": (
                None,
                lambda f: f.wb.sheet(0).add_pivot_field_item(f.pivot, f.region, "East", visible=False),
                None,
                grid_is(_WEST_ONLY),
            ),
            "add_pivot_field_shared_item": (
                add_item_at_setup,
                lambda f: f.wb.sheet(0).add_pivot_field_shared_item(f.pivot, f.region, 2, visible=False),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "総計|42"]),
            ),
            "clear_pivot_field_items": (
                _hide_east,
                lambda f: f.wb.sheet(0).clear_pivot_field_items(f.pivot, f.region),
                None,
                grid_is(_BASE_GRID),
            ),
            "set_pivot_field_item_visible": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_field_item_visible(f.pivot, f.region, 1, False),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|15", "総計|15"]),
            ),
            "add_pivot_field_subtotal_fn": (
                _two_levels,
                lambda f: f.wb.sheet(0).add_pivot_field_subtotal_fn(f.pivot, f.region, W.MAX),
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
            "clear_pivot_field_subtotal_fns": (
                clear_subtotal_setup,
                lambda f: f.wb.sheet(0).clear_pivot_field_subtotal_fns(f.pivot, f.region),
                None,
                grid_is(_TWO_LEVEL_GRID),
            ),
            "set_pivot_field_date_group": (
                _date_rows,
                lambda f: f.wb.sheet(0).set_pivot_field_date_group(
                    f.pivot, f.date, PivotDatePeriod.YEARS, PivotCalendar.JAPANESE
                ),
                None,
                grid_is(["行ラベル|Sum of Amount", "令和6年|40", "令和7年|12", "総計|52"]),
            ),
            "clear_pivot_field_date_group": (
                clear_date_group_setup,
                lambda f: f.wb.sheet(0).clear_pivot_field_date_group(f.pivot, f.date),
                None,
                grid_is(["行ラベル|Sum of Amount", "45292|10", "45323|30", "45658|5", "45689|7", "総計|52"]),
            ),
            "set_pivot_field_number_format": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_field_number_format(f.pivot, f.amount, "4"),
                None,
                number_format,
            ),
            "set_pivot_row_field_order": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_row_field_order(f.pivot, [f.product]),
                None,
                grid_is(["行ラベル|Sum of Amount", "Apple|40", "Pear|12", "総計|52"]),
            ),
            "set_pivot_column_field_order": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_column_field_order(f.pivot, [f.product]),
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
            "add_pivot_data_field": (
                None,
                lambda f: f.wb.sheet(0).add_pivot_data_field(
                    f.pivot, PivotDataFieldInput(name="Max of Amount", field_index=f.amount, aggregation=W.MAX)
                ),
                None,
                data_field_add,
            ),
            "clear_pivot_data_fields": (
                None,
                lambda f: f.wb.sheet(0).clear_pivot_data_fields(f.pivot),
                None,
                grid_is(["行ラベル", "East", "West", "総計"]),
            ),
            "set_pivot_data_field": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_data_field(
                    f.pivot,
                    0,
                    PivotDataFieldInput(
                        name="Count of Amount", field_index=f.amount, aggregation=W.COUNT, number_format="2"
                    ),
                ),
                None,
                data_field_set,
            ),
            "add_pivot_filter": (
                None,
                lambda f: f.wb.sheet(0).add_pivot_filter(f.pivot, _GREATER_THAN_20),
                None,
                grid_is(_WEST_ONLY),
            ),
            "clear_pivot_filters": (
                lambda f: f.wb.sheet(0).add_pivot_filter(f.pivot, _GREATER_THAN_20),
                lambda f: f.wb.sheet(0).clear_pivot_filters(f.pivot),
                None,
                grid_is(_BASE_GRID),
            ),
            "remove_pivot_filter": (
                filter_remove_setup,
                lambda f: f.wb.sheet(0).remove_pivot_filter(f.pivot, 1),
                None,
                grid_is(_WEST_ONLY),
            ),
            "set_pivot_field_default_subtotal": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_field_default_subtotal(f.pivot, f.region, False),
                None,
                lambda t, f, r: (
                    eq(f.wb.sheet(0).list_pivot_fields(f.pivot)[f.region].default_subtotal, False),
                    rejects(lambda: f.wb.sheet(0).set_pivot_field_default_subtotal(f.pivot, 99, True), IDX),
                ),
            ),
            "set_pivot_field_insert_blank_row": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_field_insert_blank_row(f.pivot, f.region, True),
                None,
                lambda t, f, r: (
                    eq(f.wb.sheet(0).list_pivot_fields(f.pivot)[f.region].insert_blank_row, True),
                    rejects(lambda: f.wb.sheet(0).set_pivot_field_insert_blank_row(f.pivot, 99, True), IDX),
                ),
            ),
            "set_pivot_field_fill_down_labels": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_field_fill_down_labels(f.pivot, f.region, True),
                None,
                lambda t, f, r: (
                    eq(f.wb.sheet(0).list_pivot_fields(f.pivot)[f.region].fill_down_labels, True),
                    rejects(lambda: f.wb.sheet(0).set_pivot_field_fill_down_labels(f.pivot, 99, True), IDX),
                ),
            ),
            "set_pivot_field_show_detail": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_field_show_detail(f.pivot, f.region, False),
                None,
                lambda t, f, r: (
                    eq(
                        [f.wb.sheet(0).get_pivot_field_item(f.pivot, f.region, i).show_detail for i in (0, 1)],
                        [False, False],
                    ),
                    rejects(lambda: f.wb.sheet(0).set_pivot_field_show_detail(f.pivot, 99, False), IDX),
                ),
            ),
            "set_pivot_field_item_show_detail": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_field_item_show_detail(f.pivot, f.region, 1, False),
                None,
                lambda t, f, r: (
                    eq(
                        [f.wb.sheet(0).get_pivot_field_item(f.pivot, f.region, i).show_detail for i in (0, 1)],
                        [True, False],
                    ),
                    rejects(lambda: f.wb.sheet(0).set_pivot_field_item_show_detail(f.pivot, f.region, 9, False), IDX),
                ),
            ),
            "set_pivot_field_item_caption": (
                None,
                lambda f: f.wb.sheet(0).set_pivot_field_item_caption(f.pivot, f.region, 0, "Eastern"),
                None,
                lambda t, f, r: (
                    eq(f.wb.sheet(0).get_pivot_field_item(f.pivot, f.region, 0).caption, "Eastern"),
                    eq(f.wb.sheet(0).get_pivot_field_item(f.pivot, f.region, 1).caption, None),
                    f.wb.sheet(0).set_pivot_field_item_caption(f.pivot, f.region, 0, None),
                    eq(f.wb.sheet(0).get_pivot_field_item(f.pivot, f.region, 0).caption, None),
                    rejects(lambda: f.wb.sheet(0).set_pivot_field_item_caption(f.pivot, f.region, 9, "x"), IDX),
                    rejects(lambda: f.wb.sheet(0).set_pivot_field_item_caption(f.pivot, f.region, 0, "West")),
                ),
            ),
            "show_pivot_details": (
                None,
                lambda f: f.wb.sheet(0).show_pivot_details(f.pivot, *_first_number_cell(f)),
                None,
                lambda t, f, r: (
                    eq(r.index, 0),
                    eq(len(f.wb), 2),
                    rejects(lambda: f.wb.sheet(1).show_pivot_details(f.pivot, 0, 0)),
                ),
            ),
        }

    def test_table_covers_every_pivot_mutator(self) -> None:
        declared = {
            name
            for cls in (Workbook, Sheet)
            for name, member in inspect.getmembers(cls, inspect.isfunction)
            if "pivot" in name and not name.startswith("_") and name not in _PIVOT_READERS
        }
        self.assertGreater(len(declared), 40)
        self.assertEqual(sorted(declared - set(self._rows())), [], "pivot mutators without a row")
        self.assertEqual(sorted(set(self._rows()) - declared), [], "rows for undeclared pivot mutators")

    def test_manual_sort_order_keeps_item_order(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            wb.sheet(0).clear_pivot_field_items(f.pivot, f.region)
            wb.sheet(0).add_pivot_field_item(f.pivot, f.region, "West", visible=True)
            wb.sheet(0).add_pivot_field_item(f.pivot, f.region, "East", visible=True)
            wb.sheet(0).set_pivot_field_sort(f.pivot, f.region, PivotSortOrder.MANUAL)
            self.assertEqual(f.grid(), ["行ラベル|Sum of Amount", "West|37", "East|15", "総計|52"])
            wb.sheet(0).set_pivot_field_sort(f.pivot, f.region, PivotSortOrder.ASCENDING)
            self.assertEqual(f.grid(), _BASE_GRID)
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_pivot_field_sort(f.pivot, f.region, 3)
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_pivot_field_sort(f.pivot, f.region, PivotSortOrder.MANUAL, "Sum of Amount")

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
        eq(wb.sheet(0).get_pivot_name(pivot), "Pivot1")
        eq(wb.sheet(0).get_pivot_grand_totals(pivot), PivotGrandTotals(rows=True, columns=True))
        eq(wb.sheet(0).get_pivot_row_field_order(pivot), [f.region])
        eq(wb.sheet(0).get_pivot_column_field_order(pivot), [])

        region = wb.sheet(0).list_pivot_fields(pivot)[f.region]
        eq(region.source_name, "Region")
        eq(region.axis, PivotAxis.ROW)
        eq(region.cache_field_index, 0)
        eq(region.sort_order, PivotSortOrder.ASCENDING)
        eq(region.date_periods, PivotDatePeriod(0))
        eq(region.group_base_index, None)
        eq((region.insert_blank_row, region.fill_down_labels), (False, False))

        eq(wb.sheet(0).get_pivot_field_item_count(pivot, f.region), 2)
        for index, label in enumerate(("East", "West")):
            item = wb.sheet(0).get_pivot_field_item(pivot, f.region, index)
            eq((item.name, item.cache_index, item.visible), (label, index, True))

        eq(len(wb.sheet(0).list_pivot_data_fields(pivot)), 1)
        data_field = wb.sheet(0).list_pivot_data_fields(pivot)[0]
        eq(data_field.name, "Sum of Amount")
        eq(data_field.field_index, f.amount)
        eq(data_field.aggregation, PivotAggregation.SUM)
        eq((data_field.show_as_base_field, data_field.show_as_base_item), (None, None))

    def test_reads_back_what_the_setters_wrote(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            self.assertEqual(wb.sheet(0).get_pivot_cache_id(f.pivot), f.cache)
            self._assert_fixture_definition(wb, f, f.pivot)

            wb.sheet(0).set_pivot_grand_totals(f.pivot, rows=False, columns=True)
            wb.sheet(0).set_pivot_field_sort(f.pivot, f.region, PivotSortOrder.DESCENDING, "Sum of Amount")
            wb.sheet(0).set_pivot_field_item_visible(f.pivot, f.region, 1, False)
            wb.sheet(0).set_pivot_field_date_group(f.pivot, f.date, PivotDatePeriod.YEARS, PivotCalendar.GREGORIAN)
            self.assertEqual(wb.sheet(0).get_pivot_grand_totals(f.pivot), PivotGrandTotals(rows=False, columns=True))
            region = wb.sheet(0).list_pivot_fields(f.pivot)[f.region]
            self.assertEqual((region.sort_order, region.sort_by_field), (PivotSortOrder.DESCENDING, "Sum of Amount"))
            self.assertFalse(wb.sheet(0).get_pivot_field_item(f.pivot, f.region, 1).visible)
            date = wb.sheet(0).list_pivot_fields(f.pivot)[f.date]
            self.assertEqual(date.date_periods, PivotDatePeriod.YEARS)
            self.assertEqual(date.date_calendar, PivotCalendar.GREGORIAN)
            self.assertEqual(date.axis, PivotAxis.VALUE)

            wb.sheet(0).set_pivot_field_axis(f.pivot, f.product, PivotAxis.COL)
            self.assertEqual(wb.sheet(0).get_pivot_column_field_order(f.pivot), [f.product])
            wb.sheet(0).set_pivot_row_field_order(f.pivot, [f.product, f.region])
            self.assertEqual(wb.sheet(0).get_pivot_row_field_order(f.pivot), [f.product, f.region])
            self.assertEqual(wb.sheet(0).get_pivot_column_field_order(f.pivot), [])

    def test_values_level_marker_round_trips_in_field_order(self) -> None:
        self.assertEqual(PIVOT_VALUES_FIELD, -2)
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            wb.sheet(0).add_pivot_data_field(
                f.pivot,
                PivotDataFieldInput(name="Count of Amount", field_index=f.amount, aggregation=PivotAggregation.COUNT),
            )
            wb.sheet(0).set_pivot_column_field_order(f.pivot, [PIVOT_VALUES_FIELD])
            self.assertEqual(wb.sheet(0).get_pivot_column_field_order(f.pivot), [PIVOT_VALUES_FIELD])
            with self.assertRaises(ValueError):
                wb.sheet(0).set_pivot_column_field_order(f.pivot, [-3])

    def test_style_info_defaults_on_create_and_round_trips(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            style = wb.sheet(0).get_pivot_style_info(f.pivot)
            self.assertIsNotNone(style)
            self.assertEqual(style.name, "PivotStyleLight16")
            self.assertEqual(
                (
                    style.show_row_headers,
                    style.show_col_headers,
                    style.show_row_stripes,
                    style.show_col_stripes,
                    style.show_last_column,
                ),
                (True, True, False, False, True),
            )
            wb.sheet(0).set_pivot_style_info(f.pivot, PivotStyleInput("PivotStyleDark3", show_row_stripes=True))
            data = wb.save()
            # An omitted name keeps the element without a style name.
            wb.sheet(0).set_pivot_style_info(f.pivot, PivotStyleInput())
            bare = wb.sheet(0).get_pivot_style_info(f.pivot)
            self.assertIsNotNone(bare)
            self.assertIsNone(bare.name)
        with Workbook.load(data) as loaded:
            style = loaded.sheet(0).get_pivot_style_info(0)
            self.assertEqual(style.name, "PivotStyleDark3")
            self.assertEqual(
                (style.show_row_headers, style.show_row_stripes, style.show_last_column),
                (False, True, False),
            )
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            with self.assertRaises(FormulonError):
                wb.sheet(0).get_pivot_style_info(9)
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_pivot_style_info(9, PivotStyleInput("X"))

    def test_reads_back_a_saved_and_reloaded_pivot(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            data = wb.save()
        with Workbook.load(data) as loaded:
            self.assertEqual(loaded.sheet(0).get_pivot_count(), 1)
            self.assertEqual(loaded.sheet(0).get_pivot_cache_id(0), f.cache)
            self._assert_fixture_definition(loaded, f, 0)

    def test_filters_and_subtotal_functions_survive_save_and_reload(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            wb.sheet(0).add_pivot_filter(f.pivot, _GREATER_THAN_20)
            wb.sheet(0).add_pivot_field_subtotal_fn(f.pivot, f.region, PivotAggregation.AVERAGE)
            wb.sheet(0).add_pivot_field_subtotal_fn(f.pivot, f.region, PivotAggregation.SUM)
            region = wb.sheet(0).list_pivot_fields(f.pivot)[f.region]
            self.assertFalse(region.default_subtotal)
            self.assertEqual(region.subtotal_fns, [PivotAggregation.SUM, PivotAggregation.AVERAGE])
            amount = wb.sheet(0).list_pivot_fields(f.pivot)[f.amount]
            self.assertEqual((amount.default_subtotal, amount.subtotal_fns), (True, []))
            data = wb.save()
        self.assertIn(
            '<filters count="1"><filter fld="0" type="valueGreaterThan"',
            _package_part(data, "xl/pivotTables/pivotTable1.xml"),
        )
        with Workbook.load(data) as loaded:
            loaded.set_excel_profile_id("win-365-ja_JP")
            self.assertEqual(len(loaded.sheet(0).list_pivot_filters(0)), 1)
            spec = loaded.sheet(0).list_pivot_filters(0)[0]
            self.assertEqual((spec.type, spec.field_name, spec.value_double), (_GREATER_THAN_20.type, "Region", 20.0))
            self.assertEqual(_grid(loaded), _WEST_ONLY)
            region = loaded.sheet(0).list_pivot_fields(0)[f.region]
            self.assertFalse(region.default_subtotal)
            self.assertEqual(region.subtotal_fns, [PivotAggregation.SUM, PivotAggregation.AVERAGE])
            loaded.sheet(0).clear_pivot_filters(0)
            self.assertNotIn("<filters", _package_part(loaded.save(), "xl/pivotTables/pivotTable1.xml"))

    def test_out_of_range_indices_raise(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            invalid, bad_sheet = StatusCode.INDEX_OUT_OF_RANGE, StatusCode.SHEET_INDEX_OUT_OF_RANGE
            for call, code in (
                (lambda: wb.sheet(0).get_pivot_cache_id(9), invalid),
                (lambda: wb.sheet(0).get_pivot_name(9), invalid),
                (lambda: wb.sheet(9).get_pivot_grand_totals(f.pivot), bad_sheet),
                (lambda: wb.sheet(0).get_pivot_row_field_order(9), invalid),
                (lambda: wb.sheet(0).get_pivot_column_field_order(9), invalid),
                (lambda: wb.sheet(0).get_pivot_field_item_count(f.pivot, 99), invalid),
                (lambda: wb.sheet(0).get_pivot_field_item(f.pivot, f.region, 99), invalid),
            ):
                with self.assertRaises(FormulonError) as caught:
                    call()
                self.assertEqual(caught.exception.code, code)

    def test_empty_data_field_name_is_derived_from_the_profile(self) -> None:
        for profile, expected in (("win-365-ja_JP", "合計 / Amount"), ("win-365-en_US", "Sum of Amount")):
            for name in ("", None):
                with Workbook.create_default() as wb:
                    f = _PivotFixture(wb)
                    wb.set_excel_profile_id(profile)
                    wb.sheet(0).clear_pivot_data_fields(f.pivot)
                    self.assertEqual(
                        wb.sheet(0).get_pivot_data_field_default_name(f.pivot, f.amount, PivotAggregation.SUM), expected
                    )
                    idx = wb.sheet(0).add_pivot_data_field(
                        f.pivot, PivotDataFieldInput(name=name, field_index=f.amount)
                    )
                    self.assertEqual(wb.sheet(0).list_pivot_data_fields(f.pivot)[idx].name, expected)
                    # The same source twice: the second name is numbered.
                    self.assertEqual(
                        wb.sheet(0).get_pivot_data_field_default_name(f.pivot, f.amount, PivotAggregation.SUM),
                        expected + "2",
                    )
                    idx = wb.sheet(0).add_pivot_data_field(
                        f.pivot, PivotDataFieldInput(name=name, field_index=f.amount)
                    )
                    self.assertEqual(wb.sheet(0).list_pivot_data_fields(f.pivot)[idx].name, expected + "2")
                    # Replacing an entry does not count its own name as taken.
                    self.assertEqual(
                        wb.sheet(0).get_pivot_data_field_default_name(f.pivot, f.amount, PivotAggregation.SUM, 0),
                        expected,
                    )
                    wb.sheet(0).set_pivot_data_field(
                        f.pivot,
                        0,
                        PivotDataFieldInput(name=name, field_index=f.amount, aggregation=PivotAggregation.COUNT),
                    )
                    self.assertTrue(wb.sheet(0).list_pivot_data_fields(f.pivot)[0].name.endswith(" Amount"))

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
                    wb.sheet(args[0]).get_pivot_data_field_default_name(*args[1:])
                self.assertEqual(caught.exception.code, code)
            with self.assertRaises(TypeError):
                wb.sheet(0).get_pivot_data_field_default_name(f.pivot, "0", 0)

    def test_projection_is_pending_after_a_cache_change(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            self.assertFalse(wb.sheet(0).is_pivot_projection_pending(f.pivot))
            wb.add_pivot_cache_field_shared_item_text(f.cache, 0, "North")
            self.assertTrue(wb.sheet(0).is_pivot_projection_pending(f.pivot))
            wb.sheet(0).project_pivot(f.pivot)
            self.assertFalse(wb.sheet(0).is_pivot_projection_pending(f.pivot))

    def test_options_round_trip_through_save(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            options = replace(
                wb.sheet(0).get_pivot_options(f.pivot),
                show_error=True,
                error_caption="N/A",
                grid_drop_zones=True,
                indent=2,
                alt_text="Sales",
                grand_total_caption="All",
            )
            wb.sheet(0).set_pivot_options(f.pivot, options)
            self.assertEqual(wb.sheet(0).get_pivot_options(f.pivot), options)
            # Members left out keep their value; "" clears a string option.
            wb.sheet(0).set_pivot_options(f.pivot, PivotOptions(indent=5, error_caption=""))
            kept = wb.sheet(0).get_pivot_options(f.pivot)
            self.assertEqual((kept.indent, kept.error_caption, kept.alt_text), (5, "", "Sales"))
            self.assertEqual((kept.show_error, kept.grid_drop_zones, kept.grand_total_caption), (True, True, "All"))
            options = kept
            data = wb.save()
        with Workbook.load(data) as loaded:
            self.assertEqual(loaded.sheet(0).get_pivot_options(0), options)

    def test_new_filter_types_and_high_text_round_trip(self) -> None:
        self.assertEqual(len(PivotFilterType), 68)
        self.assertEqual(PivotFilterType.DATE_M12, 67)
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            wb.sheet(0).add_pivot_filter(
                f.pivot,
                PivotFilterInput(
                    axis=PivotAxis.ROW,
                    field_name="Region",
                    type=PivotFilterType.LABEL_BETWEEN,
                    value_kind=2,
                    value_text="A",
                    value_high_kind=2,
                    value_high_text="Z",
                ),
            )
            spec = wb.sheet(0).list_pivot_filters(f.pivot)[0]
            self.assertEqual(
                (spec.type, spec.value_text, spec.value_high_text), (PivotFilterType.LABEL_BETWEEN, "A", "Z")
            )

    def test_derived_name_needs_an_in_range_field(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_pivot_data_field(f.pivot, PivotDataFieldInput(name="", field_index=99))
            self.assertEqual(len(wb.sheet(0).list_pivot_data_fields(f.pivot)), 1)


class PivotFieldBindingTests(unittest.TestCase):
    def _wide_pivot(self, wb: Workbook) -> int:
        cache = wb.add_pivot_cache()
        wb.set_pivot_cache_worksheet_source(cache, PivotWorksheetSource(ref="A1:D3", sheet="Sheet1"))
        for name in ("Region", "Product", "Sales", "Qty"):
            wb.add_pivot_cache_field(cache, name)
        for sales in (100, 200):
            rec = wb.add_pivot_cache_record(cache)
            for col, value in enumerate((0, 0, sales, 1)):
                wb.set_pivot_cache_record_number(cache, rec, col, value)
        return wb.sheet(0).add_pivot("PT", cache, 0, 6)

    def test_field_add_binds_to_the_cache_column_by_name(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._wide_pivot(wb)
            self.assertEqual(
                wb.sheet(0).add_pivot_field(pivot, PivotFieldInput(source_name="Region", axis=PivotAxis.ROW)), 0
            )
            sales = wb.sheet(0).add_pivot_field(pivot, PivotFieldInput(source_name="Sales", axis=PivotAxis.VALUE))
            self.assertEqual(sales, 2)
            self.assertEqual(len(wb.sheet(0).list_pivot_fields(pivot)), 4)
            self.assertIsNone(wb.sheet(0).list_pivot_fields(pivot)[1].axis)
            # Adding a field again updates it instead of appending another.
            self.assertEqual(
                wb.sheet(0).add_pivot_field(pivot, PivotFieldInput(source_name="Sales", axis=PivotAxis.VALUE)), 2
            )
            self.assertEqual(len(wb.sheet(0).list_pivot_fields(pivot)), 4)

    def test_axis_none_takes_a_field_off_the_report(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._wide_pivot(wb)
            wb.sheet(0).add_pivot_field(pivot, PivotFieldInput(source_name="Region", axis=PivotAxis.ROW))
            sales = wb.sheet(0).add_pivot_field(pivot, PivotFieldInput(source_name="Sales", axis=PivotAxis.VALUE))
            wb.sheet(0).add_pivot_data_field(pivot, PivotDataFieldInput(name="Total", field_index=sales))
            wb.sheet(0).set_pivot_field_axis(pivot, sales, None)
            self.assertIsNone(wb.sheet(0).list_pivot_fields(pivot)[sales].axis)
            self.assertEqual(len(wb.sheet(0).list_pivot_data_fields(pivot)), 0)
            wb.sheet(0).set_pivot_field_axis(pivot, 0, None)
            self.assertEqual(wb.sheet(0).get_pivot_row_field_order(pivot), [])
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_pivot_field_axis(pivot, 0, 4)

    def test_save_drops_a_cache_no_pivot_uses(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._wide_pivot(wb)
            self.assertIn(b"pivotCacheDefinition1.xml", wb.save())
            wb.sheet(0).remove_pivot(pivot)
            self.assertEqual(len(wb.list_pivot_caches()), 1)
            self.assertNotIn(b"pivotCacheDefinition1.xml", wb.save())


if __name__ == "__main__":
    unittest.main()
