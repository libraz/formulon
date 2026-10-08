"""Pivot surface and marshalling tests."""

from __future__ import annotations

import inspect
import io
import unittest
import zipfile
from typing import Any, Dict, List, Tuple

from formulon import (
    FormulonError,
    PivotAggregation,
    PivotAxis,
    PivotCalendar,
    PivotDataFieldSpec,
    PivotDateGrouping,
    PivotFieldSpec,
    PivotFilterSpec,
    PivotFilterType,
    PivotReportLayout,
    PivotWorksheetSource,
    ValueKind,
    Workbook,
)


class PivotTests(unittest.TestCase):
    def test_build_and_project_pivot(self) -> None:
        with Workbook.create_default() as wb:
            cache_id = wb.pivot_cache_create()
            wb.pivot_cache_field_add(cache_id, "Region")
            wb.pivot_cache_field_add(cache_id, "Amount")
            for region, amount in [("East", 10.0), ("East", 20.0), ("West", 30.0)]:
                rec = wb.pivot_cache_record_add(cache_id)
                wb.pivot_cache_record_set_text(cache_id, rec, 0, region)
                wb.pivot_cache_record_set_number(cache_id, rec, 1, amount)
            self.assertEqual(wb.pivot_cache_record_count(cache_id), 3)

            pivot = wb.pivot_create(0, "Pivot1", cache_id, 0, 4)
            region_field = wb.pivot_field_add(0, pivot, PivotFieldSpec(source_name="Region", axis=PivotAxis.ROW))
            amount_field = wb.pivot_field_add(0, pivot, PivotFieldSpec(source_name="Amount", axis=PivotAxis.VALUE))
            self.assertEqual(region_field, 0)
            wb.pivot_data_field_add(
                0,
                pivot,
                PivotDataFieldSpec(
                    name="Sum of Amount",
                    field_index=amount_field,
                    aggregation=PivotAggregation.SUM,
                ),
            )

            layout = wb.pivot_layout(0, pivot)
            self.assertGreater(len(layout.cells), 0)
            numbers = [c.value.to_python() for c in layout.cells if c.value.kind == ValueKind.NUMBER]
            # The single grand total of 10 + 20 + 30 must appear.
            self.assertIn(60.0, numbers)

    def test_cache_source_and_report_layout_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            cache = wb.pivot_cache_create()
            wb.set_pivot_cache_worksheet_source(cache, PivotWorksheetSource(ref="A1:B9", sheet="Sheet1"))
            self.assertEqual(wb.get_pivot_cache_worksheet_source(cache).sheet, "Sheet1")
            pivot = wb.pivot_create(0, "Layout", cache, 0, 0)
            wb.set_pivot_report_layout(0, pivot, PivotReportLayout.TABULAR)
            self.assertEqual(wb.get_pivot_report_layout(0, pivot), PivotReportLayout.TABULAR)
            wb.set_pivot_cache_worksheet_source(cache, None)
            self.assertIsNone(wb.get_pivot_cache_worksheet_source(cache))


class _PivotFixture:
    """Region / Product carry shared items their records index; Amount and
    Date are plain numbers. Only Region is laid out."""

    def __init__(self, wb: Workbook) -> None:
        self.wb = wb
        wb.set_excel_profile_id("win-365-ja_JP")
        self.cache = wb.pivot_cache_create()
        wb.set_pivot_cache_worksheet_source(self.cache, PivotWorksheetSource(ref="A1:D5", sheet="Sheet1"))
        for name in ("Region", "Product", "Amount", "Date"):
            wb.pivot_cache_field_add(self.cache, name)
        for item in ("East", "West"):
            wb.pivot_cache_field_add_shared_item_text(self.cache, 0, item)
        for item in ("Apple", "Pear"):
            wb.pivot_cache_field_add_shared_item_text(self.cache, 1, item)
        for values in ((0, 0, 10, 45292), (1, 0, 30, 45323), (0, 1, 5, 45658), (1, 1, 7, 45689)):
            rec = wb.pivot_cache_record_add(self.cache)
            for field, value in enumerate(values):
                wb.pivot_cache_record_set_number(self.cache, rec, field, value)
        self.pivot = wb.pivot_create(0, "Pivot1", self.cache, 0, 6)
        self.region = wb.pivot_field_add(0, self.pivot, PivotFieldSpec(source_name="Region", axis=PivotAxis.ROW))
        wb.pivot_field_add_item(0, self.pivot, self.region, "East", True)
        wb.pivot_field_add_item(0, self.pivot, self.region, "West", True)
        self.product = wb.pivot_field_add(0, self.pivot, PivotFieldSpec(source_name="Product", axis=PivotAxis.VALUE))
        wb.pivot_field_add_item(0, self.pivot, self.product, "Apple", True)
        wb.pivot_field_add_item(0, self.pivot, self.product, "Pear", True)
        self.amount = wb.pivot_field_add(0, self.pivot, PivotFieldSpec(source_name="Amount", axis=PivotAxis.VALUE))
        self.date = wb.pivot_field_add(0, self.pivot, PivotFieldSpec(source_name="Date", axis=PivotAxis.VALUE))
        wb.pivot_set_row_field_order(0, self.pivot, [self.region])
        wb.pivot_data_field_add(
            0, self.pivot, PivotDataFieldSpec(name="Sum of Amount", field_index=self.amount, aggregation=0)
        )

    def grid(self, pivot: int = 0) -> List[str]:
        """The projected grid as one ``|``-joined string per sheet row."""
        layout = self.wb.pivot_layout(0, pivot)
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
            rows.setdefault(cell.row, [""] * layout.cols)[cell.col - layout.left] = text
        return ["|".join(rows[r]) for r in sorted(rows)]


_BASE_GRID = ["行ラベル|Sum of Amount", "East|15", "West|37", "総計|52"]


_TWO_LEVEL_GRID = [
    "行ラベル|Sum of Amount",
    "Apple|10",
    "Pear|5",
    "East|15",
    "Apple|30",
    "Pear|7",
    "West|37",
    "総計|52",
]


_WEST_ONLY = ["行ラベル|Sum of Amount", "West|37", "総計|37"]


_GREATER_THAN_20 = PivotFilterSpec(
    axis=PivotAxis.ROW, field_name="Region", type=PivotFilterType.VALUE_GREATER_THAN, value_kind=1, value_double=20.0
)


_BETWEEN_10_20 = PivotFilterSpec(
    axis=PivotAxis.ROW,
    field_name="Region",
    type=PivotFilterType.VALUE_BETWEEN,
    value_kind=1,
    value_double=10.0,
    value_high_kind=1,
    value_high_double=20.0,
)


_PIVOT_READERS = {
    "pivot_count",
    "pivot_layout",
    "pivot_cache_count",
    "pivot_cache_field_count",
    "pivot_cache_field_name",
    "pivot_cache_field_shared_item_count",
    "pivot_cache_record_count",
    "pivot_field_count",
    "pivot_data_field_count",
    "pivot_filter_count",
    "pivot_filter_at",
    "get_pivot_cache_worksheet_source",
    "get_pivot_report_layout",
}


def _two_levels(f: _PivotFixture) -> None:
    f.wb.pivot_set_row_field_order(0, f.pivot, [f.region, f.product])


def _date_rows(f: _PivotFixture) -> None:
    f.wb.pivot_set_row_field_order(0, f.pivot, [f.date])


def _hide_east(f: _PivotFixture) -> None:
    f.wb.pivot_field_set_item_visible(0, f.pivot, f.region, 0, False)


def _point_at_new_item(f: _PivotFixture) -> None:
    f.wb.pivot_cache_record_set_number(f.cache, 0, 0, 2)


def _package_part(data: bytes, name: str) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as package:
        return package.read(name).decode("utf-8")


class PivotMutatorTableTests(unittest.TestCase):
    # name -> (setup, act, after, expect). `act` is the one call under test
    # and returns its result; `expect(test, fixture, result)` reads it back.
    def _rows(self) -> Dict[str, Tuple[Any, Any, Any, Any]]:
        eq = self.assertEqual

        def grid_is(expected: List[str]) -> Any:
            return lambda t, f, r: eq(f.grid(), expected)

        def rejects(call: Any) -> None:
            with self.assertRaises(FormulonError) as caught:
                call()
            eq(caught.exception.status, 2)

        def cache_create(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 77)
            eq(f.wb.pivot_cache_count(), 2)

        def cache_remove_setup(f: _PivotFixture) -> None:
            f.spare = f.wb.pivot_cache_create()  # type: ignore[attr-defined]

        def cache_remove(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.pivot_cache_count(), 1)
            eq(f.wb.pivot_cache_id_at(0), f.cache)

        def cache_source(t: Any, f: _PivotFixture, r: None) -> None:
            source = f.wb.get_pivot_cache_worksheet_source(f.cache)
            eq((source.ref, source.sheet), ("B2:E6", "Data"))
            eq(f.grid(), _BASE_GRID)

        def cache_field_add(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 4)
            eq(f.wb.pivot_cache_field_name(f.cache, 4), "Extra")

        def cache_field_clear(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.pivot_cache_field_count(f.cache), 0)
            # The data field now names a cache field that no longer exists.
            with self.assertRaises(FormulonError):
                f.wb.pivot_layout(0, f.pivot)

        def record_add(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 4)
            eq(f.grid(), ["行ラベル|Sum of Amount", "East|15", "West|37", "(空白)|0", "総計|52"])

        def pivot_create(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 1)
            layout = f.wb.pivot_layout(0, 1)
            eq((layout.top, layout.left), (20, 1))
            eq(f.grid(0), _BASE_GRID)

        def pivot_remove(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.pivot_count(0), 0)
            with self.assertRaises(FormulonError):
                f.wb.pivot_layout(0, 0)

        def set_name(t: Any, f: _PivotFixture, r: None) -> None:
            # A pivot's name is only observable in the part the writer emits.
            part = _package_part(f.wb.save(), "xl/pivotTables/pivotTable1.xml")
            self.assertRegex(part, r'<pivotTableDefinition [^>]*name="Renamed"')

        def set_anchor(t: Any, f: _PivotFixture, r: None) -> None:
            layout = f.wb.pivot_layout(0, f.pivot)
            eq((layout.top, layout.left), (10, 2))
            eq(f.grid(), _BASE_GRID)

        def set_layout(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.get_pivot_report_layout(0, f.pivot), PivotReportLayout.TABULAR)
            eq(f.grid(), ["Region|Sum of Amount", "East|15", "West|37", "総計|52"])

        def field_add(t: Any, f: _PivotFixture, r: int) -> None:
            eq(r, 4)
            eq(f.grid(), ["Fruit|(すべて)", "|", *_BASE_GRID])

        def field_clear(t: Any, f: _PivotFixture, r: None) -> None:
            eq(f.wb.pivot_field_count(0, f.pivot), 0)
            eq(f.grid(), _BASE_GRID)

        def add_item_at_setup(f: _PivotFixture) -> None:
            # Only the index-addressed form can name the blank item.
            f.wb.pivot_cache_field_add_shared_item_blank(f.cache, 0)
            f.wb.pivot_cache_record_set_blank(f.cache, 0, 0)

        def clear_subtotal_setup(f: _PivotFixture) -> None:
            _two_levels(f)
            f.wb.pivot_field_add_subtotal_fn(0, f.pivot, f.region, PivotAggregation.MAX)

        def clear_date_group_setup(f: _PivotFixture) -> None:
            _date_rows(f)
            f.wb.pivot_field_set_date_group(0, f.pivot, f.date, PivotDateGrouping.YEAR, PivotCalendar.GREGORIAN)

        def number_format(t: Any, f: _PivotFixture, r: None) -> None:
            rejects(lambda: f.wb.pivot_field_set_number_format(0, f.pivot, 99, "4"))
            rejects(lambda: f.wb.pivot_field_set_number_format(0, f.pivot, f.amount, "0.00"))
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
            formats = [c.number_format for c in f.wb.pivot_layout(0, f.pivot).cells if c.value.kind == ValueKind.NUMBER]
            eq(formats, ["2", "2", "2"])

        def filter_remove_setup(f: _PivotFixture) -> None:
            # Each filter alone keeps a different region, so removing the
            # wrong index leaves the wrong one.
            f.wb.pivot_filter_add(0, f.pivot, _GREATER_THAN_20)
            f.wb.pivot_filter_add(0, f.pivot, _BETWEEN_10_20)

        W = PivotAggregation
        return {
            "pivot_cache_create": (None, lambda f: f.wb.pivot_cache_create(77), None, cache_create),
            "pivot_cache_id_at": (None, lambda f: f.wb.pivot_cache_id_at(0), None, lambda t, f, r: eq(r, f.cache)),
            "pivot_cache_remove": (cache_remove_setup, lambda f: f.wb.pivot_cache_remove(f.spare), None, cache_remove),
            "set_pivot_cache_worksheet_source": (
                None,
                lambda f: f.wb.set_pivot_cache_worksheet_source(
                    f.cache, PivotWorksheetSource(ref="B2:E6", sheet="Data")
                ),
                None,
                cache_source,
            ),
            "pivot_cache_field_add": (
                None,
                lambda f: f.wb.pivot_cache_field_add(f.cache, "Extra"),
                None,
                cache_field_add,
            ),
            "pivot_cache_field_clear": (None, lambda f: f.wb.pivot_cache_field_clear(f.cache), None, cache_field_clear),
            "pivot_cache_field_add_shared_item_number": (
                None,
                lambda f: f.wb.pivot_cache_field_add_shared_item_number(f.cache, 0, 42.5),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "42.5|10", "East|5", "West|37", "総計|52"]),
            ),
            "pivot_cache_field_add_shared_item_text": (
                None,
                lambda f: f.wb.pivot_cache_field_add_shared_item_text(f.cache, 0, "North"),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "North|10", "West|37", "総計|52"]),
            ),
            "pivot_cache_field_add_shared_item_bool": (
                None,
                lambda f: f.wb.pivot_cache_field_add_shared_item_bool(f.cache, 0, True),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "TRUE|10", "総計|52"]),
            ),
            "pivot_cache_field_add_shared_item_blank": (
                None,
                lambda f: f.wb.pivot_cache_field_add_shared_item_blank(f.cache, 0),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "(空白)|10", "総計|52"]),
            ),
            "pivot_cache_field_add_shared_item_error": (
                None,
                lambda f: f.wb.pivot_cache_field_add_shared_item_error(f.cache, 0, 6),
                _point_at_new_item,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "#N/A|10", "総計|52"]),
            ),
            "pivot_cache_field_clear_shared_items": (
                None,
                lambda f: f.wb.pivot_cache_field_clear_shared_items(f.cache, 0),
                None,
                # With no shared items the records' indices render as themselves.
                grid_is(["行ラベル|Sum of Amount", "0|15", "1|37", "総計|52"]),
            ),
            "pivot_cache_record_add": (None, lambda f: f.wb.pivot_cache_record_add(f.cache), None, record_add),
            "pivot_cache_record_clear": (
                None,
                lambda f: f.wb.pivot_cache_record_clear(f.cache),
                None,
                grid_is(["行ラベル|Sum of Amount", "総計|0"]),
            ),
            "pivot_cache_record_set_number": (
                None,
                lambda f: f.wb.pivot_cache_record_set_number(f.cache, 0, 2, 1000),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|1005", "West|37", "総計|1042"]),
            ),
            "pivot_cache_record_set_text": (
                None,
                lambda f: f.wb.pivot_cache_record_set_text(f.cache, 0, 0, "West"),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|47", "総計|52"]),
            ),
            "pivot_cache_record_set_bool": (
                None,
                lambda f: f.wb.pivot_cache_record_set_bool(f.cache, 0, 2, True),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|6", "West|37", "総計|43"]),
            ),
            "pivot_cache_record_set_blank": (
                None,
                lambda f: f.wb.pivot_cache_record_set_blank(f.cache, 0, 2),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "総計|42"]),
            ),
            "pivot_cache_record_set_error": (
                None,
                lambda f: f.wb.pivot_cache_record_set_error(f.cache, 0, 2, 1),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|#1", "West|37", "総計|#1"]),
            ),
            "pivot_create": (None, lambda f: f.wb.pivot_create(0, "Pivot2", f.cache, 20, 1), None, pivot_create),
            "pivot_remove": (None, lambda f: f.wb.pivot_remove(0, f.pivot), None, pivot_remove),
            "pivot_set_name": (None, lambda f: f.wb.pivot_set_name(0, f.pivot, "Renamed"), None, set_name),
            "pivot_set_anchor": (None, lambda f: f.wb.pivot_set_anchor(0, f.pivot, 10, 2, 4, 2), None, set_anchor),
            "pivot_set_grand_totals": (
                lambda f: f.wb.pivot_set_col_field_order(0, f.pivot, [f.product]),
                # Rows off drops the per-row total column; columns on keeps
                # the bottom total row, so a swapped pair shows the opposite.
                lambda f: f.wb.pivot_set_grand_totals(0, f.pivot, False, True),
                None,
                grid_is(["Sum of Amount|列ラベル|", "行ラベル|Apple|Pear", "East|10|5", "West|30|7", "総計|40|12"]),
            ),
            "set_pivot_report_layout": (
                None,
                lambda f: f.wb.set_pivot_report_layout(0, f.pivot, PivotReportLayout.TABULAR),
                None,
                set_layout,
            ),
            "pivot_field_add": (
                None,
                lambda f: f.wb.pivot_field_add(
                    0, f.pivot, PivotFieldSpec(source_name="Product", custom_name="Fruit", axis=PivotAxis.PAGE)
                ),
                None,
                field_add,
            ),
            "pivot_field_clear": (_hide_east, lambda f: f.wb.pivot_field_clear(0, f.pivot), None, field_clear),
            "pivot_field_set_axis": (
                None,
                lambda f: f.wb.pivot_field_set_axis(0, f.pivot, f.product, PivotAxis.PAGE),
                None,
                grid_is(["Product|(すべて)", "|", *_BASE_GRID]),
            ),
            "pivot_field_set_sort": (
                # East now outsums West, so sorting ascending by the data
                # field differs from both label orders.
                lambda f: f.wb.pivot_cache_record_set_number(f.cache, 0, 2, 40),
                lambda f: f.wb.pivot_field_set_sort(0, f.pivot, f.region, True, "Sum of Amount"),
                None,
                grid_is(["行ラベル|Sum of Amount", "West|37", "East|45", "総計|82"]),
            ),
            "pivot_field_set_subtotal_top": (
                _two_levels,
                lambda f: f.wb.pivot_field_set_subtotal_top(0, f.pivot, f.region, True),
                None,
                grid_is(
                    [
                        "行ラベル|Sum of Amount",
                        "East|15",
                        "Apple|10",
                        "Pear|5",
                        "West|37",
                        "Apple|30",
                        "Pear|7",
                        "総計|52",
                    ]
                ),
            ),
            "pivot_field_add_item": (
                None,
                lambda f: f.wb.pivot_field_add_item(0, f.pivot, f.region, "East", False),
                None,
                grid_is(_WEST_ONLY),
            ),
            "pivot_field_add_item_at": (
                add_item_at_setup,
                lambda f: f.wb.pivot_field_add_item_at(0, f.pivot, f.region, 2, False),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|5", "West|37", "総計|42"]),
            ),
            "pivot_field_clear_items": (
                _hide_east,
                lambda f: f.wb.pivot_field_clear_items(0, f.pivot, f.region),
                None,
                grid_is(_BASE_GRID),
            ),
            "pivot_field_set_item_visible": (
                None,
                lambda f: f.wb.pivot_field_set_item_visible(0, f.pivot, f.region, 1, False),
                None,
                grid_is(["行ラベル|Sum of Amount", "East|15", "総計|15"]),
            ),
            "pivot_field_add_subtotal_fn": (
                _two_levels,
                lambda f: f.wb.pivot_field_add_subtotal_fn(0, f.pivot, f.region, W.MAX),
                None,
                grid_is(
                    [
                        "行ラベル|Sum of Amount",
                        "Apple|10",
                        "Pear|5",
                        "East|10",
                        "Apple|30",
                        "Pear|7",
                        "West|30",
                        "総計|52",
                    ]
                ),
            ),
            "pivot_field_clear_subtotal_fns": (
                clear_subtotal_setup,
                lambda f: f.wb.pivot_field_clear_subtotal_fns(0, f.pivot, f.region),
                None,
                grid_is(_TWO_LEVEL_GRID),
            ),
            "pivot_field_set_date_group": (
                _date_rows,
                lambda f: f.wb.pivot_field_set_date_group(
                    0, f.pivot, f.date, PivotDateGrouping.YEAR, PivotCalendar.JAPANESE, -1, -1
                ),
                None,
                grid_is(["行ラベル|Sum of Amount", "令和6年|40", "令和7年|12", "総計|52"]),
            ),
            "pivot_field_clear_date_group": (
                clear_date_group_setup,
                lambda f: f.wb.pivot_field_clear_date_group(0, f.pivot, f.date),
                None,
                grid_is(["行ラベル|Sum of Amount", "45292|10", "45323|30", "45658|5", "45689|7", "総計|52"]),
            ),
            "pivot_field_set_number_format": (
                None,
                lambda f: f.wb.pivot_field_set_number_format(0, f.pivot, f.amount, "4"),
                None,
                number_format,
            ),
            "pivot_set_row_field_order": (
                None,
                lambda f: f.wb.pivot_set_row_field_order(0, f.pivot, [f.product]),
                None,
                grid_is(["行ラベル|Sum of Amount", "Apple|40", "Pear|12", "総計|52"]),
            ),
            "pivot_set_col_field_order": (
                None,
                lambda f: f.wb.pivot_set_col_field_order(0, f.pivot, [f.product]),
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
            "pivot_data_field_add": (
                None,
                lambda f: f.wb.pivot_data_field_add(
                    0, f.pivot, PivotDataFieldSpec(name="Max of Amount", field_index=f.amount, aggregation=W.MAX)
                ),
                None,
                data_field_add,
            ),
            "pivot_data_field_clear": (
                None,
                lambda f: f.wb.pivot_data_field_clear(0, f.pivot),
                None,
                grid_is(["行ラベル", "East", "West", "総計"]),
            ),
            "pivot_data_field_set": (
                None,
                lambda f: f.wb.pivot_data_field_set(
                    0,
                    f.pivot,
                    0,
                    PivotDataFieldSpec(
                        name="Count of Amount", field_index=f.amount, aggregation=W.COUNT, number_format="2"
                    ),
                ),
                None,
                data_field_set,
            ),
            "pivot_filter_add": (
                None,
                lambda f: f.wb.pivot_filter_add(0, f.pivot, _GREATER_THAN_20),
                None,
                grid_is(_WEST_ONLY),
            ),
            "pivot_filter_clear": (
                lambda f: f.wb.pivot_filter_add(0, f.pivot, _GREATER_THAN_20),
                lambda f: f.wb.pivot_filter_clear(0, f.pivot),
                None,
                grid_is(_BASE_GRID),
            ),
            "pivot_filter_remove_at": (
                filter_remove_setup,
                lambda f: f.wb.pivot_filter_remove_at(0, f.pivot, 1),
                None,
                grid_is(_WEST_ONLY),
            ),
        }

    def test_table_covers_every_pivot_mutator(self) -> None:
        declared = {
            name
            for name, member in inspect.getmembers(Workbook, inspect.isfunction)
            if "pivot" in name and not name.startswith("_") and name not in _PIVOT_READERS
        }
        self.assertGreater(len(declared), 40)
        self.assertEqual(sorted(declared - set(self._rows())), [], "pivot mutators without a row")
        self.assertEqual(sorted(set(self._rows()) - declared), [], "rows for undeclared pivot mutators")

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


if __name__ == "__main__":
    unittest.main()
