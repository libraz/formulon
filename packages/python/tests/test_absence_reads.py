"""Absent values read back as ``None``, and the write forms that clear them."""

from __future__ import annotations

import datetime
import unittest

import _chart_input
from test_surface_pivot import _date_rows, _PivotFixture

from formulon import (
    AutoFilter,
    AutoFilterColumn,
    CellRect,
    CfCell,
    CfTimePeriodRuleInput,
    Color,
    ColorKind,
    DatePeriod,
    DxfInput,
    FilterKind,
    InvalidArgumentError,
    PageOrientation,
    Pivot,
    PivotAxis,
    PivotCalendar,
    PivotDataFieldInput,
    PivotDateGroupInput,
    PivotDatePeriod,
    PivotShowAsBase,
    PivotShowValuesAs,
    PrintTitles,
    PrintTitlesInput,
    SparklineGroupInput,
    SparklineInput,
    TableColumnPatch,
    TableInput,
    Workbook,
    get_function_metadata,
)


class CellStyleIndexTests(unittest.TestCase):
    def test_a_cell_with_no_style_of_its_own_reads_none(self) -> None:
        with Workbook.create_default() as wb:
            ooxml = wb.sheet(0).ooxml
            self.assertIsNone(ooxml.get_cell_xf_index(3, 3))
            self.assertIsNone(ooxml.get_cell_xf_index("D4"))

    def test_a_styled_cell_reads_its_index(self) -> None:
        with Workbook.create_default() as wb:
            ooxml = wb.sheet(0).ooxml
            ooxml.set_cell_xf_index(0, 0, 0)
            self.assertEqual(ooxml.get_cell_xf_index(0, 0), 0)
            self.assertEqual(ooxml.get_cell_xf_index("A1"), 0)


class SpillReadTests(unittest.TestCase):
    def test_a_cell_in_no_spill_reads_none(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            self.assertIsNone(wb.sheet(0).get_spill(0, 0))
            self.assertIsNone(wb.sheet(0).get_spill("A1"))

    def test_an_anchor_and_a_spilled_cell_read_the_region(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_formula(0, 0, "=SEQUENCE(3)")
            wb.recalc()
            anchor = wb.sheet(0).get_spill("A1")
            inner = wb.sheet(0).get_spill("A3")
            assert anchor is not None
            self.assertEqual(anchor, inner)
            self.assertEqual(anchor.range, CellRect(0, 0, 2, 0))
            self.assertFalse(anchor.blocked)

    def test_a_blocked_anchor_reads_its_footprint_and_the_cells_below_read_none(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_formula(0, 0, "=SEQUENCE(5)")
            wb.sheet(0).set_number(3, 0, 1.0)
            wb.recalc()
            anchor = wb.sheet(0).get_spill(0, 0)
            assert anchor is not None
            self.assertTrue(anchor.blocked)
            self.assertEqual(anchor.range, CellRect(0, 0, 4, 0))
            self.assertIsNone(wb.sheet(0).get_spill(1, 0))


class PivotCacheAbsenceTests(unittest.TestCase):
    def test_missing_items_limit_is_none_for_automatic_and_clears_with_none(self) -> None:
        with Workbook.create_default() as wb:
            cache = wb.add_pivot_cache()
            self.assertIsNone(cache.get_options().missing_items_limit)
            cache.update_options(missing_items_limit=0)
            self.assertEqual(cache.get_options().missing_items_limit, 0)
            cache.update_options(missing_items_limit=32500)
            self.assertEqual(cache.get_options().missing_items_limit, 32500)
            cache.update_options(missing_items_limit=None)
            self.assertIsNone(cache.get_options().missing_items_limit)

    def test_requested_id_is_optional_and_never_zero(self) -> None:
        with Workbook.create_default() as wb:
            assigned = wb.add_pivot_cache(None)
            self.assertGreater(assigned.cache_id, 0)
            self.assertEqual(wb.add_pivot_cache(requested_id=41).cache_id, 41)
            with self.assertRaises(InvalidArgumentError):
                wb.add_pivot_cache(0)


class PivotFieldAbsenceTests(unittest.TestCase):
    def test_axis_and_binding_read_none_when_absent(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            pivot = wb.sheet(0).pivot(f.pivot)
            fields = pivot.list_fields()
            self.assertEqual(fields[f.region].axis, PivotAxis.ROW)
            self.assertEqual(fields[f.region].cache_field_index, f.region)
            self.assertIsNone(fields[f.region].group_base_index)
            pivot.set_field_axis(f.region, None)
            self.assertIsNone(pivot.list_fields()[f.region].axis)

    def test_a_date_group_field_names_the_field_it_groups(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            _date_rows(f)
            pivot = wb.sheet(0).pivot(f.pivot)
            pivot.set_field_date_group(
                f.date, PivotDateGroupInput(periods=PivotDatePeriod.YEARS, calendar=PivotCalendar.GREGORIAN)
            )
            self.assertTrue(any(x.group_base_index == f.date for x in pivot.list_fields()))

    def test_set_field_date_group_none_ungroups(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            _date_rows(f)
            pivot = wb.sheet(0).pivot(f.pivot)
            before = pivot.list_fields()
            pivot.set_field_date_group(
                f.date, PivotDateGroupInput(periods=PivotDatePeriod.YEARS, calendar=PivotCalendar.GREGORIAN)
            )
            self.assertNotEqual(pivot.list_fields(), before)
            pivot.set_field_date_group(f.date, None)
            self.assertEqual(pivot.list_fields(), before)
            self.assertFalse(hasattr(Pivot, "clear_field_date_group"))


class PivotShowAsBaseTests(unittest.TestCase):
    def test_base_is_none_when_unset_and_takes_members_and_indices(self) -> None:
        with Workbook.create_default() as wb:
            f = _PivotFixture(wb)
            pivot = wb.sheet(0).pivot(f.pivot)
            self.assertEqual((PivotShowAsBase.PREVIOUS.name, PivotShowAsBase.NEXT.name), ("PREVIOUS", "NEXT"))
            self.assertNotEqual(PivotShowAsBase.PREVIOUS, PivotShowAsBase.NEXT)
            self.assertIsNone(pivot.list_data_fields()[0].show_as_base_item)
            for base in (PivotShowAsBase.PREVIOUS, PivotShowAsBase.NEXT, 0, 1):
                with self.subTest(base=base):
                    pivot.set_data_field(
                        0,
                        PivotDataFieldInput(
                            name="Diff",
                            field_index=f.amount,
                            show_as=PivotShowValuesAs.DIFFERENCE_FROM,
                            show_as_base_field=f.region,
                            show_as_base_item=base,
                        ),
                    )
                    got = pivot.list_data_fields()[0]
                    self.assertEqual(got.show_as_base_field, f.region)
                    self.assertEqual(got.show_as_base_item, base)


class AutoFilterAbsenceTests(unittest.TestCase):
    def test_dxf_index_is_none_when_absent_and_a_registered_index_round_trips(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            for r in range(4):
                sheet.set_number(r, 0, float(r))
            registered = wb.ooxml.add_dxf(DxfInput())
            for dxf in (None, registered):
                with self.subTest(dxf=dxf):
                    sheet.set_auto_filter(
                        AutoFilter(
                            range=CellRect(0, 0, 3, 0),
                            columns=[
                                AutoFilterColumn(column_offset=0, kind=FilterKind.COLOR, cell_color=True, dxf_index=dxf)
                            ],
                        )
                    )
                    got = sheet.get_auto_filter()
                    assert got is not None
                    self.assertEqual(got.columns[0].dxf_index, dxf)


class TableAbsenceTests(unittest.TestCase):
    def test_totals_strings_are_none_until_set(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.set_text(0, 0, "A")
            sheet.set_number(1, 0, 1.0)
            table = sheet.add_table(TableInput(name="Totals", range=CellRect(0, 0, 1, 0)))
            column = table.get_properties().columns[0]
            self.assertEqual((column.totals_row_label, column.totals_row_formula), (None, None))
            table.update(columns=[TableColumnPatch(totals_row_label="Total")])
            column = table.get_properties().columns[0]
            self.assertEqual((column.totals_row_label, column.totals_row_formula), ("Total", None))


class PrintAbsenceTests(unittest.TestCase):
    def test_orientation_reads_none_until_stated_and_clears_with_none(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            self.assertIsNone(sheet.get_page_setup().orientation)
            sheet.update_page_setup(orientation=PageOrientation.LANDSCAPE)
            self.assertEqual(sheet.get_page_setup().orientation, PageOrientation.LANDSCAPE)
            sheet.update_page_setup(orientation=None)
            self.assertIsNone(sheet.get_page_setup().orientation)

    def test_pagination_print_titles_is_none_without_titles(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.set_number(0, 0, 1.0)
            self.assertIsNone(sheet.get_pagination().print_titles)
            sheet.set_print_titles(PrintTitlesInput(rows="1:2", columns=None))
            titles = sheet.get_pagination().print_titles
            assert titles is not None
            self.assertIsInstance(titles, PrintTitles)
            self.assertIsNone(titles.columns)
            self.assertEqual((titles.rows.first, titles.rows.last), (0, 1))


class WorkbookAbsenceTests(unittest.TestCase):
    def test_file_name_is_none_until_set(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIsNone(wb.get_file_name())
            wb.set_file_name("Book.xlsx")
            self.assertEqual(wb.get_file_name(), "Book.xlsx")
            wb.set_file_name(None)
            self.assertIsNone(wb.get_file_name())

    def test_open_ended_function_has_no_max_arity(self) -> None:
        self.assertIsNone(get_function_metadata("SUM").max_arity)
        self.assertEqual(get_function_metadata("PI").max_arity, 0)
        self.assertEqual(get_function_metadata("IF").max_arity, 3)
        self.assertIsNone(get_function_metadata("NOT_A_FUNCTION"))


class ChartTitleAbsenceTests(unittest.TestCase):
    def test_title_and_axis_title_read_none_when_removed(self) -> None:
        wb = Workbook.create_default()
        for r, label in enumerate("abc"):
            wb.sheet(0).set_text(r + 1, 0, label)
            wb.sheet(0).set_number(r + 1, 1, float(r))
        wb.sheet(0).set_text(0, 1, "S1")
        with wb:
            handle = _chart_input.add_chart(wb.sheet(0), _chart_input.COLUMN, "Sheet1!$A$1:$B$4", title="T")
            chart = wb.sheet(0).chart(handle.object_id)
            self.assertEqual(chart.get_properties().title.text, "T")
            chart.set_title(None)
            self.assertIsNone(chart.get_properties().title)
            axis = chart.get_properties().axes[0]
            chart.update_axis(axis.id, title="Units")
            self.assertEqual(chart.get_properties().axes[0].title.text, "Units")
            chart.update_axis(axis.id, title=None)
            self.assertIsNone(chart.get_properties().axes[0].title)


class ChartXNumberFormatAbsenceTests(unittest.TestCase):
    def test_non_xy_series_reads_none_x_number_format(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.set_text(0, 1, "S1")
            for r in range(1, 4):
                sheet.set_text(r, 0, f"c{r}")
                sheet.set_number(r, 1, float(r))
            handle = _chart_input.add_chart(sheet, _chart_input.COLUMN, "Sheet1!$A$1:$B$4")
            series = sheet.chart(handle.object_id).get_data().series[0]
            self.assertIsNone(series.x_number_format)
            self.assertEqual(series.number_format, "General")


class SparklineColorAbsenceTests(unittest.TestCase):
    def test_unset_colors_read_none_and_clear_with_none(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            sheet.add_sparkline_group(
                SparklineGroupInput(
                    color_series=Color(kind=ColorKind.RGB, argb=0xFF336699),
                    sparklines=[SparklineInput("Sheet1!A1:D1", "F1")],
                )
            )
            group = sheet.list_sparkline_groups()[0]
            self.assertEqual(group.color_series.argb, 0xFF336699)
            self.assertIsNone(group.color_high)
            self.assertIsNone(group.color_low)
            sheet.update_sparkline_group(0, color_series=None)
            self.assertIsNone(sheet.list_sparkline_groups()[0].color_series)


class EvaluateCfRangeOptionTests(unittest.TestCase):
    def _sheet_with_today_rule(self, wb: Workbook):
        sheet = wb.sheet(0)
        sheet.set_number(0, 0, 45000.0)
        dxf = wb.ooxml.add_dxf(DxfInput())
        sheet.add_cf_rule(
            CfTimePeriodRuleInput(ranges=[CellRect(0, 0, 0, 0)], time_period=DatePeriod.TODAY, dxf_index=dxf)
        )
        return sheet

    def test_options_are_keyword_only(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            with self.assertRaises(TypeError):
                wb.sheet(0).evaluate_cf_range("A1", datetime.date(2023, 3, 15))  # type: ignore[misc]
            self.assertEqual(wb.sheet(0).evaluate_cf_range("A1"), [])

    def test_time_period_rules_off_matches_nothing_and_today_picks_the_day(self) -> None:
        with Workbook.create_default() as wb:
            sheet = self._sheet_with_today_rule(wb)
            day = datetime.date(1899, 12, 30) + datetime.timedelta(days=45000)
            hit = sheet.evaluate_cf_range("A1", today=day)
            miss = sheet.evaluate_cf_range("A1", today=day + datetime.timedelta(days=3))
            off = sheet.evaluate_cf_range("A1", today=day, time_period_rules=False)
            self.assertTrue(all(isinstance(c, CfCell) for c in hit + miss + off))
            self.assertEqual(len(hit[0].matches) if hit else 0, 1)
            self.assertEqual(len(miss[0].matches) if miss else 0, 0)
            self.assertEqual(len(off[0].matches) if off else 0, 0)


if __name__ == "__main__":
    unittest.main()
