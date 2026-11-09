"""Setters take record arguments positionally, and the record a getter returns passes back into its setter."""

from __future__ import annotations

import dataclasses
import datetime
import unittest
from typing import Any

import formulon
from formulon import (
    CellAddress,
    CellRect,
    ChartBarDirection,
    ChartPlotType,
    FontScheme,
    FormulonError,
    PivotAxis,
    PivotFieldInput,
    PivotFilterKind,
    PivotInput,
    SheetRect,
    TableInput,
    Workbook,
)


def _grid(ws: Any) -> None:
    """A1:C4 header plus numbers, a table-shaped block."""
    for c, name in enumerate(("Name", "North", "South")):
        ws.set_text(0, c, name)
    for r in range(1, 4):
        ws.set_text(r, 0, f"row{r}")
        ws.set_number(r, 1, float(r))
        ws.set_number(r, 2, float(r * 2))


def _chart(wb: Workbook) -> Any:
    _grid(wb.sheet(0))
    return wb.sheet(0).add_chart(formulon.ChartInsertInput(type=ChartPlotType.BAR, range="Sheet1!A1:C4"))


def _pivot_cache(wb: Workbook) -> Any:
    _grid(wb.sheet(0))
    cache = wb.add_pivot_cache()
    cache.set_worksheet_source(formulon.PivotWorksheetSourceInput(range="A1:C4", sheet=0))
    return cache


class PhoneticPropertiesSetterTests(unittest.TestCase):
    def test_setter_takes_the_record_the_getter_returns(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            props = formulon.PhoneticProperties(font_index=0, type=3, alignment=1)
            ws.set_cell_phonetic_properties("B1", props)
            self.assertEqual(ws.get_cell_phonetic_properties("B1").type, 3)
            ws.set_cell_phonetic_properties("A1", ws.get_cell_phonetic_properties("B1"))
            self.assertEqual(ws.get_cell_phonetic_properties("A1").alignment, 1)
            ws.set_cell_phonetic_properties(2, 2, props)
            self.assertEqual(ws.get_cell_phonetic_properties(2, 2).type, 3)

    def test_record_fields_all_have_defaults(self) -> None:
        for f in dataclasses.fields(formulon.PhoneticProperties):
            self.assertIsNot(f.default, dataclasses.MISSING, f.name)

    def test_old_keyword_form_is_refused(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.sheet(0).set_cell_phonetic_properties(0, 0, font_index=0, type=1, alignment=1)  # type: ignore[call-overload]


class ChartInputTests(unittest.TestCase):
    def test_add_chart_takes_an_insert_record(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb.sheet(0))
            chart = wb.sheet(0).add_chart(
                formulon.ChartInsertInput(
                    type=ChartPlotType.BAR, range="Sheet1!A1:C4", bar_direction=ChartBarDirection.BAR
                )
            )
            plot = chart.get_properties().plots[0]
            self.assertEqual(plot.type, ChartPlotType.BAR)
            self.assertEqual(plot.bar_direction, ChartBarDirection.BAR)

    def test_insert_range_accepts_a_sheet_rect(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb.sheet(0))
            a = wb.sheet(0).add_chart(
                formulon.ChartInsertInput(type=ChartPlotType.LINE, range=SheetRect(0, CellRect(0, 0, 3, 2)))
            )
            b = wb.sheet(0).add_chart(formulon.ChartInsertInput(type=ChartPlotType.LINE, range="Sheet1!A1:C4"))
            self.assertEqual(a.get_properties().source, b.get_properties().source)

    def test_insert_legend_defaults_to_bottom(self) -> None:
        self.assertEqual(
            formulon.ChartInsertInput(type=ChartPlotType.BAR, range="A1:B2").legend, formulon.ChartLegendPosition.BOTTOM
        )

    def test_old_positional_kind_form_is_refused(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb.sheet(0))
            with self.assertRaises(TypeError):
                wb.sheet(0).add_chart(0, "Sheet1!A1:C4")  # type: ignore[call-overload]

    def test_set_source_takes_the_record_the_getter_returns(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            source = chart.get_properties().source
            self.assertIsInstance(source.range, SheetRect)
            chart.set_source(source)
            self.assertEqual(chart.get_properties().source, source)

    def test_set_source_takes_an_input_record(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            chart.set_source(formulon.ChartSourceInput(range="Sheet1!A1:B4"))
            self.assertEqual(chart.get_properties().source.range, SheetRect(0, CellRect(0, 0, 3, 1)))
            chart.set_source(formulon.ChartSourceInput(range=SheetRect(0, CellRect(0, 0, 3, 2))))
            self.assertEqual(chart.get_properties().source.range, SheetRect(0, CellRect(0, 0, 3, 2)))

    def test_set_source_refuses_the_old_keyword_form(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            with self.assertRaises(TypeError):
                chart.set_source("Sheet1!A1:B4", direction=formulon.ChartSeriesDirection.AUTO)  # type: ignore[call-overload]

    def test_set_type_takes_a_plot_type_and_bar_direction(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            chart.set_type(0, ChartPlotType.BAR, bar_direction=ChartBarDirection.BAR)
            plot = chart.get_properties().plots[0]
            self.assertEqual((plot.type, plot.bar_direction), (ChartPlotType.BAR, ChartBarDirection.BAR))
            chart.set_type(0, plot.type, bar_direction=plot.bar_direction, grouping=plot.grouping)
            chart.set_type(0, ChartPlotType.LINE)
            self.assertEqual(chart.get_properties().plots[0].type, ChartPlotType.LINE)

    def test_set_type_refuses_an_unwritable_type(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            with self.assertRaises(FormulonError):
                chart.set_type(0, ChartPlotType.STOCK)
            with self.assertRaises(TypeError):
                chart.set_type(0, ChartPlotType.LINE, bar_direction=ChartBarDirection.BAR)

    def test_insert_refuses_a_bar_direction_on_a_non_bar_type(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb.sheet(0))
            with self.assertRaises(TypeError):
                wb.sheet(0).add_chart(
                    formulon.ChartInsertInput(
                        type=ChartPlotType.LINE, range="Sheet1!A1:C4", bar_direction=ChartBarDirection.BAR
                    )
                )
            self.assertEqual(wb.sheet(0).list_charts(), [])

    def test_the_retired_insert_kind_enum_is_gone(self) -> None:
        self.assertFalse(hasattr(formulon, "ChartInsertKind"))


class PivotInputTests(unittest.TestCase):
    def test_worksheet_source_takes_the_record_it_returns(self) -> None:
        with Workbook.create_default() as wb:
            cache = _pivot_cache(wb)
            source = cache.get_worksheet_source()
            self.assertIsInstance(source.range, SheetRect)
            self.assertFalse(hasattr(source, "ref"))
            self.assertFalse(hasattr(source, "sheet"))
            cache.set_worksheet_source(source)
            self.assertEqual(cache.get_worksheet_source(), source)

    def test_worksheet_source_resolves_the_three_range_spellings(self) -> None:
        with Workbook.create_default() as wb:
            cache = _pivot_cache(wb)
            expected = SheetRect(0, CellRect(0, 0, 3, 2))
            for given in (
                formulon.PivotWorksheetSourceInput(range="A1:C4", sheet="Sheet1"),
                formulon.PivotWorksheetSourceInput(range="Sheet1!A1:C4"),
                formulon.PivotWorksheetSourceInput(range=expected),
            ):
                cache.set_worksheet_source(given)
                self.assertEqual(cache.get_worksheet_source().range, expected)

    def test_worksheet_source_with_a_disagreeing_sheet_is_refused(self) -> None:
        with Workbook.create_default() as wb:
            cache = _pivot_cache(wb)
            wb.add_sheet("Other")
            with self.assertRaises(FormulonError):
                cache.set_worksheet_source(formulon.PivotWorksheetSourceInput(range="Sheet1!A1:C4", sheet="Other"))

    def test_worksheet_source_none_clears(self) -> None:
        with Workbook.create_default() as wb:
            cache = _pivot_cache(wb)
            cache.set_worksheet_source(None)
            self.assertIsNone(cache.get_worksheet_source())

    def test_worksheet_source_refuses_the_old_keyword_form(self) -> None:
        with Workbook.create_default() as wb:
            cache = _pivot_cache(wb)
            with self.assertRaises(TypeError):
                cache.set_worksheet_source(ref="A1:C4", sheet="Sheet1")  # type: ignore[call-overload]

    def test_date_group_setter_takes_a_group_record(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            cache = wb.add_pivot_cache()
            cache.add_field("When")
            cache.add_field("Amount")
            for day in (datetime.date(2024, 1, 5), datetime.date(2024, 3, 9)):
                record = cache.add_record()
                cache.set_record_value(record, 0, day)
                cache.set_record_value(record, 1, 1.0)
            pivot = sheet.pivot(
                sheet.add_pivot(PivotInput(name="P", cache_id=cache.cache_id, anchor=CellAddress(0, 4)))
            )
            pivot.add_field(PivotFieldInput(field_index="When", axis=PivotAxis.ROW))
            pivot.set_field_date_group(
                0,
                formulon.PivotDateGroupInput(
                    periods=formulon.PivotDatePeriod.MONTHS,
                    start=datetime.date(2024, 1, 1),
                    end=datetime.date(2024, 12, 31),
                ),
            )
            self.assertTrue(pivot.list_fields()[0].date_periods & formulon.PivotDatePeriod.MONTHS)
            with self.assertRaises(TypeError):
                pivot.set_field_date_group(0, periods=formulon.PivotDatePeriod.YEARS)  # type: ignore[call-overload]

    def test_date_group_record_defaults(self) -> None:
        group = formulon.PivotDateGroupInput(periods=formulon.PivotDatePeriod.YEARS)
        self.assertEqual(group.calendar, formulon.PivotCalendar.GREGORIAN)
        self.assertEqual(group.interval_days, 1)
        self.assertIsNone(group.start)
        self.assertIsNone(group.end)

    def test_filter_records_are_keyword_only(self) -> None:
        for cls in (formulon.PivotFilterInput, formulon.PivotFilter):
            with self.assertRaises(TypeError):
                cls("Product", PivotFilterKind.LABEL_EQUAL)  # type: ignore[call-arg]
        built = formulon.PivotFilterInput(field_index="Product", type=PivotFilterKind.LABEL_EQUAL)
        self.assertEqual(built.field_index, "Product")

    def test_field_axis_is_required(self) -> None:
        with self.assertRaises(TypeError):
            PivotFieldInput(field_index="Product")  # type: ignore[call-arg]
        with self.assertRaises(TypeError):
            PivotFieldInput("Product")  # type: ignore[call-arg]

    def test_set_column_totals_takes_the_column_record(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            _grid(ws)
            table = ws.add_table(TableInput(range="A1:C4", name="T"))
            table.set_column_totals(
                1,
                formulon.TableColumnTotalsInput(totals_row_function=formulon.TableTotalsRowFunction.SUM),
            )
            column = table.get_properties().columns[1]
            self.assertEqual(column.totals_row_function, formulon.TableTotalsRowFunction.SUM)
            table.set_column_totals(2, column)
            self.assertEqual(table.get_properties().columns[2].totals_row_function, formulon.TableTotalsRowFunction.SUM)

    def test_set_column_totals_refuses_the_old_keyword_form(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            _grid(ws)
            table = ws.add_table(TableInput(range="A1:C4", name="T"))
            with self.assertRaises(TypeError):
                table.set_column_totals(1, function=formulon.TableTotalsRowFunction.SUM)  # type: ignore[call-overload]


class SheetRecordInputTests(unittest.TestCase):
    def test_selection_takes_the_record_the_getter_returns(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_selection(formulon.SheetSelectionInput("C3", "A1:D5"))
            got = ws.get_selection()
            self.assertEqual(got.active, CellAddress(2, 2))
            self.assertEqual(got.ranges, [CellRect(0, 0, 4, 3)])
            ws.set_selection(formulon.SheetSelectionInput(CellAddress(1, 1), ["B2:C3"]))
            ws.set_selection(got)
            self.assertEqual(ws.get_selection(), got)
            with self.assertRaises(TypeError):
                ws.set_selection(active_row=1, active_col=1, ranges="B2:C3")  # type: ignore[call-overload]

    def test_auto_filter_input_takes_a_string_range(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            _grid(ws)
            ws.set_auto_filter(formulon.AutoFilterInput("A1:C4"))
            self.assertEqual(ws.get_auto_filter().range, CellRect(0, 0, 3, 2))
            ws.set_auto_filter(ws.get_auto_filter())
            self.assertEqual(ws.get_auto_filter().range, CellRect(0, 0, 3, 2))

    def test_table_auto_filter_takes_the_input_record(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            _grid(ws)
            table = ws.add_table(TableInput(range="A1:C4", name="T"))
            table.set_auto_filter(formulon.AutoFilterInput("A1:C4"))
            self.assertIsNotNone(table.get_auto_filter())

    def test_sparkline_group_input_takes_string_sources(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            _grid(ws)
            group_id = ws.add_sparkline_group(
                formulon.SparklineGroupInput(sparklines=[formulon.SparklineInput("Sheet1!B2:C2", "E1")])
            )
            group = ws.list_sparkline_groups()[0]
            self.assertEqual(group.id, group_id)
            spark = group.sparklines[0]
            self.assertEqual(spark.source, SheetRect(0, CellRect(1, 1, 1, 2)))
            self.assertEqual(spark.location, CellAddress(0, 4))
            ws.add_sparkline_group(group)

    def test_sparkline_group_date_axis_is_a_range(self) -> None:
        group = formulon.SparklineGroupInput(date_axis_range="Sheet1!A2:C2")
        self.assertEqual(group.date_axis_range, "Sheet1!A2:C2")
        self.assertFalse(hasattr(formulon.SparklineGroup(), "date_axis_ref"))

    def test_protected_range_input_takes_a_string(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.add_protected_range(formulon.ProtectedRangeInput(name="p", ranges="A1:B2"))
            read = ws.list_protected_ranges()[0]
            self.assertEqual(read.ranges, [CellRect(0, 0, 1, 1)])
            ws.add_protected_range(dataclasses.replace(read, name="q"))
            self.assertEqual(len(ws.list_protected_ranges()), 2)

    def test_update_protected_range_takes_a_string(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.add_protected_range(formulon.ProtectedRangeInput(name="p", ranges="A1:B2"))
            ws.update_protected_range(0, ranges=["B2"])
            self.assertEqual(ws.list_protected_ranges()[0].ranges, [CellRect(1, 1, 1, 1)])

    def test_array_formula_takes_an_a1_range(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_array_formula(formulon.ArrayFormula(range="A5:A6", formula="=SEQUENCE(2)"))
            read = ws.get_array_formula("A5")
            self.assertEqual(read.range, CellRect(4, 0, 5, 0))
            ws.set_array_formula(read)
            self.assertFalse(hasattr(read, "rect"))

    def test_print_titles_take_span_strings(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_print_titles(formulon.PrintTitlesInput(rows="1:2", columns="A:B"))
            read = ws.get_print_titles()
            self.assertEqual((read.rows.first, read.rows.last), (0, 1))
            self.assertEqual((read.columns.first, read.columns.last), (0, 1))
            ws.set_print_titles(read)
            self.assertEqual(ws.get_print_titles(), read)
            with self.assertRaises(FormulonError):
                ws.set_print_titles(formulon.PrintTitlesInput(rows="A:B", columns=None))

    def test_data_table_input_cell_takes_a_cell_address(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.add_data_table(formulon.DataTableInput(range="B2:D6", row_input=CellAddress(0, 0)))
            table = ws.get_data_table("B2")
            self.assertEqual(table.range, CellRect(1, 1, 5, 3))
            ws.add_data_table(formulon.DataTableInput(range="F2:H6", row_input=table.row_input))

    def test_comment_author_is_keyword_only_and_optional(self) -> None:
        self.assertEqual(formulon.Comment("check this").author, "")
        self.assertEqual(formulon.Comment("check this", author="ann").author, "ann")
        with self.assertRaises(TypeError):
            formulon.Comment("ann", "check this")  # type: ignore[misc]
        with self.assertRaises(TypeError):
            formulon.Comment("check this", "ann")  # type: ignore[misc]

    def test_comment_round_trips_through_the_setter(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_comment("B2", formulon.Comment("check this"))
            self.assertEqual(ws.get_comment("B2"), formulon.Comment(text="check this", author=""))
            ws.set_comment("C3", ws.get_comment("B2"))
            self.assertEqual(ws.get_comment("C3").text, "check this")

    def test_workbook_protection_input_defaults_and_union(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_workbook_protection(formulon.WorkbookProtectionInput(lock_structure=True))
            read = wb.get_workbook_protection()
            self.assertTrue(read.lock_structure)
            self.assertFalse(read.lock_windows)
            wb.set_workbook_protection(None)
            wb.set_workbook_protection(read)
            self.assertEqual(wb.get_workbook_protection(), read)


class ThemeInputTests(unittest.TestCase):
    def test_theme_fonts_have_no_defaults(self) -> None:
        with self.assertRaises(TypeError):
            formulon.ThemeFonts(major_latin="X")  # type: ignore[call-arg]
        for f in dataclasses.fields(formulon.ThemeFonts):
            self.assertIs(f.default, dataclasses.MISSING, f.name)
            self.assertIs(f.default_factory, dataclasses.MISSING, f.name)

    def test_theme_fonts_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            fonts = wb.get_theme_fonts()
            wb.set_theme_fonts(fonts)
            self.assertEqual(wb.get_theme_fonts(), fonts)
            wb.set_theme_fonts(formulon.ThemeFonts("A", "B", "C", "D"))
            self.assertEqual(wb.get_theme_fonts(), formulon.ThemeFonts("A", "B", "C", "D"))

    def test_theme_colors_take_the_record_the_getter_returns(self) -> None:
        with Workbook.create_default() as wb:
            colors = list(wb.get_theme_colors().colors)
            colors[4] = 0xFF123456
            wb.set_theme_colors(formulon.ThemeColorsInput(colors=colors))
            read = wb.get_theme_colors()
            self.assertEqual(read.colors[4], 0xFF123456)
            wb.set_theme_colors(read)
            self.assertEqual(wb.get_theme_colors().colors, read.colors)
            with self.assertRaises(TypeError):
                wb.set_theme_colors(colors=colors)  # type: ignore[call-overload]

    def test_theme_script_font_takes_the_record_the_getter_returns(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_theme_script_font(
                formulon.ThemeScriptFont(scheme=FontScheme.MAJOR, script="Jpan", typeface="Meiryo")
            )
            read = wb.get_theme_script_font(scheme=FontScheme.MAJOR, script="Jpan")
            self.assertEqual(read.typeface, "Meiryo")
            wb.set_theme_script_font(dataclasses.replace(read, script="Hang"))
            self.assertEqual(wb.get_theme_script_font(scheme=FontScheme.MAJOR, script="Hang").typeface, "Meiryo")
            with self.assertRaises(TypeError):
                wb.set_theme_script_font(scheme=FontScheme.MAJOR, script="Jpan", typeface="X")  # type: ignore[call-overload]


class HandleConstructorTests(unittest.TestCase):
    def test_handles_cannot_be_constructed_directly(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb.sheet(0))
            with self.assertRaises(TypeError):
                formulon.Sheet(wb, 0, "X")  # type: ignore[call-arg]
            with self.assertRaises(TypeError):
                formulon.Table(wb, 0)  # type: ignore[call-arg]
            with self.assertRaises(TypeError):
                formulon.Pivot(wb.sheet(0), 0)  # type: ignore[call-arg]
            with self.assertRaises(TypeError):
                formulon.PivotCache(wb, 0)  # type: ignore[call-arg]

    def test_handle_accessors_still_work(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIs(wb[0], wb.sheet(0))
            self.assertIsInstance(wb.sheet(0), formulon.Sheet)
            chart = _chart(wb)
            self.assertIsInstance(chart, formulon.Chart)
            with self.assertRaises(TypeError):
                formulon.Chart(wb.sheet(0), chart.object_id, formulon.DrawingObjectKind.CHART)  # type: ignore[call-arg]
