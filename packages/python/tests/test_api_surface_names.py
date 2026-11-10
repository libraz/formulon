"""Public names, shapes and defaults of the Python surface.

Renamed members exist under their final spelling and the old spelling is gone; ``update_*`` members merge keyword
arguments and keep what is ``UNCHANGED``; singleton ``set_*`` members take a record or ``None``; dimensioned fields
carry their unit; the table lifecycle sits on :class:`Sheet`.
"""

from __future__ import annotations

import dataclasses
import datetime
import unittest
from typing import Any, Iterable, Set

import _chart_input

import formulon
from formulon import (
    AlreadyExistsError,
    CellRect,
    ChartLegend,
    ChartLegendPosition,
    NotFoundError,
    PivotAxis,
    SheetProtection,
    ThemeColorsInput,
    ThemeSource,
    Workbook,
)

# ---- Fixtures ---------------------------------------------------------------------


def fill_source(sheet: formulon.Sheet) -> None:
    """Region / Amount rows in A1:B3 plus a few more rows so pagination has a body."""
    sheet.set_text(0, 0, "Region")
    sheet.set_text(0, 1, "Amount")
    for row, (region, amount) in enumerate((("East", 10), ("West", 30)), start=1):
        sheet.set_text(row, 0, region)
        sheet.set_number(row, 1, amount)
    for row in range(3, 12):
        sheet.set_number(row, 0, row)
        sheet.set_number(row, 1, row * 2)


def add_cache(wb: Workbook, source_sheet: str) -> formulon.PivotCache:
    cache = wb.add_pivot_cache()
    cache.set_worksheet_source(formulon.PivotWorksheetSourceInput(range="A1:B3", sheet=source_sheet))
    cache.add_field("Region")
    cache.add_field("Amount")
    for region, amount in (("East", 10), ("West", 30)):
        record = cache.add_record()
        cache.set_record_value(record, 0, region)
        cache.set_record_value(record, 1, amount)
    return cache


def add_pivot(sheet: formulon.Sheet, cache: formulon.PivotCache, name: str = "P1") -> formulon.Pivot:
    pivot = sheet.add_pivot(formulon.PivotInput(name=name, cache_id=cache.cache_id, anchor=formulon.CellAddress(0, 4)))
    pivot.set_field_axis("Region", PivotAxis.ROW)
    pivot.set_field_axis("Amount", PivotAxis.VALUE)
    amount = pivot.get_field("Amount").cache_field_index
    pivot.add_data_field(
        formulon.PivotDataFieldInput(
            name="Sum of Amount", field_index=amount, aggregation=formulon.PivotAggregation.SUM
        )
    )
    return pivot


def table_cells(sheet: formulon.Sheet, first_row: int = 0) -> str:
    """A header plus one body row at columns A:B, returning the unqualified A1 range."""
    sheet.set_text(first_row, 0, "Product")
    sheet.set_text(first_row, 1, "Amount")
    sheet.set_text(first_row + 1, 0, "Pen")
    sheet.set_number(first_row + 1, 1, 5)
    return f"A{first_row + 1}:B{first_row + 2}"


def field_names(cls: type) -> Set[str]:
    """Field names of a dataclass, or the keys of a ``TypedDict``."""
    if dataclasses.is_dataclass(cls):
        return {f.name for f in dataclasses.fields(cls)}
    return set(cls.__annotations__)


class SurfaceTestCase(unittest.TestCase):
    def assertHasMembers(self, owner: Any, names: Iterable[str]) -> None:
        for name in names:
            self.assertTrue(hasattr(owner, name), f"{type(owner).__name__} lacks {name}")

    def assertLacksMembers(self, owner: Any, names: Iterable[str]) -> None:
        for name in names:
            self.assertFalse(hasattr(owner, name), f"{type(owner).__name__} still has {name}")

    def assertFieldsRenamed(self, cls: type, new: Iterable[str], old: Iterable[str]) -> None:
        names = field_names(cls)
        for name in new:
            self.assertIn(name, names, f"{cls.__name__} lacks field {name}")
        for name in old:
            self.assertNotIn(name, names, f"{cls.__name__} still has field {name}")


# ---- Renamed members --------------------------------------------------------------


class RenamedMemberTests(SurfaceTestCase):
    def test_sheet_merge_members_are_spelled_update(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            self.assertHasMembers(
                ws,
                (
                    "update_page_setup",
                    "update_page_margins",
                    "update_print_options",
                    "update_header_footer",
                    "update_view",
                    "update_format_defaults",
                    "update_cell_format",
                    "update_sheet_protection",
                ),
            )
            self.assertLacksMembers(
                ws,
                (
                    "set_page_setup",
                    "set_page_margins",
                    "set_print_options",
                    "set_header_footer",
                    "set_view",
                    "set_format_defaults",
                    "set_sheet_format_defaults",
                    "set_cell_format",
                ),
            )

    def test_workbook_merge_members_are_spelled_update(self) -> None:
        with Workbook.create_default() as wb:
            self.assertHasMembers(wb, ("update_iterative", "update_table_style_defaults"))
            self.assertLacksMembers(wb, ("set_iterative", "set_table_style_defaults"))

    def test_pivot_and_cache_merge_members_are_spelled_update(self) -> None:
        with Workbook.create_default() as wb:
            fill_source(wb.sheet(0))
            cache = add_cache(wb, "Sheet1")
            pivot = add_pivot(wb.sheet(0), cache)
            self.assertHasMembers(
                pivot,
                (
                    "update_grand_totals",
                    "update_style_options",
                    "update_options",
                    "update_field_item",
                    "set_field_subtotal_functions",
                    "set_anchor",
                ),
            )
            self.assertLacksMembers(
                pivot,
                (
                    "set_grand_totals",
                    "set_style_options",
                    "set_options",
                    "update_item",
                    "add_field_subtotal_fn",
                    "clear_field_subtotal_fns",
                    "add_field_subtotal_function",
                    "clear_field_subtotal_functions",
                    "add_field_shared_item",
                ),
            )
            self.assertHasMembers(cache, ("update_options", "add_field_shared_item"))
            self.assertLacksMembers(cache, ("set_options",))

    def test_chart_merge_members_are_spelled_update(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            chart = _chart_input.add_chart(ws, _chart_input.COLUMN, "Sheet1!$A$1:$B$3")
            self.assertHasMembers(
                chart,
                (
                    "update_axis",
                    "update_display",
                    "update_series",
                    "update_plot",
                    "update_legend",
                    "update_point",
                    "set_legend",
                    "set_point",
                ),
            )
            self.assertLacksMembers(chart, ("set_axis", "set_display", "set_series", "set_plot"))

    def test_clear_members_that_restore_a_default_are_spelled_reset(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            self.assertHasMembers(ws, ("reset_row_height", "reset_column_width"))
            self.assertLacksMembers(ws, ("clear_row_height", "clear_column_width"))
            self.assertHasMembers(
                ws.ooxml,
                ("reset_cell_xf_index", "reset_column_xf_index", "reset_row_xf_index"),
            )
            self.assertLacksMembers(
                ws.ooxml,
                ("clear_cell_xf_index", "clear_column_style", "clear_row_style", "clear_column_xf_index"),
            )
            self.assertHasMembers(wb, ("reset_theme",))
            self.assertLacksMembers(wb, ("clear_theme", "clear_pinned_now"))

    def test_xf_index_family_is_not_named_style(self) -> None:
        with Workbook.create_default() as wb:
            ooxml = wb.sheet(0).ooxml
            self.assertHasMembers(
                ooxml,
                (
                    "set_column_xf_index",
                    "get_column_xf_index",
                    "set_row_xf_index",
                    "get_row_xf_index",
                    "get_range_xf_bases",
                    "get_range_xf_bases_excluding",
                    "get_cell_xf_indices",
                    "get_effective_style",
                ),
            )
            self.assertLacksMembers(
                ooxml,
                ("get_column_style", "get_row_style", "get_range_style_bases", "get_range_style_bases_excluding"),
            )

    def test_workbook_accessors_follow_the_final_names(self) -> None:
        with Workbook.create_default() as wb:
            self.assertHasMembers(wb, ("active_sheet", "get_default_font"))
            self.assertLacksMembers(wb, ("get_active_sheet",))
            self.assertHasMembers(wb.ooxml, ("has_passthrough", "get_number_format", "add_number_format"))
            self.assertLacksMembers(wb.ooxml, ("get_passthrough", "get_num_fmt_string", "add_num_fmt"))

    def test_cell_style_registry_lives_on_workbook(self) -> None:
        with Workbook.create_default() as wb:
            names = ("list_cell_styles", "get_cell_style", "set_cell_style", "remove_cell_style")
            self.assertHasMembers(wb, names)
            self.assertLacksMembers(wb.ooxml, names)

    def test_hyperlink_and_break_adders_are_gone(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            self.assertHasMembers(ws, ("set_hyperlink", "set_row_break", "set_column_break"))
            self.assertLacksMembers(
                ws,
                (
                    "add_hyperlink",
                    "add_hyperlink_range",
                    "set_hyperlink_range",
                    "add_column_break",
                    "add_col_break",
                    "remove_row_break",
                    "remove_column_break",
                ),
            )

    def test_table_members_moved_from_workbook_to_sheet(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            self.assertHasMembers(ws, ("add_table", "table", "list_tables", "remove_table"))
            self.assertLacksMembers(wb, ("add_table", "list_tables", "remove_table"))
            self.assertHasMembers(wb, ("table",))

    def test_table_handle_members_follow_the_final_names(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = ws.add_table(formulon.TableInput(range=table_cells(ws), name="Rev"))
            self.assertHasMembers(
                table, ("clear_auto_filter_criteria", "get_style_options", "update_style_options", "rename")
            )
            self.assertLacksMembers(table, ("clear_auto_filter", "display_name"))

    def test_sheet_clears_auto_filter_criteria(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            self.assertHasMembers(ws, ("clear_auto_filter_criteria",))
            self.assertLacksMembers(ws, ("clear_auto_filter", "remove_auto_filter"))

    def test_differential_format_is_dxf(self) -> None:
        self.assertTrue(hasattr(formulon, "Dxf"))
        self.assertFalse(hasattr(formulon, "DifferentialFormat"))
        self.assertFieldsRenamed(
            formulon.Dxf,
            ("number_format", "alignment", "protection"),
            ("num_fmt_code", "num_fmt_id", "alignment_xml", "protection_xml"),
        )

    def test_pivot_style_and_table_style_records_are_named_style_options(self) -> None:
        self.assertTrue(hasattr(formulon, "PivotStyleOptions"))
        self.assertTrue(hasattr(formulon, "TableStyleOptions"))
        self.assertFalse(hasattr(formulon, "PivotStyle"))
        self.assertFieldsRenamed(formulon.PivotStyleOptions, ("style_name",), ("name",))
        self.assertFieldsRenamed(formulon.TableStyleOptions, ("style_name",), ())

    def test_picture_insertion_is_add_picture(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            self.assertHasMembers(ws, ("add_picture",))
            self.assertLacksMembers(ws, ("insert_image", "add_image"))

    def test_restore_drawing_takes_assign_new_id(self) -> None:
        import inspect

        with Workbook.create_default() as wb:
            parameters = inspect.signature(wb.sheet(0).restore_drawing).parameters
            self.assertIn("assign_new_id", parameters)
            self.assertNotIn("new_id", parameters)

    def test_function_metadata_names_its_signature(self) -> None:
        self.assertFieldsRenamed(formulon.FunctionMetadata, ("signature",), ("signature_template",))


# ---- update_* merges and keeps UNCHANGED ------------------------------------------


class UpdateMergeTests(SurfaceTestCase):
    def test_update_page_setup_changes_only_the_given_fields(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_page_setup(orientation=formulon.PageOrientation.LANDSCAPE, paper_size=9)
            ws.update_page_setup(scale_pct=120)
            setup = ws.get_page_setup()
            self.assertEqual(setup.scale_pct, 120)
            self.assertTrue(setup.scale_pct_stated)
            self.assertEqual(setup.orientation, formulon.PageOrientation.LANDSCAPE)
            self.assertEqual(setup.paper_size, 9)

    def test_update_page_setup_rejects_the_old_scale_keyword(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.sheet(0).update_page_setup(scale=120)  # type: ignore[call-arg]

    def test_update_page_margins_takes_inch_keywords_and_keeps_the_rest(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_page_margins(left_in=0.4, right_in=0.5)
            ws.update_page_margins(left_in=0.9)
            margins = ws.get_page_margins()
            self.assertAlmostEqual(margins.left_in, 0.9)
            self.assertAlmostEqual(margins.right_in, 0.5)
            self.assertTrue(margins.left_in_stated)
            self.assertTrue(margins.right_in_stated)

    def test_update_view_keeps_the_fields_left_unchanged(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_view(freeze_rows=2, freeze_cols=1)
            ws.update_view(zoom_pct=130)
            view = ws.get_view()
            self.assertEqual(view.zoom_pct, 130)
            self.assertEqual((view.freeze_rows, view.freeze_cols), (2, 1))

    def test_update_format_defaults_uses_unit_suffixed_keywords(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_format_defaults(default_column_width_chars=12.0, base_column_width_chars=9.0)
            ws.update_format_defaults(default_row_height_pt=20.0)
            defaults = ws.get_format_defaults()
            self.assertAlmostEqual(defaults.default_column_width_chars, 12.0)
            self.assertAlmostEqual(defaults.base_column_width_chars, 9.0)
            self.assertAlmostEqual(defaults.default_row_height_pt, 20.0)
            self.assertGreater(defaults.effective_default_column_width_chars, 0.0)
            self.assertGreater(defaults.effective_default_row_height_pt, 0.0)

    def test_update_format_defaults_none_clears_the_stored_value(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_format_defaults(default_row_height_pt=20.0, default_column_width_chars=12.0)
            ws.update_format_defaults(default_row_height_pt=None)
            defaults = ws.get_format_defaults()
            self.assertIsNone(defaults.default_row_height_pt)
            self.assertAlmostEqual(defaults.default_column_width_chars, 12.0)

    def test_update_iterative_keeps_the_settings_left_unchanged(self) -> None:
        with Workbook.create_default() as wb:
            wb.update_iterative(enabled=True, max_iterations=50)
            wb.update_iterative(max_change=0.01)
            settings = wb.get_iterative()
            self.assertTrue(settings.enabled)
            self.assertEqual(settings.max_iterations, 50)
            self.assertAlmostEqual(settings.max_change, 0.01)

    def test_update_table_style_defaults_keeps_the_name_left_unchanged(self) -> None:
        with Workbook.create_default() as wb:
            wb.update_table_style_defaults(table_style_name="TableStyleLight9", pivot_style_name="PivotStyleLight16")
            wb.update_table_style_defaults(pivot_style_name=None)
            defaults = wb.get_table_style_defaults()
            self.assertEqual(defaults.table_style_name, "TableStyleLight9")
            self.assertIsNone(defaults.pivot_style_name)

    def test_update_grand_totals_keeps_the_other_band(self) -> None:
        with Workbook.create_default() as wb:
            fill_source(wb.sheet(0))
            pivot = add_pivot(wb.sheet(0), add_cache(wb, "Sheet1"))
            pivot.update_grand_totals(rows=False, columns=True)
            pivot.update_grand_totals(columns=False)
            totals = pivot.get_grand_totals()
            self.assertFalse(totals.rows)
            self.assertFalse(totals.columns)

    def test_update_style_options_uses_style_name_and_keeps_the_rest(self) -> None:
        with Workbook.create_default() as wb:
            fill_source(wb.sheet(0))
            pivot = add_pivot(wb.sheet(0), add_cache(wb, "Sheet1"))
            pivot.update_style_options(style_name="PivotStyleMedium9", show_row_stripes=True)
            pivot.update_style_options(show_last_column=True)
            info = pivot.get_style_options()
            self.assertIsNotNone(info)
            assert info is not None
            self.assertEqual(info.style_name, "PivotStyleMedium9")
            self.assertTrue(info.show_row_stripes)
            self.assertTrue(info.show_last_column)

    def test_update_options_keeps_the_options_left_unchanged(self) -> None:
        with Workbook.create_default() as wb:
            fill_source(wb.sheet(0))
            pivot = add_pivot(wb.sheet(0), add_cache(wb, "Sheet1"))
            pivot.update_options(show_error=True, error_caption="oops")
            pivot.update_options(show_missing=True)
            options = pivot.get_options()
            self.assertTrue(options.show_error)
            self.assertEqual(options.error_caption, "oops")
            self.assertTrue(options.show_missing)

    def test_chart_update_axis_hides_an_axis_and_keeps_the_rest(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            chart = _chart_input.add_chart(ws, _chart_input.COLUMN, "Sheet1!$A$1:$B$3")
            axes = chart.get_properties().axes
            axis_id = axes[0].id
            chart.update_axis(axis_id, visible=False)
            hidden = [a for a in chart.get_properties().axes if a.id == axis_id][0]
            self.assertFalse(hidden.visible)
            chart.update_axis(axis_id, visible=True)
            shown = [a for a in chart.get_properties().axes if a.id == axis_id][0]
            self.assertTrue(shown.visible)

    def test_chart_update_plot_uses_unit_suffixed_keywords(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            chart = _chart_input.add_chart(ws, _chart_input.COLUMN, "Sheet1!$A$1:$B$3")
            chart.update_plot(0, gap_width_pct=80)
            chart.update_plot(0, overlap_pct=10)
            plot = chart.get_properties().plots[0]
            self.assertEqual(plot.gap_width_pct, 80)
            self.assertEqual(plot.overlap_pct, 10)
            with self.assertRaises(TypeError):
                chart.update_plot(0, gap_width=80)  # type: ignore[call-arg]

    def test_chart_update_legend_merges_and_set_legend_replaces_or_removes(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            chart = _chart_input.add_chart(
                ws, _chart_input.COLUMN, "Sheet1!$A$1:$B$3", legend=ChartLegendPosition.RIGHT
            )
            chart.update_legend(overlay=True)
            legend = chart.get_properties().legend
            assert legend is not None
            self.assertTrue(legend.overlay)
            self.assertEqual(legend.position, ChartLegendPosition.RIGHT)

            chart.set_legend(ChartLegend(position=ChartLegendPosition.TOP, overlay=False))
            legend = chart.get_properties().legend
            assert legend is not None
            self.assertEqual(legend.position, ChartLegendPosition.TOP)
            self.assertFalse(legend.overlay)

            chart.set_legend(None)
            self.assertIsNone(chart.get_properties().legend)
            chart.set_legend(None)
            self.assertIsNone(chart.get_properties().legend)

    def test_chart_set_point_none_is_a_no_op_when_no_override_exists(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            chart = _chart_input.add_chart(ws, _chart_input.COLUMN, "Sheet1!$A$1:$B$3")
            chart.set_point(0, 0, 0, None)

    def test_update_cell_format_merges_groups_into_the_existing_format(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_cell_format("A1", number_format="0.00")
            ws.update_cell_format("A1", font={"bold": True})
            fmt = ws.get_cell_format("A1")
            self.assertEqual(fmt.number_format, "0.00")
            self.assertTrue(fmt.font.bold)

    def test_unspecified_keywords_default_to_the_unchanged_sentinel(self) -> None:
        import inspect

        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for method in (
                ws.update_page_setup,
                ws.update_page_margins,
                ws.update_print_options,
                ws.update_header_footer,
                ws.update_view,
                ws.update_format_defaults,
                ws.update_sheet_protection,
                wb.update_iterative,
                wb.update_table_style_defaults,
            ):
                parameters = list(inspect.signature(method).parameters.values())
                self.assertTrue(parameters, method.__name__)
                for parameter in parameters:
                    self.assertEqual(parameter.kind, inspect.Parameter.KEYWORD_ONLY, method.__name__)
                    self.assertIs(parameter.default, formulon.UNCHANGED, f"{method.__name__}.{parameter.name}")


# ---- Singleton set_*(None) removal -----------------------------------------------


class SingletonRemovalTests(SurfaceTestCase):
    def test_set_sheet_protection_none_unprotects(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_sheet_protection(SheetProtection(allow_format_cells=True))
            self.assertTrue(ws.get_sheet_protection().allow_format_cells)  # type: ignore[union-attr]
            ws.set_sheet_protection(None)
            self.assertIsNone(ws.get_sheet_protection())

    def test_set_sheet_protection_none_is_a_no_op_when_unprotected(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_sheet_protection(None)
            self.assertIsNone(ws.get_sheet_protection())

    def test_set_sheet_protection_replaces_rather_than_merges(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_sheet_protection(SheetProtection(allow_sort=True, allow_format_rows=True))
            ws.set_sheet_protection(SheetProtection(allow_auto_filter=True))
            protection = ws.get_sheet_protection()
            self.assertTrue(protection.allow_auto_filter)
            self.assertFalse(protection.allow_sort)
            self.assertFalse(protection.allow_format_rows)

    def test_update_sheet_protection_merges_into_the_existing_flags(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_sheet_protection(SheetProtection(allow_sort=True))
            ws.update_sheet_protection(allow_insert_rows=True, allow_select_locked_cells=False)
            protection = ws.get_sheet_protection()
            self.assertTrue(protection.allow_sort)
            self.assertTrue(protection.allow_insert_rows)
            self.assertFalse(protection.allow_select_locked_cells)
            self.assertTrue(protection.allow_select_unlocked_cells)

    def test_set_auto_filter_none_removes_the_filter_and_is_a_no_op_when_absent(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            ws.set_auto_filter(None)
            self.assertIsNone(ws.get_auto_filter())
            ws.set_auto_filter(formulon.AutoFilter(range=CellRect(0, 0, 2, 1)))
            self.assertIsNotNone(ws.get_auto_filter())
            ws.set_auto_filter(None)
            self.assertIsNone(ws.get_auto_filter())

    def test_clear_auto_filter_criteria_keeps_the_filter(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            ws.set_auto_filter(formulon.AutoFilter(range=CellRect(0, 0, 2, 1)))
            ws.clear_auto_filter_criteria()
            auto_filter = ws.get_auto_filter()
            assert auto_filter is not None
            self.assertEqual(auto_filter.columns, [])

    def test_clear_auto_filter_criteria_without_a_filter_is_not_found(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.NotFoundError):
                wb.sheet(0).clear_auto_filter_criteria()

    def test_table_set_auto_filter_none_removes_the_filter(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = ws.add_table(formulon.TableInput(range=table_cells(ws), name="Rev"))
            table.set_auto_filter(None)
            self.assertIsNone(table.get_auto_filter())
            table.set_auto_filter(None)
            self.assertIsNone(table.get_auto_filter())

    def test_set_pinned_now_none_unpins(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_pinned_now(datetime.datetime(2024, 5, 6, 7, 8, 9))
            self.assertIsNotNone(wb.get_pinned_now())
            wb.set_pinned_now(None)
            self.assertIsNone(wb.get_pinned_now())
            wb.set_pinned_now(None)
            self.assertIsNone(wb.get_pinned_now())

    def test_reset_theme_restores_the_default_theme(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_theme_colors(ThemeColorsInput(colors=[0xFF000000 | i for i in range(12)]))
            self.assertEqual(wb.get_theme_colors().source, ThemeSource.PART)
            wb.reset_theme()
            self.assertEqual(wb.get_theme_colors().source, ThemeSource.DEFAULT)

    def test_set_hyperlink_none_removes_and_a_same_range_link_replaces_the_first(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            cell = CellRect(0, 0, 0, 0)
            ws.set_hyperlink(
                "A1",
                formulon.Hyperlink(
                    range=cell, target="https://example.com/a", location=None, display=None, tooltip=None
                ),
            )
            ws.set_hyperlink(
                "A1",
                formulon.Hyperlink(
                    range=cell, target="https://example.com/b", location=None, display=None, tooltip=None
                ),
            )
            link = ws.get_hyperlink("A1")
            assert link is not None
            self.assertEqual(link.target, "https://example.com/b")
            # A link over a different range from the same cell is appended.
            ws.set_hyperlink(
                "A1",
                formulon.Hyperlink(
                    range=CellRect(0, 0, 1, 1),
                    target="https://example.com/c",
                    location=None,
                    display=None,
                    tooltip=None,
                ),
            )
            self.assertEqual(
                [h.target for h in ws.list_hyperlinks()], ["https://example.com/b", "https://example.com/c"]
            )
            ws.set_hyperlink("A1", None)
            self.assertIsNone(ws.get_hyperlink("A1"))
            self.assertEqual(ws.list_hyperlinks(), [])
            ws.set_hyperlink("A1", None)


# ---- SheetProtection and DataValidation defaults ---------------------------------


class ProtectionAndValidationDefaultsTests(SurfaceTestCase):
    def test_sheet_protection_fields_are_allow_flags_without_enabled_or_sheet(self) -> None:
        allow = {
            "allow_edit_objects",
            "allow_edit_scenarios",
            "allow_format_cells",
            "allow_format_columns",
            "allow_format_rows",
            "allow_insert_columns",
            "allow_insert_rows",
            "allow_insert_hyperlinks",
            "allow_delete_columns",
            "allow_delete_rows",
            "allow_select_locked_cells",
            "allow_select_unlocked_cells",
            "allow_sort",
            "allow_auto_filter",
            "allow_pivot_tables",
        }
        names = field_names(SheetProtection)
        self.assertTrue(allow <= names)
        for old in (
            "enabled",
            "sheet",
            "objects",
            "scenarios",
            "format_cells",
            "format_columns",
            "format_rows",
            "insert_columns",
            "insert_rows",
            "insert_hyperlinks",
            "delete_columns",
            "delete_rows",
            "select_locked_cells",
            "select_unlocked_cells",
            "sort",
            "auto_filter",
            "pivot_tables",
            "legacy_password",
        ):
            self.assertNotIn(old, names)
        self.assertIn("legacy_password_hash", names)

    def test_sheet_protection_defaults_match_the_excel_dialog(self) -> None:
        default = SheetProtection()
        self.assertTrue(default.allow_select_locked_cells)
        self.assertTrue(default.allow_select_unlocked_cells)
        for flag in (
            "allow_edit_objects",
            "allow_edit_scenarios",
            "allow_format_cells",
            "allow_format_columns",
            "allow_format_rows",
            "allow_insert_columns",
            "allow_insert_rows",
            "allow_insert_hyperlinks",
            "allow_delete_columns",
            "allow_delete_rows",
            "allow_sort",
            "allow_auto_filter",
            "allow_pivot_tables",
        ):
            self.assertIs(getattr(default, flag), False, flag)

    def test_sheet_protection_round_trips_the_legacy_password_hash(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_sheet_protection(SheetProtection(legacy_password_hash="CC1A"))
            self.assertEqual(ws.get_sheet_protection().legacy_password_hash, "CC1A")

    def test_validation_input_allows_blank_and_shows_the_dropdown_by_default(self) -> None:
        item = formulon.DataValidationInput(type=formulon.ValidationType.LIST, ranges="B1:B10", formula1='"a,b"')
        self.assertIs(item.allow_blank, True)
        self.assertIs(item.show_in_cell_dropdown, True)
        self.assertFalse(hasattr(item, "show_dropdown"))

    def test_added_validation_reads_back_with_the_dialog_defaults(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.add_validation(
                formulon.DataValidationInput(type=formulon.ValidationType.LIST, ranges="B1:B10", formula1='"a,b"')
            )
            rule = ws.list_validations()[0]
            self.assertTrue(rule.allow_blank)
            self.assertTrue(rule.show_in_cell_dropdown)

    def test_validation_explicit_false_is_kept(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.add_validation(
                formulon.DataValidationInput(
                    type=formulon.ValidationType.LIST,
                    ranges="B1:B10",
                    formula1='"a,b"',
                    allow_blank=False,
                    show_in_cell_dropdown=False,
                )
            )
            rule = ws.list_validations()[0]
            self.assertFalse(rule.allow_blank)
            self.assertFalse(rule.show_in_cell_dropdown)

    def test_operator_fields_are_spelled_out(self) -> None:
        for cls in (
            formulon.DataValidation,
            formulon.DataValidationInput,
            formulon.ConditionalFormat,
            formulon.CfCellIsRuleInput,
        ):
            self.assertFieldsRenamed(cls, ("operator",), ("op",))
        self.assertFieldsRenamed(formulon.DataValidation, ("show_in_cell_dropdown",), ("show_dropdown",))


# ---- Unit-suffixed fields ---------------------------------------------------------


class UnitSuffixTests(SurfaceTestCase):
    def test_margin_fields_are_inches(self) -> None:
        inches = {f"{side}_in" for side in ("left", "right", "top", "bottom", "header", "footer")}
        self.assertTrue(inches <= field_names(formulon.PageMargins))
        self.assertTrue({f"{name}_stated" for name in inches} <= field_names(formulon.PageMargins))
        self.assertTrue(
            {f"{side}_in" for side in ("left", "right", "top", "bottom", "header", "footer")}
            <= field_names(formulon.PaginationMargins)
        )
        for side in ("left", "right", "top", "bottom", "header", "footer"):
            self.assertNotIn(side, field_names(formulon.PageMargins))
            self.assertNotIn(side, field_names(formulon.PaginationMargins))

    def test_pagination_reports_percent_scale_and_inch_margins(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            ws.update_page_margins(left_in=0.5)
            pagination = ws.get_pagination()
            self.assertEqual(pagination.scale_pct, 100)
            self.assertAlmostEqual(pagination.margins.left_in, 0.5)
            ws.update_page_setup(scale_pct=50)
            self.assertEqual(ws.get_pagination().scale_pct, 50)
            self.assertNotIn("scale", field_names(formulon.Pagination))

    def test_page_setup_scale_is_a_percent(self) -> None:
        self.assertFieldsRenamed(formulon.PageSetup, ("scale_pct", "scale_pct_stated"), ("scale", "scale_stated"))

    def test_rect_and_layout_fields_carry_their_unit(self) -> None:
        self.assertFieldsRenamed(formulon.RectPt, ("width_pt", "height_pt"), ("width", "height"))
        self.assertFieldsRenamed(formulon.ColumnLayout, ("width_chars",), ("width",))
        self.assertFieldsRenamed(formulon.RowLayout, ("height_pt",), ("height",))
        self.assertFieldsRenamed(formulon.WidthModel, ("normal_font_size_pt",), ("normal_font_size",))
        self.assertFieldsRenamed(
            formulon.SheetView, ("zoom_pct", "split_x_twips", "split_y_twips"), ("zoom_scale", "split_x", "split_y")
        )

    def test_sheet_format_defaults_fields_carry_their_unit(self) -> None:
        self.assertFieldsRenamed(
            formulon.SheetFormatDefaults,
            (
                "default_column_width_chars",
                "base_column_width_chars",
                "effective_default_column_width_chars",
                "default_row_height_pt",
                "effective_default_row_height_pt",
            ),
            (
                "default_col_width",
                "base_col_width",
                "effective_default_col_width",
                "default_row_height",
                "effective_default_row_height",
            ),
        )

    def test_set_column_width_and_row_height_take_unit_keywords(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_column_width(0, 0, width_chars=12.0)
            ws.set_row_height(2, 2, height_pt=30.0)
            column = ws.list_column_overrides()[0]
            self.assertIsNotNone(column.width_chars)
            self.assertGreaterEqual(column.width_chars, 12.0)
            self.assertLess(column.width_chars, 14.0)
            row = ws.list_row_overrides()[0]
            self.assertAlmostEqual(row.height_pt, 30.0)
            with self.assertRaises(TypeError):
                ws.set_row_height(3, 3, height=30.0)  # type: ignore[call-arg]

    def test_set_row_auto_height_takes_a_points_keyword(self) -> None:
        import inspect

        with Workbook.create_default() as wb:
            parameters = inspect.signature(wb.sheet(0).set_row_auto_height).parameters
            self.assertIn("height_pt", parameters)
            self.assertNotIn("height", parameters)

    def test_reset_row_height_and_column_width_restore_the_default(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_column_width(0, 0, width_chars=12.0)
            ws.set_row_height(2, 2, height_pt=30.0)
            ws.reset_column_width(0, 0)
            ws.reset_row_height(2, 2)
            self.assertEqual(ws.list_column_overrides(), [])
            self.assertEqual(ws.list_row_overrides(), [])

    def test_image_pixel_size_is_width_px_and_height_px(self) -> None:
        self.assertFieldsRenamed(formulon.PictureProperties, ("width_px", "height_px"), ("px_width", "px_height"))
        self.assertFieldsRenamed(formulon.ImageInfo, ("width_px", "height_px"), ("px_width", "px_height"))

    def test_shape_fields_carry_degrees_and_points(self) -> None:
        self.assertFieldsRenamed(formulon.ShapeProperties, ("rotation_deg",), ("rotation", "line_width_emu"))
        self.assertFieldsRenamed(formulon.ChartLine, ("width_pt",), ("width_emu",))

    def test_shape_update_takes_rotation_degrees(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            shape = ws.add_shape(
                formulon.ShapeInput(
                    preset_geometry="rect",
                    anchor=formulon.DrawingAnchorInput(
                        kind=formulon.AnchorKind.ONE_CELL, from_row=2, width_emu=914400, height_emu=457200
                    ),
                    rotation_deg=45.0,
                )
            )
            self.assertAlmostEqual(shape.get_properties().rotation_deg, 45.0)
            shape.update(rotation_deg=90.0)
            self.assertAlmostEqual(shape.get_properties().rotation_deg, 90.0)

    def test_chart_fields_carry_percent_degrees_and_points(self) -> None:
        self.assertFieldsRenamed(formulon.ChartFill, ("transparency_pct",), ("transparency",))
        self.assertFieldsRenamed(formulon.ChartLine, ("transparency_pct",), ("transparency",))
        self.assertFieldsRenamed(formulon.ChartMarker, ("size_pt",), ("size",))
        self.assertFieldsRenamed(
            formulon.ChartPlot,
            ("gap_width_pct", "first_slice_angle_deg"),
            ("gap_width", "first_slice_angle"),
        )

    def test_font_size_is_in_points(self) -> None:
        self.assertFieldsRenamed(formulon.Font, ("size_pt",), ("size",))
        self.assertFieldsRenamed(formulon.FontPatch, ("size_pt",), ("size",))

    def test_chart_axis_is_visible_not_deleted(self) -> None:
        self.assertFieldsRenamed(formulon.ChartAxis, ("visible",), ("deleted",))

    def test_phonetic_font_and_icon_and_dxf_references_are_indexes(self) -> None:
        self.assertFieldsRenamed(formulon.PhoneticProperties, ("font_index",), ("font_id",))
        self.assertFieldsRenamed(formulon.ConditionalFormat, ("dxf_index",), ("dxf_id",))
        self.assertFieldsRenamed(formulon.CfExpressionRuleInput, ("dxf_index",), ("dxf_id",))
        self.assertFieldsRenamed(formulon.CellXf, ("style_xf_index",), ("xf_id",))
        self.assertFieldsRenamed(formulon.CfIcon, ("icon_index",), ("icon_id",))
        self.assertFieldsRenamed(formulon.TableStyleElement, ("dxf_index",), ("dxf_id",))

    def test_protected_range_and_workbook_protection_name_the_password_hash(self) -> None:
        self.assertFieldsRenamed(formulon.ProtectedRange, ("legacy_password_hash",), ("password",))
        self.assertFieldsRenamed(
            formulon.WorkbookProtection,
            ("workbook_legacy_password_hash", "revisions_legacy_password_hash"),
            ("workbook_password", "revisions_password"),
        )


# ---- Renamed record fields --------------------------------------------------------


class RenamedFieldTests(SurfaceTestCase):
    def test_auto_filter_column_spells_out_its_attributes(self) -> None:
        self.assertFieldsRenamed(
            formulon.AutoFilterColumn,
            (
                "column_offset",
                "top_value",
                "filter_value",
                "dynamic_value",
                "dynamic_max_value",
                "value_iso",
                "max_value_iso",
                "dxf_index",
                "icon_index",
            ),
            (
                "col_id",
                "op1",
                "val1",
                "op2",
                "val2",
                "top_val",
                "filter_val",
                "dyn_val",
                "dyn_max_val",
                "val_iso",
                "max_val_iso",
                "dxf_id",
                "icon_id",
            ),
        )

    def test_cell_alignment_fields_drop_the_align_suffix(self) -> None:
        self.assertFieldsRenamed(formulon.CellXf, ("horizontal", "vertical"), ("horizontal_align", "vertical_align"))
        self.assertFieldsRenamed(
            formulon.CellXfInput,
            ("horizontal", "vertical"),
            ("horizontal_align", "vertical_align", "has_horizontal_align", "has_vertical_align"),
        )

    def test_fill_and_color_fields_are_spelled_out(self) -> None:
        for cls in (formulon.Fill, formulon.FillPatch):
            self.assertFieldsRenamed(cls, ("foreground", "background"), ("fg", "bg"))
        self.assertFieldsRenamed(formulon.Color, ("argb",), ("rgb",))

    def test_effective_style_names_its_fields(self) -> None:
        self.assertFieldsRenamed(
            formulon.EffectiveStyle,
            ("font_color", "number_format"),
            ("font", "num_fmt_code", "locked", "hidden"),
        )

    def test_image_format_is_a_picture_field_and_not_a_common_drawing_field(self) -> None:
        self.assertFieldsRenamed(formulon.PictureProperties, ("format",), ("image_format", "media_path"))
        self.assertFieldsRenamed(formulon.DrawingProperties, (), ("format", "image_format", "media_path"))

    def test_phonetic_run_spans_start_and_end(self) -> None:
        self.assertFieldsRenamed(formulon.PhoneticRun, ("start", "end"), ("sb", "eb"))

    def test_pivot_field_lists_subtotal_functions(self) -> None:
        self.assertFieldsRenamed(formulon.PivotField, ("subtotal_functions",), ("subtotal_fns",))

    def test_table_properties_hold_one_name_and_no_style_fields(self) -> None:
        self.assertFieldsRenamed(
            formulon.TableProperties,
            ("name",),
            (
                "display_name",
                "style_name",
                "show_first_column",
                "show_last_column",
                "show_row_stripes",
                "show_column_stripes",
            ),
        )

    def test_dxf_number_format_and_alignment_are_typed(self) -> None:
        dxf = formulon.Dxf(number_format="0.0")
        self.assertEqual(dxf.number_format, "0.0")
        self.assertIsNone(dxf.alignment)
        self.assertIsNone(dxf.protection)


# ---- Page breaks ------------------------------------------------------------------


class PageBreakTests(SurfaceTestCase):
    def test_row_and_column_break_records_name_their_span(self) -> None:
        self.assertEqual(field_names(formulon.RowBreak), {"row", "first", "last", "manual"})
        self.assertEqual(field_names(formulon.ColumnBreak), {"col", "first", "last", "manual"})
        self.assertFalse(hasattr(formulon, "PageBreak"))

    def test_set_row_break_upserts_and_none_removes(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            ws.set_row_break(5, formulon.RowBreak(row=5))
            breaks = [b for b in ws.list_row_breaks() if b.row == 5]
            self.assertEqual(len(breaks), 1)
            self.assertTrue(breaks[0].manual)

            ws.set_row_break(5, formulon.RowBreak(row=5, first=0, last=3, manual=False))
            breaks = [b for b in ws.list_row_breaks() if b.row == 5]
            self.assertEqual(len(breaks), 1)
            self.assertEqual((breaks[0].first, breaks[0].last, breaks[0].manual), (0, 3, False))

            ws.set_row_break(5, None)
            self.assertEqual([b for b in ws.list_row_breaks() if b.row == 5], [])
            ws.set_row_break(5, None)

    def test_set_column_break_upserts_and_none_removes(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            ws.set_column_break(1, formulon.ColumnBreak(col=1))
            breaks = [b for b in ws.list_column_breaks() if b.col == 1]
            self.assertEqual(len(breaks), 1)
            self.assertIsInstance(breaks[0], formulon.ColumnBreak)
            ws.set_column_break(1, None)
            self.assertEqual([b for b in ws.list_column_breaks() if b.col == 1], [])
            ws.set_column_break(1, None)

    def test_pagination_lists_typed_breaks_without_parallel_manual_lists(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            fill_source(ws)
            ws.set_row_break(4, formulon.RowBreak(row=4))
            ws.set_column_break(1, formulon.ColumnBreak(col=1))
            pagination = ws.get_pagination()
            self.assertTrue(all(isinstance(b, formulon.RowBreak) for b in pagination.row_breaks))
            self.assertTrue(all(isinstance(b, formulon.ColumnBreak) for b in pagination.column_breaks))
            self.assertIn(4, [b.row for b in pagination.row_breaks])
            self.assertIn(1, [b.col for b in pagination.column_breaks])
            self.assertFalse(hasattr(pagination, "row_break_manual"))
            self.assertFalse(hasattr(pagination, "column_break_manual"))


# ---- Table lifecycle on Sheet -----------------------------------------------------


class SheetTableLifecycleTests(SurfaceTestCase):
    def test_add_table_returns_a_handle_found_by_name_index_and_workbook_lookup(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = ws.add_table(formulon.TableInput(range=table_cells(ws), name="Rev"))
            self.assertIsInstance(table, formulon.Table)
            self.assertEqual(table.name, "Rev")
            self.assertIs(ws.table("Rev"), table)
            self.assertIs(ws.table(0), table)
            self.assertIs(ws.table(table), table)
            self.assertIs(ws.list_tables()[0], table)
            self.assertIs(wb.table("Rev"), table)

    def test_table_indexes_are_local_to_the_sheet(self) -> None:
        with Workbook.create_default() as wb:
            first = wb.sheet(0)
            second = wb.add_sheet("Other")
            a = first.add_table(formulon.TableInput(range=table_cells(first), name="TblA"))
            b = second.add_table(formulon.TableInput(range=table_cells(second), name="TblB"))
            self.assertEqual(a.index, 0)
            self.assertEqual(b.index, 0)
            self.assertEqual([t.name for t in first.list_tables()], ["TblA"])
            self.assertEqual([t.name for t in second.list_tables()], ["TblB"])
            self.assertIs(wb.table("TblB"), b)

    def test_remove_table_detaches_the_handle(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = ws.add_table(formulon.TableInput(range=table_cells(ws), name="Rev"))
            ws.remove_table(table)
            self.assertEqual(ws.list_tables(), [])
            with self.assertRaises(NotFoundError):
                wb.table("Rev")
            with self.assertRaises(NotFoundError):
                table.get_properties()

    def test_remove_table_of_an_unknown_name_raises_not_found(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(NotFoundError):
                wb.sheet(0).remove_table("Missing")

    def test_table_has_one_name_and_rename_changes_it(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = ws.add_table(formulon.TableInput(range=table_cells(ws), name="Rev"))
            table.rename("Sales")
            self.assertEqual(table.name, "Sales")
            self.assertEqual(table.get_properties().name, "Sales")
            self.assertIs(wb.table("Sales"), table)
            with self.assertRaises(NotFoundError):
                wb.table("Rev")

    def test_add_table_without_a_style_applies_medium2_with_row_stripes(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = ws.add_table(formulon.TableInput(range=table_cells(ws), name="Rev"))
            info = table.get_style_options()
            assert info is not None
            self.assertEqual(info.style_name, "TableStyleMedium2")
            self.assertTrue(info.show_row_stripes)

    def test_add_table_takes_the_workbook_table_style_default(self) -> None:
        with Workbook.create_default() as wb:
            wb.update_table_style_defaults(table_style_name="TableStyleLight9")
            ws = wb.sheet(0)
            table = ws.add_table(formulon.TableInput(range=table_cells(ws), name="Rev"))
            info = table.get_style_options()
            assert info is not None
            self.assertEqual(info.style_name, "TableStyleLight9")

    def test_update_style_options_merges_into_the_table_style(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = ws.add_table(formulon.TableInput(range=table_cells(ws), name="Rev"))
            table.update_style_options(show_column_stripes=True)
            info = table.get_style_options()
            assert info is not None
            self.assertEqual(info.style_name, "TableStyleMedium2")
            self.assertTrue(info.show_row_stripes)
            self.assertTrue(info.show_column_stripes)
            table.update_style_options(style_name="TableStyleLight9")
            info = table.get_style_options()
            assert info is not None
            self.assertEqual(info.style_name, "TableStyleLight9")
            self.assertTrue(info.show_column_stripes)

    def test_add_table_rejects_a_separate_display_name(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            with self.assertRaises(TypeError):
                ws.add_table(
                    formulon.TableInput(range=table_cells(ws), name="Rev", display_name="Other")  # type: ignore[call-arg]
                )


# ---- Pivot behaviour --------------------------------------------------------------


class PivotFieldItemTests(SurfaceTestCase):
    def _pivot(self, wb: Workbook) -> formulon.Pivot:
        fill_source(wb.sheet(0))
        return add_pivot(wb.sheet(0), add_cache(wb, "Sheet1"))

    def test_add_field_item_by_label(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._pivot(wb)
            pivot.clear_field_items(0)
            index = pivot.add_field_item(0, "East")
            self.assertEqual(pivot.get_field_item(0, index).name, "East")
            with self.assertRaises(AlreadyExistsError):
                pivot.add_field_item(0, "East")
            self.assertEqual(len(pivot.list_field_items(0)), 1)

    def test_add_field_item_by_cache_index(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._pivot(wb)
            pivot.clear_field_items(0)
            index = pivot.add_field_item(0, cache_index=1)
            item = pivot.get_field_item(0, index)
            self.assertEqual(item.cache_index, 1)
            self.assertEqual(item.name, "West")

    def test_add_field_item_requires_exactly_one_of_name_and_cache_index(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._pivot(wb)
            with self.assertRaises(TypeError):
                pivot.add_field_item(0)  # type: ignore[call-overload]
            with self.assertRaises(TypeError):
                pivot.add_field_item(0, "East", cache_index=0)  # type: ignore[call-overload]

    def test_add_field_item_visible_is_keyword_only(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._pivot(wb)
            pivot.clear_field_items(0)
            index = pivot.add_field_item(0, "East", visible=False)
            self.assertFalse(pivot.get_field_item(0, index).visible)

    def test_update_field_item_changes_only_the_given_attributes(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._pivot(wb)
            index = 0
            pivot.update_field_item(0, index, visible=False)
            pivot.update_field_item(0, index, caption="Eastern")
            item = pivot.get_field_item(0, index)
            self.assertFalse(item.visible)
            self.assertEqual(item.caption, "Eastern")

    def test_subtotal_functions_are_listed_under_the_spelled_out_name(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._pivot(wb)
            pivot.set_field_subtotal_functions(0, [formulon.PivotAggregation.SUM])
            self.assertEqual(pivot.list_fields()[0].subtotal_functions, [formulon.PivotAggregation.SUM])
            pivot.set_field_subtotal_functions(0, [])
            self.assertEqual(pivot.list_fields()[0].subtotal_functions, [])

    def test_set_anchor_takes_one_cell(self) -> None:
        with Workbook.create_default() as wb:
            pivot = self._pivot(wb)
            pivot.set_anchor("H10")
            rect = pivot.get_layout().range
            self.assertEqual((rect.first_row, rect.first_col), (9, 7))
            pivot.set_anchor(formulon.CellAddress(2, 3))
            rect = pivot.get_layout().range
            self.assertEqual((rect.first_row, rect.first_col), (2, 3))


# ---- Workbook accessors -----------------------------------------------------------


class WorkbookAccessorTests(SurfaceTestCase):
    def test_active_sheet_returns_the_sheet_handle(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIs(wb.active_sheet(), wb.sheet(0))

    def test_has_passthrough_answers_with_a_bool(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIs(wb.ooxml.has_passthrough("customXml/absent.xml"), False)

    def test_get_default_font_reads_font_zero(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.get_default_font(), wb.ooxml.get_font(0))

    def test_number_format_round_trips_through_add_and_get(self) -> None:
        with Workbook.create_default() as wb:
            num_fmt_id = wb.ooxml.add_number_format("0.000%")
            self.assertEqual(wb.ooxml.get_number_format(num_fmt_id), "0.000%")


# ---- Enum members -----------------------------------------------------------------


class EnumMemberTests(SurfaceTestCase):
    def test_pivot_cell_kind_spells_column_out(self) -> None:
        self.assertTrue(hasattr(formulon.PivotCellKind, "COLUMN_LABEL"))
        self.assertTrue(hasattr(formulon.PivotCellKind, "COLUMN_SUBTOTAL"))
        self.assertFalse(hasattr(formulon.PivotCellKind, "COL_LABEL"))
        self.assertFalse(hasattr(formulon.PivotCellKind, "COL_SUBTOTAL"))

    def test_chart_bar_direction_spells_column_out(self) -> None:
        self.assertTrue(hasattr(formulon.ChartBarDirection, "COLUMN"))
        self.assertFalse(hasattr(formulon.ChartBarDirection, "COL"))

    def test_pivot_filter_kind_members_describe_their_operand(self) -> None:
        self.assertTrue(hasattr(formulon.PivotFilterKind, "VALUE_TOP_COUNT"))
        self.assertTrue(hasattr(formulon.PivotFilterKind, "DATE_BETWEEN"))
        self.assertFalse(hasattr(formulon.PivotFilterKind, "VALUE_TOP_10"))
        self.assertFalse(hasattr(formulon.PivotFilterKind, "LABEL_DATE"))
        self.assertEqual(formulon.PivotFilterKind.VALUE_TOP_COUNT, 0)

    def test_pivot_filter_operand_enum_is_not_public(self) -> None:
        self.assertFalse(hasattr(formulon, "PivotFilterOperandKind"))
        self.assertTrue(hasattr(formulon, "PivotFilterKind"))

    def test_extended_chart_plot_type_and_layout(self) -> None:
        self.assertEqual(formulon.ChartPlotType.EXTENDED, 16)
        self.assertFalse(hasattr(formulon.ChartPlotType, "CHARTEX"))
        layout = formulon.ChartExtendedLayout
        self.assertEqual(
            [member.name for member in layout],
            [
                "BOX_WHISKER",
                "CLUSTERED_COLUMN",
                "FUNNEL",
                "PARETO_LINE",
                "REGION_MAP",
                "SUNBURST",
                "TREEMAP",
                "WATERFALL",
            ],
        )
        self.assertEqual([member.value for member in layout], list(range(8)))

    def test_chart_extended_fields_are_spelled_extended(self) -> None:
        self.assertFieldsRenamed(formulon.ChartPlot, ("extended_layout",), ("chartex_layout",))
        self.assertFieldsRenamed(formulon.ChartProperties, ("is_extended",), ("is_chartex",))


if __name__ == "__main__":
    unittest.main()
