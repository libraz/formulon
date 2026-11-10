"""Patch semantics: atomic multi-key updates, empty patches that change nothing, and read-to-patch round trips."""

from __future__ import annotations

import unittest
from typing import Any

import formulon
from formulon import (
    UNCHANGED,
    CellAddress,
    ChartPlotType,
    FormulonError,
    PivotAxis,
    PivotInput,
    Workbook,
)


def _grid(ws: Any) -> None:
    for c, name in enumerate(("Name", "North", "South")):
        ws.set_text(0, c, name)
    for r in range(1, 4):
        ws.set_text(r, 0, f"row{r}")
        ws.set_number(r, 1, float(r))
        ws.set_number(r, 2, float(r * 2))


def _chart(wb: Workbook) -> Any:
    _grid(wb.sheet(0))
    return wb.sheet(0).add_chart(formulon.ChartInput(type=ChartPlotType.BAR, range="Sheet1!A1:C4"))


def _pivot(wb: Workbook) -> Any:
    """A pivot whose Region field carries the items East and West and whose Product field is a second field."""
    cache = wb.add_pivot_cache()
    for name in ("Region", "Product", "Amount"):
        cache.add_field(name)
    for item in ("East", "West"):
        cache.add_field_shared_item(0, item)
    for values in ((0, "Pen", 10.0), (1, "Ink", 30.0)):
        record = cache.add_record()
        for field_index, value in enumerate(values):
            cache.set_record_value(record, field_index, value)
    pivot = wb.sheet(0).pivot(
        wb.sheet(0).add_pivot(PivotInput(name="P1", cache_id=cache.cache_id, anchor=CellAddress(0, 6)))
    )
    pivot.set_field_axis("Region", PivotAxis.ROW)
    pivot.set_field_axis("Product", PivotAxis.ROW)
    pivot.set_field_axis("Amount", PivotAxis.VALUE)
    return pivot


class PivotFieldUpdateTests(unittest.TestCase):
    def test_update_field_changes_nothing_when_a_key_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb)
            before = pivot.list_fields()[0]
            self.assertFalse(before.insert_blank_row)
            with self.assertRaises(FormulonError):
                pivot.update_field(0, insert_blank_row=True, show_all=True, custom_name="Product")
            self.assertEqual(pivot.list_fields()[0], before)

    def test_update_field_changes_nothing_when_a_later_key_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb)
            before = pivot.list_fields()[0]
            with self.assertRaises(FormulonError):
                pivot.update_field(
                    0,
                    insert_blank_row=True,
                    sort_order=formulon.PivotSortOrder.MANUAL,
                    sort_by_field="Amount",
                )
            self.assertEqual(pivot.list_fields()[0], before)

    def test_update_field_applies_every_key_at_once(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb)
            pivot.update_field(0, insert_blank_row=True, custom_name="Area", show_all=True)
            read = pivot.list_fields()[0]
            self.assertTrue(read.insert_blank_row)
            self.assertTrue(read.show_all)
            self.assertEqual(read.custom_name, "Area")

    def test_update_field_with_no_key_changes_nothing(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb)
            before = pivot.list_fields()
            pivot.update_field(0)
            self.assertEqual(pivot.list_fields(), before)

    def test_update_field_takes_sort_and_show_detail(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb)
            self.assertTrue(pivot.list_fields()[0].show_detail)
            pivot.update_field(0, show_detail=False)
            self.assertFalse(pivot.list_fields()[0].show_detail)
            self.assertFalse(any(item.show_detail for item in pivot.list_field_items(0)))
            pivot.update_field(0, sort_order=formulon.PivotSortOrder.DESCENDING)
            self.assertEqual(pivot.list_fields()[0].sort_order, formulon.PivotSortOrder.DESCENDING)
            self.assertIsNone(pivot.list_fields()[0].sort_by_field)
            pivot.update_field(0, sort_by_field=None)
            self.assertIsNone(pivot.list_fields()[0].sort_by_field)

    def test_update_field_round_trips_the_read_record(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb)
            read = pivot.list_fields()[0]
            keys = (
                "custom_name",
                "number_format",
                "default_subtotal",
                "insert_blank_row",
                "fill_down_labels",
                "show_all",
                "subtotal_top",
                "sort_order",
                "sort_by_field",
                "show_detail",
            )
            pivot.update_field(0, **{key: getattr(read, key) for key in keys})
            self.assertEqual(pivot.list_fields()[0], read)

    def test_update_field_rejects_a_bad_value_type_before_any_change(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb)
            before = pivot.list_fields()
            with self.assertRaises(TypeError):
                pivot.update_field(0, insert_blank_row=True, custom_name=5)  # type: ignore[arg-type]
            self.assertEqual(pivot.list_fields(), before)

    def test_update_field_item_changes_nothing_when_a_key_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb)
            before = pivot.list_field_items(0)
            with self.assertRaises(FormulonError):
                pivot.update_field_item(0, 0, visible=False, caption="West")
            self.assertEqual(pivot.list_field_items(0), before)

    def test_update_field_item_applies_every_key_at_once(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb)
            pivot.update_field_item(0, 0, visible=False, caption="Eastern", show_detail=False)
            item = pivot.get_field_item(0, 0)
            self.assertFalse(item.visible)
            self.assertFalse(item.show_detail)
            self.assertEqual(item.caption, "Eastern")
            pivot.update_field_item(0, 0, caption=None)
            self.assertIsNone(pivot.get_field_item(0, 0).caption)

    def test_the_per_property_sort_and_detail_setters_are_gone(self) -> None:
        with Workbook.create_default() as wb:
            pivot = _pivot(wb)
            self.assertFalse(hasattr(pivot, "set_field_sort"))
            self.assertFalse(hasattr(pivot, "set_field_show_detail"))
        self.assertFalse(hasattr(formulon.Pivot, "set_field_sort"))
        self.assertFalse(hasattr(formulon.Pivot, "set_field_show_detail"))


class ChartPatchTests(unittest.TestCase):
    def test_update_point_without_a_keyword_keeps_the_override(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            fill = {"color": "#FF0000"}
            chart.update_point(0, 0, 1, fill=fill)
            before = chart.get_properties().plots[0].series[0].points
            self.assertEqual(len(before), 1)
            chart.update_point(0, 0, 1)
            self.assertEqual(chart.get_properties().plots[0].series[0].points, before)

    def test_set_point_none_removes_the_override(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            fill = {"color": "#FF0000"}
            chart.update_point(0, 0, 1, fill=fill)
            chart.set_point(0, 0, 1, None)
            self.assertEqual(chart.get_properties().plots[0].series[0].points, [])
            chart.set_point(0, 0, 1, None)

    def test_set_point_round_trips_a_read_point(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            fill = {"color": "#00FF00"}
            chart.update_point(0, 0, 1, fill=fill)
            point = chart.get_properties().plots[0].series[0].points[0]
            chart.set_point(0, 0, 2, point)
            self.assertEqual(len(chart.get_properties().plots[0].series[0].points), 2)

    def test_legend_hidden_entries_replace_hidden_series(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            chart.update_legend(hidden_entries=[1])
            legend = chart.get_properties().legend
            self.assertEqual(legend.hidden_entries, [1])
            self.assertFalse(hasattr(legend, "hidden_series"))
            with self.assertRaises(TypeError):
                chart.update_legend(hidden_series=[1])  # type: ignore[call-arg]

    def test_filter_is_a_record_independent_of_legend_entries(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            chart.update_legend(hidden_entries=[0])
            chart.set_filter(formulon.ChartFilter(hidden_series=[1], hidden_categories=[2]))
            props = chart.get_properties()
            self.assertEqual(props.filter.hidden_series, [1])
            self.assertEqual(props.filter.hidden_categories, [2])
            self.assertEqual(props.legend.hidden_entries, [0])
            self.assertFalse(hasattr(props, "hidden_categories"))
            chart.set_filter(props.filter)
            self.assertEqual(chart.get_properties().filter, props.filter)

    def test_unfiltered_chart_reads_an_empty_filter(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            self.assertEqual(chart.get_properties().filter, formulon.ChartFilter())
            self.assertEqual(formulon.ChartFilter().hidden_series, [])

    def test_set_filter_refuses_the_old_positional_form(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            with self.assertRaises(TypeError):
                chart.set_filter([1], [2])  # type: ignore[call-arg]

    def test_axis_patch_carries_every_readable_key(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            axis = next(a for a in chart.get_properties().axes if a.kind == formulon.ChartAxisKind.VALUE)
            chart.update_axis(
                axis.id,
                minor_unit=0.5,
                minor_gridlines=True,
                position=formulon.ChartAxisPosition.RIGHT,
                source_linked=False,
            )
            read = next(a for a in chart.get_properties().axes if a.id == axis.id)
            self.assertEqual(read.minor_unit, 0.5)
            self.assertTrue(read.minor_gridlines)
            self.assertEqual(read.position, formulon.ChartAxisPosition.RIGHT)
            self.assertFalse(read.source_linked)
            chart.update_axis(axis.id, minor_unit=None)
            self.assertIsNone(next(a for a in chart.get_properties().axes if a.id == axis.id).minor_unit)

    def test_axis_patch_rejects_clearing_a_flag(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            axis_id = chart.get_properties().axes[0].id
            with self.assertRaises(TypeError):
                chart.update_axis(axis_id, minor_gridlines=None)  # type: ignore[arg-type]

    def test_axis_update_with_the_read_values_changes_nothing(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            before = chart.get_properties().axes
            for axis in before:
                chart.update_axis(
                    axis.id,
                    minor_unit=axis.minor_unit,
                    minor_gridlines=axis.minor_gridlines,
                    position=axis.position,
                    source_linked=axis.source_linked,
                )
            self.assertEqual(chart.get_properties().axes, before)

    def test_display_style_sets_and_clears(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            chart.update_display(style=26)
            self.assertEqual(chart.get_properties().style, 26)
            chart.update_display(style=None)
            self.assertIsNone(chart.get_properties().style)
            for bad in (0, 49):
                with self.assertRaises(FormulonError):
                    chart.update_display(style=bad)
            self.assertIsNone(chart.get_properties().style)

    def test_series_literal_source_round_trips(self) -> None:
        with Workbook.create_default() as wb:
            chart = _chart(wb)
            series = chart.get_properties().plots[0].series[0]
            chart.update_series(0, series.series_id, values=series.values)
            self.assertEqual(chart.get_properties().plots[0].series[0].values, series.values)


class SheetViewPatchTests(unittest.TestCase):
    def test_top_left_is_a_cell_address(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_view(top_left="C5")
            view = ws.get_view()
            self.assertEqual(view.top_left, CellAddress(4, 2))
            self.assertFalse(hasattr(view, "top_left_row"))
            ws.update_view(freeze_rows=1, freeze_cols=1)
            ws.update_view(top_left=CellAddress(1, 1), pane_top_left="B2")
            view = ws.get_view()
            self.assertEqual(view.top_left, CellAddress(1, 1))
            self.assertEqual(view.pane_top_left, CellAddress(1, 1))
            ws.update_view(top_left=view.top_left, pane_top_left=view.pane_top_left)
            self.assertEqual(ws.get_view(), view)

    def test_old_row_column_keywords_are_refused(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.sheet(0).update_view(top_left_row=1)  # type: ignore[call-arg]
            with self.assertRaises(TypeError):
                wb.sheet(0).update_view(pane_top_left_col=1)  # type: ignore[call-arg]

    def test_unchanged_sentinel_still_keeps_the_view(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_view(top_left="B2")
            ws.update_view(top_left=UNCHANGED)
            self.assertEqual(ws.get_view().top_left, CellAddress(1, 1))


class DefinedNameScopeTests(unittest.TestCase):
    def test_defined_name_scope_round_trips(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Second")
            wb.set_defined_name(formulon.DefinedNameInput(name="X", formula="=1", scope=1))
            read = wb.get_defined_name("X", 1)
            self.assertEqual(read.scope, 1)
            self.assertFalse(hasattr(read, "local_sheet_id"))
            wb.set_defined_name(read)
            self.assertEqual(wb.get_defined_name("X", 1), read)
            wb.set_defined_name(formulon.DefinedNameInput(name="G", formula="=2"))
            self.assertIsNone(wb.get_defined_name("G").scope)
