"""Chart surface tests: insert, typed reads, resolved data, typed edits, drawing operations and error paths."""

from __future__ import annotations

import unittest
from dataclasses import replace

from formulon import (
    AnchorEditAs,
    AnchorKind,
    ChartAxisCrosses,
    ChartAxisKind,
    ChartBlanksAs,
    ChartGrouping,
    ChartInfo,
    ChartInsertKind,
    ChartLabelPosition,
    ChartLegendPosition,
    ChartLineDash,
    ChartMarkerSymbol,
    ChartPlotType,
    ChartPointState,
    ChartSeriesDirection,
    ChartTitleInfo,
    ChartTitleKind,
    DrawingObjectKind,
    FormulonError,
    Workbook,
)

SOURCE = "Sheet1!$A$1:$D$5"


def _grid(wb: Workbook) -> None:
    """A1 blank, B1:D1 names, A2:A5 text categories, B2:D5 numbers, F1:G5 scatter pairs."""
    for c, name in enumerate(("S1", "S2", "S3")):
        wb.set_text(0, 0, c + 1, name)
    for r, cat in enumerate("abcd"):
        wb.set_text(0, r + 1, 0, cat)
        for c in range(3):
            wb.set_number(0, r + 1, c + 1, float((r + 1) * (c + 2)))
        wb.set_number(0, r + 1, 5, float(r * 3))
        wb.set_number(0, r + 1, 6, float(r * r))
    wb.set_text(0, 0, 5, "X")
    wb.set_text(0, 0, 6, "Y")


def _chart(kind: int = ChartInsertKind.COLUMN, source: str = SOURCE, **options) -> "tuple[Workbook, int]":
    wb = Workbook.create_default()
    _grid(wb)
    return wb, wb.insert_chart(0, kind, source, row=6, **options)


def _axis_id(info: ChartInfo, kind: ChartAxisKind) -> int:
    return next(a.id for a in info.axes if a.kind == kind)


class ChartEnumTests(unittest.TestCase):
    def test_ordinals_match_the_c_header(self) -> None:
        self.assertEqual((ChartPlotType.BAR, ChartPlotType.OF_PIE, ChartPlotType.CHARTEX), (0, 9, 16))
        self.assertEqual((ChartInsertKind.COLUMN, ChartInsertKind.SCATTER, ChartInsertKind.DOUGHNUT), (0, 5, 6))
        self.assertEqual((ChartGrouping.CLUSTERED, ChartGrouping.STANDARD), (0, 3))
        self.assertEqual((ChartBlanksAs.ABSENT, ChartBlanksAs.GAP, ChartBlanksAs.SPAN), (0, 1, 3))
        self.assertEqual((ChartPointState.VALUE, ChartPointState.GAP, ChartPointState.CONNECT), (0, 1, 2))
        self.assertEqual((ChartAxisKind.CATEGORY, ChartAxisKind.SERIES), (0, 3))
        self.assertEqual((ChartAxisCrosses.ABSENT, ChartAxisCrosses.AT), (0, 4))
        self.assertEqual((ChartTitleKind.NONE, ChartTitleKind.AUTO), (0, 3))
        self.assertEqual((ChartSeriesDirection.AUTO, ChartSeriesDirection.ROWS), (0, 2))
        self.assertEqual((ChartLegendPosition.NONE, ChartLegendPosition.TOP_RIGHT), (0, 5))
        self.assertEqual((ChartLineDash.SOLID, ChartLineDash.OTHER), (0, 4))
        self.assertEqual((ChartMarkerSymbol.NONE, ChartMarkerSymbol.PICTURE), (0, 11))
        self.assertEqual((ChartLabelPosition.BEST_FIT, ChartLabelPosition.RIGHT), (0, 8))


class ChartReadTests(unittest.TestCase):
    def test_insert_and_typed_read(self) -> None:
        wb, oid = _chart(legend=ChartLegendPosition.RIGHT, name="Sales", description="by quarter")
        with wb:
            objs = wb.list_drawing_objects(0)
            self.assertEqual([(o.object_id, o.kind) for o in objs], [(oid, DrawingObjectKind.CHART)])
            self.assertEqual(objs[0].anchor_kind, AnchorKind.TWO_CELL)

            info = wb.get_chart(0, oid)
            self.assertEqual((info.object_id, info.name, info.description), (oid, "Sales", "by quarter"))
            self.assertFalse(info.is_chartex or info.is_pivot_chart or info.on_chartsheet)
            self.assertEqual(info.title.kind, ChartTitleKind.AUTO)
            self.assertEqual(info.legend.position, ChartLegendPosition.RIGHT)
            self.assertEqual((info.legend.overlay, info.legend.hidden_series), (False, []))
            self.assertEqual(info.hidden_categories, [])
            self.assertIsNotNone(info.source)
            self.assertEqual(
                (info.source.range, info.source.direction, info.source.header_row, info.source.header_column),
                (SOURCE, ChartSeriesDirection.COLUMNS, True, True),
            )

            self.assertEqual(len(info.plots), 1)
            plot = info.plots[0]
            self.assertEqual((plot.type, plot.bar_direction, plot.grouping), (ChartPlotType.BAR, "col", 0))
            self.assertIsNone(plot.chartex_layout)
            self.assertEqual(len(plot.series), 3)
            self.assertEqual(len(plot.axis_ids), 2)
            s = plot.series[1]
            self.assertEqual((s.idx, s.order, s.filtered), (1, 1, False))
            self.assertEqual((s.name.kind, s.name.formula), ("ref", "Sheet1!$C$1"))
            self.assertEqual(s.categories.formula, "Sheet1!$A$2:$A$5")
            self.assertEqual(s.values.formula, "Sheet1!$C$2:$C$5")
            self.assertEqual((s.x.kind, s.size.kind), ("none", "none"))
            self.assertEqual(s.points, [])

            cat = next(a for a in info.axes if a.kind == ChartAxisKind.CATEGORY)
            val = next(a for a in info.axes if a.kind == ChartAxisKind.VALUE)
            self.assertEqual((cat.cross_axis, val.cross_axis), (val.id, cat.id))
            self.assertEqual((cat.position, val.position), ("b", "l"))
            self.assertIsNone(val.min)
            self.assertIsNone(val.max)
            self.assertEqual(sorted(plot.axis_ids), sorted((cat.id, val.id)))

    def test_title_text_and_none_inserts(self) -> None:
        wb, oid = _chart(title="Quarterly")
        with wb:
            title = wb.get_chart(0, oid).title
            self.assertEqual((title.kind, title.text), (ChartTitleKind.TEXT, "Quarterly"))
        wb, oid = _chart(title="")
        with wb:
            self.assertEqual(wb.get_chart(0, oid).title.kind, ChartTitleKind.NONE)

    def test_resolved_data(self) -> None:
        wb, oid = _chart()
        with wb:
            data = wb.get_chart_data(0, oid)
            self.assertEqual((data.domain, data.point_count, len(data.series)), ("category", 4, 3))
            self.assertEqual(data.categories.labels, ["a", "b", "c", "d"])
            self.assertIsNone(data.categories.values)
            self.assertEqual(data.categories.levels, [])
            self.assertFalse(data.categories.generated)
            self.assertEqual(data.unresolved, [])
            s0 = data.series[0]
            self.assertEqual((s0.plot_index, s0.idx, s0.name, s0.filtered), (0, 0, "S1", False))
            self.assertEqual(s0.values, [2.0, 4.0, 6.0, 8.0])
            self.assertEqual(s0.states, [ChartPointState.VALUE] * 4)
            self.assertIsNone(s0.x)
            self.assertIsNone(s0.sizes)
            self.assertEqual(len(s0.ranges.name), 1)
            self.assertEqual(
                (s0.ranges.values[0].sheet, s0.ranges.values[0].top, s0.ranges.values[0].left),
                (0, 1, 1),
            )
            self.assertEqual((s0.ranges.values[0].bottom, s0.ranges.values[0].right), (4, 1))
            self.assertEqual(s0.ranges.x, [])

            # The handle is a snapshot of the cells at the call.
            wb.set_number(0, 2, 1, 40.0)
            self.assertEqual(s0.values[1], 4.0)
            self.assertEqual(wb.get_chart_data(0, oid).series[0].values[1], 40.0)

    def test_blank_cell_is_a_gap(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.set_text(0, 2, 1, "")
            wb.set_number(0, 3, 1, 1.0)
            states = wb.get_chart_data(0, oid).series[0].states
            self.assertEqual(len(states), 4)
            self.assertIn(states[1], (ChartPointState.GAP, ChartPointState.VALUE))

    def test_scatter_is_the_xy_domain(self) -> None:
        wb, oid = _chart(ChartInsertKind.SCATTER, "Sheet1!$F$1:$G$5")
        with wb:
            wb.set_range_xf_index(0, 1, 5, 4, 5, wb.add_cell_xf(replace(wb.get_cell_xf(0), num_fmt_id=2)))
            info = wb.get_chart(0, oid)
            self.assertEqual(info.plots[0].type, ChartPlotType.SCATTER)
            data = wb.get_chart_data(0, oid)
            self.assertEqual(data.domain, "xy")
            self.assertEqual(len(data.series), 1)
            self.assertEqual(data.series[0].x, [0.0, 3.0, 6.0, 9.0])
            self.assertEqual(data.series[0].values, [0.0, 1.0, 4.0, 9.0])
            self.assertEqual((data.series[0].x_format_code, data.series[0].format_code), ("0.00", "General"))

    def test_preview_applies_overrides_without_editing(self) -> None:
        wb, oid = _chart()
        with wb:
            self.assertEqual(len(wb.preview_chart_data(0, oid).series), 3)
            rows = wb.preview_chart_data(0, oid, {"direction": ChartSeriesDirection.ROWS})
            self.assertEqual(len(rows.series), 4)
            narrow = wb.preview_chart_data(0, oid, {"source": "Sheet1!$A$1:$B$5"})
            self.assertEqual(len(narrow.series), 1)
            self.assertEqual(narrow.series[0].name, "S1")
            headerless = wb.preview_chart_data(0, oid, {"source": "Sheet1!$B$2:$B$5", "header_row": False})
            self.assertEqual(headerless.series[0].values, [2.0, 4.0, 6.0, 8.0])
            self.assertEqual(len(wb.get_chart(0, oid).plots[0].series), 3)
            self.assertEqual(wb.get_chart(0, oid).source.range, SOURCE)

            with self.assertRaises(FormulonError):
                wb.preview_chart_data(0, oid, {"source": "not a range"})
            with self.assertRaises(FormulonError):
                wb.preview_chart_data(0, oid, {"direction": None})
            with self.assertRaises(TypeError):
                wb.preview_chart_data(0, oid, {"colour": 1})  # type: ignore[typeddict-unknown-key]

    def test_save_and_load_keep_the_chart(self) -> None:
        wb, oid = _chart(title="Kept")
        with wb:
            loaded = Workbook.load(wb.save())
        with loaded:
            objs = loaded.list_drawing_objects(0)
            self.assertEqual(len(objs), 1)
            self.assertEqual(objs[0].kind, DrawingObjectKind.CHART)
            info = loaded.get_chart(0, objs[0].object_id)
            self.assertEqual((info.title.kind, info.title.text), (ChartTitleKind.TEXT, "Kept"))
            self.assertEqual(len(info.plots[0].series), 3)


class ChartEditTests(unittest.TestCase):
    def test_title_legend_display(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.set_chart_title(0, oid, "Revenue")
            info = wb.get_chart(0, oid)
            self.assertEqual((info.title.kind, info.title.text), (ChartTitleKind.TEXT, "Revenue"))
            wb.set_chart_title(0, oid, None)
            self.assertEqual(wb.get_chart(0, oid).title.kind, ChartTitleKind.NONE)
            wb.set_chart_title(0, oid, ChartTitleInfo(ChartTitleKind.AUTO))
            self.assertEqual(wb.get_chart(0, oid).title.kind, ChartTitleKind.AUTO)
            with self.assertRaises(FormulonError):
                wb.set_chart_title(0, oid, "")

            wb.set_chart_legend(0, oid, ChartLegendPosition.BOTTOM, True, hidden_series=[1])
            legend = wb.get_chart(0, oid).legend
            self.assertEqual(
                (legend.position, legend.overlay, legend.hidden_series), (ChartLegendPosition.BOTTOM, True, [1])
            )
            wb.set_chart_legend(0, oid, ChartLegendPosition.TOP)
            legend = wb.get_chart(0, oid).legend
            self.assertEqual(
                (legend.position, legend.overlay, legend.hidden_series), (ChartLegendPosition.TOP, False, [1])
            )
            wb.set_chart_legend(0, oid, ChartLegendPosition.NONE, hidden_series=[])
            self.assertEqual(wb.get_chart(0, oid).legend.position, ChartLegendPosition.NONE)
            with self.assertRaises(ValueError):
                wb.set_chart_legend(0, oid, ChartLegendPosition.TOP, hidden_series=[-1])

            wb.set_chart_display(0, oid, ChartBlanksAs.SPAN, False)
            info = wb.get_chart(0, oid)
            self.assertEqual((info.blanks_as, info.plot_visible_only), (ChartBlanksAs.SPAN, False))
            wb.set_chart_display(0, oid, ChartBlanksAs.ABSENT, None)
            info = wb.get_chart(0, oid)
            self.assertEqual((info.blanks_as, info.plot_visible_only), (ChartBlanksAs.ABSENT, None))

    def test_axis_patch(self) -> None:
        wb, oid = _chart()
        with wb:
            val_id = _axis_id(wb.get_chart(0, oid), ChartAxisKind.VALUE)
            wb.set_chart_axis(
                0,
                oid,
                val_id,
                {
                    "min": 0.0,
                    "max": 50.0,
                    "major_unit": 10.0,
                    "title": "Units",
                    "number_format": "0.0",
                    "major_gridlines": False,
                },
            )
            axis = next(a for a in wb.get_chart(0, oid).axes if a.id == val_id)
            self.assertEqual((axis.min, axis.max, axis.major_unit), (0.0, 50.0, 10.0))
            self.assertEqual((axis.title.kind, axis.title.text), (ChartTitleKind.TEXT, "Units"))
            self.assertEqual((axis.number_format, axis.major_gridlines), ("0.0", False))

            wb.set_chart_axis(
                0,
                oid,
                val_id,
                {"min": None, "title": None, "reversed": True, "log_base": 10.0, "tick_label_position": "low"},
            )
            axis = next(a for a in wb.get_chart(0, oid).axes if a.id == val_id)
            self.assertIsNone(axis.min)
            self.assertEqual(axis.max, 50.0)
            self.assertEqual(axis.title.kind, ChartTitleKind.NONE)
            self.assertEqual((axis.reversed, axis.log_base, axis.tick_label_position), (True, 10.0, "low"))

            wb.set_chart_axis(0, oid, val_id, {"title": ChartTitleInfo(ChartTitleKind.AUTO, "", "")})
            axis = next(a for a in wb.get_chart(0, oid).axes if a.id == val_id)
            self.assertEqual(axis.title.kind, ChartTitleKind.AUTO)

            wb.set_chart_axis(0, oid, val_id, {"crosses": ChartAxisCrosses.AT, "crosses_at": 5.0})
            axis = next(a for a in wb.get_chart(0, oid).axes if a.id == val_id)
            self.assertEqual((axis.crosses, axis.crosses_at), (ChartAxisCrosses.AT, 5.0))

            wb.set_chart_axis(0, oid, val_id, {})
            with self.assertRaises(FormulonError):
                wb.set_chart_axis(0, oid, val_id, {"visible": None})
            with self.assertRaises(FormulonError):
                wb.set_chart_axis(0, oid, 987654, {"min": 1.0})

    def test_series_patch(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.set_chart_series(
                0,
                oid,
                0,
                1,
                {
                    "name": "Sheet1!$D$1",
                    "fill": {"color": 0xFFFF0000, "transparency": 25.0},
                    "line": {"color": 0xFF00FF00, "width_pt": 2.5, "dash": ChartLineDash.DASH},
                },
            )
            s = wb.get_chart(0, oid).plots[0].series[1]
            self.assertEqual(s.name.formula, "Sheet1!$D$1")
            self.assertEqual((s.fill.color, s.fill.transparency), (0xFFFF0000, 25.0))
            self.assertEqual((s.line.color, s.line.width_pt, s.line.dash), (0xFF00FF00, 2.5, ChartLineDash.DASH))
            self.assertEqual(wb.get_chart_data(0, oid).series[1].name, "S3")

            wb.set_chart_series(0, oid, 0, 1, {"fill": {"transparency": None}, "name": ""})
            s = wb.get_chart(0, oid).plots[0].series[1]
            self.assertEqual(s.name.kind, "none")
            self.assertIsNone(s.fill.transparency)
            self.assertEqual(s.fill.color, 0xFFFF0000)

            wb.set_chart_series(
                0,
                oid,
                0,
                0,
                {"labels": {"show_value": True, "separator": "; ", "position": ChartLabelPosition.OUTSIDE_END}},
            )
            labels = wb.get_chart(0, oid).plots[0].series[0].labels
            self.assertIsNotNone(labels)
            self.assertEqual((labels.show_value, labels.show_category, labels.separator), (True, False, "; "))
            self.assertEqual(labels.position, ChartLabelPosition.OUTSIDE_END)
            wb.set_chart_series(0, oid, 0, 0, {"labels": None})
            self.assertIsNone(wb.get_chart(0, oid).plots[0].series[0].labels)

            wb.set_chart_series(0, oid, 0, 2, {"values": "Sheet1!$B$2:$B$5"})
            self.assertEqual(wb.get_chart_data(0, oid).series[2].values, [2.0, 4.0, 6.0, 8.0])

            with self.assertRaises(FormulonError):
                wb.set_chart_series(0, oid, 0, 2, {"values": ""})
            with self.assertRaises(FormulonError):
                wb.set_chart_series(0, oid, 0, 2, {"values": "not a range"})
            with self.assertRaises(FormulonError):
                wb.set_chart_series(0, oid, 0, 77, {"name": "Sheet1!$B$1"})
            with self.assertRaises(TypeError):
                wb.set_chart_series(0, oid, 0, 0, {"fil": {}})  # type: ignore[typeddict-unknown-key]
            with self.assertRaises(TypeError):
                wb.set_chart_series(0, oid, 0, 0, {"fill": {"colour": 1}})  # type: ignore[typeddict-unknown-key]

    def test_line_series_marker_and_smooth(self) -> None:
        wb, oid = _chart(ChartInsertKind.LINE)
        with wb:
            wb.set_chart_series(
                0,
                oid,
                0,
                0,
                {
                    "marker": {
                        "symbol": ChartMarkerSymbol.DIAMOND,
                        "size": 9.0,
                        "fill": {"color": 0xFF112233},
                        "line": {"none": True},
                    },
                    "smooth": True,
                },
            )
            s = wb.get_chart(0, oid).plots[0].series[0]
            self.assertIsNotNone(s.marker)
            self.assertEqual((s.marker.symbol, s.marker.size), (ChartMarkerSymbol.DIAMOND, 9))
            self.assertEqual(s.marker.fill.color, 0xFF112233)
            self.assertTrue(s.marker.line.none)
            self.assertIs(s.smooth, True)
            wb.set_chart_series(0, oid, 0, 0, {"marker": None})
            self.assertIsNone(wb.get_chart(0, oid).plots[0].series[0].marker)
            with self.assertRaises(FormulonError):
                wb.set_chart_series(0, oid, 0, 0, {"smooth": None})

    def test_point_patch_and_removal(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.set_chart_point(0, oid, 0, 0, 2, {"fill": {"color": 0xFF0000FF}})
            points = wb.get_chart(0, oid).plots[0].series[0].points
            self.assertEqual([(p.idx, p.fill.color) for p in points], [(2, 0xFF0000FF)])
            self.assertIsNone(points[0].marker)
            wb.set_chart_point(0, oid, 0, 0, 2, None)
            self.assertEqual(wb.get_chart(0, oid).plots[0].series[0].points, [])
            with self.assertRaises(TypeError):
                wb.set_chart_point(0, oid, 0, 0, 2, {"colour": 1})  # type: ignore[typeddict-unknown-key]

    def test_plot_patch(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.set_chart_plot(0, oid, 0, {"gap_width": 80, "overlap": -20})
            plot = wb.get_chart(0, oid).plots[0]
            self.assertEqual((plot.gap_width, plot.overlap), (80, -20))
            with self.assertRaises(FormulonError):
                wb.set_chart_plot(0, oid, 0, {"hole_size": 50})
            with self.assertRaises(FormulonError):
                wb.set_chart_plot(0, oid, 5, {"gap_width": 80})
        wb, oid = _chart(ChartInsertKind.DOUGHNUT, "Sheet1!$A$1:$B$5")
        with wb:
            wb.set_chart_plot(0, oid, 0, {"hole_size": 60, "first_slice_angle": 90, "vary_colors": True})
            plot = wb.get_chart(0, oid).plots[0]
            self.assertEqual((plot.type, plot.hole_size, plot.first_slice_angle), (ChartPlotType.DOUGHNUT, 60, 90))
            self.assertTrue(plot.vary_colors)

    def test_filter(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.set_chart_filter(0, oid, [1], [2])
            info = wb.get_chart(0, oid)
            self.assertEqual(info.hidden_categories, [2])
            self.assertEqual([s.idx for s in info.plots[0].series if s.filtered], [1])
            data = wb.get_chart_data(0, oid)
            self.assertEqual([s.idx for s in data.series if s.filtered], [1])
            wb.set_chart_filter(0, oid, [], [])
            info = wb.get_chart(0, oid)
            self.assertEqual(info.hidden_categories, [])
            self.assertFalse(any(s.filtered for s in info.plots[0].series))

    def test_add_and_remove_series(self) -> None:
        wb, oid = _chart(source="Sheet1!$A$1:$B$5")
        with wb:
            self.assertEqual(len(wb.get_chart(0, oid).plots[0].series), 1)
            idx = wb.add_chart_series(
                0, oid, 0, {"name": "Sheet1!$C$1", "values": "Sheet1!$C$2:$C$5", "categories": "Sheet1!$A$2:$A$5"}
            )
            series = wb.get_chart(0, oid).plots[0].series
            self.assertEqual([s.idx for s in series], [0, idx])
            self.assertEqual(wb.get_chart_data(0, oid).series[1].values, [3.0, 6.0, 9.0, 12.0])
            wb.remove_chart_series(0, oid, 0, 0)
            self.assertEqual([s.idx for s in wb.get_chart(0, oid).plots[0].series], [idx])
            with self.assertRaises(FormulonError):
                wb.add_chart_series(0, oid, 0, {"name": "Sheet1!$C$1"})
            with self.assertRaises(FormulonError):
                wb.remove_chart_series(0, oid, 0, 99)

    def test_set_source(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.set_chart_source(0, oid, "Sheet1!$A$1:$C$5")
            info = wb.get_chart(0, oid)
            self.assertEqual(len(info.plots[0].series), 2)
            self.assertEqual(info.source.range, "Sheet1!$A$1:$C$5")
            wb.set_chart_source(0, oid, "Sheet1!$A$1:$D$5", ChartSeriesDirection.ROWS, True, True)
            info = wb.get_chart(0, oid)
            self.assertEqual(info.source.direction, ChartSeriesDirection.ROWS)
            self.assertEqual(len(info.plots[0].series), 4)
            with self.assertRaises(FormulonError):
                wb.set_chart_source(0, oid, "not a range")

    def test_set_type(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.set_chart_type(0, oid, 0, ChartInsertKind.LINE)
            plot = wb.get_chart(0, oid).plots[0]
            self.assertEqual((plot.type, len(plot.series)), (ChartPlotType.LINE, 3))
            wb.set_chart_type(0, oid, 0, ChartInsertKind.COLUMN, ChartGrouping.STACKED)
            plot = wb.get_chart(0, oid).plots[0]
            self.assertEqual(
                (plot.type, plot.bar_direction, plot.grouping), (ChartPlotType.BAR, "col", ChartGrouping.STACKED)
            )
            wb.set_chart_type(0, oid, 0, ChartInsertKind.PIE)
            self.assertEqual(wb.get_chart(0, oid).plots[0].type, ChartPlotType.PIE)
            with self.assertRaises(FormulonError):
                wb.set_chart_type(0, oid, 3, ChartInsertKind.LINE)
            with self.assertRaises(FormulonError):
                wb.set_chart_type(0, oid, 0, 99)


class ChartDrawingTests(unittest.TestCase):
    def test_anchor_and_z_order(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.set_drawing_anchor(
                0,
                oid,
                anchor_kind=AnchorKind.TWO_CELL,
                edit_as=AnchorEditAs.ONE_CELL,
                row=3,
                col=2,
                width_emu=2_000_000,
                height_emu=1_000_000,
            )
            obj = wb.list_drawing_objects(0)[0]
            self.assertEqual((obj.object_id, obj.kind), (oid, DrawingObjectKind.CHART))
            self.assertEqual((obj.from_row, obj.from_col, obj.cx, obj.cy), (3, 2, 2_000_000, 1_000_000))
            self.assertEqual(obj.edit_as, AnchorEditAs.ONE_CELL)

            other = wb.insert_chart(0, ChartInsertKind.PIE, "Sheet1!$A$1:$B$5", row=20)
            self.assertEqual([o.object_id for o in wb.list_drawing_objects(0)], [oid, other])
            wb.set_drawing_z_order(0, other, 0)
            self.assertEqual([o.object_id for o in wb.list_drawing_objects(0)], [other, oid])
            with self.assertRaises(FormulonError):
                wb.set_drawing_z_order(0, oid, 2)

    def test_default_frame_and_explicit_size(self) -> None:
        wb, oid = _chart()
        with wb:
            obj = wb.list_drawing_objects(0)[0]
            self.assertEqual((obj.cx, obj.cy), (360 * 12700, 216 * 12700))
            sized = wb.insert_chart(
                0, ChartInsertKind.AREA, SOURCE, anchor_kind=AnchorKind.ONE_CELL, width_emu=500_000, height_emu=400_000
            )
            by_id = {o.object_id: o for o in wb.list_drawing_objects(0)}
            self.assertEqual((by_id[sized].cx, by_id[sized].cy), (500_000, 400_000))
            self.assertEqual(by_id[sized].anchor_kind, AnchorKind.ONE_CELL)

    def test_snapshot_remove_restore(self) -> None:
        wb, a = _chart(title="A")
        with wb:
            b = wb.insert_chart(0, ChartInsertKind.PIE, "Sheet1!$A$1:$B$5", row=20, title="B")
            snap = wb.snapshot_drawing(0, a)
            self.assertIsInstance(snap, bytes)
            self.assertGreater(len(snap), 0)

            wb.remove_drawing(0, a)
            self.assertEqual([o.object_id for o in wb.list_drawing_objects(0)], [b])
            with self.assertRaises(FormulonError):
                wb.get_chart(0, a)
            with self.assertRaises(FormulonError):
                wb.remove_drawing(0, a)

            self.assertEqual(wb.restore_drawing(0, snap), a)
            self.assertEqual([o.object_id for o in wb.list_drawing_objects(0)], [a, b])
            restored = wb.get_chart(0, a)
            self.assertEqual((restored.title.text, len(restored.plots[0].series)), ("A", 3))

            wb.set_chart_title(0, a, "Changed")
            self.assertEqual(wb.restore_drawing(0, snap), a)
            self.assertEqual(wb.get_chart(0, a).title.text, "A")

            copy = wb.restore_drawing(0, snap, new_id=True)
            self.assertNotIn(copy, (a, b))
            self.assertEqual([o.object_id for o in wb.list_drawing_objects(0)], [a, b, copy])
            self.assertEqual(wb.get_chart(0, copy).title.text, "A")

            with self.assertRaises(FormulonError):
                wb.restore_drawing(0, b"junk")
            with self.assertRaises(FormulonError):
                wb.snapshot_drawing(0, 9999)


class ChartErrorTests(unittest.TestCase):
    def test_unknown_chart_and_sheet(self) -> None:
        wb, oid = _chart()
        with wb:
            for call in (
                lambda: wb.get_chart(0, 9999),
                lambda: wb.get_chart_data(0, 9999),
                lambda: wb.get_chart(9, oid),
                lambda: wb.set_chart_title(0, 9999, "x"),
                lambda: wb.set_chart_filter(0, 9999, [], []),
            ):
                with self.assertRaises(FormulonError):
                    call()

    def test_insert_rejects_bad_input(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb)
            with self.assertRaises(FormulonError):
                wb.insert_chart(0, ChartInsertKind.COLUMN, "not a range")
            with self.assertRaises(FormulonError):
                wb.insert_chart(0, 99, SOURCE)
            with self.assertRaises(FormulonError):
                wb.insert_chart(0, ChartInsertKind.COLUMN, SOURCE, anchor_kind=AnchorKind.ABSOLUTE)
            with self.assertRaises(ValueError):
                wb.insert_chart(0, ChartInsertKind.COLUMN, SOURCE, row=-1)
            self.assertEqual(wb.list_drawing_objects(0), [])

    def test_failed_edit_leaves_the_chart_unchanged(self) -> None:
        wb, oid = _chart()
        with wb:
            before = wb.get_chart(0, oid)
            with self.assertRaises(FormulonError):
                wb.set_chart_series(0, oid, 0, 0, {"fill": {"color": 0xFF000000}, "values": "not a range"})
            self.assertEqual(wb.get_chart(0, oid), before)


if __name__ == "__main__":
    unittest.main()
