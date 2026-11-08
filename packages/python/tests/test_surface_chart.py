"""Chart surface tests: insert, typed reads, resolved data, typed edits, drawing operations and error paths."""

from __future__ import annotations

import unittest
from dataclasses import replace

from _xf_input import input_of

from formulon import (
    UNCHANGED,
    AnchorEditAs,
    AnchorKind,
    CellRect,
    Chart,
    ChartAxisCrosses,
    ChartAxisKind,
    ChartAxisPosition,
    ChartBarDirection,
    ChartBlanksAs,
    ChartGrouping,
    ChartInsertKind,
    ChartLabelPosition,
    ChartLegendPosition,
    ChartLineDash,
    ChartMarkerSymbol,
    ChartPlotType,
    ChartPointState,
    ChartSeriesDirection,
    ChartSourceKind,
    ChartTitle,
    ChartTitleKind,
    Color,
    ColorKind,
    DrawingAnchor,
    DrawingObjectKind,
    FormulonError,
    InvalidArgumentError,
    NotFoundError,
    ShapePaint,
    SheetRect,
    Workbook,
)

SOURCE = "Sheet1!$A$1:$D$5"


def _grid(wb: Workbook) -> None:
    """A1 blank, B1:D1 names, A2:A5 text categories, B2:D5 numbers, F1:G5 scatter pairs."""
    for c, name in enumerate(("S1", "S2", "S3")):
        wb.sheet(0).set_text(0, c + 1, name)
    for r, cat in enumerate("abcd"):
        wb.sheet(0).set_text(r + 1, 0, cat)
        for c in range(3):
            wb.sheet(0).set_number(r + 1, c + 1, float((r + 1) * (c + 2)))
        wb.sheet(0).set_number(r + 1, 5, float(r * 3))
        wb.sheet(0).set_number(r + 1, 6, float(r * r))
    wb.sheet(0).set_text(0, 5, "X")
    wb.sheet(0).set_text(0, 6, "Y")


def _chart(kind: int = ChartInsertKind.COLUMN, source: str = SOURCE, **options) -> "tuple[Workbook, int]":
    wb = Workbook.create_default()
    _grid(wb)
    return wb, wb.sheet(0).add_chart(
        kind, source, anchor=DrawingAnchor(anchor_kind=AnchorKind.TWO_CELL, row=6), **options
    ).object_id


def _axis_id(info: Chart, kind: ChartAxisKind) -> int:
    return next(a.id for a in info.axes if a.kind == kind)


def _rgb(value: int) -> Color:
    return Color(ColorKind.RGB, rgb=value)


class ChartEnumTests(unittest.TestCase):
    def test_ordinals_match_the_c_header(self) -> None:
        self.assertEqual((ChartPlotType.BAR, ChartPlotType.OF_PIE, ChartPlotType.CHARTEX), (0, 9, 16))
        self.assertEqual((ChartInsertKind.COLUMN, ChartInsertKind.SCATTER, ChartInsertKind.DOUGHNUT), (0, 5, 6))
        self.assertEqual((ChartGrouping.CLUSTERED, ChartGrouping.STANDARD), (0, 3))
        self.assertEqual((ChartBlanksAs.GAP, ChartBlanksAs.SPAN), (0, 2))
        self.assertEqual((ChartPointState.VALUE, ChartPointState.GAP, ChartPointState.CONNECT), (0, 1, 2))
        self.assertEqual((ChartAxisKind.CATEGORY, ChartAxisKind.SERIES), (0, 3))
        self.assertEqual((ChartAxisCrosses.AUTO_ZERO, ChartAxisCrosses.AT), (0, 3))
        self.assertEqual((ChartTitleKind.TEXT, ChartTitleKind.AUTO), (1, 3))
        self.assertEqual((ChartSeriesDirection.AUTO, ChartSeriesDirection.ROWS), (0, 2))
        self.assertEqual((ChartLegendPosition.BOTTOM, ChartLegendPosition.TOP_RIGHT), (0, 4))
        self.assertEqual((ChartAxisPosition.BOTTOM, ChartAxisPosition.RIGHT), (0, 3))
        self.assertEqual((ChartBarDirection.COL, ChartBarDirection.BAR), (0, 1))
        for absent in ("ABSENT", "NONE"):
            self.assertFalse(hasattr(ChartBlanksAs, absent))
            self.assertFalse(hasattr(ChartLegendPosition, absent))
            self.assertFalse(hasattr(ChartTitleKind, absent))
        self.assertEqual((ChartLineDash.SOLID, ChartLineDash.OTHER), (0, 4))
        self.assertEqual((ChartMarkerSymbol.NONE, ChartMarkerSymbol.PICTURE), (0, 11))
        self.assertEqual((ChartLabelPosition.BEST_FIT, ChartLabelPosition.RIGHT), (0, 8))


class ChartReadTests(unittest.TestCase):
    def test_insert_and_typed_read(self) -> None:
        wb, oid = _chart(legend=ChartLegendPosition.RIGHT, name="Sales", description="by quarter")
        with wb:
            objs = wb.sheet(0).list_drawings()
            self.assertEqual([(o.object_id, o.kind) for o in objs], [(oid, DrawingObjectKind.CHART)])
            self.assertEqual(objs[0].get_drawing_properties().anchor_kind, AnchorKind.TWO_CELL)

            info = wb.sheet(0).chart(oid).get_properties()
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
            self.assertEqual(
                (plot.type, plot.bar_direction, plot.grouping), (ChartPlotType.BAR, ChartBarDirection.COL, 0)
            )
            self.assertIsNone(plot.chartex_layout)
            self.assertEqual(len(plot.series), 3)
            self.assertEqual(len(plot.axis_ids), 2)
            s = plot.series[1]
            self.assertEqual((s.index, s.order, s.filtered), (1, 1, False))
            self.assertEqual((s.name.kind, s.name.formula), (ChartSourceKind.REF, "Sheet1!$C$1"))
            self.assertEqual(s.categories.formula, "Sheet1!$A$2:$A$5")
            self.assertEqual(s.values.formula, "Sheet1!$C$2:$C$5")
            self.assertEqual((s.x.kind, s.size.kind), (ChartSourceKind.NONE, ChartSourceKind.NONE))
            self.assertEqual(s.points, [])

            cat = next(a for a in info.axes if a.kind == ChartAxisKind.CATEGORY)
            val = next(a for a in info.axes if a.kind == ChartAxisKind.VALUE)
            self.assertEqual((cat.cross_axis, val.cross_axis), (val.id, cat.id))
            self.assertEqual((cat.position, val.position), (ChartAxisPosition.BOTTOM, ChartAxisPosition.LEFT))
            self.assertIsNone(val.min)
            self.assertIsNone(val.max)
            self.assertEqual(sorted(plot.axis_ids), sorted((cat.id, val.id)))

    def test_title_text_and_none_inserts(self) -> None:
        wb, oid = _chart(title="Quarterly")
        with wb:
            title = wb.sheet(0).chart(oid).get_properties().title
            self.assertEqual((title.kind, title.text), (ChartTitleKind.TEXT, "Quarterly"))
        wb, oid = _chart(title="")
        with wb:
            self.assertIsNone(wb.sheet(0).chart(oid).get_properties().title)

    def test_resolved_data(self) -> None:
        wb, oid = _chart()
        with wb:
            data = wb.sheet(0).chart(oid).get_data()
            self.assertEqual((data.domain, data.point_count, len(data.series)), ("category", 4, 3))
            self.assertEqual(data.categories.labels, ["a", "b", "c", "d"])
            self.assertIsNone(data.categories.values)
            self.assertEqual(data.categories.levels, [])
            self.assertFalse(data.categories.generated)
            self.assertEqual(data.unresolved, [])
            s0 = data.series[0]
            self.assertEqual((s0.plot_index, s0.index, s0.name, s0.filtered), (0, 0, "S1", False))
            self.assertEqual(s0.values, [2.0, 4.0, 6.0, 8.0])
            self.assertEqual(s0.states, [ChartPointState.VALUE] * 4)
            self.assertIsNone(s0.x)
            self.assertIsNone(s0.sizes)
            self.assertEqual(len(s0.ranges.name), 1)
            self.assertEqual(s0.ranges.values[0], SheetRect(sheet=0, rect=CellRect(1, 1, 4, 1)))
            self.assertEqual(s0.ranges.x, [])

            # The handle is a snapshot of the cells at the call.
            wb.sheet(0).set_number(2, 1, 40.0)
            self.assertEqual(s0.values[1], 4.0)
            self.assertEqual(wb.sheet(0).chart(oid).get_data().series[0].values[1], 40.0)

    def test_blank_cell_is_a_gap(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.sheet(0).set_text(2, 1, "")
            wb.sheet(0).set_number(3, 1, 1.0)
            states = wb.sheet(0).chart(oid).get_data().series[0].states
            self.assertEqual(len(states), 4)
            self.assertIn(states[1], (ChartPointState.GAP, ChartPointState.VALUE))

    def test_scatter_is_the_xy_domain(self) -> None:
        wb, oid = _chart(ChartInsertKind.SCATTER, "Sheet1!$F$1:$G$5")
        with wb:
            wb.sheet(0).ooxml.set_range_xf_index(
                CellRect(1, 5, 4, 5), wb.ooxml.add_cell_xf(replace(input_of(wb.ooxml.get_cell_xf(0)), num_fmt_id=2))
            )
            info = wb.sheet(0).chart(oid).get_properties()
            self.assertEqual(info.plots[0].type, ChartPlotType.SCATTER)
            data = wb.sheet(0).chart(oid).get_data()
            self.assertEqual(data.domain, "xy")
            self.assertEqual(len(data.series), 1)
            self.assertEqual(data.series[0].x, [0.0, 3.0, 6.0, 9.0])
            self.assertEqual(data.series[0].values, [0.0, 1.0, 4.0, 9.0])
            self.assertEqual((data.series[0].x_format_code, data.series[0].format_code), ("0.00", "General"))

    def test_preview_applies_overrides_without_editing(self) -> None:
        wb, oid = _chart()
        with wb:
            self.assertEqual(len(wb.sheet(0).chart(oid).get_preview().series), 3)
            rows = wb.sheet(0).chart(oid).get_preview(direction=ChartSeriesDirection.ROWS)
            self.assertEqual(len(rows.series), 4)
            narrow = wb.sheet(0).chart(oid).get_preview(source="Sheet1!$A$1:$B$5")
            self.assertEqual(len(narrow.series), 1)
            self.assertEqual(narrow.series[0].name, "S1")
            headerless = wb.sheet(0).chart(oid).get_preview(source="Sheet1!$B$2:$B$5", header_row=False)
            self.assertEqual(headerless.series[0].values, [2.0, 4.0, 6.0, 8.0])
            self.assertEqual(len(wb.sheet(0).chart(oid).get_properties().plots[0].series), 3)
            self.assertEqual(wb.sheet(0).chart(oid).get_properties().source.range, SOURCE)

            plain = wb.sheet(0).chart(oid).get_preview()
            self.assertEqual(wb.sheet(0).chart(oid).get_preview(direction=UNCHANGED), plain)

            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).get_preview(source="not a range")
            with self.assertRaises(TypeError):
                wb.sheet(0).chart(oid).get_preview(colour=1)  # type: ignore[call-arg]

    def test_save_and_load_keep_the_chart(self) -> None:
        wb, oid = _chart(title="Kept")
        with wb:
            loaded = Workbook.load(wb.save())
        with loaded:
            objs = loaded.sheet(0).list_drawings()
            self.assertEqual(len(objs), 1)
            self.assertEqual(objs[0].kind, DrawingObjectKind.CHART)
            info = loaded.sheet(0).chart(objs[0].object_id).get_properties()
            self.assertEqual((info.title.kind, info.title.text), (ChartTitleKind.TEXT, "Kept"))
            self.assertEqual(len(info.plots[0].series), 3)


class ChartEditTests(unittest.TestCase):
    def test_title_legend_display(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.sheet(0).chart(oid).set_title("Revenue")
            info = wb.sheet(0).chart(oid).get_properties()
            self.assertEqual((info.title.kind, info.title.text), (ChartTitleKind.TEXT, "Revenue"))
            wb.sheet(0).chart(oid).set_title(None)
            self.assertIsNone(wb.sheet(0).chart(oid).get_properties().title)
            wb.sheet(0).chart(oid).set_title(ChartTitle(ChartTitleKind.AUTO))
            self.assertEqual(wb.sheet(0).chart(oid).get_properties().title.kind, ChartTitleKind.AUTO)
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_title("")

            wb.sheet(0).chart(oid).set_legend(position=ChartLegendPosition.BOTTOM, overlay=True, hidden_series=[1])
            legend = wb.sheet(0).chart(oid).get_properties().legend
            self.assertEqual(
                (legend.position, legend.overlay, legend.hidden_series), (ChartLegendPosition.BOTTOM, True, [1])
            )
            wb.sheet(0).chart(oid).set_legend(position=ChartLegendPosition.TOP)
            legend = wb.sheet(0).chart(oid).get_properties().legend
            self.assertEqual(
                (legend.position, legend.overlay, legend.hidden_series), (ChartLegendPosition.TOP, True, [1])
            )
            wb.sheet(0).chart(oid).set_legend(None)
            self.assertIsNone(wb.sheet(0).chart(oid).get_properties().legend)
            with self.assertRaises(ValueError):
                wb.sheet(0).chart(oid).set_legend(position=ChartLegendPosition.TOP, hidden_series=[-1])

            wb.sheet(0).chart(oid).set_display(blanks_as=ChartBlanksAs.SPAN, plot_visible_only=False)
            info = wb.sheet(0).chart(oid).get_properties()
            self.assertEqual((info.blanks_as, info.plot_visible_only), (ChartBlanksAs.SPAN, False))
            wb.sheet(0).chart(oid).set_display(blanks_as=None, plot_visible_only=None)
            info = wb.sheet(0).chart(oid).get_properties()
            self.assertEqual((info.blanks_as, info.plot_visible_only), (None, None))

    def test_display_patch_keeps_omitted_keys(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.sheet(0).chart(oid).set_display(blanks_as=ChartBlanksAs.ZERO, plot_visible_only=True)
            wb.sheet(0).chart(oid).set_display(plot_visible_only=False)
            info = wb.sheet(0).chart(oid).get_properties()
            self.assertEqual((info.blanks_as, info.plot_visible_only), (ChartBlanksAs.ZERO, False))
            wb.sheet(0).chart(oid).set_display()
            info = wb.sheet(0).chart(oid).get_properties()
            self.assertEqual((info.blanks_as, info.plot_visible_only), (ChartBlanksAs.ZERO, False))
            with self.assertRaises(TypeError):
                wb.sheet(0).chart(oid).set_display(nope=1)  # type: ignore[call-arg]
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_display(blanks_as=-1)

    def test_na_as_blank_patch(self) -> None:
        wb, oid = _chart()
        with wb:
            self.assertIs(wb.sheet(0).chart(oid).get_properties().na_as_blank, True)
            wb.sheet(0).chart(oid).set_display(na_as_blank=False)
            self.assertIsNone(wb.sheet(0).chart(oid).get_properties().na_as_blank)
            wb.sheet(0).chart(oid).set_display(na_as_blank=True)
            self.assertIs(wb.sheet(0).chart(oid).get_properties().na_as_blank, True)
            wb.sheet(0).chart(oid).set_display(blanks_as=ChartBlanksAs.GAP)
            self.assertIs(wb.sheet(0).chart(oid).get_properties().na_as_blank, True)
            wb.sheet(0).chart(oid).set_display(na_as_blank=None)
            self.assertIsNone(wb.sheet(0).chart(oid).get_properties().na_as_blank)

    def test_preview_na_as_blank_override(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.sheet(0).set_formula(1, 1, "=NA()")
            wb.sheet(0).set_blank(2, 1)
            wb.recalc()
            # Blanks plot as zero, so a #N/A shown as blank differs from one left as a gap.
            wb.sheet(0).chart(oid).set_display(blanks_as=ChartBlanksAs.ZERO, na_as_blank=False)
            today = wb.sheet(0).chart(oid).get_data().series[0].states
            on = wb.sheet(0).chart(oid).get_preview(na_as_blank=True).series[0].states
            off = wb.sheet(0).chart(oid).get_preview(na_as_blank=False).series[0].states
            self.assertEqual(on[0], on[1])
            self.assertNotEqual(off[0], off[1])
            self.assertEqual(off, today)
            self.assertIsNone(wb.sheet(0).chart(oid).get_properties().na_as_blank)

    def test_axis_patch(self) -> None:
        wb, oid = _chart()
        with wb:
            val_id = _axis_id(wb.sheet(0).chart(oid).get_properties(), ChartAxisKind.VALUE)
            wb.sheet(0).chart(oid).set_axis(
                val_id, min=0.0, max=50.0, major_unit=10.0, title="Units", number_format="0.0", major_gridlines=False
            )
            axis = next(a for a in wb.sheet(0).chart(oid).get_properties().axes if a.id == val_id)
            self.assertEqual((axis.min, axis.max, axis.major_unit), (0.0, 50.0, 10.0))
            self.assertEqual((axis.title.kind, axis.title.text), (ChartTitleKind.TEXT, "Units"))
            self.assertEqual((axis.number_format, axis.major_gridlines), ("0.0", False))

            wb.sheet(0).chart(oid).set_axis(
                val_id, min=None, title=None, reversed=True, log_base=10.0, tick_label_position="low"
            )
            axis = next(a for a in wb.sheet(0).chart(oid).get_properties().axes if a.id == val_id)
            self.assertIsNone(axis.min)
            self.assertEqual(axis.max, 50.0)
            self.assertIsNone(axis.title)
            self.assertEqual((axis.reversed, axis.log_base, axis.tick_label_position), (True, 10.0, "low"))

            wb.sheet(0).chart(oid).set_axis(val_id, title=ChartTitle(ChartTitleKind.AUTO, "", ""))
            axis = next(a for a in wb.sheet(0).chart(oid).get_properties().axes if a.id == val_id)
            self.assertEqual(axis.title.kind, ChartTitleKind.AUTO)

            wb.sheet(0).chart(oid).set_axis(val_id, crosses=ChartAxisCrosses.AT, crosses_at=5.0)
            axis = next(a for a in wb.sheet(0).chart(oid).get_properties().axes if a.id == val_id)
            self.assertEqual((axis.crosses, axis.crosses_at), (ChartAxisCrosses.AT, 5.0))

            wb.sheet(0).chart(oid).set_axis(val_id)
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_axis(val_id, visible=None)
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_axis(987654, min=1.0)

    def test_theme_colour_round_trip_and_indexed_refused(self) -> None:
        wb, oid = _chart()
        with wb:
            theme = Color(ColorKind.THEME, theme=4, tint=0.4)
            wb.sheet(0).chart(oid).set_series(0, 0, fill={"color": theme})
            got = wb.sheet(0).chart(oid).get_properties().plots[0].series[0].fill.color
            self.assertEqual((got.kind, got.theme), (ColorKind.THEME, 4))
            self.assertAlmostEqual(got.tint, 0.4, delta=1e-5)
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0).chart(oid).set_series(0, 0, fill={"color": Color(ColorKind.INDEXED, indexed=5)})

            shape = wb.sheet(0).add_shape(
                anchor_kind=AnchorKind.ONE_CELL,
                from_row=2,
                width_emu=9144,
                height_emu=9144,
                preset_geometry="rect",
                fill_kind=ShapePaint.SOLID,
                fill_color=theme,
            )
            paint = shape.get_properties().fill_color
            self.assertEqual((paint.kind, paint.theme), (ColorKind.THEME, 4))
            self.assertAlmostEqual(paint.tint, 0.4, delta=1e-5)
            with self.assertRaises(InvalidArgumentError):
                shape.update(fill_kind=ShapePaint.SOLID, fill_color=Color(ColorKind.INDEXED, indexed=5))
            # A colour is never dropped: without its kind it is a TypeError, with NONE a status error.
            with self.assertRaises(TypeError):
                shape.update(fill_color=theme)
            with self.assertRaises(TypeError):
                shape.update(line_kind=None, line_color=theme)
            with self.assertRaises(InvalidArgumentError):
                shape.update(fill_kind=ShapePaint.NONE, fill_color=theme)
            self.assertEqual(shape.get_properties().fill_color.theme, 4)

    def test_series_patch(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.sheet(0).chart(oid).set_series(
                0,
                1,
                name="Sheet1!$D$1",
                fill={"color": "#FFFF0000", "transparency": 25.0},
                line={"color": "#FF00FF00", "width_pt": 2.5, "dash": ChartLineDash.DASH},
            )
            s = wb.sheet(0).chart(oid).get_properties().plots[0].series[1]
            self.assertEqual(s.name.formula, "Sheet1!$D$1")
            self.assertEqual((s.fill.color, s.fill.transparency), (_rgb(0xFFFF0000), 25.0))
            self.assertEqual((s.line.color, s.line.width_pt, s.line.dash), (_rgb(0xFF00FF00), 2.5, ChartLineDash.DASH))
            self.assertEqual(wb.sheet(0).chart(oid).get_data().series[1].name, "S3")

            wb.sheet(0).chart(oid).set_series(0, 1, fill={"transparency": None}, name="")
            s = wb.sheet(0).chart(oid).get_properties().plots[0].series[1]
            self.assertEqual(s.name.kind, ChartSourceKind.NONE)
            self.assertIsNone(s.fill.transparency)
            self.assertEqual(s.fill.color, _rgb(0xFFFF0000))

            wb.sheet(0).chart(oid).set_series(
                0, 0, labels={"show_value": True, "separator": "; ", "position": ChartLabelPosition.OUTSIDE_END}
            )
            labels = wb.sheet(0).chart(oid).get_properties().plots[0].series[0].labels
            self.assertIsNotNone(labels)
            self.assertEqual((labels.show_value, labels.show_category, labels.separator), (True, False, "; "))
            self.assertEqual(labels.position, ChartLabelPosition.OUTSIDE_END)
            wb.sheet(0).chart(oid).set_series(0, 0, labels=None)
            self.assertIsNone(wb.sheet(0).chart(oid).get_properties().plots[0].series[0].labels)

            wb.sheet(0).chart(oid).set_series(0, 2, values="Sheet1!$B$2:$B$5")
            self.assertEqual(wb.sheet(0).chart(oid).get_data().series[2].values, [2.0, 4.0, 6.0, 8.0])

            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_series(0, 2, values="")
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_series(0, 2, values="not a range")
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_series(0, 77, name="Sheet1!$B$1")
            with self.assertRaises(TypeError):
                wb.sheet(0).chart(oid).set_series(0, 0, fil={})  # type: ignore[call-arg]
            with self.assertRaises(TypeError):
                wb.sheet(0).chart(oid).set_series(0, 0, fill={"colour": 1})  # type: ignore[typeddict-unknown-key]

    def test_line_series_marker_and_smooth(self) -> None:
        wb, oid = _chart(ChartInsertKind.LINE)
        with wb:
            wb.sheet(0).chart(oid).set_series(
                0,
                0,
                marker={
                    "symbol": ChartMarkerSymbol.DIAMOND,
                    "size": 9.0,
                    "fill": {"color": "#FF112233"},
                    "line": {"none": True},
                },
                smooth=True,
            )
            s = wb.sheet(0).chart(oid).get_properties().plots[0].series[0]
            self.assertIsNotNone(s.marker)
            self.assertEqual((s.marker.symbol, s.marker.size), (ChartMarkerSymbol.DIAMOND, 9))
            self.assertEqual(s.marker.fill.color, _rgb(0xFF112233))
            self.assertTrue(s.marker.line.none)
            self.assertIs(s.smooth, True)
            wb.sheet(0).chart(oid).set_series(0, 0, marker=None)
            self.assertIsNone(wb.sheet(0).chart(oid).get_properties().plots[0].series[0].marker)
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_series(0, 0, smooth=None)

    def test_point_patch_and_removal(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.sheet(0).chart(oid).set_point(0, 0, 2, fill={"color": "#FF0000FF"})
            points = wb.sheet(0).chart(oid).get_properties().plots[0].series[0].points
            self.assertEqual([(p.index, p.fill.color) for p in points], [(2, _rgb(0xFF0000FF))])
            self.assertIsNone(points[0].marker)
            wb.sheet(0).chart(oid).set_point(0, 0, 2)
            self.assertEqual(wb.sheet(0).chart(oid).get_properties().plots[0].series[0].points, [])
            with self.assertRaises(TypeError):
                wb.sheet(0).chart(oid).set_point(0, 0, 2, colour=1)  # type: ignore[call-arg]

    def test_plot_patch(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.sheet(0).chart(oid).set_plot(0, gap_width=80, overlap=-20)
            plot = wb.sheet(0).chart(oid).get_properties().plots[0]
            self.assertEqual((plot.gap_width, plot.overlap), (80, -20))
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_plot(0, hole_size=50)
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_plot(5, gap_width=80)
        wb, oid = _chart(ChartInsertKind.DOUGHNUT, "Sheet1!$A$1:$B$5")
        with wb:
            wb.sheet(0).chart(oid).set_plot(0, hole_size=60, first_slice_angle=90, vary_colors=True)
            plot = wb.sheet(0).chart(oid).get_properties().plots[0]
            self.assertEqual((plot.type, plot.hole_size, plot.first_slice_angle), (ChartPlotType.DOUGHNUT, 60, 90))
            self.assertTrue(plot.vary_colors)

    def test_filter(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.sheet(0).chart(oid).set_filter([1], [2])
            info = wb.sheet(0).chart(oid).get_properties()
            self.assertEqual(info.hidden_categories, [2])
            self.assertEqual([s.index for s in info.plots[0].series if s.filtered], [1])
            data = wb.sheet(0).chart(oid).get_data()
            self.assertEqual([s.index for s in data.series if s.filtered], [1])
            wb.sheet(0).chart(oid).set_filter([], [])
            info = wb.sheet(0).chart(oid).get_properties()
            self.assertEqual(info.hidden_categories, [])
            self.assertFalse(any(s.filtered for s in info.plots[0].series))

    def test_add_and_remove_series(self) -> None:
        wb, oid = _chart(source="Sheet1!$A$1:$B$5")
        with wb:
            self.assertEqual(len(wb.sheet(0).chart(oid).get_properties().plots[0].series), 1)
            idx = (
                wb.sheet(0)
                .chart(oid)
                .add_series(0, name="Sheet1!$C$1", values="Sheet1!$C$2:$C$5", categories="Sheet1!$A$2:$A$5")
            )
            series = wb.sheet(0).chart(oid).get_properties().plots[0].series
            self.assertEqual([s.index for s in series], [0, idx])
            self.assertEqual(wb.sheet(0).chart(oid).get_data().series[1].values, [3.0, 6.0, 9.0, 12.0])
            wb.sheet(0).chart(oid).remove_series(0, 0)
            self.assertEqual([s.index for s in wb.sheet(0).chart(oid).get_properties().plots[0].series], [idx])
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).add_series(0, name="Sheet1!$C$1")
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).remove_series(0, 99)

    def test_set_source(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.sheet(0).chart(oid).set_source("Sheet1!$A$1:$C$5")
            info = wb.sheet(0).chart(oid).get_properties()
            self.assertEqual(len(info.plots[0].series), 2)
            self.assertEqual(info.source.range, "Sheet1!$A$1:$C$5")
            wb.sheet(0).chart(oid).set_source(
                "Sheet1!$A$1:$D$5", direction=ChartSeriesDirection.ROWS, header_row=True, header_column=True
            )
            info = wb.sheet(0).chart(oid).get_properties()
            self.assertEqual(info.source.direction, ChartSeriesDirection.ROWS)
            self.assertEqual(len(info.plots[0].series), 4)
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_source("not a range")

    def test_set_type(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.sheet(0).chart(oid).set_type(0, ChartInsertKind.LINE)
            plot = wb.sheet(0).chart(oid).get_properties().plots[0]
            self.assertEqual((plot.type, len(plot.series)), (ChartPlotType.LINE, 3))
            wb.sheet(0).chart(oid).set_type(0, ChartInsertKind.COLUMN, ChartGrouping.STACKED)
            plot = wb.sheet(0).chart(oid).get_properties().plots[0]
            self.assertEqual(
                (plot.type, plot.bar_direction, plot.grouping),
                (ChartPlotType.BAR, ChartBarDirection.COL, ChartGrouping.STACKED),
            )
            wb.sheet(0).chart(oid).set_type(0, ChartInsertKind.PIE)
            self.assertEqual(wb.sheet(0).chart(oid).get_properties().plots[0].type, ChartPlotType.PIE)
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_type(3, ChartInsertKind.LINE)
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_type(0, 99)


class ChartDrawingTests(unittest.TestCase):
    def test_anchor_and_z_order(self) -> None:
        wb, oid = _chart()
        with wb:
            wb.sheet(0).drawing(oid).set_anchor(
                DrawingAnchor(
                    anchor_kind=AnchorKind.TWO_CELL,
                    edit_as=AnchorEditAs.ONE_CELL,
                    row=3,
                    col=2,
                    width_emu=2_000_000,
                    height_emu=1_000_000,
                )
            )
            obj = wb.sheet(0).list_drawings()[0].get_drawing_properties()
            self.assertEqual((obj.object_id, obj.kind), (oid, DrawingObjectKind.CHART))
            self.assertEqual((obj.from_row, obj.from_col, obj.width_emu, obj.height_emu), (3, 2, 2_000_000, 1_000_000))
            self.assertEqual(obj.edit_as, AnchorEditAs.ONE_CELL)

            other = (
                wb.sheet(0).add_chart(ChartInsertKind.PIE, "Sheet1!$A$1:$B$5", anchor=DrawingAnchor(row=20)).object_id
            )
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawings()], [oid, other])
            wb.sheet(0).drawing(other).set_z_order(0)
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawings()], [other, oid])
            with self.assertRaises(FormulonError):
                wb.sheet(0).drawing(oid).set_z_order(2)

    def test_default_frame_and_explicit_size(self) -> None:
        wb, oid = _chart()
        with wb:
            obj = wb.sheet(0).list_drawings()[0].get_drawing_properties()
            self.assertEqual((obj.width_emu, obj.height_emu), (360 * 12700, 216 * 12700))
            sized = (
                wb.sheet(0)
                .add_chart(
                    ChartInsertKind.AREA,
                    SOURCE,
                    anchor=DrawingAnchor(anchor_kind=AnchorKind.ONE_CELL, width_emu=500_000, height_emu=400_000),
                )
                .object_id
            )
            by_id = {o.object_id: o.get_drawing_properties() for o in wb.sheet(0).list_drawings()}
            self.assertEqual((by_id[sized].width_emu, by_id[sized].height_emu), (500_000, 400_000))
            self.assertEqual(by_id[sized].anchor_kind, AnchorKind.ONE_CELL)

    def test_snapshot_remove_restore(self) -> None:
        wb, a = _chart(title="A")
        with wb:
            b = (
                wb.sheet(0)
                .add_chart(ChartInsertKind.PIE, "Sheet1!$A$1:$B$5", anchor=DrawingAnchor(row=20), title="B")
                .object_id
            )
            snap = wb.sheet(0).drawing(a).snapshot()
            self.assertIsInstance(snap, bytes)
            self.assertGreater(len(snap), 0)

            wb.sheet(0).remove_drawing(a)
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawings()], [b])
            with self.assertRaises(NotFoundError):
                wb.sheet(0).chart(a).get_properties()
            with self.assertRaises(FormulonError):
                wb.sheet(0).remove_drawing(a)

            self.assertEqual(wb.sheet(0).restore_drawing(snap).object_id, a)
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawings()], [a, b])
            restored = wb.sheet(0).chart(a).get_properties()
            self.assertEqual((restored.title.text, len(restored.plots[0].series)), ("A", 3))

            wb.sheet(0).chart(a).set_title("Changed")
            self.assertEqual(wb.sheet(0).restore_drawing(snap).object_id, a)
            self.assertEqual(wb.sheet(0).chart(a).get_properties().title.text, "A")

            copy = wb.sheet(0).restore_drawing(snap, new_id=True).object_id
            self.assertNotIn(copy, (a, b))
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawings()], [a, b, copy])
            self.assertEqual(wb.sheet(0).chart(copy).get_properties().title.text, "A")

            with self.assertRaises(FormulonError):
                wb.sheet(0).restore_drawing(b"junk")
            with self.assertRaises(FormulonError):
                wb.sheet(0).drawing(9999).snapshot()


class ChartErrorTests(unittest.TestCase):
    def test_unknown_chart_and_sheet(self) -> None:
        wb, oid = _chart()
        with wb:
            with self.assertRaises(NotFoundError):
                wb.sheet(0).chart(9999)
            for call in (
                lambda: wb.sheet(0).chart(9999).get_data(),
                lambda: wb.sheet(9).chart(oid).get_properties(),
                lambda: wb.sheet(0).chart(9999).set_title("x"),
                lambda: wb.sheet(0).chart(9999).set_filter([], []),
            ):
                with self.assertRaises(FormulonError):
                    call()

    def test_insert_rejects_bad_input(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb)
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_chart(ChartInsertKind.COLUMN, "not a range")
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_chart(99, SOURCE)
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_chart(
                    ChartInsertKind.COLUMN, SOURCE, anchor=DrawingAnchor(anchor_kind=AnchorKind.ABSOLUTE)
                )
            with self.assertRaises(ValueError):
                wb.sheet(0).add_chart(ChartInsertKind.COLUMN, SOURCE, anchor=DrawingAnchor(row=-1))
            self.assertEqual(wb.sheet(0).list_drawings(), [])

    def test_failed_edit_leaves_the_chart_unchanged(self) -> None:
        wb, oid = _chart()
        with wb:
            before = wb.sheet(0).chart(oid).get_properties()
            with self.assertRaises(FormulonError):
                wb.sheet(0).chart(oid).set_series(0, 0, fill={"color": "#FF000000"}, values="not a range")
            self.assertEqual(wb.sheet(0).chart(oid).get_properties(), before)


if __name__ == "__main__":
    unittest.main()
