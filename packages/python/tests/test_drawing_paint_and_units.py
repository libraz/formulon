"""Drawing paint records, unit-suffixed names and flat chart input."""

from __future__ import annotations

import dataclasses
import inspect
import struct
import unittest
import zlib

import formulon
from formulon import (
    AnchorKind,
    CellRect,
    Chart,
    ChartBarDirection,
    ChartFill,
    ChartInput,
    ChartLine,
    ChartPlotType,
    Color,
    ColorKind,
    DrawingAnchorInput,
    DrawingPlacementInput,
    FormulonError,
    GradientFill,
    ImageFormat,
    ImageInfo,
    InvalidArgumentError,
    Picture,
    PictureProperties,
    Shape,
    ShapeInput,
    Slicer,
    SlicerInput,
    SparklineGroupInput,
    SparklineInput,
    TableInput,
    Workbook,
)

SOURCE = "Sheet1!$A$1:$C$4"
RED = Color(ColorKind.RGB, argb=0xFFFF0000)


def _grid(wb: Workbook) -> None:
    for c, name in enumerate(("S1", "S2")):
        wb.sheet(0).set_text(0, c + 1, name)
    for r, cat in enumerate("abc"):
        wb.sheet(0).set_text(r + 1, 0, cat)
        for c in range(2):
            wb.sheet(0).set_number(r + 1, c + 1, float((r + 1) * (c + 2)))


def _png(width: int, height: int) -> bytes:
    def chunk(tag: bytes, data: bytes) -> bytes:
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

    raw = b"".join(b"\x00" + b"\x00\x00\x00" * width for _ in range(height))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )


def _box(**overrides):
    fields = dict(
        anchor=DrawingAnchorInput(kind=AnchorKind.TWO_CELL, from_row=1, from_col=1, to_row=5, to_col=4),
        name="Box",
        preset_geometry="rect",
        fill=ChartFill(color=RED),
        line=ChartLine(width_pt=2.0),
    )
    fields.update(overrides)
    return fields


class DrawingUpdateTests(unittest.TestCase):
    def test_shape_update_atomic(self) -> None:
        with Workbook.create_default() as wb:
            shape = wb.sheet(0).add_shape(ShapeInput(**_box()))
            with self.assertRaises(FormulonError):
                shape.update(name="Renamed", line={"width_pt": -1.0})
            self.assertEqual(shape.get_drawing_properties().name, "Box")
            connector = wb.sheet(0).add_shape(ShapeInput(**_box(name="Wire", preset_geometry="line", fill=None)))
            with self.assertRaises(FormulonError):
                connector.update(name="Renamed", text="not allowed on a connector")
            self.assertEqual(connector.get_drawing_properties().name, "Wire")

    def test_slicer_update_atomic(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "Product")
            wb.sheet(0).set_text(0, 1, "Amount")
            wb.sheet(0).set_number(1, 1, 10.0)
            wb.sheet(0).add_table(TableInput(range=CellRect(0, 0, 1, 1), name="Sales", columns=["Product", "Amount"]))
            slicer = wb.sheet(0).add_slicer(
                "Sales", 0, SlicerInput(anchor=DrawingAnchorInput(from_row=4, from_col=3, to_row=14, to_col=5))
            )
            with self.assertRaises(FormulonError):
                slicer.update(hidden=True, column_count=0)
            self.assertFalse(slicer.get_drawing_properties().hidden)


class ShapePaintTests(unittest.TestCase):
    def test_fill_line_records(self) -> None:
        with Workbook.create_default() as wb:
            shape = wb.sheet(0).add_shape(
                ShapeInput(
                    **_box(fill=ChartFill(color=RED, transparency_pct=25.0), line=ChartLine(color=RED, width_pt=2.0))
                )
            )
            props = shape.get_properties()
            self.assertIsInstance(props.fill, ChartFill)
            self.assertIsInstance(props.line, ChartLine)
            self.assertEqual((props.fill.color, props.fill.transparency_pct), (RED, 25.0))
            self.assertEqual((props.line.color, props.line.width_pt), (RED, 2.0))
            self.assertFalse(props.fill.none or props.fill.other or props.line.other)
            names = {f.name for f in dataclasses.fields(type(props))}
            self.assertFalse(names & {"fill_kind", "fill_color", "line_kind", "line_color", "line_width_pt"})
            self.assertFalse(hasattr(formulon, "ShapePaint"))
            # A read record feeds back into the patch.
            shape.update(fill=props.fill, line=props.line)
            self.assertEqual(shape.get_properties().fill, props.fill)

    def test_fill_none_clears(self) -> None:
        with Workbook.create_default() as wb:
            shape = wb.sheet(0).add_shape(ShapeInput(**_box()))
            shape.update(fill=ChartFill(none=True))
            self.assertTrue(shape.get_properties().fill.none)
            shape.update(fill=None, line=None)
            props = shape.get_properties()
            self.assertIsNone(props.fill)
            self.assertIsNone(props.line)
            with Workbook.load(wb.save()) as loaded:
                again = loaded.sheet(0).list_shapes()[0].get_properties()
                self.assertIsNone(again.fill)
            plain = wb.sheet(0).add_shape(
                ShapeInput(anchor=DrawingAnchorInput(from_row=1, from_col=1, to_row=3, to_col=3))
            )
            self.assertIsNone(plain.get_properties().fill)
            self.assertIsNone(plain.get_properties().line)

    def test_series_fill_null_when_unstated(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb)
            chart = wb.sheet(0).add_chart(ChartInput(type=ChartPlotType.BAR, range=SOURCE))
            series = chart.get_properties().plots[0].series[0]
            # A new chart states its series paint, as Excel's does; None removes it.
            chart.update_series(0, series.series_id, fill=None, line=None)
            cleared = chart.get_properties().plots[0].series[0]
            self.assertIsNone(cleared.fill)
            self.assertIsNone(cleared.line)
            chart.update_series(0, series.series_id, fill=ChartFill(color=RED))
            self.assertEqual(chart.get_properties().plots[0].series[0].fill.color, RED)
            chart.update_series(0, series.series_id, fill=None)
            self.assertIsNone(chart.get_properties().plots[0].series[0].fill)


class SlicerTests(unittest.TestCase):
    def test_row_height_pt(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "Product")
            wb.sheet(0).set_text(0, 1, "Amount")
            wb.sheet(0).set_number(1, 1, 10.0)
            wb.sheet(0).add_table(TableInput(range=CellRect(0, 0, 1, 1), name="Sales", columns=["Product", "Amount"]))
            anchor = DrawingAnchorInput(from_row=4, from_col=3, to_row=14, to_col=5)
            slicer = wb.sheet(0).add_slicer("Sales", 0, SlicerInput(anchor=anchor, row_height_pt=30.0))
            self.assertEqual(slicer.get_properties().row_height_pt, 30.0)
            slicer.update(row_height_pt=19.0)
            self.assertEqual(slicer.get_properties().row_height_pt, 19.0)
            names = {f.name for f in dataclasses.fields(formulon.SlicerProperties)}
            self.assertNotIn("row_height_emu", names)
            with Workbook.load(wb.save()) as loaded:
                self.assertEqual(loaded.sheet(0).list_slicers()[0].get_properties().row_height_pt, 19.0)
            with self.assertRaises(FormulonError):
                slicer.update(row_height_pt=-1.0)
            with self.assertRaises(TypeError):
                slicer.update(row_height_emu=1)  # type: ignore[call-arg]

    def test_update_has_no_name_keyword(self) -> None:
        self.assertNotIn("name", inspect.signature(Slicer.update).parameters)
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "Product")
            wb.sheet(0).set_text(0, 1, "Amount")
            wb.sheet(0).set_number(1, 1, 10.0)
            wb.sheet(0).add_table(TableInput(range=CellRect(0, 0, 1, 1), name="Sales", columns=["Product", "Amount"]))
            slicer = wb.sheet(0).add_slicer(
                "Sales", 0, SlicerInput(anchor=DrawingAnchorInput(from_row=4, from_col=3, to_row=14, to_col=5))
            )
            with self.assertRaises(TypeError):
                slicer.update(name="X")  # type: ignore[call-arg]
            self.assertEqual(slicer.name, "Product")


class DrawingHandleTests(unittest.TestCase):
    def test_restore_drawing_empty_bytes_is_invalid_argument(self) -> None:
        with Workbook.create_default() as wb:
            for empty in (b"", bytearray(), memoryview(b"")):
                with self.assertRaises(InvalidArgumentError):
                    wb.sheet(0).restore_drawing(empty)
            self.assertEqual(wb.sheet(0).list_drawings(), [])

    def test_z_order_and_snapshot_are_inherited(self) -> None:
        for cls in (Picture, Chart, Shape):
            self.assertNotIn("set_z_order", vars(cls))
            self.assertNotIn("snapshot", vars(cls))
            self.assertTrue(callable(getattr(cls, "set_z_order")))
            self.assertTrue(callable(getattr(cls, "snapshot")))
        self.assertNotIn("set_z_order", vars(Slicer))
        self.assertTrue(callable(getattr(Slicer, "set_z_order")))
        with Workbook.create_default() as wb:
            _grid(wb)
            ws = wb.sheet(0)
            pic = ws.add_picture(_png(2, 2), DrawingPlacementInput(name="logo"))
            chart = ws.add_chart(ChartInput(type=ChartPlotType.BAR, range=SOURCE))
            shape = ws.add_shape(ShapeInput(**_box()))
            self.assertEqual(
                [d.object_id for d in ws.list_drawings()], [pic.object_id, chart.object_id, shape.object_id]
            )
            shape.set_z_order(0)
            self.assertEqual(
                [d.object_id for d in ws.list_drawings()], [shape.object_id, pic.object_id, chart.object_id]
            )
            pic.set_z_order(2)
            self.assertEqual(
                [d.object_id for d in ws.list_drawings()], [shape.object_id, chart.object_id, pic.object_id]
            )
            for handle in (pic, chart, shape):
                captured = handle.snapshot()
                self.assertIsInstance(captured, bytes)
                self.assertGreater(len(captured), 0)
            captured = shape.snapshot()
            ws.remove_drawing(shape)
            restored = ws.restore_drawing(captured)
            self.assertIsInstance(restored, Shape)
            self.assertEqual(restored.object_id, shape.object_id)

            ws.set_text(20, 0, "Product")
            ws.set_text(20, 1, "Amount")
            ws.set_number(21, 1, 10.0)
            ws.add_table(TableInput(range=CellRect(20, 0, 21, 1), name="Sales", columns=["Product", "Amount"]))
            slicer = ws.add_slicer(
                "Sales", 0, SlicerInput(anchor=DrawingAnchorInput(from_row=4, from_col=9, to_row=14, to_col=11))
            )
            slicer.set_z_order(0)
            self.assertEqual(ws.list_drawings()[0].object_id, slicer.object_id)


class ChartInputTests(unittest.TestCase):
    def test_chart_input_flat_anchor(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb)
            chart = wb.sheet(0).add_chart(
                ChartInput(
                    type=ChartPlotType.BAR,
                    bar_direction=ChartBarDirection.COLUMN,
                    range=SOURCE,
                    name="Sales",
                    description="d",
                    anchor=DrawingAnchorInput(kind=AnchorKind.TWO_CELL, from_row=6),
                )
            )
            props = chart.get_drawing_properties()
            self.assertEqual((props.name, props.description), ("Sales", "d"))
            self.assertEqual((props.anchor.kind, props.anchor.from_row), (AnchorKind.TWO_CELL, 6))
            self.assertTrue(issubclass(ChartInput, DrawingPlacementInput))
            self.assertNotIn("placement", {f.name for f in dataclasses.fields(ChartInput)})
            self.assertFalse(hasattr(formulon, "ChartInsertInput"))
            with self.assertRaises(TypeError):
                ChartInput(  # type: ignore[call-arg]
                    type=ChartPlotType.BAR, range=SOURCE, placement=DrawingPlacementInput(name="x")
                )


class UnitSuffixTests(unittest.TestCase):
    def test_plot_sparkline_gradient(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb)
            chart = wb.sheet(0).add_chart(ChartInput(type=ChartPlotType.BAR, range=SOURCE))
            chart.update_plot(0, gap_width_pct=80, overlap_pct=-20)
            plot = chart.get_properties().plots[0]
            self.assertEqual((plot.gap_width_pct, plot.overlap_pct), (80, -20))
            doughnut = wb.sheet(0).add_chart(ChartInput(type=ChartPlotType.DOUGHNUT, range="Sheet1!$A$1:$B$4"))
            doughnut.update_plot(0, hole_size_pct=60)
            self.assertEqual(doughnut.get_properties().plots[0].hole_size_pct, 60)
            names = {f.name for f in dataclasses.fields(type(plot))}
            self.assertFalse(names & {"overlap", "hole_size"})

            group_id = wb.sheet(0).add_sparkline_group(
                SparklineGroupInput(line_width_pt=1.25, sparklines=[SparklineInput("Sheet1!B2:B4", "E1")])
            )
            self.assertEqual(wb.sheet(0).list_sparkline_groups()[0].line_width_pt, 1.25)
            wb.sheet(0).update_sparkline_group(group_id, line_width_pt=2.0)
            self.assertEqual(wb.sheet(0).list_sparkline_groups()[0].line_width_pt, 2.0)
            self.assertFalse(hasattr(wb.sheet(0).list_sparkline_groups()[0], "line_weight"))

            names = {f.name for f in dataclasses.fields(GradientFill)}
            self.assertIn("angle_deg", names)
            self.assertNotIn("degree", names)
            self.assertEqual(GradientFill(angle_deg=45.0).angle_deg, 45.0)
            fill = formulon.Fill(gradient=GradientFill(angle_deg=45.0, stops=[]))
            self.assertEqual(fill.gradient.angle_deg, 45.0)


class ChartDataTests(unittest.TestCase):
    def test_unresolved_role(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb)
            chart = wb.sheet(0).add_chart(ChartInput(type=ChartPlotType.BAR, range=SOURCE))
            self.assertEqual(chart.get_data().unresolved, [])
            names = {f.name for f in dataclasses.fields(formulon.ChartUnresolved)}
            self.assertIn("role", names)
            self.assertNotIn("source", names)

    def test_series_number_format(self) -> None:
        with Workbook.create_default() as wb:
            _grid(wb)
            chart = wb.sheet(0).add_chart(ChartInput(type=ChartPlotType.BAR, range=SOURCE))
            series = chart.get_data().series[0]
            self.assertEqual(series.number_format, "General")
            self.assertIsNone(series.x_number_format)
            names = {f.name for f in dataclasses.fields(formulon.ChartDataSeries)}
            self.assertFalse(names & {"format_code", "x_format_code"})


class PictureTests(unittest.TestCase):
    def test_picture_properties_is_image_info(self) -> None:
        self.assertTrue(issubclass(PictureProperties, ImageInfo))
        with Workbook.create_default() as wb:
            png = _png(4, 3)
            pic = wb.sheet(0).add_picture(png, DrawingPlacementInput(name="logo"))
            props = pic.get_properties()
            self.assertIsInstance(props, ImageInfo)
            self.assertEqual((props.format, props.width_px, props.height_px, props.bytes), (ImageFormat.PNG, 4, 3, png))
            self.assertEqual(
                {f.name for f in dataclasses.fields(PictureProperties)}, {"format", "width_px", "height_px", "bytes"}
            )


if __name__ == "__main__":
    unittest.main()
