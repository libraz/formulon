"""Drawing shape, sparkline group and legacy note shape surface tests."""

from __future__ import annotations

import io
import re
import unittest
import zipfile

from formulon import (
    AnchorKind,
    BlanksAs,
    Color,
    ColorKind,
    Comment,
    DrawingObjectKind,
    FormulonError,
    InvalidArgumentError,
    Shape,
    ShapeInput,
    ShapePaint,
    ShapeProperties,
    Sparkline,
    SparklineAxisType,
    SparklineGroup,
    SparklineType,
    Workbook,
)


def _box(**overrides):
    fields = dict(
        anchor_kind=AnchorKind.TWO_CELL,
        from_row=1,
        from_col=1,
        to_row=5,
        to_col=4,
        name="Box",
        preset_geometry="roundRect",
        text="Hello\nWorld",
        fill_kind=ShapePaint.SOLID,
        fill_color="#FF0000",
        line_kind=ShapePaint.NONE,
        rotation_deg=15.0,
    )
    fields.update(overrides)
    return fields


class ShapeTests(unittest.TestCase):
    def test_insert_list_update_remove_and_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.sheet(0).list_shapes(), [])
            rect = wb.sheet(0).add_shape(ShapeInput(**_box()))
            line = wb.sheet(0).add_shape(
                ShapeInput(
                    anchor_kind=AnchorKind.ONE_CELL,
                    from_row=10,
                    from_col_offset_emu=500,
                    width_emu=914400,
                    height_emu=0,
                    preset_geometry="line",
                    line_kind=ShapePaint.SOLID,
                    line_color="#00FF00",
                    line_width_pt=1.5,
                ),
            )
            self.assertIsInstance(rect, Shape)
            self.assertNotEqual(rect.object_id, line.object_id)
            with Workbook.load(wb.save()) as loaded:
                handles = loaded.sheet(0).list_shapes()
                self.assertEqual(len(handles), 2)
                shapes = [h.get_properties() for h in handles]
                self.assertIsInstance(shapes[0], ShapeProperties)
                self.assertEqual(shapes[0].object_id, rect.object_id)
                self.assertEqual(shapes[0].kind, DrawingObjectKind.SHAPE)
                self.assertEqual(
                    (shapes[0].name, shapes[0].preset_geometry, shapes[0].text), ("Box", "roundRect", "Hello\nWorld")
                )
                self.assertEqual(
                    (shapes[0].fill_kind, shapes[0].fill_color),
                    (ShapePaint.SOLID, Color(ColorKind.RGB, argb=0xFFFF0000)),
                )
                self.assertEqual(shapes[0].line_kind, ShapePaint.NONE)
                self.assertEqual((shapes[0].rotation_deg, shapes[0].hidden, shapes[0].text_box), (15.0, False, False))
                self.assertEqual((shapes[0].from_row, shapes[0].to_row, shapes[0].to_col), (1, 5, 4))
                self.assertEqual(shapes[1].kind, DrawingObjectKind.CONNECTOR)
                self.assertEqual(
                    (shapes[1].from_col_offset_emu, shapes[1].width_emu, shapes[1].line_color),
                    (500, 914400, Color(ColorKind.RGB, argb=0xFF00FF00)),
                )
                self.assertEqual(shapes[1].line_width_pt, 1.5)

                loaded.sheet(0).shape(rect.object_id).update(name="Renamed", hidden=True, fill_kind=ShapePaint.NONE)
                loaded.sheet(0).insert_rows(0, 3)
                moved = loaded.sheet(0).list_shapes()[0].get_properties()
                self.assertEqual((moved.name, moved.hidden, moved.fill_kind), ("Renamed", True, ShapePaint.NONE))
                self.assertEqual((moved.preset_geometry, moved.rotation_deg), ("roundRect", 15.0))
                self.assertEqual((moved.from_row, moved.to_row), (4, 8))

                loaded.sheet(0).remove_drawing(rect.object_id)
                self.assertEqual([s.object_id for s in loaded.sheet(0).list_shapes()], [line.object_id])

    def test_arguments_are_validated(self) -> None:
        with Workbook.create_default() as wb:
            for bad in (
                _box(anchor_kind=9),
                _box(fill_kind=ShapePaint.OTHER),
                _box(to_row=0),
                _box(preset_geometry='a"/>'),
            ):
                with self.assertRaises(FormulonError):
                    wb.sheet(0).add_shape(ShapeInput(**bad))
            with self.assertRaises(FormulonError):
                wb.sheet(5).add_shape(ShapeInput(**_box()))
            self.assertEqual(wb.sheet(0).list_shapes(), [])
            with self.assertRaises(FormulonError):
                wb.sheet(0).shape(99).update(name="x")
            with self.assertRaises(FormulonError):
                wb.sheet(0).remove_drawing(99)


def _group(**overrides) -> SparklineGroup:
    group = SparklineGroup(
        type=SparklineType.COLUMN,
        color_series=Color(kind=1, argb=0xFF336699),
        color_negative=Color(kind=2, theme=5, tint=-0.25),
        color_low=Color(kind=3, indexed=12),
        markers=True,
        negative=True,
        display_x_axis=True,
        min_axis_type=SparklineAxisType.CUSTOM,
        manual_min=-1.5,
        line_weight=1.25,
        empty_cells_as=BlanksAs.ZERO,
        sparklines=[Sparkline("Sheet1!A1:D1", "F1"), Sparkline("Sheet1!A2:D2", "F2")],
    )
    for key, value in overrides.items():
        setattr(group, key, value)
    return group


class SparklineGroupTests(unittest.TestCase):
    def test_add_list_update_replace_remove_and_row_edit_shift(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.sheet(0).list_sparkline_groups(), [])
            wb.sheet(0).add_sparkline_group(_group())
            wb.sheet(0).add_sparkline_group(SparklineGroup(sparklines=[Sparkline("Sheet1!A9:D9", "G9")]))
            with Workbook.load(wb.save()) as loaded:
                groups = loaded.sheet(0).list_sparkline_groups()
                self.assertEqual(len(groups), 2)
                first, second = groups
                self.assertEqual(first.type, SparklineType.COLUMN)
                self.assertEqual((first.color_series.kind, first.color_series.argb), (1, 0xFF336699))
                self.assertEqual((first.color_negative.theme, first.color_negative.tint), (5, -0.25))
                self.assertEqual(first.color_low.indexed, 12)
                self.assertEqual(first.color_high.kind, 0)
                self.assertEqual((first.markers, first.high, first.negative), (True, False, True))
                self.assertEqual(first.min_axis_type, SparklineAxisType.CUSTOM)
                self.assertEqual((first.manual_min, first.manual_max), (-1.5, None))
                self.assertEqual(first.line_weight, 1.25)
                self.assertEqual(first.empty_cells_as, BlanksAs.ZERO)
                self.assertEqual(first.sparklines, [Sparkline("=Sheet1!A1:D1", "F1"), Sparkline("=Sheet1!A2:D2", "F2")])
                self.assertEqual((second.type, second.line_weight), (SparklineType.LINE, 0.75))

                loaded.sheet(0).insert_rows(0, 2)
                self.assertEqual(
                    loaded.sheet(0).list_sparkline_groups()[0].sparklines[1],
                    Sparkline("=Sheet1!A4:D4", "F4"),
                )

                loaded.sheet(0).update_sparkline_group(
                    0, type=SparklineType.STACKED, high=True, date_axis=True, date_axis_ref="Sheet1!A4:D4"
                )
                updated = loaded.sheet(0).list_sparkline_groups()[0]
                # A partial update leaves every omitted field as it was.
                self.assertEqual((updated.markers, updated.negative, updated.manual_min), (True, True, -1.5))
                self.assertEqual((updated.type, updated.high, updated.date_axis), (SparklineType.STACKED, True, True))
                self.assertEqual(updated.date_axis_ref, "=Sheet1!A4:D4")
                self.assertEqual(len(updated.sparklines), 2)

                loaded.sheet(0).update_sparkline_group(0, sparklines=[Sparkline("Sheet1!B1:B5", "H1")])
                self.assertEqual(
                    loaded.sheet(0).list_sparkline_groups()[0].sparklines,
                    [Sparkline("=Sheet1!B1:B5", "H1")],
                )
                with self.assertRaises(InvalidArgumentError):
                    loaded.sheet(0).update_sparkline_group(0, sparklines=[])
                self.assertEqual(
                    loaded.sheet(0).list_sparkline_groups()[0].sparklines,
                    [Sparkline("=Sheet1!B1:B5", "H1")],
                )
                loaded.sheet(0).remove_sparkline_group(0)
                remaining = loaded.sheet(0).list_sparkline_groups()
                self.assertEqual(len(remaining), 1)
                self.assertEqual(remaining[0].sparklines[0].location, "G11")

    def test_arguments_are_validated(self) -> None:
        good = [Sparkline("Sheet1!A1:D1", "F1")]
        with Workbook.create_default() as wb:
            for bad in (
                SparklineGroup(),
                SparklineGroup(sparklines=[Sparkline("Sheet1!A1:D1", "F1:F2")]),
                SparklineGroup(sparklines=[Sparkline("", "F1")]),
                SparklineGroup(type=7, sparklines=good),
                SparklineGroup(line_weight=-1.0, sparklines=good),
            ):
                with self.assertRaises(FormulonError):
                    wb.sheet(0).add_sparkline_group(bad)
            with self.assertRaises(FormulonError):
                wb.sheet(4).add_sparkline_group(SparklineGroup(sparklines=good))
            self.assertEqual(wb.sheet(0).list_sparkline_groups(), [])
            with self.assertRaises(FormulonError):
                wb.sheet(0).update_sparkline_group(0, markers=True)
            with self.assertRaises(FormulonError):
                wb.sheet(0).update_sparkline_group(0, sparklines=good)
            with self.assertRaises(FormulonError):
                wb.sheet(0).remove_sparkline_group(0)


class NoteShapeTests(unittest.TestCase):
    def test_surviving_note_keeps_its_own_shape_when_notes_are_added(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_comment(1, 1, Comment("Alice", "one"))
            wb.sheet(0).set_comment(3, 3, Comment("Alice", "two"))
            wb.sheet(0).set_comment_visible(1, 1, True)
            with Workbook.load(wb.save()) as first:
                first.sheet(0).set_comment(8, 8, Comment("Alice", "three"))
                first.sheet(0).set_comment_visible(3, 3, True)
                saved = first.save()
                vml = zipfile.ZipFile(io.BytesIO(saved)).read("xl/drawings/vmlDrawing1.vml").decode("utf-8")
                self.assertEqual(len(re.findall(r"<v:shape ", vml)), 3)
                self.assertEqual(len(set(re.findall(r'id="_x0000_s\d+"', vml))), 3)
                self.assertEqual(vml.count("visibility:visible"), 2)
                with Workbook.load(saved) as again:
                    self.assertEqual(
                        [again.sheet(0).get_comment_visible(r, c) for r, c in ((1, 1), (3, 3), (8, 8))],
                        [True, True, False],
                    )


if __name__ == "__main__":
    unittest.main()
