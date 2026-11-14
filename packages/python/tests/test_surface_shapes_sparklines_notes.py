"""Drawing shape, sparkline group and legacy note shape surface tests."""

from __future__ import annotations

import dataclasses
import io
import re
import unittest
import zipfile

from formulon import (
    AnchorKind,
    BlanksAs,
    CellAddress,
    CellRect,
    ChartFill,
    ChartLine,
    Color,
    ColorKind,
    Comment,
    DrawingAnchorInput,
    DrawingObjectKind,
    FormulonError,
    InvalidArgumentError,
    NotFoundError,
    Shape,
    ShapeInput,
    ShapeProperties,
    SheetRect,
    Sparkline,
    SparklineAxisType,
    SparklineGroupInput,
    SparklineInput,
    SparklineType,
    Workbook,
)


def _box(**overrides):
    fields = dict(
        anchor=DrawingAnchorInput(kind=AnchorKind.TWO_CELL, from_row=1, from_col=1, to_row=5, to_col=4),
        name="Box",
        preset_geometry="roundRect",
        text="Hello\nWorld",
        fill=ChartFill(color="#FF0000"),
        line=ChartLine(none=True),
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
                    anchor=DrawingAnchorInput(
                        kind=AnchorKind.ONE_CELL,
                        from_row=10,
                        from_col_offset_emu=500,
                        width_emu=914400,
                        height_emu=12700,
                    ),
                    preset_geometry="line",
                    line=ChartLine(color="#00FF00", width_pt=1.5),
                ),
            )
            self.assertIsInstance(rect, Shape)
            self.assertNotEqual(rect.object_id, line.object_id)
            with Workbook.load(wb.save()) as loaded:
                handles = loaded.sheet(0).list_shapes()
                self.assertEqual(len(handles), 2)
                shapes = [h.get_properties() for h in handles]
                common = [h.get_drawing_properties() for h in handles]
                self.assertIsInstance(shapes[0], ShapeProperties)
                self.assertEqual(shapes[0].object_id, rect.object_id)
                self.assertEqual(shapes[0].kind, DrawingObjectKind.SHAPE)
                self.assertEqual(
                    (common[0].name, shapes[0].preset_geometry, shapes[0].text), ("Box", "roundRect", "Hello\nWorld")
                )
                self.assertEqual(shapes[0].fill, ChartFill(color=Color(ColorKind.RGB, argb=0xFFFF0000)))
                self.assertEqual(shapes[0].line, ChartLine(none=True))
                self.assertEqual((shapes[0].rotation_deg, common[0].hidden, shapes[0].text_box), (15.0, False, False))
                self.assertEqual(
                    (common[0].anchor.from_row, common[0].anchor.to_row, common[0].anchor.to_col), (1, 5, 4)
                )
                self.assertEqual(shapes[1].kind, DrawingObjectKind.CONNECTOR)
                self.assertEqual(
                    (common[1].anchor.from_col_offset_emu, common[1].anchor.width_emu, shapes[1].line.color),
                    (500, 914400, Color(ColorKind.RGB, argb=0xFF00FF00)),
                )
                self.assertEqual(shapes[1].line.width_pt, 1.5)

                loaded.sheet(0).shape(rect.object_id).update(name="Renamed", hidden=True, fill=ChartFill(none=True))
                loaded.sheet(0).insert_rows(0, 2)
                moved_handle = loaded.sheet(0).list_shapes()[0]
                moved = moved_handle.get_properties()
                moved_common = moved_handle.get_drawing_properties()
                self.assertEqual(
                    (moved_common.name, moved_common.hidden, moved.fill), ("Renamed", True, ChartFill(none=True))
                )
                self.assertEqual((moved.preset_geometry, moved.rotation_deg), ("roundRect", 15.0))
                self.assertEqual((moved_common.anchor.from_row, moved_common.anchor.to_row), (4, 8))

                loaded.sheet(0).remove_drawing(rect.object_id)
                self.assertEqual([s.object_id for s in loaded.sheet(0).list_shapes()], [line.object_id])

    def test_a_shape_without_an_anchor_kind_is_two_cell_and_has_no_flat_keys(self) -> None:
        with Workbook.create_default() as wb:
            shape = wb.sheet(0).add_shape(
                ShapeInput(anchor=DrawingAnchorInput(from_row=1, from_col=1, to_row=4, to_col=4))
            )
            anchor = shape.get_drawing_properties().anchor
            self.assertEqual(
                (anchor.kind, anchor.from_row, anchor.from_col, anchor.to_row, anchor.to_col),
                (AnchorKind.TWO_CELL, 1, 1, 4, 4),
            )
            self.assertIsNone(shape.get_properties().fill)
            self.assertIsNone(shape.get_properties().line)
            for flat in ("anchor_kind", "from_row", "from_col", "to_row", "to_col", "width_emu", "height_emu"):
                with self.subTest(key=flat):
                    with self.assertRaises(TypeError):
                        ShapeInput(**{flat: 1})  # type: ignore[arg-type]
                    with self.assertRaises(TypeError):
                        shape.update(**{flat: 1})  # type: ignore[arg-type]
            names = {f.name for f in dataclasses.fields(ShapeProperties)}
            self.assertFalse(names & {"name", "description", "hidden", "from_row", "anchor_kind", "width_emu"})

    def test_anchor_update_merges_per_key_and_keeps_the_other_corners(self) -> None:
        with Workbook.create_default() as wb:
            shape = wb.sheet(0).add_shape(ShapeInput(**_box()))
            shape.update(anchor=DrawingAnchorInput(to_col=6))
            anchor = shape.get_drawing_properties().anchor
            self.assertEqual(
                (anchor.kind, anchor.from_row, anchor.from_col, anchor.to_row, anchor.to_col),
                (AnchorKind.TWO_CELL, 1, 1, 5, 6),
            )
            shape.update(anchor=DrawingAnchorInput(from_row=3))
            anchor = shape.get_drawing_properties().anchor
            self.assertEqual(
                (anchor.kind, anchor.from_row, anchor.from_col, anchor.to_row), (AnchorKind.TWO_CELL, 3, 1, 5)
            )

    def test_common_keys_update_before_the_shape_keys(self) -> None:
        with Workbook.create_default() as wb:
            shape = wb.sheet(0).add_shape(ShapeInput(**_box()))
            shape.update(name="N", description="D", hidden=True, text="changed", rotation_deg=30.0)
            common = shape.get_drawing_properties()
            self.assertEqual((common.name, common.description, common.hidden), ("N", "D", True))
            self.assertEqual((shape.get_properties().text, shape.get_properties().rotation_deg), ("changed", 30.0))
            self.assertEqual(shape.name, "N")
            # Every argument is checked before anything changes.
            with self.assertRaises(TypeError):
                shape.update(name="M", text=None)  # type: ignore[arg-type]
            self.assertEqual(shape.name, "N")
            # The preset geometry has no cleared form.
            with self.assertRaises(TypeError):
                shape.update(preset_geometry=None)  # type: ignore[arg-type]
            self.assertEqual(shape.get_properties().preset_geometry, "roundRect")

    def test_a_shape_paint_default_is_a_real_state_and_none_removes_the_own_paint(self) -> None:
        with Workbook.create_default() as wb:
            shape = wb.sheet(0).add_shape(ShapeInput(**_box()))
            self.assertIsNotNone(shape.get_properties().fill)
            shape.update(fill=None, line=None)
            props = shape.get_properties()
            self.assertEqual((props.fill, props.line), (None, None))

    def test_a_shape_and_a_group_member_are_removed_by_the_one_remover(self) -> None:
        with Workbook.create_default() as wb:
            shape = wb.sheet(0).add_shape(ShapeInput(**_box()))
            self.assertEqual([d.object_id for d in wb.sheet(0).list_drawings()], [shape.object_id])
            wb.sheet(0).remove_drawing(shape)
            self.assertEqual(wb.sheet(0).list_drawings(), [])
            with self.assertRaises(NotFoundError):
                shape.update(name="x")
            with self.assertRaises(NotFoundError):
                shape.name
            with self.assertRaises(NotFoundError):
                wb.sheet(0).remove_drawing(shape.object_id)

    def test_arguments_are_validated(self) -> None:
        with Workbook.create_default() as wb:
            for bad in (
                _box(anchor=DrawingAnchorInput(kind=9, from_row=1, from_col=1, to_row=5, to_col=4)),
                _box(fill=ChartFill(other=True)),
                _box(anchor=DrawingAnchorInput(kind=AnchorKind.TWO_CELL, from_row=1, from_col=1, to_row=0, to_col=4)),
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


def _group(**overrides) -> SparklineGroupInput:
    group = SparklineGroupInput(
        type=SparklineType.COLUMN,
        color_series=Color(kind=1, argb=0xFF336699),
        color_negative=Color(kind=2, theme=5, tint=-0.25),
        color_low=Color(kind=3, indexed=12),
        markers=True,
        negative=True,
        display_x_axis=True,
        right_to_left=True,
        min_axis_type=SparklineAxisType.CUSTOM,
        manual_min=-1.5,
        line_width_pt=1.25,
        empty_cells_as=BlanksAs.ZERO,
        sparklines=[SparklineInput("Sheet1!A1:D1", "F1"), SparklineInput("Sheet1!A2:D2", "F2")],
    )
    for key, value in overrides.items():
        setattr(group, key, value)
    return group


class SparklineGroupTests(unittest.TestCase):
    def test_add_list_update_replace_remove_and_row_edit_shift(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.sheet(0).list_sparkline_groups(), [])
            wb.sheet(0).add_sparkline_group(_group())
            wb.sheet(0).add_sparkline_group(SparklineGroupInput(sparklines=[SparklineInput("Sheet1!A9:D9", "G9")]))
            with Workbook.load(wb.save()) as loaded:
                groups = loaded.sheet(0).list_sparkline_groups()
                self.assertEqual(len(groups), 2)
                first, second = groups
                self.assertEqual(first.type, SparklineType.COLUMN)
                self.assertEqual((first.color_series.kind, first.color_series.argb), (1, 0xFF336699))
                self.assertEqual((first.color_negative.theme, first.color_negative.tint), (5, -0.25))
                self.assertEqual(first.color_low.indexed, 12)
                self.assertIsNone(first.color_high)
                self.assertEqual((first.markers, first.high, first.negative), (True, False, True))
                self.assertEqual((first.right_to_left, second.right_to_left), (True, False))
                self.assertEqual(first.min_axis_type, SparklineAxisType.CUSTOM)
                self.assertEqual((first.manual_min, first.manual_max), (-1.5, None))
                self.assertEqual(first.line_width_pt, 1.25)
                self.assertEqual(first.empty_cells_as, BlanksAs.ZERO)
                self.assertEqual(
                    first.sparklines,
                    [
                        Sparkline(SheetRect(0, CellRect(0, 0, 0, 3)), CellAddress(0, 5)),
                        Sparkline(SheetRect(0, CellRect(1, 0, 1, 3)), CellAddress(1, 5)),
                    ],
                )
                self.assertEqual((second.type, second.line_width_pt), (SparklineType.LINE, 0.75))

                loaded.sheet(0).insert_rows(0, 1)
                self.assertEqual(
                    loaded.sheet(0).list_sparkline_groups()[0].sparklines[1],
                    Sparkline(SheetRect(0, CellRect(3, 0, 3, 3)), CellAddress(3, 5)),
                )

                loaded.sheet(0).update_sparkline_group(
                    0, type=SparklineType.STACKED, high=True, date_axis=True, date_axis_range="Sheet1!A4:D4"
                )
                updated = loaded.sheet(0).list_sparkline_groups()[0]
                # A partial update leaves every omitted field as it was.
                self.assertEqual((updated.markers, updated.negative, updated.manual_min), (True, True, -1.5))
                self.assertEqual((updated.type, updated.high, updated.date_axis), (SparklineType.STACKED, True, True))
                self.assertEqual(updated.date_axis_range, SheetRect(0, CellRect(3, 0, 3, 3)))
                self.assertTrue(updated.right_to_left)
                loaded.sheet(0).update_sparkline_group(0, right_to_left=False)
                self.assertFalse(loaded.sheet(0).list_sparkline_groups()[0].right_to_left)
                loaded.sheet(0).update_sparkline_group(0, right_to_left=True)
                self.assertEqual(len(updated.sparklines), 2)

                loaded.sheet(0).update_sparkline_group(0, sparklines=[SparklineInput("Sheet1!B1:B5", "H1")])
                self.assertEqual(
                    loaded.sheet(0).list_sparkline_groups()[0].sparklines,
                    [Sparkline(SheetRect(0, CellRect(0, 1, 4, 1)), CellAddress(0, 7))],
                )
                with self.assertRaises(InvalidArgumentError):
                    loaded.sheet(0).update_sparkline_group(0, sparklines=[])
                self.assertEqual(
                    loaded.sheet(0).list_sparkline_groups()[0].sparklines,
                    [Sparkline(SheetRect(0, CellRect(0, 1, 4, 1)), CellAddress(0, 7))],
                )
                loaded.sheet(0).remove_sparkline_group(0)
                remaining = loaded.sheet(0).list_sparkline_groups()
                self.assertEqual(len(remaining), 1)
                self.assertEqual(remaining[0].sparklines[0].location, CellAddress(10, 6))

    def test_arguments_are_validated(self) -> None:
        good = [SparklineInput("Sheet1!A1:D1", "F1")]
        with Workbook.create_default() as wb:
            for bad in (
                SparklineGroupInput(),
                SparklineGroupInput(sparklines=[SparklineInput("Sheet1!A1:D1", "F1:F2")]),
                SparklineGroupInput(type=7, sparklines=good),
                SparklineGroupInput(line_width_pt=-1.0, sparklines=good),
            ):
                with self.assertRaises(FormulonError):
                    wb.sheet(0).add_sparkline_group(bad)
            with self.assertRaises(FormulonError):
                wb.sheet(4).add_sparkline_group(SparklineGroupInput(sparklines=good))
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
            wb.sheet(0).set_comment(1, 1, Comment("one", author="Alice"))
            wb.sheet(0).set_comment(3, 3, Comment("two", author="Alice"))
            wb.sheet(0).set_comment_visible(1, 1, True)
            with Workbook.load(wb.save()) as first:
                first.sheet(0).set_comment(8, 8, Comment("three", author="Alice"))
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
