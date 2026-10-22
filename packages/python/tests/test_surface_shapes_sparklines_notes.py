"""Drawing shape, sparkline group and legacy note shape surface tests."""

from __future__ import annotations

import io
import re
import unittest
import zipfile

from formulon import (
    AnchorKind,
    ColorSpec,
    DrawingObjectKind,
    FormulonError,
    Shape,
    ShapePaint,
    Sparkline,
    SparklineAxisType,
    SparklineEmptyCells,
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
        prst="roundRect",
        text="Hello\nWorld",
        fill_kind=ShapePaint.SOLID,
        fill_rgb=0xFF0000,
        line_kind=ShapePaint.NONE,
        rotation=900000,
    )
    fields.update(overrides)
    return fields


class ShapeTests(unittest.TestCase):
    def test_insert_list_update_remove_and_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.list_shapes(0), [])
            rect = wb.insert_shape(0, **_box())
            line = wb.insert_shape(
                0,
                anchor_kind=AnchorKind.ONE_CELL,
                from_row=10,
                from_col_off=500,
                cx=914400,
                cy=0,
                prst="line",
                line_kind=ShapePaint.SOLID,
                line_rgb=0x00FF00,
                line_width_emu=19050,
            )
            self.assertNotEqual(rect, line)
            with Workbook.load(wb.save()) as loaded:
                shapes = loaded.list_shapes(0)
                self.assertEqual(len(shapes), 2)
                self.assertIsInstance(shapes[0], Shape)
                self.assertEqual(shapes[0].object_id, rect)
                self.assertEqual(shapes[0].kind, DrawingObjectKind.SHAPE)
                self.assertEqual((shapes[0].name, shapes[0].prst, shapes[0].text), ("Box", "roundRect", "Hello\nWorld"))
                self.assertEqual((shapes[0].fill_kind, shapes[0].fill_rgb), (ShapePaint.SOLID, 0xFF0000))
                self.assertEqual(shapes[0].line_kind, ShapePaint.NONE)
                self.assertEqual((shapes[0].rotation, shapes[0].hidden, shapes[0].text_box), (900000, False, False))
                self.assertEqual((shapes[0].from_row, shapes[0].to_row, shapes[0].to_col), (1, 5, 4))
                self.assertEqual(shapes[1].kind, DrawingObjectKind.CONNECTOR)
                self.assertEqual((shapes[1].from_col_off, shapes[1].cx, shapes[1].line_rgb), (500, 914400, 0x00FF00))
                self.assertEqual(shapes[1].line_width_emu, 19050)

                loaded.update_shape(0, rect, name="Renamed", hidden=True, fill_kind=ShapePaint.NONE)
                loaded.insert_rows(0, 0, 3)
                moved = loaded.list_shapes(0)[0]
                self.assertEqual((moved.name, moved.hidden, moved.fill_kind), ("Renamed", True, ShapePaint.NONE))
                self.assertEqual((moved.prst, moved.rotation), ("roundRect", 900000))
                self.assertEqual((moved.from_row, moved.to_row), (4, 8))

                loaded.remove_shape(0, rect)
                self.assertEqual([s.object_id for s in loaded.list_shapes(0)], [line])

    def test_arguments_are_validated(self) -> None:
        with Workbook.create_default() as wb:
            for bad in (
                _box(anchor_kind=9),
                _box(fill_kind=ShapePaint.OTHER),
                _box(to_row=0),
                _box(prst='a"/>'),
            ):
                with self.assertRaises(FormulonError):
                    wb.insert_shape(0, **bad)
            with self.assertRaises(FormulonError):
                wb.insert_shape(5, **_box())
            self.assertEqual(wb.list_shapes(0), [])
            with self.assertRaises(FormulonError):
                wb.update_shape(0, 99, name="x")
            with self.assertRaises(FormulonError):
                wb.remove_shape(0, 99)


def _group(**overrides) -> SparklineGroup:
    group = SparklineGroup(
        type=SparklineType.COLUMN,
        color_series=ColorSpec(kind=1, rgb=0xFF336699),
        color_negative=ColorSpec(kind=2, theme=5, tint=-0.25),
        color_low=ColorSpec(kind=3, indexed=12),
        markers=True,
        negative=True,
        display_x_axis=True,
        min_axis_type=SparklineAxisType.CUSTOM,
        manual_min=-1.5,
        line_weight=1.25,
        empty_cells_as=SparklineEmptyCells.ZERO,
        sparklines=[Sparkline("Sheet1!A1:D1", "F1"), Sparkline("Sheet1!A2:D2", "F2")],
    )
    for key, value in overrides.items():
        setattr(group, key, value)
    return group


class SparklineGroupTests(unittest.TestCase):
    def test_add_list_update_replace_remove_and_row_edit_shift(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.list_sparkline_groups(0), [])
            wb.add_sparkline_group(0, _group())
            wb.add_sparkline_group(0, SparklineGroup(sparklines=[Sparkline("Sheet1!A9:D9", "G9")]))
            with Workbook.load(wb.save()) as loaded:
                groups = loaded.list_sparkline_groups(0)
                self.assertEqual(len(groups), 2)
                first, second = groups
                self.assertEqual(first.type, SparklineType.COLUMN)
                self.assertEqual((first.color_series.kind, first.color_series.rgb), (1, 0xFF336699))
                self.assertEqual((first.color_negative.theme, first.color_negative.tint), (5, -0.25))
                self.assertEqual(first.color_low.indexed, 12)
                self.assertEqual(first.color_high.kind, 0)
                self.assertEqual((first.markers, first.high, first.negative), (True, False, True))
                self.assertEqual(first.min_axis_type, SparklineAxisType.CUSTOM)
                self.assertEqual((first.manual_min, first.manual_max), (-1.5, None))
                self.assertEqual(first.line_weight, 1.25)
                self.assertEqual(first.empty_cells_as, SparklineEmptyCells.ZERO)
                self.assertEqual(first.sparklines, [Sparkline("Sheet1!A1:D1", "F1"), Sparkline("Sheet1!A2:D2", "F2")])
                self.assertEqual((second.type, second.line_weight), (SparklineType.LINE, 0.75))

                loaded.insert_rows(0, 0, 2)
                self.assertEqual(loaded.list_sparkline_groups(0)[0].sparklines[1], Sparkline("Sheet1!A4:D4", "F4"))

                first = loaded.list_sparkline_groups(0)[0]
                first.type = SparklineType.STACKED
                first.high = True
                first.date_axis = True
                first.date_axis_ref = "Sheet1!A4:D4"
                loaded.update_sparkline_group(0, 0, first)
                updated = loaded.list_sparkline_groups(0)[0]
                self.assertEqual((updated.type, updated.high, updated.date_axis), (SparklineType.STACKED, True, True))
                self.assertEqual(updated.date_axis_ref, "Sheet1!A4:D4")
                self.assertEqual(len(updated.sparklines), 2)

                loaded.set_sparkline_group_sparklines(0, 0, [Sparkline("Sheet1!B1:B5", "H1")])
                self.assertEqual(loaded.list_sparkline_groups(0)[0].sparklines, [Sparkline("Sheet1!B1:B5", "H1")])
                loaded.remove_sparkline_group(0, 0)
                remaining = loaded.list_sparkline_groups(0)
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
                    wb.add_sparkline_group(0, bad)
            with self.assertRaises(FormulonError):
                wb.add_sparkline_group(4, SparklineGroup(sparklines=good))
            self.assertEqual(wb.list_sparkline_groups(0), [])
            with self.assertRaises(FormulonError):
                wb.update_sparkline_group(0, 0, SparklineGroup())
            with self.assertRaises(FormulonError):
                wb.set_sparkline_group_sparklines(0, 0, good)
            with self.assertRaises(FormulonError):
                wb.remove_sparkline_group(0, 0)


class NoteShapeTests(unittest.TestCase):
    def test_surviving_note_keeps_its_own_shape_when_notes_are_added(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_comment(0, 1, 1, "Alice", "one")
            wb.set_comment(0, 3, 3, "Alice", "two")
            wb.set_comment_visible(0, 1, 1, True)
            with Workbook.load(wb.save()) as first:
                first.set_comment(0, 8, 8, "Alice", "three")
                first.set_comment_visible(0, 3, 3, True)
                saved = first.save()
                vml = zipfile.ZipFile(io.BytesIO(saved)).read("xl/drawings/vmlDrawing1.vml").decode("utf-8")
                self.assertEqual(len(re.findall(r"<v:shape ", vml)), 3)
                self.assertEqual(len(set(re.findall(r'id="_x0000_s\d+"', vml))), 3)
                self.assertEqual(vml.count("visibility:visible"), 2)
                with Workbook.load(saved) as again:
                    self.assertEqual(
                        [again.get_comment_visible(0, r, c) for r, c in ((1, 1), (3, 3), (8, 8))],
                        [True, True, False],
                    )


if __name__ == "__main__":
    unittest.main()
