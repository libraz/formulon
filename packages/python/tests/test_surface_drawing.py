"""Drawing-image surface tests: probe, insert, list, get, round trip, remove, move, reorder and restore."""

from __future__ import annotations

import dataclasses
import struct
import unittest
import zlib

from formulon import (
    AnchorEditAs,
    AnchorKind,
    ChartInsertInput,
    ChartPlotType,
    DrawingAnchor,
    DrawingAnchorInput,
    DrawingObjectKind,
    DrawingPlacementInput,
    FormulonError,
    ImageFormat,
    ImageInfo,
    IndexOutOfRangeError,
    NotFoundError,
    Picture,
    ShapeInput,
    Workbook,
    get_image_info,
)


def _png(width: int, height: int) -> bytes:
    """Build a minimal valid RGB PNG of the given size."""

    def chunk(tag: bytes, body: bytes) -> bytes:
        return struct.pack(">I", len(body)) + tag + body + struct.pack(">I", zlib.crc32(tag + body))

    rows = b"".join(b"\x00" + b"\xff\x00\x00" * width for _ in range(height))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(rows))
        + chunk(b"IEND", b"")
    )


class DrawingEnumTests(unittest.TestCase):
    def test_ordinals_match_the_c_header(self) -> None:
        self.assertEqual((ImageFormat.PNG, ImageFormat.JPEG, ImageFormat.GIF, ImageFormat.BMP), (1, 2, 3, 4))
        self.assertEqual((DrawingObjectKind.PICTURE, DrawingObjectKind.OTHER), (0, 6))
        self.assertEqual(
            (AnchorKind.DEFAULT, AnchorKind.ONE_CELL, AnchorKind.TWO_CELL, AnchorKind.ABSOLUTE), (0, 1, 2, 3)
        )
        self.assertEqual(
            (AnchorEditAs.DEFAULT, AnchorEditAs.ONE_CELL, AnchorEditAs.TWO_CELL, AnchorEditAs.ABSOLUTE),
            (0, 1, 2, 3),
        )


class DrawingImageTests(unittest.TestCase):
    def test_probe_image(self) -> None:
        self.assertEqual(get_image_info(_png(3, 2)), ImageInfo(ImageFormat.PNG, 3, 2))
        with self.assertRaises(FormulonError):
            get_image_info(b"not an image")

    def test_insert_list_get_save_load_remove(self) -> None:
        png = _png(4, 3)
        wb = Workbook.create_default()
        try:
            self.assertEqual(wb.sheet(0).list_drawings(), [])
            pic = wb.sheet(0).add_picture(
                png,
                DrawingPlacementInput(
                    name="logo",
                    description="a logo",
                    anchor=DrawingAnchorInput(from_row=2, from_col=1, from_row_offset_emu=100),
                ),
            )
            self.assertIsInstance(pic, Picture)
            objs = wb.sheet(0).list_drawings()
            self.assertEqual(len(objs), 1)
            self.assertIs(objs[0], pic)
            obj = objs[0].get_drawing_properties()
            self.assertEqual(obj.object_id, pic.object_id)
            self.assertEqual(obj.kind, DrawingObjectKind.PICTURE)
            self.assertEqual(obj.anchor.kind, AnchorKind.ONE_CELL)
            self.assertEqual((obj.anchor.from_row, obj.anchor.from_col, obj.anchor.from_row_offset_emu), (2, 1, 100))
            self.assertEqual((obj.anchor.width_emu, obj.anchor.height_emu), (4 * 9525, 3 * 9525))
            self.assertEqual((obj.name, obj.description, obj.hidden), ("logo", "a logo", False))
            props = wb.sheet(0).picture(pic.object_id).get_properties()
            self.assertEqual((props.format, props.width_px, props.height_px, props.bytes), (ImageFormat.PNG, 4, 3, png))

            loaded = Workbook.load(wb.save())
            try:
                again = loaded.sheet(0).list_drawings()
                self.assertEqual(len(again), 1)
                again_obj = again[0].get_drawing_properties()
                self.assertEqual(
                    (again_obj.name, again_obj.anchor.width_emu, again_obj.anchor.height_emu),
                    ("logo", 4 * 9525, 3 * 9525),
                )
                self.assertEqual(loaded.sheet(0).picture(again_obj.object_id).get_properties().bytes, png)
            finally:
                loaded.close()

            wb.sheet(0).remove_drawing(pic)
            self.assertEqual(wb.sheet(0).list_drawings(), [])
            with self.assertRaises(FormulonError):
                wb.sheet(0).remove_drawing(pic)
        finally:
            wb.close()

    def test_two_cell_anchor_and_explicit_size(self) -> None:
        wb = Workbook.create_default()
        try:
            wb.sheet(0).add_picture(
                _png(2, 2),
                DrawingPlacementInput(
                    anchor=DrawingAnchorInput(
                        kind=AnchorKind.TWO_CELL,
                        edit_as=AnchorEditAs.ONE_CELL,
                        width_emu=1_000_000,
                        height_emu=2_000_000,
                    )
                ),
            )
            obj = wb.sheet(0).list_drawings()[0].get_drawing_properties()
            self.assertEqual((obj.anchor.kind, obj.anchor.edit_as), (AnchorKind.TWO_CELL, AnchorEditAs.ONE_CELL))
            self.assertEqual((obj.anchor.width_emu, obj.anchor.height_emu), (1_000_000, 2_000_000))
        finally:
            wb.close()

    def test_typed_lists_are_filtered_from_the_one_drawing_list(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_text(0, 0, "x")
            ws.set_number(1, 0, 1.0)
            pic = ws.add_picture(_png(2, 2))
            shape = ws.add_shape(ShapeInput(anchor=DrawingAnchorInput(from_row=3, to_row=6, to_col=3)))
            chart = ws.add_chart(ChartInsertInput(type=ChartPlotType.BAR, range="A1:A2"))
            self.assertEqual(
                [d.object_id for d in ws.list_drawings()], [pic.object_id, shape.object_id, chart.object_id]
            )
            self.assertEqual([d.object_id for d in ws.list_pictures()], [pic.object_id])
            self.assertEqual([d.object_id for d in ws.list_shapes()], [shape.object_id])
            self.assertEqual([d.object_id for d in ws.list_charts()], [chart.object_id])
            self.assertEqual(ws.list_slicers(), [])
            for handle in ws.list_drawings():
                props = handle.get_drawing_properties()
                self.assertEqual(props.object_id, handle.object_id)
                self.assertEqual(props.name, handle.name)
                self.assertIsInstance(props.anchor.kind, AnchorKind)
                self.assertNotEqual(props.anchor.kind, AnchorKind.DEFAULT)
                self.assertNotEqual(props.anchor.edit_as, AnchorEditAs.DEFAULT)
            for handle in (pic, shape, chart):
                ws.remove_drawing(handle)
            self.assertEqual(ws.list_drawings(), [])

    def test_a_picture_without_input_is_one_cell_and_an_anchor_kind_default_is_the_operation_default(self) -> None:
        with Workbook.create_default() as wb:
            plain = wb.sheet(0).add_picture(_png(2, 2))
            self.assertEqual(plain.get_drawing_properties().anchor.kind, AnchorKind.ONE_CELL)
            explicit = wb.sheet(0).add_picture(
                _png(2, 2), DrawingPlacementInput(anchor=DrawingAnchorInput(kind=AnchorKind.DEFAULT, from_row=4))
            )
            self.assertEqual(explicit.get_drawing_properties().anchor.kind, AnchorKind.ONE_CELL)
            self.assertEqual(explicit.get_drawing_properties().anchor.from_row, 4)
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_picture(
                    _png(2, 2), DrawingPlacementInput(anchor=DrawingAnchorInput(kind=AnchorKind.ABSOLUTE))
                )
            with self.assertRaises(TypeError):
                DrawingPlacementInput(anchor=DrawingAnchorInput(row=1))  # type: ignore[call-arg]
            with self.assertRaises(TypeError):
                DrawingPlacementInput(anchor_kind=AnchorKind.ONE_CELL)  # type: ignore[call-arg]

    def test_the_anchor_read_is_a_frozen_record_with_every_field(self) -> None:
        with Workbook.create_default() as wb:
            anchor = wb.sheet(0).add_picture(_png(2, 2)).get_drawing_properties().anchor
            self.assertIsInstance(anchor, DrawingAnchor)
            self.assertEqual(
                [f.name for f in dataclasses.fields(DrawingAnchor)],
                [
                    "kind",
                    "edit_as",
                    "from_row",
                    "from_col",
                    "from_row_offset_emu",
                    "from_col_offset_emu",
                    "to_row",
                    "to_col",
                    "to_row_offset_emu",
                    "to_col_offset_emu",
                    "width_emu",
                    "height_emu",
                ],
            )
            with self.assertRaises(dataclasses.FrozenInstanceError):
                anchor.from_row = 3  # type: ignore[misc]

    def test_invalid_inputs_raise(self) -> None:
        wb = Workbook.create_default()
        try:
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_picture(b"junk")
            with self.assertRaises(NotFoundError):
                wb.sheet(0).picture(12345)
            with self.assertRaises(IndexOutOfRangeError):
                wb.sheet(9).picture(1).get_properties()
        finally:
            wb.close()

    def test_update_anchor_merges_into_the_current_anchor(self) -> None:
        with Workbook.create_default() as wb:
            pic = wb.sheet(0).add_picture(
                _png(4, 3), DrawingPlacementInput(name="logo", anchor=DrawingAnchorInput(from_row=1, from_col=1))
            )
            pic.update(anchor=DrawingAnchorInput(from_row=5, from_col=3))
            anchor = pic.get_drawing_properties().anchor
            self.assertEqual((anchor.kind, anchor.from_row, anchor.from_col), (AnchorKind.ONE_CELL, 5, 3))
            self.assertEqual((anchor.width_emu, anchor.height_emu), (4 * 9525, 3 * 9525))

            pic.update(anchor=DrawingAnchorInput(from_row_offset_emu=1000, width_emu=200_000))
            anchor = pic.get_drawing_properties().anchor
            self.assertEqual((anchor.from_row, anchor.from_col, anchor.from_row_offset_emu), (5, 3, 1000))
            self.assertEqual((anchor.width_emu, anchor.height_emu), (200_000, 3 * 9525))
            self.assertEqual(pic.name, "logo")

    def test_update_anchor_kind_converts_and_omitted_kind_keeps(self) -> None:
        with Workbook.create_default() as wb:
            pic = wb.sheet(0).add_picture(
                _png(4, 3), DrawingPlacementInput(anchor=DrawingAnchorInput(from_row=1, from_col=1))
            )
            pic.update(anchor=DrawingAnchorInput(kind=AnchorKind.TWO_CELL, edit_as=AnchorEditAs.ONE_CELL, from_row=5))
            anchor = pic.get_drawing_properties().anchor
            self.assertEqual((anchor.kind, anchor.edit_as), (AnchorKind.TWO_CELL, AnchorEditAs.ONE_CELL))
            self.assertEqual((anchor.from_row, anchor.from_col), (5, 1))
            self.assertGreater(anchor.to_row + anchor.to_col, 0)

            pic.update(anchor=DrawingAnchorInput(from_row=7))
            anchor = pic.get_drawing_properties().anchor
            self.assertEqual((anchor.kind, anchor.from_row, anchor.from_col), (AnchorKind.TWO_CELL, 7, 1))

    def test_update_anchor_rejects_a_missing_object_and_a_mismatched_field(self) -> None:
        with Workbook.create_default() as wb:
            pic = wb.sheet(0).add_picture(_png(4, 3))
            with self.assertRaises(FormulonError):
                pic.update(anchor=DrawingAnchorInput(to_col=4))
            with self.assertRaises(TypeError):
                pic.update(anchor=None)  # type: ignore[arg-type]
            wb.sheet(0).remove_drawing(pic)
            with self.assertRaises(NotFoundError):
                pic.update(name="x")

    def test_update_name_description_and_hidden_on_a_picture(self) -> None:
        with Workbook.create_default() as wb:
            pic = wb.sheet(0).add_picture(_png(2, 2), DrawingPlacementInput(name="logo"))
            self.assertEqual(pic.name, "logo")
            pic.update(name="renamed", description="alt text", hidden=True)
            props = wb.sheet(0).list_drawings()[0].get_drawing_properties()
            self.assertEqual((props.name, props.description, props.hidden), ("renamed", "alt text", True))
            self.assertEqual(pic.name, "renamed")
            pic.update(hidden=False)
            props = pic.get_drawing_properties()
            self.assertEqual((props.name, props.description, props.hidden), ("renamed", "alt text", False))
            with self.assertRaises(FormulonError):
                pic.update(name="")

    def test_a_picture_has_no_rename_or_set_anchor(self) -> None:
        with Workbook.create_default() as wb:
            pic = wb.sheet(0).add_picture(_png(2, 2))
            self.assertFalse(hasattr(pic, "rename"))
            self.assertFalse(hasattr(pic, "set_anchor"))

    def test_drawing_properties_carry_the_picture_format_on_the_picture_only(self) -> None:
        with Workbook.create_default() as wb:
            pic = wb.sheet(0).add_picture(_png(2, 2))
            self.assertFalse(hasattr(pic.get_drawing_properties(), "format"))
            self.assertEqual(pic.get_properties().format, ImageFormat.PNG)

    def test_set_drawing_z_order_reorders_the_list(self) -> None:
        with Workbook.create_default() as wb:
            a = wb.sheet(0).add_picture(_png(2, 2), DrawingPlacementInput(anchor=DrawingAnchorInput(from_row=0)))
            b = wb.sheet(0).add_picture(_png(2, 2), DrawingPlacementInput(anchor=DrawingAnchorInput(from_row=3)))
            c = wb.sheet(0).add_picture(_png(2, 2), DrawingPlacementInput(anchor=DrawingAnchorInput(from_row=6)))
            self.assertEqual(
                [o.object_id for o in wb.sheet(0).list_drawings()], [a.object_id, b.object_id, c.object_id]
            )
            wb.sheet(0).drawing(c.object_id).set_z_order(0)
            self.assertEqual(
                [o.object_id for o in wb.sheet(0).list_drawings()], [c.object_id, a.object_id, b.object_id]
            )
            with self.assertRaises(FormulonError):
                c.set_z_order(3)
            with self.assertRaises(FormulonError):
                wb.sheet(0).drawing(9999).set_z_order(0)

    def test_snapshot_and_restore_drawing(self) -> None:
        png = _png(4, 3)
        with Workbook.create_default() as wb:
            a = wb.sheet(0).add_picture(png, DrawingPlacementInput(name="A", anchor=DrawingAnchorInput(from_row=1)))
            b = wb.sheet(0).add_picture(
                _png(2, 2), DrawingPlacementInput(name="B", anchor=DrawingAnchorInput(from_row=4))
            )
            snap = wb.sheet(0).drawing(a.object_id).snapshot()
            self.assertIsInstance(snap, bytes)
            self.assertGreater(len(snap), 0)

            wb.sheet(0).remove_drawing(a)
            restored = wb.sheet(0).restore_drawing(snap)
            self.assertIs(restored, a)
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawings()], [a.object_id, b.object_id])
            self.assertEqual(wb.sheet(0).picture(a.object_id).get_properties().bytes, png)

            a.update(anchor=DrawingAnchorInput(from_row=9))
            wb.sheet(0).drawing(a.object_id).set_z_order(1)
            self.assertIs(wb.sheet(0).restore_drawing(snap), a)
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawings()], [a.object_id, b.object_id])
            self.assertEqual(wb.sheet(0).list_drawings()[0].get_drawing_properties().anchor.from_row, 1)

            copy = wb.sheet(0).restore_drawing(snap, assign_new_id=True)
            self.assertNotIn(copy.object_id, (a.object_id, b.object_id))
            self.assertEqual(
                [o.object_id for o in wb.sheet(0).list_drawings()], [a.object_id, b.object_id, copy.object_id]
            )

            with self.assertRaises(FormulonError):
                wb.sheet(0).restore_drawing(b"junk")
            with self.assertRaises(FormulonError):
                wb.sheet(0).drawing(9999).snapshot()


if __name__ == "__main__":
    unittest.main()
