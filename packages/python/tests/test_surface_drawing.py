"""Drawing-image surface tests: probe, insert, list, get, round trip, remove, move, reorder and restore."""

from __future__ import annotations

import struct
import unittest
import zlib

from formulon import (
    AnchorEditAs,
    AnchorKind,
    DrawingAnchor,
    DrawingObjectKind,
    FormulonError,
    ImageFormat,
    ImageInfo,
    IndexOutOfRangeError,
    NotFoundError,
    Picture,
    PictureInput,
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
        self.assertEqual((AnchorKind.ONE_CELL, AnchorKind.TWO_CELL, AnchorKind.ABSOLUTE), (0, 1, 2))
        self.assertEqual((AnchorEditAs.ONE_CELL, AnchorEditAs.TWO_CELL, AnchorEditAs.ABSOLUTE), (0, 1, 2))


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
                PictureInput(name="logo", description="a logo", anchor=DrawingAnchor(row=2, col=1, row_offset_emu=100)),
            )
            self.assertIsInstance(pic, Picture)
            objs = wb.sheet(0).list_drawings()
            self.assertEqual(len(objs), 1)
            self.assertIs(objs[0], pic)
            obj = objs[0].get_drawing_properties()
            self.assertEqual(obj.object_id, pic.object_id)
            self.assertEqual(obj.kind, DrawingObjectKind.PICTURE)
            self.assertEqual(obj.anchor_kind, AnchorKind.ONE_CELL)
            self.assertEqual((obj.from_row, obj.from_col, obj.from_row_offset_emu), (2, 1, 100))
            self.assertEqual((obj.width_emu, obj.height_emu), (4 * 9525, 3 * 9525))
            self.assertEqual((obj.name, obj.description), ("logo", "a logo"))
            self.assertEqual(obj.format, ImageFormat.PNG)
            props = wb.sheet(0).picture(pic.object_id).get_properties()
            self.assertEqual((props.format, props.width_px, props.height_px, props.bytes), (ImageFormat.PNG, 4, 3, png))

            loaded = Workbook.load(wb.save())
            try:
                again = loaded.sheet(0).list_drawings()
                self.assertEqual(len(again), 1)
                again_obj = again[0].get_drawing_properties()
                self.assertEqual(
                    (again_obj.name, again_obj.width_emu, again_obj.height_emu), ("logo", 4 * 9525, 3 * 9525)
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
                PictureInput(
                    anchor=DrawingAnchor(
                        anchor_kind=AnchorKind.TWO_CELL,
                        edit_as=AnchorEditAs.ONE_CELL,
                        width_emu=1_000_000,
                        height_emu=2_000_000,
                    )
                ),
            )
            obj = wb.sheet(0).list_drawings()[0].get_drawing_properties()
            self.assertEqual((obj.anchor_kind, obj.edit_as), (AnchorKind.TWO_CELL, AnchorEditAs.ONE_CELL))
            self.assertEqual((obj.width_emu, obj.height_emu), (1_000_000, 2_000_000))
        finally:
            wb.close()

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

    def test_set_drawing_anchor_moves_and_resizes(self) -> None:
        with Workbook.create_default() as wb:
            pic = wb.sheet(0).add_picture(_png(4, 3), PictureInput(name="logo", anchor=DrawingAnchor(row=1, col=1)))
            wb.sheet(0).drawing(pic.object_id).set_anchor(
                DrawingAnchor(
                    anchor_kind=AnchorKind.TWO_CELL,
                    edit_as=AnchorEditAs.ONE_CELL,
                    row=5,
                    col=3,
                    row_offset_emu=1000,
                    col_offset_emu=2000,
                    width_emu=200_000,
                )
            )
            obj = wb.sheet(0).list_drawings()[0].get_drawing_properties()
            self.assertEqual((obj.object_id, obj.name), (pic.object_id, "logo"))
            self.assertEqual((obj.anchor_kind, obj.edit_as), (AnchorKind.TWO_CELL, AnchorEditAs.ONE_CELL))
            self.assertEqual(
                (obj.from_row, obj.from_col, obj.from_row_offset_emu, obj.from_col_offset_emu), (5, 3, 1000, 2000)
            )
            self.assertEqual((obj.width_emu, obj.height_emu), (200_000, 3 * 9525))
            with self.assertRaises(FormulonError):
                pic.set_anchor(DrawingAnchor(anchor_kind=AnchorKind.ABSOLUTE))
            with self.assertRaises(FormulonError):
                wb.sheet(0).drawing(9999).set_anchor(DrawingAnchor())

    def test_set_drawing_name_renames_the_object(self) -> None:
        wb = Workbook.create_default()
        try:
            pic = wb.sheet(0).add_picture(_png(2, 2), PictureInput(name="logo"))
            wb.sheet(0).drawing(pic.object_id).rename("renamed")
            self.assertEqual(wb.sheet(0).list_drawings()[0].get_drawing_properties().name, "renamed")
            with self.assertRaises(FormulonError):
                pic.rename("")
            with self.assertRaises(FormulonError):
                wb.sheet(0).drawing(9999).rename("x")
        finally:
            wb.close()

    def test_set_drawing_z_order_reorders_the_list(self) -> None:
        with Workbook.create_default() as wb:
            a = wb.sheet(0).add_picture(_png(2, 2), PictureInput(anchor=DrawingAnchor(row=0)))
            b = wb.sheet(0).add_picture(_png(2, 2), PictureInput(anchor=DrawingAnchor(row=3)))
            c = wb.sheet(0).add_picture(_png(2, 2), PictureInput(anchor=DrawingAnchor(row=6)))
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
            a = wb.sheet(0).add_picture(png, PictureInput(name="A", anchor=DrawingAnchor(row=1)))
            b = wb.sheet(0).add_picture(_png(2, 2), PictureInput(name="B", anchor=DrawingAnchor(row=4)))
            snap = wb.sheet(0).drawing(a.object_id).snapshot()
            self.assertIsInstance(snap, bytes)
            self.assertGreater(len(snap), 0)

            wb.sheet(0).remove_drawing(a)
            restored = wb.sheet(0).restore_drawing(snap)
            self.assertIs(restored, a)
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawings()], [a.object_id, b.object_id])
            self.assertEqual(wb.sheet(0).picture(a.object_id).get_properties().bytes, png)

            wb.sheet(0).drawing(a.object_id).set_anchor(DrawingAnchor(row=9))
            wb.sheet(0).drawing(a.object_id).set_z_order(1)
            self.assertIs(wb.sheet(0).restore_drawing(snap), a)
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawings()], [a.object_id, b.object_id])
            self.assertEqual(wb.sheet(0).list_drawings()[0].get_drawing_properties().from_row, 1)

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
