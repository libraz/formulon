"""Drawing-image surface tests: probe, insert, list, get, round trip, remove, move, reorder and restore."""

from __future__ import annotations

import struct
import unittest
import zlib

from formulon import (
    AnchorEditAs,
    AnchorKind,
    DrawingObjectKind,
    FormulonError,
    ImageFormat,
    ImageInfo,
    IndexOutOfRangeError,
    Workbook,
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
        wb = Workbook.create_default()
        try:
            self.assertEqual(wb.get_image_info(_png(3, 2)), ImageInfo(ImageFormat.PNG, 3, 2))
            with self.assertRaises(FormulonError):
                wb.get_image_info(b"not an image")
        finally:
            wb.close()

    def test_insert_list_get_save_load_remove(self) -> None:
        png = _png(4, 3)
        wb = Workbook.create_default()
        try:
            self.assertEqual(wb.sheet(0).list_drawing_objects(), [])
            oid = wb.sheet(0).add_image(png, name="logo", description="a logo", row=2, col=1, row_offset_emu=100)
            objs = wb.sheet(0).list_drawing_objects()
            self.assertEqual(len(objs), 1)
            obj = objs[0]
            self.assertEqual(obj.object_id, oid)
            self.assertEqual(obj.kind, DrawingObjectKind.PICTURE)
            self.assertEqual(obj.anchor_kind, AnchorKind.ONE_CELL)
            self.assertEqual((obj.from_row, obj.from_col, obj.from_row_offset_emu), (2, 1, 100))
            self.assertEqual((obj.width_emu, obj.height_emu), (4 * 9525, 3 * 9525))
            self.assertEqual((obj.name, obj.description), ("logo", "a logo"))
            self.assertEqual(obj.image_format, ImageFormat.PNG)
            self.assertTrue(obj.media_path)
            self.assertEqual(wb.sheet(0).get_image(oid), png)

            loaded = Workbook.load(wb.save())
            try:
                again = loaded.sheet(0).list_drawing_objects()
                self.assertEqual(len(again), 1)
                self.assertEqual((again[0].name, again[0].width_emu, again[0].height_emu), ("logo", 4 * 9525, 3 * 9525))
                self.assertEqual(loaded.sheet(0).get_image(again[0].object_id), png)
            finally:
                loaded.close()

            wb.sheet(0).remove_drawing_object(oid)
            self.assertEqual(wb.sheet(0).list_drawing_objects(), [])
            with self.assertRaises(FormulonError):
                wb.sheet(0).remove_drawing_object(oid)
        finally:
            wb.close()

    def test_two_cell_anchor_and_explicit_size(self) -> None:
        wb = Workbook.create_default()
        try:
            wb.sheet(0).add_image(
                _png(2, 2),
                anchor_kind=AnchorKind.TWO_CELL,
                edit_as=AnchorEditAs.ONE_CELL,
                width_emu=1_000_000,
                height_emu=2_000_000,
            )
            obj = wb.sheet(0).list_drawing_objects()[0]
            self.assertEqual((obj.anchor_kind, obj.edit_as), (AnchorKind.TWO_CELL, AnchorEditAs.ONE_CELL))
            self.assertEqual((obj.width_emu, obj.height_emu), (1_000_000, 2_000_000))
        finally:
            wb.close()

    def test_invalid_inputs_raise(self) -> None:
        wb = Workbook.create_default()
        try:
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_image(b"junk")
            self.assertIsNone(wb.sheet(0).get_image(12345))
            with self.assertRaises(IndexOutOfRangeError):
                wb.sheet(9).get_image(1)
        finally:
            wb.close()

    def test_set_drawing_anchor_moves_and_resizes(self) -> None:
        with Workbook.create_default() as wb:
            oid = wb.sheet(0).add_image(_png(4, 3), name="logo", row=1, col=1)
            wb.sheet(0).set_drawing_anchor(
                oid,
                anchor_kind=AnchorKind.TWO_CELL,
                edit_as=AnchorEditAs.ONE_CELL,
                row=5,
                col=3,
                row_offset_emu=1000,
                col_offset_emu=2000,
                width_emu=200_000,
            )
            obj = wb.sheet(0).list_drawing_objects()[0]
            self.assertEqual((obj.object_id, obj.name), (oid, "logo"))
            self.assertEqual((obj.anchor_kind, obj.edit_as), (AnchorKind.TWO_CELL, AnchorEditAs.ONE_CELL))
            self.assertEqual(
                (obj.from_row, obj.from_col, obj.from_row_offset_emu, obj.from_col_offset_emu), (5, 3, 1000, 2000)
            )
            self.assertEqual((obj.width_emu, obj.height_emu), (200_000, 3 * 9525))
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_drawing_anchor(oid, anchor_kind=AnchorKind.ABSOLUTE)
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_drawing_anchor(9999)

    def test_set_drawing_name_renames_the_object(self) -> None:
        wb = Workbook.create_default()
        try:
            oid = wb.sheet(0).add_image(_png(2, 2), name="logo")
            wb.sheet(0).set_drawing_name(oid, "renamed")
            self.assertEqual(wb.sheet(0).list_drawing_objects()[0].name, "renamed")
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_drawing_name(oid, "")
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_drawing_name(9999, "x")
        finally:
            wb.close()

    def test_set_drawing_z_order_reorders_the_list(self) -> None:
        with Workbook.create_default() as wb:
            a = wb.sheet(0).add_image(_png(2, 2), row=0)
            b = wb.sheet(0).add_image(_png(2, 2), row=3)
            c = wb.sheet(0).add_image(_png(2, 2), row=6)
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawing_objects()], [a, b, c])
            wb.sheet(0).set_drawing_z_order(c, 0)
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawing_objects()], [c, a, b])
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_drawing_z_order(c, 3)
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_drawing_z_order(9999, 0)

    def test_snapshot_and_restore_drawing(self) -> None:
        png = _png(4, 3)
        with Workbook.create_default() as wb:
            a = wb.sheet(0).add_image(png, name="A", row=1)
            b = wb.sheet(0).add_image(_png(2, 2), name="B", row=4)
            snap = wb.sheet(0).snapshot_drawing(a)
            self.assertIsInstance(snap, bytes)
            self.assertGreater(len(snap), 0)

            wb.sheet(0).remove_drawing_object(a)
            self.assertEqual(wb.sheet(0).restore_drawing(snap), a)
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawing_objects()], [a, b])
            self.assertEqual(wb.sheet(0).get_image(a), png)

            wb.sheet(0).set_drawing_anchor(a, row=9)
            wb.sheet(0).set_drawing_z_order(a, 1)
            self.assertEqual(wb.sheet(0).restore_drawing(snap), a)
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawing_objects()], [a, b])
            self.assertEqual(wb.sheet(0).list_drawing_objects()[0].from_row, 1)

            copy = wb.sheet(0).restore_drawing(snap, new_id=True)
            self.assertNotIn(copy, (a, b))
            self.assertEqual([o.object_id for o in wb.sheet(0).list_drawing_objects()], [a, b, copy])

            with self.assertRaises(FormulonError):
                wb.sheet(0).restore_drawing(b"junk")
            with self.assertRaises(FormulonError):
                wb.sheet(0).snapshot_drawing(9999)


if __name__ == "__main__":
    unittest.main()
