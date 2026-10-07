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
        self.assertEqual((AnchorEditAs.TWO_CELL, AnchorEditAs.ONE_CELL, AnchorEditAs.ABSOLUTE), (0, 1, 2))


class DrawingImageTests(unittest.TestCase):
    def test_probe_image(self) -> None:
        wb = Workbook.create_default()
        try:
            self.assertEqual(wb.probe_image(_png(3, 2)), ImageInfo(ImageFormat.PNG, 3, 2))
            with self.assertRaises(FormulonError):
                wb.probe_image(b"not an image")
        finally:
            wb.close()

    def test_insert_list_get_save_load_remove(self) -> None:
        png = _png(4, 3)
        wb = Workbook.create_default()
        try:
            self.assertEqual(wb.list_drawing_objects(0), [])
            oid = wb.insert_image(0, png, name="logo", descr="a logo", row=2, col=1, row_off_emu=100)
            objs = wb.list_drawing_objects(0)
            self.assertEqual(len(objs), 1)
            obj = objs[0]
            self.assertEqual(obj.object_id, oid)
            self.assertEqual(obj.kind, DrawingObjectKind.PICTURE)
            self.assertEqual(obj.anchor_kind, AnchorKind.ONE_CELL)
            self.assertEqual((obj.from_row, obj.from_col, obj.from_row_off), (2, 1, 100))
            self.assertEqual((obj.cx, obj.cy), (4 * 9525, 3 * 9525))
            self.assertEqual((obj.name, obj.descr), ("logo", "a logo"))
            self.assertEqual(obj.image_format, ImageFormat.PNG)
            self.assertTrue(obj.media_path)
            self.assertEqual(wb.get_image(0, oid), png)

            loaded = Workbook.load(wb.save())
            try:
                again = loaded.list_drawing_objects(0)
                self.assertEqual(len(again), 1)
                self.assertEqual((again[0].name, again[0].cx, again[0].cy), ("logo", 4 * 9525, 3 * 9525))
                self.assertEqual(loaded.get_image(0, again[0].object_id), png)
            finally:
                loaded.close()

            wb.remove_image(0, oid)
            self.assertEqual(wb.list_drawing_objects(0), [])
            with self.assertRaises(FormulonError):
                wb.remove_image(0, oid)
        finally:
            wb.close()

    def test_two_cell_anchor_and_explicit_size(self) -> None:
        wb = Workbook.create_default()
        try:
            wb.insert_image(
                0,
                _png(2, 2),
                anchor_kind=AnchorKind.TWO_CELL,
                edit_as=AnchorEditAs.ONE_CELL,
                width_emu=1_000_000,
                height_emu=2_000_000,
            )
            obj = wb.list_drawing_objects(0)[0]
            self.assertEqual((obj.anchor_kind, obj.edit_as), (AnchorKind.TWO_CELL, AnchorEditAs.ONE_CELL))
            self.assertEqual((obj.cx, obj.cy), (1_000_000, 2_000_000))
        finally:
            wb.close()

    def test_invalid_inputs_raise(self) -> None:
        wb = Workbook.create_default()
        try:
            with self.assertRaises(FormulonError):
                wb.insert_image(0, b"junk")
            with self.assertRaises(FormulonError):
                wb.get_image(0, 12345)
        finally:
            wb.close()

    def test_set_image_anchor_moves_and_resizes(self) -> None:
        with Workbook.create_default() as wb:
            oid = wb.insert_image(0, _png(4, 3), name="logo", row=1, col=1)
            wb.set_image_anchor(
                0,
                oid,
                anchor_kind=AnchorKind.TWO_CELL,
                edit_as=AnchorEditAs.ONE_CELL,
                row=5,
                col=3,
                row_off_emu=1000,
                col_off_emu=2000,
                width_emu=200_000,
            )
            obj = wb.list_drawing_objects(0)[0]
            self.assertEqual((obj.object_id, obj.name), (oid, "logo"))
            self.assertEqual((obj.anchor_kind, obj.edit_as), (AnchorKind.TWO_CELL, AnchorEditAs.ONE_CELL))
            self.assertEqual((obj.from_row, obj.from_col, obj.from_row_off, obj.from_col_off), (5, 3, 1000, 2000))
            self.assertEqual((obj.cx, obj.cy), (200_000, 3 * 9525))
            with self.assertRaises(FormulonError):
                wb.set_image_anchor(0, oid, anchor_kind=AnchorKind.ABSOLUTE)
            with self.assertRaises(FormulonError):
                wb.set_image_anchor(0, 9999)

    def test_set_image_z_order_reorders_the_list(self) -> None:
        with Workbook.create_default() as wb:
            a = wb.insert_image(0, _png(2, 2), row=0)
            b = wb.insert_image(0, _png(2, 2), row=3)
            c = wb.insert_image(0, _png(2, 2), row=6)
            self.assertEqual([o.object_id for o in wb.list_drawing_objects(0)], [a, b, c])
            wb.set_image_z_order(0, c, 0)
            self.assertEqual([o.object_id for o in wb.list_drawing_objects(0)], [c, a, b])
            with self.assertRaises(FormulonError):
                wb.set_image_z_order(0, c, 3)
            with self.assertRaises(FormulonError):
                wb.set_image_z_order(0, 9999, 0)

    def test_snapshot_and_restore_image(self) -> None:
        png = _png(4, 3)
        with Workbook.create_default() as wb:
            a = wb.insert_image(0, png, name="A", row=1)
            b = wb.insert_image(0, _png(2, 2), name="B", row=4)
            snap = wb.snapshot_image(0, a)
            self.assertIsInstance(snap, bytes)
            self.assertGreater(len(snap), 0)

            wb.remove_image(0, a)
            self.assertEqual(wb.restore_image(0, snap), a)
            self.assertEqual([o.object_id for o in wb.list_drawing_objects(0)], [a, b])
            self.assertEqual(wb.get_image(0, a), png)

            wb.set_image_anchor(0, a, row=9)
            wb.set_image_z_order(0, a, 1)
            self.assertEqual(wb.restore_image(0, snap), a)
            self.assertEqual([o.object_id for o in wb.list_drawing_objects(0)], [a, b])
            self.assertEqual(wb.list_drawing_objects(0)[0].from_row, 1)

            copy = wb.restore_image(0, snap, new_id=True)
            self.assertNotIn(copy, (a, b))
            self.assertEqual([o.object_id for o in wb.list_drawing_objects(0)], [a, b, copy])

            with self.assertRaises(FormulonError):
                wb.restore_image(0, b"junk")
            with self.assertRaises(FormulonError):
                wb.snapshot_image(0, 9999)


if __name__ == "__main__":
    unittest.main()
