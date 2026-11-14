"""Sheet.get_cell_picture / set_cell_picture: a picture placed in a cell reads back with its format, size and alt
text; None removes it; a refused picture leaves the cell as it was."""

from __future__ import annotations

import dataclasses
import unittest

from formulon import (
    CellPicture,
    CellPictureInput,
    ErrorCode,
    ExcelError,
    ImageFormat,
    InvalidArgumentError,
    IoError,
    SecurityError,
    StatusCode,
    Workbook,
    WorkbookFormat,
)


def png(width: int, height: int) -> bytes:
    return (
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
        + width.to_bytes(4, "big")
        + height.to_bytes(4, "big")
        + b"\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
    )


def jpeg(width: int, height: int) -> bytes:
    frame = b"\xff\xc0\x00\x11\x08" + height.to_bytes(2, "big") + width.to_bytes(2, "big")
    return b"\xff\xd8\xff" + frame[1:] + b"\x03\x01\x11\x00\x02\x11\x01\x03\x11\x01\xff\xd9"


GIF = b"GIF89a\x04\x00\x04\x00\x00\x00\x00;"


class CellPictureTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_places_a_png_with_alt_text_and_reads_it_back(self) -> None:
        data = png(16, 8)
        self.ws.set_cell_picture(1, 2, CellPictureInput(data, "A red square"))
        picture = self.ws.get_cell_picture(1, 2)
        assert picture is not None
        self.assertIsInstance(picture, CellPicture)
        self.assertEqual(picture.bytes, data)
        self.assertEqual(picture.alt_text, "A red square")
        self.assertEqual(picture.format, ImageFormat.PNG)
        self.assertEqual((picture.width_px, picture.height_px), (16, 8))
        self.assertEqual(self.ws.get_cell_picture("C2"), picture)

    def test_a_jpeg_without_alt_text_is_stored_as_given(self) -> None:
        data = jpeg(40, 24)
        self.ws.set_cell_picture("A1", CellPictureInput(data))
        self.ws.set_cell_picture("A2", CellPictureInput(data, alt_text=None))
        for a1 in ("A1", "A2"):
            picture = self.ws.get_cell_picture(a1)
            assert picture is not None
            self.assertEqual(picture.bytes, data)
            self.assertIsNone(picture.alt_text)
            self.assertEqual(picture.format, ImageFormat.JPEG)
            self.assertEqual((picture.width_px, picture.height_px), (40, 24))

    def test_row_col_and_a1_forms_address_the_same_cell(self) -> None:
        self.ws.set_cell_picture("D5", CellPictureInput(png(2, 2), "by a1"))
        picture = self.ws.get_cell_picture(4, 3)
        assert picture is not None
        self.assertEqual(picture.alt_text, "by a1")
        self.ws.set_cell_picture(5, 3, CellPictureInput(png(3, 3), "by row and col"))
        picture = self.ws.get_cell_picture("D6")
        assert picture is not None
        self.assertEqual(picture.alt_text, "by row and col")

    def test_bytes_may_be_a_bytearray_or_a_memoryview(self) -> None:
        data = png(5, 5)
        self.ws.set_cell_picture("A1", CellPictureInput(bytearray(data)))
        self.ws.set_cell_picture("A2", CellPictureInput(memoryview(data)))
        for a1 in ("A1", "A2"):
            picture = self.ws.get_cell_picture(a1)
            assert picture is not None
            self.assertEqual(picture.bytes, data)

    def test_a_picture_read_back_is_accepted_and_read_only_fields_are_ignored(self) -> None:
        self.ws.set_cell_picture("A1", CellPictureInput(png(7, 3), "copy me"))
        picture = self.ws.get_cell_picture("A1")
        assert picture is not None
        self.ws.set_cell_picture("B1", picture)
        self.assertEqual(self.ws.get_cell_picture("B1"), picture)

        wrong = dataclasses.replace(picture, bytes=png(9, 4), format=ImageFormat.GIF, width_px=999, height_px=999)
        self.ws.set_cell_picture("C1", wrong)
        stored = self.ws.get_cell_picture("C1")
        assert stored is not None
        self.assertEqual(stored.format, ImageFormat.PNG)
        self.assertEqual((stored.width_px, stored.height_px), (9, 4))

    def test_a_cell_without_a_picture_reads_none(self) -> None:
        self.assertIsNone(self.ws.get_cell_picture("A1"))
        self.ws.set_value("A1", 12)
        self.assertIsNone(self.ws.get_cell_picture(0, 0))

    def test_none_removes_the_picture_empties_the_cell_and_is_a_no_op_when_absent(self) -> None:
        self.ws.set_cell_picture("B2", CellPictureInput(png(2, 2), "x"))
        self.ws.set_cell_picture("B2", None)
        self.assertIsNone(self.ws.get_cell_picture("B2"))
        self.assertIsNone(self.ws.get_value("B2"))
        self.ws.set_cell_picture("B2", None)
        self.ws.set_value("C3", 5)
        self.ws.set_cell_picture("C3", None)
        self.assertEqual(self.ws.get_value("C3"), 5)

    def test_a_value_written_over_a_picture_drops_it_and_a_picture_replaces_a_formula(self) -> None:
        self.ws.set_cell_picture("A1", CellPictureInput(png(2, 2)))
        self.ws.set_value("A1", 3)
        self.assertIsNone(self.ws.get_cell_picture("A1"))
        self.ws.set_formula("A2", "=1+1")
        self.ws.set_cell_picture("A2", CellPictureInput(png(2, 2)))
        self.assertIsNotNone(self.ws.get_cell_picture("A2"))

    def test_a_formula_reading_a_picture_cell_evaluates_it_as_value_error(self) -> None:
        self.ws.set_cell_picture("A1", CellPictureInput(png(2, 2), "alt"))
        self.ws.set_formula("B1", "=A1")
        self.ws.set_formula("C1", "=ISERROR(A1)")
        self.wb.recalc()
        self.assertEqual(self.ws.get_value("A1"), ExcelError(ErrorCode.VALUE))
        self.assertEqual(self.ws.get_value("B1"), ExcelError(ErrorCode.VALUE))
        self.assertIs(self.ws.get_value("C1"), True)

    def test_empty_data_and_other_image_formats_are_refused_and_change_nothing(self) -> None:
        original = png(2, 2)
        self.ws.set_cell_picture("A1", CellPictureInput(original, "keep"))
        for data in (b"", GIF, b"not an image"):
            for a1 in ("A1", "B1"):
                with self.assertRaises(IoError) as raised:
                    self.ws.set_cell_picture(a1, CellPictureInput(data))
                self.assertEqual(raised.exception.code, StatusCode.IO_IMAGE_UNSUPPORTED)
        self.assertIsNone(self.ws.get_cell_picture("B1"))
        kept = self.ws.get_cell_picture("A1")
        assert kept is not None
        self.assertEqual((kept.bytes, kept.alt_text), (original, "keep"))

    def test_data_over_32_mib_is_refused_as_a_resource_limit(self) -> None:
        data = png(2, 2) + bytes(32 * 1024 * 1024)
        with self.assertRaises(SecurityError):
            self.ws.set_cell_picture("A1", CellPictureInput(data))
        self.assertIsNone(self.ws.get_cell_picture("A1"))

    def test_empty_alt_text_and_coordinates_past_the_grid_are_invalid_arguments(self) -> None:
        with self.assertRaises(InvalidArgumentError):
            self.ws.set_cell_picture("A1", CellPictureInput(png(2, 2), ""))
        self.assertIsNone(self.ws.get_cell_picture("A1"))
        with self.assertRaises(InvalidArgumentError):
            self.ws.get_cell_picture(1048576, 0)
        with self.assertRaises(InvalidArgumentError):
            self.ws.set_cell_picture(0, 16384, None)
        with self.assertRaises(InvalidArgumentError):
            self.ws.get_cell_picture("Sheet1!A1")

    def test_a_malformed_picture_argument_is_a_type_error(self) -> None:
        data = png(2, 2)
        for bad in ("picture", b"bytes", {"bytes": data}, 5):
            with self.assertRaises(TypeError):
                self.ws.set_cell_picture("A1", bad)  # type: ignore[call-overload]
        with self.assertRaises(TypeError):
            self.ws.set_cell_picture("A1", CellPictureInput("abc"))  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.set_cell_picture("A1", CellPictureInput(5))  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.set_cell_picture("A1", CellPictureInput(data, 5))  # type: ignore[arg-type]
        self.assertIsNone(self.ws.get_cell_picture("A1"))

    def test_records_are_frozen_where_read_and_mutable_where_input(self) -> None:
        self.ws.set_cell_picture("A1", CellPictureInput(png(2, 2)))
        picture = self.ws.get_cell_picture("A1")
        assert picture is not None
        with self.assertRaises(dataclasses.FrozenInstanceError):
            picture.alt_text = "x"  # type: ignore[misc]

    def test_a_picture_survives_a_save_and_load_as_xlsx_and_xlsb(self) -> None:
        first = png(12, 9)
        second = jpeg(30, 20)
        self.ws.set_cell_picture("B2", CellPictureInput(first, "a chart"))
        self.ws.set_cell_picture("D3", CellPictureInput(second))
        for fmt in (WorkbookFormat.XLSX, WorkbookFormat.XLSB):
            with self.subTest(format=fmt.name):
                loaded = Workbook.load(self.wb.save(format=fmt))
                self.addCleanup(loaded.close)
                sheet = loaded.sheet(0)
                a = sheet.get_cell_picture("B2")
                assert a is not None
                self.assertEqual((a.bytes, a.alt_text, a.format), (first, "a chart", ImageFormat.PNG))
                b = sheet.get_cell_picture("D3")
                assert b is not None
                self.assertEqual((b.bytes, b.alt_text, b.format), (second, None, ImageFormat.JPEG))
                self.assertIsNone(sheet.get_cell_picture("A1"))


if __name__ == "__main__":
    unittest.main()
