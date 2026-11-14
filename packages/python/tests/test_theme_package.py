"""Workbook.apply_theme / save_theme: a theme file (.thmx) round-trips between workbooks; a refused file leaves the
theme unchanged and raises with the status of its defect."""

from __future__ import annotations

import unittest

from formulon import InvalidArgumentError, InvalidFileError, StatusCode, ThemeFonts, Workbook

FONTS = ThemeFonts("Major Source", "Major EA", "Minor Source", "Minor EA")


class ThemePackageTests(unittest.TestCase):
    def setUp(self) -> None:
        self.source = Workbook.create_default()
        self.target = Workbook.create_default()
        self.addCleanup(self.source.close)
        self.addCleanup(self.target.close)

    def test_save_theme_of_a_fresh_workbook_is_a_zip_with_the_theme_part(self) -> None:
        data = self.source.save_theme()
        self.assertIsInstance(data, bytes)
        self.assertEqual(data[:2], b"PK")
        self.assertIn(b"theme/theme/theme1.xml", data)

    def test_save_then_apply_carries_the_theme_fonts(self) -> None:
        self.source.set_theme_fonts(FONTS)
        expected = self.source.get_theme_fonts()
        self.target.apply_theme(self.source.save_theme())
        self.assertEqual(self.target.get_theme_fonts(), expected)
        self.assertEqual(self.target.get_theme_fonts().minor_latin, "Minor Source")

    def test_apply_accepts_a_bytearray_and_a_memoryview(self) -> None:
        self.source.set_theme_fonts(FONTS)
        data = self.source.save_theme()
        self.target.apply_theme(bytearray(data))
        self.assertEqual(self.target.get_theme_fonts().major_latin, "Major Source")
        self.target.reset_theme()
        self.target.apply_theme(memoryview(data))
        self.assertEqual(self.target.get_theme_fonts().major_latin, "Major Source")

    def test_empty_bytes_raise_invalid_argument_and_change_nothing(self) -> None:
        self.target.set_theme_fonts(FONTS)
        before = self.target.get_theme_fonts()
        with self.assertRaises(InvalidArgumentError):
            self.target.apply_theme(b"")
        self.assertEqual(self.target.get_theme_fonts(), before)

    def test_bytes_that_are_not_a_zip_raise_zip_corrupt_and_change_nothing(self) -> None:
        self.target.set_theme_fonts(FONTS)
        before = self.target.get_theme_fonts()
        with self.assertRaises(InvalidFileError) as caught:
            self.target.apply_theme(b"not a zip archive")
        self.assertEqual(caught.exception.code, StatusCode.IO_ZIP_CORRUPT)
        self.assertEqual(self.target.get_theme_fonts(), before)


if __name__ == "__main__":
    unittest.main()
