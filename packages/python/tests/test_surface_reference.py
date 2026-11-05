"""Reference text surface: ``parse_reference`` / ``format_reference``."""

from __future__ import annotations

import unittest

from formulon import CellRect, FormulonError, ParsedReference, Workbook

MAX_ROW = 1048575
MAX_COL = 16383


class ParseReferenceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)

    def test_cells_and_ranges(self) -> None:
        self.assertEqual(self.wb.parse_reference("A1"), ParsedReference(None, CellRect(0, 0, 0, 0)))
        self.assertEqual(self.wb.parse_reference("$B$2:C3"), ParsedReference(None, CellRect(1, 1, 2, 2)))
        self.assertEqual(self.wb.parse_reference("C3:B2"), ParsedReference(None, CellRect(1, 1, 2, 2)))
        self.assertEqual(self.wb.parse_reference("B:D"), ParsedReference(None, CellRect(0, 1, MAX_ROW, 3)))
        self.assertEqual(self.wb.parse_reference("2:11"), ParsedReference(None, CellRect(1, 0, 10, MAX_COL)))

    def test_sheet_qualifiers(self) -> None:
        self.assertEqual(self.wb.parse_reference("'My Sheet'!A1:B2"), ParsedReference("My Sheet", CellRect(0, 0, 1, 1)))
        self.assertEqual(self.wb.parse_reference("'O''Brien'!A1").sheet_name, "O'Brien")
        self.assertEqual(self.wb.parse_reference("売上!A1"), ParsedReference("売上", CellRect(0, 0, 0, 0)))
        self.assertEqual(self.wb.parse_reference("Sheet.1!A1").sheet_name, "Sheet.1")

    def test_invalid_text_is_none(self) -> None:
        for text in ("", " A1", "A1 ", "XFE1", "A1048577", "A0", "MyName", "A1,B2", "Sheet1:Sheet2!A1"):
            self.assertIsNone(self.wb.parse_reference(text), text)
        self.assertIsNone(self.wb.parse_reference("'" + "a" * 32 + "'!A1"))

    def test_r1c1(self) -> None:
        self.assertEqual(self.wb.parse_reference("R2C3", r1c1=True), ParsedReference(None, CellRect(1, 2, 1, 2)))
        self.assertEqual(
            self.wb.parse_reference("R[-1]C", r1c1=True, base_row=5, base_col=5),
            ParsedReference(None, CellRect(4, 5, 4, 5)),
        )
        self.assertIsNone(self.wb.parse_reference("R[-1]C", r1c1=True))

    def test_base_needs_both_axes(self) -> None:
        with self.assertRaises(ValueError):
            self.wb.parse_reference("R[-1]C", r1c1=True, base_row=1)
        with self.assertRaises(ValueError):
            self.wb.parse_reference("R[-1]C", r1c1=True, base_col=1)

    def test_base_outside_the_grid_raises(self) -> None:
        with self.assertRaises(FormulonError):
            self.wb.parse_reference("A1", base_row=MAX_ROW + 1, base_col=0)


class FormatReferenceTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)

    def test_shapes(self) -> None:
        self.assertEqual(self.wb.format_reference(CellRect(0, 0, 0, 0)), "A1")
        self.assertEqual(self.wb.format_reference(CellRect(0, 1, MAX_ROW, 3)), "B:D")
        self.assertEqual(self.wb.format_reference(CellRect(1, 0, 10, MAX_COL)), "2:11")
        self.assertEqual(self.wb.format_reference(CellRect(0, 0, 1, 1), absolute=True), "$A$1:$B$2")

    def test_sheet_quoting(self) -> None:
        cases = {"My Sheet": "'My Sheet'!A1", "A1": "'A1'!A1", "TRUE": "'TRUE'!A1", "1a": "'1a'!A1"}
        for name, expected in cases.items():
            self.assertEqual(self.wb.format_reference(CellRect(0, 0, 0, 0), sheet_name=name), expected)
        self.assertEqual(self.wb.format_reference(CellRect(0, 0, 0, 0), sheet_name="売上"), "売上!A1")
        self.assertEqual(self.wb.format_reference(CellRect(0, 0, 0, 0), sheet_name=""), "A1")

    def test_out_of_grid_and_reversed_raise(self) -> None:
        with self.assertRaises(FormulonError):
            self.wb.format_reference(CellRect(0, 0, MAX_ROW + 1, 0))
        with self.assertRaises(FormulonError):
            self.wb.format_reference(CellRect(0, 0, 0, MAX_COL + 1))
        with self.assertRaises(FormulonError):
            self.wb.format_reference(CellRect(2, 0, 1, 0))

    def test_rect_is_required(self) -> None:
        with self.assertRaises(TypeError):
            self.wb.format_reference((0, 0, 0, 0))  # type: ignore[arg-type]

    def test_round_trip(self) -> None:
        for name in (None, "Sheet1", "My Sheet", "O'Brien", "売上", "a!b"):
            text = self.wb.format_reference(CellRect(1, 2, 5, 7), sheet_name=name)
            self.assertEqual(self.wb.parse_reference(text), ParsedReference(name, CellRect(1, 2, 5, 7)), text)


if __name__ == "__main__":
    unittest.main()
