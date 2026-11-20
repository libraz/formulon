"""A workbook-scope name hidden from a sheet by a same-name sheet-local name reads as ``[0]!Name``."""

from __future__ import annotations

import unittest
from pathlib import Path

from formulon import Workbook

_FIXTURE = Path(__file__).resolve().parents[3] / "tests/fixtures/excel/names/workbook_print_area_hidden_by_local.xlsb"


class HiddenWorkbookNameTests(unittest.TestCase):
    def test_loaded_formula_reads_and_reenters_as_the_self_book_form(self) -> None:
        with Workbook.load(_FIXTURE) as wb:
            sheet = wb.sheet(0)
            shown = sheet.get_formula("E1")
            self.assertEqual(shown, "=ROWS([0]!Print_Area)")
            self.assertEqual(sheet.get_value("E1"), 3)

            sheet.set_formula("F1", shown)
            self.assertEqual(sheet.get_formula("F1"), shown)
            self.assertEqual(sheet.get_value("F1"), 3)

    def test_bare_name_still_reads_the_sheets_own(self) -> None:
        with Workbook.load(_FIXTURE) as wb:
            sheet = wb.sheet(0)
            sheet.set_formula("F1", "=ROWS(Print_Area)")
            self.assertEqual(sheet.get_formula("F1"), "=ROWS(Print_Area)")
            self.assertEqual(sheet.get_value("F1"), 5)


if __name__ == "__main__":
    unittest.main()
