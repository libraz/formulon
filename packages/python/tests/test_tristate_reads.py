"""Tri-state flags read back as booleans: ``FALSE`` is ``False``, not a non-zero (truthy) integer."""

from __future__ import annotations

import unittest

from formulon import RowBreak, Workbook


class BreakManualFlagTests(unittest.TestCase):
    def test_pagination_reports_automatic_breaks_as_not_manual(self) -> None:
        sheet = Workbook.create_default().sheet(0)
        for row in range(200):
            for col in range(30):
                sheet.set_number(row, col, row * col)
        pagination = sheet.get_pagination()
        self.assertTrue(pagination.row_breaks)
        self.assertTrue(pagination.column_breaks)
        self.assertTrue(all(brk.manual is False for brk in pagination.row_breaks))
        self.assertTrue(all(brk.manual is False for brk in pagination.column_breaks))

    def test_pagination_reports_a_placed_break_as_manual(self) -> None:
        sheet = Workbook.create_default().sheet(0)
        for row in range(200):
            sheet.set_number(row, 0, row)
        sheet.set_row_break(10, RowBreak(row=10, first=0, last=0, manual=True))
        pagination = sheet.get_pagination()
        self.assertEqual({brk.row: brk.manual for brk in pagination.row_breaks}.get(10), True)
        self.assertTrue(any(brk.manual is False for brk in pagination.row_breaks))


if __name__ == "__main__":
    unittest.main()
