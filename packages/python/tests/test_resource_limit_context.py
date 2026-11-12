"""Engine ceilings raise ``SecurityError`` with ``StatusCode.SECURITY_RESOURCE_LIMIT``.

The error context leads with ``part=<name> requested=<n> limit=<n>``.
"""

from __future__ import annotations

import unittest

from formulon import CellRect, FormulonError, SecurityError, StatusCode, Workbook


class ResourceLimitContextTests(unittest.TestCase):
    def _assert_limit(self, caught: FormulonError, part: str, requested: int, limit: int) -> None:
        self.assertIsInstance(caught, SecurityError)
        self.assertEqual(caught.code, StatusCode.SECURITY_RESOURCE_LIMIT)
        self.assertEqual(" ".join(caught.context.split(" ")[:3]), f"part={part} requested={requested} limit={limit}")

    def test_values_over_the_cell_cap(self) -> None:
        wide = CellRect(first_row=0, first_col=0, last_row=2048, last_col=2047)
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for read in (ws.get_values, ws.get_typed_values):
                with self.subTest(read=read.__name__):
                    with self.assertRaises(SecurityError) as caught:
                        read(wide)
                    self._assert_limit(caught.exception, "values_range", 2049 * 2048, 1 << 22)

    def test_print_fragment_over_its_limit(self) -> None:
        sheet_pr = "<sheetPr>" + "x" * 9000 + "</sheetPr>"
        with Workbook.create_default() as wb:
            with self.assertRaises(SecurityError) as caught:
                wb.sheet(0).ooxml.set_sheet_pr_xml(sheet_pr)
            self._assert_limit(caught.exception, "sheet_pr_xml", len(sheet_pr), 8192)


if __name__ == "__main__":
    unittest.main()
