"""A save leaves the live differential-format table alone, and an unresolvable dxf index is refused."""

from __future__ import annotations

import unittest

import formulon
from formulon import (
    AutoFilter,
    AutoFilterColumn,
    CellRect,
    DxfInput,
    FilterKind,
    Font,
    Workbook,
)


class DxfSaveTests(unittest.TestCase):
    def test_save_does_not_change_the_live_dxf_table(self) -> None:
        with Workbook.create_default() as wb:
            wb.ooxml.add_dxf(DxfInput(font=Font(bold=True)))
            wb.ooxml.add_dxf(DxfInput(number_format="0.000"))
            before = wb.ooxml.list_dxfs()
            self.assertGreaterEqual(len(before), 2)
            saved = wb.save()
            after = wb.ooxml.list_dxfs()
            self.assertEqual(len(after), len(before))
            for i, dxf in enumerate(before):
                self.assertEqual(wb.ooxml.get_dxf(i), dxf)
            self.assertEqual(after, before)
            with Workbook.load(saved) as reloaded:
                self.assertLessEqual(len(reloaded.ooxml.list_dxfs()), len(before))

    def test_color_filter_dxf_index_past_the_table_is_refused(self) -> None:
        with Workbook.create_default() as wb:
            sheet = wb.sheet(0)
            count = len(wb.ooxml.list_dxfs())

            def color(dxf: int) -> AutoFilter:
                column = AutoFilterColumn(column_offset=0, kind=FilterKind.COLOR, cell_color=True, dxf_index=dxf)
                return AutoFilter(range=CellRect(0, 0, 3, 0), columns=[column])

            with self.assertRaises(formulon.IndexOutOfRangeError) as ctx:
                sheet.set_auto_filter(color(count))
            self.assertEqual(ctx.exception.code, formulon.StatusCode.INDEX_OUT_OF_RANGE)
            self.assertIsNone(sheet.get_auto_filter())


if __name__ == "__main__":
    unittest.main()
