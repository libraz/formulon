"""Sheet.kind and Sheet.get_chart on a worksheet and on an Excel-saved chartsheet, from both containers."""

from __future__ import annotations

import unittest
from enum import IntEnum
from pathlib import Path

from formulon import NotFoundError, SheetKind, Workbook

_FIXTURES = Path(__file__).resolve().parents[3] / "tests/fixtures/excel/chartsheet_twins"


def _fixture(ext: str) -> bytes:
    """The Excel-saved workbook of the worksheet `Sheet1` then the chartsheet `Chart1`, which holds one chart."""
    return (_FIXTURES / f"tab_color_rgb_red.{ext}").read_bytes()


class SheetKindTests(unittest.TestCase):
    def test_sheet_kind_is_a_numeric_enum(self) -> None:
        self.assertTrue(issubclass(SheetKind, IntEnum))
        self.assertEqual(
            [(m.name, m.value) for m in SheetKind],
            [("WORKSHEET", 0), ("CHARTSHEET", 1), ("DIALOGSHEET", 2), ("MACROSHEET", 3), ("OTHER", 4)],
        )

    def test_kind_and_get_chart_tell_a_worksheet_from_a_chartsheet(self) -> None:
        for ext in ("xlsx", "xlsb"):
            with self.subTest(ext=ext), Workbook.load(_fixture(ext)) as wb:
                worksheet, chartsheet = wb.list_sheets()
                self.assertIs(worksheet.kind, SheetKind.WORKSHEET)
                self.assertIs(chartsheet.kind, SheetKind.CHARTSHEET)

                self.assertIsNone(worksheet.get_chart())
                chart = chartsheet.get_chart()
                self.assertIsNotNone(chart)
                assert chart is not None
                self.assertTrue(chart.get_properties().on_chartsheet)
                # The handle `chart(object_id)` returns for that chart.
                self.assertIs(chartsheet.chart(chart.object_id), chart)
                self.assertIs(chartsheet.get_chart(), chart)

    def test_a_new_workbook_holds_worksheets_and_kind_is_read_only(self) -> None:
        with Workbook.create_default() as wb:
            added = wb.add_sheet("Second")
            self.assertIs(wb.sheet(0).kind, SheetKind.WORKSHEET)
            self.assertIs(added.kind, SheetKind.WORKSHEET)
            self.assertIsNone(added.get_chart())
            with self.assertRaises(AttributeError):
                added.kind = SheetKind.CHARTSHEET  # type: ignore[misc]
            wb.remove_sheet(added)
            with self.assertRaises(NotFoundError):
                _ = added.kind
            with self.assertRaises(NotFoundError):
                added.get_chart()


if __name__ == "__main__":
    unittest.main()
