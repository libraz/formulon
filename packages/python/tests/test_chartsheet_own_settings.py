"""A chartsheet's protection, margins, header and footer, paper size and orientation, set and saved."""

from __future__ import annotations

import unittest
from pathlib import Path

from formulon import (
    ChartsheetProtection,
    ChartsheetProtectionInput,
    InvalidArgumentError,
    PageOrientation,
    Workbook,
    WorkbookFormat,
)

_FIXTURES = Path(__file__).resolve().parents[3] / "tests/fixtures/excel/chartsheet_twins"
_FORMATS = (("xlsx", WorkbookFormat.XLSX), ("xlsb", WorkbookFormat.XLSB))


def _load(name: str, ext: str) -> Workbook:
    """The Excel-saved workbook of the worksheet `Sheet1` then the chartsheet, from the twin `name`."""
    return Workbook.load((_FIXTURES / f"{name}.{ext}").read_bytes())


class ChartsheetOwnSettingsTests(unittest.TestCase):
    def test_set_takes_back_what_get_reads(self) -> None:
        for ext, _fmt in _FORMATS:
            with self.subTest(ext=ext), _load("cs_protect_hash", ext) as wb:
                chartsheet = wb.sheet(1)
                read = chartsheet.get_chartsheet_protection()
                assert read is not None
                self.assertEqual(read.algorithm_name, "SHA-512")
                chartsheet.set_chartsheet_protection(read)
                self.assertEqual(chartsheet.get_chartsheet_protection(), read)
                chartsheet.set_chartsheet_protection(ChartsheetProtectionInput(allow_edit_objects=True))
                self.assertEqual(
                    chartsheet.get_chartsheet_protection(),
                    ChartsheetProtection(allow_edit_content=False, allow_edit_objects=True),
                )
                with self.assertRaises(InvalidArgumentError):
                    chartsheet.set_chartsheet_protection(ChartsheetProtectionInput(hash_value=""))
                with self.assertRaises(TypeError):
                    chartsheet.set_chartsheet_protection({})  # type: ignore[arg-type]

    def test_a_protection_is_never_removed(self) -> None:
        unprotected = ChartsheetProtectionInput(allow_edit_content=True, allow_edit_objects=True)
        for ext, _fmt in _FORMATS:
            with self.subTest(ext=ext), _load("cs_control", ext) as wb:
                chartsheet = wb.sheet(1)
                chartsheet.set_chartsheet_protection(None)
                chartsheet.set_chartsheet_protection(unprotected)
                self.assertIsNone(chartsheet.get_chartsheet_protection())
                chartsheet.set_chartsheet_protection(ChartsheetProtectionInput())
                for refused in (None, unprotected):
                    with self.assertRaisesRegex(InvalidArgumentError, "Chartsheet protection cannot be removed"):
                        chartsheet.set_chartsheet_protection(refused)
                with self.assertRaisesRegex(InvalidArgumentError, "needs a chartsheet"):
                    wb.sheet(0).set_chartsheet_protection(ChartsheetProtectionInput())

    def test_paper_size_margins_and_header_apply_to_a_chartsheet(self) -> None:
        for ext, fmt in _FORMATS:
            with self.subTest(ext=ext), _load("cs_control", ext) as wb:
                chartsheet = wb.sheet(1)
                chartsheet.update_page_setup(paper_size=5)
                chartsheet.update_page_margins(
                    left_in=0.5, right_in=0.75, top_in=1.0, bottom_in=1.25, header_in=0.25, footer_in=0.375
                )
                chartsheet.update_header_footer(odd_header="&CHDR")
                with self.assertRaisesRegex(InvalidArgumentError, "scale_pct states the scale"):
                    chartsheet.update_page_setup(scale_pct=50)
                with Workbook.load(wb.save(format=fmt)) as reloaded:
                    again = reloaded.sheet(1)
                    self.assertEqual(again.get_page_setup().paper_size, 5)
                    self.assertIs(again.get_page_setup().orientation, PageOrientation.PORTRAIT)
                    self.assertEqual(again.get_page_margins().left_in, 0.5)
                    self.assertEqual(again.get_header_footer().odd_header, "&CHDR")


if __name__ == "__main__":
    unittest.main()
