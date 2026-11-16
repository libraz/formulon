"""A stored sparkline rectangle loads as its top-left cell; a location that is not one range drops its group, counted."""

from __future__ import annotations

import io
import unittest
import zipfile
from pathlib import Path

from formulon import SparklineGroupInput, SparklineInput, Workbook, WorkbookFormat

_SHEET = "xl/worksheets/sheet1.xml"
_EXCEL_XLSB = (
    Path(__file__).resolve().parents[3] / "tests/fixtures/excel/xlsb_parity/sparklines/multicell_location.xlsb"
)


def _rewrite_sheet(data: bytes, edit: dict[str, str]) -> bytes:
    """A copy of the package whose first sheet part has each key of ``edit`` replaced by its value."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            body = src.read(info.filename)
            if info.filename == _SHEET:
                text = body.decode("utf-8")
                for old, new in edit.items():
                    assert old in text
                    text = text.replace(old, new)
                body = text.encode("utf-8")
            dst.writestr(info.filename, body)
    return out.getvalue()


def _book_with_groups(locations: list[str]) -> bytes:
    with Workbook.create_default() as wb:
        ws = wb.sheet(0)
        for i, location in enumerate(locations):
            source = f"Sheet1!A{i + 2}:D{i + 2}"
            ws.add_sparkline_group(SparklineGroupInput(sparklines=[SparklineInput(source, location)]))
        return wb.save()


class SparklineLocationOnLoadTests(unittest.TestCase):
    def _locations(self, wb: Workbook) -> list[tuple[int, int]]:
        return [
            (g.sparklines[0].location.row, g.sparklines[0].location.col) for g in wb.sheet(0).list_sparkline_groups()
        ]

    def test_a_rectangle_loads_as_its_top_left_cell_without_a_diagnostic(self) -> None:
        data = _rewrite_sheet(
            _book_with_groups(["F2", "B2"]), {"<xm:sqref>B2</xm:sqref>": "<xm:sqref>B3:C4</xm:sqref>"}
        )
        with Workbook.load(data) as wb:
            self.assertEqual(self._locations(wb), [(1, 5), (2, 1)])
            self.assertEqual(wb.get_diagnostics().skipped_feature_count, 0)

    def test_groups_whose_location_is_not_one_range_are_dropped_and_counted(self) -> None:
        data = _rewrite_sheet(
            _book_with_groups(["F2", "B2", "C2", "D2"]),
            {
                "<xm:sqref>B2</xm:sqref>": "<xm:sqref>A1:B2</xm:sqref>",
                "<xm:sqref>C2</xm:sqref>": "<xm:sqref>zz</xm:sqref>",
                "<xm:sqref>D2</xm:sqref>": "<xm:sqref>A1 C3</xm:sqref>",
            },
        )
        with Workbook.load(data) as wb:
            self.assertEqual(self._locations(wb), [(1, 5), (0, 0)])
            self.assertEqual(wb.get_diagnostics().skipped_feature_count, 2)
            with Workbook.load(wb.save()) as reloaded:
                self.assertEqual(self._locations(reloaded), [(1, 5), (0, 0)])
                self.assertEqual(reloaded.get_diagnostics().skipped_feature_count, 0)

    def test_a_book_whose_only_group_is_unreadable_saves_no_sparkline_block(self) -> None:
        data = _rewrite_sheet(_book_with_groups(["B2"]), {"<xm:sqref>B2</xm:sqref>": "<xm:sqref>zz</xm:sqref>"})
        with Workbook.load(data) as wb:
            self.assertEqual(wb.sheet(0).list_sparkline_groups(), [])
            self.assertEqual(wb.get_diagnostics().skipped_feature_count, 1)
            with zipfile.ZipFile(io.BytesIO(wb.save())) as saved:
                self.assertNotIn("sparklineGroups", saved.read(_SHEET).decode("utf-8"))

    def test_an_excel_xlsb_rectangle_loads_as_its_top_left_cell(self) -> None:
        with Workbook.load(_EXCEL_XLSB.read_bytes()) as wb:
            self.assertEqual(self._locations(wb), [(0, 6)])
            self.assertEqual(wb.get_diagnostics().skipped_feature_count, 0)

    def test_an_xlsx_rectangle_is_written_to_xlsb_as_its_top_left_cell(self) -> None:
        data = _rewrite_sheet(_book_with_groups(["B2"]), {"<xm:sqref>B2</xm:sqref>": "<xm:sqref>A1:B2</xm:sqref>"})
        with Workbook.load(data) as wb:
            with Workbook.load(wb.save(format=WorkbookFormat.XLSB)) as back:
                self.assertEqual(self._locations(back), [(0, 0)])
                self.assertEqual(back.get_diagnostics().skipped_feature_count, 0)


if __name__ == "__main__":
    unittest.main()
