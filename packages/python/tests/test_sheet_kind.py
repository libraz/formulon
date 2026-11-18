"""Sheet.kind and Sheet.get_chart on a worksheet and on an Excel-saved chartsheet, from both containers."""

from __future__ import annotations

import unittest
from enum import IntEnum
from pathlib import Path

from formulon import (
    ChartsheetProtection,
    InvalidArgumentError,
    NotFoundError,
    SheetKind,
    SheetProtectionInput,
    SheetVisibility,
    Workbook,
    WorkbookFormat,
)

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

    def test_zoom_to_fit_reads_the_stored_fit_and_a_new_zoom_ends_it(self) -> None:
        for ext in ("xlsx", "xlsb"):
            with self.subTest(ext=ext), Workbook.load(_fixture(ext)) as wb:
                worksheet, chartsheet = wb.list_sheets()
                self.assertFalse(worksheet.get_view().zoom_to_fit)
                self.assertTrue(chartsheet.get_view().zoom_to_fit)
                with self.assertRaises(TypeError):
                    chartsheet.update_view(zoom_to_fit=False)  # type: ignore[call-arg]
                with self.assertRaises(InvalidArgumentError):
                    chartsheet.update_view(show_grid_lines=not chartsheet.get_view().show_grid_lines)
                chartsheet.update_view(zoom_pct=80)
                view = chartsheet.get_view()
                self.assertEqual(view.zoom_pct, 80)
                self.assertFalse(view.zoom_to_fit)

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


class ChartsheetProtectionTests(unittest.TestCase):
    def test_reads_what_a_chartsheet_part_states(self) -> None:
        for ext in ("xlsx", "xlsb"):
            contents_bytes = (_FIXTURES / f"prot_contents_only.{ext}").read_bytes()
            both_bytes = (_FIXTURES / f"prot_both_pw_abc.{ext}").read_bytes()
            with (
                self.subTest(ext=ext),
                Workbook.load(contents_bytes) as contents_only,
                Workbook.load(both_bytes) as both,
                Workbook.load(_fixture(ext)) as plain,
            ):
                self.assertEqual(
                    contents_only.sheet(1).get_chartsheet_protection(),
                    ChartsheetProtection(allow_edit_content=False, allow_edit_objects=True),
                )
                hashed = both.sheet(1).get_chartsheet_protection()
                assert hashed is not None
                self.assertEqual(hashed.algorithm_name, "SHA-512")
                self.assertEqual(hashed.salt_value, "BAMwdck9CSCcwdaFLEANaA==")
                self.assertEqual(hashed.spin_count, 100000)
                self.assertIsNone(hashed.legacy_password_hash)
                self.assertFalse(hashed.allow_edit_content)
                self.assertFalse(hashed.allow_edit_objects)
                # A worksheet, and a chartsheet that states none, read None.
                self.assertIsNone(contents_only.sheet(0).get_chartsheet_protection())
                self.assertIsNone(plain.sheet(1).get_chartsheet_protection())
                # A chartsheet refuses the worksheet protection calls, and the record is not their input.
                with self.assertRaises(InvalidArgumentError):
                    contents_only.sheet(1).get_sheet_protection()
                with self.assertRaises(TypeError):
                    plain.sheet(0).set_sheet_protection(hashed)  # type: ignore[arg-type]

    def test_a_worksheet_reads_none_and_a_removed_sheet_raises(self) -> None:
        with Workbook.create_default() as wb:
            added = wb.add_sheet("Second")
            self.assertIsNone(wb.sheet(0).get_chartsheet_protection())
            added.set_sheet_protection(SheetProtectionInput())
            self.assertIsNone(added.get_chartsheet_protection())
            wb.remove_sheet(added)
            with self.assertRaises(NotFoundError):
                added.get_chartsheet_protection()

    def test_reads_objects_only_legacy_only_and_bare_parts(self) -> None:
        authored = Path(__file__).resolve().parents[3] / "tests/fixtures/chartsheet_protection"
        cases = (
            ("objects_only", ChartsheetProtection(allow_edit_content=True, allow_edit_objects=False), True),
            (
                "legacy_only",
                ChartsheetProtection(legacy_password_hash="CBEB", allow_edit_content=False, allow_edit_objects=True),
                True,
            ),
            # An xlsb part states protection only through set flags or a hash, so a bare element reads None there.
            ("bare", ChartsheetProtection(allow_edit_content=True, allow_edit_objects=True), False),
        )
        for name, expected, survives_xlsb in cases:
            with self.subTest(name=name), Workbook.load((authored / f"{name}.xlsx").read_bytes()) as wb:
                self.assertEqual(wb.sheet(1).get_chartsheet_protection(), expected)
                with Workbook.load(wb.save(format=WorkbookFormat.XLSB)) as reloaded:
                    self.assertEqual(reloaded.sheet(1).get_chartsheet_protection(), expected if survives_xlsb else None)


if __name__ == "__main__":
    unittest.main()

    def test_a_chartsheet_answers_from_its_part_and_refuses_worksheet_content(self) -> None:
        for ext in ("xlsx", "xlsb"):
            with self.subTest(ext=ext), Workbook.load(_fixture(ext)) as wb:
                chartsheet = wb[1]
                # Cell reads and writes need a worksheet.
                with self.assertRaises(InvalidArgumentError):
                    chartsheet.get_value("A1")
                with self.assertRaises(InvalidArgumentError):
                    chartsheet.set_value("A1", 1)
                # A list of worksheet content is empty.
                self.assertEqual(chartsheet.list_merges(), [])
                # The tab colour is read from the part; a new one is refused.
                self.assertIsNotNone(chartsheet.get_tab_color())
                with self.assertRaises(InvalidArgumentError):
                    chartsheet.set_tab_color(None)
                # Visibility serves every kind.
                self.assertEqual(chartsheet.get_visibility(), SheetVisibility.VISIBLE)
                # The view takes a new zoom and tab selection, and refuses any other change.
                with self.assertRaises(InvalidArgumentError):
                    chartsheet.update_view(show_grid_lines=not chartsheet.get_view().show_grid_lines)
                chartsheet.update_view(tab_selected=True)
                self.assertTrue(chartsheet.get_view().tab_selected)
