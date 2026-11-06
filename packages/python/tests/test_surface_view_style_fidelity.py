"""Outline, view, font, gradient, custom-icon, calc and defined-name surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    CellRect,
    CfIcon,
    Cfvo,
    ConditionalFormatInput,
    DifferentialFormat,
    Fill,
    Font,
    FormulonError,
    GradientFill,
    GradientStop,
    IconSet,
    RefMode,
    SheetFormatDefaults,
    Workbook,
)
from formulon.workbook import Color


class OutlineTests(unittest.TestCase):
    def test_collapsed_is_read_back_and_survives_save_load(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_row_outline(0, 1, 1, 2)
            wb.set_row_collapsed(0, 3, 3, True)
            wb.set_column_outline(0, 2, 3, 3)
            wb.set_column_collapsed(0, 4, 4, True)
            rows = {r.row: r for r in wb.list_row_overrides(0)}
            self.assertTrue(rows[3].collapsed)
            self.assertFalse(rows[1].collapsed)
            with Workbook.load(wb.save()) as reloaded:
                rows = {r.row: r for r in reloaded.list_row_overrides(0)}
                self.assertTrue(rows[3].collapsed)
                self.assertEqual(rows[1].outline_level, 2)
                self.assertEqual(sum(1 for c in reloaded.list_column_overrides(0) if c.collapsed), 1)

    def test_clearing_collapsed_drops_an_empty_override(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_row_collapsed(0, 5, 5, True)
            wb.set_row_collapsed(0, 5, 5, False)
            wb.set_column_collapsed(0, 1, 2, True)
            wb.set_column_collapsed(0, 1, 2, False)
            self.assertEqual(wb.list_row_overrides(0), [])
            self.assertEqual(wb.list_column_overrides(0), [])

    def test_bad_arguments_are_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.set_row_collapsed(9, 0, 0, True)
            with self.assertRaises(FormulonError):
                wb.set_column_collapsed(0, 3, 2, True)


class SheetViewFlagTests(unittest.TestCase):
    def test_show_formulas_and_outline_symbols(self) -> None:
        with Workbook.create_default() as wb:
            view = wb.get_view(0)
            self.assertFalse(view.show_formulas)
            self.assertTrue(view.show_outline_symbols)
            wb.set_show_formulas(0, True)
            wb.set_show_outline_symbols(0, False)
            with Workbook.load(wb.save()) as reloaded:
                view = reloaded.get_view(0)
                self.assertTrue(view.show_formulas)
                self.assertFalse(view.show_outline_symbols)


class FontToggleTests(unittest.TestCase):
    def test_toggles_round_trip_and_dedup(self) -> None:
        record = Font(
            name="Arial",
            size=12.0,
            color_argb=0xFF000000,
            has_outline=True,
            outline=True,
            has_shadow=True,
            shadow=False,
            has_condense=True,
            condense=True,
            has_extend=True,
            extend=True,
        )
        with Workbook.create_default() as wb:
            index = wb.add_font(record)
            self.assertEqual(wb.add_font(record), index)
            back = wb.get_font(index)
            self.assertTrue(back.outline and back.has_outline)
            self.assertTrue(back.has_shadow and not back.shadow)
            self.assertTrue(back.condense and back.extend)


class GradientFillTests(unittest.TestCase):
    @staticmethod
    def _fill() -> Fill:
        return Fill(
            gradient=GradientFill(
                type=1,
                left=0.5,
                right=0.5,
                top=0.25,
                bottom=0.75,
                stops=[
                    GradientStop(0.0, 0xFFFF0000, Color(kind=1, rgb=0xFFFF0000)),
                    GradientStop(1.0, 0, Color(kind=2, theme=4, tint=0.5)),
                ],
            )
        )

    def test_add_get_and_dedup(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.add_fill(self._fill())
            self.assertEqual(wb.add_fill(self._fill()), index)
            back = wb.get_fill(index)
            self.assertIsNotNone(back.gradient)
            assert back.gradient is not None
            self.assertEqual(back.gradient.type, 1)
            self.assertEqual((back.gradient.left, back.gradient.top, back.gradient.bottom), (0.5, 0.25, 0.75))
            self.assertEqual(len(back.gradient.stops), 2)
            self.assertEqual(back.gradient.stops[0].color_argb, 0xFFFF0000)
            self.assertEqual(back.gradient.stops[1].color.theme, 4)
            self.assertIsNone(wb.get_fill(0).gradient)

    def test_survives_save_load_in_cells_and_dxfs(self) -> None:
        with Workbook.create_default() as wb:
            fill_index = wb.add_fill(self._fill())
            dxf_index = wb.add_dxf(DifferentialFormat(fill=self._fill()))
            with Workbook.load(wb.save()) as reloaded:
                gradient = reloaded.get_fill(fill_index).gradient
                assert gradient is not None
                self.assertEqual(len(gradient.stops), 2)
                dxf_fill = reloaded.get_dxf(dxf_index).fill
                assert dxf_fill is not None and dxf_fill.gradient is not None
                self.assertEqual(dxf_fill.gradient.type, 1)


class CustomIconTests(unittest.TestCase):
    @staticmethod
    def _rule(icons: list[CfIcon]) -> ConditionalFormatInput:
        return ConditionalFormatInput(
            ranges=[CellRect(0, 0, 9, 0)],
            type=4,
            icon_set=IconSet(
                name=0,
                thresholds=[Cfvo(1, "33"), Cfvo(1, "67")],
                icons=icons,
            ),
        )

    def test_icons_round_trip_through_save_load(self) -> None:
        icons = [CfIcon("3Triangles", 0), CfIcon("NoIcons", 0), CfIcon("5Boxes", 4)]
        with Workbook.create_default() as wb:
            wb.add_cf_rule(0, self._rule(icons))
            with Workbook.load(wb.save()) as reloaded:
                back = reloaded.get_cf_rule(0, 0).icon_set
                assert back is not None
                self.assertEqual(back.icons, icons)

    def test_rules_without_icons_report_none(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_cf_rule(0, self._rule([]))
            back = wb.get_cf_rule(0, 0).icon_set
            assert back is not None
            self.assertEqual(back.icons, [])

    def test_malformed_icon_lists_are_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.add_cf_rule(0, self._rule([CfIcon("3Arrows", 0), CfIcon("3Arrows", 1)]))
            with self.assertRaises(FormulonError):
                wb.add_cf_rule(0, self._rule([CfIcon("3Arrows", 0), CfIcon("7Dwarfs", 1), CfIcon("3Arrows", 2)]))
            self.assertEqual([wb.get_cf_rule(0, i) for i in range(wb.get_cf_rule_count(0))], [])


class CalcOptionTests(unittest.TestCase):
    def test_defaults_and_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            self.assertTrue(wb.get_full_precision())
            self.assertTrue(wb.get_calc_on_save())
            self.assertFalse(wb.get_date1904())
            self.assertEqual(wb.get_ref_mode(), RefMode.A1)
            wb.set_full_precision(False)
            wb.set_calc_on_save(False)
            wb.set_ref_mode(RefMode.R1C1)
            wb.set_date1904(True)
            with Workbook.load(wb.save()) as reloaded:
                self.assertFalse(reloaded.get_full_precision())
                self.assertFalse(reloaded.get_calc_on_save())
                self.assertTrue(reloaded.get_date1904())
                self.assertEqual(reloaded.get_ref_mode(), RefMode.R1C1)

    def test_unknown_ref_mode_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.set_ref_mode(9)


class FileNameTests(unittest.TestCase):
    def test_file_name_spells_and_binds_the_workbooks_own_names(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.get_file_name(), "")
            wb.set_defined_name("LN", "10")
            wb.set_file_name("M5.xlsx")
            self.assertEqual(wb.get_file_name(), "M5.xlsx")
            wb.set_formula(0, 0, 0, "=M5.xlsx!LN+1")
            self.assertEqual(wb.get_formula(0, 0, 0), "=M5.xlsx!LN+1")
            self.assertEqual([wb.get_cell_at(0, i) for i in range(wb.get_cell_count(0))][0].formula, "=M5.xlsx!LN+1")
            wb.set_file_name("1M5.xlsx")
            self.assertEqual(wb.get_formula(0, 0, 0), "='1M5.xlsx'!LN+1")
            with self.assertRaises(FormulonError):
                wb.set_file_name("Book1")
            self.assertEqual(wb.get_file_name(), "1M5.xlsx")
            wb.set_file_name("")
            self.assertEqual(wb.get_formula(0, 0, 0), "=[0]!LN+1")


class FormatDefaultFlagTests(unittest.TestCase):
    def test_custom_height_and_zero_height_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            defaults = wb.get_format_defaults(0)
            self.assertFalse(defaults.custom_height)
            self.assertFalse(defaults.zero_height)
            wb.set_format_defaults(
                0,
                SheetFormatDefaults(
                    default_row_height=18.0, has_default_row_height=True, custom_height=True, zero_height=True
                ),
            )
            with Workbook.load(wb.save()) as reloaded:
                back = reloaded.get_format_defaults(0)
                self.assertTrue(back.custom_height)
                self.assertTrue(back.zero_height)
                self.assertEqual(back.default_row_height, 18.0)
                reloaded.set_format_defaults(
                    0, SheetFormatDefaults(default_row_height=20.0, has_default_row_height=True)
                )
                kept = reloaded.get_format_defaults(0)
                self.assertTrue(kept.custom_height)
                self.assertTrue(kept.zero_height)
                self.assertEqual(kept.default_row_height, 20.0)
                reloaded.set_format_defaults(0, SheetFormatDefaults(custom_height=False, zero_height=False))
                cleared = reloaded.get_format_defaults(0)
                self.assertFalse(cleared.custom_height)
                self.assertFalse(cleared.zero_height)


class DefinedNameAttrTests(unittest.TestCase):
    def test_comment_and_hidden(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_defined_name("Rate", "Sheet1!$A$1")
            wb.set_defined_name_attrs("rate", comment="Tax rate", hidden=True)
            entry = wb.list_defined_names()[0]
            self.assertEqual((entry.comment, entry.hidden), ("Tax rate", True))

            # A replacement keeps both; an omitted argument leaves one alone.
            wb.set_defined_name_scoped("Rate", "Sheet1!$B$2", -1)
            wb.set_defined_name_attrs("Rate", hidden=False)
            entry = wb.list_defined_names()[0]
            self.assertEqual((entry.comment, entry.hidden), ("Tax rate", False))

            wb.set_defined_name_attrs("Rate", comment="", hidden=True)
            with Workbook.load(wb.save()) as reloaded:
                entry = reloaded.list_defined_names()[0]
                self.assertEqual((entry.formula, entry.comment, entry.hidden), ("Sheet1!$B$2", "", True))

    def test_missing_name_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.set_defined_name_attrs("Nope", comment="x")


if __name__ == "__main__":
    unittest.main()
