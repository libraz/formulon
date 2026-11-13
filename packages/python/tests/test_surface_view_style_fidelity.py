"""Outline, view, font, gradient, custom-icon, calc and defined-name surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    CellRect,
    CfExpressionRuleInput,
    CfIcon,
    CfIconSetInput,
    CfIconSetRuleInput,
    CfValueObject,
    Color,
    ColorKind,
    ConditionalFormatInput,
    DefinedNameInput,
    DxfInput,
    Fill,
    Font,
    FormulonError,
    GradientFill,
    GradientStop,
    IconSetName,
    InvalidArgumentError,
    RefMode,
    Workbook,
    WorkbookLocation,
    WorkbookLocationInput,
)


class OutlineTests(unittest.TestCase):
    def test_collapsed_is_read_back_and_survives_save_load(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_row_outline(1, 1, 2)
            wb.sheet(0).set_row_collapsed(3, 3, True)
            wb.sheet(0).set_column_outline(2, 3, 3)
            wb.sheet(0).set_column_collapsed(4, 4, True)
            rows = {r.row: r for r in wb.sheet(0).list_row_overrides()}
            self.assertTrue(rows[3].collapsed)
            self.assertFalse(rows[1].collapsed)
            with Workbook.load(wb.save()) as reloaded:
                rows = {r.row: r for r in reloaded.sheet(0).list_row_overrides()}
                self.assertTrue(rows[3].collapsed)
                self.assertEqual(rows[1].outline_level, 2)
                self.assertEqual(sum(1 for c in reloaded.sheet(0).list_column_overrides() if c.collapsed), 1)

    def test_clearing_collapsed_drops_an_empty_override(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_row_collapsed(5, 5, True)
            wb.sheet(0).set_row_collapsed(5, 5, False)
            wb.sheet(0).set_column_collapsed(1, 2, True)
            wb.sheet(0).set_column_collapsed(1, 2, False)
            self.assertEqual(wb.sheet(0).list_row_overrides(), [])
            self.assertEqual(wb.sheet(0).list_column_overrides(), [])

    def test_bad_arguments_are_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(9).set_row_collapsed(0, 0, True)
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_column_collapsed(3, 2, True)


class SheetViewFlagTests(unittest.TestCase):
    def test_show_formulas_and_outline_symbols(self) -> None:
        with Workbook.create_default() as wb:
            view = wb.sheet(0).get_view()
            self.assertFalse(view.show_formulas)
            self.assertTrue(view.show_outline_symbols)
            wb.sheet(0).update_view(show_formulas=True)
            wb.sheet(0).update_view(show_outline_symbols=False)
            with Workbook.load(wb.save()) as reloaded:
                view = reloaded.sheet(0).get_view()
                self.assertTrue(view.show_formulas)
                self.assertFalse(view.show_outline_symbols)


class FontToggleTests(unittest.TestCase):
    def test_toggles_round_trip_and_dedup(self) -> None:
        record = Font(
            name="Arial",
            size_pt=12.0,
            color=Color(ColorKind.RGB, argb=0xFF000000),
            outline=True,
            shadow=False,
            condense=True,
            extend=True,
        )
        with Workbook.create_default() as wb:
            index = wb.ooxml.add_font(record)
            self.assertEqual(wb.ooxml.add_font(record), index)
            back = wb.ooxml.get_font(index)
            self.assertIs(back.outline, True)
            self.assertIs(back.shadow, False)
            self.assertTrue(back.condense and back.extend)
            self.assertIsNone(back.bold)


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
                    GradientStop(0.0, Color(kind=1, argb=0xFFFF0000)),
                    GradientStop(1.0, Color(kind=2, theme=4, tint=0.5)),
                ],
            )
        )

    def test_add_get_and_dedup(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.ooxml.add_fill(self._fill())
            self.assertEqual(wb.ooxml.add_fill(self._fill()), index)
            back = wb.ooxml.get_fill(index)
            self.assertIsNotNone(back.gradient)
            assert back.gradient is not None
            self.assertEqual(back.gradient.type, 1)
            self.assertEqual((back.gradient.left, back.gradient.top, back.gradient.bottom), (0.5, 0.25, 0.75))
            self.assertEqual(len(back.gradient.stops), 2)
            self.assertEqual(back.gradient.stops[0].color, Color(kind=1, argb=0xFFFF0000))
            self.assertEqual(back.gradient.stops[1].color.theme, 4)
            self.assertIsNone(wb.ooxml.get_fill(0).gradient)

    def test_survives_save_load_in_cells_and_dxfs(self) -> None:
        with Workbook.create_default() as wb:
            fill_index = wb.ooxml.add_fill(self._fill())
            dxf_index = wb.ooxml.add_dxf(DxfInput(fill=self._fill()))
            wb.sheet(0).add_cf_rule(CfExpressionRuleInput(ranges="A1", formula1="TRUE", dxf_index=dxf_index))
            with Workbook.load(wb.save()) as reloaded:
                gradient = reloaded.ooxml.get_fill(fill_index).gradient
                assert gradient is not None
                self.assertEqual(len(gradient.stops), 2)
                dxf_fill = reloaded.ooxml.get_dxf(dxf_index).fill
                assert dxf_fill is not None and dxf_fill.gradient is not None
                self.assertEqual(dxf_fill.gradient.type, 1)


class CustomIconTests(unittest.TestCase):
    @staticmethod
    def _rule(icons: list[CfIcon]) -> ConditionalFormatInput:
        return CfIconSetRuleInput(
            ranges=[CellRect(0, 0, 9, 0)],
            icon_set=CfIconSetInput(
                name=0,
                thresholds=[CfValueObject(1, "33"), CfValueObject(1, "67")],
                icons=icons,
            ),
        )

    def test_icons_round_trip_through_save_load(self) -> None:
        icons = [
            CfIcon(IconSetName.THREE_TRIANGLES, 0),
            CfIcon(IconSetName.NO_ICONS, 0),
            CfIcon(IconSetName.FIVE_BOXES, 4),
        ]
        with Workbook.create_default() as wb:
            wb.sheet(0).add_cf_rule(self._rule(icons))
            with Workbook.load(wb.save()) as reloaded:
                back = reloaded.sheet(0).list_cf_rules()[0].icon_set
                assert back is not None
                self.assertEqual(back.icons, icons)

    def test_rules_without_icons_report_none(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_cf_rule(self._rule([]))
            back = wb.sheet(0).list_cf_rules()[0].icon_set
            assert back is not None
            self.assertEqual(back.icons, [])

    def test_malformed_icon_lists_are_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_cf_rule(
                    self._rule([CfIcon(IconSetName.THREE_ARROWS, 0), CfIcon(IconSetName.THREE_ARROWS, 1)])
                )
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_cf_rule(
                    self._rule(
                        [CfIcon(IconSetName.THREE_ARROWS, 0), CfIcon(99, 1), CfIcon(IconSetName.THREE_ARROWS, 2)]
                    )
                )
            self.assertEqual(wb.sheet(0).list_cf_rules(), [])


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


class LocationFileNameTests(unittest.TestCase):
    def test_location_file_name_spells_and_binds_the_workbooks_own_names(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIsNone(wb.get_location())
            wb.set_defined_name(DefinedNameInput("LN", "10"))
            wb.set_location(WorkbookLocationInput("M5.xlsx"))
            self.assertEqual(wb.get_location(), WorkbookLocation(None, "M5.xlsx"))
            wb.sheet(0).set_formula(0, 0, "=M5.xlsx!LN+1")
            self.assertEqual(wb.sheet(0).get_formula(0, 0), "=M5.xlsx!LN+1")
            self.assertEqual(wb.sheet(0).list_cells()[0].formula, "=M5.xlsx!LN+1")
            wb.set_location(WorkbookLocationInput("1M5.xlsx"))
            self.assertEqual(wb.sheet(0).get_formula(0, 0), "='1M5.xlsx'!LN+1")
            with self.assertRaises(FormulonError):
                wb.set_location(WorkbookLocationInput("Book1"))
            self.assertEqual(wb.get_location(), WorkbookLocation(None, "1M5.xlsx"))
            wb.set_location(None)
            self.assertEqual(wb.sheet(0).get_formula(0, 0), "=[0]!LN+1")


class FormatDefaultFlagTests(unittest.TestCase):
    def test_custom_height_and_zero_height_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            defaults = wb.sheet(0).get_format_defaults()
            self.assertFalse(defaults.custom_height)
            self.assertFalse(defaults.zero_height)
            wb.sheet(0).update_format_defaults(default_row_height_pt=18.0, custom_height=True, zero_height=True)
            with Workbook.load(wb.save()) as reloaded:
                back = reloaded.sheet(0).get_format_defaults()
                self.assertTrue(back.custom_height)
                self.assertTrue(back.zero_height)
                self.assertEqual(back.default_row_height_pt, 18.0)
                reloaded.sheet(0).update_format_defaults(default_row_height_pt=20.0)
                kept = reloaded.sheet(0).get_format_defaults()
                self.assertTrue(kept.custom_height)
                self.assertTrue(kept.zero_height)
                self.assertEqual(kept.default_row_height_pt, 20.0)
                reloaded.sheet(0).update_format_defaults(custom_height=False, zero_height=False)
                cleared = reloaded.sheet(0).get_format_defaults()
                self.assertFalse(cleared.custom_height)
                self.assertFalse(cleared.zero_height)


class DefinedNameAttrTests(unittest.TestCase):
    def test_comment_and_hidden(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_defined_name(DefinedNameInput("Rate", "Sheet1!$A$1", comment="Tax rate", hidden=True))
            entry = wb.list_defined_names()[0]
            self.assertEqual((entry.comment, entry.hidden), ("Tax rate", True))

            # A replacement is whole: omitted attributes return to their defaults.
            wb.set_defined_name(DefinedNameInput("Rate", "Sheet1!$B$2"))
            entry = wb.list_defined_names()[0]
            self.assertEqual((entry.comment, entry.hidden), (None, False))

            wb.set_defined_name(DefinedNameInput("Rate", "Sheet1!$B$2", hidden=True))
            with Workbook.load(wb.save()) as reloaded:
                entry = reloaded.list_defined_names()[0]
                self.assertEqual((entry.formula, entry.comment, entry.hidden), ("=Sheet1!$B$2", None, True))

    def test_empty_formula_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(InvalidArgumentError):
                wb.set_defined_name(DefinedNameInput("Nope", ""))
            self.assertEqual(wb.list_defined_names(), [])


if __name__ == "__main__":
    unittest.main()
