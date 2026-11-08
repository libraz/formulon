"""Theme, colour resolution, effective style and named-style removal surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    CellStyle,
    CellXfInput,
    Color,
    ColorContext,
    ColorKind,
    ColorResolution,
    EffectiveStyleSource,
    Font,
    FormulonError,
    ThemeFonts,
    ThemeScriptFont,
    ThemeSource,
    Workbook,
)

_THEME_KIND = 2
_RGB_KIND = 1
_AUTO_KIND = 4


def _xf(font_index: int) -> CellXfInput:
    return CellXfInput(
        font_index=font_index,
        fill_index=0,
        border_index=0,
        num_fmt_id=0,
        horizontal=0,
        vertical=2,
        wrap_text=False,
    )


class ThemeTests(unittest.TestCase):
    def test_get_theme_reports_twelve_colors_and_fonts(self) -> None:
        with Workbook.create_default() as wb:
            colors = wb.get_theme_colors()
            fonts = wb.get_theme_fonts()
        self.assertIn(colors.source, (ThemeSource.PART, ThemeSource.DEFAULT))
        self.assertEqual(len(colors.colors), 12)
        self.assertTrue(fonts.minor_latin)

    def test_set_theme_colors_roundtrip(self) -> None:
        colors = [0xFF000000 | (0x112233 * (i + 1) & 0xFFFFFF) for i in range(12)]
        with Workbook.create_default() as wb:
            wb.set_theme_colors(colors=colors)
            theme = wb.get_theme_colors()
            with self.assertRaises(ValueError):
                wb.set_theme_colors(colors=colors[:11])
        self.assertEqual(theme.source, ThemeSource.PART)
        self.assertEqual(theme.colors, colors)

    def test_set_theme_fonts_roundtrip(self) -> None:
        fonts = ThemeFonts("Arial", "MS Gothic", "Verdana", "Meiryo")
        with Workbook.create_default() as wb:
            wb.set_theme_fonts(fonts)
            self.assertEqual(wb.get_theme_fonts(), fonts)

    def test_reset_theme_returns_to_default(self) -> None:
        colors = [0xFF000000 | (0x112233 * (i + 1) & 0xFFFFFF) for i in range(12)]
        with Workbook.create_default() as wb:
            wb.reset_theme()
            self.assertEqual(wb.get_theme_colors().source, ThemeSource.DEFAULT)
            wb.set_theme_colors(colors=colors)
            self.assertEqual(wb.get_theme_colors().source, ThemeSource.PART)
            wb.reset_theme()
            theme = wb.get_theme_colors()
            self.assertEqual(theme.source, ThemeSource.DEFAULT)
            self.assertNotEqual(theme.colors, colors)
            self.assertNotIn(b"xl/theme/theme1.xml", wb.save())

    def test_resolve_color(self) -> None:
        with Workbook.create_default() as wb:
            rgb = wb.resolve_color(Color(kind=_RGB_KIND, argb=0xFF123456))
            self.assertEqual(rgb.argb, 0xFF123456)
            self.assertEqual(rgb.resolution, ColorResolution.EXACT)

            colors = [0xFF000000 + i * 0x010101 for i in range(12)]
            wb.set_theme_colors(colors=colors)
            # Theme index 1 is dk1, which is clrScheme slot 0.
            theme = wb.resolve_color(Color(kind=_THEME_KIND, theme=1))
            self.assertEqual(theme.argb, colors[0])
            self.assertEqual(theme.resolution, ColorResolution.EXACT)

            auto_font = wb.resolve_color(Color(kind=_AUTO_KIND), ColorContext.FONT)
            auto_fill = wb.resolve_color(Color(kind=_AUTO_KIND), ColorContext.FILL_FOREGROUND)
            self.assertEqual(auto_font.argb, 0xFF000000)
            self.assertEqual(auto_fill.argb, 0xFFFFFFFF)
            self.assertEqual(auto_font.resolution, ColorResolution.AUTO_CONTEXT)

            with self.assertRaises(FormulonError):
                wb.resolve_color(Color(kind=_RGB_KIND), 99)


class EffectiveStyleTests(unittest.TestCase):
    def test_default_cell_reads_default_xf(self) -> None:
        with Workbook.create_default() as wb:
            style = wb.sheet(0).ooxml.get_effective_style(4, 4)
        self.assertEqual(style.xf_index, 0)
        self.assertEqual(style.source, EffectiveStyleSource.DEFAULT)
        self.assertEqual(len(style.borders), 5)
        self.assertTrue(style.protection.locked)
        self.assertEqual(style.number_format, "General")

    def test_cell_xf_is_reported(self) -> None:
        with Workbook.create_default() as wb:
            font = wb.ooxml.add_font(
                Font(name="Calibri", size_pt=11.0, bold=True, color=Color(ColorKind.RGB, argb=0xFFFF0000))
            )
            xf = wb.ooxml.add_cell_xf(_xf(font))
            wb.sheet(0).ooxml.set_cell_xf_index(1, 1, xf)
            style = wb.sheet(0).ooxml.get_effective_style(1, 1)
            with self.assertRaises(FormulonError):
                wb.sheet(99).ooxml.get_effective_style(0, 0)
        self.assertEqual(style.xf_index, xf)
        self.assertEqual(style.source, EffectiveStyleSource.CELL)
        self.assertEqual(style.font_index, font)


class ClearCellXfIndexTests(unittest.TestCase):
    def test_clear_restores_inheritance(self) -> None:
        with Workbook.create_default() as wb:
            font = wb.ooxml.add_font(
                Font(name="Calibri", size_pt=11.0, bold=True, color=Color(ColorKind.RGB, argb=0xFFFF0000))
            )
            col_xf = wb.ooxml.add_cell_xf(_xf(font))
            cell_xf = wb.ooxml.add_cell_xf(_xf(font))
            wb.sheet(0).ooxml.set_column_xf_index(0, 0, col_xf)

            wb.sheet(0).ooxml.set_cell_xf_index(0, 0, cell_xf)
            style = wb.sheet(0).ooxml.get_effective_style(0, 0)
            self.assertEqual((style.source, style.xf_index), (EffectiveStyleSource.CELL, cell_xf))

            wb.sheet(0).ooxml.reset_cell_xf_index(0, 0)
            style = wb.sheet(0).ooxml.get_effective_style(0, 0)
            self.assertEqual((style.source, style.xf_index), (EffectiveStyleSource.COLUMN, col_xf))

            wb.sheet(0).ooxml.set_cell_xf_index(0, 0, 0)
            style = wb.sheet(0).ooxml.get_effective_style(0, 0)
            self.assertEqual((style.source, style.xf_index), (EffectiveStyleSource.CELL, 0))

    def test_absent_cell_is_noop(self) -> None:
        with Workbook.create_default() as wb:
            font = wb.ooxml.add_font(
                Font(name="Calibri", size_pt=11.0, bold=True, color=Color(ColorKind.RGB, argb=0xFFFF0000))
            )
            col_xf = wb.ooxml.add_cell_xf(_xf(font))
            wb.sheet(0).ooxml.set_column_xf_index(0, 0, col_xf)
            wb.sheet(0).ooxml.reset_cell_xf_index(5, 0)
            wb.sheet(0).ooxml.reset_cell_xf_index(0, 0)
            style = wb.sheet(0).ooxml.get_effective_style(0, 0)
            self.assertEqual((style.source, style.xf_index), (EffectiveStyleSource.COLUMN, col_xf))
            with self.assertRaises(FormulonError):
                wb.sheet(99).ooxml.reset_cell_xf_index(0, 0)

    def test_value_is_kept(self) -> None:
        with Workbook.create_default() as wb:
            font = wb.ooxml.add_font(
                Font(name="Calibri", size_pt=11.0, bold=True, color=Color(ColorKind.RGB, argb=0xFFFF0000))
            )
            xf = wb.ooxml.add_cell_xf(_xf(font))
            wb.sheet(0).set_number(0, 0, 7.0)
            wb.sheet(0).ooxml.set_cell_xf_index(0, 0, xf)
            wb.sheet(0).ooxml.reset_cell_xf_index(0, 0)
            self.assertEqual(wb.sheet(0).get_value(0, 0), 7.0)
            self.assertEqual(wb.sheet(0).ooxml.get_cell_xf_index(0, 0), 0)


class NamedStyleTests(unittest.TestCase):
    def test_set_and_remove_cell_style(self) -> None:
        with Workbook.create_default() as wb:
            style_xf_index = wb.ooxml.add_cell_style_xf(_xf(0))
            wb.set_cell_style(CellStyle("Mine", style_xf_index, hidden=True))
            names = {style.name: style for style in wb.list_cell_styles()}
            self.assertEqual(wb.get_cell_style("Mine"), names["Mine"])
            self.assertIn("Mine", names)
            self.assertEqual(names["Mine"].style_xf_index, style_xf_index)
            self.assertTrue(names["Mine"].hidden)
            wb.remove_cell_style("Mine")
            after = [style.name for style in wb.list_cell_styles()]
            self.assertIsNone(wb.get_cell_style("Mine"))
            self.assertNotIn("Mine", after)
            with self.assertRaises(FormulonError):
                wb.remove_cell_style("Mine")
            with self.assertRaises(FormulonError):
                wb.set_cell_style(CellStyle("", style_xf_index))


class CreateForProfileTests(unittest.TestCase):
    def test_mac_ja_jp_defaults(self) -> None:
        with Workbook.create_for_profile("mac-365-ja_JP") as wb:
            self.assertEqual(wb.get_profile_id(), "mac-365-ja_JP")
            body = wb.ooxml.get_font(0)
            self.assertEqual(body.name, "游ゴシック")
            self.assertEqual(body.size_pt, 12.0)
            self.assertEqual(body.charset, 128)
            self.assertEqual(body.scheme, 2)
            self.assertEqual(wb.ooxml.get_font(1).size_pt, 6.0)

    def test_unknown_profile_raises(self) -> None:
        with self.assertRaises(FormulonError):
            Workbook.create_for_profile("mac-365-xx_XX")


class ThemeScriptFontTests(unittest.TestCase):
    def test_default_theme_reports_its_scripts(self) -> None:
        with Workbook.create_for_profile("mac-365-ja_JP") as wb:
            jpan = wb.get_theme_script_font(major=False, script="Jpan")
            hang = wb.get_theme_script_font(major=False, script="Hang")
            missing = wb.get_theme_script_font(major=True, script="Zzzz")
        self.assertEqual(jpan, ThemeScriptFont(major=False, script="Jpan", typeface="游ゴシック"))
        self.assertEqual(hang, ThemeScriptFont(major=False, script="Hang", typeface="맑은 고딕"))
        self.assertIsNone(missing)

    def test_set_script_font_roundtrip(self) -> None:
        with Workbook.create_for_profile("mac-365-en_US") as wb:
            jpan = wb.get_theme_script_font(major=False, script="Jpan")
            wb.set_theme_script_font(major=False, script="Hang", typeface="Test Face")
            data = wb.save()
        with Workbook.load(data) as wb2:
            self.assertEqual(
                wb2.get_theme_script_font(major=False, script="Hang"),
                ThemeScriptFont(major=False, script="Hang", typeface="Test Face"),
            )
            self.assertEqual(wb2.get_theme_script_font(major=False, script="Jpan"), jpan)

    def test_invalid_script_raises(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.set_theme_script_font(major=True, script="Jp", typeface="Face")
            with self.assertRaises(FormulonError):
                wb.set_theme_script_font(major=True, script="Jpan", typeface="")


if __name__ == "__main__":
    unittest.main()
