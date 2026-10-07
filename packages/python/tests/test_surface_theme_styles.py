"""Theme, colour resolution, effective style and named-style removal surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    CellStyle,
    CellXf,
    ColorContext,
    ColorResolution,
    ColorSpec,
    EffectiveStyleSource,
    FontRecord,
    FormulonError,
    ThemeFonts,
    ThemeSource,
    Workbook,
)

_THEME_KIND = 2
_RGB_KIND = 1
_AUTO_KIND = 4


def _xf(font_index: int) -> CellXf:
    return CellXf(
        font_index=font_index,
        fill_index=0,
        border_index=0,
        num_fmt_id=0,
        horizontal_align=0,
        vertical_align=2,
        wrap_text=False,
    )


class ThemeTests(unittest.TestCase):
    def test_get_theme_reports_twelve_colors_and_fonts(self) -> None:
        with Workbook.create_default() as wb:
            theme = wb.get_theme()
        self.assertIn(theme.source, (ThemeSource.PART, ThemeSource.DEFAULT))
        self.assertEqual(len(theme.colors), 12)
        self.assertTrue(theme.fonts.minor_latin)

    def test_set_theme_colors_roundtrip(self) -> None:
        colors = [0xFF000000 | (0x112233 * (i + 1) & 0xFFFFFF) for i in range(12)]
        with Workbook.create_default() as wb:
            wb.set_theme_colors(colors)
            theme = wb.get_theme()
            with self.assertRaises(ValueError):
                wb.set_theme_colors(colors[:11])
        self.assertEqual(theme.source, ThemeSource.PART)
        self.assertEqual(theme.colors, colors)

    def test_set_theme_fonts_roundtrip(self) -> None:
        fonts = ThemeFonts("Arial", "MS Gothic", "Verdana", "Meiryo")
        with Workbook.create_default() as wb:
            wb.set_theme_fonts(fonts)
            self.assertEqual(wb.get_theme().fonts, fonts)

    def test_reset_theme_returns_to_default(self) -> None:
        colors = [0xFF000000 | (0x112233 * (i + 1) & 0xFFFFFF) for i in range(12)]
        with Workbook.create_default() as wb:
            wb.reset_theme()
            self.assertEqual(wb.get_theme().source, ThemeSource.DEFAULT)
            wb.set_theme_colors(colors)
            self.assertEqual(wb.get_theme().source, ThemeSource.PART)
            wb.reset_theme()
            theme = wb.get_theme()
            self.assertEqual(theme.source, ThemeSource.DEFAULT)
            self.assertNotEqual(theme.colors, colors)
            self.assertNotIn(b"xl/theme/theme1.xml", wb.save())

    def test_resolve_color(self) -> None:
        with Workbook.create_default() as wb:
            rgb = wb.resolve_color(ColorSpec(kind=_RGB_KIND, rgb=0xFF123456))
            self.assertEqual(rgb.argb, 0xFF123456)
            self.assertEqual(rgb.resolution, ColorResolution.EXACT)

            colors = [0xFF000000 + i * 0x010101 for i in range(12)]
            wb.set_theme_colors(colors)
            # Theme index 1 is dk1, which is clrScheme slot 0.
            theme = wb.resolve_color(ColorSpec(kind=_THEME_KIND, theme=1))
            self.assertEqual(theme.argb, colors[0])
            self.assertEqual(theme.resolution, ColorResolution.EXACT)

            auto_font = wb.resolve_color(ColorSpec(kind=_AUTO_KIND), ColorContext.FONT)
            auto_fill = wb.resolve_color(ColorSpec(kind=_AUTO_KIND), ColorContext.FILL_FOREGROUND)
            self.assertEqual(auto_font.argb, 0xFF000000)
            self.assertEqual(auto_fill.argb, 0xFFFFFFFF)
            self.assertEqual(auto_font.resolution, ColorResolution.AUTO_CONTEXT)

            with self.assertRaises(FormulonError):
                wb.resolve_color(ColorSpec(kind=_RGB_KIND), 99)


class EffectiveStyleTests(unittest.TestCase):
    def test_default_cell_reads_default_xf(self) -> None:
        with Workbook.create_default() as wb:
            style = wb.get_effective_style(0, 4, 4)
        self.assertEqual(style.xf_index, 0)
        self.assertEqual(style.source, EffectiveStyleSource.DEFAULT)
        self.assertEqual(len(style.borders), 5)
        self.assertTrue(style.locked)
        self.assertEqual(style.num_fmt_code, "General")

    def test_cell_xf_is_reported(self) -> None:
        with Workbook.create_default() as wb:
            font = wb.add_font(FontRecord(name="Calibri", size=11.0, bold=True, color_argb=0xFFFF0000))
            xf = wb.add_cell_xf(_xf(font))
            wb.set_cell_xf_index(0, 1, 1, xf)
            style = wb.get_effective_style(0, 1, 1)
            with self.assertRaises(FormulonError):
                wb.get_effective_style(99, 0, 0)
        self.assertEqual(style.xf_index, xf)
        self.assertEqual(style.source, EffectiveStyleSource.CELL)
        self.assertEqual(style.font_index, font)


class NamedStyleTests(unittest.TestCase):
    def test_set_and_remove_cell_style(self) -> None:
        with Workbook.create_default() as wb:
            xf_id = wb.add_cell_style_xf(_xf(0))
            wb.set_cell_style(CellStyle("Mine", xf_id, hidden=True))
            names = {wb.get_cell_style(i).name: wb.get_cell_style(i) for i in range(wb.cell_style_count())}
            self.assertIn("Mine", names)
            self.assertEqual(names["Mine"].xf_id, xf_id)
            self.assertTrue(names["Mine"].hidden)
            wb.remove_cell_style("Mine")
            after = [wb.get_cell_style(i).name for i in range(wb.cell_style_count())]
            self.assertNotIn("Mine", after)
            with self.assertRaises(FormulonError):
                wb.remove_cell_style("Mine")
            with self.assertRaises(FormulonError):
                wb.set_cell_style(CellStyle("", xf_id))


if __name__ == "__main__":
    unittest.main()
