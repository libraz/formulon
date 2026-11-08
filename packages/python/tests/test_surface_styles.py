"""Style surface tests."""

from __future__ import annotations

import unittest
from dataclasses import replace

from _xf_input import input_of

from formulon import (
    Alignment,
    Border,
    BorderSide,
    CellProtection,
    CellXfInput,
    Color,
    ColorKind,
    DxfInput,
    Fill,
    Font,
    HorizontalAlign,
    VerticalAlign,
    Workbook,
)


class StyleTests(unittest.TestCase):
    def test_font_numfmt_xf_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            fi = wb.ooxml.add_font(Font(name="Calibri", size_pt=12.0, bold=True))
            self.assertGreaterEqual(fi, 0)
            font = wb.ooxml.get_font(fi)
            self.assertEqual(font.name, "Calibri")
            self.assertTrue(font.bold)

            fill = wb.ooxml.add_fill(Fill(pattern=1, foreground=Color(ColorKind.RGB, argb=0xFFFF0000)))
            nf = wb.ooxml.add_number_format("0.00")
            self.assertGreater(nf, 0)
            self.assertEqual(wb.ooxml.get_number_format(nf), "0.00")

            # add_cell_xf validates indices against the parallel tables, so a
            # border must exist before it can be referenced.
            border = wb.ooxml.add_border(Border(left=BorderSide(style=1, color=Color(ColorKind.RGB, argb=0xFF000000))))

            xf = wb.ooxml.add_cell_xf(
                CellXfInput(
                    font_index=fi,
                    fill_index=fill,
                    border_index=border,
                    num_fmt_id=nf,
                    horizontal=None,
                    vertical=VerticalAlign.TOP,
                    wrap_text=None,
                    text_rotation=255,
                    indent=0,
                    relative_indent=-3,
                    shrink_to_fit=False,
                    reading_order=0,
                    justify_last_line=True,
                )
            )
            wb.sheet(0).ooxml.set_cell_xf_index(0, 0, xf)
            self.assertEqual(wb.sheet(0).ooxml.get_cell_xf_index(0, 0), xf)
            resolved = wb.ooxml.get_cell_xf(xf)
            self.assertEqual(resolved.font_index, fi)
            self.assertEqual(resolved.num_fmt_id, nf)
            self.assertEqual(resolved.text_rotation, 255)
            self.assertEqual(resolved.indent, 0)
            self.assertEqual(resolved.relative_indent, -3)
            self.assertIs(resolved.shrink_to_fit, False)
            self.assertEqual(resolved.reading_order, 0)
            self.assertIs(resolved.justify_last_line, True)
            self.assertIsNone(resolved.horizontal)
            self.assertEqual(resolved.vertical, 0)
            self.assertIsNone(resolved.wrap_text)
            self.assertEqual(wb.ooxml.add_cell_xf(input_of(resolved)), xf)

            omitted = wb.ooxml.add_cell_xf(
                CellXfInput(
                    font_index=fi,
                    fill_index=fill,
                    border_index=border,
                    num_fmt_id=nf,
                    horizontal=None,
                    vertical=None,
                    wrap_text=None,
                )
            )
            got = wb.ooxml.get_cell_xf(omitted)
            self.assertEqual((got.horizontal, got.vertical, got.wrap_text), (None, None, None))

            stated = wb.ooxml.add_cell_xf(
                CellXfInput(
                    font_index=fi,
                    fill_index=fill,
                    border_index=border,
                    num_fmt_id=nf,
                    horizontal=0,
                    vertical=2,
                    wrap_text=False,
                )
            )
            self.assertNotEqual(stated, omitted)
            got = wb.ooxml.get_cell_xf(stated)
            self.assertEqual((got.horizontal, got.vertical, got.wrap_text), (0, 2, False))

    def test_font_vert_align_roundtrip_is_the_identity(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.ooxml.add_font(
                Font(name="Arial", size_pt=12.0, vert_align=1, color=Color(ColorKind.RGB, argb=0xFF112233))
            )
            self.assertEqual(wb.ooxml.get_font(index).vert_align, 1)

            before = len(wb.ooxml.list_fonts())
            self.assertEqual(wb.ooxml.add_font(wb.ooxml.get_font(index)), index)
            self.assertEqual(len(wb.ooxml.list_fonts()), before)

    def test_font_one_field_rewrite_preserves_the_superscript(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.ooxml.add_font(
                Font(name="Arial", size_pt=12.0, vert_align=1, color=Color(ColorKind.RGB, argb=0xFF112233))
            )
            edited = wb.ooxml.get_font(index)
            edited = replace(edited, color=Color(ColorKind.RGB, argb=0xFF00FF00))
            recolored = wb.ooxml.add_font(edited)
            self.assertNotEqual(recolored, index)
            reread = wb.ooxml.get_font(recolored)
            self.assertEqual(reread.vert_align, 1)
            self.assertEqual(reread.color, Color(ColorKind.RGB, argb=0xFF00FF00))

    def test_dxf_font_vert_align_roundtrip_is_the_identity(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.ooxml.add_dxf(DxfInput(font=Font(name="Calibri", size_pt=9.0, vert_align=1)))
            got = wb.ooxml.get_dxf(index)
            self.assertIsNotNone(got.font)
            self.assertEqual(got.font.vert_align, 1)

            before = len(wb.ooxml.list_dxfs())
            self.assertEqual(wb.ooxml.add_dxf(got), index)
            self.assertEqual(len(wb.ooxml.list_dxfs()), before)

    def test_dxf_font_without_size_reads_back_without_size(self) -> None:
        with Workbook.create_default() as wb:
            bare = wb.ooxml.add_dxf(DxfInput(font=Font(bold=True, size_pt=None)))
            sized = wb.ooxml.add_dxf(DxfInput(font=Font(bold=True, size_pt=13.0)))
            self.assertNotEqual(bare, sized)
            self.assertIsNone(wb.ooxml.get_dxf(bare).font.size_pt)
            self.assertEqual(wb.ooxml.get_dxf(sized).font.size_pt, 13.0)
            self.assertEqual(wb.ooxml.get_font(wb.ooxml.add_font(Font(name="Arial", size_pt=14.0))).size_pt, 14.0)

    def test_dxf_roundtrip_and_dedup(self) -> None:
        with Workbook.create_default() as wb:
            record = DxfInput(
                font=Font(name="Arial", size_pt=12.0, bold=True, color=Color(ColorKind.RGB, argb=0xFFFF0000)),
                fill=Fill(pattern=1, foreground=Color(ColorKind.RGB, argb=0xFFFFFF00)),
                number_format="0.00",
            )
            index = wb.ooxml.add_dxf(record)
            self.assertEqual(wb.ooxml.add_dxf(record), index)
            self.assertEqual(len(wb.ooxml.list_dxfs()), 1)
            got = wb.ooxml.get_dxf(index)
            self.assertIsNotNone(got.font)
            self.assertEqual(got.font.name, "Arial")
            self.assertTrue(got.font.bold)
            self.assertIsNotNone(got.fill)
            self.assertEqual(got.fill.foreground, Color(ColorKind.RGB, argb=0xFFFFFF00))
            self.assertEqual(got.number_format, "0.00")

    def test_dxf_alignment_and_protection_survive_get_add_and_save_load(self) -> None:
        alignment = Alignment(horizontal=HorizontalAlign.CENTER, wrap_text=True)
        protection = CellProtection(locked=False, hidden=True)
        with Workbook.create_default() as wb:
            alignment_index = wb.ooxml.add_dxf(DxfInput(alignment=alignment))
            protection_index = wb.ooxml.add_dxf(DxfInput(protection=protection))
            self.assertNotEqual(alignment_index, protection_index)

            got_alignment = wb.ooxml.get_dxf(alignment_index)
            self.assertEqual(got_alignment.alignment.horizontal, HorizontalAlign.CENTER)
            self.assertIs(got_alignment.alignment.wrap_text, True)
            self.assertIsNone(got_alignment.protection)
            got_protection = wb.ooxml.get_dxf(protection_index)
            self.assertIsNone(got_protection.alignment)
            self.assertEqual(got_protection.protection, protection)
            self.assertEqual(wb.ooxml.add_dxf(got_alignment), alignment_index)
            self.assertEqual(wb.ooxml.add_dxf(got_protection), protection_index)

            with Workbook.load(wb.save()) as reloaded:
                reloaded_alignment = reloaded.ooxml.get_dxf(alignment_index)
                reloaded_protection = reloaded.ooxml.get_dxf(protection_index)
                self.assertEqual(reloaded_alignment.alignment, got_alignment.alignment)
                self.assertIsNone(reloaded_alignment.protection)
                self.assertIsNone(reloaded_protection.alignment)
                self.assertEqual(reloaded_protection.protection, protection)
                self.assertEqual(reloaded.ooxml.add_dxf(reloaded_alignment), alignment_index)
                self.assertEqual(reloaded.ooxml.add_dxf(reloaded_protection), protection_index)

    def test_selector_colours_survive_get_add_identity_and_save_load(self) -> None:
        theme = Color(kind=2, theme=3, tint=0.5)
        indexed = Color(kind=3, indexed=9)
        automatic = Color(kind=4)
        with Workbook.create_default() as wb:
            font_index = wb.ooxml.add_font(Font(name="SelectorFont", size_pt=11.0, color=theme))
            got_font = wb.ooxml.get_font(font_index)
            self.assertEqual(got_font.color, theme)
            self.assertEqual(wb.ooxml.add_font(got_font), font_index)

            fill_index = wb.ooxml.add_fill(
                Fill(
                    pattern=1,
                    foreground=indexed,
                    background=automatic,
                )
            )
            got_fill = wb.ooxml.get_fill(fill_index)
            self.assertEqual(got_fill.foreground, indexed)
            self.assertEqual(got_fill.background, automatic)
            self.assertEqual(wb.ooxml.add_fill(got_fill), fill_index)

            border = Border(
                left=BorderSide(style=1, color=theme),
                right=BorderSide(style=1, color=indexed),
                top=BorderSide(style=1, color=automatic),
            )
            border_index = wb.ooxml.add_border(border)
            got_border = wb.ooxml.get_border(border_index)
            self.assertEqual(got_border.left.color, theme)
            self.assertEqual(got_border.right.color, indexed)
            self.assertEqual(got_border.top.color, automatic)
            self.assertEqual(wb.ooxml.add_border(got_border), border_index)

            dxf_index = wb.ooxml.add_dxf(
                DxfInput(
                    font=Font(name="DxfSelector", size_pt=9.0, color=automatic),
                    fill=Fill(pattern=1, foreground=indexed),
                    border=Border(left=BorderSide(style=1, color=theme)),
                )
            )
            got_dxf = wb.ooxml.get_dxf(dxf_index)
            self.assertEqual(got_dxf.font.color, automatic)
            self.assertEqual(got_dxf.fill.foreground, indexed)
            self.assertEqual(got_dxf.border.left.color, theme)
            self.assertEqual(wb.ooxml.add_dxf(got_dxf), dxf_index)

            with Workbook.load(wb.save()) as reloaded:
                self.assertEqual(reloaded.ooxml.get_font(font_index).color, theme)
                self.assertEqual(reloaded.ooxml.get_fill(fill_index).foreground, indexed)
                self.assertEqual(reloaded.ooxml.get_border(border_index).left.color, theme)
                self.assertEqual(reloaded.ooxml.get_dxf(dxf_index).font.color, automatic)


if __name__ == "__main__":
    unittest.main()
