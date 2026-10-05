"""Style surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    ColorSpec,
    DifferentialFormat,
    FillRecord,
    FontRecord,
    Workbook,
)


class StyleTests(unittest.TestCase):
    def test_font_numfmt_xf_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            fi = wb.add_font(FontRecord(name="Calibri", size=12.0, bold=True))
            self.assertGreaterEqual(fi, 0)
            font = wb.get_font(fi)
            self.assertEqual(font.name, "Calibri")
            self.assertTrue(font.bold)

            fill = wb.add_fill(FillRecord(pattern=1, fg_argb=0xFFFF0000))
            nf = wb.add_num_fmt("0.00")
            self.assertGreater(nf, 0)
            self.assertEqual(wb.get_num_fmt(nf), "0.00")

            # add_cell_xf validates indices against the parallel tables, so a
            # border must exist before it can be referenced.
            border = wb.add_border({"left": {"style": 1, "color_argb": 0xFF000000}})

            from formulon import CellXf

            xf = wb.add_cell_xf(
                CellXf(
                    font_index=fi,
                    fill_index=fill,
                    border_index=border,
                    num_fmt_id=nf,
                    horizontal_align=0,
                    vertical_align=0,
                    wrap_text=False,
                    text_rotation=255,
                    indent=0,
                    relative_indent=-3,
                    shrink_to_fit=False,
                    reading_order=0,
                    justify_last_line=True,
                )
            )
            wb.set_cell_xf_index(0, 0, 0, xf)
            self.assertEqual(wb.get_cell_xf_index(0, 0, 0), xf)
            resolved = wb.get_cell_xf(xf)
            self.assertEqual(resolved.font_index, fi)
            self.assertEqual(resolved.num_fmt_id, nf)
            self.assertEqual(resolved.text_rotation, 255)
            self.assertTrue(resolved.has_alignment)
            self.assertEqual(resolved.indent, 0)
            self.assertEqual(resolved.relative_indent, -3)
            self.assertIs(resolved.shrink_to_fit, False)
            self.assertEqual(resolved.reading_order, 0)
            self.assertIs(resolved.justify_last_line, True)
            self.assertFalse(resolved.has_horizontal_align)
            self.assertTrue(resolved.has_vertical_align)
            self.assertFalse(resolved.has_wrap_text)
            self.assertTrue(resolved.has_justify_last_line)
            self.assertEqual(wb.add_cell_xf(resolved), xf)

            omitted = wb.add_cell_xf(
                CellXf(
                    font_index=fi,
                    fill_index=fill,
                    border_index=border,
                    num_fmt_id=nf,
                    horizontal_align=0,
                    vertical_align=2,
                    wrap_text=False,
                )
            )
            self.assertFalse(wb.get_cell_xf(omitted).has_alignment)

            explicit_empty = wb.add_cell_xf(
                CellXf(
                    font_index=fi,
                    fill_index=fill,
                    border_index=border,
                    num_fmt_id=nf,
                    horizontal_align=0,
                    vertical_align=2,
                    wrap_text=False,
                    has_alignment=True,
                )
            )
            self.assertNotEqual(explicit_empty, omitted)
            self.assertTrue(wb.get_cell_xf(explicit_empty).has_alignment)

    def test_font_vert_align_roundtrip_is_the_identity(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.add_font(FontRecord(name="Arial", size=12.0, vert_align=1, color_argb=0xFF112233))
            self.assertEqual(wb.get_font(index).vert_align, 1)

            before = wb.font_count()
            self.assertEqual(wb.add_font(wb.get_font(index)), index)
            self.assertEqual(wb.font_count(), before)

    def test_font_one_field_rewrite_preserves_the_superscript(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.add_font(FontRecord(name="Arial", size=12.0, vert_align=1, color_argb=0xFF112233))
            edited = wb.get_font(index)
            edited.color_argb = 0xFF00FF00
            recolored = wb.add_font(edited)
            self.assertNotEqual(recolored, index)
            reread = wb.get_font(recolored)
            self.assertEqual(reread.vert_align, 1)
            self.assertEqual(reread.color_argb, 0xFF00FF00)

    def test_dxf_font_vert_align_roundtrip_is_the_identity(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.add_dxf(DifferentialFormat(font=FontRecord(name="Calibri", size=9.0, vert_align=1)))
            got = wb.get_dxf(index)
            self.assertIsNotNone(got.font)
            self.assertEqual(got.font.vert_align, 1)

            before = wb.dxf_count()
            self.assertEqual(wb.add_dxf(got), index)
            self.assertEqual(wb.dxf_count(), before)

    def test_dxf_roundtrip_and_dedup(self) -> None:
        with Workbook.create_default() as wb:
            record = DifferentialFormat(
                font=FontRecord(name="Arial", size=12.0, bold=True, color_argb=0xFFFF0000),
                fill=FillRecord(pattern=1, fg_argb=0xFFFFFF00),
                num_fmt_id=164,
                num_fmt_code="0.00",
            )
            index = wb.add_dxf(record)
            self.assertEqual(wb.add_dxf(record), index)
            self.assertEqual(wb.dxf_count(), 1)
            got = wb.get_dxf(index)
            self.assertIsNotNone(got.font)
            self.assertEqual(got.font.name, "Arial")
            self.assertTrue(got.font.bold)
            self.assertIsNotNone(got.fill)
            self.assertEqual(got.fill.fg_argb, 0xFFFFFF00)
            self.assertEqual(got.num_fmt_id, 164)
            self.assertEqual(got.num_fmt_code, "0.00")

    def test_dxf_alignment_and_protection_xml_survive_get_add_and_save_load(self) -> None:
        alignment_xml = '<alignment horizontal="center" wrapText="1"/>'
        protection_xml = '<protection locked="0" hidden="1"/>'
        with Workbook.create_default() as wb:
            alignment_index = wb.add_dxf(DifferentialFormat(alignment_xml=alignment_xml))
            protection_index = wb.add_dxf(DifferentialFormat(protection_xml=protection_xml))
            self.assertNotEqual(alignment_index, protection_index)

            got_alignment = wb.get_dxf(alignment_index)
            self.assertEqual(got_alignment.alignment_xml, alignment_xml)
            self.assertEqual(got_alignment.protection_xml, "")
            got_protection = wb.get_dxf(protection_index)
            self.assertEqual(got_protection.alignment_xml, "")
            self.assertEqual(got_protection.protection_xml, protection_xml)
            self.assertEqual(wb.add_dxf(got_alignment), alignment_index)
            self.assertEqual(wb.add_dxf(got_protection), protection_index)

            with Workbook.load(wb.save()) as reloaded:
                reloaded_alignment = reloaded.get_dxf(alignment_index)
                reloaded_protection = reloaded.get_dxf(protection_index)
                self.assertEqual(reloaded_alignment.alignment_xml, alignment_xml)
                self.assertEqual(reloaded_alignment.protection_xml, "")
                self.assertEqual(reloaded_protection.alignment_xml, "")
                self.assertEqual(reloaded_protection.protection_xml, protection_xml)
                self.assertEqual(reloaded.add_dxf(reloaded_alignment), alignment_index)
                self.assertEqual(reloaded.add_dxf(reloaded_protection), protection_index)

    def test_selector_colours_survive_get_add_identity_and_save_load(self) -> None:
        theme = ColorSpec(kind=2, theme=3, tint=0.5)
        indexed = ColorSpec(kind=3, indexed=9)
        automatic = ColorSpec(kind=4)
        with Workbook.create_default() as wb:
            font_index = wb.add_font(FontRecord(name="SelectorFont", size=11.0, color_argb=0x01020304, color=theme))
            got_font = wb.get_font(font_index)
            self.assertEqual(got_font.color, theme)
            self.assertEqual(got_font.color_argb, 0x01020304)
            self.assertEqual(wb.add_font(got_font), font_index)

            fill_index = wb.add_fill(
                FillRecord(
                    pattern=1,
                    fg_argb=0x05060708,
                    bg_argb=0x090A0B0C,
                    fg=indexed,
                    bg=automatic,
                )
            )
            got_fill = wb.get_fill(fill_index)
            self.assertEqual(got_fill.fg, indexed)
            self.assertEqual(got_fill.bg, automatic)
            self.assertEqual(wb.add_fill(got_fill), fill_index)

            border = {
                "left": {"style": 1, "color_argb": 0x01020304, "color": theme},
                "right": {"style": 1, "color_argb": 0x05060708, "color": indexed},
                "top": {"style": 1, "color_argb": 0x090A0B0C, "color": automatic},
            }
            border_index = wb.add_border(border)
            got_border = wb.get_border(border_index)
            self.assertEqual(got_border["left"]["color"], theme)
            self.assertEqual(got_border["right"]["color"], indexed)
            self.assertEqual(got_border["top"]["color"], automatic)
            self.assertEqual(wb.add_border(got_border), border_index)

            dxf_index = wb.add_dxf(
                DifferentialFormat(
                    font=FontRecord(name="DxfSelector", size=9.0, color_argb=0x11121314, color=automatic),
                    fill=FillRecord(pattern=1, fg_argb=0x15161718, fg=indexed),
                    border={"left": {"style": 1, "color_argb": 0x191A1B1C, "color": theme}},
                )
            )
            got_dxf = wb.get_dxf(dxf_index)
            self.assertEqual(got_dxf.font.color, automatic)
            self.assertEqual(got_dxf.fill.fg, indexed)
            self.assertEqual(got_dxf.border["left"]["color"], theme)
            self.assertEqual(wb.add_dxf(got_dxf), dxf_index)

            with Workbook.load(wb.save()) as reloaded:
                self.assertEqual(reloaded.get_font(font_index).color, theme)
                self.assertEqual(reloaded.get_fill(fill_index).fg, indexed)
                self.assertEqual(reloaded.get_border(border_index)["left"]["color"], theme)
                self.assertEqual(reloaded.get_dxf(dxf_index).font.color, automatic)


if __name__ == "__main__":
    unittest.main()
