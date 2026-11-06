"""Cell geometry, formula read-back, range enumeration and display-text surface tests."""

from __future__ import annotations

import unittest
from dataclasses import replace

from formulon import CellXf, FormulonError, Value, ValueKind, Workbook
from formulon.workbook import (
    CellRect,
    DisplayStatus,
    DisplayText,
    EffectiveStyleSource,
    GeometryMode,
    SheetFormatDefaults,
)


class GeometryTests(unittest.TestCase):
    def test_format_defaults_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_format_defaults(
                0,
                SheetFormatDefaults(
                    default_col_width=12.5,
                    default_row_height=18.0,
                    base_col_width=9.0,
                    has_default_col_width=True,
                    has_default_row_height=True,
                ),
            )
            d = wb.get_format_defaults(0)
            self.assertEqual(d.default_col_width, 12.5)
            self.assertEqual(d.default_row_height, 18.0)
            self.assertEqual(d.base_col_width, 9.0)
            self.assertTrue(d.has_default_col_width)
            self.assertTrue(d.has_default_row_height)
            with self.assertRaises(FormulonError):
                wb.set_format_defaults(0, SheetFormatDefaults(default_col_width=-1.0, has_default_col_width=True))

    def test_clear_column_width_follows_default_and_keeps_state(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_format_defaults(
                0, SheetFormatDefaults(default_col_width=10.0, base_col_width=8.0, has_default_col_width=True)
            )
            wb.set_column_width(0, 0, 0, 10.0)
            wb.clear_column_width(0, 0, 0)
            wb.set_format_defaults(
                0, SheetFormatDefaults(default_col_width=20.0, base_col_width=8.0, has_default_col_width=True)
            )
            self.assertEqual(wb.get_column_width_pt(0, 0), wb.get_column_width_pt(0, 1))
            self.assertEqual(wb.list_column_overrides(0), [])

            wb.set_column_width(0, 0, 4, 20.0)
            wb.set_column_hidden(0, 2, 2, True)
            wb.set_column_outline(0, 2, 2, 3)
            wb.clear_column_width(0, 1, 3)
            cols = wb.list_column_overrides(0)
            self.assertEqual(
                [(c.first, c.last, c.has_width, c.hidden, c.outline_level) for c in cols],
                [(0, 0, True, False, 0), (2, 2, False, True, 3), (4, 4, True, False, 0)],
            )
            with self.assertRaises(FormulonError):
                wb.clear_column_width(0, 5, 3)
            with self.assertRaises(FormulonError):
                wb.clear_column_width(0, 0, 16384)

    def test_axis_styles_reach_entered_cells(self) -> None:
        with Workbook.create_default() as wb:
            col_xf = wb.add_cell_xf(replace(wb.get_cell_xf(0), num_fmt_id=1))
            row_xf = wb.add_cell_xf(replace(wb.get_cell_xf(0), num_fmt_id=2))
            wb.set_column_width(0, 1, 1, 20.0)
            wb.set_column_style(0, 1, 1, col_xf)
            wb.set_row_style(0, 2, 2, row_xf)
            self.assertEqual(wb.get_effective_style(0, 4, 1).xf_index, col_xf)
            wb.set_text(0, 4, 1, "column")
            wb.set_text(0, 2, 1, "both")
            self.assertEqual(wb.get_cell_xf_index(0, 4, 1), col_xf)
            self.assertEqual(wb.get_cell_xf_index(0, 2, 1), row_xf)

            wb.clear_column_style(0, 1, 1)
            wb.clear_row_style(0, 2, 2)
            cols = wb.list_column_overrides(0)
            self.assertEqual([(c.first, c.last, c.has_width, c.has_style) for c in cols], [(1, 1, True, False)])
            self.assertEqual(wb.list_row_overrides(0), [])
            with self.assertRaises(FormulonError):
                wb.set_column_style(0, 0, 0, 1000)
            with self.assertRaises(FormulonError):
                wb.set_row_style(0, 0, 0, 1000)

    def test_row_style_span(self) -> None:
        with Workbook.create_default() as wb:
            xf = wb.add_cell_xf(replace(wb.get_cell_xf(0), num_fmt_id=2))
            wb.set_row_style(0, 0, 999, xf)
            self.assertEqual(len(wb.list_row_overrides(0)), 1000)
            wb.clear_row_style(0, 0, 999)
            self.assertEqual(wb.list_row_overrides(0), [])
            with self.assertRaises(FormulonError):
                wb.set_row_style(0, 5, 3, xf)

    def test_range_style_remap_and_restore(self) -> None:
        last_row = 1048575
        with Workbook.create_default() as wb:
            r = wb.add_cell_xf(replace(wb.get_cell_xf(0), num_fmt_id=1))
            a = wb.add_cell_xf(replace(wb.get_cell_xf(0), num_fmt_id=2))
            b = wb.add_cell_xf(replace(wb.get_cell_xf(0), num_fmt_id=3))
            wb.set_row_style(0, 2, 2, r)
            self.assertEqual(wb.get_range_style_bases(0, CellRect(0, 1, last_row, 1)), [0, r])
            # Row 2 of a one-column selection is a single cell; excluding it hides the row style.
            self.assertEqual(
                wb.get_range_style_bases_excluding(0, CellRect(0, 1, last_row, 1), [(2, 1), (2, 1), (9, 9)]), [0]
            )

            snap = wb.snapshot_range_style(0, CellRect(0, 1, last_row, 1))
            wb.remap_range_style(0, CellRect(0, 1, last_row, 1), {0: a, r: b})
            e = wb.get_effective_style(0, 2, 1)
            self.assertEqual((e.source, e.xf_index), (EffectiveStyleSource.CELL, b))
            e = wb.get_effective_style(0, 2, 2)
            self.assertEqual((e.source, e.xf_index), (EffectiveStyleSource.ROW, r))
            e = wb.get_effective_style(0, 499, 1)
            self.assertEqual((e.source, e.xf_index), (EffectiveStyleSource.COLUMN, a))

            wb.restore_range_style(0, snap)
            e = wb.get_effective_style(0, 2, 1)
            self.assertEqual((e.source, e.xf_index), (EffectiveStyleSource.ROW, r))

            with self.assertRaises(FormulonError):
                wb.remap_range_style(0, CellRect(0, 1, last_row, 1), {0: 100000})

    def test_effective_format_defaults(self) -> None:
        with Workbook.create_default() as wb:
            d = wb.get_format_defaults(0)
            # Stored width of the base-8 Calibri 11 default (64 px, shown as 8.43).
            self.assertEqual(d.effective_default_col_width, 2340 / 256)
            self.assertEqual(d.effective_default_row_height, 102 / 7)
            wb.set_format_defaults(
                0,
                SheetFormatDefaults(
                    default_col_width=12.5,
                    default_row_height=20.25,
                    base_col_width=8.0,
                    has_default_col_width=True,
                    has_default_row_height=True,
                ),
            )
            d = wb.get_format_defaults(0)
            self.assertEqual(d.effective_default_col_width, 12.5)
            self.assertEqual(d.effective_default_row_height, 20.25)

    def test_row_height_override_and_clear(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_row_height(0, 2, 2, 30.0)
            rows = wb.list_row_overrides(0)
            self.assertEqual(len(rows), 1)
            self.assertTrue(rows[0].has_height)
            self.assertTrue(rows[0].custom_height)
            self.assertEqual(rows[0].height, 30.0)
            self.assertEqual(wb.get_row_height_pt(0, 2), 30.0)
            wb.clear_row_height(0, 2, 2)
            self.assertEqual(wb.list_row_overrides(0), [])
            wb.clear_row_height(0, 2, 2)

    def test_row_auto_height_stays_automatic(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_row_hidden(0, 4, 4, True)
            wb.set_row_auto_height(0, 4, 4, 31.5)
            rows = wb.list_row_overrides(0)
            self.assertEqual(len(rows), 1)
            self.assertTrue(rows[0].has_height)
            self.assertFalse(rows[0].custom_height)
            self.assertTrue(rows[0].hidden)
            self.assertEqual(rows[0].height, 31.5)

    def test_cell_rect_and_widths(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_row_height(0, 0, 0, 20.0)
            wb.set_row_height(0, 1, 1, 30.0)
            wb.set_column_width(0, 0, 0, 10.0)
            col0 = wb.get_column_width_pt(0, 0)
            col0_print = wb.get_column_width_pt(0, 0, GeometryMode.PRINT)
            self.assertGreater(col0, 0.0)
            self.assertGreater(col0_print, 0.0)
            rect = wb.get_cell_rect_pt(0, CellRect(0, 0, 1, 0))
            self.assertEqual(rect.x, 0.0)
            self.assertEqual(rect.y, 0.0)
            self.assertEqual(rect.width, col0)
            self.assertEqual(rect.height, 50.0)
            below = wb.get_cell_rect_pt(0, CellRect(2, 0, 2, 0), GeometryMode.PRINT)
            self.assertEqual(below.y, 50.0)
            with self.assertRaises(FormulonError):
                wb.get_cell_rect_pt(0, CellRect(3, 0, 1, 0))
            with self.assertRaises(FormulonError):
                wb.get_column_width_pt(0, 0, 9)

    def test_width_model_and_conversions(self) -> None:
        with Workbook.create_default() as wb:
            for mode in (GeometryMode.DISPLAY, GeometryMode.PRINT):
                model = wb.get_width_model(0, mode)
                self.assertGreater(model.points_per_char, 0.0)
                self.assertGreater(model.normal_font_size, 0.0)
                self.assertTrue(model.normal_font_name)
                self.assertEqual(model.platform, "win")
                pt = wb.convert_column_chars_to_pt(0, 10.0, mode)
                self.assertAlmostEqual(wb.convert_column_pt_to_chars(0, pt, mode), 10.0, places=6)
                self.assertEqual(wb.convert_column_chars_to_pt(0, 0.0, mode), 0.0)
            with self.assertRaises(FormulonError):
                wb.convert_column_chars_to_pt(0, -1.0)


class FormulaAndRangeTests(unittest.TestCase):
    def test_get_formula_variants(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 1.0)
            wb.set_formula(0, 1, 1, "=A1+1")
            wb.recalc()
            self.assertEqual(wb.get_formula(0, 1, 1), "=A1+1")
            self.assertEqual(wb.get_formula_r1c1(0, 1, 1), "R[-1]C[-1]+1")
            self.assertIsNone(wb.get_formula(0, 0, 0))
            self.assertIsNone(wb.get_formula_r1c1(0, 5, 5))
            with self.assertRaises(FormulonError):
                wb.get_formula(5, 0, 0)

    def test_cells_in_range_pages(self) -> None:
        with Workbook.create_default() as wb:
            for r in range(5):
                wb.set_number(0, r, 0, float(r))
            wb.set_formula(0, 2, 1, "=A3*2")
            wb.recalc()
            full = CellRect(0, 0, 9, 3)
            page_all = wb.get_cells_in_range(0, full)
            cells, nxt = page_all.cells, page_all.next_cursor
            self.assertIsNone(nxt)
            self.assertEqual([(c.row, c.col) for c in cells], [(0, 0), (1, 0), (2, 0), (2, 1), (3, 0), (4, 0)])
            self.assertEqual(cells[3].formula, "=A3*2")
            self.assertEqual(cells[3].value.number, 4.0)
            self.assertIsNone(cells[0].formula)

            seen = []
            cursor = None
            while True:
                got = wb.get_cells_in_range(0, full, cursor, 2)
                page, cursor = got.cells, got.next_cursor
                self.assertLessEqual(len(page), 2)
                seen.extend((c.row, c.col) for c in page)
                if cursor is None:
                    break
            self.assertEqual(seen, [(c.row, c.col) for c in cells])

            sub = wb.get_cells_in_range(0, CellRect(3, 0, 4, 0)).cells
            self.assertEqual([(c.row, c.col) for c in sub], [(3, 0), (4, 0)])
            with self.assertRaises(FormulonError):
                wb.get_cells_in_range(0, CellRect(4, 0, 1, 0))

    def test_merges_in_range(self) -> None:
        with Workbook.create_default() as wb:
            a = CellRect(0, 0, 1, 1)
            b = CellRect(5, 5, 6, 6)
            wb.add_merge(0, a)
            wb.add_merge(0, b)
            self.assertEqual(wb.get_merges_in_range(0, CellRect(1, 1, 3, 3)), [a])
            self.assertEqual(wb.get_merges_in_range(0, CellRect(0, 0, 10, 10)), [a, b])
            self.assertEqual(wb.get_merges_in_range(0, CellRect(8, 8, 9, 9)), [])
            with self.assertRaises(FormulonError):
                wb.get_merges_in_range(7, a)


class DisplayTextTests(unittest.TestCase):
    def test_display_text_uses_number_format(self) -> None:
        with Workbook.create_default() as wb:
            nf = wb.add_num_fmt("0.00")
            xf = wb.add_cell_xf(
                CellXf(
                    font_index=0,
                    fill_index=0,
                    border_index=0,
                    num_fmt_id=nf,
                    horizontal_align=0,
                    vertical_align=2,
                    wrap_text=False,
                )
            )
            wb.set_number(0, 0, 0, 3.14159)
            wb.set_cell_xf_index(0, 0, 0, xf)
            self.assertEqual(wb.get_display_text(0, 0, 0), DisplayText("3.14", DisplayStatus.OK))
            self.assertEqual(wb.get_display_text(0, 9, 9), DisplayText("", DisplayStatus.OK))
            with self.assertRaises(FormulonError):
                wb.get_display_text(3, 0, 0)

    def test_format_value(self) -> None:
        with Workbook.create_default() as wb:
            num = Value(kind=ValueKind.NUMBER, number=1234.5)
            self.assertEqual(wb.format_value(num, "#,##0.00"), DisplayText("1,234.50", DisplayStatus.OK))
            self.assertEqual(wb.format_value(num), DisplayText("1234.5", DisplayStatus.OK))
            text = Value(kind=ValueKind.TEXT, text="abc")
            self.assertEqual(wb.format_value(text, "General"), DisplayText("abc", DisplayStatus.OK))
            boolean = Value(kind=ValueKind.BOOL, boolean=True)
            self.assertEqual(wb.format_value(boolean).text, "TRUE")
            text_out, status = (lambda r: (r.text, r.display_status))(wb.format_value(num, "0.00["))
            self.assertEqual(status, DisplayStatus.INVALID_FORMAT)
            self.assertTrue(text_out)
            with self.assertRaises(FormulonError):
                wb.format_value(Value(kind=ValueKind.NUMBER, number=float("nan")))


class CellXfFlagTests(unittest.TestCase):
    def test_apply_and_protection_flags_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            xf = wb.add_cell_xf(
                CellXf(
                    font_index=0,
                    fill_index=0,
                    border_index=0,
                    num_fmt_id=0,
                    horizontal_align=0,
                    vertical_align=2,
                    wrap_text=False,
                    apply_number_format=True,
                    apply_font=True,
                    apply_protection=True,
                    quote_prefix=True,
                    has_protection=True,
                    locked=False,
                    hidden=True,
                )
            )
            got = wb.get_cell_xf(xf)
            self.assertTrue(got.apply_number_format)
            self.assertTrue(got.apply_font)
            self.assertFalse(got.apply_fill)
            self.assertFalse(got.apply_border)
            self.assertFalse(got.apply_alignment)
            self.assertTrue(got.apply_protection)
            self.assertTrue(got.quote_prefix)
            self.assertTrue(got.has_protection)
            self.assertFalse(got.locked)
            self.assertTrue(got.hidden)


class PaginationDetailTests(unittest.TestCase):
    def test_pagination_detail(self) -> None:
        with Workbook.create_default() as wb:
            for r in range(0, 120):
                wb.set_number(0, r, 0, float(r))
            wb.add_row_break(0, 10, True)
            p = wb.get_pagination(0)
            self.assertGreater(p.paper.width_pt, 0.0)
            self.assertGreater(p.paper.height_pt, 0.0)
            self.assertTrue(p.paper.known)
            self.assertGreater(p.margins.left, 0.0)
            self.assertGreater(p.printable.width, 0.0)
            self.assertGreater(p.scale, 0.0)
            self.assertIn(p.page_order, (0, 1))
            self.assertFalse(p.print_titles.has_rows)
            self.assertEqual(len(p.pages), p.page_count)
            self.assertGreater(p.page_count, 1)
            self.assertEqual(len(p.horizontal_break_manual), len(p.horizontal_breaks))
            self.assertEqual(len(p.vertical_break_manual), len(p.vertical_breaks))
            self.assertIn(10, p.horizontal_breaks)
            self.assertTrue(p.horizontal_break_manual[p.horizontal_breaks.index(10)])
            first = p.pages[0]
            self.assertEqual(first.rect.first_row, 0)
            self.assertGreater(first.width_pt, 0.0)
            self.assertGreater(first.height_pt, 0.0)


if __name__ == "__main__":
    unittest.main()
