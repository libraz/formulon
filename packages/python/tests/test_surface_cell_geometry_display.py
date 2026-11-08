"""Cell geometry, formula read-back, range enumeration and display-text surface tests."""

from __future__ import annotations

import unittest
from dataclasses import replace

from _xf_input import input_of

from formulon import CellXfInput, FormulonError, PrintTitles, RowBreak, Value, ValueKind, Workbook
from formulon._workbook import (
    CellRect,
    DisplayStatus,
    DisplayText,
    EffectiveStyleSource,
    GeometryMode,
)


class GeometryTests(unittest.TestCase):
    def test_format_defaults_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).update_format_defaults(
                default_col_width_chars=12.5, default_row_height_pt=18.0, base_col_width_chars=9.0
            )
            d = wb.sheet(0).get_format_defaults()
            self.assertEqual(d.default_col_width_chars, 12.5)
            self.assertEqual(d.default_row_height_pt, 18.0)
            self.assertEqual(d.base_col_width_chars, 9.0)
            wb.sheet(0).update_format_defaults(default_col_width_chars=None, default_row_height_pt=None)
            d = wb.sheet(0).get_format_defaults()
            self.assertIsNone(d.default_col_width_chars)
            self.assertIsNone(d.default_row_height_pt)
            with self.assertRaises(FormulonError):
                wb.sheet(0).update_format_defaults(default_col_width_chars=-1.0)

    def test_clear_column_width_follows_default_and_keeps_state(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).update_format_defaults(default_col_width_chars=10.0, base_col_width_chars=8.0)
            wb.sheet(0).set_column_width(0, 0, 10.0)
            wb.sheet(0).reset_column_width(0, 0)
            wb.sheet(0).update_format_defaults(default_col_width_chars=20.0, base_col_width_chars=8.0)
            self.assertEqual(wb.sheet(0).get_column_width_pt(0), wb.sheet(0).get_column_width_pt(1))
            self.assertEqual(wb.sheet(0).list_column_overrides(), [])

            wb.sheet(0).set_column_width(0, 4, 20.0)
            wb.sheet(0).set_column_hidden(2, 2, True)
            wb.sheet(0).set_column_outline(2, 2, 3)
            wb.sheet(0).reset_column_width(1, 3)
            cols = wb.sheet(0).list_column_overrides()
            self.assertEqual(
                [(c.first, c.last, c.width_chars is not None, c.hidden, c.outline_level) for c in cols],
                [(0, 0, True, False, 0), (2, 2, False, True, 3), (4, 4, True, False, 0)],
            )
            with self.assertRaises(FormulonError):
                wb.sheet(0).reset_column_width(5, 3)
            with self.assertRaises(FormulonError):
                wb.sheet(0).reset_column_width(0, 16384)

    def test_axis_styles_reach_entered_cells(self) -> None:
        with Workbook.create_default() as wb:
            col_xf = wb.ooxml.add_cell_xf(replace(input_of(wb.ooxml.get_cell_xf(0)), num_fmt_id=1))
            row_xf = wb.ooxml.add_cell_xf(replace(input_of(wb.ooxml.get_cell_xf(0)), num_fmt_id=2))
            wb.sheet(0).set_column_width(1, 1, 20.0)
            wb.sheet(0).ooxml.set_column_xf_index(1, 1, col_xf)
            wb.sheet(0).ooxml.set_row_xf_index(2, 2, row_xf)
            self.assertEqual(wb.sheet(0).ooxml.get_effective_style(4, 1).xf_index, col_xf)
            wb.sheet(0).set_text(4, 1, "column")
            wb.sheet(0).set_text(2, 1, "both")
            self.assertEqual(wb.sheet(0).ooxml.get_cell_xf_index(4, 1), col_xf)
            self.assertEqual(wb.sheet(0).ooxml.get_cell_xf_index(2, 1), row_xf)

            wb.sheet(0).ooxml.reset_column_xf_index(1, 1)
            wb.sheet(0).ooxml.reset_row_xf_index(2, 2)
            cols = wb.sheet(0).list_column_overrides()
            self.assertEqual(
                [
                    (c.first, c.last, c.width_chars is not None, wb.sheet(0).ooxml.get_column_xf_index(c.first))
                    for c in cols
                ],
                [(1, 1, True, None)],
            )
            self.assertEqual(wb.sheet(0).list_row_overrides(), [])
            with self.assertRaises(FormulonError):
                wb.sheet(0).ooxml.set_column_xf_index(0, 0, 1000)
            with self.assertRaises(FormulonError):
                wb.sheet(0).ooxml.set_row_xf_index(0, 0, 1000)

    def test_row_style_span(self) -> None:
        with Workbook.create_default() as wb:
            xf = wb.ooxml.add_cell_xf(replace(input_of(wb.ooxml.get_cell_xf(0)), num_fmt_id=2))
            wb.sheet(0).ooxml.set_row_xf_index(0, 999, xf)
            self.assertEqual(len(wb.sheet(0).list_row_overrides()), 1000)
            wb.sheet(0).ooxml.reset_row_xf_index(0, 999)
            self.assertEqual(wb.sheet(0).list_row_overrides(), [])
            with self.assertRaises(FormulonError):
                wb.sheet(0).ooxml.set_row_xf_index(5, 3, xf)

    def test_range_style_remap_and_restore(self) -> None:
        last_row = 1048575
        with Workbook.create_default() as wb:
            r = wb.ooxml.add_cell_xf(replace(input_of(wb.ooxml.get_cell_xf(0)), num_fmt_id=1))
            a = wb.ooxml.add_cell_xf(replace(input_of(wb.ooxml.get_cell_xf(0)), num_fmt_id=2))
            b = wb.ooxml.add_cell_xf(replace(input_of(wb.ooxml.get_cell_xf(0)), num_fmt_id=3))
            wb.sheet(0).ooxml.set_row_xf_index(2, 2, r)
            self.assertEqual(wb.sheet(0).ooxml.get_range_xf_bases(CellRect(0, 1, last_row, 1)), [0, r])
            # Row 2 of a one-column selection is a single cell; excluding it hides the row style.
            self.assertEqual(
                wb.sheet(0).ooxml.get_range_xf_bases_excluding(CellRect(0, 1, last_row, 1), [(2, 1), (2, 1), (9, 9)]),
                [0],
            )

            snap = wb.sheet(0).ooxml.snapshot_range_xfs(CellRect(0, 1, last_row, 1))
            wb.sheet(0).ooxml.remap_range_xfs(CellRect(0, 1, last_row, 1), {0: a, r: b})
            e = wb.sheet(0).ooxml.get_effective_style(2, 1)
            self.assertEqual((e.source, e.xf_index), (EffectiveStyleSource.CELL, b))
            e = wb.sheet(0).ooxml.get_effective_style(2, 2)
            self.assertEqual((e.source, e.xf_index), (EffectiveStyleSource.ROW, r))
            e = wb.sheet(0).ooxml.get_effective_style(499, 1)
            self.assertEqual((e.source, e.xf_index), (EffectiveStyleSource.COLUMN, a))

            wb.sheet(0).ooxml.restore_range_xfs(snap)
            e = wb.sheet(0).ooxml.get_effective_style(2, 1)
            self.assertEqual((e.source, e.xf_index), (EffectiveStyleSource.ROW, r))

            with self.assertRaises(FormulonError):
                wb.sheet(0).ooxml.remap_range_xfs(CellRect(0, 1, last_row, 1), {0: 100000})

    def test_effective_format_defaults(self) -> None:
        with Workbook.create_default() as wb:
            d = wb.sheet(0).get_format_defaults()
            # Stored width of the base-8 Calibri 11 default (64 px, shown as 8.43).
            self.assertEqual(d.effective_default_col_width_chars, 2340 / 256)
            self.assertEqual(d.effective_default_row_height_pt, 102 / 7)
            wb.sheet(0).update_format_defaults(
                default_col_width_chars=12.5, default_row_height_pt=20.25, base_col_width_chars=8.0
            )
            d = wb.sheet(0).get_format_defaults()
            self.assertEqual(d.effective_default_col_width_chars, 12.5)
            self.assertEqual(d.effective_default_row_height_pt, 20.25)

    def test_row_height_override_and_clear(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_row_height(2, 2, 30.0)
            rows = wb.sheet(0).list_row_overrides()
            self.assertEqual(len(rows), 1)
            self.assertTrue(rows[0].custom_height)
            self.assertEqual(rows[0].height_pt, 30.0)
            self.assertEqual(wb.sheet(0).get_row_height_pt(2), 30.0)
            wb.sheet(0).reset_row_height(2, 2)
            self.assertEqual(wb.sheet(0).list_row_overrides(), [])
            wb.sheet(0).reset_row_height(2, 2)

    def test_row_auto_height_stays_automatic(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_row_hidden(4, 4, True)
            wb.sheet(0).set_row_auto_height(4, 4, 31.5)
            rows = wb.sheet(0).list_row_overrides()
            self.assertEqual(len(rows), 1)
            self.assertFalse(rows[0].custom_height)
            self.assertTrue(rows[0].hidden)
            self.assertEqual(rows[0].height_pt, 31.5)

    def test_cell_rect_and_widths(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_row_height(0, 0, 20.0)
            wb.sheet(0).set_row_height(1, 1, 30.0)
            wb.sheet(0).set_column_width(0, 0, 10.0)
            col0 = wb.sheet(0).get_column_width_pt(0)
            col0_print = wb.sheet(0).get_column_width_pt(0, GeometryMode.PRINT)
            self.assertGreater(col0, 0.0)
            self.assertGreater(col0_print, 0.0)
            rect = wb.sheet(0).get_cell_rect_pt(CellRect(0, 0, 1, 0))
            self.assertEqual(rect.x, 0.0)
            self.assertEqual(rect.y, 0.0)
            self.assertEqual(rect.width_pt, col0)
            self.assertEqual(rect.height_pt, 50.0)
            below = wb.sheet(0).get_cell_rect_pt(CellRect(2, 0, 2, 0), GeometryMode.PRINT)
            self.assertEqual(below.y, 50.0)
            with self.assertRaises(FormulonError):
                wb.sheet(0).get_cell_rect_pt(CellRect(3, 0, 1, 0))
            with self.assertRaises(ValueError):
                wb.sheet(0).get_column_width_pt(0, 9)

    def test_width_model_and_conversions(self) -> None:
        with Workbook.create_default() as wb:
            for mode in (GeometryMode.DISPLAY, GeometryMode.PRINT):
                model = wb.sheet(0).get_width_model(mode)
                self.assertGreater(model.points_per_char, 0.0)
                self.assertGreater(model.normal_font_size_pt, 0.0)
                self.assertTrue(model.normal_font_name)
                self.assertEqual(model.platform, "win")
                pt = wb.sheet(0).convert_column_chars_to_pt(10.0, mode)
                self.assertAlmostEqual(wb.sheet(0).convert_column_pt_to_chars(pt, mode), 10.0, places=6)
                self.assertEqual(wb.sheet(0).convert_column_chars_to_pt(0.0, mode), 0.0)
            with self.assertRaises(FormulonError):
                wb.sheet(0).convert_column_chars_to_pt(-1.0)


class FormulaAndRangeTests(unittest.TestCase):
    def test_get_formula_variants(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            wb.sheet(0).set_formula(1, 1, "=A1+1")
            wb.recalc()
            self.assertEqual(wb.sheet(0).get_formula(1, 1), "=A1+1")
            self.assertEqual(wb.sheet(0).get_formula_r1c1(1, 1), "=R[-1]C[-1]+1")
            self.assertIsNone(wb.sheet(0).get_formula(0, 0))
            self.assertIsNone(wb.sheet(0).get_formula_r1c1(5, 5))
            with self.assertRaises(FormulonError):
                wb.sheet(5).get_formula(0, 0)

    def test_cells_in_range_pages(self) -> None:
        with Workbook.create_default() as wb:
            for r in range(5):
                wb.sheet(0).set_number(r, 0, float(r))
            wb.sheet(0).set_formula(2, 1, "=A3*2")
            wb.recalc()
            full = CellRect(0, 0, 9, 3)
            page_all = wb.sheet(0).get_cells_in_range(full)
            cells, nxt = page_all.cells, page_all.next_cursor
            self.assertIsNone(nxt)
            self.assertEqual([(c.row, c.col) for c in cells], [(0, 0), (1, 0), (2, 0), (2, 1), (3, 0), (4, 0)])
            self.assertEqual(cells[3].formula, "=A3*2")
            self.assertEqual(cells[3].value, 4.0)
            self.assertIsNone(cells[0].formula)

            seen = []
            cursor = None
            while True:
                got = wb.sheet(0).get_cells_in_range(full, cursor, 2)
                page, cursor = got.cells, got.next_cursor
                self.assertLessEqual(len(page), 2)
                seen.extend((c.row, c.col) for c in page)
                if cursor is None:
                    break
            self.assertEqual(seen, [(c.row, c.col) for c in cells])

            sub = wb.sheet(0).get_cells_in_range(CellRect(3, 0, 4, 0)).cells
            self.assertEqual([(c.row, c.col) for c in sub], [(3, 0), (4, 0)])
            with self.assertRaises(FormulonError):
                wb.sheet(0).get_cells_in_range(CellRect(4, 0, 1, 0))

    def test_merges_in_range(self) -> None:
        with Workbook.create_default() as wb:
            a = CellRect(0, 0, 1, 1)
            b = CellRect(5, 5, 6, 6)
            wb.sheet(0).add_merge(a)
            wb.sheet(0).add_merge(b)
            self.assertEqual(wb.sheet(0).get_merges_in_range(CellRect(1, 1, 3, 3)), [a])
            self.assertEqual(wb.sheet(0).get_merges_in_range(CellRect(0, 0, 10, 10)), [a, b])
            self.assertEqual(wb.sheet(0).get_merges_in_range(CellRect(8, 8, 9, 9)), [])
            with self.assertRaises(FormulonError):
                wb.sheet(7).get_merges_in_range(a)


class DisplayTextTests(unittest.TestCase):
    def test_display_text_uses_number_format(self) -> None:
        with Workbook.create_default() as wb:
            nf = wb.ooxml.add_number_format("0.00")
            xf = wb.ooxml.add_cell_xf(
                CellXfInput(
                    font_index=0,
                    fill_index=0,
                    border_index=0,
                    num_fmt_id=nf,
                    horizontal=0,
                    vertical=2,
                    wrap_text=False,
                )
            )
            wb.sheet(0).set_number(0, 0, 3.14159)
            wb.sheet(0).ooxml.set_cell_xf_index(0, 0, xf)
            self.assertEqual(wb.sheet(0).get_display_text(0, 0), DisplayText("3.14", DisplayStatus.OK))
            self.assertEqual(wb.sheet(0).get_display_text(9, 9), DisplayText("", DisplayStatus.OK))
            with self.assertRaises(FormulonError):
                wb.sheet(3).get_display_text(0, 0)

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
            xf = wb.ooxml.add_cell_xf(
                CellXfInput(
                    font_index=0,
                    fill_index=0,
                    border_index=0,
                    num_fmt_id=0,
                    horizontal=0,
                    vertical=2,
                    wrap_text=False,
                    apply_number_format=True,
                    apply_font=True,
                    apply_protection=True,
                    quote_prefix=True,
                    locked=False,
                    hidden=True,
                )
            )
            got = wb.ooxml.get_cell_xf(xf)
            self.assertTrue(got.apply_number_format)
            self.assertTrue(got.apply_font)
            self.assertFalse(got.apply_fill)
            self.assertFalse(got.apply_border)
            self.assertFalse(got.apply_alignment)
            self.assertTrue(got.apply_protection)
            self.assertTrue(got.quote_prefix)
            self.assertFalse(got.locked)
            self.assertTrue(got.hidden)


class PaginationDetailTests(unittest.TestCase):
    def test_pagination_detail(self) -> None:
        with Workbook.create_default() as wb:
            for r in range(0, 120):
                wb.sheet(0).set_number(r, 0, float(r))
            wb.sheet(0).set_row_break(10, RowBreak(10))
            p = wb.sheet(0).get_pagination()
            self.assertGreater(p.paper.width_pt, 0.0)
            self.assertGreater(p.paper.height_pt, 0.0)
            self.assertTrue(p.paper.known)
            self.assertGreater(p.margins.left_in, 0.0)
            self.assertGreater(p.printable.width_pt, 0.0)
            self.assertGreater(p.scale_pct, 0.0)
            self.assertIn(p.page_order, (0, 1))
            self.assertEqual(p.print_titles, PrintTitles(rows=None, columns=None))
            self.assertEqual(len(p.pages), p.page_count)
            self.assertGreater(p.page_count, 1)
            self.assertIn(10, [b.row for b in p.row_breaks])
            self.assertTrue(next(b for b in p.row_breaks if b.row == 10).manual)
            first = p.pages[0]
            self.assertEqual(first.rect.first_row, 0)
            self.assertGreater(first.width_pt, 0.0)
            self.assertGreater(first.height_pt, 0.0)


if __name__ == "__main__":
    unittest.main()
