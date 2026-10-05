"""Sheet structure and feature surface tests."""

from __future__ import annotations

import unittest

import formulon
from formulon import (
    DataValidationInput,
    FormulonError,
    MergeRange,
    SheetProtection,
    SheetVisibility,
    ValueKind,
    Workbook,
)


class SheetStructureTests(unittest.TestCase):
    def test_add_rename_remove_move_sheet(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Second")
            wb.add_sheet("Third")
            self.assertEqual(wb.sheet_count(), 3)
            wb.rename_sheet(1, "Renamed")
            self.assertEqual(wb.sheet_name(1), "Renamed")
            wb.move_sheet(2, 0)
            self.assertEqual(wb.sheet_name(0), "Third")
            wb.remove_sheet(0)
            self.assertEqual(wb.sheet_count(), 2)


class MatrixEditTests(unittest.TestCase):
    def test_insert_row_shifts_value_down(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 11.0)  # A1
            wb.insert_rows(0, 0, 1)
            wb.recalc()
            # A1 is now blank; the value moved to A2.
            self.assertEqual(wb.get_value(0, 0, 0).kind, ValueKind.BLANK)
            self.assertEqual(wb.get_value(0, 1, 0).to_python(), 11.0)

    def test_delete_col_shifts_value_left(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 1.0)  # A1
            wb.set_number(0, 0, 1, 2.0)  # B1
            wb.delete_cols(0, 0, 1)
            wb.recalc()
            # The old B1 collapses into A1.
            self.assertEqual(wb.get_value(0, 0, 0).to_python(), 2.0)


class MergeCommentHyperlinkTests(unittest.TestCase):
    def test_merge_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_merge(0, MergeRange(0, 0, 1, 1))
            merges = wb.get_merges(0)
            self.assertEqual(len(merges), 1)
            self.assertEqual((merges[0].first_row, merges[0].last_col), (0, 1))
            wb.clear_merges(0)
            self.assertEqual(wb.get_merges(0), [])

    def test_comment_roundtrip_and_absent(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIsNone(wb.get_comment(0, 0, 0))
            wb.set_comment(0, 0, 0, "alice", "see note")
            c = wb.get_comment(0, 0, 0)
            self.assertIsNotNone(c)
            self.assertEqual(c.author, "alice")
            self.assertEqual(c.text, "see note")

    def test_comment_invalid_sheet_raises_instead_of_looking_absent(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.FormulonError) as ctx:
                wb.get_comment(99, 0, 0)
            self.assertEqual(ctx.exception.status, 2)  # kInvalidArgument

    def test_hyperlink_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_hyperlink(0, 2, 3, "https://example.com", "Example", "tip")
            links = wb.get_hyperlinks(0)
            self.assertEqual(len(links), 1)
            self.assertEqual((links[0].row, links[0].col), (2, 3))
            self.assertEqual((links[0].last_row, links[0].last_col), (2, 3))
            self.assertEqual(links[0].target, "https://example.com")
            self.assertEqual(links[0].display, "Example")

    def test_hyperlink_range_roundtrip_at_nonzero_coordinate(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_hyperlink_range(0, 4, 6, 7, 9, "https://example.com", "Range", "tip")
            links = wb.get_hyperlinks(0)
            self.assertEqual(len(links), 1)
            self.assertEqual((links[0].row, links[0].col), (4, 6))
            self.assertEqual((links[0].last_row, links[0].last_col), (7, 9))


class ValidationTests(unittest.TestCase):
    def test_list_validation_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_validation(
                0,
                DataValidationInput(
                    type=3,  # list
                    ranges=[MergeRange(0, 0, 4, 0)],
                    allow_blank=True,
                    formula1='"a,b,c"',
                ),
            )
            self.assertEqual(wb.validation_count(0), 1)
            dv = wb.get_validation_at(0, 0)
            self.assertEqual(dv.type, 3)
            self.assertTrue(dv.allow_blank)
            self.assertEqual(dv.formula1, '"a,b,c"')
            self.assertEqual(len(dv.ranges), 1)
            wb.clear_validations(0)
            self.assertEqual(wb.validation_count(0), 0)


class SheetViewProtectionTests(unittest.TestCase):
    def test_view_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_sheet_zoom(0, 150)
            wb.set_sheet_freeze(0, 1, 2)
            view = wb.get_sheet_view(0)
            self.assertEqual(view.zoom_scale, 150)
            self.assertEqual((view.freeze_rows, view.freeze_cols), (1, 2))

    def test_view_display_defaults(self) -> None:
        with Workbook.create_default() as wb:
            view = wb.get_sheet_view(0)
            self.assertTrue(view.show_grid_lines)
            self.assertTrue(view.show_row_col_headers)
            self.assertTrue(view.show_zeros)
            self.assertFalse(view.right_to_left)
            self.assertFalse(view.tab_selected)
            self.assertEqual(view.view_mode, "")

    def test_view_display_setters_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_sheet_show_grid_lines(0, False)
            wb.set_sheet_show_row_col_headers(0, False)
            wb.set_sheet_show_zeros(0, False)
            wb.set_sheet_right_to_left(0, True)
            wb.set_sheet_tab_selected(0, True)
            wb.set_sheet_view_mode(0, "pageBreakPreview")
            view = wb.get_sheet_view(0)
            self.assertFalse(view.show_grid_lines)
            self.assertFalse(view.show_row_col_headers)
            self.assertFalse(view.show_zeros)
            self.assertTrue(view.right_to_left)
            self.assertTrue(view.tab_selected)
            self.assertEqual(view.view_mode, "pageBreakPreview")
            # An empty mode is a meaningful value (the OOXML-default
            # "normal" view), not a no-op -- it must round-trip too.
            wb.set_sheet_view_mode(0, "")
            self.assertEqual(wb.get_sheet_view(0).view_mode, "")

    def test_visibility_states_very_hidden(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.get_sheet_view(0).visibility, SheetVisibility.VISIBLE)
            self.assertFalse(wb.get_sheet_view(0).tab_hidden)

            wb.set_sheet_visibility(0, SheetVisibility.VERY_HIDDEN)
            view = wb.get_sheet_view(0)
            self.assertEqual(view.visibility, SheetVisibility.VERY_HIDDEN)
            # The two-state view stays consistent: a very-hidden sheet is
            # hidden to code that reads only the bool, never visible.
            self.assertTrue(view.tab_hidden)

            # "Hidden" says nothing a very-hidden sheet does not already
            # satisfy, so it must not weaken the author's stronger choice.
            wb.set_sheet_tab_hidden(0, True)
            self.assertEqual(wb.get_sheet_view(0).visibility, SheetVisibility.VERY_HIDDEN)

            # Demotion is the other direction the bool cannot express.
            wb.set_sheet_visibility(0, SheetVisibility.HIDDEN)
            self.assertEqual(wb.get_sheet_view(0).visibility, SheetVisibility.HIDDEN)

            wb.set_sheet_tab_hidden(0, False)
            self.assertEqual(wb.get_sheet_view(0).visibility, SheetVisibility.VISIBLE)

    def test_visibility_rejects_unknown_state(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_sheet_visibility(0, SheetVisibility.VERY_HIDDEN)
            with self.assertRaises(FormulonError):
                wb.set_sheet_visibility(0, 3)
            # A refused call leaves the sheet as it was.
            self.assertEqual(wb.get_sheet_view(0).visibility, SheetVisibility.VERY_HIDDEN)

    def test_column_row_overrides(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_column_width(0, 0, 2, 18.5)
            cols = wb.get_sheet_columns(0)
            explicit = next(c for c in cols if abs(c.width - 18.5) < 1e-9)
            self.assertTrue(explicit.has_width)
            self.assertFalse(explicit.has_style)
            wb.set_column_width(0, 4, 4, 0.0)
            zero = next(c for c in wb.get_sheet_columns(0) if c.first == 4 and c.last == 4)
            self.assertEqual(zero.width, 0.0)
            self.assertTrue(zero.has_width)
            wb.set_row_height(0, 0, 30.0)
            rows = wb.get_sheet_row_overrides(0)
            explicit_row = next(r for r in rows if abs(r.height - 30.0) < 1e-9)
            self.assertFalse(explicit_row.has_style)
            self.assertEqual(explicit_row.style_xf, 0)

    def test_protection_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            prot = SheetProtection(enabled=True, sheet=True, format_cells=True)
            wb.set_sheet_protection(0, prot)
            got = wb.get_sheet_protection(0)
            self.assertTrue(got.enabled)
            self.assertTrue(got.sheet)
            self.assertTrue(got.format_cells)


if __name__ == "__main__":
    unittest.main()
