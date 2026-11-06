"""Sheet structure and feature surface tests."""

from __future__ import annotations

import unittest

import formulon
from formulon import (
    CellRect,
    Comment,
    DataValidationInput,
    FormulonError,
    InvalidArgumentError,
    SheetProtection,
    SheetViewMode,
    SheetVisibility,
    ValueKind,
    Workbook,
)


class SheetStructureTests(unittest.TestCase):
    def test_add_rename_remove_move_sheet(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Second")
            wb.add_sheet("Third")
            self.assertEqual(len(wb), 3)
            wb.rename_sheet(1, "Renamed")
            self.assertEqual(wb.sheet(1).name, "Renamed")
            wb.move_sheet(2, 0)
            self.assertEqual(wb.sheet(0).name, "Third")
            wb.remove_sheet(0)
            self.assertEqual(len(wb), 2)

    def test_copy_sheet(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 5.0)
            wb.sheet(0).set_formula(0, 1, "=A1*2")
            self.assertEqual(wb.copy_sheet(0, 1).index, 1)
            self.assertEqual(wb.sheet(1).name, "Sheet1 (2)")
            self.assertEqual(wb.copy_sheet(0, 0, "Archive").index, 0)
            self.assertEqual(wb.sheet(0).name, "Archive")
            self.assertEqual(len(wb), 3)
            wb.recalc()
            self.assertEqual(wb.sheet(0).get_value(0, 1).number, 10.0)
            with self.assertRaises(FormulonError):
                wb.copy_sheet(3, 0)
            with self.assertRaises(FormulonError):
                wb.copy_sheet(0, 4)
            with self.assertRaises(FormulonError):
                wb.copy_sheet(0, 0, "")
            with self.assertRaises(FormulonError):
                wb.copy_sheet(0, 0, "ARCHIVE")
            self.assertEqual(len(wb), 3)


class MatrixEditTests(unittest.TestCase):
    def test_insert_row_shifts_value_down(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 11.0)  # A1
            wb.sheet(0).insert_rows(0, 1)
            wb.recalc()
            # A1 is now blank; the value moved to A2.
            self.assertEqual(wb.sheet(0).get_value(0, 0).kind, ValueKind.BLANK)
            self.assertEqual(wb.sheet(0).get_value(1, 0).to_python(), 11.0)

    def test_delete_col_shifts_value_left(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)  # A1
            wb.sheet(0).set_number(0, 1, 2.0)  # B1
            wb.sheet(0).delete_columns(0, 1)
            wb.recalc()
            # The old B1 collapses into A1.
            self.assertEqual(wb.sheet(0).get_value(0, 0).to_python(), 2.0)


class MergeCommentHyperlinkTests(unittest.TestCase):
    def test_merge_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_merge(CellRect(0, 0, 1, 1))
            merges = wb.sheet(0).list_merges()
            self.assertEqual(len(merges), 1)
            self.assertEqual((merges[0].first_row, merges[0].last_col), (0, 1))
            wb.sheet(0).clear_merges()
            self.assertEqual(wb.sheet(0).list_merges(), [])

    def test_comment_roundtrip_and_absent(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIsNone(wb.sheet(0).get_comment(0, 0))
            wb.sheet(0).set_comment(0, 0, Comment("alice", "see note"))
            c = wb.sheet(0).get_comment(0, 0)
            self.assertIsNotNone(c)
            self.assertEqual(c.author, "alice")
            self.assertEqual(c.text, "see note")

    def test_comment_invalid_sheet_raises_instead_of_looking_absent(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.IndexOutOfRangeError) as ctx:
                wb.sheet(99).get_comment(0, 0)
            self.assertEqual(ctx.exception.code, formulon.StatusCode.SHEET_INDEX_OUT_OF_RANGE)

    def test_hyperlink_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_hyperlink(2, 3, target="https://example.com", display="Example", tooltip="tip")
            links = wb.sheet(0).list_hyperlinks()
            self.assertEqual(len(links), 1)
            self.assertEqual(links[0].rect, CellRect(2, 3, 2, 3))
            self.assertEqual(links[0].target, "https://example.com")
            self.assertEqual(links[0].display, "Example")

    def test_inverted_or_off_grid_rect_is_invalid_argument(self) -> None:
        with Workbook.create_default() as wb:
            for bad in (CellRect(1, 0, 0, 0), CellRect(0, 1, 0, 0), CellRect(0, 0, 1048576, 0)):
                with self.assertRaises(InvalidArgumentError):
                    wb.sheet(0).add_merge(bad)
                with self.assertRaises(InvalidArgumentError):
                    wb.sheet(0).set_range_xf_index(bad, 0)
                with self.assertRaises(InvalidArgumentError):
                    wb.sheet(0).get_range_style_bases(bad)
            self.assertEqual(wb.sheet(0).list_merges(), [])

    def test_hyperlink_range_roundtrip_at_nonzero_coordinate(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_hyperlink_range(
                CellRect(4, 6, 7, 9), target="https://example.com", display="Range", tooltip="tip"
            )
            links = wb.sheet(0).list_hyperlinks()
            self.assertEqual(len(links), 1)
            self.assertEqual(links[0].rect, CellRect(4, 6, 7, 9))


class ValidationTests(unittest.TestCase):
    def test_list_validation_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_validation(
                DataValidationInput(
                    type=3,  # list
                    ranges=[CellRect(0, 0, 4, 0)],
                    allow_blank=True,
                    formula1='"a,b,c"',
                ),
            )
            self.assertEqual(len(wb.sheet(0).list_validations()), 1)
            dv = wb.sheet(0).list_validations()[0]
            self.assertEqual(dv.type, 3)
            self.assertTrue(dv.allow_blank)
            self.assertEqual(dv.formula1, '"a,b,c"')
            self.assertEqual(len(dv.ranges), 1)
            wb.sheet(0).clear_validations()
            self.assertEqual(len(wb.sheet(0).list_validations()), 0)


class SheetViewProtectionTests(unittest.TestCase):
    def test_view_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_zoom(150)
            wb.sheet(0).set_freeze(1, 2)
            view = wb.sheet(0).get_view()
            self.assertEqual(view.zoom_scale, 150)
            self.assertEqual((view.freeze_rows, view.freeze_cols), (1, 2))

    def test_view_display_defaults(self) -> None:
        with Workbook.create_default() as wb:
            view = wb.sheet(0).get_view()
            self.assertTrue(view.show_grid_lines)
            self.assertTrue(view.show_row_col_headers)
            self.assertTrue(view.show_zeros)
            self.assertFalse(view.right_to_left)
            self.assertFalse(view.tab_selected)
            self.assertEqual(view.view_mode, SheetViewMode.NORMAL)

    def test_view_display_setters_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_show_grid_lines(False)
            wb.sheet(0).set_show_row_col_headers(False)
            wb.sheet(0).set_show_zeros(False)
            wb.sheet(0).set_right_to_left(True)
            wb.sheet(0).set_tab_selected(True)
            wb.sheet(0).set_view_mode(SheetViewMode.PAGE_BREAK_PREVIEW)
            view = wb.sheet(0).get_view()
            self.assertFalse(view.show_grid_lines)
            self.assertFalse(view.show_row_col_headers)
            self.assertFalse(view.show_zeros)
            self.assertTrue(view.right_to_left)
            self.assertTrue(view.tab_selected)
            self.assertEqual(view.view_mode, SheetViewMode.PAGE_BREAK_PREVIEW)
            wb.sheet(0).set_view_mode(SheetViewMode.NORMAL)
            self.assertEqual(wb.sheet(0).get_view().view_mode, SheetViewMode.NORMAL)

    def test_visibility_states_very_hidden(self) -> None:
        with Workbook.create_default() as wb:
            # A workbook keeps one visible sheet, so sheet 1 stays shown.
            wb.add_sheet("Shown")
            self.assertEqual(wb.sheet(0).get_view().visibility, SheetVisibility.VISIBLE)

            wb.sheet(0).set_visibility(SheetVisibility.VERY_HIDDEN)
            self.assertEqual(wb.sheet(0).get_view().visibility, SheetVisibility.VERY_HIDDEN)

            wb.sheet(0).set_visibility(SheetVisibility.HIDDEN)
            self.assertEqual(wb.sheet(0).get_view().visibility, SheetVisibility.HIDDEN)

            wb.sheet(0).set_visibility(SheetVisibility.VISIBLE)
            self.assertEqual(wb.sheet(0).get_view().visibility, SheetVisibility.VISIBLE)

    def test_visibility_rejects_hiding_the_only_visible_sheet(self) -> None:
        with Workbook.create_default() as wb:
            for hide in (
                lambda: wb.sheet(0).set_visibility(SheetVisibility.HIDDEN),
                lambda: wb.sheet(0).set_visibility(SheetVisibility.VERY_HIDDEN),
            ):
                with self.assertRaises(FormulonError):
                    hide()
            self.assertEqual(wb.sheet(0).get_view().visibility, SheetVisibility.VISIBLE)

    def test_visibility_rejects_unknown_state(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Shown")
            wb.sheet(0).set_visibility(SheetVisibility.VERY_HIDDEN)
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_visibility(3)
            # A refused call leaves the sheet as it was.
            self.assertEqual(wb.sheet(0).get_view().visibility, SheetVisibility.VERY_HIDDEN)

    def test_column_row_overrides(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_column_width(0, 2, 18.5)
            cols = wb.sheet(0).list_column_overrides()
            explicit = next(c for c in cols if c.width is not None and abs(c.width - 18.5) < 1e-9)
            self.assertIsNone(explicit.style_xf)
            wb.sheet(0).set_column_width(4, 4, 0.0)
            zero = next(c for c in wb.sheet(0).list_column_overrides() if c.first == 4 and c.last == 4)
            self.assertEqual(zero.width, 0.0)
            wb.sheet(0).set_row_height(0, 0, 30.0)
            rows = wb.sheet(0).list_row_overrides()
            explicit_row = next(r for r in rows if r.height is not None and abs(r.height - 30.0) < 1e-9)
            self.assertIsNone(explicit_row.style_xf)

    def test_protection_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            prot = SheetProtection(enabled=True, sheet=True, format_cells=True)
            wb.sheet(0).set_sheet_protection(prot)
            got = wb.sheet(0).get_sheet_protection()
            self.assertTrue(got.enabled)
            self.assertTrue(got.sheet)
            self.assertTrue(got.format_cells)


if __name__ == "__main__":
    unittest.main()
