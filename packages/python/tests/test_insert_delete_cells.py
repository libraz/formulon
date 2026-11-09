"""Sheet.insert_cells / delete_cells: band shifts, reference rewriting, refusals and argument errors."""

from __future__ import annotations

import unittest
from typing import Any, Callable

import formulon
from formulon import (
    CellRect,
    CfExpressionRuleInput,
    ErrorCode,
    ExcelError,
    FormulonError,
    StatusCode,
    TableColumnTotalsInput,
    TableInput,
    TableTotalsRowFunction,
    Workbook,
)


class InsertDeleteCellsTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def assertCode(self, call: Callable[[], Any], code: StatusCode) -> None:
        with self.assertRaises(FormulonError) as caught:
            call()
        self.assertEqual(caught.exception.code, code)

    def test_cell_shift_ordinals(self) -> None:
        shift = formulon.CellShift
        self.assertEqual((shift.DOWN, shift.RIGHT, shift.UP, shift.LEFT), (0, 1, 2, 3))

    def test_insert_down_moves_band_and_references(self) -> None:
        self.ws.set_value("B5", 7)
        self.ws.set_value("D5", 9)
        self.ws.set_formula("A1", "=B5")
        self.ws.insert_cells("B3:C4", formulon.CellShift.DOWN)
        self.assertEqual(self.ws.get_value("B7"), 7.0)
        self.assertIsNone(self.ws.get_value("B5"))
        self.assertEqual(self.ws.get_value("D5"), 9.0)
        self.assertEqual(self.ws.get_formula("A1"), "=B7")

    def test_delete_left_ref_error(self) -> None:
        self.ws.set_value("C3", 1)
        self.ws.set_value("D3", 2)
        self.ws.set_formula("A1", "=C3")
        self.ws.set_formula("A2", "=D3")
        self.ws.delete_cells("B3:C4", formulon.CellShift.LEFT)
        self.assertIn("#REF!", self.ws.get_formula("A1"))
        self.assertEqual(self.ws.get_value("A1"), ExcelError(ErrorCode.REF))
        self.assertEqual(self.ws.get_formula("A2"), "=B3")
        self.assertEqual(self.ws.get_value("B3"), 2.0)

    def test_straddling_range_unchanged(self) -> None:
        self.ws.set_formula("A1", "=SUM(A5:D5)")
        self.ws.insert_cells("B3:C4", formulon.CellShift.DOWN)
        self.assertEqual(self.ws.get_formula("A1"), "=SUM(A5:D5)")

    def test_cf_range_splits(self) -> None:
        self.ws.add_cf_rule(CfExpressionRuleInput(ranges=["A3:D3"], formula1="TRUE"))
        self.ws.insert_cells("B3:C4", formulon.CellShift.DOWN)
        (rule,) = self.ws.list_cf_rules()
        self.assertEqual(
            frozenset(rule.ranges),
            frozenset({CellRect(2, 0, 2, 0), CellRect(2, 3, 2, 3), CellRect(4, 1, 4, 2)}),
        )

    def test_partial_merge_unmerged(self) -> None:
        self.ws.add_merge("A3:B3")
        self.ws.insert_cells("B3:C4", formulon.CellShift.DOWN)
        self.assertEqual(self.ws.list_merges(), [])

    def test_refusal_precondition_failed(self) -> None:
        for a1, header in (("A3", "a"), ("B3", "b"), ("C3", "c"), ("D3", "d")):
            self.ws.set_value(a1, header)
        table = self.ws.add_table(TableInput(range="A3:D6", name="Partial"))
        before = table.get_properties().range.rect
        self.assertCode(lambda: self.ws.insert_cells("B3:C4", formulon.CellShift.DOWN), StatusCode.PRECONDITION_FAILED)
        self.assertEqual(table.get_properties().range.rect, before)
        self.assertEqual(self.ws.get_value("B3"), "b")

    def test_wrong_direction_invalid_argument(self) -> None:
        shift = formulon.CellShift
        invalid = StatusCode.INVALID_ARGUMENT
        self.assertCode(lambda: self.ws.insert_cells("B3:C4", shift.UP), invalid)
        self.assertCode(lambda: self.ws.insert_cells("B3:C4", shift.LEFT), invalid)
        self.assertCode(lambda: self.ws.delete_cells("B3:C4", shift.DOWN), invalid)
        self.assertCode(lambda: self.ws.delete_cells("B3:C4", shift.RIGHT), invalid)
        self.assertCode(lambda: self.ws.insert_cells("B3:C4", 7), invalid)
        self.assertCode(lambda: self.ws.insert_cells("B:C", shift.DOWN), invalid)
        self.assertCode(lambda: self.ws.insert_cells("3:4", shift.RIGHT), invalid)
        self.assertCode(lambda: self.ws.insert_cells(CellRect(1048576, 1, 1048576, 2), shift.DOWN), invalid)

    def test_whole_column_right_matches_insert_columns(self) -> None:
        self.ws.set_value("C2", 5)
        self.ws.insert_cells("B:B", formulon.CellShift.RIGHT)
        self.assertEqual(self.ws.get_value("D2"), 5.0)

    def test_totals_row_on_shifts(self) -> None:
        for a1, value in (("A1", "Region"), ("B1", "Amt"), ("A2", "East"), ("B2", 10), ("A3", "West"), ("B3", 30)):
            self.ws.set_value(a1, value)
        table = self.ws.add_table(TableInput(range="A1:B3", name="Sales"))
        table.set_column_totals(1, TableColumnTotalsInput(totals_row_function=TableTotalsRowFunction.SUM))
        self.ws.set_value("B4", 5)
        table.update(totals_row=True)
        self.assertEqual(table.get_properties().range.rect, CellRect(0, 0, 3, 1))
        self.assertEqual(self.ws.get_value("B5"), 5.0)
        self.assertEqual(self.ws.get_value("B4"), 40.0)

    def test_totals_row_off_refused_when_referenced(self) -> None:
        for a1, value in (("A1", "Region"), ("B1", "Amt"), ("A2", "East"), ("B2", 10)):
            self.ws.set_value(a1, value)
        table = self.ws.add_table(TableInput(range="A1:B3", name="Sales", totals_row=True))
        self.ws.set_formula("D1", "=B3")
        self.assertCode(lambda: table.update(totals_row=False), StatusCode.PRECONDITION_FAILED)
        self.assertTrue(table.get_properties().totals_row)
        self.assertEqual(table.get_properties().range.rect, CellRect(0, 0, 2, 1))

    def test_float_or_bool_argument_type_error(self) -> None:
        shift = formulon.CellShift
        with self.assertRaises(TypeError):
            self.ws.insert_cells(CellRect(1.5, 1, 3, 2), shift.DOWN)  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.insert_cells(CellRect(True, 1, 3, 2), shift.DOWN)  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.insert_cells("B3:C4", True)  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.delete_cells("B3:C4", 1.5)  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.insert_cells(None, shift.DOWN)  # type: ignore[arg-type]
        with self.assertRaises(TypeError):
            self.ws.insert_cells("B3:C4")  # type: ignore[call-arg]


if __name__ == "__main__":
    unittest.main()
