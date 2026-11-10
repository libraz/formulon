"""Header row and totals row toggles, structural deletes and header-only creation through the public API."""

from __future__ import annotations

import unittest
from typing import Tuple

from formulon import (
    ErrorCode,
    FormulonError,
    InvalidArgumentError,
    StatusCode,
    Table,
    TableColumnPatch,
    TableInput,
    TableTotalsRowFunction,
    Workbook,
)


def sales_sheet(wb: Workbook) -> Table:
    """A1:B3 holding Product / Amount over x/10 and y/20, as the table ``Sales``."""
    ws = wb.sheet(0)
    ws.set_text("A1", "Product")
    ws.set_text("B1", "Amount")
    ws.set_text("A2", "x")
    ws.set_number("B2", 10)
    ws.set_text("A3", "y")
    ws.set_number("B3", 20)
    return ws.add_table(TableInput(range="A1:B3", name="Sales"))


def rect_of(table: Table) -> Tuple[int, int, int, int]:
    rect = table.get_properties().range.rect
    return (rect.first_row, rect.first_col, rect.last_row, rect.last_col)


class HeaderRowToggleTests(unittest.TestCase):
    def test_off_deletes_the_header_cells_and_keeps_names_on_writes_them_back(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = sales_sheet(wb)
            ws.set_formula("D1", "=Sales[[#Headers],[Amount]]")
            wb.recalc()
            self.assertEqual(ws.get_value("D1"), "Amount")

            table.update(header_row=False)
            wb.recalc()
            props = table.get_properties()
            self.assertFalse(props.header_row)
            self.assertEqual(rect_of(table), (1, 0, 2, 1))
            self.assertEqual([c.name for c in props.columns], ["Product", "Amount"])
            self.assertIsNone(ws.get_value("A1"))
            self.assertIsNone(ws.get_value("B1"))
            self.assertEqual(ws.get_value("A2"), "x")
            self.assertEqual(ws.get_value("D1").code, ErrorCode.REF)

            table.update(header_row=True)
            wb.recalc()
            self.assertTrue(table.get_properties().header_row)
            self.assertEqual(rect_of(table), (0, 0, 2, 1))
            self.assertEqual(ws.get_value("A1"), "Product")
            self.assertEqual(ws.get_value("B1"), "Amount")
            self.assertEqual(ws.get_value("B3"), 20)
            self.assertEqual(ws.get_value("D1"), "Amount")

    def test_on_shifts_the_data_down_when_the_row_above_is_occupied(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = sales_sheet(wb)
            table.update(header_row=False)
            ws.set_text("A1", "note")
            table.update(header_row=True)
            self.assertEqual(rect_of(table), (1, 0, 3, 1))
            self.assertEqual(ws.get_value("A1"), "note")
            self.assertEqual(ws.get_value("A2"), "Product")
            self.assertEqual(ws.get_value("B2"), "Amount")
            self.assertEqual(ws.get_value("A3"), "x")
            self.assertEqual(ws.get_value("B4"), 20)

    def test_a_range_with_a_toggle_applies_the_toggle_first_and_refuses_a_disagreeing_first_row(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = sales_sheet(wb)
            with self.assertRaises(InvalidArgumentError):
                table.update(header_row=False, range="A1:B3")
            self.assertTrue(table.get_properties().header_row)
            self.assertEqual(rect_of(table), (0, 0, 2, 1))
            self.assertEqual(ws.get_value("A1"), "Product")

            table.update(header_row=False, range="A2:B3")
            self.assertFalse(table.get_properties().header_row)
            self.assertEqual(rect_of(table), (1, 0, 2, 1))
            self.assertIsNone(ws.get_value("A1"))

            table.update(header_row=True, totals_row=True, range="A1:B4")
            props = table.get_properties()
            self.assertTrue(props.header_row)
            self.assertTrue(props.totals_row)
            self.assertEqual(rect_of(table), (0, 0, 3, 1))
            self.assertEqual(ws.get_value("A1"), "Product")


class TotalsRowToggleTests(unittest.TestCase):
    def test_on_writes_defaults_once_and_restores_saved_settings_afterwards(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = sales_sheet(wb)
            table.update(totals_row=True)
            wb.recalc()
            columns = table.get_properties().columns
            self.assertEqual(rect_of(table), (0, 0, 3, 1))
            self.assertEqual(columns[0].totals_row_label, "Total")
            self.assertEqual(columns[0].totals_row_function, TableTotalsRowFunction.NONE)
            self.assertEqual(columns[1].totals_row_function, TableTotalsRowFunction.SUM)
            self.assertEqual(ws.get_value("A4"), "Total")
            self.assertEqual(ws.get_value("B4"), 30)

            table.update(columns=[TableColumnPatch(), TableColumnPatch(totals_row_function=TableTotalsRowFunction.MAX)])
            table.update(totals_row=False)
            table.update(totals_row=True)
            wb.recalc()
            self.assertEqual(table.get_properties().columns[1].totals_row_function, TableTotalsRowFunction.MAX)
            self.assertEqual(ws.get_value("B4"), 20)

    def test_defaults_count_a_text_last_column_and_give_a_one_column_table_the_function_only(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_values("A1", [["Item", "Note", "Qty"], ["a", "p", "q"], ["b", "r", "s"]])
            ws.set_values("E1", [["Solo"], [1], [2]])
            wide = ws.add_table(TableInput(range="A1:C3", name="Wide"))
            solo = ws.add_table(TableInput(range="E1:E3", name="Solo"))
            wide.update(totals_row=True)
            solo.update(totals_row=True)
            self.assertEqual(wide.get_properties().columns[2].totals_row_function, TableTotalsRowFunction.COUNT)
            only = solo.get_properties().columns[0]
            self.assertEqual(only.totals_row_function, TableTotalsRowFunction.SUM)
            self.assertIsNone(only.totals_row_label)

    def test_deleting_the_totals_row_clears_the_settings_and_on_again_writes_defaults(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = sales_sheet(wb)
            table.update(totals_row=True)
            table.update(columns=[TableColumnPatch(), TableColumnPatch(totals_row_function=TableTotalsRowFunction.MAX)])
            ws.delete_rows("4")
            props = table.get_properties()
            self.assertFalse(props.totals_row)
            self.assertEqual(rect_of(table), (0, 0, 2, 1))
            for column in props.columns:
                self.assertEqual(column.totals_row_function, TableTotalsRowFunction.NONE)
                self.assertIsNone(column.totals_row_label)
                self.assertIsNone(column.totals_row_formula)
            table.update(totals_row=True)
            columns = table.get_properties().columns
            self.assertEqual(columns[1].totals_row_function, TableTotalsRowFunction.SUM)
            self.assertEqual(columns[0].totals_row_label, "Total")


class StructuralDeleteTests(unittest.TestCase):
    def test_a_delete_over_the_header_row_is_refused_and_all_data_rows_keep_the_first_empty(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = sales_sheet(wb)
            for span in ("1", "1:2"):
                with self.assertRaises(FormulonError) as caught:
                    ws.delete_rows(span)
                self.assertEqual(caught.exception.code, StatusCode.PRECONDITION_FAILED)
            self.assertEqual(rect_of(table), (0, 0, 2, 1))
            self.assertEqual(ws.get_value("A1"), "Product")
            self.assertEqual(ws.get_value("A2"), "x")

            ws.delete_rows("2:3")
            self.assertEqual(len(ws.list_tables()), 1)
            self.assertEqual(rect_of(table), (0, 0, 1, 1))
            self.assertIsNone(ws.get_value("A2"))
            self.assertIsNone(ws.get_value("B2"))


class HeaderOnlyCreationTests(unittest.TestCase):
    def test_a_header_only_range_creates_the_header_and_one_empty_data_row(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_values("A1", [["Item", "Qty"]])
            table = ws.add_table(TableInput(range="A1:B1", name="Solo"))
            self.assertTrue(table.get_properties().header_row)
            self.assertEqual(rect_of(table), (0, 0, 1, 1))
            self.assertIsNone(ws.get_value("A2"))

            ws.set_values("D1", [["Item", "Qty"], ["kept", 7]])
            shifted = ws.add_table(TableInput(range="D1:E1", name="Shifted"))
            self.assertEqual(rect_of(shifted), (0, 3, 1, 4))
            self.assertIsNone(ws.get_value("D2"))
            self.assertEqual(ws.get_value("D3"), "kept")
            self.assertEqual(ws.get_value("E3"), 7)

    def test_the_last_grid_row_and_header_with_totals_on_one_row_are_refused(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_values("A1", [["Item", "Qty"]])
            ws.set_values("A1048576", [["Item", "Qty"]])
            with self.assertRaises(FormulonError) as caught:
                ws.add_table(TableInput(range="A1048576:B1048576", name="Last"))
            self.assertEqual(caught.exception.code, StatusCode.PRECONDITION_FAILED)
            with self.assertRaises(InvalidArgumentError):
                ws.add_table(TableInput(range="A1:B1", name="Both", header_row=True, totals_row=True))
            self.assertEqual(len(ws.list_tables()), 0)


if __name__ == "__main__":
    unittest.main()
