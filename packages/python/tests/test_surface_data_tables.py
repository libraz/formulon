"""What-If data table surface tests."""

from __future__ import annotations

import io
import unittest
import zipfile

from formulon import (
    CalcMode,
    CellRect,
    DataTable,
    DataTableInputCell,
    FormulonError,
    ValueKind,
    Workbook,
)


def _sheet_xml(data: bytes) -> str:
    return zipfile.ZipFile(io.BytesIO(data)).read("xl/worksheets/sheet1.xml").decode("utf-8")


def _squares() -> Workbook:
    """B1 = 3 is the column input, B2 = B1 * B1 the formula, A3:A5 = 1..3 the values (results B3:B5)."""
    wb = Workbook.create_default()
    wb.sheet(0).set_number(0, 1, 3.0)
    wb.sheet(0).set_formula(1, 1, "=B1*B1")
    for i in range(3):
        wb.sheet(0).set_number(2 + i, 0, float(i + 1))
    return wb


def _number(wb: Workbook, row: int, col: int) -> float:
    value = wb.sheet(0).get_value(row, col)
    assert value.kind is ValueKind.NUMBER, (row, col, value)
    return value.number


class DataTableTests(unittest.TestCase):
    def test_one_variable_table_computes_reads_back_and_refuses_writes(self) -> None:
        with _squares() as wb:
            wb.sheet(0).add_data_table(CellRect(2, 1, 4, 1), col_input=(0, 1))
            wb.recalc()
            self.assertEqual([_number(wb, r, 1) for r in (2, 3, 4)], [1.0, 4.0, 9.0])
            self.assertEqual(_number(wb, 0, 1), 3.0)
            self.assertEqual(_number(wb, 1, 1), 9.0)

            table = wb.sheet(0).get_data_table(3, 1)
            self.assertEqual(
                table,
                DataTable(
                    rect=CellRect(2, 1, 4, 1),
                    two_dimensional=False,
                    row_oriented=False,
                    row_input=None,
                    col_input=DataTableInputCell(0, 1, False),
                ),
            )
            self.assertIsNone(wb.sheet(0).get_data_table(1, 1))

            with self.assertRaises(FormulonError):
                wb.sheet(0).set_number(3, 1, 7.0)
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_text(2, 1, "x")
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_formula(4, 1, "=1")
            self.assertEqual(_number(wb, 3, 1), 4.0)
            wb.sheet(0).set_number(3, 0, 10.0)
            wb.recalc()
            self.assertEqual(_number(wb, 3, 1), 100.0)

    def test_two_variable_table_and_row_edits(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 5, 1.0)
            wb.sheet(0).set_number(1, 5, 1.0)
            wb.sheet(0).set_formula(2, 2, "=F1*F2")
            for k in range(3):
                wb.sheet(0).set_number(2, 3 + k, float(k + 1))
                wb.sheet(0).set_number(3 + k, 2, float(10 * (k + 1)))
            wb.sheet(0).add_data_table(CellRect(3, 3, 5, 5), row_input=(0, 5), col_input=(1, 5))
            wb.recalc()
            self.assertEqual([_number(wb, 3 + k, 3 + k) for k in range(3)], [10.0, 40.0, 90.0])
            table = wb.sheet(0).get_data_table(4, 4)
            self.assertTrue(table.two_dimensional)
            self.assertEqual(table.row_input, DataTableInputCell(0, 5, False))
            self.assertEqual(table.col_input, DataTableInputCell(1, 5, False))

            wb.sheet(0).delete_rows(0, 1)  # the row input's row
            table = wb.sheet(0).get_data_table(4, 3)
            self.assertTrue(table.row_input.deleted)
            self.assertFalse(table.col_input.deleted)
            self.assertEqual(table.rect.first_row, 2)
            wb.recalc()
            self.assertIs(wb.sheet(0).get_value(2, 3).kind, ValueKind.ERROR)

    def test_save_load_remove_and_auto_no_table(self) -> None:
        with _squares() as wb:
            wb.sheet(0).add_data_table(CellRect(2, 1, 4, 1), col_input=(0, 1))
            wb.recalc()
            data = wb.save()
            self.assertIn('t="dataTable" ref="B3:B5"', _sheet_xml(data))
            with Workbook.load(data) as loaded:
                table = loaded.sheet(0).get_data_table(4, 1)
                self.assertEqual(table.col_input, DataTableInputCell(0, 1, False))
                with self.assertRaises(FormulonError):
                    loaded.sheet(0).set_number(4, 1, 1.0)

            wb.set_calc_mode(CalcMode.AUTO_NO_TABLE)
            wb.sheet(0).set_number(3, 0, 5.0)
            wb.recalc()
            self.assertEqual(_number(wb, 3, 1), 4.0)  # tables are left alone
            wb.recalc_data_tables()
            self.assertEqual(_number(wb, 3, 1), 25.0)

            wb.sheet(0).remove_data_table(3, 1)
            self.assertIsNone(wb.sheet(0).get_data_table(3, 1))
            self.assertIs(wb.sheet(0).get_value(3, 1).kind, ValueKind.BLANK)
            with self.assertRaises(FormulonError):
                wb.sheet(0).remove_data_table(3, 1)
            wb.sheet(0).set_number(3, 1, 7.0)

    def test_arguments_are_validated(self) -> None:
        with _squares() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_data_table(CellRect(2, 1, 4, 1))
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_data_table(CellRect(0, 1, 4, 1), col_input=(0, 1))
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_data_table(CellRect(2, 1, 4, 1), col_input=(3, 1))
            with self.assertRaises(FormulonError):
                wb.sheet(5).add_data_table(CellRect(2, 1, 4, 1), col_input=(0, 1))
            with self.assertRaises(FormulonError):
                wb.sheet(5).get_data_table(0, 0)
            with self.assertRaises(FormulonError):
                wb.sheet(5).remove_data_table(0, 0)
            with self.assertRaises((TypeError, ValueError, OverflowError)):
                wb.sheet(0).add_data_table(CellRect(-1, 1, 4, 1), col_input=(0, 1))
            wb.sheet(0).add_data_table(CellRect(2, 1, 4, 1), col_input=(0, 1))
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_data_table(CellRect(2, 1, 4, 1), col_input=(0, 1))


if __name__ == "__main__":
    unittest.main()
