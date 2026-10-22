"""What-If data table surface tests."""

from __future__ import annotations

import io
import unittest
import zipfile

from formulon import (
    CalcMode,
    DataTable,
    DataTableInput,
    FormulonError,
    ValueKind,
    Workbook,
)


def _sheet_xml(data: bytes) -> str:
    return zipfile.ZipFile(io.BytesIO(data)).read("xl/worksheets/sheet1.xml").decode("utf-8")


def _squares() -> Workbook:
    """B1 = 3 is the column input, B2 = B1 * B1 the formula, A3:A5 = 1..3 the values (results B3:B5)."""
    wb = Workbook.create_default()
    wb.set_number(0, 0, 1, 3.0)
    wb.set_formula(0, 1, 1, "=B1*B1")
    for i in range(3):
        wb.set_number(0, 2 + i, 0, float(i + 1))
    return wb


def _number(wb: Workbook, row: int, col: int) -> float:
    value = wb.get_value(0, row, col)
    assert value.kind is ValueKind.NUMBER, (row, col, value)
    return value.number


class DataTableTests(unittest.TestCase):
    def test_one_variable_table_computes_reads_back_and_refuses_writes(self) -> None:
        with _squares() as wb:
            wb.add_data_table(0, 2, 1, 4, 1, col_input=(0, 1))
            wb.recalc()
            self.assertEqual([_number(wb, r, 1) for r in (2, 3, 4)], [1.0, 4.0, 9.0])
            self.assertEqual(_number(wb, 0, 1), 3.0)
            self.assertEqual(_number(wb, 1, 1), 9.0)

            table = wb.get_data_table(0, 3, 1)
            self.assertEqual(
                table,
                DataTable(
                    first_row=2,
                    first_col=1,
                    last_row=4,
                    last_col=1,
                    two_dimensional=False,
                    row_oriented=False,
                    row_input=None,
                    col_input=DataTableInput(0, 1, False),
                ),
            )
            self.assertIsNone(wb.get_data_table(0, 1, 1))

            with self.assertRaises(FormulonError):
                wb.set_number(0, 3, 1, 7.0)
            with self.assertRaises(FormulonError):
                wb.set_text(0, 2, 1, "x")
            with self.assertRaises(FormulonError):
                wb.set_formula(0, 4, 1, "=1")
            self.assertEqual(_number(wb, 3, 1), 4.0)
            wb.set_number(0, 3, 0, 10.0)
            wb.recalc()
            self.assertEqual(_number(wb, 3, 1), 100.0)

    def test_two_variable_table_and_row_edits(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 5, 1.0)
            wb.set_number(0, 1, 5, 1.0)
            wb.set_formula(0, 2, 2, "=F1*F2")
            for k in range(3):
                wb.set_number(0, 2, 3 + k, float(k + 1))
                wb.set_number(0, 3 + k, 2, float(10 * (k + 1)))
            wb.add_data_table(0, 3, 3, 5, 5, row_input=(0, 5), col_input=(1, 5))
            wb.recalc()
            self.assertEqual([_number(wb, 3 + k, 3 + k) for k in range(3)], [10.0, 40.0, 90.0])
            table = wb.get_data_table(0, 4, 4)
            self.assertTrue(table.two_dimensional)
            self.assertEqual(table.row_input, DataTableInput(0, 5, False))
            self.assertEqual(table.col_input, DataTableInput(1, 5, False))

            wb.delete_rows(0, 0, 1)  # the row input's row
            table = wb.get_data_table(0, 4, 3)
            self.assertTrue(table.row_input.deleted)
            self.assertFalse(table.col_input.deleted)
            self.assertEqual(table.first_row, 2)
            wb.recalc()
            self.assertIs(wb.get_value(0, 2, 3).kind, ValueKind.ERROR)

    def test_save_load_remove_and_auto_no_table(self) -> None:
        with _squares() as wb:
            wb.add_data_table(0, 2, 1, 4, 1, col_input=(0, 1))
            wb.recalc()
            data = wb.save()
            self.assertIn('t="dataTable" ref="B3:B5"', _sheet_xml(data))
            with Workbook.load(data) as loaded:
                table = loaded.get_data_table(0, 4, 1)
                self.assertEqual(table.col_input, DataTableInput(0, 1, False))
                with self.assertRaises(FormulonError):
                    loaded.set_number(0, 4, 1, 1.0)

            wb.set_calc_mode(CalcMode.AUTO_NO_TABLE)
            wb.set_number(0, 3, 0, 5.0)
            wb.recalc()
            self.assertEqual(_number(wb, 3, 1), 4.0)  # tables are left alone
            wb.recalc_data_tables()
            self.assertEqual(_number(wb, 3, 1), 25.0)

            wb.remove_data_table(0, 3, 1)
            self.assertIsNone(wb.get_data_table(0, 3, 1))
            self.assertIs(wb.get_value(0, 3, 1).kind, ValueKind.BLANK)
            with self.assertRaises(FormulonError):
                wb.remove_data_table(0, 3, 1)
            wb.set_number(0, 3, 1, 7.0)

    def test_arguments_are_validated(self) -> None:
        with _squares() as wb:
            with self.assertRaises(FormulonError):
                wb.add_data_table(0, 2, 1, 4, 1)
            with self.assertRaises(FormulonError):
                wb.add_data_table(0, 0, 1, 4, 1, col_input=(0, 1))
            with self.assertRaises(FormulonError):
                wb.add_data_table(0, 2, 1, 4, 1, col_input=(3, 1))
            with self.assertRaises(FormulonError):
                wb.add_data_table(5, 2, 1, 4, 1, col_input=(0, 1))
            with self.assertRaises(FormulonError):
                wb.get_data_table(5, 0, 0)
            with self.assertRaises(FormulonError):
                wb.remove_data_table(5, 0, 0)
            with self.assertRaises((TypeError, ValueError, OverflowError)):
                wb.add_data_table(0, -1, 1, 4, 1, col_input=(0, 1))
            wb.add_data_table(0, 2, 1, 4, 1, col_input=(0, 1))
            with self.assertRaises(FormulonError):
                wb.add_data_table(0, 2, 1, 4, 1, col_input=(0, 1))


if __name__ == "__main__":
    unittest.main()
