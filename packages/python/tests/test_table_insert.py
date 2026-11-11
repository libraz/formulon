"""``Table.insert_rows`` / ``Table.insert_columns``: the omission forms, where the cells move, column naming,
calculated-column fill and the argument and refusal mapping."""

from __future__ import annotations

import io
import re
import unittest
import zipfile
from typing import List

from formulon import (
    FormulonError,
    IndexOutOfRangeError,
    InvalidArgumentError,
    NotFoundError,
    StatusCode,
    Table,
    TableInput,
    Workbook,
)


def sales_table(wb: Workbook) -> Table:
    """Item, Qty, Double over A1:C3 holding 1/2 and 10/20 in the first two columns."""
    ws = wb.sheet(0)
    for col, text in enumerate(("Item", "Qty", "Double")):
        ws.set_text(0, col, text)
    for row in (1, 2):
        ws.set_number(row, 0, row)
        ws.set_number(row, 1, 10 * row)
        ws.set_number(row, 2, 0)
    return ws.add_table(TableInput(range="A1:C3", name="Sales"))


def names(table: Table) -> List[str]:
    return [column.name for column in table.get_properties().columns]


def last_row(table: Table) -> int:
    return table.get_properties().range.rect.last_row


def last_col(table: Table) -> int:
    return table.get_properties().range.rect.last_col


class InsertRowsTests(unittest.TestCase):
    def test_no_argument_appends_one_row_and_moves_only_the_table_columns(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = sales_table(wb)
            ws.set_number("A4", 99)
            ws.set_number("D4", 77)
            self.assertIsNone(table.insert_rows())
            self.assertEqual(last_row(table), 3)
            self.assertIsNone(ws.get_value("A4"))
            self.assertEqual(ws.get_value("A5"), 99)
            self.assertEqual(ws.get_value("D4"), 77)

    def test_append_goes_above_the_totals_row(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_text("A1", "Item")
            ws.set_text("B1", "Qty")
            for row, (item, qty) in enumerate(((1, 10), (2, 20)), start=1):
                ws.set_number(row, 0, item)
                ws.set_number(row, 1, qty)
            ws.set_text("A4", "Total")
            table = ws.add_table(TableInput(range="A1:B4", name="Sales", totals_row=True))
            table.insert_rows()
            self.assertEqual(last_row(table), 4)
            self.assertTrue(table.get_properties().totals_row)
            self.assertEqual(ws.get_value("A3"), 2)
            self.assertIsNone(ws.get_value("A4"))
            self.assertEqual(ws.get_value("A5"), "Total")

    def test_first_alone_inserts_one_row_and_first_with_last_a_span(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = sales_table(wb)
            table.insert_rows(0)
            self.assertEqual(last_row(table), 3)
            self.assertIsNone(ws.get_value("A2"))
            self.assertEqual(ws.get_value("A3"), 1)
            table.insert_rows(1, 2)
            self.assertEqual(last_row(table), 5)
            self.assertEqual(ws.get_value("A5"), 1)
            self.assertEqual(ws.get_value("A6"), 2)
            table.insert_rows(first=5, last=5)
            self.assertEqual(last_row(table), 6)

    def test_none_arguments_are_omitted(self) -> None:
        with Workbook.create_default() as wb:
            table = sales_table(wb)
            table.insert_rows(None, None)
            table.insert_rows(2, None)
            self.assertEqual(last_row(table), 4)

    def test_calculated_column_fills_the_new_rows(self) -> None:
        with Workbook.create_default() as wb:
            sales_table(wb)
            package = io.BytesIO(wb.save())
        patched = io.BytesIO()
        with zipfile.ZipFile(package) as source, zipfile.ZipFile(patched, "w", zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename == "xl/tables/table1.xml":
                    text = re.sub(
                        r'(<tableColumn [^>]*name="Double"[^>]*?)/>',
                        r"\1><calculatedColumnFormula>B2*2</calculatedColumnFormula></tableColumn>",
                        data.decode("utf-8"),
                    )
                    self.assertIn("calculatedColumnFormula", text)
                    data = text.encode("utf-8")
                target.writestr(item, data)
        with Workbook.load(patched.getvalue()) as wb:
            ws = wb.sheet(0)
            table = wb.table("Sales")
            table.insert_rows(1)
            self.assertEqual(ws.get_formula("C3"), "=B3*2")
            self.assertIsNone(ws.get_formula("A3"))
            self.assertIsNone(ws.get_value("B3"))


class InsertColumnsTests(unittest.TestCase):
    def test_append_and_middle_insert_name_from_the_left_neighbour(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = sales_table(wb)
            ws.set_text("C1", "Q2")
            ws.set_number("A4", 5)
            self.assertIsNone(table.insert_columns())
            self.assertEqual(names(table), ["Item", "Qty", "Q2", "Q3"])
            self.assertEqual(ws.get_value("D1"), "Q3")
            table.insert_columns(1, 2)
            self.assertEqual(names(table), ["Item", "Column1", "Column2", "Qty", "Q2", "Q3"])
            self.assertEqual(last_col(table), 5)
            self.assertEqual(ws.get_value("D2"), 10)
            self.assertEqual(ws.get_value("A4"), 5)

    def test_position_zero_takes_the_profile_word(self) -> None:
        with Workbook.create_for_profile("mac-365-ja_JP") as wb:
            ws = wb.sheet(0)
            table = sales_table(wb)
            table.insert_columns(0)
            self.assertEqual(names(table), ["列1", "Item", "Qty", "Double"])
            self.assertEqual(ws.get_value("A1"), "列1")
            self.assertEqual(ws.get_value("B2"), 1)
            table.insert_columns(1)
            self.assertEqual(names(table), ["列1", "列2", "Item", "Qty", "Double"])


class InsertErrorTests(unittest.TestCase):
    def test_argument_errors(self) -> None:
        with Workbook.create_default() as wb:
            table = sales_table(wb)
            for method in (table.insert_rows, table.insert_columns):
                with self.subTest(method=method.__name__):
                    with self.assertRaises(TypeError):
                        method(1.5)  # type: ignore[arg-type]
                    with self.assertRaises(TypeError):
                        method(0, "1")  # type: ignore[arg-type]
                    for args in ((None, 1), (None, 0)):
                        with self.assertRaises(TypeError):
                            method(*args)
                    for args in ((-1,), (4,), (2**31,)):
                        with self.assertRaises(IndexOutOfRangeError):
                            method(*args)
                    for args in ((2, 1), (0, -1)):
                        with self.assertRaises(InvalidArgumentError):
                            method(*args)
                    name = method.__name__
                    noun = "2 rows" if name == "insert_rows" else "3 columns"
                    with self.assertRaises(TypeError) as caught_first:
                        method(None, 1)
                    self.assertEqual(str(caught_first.exception), f"Table.{name} last needs first")
                    for args, message in (((4,), f"Table.{name} first 4 is out of range ({noun})"),):
                        with self.assertRaises(IndexOutOfRangeError) as caught:
                            method(*args)
                        self.assertEqual(caught.exception.message, message)
                    with self.assertRaises(InvalidArgumentError) as caught_last:
                        method(2, 1)
                    self.assertEqual(caught_last.exception.message, f"Table.{name} last 1 is before first 2")
                    with self.assertRaises(TypeError) as caught_type:
                        method(1.5)  # type: ignore[arg-type]
                    self.assertEqual(str(caught_type.exception), f"Table.{name} first must be an integer")
            self.assertEqual(last_row(table), 2)
            self.assertEqual(last_col(table), 2)

    def test_span_past_the_grid_and_pushed_off_content_are_refused(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = sales_table(wb)
            with self.assertRaises(InvalidArgumentError):
                table.insert_rows(0, 1048576)
            with self.assertRaises(InvalidArgumentError):
                table.insert_columns(0, 16384)
            ws.set_number("B1048576", 1)
            with self.assertRaises(FormulonError) as caught:
                table.insert_rows()
            self.assertEqual(caught.exception.code, StatusCode.PRECONDITION_FAILED)
            self.assertEqual(last_row(table), 2)

    def test_the_handle_of_a_removed_table_raises_not_found(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            table = sales_table(wb)
            ws.remove_table(table)
            with self.assertRaises(NotFoundError):
                table.insert_rows()
            with self.assertRaises(NotFoundError):
                table.insert_columns()


if __name__ == "__main__":
    unittest.main()
