"""Keyed single-record reads, absence as None, and the empty-string refusal."""

from __future__ import annotations

import unittest

from formulon import (
    CellRect,
    CfExpressionRuleInput,
    DataValidationInput,
    DrawingAnchorInput,
    InvalidArgumentError,
    Person,
    ProtectedRangeInput,
    SlicerInput,
    SparklineGroupInput,
    SparklineInput,
    TableColumnTotalsInput,
    TableInput,
    TableTotalsRowFunction,
    ThreadedCommentInput,
    Workbook,
    WorkbookProtectionInput,
)

PERSON = "{11111111-1111-1111-1111-111111111111}"
THREAD = "{22222222-2222-2222-2222-222222222222}"
ANCHOR = DrawingAnchorInput(from_row=4, from_col=3, to_row=14, to_col=5)


def _sales(**table: object) -> Workbook:
    wb = Workbook.create_default()
    wb.sheet(0).set_text(0, 0, "Product")
    wb.sheet(0).set_text(0, 1, "Amount")
    wb.sheet(0).set_number(1, 1, 10.0)
    wb.sheet(0).add_table(TableInput(range=CellRect(0, 0, 1, 1), name="Sales", columns=["Product", "Amount"], **table))
    return wb


class AbsenceTests(unittest.TestCase):
    def test_workbook_protection_none(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIsNone(wb.get_workbook_protection())
            wb.set_workbook_protection(WorkbookProtectionInput())
            self.assertIsNotNone(wb.get_workbook_protection())
            wb.set_workbook_protection(None)
            self.assertIsNone(wb.get_workbook_protection())
            wb.set_workbook_protection(None)
            self.assertIsNone(wb.get_workbook_protection())


class KeyedGetTests(unittest.TestCase):
    def test_get_cf_rule(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            rule_id = ws.add_cf_rule(CfExpressionRuleInput(ranges=[CellRect(0, 0, 2, 0)], formula1="A1>1"))
            self.assertEqual(ws.get_cf_rule(rule_id), ws.list_cf_rules()[0])
            self.assertEqual(ws.get_cf_rule(0), ws.list_cf_rules()[0])
            ws.remove_cf_rule(rule_id)
            self.assertIsNone(ws.get_cf_rule(rule_id))

    def test_get_validation(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            validation_id = ws.add_validation(
                DataValidationInput(type=3, ranges=[CellRect(0, 3, 9, 3)], formula1='"Yes,No"')
            )
            self.assertEqual(ws.get_validation(validation_id), ws.list_validations()[0])
            ws.remove_validation(validation_id)
            self.assertIsNone(ws.get_validation(validation_id))

    def test_get_protected_range(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.add_protected_range(ProtectedRangeInput(name="p", ranges="A1:B2"))
            self.assertEqual(ws.get_protected_range("p"), ws.list_protected_ranges()[0])
            self.assertEqual(ws.get_protected_range(0), ws.list_protected_ranges()[0])
            ws.remove_protected_range("p")
            self.assertIsNone(ws.get_protected_range("p"))

    def test_get_sparkline_group(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            group_id = ws.add_sparkline_group(SparklineGroupInput(sparklines=[SparklineInput("Sheet1!A1:D1", "F1")]))
            self.assertEqual(ws.get_sparkline_group(group_id), ws.list_sparkline_groups()[0])
            ws.remove_sparkline_group(group_id)
            self.assertIsNone(ws.get_sparkline_group(group_id))


class EmptyStringTests(unittest.TestCase):
    def test_style_names(self) -> None:
        with _sales() as wb:
            table = wb.table("Sales")
            with self.assertRaises(InvalidArgumentError):
                table.update_style_options(style_name="")
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0).add_slicer("Sales", 0, SlicerInput(anchor=ANCHOR, style_name=""))
            slicer = wb.sheet(0).add_slicer("Sales", 0, SlicerInput(anchor=ANCHOR))
            with self.assertRaises(InvalidArgumentError):
                slicer.update(style_name="")
            with self.assertRaises(InvalidArgumentError):
                wb.update_table_style_defaults(table_style_name="")
            with self.assertRaises(InvalidArgumentError):
                wb.update_table_style_defaults(pivot_style_name="")
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0).add_table(
                    TableInput(range=CellRect(5, 0, 6, 1), name="Other", columns=["A", "B"], style_name="")
                )
            self.assertEqual(len(wb.sheet(0).list_tables()), 1)
        # An unstyled table is spelled styled=False.
        with _sales(styled=False) as wb:
            self.assertIsNone(wb.table("Sales").get_style_options())
        with self.assertRaises(InvalidArgumentError):
            with _sales(styled=False, style_name="TableStyleLight1"):
                pass

    def test_totals_and_parent_id(self) -> None:
        with _sales() as wb:
            table = wb.table("Sales")
            with self.assertRaises(InvalidArgumentError):
                table.set_column_totals(
                    1, TableColumnTotalsInput(totals_row_function=TableTotalsRowFunction.NONE, totals_row_label="")
                )
            with self.assertRaises(InvalidArgumentError):
                table.set_column_totals(
                    1,
                    TableColumnTotalsInput(totals_row_function=TableTotalsRowFunction.CUSTOM, totals_row_formula=""),
                )
            wb.add_person(Person(id=PERSON, display_name="Alice"))
            ws = wb.sheet(0)
            ws.add_threaded_comment(ThreadedCommentInput(id=THREAD, person_id=PERSON, text="root", cell="A1"))
            with self.assertRaises(InvalidArgumentError):
                ws.add_threaded_comment(ThreadedCommentInput(person_id=PERSON, text="reply", parent_id=""))
            self.assertEqual(len(ws.list_threaded_comments()), 1)


if __name__ == "__main__":
    unittest.main()
