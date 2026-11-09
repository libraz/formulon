"""Workbook protection, protected range and note visibility surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    CellRect,
    CfExpressionRuleInput,
    Comment,
    DataValidationInput,
    FormulonError,
    InvalidArgumentError,
    ProtectedRange,
    TableInput,
    Workbook,
    WorkbookProtection,
)


def _full_protection() -> WorkbookProtection:
    return WorkbookProtection(
        lock_structure=True,
        lock_windows=True,
        lock_revision=False,
        workbook_legacy_password_hash="CF1A",
        workbook_algorithm_name="SHA-512",
        workbook_hash_value="aGFzaA==",
        workbook_salt_value="c2FsdA==",
        workbook_spin_count=100000,
        revisions_legacy_password_hash="8D2B",
        revisions_algorithm_name="SHA-256",
        revisions_hash_value="cmV2",
        revisions_salt_value="c2Fs",
        revisions_spin_count=50000,
    )


class WorkbookProtectionTests(unittest.TestCase):
    def test_absent_by_default_and_every_field_round_trips(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.get_workbook_protection(), WorkbookProtection())
            wb.set_workbook_protection(_full_protection())
            self.assertEqual(wb.get_workbook_protection(), _full_protection())
            with Workbook.load(wb.save()) as reloaded:
                self.assertEqual(reloaded.get_workbook_protection(), _full_protection())
                reloaded.set_workbook_protection(WorkbookProtection())
                with Workbook.load(reloaded.save()) as cleared:
                    self.assertEqual(cleared.get_workbook_protection(), WorkbookProtection())


class ProtectedRangeTests(unittest.TestCase):
    def test_add_list_update_remove_and_row_edit_shift(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_protected_range(
                ProtectedRange(
                    name="Inputs", ranges=[CellRect(1, 1, 2, 2), CellRect(4, 4, 4, 4)], legacy_password_hash="CC1A"
                )
            )
            wb.sheet(0).add_protected_range(
                ProtectedRange(
                    name="Totals",
                    ranges=[CellRect(9, 0, 11, 0)],
                    algorithm_name="SHA-512",
                    hash_value="aA==",
                    salt_value="cw==",
                    spin_count=100000,
                    security_descriptor="O:WDG:WD",
                ),
            )
            with Workbook.load(wb.save()) as reloaded:
                ranges = reloaded.sheet(0).list_protected_ranges()
                self.assertEqual(len(reloaded.sheet(0).list_protected_ranges()), 2)
                self.assertEqual([r.name for r in ranges], ["Inputs", "Totals"])
                self.assertEqual(ranges[0].ranges, [CellRect(1, 1, 2, 2), CellRect(4, 4, 4, 4)])
                self.assertEqual(ranges[0].legacy_password_hash, "CC1A")
                self.assertEqual(ranges[1].spin_count, 100000)
                self.assertEqual(ranges[1].security_descriptor, "O:WDG:WD")

                reloaded.sheet(0).insert_rows(0, 1)
                self.assertEqual(
                    reloaded.sheet(0).list_protected_ranges()[0].ranges,
                    [CellRect(3, 1, 4, 2), CellRect(6, 4, 6, 4)],
                )
                reloaded.sheet(0).update_protected_range(0, name="Renamed", ranges=[CellRect(3, 3, 3, 3)])
                renamed = reloaded.sheet(0).list_protected_ranges()[0]
                self.assertEqual((renamed.name, renamed.ranges), ("Renamed", [CellRect(3, 3, 3, 3)]))
                # A partial update keeps the fields it does not name.
                self.assertEqual(renamed.legacy_password_hash, "CC1A")
                reloaded.sheet(0).update_protected_range(0, legacy_password_hash="")
                cleared = reloaded.sheet(0).list_protected_ranges()[0]
                self.assertEqual(
                    (cleared.name, cleared.ranges, cleared.legacy_password_hash),
                    ("Renamed", [CellRect(3, 3, 3, 3)], None),
                )
                reloaded.sheet(0).update_protected_range(0, name="Again")
                self.assertEqual(reloaded.sheet(0).list_protected_ranges()[0].ranges, [CellRect(3, 3, 3, 3)])
                reloaded.sheet(0).remove_protected_range(0)
                self.assertEqual(
                    [r.name for r in reloaded.sheet(0).list_protected_ranges()],
                    ["Totals"],
                )

    def test_arguments_are_validated(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).add_protected_range(ProtectedRange(name="Inputs", ranges=[CellRect(0, 0, 0, 0)]))
            for bad in (
                ProtectedRange(name="Inputs", ranges=[CellRect(0, 1, 0, 1)]),
                ProtectedRange(name="", ranges=[CellRect(0, 1, 0, 1)]),
                ProtectedRange(name="Other", ranges=[]),
            ):
                with self.assertRaises(FormulonError):
                    wb.sheet(0).add_protected_range(bad)
            with self.assertRaises(FormulonError):
                wb.sheet(9).add_protected_range(ProtectedRange(name="Other", ranges=[CellRect(0, 1, 0, 1)]))
            with self.assertRaises(FormulonError):
                wb.sheet(0).update_protected_range(5, name="Other", ranges=[CellRect(0, 1, 0, 1)])
            with self.assertRaises(FormulonError):
                wb.sheet(0).update_protected_range(0, ranges=())
            with self.assertRaises(FormulonError):
                wb.sheet(0).remove_protected_range(5)
            self.assertEqual(len(wb.sheet(0).list_protected_ranges()), 1)


class RangeInputTests(unittest.TestCase):
    def test_a_missing_range_is_a_type_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                DataValidationInput(type=3)  # type: ignore[call-arg]
            with self.assertRaises(TypeError):
                CfExpressionRuleInput()  # type: ignore[call-arg]
            with self.assertRaises(TypeError):
                wb.sheet(0).add_table(TableInput(name="T", columns=["A"]))  # type: ignore[call-arg]
            self.assertEqual(wb.sheet(0).list_tables(), [])

    def test_an_empty_validation_range_list_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0).add_validation(DataValidationInput(type=3, ranges=[], formula1='"A,B"'))
            self.assertEqual(len(wb.sheet(0).list_validations()), 0)


class NoteVisibilityTests(unittest.TestCase):
    def test_visibility_reads_back_and_persists(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_comment(1, 1, Comment("shown", author="Alice"))
            wb.sheet(0).set_comment(3, 3, Comment("hidden", author="Alice"))
            self.assertIs(wb.sheet(0).get_comment_visible(1, 1), False)
            wb.sheet(0).set_comment_visible(1, 1, True)
            with Workbook.load(wb.save()) as reloaded:
                self.assertIs(reloaded.sheet(0).get_comment_visible(1, 1), True)
                self.assertIs(reloaded.sheet(0).get_comment_visible(3, 3), False)
                reloaded.sheet(0).set_comment(1, 1, Comment("edited", author="Alice"))
                self.assertIs(reloaded.sheet(0).get_comment_visible(1, 1), True)
                self.assertIsNone(reloaded.sheet(0).get_comment_visible(7, 7))
                with self.assertRaises(FormulonError):
                    reloaded.sheet(0).set_comment_visible(7, 7, True)

    def test_out_of_range_sheet_raises_instead_of_returning_none(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(99).get_comment_visible(1, 1)


if __name__ == "__main__":
    unittest.main()
