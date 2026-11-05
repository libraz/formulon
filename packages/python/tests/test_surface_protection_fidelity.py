"""Workbook protection, protected range and note visibility surface tests."""

from __future__ import annotations

import unittest

from formulon import (
    CellRect,
    DataValidationInput,
    FormulonError,
    InvalidArgumentError,
    ProtectedRange,
    Workbook,
    WorkbookProtection,
)


def _full_protection() -> WorkbookProtection:
    return WorkbookProtection(
        lock_structure=True,
        lock_windows=True,
        lock_revision=False,
        workbook_password="CF1A",
        workbook_algorithm_name="SHA-512",
        workbook_hash_value="aGFzaA==",
        workbook_salt_value="c2FsdA==",
        workbook_spin_count=100000,
        revisions_password="8D2B",
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
            wb.add_protected_range(
                0, ProtectedRange(name="Inputs", ranges=[CellRect(1, 1, 2, 2), CellRect(4, 4, 4, 4)], password="CC1A")
            )
            wb.add_protected_range(
                0,
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
                ranges = reloaded.get_protected_ranges(0)
                self.assertEqual(reloaded.get_protected_range_count(0), 2)
                self.assertEqual([r.name for r in ranges], ["Inputs", "Totals"])
                self.assertEqual(ranges[0].ranges, [CellRect(1, 1, 2, 2), CellRect(4, 4, 4, 4)])
                self.assertEqual(ranges[0].password, "CC1A")
                self.assertEqual(ranges[1].spin_count, 100000)
                self.assertEqual(ranges[1].security_descriptor, "O:WDG:WD")

                reloaded.insert_rows(0, 0, 2)
                self.assertEqual(
                    reloaded.get_protected_ranges(0)[0].ranges, [CellRect(3, 1, 4, 2), CellRect(6, 4, 6, 4)]
                )
                reloaded.update_protected_range(0, 0, ProtectedRange(name="Renamed", ranges=[CellRect(3, 3, 3, 3)]))
                self.assertEqual(reloaded.get_protected_ranges(0)[0].password, "")
                reloaded.remove_protected_range(0, 0)
                self.assertEqual([r.name for r in reloaded.get_protected_ranges(0)], ["Totals"])

    def test_arguments_are_validated(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_protected_range(0, ProtectedRange(name="Inputs", ranges=[CellRect(0, 0, 0, 0)]))
            for bad in (
                ProtectedRange(name="Inputs", ranges=[CellRect(0, 1, 0, 1)]),
                ProtectedRange(name="", ranges=[CellRect(0, 1, 0, 1)]),
                ProtectedRange(name="Other", ranges=[]),
            ):
                with self.assertRaises(FormulonError):
                    wb.add_protected_range(0, bad)
            with self.assertRaises(FormulonError):
                wb.add_protected_range(9, ProtectedRange(name="Other", ranges=[CellRect(0, 1, 0, 1)]))
            with self.assertRaises(FormulonError):
                wb.update_protected_range(0, 5, ProtectedRange(name="Other", ranges=[CellRect(0, 1, 0, 1)]))
            with self.assertRaises(FormulonError):
                wb.remove_protected_range(0, 5)
            self.assertEqual(wb.get_protected_range_count(0), 1)


class RangeInputTests(unittest.TestCase):
    def test_a_missing_range_is_a_type_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                DataValidationInput(type=3)  # type: ignore[call-arg]
            with self.assertRaises(TypeError):
                wb.add_table(0, None, "T", "T", ["A"])  # type: ignore[arg-type]
            self.assertEqual(list(wb.iter_tables()), [])

    def test_an_empty_validation_range_list_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(InvalidArgumentError):
                wb.add_validation(0, DataValidationInput(type=3, ranges=[], formula1='"A,B"'))
            self.assertEqual(wb.get_validation_count(0), 0)


class NoteVisibilityTests(unittest.TestCase):
    def test_visibility_reads_back_and_persists(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_comment(0, 1, 1, "Alice", "shown")
            wb.set_comment(0, 3, 3, "Alice", "hidden")
            self.assertIs(wb.get_comment_visible(0, 1, 1), False)
            wb.set_comment_visible(0, 1, 1, True)
            with Workbook.load(wb.save()) as reloaded:
                self.assertIs(reloaded.get_comment_visible(0, 1, 1), True)
                self.assertIs(reloaded.get_comment_visible(0, 3, 3), False)
                reloaded.set_comment(0, 1, 1, "Alice", "edited")
                self.assertIs(reloaded.get_comment_visible(0, 1, 1), True)
                self.assertIsNone(reloaded.get_comment_visible(0, 7, 7))
                with self.assertRaises(FormulonError):
                    reloaded.set_comment_visible(0, 7, 7, True)


if __name__ == "__main__":
    unittest.main()
