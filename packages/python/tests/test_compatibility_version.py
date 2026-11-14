"""Workbook.get_compatibility_version / set_compatibility_version: the stored version, its effect on
text-position functions, and the refusals."""

from __future__ import annotations

import unittest

from formulon import CalcMode, CompatibilityVersion, InvalidArgumentError, Workbook


class CompatibilityVersionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_enum_values(self) -> None:
        self.assertEqual(
            {m.name: int(m) for m in CompatibilityVersion},
            {"VERSION_1": 1, "VERSION_2": 2, "VERSION_3": 3},
        )

    def test_new_workbook_reads_version_2(self) -> None:
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_2)

    def test_len_counts_utf16_units_under_version_1_and_code_points_under_version_2(self) -> None:
        self.ws.set_formula("A1", '=LEN("a\U0001f600")')
        self.wb.set_compatibility_version(CompatibilityVersion.VERSION_1)
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_1)
        self.assertEqual(self.ws.get_value("A1"), 3)
        self.wb.set_compatibility_version(CompatibilityVersion.VERSION_2)
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_2)
        self.assertEqual(self.ws.get_value("A1"), 2)

    def test_version_3_is_stored_and_marks_formulas_only_across_version_1(self) -> None:
        self.wb.set_calc_mode(CalcMode.MANUAL)
        self.ws.set_formula("A1", '=LEN("a\U0001f600")')
        self.wb.recalc()
        self.wb.set_compatibility_version(CompatibilityVersion.VERSION_3)
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_3)
        self.assertFalse(self.ws.is_cell_stale("A1"))
        self.wb.set_compatibility_version(CompatibilityVersion.VERSION_1)
        self.assertTrue(self.ws.is_cell_stale("A1"))
        self.wb.recalc()
        self.assertEqual(self.ws.get_value("A1"), 3)
        self.wb.set_compatibility_version(CompatibilityVersion.VERSION_3)
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_3)
        self.assertTrue(self.ws.is_cell_stale("A1"))
        self.wb.recalc()
        self.assertEqual(self.ws.get_value("A1"), 2)

    def test_value_outside_the_enum_is_invalid_argument(self) -> None:
        for bad in (0, 4, -1):
            with self.subTest(bad=bad):
                with self.assertRaises(InvalidArgumentError):
                    self.wb.set_compatibility_version(bad)  # type: ignore[arg-type]
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_2)

    def test_non_integer_argument_is_a_type_error(self) -> None:
        for bad in ("2", None, 1.5, float("nan"), True):
            with self.subTest(bad=bad):
                with self.assertRaises(TypeError):
                    self.wb.set_compatibility_version(bad)  # type: ignore[arg-type]
        self.assertIs(self.wb.get_compatibility_version(), CompatibilityVersion.VERSION_2)


if __name__ == "__main__":
    unittest.main()
