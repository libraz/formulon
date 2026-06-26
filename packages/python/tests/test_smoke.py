# Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
"""Smoke tests for the Formulon Python wheel.

These tests exercise the public surface only -- ``formulon.eval_formula``,
``formulon.Workbook``, the ``Value`` POD, and ``FormulonError`` -- and use
``unittest`` so we do not depend on pytest.

Run via ``make python-test``. The tests `import formulon` so they work
identically against the source tree (with PYTHONPATH=packages/python) and
against an installed wheel (after `pip install dist/formulon-*.whl`).
"""

from __future__ import annotations

import unittest

import formulon
from formulon import FormulonError, Value, ValueKind, Workbook


# Excel error code ordinals, mirrored from src/value.h ``ErrorCode``.
ERROR_DIV0 = 1
ERROR_VALUE = 2
ERROR_REF = 3
ERROR_NAME = 4
ERROR_NUM = 5
ERROR_NA = 6
ERROR_NULL = 0


class VersionTests(unittest.TestCase):
    def test_package_version_is_non_empty(self) -> None:
        # __version__ is derived from the single source of truth (the
        # pyproject [project].version, via dist metadata or a source-tree
        # fallback) -- never hardcoded -- so we only assert its shape here.
        self.assertIsInstance(formulon.__version__, str)
        self.assertRegex(formulon.__version__, r"^\d+\.\d+\.\d+")

    def test_package_version_matches_library_version(self) -> None:
        # __version__ and the C-ABI library_version() both descend from the
        # same MAJOR.MINOR.PATCH source (pyproject.toml / src/version.h are
        # bumped together by the release skill); they must agree.
        self.assertEqual(formulon.__version__, formulon.library_version())

    def test_library_version_non_empty(self) -> None:
        v = formulon.library_version()
        self.assertIsInstance(v, str)
        self.assertGreater(len(v), 0, f"expected non-empty version, got {v!r}")

    def test_version_string_alias(self) -> None:
        # Backward-compat alias matching the npm binding's name.
        self.assertEqual(formulon.version_string(), formulon.library_version())


class EvalFormulaTests(unittest.TestCase):
    def test_sum_returns_number_six(self) -> None:
        v = formulon.eval_formula("=SUM(1,2,3)")
        self.assertEqual(v.kind, ValueKind.NUMBER)
        self.assertEqual(v.to_python(), 6.0)

    def test_division_by_zero_surfaces_as_value(self) -> None:
        v = formulon.eval_formula("=1/0")
        self.assertEqual(v.kind, ValueKind.ERROR)
        # to_python returns the Value itself for ERROR; callers inspect
        # the error_code rather than catching an exception.
        self.assertIs(v.to_python(), v)
        self.assertIsNotNone(v.error_code)

    def test_text_concat(self) -> None:
        v = formulon.eval_formula('="hello "&"world"')
        self.assertEqual(v.kind, ValueKind.TEXT)
        self.assertEqual(v.to_python(), "hello world")

    def test_boolean_literal(self) -> None:
        v = formulon.eval_formula("=TRUE()")
        self.assertEqual(v.kind, ValueKind.BOOL)
        self.assertIs(v.to_python(), True)


class WorkbookLifecycleTests(unittest.TestCase):
    def test_create_default_has_one_sheet(self) -> None:
        with Workbook.create_default() as wb:
            self.assertTrue(wb.is_valid)
            self.assertEqual(wb.sheet_count(), 1)
            self.assertEqual(wb.sheet_name(0), "Sheet1")

    def test_create_empty_has_no_sheets(self) -> None:
        with Workbook.create_empty() as wb:
            self.assertEqual(wb.sheet_count(), 0)
            wb.add_sheet("S1")
            self.assertEqual(wb.sheet_count(), 1)
            self.assertEqual(wb.sheet_name(0), "S1")

    def test_close_is_idempotent(self) -> None:
        wb = Workbook.create_default()
        self.assertTrue(wb.is_valid)
        wb.close()
        self.assertFalse(wb.is_valid)
        # Second close must not crash; the C ABI accepts NULL.
        wb.close()
        self.assertFalse(wb.is_valid)

    def test_method_after_close_raises(self) -> None:
        wb = Workbook.create_default()
        wb.close()
        with self.assertRaises(FormulonError):
            wb.sheet_count()


class WorkbookRoundtripTests(unittest.TestCase):
    def test_set_get_number(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 42.0)
            wb.recalc()
            v = wb.get_value(0, 0, 0)
            self.assertEqual(v.kind, ValueKind.NUMBER)
            self.assertEqual(v.to_python(), 42.0)

    def test_set_get_text(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_text(0, 0, 0, "hello")
            wb.recalc()
            v = wb.get_value(0, 0, 0)
            self.assertEqual(v.kind, ValueKind.TEXT)
            self.assertEqual(v.to_python(), "hello")

    def test_set_get_bool(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_bool(0, 0, 0, True)
            wb.recalc()
            v = wb.get_value(0, 0, 0)
            self.assertEqual(v.kind, ValueKind.BOOL)
            self.assertIs(v.to_python(), True)

    def test_blank_cell_returns_blank_value(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 1.0)
            wb.set_blank(0, 0, 0)
            wb.recalc()
            v = wb.get_value(0, 0, 0)
            self.assertEqual(v.kind, ValueKind.BLANK)
            self.assertIsNone(v.to_python())

    def test_formula_recalc_chain(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 21.0)
            wb.set_formula(0, 0, 1, "=A1*2")
            wb.recalc()
            v = wb.get_value(0, 0, 1)
            self.assertEqual(v.kind, ValueKind.NUMBER)
            self.assertEqual(v.to_python(), 42.0)

            # Mutate, recalc, observe.
            wb.set_number(0, 0, 0, 100.0)
            wb.recalc()
            v = wb.get_value(0, 0, 1)
            self.assertEqual(v.to_python(), 200.0)


class WorkbookSaveLoadTests(unittest.TestCase):
    def test_save_then_load_preserves_values(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 7.5)
            wb.set_text(0, 0, 1, "world")
            wb.recalc()
            blob = wb.save()
            self.assertIsInstance(blob, bytes)
            # OOXML is a zip archive; first two bytes are "PK".
            self.assertGreater(len(blob), 4)
            self.assertEqual(blob[:2], b"PK")

        with Workbook.load(blob) as wb2:
            wb2.recalc()
            v0 = wb2.get_value(0, 0, 0)
            self.assertEqual(v0.kind, ValueKind.NUMBER)
            self.assertEqual(v0.to_python(), 7.5)
            v1 = wb2.get_value(0, 0, 1)
            self.assertEqual(v1.kind, ValueKind.TEXT)
            self.assertEqual(v1.to_python(), "world")

    def test_load_rejects_garbage(self) -> None:
        # Empty buffer triggers kBindingNullPointer (ABI contract: len == 0
        # rejects). Non-empty garbage triggers a kIo* parse failure. Both
        # paths must raise FormulonError with non-empty diagnostics.
        with self.assertRaises(FormulonError) as ctx:
            Workbook.load(b"not a real xlsx archive")
        err = ctx.exception
        self.assertIsInstance(err.status, int)
        self.assertNotEqual(err.status, 0)
        self.assertGreater(
            len(err.status_name), 0, "fm_status_string returned empty"
        )

    def test_value_dataclass_is_frozen(self) -> None:
        v = Value(kind=ValueKind.NUMBER, number=1.0)
        with self.assertRaises(Exception):  # FrozenInstanceError under dataclasses
            v.number = 2.0  # type: ignore[misc]


class IterationTests(unittest.TestCase):
    def test_iter_cells_yields_set_cells(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 1.0)
            wb.set_number(0, 1, 1, 2.0)
            wb.set_formula(0, 2, 0, "=A1+B2")
            wb.recalc()
            cells = list(wb.iter_cells(0))
            # We expect at least the three cells we set.
            self.assertGreaterEqual(len(cells), 3)
            row_col = {(c.row, c.col): c for c in cells}
            self.assertIn((0, 0), row_col)
            self.assertIn((1, 1), row_col)
            self.assertIn((2, 0), row_col)
            # The formula cell should expose a non-None formula string.
            self.assertIsNotNone(row_col[(2, 0)].formula)
            self.assertIn("A1", row_col[(2, 0)].formula or "")

    def test_iter_defined_names_empty_workbook(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(list(wb.iter_defined_names()), [])

    def test_iter_passthrough_empty_workbook(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(list(wb.iter_passthrough()), [])


if __name__ == "__main__":
    unittest.main()
