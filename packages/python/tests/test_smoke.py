"""Smoke tests for the Formulon Python wheel.

These tests exercise the public surface only -- ``formulon.evaluate_formula``,
``formulon.Workbook``, the ``Value`` POD, and ``FormulonError`` -- and use
``unittest`` so we do not depend on pytest.

Run via ``make python-test``. The tests `import formulon` so they work
identically against the source tree (with PYTHONPATH=packages/python) and
against an installed wheel (after `pip install dist/formulon-*.whl`).
"""

from __future__ import annotations

import threading
import unittest

import formulon
from formulon import (
    CellAddress,
    CellRect,
    Comment,
    DefinedNameInput,
    FormulonError,
    SheetComment,
    TableInput,
    Value,
    ValueKind,
    Workbook,
)
from formulon._c import LIB

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
        # __version__ and the C-ABI get_version() both descend from the
        # same MAJOR.MINOR.PATCH source (pyproject.toml / src/version.h are
        # bumped together by the release skill); they must agree.
        self.assertEqual(formulon.__version__, formulon.get_version())

    def test_library_version_non_empty(self) -> None:
        v = formulon.get_version()
        self.assertIsInstance(v, str)
        self.assertGreater(len(v), 0, f"expected non-empty version, got {v!r}")

    def test_error_display_name_returns_excel_literal(self) -> None:
        self.assertEqual(formulon.ErrorCode(1).display_name, "#DIV/0!")
        with self.assertRaises(ValueError):
            formulon.ErrorCode(999)


class EvalFormulaTests(unittest.TestCase):
    def test_sum_returns_number_six(self) -> None:
        v = formulon.evaluate_typed_formula("=SUM(1,2,3)")
        self.assertEqual(v.kind, ValueKind.NUMBER)
        self.assertEqual(formulon.to_native(v), 6.0)

    def test_division_by_zero_surfaces_as_value(self) -> None:
        v = formulon.evaluate_typed_formula("=1/0")
        self.assertEqual(v.kind, ValueKind.ERROR)
        # to_native returns an ExcelError for ERROR; callers inspect
        # the code rather than catching an exception.
        self.assertEqual(formulon.to_native(v), formulon.ExcelError(formulon.ErrorCode.DIV0))
        self.assertEqual(str(formulon.to_native(v)), "#DIV/0!")
        self.assertEqual(v.error_code, formulon.ErrorCode.DIV0)

    def test_intersection_uses_read_only_evaluation(self) -> None:
        v = formulon.evaluate_typed_formula("=A1 B1")
        self.assertEqual(v.kind, ValueKind.ERROR)
        self.assertEqual(v.error_code, 0)  # ErrorCode::Null / #NULL!

    def test_text_concat(self) -> None:
        v = formulon.evaluate_typed_formula('="hello "&"world"')
        self.assertEqual(v.kind, ValueKind.TEXT)
        self.assertEqual(formulon.to_native(v), "hello world")

    def test_boolean_literal(self) -> None:
        v = formulon.evaluate_typed_formula("=TRUE()")
        self.assertEqual(v.kind, ValueKind.BOOL)
        self.assertIs(formulon.to_native(v), True)


class WorkbookLifecycleTests(unittest.TestCase):
    def test_create_default_has_one_sheet(self) -> None:
        with Workbook.create_default() as wb:
            self.assertFalse(wb.closed)
            self.assertEqual(len(wb), 1)
            self.assertEqual(wb.sheet(0).name, "Sheet1")

    def test_create_empty_has_no_sheets(self) -> None:
        with Workbook.create_empty() as wb:
            self.assertEqual(len(wb), 0)
            wb.add_sheet("S1")
            self.assertEqual(len(wb), 1)
            self.assertEqual(wb.sheet(0).name, "S1")

    def test_unicode_sheet_identity_rejects_duplicate_and_renames_casing(self) -> None:
        with Workbook.create_empty() as wb:
            wb.add_sheet("Ä")
            wb.add_sheet("Ö")
            with self.assertRaises(FormulonError):
                wb.add_sheet("ä")
            with self.assertRaises(FormulonError):
                wb.sheet(1).rename("ä")
            self.assertEqual(wb.sheet(0).name, "Ä")
            self.assertEqual(wb.sheet(1).name, "Ö")
            wb.sheet(0).rename("ä")
            self.assertEqual(wb.sheet(0).name, "ä")
            self.assertEqual(wb.sheet(1).name, "Ö")

    def test_close_is_idempotent(self) -> None:
        wb = Workbook.create_default()
        self.assertFalse(wb.closed)
        wb.close()
        self.assertTrue(wb.closed)
        # Second close must not crash; the C ABI accepts NULL.
        wb.close()
        self.assertTrue(wb.closed)

    def test_method_after_close_raises(self) -> None:
        wb = Workbook.create_default()
        # Leave a real C diagnostic pending without constructing a Python
        # exception from it. Closing is a non-status call and must not clear
        # or expose that stale snapshot on the later closed-handle error.
        status = LIB.fm_workbook_remove_sheet(wb._handle, 99)
        self.assertNotEqual(status, 0)
        wb.close()
        with self.assertRaises(FormulonError) as ctx:
            len(wb)
        error = ctx.exception
        self.assertIsInstance(error, formulon.WorkbookClosedError)
        self.assertEqual(error.code, 7000)
        self.assertEqual(error.message, "the workbook is closed")
        self.assertEqual(error.context, "")
        self.assertNotIn("index=99", str(error))
        self.assertNotIn("sheet_count=1", str(error))

    def test_concurrent_errors_keep_their_diagnostics(self) -> None:
        indices = (99, 100, 101, 102)
        workers = {index: Workbook.create_default() for index in indices}
        try:
            main = Workbook.create_default()
            barrier = threading.Barrier(len(indices) + 1)
            contexts = {index: [] for index in indices}
            unexpected = []

            def fail(index: int) -> None:
                worker = workers[index]
                try:
                    for _ in range(200):
                        barrier.wait()
                        try:
                            worker.sheet(0).remove_table(index)
                        except FormulonError as error:
                            contexts[index].append(error.context)
                        barrier.wait()
                except BaseException as error:  # pragma: no cover - keeps barrier failures observable
                    unexpected.append((index, error))
                    try:
                        barrier.abort()
                    except threading.BrokenBarrierError:
                        pass

            threads = [threading.Thread(target=fail, args=(index,)) for index in indices]
            for thread in threads:
                thread.start()
            try:
                for _ in range(200):
                    barrier.wait()
                    main.sheet(0).get_typed_value(0, 0)
                    barrier.wait()
            finally:
                for thread in threads:
                    thread.join()
            self.assertEqual(unexpected, [])
            for index in indices:
                self.assertEqual(len(contexts[index]), 200)
                expected = f"table_index={index} count=0"
                self.assertEqual(contexts[index], [expected] * 200)
            self.assertEqual(sum(len(values) for values in contexts.values()), 800)
        finally:
            for worker in workers.values():
                worker.close()
            if "main" in locals():
                main.close()

    def test_close_waits_for_real_wasm_fetch_and_decode(self) -> None:
        """A close cannot free the handle between a native fetch and its decode."""
        wb = Workbook.create_default()
        wb.sheet(0).set_text(0, 0, "borrowed")
        decode_started = threading.Event()
        release_decode = threading.Event()
        close_finished = threading.Event()
        destroyed: list[str] = []
        result: list[Value] = []
        errors: list[BaseException] = []
        original_call = LIB._call_or_poison
        original_decode = Value.__dict__["_from_wasm"]

        def record_destroy(fn, name: str, *args):
            if name == "fm_workbook_destroy":
                destroyed.append(name)
            return original_call(fn, name, *args)

        def block_before_decode(cls, ptr: int) -> Value:
            decode_started.set()
            if not release_decode.wait(timeout=5):
                raise AssertionError("test decode was not released")
            return original_decode.__func__(cls, ptr)

        def read_value() -> None:
            try:
                result.append(wb.sheet(0).get_typed_value(0, 0))
            except BaseException as error:  # pragma: no cover - assertion aid
                errors.append(error)

        def close_workbook() -> None:
            try:
                wb.close()
            finally:
                close_finished.set()

        LIB._call_or_poison = record_destroy
        Value._from_wasm = classmethod(block_before_decode)
        reader = threading.Thread(target=read_value)
        closer = threading.Thread(target=close_workbook)
        try:
            reader.start()
            self.assertTrue(decode_started.wait(timeout=5))
            closer.start()
            self.assertFalse(close_finished.wait(timeout=0.1))
            self.assertEqual(destroyed, [])
            release_decode.set()
            reader.join(timeout=5)
            closer.join(timeout=5)
            self.assertFalse(reader.is_alive())
            self.assertFalse(closer.is_alive())
            self.assertEqual(errors, [])
            self.assertEqual(result[0].text, "borrowed")
            self.assertEqual(destroyed, ["fm_workbook_destroy"])
        finally:
            release_decode.set()
            reader.join(timeout=5)
            closer.join(timeout=5)
            LIB._call_or_poison = original_call
            Value._from_wasm = original_decode
            wb.close()


class WorkbookRoundtripTests(unittest.TestCase):
    def test_comment_enumeration_and_cf_formula_evaluation(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_comment(2, 3, Comment("review", author="Ada"))
            self.assertEqual(len(wb.sheet(0).list_comments()), 1)
            self.assertEqual(
                wb.sheet(0).list_comments(), [SheetComment(cell=CellAddress(2, 3), author="Ada", text="review")]
            )
            self.assertTrue(wb.sheet(0).evaluate_cf_formula(0, 0, 0, 0, "=A1=0"))

    def test_paginate_returns_resolved_breaks(self) -> None:
        # Asserts the envelope's shape, not the page geometry: how many rows
        # fit a page is the print engine's business and is pinned by the
        # native print suite against Excel. Row 400 is far enough down to
        # force a break under any plausible body height.
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            wb.sheet(0).set_number(400, 0, 2.0)
            result = wb.sheet(0).get_pagination()
            self.assertGreaterEqual(result.page_count, 2)
            self.assertEqual(result.print_area, [])
            self.assertTrue(result.row_breaks)
            self.assertEqual(
                result.row_breaks,
                sorted(set(result.row_breaks), key=lambda b: b.row),
            )
            self.assertEqual(result.column_breaks, [])

    def test_set_get_number(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 42.0)
            wb.recalc()
            v = wb.sheet(0).get_typed_value(0, 0)
            self.assertEqual(v.kind, ValueKind.NUMBER)
            self.assertEqual(formulon.to_native(v), 42.0)

    def test_set_get_text(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "hello")
            wb.recalc()
            v = wb.sheet(0).get_typed_value(0, 0)
            self.assertEqual(v.kind, ValueKind.TEXT)
            self.assertEqual(formulon.to_native(v), "hello")

    def test_set_get_bool(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_bool(0, 0, True)
            wb.recalc()
            v = wb.sheet(0).get_typed_value(0, 0)
            self.assertEqual(v.kind, ValueKind.BOOL)
            self.assertIs(formulon.to_native(v), True)

    def test_blank_cell_returns_blank_value(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            wb.sheet(0).set_blank(0, 0)
            wb.recalc()
            v = wb.sheet(0).get_typed_value(0, 0)
            self.assertEqual(v.kind, ValueKind.BLANK)
            self.assertIsNone(formulon.to_native(v))

    def test_formula_recalc_chain(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 21.0)
            wb.sheet(0).set_formula(0, 1, "=A1*2")
            wb.recalc()
            v = wb.sheet(0).get_typed_value(0, 1)
            self.assertEqual(v.kind, ValueKind.NUMBER)
            self.assertEqual(formulon.to_native(v), 42.0)

            # Mutate, recalc, observe.
            wb.sheet(0).set_number(0, 0, 100.0)
            wb.recalc()
            v = wb.sheet(0).get_typed_value(0, 1)
            self.assertEqual(formulon.to_native(v), 200.0)


class WorkbookSaveLoadTests(unittest.TestCase):
    def test_save_then_load_preserves_values(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 7.5)
            wb.sheet(0).set_text(0, 1, "world")
            wb.recalc()
            blob = wb.save()
            self.assertIsInstance(blob, bytes)
            # OOXML is a zip archive; first two bytes are "PK".
            self.assertGreater(len(blob), 4)
            self.assertEqual(blob[:2], b"PK")

        with Workbook.load(blob) as wb2:
            wb2.recalc()
            v0 = wb2.sheet(0).get_typed_value(0, 0)
            self.assertEqual(v0.kind, ValueKind.NUMBER)
            self.assertEqual(formulon.to_native(v0), 7.5)
            v1 = wb2.sheet(0).get_typed_value(0, 1)
            self.assertEqual(v1.kind, ValueKind.TEXT)
            self.assertEqual(formulon.to_native(v1), "world")

    def test_load_rejects_garbage(self) -> None:
        # Empty buffer triggers BINDING_NULL_POINTER (ABI contract: len == 0
        # rejects). Non-empty garbage triggers an IO_* parse failure. Both
        # paths must raise FormulonError with non-empty diagnostics.
        with self.assertRaises(FormulonError) as ctx:
            Workbook.load(b"not a real xlsx archive")
        err = ctx.exception
        self.assertIsInstance(err.code, int)
        self.assertNotEqual(err.code, 0)
        self.assertGreater(len(err.code.name), 0)

    def test_value_dataclass_is_frozen(self) -> None:
        v = Value(kind=ValueKind.NUMBER, number=1.0)
        with self.assertRaises(Exception):  # FrozenInstanceError under dataclasses
            v.number = 2.0  # type: ignore[misc]


class IterationTests(unittest.TestCase):
    def test_list_cells_reads_set_cells(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            wb.sheet(0).set_number(1, 1, 2.0)
            wb.sheet(0).set_formula(2, 0, "=A1+B2")
            wb.recalc()
            cells = wb.sheet(0).list_cells()
            # We expect at least the three cells we set.
            self.assertGreaterEqual(len(cells), 3)
            row_col = {(c.row, c.col): c for c in cells}
            self.assertIn((0, 0), row_col)
            self.assertIn((1, 1), row_col)
            self.assertIn((2, 0), row_col)
            # The formula cell should expose a non-None formula string.
            self.assertIsNotNone(row_col[(2, 0)].formula)
            self.assertIn("A1", row_col[(2, 0)].formula or "")

    def test_list_defined_names_empty_workbook(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.list_defined_names(), [])

    def test_list_passthroughs_empty_workbook(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.ooxml.list_passthroughs(), [])


class IteratorHandleLivenessTests(unittest.TestCase):
    """A read after `close()` must raise `FormulonError`, not read the
    now-freed workbook handle."""

    def test_list_cells_after_close_raises(self) -> None:
        wb = Workbook.create_default()
        wb.sheet(0).set_number(0, 0, 1.0)
        wb.sheet(0).set_number(1, 0, 2.0)
        self.assertEqual(len(wb.sheet(0).list_cells()), 2)
        wb.close()
        with self.assertRaises(FormulonError):
            wb.sheet(0).list_cells()

    def test_list_defined_names_after_close_raises(self) -> None:
        wb = Workbook.create_default()
        wb.set_defined_name(DefinedNameInput("NameOne", "=1"))
        wb.set_defined_name(DefinedNameInput("NameTwo", "=2"))
        self.assertEqual(len(wb.list_defined_names()), 2)
        wb.close()
        with self.assertRaises(FormulonError):
            wb.list_defined_names()

    def test_get_table_after_close_raises(self) -> None:
        wb = Workbook.create_default()
        wb.add_sheet("Sheet2")
        wb.sheet(0).set_text(0, 0, "A")
        wb.sheet(1).set_text(0, 0, "A")
        wb.sheet(0).add_table(TableInput(range=CellRect(0, 0, 0, 0), name="TableOne", columns=["A"]))
        wb.sheet(1).add_table(TableInput(range=CellRect(0, 0, 0, 0), name="TableTwo", columns=["A"]))
        wb.table("TableOne")
        wb.sheet(0).list_tables()[0]
        wb.close()
        with self.assertRaises(FormulonError):
            wb.table("TableTwo")
        with self.assertRaises(FormulonError):
            wb.sheet(1).list_tables()[0]

    def test_table_rename_column_rewrites_structured_references(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "Qty")
            wb.sheet(0).set_number(1, 0, 2)
            wb.sheet(0).add_table(TableInput(range=CellRect(0, 0, 1, 0), name="T", columns=["Qty"]))
            wb.sheet(0).set_formula(0, 2, "=SUM(T[Qty])")
            self.assertEqual(wb.table("T").rename_column(0, "Amount"), "Amount")
            self.assertEqual(wb.sheet(0).get_formula(0, 2), "=SUM(T[Amount])")
            self.assertEqual(wb.table("T").rename_column(0, ""), "Column1")
            with self.assertRaises(FormulonError):
                wb.table("T").rename_column(1, "X")


if __name__ == "__main__":
    unittest.main()
