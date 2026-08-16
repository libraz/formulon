#!/usr/bin/env python3
"""Unit tests for Mac Excel formula and workbook-option verification."""

from __future__ import annotations

import os
import unittest
from unittest.mock import patch

from tools.oracle import divergence_check
from tools.oracle.drivers import base, macos_excel, windows_excel


class _FakeCell:
    def __init__(
        self,
        *,
        formula2_readback: object = "=SUM(1,2)",
        formula_readback: object = None,
        formula2_error: Exception | None = None,
        value: object = None,
    ) -> None:
        self._formula2_readback = formula2_readback
        self._formula_readback = formula_readback
        self._formula2_error = formula2_error
        self.value = value
        self.assigned_formula: object = None

    @property
    def formula2(self) -> object:
        if self._formula2_error is not None:
            raise self._formula2_error
        return self._formula2_readback

    @formula2.setter
    def formula2(self, formula: object) -> None:
        self.assigned_formula = formula

    @property
    def formula(self) -> object:
        return self._formula_readback


class _FakeAnchor:
    def __init__(
        self,
        value,
        *,
        address="'[Book.xlsx]Sheet O''Clock'!$Z$1",
        row=1,
        column=26,
        address_error=None,
        offsets=None,
    ):
        self.value = value
        self._address = address
        self.row = row
        self.column = column
        self._address_error = address_error
        self._offsets = offsets or {}
        self.offset_calls = []

    def get_address(self, *, external=False):
        if not external:
            raise AssertionError("the oracle must request an external address")
        if self._address_error is not None:
            raise self._address_error
        return self._address

    def offset(self, row, column):
        self.offset_calls.append((row, column))
        return self._offsets[(row, column)]


class _FakeSheet:
    def __init__(self, anchor, address="Z1"):
        self.anchor = anchor
        self.address = address
        self.range_calls = []

    def range(self, address):
        self.range_calls.append(address)
        if address != self.address:
            raise AssertionError(f"unexpected worksheet access: {address!r}")
        return self.anchor


class _FakeMacApi:
    def __init__(self, value):
        self.value = value
        self.calls = []

    def evaluate(self, **kwargs):
        self.calls.append(kwargs)
        return self.value


class _FakeWinApi:
    def __init__(self, value):
        self.value = value
        self.calls = []

    def Evaluate(self, expression):
        self.calls.append(expression)
        return self.value


class _FakeApp:
    def __init__(self, api):
        self.api = api


class _FakeDate1904Reference:
    def __init__(self, value, *, readback=None, set_error=None, get_error=None):
        self.value = value
        self._readback_follows_set = readback is None
        self.readback = value if self._readback_follows_set else readback
        self.set_error = set_error
        self.get_error = get_error
        self.set_calls = []
        self.get_calls = 0

    def set(self, value):
        self.set_calls.append(value)
        if self.set_error is not None:
            raise self.set_error
        self.value = value
        if self._readback_follows_set:
            self.readback = value

    def get(self):
        self.get_calls += 1
        if self.get_error is not None:
            raise self.get_error
        return self.readback


class _FakeWorkbook:
    def __init__(self, date_1904, sheets=None):
        self.api = type("WorkbookApi", (), {"date_1904": date_1904})()
        self.sheets = [] if sheets is None else sheets
        self.close_calls = 0

    def close(self):
        self.close_calls += 1


class _FakeIterationReference:
    def __init__(self, value=False):
        self.value = value
        self.set_calls = []

    def get(self):
        return self.value

    def set(self, value):
        self.set_calls.append(value)
        self.value = value


class _FakeSuiteApi(_FakeMacApi):
    def __init__(self, value):
        super().__init__(value)
        self.iteration = _FakeIterationReference()


class _FakeBooks:
    def __init__(self, workbook):
        self.workbook = workbook
        self.add_calls = 0

    def add(self):
        self.add_calls += 1
        return self.workbook


class _FakeSuiteApp:
    def __init__(self, workbook):
        self.api = _FakeSuiteApi(16_385)
        self.books = _FakeBooks(workbook)
        self.calculate_calls = 0

    def calculate(self):
        self.calculate_calls += 1


class _FakeWinSuiteApp:
    """Windows counterpart of :class:`_FakeSuiteApp` (COM-style API names)."""

    def __init__(self, workbook):
        self.api = _FakeWinApi(16_385)
        self.books = _FakeBooks(workbook)
        self.calculate_calls = 0

    def calculate(self):
        self.calculate_calls += 1


class MacExcelDate1904Test(unittest.TestCase):
    def test_set_date1904_sets_and_reads_back_the_apple_script_reference(self) -> None:
        date_1904 = _FakeDate1904Reference(False)

        macos_excel._set_date1904(_FakeWorkbook(date_1904), True)

        self.assertEqual(date_1904.set_calls, [True])
        self.assertEqual(date_1904.get_calls, 1)

    def test_set_date1904_rejects_a_mismatched_readback(self) -> None:
        date_1904 = _FakeDate1904Reference(False, readback=False)

        with self.assertRaisesRegex(RuntimeError, r"date_1904=True.*False"):
            macos_excel._set_date1904(_FakeWorkbook(date_1904), True)

        self.assertEqual(date_1904.set_calls, [True])
        self.assertEqual(date_1904.get_calls, 1)

    def test_set_date1904_propagates_setter_errors(self) -> None:
        date_1904 = _FakeDate1904Reference(False, set_error=ValueError("setter failed"))

        with self.assertRaisesRegex(ValueError, "setter failed"):
            macos_excel._set_date1904(_FakeWorkbook(date_1904), True)

        self.assertEqual(date_1904.get_calls, 0)

    def test_set_date1904_propagates_readback_errors(self) -> None:
        date_1904 = _FakeDate1904Reference(False, get_error=RuntimeError("readback failed"))

        with self.assertRaisesRegex(RuntimeError, "readback failed"):
            macos_excel._set_date1904(_FakeWorkbook(date_1904), True)

        self.assertEqual(date_1904.set_calls, [True])
        self.assertEqual(date_1904.get_calls, 1)


class MacExcelDate1904RouteTest(unittest.TestCase):
    def _oracle(self):
        workbook = _FakeWorkbook(_FakeDate1904Reference(False), sheets=[_FakeSheet(_FakeAnchor(1))])
        app = _FakeSuiteApp(workbook)
        oracle = object.__new__(macos_excel.ExcelOracle)
        oracle._app = app
        return oracle, app, workbook

    def test_batch_route_uses_shared_date1904_helper(self) -> None:
        oracle, app, workbook = self._oracle()

        with patch.object(macos_excel, "_set_date1904") as set_date1904:
            self.assertEqual(oracle.run_suite("batch", [], date1904=True), [])

        set_date1904.assert_called_once_with(workbook, True)
        self.assertEqual(app.books.add_calls, 1)
        self.assertEqual(workbook.close_calls, 1)

    def test_per_case_workbook_route_uses_shared_date1904_helper(self) -> None:
        oracle, app, workbook = self._oracle()
        cases = [{"id": "cross-sheet", "formula": "=1", "setup": {}}]

        with patch.object(macos_excel, "_set_date1904") as set_date1904:
            results = oracle._run_suite_per_case_workbook("suite", cases, date1904=True, iterative=False)

        self.assertEqual(results[0].kind, "number")
        set_date1904.assert_called_once_with(workbook, True)
        self.assertEqual(app.books.add_calls, 1)
        self.assertEqual(workbook.close_calls, 1)

    def test_per_case_date1904_failure_propagates_and_closes_workbook(self) -> None:
        oracle, _app, workbook = self._oracle()
        cases = [{"id": "cross-sheet", "formula": "=1", "setup": {}}]

        with patch.object(macos_excel, "_set_date1904", side_effect=RuntimeError("date flag failed")):
            with self.assertRaisesRegex(RuntimeError, "date flag failed"):
                oracle._run_suite_per_case_workbook("suite", cases, date1904=True, iterative=False)

        self.assertEqual(workbook.close_calls, 1)


class MacExcelFormulaAssignmentTest(unittest.TestCase):
    def test_retained_formula_does_not_require_byte_identical_readback(self) -> None:
        cell = _FakeCell(formula2_readback="=SUM(1, 2)")

        macos_excel._assign_formula(cell, "=SUM(1,2)", context="case retained")

        self.assertEqual(cell.assigned_formula, "=SUM(1,2)")

    def test_silent_clear_is_reported_with_case_and_formula_context(self) -> None:
        cell = _FakeCell(formula2_readback="", formula_readback="")

        with self.assertRaisesRegex(macos_excel._FormulaRetentionError, r"case rejected.*=BROKEN\(\)"):
            macos_excel._assign_formula(cell, "=BROKEN()", context="case rejected")

    def test_prefix_readback_is_reported_as_truncation_with_both_lengths(self) -> None:
        sent = "=" + "+".join(f"A{i}" for i in range(1, 151))
        kept = "=" + "+".join(f"A{i}" for i in range(1, 67))
        cell = _FakeCell(formula2_readback=kept)

        with self.assertRaisesRegex(
            macos_excel._FormulaRetentionError,
            rf"only a prefix.*sent {len(sent)} characters, read back {len(kept)}",
        ):
            macos_excel._assign_formula(cell, sent, context="case long chain")

    def test_shorter_readback_that_is_not_a_prefix_is_accepted_as_canonicalisation(self) -> None:
        # Excel may drop insignificant whitespace, which shortens the readback
        # without dropping the tail. That is retention, not truncation.
        cell = _FakeCell(formula2_readback="=SUM(1,2)")

        macos_excel._assign_formula(cell, "=SUM(1, 2)", context="case canonicalised")

        self.assertEqual(cell.assigned_formula, "=SUM(1, 2)")

    def test_retained_formula_can_calculate_to_a_legitimate_blank(self) -> None:
        cell = _FakeCell(formula2_readback='=IF(FALSE,1,"")', value="")

        macos_excel._assign_formula(cell, '=IF(FALSE,1,"")', context="case blank-result")

        self.assertEqual(macos_excel._classify_value(cell).kind, "blank")

    def test_formula_readback_falls_back_when_formula2_is_unavailable(self) -> None:
        cell = _FakeCell(
            formula2_readback="",
            formula_readback="=SUM(1, 2)",
            formula2_error=AttributeError("formula2 is unavailable"),
        )

        macos_excel._assign_formula(cell, "=SUM(1,2)", context="case fallback")

    def test_formula2_assignment_exception_propagates_without_fallback(self) -> None:
        class SetterFailureCell(_FakeCell):
            @property
            def formula2(self) -> object:
                return "=SUM(1,2)"

            @formula2.setter
            def formula2(self, formula: object) -> None:
                raise ValueError("formula2 setter failed")

        with self.assertRaisesRegex(ValueError, "formula2 setter failed"):
            macos_excel._assign_formula(SetterFailureCell(), "=SUM(1,2)", context="case setter-error")


class FormulaPlacementTest(unittest.TestCase):
    """The `formula_cell` override reaches both driver routes."""

    def _oracle(self, sheet):
        workbook = _FakeWorkbook(_FakeDate1904Reference(False), sheets=[sheet])
        app = _FakeSuiteApp(workbook)
        oracle = object.__new__(macos_excel.ExcelOracle)
        oracle._app = app
        return oracle

    def test_absent_override_keeps_the_historical_z1_placement(self) -> None:
        self.assertEqual(base.case_formula_cell({"id": "c"}), "Z1")
        self.assertEqual(base.case_formula_cell({"id": "c", "formula_cell": None}), "Z1")

    def test_override_is_upper_cased_and_trimmed(self) -> None:
        self.assertEqual(base.case_formula_cell({"id": "c", "formula_cell": " aa5 "}), "AA5")

    def test_unusable_override_raises_instead_of_silently_falling_back(self) -> None:
        for value in ("", "   ", 5, ["Z1"]):
            with self.subTest(value=value), self.assertRaisesRegex(ValueError, "unusable formula_cell"):
                base.case_formula_cell({"id": "c", "formula_cell": value})

    def test_batch_route_writes_and_reads_the_declared_cell(self) -> None:
        sheet = _FakeSheet(_FakeAnchor(7), address="AA5")
        oracle = self._oracle(sheet)

        results = oracle.run_suite("placement", [{"id": "p", "formula": "=A:A", "formula_cell": "AA5", "setup": {}}])

        self.assertEqual([(r.kind, r.value) for r in results], [("number", 7.0)])
        self.assertEqual(set(sheet.range_calls), {"AA5"})
        self.assertEqual(sheet.anchor.formula2, "=A:A")

    def test_per_case_workbook_route_writes_and_reads_the_declared_cell(self) -> None:
        sheet = _FakeSheet(_FakeAnchor(7), address="AA5")
        oracle = self._oracle(sheet)
        cases = [{"id": "p", "formula": "=A:A", "formula_cell": "AA5", "setup": {}}]

        results = oracle._run_suite_per_case_workbook("placement", cases, date1904=False, iterative=False)

        self.assertEqual([(r.kind, r.value) for r in results], [("number", 7.0)])
        self.assertIn("AA5", sheet.range_calls)

    def test_windows_batch_route_honours_the_same_override(self) -> None:
        sheet = _FakeSheet(_FakeAnchor(7), address="AA5")
        workbook = _FakeWorkbook(_FakeDate1904Reference(False), sheets=[sheet])
        oracle = object.__new__(windows_excel.WindowsExcelOracle)
        oracle._app = _FakeWinSuiteApp(workbook)

        results = oracle.run_suite("placement", [{"id": "p", "formula": "=A:A", "formula_cell": "AA5", "setup": {}}])

        self.assertEqual([(r.kind, r.value) for r in results], [("number", 7.0)])
        self.assertEqual(set(sheet.range_calls), {"AA5"})

    def test_malformed_override_is_skipped_with_a_reason(self) -> None:
        sheet = _FakeSheet(_FakeAnchor(7), address="AA5")
        oracle = self._oracle(sheet)

        results = oracle.run_suite("placement", [{"id": "p", "formula": "=A:A", "formula_cell": "", "setup": {}}])

        self.assertEqual(results[0].kind, "skipped")
        self.assertIn("unusable formula_cell", results[0].value)


class ShapeCaptureTest(unittest.TestCase):
    """`capture: shape` records shape plus samples instead of every cell."""

    def test_samples_are_none_for_the_default_cell_walk(self) -> None:
        self.assertIsNone(base.case_shape_samples({"id": "c"}))
        self.assertIsNone(base.case_shape_samples({"id": "c", "capture": "cells"}))

    def test_samples_are_upper_cased(self) -> None:
        self.assertEqual(
            base.case_shape_samples({"id": "c", "capture": "shape", "samples": [" z1 ", "aa5"]}),
            ["Z1", "AA5"],
        )

    def test_shape_capture_without_samples_raises(self) -> None:
        for case in (
            {"id": "c", "capture": "shape"},
            {"id": "c", "capture": "shape", "samples": []},
            {"id": "c", "capture": "shape", "samples": ["Z1", 5]},
        ):
            with self.subTest(case=case), self.assertRaisesRegex(ValueError, "samples"):
                base.case_shape_samples(case)

    def test_unknown_capture_mode_raises(self) -> None:
        with self.assertRaisesRegex(ValueError, "unknown capture mode"):
            base.case_shape_samples({"id": "c", "capture": "everything"})

    def test_shape_capture_reads_only_the_declared_cells(self) -> None:
        # A shape far above the cell-walk ceiling: the probe must accept it
        # and the driver must read exactly the sampled addresses.
        anchor = _FakeAnchor(1)
        sheet = _FakeSheet(anchor, address="Z1")
        sheet.anchor = anchor
        cells = {"Z1": _FakeCell(value=1), "Z2": _FakeCell(value=2)}

        def range_(address):
            sheet.range_calls.append(address)
            if address == "Z1":
                return anchor
            return cells[address]

        sheet.range = range_  # type: ignore[method-assign]
        encoded = 1_048_576 * 16_384 + 1
        result = macos_excel._classify_shape_result(_FakeApp(_FakeMacApi(encoded)), sheet, "Z1", ["Z1", "Z2"])

        self.assertEqual(result.kind, "array_shape")
        self.assertEqual(result.array_shape, [1_048_576, 1])
        self.assertEqual(result.value, {"Z1": 1.0, "Z2": 2.0})
        # One probe read plus one read per sample; no walk of the spill.
        self.assertEqual(sheet.range_calls, ["Z1", "Z1", "Z2"])

    def test_cell_walk_still_refuses_a_shape_past_the_ceiling(self) -> None:
        anchor = _FakeAnchor(1)
        with self.assertRaisesRegex(base.SpillShapeProbeError, "capture ceiling"):
            base.probe_spill_shape(lambda _expression: 16_384 + base.MAX_CAPTURE_CELLS + 1, anchor)
        # The same shape is accepted when the caller opts out because it
        # reads a fixed sample list rather than walking.
        self.assertEqual(
            base.probe_spill_shape(lambda _expression: 16_384 + base.MAX_CAPTURE_CELLS + 1, anchor, max_cells=None),
            (1, base.MAX_CAPTURE_CELLS + 1),
        )


class SpillShapeProbeTest(unittest.TestCase):
    def test_decode_one_by_one_two_by_three_and_grid_edge(self) -> None:
        self.assertEqual(base.decode_spill_shape_probe(16_385), (1, 1))
        self.assertEqual(base.decode_spill_shape_probe(2 * 16_384 + 3), (2, 3))
        self.assertEqual(base.decode_spill_shape_probe(1_048_576 * 16_384 + 16_384), (1_048_576, 16_384))

    def test_decode_rejects_bool_nonfinite_fraction_zero_and_out_of_grid(self) -> None:
        for value in (True, float("nan"), float("inf"), 16_385.5, 0, -1, 16_384, 1_048_577 * 16_384 + 1):
            with self.subTest(value=value), self.assertRaises(base.SpillShapeProbeError):
                base.decode_spill_shape_probe(value)

    def test_probe_evaluates_once_and_preserves_external_address(self) -> None:
        anchor = _FakeAnchor(1)
        calls = []

        def evaluate(expression):
            calls.append(expression)
            return 2 * 16_384 + 3

        shape = base.probe_spill_shape(evaluate, anchor)
        self.assertEqual(shape, (2, 3))
        self.assertEqual(len(calls), 1)

    def test_invalid_external_address_is_rejected_before_evaluate(self) -> None:
        cases = (
            {"label": "get_address exception", "address_error": RuntimeError("address unavailable")},
            {"label": "none", "address": None},
            {"label": "empty", "address": ""},
            {"label": "whitespace", "address": "   "},
            {"label": "non-string", "address": 123},
        )
        for case in cases:
            with self.subTest(case=case["label"]):
                anchor = _FakeAnchor(1, **{key: value for key, value in case.items() if key != "label"})
                calls = []
                with self.assertRaises(base.SpillShapeProbeError):
                    base.probe_spill_shape(calls.append, anchor)
                self.assertEqual(calls, [])

    def test_probe_expression_and_evaluator_call_are_exact(self) -> None:
        anchor = _FakeAnchor(1)
        calls = []

        def evaluate(expression):
            calls.append(expression)
            return 2 * 16_384 + 3

        self.assertEqual(base.probe_spill_shape(evaluate, anchor), (2, 3))
        self.assertEqual(
            calls,
            ["ROWS('[Book.xlsx]Sheet O''Clock'!$Z$1#)*16384+COLUMNS('[Book.xlsx]Sheet O''Clock'!$Z$1#)"],
        )

    def test_anchor_edges_accept_exact_fit_and_reject_overflow(self) -> None:
        def evaluate(value):
            return lambda _expression: value

        last_cell = _FakeAnchor(1, row=1_048_576, column=16_384)
        self.assertEqual(base.probe_spill_shape(evaluate(16_385), last_cell), (1, 1))

        inner_edge = _FakeAnchor(1, row=1_048_575, column=16_383)
        self.assertEqual(base.probe_spill_shape(evaluate(2 * 16_384 + 2), inner_edge), (2, 2))

        with self.assertRaisesRegex(base.SpillShapeProbeError, "does not fit"):
            base.probe_spill_shape(evaluate(2 * 16_384 + 1), last_cell)
        with self.assertRaisesRegex(base.SpillShapeProbeError, "does not fit"):
            base.probe_spill_shape(evaluate(16_386), _FakeAnchor(1, row=1, column=16_384))

    def test_anchor_coordinates_are_strict_and_invalid_coordinates_skip_evaluate(self) -> None:
        for name, value in (
            ("row", True),
            ("row", 1.0),
            ("row", 0),
            ("row", 1_048_577),
            ("column", False),
            ("column", 26.0),
            ("column", 0),
            ("column", 16_385),
        ):
            anchor = _FakeAnchor(1)
            setattr(anchor, name, value)
            calls = []
            with self.subTest(name=name, value=value), self.assertRaises(base.SpillShapeProbeError):
                base.probe_spill_shape(calls.append, anchor)
            self.assertEqual(calls, [])

    def test_capture_ceiling_rejects_valid_but_huge_shape_before_offset_loop(self) -> None:
        anchor = _FakeAnchor(1)
        with self.assertRaisesRegex(base.SpillShapeProbeError, "capture ceiling"):
            base.probe_spill_shape(lambda _expression: 16_384 + base.MAX_CAPTURE_CELLS + 1, anchor)

        # The classifier must reject before it asks a worksheet for any
        # offset cell; the fake anchor intentionally has no such cells.
        with self.assertRaisesRegex(base.SpillShapeProbeError, "capture ceiling"):
            macos_excel._classify_result_cell(
                _FakeApp(_FakeMacApi(16_384 + base.MAX_CAPTURE_CELLS + 1)), _FakeSheet(anchor)
            )

    def test_mac_and_windows_adapters_use_their_application_evaluate_shape(self) -> None:
        anchor = _FakeAnchor(1)
        mac_api = _FakeMacApi(16_385)
        win_api = _FakeWinApi(16_385)

        self.assertEqual(macos_excel._evaluate_spill_shape(_FakeApp(mac_api), anchor), (1, 1))
        self.assertEqual(windows_excel._evaluate_spill_shape(_FakeApp(win_api), anchor), (1, 1))
        self.assertEqual(len(mac_api.calls), 1)
        self.assertEqual(list(mac_api.calls[0]), ["name"])
        self.assertEqual(len(win_api.calls), 1)
        self.assertEqual(win_api.calls[0], mac_api.calls[0]["name"])

    def test_scalar_error_and_blank_are_classified_without_sheet_mutation(self) -> None:
        for value, expected in ((7, "number"), ("#VALUE!", "error"), (None, "blank")):
            with self.subTest(value=value):
                anchor = _FakeAnchor(value)
                sheet = _FakeSheet(anchor)
                app = _FakeApp(_FakeMacApi(16_385))
                result = macos_excel._classify_result_cell(app, sheet)
                self.assertEqual(result.kind, expected)

    def test_array_is_row_major_and_does_not_write_clear_or_calculate(self) -> None:
        cells = {
            (0, 0): _FakeCell(value=1),
            (0, 1): _FakeCell(value=2),
            (0, 2): _FakeCell(value=3),
            (1, 0): _FakeCell(value=4),
            (1, 1): _FakeCell(value=5),
            (1, 2): _FakeCell(value=6),
        }
        anchor = _FakeAnchor(1, offsets=cells)
        sheet = _FakeSheet(anchor)
        result = macos_excel._classify_result_cell(_FakeApp(_FakeMacApi(2 * 16_384 + 3)), sheet)
        self.assertEqual(result.kind, "array")
        self.assertEqual(result.array_shape, [2, 3])
        self.assertEqual(result.value, [1.0, 2.0, 3.0, 4.0, 5.0, 6.0])
        self.assertEqual(len(anchor.offset_calls), 6)

    def test_probe_failure_is_a_distinct_error_for_suite_skip(self) -> None:
        anchor = _FakeAnchor(1)

        with self.assertRaises(base.SpillShapeProbeError):
            base.probe_spill_shape(lambda _expression: (_ for _ in ()).throw(RuntimeError("automation")), anchor)

    def test_windows_driver_is_importable_without_xlwings(self) -> None:
        # Importing the Windows adapter must remain possible on Mac/Linux so
        # these fake tests can exercise its COM call shape without xlwings.
        self.assertTrue(hasattr(windows_excel, "_evaluate_spill_shape"))


class _FakeVersionApi:
    def __init__(self, version, build):
        self.Version = version
        self.Build = build


class _FakeVersionApp:
    def __init__(self, version, build):
        self.version = version
        self.api = _FakeVersionApi(version, build)


class WindowsVersionStampTest(unittest.TestCase):
    """The product stamp a capture is promoted (or refused) on.

    `Application.Version` is the bare Office major on every SKU, so the
    build has to reach the stamp -- but `promote_capture` refuses anything
    `divergence_check.is_pending_stamp` rejects, and that allows at most
    `16.<n>.<n>`. Every shape the channels report has to land inside it.
    """

    def _stamp(self, version, build) -> str:
        oracle = object.__new__(windows_excel.WindowsExcelOracle)
        oracle._app = _FakeVersionApp(version, build)
        with patch.object(windows_excel, "detect_locale_from_app", return_value="ja-JP"):
            return oracle.probe_environment().excel_version

    def test_bare_build_is_appended(self) -> None:
        self.assertEqual(self._stamp("16.0", "18025"), "16.0.18025")

    def test_build_carrying_a_sub_build_is_cut_to_three_components(self) -> None:
        # This ja-JP Microsoft 365 install reports Build "20228.0"; the
        # naive join produced "16.0.20228.0", which promotion refused.
        self.assertEqual(self._stamp("16.0", "20228.0"), "16.0.20228")

    def test_whole_dotted_build_replaces_the_version_without_doubling(self) -> None:
        self.assertEqual(self._stamp("16.0", "16.0.20228.20190"), "16.0.20228")

    def test_build_already_present_is_not_appended_twice(self) -> None:
        self.assertEqual(self._stamp("16.0.20228", "20228"), "16.0.20228")

    def test_every_stamp_survives_the_promotion_check(self) -> None:
        for version, build in (("16.0", "18025"), ("16.0", "20228.0"), ("16.0", "16.0.20228.20190")):
            with self.subTest(build=build):
                self.assertFalse(divergence_check.is_pending_stamp(self._stamp(version, build)))


class _FakeAxisTarget:
    """A column / row handle whose size property Excel may or may not take."""

    def __init__(self, default: float, *, stubborn: bool = False, snap: float = 0.0) -> None:
        self._value = default
        self._stubborn = stubborn
        self._snap = snap

    @property
    def ColumnWidth(self) -> float:  # noqa: N802 - COM property name
        return self._value

    @ColumnWidth.setter
    def ColumnWidth(self, value: float) -> None:  # noqa: N802 - COM property name
        if not self._stubborn:
            self._value = value - self._snap

    @property
    def RowHeight(self) -> float:  # noqa: N802 - COM property name
        return self._value

    @RowHeight.setter
    def RowHeight(self, value: float) -> None:  # noqa: N802 - COM property name
        if not self._stubborn:
            self._value = value


class _FakeSizedApi:
    def __init__(self, target: _FakeAxisTarget) -> None:
        self._target = target
        self.column_keys: list = []
        self.row_keys: list = []

    def Columns(self, key):  # noqa: N802 - COM method name
        self.column_keys.append(key)
        return self._target

    def Rows(self, key):  # noqa: N802 - COM method name
        self.row_keys.append(key)
        return self._target


class _FakeSizedSheet:
    def __init__(self, target: _FakeAxisTarget) -> None:
        self.api = _FakeSizedApi(target)


class WindowsAxisSizeTest(unittest.TestCase):
    """Column-width / row-height application for print cases.

    A print case's break counts are a function of these sizes, so a size
    that never reached Excel produces a golden that looks ordinary while
    describing a default-width sheet the case never asked for.
    """

    def test_span_key_is_passed_to_columns_verbatim(self) -> None:
        sheet = _FakeSizedSheet(_FakeAxisTarget(8.43))

        windows_excel._apply_axis_size(sheet, "column", "A:H", 28.0)

        self.assertEqual(sheet.api.column_keys, ["A:H"])
        self.assertEqual(sheet.api.Columns("A:H").ColumnWidth, 28.0)

    def test_row_span_key_is_passed_to_rows_verbatim(self) -> None:
        sheet = _FakeSizedSheet(_FakeAxisTarget(18.75))

        windows_excel._apply_axis_size(sheet, "row", "3:5", 60.0)

        self.assertEqual(sheet.api.row_keys, ["3:5"])
        self.assertEqual(sheet.api.Rows("3:5").RowHeight, 60.0)

    def test_size_excel_did_not_take_is_refused_not_swallowed(self) -> None:
        sheet = _FakeSizedSheet(_FakeAxisTarget(8.43, stubborn=True))

        with self.assertRaisesRegex(RuntimeError, "does not match"):
            windows_excel._apply_axis_size(sheet, "column", "A:H", 28.0)

    def test_snapping_within_tolerance_is_accepted(self) -> None:
        # Excel rounds a size to its own grid; that is not a failed apply.
        sheet = _FakeSizedSheet(_FakeAxisTarget(8.43, snap=0.14))

        windows_excel._apply_axis_size(sheet, "column", "A", 28.0)

        self.assertAlmostEqual(sheet.api.Columns("A").ColumnWidth, 27.86)


class _FakePrinterApi:
    """`Application` as far as the printer pin is concerned."""

    def __init__(self, accepted, current: str = "FUJIFILM Apeos C2360 on Ne01:") -> None:
        self._accepted = set(accepted)
        self._current = current
        self.attempts: list = []

    @property
    def ActivePrinter(self) -> str:  # noqa: N802 - COM property name
        return self._current

    @ActivePrinter.setter
    def ActivePrinter(self, value: str) -> None:  # noqa: N802 - COM property name
        self.attempts.append(value)
        if value not in self._accepted:
            raise RuntimeError(f"Excel rejects {value}")
        self._current = value


class WindowsPrinterPinTest(unittest.TestCase):
    """Which printer's metrics a print capture paginates against.

    Excel derives automatic breaks from the active printer driver, so an
    unpinned capture records whichever device the host happened to
    default to -- and a network device that renegotiates its capabilities
    moves the breaks between two otherwise identical runs.
    """

    def _oracle(self, accepted, **env):
        oracle = object.__new__(windows_excel.WindowsExcelOracle)
        oracle._printer_pinned = False
        oracle._app = type("_App", (), {"api": _FakePrinterApi(accepted)})()
        self._env = patch.dict(os.environ, env, clear=False)
        self._env.start()
        self.addCleanup(self._env.stop)
        if not env:
            os.environ.pop(windows_excel._PRINTER_ENV_VAR, None)
        return oracle

    def test_default_device_is_pinned_when_nothing_is_requested(self) -> None:
        oracle = self._oracle({"Microsoft Print to PDF on Ne00:"})

        oracle._ensure_printer_pinned()

        self.assertEqual(oracle._app.api.ActivePrinter, "Microsoft Print to PDF on Ne00:")

    def test_port_suffix_is_searched_because_it_differs_per_machine(self) -> None:
        oracle = self._oracle({"Microsoft Print to PDF on Ne03:"})

        oracle._ensure_printer_pinned()

        self.assertEqual(oracle._app.api.ActivePrinter, "Microsoft Print to PDF on Ne03:")
        self.assertGreater(len(oracle._app.api.attempts), 1)

    def test_requested_device_excel_rejects_is_an_error(self) -> None:
        oracle = self._oracle(set(), **{windows_excel._PRINTER_ENV_VAR: "No Such Printer"})

        with self.assertRaisesRegex(RuntimeError, "names no printer"):
            oracle._ensure_printer_pinned()

    def test_absent_default_falls_back_to_the_host_device(self) -> None:
        oracle = self._oracle(set())

        oracle._ensure_printer_pinned()

        # No exception: an unset variable is a preference, not a demand.
        self.assertEqual(oracle._app.api.ActivePrinter, "FUJIFILM Apeos C2360 on Ne01:")

    def test_pin_is_attempted_once_per_session(self) -> None:
        oracle = self._oracle({"Microsoft Print to PDF on Ne00:"})

        oracle._ensure_printer_pinned()
        before = len(oracle._app.api.attempts)
        oracle._ensure_printer_pinned()

        self.assertEqual(len(oracle._app.api.attempts), before)


class _FakeBreakCollection:
    """`HPageBreaks` / `VPageBreaks`: a callable with a `.Count`."""

    def __init__(self, locations, axis: str) -> None:
        self._locations = list(locations)
        self._axis = axis

    @property
    def Count(self) -> int:  # noqa: N802 - COM property name
        return len(self._locations)

    def __call__(self, index: int):
        one_based = self._locations[index - 1] + 1
        return type("_Loc", (), {"Location": type("_Cell", (), {self._axis: one_based})()})()


class _FakePaginationSheet:
    """A worksheet whose successive break reads follow a script."""

    def __init__(self, readings) -> None:
        self._readings = list(readings)
        self._index = 0
        self.settles = 0

    def _current(self):
        return self._readings[min(self._index, len(self._readings) - 1)]

    @property
    def HPageBreaks(self):  # noqa: N802 - COM property name
        return _FakeBreakCollection(self._current()[0], "Row")

    @property
    def VPageBreaks(self):  # noqa: N802 - COM property name
        return _FakeBreakCollection(self._current()[1], "Column")

    def Calculate(self) -> None:  # noqa: N802 - COM method name
        # The settle between reads is what advances the script.
        self.settles += 1
        self._index += 1


class _FakePagesPageSetup:
    def __init__(self, sheet: _FakePaginationSheet) -> None:
        self._sheet = sheet

    @property
    def Pages(self):  # noqa: N802 - COM property name
        return type("_Pages", (), {"Count": self._sheet._current()[2]})()


class _FakePaginationApp:
    class _Api:
        ScreenUpdating = True

        @staticmethod
        def ExecuteExcel4Macro(_expression):  # noqa: N802 - COM method name
            raise RuntimeError("unused")

    api = _Api()


class _FakePaginationWorkbook:
    app = type("_App", (), {"api": _FakePaginationApp._Api()})()


class WindowsPaginationSettleTest(unittest.TestCase):
    """The guard against a stale page-break read.

    Excel answers the break collections from its last layout pass, which
    is not always the current case's: a reading that does not reproduce
    would otherwise be written into a golden that looks ordinary.
    """

    def setUp(self) -> None:
        # The real settle waits for Excel's layout thread; the fakes here
        # answer instantly, so the wait is only dead time.
        patcher = patch.object(windows_excel, "_LAYOUT_SETTLE_SECONDS", 0.0)
        patcher.start()
        self.addCleanup(patcher.stop)

    def _read(self, readings):
        sheet = _FakePaginationSheet(readings)
        value = windows_excel._read_pagination_settled(sheet, _FakePagesPageSetup(sheet), _FakePaginationWorkbook())
        return value, sheet

    def test_stable_case_is_read_twice_and_returned(self) -> None:
        value, sheet = self._read([([39], [], 2)])

        self.assertEqual(value, ([39], [], 2))
        self.assertEqual(sheet.settles, 1)

    def test_stale_first_reading_is_discarded_for_the_one_that_repeats(self) -> None:
        # The observed shape: the preceding case's break, then this case's.
        value, _sheet = self._read([([26], [], 2), ([39], [], 2), ([39], [], 2)])

        self.assertEqual(value, ([39], [], 2))

    def test_reading_that_never_repeats_is_refused(self) -> None:
        never_settles = [([i], [], 2) for i in range(windows_excel._BREAK_READ_ATTEMPTS + 2)]

        with self.assertRaisesRegex(RuntimeError, "did not settle"):
            self._read(never_settles)

    def test_break_locations_are_zero_based_and_sorted(self) -> None:
        value, _sheet = self._read([([39, 12], [7, 2], 6)])

        self.assertEqual(value, ([12, 39], [2, 7], 6))


if __name__ == "__main__":
    unittest.main()
