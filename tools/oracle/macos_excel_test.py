#!/usr/bin/env python3
"""Unit tests for Mac Excel formula-assignment verification."""

from __future__ import annotations

import unittest

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
    def __init__(self, anchor):
        self.anchor = anchor

    def range(self, address):
        if address != "Z1":
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


class MacExcelFormulaAssignmentTest(unittest.TestCase):
    def test_retained_formula_does_not_require_byte_identical_readback(self) -> None:
        cell = _FakeCell(formula2_readback="=SUM(1, 2)")

        macos_excel._assign_formula(cell, "=SUM(1,2)", context="case retained")

        self.assertEqual(cell.assigned_formula, "=SUM(1,2)")

    def test_silent_clear_is_reported_with_case_and_formula_context(self) -> None:
        cell = _FakeCell(formula2_readback="", formula_readback="")

        with self.assertRaisesRegex(macos_excel._FormulaRetentionError, r"case rejected.*=BROKEN\(\)"):
            macos_excel._assign_formula(cell, "=BROKEN()", context="case rejected")

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


if __name__ == "__main__":
    unittest.main()
