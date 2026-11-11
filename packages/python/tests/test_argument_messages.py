"""Argument-error messages: the ``<Class>.<method> <path> <constraint>`` shape.

Every argument error names the class the caller used, the parameter or keyword (with ``.field`` and ``[i]`` segments
for record members) and a constraint from a fixed vocabulary. Engine refusals use the same public words: no C field
names, no ``formulon:`` prefix, no punctuation after the call.
"""

from __future__ import annotations

import re
import unittest
from typing import Any, Callable

import formulon
from formulon import (
    AnchorKind,
    CellAddress,
    CellRect,
    DefinedNameInput,
    DrawingAnchorInput,
    IndexOutOfRangeError,
    InvalidArgumentError,
    PivotInput,
    PivotNumericGroupInput,
    ShapeInput,
    SparklineGroupInput,
    TableInput,
    Workbook,
)

_BANNED = (
    "formulon:",
    "argument ",
    "`",
    "FM_",
    "fm_",
    "NULL",
    "struct_size",
    "spec->",
    "uint32",
    "int32",
    "4294967295",
    "a object",
    "a array",
)


def _text(error: BaseException) -> str:
    """The message of a failure without the ``[CODE_NAME]`` decoration ``str()`` adds."""
    message = getattr(error, "message", None)
    return message if isinstance(message, str) else str(error)


def _raised(expected: type, call: Callable[[], Any]) -> str:
    try:
        call()
    except expected as error:  # noqa: PERF203 -- the message is the result
        return _text(error)
    raise AssertionError(f"{expected.__name__} was not raised")


class _MessageCase(unittest.TestCase):
    def assert_public_words(self, message: str) -> None:
        for token in _BANNED:
            self.assertNotIn(token, message)
        self.assertNotRegex(message, r"\b\w+_op\b")
        self.assertNotRegex(message, r"^[\w.]+: ")
        # Only the leading `Class.method` names the call; one inside the body is a relabelled word.
        self.assertNotRegex(re.sub(r"^[A-Za-z]\w*\.\w+ ", "", message), r"\b[A-Z]\w*\.[a-z_]\w* ")


class EngineProseMessageTests(_MessageCase):
    def test_a_message_leading_with_no_argument_keeps_its_wording(self) -> None:
        with Workbook.create_default() as wb:
            message = _raised(
                InvalidArgumentError, lambda: wb.set_defined_name(DefinedNameInput(name="A1", formula="1"))
            )
            self.assertEqual(message, "Workbook.set_defined_name defined_name.name is not a valid defined name")
            message = _raised(
                InvalidArgumentError,
                lambda: wb.sheet(0).add_table(TableInput(range="A1:B3", name="Has Space", columns=["a", "b"])),
            )
            self.assertEqual(message, "Sheet.add_table input.name is not a valid table name")
            self.assert_public_words(message)

    def test_a_parameter_name_inside_a_message_body_is_not_relabelled(self) -> None:
        with Workbook.create_default() as wb:
            table = wb.sheet(0).add_table(TableInput(range="A1:B3", name="T", columns=["a", "b"]))
            message = _raised(InvalidArgumentError, lambda: table.insert_rows(2, 1))
            self.assertEqual(message, "Table.insert_rows last 1 is before first 2")
            message = _raised(TypeError, lambda: table.insert_rows(None, 1))
            self.assertEqual(message, "Table.insert_rows last needs first")


class IterativeTests(_MessageCase):
    def test_max_iterations_bounds(self) -> None:
        with Workbook.create_default() as wb:
            for bad in (40000, 0):
                with self.subTest(bad=bad):
                    message = _raised(InvalidArgumentError, lambda: wb.update_iterative(max_iterations=bad))
                    self.assertEqual(message, "Workbook.update_iterative max_iterations must be in 1..32767")
            message = _raised(InvalidArgumentError, lambda: wb.update_iterative(max_iterations=-1))
            self.assertEqual(message, "Workbook.update_iterative max_iterations must not be negative")
            with self.assertRaises(TypeError) as caught:
                wb.update_iterative(max_iterations=1.5)  # type: ignore[arg-type]
            self.assertEqual(str(caught.exception), "Workbook.update_iterative max_iterations must be an integer")
            wb.update_iterative(max_iterations=32767)
            self.assertEqual(wb.get_iterative().max_iterations, 32767)
            wb.update_iterative(max_iterations=1)
            self.assertEqual(wb.get_iterative().max_iterations, 1)


class BooleanTests(_MessageCase):
    def test_int_and_str_refused(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            cases = (
                (lambda: wb.set_date1904(1), "Workbook.set_date1904 date1904 must be a boolean"),
                (lambda: wb.set_calc_on_save("yes"), "Workbook.set_calc_on_save calc_on_save must be a boolean"),
                (lambda: wb.set_full_precision(1), "Workbook.set_full_precision full_precision must be a boolean"),
                (lambda: wb.parse_reference("R1C1", r1c1=1), "Workbook.parse_reference r1c1 must be a boolean"),
                (lambda: wb.format_reference("A1", absolute=1), "Workbook.format_reference absolute must be a boolean"),
                (lambda: ws.update_view(show_grid_lines=1), "Sheet.update_view show_grid_lines must be a boolean"),
            )
            for call, expected in cases:
                with self.subTest(expected=expected):
                    with self.assertRaises(TypeError) as caught:
                        call()
                    self.assertEqual(str(caught.exception), expected)

    def test_none_is_refused_for_a_flag(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError) as caught:
                wb.set_calc_on_save(None)  # type: ignore[arg-type]
            self.assertEqual(str(caught.exception), "Workbook.set_calc_on_save calc_on_save must be a boolean")

    def test_real_booleans_are_accepted(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_date1904(True)
            wb.set_date1904(False)
            wb.set_calc_on_save(True)
            wb.sheet(0).update_view(show_grid_lines=False)
            wb.parse_reference("R1C1", r1c1=True)
            self.assertEqual(wb.format_reference("A1", absolute=True), "$A$1")


class ReferenceNotationTests(_MessageCase):
    """The notation keywords read the same on ``parse_reference``, ``localize_formula`` and ``canonicalize_formula``."""

    def test_a_lone_base_coordinate_is_refused_alike(self) -> None:
        together = "base_row and base_col must be given together"
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError) as caught:
                wb.parse_reference("R1C1", r1c1=True, base_row=1)
            self.assertEqual(str(caught.exception), f"Workbook.parse_reference {together}")
        for name in ("localize_formula", "canonicalize_formula"):
            with self.subTest(name=name):
                with self.assertRaises(TypeError) as caught:
                    getattr(formulon, name)("=1", "mac-365-de_DE", r1c1=True, base_row=1)
                self.assertEqual(str(caught.exception), f"formulon.{name} {together}")
                with self.assertRaises(TypeError) as caught:
                    getattr(formulon, name)("=1", "mac-365-de_DE", r1c1=1)  # type: ignore[arg-type]
                self.assertEqual(str(caught.exception), f"formulon.{name} r1c1 must be a boolean")

    def test_a_base_outside_the_grid_and_an_unknown_profile_are_invalid_arguments(self) -> None:
        with Workbook.create_default() as wb:
            _raised(InvalidArgumentError, lambda: wb.parse_reference("R1C1", r1c1=True, base_row=1048576, base_col=0))
            # A1 does not use the base, but a base outside the grid is refused in either notation.
            _raised(InvalidArgumentError, lambda: wb.parse_reference("A1", base_row=1048576, base_col=0))
        for name in ("localize_formula", "canonicalize_formula"):
            with self.subTest(name=name):
                convert = getattr(formulon, name)
                _raised(
                    InvalidArgumentError,
                    lambda: convert("=1", "mac-365-de_DE", r1c1=True, base_row=1048576, base_col=0),
                )
                _raised(InvalidArgumentError, lambda: convert("=1", "mac-365-de_DE", base_row=1048576, base_col=0))
                _raised(InvalidArgumentError, lambda: convert("=1", "nope", r1c1=True))


class ArgumentMessageTests(_MessageCase):
    def test_message_shape(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            message = _raised(IndexOutOfRangeError, lambda: ws.remove_validation(99))
            self.assertEqual(message, "Sheet.remove_validation validation 99 is out of range (0 items)")
            self.assert_public_words(message)

    def test_engine_bodies_use_public_words(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_text(0, 0, "Product")
            ws.set_text(0, 1, "Amount")
            ws.set_number(1, 1, 10.0)
            table = ws.add_table(TableInput(range=CellRect(0, 0, 1, 1), name="Sales", columns=["Product", "Amount"]))
            shape = ws.add_shape(
                ShapeInput(
                    preset_geometry="rect",
                    anchor=DrawingAnchorInput(
                        kind=AnchorKind.ONE_CELL, from_row=2, width_emu=914400, height_emu=457200
                    ),
                )
            )
            cache = wb.add_pivot_cache()
            pivot = ws.add_pivot(PivotInput(name="Pivot1", cache_id=cache.cache_id, anchor=CellAddress(0, 6)))
            cases = (
                (
                    InvalidArgumentError,
                    lambda: table.update_style_options(style_name=""),
                    "Table.update_style_options style_name must not be empty",
                ),
                (
                    InvalidArgumentError,
                    lambda: shape.update(name=""),
                    "Shape.update name must not be empty",
                ),
                (
                    InvalidArgumentError,
                    lambda: ws.add_sparkline_group(SparklineGroupInput()),
                    "Sheet.add_sparkline_group spec.sparklines needs at least one sparkline",
                ),
                (
                    IndexOutOfRangeError,
                    lambda: pivot.get_data_field(99),
                    "Pivot.get_data_field data_field_index 99 is out of range (0 items)",
                ),
                (
                    InvalidArgumentError,
                    lambda: pivot.set_field_numeric_group(0, PivotNumericGroupInput(interval=1.0, start=float("nan"))),
                    "Pivot.set_field_numeric_group group.start must be finite",
                ),
                (
                    InvalidArgumentError,
                    lambda: wb.update_iterative(max_iterations=40000),
                    "Workbook.update_iterative max_iterations must be in 1..32767",
                ),
            )
            for expected_class, call, expected in cases:
                with self.subTest(expected=expected):
                    message = _raised(expected_class, call)
                    self.assertEqual(message, expected)
                    self.assert_public_words(message)

    def test_range_that_is_neither_rect_nor_a1_string(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            cases = (
                (lambda: ws.get_values(5), "Sheet.get_values range"),
                (lambda: ws.get_typed_values(5), "Sheet.get_typed_values range"),
                (lambda: ws.ooxml.list_cell_xf_indices(5), "SheetOoxml.list_cell_xf_indices range"),
                (lambda: ws.recalc_viewport(5), "Sheet.recalc_viewport viewport"),
            )
            for call, label in cases:
                with self.subTest(label=label):
                    message = _raised(TypeError, call)
                    self.assertEqual(message, f"{label} must be a CellRect or an A1 string")
                    self.assert_public_words(message)


if __name__ == "__main__":
    unittest.main()
