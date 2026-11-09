"""Conformance of the high-level style API, ``to_native`` and whole-row / whole-column span strings.

The cases come from the shared tables under ``tools/api/``: ``cell_format_cases.json``, ``to_native_cases.json``
and the ``span_args`` of ``address_args.json``. Each table-driven test fails when its table is empty.
"""

from __future__ import annotations

import copy
import dataclasses
import json
import re
import unittest
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import formulon
from formulon import (
    ErrorCode,
    ExcelError,
    FormulonError,
    InvalidArgumentError,
    Sheet,
    StatusCode,
    Value,
    ValueKind,
    Workbook,
    evaluate_typed_formula,
)

REPO_ROOT = Path(__file__).resolve().parents[3]


def _read_json(relative: str) -> Any:
    with open(REPO_ROOT / relative, encoding="utf-8") as handle:
        return json.load(handle)


# ---------------------------------------------------------------------------
# update_cell_format / get_cell_format: tools/api/cell_format_cases.json
# ---------------------------------------------------------------------------

FORMAT_CASES: List[Dict[str, Any]] = _read_json("tools/api/cell_format_cases.json")["cases"]

# The enum a case names by member, by the (camelCase) key that carries it.
ENUM_KEYS = {
    "underline": formulon.UnderlineStyle,
    "vertAlign": formulon.FontVerticalAlign,
    "pattern": formulon.FillPattern,
    "style": formulon.BorderStyle,
    "horizontal": formulon.HorizontalAlign,
    "vertical": formulon.VerticalAlign,
    "readingOrder": formulon.ReadingOrder,
    "horizontalAlign": formulon.HorizontalAlign,
    "verticalAlign": formulon.VerticalAlign,
}

EXCEPTIONS: Dict[str, type] = {
    "TypeError": TypeError,
    "ValueError": ValueError,
    "InvalidArgumentError": InvalidArgumentError,
}

_ARGB = re.compile(r"^#[0-9A-F]{8}$")


def _snake(name: str) -> str:
    return re.sub(r"(?<!^)(?=[A-Z])", "_", name).lower()


def _member(enum_type: type, name: str) -> Any:
    member = re.sub(r"(?<!^)(?=[A-Z])", "_", name).upper()
    assert member in enum_type.__members__, f"{enum_type.__name__} has no member {member} (from {name})"
    return enum_type[member]


def _convert(value: Any, key: Optional[str] = None) -> Any:
    """A case value in the camelCase / member-name spelling, as the Python surface takes it."""
    if isinstance(value, dict):
        return {_snake(k): _convert(v, k) for k, v in value.items()}
    if isinstance(value, list):
        return [_convert(v, key) for v in value]
    if isinstance(value, str) and key in ENUM_KEYS:
        return _member(ENUM_KEYS[key], value)
    return value


def _argb_text(rgb: int) -> str:
    return f"#{rgb & 0xFFFFFFFF:08X}"


def _assert_subset(test: unittest.TestCase, actual: Any, expected: Any, label: str, key: Optional[str] = None) -> None:
    """``actual`` carries every key ``expected`` names; ``#AARRGGBB`` text is compared with ``.argb``."""
    if isinstance(expected, dict):
        for k, v in expected.items():
            test.assertTrue(hasattr(actual, _snake(k)), f"{label}: no attribute {_snake(k)}")
            _assert_subset(test, getattr(actual, _snake(k)), v, f"{label}.{k}", k)
    elif isinstance(expected, str) and _ARGB.match(expected):
        test.assertEqual(_argb_text(actual.argb), expected, label)
    elif isinstance(expected, str) and key in ENUM_KEYS:
        test.assertIs(actual, _member(ENUM_KEYS[key], expected), label)
    else:
        test.assertEqual(actual, expected, label)
        if isinstance(expected, bool):
            test.assertIsInstance(actual, bool, label)


class CellFormatCaseTests(unittest.TestCase):
    def test_table_is_non_empty_and_covers_successes_and_failures(self) -> None:
        self.assertTrue(FORMAT_CASES, "no cases in cell_format_cases.json")
        self.assertGreaterEqual(len([c for c in FORMAT_CASES if "throws" not in c]), 20)
        self.assertGreaterEqual(len([c for c in FORMAT_CASES if "throws" in c]), 8)
        self.assertEqual(len({c["id"] for c in FORMAT_CASES}), len(FORMAT_CASES), "duplicate case ids")

    def test_every_case(self) -> None:
        self.assertTrue(FORMAT_CASES)
        for case in FORMAT_CASES:
            with self.subTest(case=case["id"]), Workbook.create_default() as wb:
                self._run(wb, case)

    def _run(self, wb: Workbook, case: Dict[str, Any]) -> None:
        ws = wb.sheet(0)
        for step in case.get("before", []):
            ws.update_cell_format(step["range"], **_convert(step["patch"]))
        if "patch_from" in case:
            record = ws.get_cell_format(case["patch_from"])
            patch = {f.name: getattr(record, f.name) for f in dataclasses.fields(record)}
        else:
            patch = _convert(case["patch"])
        cells = {*case.get("unchanged", []), *case.get("expect", {}), "A1"}

        def snapshot() -> Dict[str, Any]:
            return {cell: copy.deepcopy(ws.get_cell_format(cell)) for cell in cells}

        before = snapshot()
        xf_count = len(wb.ooxml.list_cell_xfs())

        if "throws" in case:
            expected = EXCEPTIONS[case["throws"]["py"]]
            with self.assertRaises(expected):
                ws.update_cell_format(case["range"], **patch)
            self.assertEqual(snapshot(), before, "a failed call changed a cell")
            self.assertEqual(len(wb.ooxml.list_cell_xfs()), xf_count, "a failed call added a cell format")
            return

        ws.update_cell_format(case["range"], **patch)
        after = snapshot()
        for cell, fmt in case.get("expect", {}).items():
            _assert_subset(self, ws.get_cell_format(cell), fmt, f"{case['id']} {cell}")
        for cell, fields in case.get("expect_xf", {}).items():
            xf = wb.ooxml.get_cell_xf(ws.ooxml.get_effective_style(cell).xf_index)
            for field, want in fields.items():
                actual = getattr(xf, _snake(field))
                if isinstance(want, str) and field in ENUM_KEYS:
                    want = _member(ENUM_KEYS[field], want)
                self.assertEqual(actual, want, f"{case['id']} {cell} xf.{field}")
        for cell in case.get("unchanged", []):
            self.assertEqual(after[cell], before[cell], f"{cell} changed")
        if case.get("idempotent"):
            count = len(wb.ooxml.list_cell_xfs())
            ws.update_cell_format(case["range"], **patch)
            self.assertEqual(len(wb.ooxml.list_cell_xfs()), count, "a second application added a cell format")


class CellFormatShapeTests(unittest.TestCase):
    def test_a_record_group_is_accepted_like_a_patch(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_cell_format("A1", font={"bold": True}, number_format="0.0")
            record = ws.get_cell_format("A1")
            ws.update_cell_format(
                "C3",
                number_format=record.number_format,
                font=record.font,
                fill=record.fill,
                border=record.border,
                alignment=record.alignment,
                protection=record.protection,
            )
            self.assertEqual(ws.get_cell_format("C3"), record)

    def test_get_takes_coordinates_and_text(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_cell_format("B2", font={"italic": True})
            self.assertEqual(ws.get_cell_format(1, 1), ws.get_cell_format("B2"))
            self.assertTrue(ws.get_cell_format("B2").font.italic)

    def test_rect_takes_a_cell_rect(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_cell_format(formulon.CellRect(0, 0, 1, 1), font={"bold": True})
            self.assertTrue(ws.get_cell_format("B2").font.bold)
            self.assertFalse(ws.get_cell_format("C3").font.bold)

    def test_edge_values(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for size in (1, 409):
                with self.subTest(size=size):
                    ws.update_cell_format("A1", font={"size_pt": size})
                    self.assertEqual(ws.get_cell_format("A1").font.size_pt, size)
            for size in (0, -0.0, -1, 410):
                with self.subTest(size=size), self.assertRaises(InvalidArgumentError):
                    ws.update_cell_format("A1", font={"size_pt": size})
            ws.update_cell_format("A1", alignment={"indent": 0})
            self.assertEqual(ws.get_cell_format("A1").alignment.indent, 0)
            ws.update_cell_format("A1", alignment={"indent": 255, "text_rotation": 180})
            self.assertEqual(ws.get_cell_format("A1").alignment.indent, 255)
            for bad in (float("nan"), float("inf"), -float("inf")):
                with self.subTest(size=bad), self.assertRaises(ValueError):
                    ws.update_cell_format("A1", font={"size_pt": bad})
            for field in ("indent", "text_rotation", "relative_indent"):
                with self.subTest(field=field), self.assertRaises(ValueError):
                    ws.update_cell_format("A1", alignment={field: float("nan")})
            for value in (2**32, 2**31):
                with self.subTest(indent=value), self.assertRaises(InvalidArgumentError):
                    ws.update_cell_format("A1", alignment={"indent": value})
            with self.assertRaises(TypeError):
                ws.update_cell_format("A1", number_format=None)  # type: ignore[arg-type]
            with self.assertRaises(TypeError):
                ws.update_cell_format("A1", font={"bold": 1})  # type: ignore[typeddict-item]

    def test_numbers_that_read_as_bool_or_text_are_type_errors(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for patch in ({"size_pt": True}, {"name": 3}, {"underline": "Single"}):
                with self.subTest(patch=patch), self.assertRaises((TypeError, ValueError)):
                    ws.update_cell_format("A1", font=patch)  # type: ignore[arg-type]

    def test_non_ascii_font_name_and_number_format(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_cell_format("A1", font={"name": "游ゴシック \U0001f600"}, number_format='0.0"円"')
            fmt = ws.get_cell_format("A1")
            self.assertEqual(fmt.font.name, "游ゴシック \U0001f600")
            self.assertEqual(fmt.number_format, '0.0"円"')

    def test_failed_call_leaves_the_sheet_unchanged(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.update_cell_format("A1:B2", font={"bold": True})
            before = ws.get_cell_format("A1")
            with self.assertRaises(InvalidArgumentError):
                ws.update_cell_format("A1:B2", font={"italic": True}, alignment={"indent": 256})
            self.assertEqual(ws.get_cell_format("A1"), before)


# ---------------------------------------------------------------------------
# to_native: tools/api/to_native_cases.json
# ---------------------------------------------------------------------------

NATIVE_CASES: List[Dict[str, Any]] = _read_json("tools/api/to_native_cases.json")["cases"]


class ToNativeCaseTests(unittest.TestCase):
    def test_table_is_non_empty(self) -> None:
        self.assertGreaterEqual(len(NATIVE_CASES), 10)

    def test_every_case(self) -> None:
        self.assertTrue(NATIVE_CASES)
        for index, case in enumerate(NATIVE_CASES):
            label = case.get("formula") or f"cell {case.get('cell')}"
            with self.subTest(case=index, source=label):
                if "formula" in case:
                    value = evaluate_typed_formula(case["formula"])
                else:
                    with Workbook.create_default() as wb:
                        value = wb.sheet(0).get_typed_value(case["cell"])
                native = formulon.to_native(value)
                want = case["native"]
                if isinstance(want, dict) and "error" in want:
                    self.assertIsInstance(native, ExcelError)
                    self.assertEqual(native.code, ErrorCode[want["error"].upper()])
                    self.assertEqual(value.kind, ValueKind.ERROR)
                elif want is None:
                    self.assertIsNone(native)
                elif isinstance(want, bool):
                    self.assertIs(native, want)
                elif isinstance(want, (int, float)):
                    self.assertNotIsInstance(native, bool)
                    self.assertEqual(native, want)
                else:
                    self.assertIsInstance(native, str)
                    self.assertEqual(native, want)

    def test_edge_rows(self) -> None:
        self.assertIsNone(formulon.to_native(Value(kind=ValueKind.BLANK)))
        negative_zero = formulon.to_native(Value(kind=ValueKind.NUMBER, number=-0.0))
        self.assertEqual(str(negative_zero), "-0.0")
        self.assertEqual(str(formulon.to_native(Value(kind=ValueKind.NUMBER, number=0.0))), "0.0")
        self.assertIs(formulon.to_native(Value(kind=ValueKind.BOOL, boolean=False)), False)
        self.assertEqual(formulon.to_native(Value(kind=ValueKind.TEXT, text="")), "")
        astral = "\U0001f600x￿"
        self.assertEqual(formulon.to_native(Value(kind=ValueKind.TEXT, text=astral)), astral)
        array = Value(
            kind=ValueKind.ARRAY, array=((Value(kind=ValueKind.NUMBER, number=1.0), Value(kind=ValueKind.BLANK)),)
        )
        self.assertEqual(formulon.to_native(array), [[1.0, None]])
        error = Value(kind=ValueKind.ERROR, error_code=ErrorCode.NA)
        self.assertEqual(formulon.to_native(error), ExcelError(ErrorCode.NA))


# ---------------------------------------------------------------------------
# Spans: span_args of tools/api/address_args.json
# ---------------------------------------------------------------------------

SPAN_ARGS: Dict[str, List[List[Any]]] = _read_json("tools/api/address_args.json")["py"]["span_args"]


def _xf_of(ws: Sheet) -> int:
    return ws._workbook.ooxml.add_cell_xf(
        formulon.CellXfInput(
            font_index=0,
            fill_index=0,
            border_index=0,
            num_fmt_id=2,
            horizontal=None,
            vertical=None,
            wrap_text=None,
        )
    )


class _Row:
    """How to call one span method: the arguments after the span, a setup, and whether it must change the state."""

    def __init__(
        self,
        rest: Callable[[Sheet], Tuple[Any, ...]] = lambda ws: (),
        pre: Optional[Callable[[Sheet], None]] = None,
    ) -> None:
        self.rest = rest
        self.pre = pre


SPAN_METHODS: Dict[str, _Row] = {
    "Sheet.set_column_width": _Row(lambda ws: (12,)),
    "Sheet.set_column_hidden": _Row(lambda ws: (True,)),
    "Sheet.set_column_outline": _Row(lambda ws: (2,)),
    "Sheet.set_column_collapsed": _Row(lambda ws: (True,)),
    "Sheet.set_column_xf_index": _Row(lambda ws: (_xf_of(ws),)),
    "Sheet.reset_column_xf_index": _Row(pre=lambda ws: ws.ooxml.set_column_xf_index(1, 3, _xf_of(ws))),
    "Sheet.reset_column_width": _Row(pre=lambda ws: ws.set_column_width(1, 3, 12)),
    "Sheet.set_row_height": _Row(lambda ws: (30,)),
    "Sheet.set_row_auto_height": _Row(lambda ws: (30,)),
    "Sheet.set_row_hidden": _Row(lambda ws: (True,)),
    "Sheet.set_row_outline": _Row(lambda ws: (2,)),
    "Sheet.set_row_collapsed": _Row(lambda ws: (True,)),
    "Sheet.set_row_xf_index": _Row(lambda ws: (_xf_of(ws),)),
    "Sheet.reset_row_xf_index": _Row(pre=lambda ws: ws.ooxml.set_row_xf_index(2, 4, _xf_of(ws))),
    "Sheet.reset_row_height": _Row(pre=lambda ws: ws.set_row_height(2, 4, 30)),
    "Sheet.insert_rows": _Row(),
    "Sheet.delete_rows": _Row(),
    "Sheet.insert_columns": _Row(),
    "Sheet.delete_columns": _Row(),
}

# Column operations take B:D (1..3), row operations 3:5 (2..4).
SPAN_FORMS = {
    "col": ("B:D", 1, 3),
    "row": ("3:5", 2, 4),
}


def _state(ws: Sheet) -> Any:
    values = [ws.get_value(r, c) for r in range(8) for c in range(6)]
    return (ws.list_column_overrides(), ws.list_row_overrides(), values)


def _seed(ws: Sheet) -> None:
    for r in range(8):
        for c in range(6):
            ws.set_number(r, c, float(r * 10 + c))


class SpanCaseTests(unittest.TestCase):
    def test_table_is_non_empty_and_every_entry_has_an_argument_row(self) -> None:
        self.assertTrue(SPAN_ARGS, "no span_args in address_args.json")
        self.assertEqual(sorted(set(SPAN_ARGS) - set(SPAN_METHODS)), [], "span methods missing from the argument table")
        self.assertEqual(sorted(set(SPAN_METHODS) - set(SPAN_ARGS)), [], "rows for methods that take no span")

    def test_every_span_method_takes_the_string_like_its_numeric_form(self) -> None:
        self.assertTrue(SPAN_ARGS)
        for key, positions in SPAN_ARGS.items():
            with self.subTest(method=key):
                self.assertEqual(len(positions), 1)
                position, axis, form = positions[0]
                self.assertEqual(position, 0)
                text, first, last = SPAN_FORMS[axis]
                row = SPAN_METHODS[key]
                name = key.split(".")[1]

                def run(as_text: bool) -> Tuple[Any, Any]:
                    with Workbook.create_default() as wb:
                        ws = wb.sheet(0)
                        _seed(ws)
                        if row.pre is not None:
                            row.pre(ws)
                        before = _state(ws)
                        rest = row.rest(ws)
                        if as_text:
                            args: Tuple[Any, ...] = (text, *rest)
                        elif form == "start_count":
                            args = (first, last - first + 1, *rest)
                        else:
                            args = (first, last, *rest)
                        getattr(ws if hasattr(ws, name) else ws.ooxml, name)(*args)
                        return before, _state(ws)

                _, numeric_after = run(False)
                text_before, text_after = run(True)
                self.assertEqual(text_after, numeric_after, f"{key}: the string form differs from the numeric form")
                self.assertNotEqual(text_after, text_before, f"{key} changed nothing")

    def test_edge_rows(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            # "B" and "3" are one column and one row; the grid's last column and row are valid.
            ws.set_column_width("B", 9)
            self.assertEqual([(c.first, c.last) for c in ws.list_column_overrides()], [(1, 1)])
            ws.set_row_height("3", 25)
            self.assertEqual([r.row for r in ws.list_row_overrides()], [2])
            ws.set_column_width("XFD", 9)
            ws.set_row_height("1048576", 25)
            self.assertTrue(any(c.last == 16383 for c in ws.list_column_overrides()))
            self.assertTrue(any(r.row == 1048575 for r in ws.list_row_overrides()))
            # Dollar signs are allowed.
            ws.set_column_hidden("$B:$D", True)
            ws.set_row_hidden("$3:$5", True)
            for bad in ("0", "", "XFE", "1048577", "Ｂ:Ｄ", "Sheet1!B:D", "B:", ":D", "B:3"):
                with self.subTest(column=bad), self.assertRaises(InvalidArgumentError):
                    ws.set_column_width(bad, 9)
            for bad in ("0", "", "1048577", "B", "Ｂ:Ｄ", "Sheet1!3:5", "3:"):
                with self.subTest(row=bad), self.assertRaises(InvalidArgumentError):
                    ws.set_row_height(bad, 9)
            # A column span on a row operation and the reverse.
            with self.assertRaises(InvalidArgumentError):
                ws.set_row_height("B:D", 9)
            with self.assertRaises(InvalidArgumentError):
                ws.set_column_width("3:5", 9)
            with self.assertRaises(InvalidArgumentError):
                ws.insert_rows("B:D")
            with self.assertRaises(InvalidArgumentError):
                ws.delete_columns("3:5")
            # None is not a span; a missing argument is a TypeError too.
            with self.assertRaises(TypeError):
                ws.set_column_width(None, 9)  # type: ignore[call-overload]
            with self.assertRaises(TypeError):
                ws.insert_rows(None)  # type: ignore[call-overload]
            with self.assertRaises(TypeError):
                ws.insert_rows()  # type: ignore[call-overload]
            with self.assertRaises(TypeError):
                ws.set_column_width(1, 3)  # type: ignore[call-overload]
            # A rejected call changed nothing.
            state = _state(ws)
            with self.assertRaises(InvalidArgumentError):
                ws.delete_rows("0")
            self.assertEqual(_state(ws), state)
            # The whole sheet is both a column span and a row span (last: it makes a million row records).
            ws.set_column_width("A:XFD", 10)
            ws.set_row_height("1:1048576", 20)
            self.assertEqual((ws.list_column_overrides()[0].first, ws.list_column_overrides()[-1].last), (0, 16383))

    def test_numeric_zero_is_a_valid_first_index_but_text_zero_is_not(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_column_width(0, 0, 8)
            self.assertEqual([(c.first, c.last) for c in wb.sheet(0).list_column_overrides()], [(0, 0)])
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0).set_column_width("0", 8)

    def test_insert_and_delete_take_the_span_as_start_and_count(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            _seed(ws)
            ws.insert_rows("3:5")
            # Rows 3..5 (1-based) are three new blank rows; the old row 3 moved to row 6.
            self.assertEqual(ws.get_typed_value(2, 0).kind, ValueKind.BLANK)
            self.assertEqual(ws.get_typed_value(4, 0).kind, ValueKind.BLANK)
            self.assertEqual(ws.get_value(5, 0), 20)
            ws.delete_rows("3:5")
            self.assertEqual(ws.get_value(2, 0), 20)
            ws.insert_columns("B")
            self.assertEqual(ws.get_typed_value(0, 1).kind, ValueKind.BLANK)
            self.assertEqual(ws.get_value(0, 2), 1)
            ws.delete_columns("B:B")
            self.assertEqual(ws.get_value(0, 1), 1)

    def test_status_of_a_rejected_span_is_invalid_reference(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError) as caught:
                wb.sheet(0).set_column_width("XFE", 9)
            self.assertEqual(caught.exception.code, StatusCode.PARSER_INVALID_REFERENCE)


if __name__ == "__main__":
    unittest.main()
