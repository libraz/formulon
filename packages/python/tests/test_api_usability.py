"""Public API usability conformance: A1 arguments, bulk reads and writes, patches, refs, lists, save, parity of shapes.

One ``TestCase`` per documented success criterion. The coordinate-pair and rectangle argument sets come from
``tools/api/address_args.json`` (``py.cell_args`` / ``py.rect_args``); the per-method argument table below says how to
call each of them, and a method missing from the table (or an empty set) fails.
"""

from __future__ import annotations

import dataclasses
import datetime
import inspect
import io
import json
import os
import re
import sys
import tempfile
import typing
import unittest
import zipfile
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

import _chart_input

import formulon
from formulon import (
    UNCHANGED,
    AlreadyExistsError,
    CellAddress,
    CellRect,
    ErrorCode,
    ExcelError,
    FormulaSyntaxError,
    FormulonError,
    IndexOutOfRangeError,
    InvalidArgumentError,
    NotFoundError,
    PhoneticProperties,
    PivotInput,
    Sheet,
    Value,
    ValueKind,
    Workbook,
)

REPO_ROOT = Path(__file__).resolve().parents[3]
ADDRESS_ARGS = REPO_ROOT / "tools" / "api" / "address_args.json"

BLANK = Value(kind=ValueKind.BLANK)


def _number(x: float) -> Value:
    return Value(kind=ValueKind.NUMBER, number=float(x))


def _open() -> Workbook:
    return Workbook.create_default()


# ---------------------------------------------------------------------------
# Criteria 1 and 2: the argument table
# ---------------------------------------------------------------------------


def _pivot_book(wb: Workbook, ws: Sheet) -> None:
    """A cache with two fields and three records, and one pivot named ``Pivot1`` at E1."""
    cache = wb.add_pivot_cache().cache_id
    wb.pivot_cache(cache).add_field("Region")
    wb.pivot_cache(cache).add_field("Amount")
    for region, amount in (("East", 10.0), ("East", 20.0), ("West", 30.0)):
        rec = wb.pivot_cache(cache).add_record()
        wb.pivot_cache(cache).set_record_value(rec, 0, region)
        wb.pivot_cache(cache).set_record_value(rec, 1, amount)
    ws.add_pivot(PivotInput(name="Pivot1", cache_id=cache, anchor=CellAddress(0, 4)))


def _pivot_report_book(wb: Workbook, ws: Sheet) -> None:
    """A worksheet-sourced pivot ``Pivot1`` at E1 with Region on rows and Sum of Amount; F2 is its first value cell."""
    ws.set_text(0, 0, "Region")
    ws.set_text(0, 1, "Amount")
    cache = wb.add_pivot_cache().cache_id
    wb.pivot_cache(cache).set_worksheet_source(formulon.PivotWorksheetSourceInput(range="A1:B3", sheet="Sheet1"))
    wb.pivot_cache(cache).add_field("Region")
    wb.pivot_cache(cache).add_field("Amount")
    for row, (region, amount) in enumerate((("East", 10.0), ("West", 30.0)), start=1):
        ws.set_text(row, 0, region)
        ws.set_number(row, 1, amount)
        rec = wb.pivot_cache(cache).add_record()
        wb.pivot_cache(cache).set_record_value(rec, 0, region)
        wb.pivot_cache(cache).set_record_value(rec, 1, amount)
    pivot = ws.add_pivot(PivotInput(name="Pivot1", cache_id=cache, anchor=CellAddress(0, 4)))
    ws.pivot(pivot).set_field_axis("Region", formulon.PivotAxis.ROW)
    ws.pivot(pivot).set_field_axis("Amount", formulon.PivotAxis.VALUE)
    amount_field = ws.pivot(pivot).get_field("Amount").cache_field_index
    ws.pivot(pivot).add_data_field(formulon.PivotDataFieldInput(name="Sum of Amount", field_index=amount_field))


def _text_cell(wb: Workbook, ws: Sheet) -> None:
    ws.set_text(2, 1, "text")


def _merged(wb: Workbook, ws: Sheet) -> None:
    ws.add_merge(CellRect(1, 1, 2, 2))


def _linked(wb: Workbook, ws: Sheet) -> None:
    ws.set_hyperlink(formulon.HyperlinkInput(range=CellRect(1, 1, 2, 2), target="https://example.com"))


class Case:
    """How to call one method: ``lead`` before the address, ``tail`` after it, ``kw`` as keywords."""

    def __init__(
        self,
        lead: Tuple[Any, ...] = (),
        tail: Callable[[], Tuple[Any, ...]] = lambda: (),
        kw: Optional[Callable[[], Dict[str, Any]]] = None,
        setup: Callable[[Workbook, Sheet], None] = _text_cell,
        cell: Tuple[int, int] = (2, 1),
        recv: Callable[[Sheet], Any] = lambda ws: ws,
    ) -> None:
        self.recv = recv
        self.lead = lead
        self.tail = tail
        self.kw = kw or (lambda: {})
        self.setup = setup
        self.cell = cell


def _runs() -> Tuple[Any, ...]:
    return ([formulon.PhoneticRun(0, 1, "kana")],)


# Methods that no longer take a coordinate pair or whose pair moved into a record are listed as ``None``: the
# conformance loop skips them and ``test_non_coordinate_methods`` pins their new shape instead.
CELL_TABLE: Dict[str, Optional[Case]] = {
    "Sheet.add_pivot": None,
    "Sheet.reset_cell_xf_index": Case(),
    "Sheet.evaluate_cf_formula": Case(tail=lambda: ("=TRUE",)),
    "Sheet.evaluate_formula": Case(tail=lambda: ("=A1+1",)),
    "Sheet.evaluate_formula_array": Case(tail=lambda: ("=1+1",)),
    "Sheet.evaluate_typed_formula": Case(tail=lambda: ("=A1+1",)),
    "Sheet.evaluate_typed_formula_array": Case(tail=lambda: ("=1+1",)),
    "Sheet.get_array_formula": Case(),
    "Sheet.get_cell_format": Case(),
    "Sheet.get_cell_phonetic": Case(),
    "Sheet.get_cell_phonetic_properties": Case(),
    "Sheet.get_cell_phonetic_runs": Case(),
    "Sheet.get_cell_show_phonetic": Case(),
    "Sheet.get_cell_xf_index": Case(),
    "Sheet.get_comment": Case(),
    "Sheet.get_comment_visible": Case(),
    "Sheet.get_data_table": Case(),
    "Sheet.get_date": Case(),
    "Sheet.get_dependents": Case(),
    "Sheet.get_display_text": Case(),
    "Sheet.get_effective_style": Case(),
    "Sheet.get_formula": Case(),
    "Sheet.get_formula_r1c1": Case(),
    "Sheet.get_hyperlink": Case(),
    "Sheet.get_lambda_text": Case(),
    "Sheet.get_precedents": Case(),
    "Sheet.get_spill": Case(),
    "Sheet.get_value": Case(),
    "Sheet.get_typed_value": Case(),
    "Sheet.remove_data_table": Case(),
    "Sheet.set_blank": Case(),
    "Sheet.set_bool": Case(tail=lambda: (True,)),
    "Sheet.set_cell_phonetic": Case(tail=lambda: ("kana",)),
    "Sheet.set_cell_phonetic_properties": Case(
        tail=lambda: (formulon.PhoneticProperties(font_index=0, type=1, alignment=1),)
    ),
    "Sheet.set_cell_phonetic_runs": Case(tail=_runs),
    "Sheet.set_cell_show_phonetic": Case(tail=lambda: (True,)),
    "Sheet.set_cell_xf_index": Case(tail=lambda: (0,)),
    "Sheet.set_comment": Case(tail=lambda: (formulon.Comment("note", author="Ada"),)),
    "Sheet.set_comment_visible": Case(tail=lambda: (True,)),
    "Sheet.set_error": Case(tail=lambda: (ErrorCode.NA,)),
    "Sheet.set_formula": Case(tail=lambda: ("=1+1",)),
    "Sheet.set_number": Case(tail=lambda: (5.0,)),
    "Sheet.set_pane_top_left": None,
    "Sheet.set_text": Case(tail=lambda: ("hello",)),
    "Sheet.set_top_left": None,
    "Sheet.set_value": Case(tail=lambda: (5.0,)),
    "Sheet.set_values": Case(tail=lambda: ([[5.0, "x"]],)),
    "Pivot.show_details": Case(setup=_pivot_report_book, cell=(1, 5), recv=lambda ws: ws.pivot("Pivot1")),
    "Sheet.validate_value": Case(tail=lambda: (_number(1),)),
}

RECT_TABLE: Dict[str, Case] = {
    "Sheet.add_merge": Case(setup=lambda wb, ws: None),
    "Sheet.delete_cells": Case(tail=lambda: (formulon.CellShift.UP,)),
    "Sheet.insert_cells": Case(tail=lambda: (formulon.CellShift.DOWN,)),
    "Sheet.evaluate_cf_range": Case(),
    "Sheet.get_cell_rect_pt": Case(),
    "Sheet.get_cells_in_range": Case(),
    "Sheet.get_merges_in_range": Case(setup=_merged),
    "Sheet.iter_rows": Case(),
    "Sheet.list_cells": Case(),
    "Sheet.get_range_xf_bases": Case(),
    "Sheet.get_values": Case(),
    "Sheet.get_typed_values": Case(),
    "Sheet.get_range_xf_bases_excluding": Case(tail=lambda: ([(1, 1)],)),
    "Sheet.list_cell_xf_indices": Case(),
    "Sheet.list_show_phonetic_cells": Case(),
    "Sheet.recalc_viewport": Case(),
    "Sheet.remap_range_xfs": Case(tail=lambda: ({0: 0},)),
    "Sheet.remove_hyperlink": Case(setup=_linked),
    "Sheet.remove_merge": Case(setup=_merged),
    "Pivot.set_anchor": Case(setup=_pivot_book, recv=lambda ws: ws.pivot("Pivot1")),
    "Sheet.update_cell_format": Case(kw=lambda: {"number_format": "0.00"}),
    "Sheet.set_print_area": Case(setup=lambda wb, ws: None),
    "Sheet.set_range_xf_index": Case(tail=lambda: (0,)),
    "Sheet.snapshot_range_xfs": Case(),
    "Workbook.format_reference": Case(recv=lambda ws: ws.workbook),
}

_CLASSES: Dict[str, type] = {"Sheet": Sheet, "Pivot": formulon.Pivot}
_HANDLE_CLASSES = (
    formulon.Table,
    formulon.Pivot,
    formulon.PivotCache,
    formulon.Drawing,
    formulon.Picture,
    formulon.Chart,
    formulon.Shape,
    formulon.Slicer,
)

# These take a sequence of rectangles; the A1 form is a list of strings.
RECT_SEQUENCE = {"Sheet.set_print_area"}


def _address_args() -> Dict[str, Any]:
    with open(ADDRESS_ARGS, encoding="utf-8") as handle:
        return json.load(handle)["py"]


def _a1(row: int, col: int) -> str:
    letters = ""
    n = col + 1
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row + 1}"


def _member(recv: Any, name: str) -> Callable[..., Any]:
    """The bound method ``name`` of ``recv``, or of its ``ooxml`` layer when the facade keeps it there."""
    return getattr(recv if hasattr(recv, name) else recv.ooxml, name)


def _has_member(cls: type, name: str) -> bool:
    return hasattr(cls, name) or hasattr(_OOXML_LAYERS.get(cls), name)


_OOXML_LAYERS: Dict[type, type] = {Sheet: formulon.SheetOoxml, Workbook: formulon.WorkbookOoxml}


def _norm(result: Any) -> Any:
    if isinstance(result, Sheet):
        return ("sheet", result.name)
    if inspect.isgenerator(result):
        return list(result)
    return result


def _state(ws: Sheet) -> Any:
    cells = tuple(ws.get_value(r, c) for r in range(6) for c in range(6))
    return (cells, ws.list_hyperlinks(), ws.list_comments(), ws.list_merges(), ws.get_cell_format(1, 1))


def _outcome(case: Case, call: Callable[[Any], Any]) -> Tuple[Any, ...]:
    with _open() as wb:
        ws = wb.sheet(0)
        case.setup(wb, ws)
        try:
            result = call(case.recv(ws))
        except TypeError:
            return ("type-error",)
        except FormulonError as exc:
            return ("error", type(exc).__name__)
        return ("ok", _norm(result), _state(ws))


class CellArgumentsTests(unittest.TestCase):
    """Criterion 1: every coordinate-pair method also takes A1 text in place of each pair."""

    def test_set_is_not_empty_and_covered_by_the_table(self) -> None:
        cell_args = _address_args()["cell_args"]
        self.assertTrue(cell_args, "address_args.json lists no coordinate-pair methods")
        missing = sorted(set(cell_args) - set(CELL_TABLE))
        self.assertEqual(missing, [], "methods without an argument-table entry")

    def test_every_method_takes_both_forms(self) -> None:
        cell_args = _address_args()["cell_args"]
        self.assertTrue(cell_args)
        for key, positions in sorted(cell_args.items()):
            case = CELL_TABLE.get(key)
            if case is None:
                continue
            cls, name = key.split(".")
            with self.subTest(method=key):
                self.assertTrue(_has_member(_CLASSES[cls], name), f"{key} is missing")
                self.assertEqual(positions[0][0], len(case.lead), "the table's lead does not match the ledger position")
                pairs = [case.cell] if len(positions) == 1 else [(1, 1), (0, 0)][: len(positions)]
                coords: List[Any] = []
                texts: List[Any] = []
                for r, c in pairs:
                    coords += [r, c]
                    texts.append(_a1(r, c))
                coordinate = _outcome(
                    case, lambda recv: _member(recv, name)(*case.lead, *coords, *case.tail(), **case.kw())
                )
                self.assertNotEqual(coordinate, ("type-error",), "the coordinate form must be accepted by the table")
                spelled = _outcome(
                    case, lambda recv: _member(recv, name)(*case.lead, *texts, *case.tail(), **case.kw())
                )
                self.assertEqual(spelled, coordinate)

    def test_a1_spellings_equal_coordinates(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.set_number(0, 0, 7.0)
            for text in ("A1", "a1", "$A$1", "A01"):
                with self.subTest(text=text):
                    self.assertEqual(ws.get_value(text), ws.get_value(0, 0))
            self.assertEqual(ws.get_typed_value("XFD1048576"), BLANK)

    def test_rejected_texts(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            for text in ("Sheet1!A1", "A1:B2"):
                with self.subTest(text=text), self.assertRaises(InvalidArgumentError):
                    ws.get_value(text)
            for text in ("", "R1C1", "Ａ1", "XFE1"):
                with self.subTest(text=text), self.assertRaises(FormulaSyntaxError):
                    ws.get_value(text)

    def test_set_number_with_text(self) -> None:
        with _open() as wb:
            wb.sheet(0).set_number("A1", 5)
            self.assertEqual(wb.sheet(0).get_typed_value(0, 0), _number(5))

    def test_two_pairs_do_not_mix(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            self.assertEqual(ws.evaluate_cf_formula("B2", "A1", "=TRUE"), ws.evaluate_cf_formula(1, 1, 0, 0, "=TRUE"))
            with self.assertRaises(TypeError):
                ws.evaluate_cf_formula("B2", 0, 0, "=TRUE")  # type: ignore[call-overload]

    def test_pivot_details_forms(self) -> None:
        # Each call inserts a detail sheet, so every form runs on its own workbook.
        def details(call: Callable[[Sheet], Sheet]) -> Tuple[int, List[List[Value]]]:
            with _open() as wb:
                ws = wb.sheet(0)
                _pivot_report_book(wb, ws)
                sheet = call(ws)
                return sheet.index, sheet.get_typed_values("A1:B3")

        expected = details(lambda ws: ws.pivot(0).show_details(1, 5))
        self.assertEqual(expected[1][1][1].number, 10.0)
        self.assertEqual(details(lambda ws: ws.pivot(0).show_details("F2")), expected)
        self.assertEqual(details(lambda ws: ws.pivot("Pivot1").show_details(1, 5)), expected)
        self.assertEqual(details(lambda ws: ws.pivot("Pivot1").show_details("F2")), expected)

    def test_set_value_edges(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.set_number(0, 0, 3.0)
            ws.set_value(0, 0, None)
            self.assertEqual(ws.get_typed_value(0, 0), BLANK)
            for args in ((0, 0), (0,), ("A1",)):
                with self.subTest(args=args), self.assertRaises(TypeError):
                    ws.set_value(*args)  # type: ignore[call-overload]
            with self.assertRaises(TypeError):
                ws.set_value(row=0, col=0, value=1)  # type: ignore[call-overload]
            ws.set_value("A1", value=1)
            self.assertEqual(ws.get_typed_value(0, 0), _number(1))

    def test_signature_of_set_value(self) -> None:
        sig = inspect.signature(Sheet.set_value)
        self.assertEqual(list(sig.parameters), ["self", "row", "col", "value"])
        self.assertTrue(all(p.kind is inspect.Parameter.POSITIONAL_OR_KEYWORD for p in sig.parameters.values()))

    @unittest.skipIf(sys.version_info < (3, 11), "typing.get_overloads needs Python 3.11")
    def test_set_formula_overloads(self) -> None:
        stubs = typing.get_overloads(Sheet.set_formula)  # type: ignore[attr-defined]
        self.assertEqual(len(stubs), 2)
        a1_forms = [inspect.signature(s) for s in stubs if "a1" in inspect.signature(s).parameters]
        self.assertEqual(len(a1_forms), 1)
        self.assertEqual(list(a1_forms[0].parameters)[:3], ["self", "a1", "formula"])

    def test_non_coordinate_methods(self) -> None:
        """Methods dropped from the coordinate set: the view setters are gone and ``add_pivot`` takes one record."""
        for name in ("set_top_left", "set_pane_top_left"):
            self.assertFalse(hasattr(Sheet, name), name)
        params = inspect.signature(Sheet.add_pivot).parameters
        self.assertEqual([p for p in params if p != "self"], ["input"])
        link = inspect.signature(Sheet.set_hyperlink).parameters
        self.assertEqual([p for p in link if p != "self"], ["link"])


class RectArgumentsTests(unittest.TestCase):
    """Criterion 2: every rectangle method also takes an A1 range text."""

    def test_set_is_not_empty_and_covered_by_the_table(self) -> None:
        rect_args = _address_args()["rect_args"]
        self.assertTrue(rect_args, "address_args.json lists no rectangle methods")
        self.assertEqual(sorted(set(rect_args) - set(RECT_TABLE)), [])

    def test_every_method_takes_both_forms(self) -> None:
        rect_args = _address_args()["rect_args"]
        self.assertTrue(rect_args)
        for key, positions in sorted(rect_args.items()):
            case = RECT_TABLE[key]
            name = key.split(".")[1]
            with self.subTest(method=key):
                self.assertEqual(positions[0], len(case.lead), "the table's lead does not match the ledger position")
                rect = CellRect(1, 1, 2, 2)
                sequence = key in RECT_SEQUENCE
                as_rect: Any = [rect] if sequence else rect
                as_text: Any = ["B2:C3"] if sequence else "B2:C3"
                typed = _outcome(case, lambda recv: _member(recv, name)(*case.lead, as_rect, *case.tail(), **case.kw()))
                self.assertNotEqual(typed, ("type-error",))
                spelled = _outcome(
                    case, lambda recv: _member(recv, name)(*case.lead, as_text, *case.tail(), **case.kw())
                )
                self.assertEqual(spelled, typed)

    def test_merge_texts(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.add_merge("A1:B2")
            self.assertEqual(ws.list_merges(), [CellRect(0, 0, 1, 1)])
            ws.add_merge("D4")
            self.assertIn(CellRect(3, 3, 3, 3), ws.list_merges())


# ---------------------------------------------------------------------------
# Criterion 3: subscripts replace get/set
# ---------------------------------------------------------------------------


class SubscriptTests(unittest.TestCase):
    def test_get_and_set_are_gone(self) -> None:
        self.assertFalse(hasattr(Sheet, "get"))
        self.assertFalse(hasattr(Sheet, "set"))

    def test_subscript_reads_and_writes_native_values(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws[0, 0] = 1
            ws["A2"] = "x"
            ws["A3"] = True
            ws["A4"] = None
            self.assertEqual(ws["A1"], 1.0)
            self.assertEqual(ws[0, 0], 1.0)
            self.assertEqual(ws["A2"], "x")
            self.assertIs(ws[2, 0], True)
            self.assertIsNone(ws["A4"])
            ws["B1"] = "=1+1"
            self.assertEqual(ws["B1"], 2.0)
            self.assertEqual(ws.get_formula("B1"), "=1+1")

    def test_range_key_is_a_grid(self) -> None:
        with _open() as wb:
            self.assertEqual(wb.sheet(0)["A1:B2"], [[None, None], [None, None]])
            with self.assertRaises(TypeError):
                wb.sheet(0)["A1:B2"] = 1
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0)["A1:B2"] = [[1]]

    def test_int_and_slice_keys_are_type_errors(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            for key in (0, slice(0, 2), (slice(0, 2), 0)):
                with self.subTest(key=key):
                    with self.assertRaises(TypeError):
                        ws[key]  # type: ignore[index]
                    with self.assertRaises(TypeError):
                        ws[key] = 1  # type: ignore[index]


# ---------------------------------------------------------------------------
# Criterion 4: bulk reads and writes
# ---------------------------------------------------------------------------


class BulkValuesTests(unittest.TestCase):
    def test_blank_range_is_dense(self) -> None:
        with _open() as wb:
            self.assertEqual(wb.sheet(0).get_typed_values("A1:B2"), [[BLANK, BLANK], [BLANK, BLANK]])
            self.assertEqual(wb.sheet(0).get_values(CellRect(0, 0, 1, 1)), [[None, None], [None, None]])

    def test_range_over_one_page(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            for row in (0, 65535, 65536, 69999):
                ws.set_number(row, 0, float(row + 1))
            values = ws.get_values("A1:A70000")
            self.assertEqual(len(values), 70000)
            for row in (0, 65535, 65536, 69999):
                self.assertEqual(values[row], [float(row + 1)])
            self.assertEqual(values[1], [None])
            self.assertEqual(values[65537], [None])

    def test_spill_values_are_read(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.set_formula(0, 0, "=SEQUENCE(3)")
            wb.recalc()
            self.assertEqual(ws.get_typed_values("A1:A3"), [[_number(1)], [_number(2)], [_number(3)]])

    def test_whole_row_range_is_rejected(self) -> None:
        with _open() as wb:
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0).get_values("A:XFD")

    def test_limit_of_two_pow_22_cells_is_accepted(self) -> None:
        with _open() as wb:
            values = wb.sheet(0).get_values("A1:D1048576")
            self.assertEqual(len(values), 1048576)
            self.assertEqual(len(values[-1]), 4)

    def test_set_values_writes_four_cells(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.set_number(1, 1, 9.0)
            ws.set_values("A1", [[1, "x"], [True, None]])
            self.assertEqual(
                ws.get_values("A1:B2"),
                [[1.0, "x"], [True, None]],
            )
            ws.set_values(3, 0, [[2.5]])
            self.assertEqual(ws.get_value(3, 0), 2.5)

    def test_invalid_input_writes_nothing(self) -> None:
        bad_inputs: List[Tuple[str, Any, Any]] = [
            ("A1", [[1], [1, 2]], (TypeError, InvalidArgumentError)),
            ("A1", [[float("nan")]], (TypeError, InvalidArgumentError)),
            ("A1", [[float("inf")]], (TypeError, InvalidArgumentError)),
            ("A1", [[1, object()]], TypeError),
            (
                "A1",
                [[Value(kind=ValueKind.ARRAY, array=((Value(kind=ValueKind.NUMBER, number=1.0),),))]],
                (TypeError, InvalidArgumentError),
            ),
            ("XFD1", [[1, 2]], InvalidArgumentError),
        ]
        for anchor, rows, expected in bad_inputs:
            with self.subTest(anchor=anchor, rows=repr(rows)), _open() as wb:
                ws = wb.sheet(0)
                ws.set_text(0, 0, "keep")
                with self.assertRaises(expected):
                    ws.set_values(anchor, rows)
                self.assertEqual(ws.get_value(0, 0), "keep")
                self.assertIsNone(ws.get_value(0, 16383))

    def test_empty_values_succeed(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.set_number(0, 0, 4.0)
            ws.set_values("A1", [])
            ws.set_values("A1", [[]])
            self.assertEqual(ws.get_typed_value(0, 0), _number(4))

    def test_equals_prefix_is_a_formula_unless_the_value_is_typed_text(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.set_values("A1", [["=1+1", Value(kind=ValueKind.TEXT, text="=1+1")]])
            self.assertEqual(ws.get_formula(0, 0), "=1+1")
            self.assertEqual(ws.get_value(0, 0), 2.0)
            self.assertEqual(ws.get_typed_value(0, 1), Value(kind=ValueKind.TEXT, text="=1+1"))
            self.assertIsNone(ws.get_formula(0, 1))

    def test_overlap_with_data_table_keeps_earlier_cells(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.set_number(0, 0, 1.0)
            ws.set_formula(1, 1, "=A1*2")
            ws.set_number(2, 0, 3.0)
            ws.set_number(3, 0, 4.0)
            # B3:B4 is the data table's result range, which refuses writes.
            ws.add_data_table(formulon.DataTableInput(CellRect(2, 1, 3, 1), column_input=formulon.CellAddress(0, 0)))
            with self.assertRaises(FormulonError):
                ws.set_values("A3", [[7, 8]])
            self.assertEqual(ws.get_typed_value(2, 0), _number(7))
            self.assertIsNotNone(ws.get_data_table(2, 1))


# ---------------------------------------------------------------------------
# Criterion 5: has_dirty_cells
# ---------------------------------------------------------------------------


class DirtyCellsTests(unittest.TestCase):
    def test_clean_states(self) -> None:
        with _open() as wb:
            self.assertFalse(wb.has_dirty_cells())
            wb.sheet(0).set_number(0, 0, 1.0)
            wb.sheet(0).set_formula(0, 1, "=A1+1")
            wb.recalc()
            self.assertFalse(wb.has_dirty_cells())
        with _open() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            with Workbook.load(wb.save()) as loaded:
                self.assertFalse(loaded.has_dirty_cells())

    def test_loaded_formulas_are_dirty_until_the_first_recalc(self) -> None:
        with _open() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            wb.sheet(0).set_formula(0, 1, "=A1+1")
            wb.recalc()
            with Workbook.load(wb.save()) as loaded:
                self.assertTrue(loaded.has_dirty_cells())
                loaded.recalc()
                self.assertFalse(loaded.has_dirty_cells())

    def test_edits_make_the_workbook_dirty(self) -> None:
        def formula_book() -> Workbook:
            wb = _open()
            wb.sheet(0).set_number(0, 0, 1.0)
            wb.sheet(0).set_formula(0, 1, "=A1+1")
            wb.recalc()
            return wb

        def set_number(wb: Workbook) -> None:
            wb.sheet(0).set_number(0, 0, 2.0)

        def set_name(wb: Workbook) -> None:
            wb.sheet(0).set_formula(1, 1, "=Rate*2")
            wb.recalc()
            wb.set_defined_name(formulon.DefinedNameInput("Rate", "Sheet1!$A$1"))

        def insert_rows(wb: Workbook) -> None:
            wb.sheet(0).insert_rows(0, 1)

        def remove_sheet(wb: Workbook) -> None:
            wb.add_sheet("Other")
            wb.recalc()
            wb.remove_sheet(1)

        def other_profile(wb: Workbook) -> None:
            wb.set_profile_id("mac-365-ja_JP")

        for label, edit in (
            ("set_number", set_number),
            ("defined name", set_name),
            ("insert_rows", insert_rows),
            ("remove_sheet", remove_sheet),
            ("other profile", other_profile),
        ):
            with self.subTest(edit=label), formula_book() as wb:
                edit(wb)
                self.assertTrue(wb.has_dirty_cells())

    def test_same_profile_stays_clean(self) -> None:
        with _open() as wb:
            wb.sheet(0).set_formula(0, 0, "=1+1")
            wb.recalc()
            wb.set_profile_id(wb.get_profile_id())
            self.assertFalse(wb.has_dirty_cells())

    def test_volatile_only_workbook_is_clean_after_recalc(self) -> None:
        with _open() as wb:
            wb.sheet(0).set_formula(0, 0, "=NOW()")
            wb.recalc()
            self.assertFalse(wb.has_dirty_cells())

    def test_viewport_recalc_does_not_clean(self) -> None:
        with _open() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            wb.sheet(0).set_formula(0, 1, "=A1+1")
            wb.sheet(0).set_formula(500, 1, "=A1+2")
            wb.recalc()
            wb.sheet(0).set_number(0, 0, 2.0)
            wb.sheet(0).recalc_viewport(CellRect(0, 0, 5, 5))
            self.assertTrue(wb.has_dirty_cells())

    def test_closed_workbook_raises(self) -> None:
        wb = _open()
        wb.close()
        with self.assertRaises(formulon.WorkbookClosedError):
            wb.has_dirty_cells()


# ---------------------------------------------------------------------------
# Criteria 6 and 13: no counts or index getters, sibling alignment
# ---------------------------------------------------------------------------


class NoCountNamesTests(unittest.TestCase):
    EXEMPT = {"get_record_count"}

    def test_no_count_or_at_names_remain(self) -> None:
        offenders = []
        for owner in (Workbook, Sheet, *_HANDLE_CLASSES):
            for name in dir(owner):
                if re.fullmatch(r"get_\w+_(count|at)", name) and name not in self.EXEMPT:
                    offenders.append(f"{owner.__name__}.{name}")
        self.assertEqual(offenders, [])

    def test_removed_enumerators(self) -> None:
        removed = {
            Sheet: ["get_cell_count", "get_cell_at", "get_pivot_count", "get_pivot_field_item_count"],
            Workbook: [
                "get_font_count",
                "get_fill_count",
                "get_border_count",
                "get_cell_xf_count",
                "get_cell_style_xf_count",
                "get_dxf_count",
                "get_pivot_cache_field_count",
                "get_pivot_cache_field_name",
                "get_pivot_cache_field_formula",
                "get_pivot_cache_field_shared_item_count",
                "get_pivot_cache_field_shared_item_value",
                "is_pivot_cache_field_shared_item_date",
            ],
        }
        for owner, names in removed.items():
            for name in names:
                self.assertFalse(hasattr(owner, name), f"{owner.__name__}.{name}")

    def test_lists_exist(self) -> None:
        for name in ("list_cells", "list_pivots"):
            self.assertTrue(hasattr(Sheet, name), name)
        self.assertTrue(hasattr(formulon.Pivot, "list_field_items"))
        for name in (
            "list_fonts",
            "list_fills",
            "list_borders",
            "list_cell_xfs",
            "list_cell_style_xfs",
            "list_dxfs",
        ):
            self.assertTrue(hasattr(formulon.WorkbookOoxml, name), name)
        for name in ("list_fields", "list_field_shared_items"):
            self.assertTrue(hasattr(formulon.PivotCache, name), name)

    def test_list_cells_matches_the_cells_written(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.set_number(0, 0, 1.0)
            ws.set_text(2, 3, "x")
            cells = ws.list_cells()
            self.assertEqual({(c.row, c.col) for c in cells}, {(0, 0), (2, 3)})
            self.assertEqual({(c.row, c.col) for c in ws.list_cells("A1:B2")}, {(0, 0)})
            self.assertEqual(ws.list_cells(CellRect(5, 5, 6, 6)), [])

    def test_pivot_and_cache_lists(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            _pivot_book(wb, ws)
            cache = wb.list_pivot_caches()[0]
            pivots = ws.list_pivots()
            self.assertEqual([(p.name, p.cache().cache_id) for p in pivots], [("Pivot1", cache.cache_id)])
            self.assertIs(pivots[0].cache(), cache)
            fields = cache.list_fields()
            self.assertEqual([(f.name, f.formula) for f in fields], [("Region", None), ("Amount", None)])
            properties = pivots[0].get_properties()
            self.assertEqual((properties.name, properties.cache_id), ("Pivot1", cache.cache_id))
            self.assertTrue(dataclasses.is_dataclass(properties))
            self.assertTrue(properties.__dataclass_params__.frozen)  # type: ignore[attr-defined]

    def test_sibling_alignment(self) -> None:
        for name in ("convert_column_chars_to_pt", "convert_column_pt_to_chars"):
            mode = inspect.signature(getattr(Sheet, name)).parameters["mode"]
            self.assertIs(mode.default, formulon.GeometryMode.DISPLAY)
            self.assertEqual(list(inspect.signature(getattr(Sheet, name)).parameters)[-1], "mode")
        self.assertFalse(hasattr(Sheet, "remove_merge_at"))
        self.assertTrue(hasattr(Sheet, "remove_hyperlink_at"))
        self.assertTrue(hasattr(Sheet, "insert_hyperlink_at"))
        self.assertTrue(hasattr(Sheet, "remove_hyperlink"))
        self.assertTrue(hasattr(Sheet, "evaluate_formula"))
        self.assertFalse(hasattr(formulon, "get_status_name"))
        self.assertIn("scheme", {f.name for f in dataclasses.fields(formulon.ThemeScriptFont)})
        self.assertIn("scope", inspect.signature(Workbook.get_defined_name).parameters)
        self.assertIn("scope", inspect.signature(Workbook.remove_defined_name).parameters)
        sort_by_field = inspect.signature(formulon.Pivot.update_field).parameters["sort_by_field"]
        self.assertIs(sort_by_field.kind, inspect.Parameter.KEYWORD_ONLY)
        self.assertIs(sort_by_field.default, formulon.UNCHANGED)
        visible = inspect.signature(formulon.Pivot.add_field_item).parameters["visible"]
        self.assertIs(visible.kind, inspect.Parameter.KEYWORD_ONLY)
        self.assertIs(visible.default, True)
        cache_index = inspect.signature(formulon.Pivot.add_field_item).parameters["cache_index"]
        self.assertIs(cache_index.kind, inspect.Parameter.KEYWORD_ONLY)
        self.assertIsNone(cache_index.default)
        self.assertFalse(hasattr(formulon.Pivot, "add_field_shared_item"))
        grouping = inspect.signature(formulon.Chart.set_type).parameters["grouping"]
        self.assertEqual(grouping.default, formulon.ChartGrouping.CLUSTERED)

    def test_defined_name_scope(self) -> None:
        with _open() as wb:
            wb.set_defined_name(formulon.DefinedNameInput("X", "Sheet1!$A$1"))
            wb.set_defined_name(formulon.DefinedNameInput("X", "Sheet1!$B$1", 0))
            self.assertEqual(wb.get_defined_name("X").formula, "=Sheet1!$A$1")  # type: ignore[union-attr]
            self.assertEqual(wb.get_defined_name("X", "Sheet1").formula, "=Sheet1!$B$1")  # type: ignore[union-attr]
            self.assertEqual(wb.get_defined_name("X", scope=0).formula, "=Sheet1!$B$1")  # type: ignore[union-attr]
            wb.remove_defined_name("X", "Sheet1")
            self.assertEqual(wb.get_defined_name("X").formula, "=Sheet1!$A$1")  # type: ignore[union-attr]

    def test_geometry_mode_is_validated(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            self.assertEqual(
                ws.convert_column_chars_to_pt(10), ws.convert_column_chars_to_pt(10, formulon.GeometryMode.DISPLAY)
            )
            with self.assertRaises(ValueError):
                ws.convert_column_chars_to_pt(10, 7)  # type: ignore[arg-type]

    def test_iterative_validation(self) -> None:
        with _open() as wb:
            with self.assertRaises(InvalidArgumentError):
                wb.update_iterative(max_change=-1.0)
            with self.assertRaises(TypeError):
                wb.update_iterative(max_iterations=1.5)  # type: ignore[arg-type]

    def test_formula_text_forms(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.set_formula("A1", "1+1")
            self.assertEqual(ws.get_formula(0, 0), "=1+1")
            ws.set_formula(1, 1, "=A1+1")
            self.assertEqual(ws.get_formula_r1c1(1, 1), "=R[-1]C[-1]+1")
            ws.set_formula(2, 2, "=1.50")
            self.assertEqual(ws.get_formula(2, 2), "=1.5")


# ---------------------------------------------------------------------------
# Criterion 7: patches
# ---------------------------------------------------------------------------


class PatchShapeTests(unittest.TestCase):
    """Every patch method takes keyword-only ``UNCHANGED`` arguments named like fields of the paired record."""

    PATCHES: List[Tuple[type, str, Optional[type]]] = [
        (Sheet, "update_view", getattr(formulon, "SheetView", None)),
        (formulon.Pivot, "update_field", getattr(formulon, "PivotField", None)),
        (formulon.Pivot, "update_field_item", getattr(formulon, "PivotItem", None)),
        (formulon.Pivot, "update_style_options", getattr(formulon, "PivotStyleOptions", None)),
        (formulon.Chart, "update_legend", getattr(formulon, "ChartLegend", None)),
        (Sheet, "update_threaded_comment", None),
        (formulon.Table, "update", None),
        (formulon.Table, "update_style_options", getattr(formulon, "TableStyleOptions", None)),
        (Sheet, "update_print_options", None),
        (Sheet, "update_header_footer", None),
    ]

    def test_signatures(self) -> None:
        for owner, name, record in self.PATCHES:
            with self.subTest(method=name):
                method = getattr(owner, name, None)
                self.assertIsNotNone(method, f"{owner.__name__}.{name} is missing")
                params = inspect.signature(method).parameters
                keyword_only = {n: p for n, p in params.items() if p.kind is inspect.Parameter.KEYWORD_ONLY}
                self.assertTrue(keyword_only, "no keyword-only arguments")
                for n, p in keyword_only.items():
                    self.assertIs(p.default, UNCHANGED, f"{n} must default to UNCHANGED")
                self.assertFalse(
                    [n for n, p in params.items() if p.kind is inspect.Parameter.VAR_KEYWORD and n != "kwargs"]
                )
                if record is not None:
                    fields = {f.name for f in dataclasses.fields(record)}
                    self.assertLessEqual(set(keyword_only), fields)

    def test_unchanged_is_a_singleton_sentinel(self) -> None:
        self.assertIs(formulon.UNCHANGED, UNCHANGED)
        self.assertNotEqual(UNCHANGED, None)

    def test_typed_writers_are_gone(self) -> None:
        for name in (
            "set_zoom",
            "set_freeze",
            "set_show_grid_lines",
            "set_show_row_col_headers",
            "set_show_zeros",
            "set_show_formulas",
            "set_show_outline_symbols",
            "set_right_to_left",
            "set_tab_selected",
            "set_top_left",
            "set_pane_top_left",
            "set_split",
            "set_view_mode",
            "set_pivot_field_custom_name",
            "set_pivot_field_number_format",
            "set_pivot_field_default_subtotal",
            "set_pivot_field_insert_blank_row",
            "set_pivot_field_fill_down_labels",
            "set_pivot_field_show_all",
            "set_pivot_field_subtotal_top",
            "set_pivot_field_item_visible",
            "set_pivot_field_item_show_detail",
            "set_pivot_field_item_caption",
        ):
            self.assertFalse(hasattr(Sheet, name), name)


class SetViewTests(unittest.TestCase):
    def test_empty_patch_changes_nothing(self) -> None:
        with _open() as wb:
            before = wb.sheet(0).get_view()
            wb.sheet(0).update_view()
            self.assertEqual(wb.sheet(0).get_view(), before)

    def test_none_is_a_type_error(self) -> None:
        with _open() as wb:
            with self.assertRaises(TypeError):
                wb.sheet(0).update_view(zoom_pct=None)  # type: ignore[arg-type]

    def test_one_key_changes_one_field(self) -> None:
        with _open() as wb:
            before = wb.sheet(0).get_view()
            wb.sheet(0).update_view(show_grid_lines=False)
            self.assertEqual(wb.sheet(0).get_view(), dataclasses.replace(before, show_grid_lines=False))

    def test_freeze_rows_keeps_freeze_cols(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.update_view(freeze_rows=1, freeze_cols=1)
            ws.update_view(freeze_rows=2)
            view = ws.get_view()
            self.assertEqual((view.freeze_rows, view.freeze_cols), (2, 1))

    def test_freeze_and_split_conflict(self) -> None:
        with _open() as wb:
            with self.assertRaises(InvalidArgumentError):
                wb.sheet(0).update_view(freeze_rows=1, split_x_twips=100)

    def test_pane_corner_is_checked_after_the_merge(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.update_view(freeze_rows=1, pane_top_left="A4")
            view = ws.get_view()
            self.assertEqual((view.freeze_rows, view.pane_top_left), (1, CellAddress(3, 0)))
            with self.assertRaises(FormulonError):
                wb.sheet(0).update_view(freeze_rows=0, pane_top_left="A4")

    def test_hiding_the_only_visible_sheet_fails_and_changes_nothing(self) -> None:
        with _open() as wb:
            before = wb.sheet(0).get_view()
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_visibility(formulon.SheetVisibility.HIDDEN)
            self.assertEqual(wb.sheet(0).get_view(), before)


class ChartLegendPatchTests(unittest.TestCase):
    @staticmethod
    def _chart(wb: Workbook, **options: Any) -> int:
        ws = wb.sheet(0)
        for r in range(3):
            ws.set_text(r + 1, 0, "abc"[r])
            ws.set_number(r + 1, 1, float(r + 1))
        ws.set_text(0, 1, "S")
        return _chart_input.add_chart(
            ws, _chart_input.COLUMN, "Sheet1!$A$1:$B$4", anchor=formulon.DrawingAnchorInput(from_row=6), **options
        )

    def test_position_keeps_overlay(self) -> None:
        with _open() as wb:
            oid = self._chart(wb)
            wb.sheet(0).chart(oid).set_legend(formulon.ChartLegend(formulon.ChartLegendPosition.TOP, True))
            wb.sheet(0).chart(oid).update_legend(position=formulon.ChartLegendPosition.RIGHT)
            legend = wb.sheet(0).chart(oid).get_properties().legend  # type: ignore[union-attr]
            self.assertEqual(legend.position, formulon.ChartLegendPosition.RIGHT)
            self.assertTrue(legend.overlay)

    def test_null_position_and_hidden_entries_are_type_errors(self) -> None:
        with _open() as wb:
            oid = self._chart(wb)
            with self.assertRaises(TypeError):
                wb.sheet(0).chart(oid).update_legend(position=None)  # type: ignore[arg-type]
            with self.assertRaises(TypeError):
                wb.sheet(0).chart(oid).update_legend(hidden_entries=None)  # type: ignore[arg-type]

    def test_overlay_alone_creates_a_right_legend(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            oid = self._chart(wb, legend=None)
            self.assertIsNone(ws.chart(oid).get_properties().legend)  # type: ignore[union-attr]
            ws.chart(oid).update_legend(overlay=True)
            legend = ws.chart(oid).get_properties().legend  # type: ignore[union-attr]
            self.assertEqual(legend.position, formulon.ChartLegendPosition.RIGHT)
            self.assertTrue(legend.overlay)
            self.assertEqual(list(legend.hidden_entries), [])


class PrintPatchTests(unittest.TestCase):
    def test_header_footer_none_and_empty_string_both_clear(self) -> None:
        for cleared in (None, ""):
            with self.subTest(cleared=cleared), _open() as wb:
                ws = wb.sheet(0)
                ws.update_header_footer(odd_header="&Ctitle")
                self.assertIn("title", ws.ooxml.get_header_footer_xml() or "")
                ws.update_header_footer(odd_header=cleared)
                self.assertNotIn("title", ws.ooxml.get_header_footer_xml() or "")

    def test_omitted_keys_are_kept(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.update_header_footer(odd_header="&Ctitle", odd_footer="&Cfoot")
            ws.update_header_footer(odd_header="&Cnew")
            xml = ws.ooxml.get_header_footer_xml() or ""
            self.assertIn("new", xml)
            self.assertIn("foot", xml)

    def test_print_options_none_is_a_type_error(self) -> None:
        with _open() as wb:
            with self.assertRaises(TypeError):
                wb.sheet(0).update_print_options(grid_lines=None)  # type: ignore[arg-type]
            wb.sheet(0).update_print_options(grid_lines=True)
            wb.sheet(0).update_print_options(headings=True)
            xml = wb.sheet(0).ooxml.get_print_options_xml() or ""
            self.assertIn("gridLines", xml)
            self.assertIn("headings", xml)

    def test_print_area_unset_is_none(self) -> None:
        with _open() as wb:
            self.assertIsNone(wb.sheet(0).get_print_area())


class PivotPatchTests(unittest.TestCase):
    def test_update_pivot_field(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            _pivot_book(wb, ws)
            ws.pivot(0).set_field_axis("Region", formulon.PivotAxis.ROW)
            ws.pivot(0).update_field(0, custom_name="Area", show_all=True)
            field = ws.pivot(0).list_fields()[0]
            self.assertEqual(field.custom_name, "Area")
            self.assertTrue(field.show_all)
            ws.pivot(0).update_field(0, custom_name=None)
            field = ws.pivot(0).list_fields()[0]
            self.assertIsNone(field.custom_name)
            self.assertTrue(field.show_all)

    def test_axis_is_not_a_patch_key(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            _pivot_book(wb, ws)
            ws.pivot(0).set_field_axis("Region", formulon.PivotAxis.ROW)
            with self.assertRaises(TypeError):
                ws.pivot(0).update_field(0, axis=formulon.PivotAxis.COLUMN)  # type: ignore[call-arg]

    def test_update_pivot_item(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            _pivot_book(wb, ws)
            ws.pivot(0).set_field_axis("Region", formulon.PivotAxis.ROW)
            ws.pivot(0).update_field_item(0, 0, visible=False, caption="Eastern")
            item = ws.pivot(0).get_field_item(0, 0)
            self.assertFalse(item.visible)
            self.assertEqual(item.caption, "Eastern")
            ws.pivot(0).update_field_item(0, 0, caption=None)
            item = ws.pivot(0).get_field_item(0, 0)
            self.assertIsNone(item.caption)
            self.assertFalse(item.visible)


# ---------------------------------------------------------------------------
# Criterion 8: refs by number or by name
# ---------------------------------------------------------------------------


def _table_book(wb: Workbook) -> None:
    ws = wb.sheet(0)
    for c, head in enumerate(("A", "B")):
        ws.set_text(0, c, head)
    ws.add_table(formulon.TableInput(range=CellRect(0, 0, 2, 1), name="Shown", columns=["A", "B"]))


class RefTests(unittest.TestCase):
    def test_remove_table_by_name(self) -> None:
        with _open() as wb:
            _table_book(wb)
            wb.sheet(0).remove_table("Shown")
            self.assertEqual(wb.sheet(0).list_tables(), [])

    def test_remove_table_by_unknown_name_is_not_found(self) -> None:
        with _open() as wb:
            _table_book(wb)
            with self.assertRaises(NotFoundError):
                wb.sheet(0).remove_table("Original")
            self.assertEqual(len(wb.sheet(0).list_tables()), 1)

    def test_remove_table_edges(self) -> None:
        with _open() as wb:
            _table_book(wb)
            ws = wb.sheet(0)
            with self.assertRaises(NotFoundError):
                ws.remove_table("Missing")
            with self.assertRaises(IndexOutOfRangeError):
                ws.remove_table(99)
            with self.assertRaises(TypeError):
                ws.remove_table(True)  # type: ignore[arg-type]
            ws.remove_table(0)
            self.assertEqual(ws.list_tables(), [])

    def test_string_refs_fold_case(self) -> None:
        with _open() as wb:
            _table_book(wb)
            wb.sheet(0).remove_table("SHOWN")
            self.assertEqual(wb.sheet(0).list_tables(), [])
        with _open() as wb:
            ws = wb.sheet(0)
            _pivot_book(wb, ws)
            ws.remove_pivot("pivot1")
            self.assertEqual(ws.list_pivots(), [])
        with _open() as wb:
            ws = wb.sheet(0)
            ws.add_cf_rule(formulon.CfCellIsRuleInput(ranges=[CellRect(0, 0, 2, 0)], operator=3, formula1="5"))
            rule_id = ws.list_cf_rules()[0].id
            ws.remove_cf_rule(rule_id.lower())
            self.assertEqual(ws.list_cf_rules(), [])

    def test_digit_string_is_a_name(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.set_text(0, 0, "A")
            ws.add_table(formulon.TableInput(range=CellRect(0, 0, 1, 0), name="Sales", columns=["A"]))
            with self.assertRaises(NotFoundError):
                ws.remove_table("0")
            self.assertEqual(len(ws.list_tables()), 1)
            ws.remove_table(0)
            self.assertEqual(ws.list_tables(), [])

    def test_cf_rule_by_id(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.add_cf_rule(formulon.CfCellIsRuleInput(ranges=[CellRect(0, 0, 2, 0)], operator=3, formula1="5"))
            rule_id = ws.list_cf_rules()[0].id
            self.assertTrue(rule_id)
            ws.remove_cf_rule(rule_id)
            self.assertEqual(ws.list_cf_rules(), [])
            with self.assertRaises(NotFoundError):
                ws.remove_cf_rule(rule_id)

    def test_colliding_cf_id_gets_a_fresh_one(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.add_cf_rule(formulon.CfCellIsRuleInput(ranges=[CellRect(0, 0, 2, 0)], operator=3, formula1="5"))
            rule_id = ws.list_cf_rules()[0].id
            ws.add_cf_rule(
                formulon.CfCellIsRuleInput(ranges=[CellRect(0, 1, 2, 1)], operator=3, formula1="6", id=rule_id)
            )
            ids = [rule.id for rule in ws.list_cf_rules()]
            self.assertEqual(len(ids), 2)
            self.assertEqual(ids[0], rule_id)
            self.assertNotEqual(ids[1], rule_id)

    def test_cf_round_trip_keeps_one_equals_sign(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.add_cf_rule(formulon.CfCellIsRuleInput(ranges=[CellRect(0, 0, 2, 0)], operator=3, formula1="5"))
            rule = ws.list_cf_rules()[0]
            self.assertTrue(rule.formula1.startswith("="))
            self.assertFalse(rule.formula1.startswith("=="))
            ws.add_cf_rule(rule)
            first, second = ws.list_cf_rules()[:2]
            self.assertEqual(first.formula1, second.formula1)
            self.assertFalse(second.formula1.startswith("=="))

    def test_duplicate_pivot_name_is_refused_and_changes_nothing(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            first = wb.add_pivot_cache().cache_id
            second = wb.add_pivot_cache().cache_id
            ws.add_pivot(PivotInput(name="Dup", cache_id=first, anchor=CellAddress(0, 0)))
            for name in ("Dup", "dup"):
                with self.assertRaises(AlreadyExistsError):
                    ws.add_pivot(PivotInput(name=name, cache_id=second, anchor=CellAddress(0, 10)))
            self.assertEqual([(p.name, p.cache().cache_id) for p in ws.list_pivots()], [("Dup", first)])
            ws.remove_pivot("DUP")
            self.assertEqual(ws.list_pivots(), [])

    def test_pivot_by_name_and_number_agree(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            _pivot_book(wb, ws)
            with self.assertRaises(NotFoundError):
                ws.remove_pivot("Missing")
            with self.assertRaises(TypeError):
                ws.remove_pivot(True)  # type: ignore[arg-type]
            ws.remove_pivot("Pivot1")
            self.assertEqual(ws.list_pivots(), [])


# ---------------------------------------------------------------------------
# Criterion 9: stateless functions live on the module
# ---------------------------------------------------------------------------


class ModuleFunctionTests(unittest.TestCase):
    NAMES = [
        "get_function_metadata",
        "list_function_names",
        "localize_function_name",
        "canonicalize_function_name",
        "localize_formula",
        "canonicalize_formula",
        "get_locale_facts",
        "compare_text",
        "get_collation_classes",
        "get_image_info",
    ]

    def test_module_only(self) -> None:
        for name in self.NAMES:
            with self.subTest(name=name):
                self.assertTrue(callable(getattr(formulon, name, None)), "missing on the module")
                self.assertFalse(hasattr(Workbook, name), "still on Workbook")
                self.assertFalse(hasattr(Sheet, name), "still on Sheet")

    def test_locale_error_names_come_from_the_locale_facts_only(self) -> None:
        self.assertFalse(hasattr(formulon, "list_locale_error_names"))
        self.assertNotIn("list_locale_error_names", formulon.__all__)
        facts = formulon.get_locale_facts("mac-365-de_DE")
        na = next(e for e in facts.error_names if e.canonical == "#N/A")
        self.assertEqual(na.localized, "#NV")

    def test_get_image_info_reads_a_png(self) -> None:
        png = (
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x03\x00\x00\x00\x02\x08\x02\x00\x00\x00"
            b"\x00\x00\x00\x00\x00\x00\x00\x00IEND\xaeB`\x82"
        )
        info = formulon.get_image_info(png)
        self.assertEqual((info.width_px, info.height_px), (3, 2))
        self.assertEqual(info.format, formulon.ImageFormat.PNG)


# ---------------------------------------------------------------------------
# Criterion 10: pivot cache writers dispatch on the value type
# ---------------------------------------------------------------------------


class PivotCacheWriterTests(unittest.TestCase):
    TYPED = [
        f"add_pivot_cache_field_shared_item_{kind}" for kind in ("text", "number", "bool", "blank", "error", "date")
    ] + [f"set_pivot_cache_record_{kind}" for kind in ("text", "number", "bool", "blank", "error", "date")]

    def test_typed_writers_are_gone(self) -> None:
        for name in self.TYPED:
            self.assertFalse(hasattr(Workbook, name), name)
            self.assertFalse(hasattr(formulon.PivotCache, name), name)
        self.assertTrue(hasattr(formulon.PivotCache, "add_field_shared_item"))
        self.assertTrue(hasattr(formulon.PivotCache, "set_record_value"))

    def _cache(self, wb: Workbook) -> int:
        cache = wb.add_pivot_cache().cache_id
        wb.pivot_cache(cache).add_field("F")
        return cache

    def test_values_dispatch_on_type(self) -> None:
        with _open() as wb:
            cache = self._cache(wb)
            for value in ("East", 4.5, True, None, ExcelError(ErrorCode.NA)):
                wb.pivot_cache(cache).add_field_shared_item(0, value)
            items = wb.pivot_cache(cache).list_field_shared_items(0)
            self.assertEqual([i.value for i in items], ["East", 4.5, True, None, ExcelError(ErrorCode.NA)])
            self.assertFalse(any(i.is_date for i in items))

    def test_date_flag(self) -> None:
        with _open() as wb:
            cache = self._cache(wb)
            wb.pivot_cache(cache).add_field_shared_item(0, datetime.date(1900, 1, 5))
            wb.pivot_cache(cache).add_field_shared_item(0, 5)
            items = wb.pivot_cache(cache).list_field_shared_items(0)
            self.assertEqual([i.is_date for i in items], [True, False])
            self.assertEqual([i.value for i in items], [5.0, 5.0])

    def test_record_value(self) -> None:
        with _open() as wb:
            cache = self._cache(wb)
            rec = wb.pivot_cache(cache).add_record()
            wb.pivot_cache(cache).set_record_value(rec, 0, "x")
            wb.pivot_cache(cache).set_record_value(rec, 0, datetime.date(1900, 1, 3))
            with self.assertRaises(TypeError):
                wb.pivot_cache(cache).set_record_value(rec, 0, datetime.time(1, 2))


# ---------------------------------------------------------------------------
# Criterion 11: save
# ---------------------------------------------------------------------------


def _container(blob: bytes) -> str:
    names = zipfile.ZipFile(io.BytesIO(blob)).namelist()
    if "xl/workbook.bin" in names:
        return "xlsb"
    if "xl/workbook.xml" in names:
        return "xlsx"
    return "unknown"


class SaveTests(unittest.TestCase):
    def _save(self, name: str, **kwargs: Any) -> str:
        with tempfile.TemporaryDirectory() as tmp, _open() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            wb.save(os.path.join(tmp, name), **kwargs)
            return _container(Path(tmp, name).read_bytes())

    def test_extension_picks_the_container(self) -> None:
        self.assertEqual(self._save("o.xlsb"), "xlsb")
        self.assertEqual(self._save("O.XLSB"), "xlsb")
        self.assertEqual(self._save("o.xlsx"), "xlsx")
        with self.assertRaises(ValueError):
            self._save("o.bin")
        self.assertEqual(self._save("o.bin", format=formulon.WorkbookFormat.XLSX), "xlsx")

    def test_path_like(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, _open() as wb:
            target = Path(tmp) / "o.xlsb"
            wb.save(target)
            self.assertEqual(_container(target.read_bytes()), "xlsb")

    def test_conflicting_format_is_invalid_argument(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, _open() as wb:
            with self.assertRaises(InvalidArgumentError):
                wb.save(os.path.join(tmp, "o.xlsb"), format=formulon.WorkbookFormat.XLSX)
            with self.assertRaises(InvalidArgumentError):
                wb.save(os.path.join(tmp, "o.xlsx"), format=formulon.WorkbookFormat.XLSB)

    def test_explicit_format_without_a_decisive_extension(self) -> None:
        self.assertEqual(self._save("o.bin", format=formulon.WorkbookFormat.XLSB), "xlsb")

    def test_stream_and_bytes_default_to_xlsx(self) -> None:
        with _open() as wb:
            stream = io.BytesIO()
            wb.save(stream)
            self.assertEqual(_container(stream.getvalue()), "xlsx")
            self.assertEqual(_container(wb.save()), "xlsx")
            self.assertEqual(_container(wb.save(format=formulon.WorkbookFormat.XLSB)), "xlsb")

    def test_save_with_diagnostics_takes_dest_and_returns_bytes(self) -> None:
        with tempfile.TemporaryDirectory() as tmp, _open() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            path = os.path.join(tmp, "o.xlsx")
            diagnostics = wb.save_with_diagnostics(path)
            self.assertEqual(diagnostics.bytes, Path(path).read_bytes())
            xlsb = wb.save_with_diagnostics(os.path.join(tmp, "o.xlsb"))
            self.assertEqual(_container(xlsb.bytes), "xlsb")
            self.assertEqual(_container(wb.save_with_diagnostics().bytes), "xlsx")

    def test_formula_equals_sign_survives_a_round_trip(self) -> None:
        with _open() as wb:
            wb.sheet(0).set_formula(0, 0, "=1+1")
            with Workbook.load(wb.save()) as loaded:
                self.assertEqual(loaded.sheet(0).get_formula(0, 0), "=1+1")


# ---------------------------------------------------------------------------
# Criterion 12: missing keys
# ---------------------------------------------------------------------------


class MissingKeyTests(unittest.TestCase):
    def test_remove_missing_merge_and_data_table(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            with self.assertRaises(NotFoundError):
                ws.remove_merge(CellRect(5, 5, 6, 6))
            with self.assertRaises(NotFoundError):
                ws.remove_data_table(3, 3)

    def test_missing_drawing_object_and_chart_data(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            with self.assertRaises(NotFoundError):
                ws.remove_drawing(999)
            with self.assertRaises(NotFoundError):
                ws.chart(999).get_data()
            with self.assertRaises(NotFoundError):
                ws.chart(999).get_preview()


# ---------------------------------------------------------------------------
# Criterion 14: error display names
# ---------------------------------------------------------------------------


class ErrorDisplayNameTests(unittest.TestCase):
    def test_enum_property(self) -> None:
        self.assertEqual(ErrorCode.DIV0.display_name, "#DIV/0!")
        self.assertEqual(ErrorCode.NA.display_name, "#N/A")
        self.assertEqual(ErrorCode.NULL.display_name, "#NULL!")

    def test_str_of_error_value(self) -> None:
        self.assertEqual(str(ExcelError(ErrorCode.DIV0)), "#DIV/0!")

    def test_every_code_has_a_name(self) -> None:
        for code in ErrorCode:
            self.assertTrue(code.display_name.startswith("#"), code.name)

    def test_unknown_code_is_a_value_error(self) -> None:
        with self.assertRaises(ValueError):
            ErrorCode(999)

    def test_function_is_gone(self) -> None:
        self.assertFalse(hasattr(formulon, "get_error_display_name"))
        self.assertFalse(hasattr(Workbook, "get_error_display_name"))


# ---------------------------------------------------------------------------
# Criterion 13 leftovers: removed Python-only methods
# ---------------------------------------------------------------------------


class PlatformOnlyRemovalTests(unittest.TestCase):
    def test_removed_python_only_names(self) -> None:
        for name in ("add_batch", "rename_table", "set_table_style_options"):
            self.assertFalse(hasattr(Workbook, name), name)
        self.assertFalse(hasattr(Sheet, "has_cell_phonetic_properties"))

    def test_phonetic_properties_expose_stated(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            ws.set_text(0, 0, "x")
            self.assertFalse(ws.get_cell_phonetic_properties(0, 0).stated)
            ws.set_cell_phonetic_properties(0, 0, PhoneticProperties(font_index=0, type=1, alignment=1))
            self.assertTrue(ws.get_cell_phonetic_properties(0, 0).stated)

    def test_unset_pivot_option_strings_read_none(self) -> None:
        with _open() as wb:
            ws = wb.sheet(0)
            _pivot_book(wb, ws)
            options = ws.pivot(0).get_options()
            self.assertIsNone(options.error_caption)
            self.assertIsNone(options.missing_caption)


if __name__ == "__main__":
    unittest.main()
