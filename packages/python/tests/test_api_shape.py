"""Shape of the idiomatic Python surface.

Covers the exception classes, ``ExcelError``, ``load`` / ``save`` with paths
and streams, the ``closed`` lifecycle predicate, ``set_value``,
``get_sheet_index`` and the A1-addressed ``Sheet`` view.
"""

from __future__ import annotations

import copy
import dataclasses
import datetime
import inspect
import io
import math
import pathlib
import pickle
import tempfile
import unittest

import formulon
from formulon import (
    CellAddress,
    Comment,
    ErrorCode,
    ExcelError,
    FormulonError,
    IndexOutOfRangeError,
    InvalidArgumentError,
    NotFoundError,
    PivotInput,
    Sheet,
    StatusCode,
    Value,
    ValueKind,
    Workbook,
    WorkbookClosedError,
    WorkbookFormat,
)

# A local-file-header signature followed by zeros: not a readable ZIP.
_CORRUPT_ZIP_16 = b"PK\x03\x04" + b"\x00" * 12


class ExceptionClassTests(unittest.TestCase):
    def test_hierarchy(self) -> None:
        self.assertTrue(issubclass(NotFoundError, FormulonError))
        self.assertTrue(issubclass(NotFoundError, KeyError))
        self.assertTrue(issubclass(IndexOutOfRangeError, FormulonError))
        self.assertTrue(issubclass(IndexOutOfRangeError, IndexError))
        self.assertFalse(issubclass(IndexOutOfRangeError, NotFoundError))
        self.assertTrue(issubclass(InvalidArgumentError, ValueError))
        self.assertTrue(issubclass(WorkbookClosedError, FormulonError))
        self.assertTrue(issubclass(WorkbookClosedError, ValueError))
        self.assertTrue(issubclass(NotFoundError, LookupError))
        self.assertTrue(issubclass(IndexOutOfRangeError, LookupError))

    def test_package_exports(self) -> None:
        for name in ("ExcelError", "IndexOutOfRangeError", "WorkbookClosedError", "Sheet"):
            self.assertIn(name, formulon.__all__)

    def test_sheet_index_out_of_range_is_index_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(IndexOutOfRangeError) as cm:
                wb.sheet(5).name
            self.assertEqual(cm.exception.code, StatusCode.SHEET_INDEX_OUT_OF_RANGE)
            self.assertEqual(cm.exception.context, "sheet_index=5 sheet_count=1")
            self.assertEqual(str(cm.exception), f"{cm.exception.message} [sheet_index=5 sheet_count=1]")

    def test_collection_index_out_of_range_is_index_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(IndexOutOfRangeError) as cm:
                wb.ooxml.get_font(len(wb.ooxml.list_fonts()))
            self.assertEqual(cm.exception.code, StatusCode.INDEX_OUT_OF_RANGE)

    def test_missing_natural_key_remove_is_key_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(NotFoundError) as cm:
                wb.remove_defined_name("nope")
            self.assertIsInstance(cm.exception, KeyError)
            self.assertEqual(cm.exception.code, StatusCode.NOT_FOUND)
            # KeyError would render through repr; the text is the plain message.
            self.assertFalse(str(cm.exception).startswith(("'", '"')))

    def test_message_names_no_c_function(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError) as cm:
                wb.sheet(5).name
            self.assertNotIn("fm_", str(cm.exception))
            self.assertNotIn("SHEET_INDEX_OUT_OF_RANGE", str(cm.exception))

    def test_message_without_context_has_no_brackets(self) -> None:
        wb = Workbook.create_default()
        wb.close()
        with self.assertRaises(WorkbookClosedError) as cm:
            len(wb)
        self.assertEqual(cm.exception.code, StatusCode.BINDING_INVALID_HANDLE)
        self.assertEqual(cm.exception.message, "the workbook is closed")
        self.assertEqual(cm.exception.context, "")
        self.assertEqual(str(cm.exception), "the workbook is closed")


class ExcelErrorTests(unittest.TestCase):
    def test_frozen_dataclass_with_error_code(self) -> None:
        e = ExcelError(ErrorCode.DIV0)
        self.assertIs(e.code, ErrorCode.DIV0)
        self.assertEqual(e, ExcelError(ErrorCode.DIV0))
        with self.assertRaises(AttributeError):
            e.code = ErrorCode.NA  # type: ignore[misc]

    def test_str_is_display_literal(self) -> None:
        self.assertEqual(str(ExcelError(ErrorCode.DIV0)), "#DIV/0!")
        self.assertEqual(str(ExcelError(ErrorCode.NA)), "#N/A")

    def test_to_python_returns_excel_error(self) -> None:
        v = formulon.evaluate_typed_formula("=1/0")
        self.assertIs(v.kind, ValueKind.ERROR)
        self.assertIs(v.error_code, ErrorCode.DIV0)
        self.assertEqual(v.to_python(), ExcelError(ErrorCode.DIV0))

    def test_value_error_code_is_error_code(self) -> None:
        self.assertIs(Value(kind=ValueKind.ERROR, error_code=6).error_code, ErrorCode.NA)

    def test_array_value_converts_to_nested_lists(self) -> None:
        v = Value(kind=ValueKind.ARRAY, array=((Value(kind=ValueKind.NUMBER, number=1.0),),))
        self.assertEqual(v.to_python(), [[1.0]])


class LoadSaveTests(unittest.TestCase):
    def setUp(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 42.0)
            self.xlsx = wb.save()
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = pathlib.Path(tmp.name)

    def _assert_loaded(self, wb: Workbook) -> None:
        with wb:
            self.assertEqual(wb.sheet(0).get_typed_value(0, 0).number, 42.0)

    def test_save_without_dest_returns_bytes(self) -> None:
        self.assertIsInstance(self.xlsx, bytes)
        self.assertTrue(self.xlsx.startswith(b"PK"))

    def test_load_bytes_like(self) -> None:
        self._assert_loaded(Workbook.load(self.xlsx))
        self._assert_loaded(Workbook.load(bytearray(self.xlsx)))
        self._assert_loaded(Workbook.load(memoryview(self.xlsx)))

    def test_load_str_and_pathlike(self) -> None:
        path = self.dir / "a.xlsx"
        path.write_bytes(self.xlsx)
        self._assert_loaded(Workbook.load(str(path)))
        self._assert_loaded(Workbook.load(path))

    def test_load_binary_stream(self) -> None:
        self._assert_loaded(Workbook.load(io.BytesIO(self.xlsx)))
        path = self.dir / "b.xlsx"
        path.write_bytes(self.xlsx)
        with open(path, "rb") as f:
            self._assert_loaded(Workbook.load(f))

    def test_load_text_stream_is_type_error(self) -> None:
        with self.assertRaises(TypeError):
            Workbook.load(io.StringIO("not bytes"))

    def test_load_missing_path_is_file_not_found(self) -> None:
        with self.assertRaises(FileNotFoundError):
            Workbook.load(self.dir / "missing.xlsx")

    def test_load_other_type_is_type_error(self) -> None:
        for bad in (42, None, ["PK"]):
            with self.subTest(bad=bad), self.assertRaises(TypeError):
                Workbook.load(bad)  # type: ignore[arg-type]

    def test_load_empty_is_invalid_argument(self) -> None:
        with self.assertRaises(InvalidArgumentError) as cm:
            Workbook.load(b"")
        self.assertEqual(cm.exception.code, StatusCode.INVALID_ARGUMENT)

    def test_load_corrupt_zip(self) -> None:
        for data in (b"P", _CORRUPT_ZIP_16):
            with self.subTest(length=len(data)):
                with self.assertRaises(FormulonError) as cm:
                    Workbook.load(data)
                self.assertEqual(cm.exception.code, StatusCode.IO_ZIP_CORRUPT)

    def test_save_to_str_and_pathlike(self) -> None:
        with Workbook.load(self.xlsx) as wb:
            p1 = self.dir / "s1.xlsx"
            p2 = self.dir / "s2.xlsx"
            self.assertIsNone(wb.save(str(p1)))
            self.assertIsNone(wb.save(p2))
        self._assert_loaded(Workbook.load(p1))
        self._assert_loaded(Workbook.load(p2))

    def test_save_to_stream(self) -> None:
        buf = io.BytesIO()
        with Workbook.load(self.xlsx) as wb:
            self.assertIsNone(wb.save(buf))
        self._assert_loaded(Workbook.load(buf.getvalue()))

    def test_save_xlsb(self) -> None:
        with Workbook.load(self.xlsx) as wb:
            data = wb.save(format=WorkbookFormat.XLSB)
            buf = io.BytesIO()
            wb.save(buf, format=WorkbookFormat.XLSB)
        self.assertIsInstance(data, bytes)
        self._assert_loaded(Workbook.load(data))
        self._assert_loaded(Workbook.load(buf.getvalue()))

    def test_save_unknown_format_value_is_value_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError):
                wb.save(format=99)  # type: ignore[arg-type]

    def test_save_unknown_format_member_is_invalid_argument(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(InvalidArgumentError):
                wb.save(format=WorkbookFormat.UNKNOWN)

    def test_save_format_is_keyword_only(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.save(WorkbookFormat.XLSB)  # type: ignore[call-overload]

    def test_save_bad_dest_is_type_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(TypeError):
                wb.save(42)  # type: ignore[call-overload]

    def test_save_write_failure_is_os_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(OSError):
                wb.save(self.dir / "no-such-dir" / "x.xlsx")

    def test_save_as_is_gone(self) -> None:
        self.assertFalse(hasattr(Workbook, "save_as"))

    def test_save_with_diagnostics_takes_keyword_format(self) -> None:
        with Workbook.load(self.xlsx) as wb:
            diag = wb.save_with_diagnostics(format=WorkbookFormat.XLSB)
            self._assert_loaded(Workbook.load(diag.bytes))
            self.assertTrue(wb.save_with_diagnostics().bytes.startswith(b"PK"))
            with self.assertRaises(TypeError):
                wb.save_with_diagnostics(WorkbookFormat.XLSX)  # type: ignore[misc]


class LifecycleTests(unittest.TestCase):
    def test_direct_construction_is_type_error(self) -> None:
        with self.assertRaises(TypeError) as cm:
            Workbook()
        self.assertEqual(str(cm.exception), "use Workbook.create_default(), create_for_profile() or load()")

    def test_closed_property(self) -> None:
        wb = Workbook.create_default()
        self.assertFalse(wb.closed)
        wb.close()
        self.assertTrue(wb.closed)
        wb.close()
        self.assertTrue(wb.closed)
        self.assertFalse(hasattr(wb, "is_valid"))

    def test_with_block_closes(self) -> None:
        with Workbook.create_default() as wb:
            self.assertFalse(wb.closed)
        self.assertTrue(wb.closed)

    def test_calls_after_close_raise_workbook_closed(self) -> None:
        wb = Workbook.create_default()
        wb.close()
        with self.assertRaises(WorkbookClosedError):
            wb.sheet(0).get_value(0, 0)
        with self.assertRaises(WorkbookClosedError):
            wb.save()


_ARRAY_VALUE = Value(kind=ValueKind.ARRAY, array=((Value(kind=ValueKind.NUMBER, number=1.0),),))


class SetValueTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)

    def _stored(self, value: object) -> Value:
        self.wb.sheet(0).set_value(0, 0, value)  # type: ignore[arg-type]
        return self.wb.sheet(0).get_typed_value(0, 0)

    def test_zero_and_negative_zero(self) -> None:
        self.assertEqual(self._stored(0), Value(kind=ValueKind.NUMBER, number=0.0))
        v = self._stored(-0.0)
        self.assertIs(v.kind, ValueKind.NUMBER)
        self.assertEqual(v.number, 0.0)

    def test_non_finite_is_invalid_argument(self) -> None:
        for bad in (math.nan, math.inf, -math.inf):
            with self.subTest(bad=bad), self.assertRaises(InvalidArgumentError):
                self.wb.sheet(0).set_value(0, 0, bad)

    def test_none_is_blank(self) -> None:
        self.wb.sheet(0).set_number(0, 0, 1.0)
        self.assertIs(self._stored(None).kind, ValueKind.BLANK)

    def test_missing_value_is_type_error(self) -> None:
        with self.assertRaises(TypeError):
            self.wb.sheet(0).set_value(0, 0)  # type: ignore[call-arg]

    def test_bool_before_int(self) -> None:
        self.assertEqual(self._stored(True), Value(kind=ValueKind.BOOL, boolean=True))
        self.assertEqual(self._stored(False), Value(kind=ValueKind.BOOL, boolean=False))

    def test_int_and_float_are_numbers(self) -> None:
        self.assertEqual(self._stored(7), Value(kind=ValueKind.NUMBER, number=7.0))
        self.assertEqual(self._stored(2.5), Value(kind=ValueKind.NUMBER, number=2.5))

    def test_large_ints_round_to_float(self) -> None:
        self.assertEqual(self._stored(2**53).number, float(2**53))
        self.assertEqual(self._stored(2**63).number, float(2**63))

    def test_int_beyond_float_is_value_error(self) -> None:
        with self.assertRaises(ValueError) as cm:
            self.wb.sheet(0).set_value(0, 0, 2**1024)
        self.assertNotIsInstance(cm.exception, FormulonError)

    def test_formula_like_string_is_a_formula_and_typed_text_stays_text(self) -> None:
        self._stored("=A1")
        self.assertEqual(self.wb.sheet(0).get_formula(0, 0), "=A1")
        text = Value(kind=ValueKind.TEXT, text="=A1")
        self.assertEqual(self._stored(text), text)
        self.assertIsNone(self.wb.sheet(0).get_formula(0, 0))

    def test_excel_error(self) -> None:
        self.assertEqual(self._stored(ExcelError(ErrorCode.NA)), Value(kind=ValueKind.ERROR, error_code=ErrorCode.NA))

    def test_unsupported_types_are_type_error(self) -> None:
        for bad in (object(), [1], {"a": 1}, b"bytes", _ARRAY_VALUE, Value(kind=ValueKind.REF)):
            with self.subTest(bad=bad), self.assertRaises(TypeError):
                self.wb.sheet(0).set_value(0, 0, bad)  # type: ignore[arg-type]

    def test_scalar_value_round_trip(self) -> None:
        for v in (
            Value(kind=ValueKind.BLANK),
            Value(kind=ValueKind.NUMBER, number=3.25),
            Value(kind=ValueKind.BOOL, boolean=True),
            Value(kind=ValueKind.TEXT, text="hello"),
            Value(kind=ValueKind.ERROR, error_code=ErrorCode.DIV0),
        ):
            with self.subTest(kind=v.kind.name):
                self.wb.sheet(0).set_value(1, 1, v)
                self.assertEqual(self.wb.sheet(0).get_typed_value(1, 1), v)


class SheetIndexTests(unittest.TestCase):
    def test_sheet_by_name(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Data")
            self.assertEqual(wb.sheet("Sheet1").index, 0)
            self.assertEqual(wb.sheet("data").index, 1)
            self.assertNotIn("nope", wb)
            self.assertNotIn("", wb)
            with self.assertRaises(TypeError):
                wb.sheet(None)  # type: ignore[arg-type]


class WorkbookSheetAccessTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.wb.add_sheet("Data")
        self.addCleanup(self.wb.close)

    def test_by_name_and_index(self) -> None:
        ws = self.wb.sheet("Data")
        self.assertIsInstance(ws, Sheet)
        self.assertEqual(ws.index, 1)
        self.assertIs(ws.workbook, self.wb)
        self.assertEqual(self.wb["sheet1"].index, 0)
        self.assertEqual(self.wb[1].name, "Data")

    def test_missing_name_is_not_found(self) -> None:
        for name in ("", "nope"):
            with self.subTest(name=name):
                with self.assertRaises(NotFoundError):
                    self.wb.sheet(name)
                with self.assertRaises(KeyError):
                    self.wb[name]

    def test_negative_index_is_value_error(self) -> None:
        with self.assertRaises(ValueError) as cm:
            self.wb.sheet(-1)
        self.assertNotIsInstance(cm.exception, FormulonError)

    def test_index_at_count_is_index_out_of_range(self) -> None:
        for index in (2, 2**31, 2**32):
            with self.subTest(index=index):
                with self.assertRaises(IndexOutOfRangeError) as cm:
                    self.wb.sheet(index)
                self.assertEqual(cm.exception.code, StatusCode.SHEET_INDEX_OUT_OF_RANGE)
                with self.assertRaises(IndexError):
                    self.wb[index]

    def test_index_beyond_u64_is_value_error(self) -> None:
        with self.assertRaises(ValueError):
            self.wb.sheet(2**64)

    def test_non_integer_key_is_type_error(self) -> None:
        for bad in (1.5, True, None):
            with self.subTest(bad=bad), self.assertRaises(TypeError):
                self.wb.sheet(bad)  # type: ignore[arg-type]

    def test_len_iter_contains(self) -> None:
        self.assertEqual(len(self.wb), 2)
        self.assertEqual([(s.index, s.name) for s in self.wb], [(0, "Sheet1"), (1, "Data")])
        self.assertIn("Sheet1", self.wb)
        self.assertIn("DATA", self.wb)
        self.assertNotIn("nope", self.wb)
        self.assertIn(1, self.wb)
        self.assertNotIn(2, self.wb)
        self.assertNotIn(-1, self.wb)


class SheetViewTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_get_and_set(self) -> None:
        self.ws.set_value("B3", 5)
        self.assertEqual(self.wb.sheet(0).get_typed_value(2, 1), Value(kind=ValueKind.NUMBER, number=5.0))
        self.assertEqual(self.ws.get_typed_value("B3"), Value(kind=ValueKind.NUMBER, number=5.0))

    def test_formula(self) -> None:
        self.ws.set_formula("C1", "=1+2")
        self.assertEqual(self.ws.get_formula("C1"), "=1+2")
        self.assertIsNone(self.ws.get_formula("D1"))
        self.wb.recalc()
        self.assertEqual(self.ws["C1"], 3.0)

    def test_mapping_protocol(self) -> None:
        self.ws["A1"] = 42
        self.assertEqual(self.ws["A1"], 42.0)
        self.ws["A2"] = "x"
        self.assertEqual(self.wb["Sheet1"]["A2"], "x")
        self.ws["A3"] = None
        self.assertIsNone(self.ws["A3"])
        self.ws["A4"] = ExcelError(ErrorCode.REF)
        self.assertEqual(self.ws["A4"], ExcelError(ErrorCode.REF))

    def test_name_is_live(self) -> None:
        self.wb.sheet(0).rename("Renamed")
        self.assertEqual(self.ws.name, "Renamed")

    def test_equivalent_spellings(self) -> None:
        self.ws.set_value("B3", 1)
        for a1 in ("b3", "$B$3", "B$3", "$b3"):
            with self.subTest(a1=a1):
                self.assertEqual(self.ws.get_typed_value(a1).number, 1.0)
        self.ws.set_value("A1", 2)
        for a1 in ("A01", "a0000001"):
            with self.subTest(a1=a1):
                self.assertEqual(self.ws.get_typed_value(a1).number, 2.0)

    def test_grid_edge(self) -> None:
        self.ws.set_value("XFD1048576", 9)
        self.assertEqual(self.wb.sheet(0).get_typed_value(1048575, 16383).number, 9.0)
        with self.assertRaises(InvalidArgumentError):
            self.ws.get_value("XFE1")

    def test_rejected_addresses(self) -> None:
        # Text that is no reference fails to parse; a valid reference that is not one bare cell is an invalid argument.
        unparseable = ("", "   ", "R1C1", "Ａ１", "Aé1", "\ud800A1")
        not_a_cell = ("A1:B2", "Sheet2!A1", "Sheet1!A1")
        for a1 in unparseable + not_a_cell:
            with self.subTest(a1=a1):
                with self.assertRaises(InvalidArgumentError) as cm:
                    self.ws.get_value(a1)
                want = StatusCode.PARSER_INVALID_REFERENCE if a1 in unparseable else StatusCode.INVALID_ARGUMENT
                self.assertEqual(cm.exception.code, want)
                with self.assertRaises(InvalidArgumentError):
                    self.ws.set_value(a1, 1)

    def test_non_string_address_is_type_error(self) -> None:
        with self.assertRaises(TypeError):
            self.ws.get_value(None)  # type: ignore[arg-type]

    def test_after_close(self) -> None:
        self.wb.close()
        self.assertIs(self.ws.workbook, self.wb)
        with self.assertRaises(WorkbookClosedError):
            self.ws.index
        with self.assertRaises(WorkbookClosedError):
            self.ws.get_value("A1")
        with self.assertRaises(WorkbookClosedError):
            self.ws.set_value("A1", 1)
        with self.assertRaises(WorkbookClosedError):
            self.ws.name
        with self.assertRaises(WorkbookClosedError):
            self.ws["A1"]
        with self.assertRaises(WorkbookClosedError):
            self.ws.get_formula("A1")
        with self.assertRaises(WorkbookClosedError):
            self.ws.set_formula("A1", "=1")


# Pairwise A1 model: column x row x absolute x case x decoration. `None` marks
# an address that must be rejected; otherwise the expected 0-based (row, col).
_A1_PAIRWISE = (
    ("P1", "$z0", None),
    ("P2", "Sheet1!xfe$9", None),
    ("P3", "$XFE1:B2", None),
    ("P4", " Z$01 ", None),
    ("P5", "aa$1048576:B2", None),
    ("P6", " $a$1 ", None),
    ("P7", "Sheet1!XFD1048576", None),
    ("P8", "$xfd$1048577", None),
    ("P9", "Sheet1!z01", None),
    ("P10", "Sheet1!$A1048577", None),
    ("P11", " xfe1048577 ", None),
    ("P12", "AA9", (8, 26)),
    ("P13", "$XFE$0:B2", None),
    ("P14", " $aa01 ", None),
    ("P15", "a01:B2", None),
    ("P16", "A$1", (0, 0)),
    ("P17", "$Z$1048576", (1048575, 25)),
    ("P18", "Z$1048577:B2", None),
    ("P19", " $xfe1048576 ", None),
    ("P20", " XFD$0 ", None),
    ("P21", "Sheet1!$aa$1", None),
    ("P22", "$xfd1:B2", None),
    ("P23", "Sheet1!a0", None),
    ("P24", "XFD$9:B2", None),
    ("P25", "XFE01", None),
    ("P26", "Sheet1!aa$1048577", None),
    ("P27", " $z$9 ", None),
    ("P28", "Sheet1!z1", None),
    ("P29", " AA0 ", None),
    ("P30", "A$1048576:B2", None),
    ("P31", "Sheet1!$xfd$01", None),
    ("P32", "$a9", (8, 0)),
)


class SheetA1PairwiseTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet("Sheet1")

    def test_pairwise_cases(self) -> None:
        for case_id, a1, expected in _A1_PAIRWISE:
            with self.subTest(case=case_id, a1=a1):
                if expected is None:
                    with self.assertRaises(InvalidArgumentError):
                        self.ws.get_value(a1)
                    with self.assertRaises(InvalidArgumentError):
                        self.ws.set_value(a1, 1)
                    continue
                row, col = expected
                self.ws.set_value(a1, float(row * 100000 + col))
                self.assertEqual(self.wb.sheet(0).get_typed_value(row, col).number, float(row * 100000 + col))
                self.wb.sheet(0).set_number(row, col, -1.0)
                self.assertEqual(self.ws.get_typed_value(a1).number, -1.0)


# ---------------------------------------------------------------------------
# S3: the public names of `Workbook` and `Sheet` are the ledger's.
# ---------------------------------------------------------------------------

_REPO_ROOT = pathlib.Path(__file__).resolve().parents[3]
_LEDGER = _REPO_ROOT / "tools" / "api" / "names.yaml"

# Hand-written members that are not ledger ops.
_FACADE_EXTRA = {
    "Workbook": {"sheet", "list_sheets", "close", "closed", "active_sheet", "set_active_sheet", "ooxml"},
    "Sheet": {"workbook", "index", "name", "ooxml"},
    "WorkbookOoxml": set(),
    "SheetOoxml": set(),
}


def _derived_py_names() -> "dict[str, set[str]]":
    """Public Python names per class: ledger op / list entries by scope, plus factories and facade extras."""
    import yaml  # PyYAML is a test dependency: a missing module fails the test, it does not skip it.

    with open(_LEDGER, encoding="utf-8") as f:
        ledger = yaml.safe_load(f)
    out = {key: set(names) for key, names in _FACADE_EXTRA.items()}
    for entry in ledger["entries"].values():
        if not isinstance(entry, dict):
            continue
        surfaces = entry.get("surfaces")
        if surfaces is not None and "py" not in surfaces:
            continue
        name = entry.get("py")
        if not isinstance(name, str) or (entry.get("drop") or {}).get("py") == name:
            continue
        if entry.get("internal"):
            continue
        handle = entry.get("handle")
        if handle is not None:
            classes = handle["class"] if isinstance(handle["class"], list) else [handle["class"]]
            if "Sheet" in classes and isinstance(handle.get("py"), str):
                out["Sheet"].add(handle["py"])
            continue
        if entry.get("kind") == "factory" and name.startswith("Workbook."):
            out["Workbook"].add(name.split(".", 1)[1])
        elif entry.get("kind") in ("op", "list") and entry.get("scope") in ("workbook", "sheet"):
            owner = "Workbook" if entry["scope"] == "workbook" else "Sheet"
            out[owner + "Ooxml" if entry.get("namespace") == "ooxml" else owner].add(name)
    return out


def _public_names(cls: type) -> "set[str]":
    return {name for name in dir(cls) if not name.startswith("_")}


class LedgerSurfaceTests(unittest.TestCase):
    def test_workbook_and_sheet_match_the_ledger(self) -> None:
        derived = _derived_py_names()
        for cls in (Workbook, Sheet, formulon.WorkbookOoxml, formulon.SheetOoxml):
            with self.subTest(cls=cls.__name__):
                actual = _public_names(cls)
                want = derived[cls.__name__]
                self.assertEqual(sorted(actual - want), [], "public names the ledger does not list")
                self.assertEqual(sorted(want - actual), [], "ledger names the class lacks")

    def test_no_sheet_index_entry_points(self) -> None:
        for name in ("get_sheet_count", "get_sheet_name", "get_sheet_index"):
            self.assertFalse(hasattr(Workbook, name), name)

    def test_private_helpers_are_not_exported(self) -> None:
        for name in ("_check", "_sint"):
            self.assertNotIn(name, formulon.__all__)
            self.assertFalse(hasattr(formulon, name), name)
        self.assertEqual(sorted(n for n in formulon.__all__ if n.startswith("_") and not n.startswith("__")), [])


# ---------------------------------------------------------------------------
# S4: `Sheet` handles follow structural edits.
# ---------------------------------------------------------------------------


class SheetHandleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.wb.add_sheet("Data")
        self.wb.add_sheet("Tail")
        self.addCleanup(self.wb.close)

    def _assert_in_sync(self) -> None:
        # H11: the handle list and the engine agree on count and names.
        sheets = self.wb.list_sheets()
        self.assertEqual(self.wb._sheet_count(), len(sheets))
        self.assertEqual([self.wb._sheet_name(i) for i in range(len(sheets))], [s.name for s in sheets])
        self.assertEqual([s.index for s in sheets], list(range(len(sheets))))

    def test_identity(self) -> None:
        # H1
        ws = self.wb.sheet(1)
        self.assertIs(self.wb.sheet(1), ws)
        self.assertIs(self.wb.sheet("Data"), ws)
        self.assertIs(self.wb.sheet("data"), ws)
        self.assertIs(self.wb[1], ws)
        self.assertIs(self.wb["DATA"], ws)
        self.assertIs(self.wb.sheet(ws), ws)
        self.assertIs(self.wb.list_sheets()[1], ws)
        self.assertEqual(list(self.wb), self.wb.list_sheets())
        self.assertEqual(ws, ws)
        self.assertEqual(hash(ws), hash(self.wb.sheet("Data")))

    def test_copy_before_keeps_handle(self) -> None:
        # H2
        ws = self.wb.sheet("Data")
        ws.set_value("A1", 7)
        copy = self.wb.copy_sheet(ws, 0)
        self.assertIsInstance(copy, Sheet)
        self.assertEqual(copy.index, 0)
        self.assertEqual(copy.name, "Data (2)")
        self.assertEqual(ws.index, 2)
        self.assertEqual(ws.name, "Data")
        self.assertEqual(copy["A1"], 7.0)
        self.assertIs(self.wb.sheet(0), copy)
        self._assert_in_sync()

    def test_copy_inserts_at_returned_index(self) -> None:
        copy = self.wb.copy_sheet("Sheet1", 2, "Mid")
        self.assertIs(self.wb.sheet(2), copy)
        self.assertEqual(copy.name, "Mid")
        self._assert_in_sync()

    def test_move_keeps_handle(self) -> None:
        # H3
        ws = self.wb.sheet("Sheet1")
        self.wb.sheet(ws).move(2)
        self.assertEqual(ws.index, 2)
        self.assertEqual(ws.name, "Sheet1")
        self.assertIs(self.wb.sheet(2), ws)
        self._assert_in_sync()

    def test_rename_keeps_handle(self) -> None:
        # H4
        ws = self.wb.sheet("Data")
        self.wb.sheet(ws).rename("Renamed")
        self.assertEqual(ws.name, "Renamed")
        self.assertIs(self.wb.sheet("renamed"), ws)
        self._assert_in_sync()

    def test_removed_handle_raises_not_found(self) -> None:
        # H5
        ws = self.wb.sheet("Data")
        other = self.wb.sheet("Tail")
        self.wb.remove_sheet(ws)
        for read in (lambda: ws.index, lambda: ws.name, lambda: ws.get_value("A1"), lambda: ws.set_number(0, 0, 1.0)):
            with self.assertRaises(NotFoundError) as cm:
                read()
            self.assertEqual(cm.exception.code, StatusCode.NOT_FOUND)
        self.assertEqual(other.name, "Tail")
        self.assertEqual(other.index, 1)
        self.assertNotIn(ws, self.wb)
        self._assert_in_sync()

    def test_removing_an_earlier_sheet_shifts_the_index(self) -> None:
        # H6
        tail = self.wb.sheet("Tail")
        self.wb.remove_sheet(0)
        self.assertEqual(tail.index, 1)
        self.assertEqual(tail.name, "Tail")
        self._assert_in_sync()

    def test_after_close(self) -> None:
        # H7, H13
        ws = self.wb.sheet(0)
        self.wb.close()
        for read in (lambda: ws.index, lambda: ws.name, lambda: ws.get_value("A1"), lambda: ws.set_number(0, 0, 1.0)):
            with self.assertRaises(WorkbookClosedError):
                read()
        self.assertIs(ws.workbook, self.wb)

    def test_bad_keys(self) -> None:
        # H8
        with self.assertRaises(NotFoundError):
            self.wb.sheet("nope")
        with self.assertRaises(KeyError):
            self.wb["nope"]
        with self.assertRaises(IndexOutOfRangeError):
            self.wb.sheet(99)
        with self.assertRaises(ValueError) as cm:
            self.wb.sheet(-1)
        self.assertNotIsInstance(cm.exception, FormulonError)
        for bad in (1.5, True, None):
            with self.subTest(bad=bad), self.assertRaises(TypeError):
                self.wb.sheet(bad)  # type: ignore[arg-type]

    def test_structure_edge_cases(self) -> None:
        # H14
        count = len(self.wb)
        tail = self.wb.copy_sheet(0, count)
        self.assertEqual(tail.index, count)
        with self.assertRaises(IndexOutOfRangeError) as cm:
            self.wb.copy_sheet(0, len(self.wb) + 1)
        self.assertEqual(cm.exception.code, StatusCode.SHEET_INDEX_OUT_OF_RANGE)
        ws = self.wb.sheet("Data")
        self.wb.sheet(ws).move(ws.index)
        self.assertEqual(ws.index, 1)
        self._assert_in_sync()
        with Workbook.create_default() as single:
            with self.assertRaises(FormulonError) as cm:
                single.remove_sheet(0)
            self.assertEqual(cm.exception.code, StatusCode.CANNOT_REMOVE_LAST_SHEET)
            self.assertEqual(len(single.list_sheets()), 1)

    def test_index_after_removal_names_the_next_sheet(self) -> None:
        # H15
        tail = self.wb.sheet(2)
        self.wb.remove_sheet(1)
        self.assertIs(self.wb.sheet(1), tail)

    def test_contains(self) -> None:
        # H16
        ws = self.wb.sheet("Data")
        self.assertIn(ws, self.wb)
        self.assertIn("Data", self.wb)
        self.assertIn("tail", self.wb)
        self.assertNotIn("nope", self.wb)
        self.assertIn(0, self.wb)
        self.assertIn(2, self.wb)
        self.assertNotIn(3, self.wb)
        self.assertNotIn(-1, self.wb)
        self.assertNotIn(True, self.wb)
        with Workbook.create_default() as other:
            self.assertNotIn(other.sheet(0), self.wb)

    def test_two_workbooks_never_share_handles(self) -> None:
        # H10
        with Workbook.create_default() as other:
            self.assertIsNot(other.sheet(0), self.wb.sheet(0))
            self.assertNotEqual(other.sheet(0), self.wb.sheet(0))

    def test_sheet_ref_forms(self) -> None:
        # H12
        ws = self.wb.sheet("Tail")
        self.wb.sheet(ws).rename("T1")
        self.wb.sheet(2).rename("T2")
        self.wb.sheet("t2").rename("T3")
        self.assertEqual(ws.name, "T3")
        self.wb.set_active_sheet(ws)
        self.assertIs(self.wb.active_sheet(), ws)
        self.wb.set_active_sheet("Data")
        self.assertIs(self.wb.active_sheet(), self.wb.sheet(1))
        self.wb.set_active_sheet(0)
        self.assertIs(self.wb.active_sheet(), self.wb.sheet(0))
        with Workbook.create_default() as other:
            with self.assertRaises(TypeError):
                self.wb.remove_sheet(other.sheet(0))
            with self.assertRaises(TypeError):
                self.wb.sheet(other.sheet(0))

    def test_add_sheet_returns_handle(self) -> None:
        ws = self.wb.add_sheet("More")
        self.assertIsInstance(ws, Sheet)
        self.assertEqual(ws.index, 3)
        self.assertIs(self.wb.sheet("More"), ws)
        self._assert_in_sync()

    def test_list_sheets_is_a_new_list(self) -> None:
        # L2
        sheets = self.wb.list_sheets()
        sheets.clear()
        self.assertEqual(len(self.wb.list_sheets()), 3)

    def test_repr(self) -> None:
        ws = self.wb.sheet("Data")
        self.assertEqual(repr(self.wb), "Workbook(sheets=3)")
        self.assertEqual(repr(ws), "Sheet(index=1, name='Data')")
        self.wb.remove_sheet(ws)
        self.assertEqual(repr(ws), "Sheet(removed)")
        self.wb.close()
        self.assertEqual(repr(self.wb), "Workbook(closed)")

    def test_value_repr_is_compact(self) -> None:
        self.assertEqual(repr(Value(kind=ValueKind.NUMBER, number=3)), "Value(NUMBER, 3.0)")
        self.assertEqual(repr(Value(kind=ValueKind.ERROR, error_code=ErrorCode.DIV0)), "Value(ERROR, DIV0)")
        self.assertEqual(repr(Value(kind=ValueKind.BOOL, boolean=True)), "Value(BOOL, True)")
        self.assertEqual(repr(Value(kind=ValueKind.TEXT, text="hi")), "Value(TEXT, 'hi')")
        self.assertEqual(repr(Value(kind=ValueKind.BLANK)), "Value(BLANK)")
        self.assertEqual(repr(_ARRAY_VALUE), "Value(ARRAY, 1x1)")
        self.assertEqual(repr(Value(kind=ValueKind.TEXT, text="x" * 40)), "Value(TEXT, '" + "x" * 32 + "...')")


# ---------------------------------------------------------------------------
# Fixtures shared by the shape sections below.
# ---------------------------------------------------------------------------


def _workbook_with_pivot() -> "tuple[Workbook, int, int]":
    """A default workbook with one two-field pivot on sheet 0: ``(wb, pivot_index, cache_id)``."""
    wb = Workbook.create_default()
    cache_id = wb.add_pivot_cache().cache_id
    wb.pivot_cache(cache_id).set_worksheet_source(ref="A1:B3", sheet="Sheet1")
    wb.pivot_cache(cache_id).add_field("Region")
    wb.pivot_cache(cache_id).add_field("Amount")
    for region, amount in (("East", 10.0), ("West", 30.0)):
        record = wb.pivot_cache(cache_id).add_record()
        wb.pivot_cache(cache_id).set_record_value(record, 0, region)
        wb.pivot_cache(cache_id).set_record_value(record, 1, amount)
    ws = wb.sheet(0)
    pivot = ws.add_pivot(PivotInput(name="Pivot1", cache_id=cache_id, anchor=CellAddress(0, 4)))
    ws.pivot(pivot).add_field(formulon.PivotFieldInput(source_name="Region", axis=formulon.PivotAxis.ROW))
    amount = ws.pivot(pivot).add_field(formulon.PivotFieldInput(source_name="Amount", axis=formulon.PivotAxis.VALUE))
    ws.pivot(pivot).add_data_field(formulon.PivotDataFieldInput(name="Sum of Amount", field_index=amount))
    return wb, pivot, cache_id


def _setter_kwargs(setter: object, record: object) -> "dict[str, object]":
    """``asdict(record)`` restricted to the keyword arguments ``setter`` takes."""
    params = inspect.signature(setter).parameters  # type: ignore[arg-type]
    return {k: v for k, v in dataclasses.asdict(record).items() if k in params}  # type: ignore[call-overload]


def _stated_kwargs(setter: object, record: object) -> "dict[str, object]":
    """The ``_setter_kwargs`` of a presence-flagged record, keeping only the fields its ``*_stated`` flag marks."""
    return {k: v for k, v in _setter_kwargs(setter, record).items() if getattr(record, f"{k}_stated")}


def _input_of(input_type: type, record: object) -> object:
    """Project ``record`` onto the fields of the ``*Input`` dataclass ``input_type``."""
    names = {f.name for f in dataclasses.fields(input_type)}
    return input_type(**{k: v for k, v in dataclasses.asdict(record).items() if k in names})  # type: ignore[call-overload]


# ---------------------------------------------------------------------------
# S6: every setter takes back what its getter returns.
# ---------------------------------------------------------------------------


class SetterRoundTripTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb, self.pivot, self.cache = _workbook_with_pivot()
        self.addCleanup(self.wb.close)
        self.ws = self.wb.sheet(0)

    def test_patch_records(self) -> None:
        self.wb.update_iterative(enabled=True, max_iterations=7, max_change=0.5)
        before = self.wb.get_iterative()
        self.wb.update_iterative(**dataclasses.asdict(before))
        self.assertEqual(self.wb.get_iterative(), before)

        self.wb.update_table_style_defaults(table_style="TableStyleLight1")
        before = self.wb.get_table_style_defaults()
        self.wb.update_table_style_defaults(**dataclasses.asdict(before))
        self.assertEqual(self.wb.get_table_style_defaults(), before)

        self.ws.pivot(self.pivot).update_grand_totals(rows=False)
        before = self.ws.pivot(self.pivot).get_grand_totals()
        self.assertEqual(before, formulon.PivotGrandTotals(rows=False, columns=True))
        self.ws.pivot(self.pivot).update_grand_totals(**dataclasses.asdict(before))
        self.assertEqual(self.ws.pivot(self.pivot).get_grand_totals(), before)

        self.ws.update_page_setup(orientation=formulon.PageOrientation.LANDSCAPE, scale_pct=80)
        before = self.ws.get_page_setup()
        self.ws.update_page_setup(**_stated_kwargs(self.ws.update_page_setup, before))
        self.assertEqual(self.ws.get_page_setup(), before)

        self.ws.update_page_margins(left_in=0.5)
        before = self.ws.get_page_margins()
        self.ws.update_page_margins(**_stated_kwargs(self.ws.update_page_margins, before))
        self.assertEqual(self.ws.get_page_margins(), before)

    def test_whole_records(self) -> None:
        self.ws.set_comment(1, 1, Comment("Ann", "Note"))
        self.ws.set_comment(1, 1, self.ws.get_comment(1, 1))
        self.assertEqual(self.ws.get_comment(1, 1), Comment("Ann", "Note"))

        titles = formulon.PrintTitles(rows=formulon.Span(0, 1), columns=None)
        self.ws.set_print_titles(titles)
        self.ws.set_print_titles(self.ws.get_print_titles())
        self.assertEqual(self.ws.get_print_titles(), titles)

        selection = formulon.SheetSelection(2, 3, [formulon.CellRect(2, 3, 4, 5)])
        self.ws.set_selection(active_row=2, active_col=3, ranges=selection.ranges)
        got = self.ws.get_selection()
        self.ws.set_selection(active_row=got.active_row, active_col=got.active_col, ranges=got.ranges)
        self.assertEqual(self.ws.get_selection(), selection)

        protection = formulon.SheetProtectionInput(hash_value="abc")
        self.ws.set_sheet_protection(protection)
        self.ws.set_sheet_protection(self.ws.get_sheet_protection())
        self.assertEqual(self.ws.get_sheet_protection().hash_value, "abc")

        workbook_protection = formulon.WorkbookProtection(lock_structure=True)
        self.wb.set_workbook_protection(workbook_protection)
        self.wb.set_workbook_protection(self.wb.get_workbook_protection())
        self.assertEqual(self.wb.get_workbook_protection(), workbook_protection)

        fonts = self.wb.get_theme_fonts()
        self.wb.set_theme_fonts(fonts)
        self.assertEqual(self.wb.get_theme_fonts(), fonts)

        self.ws.set_number(0, 0, 1.0)
        self.ws.set_array_formula(formulon.ArrayFormula(formulon.CellRect(0, 8, 1, 8), "=A1*2"))
        array = self.ws.get_array_formula(0, 8)
        self.ws.set_array_formula(array)
        self.assertEqual(self.ws.get_array_formula(1, 8), array)

        self.wb.set_pinned_now(datetime.datetime(2026, 4, 23, 15, 30, 45, 999))
        self.wb.set_pinned_now(self.wb.get_pinned_now())
        self.assertEqual(self.wb.get_pinned_now(), datetime.datetime(2026, 4, 23, 15, 30, 45))

        runs = [formulon.PhoneticRun(start=0, end=1, text="カ")]
        self.ws.set_text(3, 0, "漢")
        self.ws.set_cell_phonetic_runs(3, 0, runs)
        self.ws.set_cell_phonetic_runs(3, 0, self.ws.get_cell_phonetic_runs(3, 0))
        self.assertEqual(self.ws.get_cell_phonetic_runs(3, 0), runs)

    def test_input_records(self) -> None:
        self.wb.set_defined_name(formulon.DefinedNameInput("Rate", "Sheet1!$A$1", comment="c", hidden=True))
        got = self.wb.get_defined_name("Rate")
        self.wb.set_defined_name(_input_of(formulon.DefinedNameInput, got))
        self.assertEqual(self.wb.get_defined_name("Rate"), got)

        colors = self.wb.get_theme_colors()
        self.wb.set_theme_colors(colors=colors.colors)
        self.assertEqual(self.wb.get_theme_colors().colors, colors.colors)

        script = self.wb.get_theme_script_font(major=True, script="Jpan")
        self.wb.set_theme_script_font(major=True, script=script.script, typeface=script.typeface)
        self.assertEqual(self.wb.get_theme_script_font(major=True, script="Jpan"), script)

        style = self.ws.pivot(self.pivot).get_style_options()
        self.ws.pivot(self.pivot).update_style_options(**dataclasses.asdict(style))
        self.assertEqual(self.ws.pivot(self.pivot).get_style_options(), style)


# ---------------------------------------------------------------------------
# S7: "absent" reads as None, and None / remove clears.
# ---------------------------------------------------------------------------


class AbsentValueTests(unittest.TestCase):
    """A1: a fixture workbook whose objects state none of their optional attributes."""

    @classmethod
    def setUpClass(cls) -> None:
        wb, pivot, cache = _workbook_with_pivot()
        ws = wb.sheet(0)
        ws.set_text(0, 0, "Region")
        ws.set_text(0, 1, "Amount")
        ws.set_hyperlink(10, 0, formulon.HyperlinkInput(target="https://example.test/"))
        ws.add_validation(formulon.DataValidationInput(ranges=[formulon.CellRect(11, 0, 11, 0)], type=1, formula1="1"))
        table = ws.add_table(
            formulon.TableInput(range=formulon.CellRect(20, 0, 22, 1), name="T", columns=["A", "B"], style_name="")
        )
        wb.set_defined_name(formulon.DefinedNameInput("Rate", "Sheet1!$A$1"))
        chart = ws.add_chart(formulon.ChartInsertKind.COLUMN, "Sheet1!$A$1:$B$3", legend=None)
        ws.add_sparkline_group(
            formulon.SparklineGroup(sparklines=[formulon.Sparkline(formula="Sheet1!A1:B1", location="C1")])
        )
        cls.wb, cls.pivot, cls.cache, cls.table, cls.chart = wb, pivot, cache, table, chart

    @classmethod
    def tearDownClass(cls) -> None:
        cls.wb.close()

    def test_absent_reads_are_none(self) -> None:
        wb, ws = self.wb, self.wb.sheet(0)
        self.assertIsNone(wb.get_file_name())
        self.assertIsNone(ws.get_cell_phonetic(0, 0))
        self.assertIsNone(ws.ooxml.get_auto_filter_xml())
        for getter in (
            ws.ooxml.get_page_setup_xml,
            ws.ooxml.get_page_margins_xml,
            ws.ooxml.get_print_options_xml,
            ws.ooxml.get_header_footer_xml,
        ):
            with self.subTest(getter=getter.__name__):
                self.assertIsNone(getter())
        self.assertIsNone(ws.get_print_titles())
        self.assertIsNone(ws.get_print_area())
        self.assertIsNone(wb.pivot_cache(self.cache).list_fields()[0].formula)
        self.assertIsNone(wb.get_defined_name("Rate").comment)
        self.assertIsNone(ws.list_tables()[self.table.index].get_style_options())
        self.assertEqual(wb.get_table_style_defaults(), formulon.TableStyleDefaults(None, None))
        hyperlink = ws.get_hyperlink(10, 0)
        self.assertEqual((hyperlink.display, hyperlink.tooltip, hyperlink.location), (None, None, None))
        validation = ws.list_validations()[0]
        self.assertEqual((validation.formula2, validation.error_title, validation.prompt_message), (None, None, None))
        chart = ws.chart(self.chart).get_properties()
        self.assertIsNone(chart.legend)
        drawing = ws.drawing(self.chart).get_drawing_properties()
        self.assertFalse(hasattr(drawing, "media_path"))
        self.assertIsNone(ws.list_sparkline_groups()[0].date_axis_ref)
        data_field = ws.pivot(self.pivot).list_data_fields()[0]
        self.assertEqual(
            (data_field.number_format, data_field.show_as_base_field, data_field.show_as_base_item), (None, None, None)
        )
        field = ws.pivot(self.pivot).list_fields()[0]
        self.assertEqual((field.custom_name, field.group_base_index), (None, None))
        self.assertIsNone(ws.get_page_setup().orientation)
        self.assertEqual(ws.get_view().view_mode, formulon.SheetViewMode.NORMAL)
        protection = wb.get_workbook_protection()
        self.assertEqual((protection.workbook_hash_value, protection.workbook_legacy_password_hash), (None, None))

    def test_pinned_now_none_releases(self) -> None:
        # A2
        with Workbook.create_default() as wb:
            wb.set_pinned_now(datetime.datetime(2026, 1, 2))
            wb.set_pinned_now(None)
            self.assertIsNone(wb.get_pinned_now())

    def test_defined_name_scope_none_is_workbook(self) -> None:
        # A3, A9
        with Workbook.create_default() as wb:
            wb.set_defined_name(formulon.DefinedNameInput("Rate", "Sheet1!$A$1", None, comment="x", hidden=True))
            self.assertIsNone(wb.list_defined_names()[0].local_sheet_id)
            wb.set_defined_name(formulon.DefinedNameInput("Rate", "Sheet1!$B$1"))
            self.assertEqual(
                wb.get_defined_name("Rate"),
                formulon.DefinedName("Rate", "=Sheet1!$B$1", None, comment=None, hidden=False),
            )

    def test_print_area_none_and_empty_clear(self) -> None:
        # A4
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for clear in (None, []):
                with self.subTest(clear=clear):
                    ws.set_print_area([formulon.CellRect(0, 0, 3, 3)])
                    ws.set_print_area(clear)
                    self.assertIsNone(ws.get_print_area())

    def test_pivot_field_axis_none(self) -> None:
        # A5
        wb, pivot, _ = _workbook_with_pivot()
        with wb:
            ws = wb.sheet(0)
            ws.pivot(pivot).set_field_axis(0, None)
            self.assertIsNone(ws.pivot(pivot).list_fields()[0].axis)
            with self.assertRaises(InvalidArgumentError):
                ws.pivot(pivot).set_field_axis(0, -1)  # type: ignore[arg-type]

    def test_patch_omitted_none_and_value(self) -> None:
        # A6
        with Workbook.create_default() as wb:
            wb.update_table_style_defaults(table_style="TableStyleLight1", pivot_style="PivotStyleLight2")
            wb.update_table_style_defaults(pivot_style=None)
            self.assertEqual(wb.get_table_style_defaults(), formulon.TableStyleDefaults("TableStyleLight1", None))
            wb.update_iterative(max_iterations=5)
            self.assertEqual(wb.get_iterative().max_iterations, 5)
            with self.assertRaises(TypeError):
                wb.update_iterative(enabled=None)  # type: ignore[arg-type]

    def test_empty_values(self) -> None:
        # A7, A8
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            with self.assertRaises(InvalidArgumentError):
                wb.set_defined_name(formulon.DefinedNameInput("Rate", ""))
            ws.set_comment(0, 0, Comment("Ann", ""))
            self.assertEqual(ws.get_comment(0, 0), Comment("Ann", ""))
            ws.remove_comment(0, 0)
            self.assertIsNone(ws.get_comment(0, 0))
            with self.assertRaises(NotFoundError):
                ws.remove_comment(0, 0)

    def test_none_clears_protection_and_phonetic_runs(self) -> None:
        # A10
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            wb.set_workbook_protection(formulon.WorkbookProtection(lock_structure=True))
            wb.set_workbook_protection(None)
            self.assertEqual(wb.get_workbook_protection(), formulon.WorkbookProtection())
            ws.set_sheet_protection(formulon.SheetProtectionInput())
            self.assertIsNotNone(ws.get_sheet_protection())
            ws.set_sheet_protection(None)
            self.assertIsNone(ws.get_sheet_protection())
            for clear in (None, []):
                with self.subTest(clear=clear):
                    ws.set_text(0, 0, "漢")
                    ws.set_cell_phonetic_runs(0, 0, [formulon.PhoneticRun(start=0, end=1, text="カ")])
                    ws.set_cell_phonetic_runs(0, 0, clear)
                    self.assertEqual(ws.get_cell_phonetic_runs(0, 0), [])

    def test_chart_title_and_legend_none(self) -> None:
        # A11
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for r in range(3):
                ws.set_number(r, 0, float(r))
            chart = ws.add_chart(formulon.ChartInsertKind.COLUMN, "Sheet1!$A$1:$A$3", title="T")
            ws.chart(chart).set_title(None)
            ws.chart(chart).set_legend(None)
            info = ws.chart(chart).get_properties()
            self.assertEqual((info.title, info.legend), (None, None))


# ---------------------------------------------------------------------------
# P / L / N / D boundary rows.
# ---------------------------------------------------------------------------


class PatchBoundaryTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)

    def test_empty_patch_changes_nothing(self) -> None:
        # P1
        before = (self.wb.get_iterative(), self.wb.get_table_style_defaults(), self.wb.sheet(0).get_page_setup())
        self.wb.update_iterative()
        self.wb.update_table_style_defaults()
        self.wb.sheet(0).update_page_setup()
        self.assertEqual(
            (self.wb.get_iterative(), self.wb.get_table_style_defaults(), self.wb.sheet(0).get_page_setup()), before
        )

    def test_unchanged_outside_a_patch(self) -> None:
        # P2
        with self.assertRaises(TypeError):
            self.wb.sheet(0).set_number(formulon.UNCHANGED, 0, 1.0)
        with self.assertRaises(TypeError):
            self.wb.sheet(0).set_text(0, 0, formulon.UNCHANGED)

    def test_unchanged_is_a_singleton(self) -> None:
        self.assertIs(copy.deepcopy(formulon.UNCHANGED), formulon.UNCHANGED)
        self.assertIs(pickle.loads(pickle.dumps(formulon.UNCHANGED)), formulon.UNCHANGED)
        self.assertEqual(repr(formulon.UNCHANGED), "formulon.UNCHANGED")

    def test_iterative_number_checks(self) -> None:
        # P3
        for bad in (1.5, math.nan, "3"):
            with self.subTest(bad=bad), self.assertRaises(TypeError):
                self.wb.update_iterative(max_iterations=bad)  # type: ignore[arg-type]
        for bad in (math.inf, math.nan):
            with self.subTest(bad=bad), self.assertRaises(InvalidArgumentError):
                self.wb.update_iterative(max_change=bad)

    def test_local_sheet_id_range(self) -> None:
        # P4
        with self.assertRaises(IndexError):
            self.wb.set_defined_name(formulon.DefinedNameInput("Rate", "Sheet1!$A$1", 2**31))
        with self.assertRaises(IndexError):
            self.wb.get_defined_name("Rate", 2**31)

    def test_unknown_keyword_is_type_error(self) -> None:
        # O1, O3
        with self.assertRaises(TypeError):
            self.wb.update_iterative(enabeld=True)  # type: ignore[call-arg]
        with self.assertRaises(TypeError):
            self.wb.sheet(0).set_hyperlink(0, 0, bogus=1)
        with self.assertRaises(TypeError):
            formulon.Comment(author="Ann")  # type: ignore[call-arg]


class ListBoundaryTests(unittest.TestCase):
    def test_empty_and_fresh_lists(self) -> None:
        # L1, L2
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for listed in (ws.list_cf_rules, ws.list_validations, ws.list_protected_ranges, ws.list_sparkline_groups):
                with self.subTest(listed=listed.__name__):
                    self.assertEqual(listed(), [])
            ws.add_validation(
                formulon.DataValidationInput(ranges=[formulon.CellRect(0, 0, 0, 0)], type=1, formula1="1")
            )
            first = ws.list_validations()
            first.clear()
            self.assertEqual(len(ws.list_validations()), 1)
            self.assertIsNot(ws.list_validations(), ws.list_validations())


class EnumBoundaryTests(unittest.TestCase):
    def test_families_agree(self) -> None:
        # N1, S5
        self.assertEqual(formulon.SlicerSortOrder.DESCENDING, formulon.PivotSortOrder.DESCENDING)
        self.assertEqual(formulon.SlicerSortOrder.ASCENDING, formulon.PivotSortOrder.ASCENDING)
        self.assertEqual(formulon.AnchorEditAs.ONE_CELL, formulon.AnchorKind.ONE_CELL)
        self.assertEqual(formulon.AnchorEditAs.TWO_CELL, formulon.AnchorKind.TWO_CELL)
        self.assertEqual(formulon.ChartBlanksAs.GAP, formulon.SparklineEmptyCells.GAP)
        self.assertEqual(formulon.ChartBlanksAs.SPAN, formulon.SparklineEmptyCells.SPAN)
        self.assertEqual(formulon.ChartAxisPosition.LEFT, formulon.ChartLegendPosition.LEFT)
        self.assertEqual(formulon.ChartAxisPosition.RIGHT, formulon.ChartLegendPosition.RIGHT)

    def test_out_of_range_and_absent_values(self) -> None:
        # N2, N4
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for bad in (99, -1):
                with self.subTest(bad=bad), self.assertRaises(InvalidArgumentError):
                    ws.update_view(view_mode=bad)  # type: ignore[arg-type]
            for bad in (1.5, math.nan):
                with self.subTest(bad=bad), self.assertRaises(ValueError):
                    ws.update_view(view_mode=bad)  # type: ignore[arg-type]
            with self.assertRaises(TypeError):
                ws.update_view(view_mode="0")  # type: ignore[arg-type]

    def test_absent_members_are_gone(self) -> None:
        for enum, member in (
            (formulon.ChartBlanksAs, "ABSENT"),
            (formulon.ChartAxisCrosses, "ABSENT"),
            (formulon.ShapePaint, "UNSET"),
            (formulon.PivotFilterOperandKind, "NONE"),
            (formulon.ChartTitleKind, "NONE"),
            (formulon.ChartLegendPosition, "NONE"),
        ):
            with self.subTest(enum=enum.__name__):
                self.assertFalse(hasattr(enum, member))


class DateTimeBoundaryTests(unittest.TestCase):
    def test_pinned_now_inputs(self) -> None:
        # D1
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError):
                wb.set_pinned_now(datetime.datetime(2026, 1, 1, tzinfo=datetime.timezone.utc))
            with self.assertRaises(TypeError):
                wb.set_pinned_now(datetime.date(2026, 1, 1))  # type: ignore[arg-type]
            with self.assertRaises(InvalidArgumentError):
                wb.set_pinned_now(datetime.datetime(1899, 12, 31))
            self.assertIsNone(wb.get_pinned_now())


class PivotDetailsHandleTests(unittest.TestCase):
    def test_show_pivot_details_returns_a_tracked_handle(self) -> None:
        # H9
        wb, pivot, _ = _workbook_with_pivot()
        with wb:
            ws = wb.sheet(0)
            layout = ws.pivot(pivot).get_layout()
            row, col = next((c.row, c.col) for c in layout.cells if c.value.kind == ValueKind.NUMBER)
            details = ws.pivot(pivot).show_details(row, col)
            self.assertIsInstance(details, Sheet)
            self.assertEqual(details.index, 0)
            self.assertEqual(ws.index, 1)
            self.assertIs(wb.sheet(0), details)
            self.assertIs(wb.active_sheet(), details)
            self.assertEqual([s.name for s in wb.list_sheets()], [wb._sheet_name(i) for i in range(len(wb))])


if __name__ == "__main__":
    unittest.main()
