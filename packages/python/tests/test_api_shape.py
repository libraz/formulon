"""Shape of the idiomatic Python surface.

Covers the exception classes, ``ExcelError``, ``load`` / ``save`` with paths
and streams, the ``closed`` lifecycle predicate, ``set_value``,
``get_sheet_index`` and the A1-addressed ``Sheet`` view.
"""

from __future__ import annotations

import io
import math
import pathlib
import tempfile
import unittest

import formulon
from formulon import (
    ErrorCode,
    ExcelError,
    FormulonError,
    IndexOutOfRangeError,
    InvalidArgumentError,
    NotFoundError,
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
                wb.get_sheet_name(5)
            self.assertEqual(cm.exception.code, StatusCode.SHEET_INDEX_OUT_OF_RANGE)
            self.assertEqual(cm.exception.context, "sheet_index=5 sheet_count=1")
            self.assertEqual(str(cm.exception), f"{cm.exception.message} [sheet_index=5 sheet_count=1]")

    def test_collection_index_out_of_range_is_index_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(IndexOutOfRangeError) as cm:
                wb.get_font(wb.get_font_count())
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
                wb.get_sheet_name(5)
            self.assertNotIn("fm_", str(cm.exception))
            self.assertNotIn("SHEET_INDEX_OUT_OF_RANGE", str(cm.exception))

    def test_message_without_context_has_no_brackets(self) -> None:
        wb = Workbook.create_default()
        wb.close()
        with self.assertRaises(WorkbookClosedError) as cm:
            wb.get_sheet_count()
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
        v = formulon.eval_formula("=1/0")
        self.assertIs(v.kind, ValueKind.ERROR)
        self.assertIs(v.error_code, ErrorCode.DIV0)
        self.assertEqual(v.to_python(), ExcelError(ErrorCode.DIV0))

    def test_value_error_code_is_error_code(self) -> None:
        self.assertIs(Value(kind=ValueKind.ERROR, error_code=6).error_code, ErrorCode.NA)

    def test_payloadless_kinds_return_value(self) -> None:
        v = Value(kind=ValueKind.ARRAY)
        self.assertIs(v.to_python(), v)


class LoadSaveTests(unittest.TestCase):
    def setUp(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 42.0)
            self.xlsx = wb.save()
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.dir = pathlib.Path(tmp.name)

    def _assert_loaded(self, wb: Workbook) -> None:
        with wb:
            self.assertEqual(wb.get_value(0, 0, 0).number, 42.0)

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
            wb.get_value(0, 0, 0)
        with self.assertRaises(WorkbookClosedError):
            wb.save()


class SetValueTests(unittest.TestCase):
    def setUp(self) -> None:
        self.wb = Workbook.create_default()
        self.addCleanup(self.wb.close)

    def _stored(self, value: object) -> Value:
        self.wb.set_value(0, 0, 0, value)  # type: ignore[arg-type]
        return self.wb.get_value(0, 0, 0)

    def test_zero_and_negative_zero(self) -> None:
        self.assertEqual(self._stored(0), Value(kind=ValueKind.NUMBER, number=0.0))
        v = self._stored(-0.0)
        self.assertIs(v.kind, ValueKind.NUMBER)
        self.assertEqual(v.number, 0.0)

    def test_non_finite_is_invalid_argument(self) -> None:
        for bad in (math.nan, math.inf, -math.inf):
            with self.subTest(bad=bad), self.assertRaises(InvalidArgumentError):
                self.wb.set_value(0, 0, 0, bad)

    def test_none_is_blank(self) -> None:
        self.wb.set_number(0, 0, 0, 1.0)
        self.assertIs(self._stored(None).kind, ValueKind.BLANK)

    def test_missing_value_is_type_error(self) -> None:
        with self.assertRaises(TypeError):
            self.wb.set_value(0, 0, 0)  # type: ignore[call-arg]

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
            self.wb.set_value(0, 0, 0, 2**1024)
        self.assertNotIsInstance(cm.exception, FormulonError)

    def test_formula_like_string_is_text(self) -> None:
        self.assertEqual(self._stored("=A1"), Value(kind=ValueKind.TEXT, text="=A1"))
        self.assertIsNone(self.wb.get_formula(0, 0, 0))

    def test_excel_error(self) -> None:
        self.assertEqual(self._stored(ExcelError(ErrorCode.NA)), Value(kind=ValueKind.ERROR, error_code=ErrorCode.NA))

    def test_unsupported_types_are_type_error(self) -> None:
        for bad in (object(), [1], {"a": 1}, b"bytes", Value(kind=ValueKind.ARRAY), Value(kind=ValueKind.REF)):
            with self.subTest(bad=bad), self.assertRaises(TypeError):
                self.wb.set_value(0, 0, 0, bad)  # type: ignore[arg-type]

    def test_scalar_value_round_trip(self) -> None:
        for v in (
            Value(kind=ValueKind.BLANK),
            Value(kind=ValueKind.NUMBER, number=3.25),
            Value(kind=ValueKind.BOOL, boolean=True),
            Value(kind=ValueKind.TEXT, text="hello"),
            Value(kind=ValueKind.ERROR, error_code=ErrorCode.DIV0),
        ):
            with self.subTest(kind=v.kind.name):
                self.wb.set_value(0, 1, 1, v)
                self.assertEqual(self.wb.get_value(0, 1, 1), v)


class SheetIndexTests(unittest.TestCase):
    def test_get_sheet_index(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Data")
            self.assertEqual(wb.get_sheet_index("Sheet1"), 0)
            self.assertEqual(wb.get_sheet_index("data"), 1)
            self.assertIsNone(wb.get_sheet_index("nope"))
            self.assertIsNone(wb.get_sheet_index(""))
            with self.assertRaises(TypeError):
                wb.get_sheet_index(None)  # type: ignore[arg-type]


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
        self.ws.set("B3", 5)
        self.assertEqual(self.wb.get_value(0, 2, 1), Value(kind=ValueKind.NUMBER, number=5.0))
        self.assertEqual(self.ws.get("B3"), Value(kind=ValueKind.NUMBER, number=5.0))

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
        self.wb.rename_sheet(0, "Renamed")
        self.assertEqual(self.ws.name, "Renamed")

    def test_equivalent_spellings(self) -> None:
        self.ws.set("B3", 1)
        for a1 in ("b3", "$B$3", "B$3", "$b3"):
            with self.subTest(a1=a1):
                self.assertEqual(self.ws.get(a1).number, 1.0)
        self.ws.set("A1", 2)
        for a1 in ("A01", "a0000001"):
            with self.subTest(a1=a1):
                self.assertEqual(self.ws.get(a1).number, 2.0)

    def test_grid_edge(self) -> None:
        self.ws.set("XFD1048576", 9)
        self.assertEqual(self.wb.get_value(0, 1048575, 16383).number, 9.0)
        with self.assertRaises(InvalidArgumentError):
            self.ws.get("XFE1")

    def test_rejected_addresses(self) -> None:
        for a1 in ("", "   ", "A1:B2", "Sheet2!A1", "Sheet1!A1", "R1C1", "Ａ１", "Aé1", "\ud800A1"):
            with self.subTest(a1=a1):
                with self.assertRaises(InvalidArgumentError) as cm:
                    self.ws.get(a1)
                self.assertEqual(cm.exception.code, StatusCode.INVALID_ARGUMENT)
                with self.assertRaises(InvalidArgumentError):
                    self.ws.set(a1, 1)

    def test_non_string_address_is_type_error(self) -> None:
        with self.assertRaises(TypeError):
            self.ws.get(None)  # type: ignore[arg-type]

    def test_after_close(self) -> None:
        self.wb.close()
        self.assertEqual(self.ws.index, 0)
        self.assertIs(self.ws.workbook, self.wb)
        with self.assertRaises(WorkbookClosedError):
            self.ws.get("A1")
        with self.assertRaises(WorkbookClosedError):
            self.ws.set("A1", 1)
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
                        self.ws.get(a1)
                    with self.assertRaises(InvalidArgumentError):
                        self.ws.set(a1, 1)
                    continue
                row, col = expected
                self.ws.set(a1, float(row * 100000 + col))
                self.assertEqual(self.wb.get_value(0, row, col).number, float(row * 100000 + col))
                self.wb.set_number(0, row, col, -1.0)
                self.assertEqual(self.ws.get(a1).number, -1.0)


if __name__ == "__main__":
    unittest.main()
