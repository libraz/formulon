"""Status codes and the exception family raised for them."""

from __future__ import annotations

import re
import unittest
from pathlib import Path

import formulon
from formulon import (
    FormulonError,
    IndexOutOfRangeError,
    InvalidArgumentError,
    NotFoundError,
    StatusCode,
    Workbook,
)
from formulon._c import LIB, FormulaSyntaxError, ParseError, ParseErrorCode
from formulon._status import ERROR_RANGES, range_class_name

_STATUS_HEADER = Path(__file__).resolve().parents[3] / "src" / "c_api" / "formulon_status.h"


def _header_codes() -> dict[str, int]:
    text = _STATUS_HEADER.read_text(encoding="utf-8")
    return {name: int(value) for name, value in re.findall(r"^#define FM_STATUS_(\w+) \((\d+)\)$", text, re.M)}


class StatusCodeTest(unittest.TestCase):
    def test_enum_matches_the_c_header(self) -> None:
        self.assertEqual({member.name: member.value for member in StatusCode}, _header_codes())

    def test_names_are_upper_snake(self) -> None:
        for member in StatusCode:
            with self.subTest(member=member.name):
                self.assertRegex(member.name, r"^[A-Z][A-Z0-9]*(_[A-Z0-9]+)*$")

    def test_boundary_spelling(self) -> None:
        self.assertEqual(StatusCode.SHEET_INDEX_OUT_OF_RANGE, 10)
        self.assertEqual(StatusCode.PARSER_BOM_NOT_SUPPORTED, 1013)
        self.assertEqual(StatusCode.OK, 0)

    def test_structural_codes_sit_in_the_general_band(self) -> None:
        self.assertEqual(
            (
                StatusCode.SHEET_INDEX_OUT_OF_RANGE,
                StatusCode.INVALID_SHEET_NAME,
                StatusCode.INDEX_OUT_OF_RANGE,
            ),
            (10, 11, 13),
        )
        for member in (
            StatusCode.SHEET_INDEX_OUT_OF_RANGE,
            StatusCode.INVALID_SHEET_NAME,
            StatusCode.INDEX_OUT_OF_RANGE,
        ):
            with self.subTest(member=member.name):
                self.assertEqual(range_class_name(int(member)), "FormulonError")

    def test_retired_structural_numbers_are_not_assigned(self) -> None:
        values = {int(member) for member in StatusCode}
        for retired in (12, 5050, 5051, 5053, 5054, 9005):
            self.assertNotIn(retired, values)
        for gone in ("SHEET_COUNT_LIMIT_EXCEEDED", "PRINT_PAGE_COUNT_OVERFLOW"):
            self.assertFalse(hasattr(StatusCode, gone), gone)

    def test_renamed_members_replace_the_abbreviated_ones(self) -> None:
        self.assertEqual(StatusCode.FUNCTION_OVERFLOW, 3004)
        self.assertEqual(StatusCode.SECURITY_RESOURCE_LIMIT, 6005)
        for gone in ("FN_OVERFLOW", "SEC_RESOURCE_LIMIT"):
            self.assertFalse(hasattr(StatusCode, gone), gone)
        self.assertEqual(range_class_name(3004), "FunctionError")
        self.assertEqual(range_class_name(6005), "SecurityError")

    def test_every_code_lies_in_a_module_band(self) -> None:
        for member in StatusCode:
            with self.subTest(member=member.name):
                self.assertTrue(any(first <= int(member) <= last for first, last, _ in ERROR_RANGES))

    def test_c_status_string_uses_the_same_names(self) -> None:
        for member in (
            StatusCode.OK,
            StatusCode.INVALID_ARGUMENT,
            StatusCode.IO_ZIP_BOMB,
            StatusCode.BINDING_INVALID_HANDLE,
        ):
            with self.subTest(member=member.name):
                self.assertEqual(LIB.read_cstr(LIB.fm_get_status_name(int(member))), member.name)
        self.assertEqual(LIB.read_cstr(LIB.fm_get_status_name(123456)), "UNRECOGNIZED_STATUS")


class ExceptionFamilyTest(unittest.TestCase):
    def test_hierarchy(self) -> None:
        self.assertTrue(issubclass(InvalidArgumentError, FormulonError))
        self.assertTrue(issubclass(InvalidArgumentError, ValueError))
        self.assertTrue(issubclass(NotFoundError, FormulonError))
        self.assertTrue(issubclass(NotFoundError, KeyError))
        self.assertTrue(issubclass(NotFoundError, LookupError))
        self.assertFalse(issubclass(NotFoundError, ValueError))
        self.assertTrue(issubclass(IndexOutOfRangeError, FormulonError))
        self.assertTrue(issubclass(IndexOutOfRangeError, IndexError))
        self.assertTrue(issubclass(IndexOutOfRangeError, LookupError))
        self.assertFalse(issubclass(IndexOutOfRangeError, NotFoundError))
        self.assertTrue(issubclass(formulon.WorkbookClosedError, FormulonError))
        self.assertTrue(issubclass(formulon.WorkbookClosedError, ValueError))
        self.assertIs(formulon.InvalidArgumentError, InvalidArgumentError)

    def test_code_is_the_enum_member(self) -> None:
        err = FormulonError(StatusCode.PARSER_INVALID_REFERENCE)
        self.assertIs(err.code, StatusCode.PARSER_INVALID_REFERENCE)
        self.assertEqual(err.code.name, "PARSER_INVALID_REFERENCE")
        self.assertFalse(hasattr(err, "status"))
        self.assertFalse(hasattr(err, "status_name"))
        self.assertEqual(str(err), "[PARSER_INVALID_REFERENCE]")

    def test_string_form_is_the_code_name_and_the_message_without_context(self) -> None:
        err = FormulonError(StatusCode.INVALID_ARGUMENT, "bad", "k=v")
        self.assertEqual(str(err), "[INVALID_ARGUMENT] bad")
        self.assertNotIn("k=v", str(err))
        self.assertEqual(err.context, "k=v")

    def test_host_constructor_takes_code_message_context_and_parse_error(self) -> None:
        err = FormulonError(StatusCode.NOT_FOUND, "x", "k=v")
        self.assertIsInstance(err, NotFoundError)
        self.assertEqual(
            (err.code, err.message, err.context, err.parse_error), (StatusCode.NOT_FOUND, "x", "k=v", None)
        )
        bare = FormulonError(StatusCode.INVALID_ARGUMENT)
        self.assertEqual((bare.message, bare.context, bare.parse_error), ("", "", None))
        record = ParseError(ParseErrorCode.UNEXPECTED_EOF, 3, 3, 1, 4)
        syntax = FormulonError(StatusCode.PARSER_UNTERMINATED_STRING, "eof", "", record)
        self.assertIsInstance(syntax, FormulaSyntaxError)
        self.assertIs(syntax.parse_error, record)

    def test_engine_failure_string_carries_no_context(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(IndexOutOfRangeError) as caught:
                wb.sheet(5)
        self.assertTrue(caught.exception.context)
        self.assertEqual(str(caught.exception), f"[SHEET_INDEX_OUT_OF_RANGE] {caught.exception.message}")
        self.assertNotIn(caught.exception.context, str(caught.exception))

    def test_syntax_error_keeps_only_the_parse_error_record(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulaSyntaxError) as caught:
                wb.sheet(0).set_formula("A1", "=SUM(1,")
        error = caught.exception
        self.assertIsInstance(error.parse_error, ParseError)
        self.assertGreaterEqual(error.parse_error.start, 0)
        for retired in ("start", "end", "line", "column", "parse_error_code"):
            self.assertFalse(hasattr(error, retired), retired)

    def test_dedicated_subclass_per_code(self) -> None:
        self.assertIs(type(FormulonError(StatusCode.INVALID_ARGUMENT)), InvalidArgumentError)
        for member in (StatusCode.NOT_FOUND,):
            with self.subTest(member=member.name):
                self.assertIs(type(FormulonError(member)), NotFoundError)
        self.assertIs(type(FormulonError(StatusCode.ALREADY_EXISTS)), formulon.AlreadyExistsError)
        self.assertIs(type(FormulonError(StatusCode.SHEET_INDEX_OUT_OF_RANGE)), IndexOutOfRangeError)
        self.assertIs(type(FormulonError(StatusCode.INDEX_OUT_OF_RANGE)), IndexOutOfRangeError)
        self.assertIs(type(FormulonError(StatusCode.BINDING_INVALID_HANDLE)), formulon.WorkbookClosedError)

    def test_file_not_found_codes_use_the_io_class_not_the_not_found_class(self) -> None:
        error = FormulonError(StatusCode.IO_FILE_NOT_FOUND)
        self.assertIs(type(error), formulon.IoError)
        self.assertNotIsInstance(error, NotFoundError)

    def test_every_not_found_suffix_maps_to_not_found_error(self) -> None:
        for member in StatusCode:
            expected_not_found = member.name == "NOT_FOUND" or (
                member.name.endswith("_NOT_FOUND") and member.name != "IO_FILE_NOT_FOUND"
            )
            with self.subTest(member=member.name):
                self.assertEqual(isinstance(FormulonError(member), NotFoundError), expected_not_found)

    def test_unlisted_status_keeps_the_raw_number(self) -> None:
        err = FormulonError(123456)
        self.assertEqual(err.code, 123456)
        self.assertNotIsInstance(err.code, StatusCode)
        self.assertEqual(str(err), "[123456]")

    def test_engine_failures_arrive_as_the_family(self) -> None:
        with self.assertRaises(FormulonError) as ctx:
            Workbook.load(b"not a real xlsx archive")
        self.assertNotEqual(ctx.exception.code, StatusCode.OK)
        with Workbook.create_default() as wb:
            with self.assertRaises(IndexOutOfRangeError) as missing:
                wb.sheet(len(wb)).rename("x")
            self.assertEqual(missing.exception.code, StatusCode.SHEET_INDEX_OUT_OF_RANGE)
            with self.assertRaises(FormulonError) as invalid:
                wb.sheet(0).rename("")
            self.assertIs(type(invalid.exception), FormulonError)
            self.assertEqual(invalid.exception.code, StatusCode.INVALID_SHEET_NAME)
