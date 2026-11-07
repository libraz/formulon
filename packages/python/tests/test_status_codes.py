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
from formulon._c import LIB

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
        self.assertEqual(StatusCode.SHEET_INDEX_OUT_OF_RANGE, 5050)
        self.assertEqual(StatusCode.PARSER_INVALID_R1C1, 1011)
        self.assertEqual(StatusCode.OK, 0)

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
        err = FormulonError(StatusCode.PARSER_INVALID_R1C1)
        self.assertIs(err.code, StatusCode.PARSER_INVALID_R1C1)
        self.assertEqual(err.code.name, "PARSER_INVALID_R1C1")
        self.assertFalse(hasattr(err, "status"))
        self.assertFalse(hasattr(err, "status_name"))
        self.assertEqual(str(err), err.message)

    def test_dedicated_subclass_per_code(self) -> None:
        self.assertIs(type(FormulonError(StatusCode.INVALID_ARGUMENT)), InvalidArgumentError)
        for member in (
            StatusCode.NOT_FOUND,
            StatusCode.PIVOT_FIELD_NOT_FOUND,
            StatusCode.EVAL_LAMBDA_NOT_FOUND,
        ):
            with self.subTest(member=member.name):
                self.assertIs(type(FormulonError(member)), NotFoundError)
        self.assertIs(type(FormulonError(StatusCode.SHEET_INDEX_OUT_OF_RANGE)), IndexOutOfRangeError)
        self.assertIs(type(FormulonError(StatusCode.INDEX_OUT_OF_RANGE)), IndexOutOfRangeError)
        self.assertIs(type(FormulonError(StatusCode.BINDING_INVALID_HANDLE)), formulon.WorkbookClosedError)

    def test_file_not_found_codes_stay_on_the_base(self) -> None:
        for member in (StatusCode.IO_FILE_NOT_FOUND, StatusCode.CLI_FILE_NOT_FOUND):
            with self.subTest(member=member.name):
                self.assertIs(type(FormulonError(member)), FormulonError)

    def test_every_not_found_suffix_maps_to_not_found_error(self) -> None:
        for member in StatusCode:
            expected_not_found = member.name == "NOT_FOUND" or (
                member.name.endswith("_NOT_FOUND") and member.name not in ("IO_FILE_NOT_FOUND", "CLI_FILE_NOT_FOUND")
            )
            with self.subTest(member=member.name):
                self.assertEqual(isinstance(FormulonError(member), NotFoundError), expected_not_found)

    def test_unlisted_status_keeps_the_raw_number(self) -> None:
        err = FormulonError(123456)
        self.assertEqual(err.code, 123456)
        self.assertNotIsInstance(err.code, StatusCode)
        self.assertEqual(str(err), err.message)

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
