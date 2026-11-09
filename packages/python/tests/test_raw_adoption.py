"""Raw-layer tests: struct_size packing, parse-error position, colour spec."""

from __future__ import annotations

import re
import struct
import unittest
from pathlib import Path

from formulon import FormulonError, InvalidArgumentError
from formulon import _structs as S
from formulon._c import LIB, FormulaSyntaxError, ParseErrorCode

REPO_ROOT = Path(__file__).resolve().parents[3]
STRUCT_V1_SIZES = REPO_ROOT / "src" / "c_api" / "parts" / "struct_v1_sizes.h"


class RecordingMemory:
    """Stands in for the WASM instance in packer tests."""

    def __init__(self, size: int = 1024) -> None:
        self.memory = bytearray(size)

    def write_bytes(self, ptr: int, data: bytearray) -> None:
        self.memory[ptr : ptr + len(data)] = data

    def read_bytes(self, ptr: int, length: int) -> bytes:
        return bytes(self.memory[ptr : ptr + length])


def _layouts() -> dict[str, S.Struct]:
    return {layout.name: layout for layout in vars(S).values() if isinstance(layout, S.Struct)}


class StructSizePackingTest(unittest.TestCase):
    def test_record_names_match_the_header_table(self) -> None:
        text = STRUCT_V1_SIZES.read_text(encoding="utf-8")
        header_records = set(re.findall(r"FORMULON_STRUCT_V1\((fm_\w+),", text))
        self.assertEqual(set(S.RECORD_NAMES), header_records)

    def test_every_record_layout_leads_with_struct_size(self) -> None:
        for name, layout in _layouts().items():
            with self.subTest(name):
                if name in S.RECORD_NAMES:
                    self.assertTrue(layout.is_record)
                    self.assertEqual(layout.fields[0], ("struct_size", S.U32))
                    self.assertEqual(layout.offsets["struct_size"], ("u32", 0))
                else:
                    self.assertFalse(layout.is_record)
                    self.assertNotIn("struct_size", layout.offsets)

    def test_pack_writes_struct_size_without_the_caller_setting_it(self) -> None:
        lib = RecordingMemory()
        S.COMMENT.pack(lib, 64, {"row": 3, "col": 4, "author": 0, "text": 0})

        self.assertEqual(struct.unpack_from("<I", lib.memory, 64)[0], S.COMMENT.size)
        self.assertEqual(S.COMMENT.size, 20)
        decoded = S.COMMENT.unpack(lib, 64)
        self.assertEqual((decoded["row"], decoded["col"]), (3, 4))
        self.assertNotIn("struct_size", decoded)

    def test_pack_with_no_values_still_writes_struct_size(self) -> None:
        lib = RecordingMemory()
        S.DATA_VALIDATION.pack(lib, 0, {})
        self.assertEqual(struct.unpack_from("<I", lib.memory, 0)[0], S.DATA_VALIDATION.size)

    def test_an_explicit_struct_size_is_kept_for_negative_cases(self) -> None:
        raw = S.COMMENT.blank()
        S.COMMENT.pack_into(raw, 0, {"struct_size": 4})
        self.assertEqual(struct.unpack_from("<I", raw, 0)[0], 4)

    def test_zero_struct_clears_the_body_and_keeps_struct_size(self) -> None:
        lib = RecordingMemory()
        lib.memory[: S.COMMENT.size] = b"\xff" * S.COMMENT.size
        S.zero_struct(lib, S.COMMENT, 0)
        self.assertEqual(struct.unpack_from("<I", lib.memory, 0)[0], S.COMMENT.size)
        self.assertEqual(bytes(lib.memory[4 : S.COMMENT.size]), b"\x00" * (S.COMMENT.size - 4))

    def test_a_value_type_has_no_struct_size(self) -> None:
        raw = S.COLOR_SPEC.blank()
        self.assertEqual(len(raw), 24)
        self.assertEqual(bytes(raw), b"\x00" * 24)

    def test_a_record_the_engine_checks_accepts_the_packed_block(self) -> None:
        # fm_get_last_parse_error rejects a record whose struct_size is below
        # its layout, so a successful read proves the packer's value is right.
        status_ptr = LIB.alloc(4 + S.PARSE_ERROR.size)
        try:
            LIB.write_bytes(status_ptr, bytes(4))
            S.PARSE_ERROR.pack(LIB, status_ptr + 4, {})
            status = LIB.fm_get_last_parse_error(status_ptr, status_ptr + 4)
            self.assertEqual(status, 0)  # not FM_STATUS_INVALID_ARGUMENT
        finally:
            LIB.free(status_ptr)


def _cstr(text: str) -> int:
    data = text.encode("utf-8") + b"\x00"
    ptr = LIB.alloc(len(data))
    LIB.write_bytes(ptr, data)
    return ptr


class _RawWorkbook:
    def __enter__(self) -> int:
        out = LIB.alloc(4)
        try:
            status = LIB.fm_workbook_create_default(out)
            if status != 0:
                raise FormulonError(status)
            self.handle = LIB.read_u32(out)
        finally:
            LIB.free(out)
        return self.handle

    def __exit__(self, *exc: object) -> None:
        LIB.fm_workbook_destroy(self.handle)


class ParseErrorPositionTest(unittest.TestCase):
    def test_syntax_failure_reports_code_and_position(self) -> None:
        with _RawWorkbook() as wb:
            formula = _cstr("=1+")
            try:
                status = LIB.fm_sheet_set_formula(wb, 0, 0, 0, formula)
            finally:
                LIB.free(formula)

            self.assertTrue(1000 <= status <= 1014, status)
            error = FormulonError._from_engine(status)

        self.assertIsInstance(error, FormulaSyntaxError)
        self.assertIsInstance(error, InvalidArgumentError)
        self.assertIsNotNone(error.parse_error)
        self.assertIsInstance(error.parse_error.code, ParseErrorCode)
        self.assertLessEqual(error.parse_error.start, error.parse_error.end)
        self.assertGreaterEqual(error.parse_error.line, 1)
        self.assertGreaterEqual(error.parse_error.column, 1)
        for retired in ("start", "end", "line", "column", "parse_error_code"):
            self.assertFalse(hasattr(error, retired), retired)

    def test_fm_get_last_parse_error_reads_the_record_directly(self) -> None:
        with _RawWorkbook() as wb:
            formula = _cstr("=SUM(")
            out = LIB.alloc(4 + S.PARSE_ERROR.size)
            try:
                status = LIB.fm_sheet_set_formula(wb, 0, 0, 0, formula)
                self.assertTrue(1000 <= status <= 1014, status)
                S.PARSE_ERROR.pack(LIB, out + 4, {})
                self.assertEqual(LIB.fm_get_last_parse_error(out, out + 4), 0)
                self.assertEqual(LIB.read_u32(out), 1)
                decoded = S.PARSE_ERROR.unpack(LIB, out + 4)
                # The record survives the read: a second read returns the same.
                S.PARSE_ERROR.pack(LIB, out + 4, {})
                self.assertEqual(LIB.fm_get_last_parse_error(out, out + 4), 0)
                self.assertEqual(LIB.read_u32(out), 1)
                self.assertEqual(S.PARSE_ERROR.unpack(LIB, out + 4), decoded)
            finally:
                LIB.free(out)
                LIB.free(formula)

        self.assertLessEqual(decoded["start"], decoded["end"])
        self.assertGreaterEqual(decoded["line"], 1)

    def test_a_later_call_clears_the_record(self) -> None:
        with _RawWorkbook() as wb:
            formula = _cstr("=1+")
            out = LIB.alloc(4 + S.PARSE_ERROR.size)
            try:
                LIB.fm_sheet_set_formula(wb, 0, 0, 0, formula)
                LIB.fm_workbook_recalc(wb)
                LIB.write_bytes(out, bytes(4))
                S.PARSE_ERROR.pack(LIB, out + 4, {})
                self.assertEqual(LIB.fm_get_last_parse_error(out, out + 4), 0)
                self.assertEqual(LIB.read_u32(out), 0)
            finally:
                LIB.free(out)
                LIB.free(formula)

    def test_a_non_syntax_failure_carries_no_parse_error(self) -> None:
        with _RawWorkbook() as wb:
            status = LIB.fm_sheet_set_number(wb, 99, 0, 0, 1.0)
            self.assertNotEqual(status, 0)
            error = FormulonError._from_engine(status)

        self.assertNotIsInstance(error, FormulaSyntaxError)
        self.assertIsNone(error.parse_error)


class ColorSpecRoundTripTest(unittest.TestCase):
    CASES = (
        {"kind": 0, "argb": 0, "theme": 0, "indexed": 0, "tint": 0.0},
        {"kind": 1, "argb": 0xFF336699, "theme": 0, "indexed": 0, "tint": 0.0},
        {"kind": 2, "argb": 0, "theme": 5, "indexed": 0, "tint": -0.25},
        {"kind": 3, "argb": 0, "theme": 0, "indexed": 12, "tint": 0.0},
        {"kind": 4, "argb": 0, "theme": 0, "indexed": 0, "tint": 0.0},
    )

    def test_layout_is_the_value_type_size(self) -> None:
        self.assertEqual(S.COLOR_SPEC.size, 24)
        self.assertEqual([name for name, _ in S.COLOR_SPEC.fields], ["tint", "argb", "theme", "indexed", "kind"])

    def test_tab_color_round_trips_every_kind(self) -> None:
        with _RawWorkbook() as wb:
            spec = LIB.alloc(S.COLOR_SPEC.size)
            try:
                for case in self.CASES:
                    with self.subTest(kind=case["kind"]):
                        S.COLOR_SPEC.pack(LIB, spec, case)
                        self.assertEqual(LIB.fm_sheet_set_tab_color(wb, 0, spec), 0)

                        S.zero_struct(LIB, S.COLOR_SPEC, spec)
                        self.assertEqual(LIB.fm_sheet_get_tab_color(wb, 0, spec), 0)
                        self.assertEqual(S.COLOR_SPEC.unpack(LIB, spec), case)
            finally:
                LIB.free(spec)


if __name__ == "__main__":
    unittest.main()
