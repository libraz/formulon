#!/usr/bin/env python3
"""Fixture tests for the engine-messages check.

Run from the repo root: `python3 -m tools.dev.binding_drift.engine_messages_test`.
"""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from tools.dev.binding_drift.engine_messages import check_engine_messages, message_problem

HEADER = """
typedef struct fm_table_spec_t { const char* name; uint32_t first_row; } fm_table_spec_t;
FM_API fm_status_t fm_sheet_insert_cells(fm_workbook_t* wb, size_t sheet_index, const char* range, int32_t shift);
"""
FIELDS = {"name", "first_row", "wb", "sheet_index", "range", "shift"}

SOURCE = """
Error a() { return make_error(FormulonErrorCode::kInvalidArgument, "invalid name", "name=x"); }
Error b() { return make_error(FormulonErrorCode::kInvalidArgument, "name is not a valid table name"); }
Error c(const char* op) { return make_error(FormulonErrorCode::kPreconditionFailed, std::string(op) + ": busy"); }
Error d() { return make_error(FormulonErrorCode::kNotFound, "no such defined name"); }
Error e() { return make_error(FormulonErrorCode::kIoSheetCorrupt, "cell parser: t='s' without <v> child"); }
Error f() { return invalid_table("shift must be down or right"); }
Error g() { return make_error(FormulonErrorCode::kPreconditionFailed, "Normal style cannot be removed"); }
Error h() { return make_error(FormulonErrorCode::kInvalidArgument, "insert_cells: range is outside the grid"); }
"""


class EngineMessagesTest(unittest.TestCase):
    def test_message_problem(self) -> None:
        self.assertIsNone(message_problem("name is not a valid defined name", FIELDS))
        self.assertIsNone(message_problem("range is outside the grid", FIELDS))
        self.assertIsNone(message_problem("Normal style cannot be removed", FIELDS))
        self.assertIn("invalid", message_problem("invalid name", FIELDS) or "")
        self.assertEqual(message_problem("insert_cells: shift", FIELDS), "leads with an operation prefix")

    def test_scan_flags_prose_and_op_prefixes_and_skips_file_diagnostics(self) -> None:
        with TemporaryDirectory() as tmp:
            src = Path(tmp) / "src"
            (src / "wasm").mkdir(parents=True)
            (src / "x.cpp").write_text(SOURCE, encoding="utf-8")
            (src / "wasm" / "glue.cpp").write_text(SOURCE, encoding="utf-8")
            problems = check_engine_messages(src, HEADER)
        lines = sorted(int(p.split(":")[2]) for p in problems)
        # a (prose), c (op prefix), d (prose), h (op prefix); the binding copy is not scanned.
        self.assertEqual(lines, [2, 4, 5, 9])


if __name__ == "__main__":
    unittest.main()
