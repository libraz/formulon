"""Tests for the `ARG_ROLES` tables.

Run: `python3 -m unittest tools/codegen/gen_facade_tables_roles_test.py`
"""

from __future__ import annotations

import sys
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import facade_scope  # noqa: E402
import gen_facade_tables as gen  # noqa: E402

DTS = """
export interface Workbook {
  removeSheet(sheetIndex: number): void;
  recalcParallel(threadCount: number): void;
  copySheet(source: number, insertBefore?: number | null): void;
}
export interface Sheet {
  insertRows(row: number, count: number, hidden: boolean): void;
  renameColumn(table: number, column: number, name: string): string;
}
"""

HEADER = """
FM_API fm_status_t fm_workbook_remove_sheet(fm_workbook_t* wb, size_t sheet_index);
FM_API fm_status_t fm_workbook_recalc_parallel(fm_workbook_t* wb, int32_t thread_count);
FM_API fm_status_t fm_workbook_copy_sheet(fm_workbook_t* wb, size_t source_index, size_t insert_before);
FM_API fm_status_t fm_sheet_insert_rows(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t count,
                                        int32_t hidden);
FM_API fm_status_t fm_sheet_rename_column(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                          size_t column_index, const char* name, const char** out_name);
"""


def _op(scope: str, c: str, js: str, py: str, **extra: object) -> dict:
    return {"kind": "op", "scope": scope, "c": c, "js": js, "py": py, **extra}


ENTRIES = {
    "sheet.remove": _op("workbook", "fm_workbook_remove_sheet", "removeSheet", "remove_sheet"),
    "recalc.parallel": _op("workbook", "fm_workbook_recalc_parallel", "recalcParallel", "recalc_parallel"),
    "sheet.copy": _op("workbook", "fm_workbook_copy_sheet", "copySheet", "copy_sheet", facade="handwritten"),
    "rows.insert": _op("sheet", "fm_sheet_insert_rows", "insertRows", "insert_rows"),
    "column.rename": _op("sheet", "fm_sheet_rename_column", "renameColumn", "rename_column"),
}


def _ledger(entries: dict) -> facade_scope.Ledger:
    return facade_scope.Ledger({"schema": 1, "collections": {}, "entries": entries})


class ArgRoleTest(unittest.TestCase):
    def test_the_c_parameter_name_decides_the_role(self) -> None:
        self.assertEqual(gen.arg_role("sheet_index"), "sheet")
        for name in ("index", "table_index", "plot_index"):
            self.assertEqual(gen.arg_role(name), "position", name)
        for name in ("row", "col", "first", "last", "first_row", "last_col"):
            self.assertEqual(gen.arg_role(name), "coordinate", name)
        for name in ("count", "depth", "thread_count", "cache_id"):
            self.assertEqual(gen.arg_role(name), "quantity", name)


class ArgRolesTest(unittest.TestCase):
    def setUp(self) -> None:
        self.js = gen.arg_roles(_ledger(ENTRIES), DTS, HEADER)

    def test_js_positions_follow_the_longest_overload(self) -> None:
        self.assertEqual(self.js["Workbook.removeSheet"], [[0, "sheet"]])
        self.assertEqual(self.js["Workbook.recalcParallel"], [[0, "quantity"]])
        self.assertEqual(self.js["Sheet.insertRows"], [[0, "coordinate"], [1, "quantity"]])

    def test_an_index_suffix_is_optional_in_the_js_name(self) -> None:
        self.assertEqual(self.js["Sheet.renameColumn"], [[0, "position"], [1, "position"]])

    def test_boolean_and_handwritten_parameters_are_left_out(self) -> None:
        self.assertEqual(len(self.js["Sheet.insertRows"]), 2)
        self.assertNotIn("Workbook.copySheet", self.js)


if __name__ == "__main__":
    unittest.main()
