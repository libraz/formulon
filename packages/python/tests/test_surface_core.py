"""Core Workbook surface tests."""

from __future__ import annotations

import inspect
import subprocess
import unittest
from pathlib import Path

from formulon import (
    CalcMode,
    CivilTime,
    FormulonError,
    ValueKind,
    Workbook,
    _c,
)
from formulon import _structs as S


class StructLayoutTests(unittest.TestCase):
    """Guards against silent field-reorder breakage in _structs.py."""

    # One pinned wasm32 size per `Struct` in `_structs.py`. The values are
    # deliberately literal rather than derived -- a derived table would agree
    # with any layout and assert nothing. What binds them to the C header is
    # `python-struct-layouts` below; this table's job is to make a silent
    # *size* change loud, and `test_every_struct_has_a_pinned_size` is what
    # stops a new struct from opting out by simply not being listed.
    EXPECTED_SIZES = {
        "MERGE_RANGE": 16,
        "HYPERLINK": 32,
        "COMMENT": 16,
        "DATA_VALIDATION": 52,
        "SHEET_PROTECTION": 88,
        "VIEWPORT": 20,
        "CELL_NODE": 12,
        "PHONETIC_RUN": 12,
        "CFVO": 12,
        "CF_CELL_RANGE": 16,
        "CF_COLOR": 4,
        "CF_MATCH": 72,
        "CF_RULE": 240,
        "PIVOT_CELL": 40,
        "PIVOT_FIELD_SPEC": 20,
        "PIVOT_DATA_FIELD_SPEC": 28,
        "PIVOT_FILTER_SPEC": 64,
        "READ_DIAGNOSTICS": 20,
        "SAVE_DIAGNOSTICS": 20,
        "SPILL_INFO": 20,
        "FUNCTION_METADATA": 24,
        "CIVIL_TIME": 24,
        "SHEET_VIEW": 44,
        "COLUMN_LAYOUT": 40,
        "ROW_LAYOUT": 32,
        "CELL_XF": 88,
        "COLOR_SPEC": 24,
        "FONT_RECORD": 88,
        "FILL_RECORD": 64,
        "BORDER_SIDE": 32,
        "BORDER_RECORD": 168,
        "DXF_RECORD": 368,
        # Fifteen pointer-or-`size_t` fields; both are four bytes on
        # wasm32, so the whole record is 15 x 4.
        "STYLES_BATCH": 60,
        "CELL_STYLE_RECORD": 24,
        "EXTERNAL_LINK_RECORD": 24,
        "PAGE_BREAK": 16,
        "PAGE_SETUP": 48,
        # Six `double`s, each preceded by an `int32` flag: the flag's four
        # bytes of tail padding are what makes this 96 rather than 72.
        "PAGE_MARGINS": 96,
        "PRINT_OPTIONS": 32,
        "HEADER_FOOTER": 56,
    }

    def test_struct_sizes(self) -> None:
        for name, expected in self.EXPECTED_SIZES.items():
            layout = getattr(S, name)
            self.assertEqual(layout.size, expected, f"{name} size drifted to {layout.size}")

    def test_every_struct_has_a_pinned_size(self) -> None:
        """No `Struct` may be absent from ``EXPECTED_SIZES``.

        Without this the table degrades quietly: a struct added to
        ``_structs.py`` and never listed here is not size-pinned, and nothing
        says so. Both directions are checked so a removed struct cannot leave
        a stale entry behind either.
        """
        declared = {name for name, value in vars(S).items() if isinstance(value, S.Struct)}
        pinned = set(self.EXPECTED_SIZES)
        self.assertEqual(
            declared,
            pinned,
            f"not size-pinned: {sorted(declared - pinned)}; stale entries: {sorted(pinned - declared)}",
        )

    def test_pivot_cell_value_offset(self) -> None:
        self.assertEqual(S.PIVOT_CELL_VALUE_OFFSET, 8)

    # Records the wrapper marshals without a `Struct` entry, so they are
    # invisible to `EXPECTED_SIZES` above. `fm_value_t` is passed around as
    # the bare `fm_value_t_size` and as `VALUE_BLOB` when it sits inline in a
    # larger record; `fm_print_range_t` is the literal 16 in
    # `Workbook.paginate`. `python-inline-structs` binds these to the C
    # header; pinning them here makes a size change loud on its own.
    def test_hand_marshalled_record_sizes(self) -> None:
        self.assertEqual(_c.fm_value_t_size, 16)
        self.assertEqual(S.VALUE_BLOB, ("blob16", 16, 8))
        source = inspect.getsource(Workbook.paginate)
        self.assertIn('struct.unpack("<IIII", LIB.read_bytes(range_ptr, 16))', source)

    def test_layouts_match_c_header(self) -> None:
        """Every binding-drift check that reads the Python package.

        Run as a group: the struct-layout check covers the `Struct` table,
        the call-signature check covers what the wrapper passes to each
        `fm_*` entry point, and the inline-struct check covers the records
        decoded with a bare `struct.unpack`. A failure in any of them is a
        Python-side ABI mismatch and belongs in this suite's result.
        """
        root = Path(__file__).resolve().parents[3]
        for check in ("python-struct-layouts", "python-call-signatures", "python-inline-structs"):
            with self.subTest(check=check):
                result = subprocess.run(
                    ["python3", str(root / "tools" / "dev" / "check_binding_drift.py"), check],
                    cwd=root,
                    text=True,
                    capture_output=True,
                    check=False,
                )
                self.assertEqual(result.returncode, 0, result.stdout + result.stderr)


class AdHocArrayEvalTests(unittest.TestCase):
    def test_sequence_matrix_preserves_all_cells(self) -> None:
        with Workbook.create_default() as wb:
            grid = wb.evaluate_formula_array(0, 0, 0, "=SEQUENCE(2,3)")
            self.assertEqual(len(grid), 2)
            self.assertEqual(len(grid[0]), 3)
            # Row-major 1..6.
            self.assertEqual(grid[0][0].to_python(), 1.0)
            self.assertEqual(grid[1][2].to_python(), 6.0)

    def test_scalar_reported_as_one_by_one(self) -> None:
        with Workbook.create_default() as wb:
            grid = wb.evaluate_formula_array(0, 0, 0, "=1+2")
            self.assertEqual(len(grid), 1)
            self.assertEqual(len(grid[0]), 1)
            self.assertEqual(grid[0][0].kind, ValueKind.NUMBER)
            self.assertEqual(grid[0][0].to_python(), 3.0)


class DefinedNameTests(unittest.TestCase):
    def test_defined_name_set_get_remove(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_defined_name("MyRef", "Sheet1!$A$1")
            names = {dn.name: dn.formula for dn in wb.iter_defined_names()}
            self.assertIn("MyRef", names)
            self.assertEqual(names["MyRef"], "Sheet1!$A$1")

            # Updating in place replaces the formula text.
            wb.set_defined_name("MyRef", "Sheet1!$B$2")
            names = {dn.name: dn.formula for dn in wb.iter_defined_names()}
            self.assertEqual(names["MyRef"], "Sheet1!$B$2")

            # An empty formula removes the entry.
            wb.set_defined_name("MyRef", "")
            self.assertNotIn("MyRef", {dn.name for dn in wb.iter_defined_names()})

    def test_scoped_defined_name_set_get_remove(self) -> None:
        with Workbook.create_empty() as wb:
            wb.add_sheet("Sheet1")
            wb.add_sheet("Sheet2")
            wb.set_defined_name("Rate", "=1")
            wb.set_defined_name_scoped("Rate", "=2", 1)
            names = {(dn.name, dn.local_sheet_id): dn.formula for dn in wb.iter_defined_names()}
            self.assertEqual(names[("Rate", -1)], "=1")
            self.assertEqual(names[("Rate", 1)], "=2")

            wb.set_defined_name_scoped("Rate", "", 1)
            self.assertNotIn(
                ("Rate", 1),
                {(dn.name, dn.local_sheet_id) for dn in wb.iter_defined_names()},
            )


class CalcPolicyTests(unittest.TestCase):
    def test_calc_mode_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.calc_mode(), CalcMode.AUTO)
            wb.set_calc_mode(CalcMode.MANUAL)
            self.assertEqual(wb.calc_mode(), CalcMode.MANUAL)

    def test_pinned_now_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            # Unpinned by default: the workbook follows the host clock.
            self.assertIsNone(wb.pinned_now())
            wb.set_pinned_now(2026, 4, 23, 15, 30, 45)
            self.assertEqual(wb.pinned_now(), CivilTime(2026, 4, 23, 15, 30, 45))
            # 2026-04-23 is serial 46135 under the 1900 date system.
            grid = wb.evaluate_formula_array(0, 0, 0, "=TODAY()")
            self.assertEqual(grid[0][0].number, 46135.0)
            wb.clear_pinned_now()
            self.assertIsNone(wb.pinned_now())

    def test_pinned_now_rejects_a_non_calendar_instant(self) -> None:
        with Workbook.create_default() as wb:
            # The pin is a calendar instant, not a normalising constructor.
            with self.assertRaises(FormulonError):
                wb.set_pinned_now(2026, 13, 1)
            with self.assertRaises(FormulonError):
                wb.set_pinned_now(2025, 2, 29)
            self.assertIsNone(wb.pinned_now())

    def test_excel_profile_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            self.assertTrue(wb.excel_profile_id())
            wb.set_excel_profile_id("mac-365-ja_JP")
            self.assertEqual(wb.excel_profile_id(), "mac-365-ja_JP")


class PartialRecalcTests(unittest.TestCase):
    def test_partial_recalc_recomputes_chain(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 10.0)  # A1
            wb.set_formula(0, 0, 1, "=A1*2")  # B1
            recomputed = wb.partial_recalc(0, 0, 0, 1, 1)
            self.assertGreater(recomputed, 0)
            self.assertEqual(wb.get_value(0, 0, 1).to_python(), 20.0)


class TraceTests(unittest.TestCase):
    def test_precedents_and_dependents(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 1.0)  # A1
            wb.set_formula(0, 0, 1, "=A1")  # B1
            wb.recalc()
            prec = wb.precedents(0, 0, 1, 1)
            self.assertIn((0, 0, 0), [(p.sheet, p.row, p.col) for p in prec])
            deps = wb.dependents(0, 0, 0, 1)
            self.assertIn((0, 0, 1), [(d.sheet, d.row, d.col) for d in deps])


class SpillTests(unittest.TestCase):
    def test_sequence_spills(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_formula(0, 0, 0, "=SEQUENCE(3)")
            wb.recalc()
            info = wb.spill_info(0, 0, 0)
            self.assertTrue(info.engaged)
            self.assertEqual(info.rows, 3)
            self.assertEqual((info.anchor_row, info.anchor_col), (0, 0))


class ExternalLinkTests(unittest.TestCase):
    def test_fresh_workbook_has_no_external_links(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.external_link_count(), 0)
            self.assertEqual(wb.get_external_links(), [])


if __name__ == "__main__":
    unittest.main()
