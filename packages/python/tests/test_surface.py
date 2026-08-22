"""Surface tests for the expanded Formulon Python Workbook API.

These exercise a representative sample of every method group added on top
of the load/edit/recalc/save core: pivots, conditional formats, styles,
matrix edits, merges/comments/hyperlinks/validations, defined names,
partial recalc, precedents/dependents, spill, function catalog, sheet
view/protection, and external links.

Run via ``make python-test`` (or ``python -m unittest`` with
``PYTHONPATH=packages/python``). Uses stdlib ``unittest`` -- no pytest.
"""

from __future__ import annotations

import inspect
import json
import struct
import subprocess
import unittest
from pathlib import Path
from unittest import mock

import formulon
from formulon import (
    CalcMode,
    CellXf,
    CfColor,
    CfValueObject,
    CivilTime,
    ColorScale,
    ColorSpec,
    ConditionalFormatInput,
    DataBar,
    DataValidationInput,
    DifferentialFormat,
    FillRecord,
    FontRecord,
    FormulonError,
    IconSet,
    MergeRange,
    PivotAggregation,
    PivotAxis,
    PivotDataFieldSpec,
    PivotFieldSpec,
    PivotReportLayout,
    PivotWorksheetSource,
    ReadDiagnostics,
    SaveDiagnostics,
    SheetProtection,
    SheetVisibility,
    ValueKind,
    Workbook,
    WorkbookFormat,
    _c,
)
from formulon import _structs as S


def _append_empty_zip_entry(data: bytes, name: str) -> bytes:
    """Add one deterministic, empty stored entry without recompressing XLSB."""
    eocd = data.rfind(b"PK\x05\x06")
    if eocd < 0:
        raise AssertionError("missing ZIP end record")
    _disk_count, count, central_size, central_offset = struct.unpack_from("<HHII", data, eocd + 8)
    encoded_name = name.encode("utf-8")
    local = struct.pack("<IHHHHHIIIHH", 0x04034B50, 20, 0, 0, 0, 0, 0, 0, 0, len(encoded_name), 0) + encoded_name
    central = (
        struct.pack(
            "<IHHHHHHIIIHHHHHII",
            0x02014B50,
            20,
            20,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            len(encoded_name),
            0,
            0,
            0,
            0,
            0,
            central_offset,
        )
        + encoded_name
    )
    new_eocd = struct.pack(
        "<IHHHHIIH",
        0x06054B50,
        0,
        0,
        count + 1,
        count + 1,
        central_size + len(central),
        central_offset + len(local),
        0,
    )
    return data[:central_offset] + local + data[central_offset : central_offset + central_size] + central + new_eocd


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
        "CF_RULE": 216,
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
        "FONT_RECORD": 80,
        "FILL_RECORD": 64,
        "BORDER_SIDE": 32,
        "BORDER_RECORD": 168,
        "DXF_RECORD": 360,
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


class SheetStructureTests(unittest.TestCase):
    def test_add_rename_remove_move_sheet(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("Second")
            wb.add_sheet("Third")
            self.assertEqual(wb.sheet_count(), 3)
            wb.rename_sheet(1, "Renamed")
            self.assertEqual(wb.sheet_name(1), "Renamed")
            wb.move_sheet(2, 0)
            self.assertEqual(wb.sheet_name(0), "Third")
            wb.remove_sheet(0)
            self.assertEqual(wb.sheet_count(), 2)


class MatrixEditTests(unittest.TestCase):
    def test_insert_row_shifts_value_down(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 11.0)  # A1
            wb.insert_rows(0, 0, 1)
            wb.recalc()
            # A1 is now blank; the value moved to A2.
            self.assertEqual(wb.get_value(0, 0, 0).kind, ValueKind.BLANK)
            self.assertEqual(wb.get_value(0, 1, 0).to_python(), 11.0)

    def test_delete_col_shifts_value_left(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 1.0)  # A1
            wb.set_number(0, 0, 1, 2.0)  # B1
            wb.delete_cols(0, 0, 1)
            wb.recalc()
            # The old B1 collapses into A1.
            self.assertEqual(wb.get_value(0, 0, 0).to_python(), 2.0)


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


class MergeCommentHyperlinkTests(unittest.TestCase):
    def test_merge_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_merge(0, MergeRange(0, 0, 1, 1))
            merges = wb.get_merges(0)
            self.assertEqual(len(merges), 1)
            self.assertEqual((merges[0].first_row, merges[0].last_col), (0, 1))
            wb.clear_merges(0)
            self.assertEqual(wb.get_merges(0), [])

    def test_comment_roundtrip_and_absent(self) -> None:
        with Workbook.create_default() as wb:
            self.assertIsNone(wb.get_comment(0, 0, 0))
            wb.set_comment(0, 0, 0, "alice", "see note")
            c = wb.get_comment(0, 0, 0)
            self.assertIsNotNone(c)
            self.assertEqual(c.author, "alice")
            self.assertEqual(c.text, "see note")

    def test_comment_invalid_sheet_raises_instead_of_looking_absent(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.FormulonError) as ctx:
                wb.get_comment(99, 0, 0)
            self.assertEqual(ctx.exception.status, 2)  # kInvalidArgument

    def test_hyperlink_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_hyperlink(0, 2, 3, "https://example.com", "Example", "tip")
            links = wb.get_hyperlinks(0)
            self.assertEqual(len(links), 1)
            self.assertEqual((links[0].row, links[0].col), (2, 3))
            self.assertEqual((links[0].last_row, links[0].last_col), (2, 3))
            self.assertEqual(links[0].target, "https://example.com")
            self.assertEqual(links[0].display, "Example")

    def test_hyperlink_range_roundtrip_at_nonzero_coordinate(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_hyperlink_range(0, 4, 6, 7, 9, "https://example.com", "Range", "tip")
            links = wb.get_hyperlinks(0)
            self.assertEqual(len(links), 1)
            self.assertEqual((links[0].row, links[0].col), (4, 6))
            self.assertEqual((links[0].last_row, links[0].last_col), (7, 9))


class ValidationTests(unittest.TestCase):
    def test_list_validation_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_validation(
                0,
                DataValidationInput(
                    type=3,  # list
                    ranges=[MergeRange(0, 0, 4, 0)],
                    allow_blank=True,
                    formula1='"a,b,c"',
                ),
            )
            self.assertEqual(wb.validation_count(0), 1)
            dv = wb.get_validation_at(0, 0)
            self.assertEqual(dv.type, 3)
            self.assertTrue(dv.allow_blank)
            self.assertEqual(dv.formula1, '"a,b,c"')
            self.assertEqual(len(dv.ranges), 1)
            wb.clear_validations(0)
            self.assertEqual(wb.validation_count(0), 0)


class SheetViewProtectionTests(unittest.TestCase):
    def test_view_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_sheet_zoom(0, 150)
            wb.set_sheet_freeze(0, 1, 2)
            view = wb.get_sheet_view(0)
            self.assertEqual(view.zoom_scale, 150)
            self.assertEqual((view.freeze_rows, view.freeze_cols), (1, 2))

    def test_view_display_defaults(self) -> None:
        with Workbook.create_default() as wb:
            view = wb.get_sheet_view(0)
            self.assertTrue(view.show_grid_lines)
            self.assertTrue(view.show_row_col_headers)
            self.assertTrue(view.show_zeros)
            self.assertFalse(view.right_to_left)
            self.assertFalse(view.tab_selected)
            self.assertEqual(view.view_mode, "")

    def test_view_display_setters_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_sheet_show_grid_lines(0, False)
            wb.set_sheet_show_row_col_headers(0, False)
            wb.set_sheet_show_zeros(0, False)
            wb.set_sheet_right_to_left(0, True)
            wb.set_sheet_tab_selected(0, True)
            wb.set_sheet_view_mode(0, "pageBreakPreview")
            view = wb.get_sheet_view(0)
            self.assertFalse(view.show_grid_lines)
            self.assertFalse(view.show_row_col_headers)
            self.assertFalse(view.show_zeros)
            self.assertTrue(view.right_to_left)
            self.assertTrue(view.tab_selected)
            self.assertEqual(view.view_mode, "pageBreakPreview")
            # An empty mode is a meaningful value (the OOXML-default
            # "normal" view), not a no-op -- it must round-trip too.
            wb.set_sheet_view_mode(0, "")
            self.assertEqual(wb.get_sheet_view(0).view_mode, "")

    def test_visibility_states_very_hidden(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.get_sheet_view(0).visibility, SheetVisibility.VISIBLE)
            self.assertFalse(wb.get_sheet_view(0).tab_hidden)

            wb.set_sheet_visibility(0, SheetVisibility.VERY_HIDDEN)
            view = wb.get_sheet_view(0)
            self.assertEqual(view.visibility, SheetVisibility.VERY_HIDDEN)
            # The two-state view stays consistent: a very-hidden sheet is
            # hidden to code that reads only the bool, never visible.
            self.assertTrue(view.tab_hidden)

            # "Hidden" says nothing a very-hidden sheet does not already
            # satisfy, so it must not weaken the author's stronger choice.
            wb.set_sheet_tab_hidden(0, True)
            self.assertEqual(wb.get_sheet_view(0).visibility, SheetVisibility.VERY_HIDDEN)

            # Demotion is the other direction the bool cannot express.
            wb.set_sheet_visibility(0, SheetVisibility.HIDDEN)
            self.assertEqual(wb.get_sheet_view(0).visibility, SheetVisibility.HIDDEN)

            wb.set_sheet_tab_hidden(0, False)
            self.assertEqual(wb.get_sheet_view(0).visibility, SheetVisibility.VISIBLE)

    def test_visibility_rejects_unknown_state(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_sheet_visibility(0, SheetVisibility.VERY_HIDDEN)
            with self.assertRaises(FormulonError):
                wb.set_sheet_visibility(0, 3)
            # A refused call leaves the sheet as it was.
            self.assertEqual(wb.get_sheet_view(0).visibility, SheetVisibility.VERY_HIDDEN)

    def test_column_row_overrides(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_column_width(0, 0, 2, 18.5)
            cols = wb.get_sheet_columns(0)
            explicit = next(c for c in cols if abs(c.width - 18.5) < 1e-9)
            self.assertTrue(explicit.has_width)
            self.assertFalse(explicit.has_style)
            wb.set_column_width(0, 4, 4, 0.0)
            zero = next(c for c in wb.get_sheet_columns(0) if c.first == 4 and c.last == 4)
            self.assertEqual(zero.width, 0.0)
            self.assertTrue(zero.has_width)
            wb.set_row_height(0, 0, 30.0)
            rows = wb.get_sheet_row_overrides(0)
            explicit_row = next(r for r in rows if abs(r.height - 30.0) < 1e-9)
            self.assertFalse(explicit_row.has_style)
            self.assertEqual(explicit_row.style_xf, 0)

    def test_protection_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            prot = SheetProtection(enabled=True, sheet=True, format_cells=True)
            wb.set_sheet_protection(0, prot)
            got = wb.get_sheet_protection(0)
            self.assertTrue(got.enabled)
            self.assertTrue(got.sheet)
            self.assertTrue(got.format_cells)


class PrintSettingsTests(unittest.TestCase):
    def test_raw_xml_roundtrip_and_removal(self) -> None:
        with Workbook.create_default() as wb:
            # An absent element reads back as "" rather than raising.
            self.assertEqual(wb.get_page_setup_xml(0), "")
            fragment = '<pageSetup paperSize="9" orientation="portrait" scale="85"/>'
            wb.set_page_setup_xml(0, fragment)
            self.assertEqual(wb.get_page_setup_xml(0), fragment)

            setup = wb.get_page_setup(0)
            self.assertEqual(setup.paper_size, 9)
            self.assertEqual(setup.scale, 85)
            self.assertTrue(setup.scale_stated)

            # Empty removes the element and restores the defaults.
            wb.set_page_setup_xml(0, "")
            self.assertEqual(wb.get_page_setup_xml(0), "")
            cleared = wb.get_page_setup(0)
            self.assertEqual(cleared.scale, 100)
            self.assertFalse(cleared.scale_stated)

    def test_malformed_fragment_is_rejected_at_set_time(self) -> None:
        with Workbook.create_default() as wb:
            for bad in ("<pageSetup/><pageSetup/>", '<pageMargins left="1"/>', '<pageSetup orientation="p"'):
                with self.assertRaises(FormulonError, msg=bad):
                    wb.set_page_setup_xml(0, bad)
            self.assertEqual(wb.get_page_setup_xml(0), "")

    def test_fit_to_page_keeps_the_rest_of_sheet_pr(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_sheet_pr_xml(0, '<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"/></sheetPr>')
            wb.set_fit_to_page(0, True)
            xml = wb.get_sheet_pr_xml(0)
            # A VBA binding and a tab colour are not print settings, and
            # toggling fit-to-page must not be how they get lost.
            self.assertIn('codeName="Sheet1"', xml)
            self.assertIn('<tabColor rgb="FFFF0000"/>', xml)
            self.assertIn('fitToPage="true"', xml)
            self.assertTrue(wb.get_page_setup(0).fit_to_page)

    def test_typed_patch_preserves_unmodelled_attributes(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_page_setup_xml(0, '<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600" copies="3"/>')
            wb.set_page_setup(0, orientation=2)
            xml = wb.get_page_setup_xml(0)
            self.assertIn('orientation="landscape"', xml)
            self.assertIn('horizontalDpi="600"', xml)
            self.assertIn('copies="3"', xml)
            self.assertIn('paperSize="9"', xml)

    def test_scale_outside_excels_range_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            # Rejected rather than clamped: a mis-stated print scale lands
            # on paper, unlike the on-screen zoom.
            with self.assertRaises(FormulonError):
                wb.set_page_setup(0, scale=9)
            with self.assertRaises(FormulonError):
                wb.set_page_setup(0, scale=401)
            wb.set_page_setup(0, scale=400)
            self.assertEqual(wb.get_page_setup(0).scale, 400)

    def test_margins_patch_reports_presence_separately(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_page_margins(0, left=0.25, top=1.5)
            margins = wb.get_page_margins(0)
            self.assertAlmostEqual(margins.left, 0.25)
            self.assertAlmostEqual(margins.top, 1.5)
            self.assertTrue(margins.left_stated)
            # An unstated margin still reports its effective default, with
            # the flag clear so the caller can tell the two apart.
            self.assertAlmostEqual(margins.right, 0.7)
            self.assertFalse(margins.right_stated)

    def test_negative_margin_is_rejected(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.set_page_margins(0, bottom=-1.0)

    def test_header_footer_sections_take_decoded_text(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_header_footer(0, odd_header="&C\u5e33\u7968", odd_footer="&R&P / &N")
            # `&` introduces Excel's codes but lives in XML element text, so
            # it reaches the file escaped and decodes back to `&C...`.
            self.assertEqual(
                wb.get_header_footer_xml(0),
                "<headerFooter><oddHeader>&amp;C\u5e33\u7968</oddHeader>"
                "<oddFooter>&amp;R&amp;P / &amp;N</oddFooter></headerFooter>",
            )
            # Omitting a section leaves it; "" clears it.
            wb.set_header_footer(0, odd_footer="")
            self.assertEqual(
                wb.get_header_footer_xml(0),
                "<headerFooter><oddHeader>&amp;C\u5e33\u7968</oddHeader></headerFooter>",
            )

    def test_print_options_patch(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_print_options(0, grid_lines=True, horizontal_centered=True)
            self.assertEqual(
                wb.get_print_options_xml(0),
                '<printOptions gridLines="true" horizontalCentered="true"/>',
            )

    def test_print_area_and_titles(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_print_area(0, "A1:F8")
            self.assertEqual(wb.get_print_area(0), "A1:F8")
            with self.assertRaises(FormulonError):
                wb.set_print_area(0, "not-a-range")
            self.assertEqual(wb.get_print_area(0), "A1:F8")
            wb.set_print_area(0, "")
            self.assertEqual(wb.get_print_area(0), "")

            wb.set_print_titles(0, "1:2", "A:A")
            self.assertEqual(wb.get_print_titles(0), ("1:2", "A:A"))
            wb.set_print_titles(0)
            self.assertEqual(wb.get_print_titles(0), ("", ""))

    def test_manual_breaks(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_row_break(0, 40)
            wb.add_row_break(0, 10)
            wb.add_col_break(0, 5)
            rows = wb.get_row_breaks(0)
            # Kept sorted by index regardless of insertion order.
            self.assertEqual([b.id for b in rows], [10, 40])
            self.assertEqual(rows[0].max, 16383)
            self.assertTrue(rows[0].manual)
            self.assertEqual(wb.get_col_breaks(0)[0].max, 1048575)

            # Re-adding an existing index replaces rather than duplicates.
            wb.add_row_break(0, 10, manual=False)
            rows = wb.get_row_breaks(0)
            self.assertEqual(len(rows), 2)
            self.assertFalse(rows[0].manual)

            # Removing an absent break is not an error.
            wb.remove_row_break(0, 999)
            wb.clear_breaks(0)
            self.assertEqual(wb.get_row_breaks(0), [])
            self.assertEqual(wb.get_col_breaks(0), [])

    def test_settings_reach_paginate_without_a_save_cycle(self) -> None:
        with Workbook.create_default() as wb:
            for row in range(200):
                for col in range(20):
                    wb.set_number(0, row, col, 1.0)
            portrait = wb.paginate(0).page_count
            wb.set_page_setup(0, orientation=2)
            self.assertNotEqual(wb.paginate(0).page_count, portrait)

    def test_authored_report_survives_save_and_load(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_text(0, 0, 0, "\u58f2\u4e0a")
            wb.set_page_setup(0, paper_size=9, orientation=1, fit_to_page=True, fit_to_width=1, fit_to_height=0)
            wb.set_page_margins(0, left=0.5, right=0.5, top=0.8, bottom=0.8)
            wb.set_header_footer(0, odd_header="&C\u6708\u6b21\u5831\u544a")
            wb.set_print_area(0, "A1:F80")
            wb.set_print_titles(0, "1:2")
            wb.add_row_break(0, 39)
            data = wb.save()

        with Workbook.load(data) as reloaded:
            setup = reloaded.get_page_setup(0)
            self.assertEqual(setup.paper_size, 9)
            self.assertEqual(setup.orientation, 1)
            self.assertTrue(setup.fit_to_page)
            self.assertEqual(setup.fit_to_height, 0)
            self.assertAlmostEqual(reloaded.get_page_margins(0).left, 0.5)
            self.assertEqual(reloaded.get_print_area(0), "A1:F80")
            self.assertEqual(reloaded.get_print_titles(0), ("1:2", ""))
            self.assertEqual([b.id for b in reloaded.get_row_breaks(0)], [39])

    def test_set_range_xf_index_materialises_the_rectangle(self) -> None:
        with Workbook.create_default() as wb:
            # Something other than the seeded default xf, so the range
            # write is observable rather than deduping to index 0.
            xf = wb.add_cell_xf(
                CellXf(
                    font_index=0,
                    fill_index=0,
                    border_index=0,
                    num_fmt_id=0,
                    horizontal_align=0,
                    vertical_align=2,
                    wrap_text=True,
                    has_wrap_text=True,
                )
            )
            wb.set_range_xf_index(0, 0, 0, 2, 2, xf)
            for row, col in ((0, 0), (1, 1), (2, 2)):
                self.assertEqual(wb.get_cell_xf_index(0, row, col), xf)
            self.assertEqual(wb.get_cell_xf_index(0, 3, 3), 0)


class StyleTests(unittest.TestCase):
    def test_font_numfmt_xf_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            fi = wb.add_font(FontRecord(name="Calibri", size=12.0, bold=True))
            self.assertGreaterEqual(fi, 0)
            font = wb.get_font(fi)
            self.assertEqual(font.name, "Calibri")
            self.assertTrue(font.bold)

            fill = wb.add_fill(FillRecord(pattern=1, fg_argb=0xFFFF0000))
            nf = wb.add_num_fmt("0.00")
            self.assertGreater(nf, 0)
            self.assertEqual(wb.get_num_fmt(nf), "0.00")

            # add_cell_xf validates indices against the parallel tables, so a
            # border must exist before it can be referenced.
            border = wb.add_border({"left": {"style": 1, "color_argb": 0xFF000000}})

            from formulon import CellXf

            xf = wb.add_cell_xf(
                CellXf(
                    font_index=fi,
                    fill_index=fill,
                    border_index=border,
                    num_fmt_id=nf,
                    horizontal_align=0,
                    vertical_align=0,
                    wrap_text=False,
                    text_rotation=255,
                    indent=0,
                    relative_indent=-3,
                    shrink_to_fit=False,
                    reading_order=0,
                    justify_last_line=True,
                )
            )
            wb.set_cell_xf_index(0, 0, 0, xf)
            self.assertEqual(wb.get_cell_xf_index(0, 0, 0), xf)
            resolved = wb.get_cell_xf(xf)
            self.assertEqual(resolved.font_index, fi)
            self.assertEqual(resolved.num_fmt_id, nf)
            self.assertEqual(resolved.text_rotation, 255)
            self.assertTrue(resolved.has_alignment)
            self.assertEqual(resolved.indent, 0)
            self.assertEqual(resolved.relative_indent, -3)
            self.assertIs(resolved.shrink_to_fit, False)
            self.assertEqual(resolved.reading_order, 0)
            self.assertIs(resolved.justify_last_line, True)
            self.assertFalse(resolved.has_horizontal_align)
            self.assertTrue(resolved.has_vertical_align)
            self.assertFalse(resolved.has_wrap_text)
            self.assertTrue(resolved.has_justify_last_line)
            self.assertEqual(wb.add_cell_xf(resolved), xf)

            omitted = wb.add_cell_xf(
                CellXf(
                    font_index=fi,
                    fill_index=fill,
                    border_index=border,
                    num_fmt_id=nf,
                    horizontal_align=0,
                    vertical_align=2,
                    wrap_text=False,
                )
            )
            self.assertFalse(wb.get_cell_xf(omitted).has_alignment)

            explicit_empty = wb.add_cell_xf(
                CellXf(
                    font_index=fi,
                    fill_index=fill,
                    border_index=border,
                    num_fmt_id=nf,
                    horizontal_align=0,
                    vertical_align=2,
                    wrap_text=False,
                    has_alignment=True,
                )
            )
            self.assertNotEqual(explicit_empty, omitted)
            self.assertTrue(wb.get_cell_xf(explicit_empty).has_alignment)

    def test_font_vert_align_roundtrip_is_the_identity(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.add_font(FontRecord(name="Arial", size=12.0, vert_align=1, color_argb=0xFF112233))
            self.assertEqual(wb.get_font(index).vert_align, 1)

            before = wb.font_count()
            self.assertEqual(wb.add_font(wb.get_font(index)), index)
            self.assertEqual(wb.font_count(), before)

    def test_font_one_field_rewrite_preserves_the_superscript(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.add_font(FontRecord(name="Arial", size=12.0, vert_align=1, color_argb=0xFF112233))
            edited = wb.get_font(index)
            edited.color_argb = 0xFF00FF00
            recolored = wb.add_font(edited)
            self.assertNotEqual(recolored, index)
            reread = wb.get_font(recolored)
            self.assertEqual(reread.vert_align, 1)
            self.assertEqual(reread.color_argb, 0xFF00FF00)

    def test_dxf_font_vert_align_roundtrip_is_the_identity(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.add_dxf(DifferentialFormat(font=FontRecord(name="Calibri", size=9.0, vert_align=1)))
            got = wb.get_dxf(index)
            self.assertIsNotNone(got.font)
            self.assertEqual(got.font.vert_align, 1)

            before = wb.dxf_count()
            self.assertEqual(wb.add_dxf(got), index)
            self.assertEqual(wb.dxf_count(), before)

    def test_dxf_roundtrip_and_dedup(self) -> None:
        with Workbook.create_default() as wb:
            record = DifferentialFormat(
                font=FontRecord(name="Arial", size=12.0, bold=True, color_argb=0xFFFF0000),
                fill=FillRecord(pattern=1, fg_argb=0xFFFFFF00),
                num_fmt_id=164,
                num_fmt_code="0.00",
            )
            index = wb.add_dxf(record)
            self.assertEqual(wb.add_dxf(record), index)
            self.assertEqual(wb.dxf_count(), 1)
            got = wb.get_dxf(index)
            self.assertIsNotNone(got.font)
            self.assertEqual(got.font.name, "Arial")
            self.assertTrue(got.font.bold)
            self.assertIsNotNone(got.fill)
            self.assertEqual(got.fill.fg_argb, 0xFFFFFF00)
            self.assertEqual(got.num_fmt_id, 164)
            self.assertEqual(got.num_fmt_code, "0.00")

    def test_dxf_alignment_and_protection_xml_survive_get_add_and_save_load(self) -> None:
        alignment_xml = '<alignment horizontal="center" wrapText="1"/>'
        protection_xml = '<protection locked="0" hidden="1"/>'
        with Workbook.create_default() as wb:
            alignment_index = wb.add_dxf(DifferentialFormat(alignment_xml=alignment_xml))
            protection_index = wb.add_dxf(DifferentialFormat(protection_xml=protection_xml))
            self.assertNotEqual(alignment_index, protection_index)

            got_alignment = wb.get_dxf(alignment_index)
            self.assertEqual(got_alignment.alignment_xml, alignment_xml)
            self.assertEqual(got_alignment.protection_xml, "")
            got_protection = wb.get_dxf(protection_index)
            self.assertEqual(got_protection.alignment_xml, "")
            self.assertEqual(got_protection.protection_xml, protection_xml)
            self.assertEqual(wb.add_dxf(got_alignment), alignment_index)
            self.assertEqual(wb.add_dxf(got_protection), protection_index)

            with Workbook.load(wb.save()) as reloaded:
                reloaded_alignment = reloaded.get_dxf(alignment_index)
                reloaded_protection = reloaded.get_dxf(protection_index)
                self.assertEqual(reloaded_alignment.alignment_xml, alignment_xml)
                self.assertEqual(reloaded_alignment.protection_xml, "")
                self.assertEqual(reloaded_protection.alignment_xml, "")
                self.assertEqual(reloaded_protection.protection_xml, protection_xml)
                self.assertEqual(reloaded.add_dxf(reloaded_alignment), alignment_index)
                self.assertEqual(reloaded.add_dxf(reloaded_protection), protection_index)

    def test_selector_colours_survive_get_add_identity_and_save_load(self) -> None:
        theme = ColorSpec(kind=2, theme=3, tint=0.5)
        indexed = ColorSpec(kind=3, indexed=9)
        automatic = ColorSpec(kind=4)
        with Workbook.create_default() as wb:
            font_index = wb.add_font(FontRecord(name="SelectorFont", size=11.0, color_argb=0x01020304, color=theme))
            got_font = wb.get_font(font_index)
            self.assertEqual(got_font.color, theme)
            self.assertEqual(got_font.color_argb, 0x01020304)
            self.assertEqual(wb.add_font(got_font), font_index)

            fill_index = wb.add_fill(
                FillRecord(
                    pattern=1,
                    fg_argb=0x05060708,
                    bg_argb=0x090A0B0C,
                    fg=indexed,
                    bg=automatic,
                )
            )
            got_fill = wb.get_fill(fill_index)
            self.assertEqual(got_fill.fg, indexed)
            self.assertEqual(got_fill.bg, automatic)
            self.assertEqual(wb.add_fill(got_fill), fill_index)

            border = {
                "left": {"style": 1, "color_argb": 0x01020304, "color": theme},
                "right": {"style": 1, "color_argb": 0x05060708, "color": indexed},
                "top": {"style": 1, "color_argb": 0x090A0B0C, "color": automatic},
            }
            border_index = wb.add_border(border)
            got_border = wb.get_border(border_index)
            self.assertEqual(got_border["left"]["color"], theme)
            self.assertEqual(got_border["right"]["color"], indexed)
            self.assertEqual(got_border["top"]["color"], automatic)
            self.assertEqual(wb.add_border(got_border), border_index)

            dxf_index = wb.add_dxf(
                DifferentialFormat(
                    font=FontRecord(name="DxfSelector", size=9.0, color_argb=0x11121314, color=automatic),
                    fill=FillRecord(pattern=1, fg_argb=0x15161718, fg=indexed),
                    border={"left": {"style": 1, "color_argb": 0x191A1B1C, "color": theme}},
                )
            )
            got_dxf = wb.get_dxf(dxf_index)
            self.assertEqual(got_dxf.font.color, automatic)
            self.assertEqual(got_dxf.fill.fg, indexed)
            self.assertEqual(got_dxf.border["left"]["color"], theme)
            self.assertEqual(wb.add_dxf(got_dxf), dxf_index)

            with Workbook.load(wb.save()) as reloaded:
                self.assertEqual(reloaded.get_font(font_index).color, theme)
                self.assertEqual(reloaded.get_fill(fill_index).fg, indexed)
                self.assertEqual(reloaded.get_border(border_index)["left"]["color"], theme)
                self.assertEqual(reloaded.get_dxf(dxf_index).font.color, automatic)


class ConditionalFormatTests(unittest.TestCase):
    def test_cf_add_get_evaluate_clear(self) -> None:
        with Workbook.create_default() as wb:
            for r in range(3):
                wb.set_number(0, r, 0, float(r * 5))  # A1=0, A2=5, A3=10
            # A rule's dxf_id must resolve against a registered dxf.
            dxf_index = wb.add_dxf(DifferentialFormat())
            wb.add_conditional_format(
                0,
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 0, 2, 0)],
                    type=1,  # cellIs
                    op_engaged=True,
                    op=4,  # greaterThan
                    formula1="4",
                    dxf_id_engaged=True,
                    dxf_id=dxf_index,
                ),
            )
            rules = wb.get_conditional_formats(0)
            self.assertEqual(len(rules), 1)
            self.assertEqual(rules[0].type, 1)
            self.assertEqual(rules[0].formula1, "4")

            wb.recalc()
            cells = wb.evaluate_cf_range(0, 0, 0, 2, 0)
            matched = {(c.row, c.col) for c in cells if c.matches}
            # A2 (=5) and A3 (=10) exceed 4; A1 (=0) does not.
            self.assertIn((1, 0), matched)
            self.assertIn((2, 0), matched)
            self.assertNotIn((0, 0), matched)

            wb.clear_conditional_formats(0)
            self.assertEqual(wb.get_conditional_formats(0), [])

    def test_visual_rules_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            rules = [
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 0, 2, 0)],
                    type=2,
                    color_scale=ColorScale(
                        [CfValueObject(3), CfValueObject(4)],
                        [CfColor(255, 0, 0), CfColor(0, 255, 0)],
                    ),
                ),
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 1, 2, 1)],
                    type=3,
                    data_bar=DataBar(CfValueObject(3), CfValueObject(4), CfColor(0, 0, 255)),
                ),
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 2, 2, 2)],
                    type=4,
                    icon_set=IconSet(0, [CfValueObject(1, "33"), CfValueObject(1, "67")]),
                ),
            ]
            for rule in rules:
                wb.add_conditional_format(0, rule)
            got = wb.get_conditional_formats(0)
            self.assertEqual(got[0].color_scale.colors[1], CfColor(0, 255, 0))
            self.assertEqual(got[1].data_bar.fill, CfColor(0, 0, 255))
            self.assertEqual(got[2].icon_set.thresholds[1].value, "67")

    # `CfValueObject.type`: 0 num, 1 percent, 2 percentile, 3 min, 4 max,
    # 5 formula, 6 autoMin, 7 autoMax. Every threshold below states a
    # `value`, which is what puts a borrowed string behind each CFVO.
    EXPECTED_CFVO_VALUES = ["0", "50", "100", "5", "95", "0", "33", "67"]

    @staticmethod
    def _cfvo_rules() -> list[ConditionalFormatInput]:
        """Build one color-scale, one data-bar and one icon-set rule."""
        return [
            ConditionalFormatInput(
                sqref=[MergeRange(0, 0, 4, 0)],
                type=2,
                color_scale=ColorScale(
                    [CfValueObject(0, "0"), CfValueObject(1, "50"), CfValueObject(0, "100")],
                    [CfColor(248, 105, 107), CfColor(255, 235, 132), CfColor(99, 190, 123)],
                ),
            ),
            ConditionalFormatInput(
                sqref=[MergeRange(0, 1, 4, 1)],
                type=3,
                data_bar=DataBar(CfValueObject(0, "5"), CfValueObject(0, "95"), CfColor(0, 112, 192)),
            ),
            ConditionalFormatInput(
                sqref=[MergeRange(0, 2, 4, 2)],
                type=4,
                icon_set=IconSet(0, [CfValueObject(1, "0"), CfValueObject(1, "33"), CfValueObject(1, "67")]),
            ),
        ]

    @staticmethod
    def _cfvo_values(rules: list) -> list[str]:
        """Flatten the threshold value strings of the three visual rules."""
        by_type = {rule.type: rule for rule in rules}
        color_scale = by_type[2].color_scale
        data_bar = by_type[3].data_bar
        icon_set = by_type[4].icon_set
        return (
            [t.value for t in color_scale.thresholds]
            + [data_bar.minimum.value, data_bar.maximum.value]
            + [t.value for t in icon_set.thresholds]
        )

    def test_cfvo_value_strings_survive_read_back_and_a_save_load_cycle(self) -> None:
        # Each CFVO `value` crosses the C ABI as a borrowed `const char*`.
        # A store that relocates while the later thresholds are pulled
        # publishes a pointer into freed bytes, so the strings come back
        # wrong or empty even though the rule count and colors look right.
        with Workbook.create_default() as wb:
            for rule in self._cfvo_rules():
                wb.add_conditional_format(0, rule)

            in_session = wb.get_conditional_formats(0)
            self.assertEqual(len(in_session), 3)
            self.assertEqual(self._cfvo_values(in_session), self.EXPECTED_CFVO_VALUES)
            by_type = {rule.type: rule for rule in in_session}
            # The type travels with the value: a percent threshold read
            # back as type 0 would mean an absolute number instead.
            self.assertEqual([t.type for t in by_type[2].color_scale.thresholds], [0, 1, 0])
            self.assertEqual([t.type for t in by_type[4].icon_set.thresholds], [1, 1, 1])
            self.assertEqual(by_type[2].color_scale.colors[2], CfColor(99, 190, 123))
            self.assertEqual(by_type[3].data_bar.fill, CfColor(0, 112, 192))
            saved = wb.save()

        with Workbook.load(saved) as reloaded:
            rules = reloaded.get_conditional_formats(0)
            self.assertEqual(len(rules), 3)
            self.assertEqual(self._cfvo_values(rules), self.EXPECTED_CFVO_VALUES)

    def test_data_bar_x14_fields_survive_save_and_load(self) -> None:
        # These six live in the `x14` extension, not the legacy `<dataBar>`
        # element. An in-session round-trip alone would not catch a writer
        # that never emits the extension, which is how they were lost
        # before: the values came back from the model and disappeared on
        # the way through the file.
        bar = DataBar(
            CfValueObject(3),
            CfValueObject(4),
            CfColor(0, 0, 255),
            gradient=False,
            axis_position=1,
            negative_fill=CfColor(255, 0, 0),
            border=CfColor(9, 9, 9),
            negative_border=CfColor(8, 8, 8),
            axis_color=CfColor(1, 2, 3),
        )
        with Workbook.create_default() as wb:
            wb.add_conditional_format(0, ConditionalFormatInput(sqref=[MergeRange(0, 0, 2, 0)], type=3, data_bar=bar))
            saved = wb.save()

        with Workbook.load(saved) as reloaded:
            got = reloaded.get_conditional_formats(0)[0].data_bar
            self.assertIs(got.gradient, False)
            self.assertEqual(got.axis_position, 1)
            self.assertEqual(got.negative_fill, CfColor(255, 0, 0))
            self.assertEqual(got.border, CfColor(9, 9, 9))
            self.assertEqual(got.negative_border, CfColor(8, 8, 8))
            self.assertEqual(got.axis_color, CfColor(1, 2, 3))
            # Feeding the decoded bar straight back must reproduce it.
            reloaded.add_conditional_format(
                0, ConditionalFormatInput(sqref=[MergeRange(4, 0, 6, 0)], type=3, data_bar=got)
            )
            self.assertEqual(reloaded.get_conditional_formats(0)[1].data_bar, got)

    def test_omitted_data_bar_x14_fields_keep_the_model_defaults(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_conditional_format(
                0,
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 0, 2, 0)],
                    type=3,
                    data_bar=DataBar(CfValueObject(3), CfValueObject(4), CfColor(0, 0, 255)),
                ),
            )
            got = wb.get_conditional_formats(0)[0].data_bar
            # The getter engages all six, so they read back as the defaults
            # rather than as `None`.
            self.assertIs(got.gradient, True)
            self.assertEqual(got.axis_position, 0)
            self.assertEqual(got.negative_fill, CfColor(0, 0, 255))


class PivotTests(unittest.TestCase):
    def test_build_and_project_pivot(self) -> None:
        with Workbook.create_default() as wb:
            cache_id = wb.pivot_cache_create()
            wb.pivot_cache_field_add(cache_id, "Region")
            wb.pivot_cache_field_add(cache_id, "Amount")
            for region, amount in [("East", 10.0), ("East", 20.0), ("West", 30.0)]:
                rec = wb.pivot_cache_record_add(cache_id)
                wb.pivot_cache_record_set_text(cache_id, rec, 0, region)
                wb.pivot_cache_record_set_number(cache_id, rec, 1, amount)
            self.assertEqual(wb.pivot_cache_record_count(cache_id), 3)

            pivot = wb.pivot_create(0, "Pivot1", cache_id, 0, 4)
            region_field = wb.pivot_field_add(0, pivot, PivotFieldSpec(source_name="Region", axis=PivotAxis.ROW))
            amount_field = wb.pivot_field_add(0, pivot, PivotFieldSpec(source_name="Amount", axis=PivotAxis.VALUE))
            self.assertEqual(region_field, 0)
            wb.pivot_data_field_add(
                0,
                pivot,
                PivotDataFieldSpec(
                    name="Sum of Amount",
                    field_index=amount_field,
                    aggregation=PivotAggregation.SUM,
                ),
            )

            layout = wb.pivot_layout(0, pivot)
            self.assertGreater(len(layout.cells), 0)
            numbers = [c.value.to_python() for c in layout.cells if c.value.kind == ValueKind.NUMBER]
            # The single grand total of 10 + 20 + 30 must appear.
            self.assertIn(60.0, numbers)

    def test_cache_source_and_report_layout_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            cache = wb.pivot_cache_create()
            wb.set_pivot_cache_worksheet_source(cache, PivotWorksheetSource(ref="A1:B9", sheet="Sheet1"))
            self.assertEqual(wb.get_pivot_cache_worksheet_source(cache).sheet, "Sheet1")
            pivot = wb.pivot_create(0, "Layout", cache, 0, 0)
            wb.set_pivot_report_layout(0, pivot, PivotReportLayout.TABULAR)
            self.assertEqual(wb.get_pivot_report_layout(0, pivot), PivotReportLayout.TABULAR)
            wb.set_pivot_cache_worksheet_source(cache, None)
            self.assertIsNone(wb.get_pivot_cache_worksheet_source(cache))


class FunctionCatalogTests(unittest.TestCase):
    def test_catalog_metadata(self) -> None:
        self.assertGreater(Workbook.function_count(), 0)
        meta = Workbook.function_metadata("SUM", 0)
        self.assertIsNotNone(meta)
        self.assertEqual(meta.name, "SUM")
        self.assertGreaterEqual(meta.min_arity, 1)
        # SUM is an unbounded variadic; the sentinel is normalized to None.
        self.assertIsNone(meta.max_arity)
        # Lazy-dispatch forms (not in the eager registry) still resolve.
        xlookup = Workbook.function_metadata("XLOOKUP", 0)
        self.assertIsNotNone(xlookup)
        self.assertEqual(xlookup.name, "XLOOKUP")
        names = {Workbook.function_name_at(i) for i in range(Workbook.function_count())}
        self.assertIn("XLOOKUP", names)
        self.assertIsNone(Workbook.function_metadata("NOT_A_REAL_FUNCTION", 0))

    def test_function_name_at(self) -> None:
        name = Workbook.function_name_at(0)
        self.assertIsInstance(name, str)
        self.assertGreater(len(name), 0)

    def test_merge_function_metadata(self) -> None:
        base = Workbook.function_metadata("XLOOKUP", 0)
        self.assertIsNotNone(base)
        # The engine leaves display metadata empty.
        self.assertIsNone(base.signature_template)
        self.assertIsNone(base.description)

        entry = {
            "signature": "XLOOKUP(lookup_value, lookup_array, return_array)",
            "description": "Searches a range or an array.",
            "aliases": {"fr-FR": "RECHERCHEX"},
            "localized": {"fr-FR": {"signature": "RECHERCHEX(...)", "description": "Recherche."}},
        }

        # Localized override wins for the matching locale.
        fr = formulon.merge_function_metadata(base, entry, "fr-FR")
        self.assertEqual(fr.signature_template, "RECHERCHEX(...)")
        self.assertEqual(fr.description, "Recherche.")
        self.assertEqual(fr.localized_name, "RECHERCHEX")
        # Structural fields survive the merge.
        self.assertEqual(fr.name, "XLOOKUP")

        # A locale with no localized/alias entry falls back to the default
        # signature/description and the canonical display name.
        de = formulon.merge_function_metadata(base, entry, "de-DE")
        self.assertEqual(
            de.signature_template,
            "XLOOKUP(lookup_value, lookup_array, return_array)",
        )
        self.assertEqual(de.description, "Searches a range or an array.")
        self.assertEqual(de.localized_name, "XLOOKUP")

        # No provider entry -> base returned verbatim; display metadata NULL.
        none = formulon.merge_function_metadata(base, None, "fr-FR")
        self.assertIs(none, base)
        self.assertIsNone(none.signature_template)
        self.assertIsNone(none.description)

    def test_shipped_example_document_conforms_to_the_engine(self) -> None:
        """The example provider document must stay usable against this engine.

        ``docs/examples/function-metadata.example.json`` is offered as a
        ready-to-use starting point, but nothing else loads it, so a
        function rename or a schema change would leave a broken example
        shipped in the docs. Checking it here also pins the two claims the
        schema doc makes about the keys -- canonical, uppercase, English --
        against the engine that decides what canonical means.
        """
        path = Path(__file__).resolve().parents[3] / "docs" / "examples" / "function-metadata.example.json"
        document = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(document["version"], 1)
        functions = document["functions"]
        self.assertTrue(functions, "the example document declares no functions")

        for name, entry in functions.items():
            self.assertEqual(name, name.upper(), f"{name} is not an uppercase key")
            base = Workbook.function_metadata(name, 0)
            self.assertIsNotNone(base, f"{name} is not a function this engine knows")
            self.assertEqual(base.name, name, f"{name} is not the canonical spelling")
            # Merging must actually reach the entry: an alias-only entry
            # (VLOOKUP here) still has to resolve its display name, which is
            # what distinguishes a real merge from returning `base` verbatim.
            for locale in sorted(set(entry.get("aliases", {})) | set(entry.get("localized", {}))):
                merged = formulon.merge_function_metadata(base, entry, locale)
                self.assertEqual(merged.name, name)
                alias = entry.get("aliases", {}).get(locale)
                self.assertEqual(merged.localized_name, alias if alias is not None else name)


class ExternalLinkTests(unittest.TestCase):
    def test_fresh_workbook_has_no_external_links(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.external_link_count(), 0)
            self.assertEqual(wb.get_external_links(), [])


class PackageDiagnosticsTests(unittest.TestCase):
    def test_unknown_save_format_raises_formulon_error(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(formulon.FormulonError) as ctx:
                wb.save_with_diagnostics(WorkbookFormat.UNKNOWN)
            self.assertIn("unsupported format", str(ctx.exception))

    def test_save_and_read_diagnostics_report_every_counter(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_formula(0, 0, 0, "=@A1:A10")
            wb.add_validation(
                0,
                DataValidationInput(
                    type=3,
                    ranges=[MergeRange(0, 0, 0, 0)],
                    formula1='"Yes,No"',
                ),
            )
            wb.add_conditional_format(
                0,
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 0, 0, 0)],
                    type=1,
                    op_engaged=True,
                    op=5,
                    formula1="10",
                ),
            )

            xlsx = wb.save_with_diagnostics(WorkbookFormat.XLSX)
            self.assertIsInstance(xlsx, SaveDiagnostics)
            self.assertIsInstance(xlsx.bytes, bytes)
            self.assertEqual(xlsx.downgraded_formula_count, 0)
            self.assertEqual(xlsx.deferred_feature_count, 0)
            self.assertEqual(xlsx.dropped_part_count, 0)
            self.assertEqual(xlsx.dropped_relationship_count, 0)
            self.assertEqual(xlsx.renumbered_part_count, 0)

            xlsb = wb.save_with_diagnostics(WorkbookFormat.XLSB)
            self.assertEqual(xlsb.downgraded_formula_count, 1)
            self.assertGreaterEqual(xlsb.deferred_feature_count, 2)
            # The binary writer never reassigns a part id.
            self.assertEqual(xlsb.renumbered_part_count, 0)

            # `save_as` remains a plain bytes return value.
            self.assertIsInstance(wb.save_as(WorkbookFormat.XLSB), bytes)

            with Workbook.load(xlsb.bytes) as loaded:
                read = loaded.read_diagnostics()
                self.assertIsInstance(read, ReadDiagnostics)
                self.assertEqual(read.undecoded_formula_count, 0)
                self.assertEqual(read.undecoded_defined_name_count, 0)
                self.assertEqual(read.undecoded_part_count, 0)
                self.assertEqual(read.skipped_feature_count, 0)
                self.assertEqual(read.unknown_content_type_count, 0)

            with Workbook.load(_append_empty_zip_entry(xlsb.bytes, "xl/preserved.bin")) as preserved_loaded:
                preserved = preserved_loaded.read_diagnostics()
                self.assertEqual(preserved.undecoded_part_count, 0)

    def test_diagnostic_scratch_is_freed_when_a_later_allocation_fails(self) -> None:
        # `save_with_diagnostics` takes three blocks (two out-pointers plus the
        # counter struct). If the second allocation raises, the first must
        # still be released rather than leaked into the WASM heap.
        with Workbook.create_default() as wb:
            original_alloc = formulon.workbook.LIB.alloc
            calls = 0

            def alloc_then_fail(size: int) -> int:
                nonlocal calls
                calls += 1
                if calls == 2:
                    raise RuntimeError("synthetic allocation failure")
                return original_alloc(size)

            with (
                mock.patch.object(formulon.workbook.LIB, "alloc", side_effect=alloc_then_fail),
                mock.patch.object(formulon.workbook.LIB, "free", wraps=formulon.workbook.LIB.free) as free,
            ):
                with self.assertRaises(RuntimeError):
                    wb.save_with_diagnostics(WorkbookFormat.XLSB)
            self.assertEqual(free.call_count, 1)


class SurfaceParityTests(unittest.TestCase):
    """Every advertised method must exist on the Workbook class."""

    METHODS = [
        "move_sheet",
        "remove_sheet",
        "rename_sheet",
        "set_defined_name",
        "set_defined_name_scoped",
        "set_error",
        "insert_rows",
        "delete_rows",
        "insert_cols",
        "delete_cols",
        "calc_mode",
        "set_calc_mode",
        "pinned_now",
        "set_pinned_now",
        "clear_pinned_now",
        "excel_profile_id",
        "set_excel_profile_id",
        "partial_recalc",
        "lambda_text_at",
        "evaluate_formula_array",
        "add_merge",
        "remove_merge",
        "remove_merge_at",
        "clear_merges",
        "get_merges",
        "merge_count",
        "add_hyperlink",
        "remove_hyperlink",
        "remove_hyperlink_at",
        "clear_hyperlinks",
        "get_hyperlinks",
        "hyperlink_count",
        "get_comment",
        "set_comment",
        "validation_count",
        "get_validation_at",
        "get_validations",
        "add_validation",
        "remove_validation_at",
        "clear_validations",
        "get_sheet_protection",
        "set_sheet_protection",
        "get_sheet_view",
        "set_sheet_zoom",
        "set_sheet_freeze",
        "set_sheet_tab_hidden",
        "set_sheet_visibility",
        "set_sheet_show_grid_lines",
        "set_sheet_show_row_col_headers",
        "set_sheet_show_zeros",
        "set_sheet_right_to_left",
        "set_sheet_tab_selected",
        "set_sheet_view_mode",
        "get_sheet_columns",
        "set_column_width",
        "set_column_hidden",
        "set_column_outline",
        "get_sheet_row_overrides",
        "set_row_height",
        "set_row_hidden",
        "set_row_outline",
        "evaluate_cf_range",
        "cf_count",
        "get_conditional_format_at",
        "get_conditional_formats",
        "add_conditional_format",
        "remove_conditional_format_at",
        "clear_conditional_formats",
        "get_cell_xf_index",
        "set_cell_xf_index",
        "get_cell_xf",
        "get_font",
        "get_fill",
        "get_border",
        "get_num_fmt",
        "font_count",
        "fill_count",
        "border_count",
        "cell_xf_count",
        "cell_style_count",
        "cell_style_xf_count",
        "add_font",
        "add_fill",
        "add_border",
        "add_num_fmt",
        "add_cell_xf",
        "get_cell_style",
        "get_cell_style_xf",
        "pivot_count",
        "pivot_layout",
        "pivot_cache_count",
        "pivot_cache_id_at",
        "pivot_cache_create",
        "pivot_cache_remove",
        "pivot_cache_field_count",
        "pivot_cache_field_name",
        "pivot_cache_field_add",
        "pivot_cache_field_clear",
        "pivot_cache_record_add",
        "pivot_cache_record_set_number",
        "pivot_cache_record_set_text",
        "pivot_create",
        "pivot_remove",
        "pivot_set_name",
        "pivot_set_anchor",
        "pivot_set_grand_totals",
        "pivot_field_add",
        "pivot_field_add_item",
        "pivot_field_add_item_at",
        "pivot_field_set_axis",
        "pivot_data_field_add",
        "pivot_data_field_set",
        "pivot_filter_add",
        "pivot_set_row_field_order",
        "pivot_set_col_field_order",
        "precedents",
        "dependents",
        "spill_info",
        "function_count",
        "function_name_at",
        "function_metadata",
        "localize_function_name",
        "canonicalize_function_name",
        "external_link_count",
        "get_external_link_at",
        "get_external_links",
        "save",
        "save_as",
        "save_with_diagnostics",
        "read_diagnostics",
    ]

    def test_all_methods_present(self) -> None:
        for name in self.METHODS:
            self.assertTrue(hasattr(Workbook, name), f"Workbook missing method: {name}")

    def test_public_types_exported(self) -> None:
        for name in (
            "MergeRange",
            "Comment",
            "Hyperlink",
            "DataValidation",
            "DataValidationInput",
            "SheetProtection",
            "SheetView",
            "ConditionalFormat",
            "ConditionalFormatInput",
            "CfMatch",
            "CfCellResult",
            "CellNode",
            "SpillInfo",
            "FunctionMetadata",
            "CellXf",
            "ColorSpec",
            "FontRecord",
            "FillRecord",
            "CellStyle",
            "ExternalLink",
            "PivotCell",
            "PivotLayout",
            "PivotFieldSpec",
            "PivotDataFieldSpec",
            "PivotFilterSpec",
            "CalcMode",
            "PivotAxis",
            "PivotAggregation",
            "MergedFunctionMetadata",
            "merge_function_metadata",
        ):
            self.assertTrue(hasattr(formulon, name), f"formulon missing: {name}")


if __name__ == "__main__":
    unittest.main()
