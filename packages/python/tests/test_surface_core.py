"""Core Workbook surface tests."""

from __future__ import annotations

import datetime
import subprocess
import unittest
from pathlib import Path

from formulon import (
    ArrayFormula,
    CalcMode,
    CellRect,
    DefinedNameInput,
    FormulonError,
    InvalidArgumentError,
    ValueKind,
    Workbook,
    _c,
    to_native,
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
        "CELL_RECT": 16,
        "XF_CELL": 12,
        "HYPERLINK": 36,
        "COMMENT": 20,
        "DATA_VALIDATION": 68,
        "DATA_VALIDATION_PATCH": 120,
        "CF_RULE_PATCH": 472,
        "SHEET_PROTECTION": 84,
        "SHEET_PROTECTION_PATCH": 164,
        "WORKBOOK_PROTECTION": 56,
        "WORKBOOK_LOCATION": 12,
        "PROTECTED_RANGE": 40,
        "CELL_NODE": 16,
        "PHONETIC_RUN": 20,
        "CFVO": 12,
        "CF_MATCH": 136,
        "CF_RULE": 352,
        "PIVOT_CELL": 48,
        "PIVOT_DATA_FIELD_SPEC": 40,
        "PIVOT_FILTER_SPEC": 64,
        "PIVOT_FIELD_INFO": 72,
        "PIVOT_ITEM_INFO": 28,
        "PIVOT_STYLE_OPTIONS": 28,
        "PIVOT_OPTIONS": 108,
        "PIVOT_CACHE_FIELD_TYPES_INFO": 36,
        "PIVOT_CACHE_OPTIONS": 20,
        "PIVOT_CALCULATED_ITEM": 20,
        "TABLE_INFO": 56,
        "TABLE_COLUMN_INFO": 24,
        "TABLE_STYLE_ELEMENT": 16,
        "TABLE_STYLE_OPTIONS": 24,
        "TABLE_STYLE_OPTIONS_PATCH": 44,
        "TABLE_STYLE": 24,
        "READ_DIAGNOSTICS": 24,
        "SAVE_DIAGNOSTICS": 24,
        "SPILL_INFO": 36,
        "ARRAY_FORMULA": 24,
        "FUNCTION_METADATA": 32,
        "LOCALE_FACTS": 60,
        "CIVIL_TIME": 24,
        "SHEET_VIEW": 80,
        "COLUMN_LAYOUT": 48,
        "ROW_LAYOUT": 48,
        "CELL_XF": 136,
        "COLOR_SPEC": 24,
        "FONT_RECORD": 128,
        "FILL_RECORD": 112,
        "GRADIENT_STOP": 32,
        "CF_ICON": 8,
        "BORDER_SIDE": 32,
        "BORDER_RECORD": 168,
        "DXF_RECORD": 448,
        "CELL_STYLE_RECORD": 28,
        "EXTERNAL_LINK_RECORD": 28,
        "ROW_BREAK": 20,
        "COLUMN_BREAK": 20,
        "PAGE_SETUP": 52,
        "PAGE_MARGINS": 96,
        "PRINT_OPTIONS": 36,
        "HEADER_FOOTER": 60,
        "SHEET_FORMAT_DEFAULTS": 64,
        "RECT_PT": 40,
        "WIDTH_MODEL": 48,
        "PAPER_INFO": 32,
        "MARGINS_PT": 56,
        "PRINT_TITLES": 28,
        "PAGE_LAYOUT": 56,
        "THEME_COLORS": 52,
        "THEME_FONTS": 20,
        "EFFECTIVE_STYLE": 100,
        "DATE_GROUP_ITEM": 12,
        "FILTER_COLUMN": 144,
        "SORT_CONDITION": 48,
        "AUTO_FILTER": 68,
        "VALIDATION_OUTCOME": 20,
        "MENTION": 24,
        "THREADED_COMMENT": 44,
        "PERSON": 20,
        "IMAGE_INFO": 16,
        "DRAWING_OBJECT": 104,
        "DRAWING_PLACEMENT_SPEC": 88,
        "DRAWING_ANCHOR_PATCH": 144,
        "DRAWING_PATCH": 176,
        "SHAPE_SPEC": 240,
        "SLICER_SPEC": 128,
        "DRAWING_ANCHOR": 72,
        "SHAPE": 160,
        "DATA_TABLE": 60,
        "SLICER": 56,
        "SPARKLINE": 8,
        "SPARKLINE_GROUP": 296,
        "CHART_TITLE": 12,
        "CHART_FILL": 48,
        "CHART_LINE": 72,
        "CHART_MARKER": 136,
        "CHART_LABELS": 88,
        "CHART_INFO": 100,
        "CHART_PLOT": 68,
        "CHART_DATA_SOURCE": 12,
        "CHART_SERIES": 456,
        "CHART_LITERAL": 32,
        "CHART_POINT": 272,
        "CHART_AXIS": 152,
        "CHART_PREVIEW_OVERRIDES": 60,
        "CHART_DATA_INFO": 48,
        "CHART_DATA_SERIES": 60,
        "CHART_DATA_POINTS": 24,
        "SHEET_RECT": 20,
        "CHART_AXIS_TITLE_TEXT": 12,
        "CHART_UNRESOLVED": 16,
        "CHART_FILL_PATCH": 56,
        "CHART_LINE_PATCH": 80,
        "CHART_MARKER_PATCH": 160,
        "CHART_LABELS_PATCH": 128,
        "CHART_SERIES_PATCH": 576,
        "CHART_POINT_PATCH": 312,
        "CHART_PLOT_PATCH": 44,
        "CHART_AXIS_PATCH": 184,
        "CHART_LEGEND": 20,
        "CHART_LEGEND_PATCH": 32,
        "CHART_DISPLAY_PATCH": 36,
        "REFERENCE": 24,
        "ITERATIVE_SETTINGS_PATCH": 32,
        "DEFINED_NAME": 24,
        "TABLE_SPEC": 64,
        "PIVOT_WORKSHEET_SOURCE": 16,
        "PIVOT_STYLE_OPTIONS_PATCH": 52,
        "PIVOT_SPEC": 20,
        "PIVOT_ITEM_PATCH": 28,
        "PIVOT_FIELD_PATCH": 84,
        "PIVOT_DATE_GROUP": 48,
        "PIVOT_NUMERIC_GROUP": 40,
        "REFERENCE_NOTATION_SPEC": 20,
        "FILTER_CONDITION": 8,
        "CHART_DATA_SOURCE_SPEC": 16,
        "THREADED_COMMENT_PATCH": 32,
        "PARSE_ERROR": 36,
        "HEADER_FOOTER_PATCH": 84,
        "ITERATIVE_SETTINGS": 24,
        "PAGE_MARGINS_PATCH": 96,
        "PAGE_SETUP_PATCH": 52,
        "PIVOT_CACHE_OPTIONS_PATCH": 28,
        "PIVOT_GRAND_TOTALS_PATCH": 20,
        "PIVOT_OPTIONS_PATCH": 212,
        "PRINT_OPTIONS_PATCH": 36,
        "PROTECTED_RANGE_PATCH": 72,
        "SHAPE_PATCH": 184,
        "SHEET_FORMAT_DEFAULTS_PATCH": 64,
        "SHEET_VIEW_PATCH": 152,
        "SLICER_PATCH": 56,
        "SPARKLINE_GROUP_PATCH": 432,
        "TABLE_PATCH": 52,
        "TABLE_COLUMN_PATCH": 32,
        "TABLE_STYLE_DEFAULTS_PATCH": 20,
        "PHONETIC_PROPERTIES": 20,
        "CELL_ADDRESS": 8,
        "XF_REMAP": 8,
        "CF_RANGE_EVALUATION_SPEC": 24,
        "NUMBER_FORMAT_INFO": 12,
        "CHART_SPEC": 144,
        "PIVOT_GRAND_TOTALS": 12,
        "TABLE_STYLE_DEFAULTS": 12,
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
        self.assertEqual(S.PIVOT_CELL_VALUE_OFFSET, 16)

    # Records the wrapper marshals without a `Struct` entry, so they are
    # invisible to `EXPECTED_SIZES` above. `fm_value_t` is passed around as
    # the bare `fm_value_t_size` and as `VALUE_BLOB` when it sits inline in a
    # larger record. `python-inline-structs` binds these to the C header;
    # pinning them here makes a size change loud on its own.
    def test_hand_marshalled_record_sizes(self) -> None:
        self.assertEqual(_c.fm_value_t_size, 16)
        self.assertEqual(S.VALUE_BLOB, ("blob16", 16, 8))

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
            grid = wb.sheet(0).evaluate_typed_formula_array(0, 0, "=SEQUENCE(2,3)")
            self.assertEqual(len(grid), 2)
            self.assertEqual(len(grid[0]), 3)
            # Row-major 1..6.
            self.assertEqual(to_native(grid[0][0]), 1.0)
            self.assertEqual(to_native(grid[1][2]), 6.0)

    def test_scalar_reported_as_one_by_one(self) -> None:
        with Workbook.create_default() as wb:
            grid = wb.sheet(0).evaluate_typed_formula_array(0, 0, "=1+2")
            self.assertEqual(len(grid), 1)
            self.assertEqual(len(grid[0]), 1)
            self.assertEqual(grid[0][0].kind, ValueKind.NUMBER)
            self.assertEqual(to_native(grid[0][0]), 3.0)


class DefinedNameTests(unittest.TestCase):
    def test_defined_name_set_get_remove(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_defined_name(DefinedNameInput("MyRef", "Sheet1!$A$1"))
            names = {dn.name: dn.formula for dn in wb.list_defined_names()}
            self.assertIn("MyRef", names)
            self.assertEqual(names["MyRef"], "=Sheet1!$A$1")

            # Updating in place replaces the formula text.
            wb.set_defined_name(DefinedNameInput("MyRef", "Sheet1!$B$2"))
            names = {dn.name: dn.formula for dn in wb.list_defined_names()}
            self.assertEqual(names["MyRef"], "=Sheet1!$B$2")

            # An empty formula is rejected; removal has its own call.
            with self.assertRaises(InvalidArgumentError):
                wb.set_defined_name(DefinedNameInput("MyRef", ""))
            wb.remove_defined_name("MyRef")
            self.assertNotIn("MyRef", {dn.name for dn in wb.list_defined_names()})

    def test_scoped_defined_name_set_get_remove(self) -> None:
        with Workbook.create_empty() as wb:
            wb.add_sheet("Sheet1")
            wb.add_sheet("Sheet2")
            wb.set_defined_name(DefinedNameInput("Rate", "=1"))
            wb.set_defined_name(DefinedNameInput("Rate", "=2", 1))
            names = {(dn.name, dn.scope): dn.formula for dn in wb.list_defined_names()}
            # The reader returns the formula with one leading "=", whatever the input had.
            self.assertEqual(names[("Rate", None)], "=1")
            self.assertEqual(names[("Rate", 1)], "=2")

            wb.remove_defined_name("Rate", 1)
            self.assertNotIn(
                ("Rate", 1),
                {(dn.name, dn.scope) for dn in wb.list_defined_names()},
            )


class CalcPolicyTests(unittest.TestCase):
    def test_calc_mode_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.get_calc_mode(), CalcMode.AUTO)
            wb.set_calc_mode(CalcMode.MANUAL)
            self.assertEqual(wb.get_calc_mode(), CalcMode.MANUAL)

    def test_pinned_now_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            # Unpinned by default: the workbook follows the host clock.
            self.assertIsNone(wb.get_pinned_now())
            wb.set_pinned_now(datetime.datetime(2026, 4, 23, 15, 30, 45))
            self.assertEqual(wb.get_pinned_now(), datetime.datetime(2026, 4, 23, 15, 30, 45))
            # 2026-04-23 is serial 46135 under the 1900 date system.
            grid = wb.sheet(0).evaluate_typed_formula_array(0, 0, "=TODAY()")
            self.assertEqual(grid[0][0].number, 46135.0)
            wb.set_pinned_now(None)
            self.assertIsNone(wb.get_pinned_now())

    def test_pinned_now_rejects_a_non_calendar_instant(self) -> None:
        with Workbook.create_default() as wb:
            # The engine accepts years 1900..9999 only.
            with self.assertRaises(InvalidArgumentError):
                wb.set_pinned_now(datetime.datetime(1899, 12, 31))
            with self.assertRaises(TypeError):
                wb.set_pinned_now("2026-01-01")  # type: ignore[arg-type]
            self.assertIsNone(wb.get_pinned_now())

    def test_excel_profile_roundtrip(self) -> None:
        with Workbook.create_default() as wb:
            self.assertTrue(wb.get_profile_id())
            wb.set_profile_id("mac-365-ja_JP")
            self.assertEqual(wb.get_profile_id(), "mac-365-ja_JP")

    def test_excel_profile_switches_live_workbook(self) -> None:
        # mac-365-* values are the oracle captures of DOLLAR(-1234.567); the
        # estimated win-* ids are checked for the id round trip only.
        dollar = {
            "mac-365-ja_JP": "\u00a5-1,235",
            "mac-365-en_US": "($1,234.57)",
            "mac-365-de_DE": "-1.234,57 \u20ac",
            "mac-365-fr_FR": "(1 234,57 \u20ac)",
            "mac-365-zh_CN": "(\u00a51,234.57)",
            "mac-365-ko_KR": "(\u20a91,235)",
            "mac-365-th_TH": "(\u0e3f1,234.57)",
            "mac-365-es_ES": "-1.234,57 \u20ac",
            "mac-365-es_MX": "-$1,234.57",
            "mac-365-pt_BR": "-R$ 1.234,57",
            "mac-365-ru_RU": "-1 234,57 \u20bd",
            "mac-365-zh_TW": "($1,234.57)",
            "mac-365-it_IT": "-1.234,57 \u20ac",
            "mac-365-nl_NL": "(\u20ac 1.234,57)",
        }
        locales = (
            "ja_JP",
            "en_US",
            "de_DE",
            "fr_FR",
            "zh_CN",
            "ko_KR",
            "th_TH",
            "es_ES",
            "es_MX",
            "pt_BR",
            "ru_RU",
            "zh_TW",
            "it_IT",
            "nl_NL",
        )
        with Workbook.create_default() as wb:
            self.assertEqual(wb.get_profile_id(), "win-365-en_US")
            wb.sheet(0).set_formula(0, 0, "=DOLLAR(-1234.567)")
            for host in ("mac", "win"):
                for locale in locales:
                    profile_id = f"{host}-365-{locale}"
                    wb.set_profile_id(profile_id)
                    self.assertEqual(wb.get_profile_id(), profile_id)
                    wb.recalc()
                    if profile_id in dollar:
                        self.assertEqual(to_native(wb.sheet(0).get_typed_value(0, 0)), dollar[profile_id], profile_id)
            with self.assertRaises(FormulonError):
                wb.set_profile_id("mac-365-xx_XX")


class PartialRecalcTests(unittest.TestCase):
    def test_partial_recalc_recomputes_chain(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 10.0)  # A1
            wb.sheet(0).set_formula(0, 1, "=A1*2")  # B1
            recomputed = wb.sheet(0).recalc_viewport(CellRect(0, 1, 0, 1))
            self.assertGreater(recomputed, 0)
            self.assertEqual(to_native(wb.sheet(0).get_typed_value(0, 1)), 20.0)
            with self.assertRaises(TypeError):
                wb.sheet(0).recalc_viewport((0, 1, 0, 1))  # type: ignore[arg-type]


class TraceTests(unittest.TestCase):
    def test_precedents_and_dependents(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)  # A1
            wb.sheet(0).set_formula(0, 1, "=A1")  # B1
            wb.recalc()
            prec = wb.sheet(0).get_precedents(0, 1, 1)
            self.assertIn((0, 0, 0), [(p.sheet, p.row, p.col) for p in prec])
            deps = wb.sheet(0).get_dependents(0, 0, 1)
            self.assertIn((0, 0, 1), [(d.sheet, d.row, d.col) for d in deps])


class SpillTests(unittest.TestCase):
    def test_sequence_spills(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_formula(0, 0, "=SEQUENCE(3)")
            wb.recalc()
            info = wb.sheet(0).get_spill(0, 0)
            self.assertIsNotNone(info)
            self.assertEqual(info.range, CellRect(0, 0, 2, 0))
            self.assertFalse(info.blocked)

    def test_blocked_spill_reports_the_attempted_footprint(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_formula(0, 0, "=SEQUENCE(5)")
            wb.sheet(0).set_number(3, 0, 1.0)
            wb.recalc()
            info = wb.sheet(0).get_spill(0, 0)
            self.assertTrue(info.blocked)
            self.assertEqual(info.range, CellRect(0, 0, 4, 0))
            self.assertIsNone(wb.sheet(0).get_spill(1, 0))
            self.assertIsNone(wb.sheet(0).get_spill(9, 9))


class ArrayFormulaTests(unittest.TestCase):
    def test_set_get_clear_and_refusals(self) -> None:
        with Workbook.create_default() as wb:
            for r in range(3):
                wb.sheet(0).set_number(r, 0, float(r + 1))
            self.assertIsNone(wb.sheet(0).get_array_formula(0, 2))
            wb.sheet(0).set_array_formula(ArrayFormula(CellRect(0, 2, 2, 2), "=A1:A3*2"))
            wb.recalc()
            for r in range(3):
                found = wb.sheet(0).get_array_formula(r, 2)
                self.assertIsNotNone(found)
                assert found is not None
                self.assertEqual(found.range, CellRect(0, 2, 2, 2))
                self.assertIn("A1:A3", found.formula)
            self.assertEqual(to_native(wb.sheet(0).get_typed_value(2, 2)), 6.0)
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_array_formula(ArrayFormula(CellRect(1, 2, 2, 2), "=1"))
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_number(1, 2, 1.0)
            wb.sheet(0).set_blank(0, 2)
            wb.recalc()
            self.assertIsNone(wb.sheet(0).get_array_formula(1, 2))


class ExternalLinkTests(unittest.TestCase):
    def test_fresh_workbook_has_no_external_links(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(len(wb.list_external_links()), 0)
            self.assertEqual(wb.list_external_links(), [])


if __name__ == "__main__":
    unittest.main()
