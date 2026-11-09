"""Contract tests for the Python wrapper's own error and scratch policies.

These cover behaviour the wrapper owns rather than the engine: which
exception type a defensive path raises, what a failure's ``op`` prefix
names, how the WASM-locator error reports its candidates, and the
allocation discipline of the collection readers.
"""

from __future__ import annotations

import ast
import datetime
import inspect
import io
import os
import struct
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

import formulon
from formulon import (
    AnchorKind,
    Border,
    BorderSide,
    CalcMode,
    CellAddress,
    CellRect,
    CellStyle,
    CellXfInput,
    CfCellIsRuleInput,
    CfDataBarInput,
    CfDataBarRuleInput,
    CfTimePeriodRuleInput,
    CfValueObject,
    Color,
    ColorKind,
    Comment,
    ComparisonOperator,
    DefinedNameInput,
    DrawingAnchorInput,
    DxfInput,
    ErrorCode,
    ExcelError,
    ExternalLinkKind,
    Fill,
    Font,
    FormulonError,
    LocaleDateOrder,
    LogLevel,
    PhoneticProperties,
    PhoneticRun,
    PivotAggregation,
    PivotAxis,
    PivotDataFieldInput,
    PivotFilterInput,
    PivotFilterKind,
    PivotGrandTotals,
    PivotInput,
    ShapeInput,
    Sheet,
    SheetRect,
    SheetSelection,
    SheetSelectionInput,
    TableColumnTotalsInput,
    TableInput,
    TableStyle,
    TableStyleDefaults,
    TableStyleElement,
    TableStyleElementType,
    TableTotalsRowFunction,
    Value,
    ValueKind,
    Workbook,
    _c,
    canonicalize_formula,
    canonicalize_function_name,
    get_function_metadata,
    get_locale_facts,
    localize_formula,
    localize_function_name,
)


def _rgb(red: int, green: int, blue: int) -> Color:
    return Color(ColorKind.RGB, argb=0xFF000000 | (red << 16) | (green << 8) | blue)


class IndexGetterTests(unittest.TestCase):
    def test_an_out_of_range_index_raises(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.ooxml.get_font(len(wb.ooxml.list_fonts()))

    def test_empty_collections_list_as_empty(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            for name, listed in (
                ("list_cf_rules", ws.list_cf_rules()),
                ("list_validations", ws.list_validations()),
                ("list_protected_ranges", ws.list_protected_ranges()),
                ("list_sparkline_groups", ws.list_sparkline_groups()),
                ("list_external_links", wb.list_external_links()),
                ("get_cell_phonetic_runs", ws.get_cell_phonetic_runs(0, 0)),
                ("list_tables", ws.list_tables()),
            ):
                with self.subTest(name=name):
                    self.assertEqual(listed, [])

    def test_counts_start_at_zero(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(
                [
                    len(wb.sheet(0).list_cf_rules()),
                    len(wb.sheet(0).list_validations()),
                    len(wb.sheet(0).list_protected_ranges()),
                    len(wb.sheet(0).list_sparkline_groups()),
                    len(wb.list_external_links()),
                    len(wb.sheet(0).get_cell_phonetic_runs(0, 0)),
                    len(wb.sheet(0).list_tables()),
                    len(wb.sheet(0).list_cells()),
                ],
                [0] * 8,
            )


class FormulonErrorStatusTests(unittest.TestCase):
    def test_paginate_null_result_raises_formulon_error_with_status(self) -> None:
        """The defensive null-handle path must not raise ValueError."""
        with Workbook.create_default() as wb:
            # Force `fm_sheet_paginate` to report success while leaving
            # the out-slot at 0, which is the branch under test.
            real_paginate = _c.LIB.fm_sheet_paginate

            def fake_paginate(handle, sheet, out):
                real_paginate(handle, sheet, out)
                _c.LIB.write_bytes(out, b"\x00\x00\x00\x00")
                return 0

            with mock.patch.object(_c.LIB, "fm_sheet_paginate", fake_paginate, create=True):
                with self.assertRaises(FormulonError) as caught:
                    wb.sheet(0).get_pagination()
        self.assertEqual(caught.exception.code, 7001)
        self.assertEqual(caught.exception.message, "returned OK with a null pagination handle")
        self.assertNotIn("fm_", str(caught.exception))


class ErrorOperationNameTests(unittest.TestCase):
    def test_count_helper_failure_carries_no_c_abi_name(self) -> None:
        """Messages are written for Python callers: no ABI symbol, no binding helper."""
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError) as caught:
                wb.sheet(99).list_merges()
        message = str(caught.exception)
        self.assertIsInstance(caught.exception, formulon.IndexOutOfRangeError)
        self.assertNotIn("fm_", message)
        self.assertNotIn("_wrapped", message)

    def test_trace_failure_carries_no_c_abi_name(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError) as caught:
                wb.sheet(99).get_precedents(0, 0)
        message = str(caught.exception)
        self.assertIsInstance(caught.exception, formulon.IndexOutOfRangeError)
        self.assertNotIn("fm_", message)
        self.assertNotIn("_wrapped", message)

    def test_exports_expose_the_c_abi_symbol_as_their_name(self) -> None:
        self.assertEqual(_c.LIB.fm_workbook_get_sheet_count.__name__, "fm_workbook_get_sheet_count")


class WasmLocatorTests(unittest.TestCase):
    def test_missing_wasm_error_lists_every_probed_candidate(self) -> None:
        env = dict(os.environ)
        env["FORMULON_WASM_PATH"] = "/nonexistent/override/formulon_capi.wasm"
        with mock.patch.dict(os.environ, env, clear=True):
            with mock.patch.object(_c.Path, "is_file", lambda self: False):
                with self.assertRaises(FileNotFoundError) as caught:
                    _c._locate_wasm()
        message = str(caught.exception)
        self.assertIn("_wasm/formulon_capi.wasm", message)
        self.assertIn("/nonexistent/override/formulon_capi.wasm", message)
        self.assertIn("FORMULON_WASM_PATH", message)

    def test_missing_wasm_error_omits_the_override_when_unset(self) -> None:
        env = {k: v for k, v in os.environ.items() if k != "FORMULON_WASM_PATH"}
        with mock.patch.dict(os.environ, env, clear=True):
            with mock.patch.object(_c.Path, "is_file", lambda self: False):
                with self.assertRaises(FileNotFoundError) as caught:
                    _c._locate_wasm()
        # The remediation sentence still names the variable; what must be
        # absent is an override entry in the probed-candidate list.
        self.assertNotIn("(from FORMULON_WASM_PATH)", str(caught.exception))


class CatalogProfileTests(unittest.TestCase):
    def test_unknown_function_returns_none(self) -> None:
        self.assertIsNone(get_function_metadata("NOPE"))

    def test_unknown_profile_raises_on_every_profile_keyed_call(self) -> None:
        for call in (
            lambda: localize_function_name("SUM", "nope"),
            lambda: canonicalize_function_name("SUM", "nope"),
            lambda: localize_formula("=SUM(1)", "nope"),
            lambda: canonicalize_formula("=SUM(1)", "nope"),
            lambda: get_locale_facts("nope"),
        ):
            with self.assertRaises(FormulonError):
                call()

    def test_formula_and_function_names_follow_the_profile(self) -> None:
        self.assertEqual(localize_formula("=SUM(1.5,2)", "mac-365-de_DE"), "=SUMME(1,5;2)")
        self.assertEqual(canonicalize_formula("=SUMME(1,5;2)", "mac-365-de_DE"), "=SUM(1.5,2)")
        self.assertEqual(localize_formula("=IF(TRUE,1,2)", "mac-365-fr_FR"), "=SI(VRAI;1;2)")
        self.assertEqual(localize_formula("=#N/A", "mac-365-de_DE"), "=#NV")
        self.assertEqual(canonicalize_formula("=#NV", "mac-365-de_DE"), "=#N/A")
        self.assertEqual(localize_formula('=A1&";"', "mac-365-de_DE"), '=A1&";"')
        self.assertEqual(localize_function_name("SUM", "win-365-de_DE"), "SUMME")
        self.assertEqual(canonicalize_function_name("SUMME", "win-365-de_DE"), "SUM")

    def test_locale_facts_report_separators_and_measurement(self) -> None:
        mac = get_locale_facts("mac-365-de_DE")
        self.assertEqual(mac.decimal_separator, ",")
        self.assertEqual(mac.list_separator, ";")
        self.assertEqual(mac.date_order, LocaleDateOrder.DMY)
        self.assertTrue(mac.measured)
        na = next(e for e in mac.error_names if e.canonical == "#N/A")
        self.assertEqual(na.localized, "#NV")
        self.assertFalse(get_locale_facts("win-365-de_DE").measured)
        self.assertEqual(get_locale_facts("win-365-en_US").decimal_separator, ".")


class NamedConstantTableTests(unittest.TestCase):
    """Ordinals crossing the ABI as plain numbers must be nameable here too.

    Both JS packages expose the same tables under the same names, held in
    step by the binding-drift check. This side asserts that the Python
    names exist, carry the documented ordinals, and are reachable from the
    package root.
    """

    def test_error_code_ordinals_match_the_engine_enum(self) -> None:
        self.assertEqual(
            [ErrorCode.NULL, ErrorCode.DIV0, ErrorCode.VALUE, ErrorCode.REF, ErrorCode.NAME, ErrorCode.NUM],
            [0, 1, 2, 3, 4, 5],
        )
        self.assertEqual(ErrorCode.NA, 6)
        self.assertEqual(ErrorCode.UNKNOWN, 16)

    def test_external_link_kind_ordinals(self) -> None:
        self.assertEqual(
            [
                ExternalLinkKind.UNKNOWN,
                ExternalLinkKind.EXTERNAL_BOOK,
                ExternalLinkKind.OLE,
                ExternalLinkKind.DDE,
            ],
            [0, 1, 2, 3],
        )

    def test_new_tables_are_exported_from_the_package_root(self) -> None:
        for name, table in (("ErrorCode", ErrorCode), ("ExternalLinkKind", ExternalLinkKind), ("CalcMode", CalcMode)):
            with self.subTest(name=name):
                self.assertIn(name, formulon.__all__)
                self.assertIs(getattr(formulon, name), table)

    def test_error_code_names_a_value_read_back_from_a_cell(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_error(0, 0, ErrorCode.DIV0)
            value = wb.sheet(0).get_value(0, 0)
        self.assertEqual(value, ExcelError(ErrorCode.DIV0))
        self.assertEqual(value.code, ErrorCode.DIV0)


class StyleAndCalcAccessorTests(unittest.TestCase):
    """Writers whose readers already existed, and vice versa."""

    @staticmethod
    def _xf() -> CellXfInput:
        return CellXfInput(
            font_index=0,
            fill_index=0,
            border_index=0,
            num_fmt_id=0,
            horizontal=0,
            vertical=2,
            wrap_text=False,
        )

    def test_named_style_xf_writer_round_trips_through_its_reader(self) -> None:
        with Workbook.create_default() as wb:
            style_xf_index = wb.ooxml.add_cell_style_xf(self._xf())
            read_back = wb.ooxml.get_cell_style_xf(style_xf_index)
        self.assertEqual(read_back.num_fmt_id, 0)
        self.assertEqual(read_back.vertical, 2)

    def test_named_style_xf_writer_deduplicates(self) -> None:
        with Workbook.create_default() as wb:
            first = wb.ooxml.add_cell_style_xf(self._xf())
            second = wb.ooxml.add_cell_style_xf(self._xf())
        self.assertEqual(first, second)

    def test_iterative_settings_read_back_what_was_set(self) -> None:
        with Workbook.create_default() as wb:
            wb.update_iterative(enabled=True, max_iterations=42, max_change=0.5)
            settings = wb.get_iterative()
        self.assertTrue(settings.enabled)
        self.assertEqual(settings.max_iterations, 42)
        self.assertEqual(settings.max_change, 0.5)

    def test_iterative_cap_and_threshold_survive_being_disabled(self) -> None:
        with Workbook.create_default() as wb:
            wb.update_iterative(enabled=True, max_iterations=42, max_change=0.5)
            wb.update_iterative(enabled=False, max_iterations=42, max_change=0.5)
            settings = wb.get_iterative()
        self.assertFalse(settings.enabled)
        self.assertEqual(settings.max_iterations, 42)
        self.assertEqual(settings.max_change, 0.5)

    def test_enable_toggle_keeps_the_cap_and_threshold(self) -> None:
        with Workbook.create_default() as wb:
            wb.update_iterative(enabled=True, max_iterations=42, max_change=0.5)
            wb.update_iterative(enabled=False)
            disabled = wb.get_iterative()
            wb.update_iterative(enabled=True)
            reenabled = wb.get_iterative()
        self.assertFalse(disabled.enabled)
        self.assertTrue(reenabled.enabled)
        self.assertEqual(reenabled.max_iterations, 42)
        self.assertEqual(reenabled.max_change, 0.5)


class StyleSeedingTests(unittest.TestCase):
    """Excel's reserved style slots survive whichever factory built the workbook."""

    def test_empty_workbook_style_records_append_after_the_reserved_slots(self) -> None:
        with Workbook.create_empty() as wb:
            wb.add_sheet("Sheet1")
            font_index = wb.ooxml.add_font(Font(name="Arial", size_pt=12.0))
            # Excel reserves fill 0 for `none` and fill 1 for `gray125`.
            fill_index = wb.ooxml.add_fill(Fill(pattern=1, foreground=Color(ColorKind.RGB, argb=0xFFFF0000)))
            border_index = wb.ooxml.add_border(
                Border(left=BorderSide(style=1, color=Color(ColorKind.RGB, argb=0xFF000000)))
            )
            xf_index = wb.ooxml.add_cell_xf(
                CellXfInput(
                    font_index=font_index,
                    fill_index=fill_index,
                    border_index=border_index,
                    num_fmt_id=0,
                    horizontal=0,
                    vertical=0,
                    wrap_text=False,
                )
            )
            data = wb.save()
        self.assertGreater(font_index, 0)
        self.assertGreater(fill_index, 1)
        self.assertGreater(border_index, 0)
        self.assertGreater(xf_index, 0)
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            styles_xml = archive.read("xl/styles.xml").decode("utf-8")
        self.assertIn("gray125", styles_xml)
        self.assertIn('name="Normal"', styles_xml)


class MemoryUsageTests(unittest.TestCase):
    def test_estimate_grows_with_stored_cells(self) -> None:
        with Workbook.create_default() as wb:
            empty = wb.get_memory_usage()
            for row in range(2000):
                wb.sheet(0).set_number(row, 0, float(row))
            filled = wb.get_memory_usage()
        self.assertGreater(filled, empty)

    def test_estimate_is_stable_across_repeated_calls(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 1.0)
            self.assertEqual(wb.get_memory_usage(), wb.get_memory_usage())

    def test_closed_handle_raises_rather_than_reporting_zero(self) -> None:
        wb = Workbook.create_default()
        wb.close()
        with self.assertRaises(FormulonError) as caught:
            wb.get_memory_usage()
        self.assertEqual(caught.exception.code, 7000)


class StructuredLogTests(unittest.TestCase):
    """The threshold is process-wide, so each test restores the default."""

    def setUp(self) -> None:
        self.addCleanup(formulon.set_log_min_level, LogLevel.OFF)

    @staticmethod
    def _save_xlsb_emitting_a_warning() -> str:
        """Drive a save that logs, capturing the engine's own stderr.

        The engine writes to file descriptor 2 from inside the WASM
        module, which ``contextlib.redirect_stderr`` cannot see, so the
        capture has to happen at the descriptor level.
        """
        with tempfile.TemporaryFile(mode="w+") as sink:
            saved = os.dup(2)
            try:
                sys.stderr.flush()
                os.dup2(sink.fileno(), 2)
                with Workbook.create_default() as wb:
                    # A structured reference the XLSB encoder
                    # cannot lower, which makes the writer emit a per-cell
                    # warn record.
                    wb.sheet(0).set_formula(0, 0, "=SUM(T[C])")
                    wb.recalc()
                    wb.save(format=formulon.WorkbookFormat.XLSB)
            finally:
                os.dup2(saved, 2)
                os.close(saved)
            sink.seek(0)
            return sink.read()

    def test_capture_sees_records_once_the_threshold_admits_them(self) -> None:
        """Guards the silence assertion below from passing vacuously."""
        formulon.set_log_min_level(LogLevel.WARN)
        self.assertIn("xlsb.writer.formula_downgraded", self._save_xlsb_emitting_a_warning())

    def test_default_threshold_is_off_and_emits_nothing(self) -> None:
        self.assertEqual(LogLevel.OFF, 4)
        self.assertEqual(self._save_xlsb_emitting_a_warning(), "")

    def test_warn_threshold_does_not_silence_the_per_cell_downgrade_record(self) -> None:
        formulon.set_log_min_level(LogLevel.ERROR)
        self.assertEqual(self._save_xlsb_emitting_a_warning(), "")

    def test_threshold_is_a_module_level_control_not_a_workbook_method(self) -> None:
        self.assertTrue(callable(formulon.set_log_min_level))
        self.assertFalse(hasattr(Workbook, "set_log_min_level"))

    def test_out_of_range_level_is_rejected(self) -> None:
        for level in (-1, 5, 99):
            with self.subTest(level=level):
                with self.assertRaises(FormulonError) as caught:
                    formulon.set_log_min_level(level)
                self.assertEqual(caught.exception.code, 2)

    def test_every_level_ordinal_is_accepted(self) -> None:
        for level in LogLevel:
            with self.subTest(level=level):
                formulon.set_log_min_level(level)


class WasmOnlyCapabilityTests(unittest.TestCase):
    """Capabilities the WASM npm package exposed before this wheel did."""

    def test_phonetic_guide_round_trips(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "日本語")
            self.assertIsNone(wb.sheet(0).get_cell_phonetic(0, 0))
            wb.sheet(0).set_cell_phonetic(0, 0, "ニホンゴ")
            self.assertEqual(wb.sheet(0).get_cell_phonetic(0, 0), "ニホンゴ")

    def test_phonetic_guide_survives_a_save_load_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "日本語")
            wb.sheet(0).set_cell_phonetic(0, 0, "ニホンゴ")
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertEqual(reloaded.sheet(0).get_cell_phonetic(0, 0), "ニホンゴ")

    def test_empty_phonetic_guide_clears_it(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "日本語")
            wb.sheet(0).set_cell_phonetic(0, 0, "ニホンゴ")
            wb.sheet(0).set_cell_phonetic(0, 0, "")
            self.assertIsNone(wb.sheet(0).get_cell_phonetic(0, 0))

    def test_phonetic_runs_keep_their_spans(self) -> None:
        runs = [PhoneticRun(0, 2, "トウキョウ"), PhoneticRun(2, 3, "ト")]
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "東京都")
            self.assertEqual(wb.sheet(0).get_cell_phonetic_runs(0, 0), [])
            wb.sheet(0).set_cell_phonetic_runs(0, 0, runs)
            self.assertEqual(wb.sheet(0).get_cell_phonetic_runs(0, 0), runs)
            # The flattening getter still reports the concatenation.
            self.assertEqual(wb.sheet(0).get_cell_phonetic(0, 0), "トウキョウト")
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertEqual(
                reloaded.sheet(0).get_cell_phonetic_runs(0, 0),
                runs,
            )

    def test_phonetic_run_with_empty_text_round_trips(self) -> None:
        # Excel itself produces a run with `text == ""` for an out-of-range
        # kana slice (see src/io/xlsb/sst_reader.cpp); the C ABI rejects a NULL
        # run text but accepts a zero-length one, so this must not raise.
        runs = [PhoneticRun(0, 2, "トウキョウ"), PhoneticRun(2, 3, "")]
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "東京都")
            wb.sheet(0).set_cell_phonetic_runs(0, 0, runs)
            self.assertEqual(wb.sheet(0).get_cell_phonetic_runs(0, 0), runs)

    def test_whole_cell_setter_collapses_the_spans(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "東京都")
            wb.sheet(0).set_cell_phonetic_runs(0, 0, [PhoneticRun(0, 2, "トウキョウ"), PhoneticRun(2, 3, "ト")])
            wb.sheet(0).set_cell_phonetic(0, 0, wb.sheet(0).get_cell_phonetic(0, 0))
            self.assertEqual(
                wb.sheet(0).get_cell_phonetic_runs(0, 0),
                [PhoneticRun(0, 3, "トウキョウト")],
            )

    def test_empty_run_list_clears_the_guide(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "東京都")
            wb.sheet(0).set_cell_phonetic_runs(0, 0, [PhoneticRun(0, 3, "トウキョウト")])
            wb.sheet(0).set_cell_phonetic_runs(0, 0, [])
            self.assertEqual(wb.sheet(0).get_cell_phonetic_runs(0, 0), [])
            self.assertIsNone(wb.sheet(0).get_cell_phonetic(0, 0))

    def test_out_of_order_runs_are_rejected(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "東京都")
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_cell_phonetic_runs(0, 0, [PhoneticRun(2, 3, "ト"), PhoneticRun(0, 2, "トウ")])

    def test_phonetic_properties_are_independent_of_the_runs(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "大阪")
            # An unannotated cell reports what Excel infers for a guide
            # written with no <phoneticPr> element at all.
            self.assertEqual(wb.sheet(0).get_cell_phonetic_properties(0, 0), PhoneticProperties(0, 0, 0))

            wb.sheet(0).set_cell_phonetic_properties(0, 0, PhoneticProperties(font_index=3, type=2, alignment=2))
            # Setting the readings must not reset the rendering.
            wb.sheet(0).set_cell_phonetic_runs(0, 0, [PhoneticRun(0, 2, "おおさか")])
            self.assertEqual(wb.sheet(0).get_cell_phonetic_properties(0, 0), PhoneticProperties(3, 2, 2, stated=True))
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertEqual(
                reloaded.sheet(0).get_cell_phonetic_properties(0, 0), PhoneticProperties(3, 2, 2, stated=True)
            )

    def test_phonetic_properties_reject_values_the_container_cannot_hold(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "大阪")
            wb.sheet(0).set_cell_phonetic_properties(0, 0, PhoneticProperties(font_index=1, type=3, alignment=3))
            # Two bits each on the binary side, so a wider ordinal is
            # refused rather than truncated into a valid-looking one.
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_cell_phonetic_properties(0, 0, PhoneticProperties(font_index=0, type=4, alignment=0))
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_cell_phonetic_properties(0, 0, PhoneticProperties(font_index=0, type=0, alignment=4))
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_cell_phonetic_properties(
                    0, 0, PhoneticProperties(font_index=0x10000, type=0, alignment=0)
                )
            self.assertEqual(wb.sheet(0).get_cell_phonetic_properties(0, 0), PhoneticProperties(1, 3, 3, stated=True))

    def test_a_value_write_discards_the_guide_and_its_rendering(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "大阪")
            wb.sheet(0).set_cell_phonetic_runs(0, 0, [PhoneticRun(0, 2, "おおさか")])
            wb.sheet(0).set_cell_phonetic_properties(0, 0, PhoneticProperties(font_index=3, type=2, alignment=2))
            wb.sheet(0).set_text(0, 0, "京都")
            self.assertEqual(wb.sheet(0).get_cell_phonetic_runs(0, 0), [])
            self.assertEqual(wb.sheet(0).get_cell_phonetic_properties(0, 0), PhoneticProperties(0, 0, 0))

    def test_default_font_replaces_what_an_unstyled_cell_saves_as(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.ooxml.get_font(0).name, "Calibri")
            # add_font can only ever append beside the seeded default.
            self.assertGreater(wb.ooxml.add_font(Font(name="游ゴシック", size_pt=11.0)), 0)
            self.assertEqual(wb.ooxml.get_font(0).name, "Calibri")

            wb.set_default_font(Font(name="游ゴシック", size_pt=11.0, charset=128))
            self.assertEqual(wb.ooxml.get_font(0).name, "游ゴシック")
            self.assertEqual(wb.ooxml.get_font(0).charset, 128)
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertEqual(reloaded.ooxml.get_font(0).name, "游ゴシック")

    def test_set_font_overwrites_an_existing_slot(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.ooxml.add_font(Font(name="Meiryo", size_pt=12.0))
            wb.ooxml.set_font(index, Font(name="MS Gothic", size_pt=9.0))
            self.assertEqual(wb.ooxml.get_font(index).name, "MS Gothic")
            self.assertEqual(len(wb.ooxml.list_fonts()), index + 1)
            with self.assertRaises(FormulonError):
                wb.ooxml.set_font(len(wb.ooxml.list_fonts()), Font(name="MS Gothic"))

    def test_table_create_update_remove(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            ws.set_text(0, 0, "A")
            ws.set_text(0, 1, "B")
            table = ws.add_table(TableInput(range=CellRect(0, 0, 2, 1), name="Table1", columns=["A", "B"]))
            self.assertEqual(
                [t.get_properties().range for t in ws.list_tables()],
                [SheetRect(0, CellRect(0, 0, 2, 1))],
            )
            self.assertEqual([t.name for t in ws.list_tables()], ["Table1"])

            table.update(range=CellRect(0, 0, 4, 1))
            self.assertEqual(
                [t.get_properties().range for t in ws.list_tables()],
                [SheetRect(0, CellRect(0, 0, 4, 1))],
            )

            ws.remove_table(table)
            self.assertEqual(ws.list_tables(), [])

    def test_table_style_toggles_survive_a_style_name_change(self) -> None:
        with Workbook.create_default() as wb:
            table = wb.sheet(0).add_table(
                TableInput(
                    range=CellRect(0, 0, 2, 1),
                    name="Sales",
                    columns=["A", "B"],
                    style_name="TableStyleLight9",
                )
            )
            style = table.get_style_options()
            self.assertEqual(style.style_name, "TableStyleLight9")
            self.assertTrue(style.show_row_stripes)
            self.assertFalse(style.show_first_column)
            properties = table.get_properties()
            self.assertEqual((properties.header_row, properties.totals_row), (True, False))
            self.assertEqual([c.name for c in properties.columns], ["A", "B"])

            table.update_style_options(show_first_column=True, show_row_stripes=False)
            table.update_style_options(style_name="TableStyleMedium4")
            with Workbook.load(wb.save()) as reloaded:
                style = reloaded.sheet(0).list_tables()[0].get_style_options()
                self.assertEqual(style.style_name, "TableStyleMedium4")
                self.assertTrue(style.show_first_column)
                self.assertFalse(style.show_row_stripes)

    def test_table_without_a_style_name_takes_the_default_style(self) -> None:
        with Workbook.create_default() as wb:
            ws = wb.sheet(0)
            plain = ws.add_table(TableInput(range=CellRect(0, 0, 2, 1), name="Plain", columns=["A", "B"]))
            options = plain.get_style_options()
            self.assertEqual(options.style_name, "TableStyleMedium2")
            self.assertTrue(options.show_row_stripes)
            self.assertFalse(options.show_first_column)

            bare = ws.add_table(TableInput(range=CellRect(0, 3, 2, 4), name="Bare", columns=["C", "D"], styled=False))
            self.assertIsNone(bare.get_style_options())

            wb.update_table_style_defaults(table_style_name="TableStyleLight9")
            styled = ws.add_table(TableInput(range=CellRect(5, 0, 7, 1), name="Styled", columns=["E", "F"]))
            self.assertEqual(styled.get_style_options().style_name, "TableStyleLight9")

    def test_table_column_totals_write_the_totals_row(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "Product")
            wb.sheet(0).set_text(0, 1, "Amount")
            wb.sheet(0).set_number(1, 1, 10.0)
            wb.sheet(0).set_number(2, 1, 20.0)
            table = wb.sheet(0).add_table(
                TableInput(
                    range=CellRect(0, 0, 3, 1),
                    name="Sales",
                    columns=["Product", "Amount"],
                    totals_row=True,
                )
            )
            table.set_column_totals(0, TableColumnTotalsInput(totals_row_label="Total"))
            table.set_column_totals(1, TableColumnTotalsInput(totals_row_function=TableTotalsRowFunction.SUM))
            self.assertEqual(wb.sheet(0).get_formula(3, 1), "=SUBTOTAL(109,[Amount])")
            wb.recalc()
            self.assertEqual(wb.sheet(0).get_value(3, 1), 30.0)
            with self.assertRaises(FormulonError):
                table.set_column_totals(1, TableColumnTotalsInput(totals_row_function=99))  # type: ignore[arg-type]
            with self.assertRaises(TypeError):
                table.set_column_totals(1, TableColumnTotalsInput(totals_row_function="sum"))  # type: ignore[arg-type]
            with Workbook.load(wb.save()) as reloaded:
                columns = reloaded.sheet(0).list_tables()[0].get_properties().columns
                self.assertEqual(columns[0].totals_row_label, "Total")
                self.assertEqual(columns[1].totals_row_function, TableTotalsRowFunction.SUM)

    def test_table_rename_rewrites_structured_references(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "Amount")
            wb.sheet(0).set_number(1, 0, 5.0)
            table = wb.sheet(0).add_table(TableInput(range=CellRect(0, 0, 1, 0), name="Sales", columns=["Amount"]))
            wb.sheet(0).set_formula(0, 3, "=SUM(Sales[Amount])")
            table.rename("Revenue")
            self.assertEqual(wb.sheet(0).get_formula(0, 3), "=SUM(Revenue[Amount])")
            self.assertEqual(wb.sheet(0).list_tables()[0].name, "Revenue")
            other = wb.sheet(0).add_table(TableInput(range=CellRect(0, 2, 1, 2), name="Other", columns=["X"]))
            with self.assertRaises(FormulonError):
                other.rename("revenue")

    def test_custom_table_style_round_trips(self) -> None:
        with Workbook.create_default() as wb:
            dxf = wb.ooxml.add_dxf(DxfInput())
            style = TableStyle(
                name="Corporate",
                pivot=False,
                elements=[
                    TableStyleElement(TableStyleElementType.WHOLE_TABLE, dxf_index=dxf),
                    TableStyleElement(TableStyleElementType.FIRST_ROW_STRIPE, size=2, dxf_index=dxf),
                    TableStyleElement(TableStyleElementType.TOTAL_ROW),
                ],
            )
            wb.set_table_style(style)
            self.assertEqual(len(wb.list_table_styles()), 1)
            self.assertEqual(
                wb.get_table_style_defaults(),
                TableStyleDefaults(table_style_name="TableStyleMedium2", pivot_style_name="PivotStyleLight16"),
            )
            table = wb.sheet(0).add_table(
                TableInput(range=CellRect(0, 0, 1, 0), name="T", columns=["X"], style_name="Corporate")
            )
            with self.assertRaises(FormulonError):
                wb.set_table_style(TableStyle(name="TableStyleMedium2"))
            with Workbook.load(wb.save()) as reloaded:
                self.assertEqual(reloaded.list_table_styles(), [style])
                self.assertEqual(reloaded.get_table_style("corporate"), style)
                self.assertIsNone(reloaded.get_table_style("Other"))
                self.assertEqual(
                    reloaded.sheet(0).list_tables()[table.index].get_style_options().style_name, "Corporate"
                )
                reloaded.update_table_style_defaults(table_style_name="Corporate")
                self.assertEqual(reloaded.get_table_style_defaults().table_style_name, "Corporate")
                reloaded.remove_table_style("corporate")
                self.assertEqual(reloaded.list_table_styles(), [])
                self.assertEqual(reloaded.get_table_style_defaults().table_style_name, "TableStyleMedium2")

    def test_table_update_none_style_name_removes_the_name_and_keeps_the_toggles(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "A")
            wb.sheet(0).set_text(0, 1, "B")
            table = wb.sheet(0).add_table(
                TableInput(
                    range=CellRect(0, 0, 2, 1),
                    name="Sales",
                    columns=["A", "B"],
                    style_name="TableStyleMedium2",
                )
            )
            table.update(range=CellRect(0, 0, 2, 1))
            self.assertEqual(table.get_style_options().style_name, "TableStyleMedium2")

            table.update_style_options(style_name=None)
            options = table.get_style_options()
            self.assertIsNone(options.style_name)
            self.assertTrue(options.show_row_stripes)
            data = wb.save()
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            parts = [name for name in archive.namelist() if name.startswith("xl/tables/")]
            self.assertEqual(len(parts), 1)
            self.assertIn(b"tableStyleInfo", archive.read(parts[0]))
        with Workbook.load(data) as reloaded:
            self.assertIsNone(reloaded.sheet(0).list_tables()[0].get_style_options().style_name)

    def test_table_create_rejects_a_column_count_that_does_not_match_ref(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.sheet(0).add_table(TableInput(range=CellRect(0, 0, 2, 1), name="Table1", columns=["OnlyOne"]))

    def test_auto_filter_xml_round_trips(self) -> None:
        fragment = '<autoFilter ref="A1:B3"/>'
        with Workbook.create_default() as wb:
            self.assertIsNone(wb.sheet(0).ooxml.get_auto_filter_xml())
            wb.sheet(0).ooxml.set_auto_filter_xml(fragment)
            self.assertEqual(wb.sheet(0).ooxml.get_auto_filter_xml(), fragment)
            wb.sheet(0).ooxml.set_auto_filter_xml("")
            self.assertIsNone(wb.sheet(0).ooxml.get_auto_filter_xml())

    def test_named_cell_style_registration_reaches_the_reader(self) -> None:
        with Workbook.create_default() as wb:
            style_xf_index = wb.ooxml.add_cell_style_xf(StyleAndCalcAccessorTests._xf())
            before = len(wb.list_cell_styles())
            wb.set_cell_style(CellStyle("MyStyle", style_xf_index))
            names = [style.name for style in wb.list_cell_styles()]
        self.assertGreater(len(names), before)
        self.assertIn("MyStyle", names)

    def test_named_cell_style_rejects_an_unregistered_xf_id(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.set_cell_style(CellStyle("Bogus", 9999))


class OneShotEvaluationTests(unittest.TestCase):
    """``evaluate_formula`` must agree value-for-value with the JS surfaces.

    Both npm packages assert the same table. The anchor-referencing cases
    are the ones that diverge if a surface writes the formula into A1 and
    recalcs instead of evaluating read-only.
    """

    CASES = (
        ("=1+2", ValueKind.NUMBER, 3.0),
        ("=A1", ValueKind.NUMBER, 0.0),
        ("=COUNTA(A1)", ValueKind.NUMBER, 0.0),
        ("=ISBLANK(A1)", ValueKind.BOOL, True),
        ("=SUM(A1:A3)", ValueKind.NUMBER, 0.0),
        ("=SEQUENCE(3)", ValueKind.ARRAY, [[1.0], [2.0], [3.0]]),
        ("=ROWS(SEQUENCE(3))", ValueKind.NUMBER, 3.0),
    )

    def test_values_match_the_shared_one_shot_table(self) -> None:
        for formula, kind, payload in self.CASES:
            with self.subTest(formula=formula):
                self.assertEqual(formulon.evaluate_formula(formula), payload)
                value = formulon.evaluate_typed_formula(formula)
                self.assertEqual(value.kind, kind)
                self.assertEqual(formulon.to_native(value), payload)
                # Writing into the anchor and recalcing would make the
                # anchor-referencing cases #REF!.
                self.assertIsNone(value.error_code)

    def test_the_anchor_is_left_blank_rather_than_written(self) -> None:
        """The distinguishing observation between the two semantics."""
        self.assertIs(formulon.evaluate_formula("=ISBLANK(A1)"), True)

    def test_array_results_keep_their_shape(self) -> None:
        self.assertEqual(formulon.evaluate_formula("=SEQUENCE(3)"), [[1.0], [2.0], [3.0]])
        self.assertEqual(formulon.evaluate_formula("=SEQUENCE(2,3)"), [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]])

    def test_a_one_by_one_result_collapses_and_larger_ones_return_every_cell(self) -> None:
        self.assertEqual(formulon.evaluate_formula("=1+2"), 3.0)
        big = formulon.evaluate_formula("=SEQUENCE(100,100,7)")
        self.assertEqual((len(big), len(big[0])), (100, 100))
        self.assertEqual((big[0][0], big[99][99]), (7.0, 10006.0))


class PivotFilterPersistenceTests(unittest.TestCase):
    """The filter list is saved with the pivot and reads back after a load.

    The npm and npm-native packages assert the same three properties, so a
    surface that stopped persisting (or started dropping) filters would break
    here and there together.
    """

    @staticmethod
    def _workbook_with_pivot():
        wb = Workbook.create_default()
        cache_id = wb.add_pivot_cache().cache_id
        # A cache with no declared source cannot be saved: Excel offers to
        # repair any package containing one, so the writer refuses instead.
        wb.pivot_cache(cache_id).set_worksheet_source(formulon.PivotWorksheetSourceInput(range="A1:B3", sheet="Sheet1"))
        wb.pivot_cache(cache_id).add_field("Region")
        wb.pivot_cache(cache_id).add_field("Amount")
        for region, amount in (("East", 10.0), ("West", 30.0)):
            rec = wb.pivot_cache(cache_id).add_record()
            wb.pivot_cache(cache_id).set_record_value(rec, 0, region)
            wb.pivot_cache(cache_id).set_record_value(rec, 1, amount)
        pivot = wb.sheet(0).add_pivot(PivotInput(name="Pivot1", cache_id=cache_id, anchor=CellAddress(0, 4)))
        wb.sheet(0).pivot(pivot).set_field_axis("Region", PivotAxis.ROW)
        wb.sheet(0).pivot(pivot).set_field_axis("Amount", PivotAxis.VALUE)
        amount_field = wb.sheet(0).pivot(pivot).get_field("Amount").cache_field_index
        wb.sheet(0).pivot(pivot).add_data_field(
            PivotDataFieldInput(name="Sum of Amount", field_index=amount_field, aggregation=PivotAggregation.SUM)
        )
        return wb, pivot

    @staticmethod
    def _filter() -> PivotFilterInput:
        return PivotFilterInput(
            field_index="Region",
            kind=PivotFilterKind.VALUE_GREATER_THAN,
            value=15.0,
        )

    def test_added_filter_reads_back_field_for_field(self) -> None:
        wb, pivot = self._workbook_with_pivot()
        try:
            wb.sheet(0).pivot(pivot).add_filter(self._filter())
            self.assertEqual(len(wb.sheet(0).pivot(pivot).list_filters()), 1)
            got = wb.sheet(0).pivot(pivot).list_filters()[0]
        finally:
            wb.close()
        self.assertEqual(got.field_index, 0)
        self.assertEqual(got.kind, PivotFilterKind.VALUE_GREATER_THAN)
        self.assertEqual(got.value, 15.0)

    def test_count_reports_the_filters_on_the_pivot(self) -> None:
        wb, pivot = self._workbook_with_pivot()
        try:
            self.assertEqual(len(wb.sheet(0).pivot(pivot).list_filters()), 0)
            wb.sheet(0).pivot(pivot).add_filter(self._filter())
            self.assertEqual(len(wb.sheet(0).pivot(pivot).list_filters()), 1)
        finally:
            wb.close()

    def test_filters_survive_save_and_load(self) -> None:
        """The list is written by save() and decodes back on load."""
        wb, pivot = self._workbook_with_pivot()
        try:
            wb.sheet(0).pivot(pivot).add_filter(self._filter())
            self.assertEqual(len(wb.sheet(0).pivot(pivot).list_filters()), 1)
            data = wb.save()
        finally:
            wb.close()
        with Workbook.load(data) as reloaded:
            self.assertEqual(len(reloaded.sheet(0).list_pivots()), 1)
            self.assertEqual(len(reloaded.sheet(0).pivot(0).list_filters()), 1)
            got = reloaded.sheet(0).pivot(0).list_filters()[0]
            self.assertEqual((got.field_index, got.kind), (0, PivotFilterKind.VALUE_GREATER_THAN))
            self.assertEqual(got.value, 15.0)

    def test_out_of_range_index_is_rejected(self) -> None:
        wb, pivot = self._workbook_with_pivot()
        try:
            with self.assertRaises(IndexError):
                wb.sheet(0).pivot(pivot).list_filters()[99]
        finally:
            wb.close()


class PivotDefinitionReadBackTests(unittest.TestCase):
    """The definition read-backs report persisted state, unlike the filter list."""

    def test_definition_survives_save_and_load(self) -> None:
        wb, pivot = PivotFilterPersistenceTests._workbook_with_pivot()
        try:
            wb.sheet(0).pivot(pivot).update_grand_totals(rows=True, columns=False)
            cache_id = wb.sheet(0).pivot(pivot).cache().cache_id
            data = wb.save()
        finally:
            wb.close()
        with Workbook.load(data) as reloaded:
            self.assertEqual(reloaded.sheet(0).pivot(0).cache().cache_id, cache_id)
            self.assertEqual(reloaded.sheet(0).pivot(0).name, "Pivot1")
            self.assertEqual(reloaded.sheet(0).pivot(0).get_grand_totals(), PivotGrandTotals(rows=True, columns=False))
            self.assertEqual(reloaded.sheet(0).pivot(0).get_row_field_order(), [0])
            self.assertEqual(reloaded.sheet(0).pivot(0).list_fields()[0].axis, PivotAxis.ROW)
            self.assertEqual(reloaded.sheet(0).pivot(0).list_data_fields()[0].name, "Sum of Amount")

    def test_out_of_range_index_is_rejected(self) -> None:
        wb, pivot = PivotFilterPersistenceTests._workbook_with_pivot()
        try:
            with self.assertRaises(IndexError):
                wb.sheet(0).pivot(pivot).list_fields()[99]
            with self.assertRaises(IndexError):
                wb.sheet(0).pivot(pivot).list_data_fields()[99]
        finally:
            wb.close()


class WorkbookViewAndPhoneticFlagTests(unittest.TestCase):
    def test_active_sheet_is_recorded_and_survives_a_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            wb.add_sheet("B")
            wb.add_sheet("C")
            self.assertEqual(wb.active_sheet().index, 0)
            wb.set_active_sheet(2)
            self.assertEqual(wb.active_sheet().index, 2)
            with self.assertRaises(FormulonError):
                wb.set_active_sheet(9)
            self.assertEqual(wb.active_sheet().index, 2)
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertEqual(reloaded.active_sheet().index, 2)
            reloaded.sheet(2).move(0)
            self.assertEqual(reloaded.active_sheet().index, 0)

    def test_show_phonetic_survives_value_writes_and_a_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            self.assertFalse(wb.sheet(0).get_cell_show_phonetic(0, 0))
            wb.sheet(0).set_text(0, 0, "漢字")
            wb.sheet(0).set_cell_show_phonetic(0, 0, True)
            wb.sheet(0).set_text(0, 0, "漢字")
            self.assertTrue(wb.sheet(0).get_cell_show_phonetic(0, 0))
            with self.assertRaises(FormulonError):
                wb.sheet(9).set_cell_show_phonetic(0, 0, True)
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertTrue(reloaded.sheet(0).get_cell_show_phonetic(0, 0))
            reloaded.sheet(0).set_cell_show_phonetic(0, 0, False)
            self.assertFalse(reloaded.sheet(0).get_cell_show_phonetic(0, 0))

    def test_explicit_all_zero_phonetic_pr_is_stated_and_round_trips(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "漢字")
            wb.sheet(0).set_cell_phonetic_runs(0, 0, [PhoneticRun(0, 2, "かんじ")])
            self.assertFalse(wb.sheet(0).get_cell_phonetic_properties(0, 0).stated)
            wb.sheet(0).set_cell_phonetic_properties(0, 0, PhoneticProperties(font_index=0, type=0, alignment=0))
            self.assertTrue(wb.sheet(0).get_cell_phonetic_properties(0, 0).stated)
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertTrue(reloaded.sheet(0).get_cell_phonetic_properties(0, 0).stated)
            self.assertEqual(
                reloaded.sheet(0).get_cell_phonetic_properties(0, 0), PhoneticProperties(0, 0, 0, stated=True)
            )
            with self.assertRaises(FormulonError):
                reloaded.sheet(9).get_cell_phonetic_properties(0, 0).stated


class SheetViewOriginAndSelectionTests(unittest.TestCase):
    def test_scroll_origin_pane_origin_and_selection_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).update_view(freeze_rows=2, freeze_cols=1)
            view = wb.sheet(0).get_view()
            self.assertEqual(view.pane_top_left, CellAddress(2, 1))
            wb.sheet(0).update_view(top_left=CellAddress(5, 0))
            wb.sheet(0).update_view(pane_top_left=CellAddress(40, 6))
            with self.assertRaises(FormulonError):
                wb.sheet(0).update_view(pane_top_left=CellAddress(1, 6))
            ranges = [CellRect(41, 7, 43, 8), CellRect(50, 9, 50, 9)]
            wb.sheet(0).set_selection(SheetSelectionInput(CellAddress(41, 7), ranges=list(ranges)))
            with self.assertRaises(FormulonError):
                wb.sheet(0).set_selection(SheetSelectionInput(CellAddress(0, 0), ranges=[CellRect(3, 0, 1, 0)]))
            data = wb.save()
        with Workbook.load(data) as reloaded:
            view = reloaded.sheet(0).get_view()
            self.assertEqual(
                (view.top_left, view.pane_top_left),
                (CellAddress(5, 0), CellAddress(40, 6)),
            )
            self.assertEqual(
                reloaded.sheet(0).get_selection(), SheetSelection(active=CellAddress(41, 7), ranges=ranges)
            )

    def test_split_reads_back_round_trips_and_exchanges_with_freeze(self) -> None:
        with Workbook.create_default() as wb:
            view = wb.sheet(0).get_view()
            self.assertEqual((view.split_x_twips, view.split_y_twips), (0.0, 0.0))
            wb.sheet(0).update_view(split_x_twips=2000, split_y_twips=1500)
            view = wb.sheet(0).get_view()
            self.assertEqual(
                (view.split_x_twips, view.split_y_twips, view.freeze_rows, view.freeze_cols), (2000.0, 1500.0, 0, 0)
            )
            data = wb.save()
            wb.sheet(0).update_view(freeze_rows=2, freeze_cols=1)
            view = wb.sheet(0).get_view()
            self.assertEqual((view.split_x_twips, view.split_y_twips), (0.0, 0.0))
            wb.sheet(0).update_view(split_x_twips=0, split_y_twips=900)
            view = wb.sheet(0).get_view()
            self.assertEqual((view.split_y_twips, view.freeze_rows, view.freeze_cols), (900.0, 0, 0))
            for bad in ((-1, 0), (0, float("nan")), (float("inf"), 0)):
                with self.assertRaises(FormulonError):
                    wb.sheet(0).update_view(split_x_twips=bad[0], split_y_twips=bad[1])
            self.assertEqual(wb.sheet(0).get_view().split_y_twips, 900.0)
            wb.sheet(0).update_view(split_x_twips=0, split_y_twips=0)
            self.assertEqual(wb.sheet(0).get_view().split_y_twips, 0.0)
        with Workbook.load(data) as reloaded:
            view = reloaded.sheet(0).get_view()
            self.assertEqual(
                (view.split_x_twips, view.split_y_twips, view.freeze_rows, view.freeze_cols), (2000.0, 1500.0, 0, 0)
            )

    def test_unset_selection_reads_as_a1(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(
                wb.sheet(0).get_selection(), SheetSelection(active=CellAddress(0, 0), ranges=[CellRect(0, 0, 0, 0)])
            )
            with self.assertRaises(FormulonError):
                wb.sheet(9).get_selection()

    def test_selection_without_ranges_is_the_active_cell(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_selection(SheetSelectionInput(CellAddress(4, 2)))
            self.assertEqual(
                wb.sheet(0).get_selection(), SheetSelection(active=CellAddress(4, 2), ranges=[CellRect(4, 2, 4, 2)])
            )
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertEqual(
                reloaded.sheet(0).get_selection(),
                SheetSelection(active=CellAddress(4, 2), ranges=[CellRect(4, 2, 4, 2)]),
            )


class CollectionReaderAllocationTests(unittest.TestCase):
    """Reading N items must cost O(1) WASM malloc/free, not O(N)."""

    def _count_allocations(self, call) -> tuple[object, int, int]:
        counts = {"alloc": 0, "free": 0}
        real_alloc, real_free = _c.LIB.alloc, _c.LIB.free

        def counting_alloc(size):
            counts["alloc"] += 1
            return real_alloc(size)

        def counting_free(ptr):
            counts["free"] += 1
            return real_free(ptr)

        with mock.patch.object(_c.LIB, "alloc", counting_alloc):
            with mock.patch.object(_c.LIB, "free", counting_free):
                result = call()
        return result, counts["alloc"], counts["free"]

    def _assert_flat_in_item_count(self, populate, read) -> None:
        small = Workbook.create_default()
        large = Workbook.create_default()
        try:
            populate(small, 2)
            populate(large, 40)
            small_out, small_alloc, small_free = self._count_allocations(lambda: read(small))
            large_out, large_alloc, large_free = self._count_allocations(lambda: read(large))
            self.assertEqual(len(small_out), 2)
            self.assertEqual(len(large_out), 40)
            self.assertEqual(small_alloc, large_alloc)
            self.assertEqual(small_free, large_free)
            self.assertEqual(small_alloc, small_free)
        finally:
            small.close()
            large.close()

    def test_get_merges_allocation_count_is_flat(self) -> None:
        self._assert_flat_in_item_count(
            lambda wb, n: [wb.sheet(0).add_merge(CellRect(r, 4, r, 6)) for r in range(n)],
            lambda wb: wb.sheet(0).list_merges(),
        )

    def test_get_hyperlinks_allocation_count_is_flat(self) -> None:
        self._assert_flat_in_item_count(
            lambda wb, n: [
                wb.sheet(0).set_hyperlink(r, 8, formulon.HyperlinkInput(target="https://example.com/"))
                for r in range(n)
            ],
            lambda wb: wb.sheet(0).list_hyperlinks(),
        )

    def test_get_comments_allocation_count_is_flat(self) -> None:
        self._assert_flat_in_item_count(
            lambda wb, n: [wb.sheet(0).set_comment(r, 10, Comment("text", author="author")) for r in range(n)],
            lambda wb: wb.sheet(0).list_comments(),
        )

    def test_list_defined_names_allocation_count_is_flat(self) -> None:
        self._assert_flat_in_item_count(
            lambda wb, n: [wb.set_defined_name(DefinedNameInput(f"Name{r}", "=1")) for r in range(n)],
            lambda wb: wb.list_defined_names(),
        )

    def test_get_sheet_row_overrides_allocation_count_is_flat(self) -> None:
        self._assert_flat_in_item_count(
            lambda wb, n: [wb.sheet(0).set_row_height(r, r, 15.0) for r in range(n)],
            lambda wb: wb.sheet(0).list_row_overrides(),
        )

    def test_precedents_allocation_count_is_flat(self) -> None:
        def populate(wb, n):
            for r in range(n):
                wb.sheet(0).set_number(r + 1, 1, float(r))
            wb.sheet(0).set_formula(0, 0, f"=SUM(B2:B{n + 1})")
            wb.recalc()

        self._assert_flat_in_item_count(populate, lambda wb: wb.sheet(0).get_precedents(0, 0))

    def test_evaluate_cf_range_allocation_count_is_flat(self) -> None:
        def populate(wb, n):
            dxf_index = wb.ooxml.add_dxf(DxfInput())
            for r in range(n):
                wb.sheet(0).set_number(r, 0, float(r + 100))
            wb.sheet(0).add_cf_rule(
                CfCellIsRuleInput(
                    ranges=[CellRect(0, 0, n - 1, 0)],
                    operator=ComparisonOperator.GREATER_THAN_OR_EQUAL,
                    formula1="10",
                    dxf_index=dxf_index,
                ),
            )
            wb.recalc()

        self._assert_flat_in_item_count(populate, lambda wb: wb.sheet(0).evaluate_cf_range(CellRect(0, 0, 39, 0)))


class AbiScratchBlockTests(unittest.TestCase):
    """The records the wrapper sizes by hand, measured against the real module.

    ``fm_value_t`` and ``fm_cell_rect_t`` have no ``_structs.Struct``
    layout: the wrapper carries their sizes as the literals
    ``_c.fm_value_t_size`` and the 16 in ``Workbook.paginate``. The static
    drift check compares those literals to the C header in the working tree,
    which says nothing about the ``.wasm`` actually loaded -- a module built
    from a wider definition of either record writes past the block the
    wrapper allocated, corrupting whatever the WASM allocator placed after
    it. Padding each block with a sentinel makes that overrun observable
    here instead of as an unrelated failure later in the session.
    """

    _SENTINEL = b"\xa5" * 16

    def _guarded_block(self, size: int) -> int:
        ptr = _c.LIB.alloc(size + len(self._SENTINEL))
        _c.LIB.write_bytes(ptr, b"\x00" * size + self._SENTINEL)
        return ptr

    def _assert_sentinel_intact(self, ptr: int, size: int, record: str) -> None:
        tail = _c.LIB.read_bytes(ptr + size, len(self._SENTINEL))
        self.assertEqual(
            tail,
            self._SENTINEL,
            f"the engine wrote past {size} bytes for a {record}; the wrapper's block is too small for the "
            f"module it loaded (tail: {tail!r})",
        )

    def test_value_out_param_stays_within_the_wrappers_block(self) -> None:
        size = _c.fm_value_t_size
        with Workbook.create_default() as wb:
            wb.sheet(0).set_number(0, 0, 6.25)
            wb.recalc()
            ptr = self._guarded_block(size)
            try:
                status = _c.LIB.fm_sheet_get_value(wb._require(), 0, 0, 0, ptr)
                self.assertEqual(status, 0)
                self._assert_sentinel_intact(ptr, size, "fm_value_t")
                # The payload must also land where `Value._from_wasm` looks
                # for it, which pins the union's offset rather than only the
                # record's total width.
                raw = _c.LIB.read_bytes(ptr, size)
                self.assertEqual(struct.unpack_from("<i", raw, 0)[0], int(ValueKind.NUMBER))
                self.assertEqual(struct.unpack_from("<d", raw, 8)[0], 6.25)
            finally:
                _c.LIB.free(ptr)

    def test_print_range_out_param_stays_within_the_wrappers_block(self) -> None:
        size = 16  # fm_cell_rect_t: four uint32_t, as Workbook.paginate reads it
        with Workbook.create_default() as wb:
            wb.set_defined_name(DefinedNameInput("_xlnm.Print_Area", "Sheet1!$A$1:$C$5", 0))
            out = _c.LIB.alloc(4)
            try:
                self.assertEqual(_c.LIB.fm_sheet_paginate(wb._require(), 0, out), 0)
                pagination = _c.LIB.read_u32(out)
                self.assertNotEqual(pagination, 0)
                try:
                    count = _c.LIB.alloc(4)
                    try:
                        self.assertEqual(_c.LIB.fm_pagination_get_print_area_count(pagination, count), 0)
                        self.assertGreater(_c.LIB.read_u32(count), 0)
                    finally:
                        _c.LIB.free(count)
                    ptr = self._guarded_block(size)
                    try:
                        self.assertEqual(_c.LIB.fm_pagination_get_print_area_at(pagination, 0, ptr), 0)
                        self._assert_sentinel_intact(ptr, size, "fm_cell_rect_t")
                        self.assertEqual(struct.unpack("<IIII", _c.LIB.read_bytes(ptr, size)), (0, 0, 4, 2))
                    finally:
                        _c.LIB.free(ptr)
                finally:
                    _c.LIB.fm_pagination_destroy(pagination)
            finally:
                _c.LIB.free(out)


class CellMutatorDocstringTests(unittest.TestCase):
    def test_cell_mutators_document_zero_based_coordinates(self) -> None:
        for name in ("set_number", "set_bool", "set_error", "set_text", "set_blank", "set_formula"):
            with self.subTest(method=name):
                doc = inspect.getdoc(getattr(Sheet, name))
                self.assertTrue(doc, f"{name} has no docstring")
                self.assertIn("0-based", doc)

    def test_set_error_documents_the_error_code_ordinal(self) -> None:
        doc = inspect.getdoc(Sheet.set_error)
        self.assertIn("ErrorCode", doc)
        self.assertIn("Value.error_code", doc)

    def test_every_public_method_has_a_docstring(self) -> None:
        for cls in (Workbook, Sheet):
            undocumented = [
                name
                for name, member in inspect.getmembers(cls, callable)
                if not name.startswith("_") and not (inspect.getdoc(member) or "").strip()
            ]
            self.assertEqual(undocumented, [], cls.__name__)


class EvaluateCfRangeDefaultTests(unittest.TestCase):
    def test_default_today_is_the_host_date_for_time_period_rules(self) -> None:
        """Omitting ``today`` must not pin the basis to 1899-12-30."""
        with Workbook.create_default() as wb:
            dxf_index = wb.ooxml.add_dxf(DxfInput())
            wb.sheet(0).set_number(0, 0, 0.0)
            wb.sheet(0).add_cf_rule(
                CfTimePeriodRuleInput(ranges=[CellRect(0, 0, 0, 0)], time_period=0, dxf_index=dxf_index),
            )
            wb.recalc()
            omitted = wb.sheet(0).evaluate_cf_range(CellRect(0, 0, 0, 0))
            explicit = wb.sheet(0).evaluate_cf_range(CellRect(0, 0, 0, 0), today=datetime.date(2024, 3, 15))
        self.assertEqual([c for c in omitted if c.matches], [])
        self.assertEqual(len(omitted), len(explicit))
        self.assertEqual([c for c in explicit if c.matches], [])

    def test_omitted_today_reads_the_pinned_date(self) -> None:
        """Omitting ``today`` evaluates ``timePeriod`` rules against the pinned date."""
        with Workbook.create_default() as wb:
            wb.set_pinned_now(datetime.datetime(2024, 3, 15, 9, 30))
            dxf_index = wb.ooxml.add_dxf(DxfInput())
            wb.sheet(0).set_number(0, 0, 45366.0)  # 2024-03-15
            wb.sheet(0).add_cf_rule(
                CfTimePeriodRuleInput(ranges=[CellRect(0, 0, 0, 0)], time_period=0, dxf_index=dxf_index),
            )
            wb.recalc()
            omitted = wb.sheet(0).evaluate_cf_range(CellRect(0, 0, 0, 0))
            elsewhere = wb.sheet(0).evaluate_cf_range(CellRect(0, 0, 0, 0), today=datetime.datetime(2024, 3, 16))
            same_day = wb.sheet(0).evaluate_cf_range(CellRect(0, 0, 0, 0), today=datetime.date(2024, 3, 15))
        self.assertEqual(len([c for c in omitted if c.matches]), 1)
        self.assertEqual([c for c in elsewhere if c.matches], [])
        self.assertEqual(len([c for c in same_day if c.matches]), 1)


class ArgumentScopeTests(unittest.TestCase):
    """Every WASM block a call allocates for its arguments is freed on every exit."""

    def _assert_rejected_without_leak(self, call, exception=FormulonError) -> int:
        """Run ``call``, expect ``exception``, and return how many blocks it allocated."""
        live: set = set()
        allocations = 0
        real_alloc, real_free = _c.LIB.alloc, _c.LIB.free

        def tracking_alloc(size):
            nonlocal allocations
            allocations += 1
            ptr = real_alloc(size)
            if ptr:
                live.add(ptr)
            return ptr

        def tracking_free(ptr):
            live.discard(ptr)
            return real_free(ptr)

        with mock.patch.object(_c.LIB, "alloc", tracking_alloc), mock.patch.object(_c.LIB, "free", tracking_free):
            with self.assertRaises(exception):
                call()
        self.assertEqual(live, set())
        return allocations

    def test_only_the_scope_allocates_or_frees_argument_memory(self) -> None:
        # A hand-rolled alloc/free pair is how a leak on an exceptional exit
        # keeps coming back; every argument block goes through `_ArgScope`.
        raw = {"alloc", "free"}
        findings = []
        package = Path(formulon.__file__).resolve().parent
        for module in sorted(package.glob("*.py")):
            if module.name == "_c.py":
                continue
            tree = ast.parse(module.read_text(encoding="utf-8"))
            exempt = {
                id(node)
                for cls in ast.walk(tree)
                if isinstance(cls, ast.ClassDef) and cls.name == "_ArgScope"
                for node in ast.walk(cls)
            }
            for node in ast.walk(tree):
                if (
                    isinstance(node, ast.Call)
                    and isinstance(node.func, ast.Attribute)
                    and isinstance(node.func.value, ast.Name)
                    and node.func.value.id in ("LIB", "lib")
                    and node.func.attr in raw
                    and id(node) not in exempt
                ):
                    findings.append(f"{module.name}:{node.lineno} {ast.unparse(node.func)}")
        self.assertEqual(findings, [])

    def test_a_packer_that_raises_after_allocating_frees_everything(self) -> None:
        runs = [PhoneticRun(i, i + 1, "\u30a2") for i in range(200)] + [PhoneticRun(-1, 1, "\u30a4")]
        ranges = [CellRect(0, 0, 0, 0)] * 200 + [CellRect(-1, 0, 0, 0)]
        bad_spec = PivotDataFieldInput(name="Sum", field_index=-1)
        wb, pivot = PivotFilterPersistenceTests._workbook_with_pivot()
        with wb:
            for name, call in (
                ("set_cell_phonetic_runs", lambda: wb.sheet(0).set_cell_phonetic_runs(0, 0, runs)),
                (
                    "set_selection",
                    lambda: wb.sheet(0).set_selection(SheetSelectionInput(CellAddress(0, 0), ranges=list(ranges))),
                ),
                ("add_pivot_data_field", lambda: wb.sheet(0).pivot(pivot).add_data_field(bad_spec)),
                ("set_pivot_data_field", lambda: wb.sheet(0).pivot(pivot).set_data_field(0, bad_spec)),
            ):
                with self.subTest(method=name):
                    self.assertGreater(self._assert_rejected_without_leak(call), 0)

    def test_repeated_rejected_calls_do_not_grow_linear_memory(self) -> None:
        ranges = [CellRect(0, 0, 0, 0)] * 20000 + [CellRect(-1, 0, 0, 0)]
        with Workbook.create_default() as wb:
            with self.assertRaises(ValueError):
                wb.sheet(0).set_selection(SheetSelectionInput(CellAddress(0, 0), ranges=list(ranges)))
            before = _c.LIB.memory.data_len(_c.LIB.store)
            for _ in range(20):
                with self.assertRaises(ValueError):
                    wb.sheet(0).set_selection(SheetSelectionInput(CellAddress(0, 0), ranges=list(ranges)))
            self.assertEqual(_c.LIB.memory.data_len(_c.LIB.store), before)

    def test_a_value_of_the_wrong_type_is_rejected_before_any_allocation(self) -> None:
        with Workbook.create_default() as wb:
            allocations = self._assert_rejected_without_leak(lambda: wb.format_value([1], "General"), TypeError)  # type: ignore[arg-type]
            self.assertEqual(allocations, 0)


class LoneSurrogateTests(unittest.TestCase):
    """A lone surrogate is stored as U+FFFD, as the JavaScript bindings store it."""

    def _assert_normalised(self, wb: Workbook) -> None:
        self.assertEqual(wb.sheet(0).get_value(0, 0), "a\ufffdb")
        self.assertEqual(wb.sheet(0).get_value(2, 0), "\U0001f600")
        self.assertEqual(wb.sheet(0).name, "S\ufffd")
        self.assertEqual(wb.sheet(0).get_formula(1, 0), '="x\ufffd"')
        formula = next(n.formula for n in wb.list_defined_names() if n.name == "Lone")
        self.assertIn("y\ufffd", formula)

    def test_text_names_and_formulas_store_the_replacement_character(self) -> None:
        with Workbook.create_default() as wb:
            wb.sheet(0).set_text(0, 0, "a\ud800b")
            # A surrogate pair is one code point, not two replacements.
            wb.sheet(0).set_text(2, 0, "\ud83d\ude00")
            wb.sheet(0).rename("S\udc00")
            wb.sheet(0).set_formula(1, 0, '="x\ud800"')
            wb.set_defined_name(DefinedNameInput("Lone", '="y\udfff"'))
            self._assert_normalised(wb)
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self._assert_normalised(reloaded)


class ValueValidationTests(unittest.TestCase):
    """``Value`` holds exactly the payload its kind takes, checked at construction."""

    def test_an_integer_kind_is_coerced_to_value_kind(self) -> None:
        value = Value(kind=1, number=5.5)
        self.assertIs(value.kind, ValueKind.NUMBER)
        with Workbook.create_default() as wb:
            self.assertEqual(wb.format_value(value, "General").text, "5.5")

    def test_a_missing_or_extra_payload_is_a_value_error(self) -> None:
        for kwargs in (
            {"kind": ValueKind.NUMBER},
            {"kind": ValueKind.TEXT},
            {"kind": ValueKind.ERROR},
            {"kind": ValueKind.NUMBER, "number": 1.0, "text": "x"},
            {"kind": ValueKind.BLANK, "boolean": True},
            {"kind": 99},
        ):
            with self.subTest(**kwargs):
                with self.assertRaises(ValueError):
                    Value(**kwargs)

    def test_a_payload_of_the_wrong_type_is_a_type_error(self) -> None:
        for kwargs in (
            {"kind": 3, "text": "abc".encode()},
            {"kind": ValueKind.TEXT, "text": 123},
            {"kind": ValueKind.BOOL, "boolean": 1},
            {"kind": ValueKind.NUMBER, "number": "5"},
            {"kind": "1", "number": 1.0},
        ):
            with self.subTest(**kwargs):
                with self.assertRaises(TypeError):
                    Value(**kwargs)


class PackedIntegerGuardTests(unittest.TestCase):
    """Packed struct fields reject a fraction, a string or an overflow as the scalar arguments do."""

    def test_shape_emu_offsets(self) -> None:
        with Workbook.create_default() as wb:
            for bad, exception in ((1.9, TypeError), (2.0, TypeError), ("5", TypeError), (2**63, ValueError)):
                with self.subTest(from_col_offset_emu=bad):
                    with self.assertRaises(exception):
                        wb.sheet(0).add_shape(
                            ShapeInput(
                                anchor=DrawingAnchorInput(
                                    kind=AnchorKind.TWO_CELL, to_row=5, to_col=4, from_col_offset_emu=bad
                                )
                            )
                        )
            self.assertEqual(wb.sheet(0).list_shapes(), [])

    def test_pivot_field_order_indices(self) -> None:
        wb, pivot = PivotFilterPersistenceTests._workbook_with_pivot()
        with wb:
            for indices in ([1.7], [-1], ["1"]):
                with self.subTest(indices=indices):
                    with self.assertRaises((TypeError, FormulonError)):
                        wb.sheet(0).pivot(pivot).set_row_field_order(indices)

    def test_data_bar_length_percentages(self) -> None:
        with Workbook.create_default() as wb:
            for bad, exception in ((1.5, TypeError), (256, ValueError), (-1, ValueError)):
                with self.subTest(min_length_pct=bad):
                    rule = CfDataBarRuleInput(
                        ranges=[CellRect(0, 0, 2, 0)],
                        data_bar=CfDataBarInput(
                            CfValueObject(3), CfValueObject(4), _rgb(0, 0, 255), min_length_pct=bad
                        ),
                    )
                    with self.assertRaises(exception):
                        wb.sheet(0).add_cf_rule(rule)
            self.assertEqual(len(wb.sheet(0).list_cf_rules()), 0)


if __name__ == "__main__":
    unittest.main()
