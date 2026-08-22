"""Contract tests for the Python wrapper's own error and scratch policies.

These cover behaviour the wrapper owns rather than the engine: which
exception type a defensive path raises, what a failure's ``op`` prefix
names, how the WASM-locator error reports its candidates, and the
allocation discipline of the collection readers.
"""

from __future__ import annotations

import inspect
import io
import os
import struct
import sys
import tempfile
import unittest
import zipfile
from unittest import mock

import formulon
from formulon import (
    CalcMode,
    CellXf,
    ConditionalFormatInput,
    DifferentialFormat,
    ErrorCode,
    ExternalLinkKind,
    FillRecord,
    FontRecord,
    FormulonError,
    LogLevel,
    MergeRange,
    PhoneticRun,
    PivotAggregation,
    PivotAxis,
    PivotDataFieldSpec,
    PivotFieldSpec,
    PivotFilterSpec,
    PivotFilterType,
    PivotFilterValueKind,
    PivotWorksheetSource,
    ValueKind,
    Workbook,
    _c,
)


class FormulonErrorStatusTests(unittest.TestCase):
    def test_paginate_null_result_raises_formulon_error_with_status(self) -> None:
        """The defensive null-handle path must not raise ValueError."""
        with Workbook.create_default() as wb:
            # Force `fm_workbook_paginate` to report success while leaving
            # the out-slot at 0, which is the branch under test.
            real_paginate = _c.LIB.fm_workbook_paginate

            def fake_paginate(handle, sheet, out):
                real_paginate(handle, sheet, out)
                _c.LIB.write_bytes(out, b"\x00\x00\x00\x00")
                return 0

            with mock.patch.object(_c.LIB, "fm_workbook_paginate", fake_paginate, create=True):
                with self.assertRaises(FormulonError) as caught:
                    wb.paginate(0)
        self.assertEqual(caught.exception.status, 7001)
        self.assertIn("fm_workbook_paginate", str(caught.exception))


class ErrorOperationNameTests(unittest.TestCase):
    def test_count_helper_failure_names_the_c_abi_entry_point(self) -> None:
        """`op` must name the ABI symbol, never a binding-internal helper."""
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError) as caught:
                wb.merge_count(99)
        message = str(caught.exception)
        self.assertIn("fm_sheet_get_merge_count", message)
        self.assertNotIn("_wrapped", message)

    def test_trace_failure_names_the_c_abi_entry_point(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError) as caught:
                wb.precedents(99, 0, 0)
        message = str(caught.exception)
        self.assertIn("fm_workbook_precedents", message)
        self.assertNotIn("_wrapped", message)

    def test_exports_expose_the_c_abi_symbol_as_their_name(self) -> None:
        self.assertEqual(_c.LIB.fm_workbook_sheet_count.__name__, "fm_workbook_sheet_count")


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


class FunctionMetadataLocaleTests(unittest.TestCase):
    def test_unknown_function_returns_none(self) -> None:
        self.assertIsNone(Workbook.function_metadata("NOPE"))

    def test_out_of_range_locale_raises_instead_of_returning_none(self) -> None:
        """An invalid locale is an API error, not a "function not found"."""
        for locale in (-1, 2, 99):
            with self.subTest(locale=locale):
                with self.assertRaises(FormulonError):
                    Workbook.function_metadata("SUM", locale)

    def test_out_of_range_locale_matches_the_sibling_catalog_methods(self) -> None:
        with self.assertRaises(FormulonError):
            Workbook.localize_function_name("SUM", 99)
        with self.assertRaises(FormulonError):
            Workbook.canonicalize_function_name("SUM", 99)

    def test_known_function_resolves_in_both_locales(self) -> None:
        for locale in (0, 1):
            with self.subTest(locale=locale):
                self.assertIsNotNone(Workbook.function_metadata("SUM", locale))


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
            wb.set_error(0, 0, 0, ErrorCode.DIV0)
            value = wb.get_value(0, 0, 0)
        self.assertEqual(value.kind, ValueKind.ERROR)
        self.assertEqual(value.error_code, ErrorCode.DIV0)


class StyleAndCalcAccessorTests(unittest.TestCase):
    """Writers whose readers already existed, and vice versa."""

    @staticmethod
    def _xf() -> CellXf:
        return CellXf(
            font_index=0,
            fill_index=0,
            border_index=0,
            num_fmt_id=0,
            horizontal_align=0,
            vertical_align=2,
            wrap_text=False,
        )

    def test_named_style_xf_writer_round_trips_through_its_reader(self) -> None:
        with Workbook.create_default() as wb:
            xf_id = wb.add_cell_style_xf(self._xf())
            read_back = wb.get_cell_style_xf(xf_id)
        self.assertEqual(read_back.num_fmt_id, 0)
        self.assertEqual(read_back.vertical_align, 2)

    def test_named_style_xf_writer_deduplicates(self) -> None:
        with Workbook.create_default() as wb:
            first = wb.add_cell_style_xf(self._xf())
            second = wb.add_cell_style_xf(self._xf())
        self.assertEqual(first, second)

    def test_iterative_settings_read_back_what_was_set(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_iterative(True, 42, 0.5)
            settings = wb.get_iterative()
        self.assertTrue(settings.enabled)
        self.assertEqual(settings.max_iterations, 42)
        self.assertEqual(settings.max_change, 0.5)

    def test_iterative_cap_and_threshold_survive_being_disabled(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_iterative(True, 42, 0.5)
            wb.set_iterative(False, 42, 0.5)
            settings = wb.get_iterative()
        self.assertFalse(settings.enabled)
        self.assertEqual(settings.max_iterations, 42)
        self.assertEqual(settings.max_change, 0.5)

    def test_enable_toggle_keeps_the_cap_and_threshold(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_iterative(True, 42, 0.5)
            wb.set_iterative_enabled(False)
            disabled = wb.get_iterative()
            wb.set_iterative_enabled(True)
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
            font_index = wb.add_font(FontRecord(name="Arial", size=12.0))
            # Excel reserves fill 0 for `none` and fill 1 for `gray125`.
            fill_index = wb.add_fill(FillRecord(pattern=1, fg_argb=0xFFFF0000))
            border_index = wb.add_border({"left": {"style": 1, "color_argb": 0xFF000000}})
            xf_index = wb.add_cell_xf(
                CellXf(
                    font_index=font_index,
                    fill_index=fill_index,
                    border_index=border_index,
                    num_fmt_id=0,
                    horizontal_align=0,
                    vertical_align=0,
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


class StyleBatchTests(unittest.TestCase):
    """The bulk style path the per-record methods' docs point callers at."""

    def test_batch_returns_one_index_per_requested_record(self) -> None:
        with Workbook.create_empty() as wb:
            wb.add_sheet("Sheet1")
            indices = wb.add_batch(
                fonts=[FontRecord(name="Arial", size=10.0), FontRecord(name="Arial", size=12.0)],
                fills=[FillRecord(pattern=1, fg_argb=0xFF00FF00)],
                borders=[{"left": {"style": 1, "color_argb": 0xFF000000}}],
                num_fmts=["0.000"],
            )
            xf = wb.add_batch(
                cell_xfs=[
                    CellXf(
                        font_index=indices.fonts[1],
                        fill_index=indices.fills[0],
                        border_index=indices.borders[0],
                        num_fmt_id=indices.num_fmts[0],
                        horizontal_align=0,
                        vertical_align=0,
                        wrap_text=False,
                    )
                ]
            )
            font_count = wb.font_count()
            read_back = wb.get_font(indices.fonts[1])
        self.assertEqual(len(indices.fonts), 2)
        self.assertEqual(len(indices.fills), 1)
        self.assertEqual(len(indices.borders), 1)
        self.assertEqual(len(indices.num_fmts), 1)
        self.assertEqual(indices.cell_xfs, [])
        self.assertEqual(len(xf.cell_xfs), 1)
        self.assertEqual(read_back.size, 12.0)
        # The seeded default font plus the two requested ones.
        self.assertEqual(font_count, 3)

    def test_batch_matches_the_per_record_path(self) -> None:
        fonts = [FontRecord(name="Arial", size=float(size)) for size in (8, 9, 10)]
        with Workbook.create_empty() as wb:
            wb.add_sheet("Sheet1")
            batched = wb.add_batch(fonts=fonts).fonts
        with Workbook.create_empty() as wb:
            wb.add_sheet("Sheet1")
            per_record = [wb.add_font(record) for record in fonts]
        self.assertEqual(batched, per_record)

    def test_batch_deduplicates_against_existing_records(self) -> None:
        record = FontRecord(name="Arial", size=12.0)
        with Workbook.create_empty() as wb:
            wb.add_sheet("Sheet1")
            first = wb.add_font(record)
            repeated = wb.add_batch(fonts=[record, record]).fonts
        self.assertEqual(repeated, [first, first])

    def test_batch_with_no_arrays_is_a_no_op(self) -> None:
        with Workbook.create_empty() as wb:
            wb.add_sheet("Sheet1")
            before = wb.font_count()
            indices = wb.add_batch()
            after = wb.font_count()
        self.assertEqual(indices.fonts, [])
        self.assertEqual(before, after)

    def test_batch_rejecting_an_xf_leaves_the_workbook_unchanged(self) -> None:
        with Workbook.create_empty() as wb:
            wb.add_sheet("Sheet1")
            before = wb.font_count()
            with self.assertRaises(FormulonError):
                wb.add_batch(
                    fonts=[FontRecord(name="Arial", size=12.0)],
                    cell_xfs=[
                        CellXf(
                            font_index=9999,
                            fill_index=0,
                            border_index=0,
                            num_fmt_id=0,
                            horizontal_align=0,
                            vertical_align=0,
                            wrap_text=False,
                        )
                    ],
                )
            self.assertEqual(wb.font_count(), before)


class MemoryUsageTests(unittest.TestCase):
    def test_estimate_grows_with_stored_cells(self) -> None:
        with Workbook.create_default() as wb:
            empty = wb.memory_usage()
            for row in range(2000):
                wb.set_number(0, row, 0, float(row))
            filled = wb.memory_usage()
        self.assertGreater(filled, empty)

    def test_estimate_is_stable_across_repeated_calls(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_number(0, 0, 0, 1.0)
            self.assertEqual(wb.memory_usage(), wb.memory_usage())

    def test_closed_handle_raises_rather_than_reporting_zero(self) -> None:
        wb = Workbook.create_default()
        wb.close()
        with self.assertRaises(FormulonError) as caught:
            wb.memory_usage()
        self.assertEqual(caught.exception.status, 7000)


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
                    # An implicit-intersection formula the XLSB encoder
                    # cannot lower, which makes the writer emit a per-cell
                    # warn record.
                    wb.set_formula(0, 0, 0, "=@A1:A10")
                    wb.recalc()
                    wb.save_as(formulon.WorkbookFormat.XLSB)
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
                self.assertEqual(caught.exception.status, 2)

    def test_every_level_ordinal_is_accepted(self) -> None:
        for level in LogLevel:
            with self.subTest(level=level):
                formulon.set_log_min_level(level)


class WasmOnlyCapabilityTests(unittest.TestCase):
    """Capabilities the WASM npm package exposed before this wheel did."""

    def test_phonetic_guide_round_trips(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_text(0, 0, 0, "日本語")
            self.assertEqual(wb.get_phonetic(0, 0, 0), "")
            wb.set_phonetic(0, 0, 0, "ニホンゴ")
            self.assertEqual(wb.get_phonetic(0, 0, 0), "ニホンゴ")

    def test_phonetic_guide_survives_a_save_load_round_trip(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_text(0, 0, 0, "日本語")
            wb.set_phonetic(0, 0, 0, "ニホンゴ")
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertEqual(reloaded.get_phonetic(0, 0, 0), "ニホンゴ")

    def test_empty_phonetic_guide_clears_it(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_text(0, 0, 0, "日本語")
            wb.set_phonetic(0, 0, 0, "ニホンゴ")
            wb.set_phonetic(0, 0, 0, "")
            self.assertEqual(wb.get_phonetic(0, 0, 0), "")

    def test_phonetic_runs_keep_their_spans(self) -> None:
        runs = [PhoneticRun(0, 2, "トウキョウ"), PhoneticRun(2, 3, "ト")]
        with Workbook.create_default() as wb:
            wb.set_text(0, 0, 0, "東京都")
            self.assertEqual(wb.get_phonetic_runs(0, 0, 0), [])
            wb.set_phonetic_runs(0, 0, 0, runs)
            self.assertEqual(wb.get_phonetic_runs(0, 0, 0), runs)
            # The flattening getter still reports the concatenation.
            self.assertEqual(wb.get_phonetic(0, 0, 0), "トウキョウト")
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertEqual(reloaded.get_phonetic_runs(0, 0, 0), runs)

    def test_whole_cell_setter_collapses_the_spans(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_text(0, 0, 0, "東京都")
            wb.set_phonetic_runs(0, 0, 0, [PhoneticRun(0, 2, "トウキョウ"), PhoneticRun(2, 3, "ト")])
            wb.set_phonetic(0, 0, 0, wb.get_phonetic(0, 0, 0))
            self.assertEqual(wb.get_phonetic_runs(0, 0, 0), [PhoneticRun(0, 3, "トウキョウト")])

    def test_empty_run_list_clears_the_guide(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_text(0, 0, 0, "東京都")
            wb.set_phonetic_runs(0, 0, 0, [PhoneticRun(0, 3, "トウキョウト")])
            wb.set_phonetic_runs(0, 0, 0, [])
            self.assertEqual(wb.get_phonetic_runs(0, 0, 0), [])
            self.assertEqual(wb.get_phonetic(0, 0, 0), "")

    def test_out_of_order_runs_are_rejected(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_text(0, 0, 0, "東京都")
            with self.assertRaises(FormulonError):
                wb.set_phonetic_runs(0, 0, 0, [PhoneticRun(2, 3, "ト"), PhoneticRun(0, 2, "トウ")])

    def test_default_font_replaces_what_an_unstyled_cell_saves_as(self) -> None:
        with Workbook.create_default() as wb:
            self.assertEqual(wb.get_font(0).name, "Calibri")
            # add_font can only ever append beside the seeded default.
            self.assertGreater(wb.add_font(FontRecord(name="游ゴシック", size=11.0)), 0)
            self.assertEqual(wb.get_font(0).name, "Calibri")

            wb.set_default_font(FontRecord(name="游ゴシック", size=11.0, has_charset=True, charset=128))
            self.assertEqual(wb.get_font(0).name, "游ゴシック")
            self.assertEqual(wb.get_font(0).charset, 128)
            data = wb.save()
        with Workbook.load(data) as reloaded:
            self.assertEqual(reloaded.get_font(0).name, "游ゴシック")

    def test_set_font_overwrites_an_existing_slot(self) -> None:
        with Workbook.create_default() as wb:
            index = wb.add_font(FontRecord(name="Meiryo", size=12.0))
            wb.set_font(index, FontRecord(name="MS Gothic", size=9.0))
            self.assertEqual(wb.get_font(index).name, "MS Gothic")
            self.assertEqual(wb.font_count(), index + 1)
            with self.assertRaises(FormulonError):
                wb.set_font(wb.font_count(), FontRecord(name="MS Gothic"))

    def test_table_create_update_remove(self) -> None:
        with Workbook.create_default() as wb:
            wb.set_text(0, 0, 0, "A")
            wb.set_text(0, 0, 1, "B")
            index = wb.table_create(0, "A1:B3", "Table1", "Table1", ["A", "B"])
            self.assertEqual([t.ref for t in wb.iter_tables()], ["A1:B3"])
            self.assertEqual([t.name for t in wb.iter_tables()], ["Table1"])

            wb.table_update(index, "A1:B5")
            self.assertEqual([t.ref for t in wb.iter_tables()], ["A1:B5"])

            wb.table_remove(index)
            self.assertEqual(list(wb.iter_tables()), [])

    def test_table_create_rejects_a_column_count_that_does_not_match_ref(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.table_create(0, "A1:B3", "Table1", "Table1", ["OnlyOne"])

    def test_auto_filter_xml_round_trips(self) -> None:
        fragment = '<autoFilter ref="A1:B3"/>'
        with Workbook.create_default() as wb:
            self.assertEqual(wb.get_auto_filter_xml(0), "")
            wb.set_auto_filter_xml(0, fragment)
            self.assertEqual(wb.get_auto_filter_xml(0), fragment)
            wb.set_auto_filter_xml(0, "")
            self.assertEqual(wb.get_auto_filter_xml(0), "")

    def test_named_cell_style_registration_reaches_the_reader(self) -> None:
        with Workbook.create_default() as wb:
            xf_id = wb.add_cell_style_xf(StyleAndCalcAccessorTests._xf())
            before = wb.cell_style_count()
            wb.set_cell_style("MyStyle", xf_id)
            names = [wb.get_cell_style(i).name for i in range(wb.cell_style_count())]
        self.assertGreater(len(names), before)
        self.assertIn("MyStyle", names)

    def test_named_cell_style_rejects_an_unregistered_xf_id(self) -> None:
        with Workbook.create_default() as wb:
            with self.assertRaises(FormulonError):
                wb.set_cell_style("Bogus", 9999)


class OneShotEvaluationTests(unittest.TestCase):
    """``eval_formula`` must agree value-for-value with the JS surfaces.

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
        ("=SEQUENCE(3)", ValueKind.NUMBER, 1.0),
        ("=ROWS(SEQUENCE(3))", ValueKind.NUMBER, 3.0),
    )

    def test_values_match_the_shared_one_shot_table(self) -> None:
        for formula, kind, payload in self.CASES:
            with self.subTest(formula=formula):
                value = formulon.eval_formula(formula)
                self.assertEqual(value.kind, kind)
                if kind is ValueKind.BOOL:
                    self.assertEqual(value.boolean, payload)
                else:
                    self.assertEqual(value.number, payload)
                # Writing into the anchor and recalcing would make the
                # anchor-referencing cases #REF!.
                self.assertIsNone(value.error_code)

    def test_the_anchor_is_left_blank_rather_than_written(self) -> None:
        """The distinguishing observation between the two semantics."""
        self.assertEqual(formulon.eval_formula("=ISBLANK(A1)").boolean, True)

    def test_array_results_reduce_to_their_top_left_element(self) -> None:
        self.assertEqual(formulon.eval_formula("=SEQUENCE(3)").number, 1.0)
        self.assertEqual(formulon.eval_formula("=SEQUENCE(2,3)").number, 1.0)


class PivotFilterSessionStateTests(unittest.TestCase):
    """The active-filter list lives only as long as the handle.

    The npm and npm-native packages assert the same three properties, so a
    surface that started persisting (or dropping) filters would break here
    and there together.
    """

    @staticmethod
    def _workbook_with_pivot():
        wb = Workbook.create_default()
        cache_id = wb.pivot_cache_create()
        # A cache with no declared source cannot be saved: Excel offers to
        # repair any package containing one, so the writer refuses instead.
        wb.set_pivot_cache_worksheet_source(cache_id, PivotWorksheetSource(ref="A1:B3", sheet="Sheet1"))
        wb.pivot_cache_field_add(cache_id, "Region")
        wb.pivot_cache_field_add(cache_id, "Amount")
        for region, amount in (("East", 10.0), ("West", 30.0)):
            rec = wb.pivot_cache_record_add(cache_id)
            wb.pivot_cache_record_set_text(cache_id, rec, 0, region)
            wb.pivot_cache_record_set_number(cache_id, rec, 1, amount)
        pivot = wb.pivot_create(0, "Pivot1", cache_id, 0, 4)
        wb.pivot_field_add(0, pivot, PivotFieldSpec(source_name="Region", axis=PivotAxis.ROW))
        amount_field = wb.pivot_field_add(0, pivot, PivotFieldSpec(source_name="Amount", axis=PivotAxis.VALUE))
        wb.pivot_data_field_add(
            0,
            pivot,
            PivotDataFieldSpec(name="Sum of Amount", field_index=amount_field, aggregation=PivotAggregation.SUM),
        )
        return wb, pivot

    @staticmethod
    def _filter() -> PivotFilterSpec:
        return PivotFilterSpec(
            axis=PivotAxis.ROW,
            field_name="Region",
            type=PivotFilterType.VALUE_GREATER_THAN,
            value_kind=PivotFilterValueKind.DOUBLE,
            value_double=15.0,
        )

    def test_added_filter_reads_back_field_for_field(self) -> None:
        wb, pivot = self._workbook_with_pivot()
        try:
            wb.pivot_filter_add(0, pivot, self._filter())
            self.assertEqual(wb.pivot_filter_count(0, pivot), 1)
            got = wb.pivot_filter_at(0, pivot, 0)
        finally:
            wb.close()
        self.assertEqual(got.axis, PivotAxis.ROW)
        self.assertEqual(got.field_name, "Region")
        self.assertEqual(got.type, PivotFilterType.VALUE_GREATER_THAN)
        self.assertEqual(got.value_kind, PivotFilterValueKind.DOUBLE)
        self.assertEqual(got.value_double, 15.0)

    def test_count_reports_only_what_this_session_added(self) -> None:
        wb, pivot = self._workbook_with_pivot()
        try:
            self.assertEqual(wb.pivot_filter_count(0, pivot), 0)
            wb.pivot_filter_add(0, pivot, self._filter())
            self.assertEqual(wb.pivot_filter_count(0, pivot), 1)
        finally:
            wb.close()

    def test_filters_do_not_survive_save_and_load(self) -> None:
        """The core of the contract: the list is not written by save()."""
        wb, pivot = self._workbook_with_pivot()
        try:
            wb.pivot_filter_add(0, pivot, self._filter())
            self.assertEqual(wb.pivot_filter_count(0, pivot), 1)
            data = wb.save()
        finally:
            wb.close()
        with Workbook.load(data) as reloaded:
            self.assertEqual(reloaded.pivot_count(0), 1)
            self.assertEqual(reloaded.pivot_filter_count(0, 0), 0)

    def test_out_of_range_index_is_rejected(self) -> None:
        wb, pivot = self._workbook_with_pivot()
        try:
            with self.assertRaises(FormulonError):
                wb.pivot_filter_at(0, pivot, 99)
        finally:
            wb.close()


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
            lambda wb, n: [wb.add_merge(0, MergeRange(r, 4, r, 6)) for r in range(n)],
            lambda wb: wb.get_merges(0),
        )

    def test_get_hyperlinks_allocation_count_is_flat(self) -> None:
        self._assert_flat_in_item_count(
            lambda wb, n: [wb.add_hyperlink(0, r, 8, "https://example.com/") for r in range(n)],
            lambda wb: wb.get_hyperlinks(0),
        )

    def test_get_comments_allocation_count_is_flat(self) -> None:
        self._assert_flat_in_item_count(
            lambda wb, n: [wb.set_comment(0, r, 10, "author", "text") for r in range(n)],
            lambda wb: wb.get_comments(0),
        )

    def test_iter_defined_names_allocation_count_is_flat(self) -> None:
        self._assert_flat_in_item_count(
            lambda wb, n: [wb.set_defined_name(f"Name{r}", "=1") for r in range(n)],
            lambda wb: list(wb.iter_defined_names()),
        )

    def test_get_sheet_row_overrides_allocation_count_is_flat(self) -> None:
        self._assert_flat_in_item_count(
            lambda wb, n: [wb.set_row_height(0, r, 15.0) for r in range(n)],
            lambda wb: wb.get_sheet_row_overrides(0),
        )

    def test_precedents_allocation_count_is_flat(self) -> None:
        def populate(wb, n):
            for r in range(n):
                wb.set_number(0, r + 1, 1, float(r))
            wb.set_formula(0, 0, 0, f"=SUM(B2:B{n + 1})")
            wb.recalc()

        self._assert_flat_in_item_count(populate, lambda wb: wb.precedents(0, 0, 0))

    def test_evaluate_cf_range_allocation_count_is_flat(self) -> None:
        def populate(wb, n):
            dxf_index = wb.add_dxf(DifferentialFormat())
            for r in range(n):
                wb.set_number(0, r, 0, float(r + 100))
            wb.add_conditional_format(
                0,
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 0, n - 1, 0)],
                    type=1,
                    op_engaged=True,
                    op=4,
                    formula1="10",
                    dxf_id_engaged=True,
                    dxf_id=dxf_index,
                ),
            )
            wb.recalc()

        self._assert_flat_in_item_count(populate, lambda wb: wb.evaluate_cf_range(0, 0, 0, 39, 0))


class AbiScratchBlockTests(unittest.TestCase):
    """The records the wrapper sizes by hand, measured against the real module.

    ``fm_value_t`` and ``fm_print_range_t`` have no ``_structs.Struct``
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
            wb.set_number(0, 0, 0, 6.25)
            wb.recalc()
            ptr = self._guarded_block(size)
            try:
                status = _c.LIB.fm_workbook_get_value(wb._require(), 0, 0, 0, ptr)
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
        size = 16  # fm_print_range_t: four uint32_t, as Workbook.paginate reads it
        with Workbook.create_default() as wb:
            wb.set_defined_name_scoped("_xlnm.Print_Area", "Sheet1!$A$1:$C$5", 0)
            out = _c.LIB.alloc(4)
            try:
                self.assertEqual(_c.LIB.fm_workbook_paginate(wb._require(), 0, out), 0)
                pagination = _c.LIB.read_u32(out)
                self.assertNotEqual(pagination, 0)
                try:
                    self.assertGreater(int(_c.LIB.fm_pagination_print_area_count(pagination)), 0)
                    ptr = self._guarded_block(size)
                    try:
                        self.assertEqual(_c.LIB.fm_pagination_print_area_at(pagination, 0, ptr), 0)
                        self._assert_sentinel_intact(ptr, size, "fm_print_range_t")
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
                doc = inspect.getdoc(getattr(Workbook, name))
                self.assertTrue(doc, f"{name} has no docstring")
                self.assertIn("0-based", doc)

    def test_set_error_documents_the_error_code_ordinal(self) -> None:
        doc = inspect.getdoc(Workbook.set_error)
        self.assertIn("ErrorCode", doc)
        self.assertIn("#DIV/0!", doc)

    def test_every_public_workbook_method_has_a_docstring(self) -> None:
        undocumented = [
            name
            for name, member in inspect.getmembers(Workbook, callable)
            if not name.startswith("_") and not (inspect.getdoc(member) or "").strip()
        ]
        self.assertEqual(undocumented, [])


class EvaluateCfRangeDefaultTests(unittest.TestCase):
    def test_default_today_serial_disables_time_period_rules(self) -> None:
        """Omitting ``today_serial`` must not pin the basis to 1899-12-30."""
        with Workbook.create_default() as wb:
            dxf_index = wb.add_dxf(DifferentialFormat())
            wb.set_number(0, 0, 0, 0.0)
            wb.add_conditional_format(
                0,
                ConditionalFormatInput(
                    sqref=[MergeRange(0, 0, 0, 0)],
                    type=15,
                    time_period_engaged=True,
                    time_period=0,
                    dxf_id_engaged=True,
                    dxf_id=dxf_index,
                ),
            )
            wb.recalc()
            omitted = wb.evaluate_cf_range(0, 0, 0, 0, 0)
            explicit_nan = wb.evaluate_cf_range(0, 0, 0, 0, 0, float("nan"))
            zero_basis = wb.evaluate_cf_range(0, 0, 0, 0, 0, 0.0)
        self.assertEqual([c for c in omitted if c.matches], [])
        self.assertEqual(len(omitted), len(explicit_nan))
        self.assertEqual(len([c for c in zero_basis if c.matches]), 1)


if __name__ == "__main__":
    unittest.main()
