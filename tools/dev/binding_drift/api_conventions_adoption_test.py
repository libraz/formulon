#!/usr/bin/env python3
"""Fixture tests for the adoption rules (R29-R35) and the pending files of the api-conventions baseline.

Run from the repo root: `python3 -m tools.dev.binding_drift_test` (which imports these classes).
"""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from tools.dev.binding_drift import api_conventions
from tools.dev.binding_drift.api_conventions_bindings_test import BindingRuleTestBase, _py_wb, _wb, hits


def _c_hits(header: str, rule_id: str) -> list[tuple[str, str]]:
    ctx = api_conventions.Context(header, api_conventions.Ledger(None))
    rules = {rule_id: api_conventions.all_rules()[rule_id]}
    return [(v.op_key, v.discriminator) for v in api_conventions.collect(ctx, rules)]


TITLE_GETTER = "FM_API fm_status_t fm_workbook_get_title(const fm_workbook_t* wb, const char** out_title);\n"


class AdoptionCRuleTest(unittest.TestCase):
    def assert_c_rule(self, rule_id: str, bad: str, good: str, expected: list[tuple[str, str]]) -> None:
        self.assertEqual(_c_hits(bad, rule_id), expected)
        self.assertEqual(_c_hits(good, rule_id), [])

    def test_r32_a_record_passed_by_pointer_starts_with_struct_size(self) -> None:
        tail = (
            "FM_API fm_status_t fm_sheet_add_thing(fm_workbook_t* wb, size_t sheet_index, const fm_thing_t* thing,"
            " const fm_cell_rect_t* rect);\n"
            "typedef struct {\n  uint32_t first_row;\n  uint32_t first_col;\n  uint32_t last_row;\n  uint32_t last_col;\n"
            "} fm_cell_rect_t;\n"
            "typedef struct {\n  int32_t unused;\n} fm_embedded_t;\n"
        )
        self.assert_c_rule(
            "R32.struct_size",
            "typedef struct {\n  int32_t kind;\n  uint32_t struct_size;\n} fm_thing_t;\n" + tail,
            "typedef struct {\n  uint32_t struct_size;\n  int32_t kind;\n} fm_thing_t;\n" + tail,
            [("fm_thing_t", "-")],
        )

    def test_r33_a_pointer_output_states_one_of_three_lifetimes(self) -> None:
        destroy = "FM_API void fm_workbook_destroy(fm_workbook_t* wb);\n"
        record = "typedef struct {\n  uint32_t struct_size;\n  const char* id;\n} fm_rule_t;\n"
        out_record = "FM_API fm_status_t fm_sheet_get_rule(const fm_workbook_t* wb, fm_rule_t* out);\n"
        names_in = "FM_API fm_status_t fm_sheet_set_names(fm_workbook_t* wb, const char* const* names, size_t n);\n"
        bad = (
            destroy
            + record
            + "/** Reads the title. */\n"
            + TITLE_GETTER
            + "/** Lifetime: borrowed until the next call. */\n"
            + "FM_API fm_status_t fm_workbook_get_author(const fm_workbook_t* wb, const char** out_author);\n"
            + "/** Lifetime: owned (until fm_workbook_free). */\n"
            + "FM_API fm_status_t fm_workbook_create(fm_workbook_t** out);\n"
            + "/** Reads a rule. */\n"
            + out_record
            + "/** Sets names. */\n"
            + names_in
        )
        good = (
            destroy
            + record
            + "/** Reads the title. Lifetime: transient. */\n"
            + TITLE_GETTER
            + "/** Lifetime: static. */\n"
            + "FM_API fm_status_t fm_workbook_get_author(const fm_workbook_t* wb, const char** out_author);\n"
            + "/** Lifetime: owned (until fm_workbook_destroy). */\n"
            + "FM_API fm_status_t fm_workbook_create(fm_workbook_t** out);\n"
            + "/** Reads a rule. `out->id` Lifetime: owned (until fm_workbook_destroy). */\n"
            + out_record
            + "/** Sets names. */\n"
            + names_in
        )
        self.assert_c_rule(
            "R33.lifetime",
            bad,
            good,
            [
                ("fm_sheet_get_rule", "-"),
                ("fm_workbook_create", "malformed"),
                ("fm_workbook_get_author", "malformed"),
                ("fm_workbook_get_title", "-"),
            ],
        )

    def test_r31_c_colours_are_colour_specs(self) -> None:
        spec = "typedef struct {\n  int32_t kind;\n  uint32_t rgb;\n} fm_color_t;\n"
        cf = "typedef struct {\n  uint8_t r;\n} fm_cf_color_t;\n"
        bad = (
            spec
            + cf
            + (
                "typedef struct {\n  uint32_t color_argb;\n  fm_cf_color_t fill;\n  int32_t has_color;\n"
                "  uint32_t font_color;\n  int32_t color_op;\n} fm_look_t;\n"
                "FM_API fm_status_t fm_sheet_set_tab_color(fm_workbook_t* wb, size_t sheet_index, uint32_t argb);\n"
            )
        )
        good = spec + (
            "typedef struct {\n  fm_color_t color;\n  fm_color_t fill;\n  fm_color_t font_color;\n"
            "  int32_t color_op;\n} fm_look_t;\n"
            "FM_API fm_status_t fm_sheet_set_tab_color(fm_workbook_t* wb, size_t sheet_index,"
            " const fm_color_t* color);\n"
        )
        self.assert_c_rule(
            "R31.color",
            bad,
            good,
            [
                ("fm_cf_color_t", "-"),
                ("fm_look_t", "color_argb"),
                ("fm_look_t", "fill"),
                ("fm_look_t", "font_color"),
                ("fm_look_t", "has_color"),
                ("fm_sheet_set_tab_color", "argb"),
            ],
        )


class AdoptionLedgerRuleTest(BindingRuleTestBase):
    def test_r29_record_ranges_take_a1_text(self) -> None:
        ledger = {
            "entries": {
                "record.table_input": {
                    "kind": "record",
                    "role": "input",
                    "js": "TableInput",
                    "py": "TableInput",
                    "a1_exempt": {"fields": ["fromRow", "fromCol"], "reason": "an anchor with an offset"},
                }
            }
        }
        bad_js = (
            "export interface TableInput {\n  range: SheetRect;\n  ranges?: readonly CellRect[];\n"
            "  anchorRow: number;\n  anchorCol: number;\n  fromRow: number;\n  fromCol: number;\n}\n"
        )
        good_js = (
            "export interface TableInput {\n  range: SheetRect | string;\n  ranges?: readonly (CellRect | string)[] | string;\n"
            "  anchor: CellAddress | string;\n  fromRow: number;\n  fromCol: number;\n}\n"
        )
        bad_py = "class TableInput:\n    ranges: list[CellRect]\n    from_row: int\n    from_col: int\n"
        good_py = (
            "class TableInput:\n    ranges: Union[str, Sequence[Union[CellRect, str]]]\n"
            "    anchor: Optional[CellAddress | str]\n    from_row: int\n    from_col: int\n"
        )
        self.assert_rule(
            "R29.record_a1",
            self.fixture.context("", ledger, wasm=bad_js, py_workbook=bad_py),
            self.fixture.context("", ledger, wasm=good_js, py_workbook=good_py),
            [
                ("py", "record.table_input", "ranges"),
                ("wasm", "record.table_input", "anchor"),
                ("wasm", "record.table_input", "range"),
                ("wasm", "record.table_input", "ranges"),
            ],
        )

    def test_r30_the_ooxml_layer_is_not_a_direct_member(self) -> None:
        ledger = {
            "entries": {
                "cell_xf.get": {"kind": "op", "scope": "workbook", "js": "getCellXf", "py": "get_cell_xf"},
                "font.get": {"kind": "op", "scope": "workbook", "js": "getFont", "namespace": "ooxml"},
                "effective_style.get": {"kind": "op", "scope": "workbook", "js": "getStyle", "resolved": True},
            }
        }
        records = (
            "export interface CellXf { readonly xfId: number }\n"
            "export interface Style { readonly xfIndex: number }\n"
            "export interface Font { readonly name: string }\n"
        )
        bad = records + _wb(
            "  getCellXf(xfIndex: number): CellXf;\n  getFont(index: number): Font;\n  getStyle(): Style;\n"
            "  getSheetPrXml(): string;\n"
        )
        good = records + _wb("  getStyle(): Style;\n  getUsedRange(): CellRect | null;\n")
        self.assert_rule(
            "R30.ooxml_namespace",
            self.fixture.context(
                "", ledger, wasm=bad, py_workbook=_py_wb("    def get_cell_xf(self, xf_index): ...\n")
            ),
            self.fixture.context("", ledger, wasm=good, py_workbook=_py_wb("")),
            [
                ("py", "cell_xf.get", "name"),
                ("py", "cell_xf.get", "xf_index"),
                ("wasm", "cell_xf.get", "name"),
                ("wasm", "cell_xf.get", "returns.xfId"),
                ("wasm", "cell_xf.get", "xfIndex"),
                ("wasm", "font.get", "namespace"),
                ("wasm", "getSheetPrXml", "xml"),
            ],
        )

    def test_r31_binding_colours_are_color_and_inputs_color_input(self) -> None:
        ledger = {
            "entries": {
                "record.font": {"kind": "record", "role": "result", "js": "Font"},
                "record.spark_input": {"kind": "record", "role": "input", "js": "SparkInput"},
                "record.resolved_color": {
                    "kind": "record",
                    "role": "result",
                    "js": "ResolvedColor",
                    "color": "resolved",
                },
                "record.cf_color": {"kind": "record", "role": "result", "was": {"js": "CfColor"}},
            }
        }
        bad = (
            "export interface Font { readonly colorArgb: number; readonly color: Color }\n"
            "export interface SparkInput { colorSeries?: Color; fill?: CfColor }\n"
            "export interface ResolvedColor { readonly argb: number }\n"
            "export interface CfColor { readonly r: number }\n" + _wb("  setTabColor(color: number): void;\n")
        )
        good = (
            "export interface Font { readonly color: Color }\n"
            "export interface SparkInput { colorSeries?: ColorInput; fill?: ColorInput }\n"
            "export interface ResolvedColor { readonly argb: number }\n"
            + _wb("  setTabColor(color: ColorInput | null): void;\n")
        )
        self.assert_rule(
            "R31.color",
            self.fixture.context("", ledger, wasm=bad),
            self.fixture.context("", ledger, wasm=good),
            [
                ("wasm", "record.cf_color", "-"),
                ("wasm", "record.font", "Font.colorArgb"),
                ("wasm", "record.spark_input", "SparkInput.colorSeries"),
                ("wasm", "record.spark_input", "SparkInput.fill"),
                ("wasm", "setTabColor", "color"),
            ],
        )

    def _parts(self, ctx: api_conventions.Context, body: str) -> api_conventions.Context:
        path = ctx.repo_root / "src" / "c_api" / "parts" / "cells.cpp"
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(body, encoding="utf-8")
        return ctx

    def test_r34_a_value_read_calls_the_gate(self) -> None:
        header = (
            "FM_API fm_status_t fm_sheet_get_value(const fm_workbook_t* wb, size_t sheet_index, fm_value_t* out);\n"
            "FM_API fm_status_t fm_sheet_get_formula(const fm_workbook_t* wb, size_t sheet_index, const char** out);\n"
        )
        ledger = {
            "entries": {
                "value.get": {"kind": "op", "c": "fm_sheet_get_value", "reads_values": True},
                "formula.get": {"kind": "op", "c": "fm_sheet_get_formula"},
            }
        }
        gate = "  if (auto st = parts::ensure_calculated(wb); st != FM_STATUS_OK) return st;\n"
        bad = (
            'extern "C" fm_status_t fm_sheet_get_value(const fm_workbook_t* wb, size_t sheet_index,\n'
            "                                       fm_value_t* out) {\n"
            "  // ensure_calculated(wb) is not called here\n"
            '  const char* s = "ensure_calculated(";\n  return read(wb, out);\n}\n'
            'extern "C" fm_status_t fm_sheet_get_formula(const fm_workbook_t* wb, size_t sheet_index, const char** out) {\n'
            + gate
            + "  return formula(wb, out);\n}\n"
        )
        good = (
            'extern "C" fm_status_t fm_sheet_get_value(const fm_workbook_t* wb, size_t sheet_index, fm_value_t* out) {\n'
            + gate
            + "  if (x) { return y; }\n  return read(wb, out);\n}\n"
            'extern "C" fm_status_t fm_sheet_get_formula(const fm_workbook_t* wb, size_t sheet_index, const char** out) {\n'
            "  return formula(wb, out);\n}\n"
        )
        bad_ctx = self._parts(self.fixture.context(header, ledger), bad)
        self.assert_rule(
            "R34.reads_values",
            bad_ctx,
            self._parts(self.fixture.context(header, ledger), good),
            [("c", "formula.get", "undeclared"), ("c", "value.get", "gate")],
        )
        sites = {
            v.discriminator: v.site
            for v in api_conventions.collect(bad_ctx, {"R34.reads_values": _rule("R34.reads_values")})
        }
        self.assertEqual(sites["gate"], "src/c_api/parts/cells.cpp:1")

    def test_r34_operators_do_not_end_a_body(self) -> None:
        header = (
            "FM_API fm_status_t fm_sheet_get_value(const fm_workbook_t* wb, size_t sheet_index, fm_value_t* out);\n"
            "FM_API fm_status_t fm_sheet_get_formula(const fm_workbook_t* wb, size_t sheet_index, const char** out);\n"
        )
        ledger = {
            "entries": {
                "value.get": {"kind": "op", "c": "fm_sheet_get_value", "reads_values": True},
                "formula.get": {"kind": "op", "c": "fm_sheet_get_formula"},
            }
        }
        body = (
            'extern "C" fm_status_t fm_sheet_get_formula(const fm_workbook_t* wb, size_t sheet_index, const char** out) {\n'
            "  if (wb->count < sheet_index || out->size > 2) return bad();\n  return formula(wb, out);\n}\n"
            'extern "C" fm_status_t fm_sheet_get_value(const fm_workbook_t* wb, size_t sheet_index, fm_value_t* out) {\n'
            "  if (auto st = parts::ensure_calculated(wb); st != FM_STATUS_OK) return st;\n  return read(wb, out);\n}\n"
        )
        ctx = self._parts(self.fixture.context(header, ledger), body)
        self.assertEqual(
            [v.discriminator for v in api_conventions.collect(ctx, {"R34.reads_values": _rule("R34.reads_values")})],
            [],
        )

    def test_r35_a_trailing_default_may_be_omitted(self) -> None:
        ledger = {
            "entries": {
                "precedents.get": {
                    "kind": "op",
                    "scope": "workbook",
                    "js": "getPrecedents",
                    "py": "get_precedents",
                    "params": {"depth": {"default": 1}},
                }
            }
        }
        bad_js = _wb("  getPrecedents(row: number, col: number, depth: number): string[];\n")
        bad_py = _py_wb("    def get_precedents(self, row: int, col: int, depth: int = 2) -> list: ...\n")
        good_js = _wb(
            "  getPrecedents(row: number, col: number): string[];\n"
            "  getPrecedents(row: number, col: number, depth: number): string[];\n"
        )
        good_py = _py_wb("    def get_precedents(self, row: int, col: int, depth: int = 1) -> list: ...\n")
        self.assert_rule(
            "R35.trailing_default",
            self.fixture.context("", ledger, wasm=bad_js, py_workbook=bad_py),
            self.fixture.context("", ledger, wasm=good_js, py_workbook=good_py),
            [("py", "precedents.get", "depth"), ("wasm", "precedents.get", "depth")],
        )
        optional = _wb("  getPrecedents(row: number, col: number, depth?: number | null): string[];\n")
        self.assertEqual(hits(self.fixture.context("", ledger, wasm=optional), "R35.trailing_default"), [])


def _rule(rule_id: str):
    return api_conventions.all_rules()[rule_id]


class PendingFileTest(unittest.TestCase):
    HEADER = (
        "FM_API fm_status_t fm_workbook_sheet_name(const fm_workbook_t* wb, size_t sheet_index);\n"
        "/** Reads the title. */\n" + TITLE_GETTER
    )

    def setUp(self) -> None:
        self._tmp = TemporaryDirectory()
        self.dir = Path(self._tmp.name) / "baseline"

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def ctx(self, header: str | None = None) -> api_conventions.Context:
        return api_conventions.Context(header if header is not None else self.HEADER, api_conventions.Ledger(None))

    def problems(self, header: str | None = None) -> list[str]:
        return api_conventions.check_api_conventions(self.dir, self.ctx(header))

    def violations(self, header: str | None = None) -> list[api_conventions.Violation]:
        return api_conventions.collect(self.ctx(header))

    def pending_path(self) -> Path:
        return self.dir / api_conventions.PENDING_DIR_NAME / "R33.lifetime.txt"

    def write_both(self) -> None:
        api_conventions.write_pending(self.violations(), self.dir, ["R33.lifetime"])
        api_conventions.write_baseline(self.violations(), self.dir)

    def test_a_listed_violation_passes_and_the_baseline_leaves_the_rule_out(self) -> None:
        self.write_both()
        self.assertEqual(self.problems(), [])
        self.assertNotIn("R33.lifetime", (self.dir / "c.txt").read_text(encoding="utf-8"))
        self.assertEqual(
            self.pending_path().read_text(encoding="utf-8").split("\t")[:3], ["c", "fm_workbook_get_title", "-"]
        )

    def test_an_unlisted_violation_fails(self) -> None:
        self.write_both()
        added = self.HEADER + "FM_API fm_status_t fm_workbook_get_author(const fm_workbook_t* wb, const char** out);\n"
        problems = self.problems(added)
        self.assertEqual(len(problems), 1, problems)
        self.assertIn("R33.lifetime [c] fm_workbook_get_author", problems[0])

    def test_a_pending_line_no_longer_violated_does_not_fail(self) -> None:
        self.write_both()
        fixed = self.HEADER.replace("Reads the title.", "Reads the title. Lifetime: transient.")
        self.assertEqual(self.problems(fixed), [])
        counts = api_conventions.counts_report(self.violations(fixed), self.dir, ["R33.lifetime"])
        self.assertIn("pending=0\tpending_clear=1", counts[0])

    def test_write_pending_removes_a_file_left_empty(self) -> None:
        self.write_both()
        fixed = self.HEADER.replace("Reads the title.", "Lifetime: static.")
        self.assertEqual(
            api_conventions.write_pending(self.violations(fixed), self.dir, ["R33.lifetime"]), {"R33.lifetime": 0}
        )
        self.assertFalse(self.pending_path().exists())

    def test_malformed_unknown_and_doubled_lines_fail(self) -> None:
        self.write_both()
        pending_dir = self.dir / api_conventions.PENDING_DIR_NAME
        (pending_dir / "R99.none.txt").write_text("c\top\t-\tm\n", encoding="utf-8")
        with self.pending_path().open("a", encoding="utf-8") as out:
            out.write("cpp\top\t-\tm\n")
        with (self.dir / "c.txt").open("a", encoding="utf-8") as out:
            out.write("R33.lifetime\tfm_workbook_get_title\t-\tm\n")
        problems = self.problems()
        self.assertTrue(any("is not `<surface>" in p for p in problems), problems)
        self.assertTrue(any("R99.none.txt names no registered rule" in p for p in problems), problems)
        self.assertTrue(any("is in both c.txt and pending/R33.lifetime.txt" in p for p in problems), problems)

    def test_the_violation_report_states_where_each_violation_stands(self) -> None:
        api_conventions.write_pending(self.violations(), self.dir, ["R33.lifetime"])
        states = {
            line.split("\t")[1]: line.split("\t")[0]
            for line in api_conventions.violations_report(self.violations(), self.dir)
        }
        self.assertEqual(states["R33.lifetime"], "pending")
        self.assertEqual(states["R2.verb"], "new")


if __name__ == "__main__":
    unittest.main()
