#!/usr/bin/env python3
"""Fixture tests for the consistency rules (R36-R44) and the ledger tools that carry their vocabulary.

Run from the repo root: `python3 -m tools.dev.binding_drift.api_conventions_consistency_test`.
"""

from __future__ import annotations

import contextlib
import importlib
import importlib.util
import io
import json
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from tools.dev.binding_drift import api_conventions
from tools.dev.binding_drift.api_conventions_bindings_test import BindingRuleTestBase, _py_wb, _wb, hits

TOOLS_DIR = Path(__file__).resolve().parents[2]


def _tool(directory: str, name: str):
    path = str(TOOLS_DIR / directory)
    if path not in sys.path:
        sys.path.insert(0, path)
    return importlib.import_module(name)


def _quiet(call, *args):
    """`call(*args)` with its progress lines kept off the test output."""
    with contextlib.redirect_stdout(io.StringIO()):
        return call(*args)


def _c_hits(header: str, rule_id: str) -> list[tuple[str, str]]:
    ctx = api_conventions.Context(header, api_conventions.Ledger({"schema": 1, "entries": {"fixture.present": {}}}))
    rules = {rule_id: api_conventions.all_rules()[rule_id]}
    return [(v.op_key, v.discriminator) for v in api_conventions.collect(ctx, rules)]


def _sheet(members: str) -> str:
    return "export interface Sheet {\n" + members + "}\n"


def _op(**fields: object) -> dict:
    return {"kind": "op", "scope": "sheet", "surfaces": ["c", "wasm", "native", "py"], **fields}


class LedgerRuleTest(BindingRuleTestBase):
    """R36, R37 and R44: rules the ledger answers, reported on the `ledger` surface."""

    def test_r36_a_patch_op_takes_the_verb_update(self) -> None:
        patch = {"patch": True, "patch_record": "PrintOptions"}
        bad = {
            "entries": {
                "print_options.set": _op(c="fm_sheet_set_print_options", js="setPrintOptions",
                                         py="set_print_options", **patch),
                "chart_axis.set": _op(c="fm_sheet_set_chart_axis", js="setChartAxis", py="set_chart_axis"),
                "chart_series.add": _op(c="fm_sheet_add_chart_series", js="addChartSeries", py="add_chart_series"),
            }
        }  # fmt: skip
        good = {
            "entries": {
                "print_options.update": _op(c="fm_sheet_update_print_options", js="updatePrintOptions",
                                            py="update_print_options", **patch),
                "chart_axis.update": _op(c="fm_sheet_update_chart_axis", js="updateChartAxis",
                                         py="update_chart_axis"),
                "chart_series.add": _op(c="fm_sheet_add_chart_series", js="addChartSeries", py="add_chart_series"),
            }
        }  # fmt: skip
        header = (
            "FM_API fm_status_t fm_sheet_{}_chart_axis(fm_workbook_t* wb, size_t sheet_index,"
            " const fm_chart_axis_patch_t* patch);\n"
            "FM_API fm_status_t fm_sheet_add_chart_series(fm_workbook_t* wb, size_t sheet_index,"
            " const fm_chart_series_patch_t* series);\n"
        )
        self.assert_rule(
            "R36.verb.update",
            self.fixture.context(header.format("set"), bad),
            self.fixture.context(header.format("update"), good),
            [("ledger", "chart_axis.set", "verb"), ("ledger", "print_options.set", "verb")],
        )

    def test_r36_reads_a_js_patch_argument_of_a_set(self) -> None:
        ledger = {"entries": {"view.set": _op(js="setView", py="set_view", surfaces=["wasm"])}}
        bad = _sheet("  setView(patch: SheetViewPatch): void;\n")
        good = _sheet("  setView(view: SheetView | null): void;\n")
        self.assert_rule(
            "R36.verb.update",
            self.fixture.context("", ledger, wasm=bad),
            self.fixture.context("", ledger, wasm=good),
            [("ledger", "view.set", "verb")],
        )

    @staticmethod
    def ledger_r37(**returns: str) -> dict:
        def add(name: str) -> dict:
            extra = {"add_returns": returns[name]} if name in returns else {}
            return _op(js=f"add{name.title()}", **extra)

        return {
            "collections": {
                "table": {"key": "index", "access": "handles", "handle_class": "Table", "plural": "tables",
                          "ops": ["table.add", "table.remove"]},
                "person": {"key": "natural", "by": "id", "plural": "people", "ops": ["person.add", "person.remove"]},
                "rule": {"key": "index", "access": "list", "plural": "rules", "ops": ["rule.add", "rule.remove"]},
            },
            "entries": {
                "table.add": add("table"),
                "person.add": add("person"),
                "rule.add": add("rule"),
                "table.remove": _op(js="removeTable"),
                "person.remove": _op(js="removePerson"),
                "rule.remove": _op(js="removeRule"),
            },
        }  # fmt: skip

    def test_r37_an_add_states_what_it_returns_and_matches_its_collection(self) -> None:
        self.assert_rule(
            "R37.add.returns",
            self.fixture.context("", self.ledger_r37(table="handle", person="index")),
            self.fixture.context("", self.ledger_r37(table="handle", person="key", rule="index")),
            [("ledger", "person.add", "remove"), ("ledger", "rule.add", "add_returns")],
        )

    def test_r44_a_singleton_is_removed_with_set_null(self) -> None:
        def ledger(*keys: str) -> dict:
            entries = {
                "theme.set": _op(js="setTheme"),
                "row_height.set": _op(js="setRowHeight"),
                "hyperlinks.clear": _op(js="clearHyperlinks"),
                "comment.add": _op(js="addComment"),
                "comment.remove": _op(js="removeComment"),
            }
            entries.update({key: _op(js=key.replace(".", "_")) for key in keys})
            return {"entries": entries}

        self.assert_rule(
            "R44.singleton.null",
            self.fixture.context("", ledger("theme.clear", "auto_filter.remove", "row_height.clear", "pinned.reset")),
            self.fixture.context("", ledger("row_height.reset")),
            [
                ("ledger", "auto_filter.remove", "remove"),
                ("ledger", "pinned.reset", "reset"),
                ("ledger", "row_height.clear", "clear"),
                ("ledger", "theme.clear", "clear"),
            ],
        )

    def test_r44_a_reset_is_satisfied_by_setters_on_every_part(self) -> None:
        def ledger(*part_verbs: str) -> dict:
            entries = {"palette.reset": _op(js="resetPalette")}
            entries.update({key: _op(js=key.replace(".", "_")) for key in part_verbs})
            return {"entries": entries}

        self.assert_rule(
            "R44.singleton.null",
            self.fixture.context("", ledger("palette_colors.set", "palette_fonts.get")),
            self.fixture.context("", ledger("palette_colors.set", "palette_fonts.set")),
            [("ledger", "palette.reset", "reset")],
        )


class BindingSurfaceRuleTest(BindingRuleTestBase):
    """R38, R39, R41 and R42: rules read off the declaration files and the Python package."""

    def test_r38_a_closed_set_is_numeric(self) -> None:
        bad = (
            "export type ProfileId = 'win-365-en_US' | 'mac-365-ja_JP';\n"
            "export type Kind = 'bar' | 'line';\n"
            "export interface Axis {\n  readonly position: 'nextTo' | 'low' | null;\n  readonly kind: Kind;\n"
            "  readonly profile: ProfileId;\n}\n"
            "export type AxisPatch = Patch<{\n  position: 'nextTo' | 'low';\n  title: string;\n}>;\n"
        )
        good = (
            "export type ProfileId = 'win-365-en_US' | 'mac-365-ja_JP';\n"
            "export interface Axis {\n  readonly position: TickLabelPosition | null;\n  readonly kind: ChartKind;\n"
            "  readonly profile: ProfileId;\n  readonly tag: 'axis';\n}\n"
            "export type AxisPatch = Patch<{\n  position: TickLabelPosition;\n  title: string;\n}>;\n"
        )
        self.assert_rule(
            "R38.enum.numeric",
            self.fixture.context("", None, wasm=bad, native=good),
            self.fixture.context("", None, wasm=good, native=good),
            [("wasm", "Axis", "kind"), ("wasm", "Axis", "position"), ("wasm", "AxisPatch", "position")],
        )

    LEDGER_R39 = {
        "entries": {
            "hyperlink.add": _op(js="addHyperlink", py="add_hyperlink", surfaces=["wasm", "py"]),
            "pinned_now.set": _op(scope="workbook", js="setPinnedNow", py="set_pinned_now", surfaces=["wasm", "py"]),
            "range.update": _op(js="updateRange", py="update_range", surfaces=["wasm", "py"],
                                ref_args=[0, "protectedRange"]),
        }
    }  # fmt: skip
    JS_R39 = (
        "export interface HyperlinkInput {\n  target: string;\n}\n"
        "export interface CivilTime {\n  year: number;\n}\n"
        + _wb("  setPinnedNow(now: CivilTime | null): void;\n")
        + _sheet(
            "  addHyperlink(row: number, col: number, link: HyperlinkInput): void;\n"
            "  updateRange(range: ProtectedRangeRef, patch: RangePatch): void;\n"
        )
    )
    PY_DATACLASS = "@dataclass\nclass HyperlinkInput:\n    target: str\n\n"

    def test_r39_python_inputs_are_one_dataclass_and_patches_keywords(self) -> None:
        bad_py = "class Sheet:\n" + (
            "    def add_hyperlink(self, row: int, col: int, *, target: str) -> None: ...\n"
            "    def update_range(self, protected_range: Union[int, str], name: str, *, password: str = None)"
            " -> None: ...\n"
        )
        good_py = "class Sheet:\n" + (
            "    def add_hyperlink(self, row: int, col: int, link: HyperlinkInput) -> None: ...\n"
            "    def update_range(self, protected_range: Union[int, str], *, name: str = UNCHANGED) -> None: ...\n"
        )
        workbook = _py_wb("    def set_pinned_now(self, now: Optional[datetime]) -> None: ...\n", self.PY_DATACLASS)
        self.assert_rule(
            "R39.py.args",
            self.fixture.context("", self.LEDGER_R39, wasm=self.JS_R39, py_workbook=workbook, py_sheet=bad_py),
            self.fixture.context("", self.LEDGER_R39, wasm=self.JS_R39, py_workbook=workbook, py_sheet=good_py),
            [
                ("py", "hyperlink.add", "input"),
                ("py", "range.update", "default:password"),
                ("py", "range.update", "positional:name"),
            ],
        )

    def test_r39_a_positional_key_is_a_field_of_the_patch_record(self) -> None:
        ledger = {
            "entries": {
                "plot.update": _op(js="updatePlot", py="update_plot", surfaces=["py"], patch=True,
                                   patch_record="ChartPlot"),
            }
        }  # fmt: skip
        record = "@dataclass(frozen=True)\nclass ChartPlot:\n    gap_width_pct: Optional[float]\n\n"
        bad = "class Sheet:\n    def update_plot(self, gap_width_pct: float, *, plot: int = UNCHANGED) -> None: ...\n"
        good = "class Sheet:\n    def update_plot(self, plot: int, *, gap_width_pct: float = UNCHANGED) -> None: ...\n"
        self.assert_rule(
            "R39.py.args",
            self.fixture.context("", ledger, py_workbook=record, py_sheet=bad),
            self.fixture.context("", ledger, py_workbook=record, py_sheet=good),
            [("py", "plot.update", "positional:gap_width_pct")],
        )

    def test_r41_protection_booleans_name_what_true_allows(self) -> None:
        ledger = {"entries": {"fixture.present": {"kind": "op"}}}
        cell = "export interface CellProtection {\n  readonly locked: boolean;\n  readonly hidden: boolean;\n}\n"
        bad_js = cell + (
            "export interface SheetProtection {\n  readonly formatCells: boolean;\n  readonly lockStructure: boolean;\n"
            "  readonly spinCount: number;\n}\n"
            "export interface SheetProtectionInput {\n  allowFormatCells?: boolean;\n}\n"
        )
        good_js = cell + (
            "export interface SheetProtection {\n  readonly allowFormatCells: boolean;\n"
            "  readonly lockStructure: boolean;\n  readonly spinCount: number;\n}\n"
            "export interface SheetProtectionInput {\n  allowFormatCells?: boolean;\n}\n"
        )
        bad_py = "@dataclass(frozen=True)\nclass SheetProtection:\n    format_cells: bool\n    spin_count: int\n"
        good_py = "@dataclass(frozen=True)\nclass SheetProtection:\n    allow_format_cells: Optional[bool]\n"
        self.assert_rule(
            "R41.bool.polarity",
            self.fixture.context("", ledger, wasm=bad_js, native=good_js, py_workbook=bad_py),
            self.fixture.context("", ledger, wasm=good_js, native=good_js, py_workbook=good_py),
            [
                ("py", "SheetProtection", "polarity:format_cells"),
                ("wasm", "SheetProtection", "polarity:formatCells"),
                ("wasm", "SheetProtectionInput", "name:allowFormatCells"),
            ],
        )

    def test_r42_a_dimension_carries_its_unit(self) -> None:
        header = (
            "typedef struct {\n  uint32_t struct_size;\n  double {w};\n  int32_t {w}_op;\n  int32_t has_{w};\n"
            "  uint32_t gap_count;\n} fm_line_patch_t;\n"
        )
        js = "export interface PageMargins {\n  readonly {left}: number;\n  readonly header: string;\n}\n" + _wb(
            "  setColumnWidth(col: number, {cw}: number): void;\n"
        )
        py = "@dataclass(frozen=True)\nclass Font:\n    {fs}: float\n    font_size_kind: int\n"

        def ctx(w: str, left: str, cw: str, fs: str, ledger: dict | None = None) -> api_conventions.Context:
            return self.fixture.context(
                header.replace("{w}", w),
                ledger or {"entries": {"fixture.present": {"kind": "op"}}},
                wasm=js.replace("{left}", left).replace("{cw}", cw),
                py_workbook=py.replace("{fs}", fs),
            )

        self.assert_rule(
            "R42.unit.suffix",
            ctx("width", "left", "width", "font_size"),
            ctx("width_pt", "leftInches", "widthChars", "font_size_pt"),
            [
                ("c", "fm_line_patch_t", "width"),
                ("py", "Font", "font_size"),
                ("wasm", "PageMargins", "left"),
                ("wasm", "Workbook", "setColumnWidth.width"),
            ],
        )
        listed = {"entries": {"fixture.present": {"kind": "op"}}, "unitless": ["width", "Left", "fontSize"]}
        self.assertEqual(hits(ctx("width", "left", "width", "font_size", listed), "R42.unit.suffix"), [])


class CHeaderRuleTest(unittest.TestCase):
    """R40 and R43: rules on the C header alone."""

    RECT = "typedef struct {\n  uint32_t first_row;\n} fm_cell_rect_t;\n"
    UPDATE = (
        "FM_API fm_status_t fm_sheet_update_shape(fm_workbook_t* wb, const fm_shape_t* spec,"
        " const fm_cell_rect_t* rect);\n"
    )

    def test_r40_a_patch_field_carries_a_patch_op(self) -> None:
        bad = (
            "/** @brief A negative field leaves that setting unchanged. */\n"
            "typedef struct {\n  uint32_t struct_size;\n  int32_t enabled;\n  int32_t width_op;\n  double width;\n"
            "  uint32_t fill_mask;\n  const fm_mention_t* mentions;\n  uint32_t mentions_count;\n"
            "  int32_t mentions_engaged;\n  int32_t has_max;\n  double max;\n} fm_iterative_patch_t;\n"
            "typedef struct {\n  int32_t name_op;\n  const char* name;\n} fm_shape_t;\n" + self.RECT + self.UPDATE
        )
        good = (
            "/** @brief Iterative settings patch. */\n"
            "typedef struct {\n  uint32_t struct_size;\n  fm_patch_op_t enabled_op;\n  int32_t enabled;\n"
            "  fm_patch_op_t width_op;\n  double width;\n  fm_patch_op_t mentions_op;\n"
            "  const fm_mention_t* mentions;\n  uint32_t mentions_count;\n  fm_line_patch_t line;\n"
            "} fm_iterative_patch_t;\n"
            "typedef struct {\n  fm_patch_op_t name_op;\n  const char* name;\n} fm_shape_t;\n" + self.RECT + self.UPDATE
        )
        self.assertEqual(
            _c_hits(bad, "R40.c.patch_op"),
            [
                ("fm_iterative_patch_t", "engaged:mentions_engaged"),
                ("fm_iterative_patch_t", "has:has_max"),
                ("fm_iterative_patch_t", "mask:fill_mask"),
                ("fm_iterative_patch_t", "negative-keep"),
                ("fm_iterative_patch_t", "op-type:width_op"),
                ("fm_iterative_patch_t", "op:enabled"),
                ("fm_iterative_patch_t", "op:max"),
                ("fm_iterative_patch_t", "op:mentions"),
                ("fm_shape_t", "op-type:name_op"),
            ],
        )
        self.assertEqual(_c_hits(good, "R40.c.patch_op"), [])

    def test_r43_a_status_is_documented(self) -> None:
        decl = "FM_API fm_status_t fm_sheet_{}(fm_workbook_t* wb, size_t sheet_index);\n"
        destroy = "/** Frees it. */\nFM_API void fm_workbook_destroy(fm_workbook_t* wb);\n"
        bad = destroy + "/** Reads the zoom. */\n" + decl.format("get_zoom") + decl.format("get_view")
        good = (
            destroy
            + "/** Reads the zoom.\n * @return `FM_STATUS_OK` (0). */\n"
            + decl.format("get_zoom")
            + "/** Reads the view; fails only with the common sheet failures. */\n"
            + decl.format("get_view")
        )
        self.assertEqual(_c_hits(bad, "R43.c.status_doc"), [("fm_sheet_get_view", "-"), ("fm_sheet_get_zoom", "-")])
        self.assertEqual(_c_hits(good, "R43.c.status_doc"), [])


class RegistryTest(unittest.TestCase):
    def test_the_consistency_rules_are_listed(self) -> None:
        expected = {
            "R36.verb.update", "R37.add.returns", "R38.enum.numeric", "R39.py.args", "R40.c.patch_op",
            "R41.bool.polarity", "R42.unit.suffix", "R43.c.status_doc", "R44.singleton.null",
        }  # fmt: skip
        self.assertLessEqual(expected, set(api_conventions.all_rules()))

    def test_a_ledger_violation_is_pending_or_new_never_baselined(self) -> None:
        violation = api_conventions.Violation("R44.singleton.null", "ledger", "theme.clear", "clear", "m")
        with TemporaryDirectory() as tmp:
            directory = Path(tmp)
            api_conventions.write_baseline([violation], directory)
            self.assertEqual(
                sorted(p.name for p in directory.glob("*.txt")), ["c.txt", "native.txt", "py.txt", "wasm.txt"]
            )
            self.assertEqual(len(api_conventions.compare([violation], directory, ["R44.singleton.null"])), 1)
            api_conventions.write_pending([violation], directory, ["R44.singleton.null"])
            self.assertEqual(api_conventions.compare([violation], directory, ["R44.singleton.null"]), [])
            pending = directory / api_conventions.PENDING_DIR_NAME / "R44.singleton.null.txt"
            self.assertEqual(pending.read_text(encoding="utf-8").split("\t")[0], "ledger")


class ValidateNamesVocabularyTest(unittest.TestCase):
    def setUp(self) -> None:
        self.validate = _tool("api", "validate_names")

    def schema(self, entries: dict, **top: object) -> list[str]:
        return self.validate.check_schema({"schema": 1, "collections": {}, "entries": entries, **top})

    def test_add_returns_clearable_unitless_and_enum_members_are_accepted(self) -> None:
        entries = {
            "rule.add": {"kind": "op", "scope": "sheet", "js": "addRule", "surfaces": ["wasm"], "add_returns": "index"},
            "record.legend": {"kind": "record", "role": "result", "js": "ChartLegend", "surfaces": ["wasm"],
                              "clearable": ["position", "hidden_series"]},
            "record.plot_patch": {"kind": "record", "js": "ChartPlotPatch", "surfaces": ["wasm"], "clearable": []},
            "enum.sort_order": {"kind": "enum", "js": "SortOrder", "py": "SortOrder", "surfaces": ["wasm", "py"]},
            "constant.sort_order_ascending": {"kind": "constant", "enum": "enum.sort_order", "js": "Ascending",
                                              "py": "ASCENDING", "surfaces": ["wasm", "py"],
                                              "was": {"js": "Asc", "py": "ASC"}},
        }  # fmt: skip
        self.assertEqual(self.schema(entries, unitless=["zoom", "gapWidth"]), [])

    def test_malformed_vocabulary_is_refused(self) -> None:
        entries = {
            "rule.add": {"kind": "op", "scope": "sheet", "js": "addRule", "surfaces": ["wasm"], "add_returns": "id"},
            "legend.get": {"kind": "op", "scope": "sheet", "js": "getLegend", "surfaces": ["wasm"],
                           "clearable": ["position"]},
            "record.legend": {"kind": "record", "js": "ChartLegend", "surfaces": ["wasm"],
                              "clearable": ["hiddenSeries", "hiddenSeries"]},
        }  # fmt: skip
        problems = self.schema(entries, unitless=["gapWidth", "gap_width"])
        self.assertTrue(any("add_returns is one of" in p for p in problems), problems)
        self.assertTrue(any("clearable belongs to a record" in p for p in problems), problems)
        self.assertTrue(any("clearable lists distinct field names" in p for p in problems), problems)
        self.assertTrue(any("unitless: gap_width is listed in two casings" in p for p in problems), problems)

    def test_ref_exempt_names_arguments_and_a_reason(self) -> None:
        def add(exempt: object) -> dict:
            return {"slicer.add": {"kind": "op", "scope": "sheet", "js": "addSlicer", "surfaces": ["wasm"],
                                   "ref_exempt": exempt}}  # fmt: skip

        self.assertEqual(self.schema(add({"params": ["table"], "reason": "a table on another sheet"})), [])
        for exempt in ({"params": ["table"]}, {"params": [], "reason": "x"}, ["table"]):
            problems = self.schema(add(exempt))
            self.assertTrue(any("ref_exempt is `{params: [name, ...], reason}`" in p for p in problems), problems)

    def test_one_entry_naming_a_field_alike_in_every_language_is_one_claim(self) -> None:
        entries = {
            "field.xf_vertical": {"kind": "field", "record": "record.xf", "c": "vertical", "js": "vertical",
                                  "py": "vertical", "surfaces": ["c", "wasm", "py"],
                                  "was": {"c": "vertical_align", "js": "verticalAlign", "py": "vertical_align"}},
        }  # fmt: skip
        self.assertEqual(self.validate.check_chains(entries), [])

    def test_clearable_names_fields_of_the_record_interface(self) -> None:
        entries = {
            "record.legend": {"kind": "record", "js": "ChartLegend", "surfaces": ["wasm"],
                              "clearable": ["position", "overlay_mode"]},
        }  # fmt: skip
        dts = "export interface ChartLegend {\n  readonly position: number | null;\n  readonly overlay: boolean;\n}\n"
        self.assertEqual(
            self.validate.check_clearable(entries, dts),
            ["entries.record.legend: clearable field overlay_mode is not a field of ChartLegend"],
        )

    def test_two_entries_renamed_to_one_name_are_refused(self) -> None:
        def member(enum: str, js: str, was: str) -> dict:
            return {"kind": "constant", "enum": enum, "js": js, "surfaces": ["wasm"], "was": {"js": was}}

        entries = {
            "a.set": {"kind": "op", "scope": "sheet", "js": "updateView", "surfaces": ["wasm"],
                      "was": {"js": "setView"}},
            "b.set": {"kind": "op", "scope": "sheet", "js": "updateView", "surfaces": ["wasm"],
                      "was": {"js": "setSheetView"}},
            "m.sort_asc": member("enum.sort", "Ascending", "Asc"),
            "m.sort_up": member("enum.sort", "Ascending", "Up"),
            "m.pivot_asc": member("enum.pivot_sort", "Ascending", "Asc"),
        }  # fmt: skip
        problems = self.validate.check_chains(entries)
        self.assertIn("js@sheet name updateView is the new name of several entries: a.set, b.set", problems)
        self.assertIn(
            "member:js:enum.sort name Ascending is the new name of several entries: m.sort_asc, m.sort_up", problems
        )
        self.assertFalse(any("enum.pivot_sort" in p for p in problems), problems)

    def test_an_enum_member_claims_no_symbol_name(self) -> None:
        entries = {
            "constant.a": {"kind": "constant", "enum": "enum.a", "js": "None", "surfaces": ["wasm"]},
            "constant.b": {"kind": "constant", "enum": "enum.b", "js": "None", "surfaces": ["wasm"]},
        }
        self.assertEqual(self.validate.check_ledger_index({"schema": 1, "collections": {}, "entries": entries}), [])


ENUM_LEDGER = {
    "schema": 1,
    "collections": {},
    "entries": {
        "enum.sort_order": {"kind": "enum", "js": "SortOrder", "py": "SortOrder", "surfaces": ["wasm", "py"],
                            "was": {"js": "SortDirection", "py": "SortOrder"}},
        "constant.sort_order_ascending": {"kind": "constant", "enum": "enum.sort_order", "js": "Ascending",
                                          "py": "ASCENDING", "surfaces": ["wasm", "py"],
                                          "was": {"js": "Asc", "py": "ASC"}},
        "record.legend": {"kind": "record", "js": "ChartLegend", "py": "ChartLegend", "surfaces": ["wasm", "py"],
                          "was": {"js": "Legend", "py": "Legend"}, "clearable": ["position"]},
        "field.legend_overlay": {"kind": "field", "record": "record.legend", "js": "overlay",
                                 "surfaces": ["wasm"], "was": {"js": "overlap"}},
    },
}  # fmt: skip


class ApplyRenamesMemberTest(unittest.TestCase):
    FILES = {
        "packages/npm/common.mjs": "const order = SortDirection.Asc;\nconst bare = Asc;\nx.Asc;\n",
        "src/wasm/formulon.d.ts": "/** Defaults to `SortOrder.Asc`. */\nexport const Asc = 0;\n",
        "docs/guide.md": "```js\nws.sort(SortDirection.Asc);\n```\n\n```python\nws.sort(SortOrder.ASC)\n```\n"
        "Prose SortOrder.ASC stays.\n",
        "packages/python/formulon/_sheet.py": "def f(order=SortOrder.ASC):\n    return ASC\n",
    }

    def setUp(self) -> None:
        self.apply = _tool("api", "apply_renames")
        self._tmp = TemporaryDirectory()
        self.root = Path(self._tmp.name)
        for rel, text in self.FILES.items():
            path = self.root / rel
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_enum_members_are_rewritten_where_qualified_and_checked(self) -> None:
        renames = self.apply.build_renames(ENUM_LEDGER)
        self.assertEqual(renames.js_members, {("SortDirection", "SortOrder"): {"Asc": "Ascending"}})
        self.assertNotIn("Asc", renames.js_types)
        self.assertEqual(_quiet(self.apply.cmd_check, [self.root], renames), 1)
        self.assertEqual(_quiet(self.apply.cmd_apply, [self.root], renames), 0)
        read = {rel: (self.root / rel).read_text(encoding="utf-8") for rel in self.FILES}
        self.assertEqual(
            read["packages/npm/common.mjs"], "const order = SortOrder.Ascending;\nconst bare = Asc;\nx.Asc;\n"
        )
        self.assertIn("`SortOrder.Ascending`", read["src/wasm/formulon.d.ts"])
        self.assertIn("export const Asc = 0;", read["src/wasm/formulon.d.ts"])
        self.assertIn("ws.sort(SortOrder.Ascending);", read["docs/guide.md"])
        self.assertIn("ws.sort(SortOrder.ASCENDING)", read["docs/guide.md"])
        self.assertIn("Prose SortOrder.ASC stays.", read["docs/guide.md"])
        self.assertEqual(
            read["packages/python/formulon/_sheet.py"], "def f(order=SortOrder.ASCENDING):\n    return ASC\n"
        )
        self.assertEqual(_quiet(self.apply.cmd_check, [self.root], renames), 0)


class MigrationTableTest(unittest.TestCase):
    def test_fields_and_enum_members_have_rows(self) -> None:
        table = _tool("api", "gen_migration_table").rows(ENUM_LEDGER)
        self.assertIn(("`SortDirection.Asc`", "`SortOrder.Ascending`"), table["js"])
        self.assertIn(("`SortOrder.ASC`", "`SortOrder.ASCENDING`"), table["py"])
        self.assertIn(("`ChartLegend.overlap`", "`ChartLegend.overlay`"), table["js"])
        self.assertIn(("`Legend`", "`ChartLegend`"), table["js"])


class PatchClearCasesTest(unittest.TestCase):
    DTS = (
        "export interface ChartLegend {\n  readonly position: number | null;\n  readonly overlay: boolean;\n"
        "  readonly hiddenSeries: readonly number[];\n}\n"
    )

    def test_the_table_splits_each_record_into_clearable_and_not(self) -> None:
        gen = _tool("api", "gen_patch_clear_cases")
        table = gen.cases(ENUM_LEDGER, self.DTS)
        self.assertEqual(
            table, {"ChartLegend": {"clearable": ["position"], "non_clearable": ["hidden_series", "overlay"]}}
        )
        with TemporaryDirectory() as tmp:
            ledger = Path(tmp) / "names.yaml"
            dts = Path(tmp) / "formulon.d.ts"
            out = Path(tmp) / "patch_clear_cases.json"
            ledger.write_text(json.dumps(ENUM_LEDGER), encoding="utf-8")
            dts.write_text(self.DTS, encoding="utf-8")
            args = ["--ledger", str(ledger), "--dts", str(dts), "--output", str(out)]
            self.assertEqual(_quiet(gen.main, [*args, "--check"]), 1)
            self.assertEqual(_quiet(gen.main, [*args, "--write"]), 0)
            self.assertEqual(_quiet(gen.main, [*args, "--check"]), 0)
            self.assertEqual(json.loads(out.read_text(encoding="utf-8")), table)


class StatusRangeTest(unittest.TestCase):
    ERROR_H = (
        "enum class FormulonErrorCode : int32_t {\n"
        "  // ===== 0-999: General =====\n  kOk = 0,\n"
        "  // ===== 1000-1999: Parser / Tokenizer =====\n  kParserX = 1000,\n"
        "  // ===== 2000-2999: Evaluator =====\n  // ===== 3000-3999: Functions =====\n"
        "  // ===== 4000-4999: Dependency graph / Recalc =====\n  // ===== 5000-5999: I/O (OOXML) =====\n"
        "  // ----- Workbook structural mutation (5050-5069) -----\n"
        "  // ===== 6000-6999: Crypto / Security =====\n  // ===== 7000-7999: Bindings / C API =====\n"
        "  // ===== 8000-8999: Pivot / Advanced =====\n  // ===== 9000-9999: UI integration / Printing =====\n"
        "};\n"
    )

    def test_each_module_range_maps_to_its_exception_class(self) -> None:
        gen = _tool("codegen", "gen_status_codes")
        ranges = gen.parse_ranges(self.ERROR_H)
        self.assertEqual(ranges[0], (0, 999, "FormulonError"))
        self.assertEqual(ranges[5], (5000, 5999, "IoError"))
        self.assertEqual(ranges[-1], (9000, 9999, "PrintError"))
        text = gen.render_py([("Ok", "OK", 0), ("NotFound", "NOT_FOUND", 6)], ranges)
        self.assertIn("NOT_FOUND_CODES", text)
        region = text[text.index(gen.RANGES_BEGIN) : text.index(gen.RANGES_END) + len(gen.RANGES_END)]
        with TemporaryDirectory() as tmp:
            path = Path(tmp) / "status_ranges.py"
            path.write_text(region + "\n", encoding="utf-8")
            spec = importlib.util.spec_from_file_location("status_ranges", path)
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
        name = module.range_class_name
        self.assertEqual(
            [name(2), name(1003), name(4005), name(5050), name(6001), name(7001), name(8000), name(9999)],
            ["FormulonError", "FormulaSyntaxError", "RecalcError", "IoError", "SecurityError", "BindingError",
             "PivotError", "PrintError"],
        )  # fmt: skip
        self.assertEqual(name(10000), "FormulonError")

    def test_an_unknown_range_is_refused(self) -> None:
        gen = _tool("codegen", "gen_status_codes")
        with self.assertRaises(SystemExit):
            gen.parse_ranges(self.ERROR_H.replace("Crypto / Security", "Telemetry"))


if __name__ == "__main__":
    unittest.main()
