#!/usr/bin/env python3
"""Fixture tests for the ledger-bound API convention rules.

Run from the repo root: `python3 -m tools.dev.binding_drift.api_conventions_bindings_test`.
"""

from __future__ import annotations

import copy
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from tools.dev.binding_drift import api_conventions, api_conventions_bindings
from tools.dev.binding_drift.surface_files import _SKIPPED

WASM = "src/wasm/formulon.d.ts"
NATIVE = "packages/npm-native/index.d.ts"
PY_INIT = "packages/python/formulon/__init__.py"
PY_WORKBOOK = "packages/python/formulon/workbook.py"

RECT = (
    "typedef struct {\n  uint32_t first_row;\n  uint32_t first_col;\n  uint32_t last_row;\n  uint32_t last_col;\n"
    "} fm_cell_rect_t;\n"
)

# The static rule-ids of the checker, ledger-free and ledger-bound together.
STATIC_RULE_IDS = {
    "R0.coverage", "R1.form", "R1.scope", "R2.verb", "R3.suffix", "R3.keys", "R4.unique", "R5.derive",
    "R5.list", "R6.c", "R6.js", "R6.py", "R7.c.byptr", "R7.c.typedef", "R7.js", "R7.py", "R8.c", "R8.js",
    "R8.py", "R10.c.doc", "R11.axis", "R11.rect", "R11.multi", "R11.sheet", "R12",
}  # fmt: skip


def _wb(members: str) -> str:
    return "export interface Workbook {\n" + members + "}\n"


def _py_wb(methods: str, prelude: str = "") -> str:
    return prelude + "class Workbook:\n" + (methods or "    pass\n")


class Fixture:
    """A throwaway repo root holding the binding surfaces a rule reads."""

    def __init__(self, tmp: Path) -> None:
        self.tmp = tmp
        self.count = 0

    def context(self, header: str = "", ledger: dict | None = None, **files: str) -> api_conventions.Context:
        self.count += 1
        root = self.tmp / str(self.count)
        paths = {"wasm": WASM, "native": NATIVE, "py_init": PY_INIT, "py_workbook": PY_WORKBOOK}
        defaults = {"py_init": "__all__ = []\n"}
        for name, rel in paths.items():
            path = root / rel
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(files.get(name, defaults.get(name, "")), encoding="utf-8")
        data = {"schema": 1, "collections": {}, **(ledger or {})}
        data.setdefault("entries", {"fixture.present": {"kind": "op"}})
        return api_conventions.Context(header, api_conventions.Ledger(data), root)


def hits(ctx: api_conventions.Context, rule_id: str) -> list[tuple[str, str, str]]:
    rules = {rule_id: api_conventions.all_rules()[rule_id]}
    return [(v.surface, v.op_key, v.discriminator) for v in api_conventions.collect(ctx, rules)]


class BindingRuleTestBase(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = TemporaryDirectory()
        self.fixture = Fixture(Path(self._tmp.name))

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def assert_rule(self, rule_id: str, bad: api_conventions.Context, good: api_conventions.Context, expected) -> None:
        found = hits(bad, rule_id)
        self.assertEqual(found, sorted(expected))
        self.assertTrue(found)
        self.assertEqual(hits(good, rule_id), [])


class RegistryTest(unittest.TestCase):
    def test_rule_ids_are_exactly_the_static_set(self) -> None:
        self.assertEqual(set(api_conventions.all_rules()), STATIC_RULE_IDS)

    def test_without_a_ledger_the_bound_rules_are_skipped_visibly(self) -> None:
        ctx = api_conventions.Context("FM_API fm_status_t fm_workbook_sheet_name(const fm_workbook_t* wb);\n")
        bound = {r: fn for r, fn in api_conventions.all_rules().items() if fn.__module__.endswith("_bindings")}
        self.assertEqual(len(bound), 16)
        self.assertEqual(api_conventions.collect(ctx, bound), [])
        self.assertTrue(any("no name ledger" in note for note in _SKIPPED))


class DeclarationParserTest(BindingRuleTestBase):
    def test_typescript_members_params_and_returns(self) -> None:
        js = api_conventions_bindings.JsSurface(
            "/** doc with (parens); */\n"
            "export interface Workbook {\n"
            "  /** Sets a sink. */\n"
            "  setLogSink(sink: ((record: Uint8Array) => void) | null,\n    level?: number): Status;\n"
            "  getRange<T extends string = string>(sheet: number, opts: { a: number; b: string }): ListResult<T>;\n"
            "  [Symbol.dispose](): void;\n"
            "}\n"
            "export interface WorkbookCtor {\n  load(bytes: Uint8Array): OptionalResult<Workbook>;\n}\n"
            "export interface Cell { readonly row: number; col?: number }\n"
            "export type Patch<T> = { [K in keyof T]?: T[K] | null };\n"
            "export const StatusCode: Readonly<{ Ok: 0; Err: 1 }>;\n"
            "export function evalFormula(formula: string): EvalResult;\n"
        )
        sink = js.workbook["setLogSink"]
        self.assertEqual(sink.params, (("sink", "((record: Uint8Array) => void) | null"), ("level", "number")))
        self.assertEqual(sink.type, "Status")
        self.assertEqual(js.workbook["getRange"].type, "ListResult<T>")
        self.assertEqual(js.workbook["getRange"].params[1], ("opts", "{ a: number; b: string }"))
        self.assertIn("[Symbol.dispose]", js.workbook)
        self.assertEqual(
            set(js.callables()), {"setLogSink", "getRange", "[Symbol.dispose]", "Workbook.load", "evalFormula"}
        )
        self.assertEqual([(m.name, m.readonly) for m in js.decls["Cell"].members], [("row", True), ("col", False)])
        self.assertEqual(js.decls["Patch"].text, "{ [K in keyof T]?: T[K] | null }")
        self.assertEqual(api_conventions_bindings._constant_tables(js), ["StatusCode"])

    def test_python_workbook_methods_and_exports(self) -> None:
        py = api_conventions_bindings.PySurface(
            {
                "__init__.py": "__all__ = ['Workbook', 'get_version']\n",
                "workbook.py": "def get_version() -> str: ...\n"
                + _py_wb(
                    "    @classmethod\n    def load(cls, data: bytes) -> 'Workbook': ...\n"
                    "    def get_zoom(self, sheet: int) -> int: ...\n    def _private(self) -> None: ...\n"
                ),
            }
        )
        self.assertEqual(set(py.workbook), {"Workbook.load", "get_zoom"})
        self.assertEqual(py.workbook["get_zoom"].params, (("sheet", "int"),))
        self.assertEqual(set(py.symbols()), {"Workbook", "get_version", "Workbook.load", "get_zoom"})


class CoverageTest(BindingRuleTestBase):
    LEDGER = {
        "entries": {
            "workbook.class": {
                "kind": "record",
                "js": "Workbook",
                "py": "Workbook",
                "surfaces": ["wasm", "native", "py"],
            },
            "sheet.get_name": {
                "kind": "op",
                "c": "fm_workbook_get_sheet_name",
                "js": "getSheetName",
                "py": "get_sheet_name",
                "surfaces": ["c", "wasm", "native", "py"],
                "was": {"c": "fm_workbook_sheet_name", "js": "sheetName", "py": "sheet_name"},
            },
            "hyperlink.get_at": {
                "kind": "op",
                "js": "getHyperlinkAt",
                "surfaces": ["c"],
                "drop": {"js": "getHyperlinkAt", "closes": "R5.list"},
            },
            "memory_usage.get": {
                "kind": "op",
                "js": "getMemoryUsage",
                "py": "get_memory_usage",
                "surfaces": ["native"],
            },
        }
    }

    def test_r0_coverage(self) -> None:
        decl = "FM_API fm_status_t fm_workbook_{}(const fm_workbook_t* wb, size_t sheet_index);\n"
        bad = self.fixture.context(
            decl.format("sheet_name") + decl.format("orphan"),
            self.LEDGER,
            wasm=_wb("  sheetName(sheet: number): StringResult;\n  getHyperlinkAt(i: number): Status;\n"),
            native=_wb("  getSheetName(sheet: number): StringResult;\n  extra(): void;\n"),
            py_init="__all__ = ['Workbook']\n",
            py_workbook=_py_wb(
                "    def get_sheet_name(self, sheet: int) -> str: ...\n    def get_memory_usage(self) -> int: ...\n"
            ),
        )
        good = self.fixture.context(
            decl.format("get_sheet_name"),
            self.LEDGER,
            wasm=_wb("  getSheetName(sheet: number): StringResult;\n  getHyperlinkAt(i: number): Status;\n"),
            native=_wb("  getSheetName(sheet: number): StringResult;\n  getMemoryUsage(): number;\n"),
            py_init="__all__ = ['Workbook']\n",
            py_workbook=_py_wb("    def get_sheet_name(self, sheet: int) -> str: ...\n"),
        )
        self.assert_rule(
            "R0.coverage",
            bad,
            good,
            [
                ("c", "fm_workbook_orphan", "fm_workbook_orphan"),
                ("c", "sheet.get_name", "fm_workbook_sheet_name"),
                ("native", "extra", "extra"),
                ("py", "memory_usage.get", "get_memory_usage"),
                ("wasm", "sheet.get_name", "sheetName"),
            ],
        )

    def test_header_constants_and_replaced_types(self) -> None:
        ledger = {
            "entries": {
                "record.merge_range": {"kind": "typedef", "c": "fm_merge_range", "replaced_by": "record.cell_rect"},
                "constant.a": {"kind": "constant", "c": "FM_MODE_A", "surfaces": ["c"]},
            }
        }
        header = (
            "typedef struct {\n  uint32_t row;\n} fm_merge_range;\n"
            "typedef enum { FM_MODE_A = 0, FM_MODE_B = 1 } fm_mode_t;\n#define FM_PHONETIC_LEFT 1u\n"
        )
        found = hits(self.fixture.context(header, ledger), "R0.coverage")
        self.assertEqual(
            found,
            [
                ("c", "FM_MODE_B", "FM_MODE_B"),
                ("c", "FM_PHONETIC_LEFT", "FM_PHONETIC_LEFT"),
                ("c", "fm_mode_t", "fm_mode_t"),
            ],
        )


class ScopeTest(BindingRuleTestBase):
    LEDGER = {
        "entries": {
            "sheet.zoom.set": {"kind": "op", "c": "fm_sheet_set_zoom", "scope": "sheet", "noun": "zoom"},
            "tab_color.set": {"kind": "op", "c": "fm_workbook_set_tab_color", "scope": "workbook", "noun": "tab_color"},
            "sheet.remove": {"kind": "op", "c": "fm_workbook_remove_sheet", "scope": "workbook", "noun": "sheet"},
            "chart.get_title": {"kind": "op", "c": "fm_chart_get_title", "scope": "chart", "noun": "title"},
        }
    }

    def test_r1_scope(self) -> None:
        remove = "FM_API fm_status_t fm_workbook_remove_sheet(fm_workbook_t* wb, size_t sheet_index);\n"
        bad = self.fixture.context(
            "FM_API fm_status_t fm_sheet_set_zoom(fm_workbook_t* wb, uint32_t sheet, uint32_t zoom);\n"
            "FM_API fm_status_t fm_workbook_set_tab_color(fm_workbook_t* wb, size_t sheet_index, uint32_t c);\n"
            "FM_API fm_status_t fm_chart_get_title(const fm_workbook_t* wb, const char** out_title);\n" + remove,
            self.LEDGER,
        )
        good = self.fixture.context(
            "FM_API fm_status_t fm_sheet_set_zoom(fm_workbook_t* wb, size_t sheet_index, uint32_t zoom);\n"
            "FM_API fm_status_t fm_workbook_set_tab_color(fm_workbook_t* wb, uint32_t c);\n"
            "FM_API fm_status_t fm_chart_get_title(const fm_chart_t* chart, const char** out_title);\n" + remove,
            self.LEDGER,
        )
        self.assert_rule(
            "R1.scope",
            bad,
            good,
            [
                ("c", "chart.get_title", "handle"),
                ("c", "sheet.zoom.set", "sheet_index"),
                ("c", "tab_color.set", "scope"),
            ],
        )

    def test_r1_scope_sheet_arg_exempts_a_workbook_scope_sheet_index(self) -> None:
        def ledger(sheet_arg: bool) -> dict:
            add = {"kind": "op", "c": "fm_workbook_add_table", "scope": "workbook", "noun": "table"}
            if sheet_arg:
                add["sheet_arg"] = "tables are indexed across the workbook"
            return {"entries": {"table.add": add}}

        header = "FM_API fm_status_t fm_workbook_add_table(fm_workbook_t* wb, size_t sheet_index, size_t* out_index);\n"
        self.assert_rule(
            "R1.scope",
            self.fixture.context(header, ledger(False)),
            self.fixture.context(header, ledger(True)),
            [("c", "table.add", "scope")],
        )

    def test_r1_scope_buffer_handle_is_the_bytes_pointer(self) -> None:
        ledger = {
            "entries": {"buffer.destroy": {"kind": "op", "c": "fm_buffer_destroy", "scope": "buffer", "noun": "buffer"}}
        }
        self.assert_rule(
            "R1.scope",
            self.fixture.context("FM_API void fm_buffer_destroy(char* bytes);\n", ledger),
            self.fixture.context("FM_API void fm_buffer_destroy(uint8_t* bytes);\n", ledger),
            [("c", "buffer.destroy", "handle")],
        )


class CollectionTest(BindingRuleTestBase):
    def test_r3_suffix(self) -> None:
        ledger = {
            "collections": {
                "table": {"key": "index", "ops": ["table.get", "table.get_count"]},
                "comment": {"key": "natural", "ops": ["comment.get_at", "comment.get"]},
            },
            "entries": {
                "table.get": {"kind": "op", "c": "fm_workbook_get_table", "was": {"c": "fm_workbook_get_table_at"}},
                "comment.get_at": {"kind": "op", "c": "fm_sheet_get_comment_at"},
                "note.get_count": {"kind": "op", "c": "fm_sheet_get_note_count"},
            },
        }
        bad = self.fixture.context(
            "FM_API fm_status_t fm_workbook_get_table_at(const fm_workbook_t* wb, size_t index);\n"
            "FM_API fm_status_t fm_sheet_get_comment_at(const fm_workbook_t* wb, size_t sheet_index, size_t position);\n"
            "FM_API fm_status_t fm_sheet_get_note_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count);\n",
            ledger,
        )
        good = self.fixture.context(
            "FM_API fm_status_t fm_workbook_get_table(const fm_workbook_t* wb, size_t table_index);\n"
            "FM_API fm_status_t fm_sheet_get_comment_at(const fm_workbook_t* wb, size_t sheet_index, size_t index);\n",
            ledger,
        )
        self.assert_rule(
            "R3.suffix",
            bad,
            good,
            [("c", "comment.get_at", "at"), ("c", "note.get_count", "collection"), ("c", "table.get", "at")],
        )

    def test_r3_keys(self) -> None:
        ledger = {
            "collections": {
                "comment": {"key": "natural", "ops": ["comment.get_at", "comment.get", "comment.remove"]},
                "table": {"key": "index", "ops": ["table.add"]},
                "slicer": {"key": "natural", "ops": ["slicer.add", "slicer.update", "slicer.get"]},
            },
            "entries": {
                "comment.get_at": {"kind": "op", "c": "fm_sheet_get_comment_at"},
                "comment.get": {"kind": "op", "c": "fm_sheet_get_comment"},
                "comment.remove": {"kind": "op", "c": "fm_sheet_remove_comment"},
                "table.add": {"kind": "op", "c": "fm_sheet_add_table"},
                "slicer.add": {"kind": "op", "c": "fm_sheet_add_slicer"},
                "slicer.update": {"kind": "op", "c": "fm_sheet_update_slicer"},
                "slicer.get": {"kind": "op", "c": "fm_sheet_get_slicer"},
            },
        }
        head = "FM_API fm_status_t fm_sheet_"
        shared = (
            head + "get_comment_at(const fm_workbook_t* wb, size_t sheet_index, size_t index);\n"
            + head + "update_slicer(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, const char* name);\n"
            + head + "get_slicer(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, const char** out);\n"
        )  # fmt: skip
        bad = self.fixture.context(
            shared
            + head + "remove_comment(fm_workbook_t* wb, size_t sheet_index, size_t index);\n"
            + head + "add_table(fm_workbook_t* wb, size_t sheet_index, const char* name);\n"
            + head + "add_slicer(fm_workbook_t* wb, size_t sheet_index, const char* name);\n",
            ledger,
        )  # fmt: skip
        good = self.fixture.context(
            shared
            + head + "get_comment(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col, const char** out);\n"
            + head + "remove_comment(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col);\n"
            + head + "add_table(fm_workbook_t* wb, size_t sheet_index, const char* name, size_t* out_table_index);\n"
            + head + "add_slicer(fm_workbook_t* wb, size_t sheet_index, const char* name, uint32_t* out_object_id);\n",
            ledger,
        )  # fmt: skip
        self.assert_rule(
            "R3.keys",
            bad,
            good,
            [
                ("c", "comment", "getter"),
                ("c", "comment.remove", "index"),
                ("c", "slicer.add", "out-key"),
                ("c", "table.add", "out-key"),
            ],
        )

    def test_r3_keys_no_getter_exempts_the_getter(self) -> None:
        def ledger(no_getter: bool) -> dict:
            merge = {"key": "natural", "ops": ["merge.get_at"]}
            if no_getter:
                merge["no_getter"] = "the element is its rectangle"
            return {
                "collections": {"merge": merge},
                "entries": {"merge.get_at": {"kind": "op", "c": "fm_sheet_get_merge_at"}},
            }

        header = (
            "FM_API fm_status_t fm_sheet_get_merge_at(const fm_workbook_t* wb, size_t sheet_index, size_t index);\n"
        )
        self.assert_rule(
            "R3.keys",
            self.fixture.context(header, ledger(False)),
            self.fixture.context(header, ledger(True)),
            [("c", "merge", "getter")],
        )


class NounTest(BindingRuleTestBase):
    def test_r4_unique(self) -> None:
        def ledger(qualified: bool) -> dict:
            cell = {
                "kind": "op",
                "c": "fm_sheet_get_font",
                "scope": "sheet",
                "noun": "cell_font" if qualified else "font",
            }
            return {
                "entries": {
                    "font.get": {"kind": "op", "c": "fm_workbook_get_font", "scope": "workbook", "noun": "font"},
                    "cell_font.get": cell,
                }
            }

        decl = "FM_API fm_status_t fm_pagination_get_{}_count(const fm_pagination_t* p, size_t* out_count);\n"
        bad = self.fixture.context(
            decl.format("horizontal_break"), ledger(False), wasm=_wb("  getHorizontalBreaks(): ListResult<number>;\n")
        )
        good = self.fixture.context(
            decl.format("row_break"), ledger(True), wasm=_wb("  getRowBreaks(): ListResult<number>;\n")
        )
        self.assert_rule(
            "R4.unique",
            bad,
            good,
            [
                ("c", "cell_font.get", "noun"),
                ("c", "fm_pagination_get_horizontal_break_count", "horizontal_break"),
                ("c", "font.get", "noun"),
                ("wasm", "getHorizontalBreaks", "horizontal_break"),
            ],
        )


class DerivationTest(BindingRuleTestBase):
    LEDGER = {
        "entries": {
            "sheet.get_name": {
                "kind": "op",
                "c": "fm_workbook_get_sheet_name",
                "js": "getSheetName",
                "py": "get_sheet_name",
                "was": {"c": "fm_workbook_sheet_name", "js": "sheetName", "py": "sheet_name"},
            },
            "record.hyperlink": {
                "kind": "record",
                "c": "fm_hyperlink_t",
                "js": "Hyperlink",
                "was": {"js": "HyperlinkEntry"},
            },
            "workbook.load": {
                "kind": "factory",
                "c": "fm_workbook_load",
                "js": "Workbook.load",
                "was": {"js": "Workbook.loadBytes"},
            },
            "memory.get": {"kind": "op", "c": "fm_workbook_get_memory", "js": "memoryStats", "exception": "fixture"},
        }
    }
    HEADER = (
        "typedef struct {\n  uint32_t row;\n} fm_hyperlink_t;\n"
        "FM_API fm_status_t fm_workbook_get_sheet_name(const fm_workbook_t* wb, size_t sheet_index);\n"
        "FM_API fm_status_t fm_workbook_load(const uint8_t* bytes, size_t len, fm_workbook_t** out);\n"
        "FM_API fm_status_t fm_workbook_get_memory(const fm_workbook_t* wb, size_t* out);\n"
    )

    def test_r5_derive(self) -> None:
        bad = self.fixture.context(
            self.HEADER,
            self.LEDGER,
            wasm=_wb("  sheetName(sheet: number): StringResult;\n  memoryStats(): number;\n")
            + "export interface WorkbookCtor {\n  loadBytes(b: Uint8Array): Workbook;\n}\n"
            + "export interface HyperlinkEntry { readonly row: number }\n",
            py_init="__all__ = []\n",
            py_workbook=_py_wb("    def sheet_name(self, sheet: int) -> str: ...\n"),
        )
        good = self.fixture.context(
            self.HEADER,
            self.LEDGER,
            wasm=_wb("  getSheetName(sheet: number): StringResult;\n  memoryStats(): number;\n")
            + "export interface WorkbookCtor {\n  load(b: Uint8Array): Workbook;\n}\n"
            + "export interface Hyperlink { readonly row: number }\n",
            py_workbook=_py_wb("    def get_sheet_name(self, sheet: int) -> str: ...\n"),
        )
        self.assert_rule(
            "R5.derive",
            bad,
            good,
            [
                ("py", "sheet.get_name", "-"),
                ("wasm", "record.hyperlink", "-"),
                ("wasm", "sheet.get_name", "-"),
                ("wasm", "workbook.load", "-"),
            ],
        )

    def test_r5_list(self) -> None:
        def ledger(py_name: str) -> dict:
            return {
                "collections": {
                    "hyperlink": {
                        "key": "natural",
                        "plural": "hyperlinks",
                        "ops": ["hyperlink.get_count", "hyperlink.get_at"],
                    }
                },
                "entries": {
                    "hyperlink.get_at": {"kind": "op", "drop": {"js": "getHyperlinkAt", "py": "get_hyperlink_at"}},
                    "hyperlinks.list": {
                        "kind": "list",
                        "from": ["hyperlink.get_count", "hyperlink.get_at"],
                        "js": "listHyperlinks",
                        "py": py_name,
                        "surfaces": ["wasm", "native", "py"],
                    },
                },
            }

        bad = self.fixture.context(
            "",
            ledger("list_hyperlink"),
            wasm=_wb("  getHyperlinkAt(sheet: number, index: number): Status;\n"),
            py_workbook=_py_wb("    def list_hyperlink(self, sheet: int) -> list: ...\n"),
        )
        listed = _wb("  listHyperlinks(sheet: number): ListResult<Hyperlink>;\n")
        good = self.fixture.context(
            "",
            ledger("list_hyperlinks"),
            wasm=listed,
            native=listed,
            py_workbook=_py_wb("    def list_hyperlinks(self, sheet: int) -> list: ...\n"),
        )
        self.assert_rule(
            "R5.list",
            bad,
            good,
            [
                ("native", "hyperlinks.list", "missing"),
                ("py", "hyperlinks.list", "name"),
                ("wasm", "hyperlink.get_at", "getHyperlinkAt"),
                ("wasm", "hyperlinks.list", "missing"),
            ],
        )

    def test_a_process_list_may_be_a_module_function(self) -> None:
        ledger = {
            "collections": {"function": {"key": "natural", "plural": "function_names", "ops": ["function.get_count"]}},
            "entries": {
                "function_names.list": {
                    "kind": "list",
                    "from": ["function.get_count"],
                    "py": "list_function_names",
                    "surfaces": ["py"],
                },
            },
        }
        ctx = self.fixture.context(
            "",
            ledger,
            py_init="__all__ = ['list_function_names']\n",
            py_workbook="def list_function_names() -> list: ...\n",
        )
        self.assertEqual(hits(ctx, "R5.list"), [])

    def test_a_drop_reports_under_the_rule_it_closes(self) -> None:
        ledger = {"entries": {"cf.get_at": {"kind": "op", "drop": {"js": "getCfAt", "closes": "R12"}}}}
        ctx = self.fixture.context(
            "", ledger, wasm=_wb("  getCfAt(i: number): Status;\n"), native=_wb("  getCfAt(i: number): Status;\n")
        )
        self.assertEqual(hits(ctx, "R5.list"), [])
        self.assertEqual(hits(ctx, "R12"), [("native", "cf.get_at", "getCfAt"), ("wasm", "cf.get_at", "getCfAt")])
        self.assertNotIn(("wasm", "cf.get_at", "getCfAt"), hits(ctx, "R0.coverage"))


class RoleSuffixDerivationTest(BindingRuleTestBase):
    LEDGER = {
        "entries": {
            "record.print_options": {
                "kind": "record",
                "role": "input",
                "c": "fm_print_options_t",
                "js": "PrintOptionsInput",
                "was": {"js": "PrintOptions"},
            },
        }
    }
    HEADER = "typedef struct {\n  uint32_t row;\n} fm_print_options_t;\n"

    def test_r5_derive_applies_the_role_suffix(self) -> None:
        bad = self.fixture.context(
            self.HEADER, self.LEDGER, wasm="export interface PrintOptions { readonly row: number }\n"
        )
        good = self.fixture.context(
            self.HEADER, self.LEDGER, wasm="export interface PrintOptionsInput { readonly row: number }\n"
        )
        self.assert_rule("R5.derive", bad, good, [("wasm", "record.print_options", "-")])


class ScalarTypeTest(BindingRuleTestBase):
    LEDGER = {
        "entries": {
            "workbook.get_date1904": {"kind": "op", "js": "getDate1904", "py": "get_date1904", "returns": "bool"},
            "page_setup.set_orientation": {
                "kind": "op",
                "js": "setOrientation",
                "py": "set_orientation",
                "params": {"orientation": "enum:page_orientation"},
            },
            "full_precision.set": {
                "kind": "op",
                "js": "setFullPrecision",
                "py": "set_full_precision",
                "params": {"enabled": "bool"},
            },
            "enum.page_orientation": {"kind": "enum", "js": "PageOrientation", "py": "PageOrientation"},
        }
    }
    GOOD_JS = _wb(
        "  getDate1904(): BoolResult;\n"
        "  setOrientation(sheet: number, orientation: PageOrientation): Status;\n"
        "  setFullPrecision(enabled: boolean): Status;\n"
    )
    GOOD_PY = _py_wb(
        "    def get_date1904(self) -> bool: ...\n"
        "    def set_orientation(self, sheet: int, orientation: PageOrientation) -> None: ...\n"
        "    def set_full_precision(self, enabled: bool) -> None: ...\n"
    )

    def contexts(self):
        bad = self.fixture.context(
            "",
            self.LEDGER,
            wasm=_wb(
                "  getDate1904(): NumberResult;\n"
                "  setOrientation(sheet: number, orientation: number): Status;\n"
                "  setFullPrecision(enabled: number): Status;\n"
            ),
            native=self.GOOD_JS,
            py_workbook=_py_wb(
                "    def get_date1904(self) -> int: ...\n"
                "    def set_orientation(self, sheet: int, orientation: int) -> None: ...\n"
                "    def set_full_precision(self, enabled: bool) -> None: ...\n"
            ),
        )
        good = self.fixture.context("", self.LEDGER, wasm=self.GOOD_JS, native=self.GOOD_JS, py_workbook=self.GOOD_PY)
        return bad, good

    def test_r6_js(self) -> None:
        bad, good = self.contexts()
        self.assert_rule(
            "R6.js",
            bad,
            good,
            [
                ("wasm", "full_precision.set", "enabled"),
                ("wasm", "page_setup.set_orientation", "orientation"),
                ("wasm", "workbook.get_date1904", "return"),
            ],
        )

    def test_r6_py(self) -> None:
        bad, good = self.contexts()
        self.assert_rule(
            "R6.py",
            bad,
            good,
            [("py", "page_setup.set_orientation", "orientation"), ("py", "workbook.get_date1904", "return")],
        )


class RecordRoleTest(BindingRuleTestBase):
    LEDGER = {
        "entries": {
            "record.hyperlink": {"kind": "record", "role": "result", "js": "Hyperlink", "py": "Hyperlink"},
            "record.hyperlink_input": {
                "kind": "record",
                "role": "input",
                "js": "HyperlinkInput",
                "py": "HyperlinkInput",
                "was": {"js": "HyperlinkSpec", "py": "HyperlinkSpec"},
            },
            "record.iterative_patch": {
                "kind": "record",
                "role": "patch",
                "js": "IterativePatch",
                "py": "IterativePatch",
            },
        }
    }

    def test_r7_js(self) -> None:
        bad = self.fixture.context(
            "",
            self.LEDGER,
            wasm="export interface Hyperlink { readonly target: string; display: string }\n"
            "export interface HyperlinkSpec { target: string }\n"
            "export interface IterativePatch { enabled?: boolean }\n",
        )
        good = self.fixture.context(
            "",
            self.LEDGER,
            wasm="export interface Hyperlink { readonly target: string; readonly display: string }\n"
            "export interface HyperlinkInput { target: string }\n"
            "export type Patch<T> = { [K in keyof T]?: T[K] | null };\n"
            "export type IterativePatch = Patch<Iterative>;\n",
            native="export type Patch<T> = {\n  [K in keyof T]?: T[K] | null\n};\n",
        )
        self.assert_rule(
            "R7.js",
            bad,
            good,
            [
                ("native", "Patch", "alias"),
                ("wasm", "Patch", "alias"),
                ("wasm", "record.hyperlink", "readonly"),
                ("wasm", "record.hyperlink_input", "name"),
                ("wasm", "record.iterative_patch", "patch"),
            ],
        )

    def test_r7_py(self) -> None:
        bad = self.fixture.context(
            "",
            self.LEDGER,
            py_workbook="class Hyperlink(NamedTuple):\n    target: str\n\n"
            "class HyperlinkSpec:\n    target: str\n\n"
            "class IterativePatch(TypedDict):\n    enabled: bool\n",
        )
        good = self.fixture.context(
            "",
            self.LEDGER,
            py_workbook="@dataclass(frozen=True)\nclass Hyperlink:\n    target: str\n\n"
            "@dataclass\nclass HyperlinkInput:\n    target: str\n\n"
            "class IterativePatch(TypedDict, total=False):\n    enabled: Optional[bool]\n    limit: int | None\n",
        )
        self.assert_rule(
            "R7.py",
            bad,
            good,
            [
                ("py", "record.hyperlink", "dataclass"),
                ("py", "record.hyperlink_input", "dataclass"),
                ("py", "record.hyperlink_input", "name"),
                ("py", "record.iterative_patch", "enabled"),
                ("py", "record.iterative_patch", "typed-dict"),
            ],
        )


class ResultShapeTest(BindingRuleTestBase):
    LEDGER = {
        "collections": {"comment": {"key": "natural", "ops": ["comment.get"]}},
        "entries": {
            "record.status": {"kind": "record", "js": "Status", "py": "FormulonError"},
            "field.status_code": {
                "kind": "field",
                "record": "record.status",
                "js": "code",
                "py": "code",
                "was": {"js": "status", "py": "status"},
            },
            "sheet.get_name": {"kind": "op", "js": "getSheetName", "py": "get_sheet_name", "returns": "string"},
            "table.add": {"kind": "op", "c": "fm_sheet_add_table", "js": "addTable"},
            "workbook.load": {"kind": "factory", "js": "Workbook.load"},
            "comment.get": {"kind": "op", "c": "fm_sheet_get_comment", "js": "getComment", "py": "get_comment"},
        },
    }
    HEADER = (
        "FM_API fm_status_t fm_sheet_add_table(fm_workbook_t* wb, size_t sheet_index, size_t* out_table_index);\n"
        "FM_API fm_status_t fm_sheet_get_comment(const fm_workbook_t* wb, size_t sheet_index, uint32_t row);\n"
    )

    def test_r8_js(self) -> None:
        bad = self.fixture.context(
            self.HEADER,
            self.LEDGER,
            wasm="export interface Status { readonly ok: boolean; readonly status: number }\n"
            "export interface StringResult { readonly status: Status; readonly text: string }\n"
            + _wb(
                "  getSheetName(sheet: number): string;\n"
                "  addTable(sheet: number): Status;\n"
                "  getComment(sheet: number, row: number): Comment | null;\n"
            )
            + "export interface WorkbookCtor {\n  load(bytes: Uint8Array): Workbook;\n}\n",
        )
        good = self.fixture.context(
            self.HEADER,
            self.LEDGER,
            wasm="export interface Status { readonly ok: boolean; readonly code: StatusCode }\n"
            "export interface StringResult { readonly status: Status; readonly value: string }\n"
            + _wb(
                "  getSheetName(sheet: number): StringResult;\n"
                "  addTable(sheet: number): IndexResult;\n"
                "  getComment(sheet: number, row: number): OptionalResult<Comment>;\n"
            )
            + "export interface WorkbookCtor {\n  load(bytes: Uint8Array): OptionalResult<Workbook>;\n}\n",
        )
        self.assert_rule(
            "R8.js",
            bad,
            good,
            [
                ("wasm", "StringResult", "value"),
                ("wasm", "comment.get", "null"),
                ("wasm", "comment.get", "result"),
                ("wasm", "field.status_code", "code"),
                ("wasm", "sheet.get_name", "result"),
                ("wasm", "table.add", "result"),
                ("wasm", "workbook.load", "result"),
            ],
        )

    def test_r8_py(self) -> None:
        error = (
            "class FormulonError(Exception):\n    def __init__(self, value: int) -> None:\n        self.{} = value\n\n"
        )
        bad = self.fixture.context(
            self.HEADER,
            self.LEDGER,
            py_workbook=error.format("status")
            + _py_wb(
                "    def get_sheet_name(self, sheet: int) -> bytes: ...\n"
                "    def get_comment(self, sheet: int, row: int) -> Comment: ...\n"
                "    def get_print_titles(self, sheet: int) -> Tuple[str, str]: ...\n"
            ),
        )
        good = self.fixture.context(
            self.HEADER,
            self.LEDGER,
            py_workbook=error.format("code")
            + _py_wb(
                "    def get_sheet_name(self, sheet: int) -> str: ...\n"
                "    def get_comment(self, sheet: int, row: int) -> Optional[Comment]: ...\n"
                "    def get_print_titles(self, sheet: int) -> PrintTitles: ...\n"
            ),
        )
        self.assert_rule(
            "R8.py",
            bad,
            good,
            [
                ("py", "comment.get", "optional"),
                ("py", "field.status_code", "code"),
                ("py", "get_print_titles", "tuple"),
                ("py", "sheet.get_name", "result"),
            ],
        )

    def test_r8_total_lookup_is_not_optional(self) -> None:
        ledger = copy.deepcopy(self.LEDGER)

        def ctx() -> api_conventions.Context:
            return self.fixture.context(
                self.HEADER,
                ledger,
                wasm=_wb("  getComment(sheet: number, row: number): CommentResult;\n"),
                py_workbook=_py_wb("    def get_comment(self, sheet: int, row: int) -> Comment: ...\n"),
            )

        self.assertIn(("wasm", "comment.get", "result"), hits(ctx(), "R8.js"))
        self.assertIn(("py", "comment.get", "optional"), hits(ctx(), "R8.py"))
        ledger["entries"]["comment.get"]["total"] = "every cell reads as some comment"
        self.assertNotIn(("wasm", "comment.get", "result"), hits(ctx(), "R8.js"))
        self.assertNotIn(("py", "comment.get", "optional"), hits(ctx(), "R8.py"))


class RangeTest(BindingRuleTestBase):
    def test_r11_multi(self) -> None:
        parse = "FM_API fm_status_t fm_workbook_parse_sqref(const fm_workbook_t* wb, const char* sqref);\n"
        bad = self.fixture.context(
            RECT + parse + "typedef struct {\n  const char* sqref;\n} fm_protected_range_t;\n"
            "typedef struct {\n  const fm_cell_rect_t* sqref;\n  uint32_t sqref_count;\n} fm_cf_rule_t;\n"
            "FM_API fm_status_t fm_sheet_set_print_area(fm_workbook_t* wb, size_t sheet_index, const char* ranges_a1);\n"
            "FM_API fm_status_t fm_workbook_table_create(fm_workbook_t* wb, size_t sheet_index, const char* ref);\n",
            wasm="export interface ProtectedRange { readonly sqref: string }\n"
            + _wb("  setPrintArea(sheet: number, ranges: string): Status;\n"),
            py_workbook="class ProtectedRange:\n    sqref: str\n\n"
            + _py_wb("    def set_print_area(self, sheet: int, ranges: str) -> None: ...\n"),
        )
        good = self.fixture.context(
            RECT
            + parse
            + "typedef struct {\n  const fm_cell_rect_t* ranges;\n  size_t range_count;\n} fm_protected_range_t;\n"
            "FM_API fm_status_t fm_sheet_set_print_area(fm_workbook_t* wb, size_t sheet_index,"
            " const fm_cell_rect_t* ranges, size_t range_count);\n"
            "FM_API fm_status_t fm_sheet_recalc_viewport(fm_workbook_t* wb, size_t sheet_index,"
            " const fm_cell_rect_t* viewport, uint32_t* out_recomputed_count);\n",
            wasm="export interface ProtectedRange { readonly ranges: CellRect[] }\n"
            + _wb("  setPrintArea(sheet: number, ranges: CellRect[]): Status;\n"),
            py_workbook="class ProtectedRange:\n    ranges: List[CellRect]\n\n"
            + _py_wb("    def set_print_area(self, sheet: int, ranges: List[CellRect]) -> None: ...\n"),
        )
        self.assert_rule(
            "R11.multi",
            bad,
            good,
            [
                ("c", "fm_cf_rule_t", "sqref"),
                ("c", "fm_protected_range_t", "sqref"),
                ("c", "fm_sheet_set_print_area", "ranges_a1"),
                ("c", "fm_workbook_table_create", "ref"),
                ("py", "ProtectedRange", "sqref"),
                ("py", "set_print_area", "ranges"),
                ("wasm", "ProtectedRange", "sqref"),
                ("wasm", "setPrintArea", "ranges"),
            ],
        )

    def test_r11_sheet(self) -> None:
        edges = "  uint32_t top;\n  uint32_t left;\n  uint32_t bottom;\n  uint32_t right;\n"
        cell_entry = (
            "export interface CellEntry { readonly sheet: number; readonly row: number; readonly col: number }\n"
        )
        bad = self.fixture.context(
            RECT
            + "typedef struct {\n  size_t sheet;\n" + edges + "} fm_chart_sheet_range_t;\n"
            "typedef struct {\n  size_t sheet;\n  fm_cell_rect_t rect;\n} fm_sheet_rect_t;\n",
            wasm=cell_entry
            + "export interface ChartSheetRange { sheet: number; top: number; left: number; bottom: number; right: number }\n",
            py_workbook="class ChartSheetRange:\n    sheet: int\n    top: int\n    left: int\n    bottom: int\n    right: int\n",
        )  # fmt: skip
        good = self.fixture.context(
            RECT + "typedef struct {\n  size_t sheet_index;\n  fm_cell_rect_t rect;\n} fm_sheet_rect_t;\n",
            wasm=cell_entry + "export interface SheetRect { readonly sheet: number; readonly rect: CellRect }\n",
            py_workbook="class SheetRect:\n    sheet: int\n    rect: CellRect\n",
        )
        self.assert_rule(
            "R11.sheet",
            bad,
            good,
            [
                ("c", "fm_chart_sheet_range_t", "-"),
                ("c", "fm_sheet_rect_t", "-"),
                ("py", "ChartSheetRange", "-"),
                ("wasm", "ChartSheetRange", "-"),
            ],
        )

    def test_r11_sheet_named_reference(self) -> None:
        bad = self.fixture.context(
            RECT,
            wasm="export interface ParsedReference { sheetName: string | null; row0: number; col0: number;"
            " row1: number; col1: number }\n",
            py_workbook="class ParsedReference:\n    sheet_name: Optional[str]\n    row0: int\n    col0: int\n"
            "    row1: int\n    col1: int\n",
        )
        good = self.fixture.context(
            RECT,
            wasm="export interface ParsedReference { sheetName: string | null; rect: CellRect }\n",
            py_workbook="class ParsedReference:\n    sheet_name: Optional[str]\n    rect: CellRect\n",
        )
        self.assert_rule(
            "R11.sheet",
            bad,
            good,
            [("py", "ParsedReference", "rect"), ("wasm", "ParsedReference", "rect")],
        )


class EntryPointTest(BindingRuleTestBase):
    LEDGER = {"entries": {"memory_usage.get": {"kind": "op", "js": "getMemoryUsage", "surfaces": ["native"]}}}

    def test_r12(self) -> None:
        enum = "export enum ValueKind { Blank = 0 }\n"
        native_table = "export const ValueKind: Readonly<{ Blank: 0 }>;\n"
        native_default = "declare const _default: {\n  Workbook: WorkbookCtor;\n  ValueKind: typeof ValueKind;\n};\n"
        bad = self.fixture.context(
            "",
            self.LEDGER,
            wasm=enum
            + _wb("  addCellStyleXf(): Status;\n  delete(): void;\n")
            + "export interface FormulonModule {\n  Workbook: WorkbookCtor;\n  evalFormula(f: string): EvalResult;\n}\n",
            native=native_table + _wb("  getMemoryUsage(): number;\n") + native_default,
        )
        good = self.fixture.context(
            "",
            self.LEDGER,
            wasm=enum
            + _wb("  addCellStyleXf(): Status;\n")
            + "export interface FormulonModule {\n  ValueKind: typeof ValueKind;\n  evalFormula(f: string): EvalResult;\n}\n",
            native=native_table
            + _wb("  addCellStyleXf(): Status;\n  getMemoryUsage(): number;\n")
            + "export function evalFormula(f: string): EvalResult;\n"
            + native_default,
        )
        self.assert_rule(
            "R12",
            bad,
            good,
            [
                ("native", "addCellStyleXf", "method"),
                ("native", "delete", "method"),
                ("native", "evalFormula", "function"),
                ("wasm", "ValueKind", "module"),
            ],
        )


if __name__ == "__main__":
    unittest.main()
