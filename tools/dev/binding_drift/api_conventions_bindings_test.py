#!/usr/bin/env python3
"""Fixture tests for the ledger-bound API convention rules.

Run from the repo root: `python3 -m tools.dev.binding_drift.api_conventions_bindings_test`.
"""

from __future__ import annotations

import copy
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from tools.dev.binding_drift import api_conventions, api_conventions_bindings, api_conventions_shapes
from tools.dev.binding_drift.surface_files import _SKIPPED

WASM = "src/wasm/formulon.d.ts"
NATIVE = "packages/npm-native/index.d.ts"
PY_INIT = "packages/python/formulon/__init__.py"
PY_WORKBOOK = "packages/python/formulon/workbook.py"
PY_SHEET = "packages/python/formulon/sheet.py"
COMMON = "packages/npm/common.mjs"

RECT = (
    "typedef struct {\n  uint32_t first_row;\n  uint32_t first_col;\n  uint32_t last_row;\n  uint32_t last_col;\n"
    "} fm_cell_rect_t;\n"
)

# The static rule-ids of the checker, ledger-free and ledger-bound together.
STATIC_RULE_IDS = {
    "R0.coverage", "R1.form", "R1.scope", "R2.verb", "R3.suffix", "R3.keys", "R4.unique", "R5.derive",
    "R5.list", "R6.c", "R6.js", "R6.py", "R7.c.byptr", "R7.c.typedef", "R7.js", "R7.py", "R8.c", "R8.js",
    "R8.py", "R10.c.doc", "R11.axis", "R11.rect", "R11.multi", "R11.sheet", "R12", "R13.enum", "R14.abbrev",
    "R16.js", "R17.entry",
    "R6.c.field", "R6.closed.js", "R6.closed.py", "R7.opts.js", "R7.opts.py", "R7.absent.js", "R7.absent.py",
    "R15.header", "R18.js", "R18.py", "R19.c", "R19.js", "R19.py", "R20.js", "R20.py", "R21.role", "R21.name",
    "R22.c", "R22.js", "R22.py", "R23.js", "R23.py", "R23.track", "R23.list", "R24.missing.js",
    "R24.missing.py",
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
        paths = {"wasm": WASM, "native": NATIVE, "py_init": PY_INIT, "py_workbook": PY_WORKBOOK, "common": COMMON}
        defaults = {"py_init": "__all__ = []\n"}
        if "py_sheet" in files:
            paths["py_sheet"] = PY_SHEET
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
        self.assertEqual(len(bound), 44)
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

    def test_symbol_dispose_is_covered_by_the_dispose_entry(self) -> None:
        ledger = {
            "entries": {
                "workbook.class": {"kind": "class", "js": "Workbook"},
                "workbook.destroy": {"kind": "op", "scope": "workbook", "noun": "workbook", "js": "dispose"},
            }
        }
        body = "  dispose(): void;\n  [Symbol.dispose](): void;\n"
        found = hits(self.fixture.context("", ledger, wasm=_wb(body), native=_wb(body)), "R0.coverage")
        self.assertEqual(found, [])
        body = "  [Symbol.dispose](): void;\n"
        found = hits(self.fixture.context("", {"entries": {}}, wasm=_wb(body), native=_wb(body)), "R0.coverage")
        self.assertIn(("wasm", "[Symbol.dispose]", "[Symbol.dispose]"), found)


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
        "  getDate1904(): boolean;\n"
        "  setOrientation(sheet: number, orientation: PageOrientation): void;\n"
        "  setFullPrecision(enabled: boolean): void;\n"
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
                "  getDate1904(): number;\n"
                "  setOrientation(sheet: number, orientation: number): void;\n"
                "  setFullPrecision(enabled: number): void;\n"
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


class OptionalBoolTest(BindingRuleTestBase):
    LEDGER = {
        "entries": {
            "sheet.class": {"kind": "class", "js": "Workbook", "py": "Workbook"},
            "flag.set": {
                "kind": "op",
                "js": "setFlag",
                "py": "set_flag",
                "params": {"enabled": "bool"},
            },
            "comment_visible.get": {
                "kind": "op",
                "scope": "sheet",
                "noun": "comment_visible",
                "c": "fm_sheet_get_comment_visible",
                "js": "getCommentVisible",
                "py": "get_comment_visible",
                "returns": "bool",
            },
            "chart_title.set": {
                "kind": "op",
                "js": "setChartTitle",
                "py": "set_chart_title",
                "params": {"title": "union:ChartTitle"},
            },
            "record.chart_title": {"kind": "record", "role": "result", "py": "ChartTitle"},
        },
        "collections": {"comment_visible": {"key": "natural", "ops": ["comment_visible.get"]}},
    }
    JS = _wb(
        "  setFlag(enabled?: boolean): Status;\n  setChartTitle(title: string | null | { kind: number }): Status;\n"
    )

    def py_context(self, body: str):
        return self.fixture.context("", self.LEDGER, wasm=self.JS, native=self.JS, py_workbook=_py_wb(body))

    def test_optional_bool_with_none_default_is_accepted(self) -> None:
        ctx = self.py_context(
            "    def set_flag(self, *, enabled: Optional[bool] = None) -> None: ...\n"
            "    def set_chart_title(self, title: Union[str, None, ChartTitle]) -> None: ...\n"
        )
        self.assertEqual(hits(ctx, "R6.py"), [])
        ctx = self.py_context(
            "    def set_flag(self, enabled: bool | None = None) -> None: ...\n"
            "    def set_chart_title(self, title: Union[str, None, ChartTitle]) -> None: ...\n"
        )
        self.assertEqual(hits(ctx, "R6.py"), [])

    def test_optional_bool_without_none_default_is_flagged(self) -> None:
        for sig in ("enabled: Optional[bool]", "enabled: Optional[bool] = True"):
            ctx = self.py_context(
                f"    def set_flag(self, {sig}) -> None: ...\n"
                "    def set_chart_title(self, title: Union[str, None, ChartTitle]) -> None: ...\n"
            )
            self.assertEqual(hits(ctx, "R6.py"), [("py", "flag.set", "enabled")])

    def test_plain_bool_is_still_accepted(self) -> None:
        ctx = self.py_context(
            "    def set_flag(self, enabled: bool) -> None: ...\n"
            "    def set_chart_title(self, title: Union[str, None, ChartTitle]) -> None: ...\n"
        )
        self.assertEqual(hits(ctx, "R6.py"), [])

    def test_record_union_names_the_record_on_py_only(self) -> None:
        ctx = self.py_context(
            "    def set_flag(self, enabled: bool) -> None: ...\n"
            "    def set_chart_title(self, title: str) -> None: ...\n"
        )
        self.assertEqual(hits(ctx, "R6.py"), [("py", "chart_title.set", "title")])

    def test_boolean_lookup_accepts_a_nullable_boolean_on_js(self) -> None:
        py = _py_wb(
            "    def set_flag(self, enabled: bool) -> None: ...\n"
            "    def set_chart_title(self, title: Union[str, None, ChartTitle]) -> None: ...\n"
            "    def get_comment_visible(self, sheet: int, cell: str) -> Optional[bool]: ...\n"
        )
        head = (
            "  setFlag(enabled?: boolean): Status;\n  setChartTitle(title: string | null | { kind: number }): Status;\n"
        )
        good = _wb(head + "  getCommentVisible(sheet: number, cell: string): boolean | null;\n")
        bad = _wb(head + "  getCommentVisible(sheet: number, cell: string): number;\n")
        self.assertEqual(
            hits(self.fixture.context("", self.LEDGER, wasm=good, native=good, py_workbook=py), "R6.js"), []
        )
        self.assertEqual(
            hits(self.fixture.context("", self.LEDGER, wasm=bad, native=good, py_workbook=py), "R6.js"),
            [("wasm", "comment_visible.get", "return")],
        )

    def test_boolean_lookup_returns_optional_bool(self) -> None:
        body = (
            "    def set_flag(self, enabled: bool) -> None: ...\n"
            "    def set_chart_title(self, title: Union[str, None, ChartTitle]) -> None: ...\n"
        )
        bad = self.py_context(body + "    def get_comment_visible(self, sheet: int, cell: str) -> bool: ...\n")
        good = self.py_context(
            body + "    def get_comment_visible(self, sheet: int, cell: str) -> Optional[bool]: ...\n"
        )
        self.assertEqual(hits(bad, "R6.py"), [("py", "comment_visible.get", "return")])
        self.assertEqual(hits(good, "R6.py"), [])


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

    def test_r7_exception_keeps_name_and_form(self) -> None:
        ledger = copy.deepcopy(self.LEDGER)
        ledger["entries"]["record.hyperlink_input"]["exception"] = "one shape shared by an external file format"

        def ctx() -> api_conventions.Context:
            return self.fixture.context(
                "",
                ledger,
                wasm="export interface HyperlinkSpec { target: string }\n",
                py_workbook="class HyperlinkSpec(TypedDict):\n    target: str\n",
            )

        def own(rule_id: str) -> list[tuple[str, str, str]]:
            return [h for h in hits(ctx(), rule_id) if h[1] == "record.hyperlink_input"]

        self.assertEqual(own("R7.js"), [])
        self.assertEqual(own("R7.py"), [])
        del ledger["entries"]["record.hyperlink_input"]["exception"]
        self.assertIn(("wasm", "record.hyperlink_input", "name"), hits(ctx(), "R7.js"))
        self.assertIn(("py", "record.hyperlink_input", "dataclass"), hits(ctx(), "R7.py"))


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
            "export interface StringResult { readonly status: Status; readonly value: string }\n"
            + _wb(
                "  getSheetName(sheet: number): StringResult;\n"
                "  addTable(sheet: number): Status;\n"
                "  getComment(sheet: number, row: number): Comment;\n"
            )
            + "export interface WorkbookCtor {\n  load(bytes: Uint8Array): Workbook;\n}\n",
        )
        good = self.fixture.context(
            self.HEADER,
            self.LEDGER,
            wasm="export interface Comment { readonly text: string }\n"
            + _wb(
                "  getSheetName(sheet: number): string;\n"
                "  addTable(sheet: number): number;\n"
                "  getComment(sheet: number, row: number): Comment | null;\n"
            )
            + "export interface WorkbookCtor {\n  load(bytes: Uint8Array): Workbook;\n}\n",
        )
        self.assert_rule(
            "R8.js",
            bad,
            good,
            [
                ("wasm", "StringResult", "StringResult"),
                ("wasm", "StringResult", "status"),
                ("wasm", "comment.get", "optional"),
                ("wasm", "field.status_code", "code"),
                ("wasm", "record.status", "Status"),
                ("wasm", "record.status", "status"),
                ("wasm", "sheet.get_name", "StringResult"),
                ("wasm", "sheet.get_name", "result"),
                ("wasm", "table.add", "Status"),
                ("wasm", "table.add", "result"),
            ],
        )

    def test_r8_js_retires_every_envelope_type_and_status_field(self) -> None:
        retired = (
            "Status", "OptionalResult", "ListResult", "StringResult", "NumberResult", "BoolResult", "IndexResult",
            "ObjectIdResult",
        )  # fmt: skip
        for name in retired:
            ctx = self.fixture.context("", {"entries": {}}, wasm=f"export type {name}<T = void> = T;\n")
            self.assertEqual(hits(ctx, "R8.js"), [("wasm", name, name)])
        record = self.fixture.context(
            "", {"entries": {}}, wasm="export interface SaveDiagnostics { readonly status: number }\n"
        )
        self.assertEqual(hits(record, "R8.js"), [("wasm", "SaveDiagnostics", "status")])
        clean = self.fixture.context(
            "", {"entries": {}}, wasm="export interface SaveDiagnostics { readonly code: number }\n"
        )
        self.assertEqual(hits(clean, "R8.js"), [])

    def test_r8_js_flags_a_retired_type_in_any_return(self) -> None:
        ctx = self.fixture.context(
            "",
            {"entries": {}},
            wasm=_wb("  recalc(): Status;\n  list(): ListResult<number>;\n  count(): number;\n"),
        )
        self.assertEqual(hits(ctx, "R8.js"), [("wasm", "list", "ListResult"), ("wasm", "recalc", "Status")])

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

    def test_r8_flags_quoted_and_unledgered_tuple_returns(self) -> None:
        def ctx(body: str) -> api_conventions.Context:
            return self.fixture.context(self.HEADER, {"entries": {}}, py_workbook=_py_wb(body))

        bad = ctx(
            '    def get_display_text(self, sheet: int) -> "tuple[str, int]": ...\n'
            "    def format_value(self, value: int) -> Tuple[str, int]: ...\n"
            "    def _helper(self) -> tuple[int, int]: ...\n"
        )
        found = [h for h in hits(bad, "R8.py") if h[2] == "tuple"]
        self.assertEqual([h[1] for h in found], ["format_value", "get_display_text"])
        good = ctx("    def get_display_text(self, sheet: int) -> DisplayText: ...\n")
        self.assertEqual([h for h in hits(good, "R8.py") if h[2] == "tuple"], [])

    def test_r8_js_allocated_key_may_be_a_typed_handle(self) -> None:
        header = (
            "FM_API fm_status_t fm_workbook_add_sheet(fm_workbook_t* wb, const char* name, size_t* out_sheet_index);\n"
        )
        ledger = {"entries": {"sheet.add": {"kind": "op", "c": "fm_workbook_add_sheet", "js": "addSheet"}}}

        def ctx() -> api_conventions.Context:
            wasm = "export interface Sheet { readonly name: string }\n" + _wb("  addSheet(name: string): Sheet;\n")
            return self.fixture.context(header, ledger, wasm=wasm)

        self.assertEqual(hits(ctx(), "R8.js"), [("wasm", "sheet.add", "result")])
        ledger["entries"]["sheet.add"]["returns"] = "union:Sheet"
        self.assertEqual(hits(ctx(), "R8.js"), [])

    def test_r8_total_lookup_is_not_optional(self) -> None:
        ledger = copy.deepcopy(self.LEDGER)

        def ctx() -> api_conventions.Context:
            return self.fixture.context(
                self.HEADER,
                ledger,
                wasm=_wb("  getComment(sheet: number, row: number): Comment;\n"),
                py_workbook=_py_wb("    def get_comment(self, sheet: int, row: int) -> Comment: ...\n"),
            )

        self.assertIn(("wasm", "comment.get", "optional"), hits(ctx(), "R8.js"))
        self.assertIn(("py", "comment.get", "optional"), hits(ctx(), "R8.py"))
        ledger["entries"]["comment.get"]["total"] = "every cell reads as some comment"
        self.assertNotIn(("wasm", "comment.get", "optional"), hits(ctx(), "R8.js"))
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
            + "typedef struct {\n  const fm_cell_rect_t* ranges;\n  uint32_t range_count;\n} fm_protected_range_t;\n"
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
        bad = self.fixture.context(
            "",
            self.LEDGER,
            wasm=enum
            + _wb("  addCellStyleXf(): void;\n  delete(): void;\n")
            + "export interface Formulon {\n  Workbook: WorkbookCtor;\n  evalFormula(f: string): number;\n}\n",
            native=native_table
            + _wb("  getMemoryUsage(): number;\n")
            + "export interface Formulon {\n  Workbook: WorkbookCtor;\n  ValueKind: typeof ValueKind;\n}\n",
        )
        good = self.fixture.context(
            "",
            self.LEDGER,
            wasm=enum
            + _wb("  addCellStyleXf(): void;\n")
            + "export interface Formulon {\n  ValueKind: typeof ValueKind;\n  evalFormula(f: string): number;\n}\n",
            native=native_table
            + _wb("  addCellStyleXf(): void;\n  getMemoryUsage(): number;\n")
            + "export interface Formulon {\n  ValueKind: typeof ValueKind;\n  evalFormula(f: string): number;\n}\n",
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


def _value_union(
    count: int = 8, *, bool_type: str = "boolean", error_type: str = "ErrorCode", tag: str = "readonly"
) -> str:
    kinds = ["Blank", "Number", "Bool", "Text", "Error", "Array", "Ref", "Lambda"][:count]
    fields = {"Number": "number: number", "Bool": f"boolean: {bool_type}", "Text": "text: string"}
    fields["Error"] = f"errorCode: {error_type}"
    out = ""
    for kind in kinds:
        extra = f"; readonly {fields[kind]}" if kind in fields else ""
        out += f"export interface {kind}Value {{ {tag} kind: typeof ValueKind.{kind}{extra} }}\n"
    return out + "export type Value =\n" + "".join(f"  | {kind}Value\n" for kind in kinds).rstrip("\n") + ";\n"


class ValueUnionTest(BindingRuleTestBase):
    def test_r16_js(self) -> None:
        def ctx(wasm: str, native: str):
            return self.fixture.context("", None, wasm=wasm, native=native)

        good = ctx(_value_union(), _value_union())
        bad = ctx(
            _value_union(7),
            _value_union(bool_type="number", error_type="number", tag="")
            + "export interface TextValue2 { readonly kind: typeof ValueKind.Text }\n",
        )
        self.assert_rule(
            "R16.js",
            bad,
            good,
            [
                ("native", "Value", "BlankValue.kind"),
                ("native", "Value", "NumberValue.kind"),
                ("native", "Value", "BoolValue.kind"),
                ("native", "Value", "TextValue.kind"),
                ("native", "Value", "ErrorValue.kind"),
                ("native", "Value", "ArrayValue.kind"),
                ("native", "Value", "RefValue.kind"),
                ("native", "Value", "LambdaValue.kind"),
                ("wasm", "Value", "union"),
            ],
        )

    def test_r16_js_field_types_and_duplicate_kinds(self) -> None:
        bad = self.fixture.context(
            "",
            None,
            wasm=_value_union(bool_type="number", error_type="number"),
            native=_value_union().replace("typeof ValueKind.Ref", "typeof ValueKind.Array"),
        )
        self.assertEqual(
            hits(bad, "R16.js"),
            [
                ("native", "Value", "typeof ValueKind.Array"),
                ("wasm", "Value", "BoolValue.boolean"),
                ("wasm", "Value", "ErrorValue.errorCode"),
            ],
        )

    def test_a_missing_value_is_reported(self) -> None:
        ctx = self.fixture.context("", None, wasm="export interface Value { readonly kind: number }\n", native="")
        self.assertEqual(hits(ctx, "R16.js"), [("native", "Value", "union"), ("wasm", "Value", "union")])

    def test_r7_js_accepts_a_union_of_result_interfaces(self) -> None:
        ledger = {"entries": {"record.value": {"kind": "record", "role": "result", "js": "Value"}}}
        good = self.fixture.context("", ledger, wasm=_value_union(), native=_value_union())
        bad = self.fixture.context("", ledger, wasm="export type Value = number | string;\n", native=_value_union())
        self.assertEqual(hits(good, "R7.js"), [])
        self.assertEqual(hits(bad, "R7.js"), [("wasm", "record.value", "interface")])


class EntryShapeTest(BindingRuleTestBase):
    COMMON = (
        "export const FORMULON_MEMBERS = Object.freeze([\n"
        "  'Workbook', // factory\n  'ValueKind',\n  'FormulonError',\n  'getVersion',\n]);\n"
    )
    GOOD = (
        "export const ValueKind: Readonly<{ Blank: 0 }>;\n"
        "export const PIVOT_VALUES_FIELD = -2;\n"
        "export class FormulonError extends Error { readonly code: number }\n"
        "export interface Formulon {\n  Workbook: WorkbookCtor;\n  ValueKind: typeof ValueKind;\n"
        "  FormulonError: typeof FormulonError;\n  getVersion(): string;\n}\n"
        "export default function createFormulon(opts?: FormulonOptions): Promise<Formulon>;\n"
    )

    def test_r17_entry(self) -> None:
        bad_native = (
            "export const ValueKind: Readonly<{ Blank: 0 }>;\n"
            "export const Workbook: WorkbookCtor;\n"
            "export function getVersion(): string;\n"
            "export class FormulonError extends Error { readonly code: number }\n"
            "export interface Formulon {\n  Workbook: WorkbookCtor;\n  ValueKind: typeof ValueKind;\n"
            "  FormulonError: typeof FormulonError;\n  evalFormula(f: string): number;\n}\n"
        )
        bad = self.fixture.context("", None, wasm=self.GOOD, native=bad_native, common=self.COMMON)
        good = self.fixture.context("", None, wasm=self.GOOD, native=self.GOOD, common=self.COMMON)
        self.assert_rule(
            "R17.entry",
            bad,
            good,
            [
                ("native", "Formulon", "evalFormula"),
                ("native", "Formulon", "getVersion"),
                ("native", "Workbook", "Workbook"),
                ("native", "createFormulon", "createFormulon"),
                ("native", "getVersion", "getVersion"),
            ],
        )

    def test_r17_entry_without_a_member_list(self) -> None:
        ctx = self.fixture.context("", None, wasm=self.GOOD, native=self.GOOD, common="export const X = 1;\n")
        self.assertEqual(
            hits(ctx, "R17.entry"),
            [("native", "Formulon", "FORMULON_MEMBERS"), ("wasm", "Formulon", "FORMULON_MEMBERS")],
        )


class EnumTypeTest(BindingRuleTestBase):
    def test_r13_enum_js(self) -> None:
        bad = self.fixture.context(
            "",
            None,
            wasm="export interface ImageInfo { readonly format: number; readonly level: LogLevel }\n"
            + _wb(
                "  save(format?: number): Uint8Array;\n"
                "  setError(sheet: number, errorCode: ErrorCode | number): void;\n"
                "  setLogMinLevel(level: number): void;\n"
                "  setRowOutline(sheet: number, row: number, level: number): void;\n"
                "  setColumnOutline(sheet: number, column: number, level: number): void;\n"
            ),
        )
        good = self.fixture.context(
            "",
            None,
            wasm="export interface ImageInfo { readonly format: ImageFormat; readonly level: LogLevel }\n"
            + _wb(
                "  save(format?: WorkbookFormat): Uint8Array;\n"
                "  setError(sheet: number, errorCode: ErrorCode): void;\n"
                "  setLogMinLevel(level: LogLevel): void;\n"
                "  setRowOutline(sheet: number, row: number, level: number): void;\n"
                "  setColumnOutline(sheet: number, column: number, level: number): void;\n"
            ),
        )
        self.assert_rule(
            "R13.enum",
            bad,
            good,
            [
                ("wasm", "ImageInfo", "ImageInfo.format"),
                ("wasm", "save", "Workbook.save(format)"),
                ("wasm", "setError", "Workbook.setError(errorCode)"),
                ("wasm", "setLogMinLevel", "Workbook.setLogMinLevel(level)"),
            ],
        )

    def test_r13_enum_py(self) -> None:
        bad = "@dataclass(frozen=True)\nclass StatusBox:\n    status: int\n    code: int\n\n" + _py_wb(
            "    def save(self, format: int = 0) -> bytes: ...\n"
            "    def set_error(self, sheet: int, error_code: ErrorCode | int) -> None: ...\n"
            "    def set_log_min_level(self, level: Optional[int]) -> None: ...\n"
            "    def set_row_outline(self, sheet: int, row: int, level: int) -> None: ...\n"
        )
        good = "@dataclass(frozen=True)\nclass StatusBox:\n    status: StatusCode\n    code: int\n\n" + _py_wb(
            "    def save(self, format: WorkbookFormat = WorkbookFormat.XLSX) -> bytes: ...\n"
            "    def set_error(self, sheet: int, error_code: ErrorCode) -> None: ...\n"
            "    def set_log_min_level(self, level: LogLevel) -> None: ...\n"
            "    def set_row_outline(self, sheet: int, row: int, level: int) -> None: ...\n"
        )
        init = "__all__ = ['StatusBox']\n"
        self.assert_rule(
            "R13.enum",
            self.fixture.context("", None, py_init=init, py_workbook=bad),
            self.fixture.context("", None, py_init=init, py_workbook=good),
            [
                ("py", "StatusBox", "StatusBox.status"),
                ("py", "save", "Workbook.save(format)"),
                ("py", "set_error", "Workbook.set_error(error_code)"),
                ("py", "set_log_min_level", "Workbook.set_log_min_level(level)"),
            ],
        )


class AbbreviationTest(BindingRuleTestBase):
    def test_r14_abbrev_js(self) -> None:
        bad = self.fixture.context(
            "",
            None,
            wasm="export interface Anchor {\n  readonly fromColOff: number;\n  readonly firstCol: number;\n"
            "  readonly idx: number;\n}\n"
            + _wb(
                "  addColBreak(sheet: number, col: number): void;\n"
                "  insertCols(sheet: number, pivotIdx: number): void;\n"
                "  getByIdx(): number;\n"
                "  setShowRowColHeaders(show: boolean): void;\n"
                "  addColumnBreak(sheet: number, col: number): void;\n"
            ),
        )
        good = self.fixture.context(
            "",
            None,
            wasm="export interface Anchor { readonly fromColOffset: number; readonly firstCol: number }\n"
            "export interface View { readonly showRowColHeaders: boolean; readonly colOffsetEmu: number }\n"
            + _wb(
                "  addColumnBreak(sheet: number, col: number): void;\n"
                "  insertColumns(sheet: number, pivotIndex: number): void;\n"
                "  getByIndex(): number;\n"
                "  setShowRowColHeaders(show: boolean): void;\n"
            ),
        )
        self.assert_rule(
            "R14.abbrev",
            bad,
            good,
            [
                ("wasm", "Anchor", "idx:Anchor.idx"),
                ("wasm", "Anchor", "off:Anchor.fromColOff"),
                ("wasm", "addColBreak", "col:Workbook.addColBreak"),
                ("wasm", "getByIdx", "idx:Workbook.getByIdx"),
                ("wasm", "insertCols", "col:Workbook.insertCols"),
                ("wasm", "insertCols", "idx:Workbook.insertCols(pivotIdx)"),
                ("wasm", "setShowRowColHeaders", "col:Workbook.setShowRowColHeaders"),
            ],
        )

    def test_r14_abbrev_py(self) -> None:
        bad = "@dataclass(frozen=True)\nclass Anchor:\n    from_col_off: int\n    first_col: int\n\n" + _py_wb(
            "    def add_col_break(self, sheet: int, col: int) -> None: ...\n"
            "    def get_pivot_idx(self, pivot_idx: int) -> int: ...\n"
            "    def set_show_row_col_headers(self, show: bool) -> None: ...\n"
        )
        good = (
            "@dataclass(frozen=True)\nclass Anchor:\n    from_col_offset: int\n    first_col: int\n\n"
            "@dataclass(frozen=True)\nclass View:\n    show_row_col_headers: bool\n\n"
            + _py_wb(
                "    def add_column_break(self, sheet: int, col: int) -> None: ...\n"
                "    def get_pivot_index(self, pivot_index: int) -> int: ...\n"
                "    def set_show_row_col_headers(self, show: bool) -> None: ...\n"
            )
        )
        init = "__all__ = ['Anchor', 'View']\n"
        self.assert_rule(
            "R14.abbrev",
            self.fixture.context("", None, py_init=init, py_workbook=bad),
            self.fixture.context("", None, py_init=init, py_workbook=good),
            [
                ("py", "Anchor", "off:Anchor.from_col_off"),
                ("py", "add_col_break", "col:Workbook.add_col_break"),
                ("py", "get_pivot_idx", "idx:Workbook.get_pivot_idx"),
                ("py", "get_pivot_idx", "idx:Workbook.get_pivot_idx(pivot_idx)"),
                ("py", "set_show_row_col_headers", "col:Workbook.set_show_row_col_headers"),
            ],
        )

    def test_r4_synonym_for_a_vertical_break_is_the_column_break(self) -> None:
        decl = "FM_API fm_status_t fm_pagination_get_{}_count(const fm_pagination_t* p, size_t* out_count);\n"
        bad = self.fixture.context(decl.format("vertical_break"), None, wasm=_wb("  getVerticalBreaks(): number[];\n"))
        good = self.fixture.context(decl.format("column_break"), None, wasm=_wb("  getColumnBreaks(): number[];\n"))
        self.assert_rule(
            "R4.unique",
            bad,
            good,
            [
                ("c", "fm_pagination_get_vertical_break_count", "vertical_break"),
                ("wasm", "getVerticalBreaks", "vertical_break"),
            ],
        )


def _sheet(members: str) -> str:
    return "export interface Sheet {\n" + members + "}\n"


def _py_sheet(methods: str) -> str:
    return "class Sheet:\n" + (methods or "    pass\n")


FACADE = "function makeFacade(raw) { return raw; }\n"


class SheetSurfaceTest(BindingRuleTestBase):
    """The ledger rules read `Workbook` and `Sheet` together."""

    LEDGER = {
        "entries": {
            "type.sheet": {"kind": "record", "js": "Sheet", "py": "Sheet", "surfaces": ["wasm", "native", "py"]},
            "zoom.get": {
                "kind": "op", "scope": "sheet", "c": "fm_sheet_get_zoom", "js": "getZoom", "py": "get_zoom",
                "returns": "int", "params": {"visible": "bool"}, "surfaces": ["c", "wasm", "native", "py"],
            },
        }
    }  # fmt: skip

    def test_r0_coverage_reads_sheet_and_skips_its_facade_members(self) -> None:
        body = _sheet("  readonly name: string;\n  getZoom(): number;\n  get(a1: string): Value;\n")
        good = self.fixture.context("", self.LEDGER, wasm=body, native=body,
                                    py_sheet=_py_sheet("    def get_zoom(self) -> int: ...\n"
                                                       "    def get(self, a1: str) -> int: ...\n"))  # fmt: skip
        self.assertEqual(hits(good, "R0.coverage"), [])
        bad = self.fixture.context("", self.LEDGER, wasm=_sheet("  getZoomLevel(): number;\n"),
                                   py_sheet=_py_sheet("    def get_zoom_level(self) -> int: ...\n"))  # fmt: skip
        self.assertEqual(
            hits(bad, "R0.coverage"),
            [("py", "get_zoom_level", "get_zoom_level"), ("wasm", "getZoomLevel", "getZoomLevel")],
        )

    def test_r6_js_judges_a_sheet_method(self) -> None:
        bad = self.fixture.context("", self.LEDGER, wasm=_sheet("  getZoom(visible: number): number;\n"))
        good = self.fixture.context("", self.LEDGER, wasm=_sheet("  getZoom(visible: boolean): number;\n"))
        self.assert_rule("R6.js", bad, good, [("wasm", "zoom.get", "visible")])


class ClassScopeTest(BindingRuleTestBase):
    LEDGER = {
        "entries": {
            "zoom.get": {"kind": "op", "scope": "sheet", "js": "getZoom", "py": "get_zoom",
                         "surfaces": ["wasm", "native", "py"]},
            "sheet.add": {"kind": "op", "scope": "workbook", "js": "addSheet", "py": "add_sheet",
                          "surfaces": ["wasm", "native", "py"]},
        }
    }  # fmt: skip

    def test_r23_js(self) -> None:
        bad = _wb("  getZoom(sheet: number): number;\n") + _sheet("  addSheet(name: string): Sheet;\n")
        good = _wb("  addSheet(name: string): Sheet;\n") + _sheet("  getZoom(): number;\n")
        self.assert_rule(
            "R23.js",
            self.fixture.context("", self.LEDGER, wasm=bad, native=_sheet("  getZoom(sheet: number): number;\n")),
            self.fixture.context("", self.LEDGER, wasm=good, native=good, common=FACADE),
            [
                ("native", "zoom.get", "sheet-arg"),
                ("wasm", "Workbook", "makeFacade"),
                ("wasm", "sheet.add", "Sheet"),
                ("wasm", "zoom.get", "Workbook"),
            ],
        )

    def test_r23_py(self) -> None:
        self.assert_rule(
            "R23.py",
            self.fixture.context(
                "", self.LEDGER, py_workbook=_py_wb("    def get_zoom(self, sheet: int) -> int: ...\n"),
                py_sheet=_py_sheet("    def add_sheet(self, name: str) -> 'Sheet': ...\n"),
            ),
            self.fixture.context(
                "", self.LEDGER, py_workbook=_py_wb("    def add_sheet(self, name: str) -> 'Sheet': ...\n"),
                py_sheet=_py_sheet("    def get_zoom(self) -> int: ...\n"),
            ),
            [("py", "sheet.add", "Sheet"), ("py", "zoom.get", "Workbook")],
        )  # fmt: skip

    def test_r23_py_flags_a_leading_sheet_argument(self) -> None:
        ctx = self.fixture.context(
            "", self.LEDGER, py_sheet=_py_sheet("    def get_zoom(self, sheet: int) -> int: ...\n")
        )
        self.assertEqual(hits(ctx, "R23.py"), [("py", "zoom.get", "sheet-arg")])

    def test_r23_track(self) -> None:
        header = (
            "FM_API fm_status_t fm_workbook_add_sheet(fm_workbook_t* wb, const char* name, size_t* out_sheet_index);\n"
            "FM_API fm_status_t fm_sheet_show_pivot_details(fm_workbook_t* wb, size_t sheet_index,"
            " size_t* out_sheet_index);\n"
            "FM_API fm_status_t fm_workbook_get_table(const fm_workbook_t* wb, size_t table_index,"
            " size_t* out_sheet_index);\n"
        )
        ledger = {
            "collections": {"sheet": {"key": "index", "ops": ["sheet.add"]}},
            "entries": {
                key: {"kind": "op", "c": c, "js": js, "py": py, "surfaces": ["c", "wasm", "native", "py"]}
                for key, c, js, py in (
                    ("sheet.add", "fm_workbook_add_sheet", "addSheet", "add_sheet"),
                    ("pivot_details.show", "fm_sheet_show_pivot_details", "showPivotDetails", "show_pivot_details"),
                    ("table.get", "fm_workbook_get_table", "getTable", "get_table"),
                )
            },
        }
        common = (
            "// <facade-tables>\nconst SHEET_STRUCTURE_OPS = Object.freeze(['addSheet', 'showPivotDetails']);\n"
            "// </facade-tables>\n"
        )
        sheet_py = '# <facade-tables>\nSHEET_STRUCTURE_OPS = frozenset({"add_sheet", "show_pivot_details"})\n# </facade-tables>\n'
        self.assert_rule(
            "R23.track",
            self.fixture.context(header, ledger, common=common.replace(", 'showPivotDetails'", "")),
            self.fixture.context(header, ledger, common=common, py_sheet=sheet_py),
            [
                ("py", "pivot_details.show", "SHEET_STRUCTURE_OPS"),
                ("py", "sheet.add", "SHEET_STRUCTURE_OPS"),
                ("wasm", "pivot_details.show", "SHEET_STRUCTURE_OPS"),
            ],
        )


class CollectionAccessTest(BindingRuleTestBase):
    def ledger(self) -> dict:
        def op(c: str, js: str) -> dict:
            return {"kind": "op", "scope": "workbook", "c": c, "js": js, "surfaces": ["c", "wasm"]}

        return {
            "collections": {
                "table": {"key": "index", "access": "list", "plural": "tables", "ops": ["table.get", "table.get_count"]},
                "font": {"key": "index", "access": "indexed", "plural": "fonts", "ops": ["font.get", "font.get_count"]},
                "run": {"key": "index", "access": "property", "plural": "cell_phonetic_runs",
                        "ops": ["run.get", "run.get_count"]},
                "pivot": {"key": "index", "access": "exempt", "plural": "pivots", "exempt": "x",
                          "ops": ["pivot.get_count"]},
            },
            "entries": {
                "table.get": op("fm_workbook_get_table", "getTable"),
                "table.get_count": op("fm_workbook_get_table_count", "getTableCount"),
                "font.get": op("fm_workbook_get_font", "getFont"),
                "font.get_count": op("fm_workbook_get_font_count", "getFontCount"),
                "run.get": op("fm_sheet_get_cell_phonetic_run", "getCellPhoneticRun"),
                "run.get_count": op("fm_sheet_get_cell_phonetic_run_count", "getCellPhoneticRunCount"),
                "pivot.get_count": op("fm_sheet_get_pivot_count", "getPivotCount"),
            },
        }  # fmt: skip

    def test_r23_list(self) -> None:
        bad = _wb(
            "  getTable(i: number): Table;\n  getTableCount(): number;\n  getFont(i: number): Font;\n"
            "  listFonts(): readonly Font[];\n  getCellPhoneticRun(i: number): Run;\n  getCellPhoneticRunCount(): number;\n"
            "  getPivotCount(): number;\n"
        )
        good = _wb(
            "  listTables(): readonly Table[];\n  getFont(i: number): Font;\n  getFontCount(): number;\n"
            "  getCellPhoneticRuns(): readonly Run[];\n  setCellPhoneticRuns(runs: readonly Run[] | null): void;\n"
            "  getPivotCount(): number;\n"
        )
        self.assert_rule(
            "R23.list",
            self.fixture.context("", self.ledger(), wasm=bad),
            self.fixture.context("", self.ledger(), wasm=good),
            [
                ("wasm", "collection.font", "indexed"),
                ("wasm", "collection.font", "listFonts"),
                ("wasm", "collection.run", "getCellPhoneticRuns"),
                ("wasm", "collection.run", "setCellPhoneticRuns"),
                ("wasm", "collection.table", "listTables"),
                ("wasm", "run.get", "getCellPhoneticRun"),
                ("wasm", "run.get_count", "getCellPhoneticRunCount"),
                ("wasm", "table.get", "getTable"),
                ("wasm", "table.get_count", "getTableCount"),
            ],
        )


class OptionsObjectTest(BindingRuleTestBase):
    def test_r7_opts_js(self) -> None:
        bad = _wb(
            "  addLink(sheet: number, row: number, col: number, target: string, display: string, tip: string): void;\n"
            "  setTotals(sheet: number, rows: boolean, cols: boolean): void;\n  setShowZeros(sheet: number, s: boolean): void;\n"
        )
        good = _wb(
            "  addLink(sheet: number, row: number, col: number, link: HyperlinkInput): void;\n"
            "  setTotals(sheet: number, totals: GrandTotalsPatch): void;\n  setShowZeros(sheet: number, s: boolean): void;\n"
        )
        self.assert_rule(
            "R7.opts.js",
            self.fixture.context("", None, wasm=bad),
            self.fixture.context("", None, wasm=good),
            [("wasm", "addLink", "run"), ("wasm", "setTotals", "bool")],
        )

    def test_r7_opts_py(self) -> None:
        bad = _py_wb(
            "    def add_link(self, sheet: int, row: int, col: int, target: str, display: str, tip: str) -> None: ...\n"
            "    def set_totals(self, sheet: int, rows: bool, cols: bool) -> None: ...\n"
            "    def set_show_zeros(self, sheet: int, show: bool) -> None: ...\n"
        )
        good = _py_wb(
            "    def add_link(self, sheet: int, row: int, col: int, target: str, *, display: str = '', tip: str = '')"
            " -> None: ...\n"
            "    def set_totals(self, sheet: int, *, rows: bool, cols: bool) -> None: ...\n"
            "    def set_show_zeros(self, sheet: int, show: bool) -> None: ...\n"
        )
        self.assert_rule(
            "R7.opts.py",
            self.fixture.context("", None, py_workbook=bad),
            self.fixture.context("", None, py_workbook=good),
            [("py", "add_link", "run"), ("py", "set_totals", "bool:cols"), ("py", "set_totals", "bool:rows")],
        )

    def test_single_boolean_setters_are_counted_by_shape(self) -> None:
        self.assertTrue(
            api_conventions_shapes.single_boolean_setter("setShowZeros", (("sheet", "number"), ("s", "boolean")), True)
        )
        self.assertFalse(
            api_conventions_shapes.single_boolean_setter("addBreak", (("row", "number"), ("m", "boolean")), True)
        )
        self.assertTrue(
            api_conventions_shapes.single_boolean_setter("set_hidden", (("sheet", "int"), ("h", "bool")), False)
        )


class SetterSymmetryTest(BindingRuleTestBase):
    RECORDS = (
        "export interface IterativeSettings {\n  readonly enabled: boolean;\n  readonly maxIterations: number;\n}\n"
        "export interface Thing {\n  readonly a: number;\n}\nexport interface ThingInput {\n  a: number;\n  b?: number;\n}\n"
        "export type IterativeSettingsPatch = Patch<IterativeSettings>;\n"
    )

    def test_r18_js(self) -> None:
        bad = self.RECORDS + _wb(
            "  getIterative(): IterativeSettings;\n  setIterative(enabled: boolean, maxIterations: number): void;\n"
            "  getThing(sheet: number): Thing | null;\n  setThing(sheet: number, t: ThingInput): void;\n"
            "  setOptions(o: Partial<Thing>): void;\n"
        )
        good = self.RECORDS.replace("  b?: number;\n", "") + _wb(
            "  getIterative(): IterativeSettings;\n  setIterative(patch: IterativeSettingsPatch): void;\n"
            "  getThing(sheet: number): Thing | null;\n  setThing(sheet: number, t: ThingInput | null): void;\n"
        )
        self.assert_rule(
            "R18.js",
            self.fixture.context("", None, wasm=bad),
            self.fixture.context("", None, wasm=good),
            [("wasm", "setIterative", "shape"), ("wasm", "setOptions", "Partial:o"), ("wasm", "setThing", "fields")],
        )

    def test_r18_js_reads_a_sheet_argument_through_its_alias(self) -> None:
        records = (
            "export interface Sheet {\n  readonly name: string;\n}\nexport type SheetRef = Sheet | number | string;\n"
        )
        bad = records + _wb("  getActiveSheet(): Sheet;\n  setActiveSheet(sheet: number): void;\n")
        good = records + _wb("  getActiveSheet(): Sheet;\n  setActiveSheet(sheet: SheetRef): void;\n")
        self.assert_rule(
            "R18.js",
            self.fixture.context("", None, wasm=bad),
            self.fixture.context("", None, wasm=good),
            [("wasm", "setActiveSheet", "shape")],
        )

    def test_r18_py(self) -> None:
        records = "@dataclass(frozen=True)\nclass IterativeSettings:\n    enabled: bool\n    max_iterations: int\n\n"
        bad = records + _py_wb(
            "    def get_iterative(self) -> IterativeSettings: ...\n"
            "    def set_iterative(self, enabled: bool, max_iterations: int) -> None: ...\n"
        )
        good = records + _py_wb(
            "    def get_iterative(self) -> IterativeSettings: ...\n"
            "    def set_iterative(self, *, enabled: bool = UNCHANGED, max_iterations: int = UNCHANGED) -> None: ...\n"
        )
        self.assert_rule(
            "R18.py",
            self.fixture.context("", None, py_workbook=bad),
            self.fixture.context("", None, py_workbook=good),
            [("py", "set_iterative", "shape")],
        )


class AbsentValueTest(BindingRuleTestBase):
    LEDGER = {
        "entries": {
            "record.font": {"kind": "record", "role": "result", "js": "Font", "py": "Font",
                            "surfaces": ["wasm", "native", "py"]},
            "protection.get": {"kind": "op", "scope": "sheet", "js": "getProtection", "py": "get_protection",
                               "total": "reads as the all-off record", "surfaces": ["wasm", "native", "py"]},
        }
    }  # fmt: skip

    def test_r7_absent_js(self) -> None:
        bad = "export interface Font {\n  readonly hasBold: boolean;\n  readonly bold: boolean;\n}\n" + _wb(
            '  /** Reads the file name; `""` when unset. */\n  getFileName(): string;\n'
            '  /** Sets the file name; `""` clears it. */\n  setFileName(name: string): void;\n'
            "  setAxis(sheet: number, axis: PivotAxis | -1): void;\n"
            "  /** Every field is zero or empty when absent. */\n  getProtection(sheet: number): Protection;\n"
        )
        good = "export interface Font {\n  readonly bold: boolean | null;\n}\n" + _wb(
            "  /** Reads the file name; `null` when unset. */\n  getFileName(): string | null;\n"
            '  /** Sets the file name; `null` or `""` clears it. */\n  setFileName(name: string | null): void;\n'
            "  setAxis(sheet: number, axis: PivotAxis | null): void;\n"
        )
        self.assert_rule(
            "R7.absent.js",
            self.fixture.context("", self.LEDGER, wasm=bad),
            self.fixture.context("", self.LEDGER, wasm=good),
            [
                ("wasm", "record.font", "has:hasBold"),
                ("wasm", "getFileName", "doc:Workbook.getFileName"),
                ("wasm", "setAxis", "-1:Workbook.setAxis(axis)"),
                ("wasm", "setFileName", "clear:Workbook.setFileName"),
            ],
        )

    def test_r7_absent_py(self) -> None:
        bad = "@dataclass(frozen=True)\nclass Font:\n    has_bold: bool\n    bold: bool\n\n" + _py_wb(
            '    def get_file_name(self) -> str:\n        """Returns ``""`` when unset."""\n'
            '    def set_file_name(self, name: str) -> None:\n        """Pass ``""`` to clear the name."""\n'
            "    def get_name(self, name: str, local_sheet_id: int = -1) -> str: ...\n"
        )
        good = "@dataclass(frozen=True)\nclass Font:\n    bold: Optional[bool]\n\n" + _py_wb(
            '    def get_file_name(self) -> Optional[str]:\n        """Returns None when unset."""\n'
            '    def set_file_name(self, name: Optional[str]) -> None:\n        """Pass None or ``""`` to clear it."""\n'
            "    def get_name(self, name: str, local_sheet_id: Optional[int] = None) -> str: ...\n"
        )
        self.assert_rule(
            "R7.absent.py",
            self.fixture.context("", self.LEDGER, py_workbook=bad),
            self.fixture.context("", self.LEDGER, py_workbook=good),
            [
                ("py", "record.font", "has:has_bold"),
                ("py", "get_file_name", "doc"),
                ("py", "get_name", "-1:local_sheet_id"),
                ("py", "set_file_name", "clear"),
            ],
        )


class ClosedSetTest(BindingRuleTestBase):
    HEADER = (
        "typedef enum { FM_MODE_A = 0, FM_MODE_B = 1 } fm_mode_t;\n"
        "typedef struct {\n  int32_t mode; /* fm_mode_t */\n  int32_t count;\n} fm_thing_t;\n"
    )
    LEDGER = {
        "entries": {
            "enum.mode": {"kind": "enum", "c": "fm_mode_t", "js": "Mode", "py": "Mode", "surfaces": ["c", "wasm", "py"]},
            "record.thing": {"kind": "record", "role": "result", "c": "fm_thing_t", "js": "Thing", "py": "Thing",
                             "surfaces": ["c", "wasm", "py"]},
            "mode.set": {"kind": "op", "scope": "workbook", "js": "setMode", "py": "set_mode",
                         "params": {"mode": "enum:mode"}, "surfaces": ["wasm", "py"]},
        }
    }  # fmt: skip
    TABLE = "export const Mode: Readonly<{\n  A: 0;\n  B: 1;\n}>;\nexport type Mode = 0 | 1;\n"

    def test_r6_closed_js(self) -> None:
        bad = (
            self.TABLE
            + "export interface Thing {\n  readonly mode: number;\n  readonly count: number;\n}\n"
            + _wb("  setMode(mode: Mode | number): void;\n")
            + "export interface WorkbookCtor {\n  createForProfile(profileId: string): Workbook;\n}\n"
        )
        good = (
            self.TABLE
            + "export interface Thing {\n  readonly mode: Mode;\n  readonly count: number;\n}\n"
            + _wb("  setMode(mode: Mode): void;\n")
            + "export interface WorkbookCtor {\n  createForProfile(profileId: ExcelProfileId): Workbook;\n}\n"
        )
        self.assert_rule(
            "R6.closed.js",
            self.fixture.context(self.HEADER, self.LEDGER, wasm=bad),
            self.fixture.context(self.HEADER, self.LEDGER, wasm=good),
            [
                ("wasm", "Workbook.createForProfile", "Workbook.createForProfile(profileId)"),
                ("wasm", "mode.set", "setMode(mode)"),
                ("wasm", "record.thing", "Thing.mode"),
            ],
        )

    def test_r6_closed_py(self) -> None:
        prelude = "class Mode(IntEnum):\n    A = 0\n    B = 1\n\n"
        bad = (
            prelude
            + "@dataclass(frozen=True)\nclass Thing:\n    mode: int\n    count: int\n\n"
            + _py_wb("    def set_mode(self, mode: 'Mode | int') -> None: ...\n")
        )
        good = (
            prelude
            + "@dataclass(frozen=True)\nclass Thing:\n    mode: Mode\n    count: int\n\n"
            + _py_wb("    def set_mode(self, mode: Mode) -> None: ...\n")
        )
        self.assert_rule(
            "R6.closed.py",
            self.fixture.context(self.HEADER, self.LEDGER, py_workbook=bad),
            self.fixture.context(self.HEADER, self.LEDGER, py_workbook=good),
            [("py", "mode.set", "set_mode(mode)"), ("py", "record.thing", "Thing.mode")],
        )


class EnumFamilyTest(BindingRuleTestBase):
    HEADER = (
        "typedef enum { FM_PIVOT_SORT_DESCENDING = 0, FM_PIVOT_SORT_ASCENDING = 1 } fm_pivot_sort_t;\n"
        "typedef enum { FM_SLICER_SORT_ASCENDING = 0, FM_SLICER_SORT_DESCENDING = 1 } fm_slicer_sort_t;\n"
        "typedef enum { FM_BLANKS_ABSENT = 0, FM_BLANKS_GAP = 1, FM_BLANKS_ZERO = 2 } fm_chart_blanks_as_t;\n"
        "typedef enum { FM_EMPTY_GAP = 0, FM_EMPTY_ZERO = 1 } fm_empty_cells_t;\n"
    )

    def ledger(self, distinct: bool = False) -> dict:
        def enum(c: str, name: str) -> dict:
            return {"kind": "enum", "c": c, "js": name, "py": name, "surfaces": ["c", "wasm", "py"]}

        data = {
            "enum_families": {"sort_order": ["enum.pivot_sort", "enum.slicer_sort"]},
            "entries": {
                "enum.pivot_sort": enum("fm_pivot_sort_t", "PivotSort"),
                "enum.slicer_sort": enum("fm_slicer_sort_t", "SlicerSort"),
                "enum.chart_blanks_as": enum("fm_chart_blanks_as_t", "ChartBlanksAs"),
                "enum.empty_cells": enum("fm_empty_cells_t", "EmptyCells"),
            },
        }
        if distinct:
            data["enum_distinct"] = [{"enums": ["enum.chart_blanks_as", "enum.empty_cells"], "reason": "fixture"}]
        return data

    def test_r19_c(self) -> None:
        good = self.HEADER.replace("FM_SLICER_SORT_ASCENDING = 0, FM_SLICER_SORT_DESCENDING = 1",
                                   "FM_SLICER_SORT_DESCENDING = 0, FM_SLICER_SORT_ASCENDING = 1")  # fmt: skip
        good = good.replace("FM_BLANKS_ABSENT = 0, FM_BLANKS_GAP = 1, FM_BLANKS_ZERO = 2",
                            "FM_BLANKS_ABSENT = -1, FM_BLANKS_GAP = 0, FM_BLANKS_ZERO = 1")  # fmt: skip
        self.assert_rule(
            "R19.c",
            self.fixture.context(self.HEADER, self.ledger()),
            self.fixture.context(good, self.ledger(distinct=True)),
            [
                ("c", "enum.chart_blanks_as", "absent:Absent"),
                ("c", "enum.empty_cells", "pair:enum.chart_blanks_as"),
                ("c", "enum.slicer_sort", "ascending"),
                ("c", "enum.slicer_sort", "descending"),
            ],
        )

    def test_r19_js(self) -> None:
        def table(name: str, members: str) -> str:
            return f"export const {name}: Readonly<{{\n{members}}}>;\n"

        bad = (
            table("PivotSort", "  Descending: 0;\n  Ascending: 1;\n")
            + table("SlicerSort", "  Ascending: 0;\n  Descending: 1;\n")
            + table("ChartBlanksAs", "  Absent: 0;\n  Gap: 1;\n")
        )
        good = (
            table("PivotSort", "  Descending: 0;\n  Ascending: 1;\n")
            + table("SlicerSort", "  Descending: 0;\n  Ascending: 1;\n")
            + table("ChartBlanksAs", "  Gap: 0;\n")
        )
        self.assert_rule(
            "R19.js",
            self.fixture.context(self.HEADER, self.ledger(), wasm=bad),
            self.fixture.context(self.HEADER, self.ledger(), wasm=good),
            [
                ("wasm", "enum.chart_blanks_as", "absent:Absent"),
                ("wasm", "enum.slicer_sort", "ascending"),
                ("wasm", "enum.slicer_sort", "descending"),
            ],
        )

    def test_r19_py(self) -> None:
        bad = (
            "class PivotSort(IntEnum):\n    DESCENDING = 0\n    ASCENDING = 1\n\n"
            "class SlicerSort(IntEnum):\n    ASCENDING = 0\n    DESCENDING = 1\n\n"
            "class ChartBlanksAs(IntEnum):\n    ABSENT = 0\n    GAP = 1\n"
        )
        good = bad.replace("ASCENDING = 0\n    DESCENDING = 1", "DESCENDING = 0\n    ASCENDING = 1").replace(
            "    ABSENT = 0\n    GAP = 1\n", "    GAP = 0\n"
        )
        self.assert_rule(
            "R19.py",
            self.fixture.context(self.HEADER, self.ledger(), py_workbook=bad),
            self.fixture.context(self.HEADER, self.ledger(), py_workbook=good),
            [
                ("py", "enum.chart_blanks_as", "absent:Absent"),
                ("py", "enum.slicer_sort", "ascending"),
                ("py", "enum.slicer_sort", "descending"),
            ],
        )


class ReadonlyResultTest(BindingRuleTestBase):
    LEDGER = {
        "entries": {
            "record.thing": {"kind": "record", "role": "result", "js": "Thing", "py": "Thing",
                             "surfaces": ["wasm", "native", "py"]},
        }
    }  # fmt: skip

    def test_r20_js(self) -> None:
        bad = "export interface Thing {\n  readonly items: Item[];\n}\n" + _wb(
            "  listPeople(): Person[];\n  getCells(): ReadonlyArray<Cell>;\n"
        )
        good = "export interface Thing {\n  readonly items: readonly Item[];\n}\n" + _wb(
            "  listPeople(): readonly Person[];\n  getCells(): readonly Cell[];\n"
        )
        self.assert_rule(
            "R20.js",
            self.fixture.context("", self.LEDGER, wasm=bad),
            self.fixture.context("", self.LEDGER, wasm=good),
            [
                ("wasm", "record.thing", "field:items"),
                ("wasm", "getCells", "ReadonlyArray:Workbook.getCells"),
                ("wasm", "listPeople", "return"),
            ],
        )

    def test_r20_py(self) -> None:
        bad = (
            "@dataclass(frozen=True)\nclass Thing:\n    items: Tuple[Item, ...]\n    rgba: Tuple[int, int, int, int]\n"
        )
        good = "@dataclass(frozen=True)\nclass Thing:\n    items: List[Item]\n    rgba: Tuple[int, int, int, int]\n"
        self.assert_rule(
            "R20.py",
            self.fixture.context("", self.LEDGER, py_workbook=bad),
            self.fixture.context("", self.LEDGER, py_workbook=good),
            [("py", "record.thing", "field:items")],
        )


class RecordNameTest(BindingRuleTestBase):
    def test_r21_role(self) -> None:
        record = {"kind": "record", "js": "Font", "surfaces": ["c", "wasm"]}
        self.assert_rule(
            "R21.role",
            self.fixture.context("", {"entries": {"record.font": record}}),
            self.fixture.context("", {"entries": {"record.font": {**record, "role": "both"}}}),
            [("wasm", "record.font", "role")],
        )

    def test_r21_name(self) -> None:
        def ledger(js: str, py: str) -> dict:
            entry = {"kind": "record", "role": "result", "js": js, "py": py, "surfaces": ["wasm", "native", "py"]}
            return {"entries": {"record.font": entry}}

        self.assert_rule(
            "R21.name",
            self.fixture.context("", ledger("FontRecord", "FontEntry"), wasm="export interface FontRecord {\n}\n",
                                 py_workbook="class FontEntry:\n    pass\n"),
            self.fixture.context("", ledger("Font", "Font"), wasm="export interface Font {\n}\n",
                                 py_workbook="class Font:\n    pass\n"),
            [("py", "record.font", "py-name"), ("py", "record.font", "suffix"), ("wasm", "record.font", "suffix")],
        )  # fmt: skip


class IndexNameTest(BindingRuleTestBase):
    def test_r22_c(self) -> None:
        bad = (
            "FM_API fm_status_t fm_sheet_remove_validation(fm_workbook_t* wb, size_t sheet_index, size_t index);\n"
            "FM_API fm_status_t fm_sheet_get_comment_at(const fm_workbook_t* wb, size_t sheet_index, size_t index);\n"
            "FM_API fm_status_t fm_sheet_set_cell_phonetic(fm_workbook_t* wb, size_t sheet_index, const char* utf8);\n"
            "FM_API fm_status_t fm_workbook_add_font(fm_workbook_t* wb, size_t* out_index);\n"
        )
        good = (
            bad.replace("size_t index);\nFM_API fm_status_t fm_sheet_get", "size_t validation_index);\nFM_API fm_status_t fm_sheet_get")
            .replace("utf8", "text")
            .replace("out_index", "out_font_index")
        )  # fmt: skip
        self.assert_rule(
            "R22.c",
            self.fixture.context(bad, None),
            self.fixture.context(good, None),
            [
                ("c", "fm_sheet_remove_validation", "index"),
                ("c", "fm_sheet_set_cell_phonetic", "utf8"),
                ("c", "fm_workbook_add_font", "out_index"),
            ],
        )

    def test_r22_js_and_py(self) -> None:
        bad_js = _wb(
            "  removeValidation(sheet: number, index: number): void;\n  getCommentAt(sheet: number, index: number): Comment;\n"
            "  setNumberFormat(sheet: number, utf8: string): void;\n"
        )
        good_js = bad_js.replace(
            "removeValidation(sheet: number, index", "removeValidation(sheet: number, validationIndex"
        )
        good_js = good_js.replace("utf8", "formatCode")
        self.assert_rule(
            "R22.js",
            self.fixture.context("", None, wasm=bad_js),
            self.fixture.context("", None, wasm=good_js),
            [("wasm", "removeValidation", "index"), ("wasm", "setNumberFormat", "utf8")],
        )
        bad_py = _py_wb("    def remove_validation(self, sheet: int, index: int) -> None: ...\n"
                        "    def get_comment_at(self, sheet: int, index: int) -> None: ...\n")  # fmt: skip
        self.assert_rule(
            "R22.py",
            self.fixture.context("", None, py_workbook=bad_py),
            self.fixture.context("", None, py_workbook=bad_py.replace("int, index: int) -> None: ...\n    def get",
                                                                     "int, validation_index: int) -> None: ...\n    def get")),
            [("py", "remove_validation", "index")],
        )  # fmt: skip


class MissingLookupTest(BindingRuleTestBase):
    LEDGER = {
        "entries": {
            "chart_data.get": {"kind": "op", "scope": "sheet", "js": "getChartData", "py": "get_chart_data",
                               "missing": "throw", "surfaces": ["wasm", "native", "py"]},
            "image_info.get": {"kind": "op", "scope": "sheet", "js": "getImageInfo", "py": "get_image_info",
                               "missing": None, "surfaces": ["wasm", "native", "py"]},
        }
    }  # fmt: skip

    def test_r24_missing_js(self) -> None:
        self.assert_rule(
            "R24.missing.js",
            self.fixture.context("", self.LEDGER, wasm=_sheet("  getChartData(id: number): ChartData | null;\n"
                                                              "  getImageInfo(id: number): ImageInfo;\n")),
            self.fixture.context("", self.LEDGER, wasm=_sheet("  getChartData(id: number): ChartData;\n"
                                                              "  getImageInfo(id: number): ImageInfo | null;\n")),
            [("wasm", "chart_data.get", "throw"), ("wasm", "image_info.get", "null")],
        )  # fmt: skip

    def test_r24_missing_py(self) -> None:
        bad = _py_sheet('    def get_chart_data(self, object_id: int) -> "ChartData":\n        """Opens the data."""\n'
                        "    def get_image_info(self, object_id: int) -> ImageInfo: ...\n")  # fmt: skip
        good = _py_sheet(
            '    def get_chart_data(self, object_id: int) -> "ChartData":\n'
            '        """Opens the data.\n\n        Raises:\n            NotFoundError: no chart has the id.\n        """\n'
            "    def get_image_info(self, object_id: int) -> Optional[ImageInfo]: ...\n"
        )
        self.assert_rule(
            "R24.missing.py",
            self.fixture.context("", self.LEDGER, py_sheet=bad),
            self.fixture.context("", self.LEDGER, py_sheet=good),
            [("py", "chart_data.get", "throw"), ("py", "image_info.get", "null")],
        )


if __name__ == "__main__":
    unittest.main()
