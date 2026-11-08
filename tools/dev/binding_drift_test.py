#!/usr/bin/env python3
"""Regression tests for the JavaScript binding-drift seams."""

from __future__ import annotations

import importlib.util
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch

from tools.dev.binding_drift import (
    abi_checks,
    api_conventions,
    declaration_checks,
    enum_decl,
    python_checks,
    readme_checks,
    readme_identifiers,
    shape_checks,
    ts_decl,
)

# The fixture tests of the adoption rules (R29-R35) and the pending files run here as well.
from tools.dev.binding_drift.api_conventions_adoption_test import (  # noqa: F401
    AdoptionCRuleTest,
    AdoptionLedgerRuleTest,
    PendingFileTest,
)

# The fixture tests of the ledger-bound rules (R4, R7.js, R8.js, R12, R13, R14, R16, R17) run here as well.
from tools.dev.binding_drift.api_conventions_bindings_test import (  # noqa: F401
    AbbreviationTest,
    EntryPointTest,
    EntryShapeTest,
    EnumTypeTest,
    ResultShapeTest,
    ValueUnionTest,
)


def _load_codegen(name: str):
    path = Path(__file__).resolve().parents[1] / "codegen" / f"{name}.py"
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    # A dataclass resolves its annotations through the module it was defined in.
    sys.modules.setdefault(name, module)
    spec.loader.exec_module(module)
    return module


def _load_gen_native_dts():
    return _load_codegen("gen_native_dts")


class JavaScriptSurfaceTest(unittest.TestCase):
    def _fixture(self, common: str, native: str) -> tuple[Path, Path, Path]:
        root = Path(self._tmp.name)
        common_path = root / "packages" / "npm" / "common.mjs"
        native_path = root / "packages" / "npm-native" / "index.mjs"
        common_path.parent.mkdir(parents=True)
        native_path.parent.mkdir(parents=True)
        common_path.write_text(common, encoding="utf-8")
        native_path.write_text(native, encoding="utf-8")
        return root, common_path, native_path

    def setUp(self) -> None:
        self._tmp = TemporaryDirectory()

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_only_explicitly_reexported_common_names_are_counted(self) -> None:
        _, _, native = self._fixture(
            "export const ValueKind = Object.freeze({ Blank: 0 });\n"
            "export const PrivateKind = Object.freeze({ Secret: 1 });\n",
            "import { ValueKind, PrivateKind } from '../npm/common.mjs';\nexport { ValueKind };\n",
        )

        self.assertEqual(ts_decl._parse_js_constants(native), {"ValueKind": {"Blank": 0}})

    def test_missing_import_or_reexport_is_not_treated_as_a_common_surface(self) -> None:
        _, _, native = self._fixture(
            "export const ValueKind = Object.freeze({ Blank: 0 });\n",
            "export { ValueKind };\n",
        )
        self.assertEqual(ts_decl._parse_js_constants(native), {})

        native.write_text(
            "import { ValueKind } from '../npm/common.mjs';\n",
            encoding="utf-8",
        )
        self.assertEqual(ts_decl._parse_js_constants(native), {})

    def test_aliased_reexport_uses_canonical_value(self) -> None:
        _, _, native = self._fixture(
            "export const ValueKind = Object.freeze({ Blank: 0, Number: 1 });\n",
            "import { ValueKind as CanonicalValueKind } from '../npm/common.mjs';\n"
            "export { CanonicalValueKind as ValueKind };\n",
        )

        self.assertEqual(
            ts_decl._parse_js_constants(native),
            {"ValueKind": {"Blank": 0, "Number": 1}},
        )

    def test_wrong_common_value_remains_visible_to_drift_comparison(self) -> None:
        _, _, native = self._fixture(
            "export const ValueKind = Object.freeze({ Blank: 0, Number: 99 });\n",
            "import { ValueKind } from '../npm/common.mjs';\nexport { ValueKind };\n",
        )

        parsed = ts_decl._parse_js_constants(native)
        self.assertNotEqual(parsed, {"ValueKind": {"Blank": 0, "Number": 1}})

    def test_helper_requires_the_canonical_import(self) -> None:
        _, _, native = self._fixture(
            "export function mergeFunctionMetadata(base) { return base; }\n",
            "export function mergeFunctionMetadata(base) { return base; }\n",
        )
        self.assertIsNone(ts_decl._js_common_imported_function_source(native, "mergeFunctionMetadata"))

        native.write_text("import { mergeFunctionMetadata } from './elsewhere.mjs';\n", encoding="utf-8")
        self.assertIsNone(ts_decl._js_common_imported_function_source(native, "mergeFunctionMetadata"))

        native.write_text("import { mergeFunctionMetadata } from '../npm/common.mjs';\n", encoding="utf-8")
        self.assertEqual(ts_decl._js_common_imported_function_source(native, "mergeFunctionMetadata"), "return base;")

        native.write_text("import { other } from '../npm/common.mjs';\n", encoding="utf-8")
        self.assertIsNone(ts_decl._js_common_imported_function_source(native, "mergeFunctionMetadata"))

    def test_export_from_common_is_a_reexport(self) -> None:
        _, _, native = self._fixture(
            "export const ValueKind = Object.freeze({ Blank: 0 });\n"
            "export const Other = Object.freeze({ A: 1 });\n"
            "export function buildFormulon() {}\n",
            "export { ValueKind } from '../npm/common.mjs';\n",
        )
        self.assertEqual(ts_decl._parse_js_constants(native), {"ValueKind": {"Blank": 0}})

        native.write_text("export { ValueKind } from './elsewhere.mjs';\n", encoding="utf-8")
        self.assertEqual(ts_decl._parse_js_constants(native), {})

    def test_namespace_wiring_reaches_the_helper(self) -> None:
        common = (
            "export const FORMULON_MEMBERS = Object.freeze(['Workbook', 'mergeFunctionMetadata']);\n"
            "export function mergeFunctionMetadata(base) { return base; }\n"
        )
        wired = "import * as common from '../npm/common.mjs';\nconst f = common.buildFormulon(native, common);\n"
        _, common_path, native = self._fixture(common, wired)
        self.assertEqual(ts_decl._js_common_wired_function_source(native, "mergeFunctionMetadata"), "return base;")

        native.write_text("import * as common from '../npm/common.mjs';\n", encoding="utf-8")
        self.assertIsNone(ts_decl._js_common_wired_function_source(native, "mergeFunctionMetadata"))

        native.write_text("import * as common from './elsewhere.mjs';\ncommon.buildFormulon();\n", encoding="utf-8")
        self.assertIsNone(ts_decl._js_common_wired_function_source(native, "mergeFunctionMetadata"))

        native.write_text(wired, encoding="utf-8")
        common_path.write_text(common.replace(", 'mergeFunctionMetadata'", ""), encoding="utf-8")
        self.assertIsNone(ts_decl._js_common_wired_function_source(native, "mergeFunctionMetadata"))

    def test_typeof_member_counts_as_declared(self) -> None:
        body = "  mergeFunctionMetadata: typeof mergeFunctionMetadata;\n  other(): void;\n  plain: number;\n"
        members = ts_decl._extract_ts_callable_members(body)
        self.assertEqual(members, {"mergeFunctionMetadata", "other"})
        self.assertNotIn("mergeFunctionMetadata", ts_decl._extract_ts_callable_members("  plain: number;\n"))
        self.assertNotIn("x", ts_decl._extract_ts_callable_members("  x: typeof y;\n"))


class StagedSurfaceTest(unittest.TestCase):
    def _fixture(self) -> tuple[Path, Path, Path, Path, Path, Path, Path]:
        root = Path(self._tmp.name)
        source_index = root / "src-index.mjs"
        staged_index = root / "dist-index.mjs"
        source_common = root / "src-common.mjs"
        staged_common = root / "dist-common.mjs"
        source_dts = root / "src-index.d.ts"
        staged_dts = root / "dist-index.d.ts"
        source_index.write_text(
            "import { ValueKind } from '../npm/common.mjs';\nexport { ValueKind };\n",
            encoding="utf-8",
        )
        source_common.write_bytes(b"export const ValueKind = Object.freeze({ Blank: 0 });\n")
        source_dts.write_bytes(b"export const ValueKind: unknown;\n")
        staged_index.write_bytes(source_index.read_bytes().replace(b"from '../npm/common.mjs'", b"from './common.mjs'"))
        staged_common.write_bytes(source_common.read_bytes())
        staged_dts.write_bytes(source_dts.read_bytes())
        (root / "stage.mjs").write_text("if (commonImportSites !== 1) { process.exit(1); }\n", encoding="utf-8")
        return root, source_index, staged_index, source_common, staged_common, source_dts, staged_dts

    def setUp(self) -> None:
        self._tmp = TemporaryDirectory()

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_expected_rewrite_and_stale_shim_and_common_are_checked(self) -> None:
        root, source_index, staged_index, source_common, staged_common, source_dts, staged_dts = self._fixture()
        copies = (
            (source_dts, staged_dts),
            (source_index, staged_index),
            (source_common, staged_common),
        )
        with (
            patch.object(declaration_checks, "REPO_ROOT", root),
            patch.object(declaration_checks, "NODE_INDEX_MJS", source_index),
            patch.object(declaration_checks, "NODE_STAGE_MJS", root / "stage.mjs"),
            patch.object(declaration_checks, "_NODE_STAGED_COPIES", copies),
        ):
            self.assertEqual(declaration_checks.check_staged_dist(), [])

            expected, error = declaration_checks._staged_bytes(source_index, source_index.read_bytes())
            self.assertIsNone(error)
            self.assertEqual(staged_index.read_bytes(), expected)

            staged_index.write_bytes(staged_index.read_bytes() + b"// stale shim\n")
            staged_common.write_bytes(staged_common.read_bytes() + b"// stale common\n")
            problems = declaration_checks.check_staged_dist()

        self.assertEqual(len(problems), 2)
        self.assertTrue(any("dist-index.mjs" in problem for problem in problems))
        self.assertTrue(any("dist-common.mjs" in problem for problem in problems))

    def test_partial_staging_missing_common_or_shim_is_drift(self) -> None:
        for missing_name in ("common", "index"):
            with self.subTest(missing=missing_name):
                (
                    root,
                    source_index,
                    staged_index,
                    source_common,
                    staged_common,
                    source_dts,
                    staged_dts,
                ) = self._fixture()
                missing = staged_common if missing_name == "common" else staged_index
                missing.unlink()
                copies = (
                    (source_dts, staged_dts),
                    (source_index, staged_index),
                    (source_common, staged_common),
                )
                with (
                    patch.object(declaration_checks, "REPO_ROOT", root),
                    patch.object(declaration_checks, "NODE_INDEX_MJS", source_index),
                    patch.object(declaration_checks, "NODE_STAGE_MJS", root / "stage.mjs"),
                    patch.object(declaration_checks, "NODE_STAGE_MJS", root / "stage.mjs"),
                    patch.object(declaration_checks, "_NODE_STAGED_COPIES", copies),
                    patch.object(declaration_checks, "_SKIPPED", []),
                ):
                    problems = declaration_checks.check_staged_dist()
                self.assertEqual(len(problems), 1)
                self.assertIn(missing.name, problems[0])

    def test_import_site_count_follows_stage_script(self) -> None:
        root, source_index, _, _, _, _, _ = self._fixture()
        stage = root / "stage.mjs"
        two_sites = (
            "import * as common from '../npm/common.mjs';\nexport { ValueKind } from '../npm/common.mjs';\n"
        ).encode()
        with (
            patch.object(declaration_checks, "REPO_ROOT", root),
            patch.object(declaration_checks, "NODE_INDEX_MJS", source_index),
            patch.object(declaration_checks, "NODE_STAGE_MJS", stage),
        ):
            stage.write_text("if (commonImportSites !== 2) {}\n", encoding="utf-8")
            expected, error = declaration_checks._staged_bytes(source_index, two_sites)
            self.assertIsNone(error)
            self.assertEqual(expected.count(b"from './common.mjs'"), 2)

            stage.write_text("if (commonImportSites !== 1) {}\n", encoding="utf-8")
            expected, error = declaration_checks._staged_bytes(source_index, two_sites)
            self.assertIsNone(expected)
            self.assertIn("exactly 1", error)

            stage.write_text("// no count\n", encoding="utf-8")
            expected, error = declaration_checks._staged_bytes(source_index, two_sites)
            self.assertIsNone(expected)
            self.assertIn("cannot read", error)

    def test_entirely_unstaged_dist_is_skipped(self) -> None:
        (
            root,
            source_index,
            staged_index,
            source_common,
            staged_common,
            source_dts,
            staged_dts,
        ) = self._fixture()
        for staged in (staged_index, staged_common, staged_dts):
            staged.unlink()
        copies = (
            (source_dts, staged_dts),
            (source_index, staged_index),
            (source_common, staged_common),
        )
        with (
            patch.object(declaration_checks, "REPO_ROOT", root),
            patch.object(declaration_checks, "NODE_INDEX_MJS", source_index),
            patch.object(declaration_checks, "NODE_STAGE_MJS", root / "stage.mjs"),
            patch.object(declaration_checks, "_NODE_STAGED_COPIES", copies),
            patch.object(declaration_checks, "_SKIPPED", []),
        ):
            problems = declaration_checks.check_staged_dist()
            skipped = declaration_checks._SKIPPED
        self.assertEqual(problems, [])
        self.assertEqual(len(skipped), 1)
        self.assertIn("has not been staged", skipped[0])


class ReadmeFactsTest(unittest.TestCase):
    def _problems_after(self, old: str, new: str) -> list[str]:
        original = readme_checks._read
        with patch.object(readme_checks, "_read", lambda path: original(path).replace(old, new)):
            return readme_checks.check_readme_facts()

    def test_current_readmes_match_their_sources(self) -> None:
        self.assertEqual(readme_checks.check_readme_facts(), [])

    def test_each_fact_catches_a_stale_claim(self) -> None:
        stale = {
            "function-counts": ("526 Excel function names", "525 Excel function names"),
            "size-ceilings": ("1344 KiB Brotli", "1280 KiB Brotli"),
            "default-profile": ("New workbooks use `win-365-en_US`", "New workbooks use `mac-365-ja_JP`"),
            "profile-ids": (",nl_NL}", ",nl_NL,pl_PL}"),
            "node-min": ("Node 22+", "Node 20+"),
            "cli-platforms": ("`linux-arm64` are on", "`win-x64` are on"),
        }
        for fact, (old, new) in stale.items():
            with self.subTest(fact=fact):
                problems = self._problems_after(old, new)
                self.assertTrue(any(f" {fact} says " in p for p in problems), problems)

    def test_a_reworded_sentence_fails_instead_of_passing_silently(self) -> None:
        problems = self._problems_after("Excel behavior profiles across", "profiles over")
        self.assertTrue(any("profile-count sentence not found" in p for p in problems), problems)


class ReadmeIdentifiersTest(unittest.TestCase):
    def _problems_after(self, readme: str, old: str, new: str) -> list[str]:
        original = readme_identifiers._read
        target = readme_identifiers.REPO_ROOT / readme

        def read(path: Path) -> str:
            text = original(path)
            return text.replace(old, new) if path == target else text

        with patch.object(readme_identifiers, "_read", read):
            return readme_identifiers.check_readme_identifiers()

    def test_current_documents_resolve(self) -> None:
        self.assertEqual(readme_identifiers.check_readme_identifiers(), [])

    def test_unknown_c_name_is_reported(self) -> None:
        problems = self._problems_after("README.md", "formulon eval '=SUM(1,2,3)'", "fm_no_such_function(wb);")
        self.assertTrue(any("`fm_no_such_function`" in p for p in problems), problems)

    def test_unknown_js_import_is_reported(self) -> None:
        problems = self._problems_after(
            "packages/npm-native/README.md",
            "import createFormulon from '@libraz/formulon-native';",
            "import { NoSuchExport } from '@libraz/formulon-native';",
        )
        self.assertTrue(any("`NoSuchExport`" in p for p in problems), problems)

    def test_stale_workbook_method_is_reported(self) -> None:
        problems = self._problems_after("packages/npm/README.md", "wb.dispose();", "wb.delete();")
        self.assertTrue(any("`wb.delete`" in p for p in problems), problems)

    def test_unknown_factory_is_reported(self) -> None:
        problems = self._problems_after("packages/npm/README.md", "Workbook.load(", "Workbook.open(")
        self.assertTrue(any("Workbook.open" in p for p in problems), problems)

    def test_stale_python_method_is_reported(self) -> None:
        problems = self._problems_after("packages/python/README.md", "wb.recalc()", "wb.recalculate()")
        self.assertTrue(any("`wb.recalculate`" in p for p in problems), problems)

    def test_link_to_ignored_path_is_reported(self) -> None:
        problems = self._problems_after("CONTRIBUTING.md", "[Apache 2.0 license](LICENSE)", "[rules](CLAUDE.md)")
        self.assertTrue(any("git-ignored path `CLAUDE.md`" in p for p in problems), problems)


def _conventions(header: str, ledger: dict | None = None) -> api_conventions.Context:
    return api_conventions.Context(header, api_conventions.Ledger(ledger))


def _rule_hits(ctx: api_conventions.Context, rule_id: str) -> list[tuple[str, str]]:
    rules = {rule_id: api_conventions.all_rules()[rule_id]}
    return [(v.op_key, v.discriminator) for v in api_conventions.collect(ctx, rules)]


class LedgerRenameChainTest(unittest.TestCase):
    @staticmethod
    def _validate_names():
        import importlib
        import sys

        api_dir = str(Path(__file__).resolve().parents[1] / "api")
        if api_dir not in sys.path:
            sys.path.insert(0, api_dir)
        return importlib.import_module("validate_names")

    @staticmethod
    def _surfaces(**param_counts: int) -> SimpleNamespace:
        fns = {name: SimpleNamespace(params=[None] * count) for name, count in param_counts.items()}
        return SimpleNamespace(functions=fns)

    @staticmethod
    def _entries(surface: str = "c") -> dict:
        return {
            "a": {"kind": "op", "c": "fm_sheet_get_x", "surfaces": [surface], "was": {"c": "fm_sheet_get_x_at"}},
            "b": {"kind": "op", "c": "fm_sheet_get_x_at", "surfaces": [surface], "was": {"c": "fm_sheet_get_x_idx"}},
        }

    def test_c_name_handed_over_with_other_parameters_is_allowed(self) -> None:
        validate = self._validate_names()
        surfaces = self._surfaces(fm_sheet_get_x_at=5, fm_sheet_get_x_idx=4)
        self.assertEqual(validate.check_chains(self._entries(), surfaces), [])

    def test_two_cycle_with_other_parameters_is_allowed(self) -> None:
        validate = self._validate_names()
        entries = {
            "a": {"kind": "op", "c": "fm_sheet_get_p", "surfaces": ["c"], "was": {"c": "fm_sheet_get_q"}},
            "b": {"kind": "op", "c": "fm_sheet_get_q", "surfaces": ["c"], "was": {"c": "fm_sheet_get_p"}},
        }
        self.assertEqual(validate.check_chains(entries, self._surfaces(fm_sheet_get_p=4, fm_sheet_get_q=5)), [])
        self.assertTrue(validate.check_chains(entries, self._surfaces(fm_sheet_get_p=4, fm_sheet_get_q=4)))

    def test_c_name_handed_over_is_allowed_once_the_tree_is_renamed(self) -> None:
        validate = self._validate_names()
        surfaces = self._surfaces(fm_sheet_get_x=5, fm_sheet_get_x_at=4)
        self.assertEqual(validate.check_chains(self._entries(), surfaces), [])

    def test_coverage_accepts_the_former_or_the_current_name(self) -> None:
        validate = self._validate_names()
        entries = self._entries()
        for exported in (("fm_sheet_get_x_at", "fm_sheet_get_x_idx"), ("fm_sheet_get_x", "fm_sheet_get_x_at")):
            surfaces = SimpleNamespace(exported=lambda names=exported: [("c", n) for n in names])
            self.assertEqual(validate.check_coverage(entries, surfaces), [])
        surfaces = SimpleNamespace(exported=lambda: [("c", "fm_sheet_get_y")])
        self.assertTrue(validate.check_coverage(entries, surfaces))

    def test_equal_parameter_counts_stay_a_chain(self) -> None:
        validate = self._validate_names()
        problems = validate.check_chains(self._entries(), self._surfaces(fm_sheet_get_x_at=4, fm_sheet_get_x_idx=4))
        self.assertTrue(any("rename chain" in p for p in problems), problems)

    def test_binding_name_handover_stays_a_chain(self) -> None:
        validate = self._validate_names()
        entries = {
            "a": {"kind": "op", "js": "getX", "surfaces": ["wasm"], "was": {"js": "getXAt"}},
            "b": {"kind": "op", "js": "getXAt", "surfaces": ["wasm"], "was": {"js": "getXIdx"}},
        }
        problems = validate.check_chains(entries, self._surfaces())
        self.assertTrue(any("rename chain" in p for p in problems), problems)

    def test_record_role_suffix_is_checked_against_the_type_name(self) -> None:
        validate = self._validate_names()
        entry = {
            "kind": "record",
            "role": "input",
            "c": "fm_print_options_t",
            "js": "PrintOptions",
            "surfaces": ["c", "wasm"],
        }
        problems = validate.check_rules({"record.print_options": entry})
        self.assertTrue(any("PrintOptionsInput" in p for p in problems), problems)
        entry["js"] = "PrintOptionsInput"
        self.assertEqual(validate.check_rules({"record.print_options": entry}), [])


class LedgerHandleTest(unittest.TestCase):
    """The handle layer of the ledger: namespaces, `internal`, `handle:` and the `Properties` derivation."""

    COLLECTIONS = {
        "pivot": {"key": "index", "access": "handles", "handle_class": "Pivot", "plural": "pivots", "ops": []},
        "drawing_object": {"key": "natural", "by": "object_id", "access": "handles", "handle_class": "Drawing",
                           "plural": "drawings", "ops": []},
        "image": {"key": "natural", "by": "object_id", "access": "handles", "handle_class": "Picture",
                  "handle_base": "Drawing", "plural": "pictures", "ops": []},
        "chart": {"key": "natural", "by": "object_id", "access": "handles", "handle_class": "Chart",
                  "handle_base": "Drawing", "plural": "charts", "ops": []},
        "sheet": {"key": "index", "access": "handles", "plural": "sheets", "ops": []},
    }  # fmt: skip

    def setUp(self) -> None:
        self.validate = LedgerRenameChainTest._validate_names()

    @staticmethod
    def op(scope: str, js: str, **fields: object) -> dict:
        return {"kind": "op", "scope": scope, "js": js, "surfaces": ["wasm"], "exception": "x", **fields}

    def schema(self, entries: dict) -> list:
        return self.validate.check_schema({"schema": 1, "collections": self.COLLECTIONS, "entries": entries})

    def test_a_public_name_is_unique_per_scope_or_handle_class(self) -> None:
        chains = lambda entries: self.validate.check_chains(entries, None, self.COLLECTIONS)  # noqa: E731
        process = {"kind": "free", "scope": "process", "js": "evaluateFormula", "surfaces": ["wasm"]}
        sheet = self.op("sheet", "evaluateFormula")
        self.assertEqual(chains({"formula.eval": process, "formula.evaluate": sheet}), [])
        self.assertTrue(chains({"a": sheet, "b": dict(sheet)}))
        pivot = {"class": "Pivot", "supplies": ["sheet", "pivot"]}
        a = self.op("sheet", "setPivotName", handle={**pivot, "js": "rename"})
        b = self.op("sheet", "setPivotTitle", handle={**pivot, "js": "rename"})
        problems = chains({"a": a, "b": b})
        self.assertTrue(any("js#Pivot name rename" in p for p in problems), problems)
        c = self.op("sheet", "setDrawingName", handle={"class": ["Picture", "Chart"], "js": "rename"})
        self.assertEqual(chains({"a": a, "c": c}), [])
        renamed = self.op("workbook", "renameSheet", handle={"class": "Sheet", "js": "rename"})
        self.assertTrue(chains({"renamed": renamed, "clash": self.op("sheet", "rename")}))

    def test_a_record_name_moves_to_its_handle_class(self) -> None:
        record = {"kind": "record", "role": "result", "js": "ShapeProperties", "surfaces": ["wasm"],
                  "was": {"js": "Shape"}}  # fmt: skip
        handle = {"kind": "typedef", "js": "Shape", "surfaces": ["wasm"]}
        collections = {**self.COLLECTIONS, "shape": {"key": "natural", "by": "object_id", "access": "handles",
                                                     "handle_class": "Shape", "plural": "shapes"}}  # fmt: skip
        entries = {"record.shape": record, "type.shape_handle": handle}
        self.assertEqual(self.validate.check_chains(entries, None, collections), [])
        self.assertTrue(self.validate.check_chains(entries, None, self.COLLECTIONS))
        ledger = api_conventions.Ledger({"schema": 1, "collections": collections, "entries": entries})
        self.assertEqual(ledger.key_for("js", "Shape"), "type.shape_handle")
        self.assertEqual(ledger.problems, [])

    def test_handle_names_derive_from_the_raw_name_and_the_class_noun(self) -> None:
        derive = api_conventions.derive_handle_name
        self.assertEqual(derive("addPivotField", "pivot", "js"), "addField")
        self.assertEqual(derive("get_pivot_cache_worksheet_source", "pivot_cache", "py"), "get_worksheet_source")
        self.assertEqual(derive("getChart", "chart", "js"), "getProperties")
        self.assertEqual(derive("get_shape", "shape", "py"), "get_properties")
        self.assertEqual(derive("setDrawingZOrder", "drawing_object", "js"), "setZOrder")
        self.assertEqual(derive("listPivotFields", "pivot", "js"), "listFields")
        self.assertEqual(derive("renameSheet", "sheet", "js"), "rename")
        shared = {"kind": "op", "js": "setDrawingAnchor", "py": "set_drawing_anchor",
                  "handle": {"class": ["Picture", "Chart"]}}  # fmt: skip
        self.assertEqual(api_conventions.handle_name(shared, "js", self.COLLECTIONS), "setAnchor")
        shared["handle"]["js"] = "moveTo"
        self.assertEqual(api_conventions.handle_name(shared, "js", self.COLLECTIONS), "moveTo")
        self.assertEqual(
            api_conventions.common_base(["Picture", "Chart"], api_conventions.handle_bases(self.COLLECTIONS)),
            "Drawing",
        )

    def test_internal_and_handle_schema(self) -> None:
        handle = {"class": "Pivot", "supplies": ["sheet", "pivot"]}
        self.assertEqual(self.schema({"a": self.op("sheet", "addPivotField", handle=handle)}), [])
        self.assertEqual(
            self.schema({"a": self.op("sheet", "setDrawingName", handle={"class": ["Picture", "Chart"]})}), []
        )
        self.assertTrue(self.schema({"a": self.op("sheet", "addPivotField", handle={"class": "Pivots"})}))
        self.assertTrue(self.schema({"a": self.op("sheet", "addPivotField", handle={**handle, "property": False})}))
        self.assertTrue(self.schema({"a": self.op("sheet", "addPivotField", handle={**handle, "label": "x"})}))
        self.assertTrue(self.schema({"a": self.op("sheet", "x", handle=handle, internal=True)}))
        self.assertTrue(self.schema({"a": self.op("sheet", "x", internal="yes")}))
        raw_less = {"kind": "op", "scope": "sheet", "surfaces": ["wasm"], "exception": "x", "handle": handle}
        self.assertTrue(self.schema({"a": raw_less}))
        raw_less["handle"] = {**handle, "js": "getProperties"}
        self.assertEqual(self.schema({"a": raw_less}), [])
        absorbed = self.op("sheet", "addImage", internal=True, replaced_by="b", closes="R28.handles")
        self.assertEqual(self.schema({"a": absorbed, "b": self.op("sheet", "addPicture")}), [])
        record = {"kind": "record", "role": "result", "js": "X", "surfaces": ["wasm"]}
        self.assertTrue(self.schema({"a": absorbed, "b": record}))

    def test_collection_handle_fields(self) -> None:
        def problems(**fields: object) -> list:
            coll = {"key": "natural", "by": "object_id", "plural": "charts", **fields}
            data = {"schema": 1, "collections": {**self.COLLECTIONS, "chart": coll}, "entries": {}}
            return self.validate.check_schema(data)

        self.assertEqual(problems(access="handles", handle_class="Chart", handle_base="Drawing"), [])
        self.assertTrue(problems(access="list"))
        self.assertTrue(problems(handle_class="Chart"))
        self.assertTrue(problems(access="handles", handle_class="Chart", handle_base="Nothing"))
        self.assertTrue(problems(access="handles", handle_class="Chart", colour="red"))

    def test_handle_collection_records_derive_to_properties(self) -> None:
        entry = {"kind": "record", "role": "result", "c": "fm_chart_info_t", "js": "ChartProperties",
                 "surfaces": ["c", "wasm"]}  # fmt: skip
        self.assertEqual(self.validate.check_rules({"record.chart_info": entry}, self.COLLECTIONS), [])
        self.assertTrue(self.validate.check_rules({"record.chart_info": entry}, {}))
        entry["js"] = "Chart"
        self.assertTrue(self.validate.check_rules({"record.chart_info": entry}, self.COLLECTIONS))
        self.assertEqual(
            api_conventions.derive_record_type("fm_drawing_object_t", "result", {"drawing_object": "Drawing"}),
            "DrawingProperties",
        )
        self.assertEqual(
            api_conventions.derive_record_type("fm_drawing_object_t", "input", {"drawing_object": "Drawing"}),
            "DrawingObjectInput",
        )

    def test_the_ledger_only_mode_reads_no_tree(self) -> None:
        import contextlib
        import io

        try:
            import yaml  # type: ignore  # noqa: F401
        except ImportError:
            self.skipTest("PyYAML is not importable")
        with patch.object(self.validate.extract_names, "scan", side_effect=AssertionError("scanned the tree")):
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(self.validate.main(["--ledger-only"]), 0)


class MigrationTableTest(unittest.TestCase):
    """Rows of the CHANGELOG migration table for handle methods, absorbed raws and removed ops."""

    @staticmethod
    def _gen():
        import importlib

        api_dir = str(Path(__file__).resolve().parents[1] / "api")
        if api_dir not in sys.path:
            sys.path.insert(0, api_dir)
        return importlib.import_module("gen_migration_table")

    def test_handle_internal_and_removed_rows(self) -> None:
        data = {
            "schema": 1,
            "collections": {
                "pivot": {"key": "index", "access": "handles", "handle_class": "Pivot", "plural": "pivots",
                          "ops": ["binding.pivot"]},
            },
            "entries": {
                "binding.pivot": {"kind": "op", "scope": "sheet", "js": "pivot", "py": "pivot"},
                "pivot_field.add": {"kind": "op", "scope": "sheet", "c": "fm_sheet_add_pivot_field",
                                    "js": "addPivotField", "py": "add_pivot_field",
                                    "was": {"c": "fm_sheet_add_pivot_field", "js": "pivotFieldAdd"},
                                    "handle": {"class": "Pivot", "supplies": ["sheet", "pivot"]}},
                "binding.remove_drawing": {"kind": "op", "scope": "sheet", "js": "removeDrawing"},
                "shape.remove": {"kind": "op", "scope": "sheet", "c": "fm_sheet_remove_shape", "js": "removeShape",
                                 "was": {"c": "fm_sheet_remove_shape", "js": "removeShape"}, "internal": True,
                                 "replaced_by": "binding.remove_drawing"},
                "runtime_id.get": {"kind": "op", "scope": "sheet", "js": "getPivotRuntimeId", "internal": True,
                                   "was": {"js": "getPivotRuntimeId"}},
                "page_setup.set": {"kind": "op", "scope": "sheet", "c": "fm_sheet_set_page_setup",
                                   "js": "setPageSetup"},
                "fit_to_page.set": {"kind": "op", "scope": "sheet", "c": "fm_sheet_set_fit_to_page",
                                    "js": "setFitToPage", "was": {"c": "fm_sheet_set_fit_to_page",
                                                                  "js": "setFitToPage"},
                                    "drop": {"c": "fm_sheet_set_fit_to_page", "js": "setFitToPage",
                                             "closes": "R18.all"},
                                    "replaced_by": "page_setup.set"},
            },
        }  # fmt: skip
        rows = self._gen().rows(data)
        self.assertEqual(
            rows["js"],
            [
                ("`pivotFieldAdd()`", "`ws.pivot(...).addField()`"),
                ("`removeShape()`", "`removeDrawing()`"),
                ("`setFitToPage()`", "removed; use `setPageSetup()`"),
            ],
        )
        self.assertEqual(rows["c"], [("`fm_sheet_set_fit_to_page`", "removed; use `fm_sheet_set_page_setup`")])


class CountHandleOpsTest(unittest.TestCase):
    def test_ledger_and_python_counts(self) -> None:
        import importlib

        api_dir = str(Path(__file__).resolve().parents[1] / "api")
        if api_dir not in sys.path:
            sys.path.insert(0, api_dir)
        count = importlib.import_module("count_handle_ops")
        pivot = {"class": "Pivot", "supplies": ["sheet", "pivot"]}
        ledger = api_conventions.Ledger(
            {
                "schema": 1,
                "collections": {
                    "pivot": {"key": "index", "access": "handles", "handle_class": "Pivot", "plural": "pivots"},
                },
                "entries": {
                    "pivot_field.add": {"kind": "op", "scope": "sheet", "py": "add_pivot_field",
                                        "surfaces": ["py"], "handle": pivot},
                    "pivot_name.get": {"kind": "op", "scope": "sheet", "py": "get_pivot_name", "surfaces": ["py"],
                                       "handle": {**pivot, "py": "name", "property": True}},
                },
            }
        )  # fmt: skip
        with TemporaryDirectory() as tmp:
            package = Path(tmp)
            (package / "_pivot.py").write_text(
                "class Pivot:\n"
                "    @property\n    def sheet(self): ...\n"
                "    @property\n    def name(self): ...\n"
                "    def add_field(self, spec): ...\n"
                "    def _resolve(self): ...\n",
                encoding="utf-8",
            )
            result = count.compare("Pivot", ledger, package)
            self.assertEqual((result["ledger"], result["python"], result["missing"]), (2, 2, []))
            (package / "_pivot.py").write_text("class Pivot:\n    def add_field(self, spec): ...\n", encoding="utf-8")
            self.assertEqual(count.compare("Pivot", ledger, package)["missing"], ["name"])
            self.assertTrue(count.compare("Chart", ledger, package)["absent"])


class HandleDeclarationTest(unittest.TestCase):
    LEDGER = {
        "schema": 1,
        "collections": {
            "pivot": {"key": "index", "access": "handles", "handle_class": "Pivot", "plural": "pivots", "ops": []},
        },
        "entries": {
            "pivot_field.add": {"kind": "op", "scope": "sheet", "js": "addPivotField", "surfaces": ["wasm"],
                                "handle": {"class": "Pivot", "supplies": ["sheet", "pivot"]}},
            "pivot_runtime_id.get": {"kind": "op", "scope": "sheet", "js": "getPivotRuntimeId", "surfaces": ["wasm"],
                                     "internal": True},
        },
    }  # fmt: skip

    def test_internal_and_handle_raws_need_no_workbook_declaration(self) -> None:
        raw_only, _facade_only = declaration_checks.facade_split(api_conventions.Ledger(self.LEDGER))
        self.assertTrue({"addPivotField", "getPivotRuntimeId"} <= raw_only)

    def test_a_handle_method_is_declared_on_its_class_and_forwards_to_a_registration(self) -> None:
        ledger = api_conventions.Ledger(self.LEDGER)
        dts = Path("src/wasm/formulon.d.ts")
        good = "export interface Pivot {\n  addField(spec: PivotFieldInput): number;\n}\n"
        check = declaration_checks.handle_declaration_problems
        self.assertEqual(check("dts-wasm", ledger, {"addPivotField"}, good, declaration_checks.REPO_ROOT / dts), [])
        problems = check("dts-wasm", ledger, set(), "export interface Pivot {\n}\n", declaration_checks.REPO_ROOT / dts)
        self.assertEqual(len(problems), 2, problems)
        self.assertTrue(any("`Pivot.addField` is not declared" in p for p in problems), problems)


class HandleDeclarationShapeTest(unittest.TestCase):
    LEDGER = {
        "schema": 1,
        "collections": {
            "pivot": {"key": "index", "access": "handles", "handle_class": "Pivot", "plural": "pivots", "ops": []},
        },
        "entries": {
            "pivot_name.get": {"kind": "op", "scope": "sheet", "c": "fm_sheet_get_pivot_name", "js": "getPivotName",
                               "surfaces": ["wasm"],
                               "handle": {"class": "Pivot", "js": "name", "property": True,
                                          "supplies": ["sheet", "pivot"]}},
            "sheet.rename": {"kind": "op", "scope": "workbook", "c": "fm_workbook_rename_sheet", "js": "renameSheet",
                             "surfaces": ["wasm"], "handle": {"class": "Sheet", "js": "rename", "supplies": ["sheet"]}},
            "drawing_objects.list": {"kind": "list", "scope": "sheet", "js": "listDrawingObjects",
                                     "surfaces": ["wasm"], "internal": True},
        },
    }  # fmt: skip

    def test_interfaces_with_a_base_list_are_found(self) -> None:
        text = "export interface Chart extends Drawing {\n  setTitle(t: string): void;\n}\n"
        self.assertIn("setTitle", ts_decl._find_interface_body(text, "Chart", Path("x.d.ts")))

    def test_properties_sheet_methods_and_internal_lists(self) -> None:
        ledger = api_conventions.Ledger(self.LEDGER)
        raw_only, facade_only = declaration_checks.facade_split(ledger)
        self.assertTrue({"renameSheet", "listDrawingObjects", "getPivotName"} <= raw_only)
        self.assertIn("rename", facade_only)
        dts = declaration_checks.REPO_ROOT / "src/wasm/formulon.d.ts"
        text = "export interface Pivot {\n  readonly name: string;\n}\nexport interface Sheet {\n  rename(n: string): void;\n}\n"
        registered = {"getPivotName", "renameSheet"}
        self.assertEqual(declaration_checks.handle_declaration_problems("dts-wasm", ledger, registered, text, dts), [])
        as_method = text.replace("readonly name: string;", "name(): string;")
        self.assertTrue(declaration_checks.handle_declaration_problems("dts-wasm", ledger, registered, as_method, dts))


class ExtractHandleMembersTest(unittest.TestCase):
    def test_generated_handle_tables_name_the_js_members(self) -> None:
        import importlib

        api_dir = str(Path(__file__).resolve().parents[1] / "api")
        if api_dir not in sys.path:
            sys.path.insert(0, api_dir)
        extract = importlib.import_module("extract_names")
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            self.assertEqual(extract._js_handle_members(root), [])
            common = root / extract.NPM_COMMON
            common.parent.mkdir(parents=True)
            common.write_text(
                "// <facade-tables>\n"
                "const HANDLE_FORWARDERS = Object.freeze({\n  'Pivot.addField': 'addPivotField',\n"
                "  'Pivot.getProperties': null,\n});\n"
                "const HANDLE_PROPERTIES = Object.freeze({\n  'Pivot.name': 'getPivotName',\n});\n"
                "// </facade-tables>\n",
                encoding="utf-8",
            )
            self.assertEqual(extract._js_handle_members(root), ["Pivot.addField", "Pivot.getProperties", "Pivot.name"])


class ApiConventionNameTest(unittest.TestCase):
    def test_binding_names_derive_from_the_c_name(self) -> None:
        derive = api_conventions
        self.assertEqual(derive.derive_js("fm_sheet_get_hyperlink_at"), "getHyperlinkAt")
        self.assertEqual(derive.derive_py("fm_sheet_get_hyperlink_at"), "get_hyperlink_at")
        self.assertEqual(derive.derive_js("fm_workbook_get_formula_r1c1"), "getFormulaR1C1")
        self.assertEqual(derive.derive_js("fm_chart_data_get_series_count"), "getSeriesCount")
        self.assertEqual(derive.derive_js("fm_get_version"), "getVersion")
        self.assertEqual(derive.derive_type("fm_hyperlink_t"), "Hyperlink")
        self.assertEqual(derive.derive_type("fm_cell_rect_t"), "CellRect")
        self.assertEqual(derive.derive_type("fm_value_kind_t"), "ValueKind")
        self.assertEqual(derive.derive_type("fm_cf_rule_t"), "CfRule")

    def test_record_role_adds_a_suffix_once(self) -> None:
        derive = api_conventions.derive_type
        self.assertEqual(derive("fm_hyperlink_t", "result"), "Hyperlink")
        self.assertEqual(derive("fm_print_options_t", "input"), "PrintOptionsInput")
        self.assertEqual(derive("fm_hyperlink_input_t", "input"), "HyperlinkInput")
        self.assertEqual(derive("fm_iterative_t", "patch"), "IterativePatch")
        self.assertEqual(derive("fm_iterative_patch_t", "patch"), "IterativePatch")

    def test_project_and_show_are_c_verbs(self) -> None:
        self.assertEqual(api_conventions.check_c_name("fm_sheet_project_pivot"), [])
        self.assertEqual(api_conventions.check_c_name("fm_sheet_show_pivot_details"), [])

    def test_check_c_name_separates_form_from_verb(self) -> None:
        self.assertEqual(api_conventions.check_c_name("fm_get_version"), [])
        self.assertEqual(api_conventions.check_c_name("fm_workbook_load"), [])
        self.assertEqual(api_conventions.check_c_name("fm_chart_data_destroy"), [])
        self.assertEqual([r for r, _ in api_conventions.check_c_name("fm_version_string")], ["R1.form"])
        self.assertEqual([r for r, _ in api_conventions.check_c_name("fm_Sheet_get_x")], ["R1.form"])
        self.assertEqual([r for r, _ in api_conventions.check_c_name("fm_workbook")], ["R1.form"])
        self.assertEqual([r for r, _ in api_conventions.check_c_name("fm_sheet_list_invalid_cells")], ["R2.verb"])


class ApiConventionRuleTest(unittest.TestCase):
    """One violating and one conforming fixture per rule-id."""

    def assert_rule(self, rule_id: str, bad: str, good: str, expected: list[tuple[str, str]]) -> None:
        self.assertEqual(_rule_hits(_conventions(bad), rule_id), expected)
        self.assertEqual(_rule_hits(_conventions(good), rule_id), [])

    def test_r1_form(self) -> None:
        self.assert_rule(
            "R1.form",
            "FM_API fm_status_t fm_styles_get_font(const fm_workbook_t* wb, size_t font_index);\n",
            "FM_API fm_status_t fm_workbook_get_font(const fm_workbook_t* wb, size_t font_index);\n",
            [("fm_styles_get_font", "-")],
        )

    def test_r2_verb(self) -> None:
        self.assert_rule(
            "R2.verb",
            "FM_API fm_status_t fm_workbook_sheet_name(const fm_workbook_t* wb, size_t sheet_index);\n",
            "FM_API fm_status_t fm_workbook_get_sheet_name(const fm_workbook_t* wb, size_t sheet_index);\n",
            [("fm_workbook_sheet_name", "-")],
        )

    def test_r6_c(self) -> None:
        cases = {
            "sheet": (
                "FM_API fm_status_t fm_sheet_set_zoom(fm_workbook_t* wb, uint32_t sheet, uint32_t zoom);\n",
                "FM_API fm_status_t fm_sheet_set_zoom(fm_workbook_t* wb, size_t sheet_index, uint32_t zoom);\n",
                [("fm_sheet_set_zoom", "sheet")],
            ),
            "index": (
                "FM_API fm_status_t fm_workbook_remove_table(fm_workbook_t* wb, uint32_t table_index);\n",
                "FM_API fm_status_t fm_workbook_remove_table(fm_workbook_t* wb, size_t table_index);\n",
                [("fm_workbook_remove_table", "table_index")],
            ),
            "id": (
                "FM_API fm_status_t fm_workbook_get_num_fmt(const fm_workbook_t* wb, uint16_t num_fmt_id);\n",
                "FM_API fm_status_t fm_workbook_get_num_fmt(const fm_workbook_t* wb, uint32_t num_fmt_id);\n",
                [("fm_workbook_get_num_fmt", "num_fmt_id")],
            ),
            "row": (
                "FM_API fm_status_t fm_sheet_insert_rows(fm_workbook_t* wb, size_t sheet_index, int32_t row);\n",
                "FM_API fm_status_t fm_sheet_insert_rows(fm_workbook_t* wb, size_t sheet_index, uint32_t row);\n",
                [("fm_sheet_insert_rows", "row")],
            ),
            "enum": (
                "typedef enum { FM_A = 0 } fm_mode_t;\n"
                "FM_API fm_status_t fm_workbook_get_mode(const fm_workbook_t* wb, fm_mode_t* out_mode);\n",
                "typedef enum { FM_A = 0 } fm_mode_t;\n"
                "FM_API fm_status_t fm_workbook_get_mode(const fm_workbook_t* wb, int32_t* out_mode);\n",
                [("fm_workbook_get_mode", "out_mode")],
            ),
            "enum-field": (
                "typedef enum { FM_A = 0 } fm_mode_t;\ntypedef struct {\n  fm_mode_t mode;\n} fm_view_t;\n",
                "typedef enum { FM_A = 0 } fm_mode_t;\ntypedef struct {\n  int32_t mode;\n} fm_view_t;\n",
                [("fm_view_t", "mode")],
            ),
            "count": (
                "FM_API size_t fm_workbook_get_table_count(const fm_workbook_t* wb);\n",
                "FM_API fm_status_t fm_workbook_get_table_count(const fm_workbook_t* wb, size_t* out_count);\n",
                [("fm_workbook_get_table_count", "count")],
            ),
            "const-handle": (
                "FM_API fm_status_t fm_workbook_get_zoom(fm_workbook_t* wb, uint32_t* out_zoom);\n",
                "FM_API fm_status_t fm_workbook_get_zoom(const fm_workbook_t* wb, uint32_t* out_zoom);\n",
                [("fm_workbook_get_zoom", "const-handle")],
            ),
            "buffer": (
                "FM_API fm_status_t fm_sheet_match(const fm_workbook_t* wb, uint8_t* out_match, size_t cap,"
                " size_t* out_len);\n",
                "FM_API fm_status_t fm_sheet_match(const fm_workbook_t* wb, uint8_t* buf, size_t cap,"
                " size_t* out_len);\n",
                [("fm_sheet_match", "out_match")],
            ),
        }
        for name, (bad, good, expected) in cases.items():
            with self.subTest(case=name):
                self.assert_rule("R6.c", bad, good, expected)

    def test_r6_c_field(self) -> None:
        enum = "typedef enum { FM_MODE_A = 0, FM_MODE_B = 1 } fm_mode_t;\n#define FM_ORIENTATION_PORTRAIT 1u\n"
        self.assert_rule(
            "R6.c.field",
            enum + "typedef struct {\n  size_t range_count;\n  uint32_t orientation; /* FM_ORIENTATION_* */\n"
            "  uint32_t mode; /* fm_mode_t */\n  uint8_t kind; /* fm_mode_t ordinal */\n} fm_page_t;\n",
            enum + "typedef struct {\n  uint32_t range_count;\n  int32_t orientation; /* FM_ORIENTATION_* */\n"
            "  int32_t mode; /* fm_mode_t */\n  uint8_t kind; /* fm_mode_t ordinal */\n  uint64_t byte_length;\n"
            "} fm_page_t;\n",
            [("fm_page_t", "mode"), ("fm_page_t", "orientation"), ("fm_page_t", "range_count")],
        )

    def test_r15_header(self) -> None:
        decl = (
            "typedef struct fm_workbook fm_workbook_t;\n"
            "FM_API fm_status_t fm_workbook_add_font(fm_workbook_t* wb, size_t* out_font_index);\n"
        )
        self.assert_rule(
            "R15.header",
            "/* See fm_styles_add_* and fm_workbook_gone. */\ntypedef struct fm_workbook fm_workbook_t;\n"
            "/* TODO: for now only fonts. */\n"
            "FM_API fm_status_t fm_workbook_add_font(fm_workbook_t* wb, size_t* out_font_index);\n",
            "/* See fm_workbook_add_* and fm_workbook_add_font on an fm_workbook_t*; a later call adds more. */\n"
            + decl,
            [
                ("fm_styles_add_*", "unresolved"),
                ("fm_workbook_add_font", "TODO"),
                ("fm_workbook_add_font", "for now"),
                ("fm_workbook_gone", "unresolved"),
            ],
        )

    def test_r7_c_byptr(self) -> None:
        record = "typedef struct {\n  uint32_t row;\n} fm_anchor_t;\n"
        self.assert_rule(
            "R7.c.byptr",
            record + "FM_API fm_status_t fm_sheet_add_anchor(fm_workbook_t* wb, fm_anchor_t anchor);\n",
            record + "FM_API fm_status_t fm_sheet_add_anchor(fm_workbook_t* wb, const fm_anchor_t* anchor);\n",
            [("fm_sheet_add_anchor", "anchor")],
        )

    def test_r7_c_typedef(self) -> None:
        self.assert_rule(
            "R7.c.typedef",
            "typedef struct {\n  uint32_t row;\n} fm_comment;\ntypedef struct fm_chart fm_chart;\n",
            "typedef struct {\n  uint32_t row;\n} fm_comment_t;\ntypedef struct fm_chart fm_chart_t;\n",
            [("fm_chart", "-"), ("fm_comment", "-")],
        )

    def test_r7_c_typedef_exempts_a_replaced_type(self) -> None:
        header = "typedef struct {\n  uint32_t row;\n} fm_merge_range;\n"
        ledger = {
            "entries": {
                "record.merge_range": {
                    "kind": "typedef",
                    "c": "fm_merge_range",
                    "replaced_by": "record.cell_rect",
                    "closes": "R11.rect",
                }
            }
        }
        self.assertEqual(_rule_hits(_conventions(header, ledger), "R7.c.typedef"), [])

    def test_r8_c(self) -> None:
        self.assert_rule(
            "R8.c",
            "FM_API void fm_buffer_free(uint8_t* bytes);\nFM_API const char* fm_version_string(void);\n",
            "FM_API void fm_buffer_destroy(uint8_t* bytes);\nFM_API const char* fm_get_version(void);\n"
            "FM_API fm_status_t fm_workbook_save(fm_workbook_t* wb);\n",
            [("fm_buffer_free", "-"), ("fm_version_string", "-")],
        )

    def test_r8_c_allows_the_process_version_number(self) -> None:
        self.assert_rule(
            "R8.c",
            "FM_API uint32_t fm_get_build_number(void);\n",
            "FM_API uint32_t fm_get_version_number(void);\nFM_API const char* fm_get_version(void);\n",
            [("fm_get_build_number", "-")],
        )

    def test_r8_c_judges_the_ledger_target_name(self) -> None:
        header = "FM_API const char* fm_version_string(void);\n"
        ledger = {
            "entries": {"version.get": {"kind": "free", "c": "fm_get_version", "was": {"c": "fm_version_string"}}}
        }
        self.assertEqual(_rule_hits(_conventions(header), "R8.c"), [("fm_version_string", "-")])
        self.assertEqual(_rule_hits(_conventions(header, ledger), "R8.c"), [])

    def test_r11_axis(self) -> None:
        column = "FM_API fm_status_t fm_sheet_set_column_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last, int32_t hidden);\n"  # noqa: E501
        self.assert_rule(
            "R11.axis",
            column
            + "FM_API fm_status_t fm_sheet_set_row_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t hidden);\n",  # noqa: E501
            column
            + "FM_API fm_status_t fm_sheet_set_row_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last, int32_t hidden);\n",  # noqa: E501
            [("fm_sheet_set_row_hidden", "-")],
        )

    def test_r11_axis_auto_height_and_clear_height(self) -> None:
        width = (
            "FM_API fm_status_t fm_sheet_set_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last, double width);\n"  # noqa: E501
            "FM_API fm_status_t fm_sheet_clear_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);\n"  # noqa: E501
        )
        self.assert_rule(
            "R11.axis",
            width
            + "FM_API fm_status_t fm_sheet_set_row_auto_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row, double height);\n"  # noqa: E501
            + "FM_API fm_status_t fm_sheet_clear_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row);\n",  # noqa: E501
            width
            + "FM_API fm_status_t fm_sheet_set_row_auto_height(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last, double height);\n"  # noqa: E501
            + "FM_API fm_status_t fm_sheet_clear_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last);\n",  # noqa: E501
            [("fm_sheet_clear_row_height", "-"), ("fm_sheet_set_row_auto_height", "-")],
        )

    def test_r11_rect(self) -> None:
        rect = "typedef struct {\n  uint32_t first_row;\n  uint32_t first_col;\n  uint32_t last_row;\n  uint32_t last_col;\n}"
        self.assert_rule(
            "R11.rect",
            rect + " fm_merge_range;\n"
            "FM_API fm_status_t fm_sheet_set_area(fm_workbook_t* wb, size_t sheet_index, uint32_t row0, uint32_t col0,"
            " uint32_t row1, uint32_t col1);\n",
            rect + " fm_cell_rect_t;\n"
            "FM_API fm_status_t fm_sheet_set_area(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* rect);\n"
            # Two independent spans are not a rectangle.
            "typedef struct {\n  uint32_t first_row;\n  uint32_t last_row;\n  int32_t has_cols;\n"
            "  uint32_t first_col;\n  uint32_t last_col;\n} fm_titles_t;\n",
            [("fm_merge_range", "first_row"), ("fm_sheet_set_area", "row0")],
        )

    def test_r11_rect_anchor_and_span(self) -> None:
        self.assert_rule(
            "R11.rect",
            "FM_API fm_status_t fm_sheet_set_pivot_anchor(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,"
            " uint32_t anchor_row, uint32_t anchor_col, uint32_t span_rows, uint32_t span_cols);\n",
            "FM_API fm_status_t fm_sheet_set_pivot_anchor(fm_workbook_t* wb, size_t sheet_index, size_t pivot_index,"
            " const fm_cell_rect_t* rect);\n",
            [("fm_sheet_set_pivot_anchor", "anchor_row")],
        )


class ApiConventionBaselineTest(unittest.TestCase):
    HEADER = (
        "FM_API fm_status_t fm_workbook_sheet_name(const fm_workbook_t* wb, size_t sheet_index);\n"
        "FM_API void fm_buffer_free(uint8_t* bytes);\n"
    )

    def setUp(self) -> None:
        self._tmp = TemporaryDirectory()
        self.dir = Path(self._tmp.name) / "baseline"

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def problems(self, header: str | None = None) -> list[str]:
        return api_conventions.check_api_conventions(self.dir, _conventions(header or self.HEADER))

    def test_missing_files_read_as_empty_and_fail(self) -> None:
        self.assertEqual(len(self.problems()), 3)

    def test_written_baseline_passes_and_holds_one_line_per_violation(self) -> None:
        violations = api_conventions.collect(_conventions(self.HEADER))
        counts = api_conventions.write_baseline(violations, self.dir)
        self.assertEqual(counts, {"c.txt": 3, "wasm.txt": 0, "native.txt": 0, "py.txt": 0})
        self.assertEqual(self.problems(), [])
        for line in (self.dir / "c.txt").read_text(encoding="utf-8").splitlines():
            self.assertEqual(len(line.split("\t")), 4)

    def test_stale_line_fails(self) -> None:
        api_conventions.write_baseline(api_conventions.collect(_conventions(self.HEADER)), self.dir)
        fixed = self.HEADER.replace("fm_workbook_sheet_name", "fm_workbook_get_sheet_name")
        problems = self.problems(fixed)
        self.assertEqual(len(problems), 1)
        self.assertIn("stale line in c.txt", problems[0])

    def test_message_is_not_part_of_the_identity(self) -> None:
        api_conventions.write_baseline(api_conventions.collect(_conventions(self.HEADER)), self.dir)
        path = self.dir / "c.txt"
        rows = [line.split("\t") for line in path.read_text(encoding="utf-8").splitlines()]
        path.write_text("".join("\t".join(row[:3] + ["outdated"]) + "\n" for row in rows), encoding="utf-8")
        self.assertEqual(self.problems(), [])

    def test_line_in_another_surface_file_does_not_count(self) -> None:
        api_conventions.write_baseline(api_conventions.collect(_conventions(self.HEADER)), self.dir)
        (self.dir / "py.txt").write_text((self.dir / "c.txt").read_text(encoding="utf-8"), encoding="utf-8")
        (self.dir / "c.txt").write_text("", encoding="utf-8")
        self.assertEqual(len(self.problems()), 6)

    def test_malformed_and_repeated_lines_fail(self) -> None:
        self.dir.mkdir(parents=True)
        (self.dir / "c.txt").write_text(
            "# comment\nR8.c\tfm_buffer_free\t-\tm\nR8.c\tfm_buffer_free\t-\tm\n", encoding="utf-8"
        )
        problems = api_conventions.check_api_conventions(
            self.dir, _conventions("FM_API void fm_buffer_free(uint8_t* b);\n")
        )
        self.assertTrue(any("is not `<rule-id>" in p for p in problems), problems)
        self.assertTrue(any("repeats" in p for p in problems), problems)

    def test_op_key_comes_from_the_ledger_through_the_former_name(self) -> None:
        ledger = {
            "entries": {
                "sheet.get_name": {
                    "kind": "op",
                    "c": "fm_workbook_get_sheet_name",
                    "was": {"c": "fm_workbook_sheet_name"},
                }
            }
        }  # noqa: E501
        violations = api_conventions.collect(_conventions(self.HEADER, ledger))
        self.assertIn(("R2.verb", "sheet.get_name"), [(v.rule, v.op_key) for v in violations])

    def test_two_entries_claiming_one_name_is_a_ledger_problem(self) -> None:
        ledger = {"entries": {"a": {"c": "fm_workbook_sheet_name"}, "b": {"was": {"c": "fm_workbook_sheet_name"}}}}
        ctx = _conventions(self.HEADER, ledger)
        problems = api_conventions.check_api_conventions(self.dir, ctx)
        self.assertTrue(any("both claim" in p for p in problems), problems)

    def _handoff_ledger(self) -> dict:
        return {
            "entries": {
                "by_cell": {"c": "fm_sheet_get_comment", "was": {"c": "fm_sheet_get_comment_at"}},
                "by_index": {"c": "fm_sheet_get_comment_at", "was": {"c": "fm_sheet_get_comment_at_index"}},
            }
        }

    def test_name_handed_to_a_function_with_other_parameters_is_not_a_double_claim(self) -> None:
        header = (
            "FM_API fm_status_t fm_sheet_get_comment_at(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col);\n"
            "FM_API fm_status_t fm_sheet_get_comment_at_index(fm_workbook_t* wb, uint32_t sheet, uint32_t index);\n"
        )
        ctx = _conventions(header, self._handoff_ledger())
        ctx.ledger.key_for("c", "")
        self.assertEqual(ctx.ledger.problems, [])

    def test_name_handed_to_a_function_with_equal_parameters_is_a_double_claim(self) -> None:
        header = (
            "FM_API fm_status_t fm_sheet_get_comment_at(fm_workbook_t* wb, uint32_t sheet, uint32_t row);\n"
            "FM_API fm_status_t fm_sheet_get_comment_at_index(fm_workbook_t* wb, uint32_t sheet, uint32_t index);\n"
        )
        ctx = _conventions(header, self._handoff_ledger())
        ctx.ledger.key_for("c", "")
        self.assertTrue(any("both claim" in p for p in ctx.ledger.problems), ctx.ledger.problems)

    def test_absent_ledger_is_empty(self) -> None:
        ledger = api_conventions.load_ledger(Path(self._tmp.name) / "names.yaml")
        self.assertFalse(ledger.present)
        self.assertEqual(ledger.problems, [])

    def test_ledger_file_is_read(self) -> None:
        try:
            import yaml  # type: ignore  # noqa: F401
        except ImportError:
            self.skipTest("PyYAML not importable")
        path = Path(self._tmp.name) / "names.yaml"
        path.write_text("schema: 1\nentries:\n  version.get: { kind: free, c: fm_get_version }\n", encoding="utf-8")
        self.assertEqual(api_conventions.load_ledger(path).key_for("c", "fm_get_version"), "version.get")
        path.write_text("entries: {}\n", encoding="utf-8")
        self.assertEqual(len(api_conventions.load_ledger(path).problems), 1)

    def test_another_module_registers_into_the_same_registry(self) -> None:
        @api_conventions.rule("R99.fixture")
        def _fixture(ctx: api_conventions.Context):
            yield api_conventions.Violation("R99.fixture", "py", "op", "-", "fixture")

        try:
            self.assertIn("R99.fixture", api_conventions.all_rules())
            self.assertEqual(len(self.problems()), 4)
        finally:
            del api_conventions._RULES["R99.fixture"]

    def test_every_c_rule_is_registered(self) -> None:
        expected = {
            "R1.form",
            "R2.verb",
            "R6.c",
            "R7.c.byptr",
            "R7.c.typedef",
            "R8.c",
            "R11.axis",
            "R11.rect",
            "R6.c.field",
            "R15.header",
            "R32.struct_size",
            "R33.lifetime",
        }
        self.assertLessEqual(expected, set(api_conventions.all_rules()))


class HeaderErrorCodesTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = TemporaryDirectory()
        root = Path(self._tmp.name)
        self.header = root / "formulon_c.h"
        self.status = root / "formulon_status.h"
        self.error_h = root / "error.h"
        self.error_h.write_text("enum class FormulonErrorCode : int32_t {\n  kOk = 0,\n  kInvalidArgument = 2,\n};\n")
        self.status.write_text("#define FM_STATUS_OK (0)\n#define FM_STATUS_INVALID_ARGUMENT (2)\n")

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def problems(self, header: str) -> list[str]:
        self.header.write_text(header, encoding="utf-8")
        with (
            patch.object(abi_checks, "REPO_ROOT", Path(self._tmp.name)),
            patch.object(abi_checks, "CAPI_HEADER", self.header),
            patch.object(abi_checks, "CAPI_STATUS_HEADER", self.status),
            patch.object(abi_checks, "ERROR_H", self.error_h),
        ):
            return abi_checks.check_header_error_codes()

    def test_status_macro_quotes_match_the_defines(self) -> None:
        self.assertEqual(self.problems("/** @return `FM_STATUS_INVALID_ARGUMENT`\n *         (2). */\n"), [])
        self.assertEqual(len(self.problems("/** @return `FM_STATUS_INVALID_ARGUMENT` (3). */\n")), 1)
        self.assertEqual(len(self.problems("/** @return `FM_STATUS_NO_SUCH` (2). */\n")), 1)

    def test_enumerator_quotes_are_still_accepted(self) -> None:
        self.assertEqual(self.problems("/** @return `kInvalidArgument` (2). */\n"), [])
        self.assertEqual(len(self.problems("/** @return `kInvalidArgument` (7). */\n")), 1)

    def test_a_header_quoting_nothing_fails(self) -> None:
        problems = self.problems("/** @return `FM_STATUS_OK` on success. */\n")
        self.assertEqual(len(problems), 1)
        self.assertIn("quotes no status code", problems[0])


class DeclarationTableTest(unittest.TestCase):
    def test_enum_ordinals_are_read_from_a_const_table(self) -> None:
        text = "export const Mode: Readonly<{\n  A: 0;\n  B: 1;\n  C: 5;\n}>;\nexport type Mode = 0 | 1 | 5;\n"
        self.assertEqual(enum_decl._extract_ts_enum(text, "Mode"), [0, 1, 5])
        self.assertIsNone(enum_decl._extract_ts_enum(text, "Other"))

    def test_record_fields_are_read_through_readonly(self) -> None:
        body = "  readonly alpha: number;\n  beta?: string;\n  readonly gamma?: boolean;\n"
        self.assertEqual(set(shape_checks._TS_FIELD_RE.findall(body)), {"alpha", "beta", "gamma"})


class DeclarationEntryShapeTest(unittest.TestCase):
    """The registration checks against a throwaway tree shaped like the published entry points."""

    WASM_CPP = (
        'class_<Workbook>("Workbook")\n'
        '  .function("getSheetIndex", &W::a)\n  .function("save", &W::b)\n  .function("getMemoryUsage", &W::e)\n'
        '  .property("disposed", &W::c)\n  .class_function("load", &W::d);\n'
        'function("getVersion", &getVersion);\n'
    )
    NODE_CLASS = (
        'InstanceMethod<&Workbook::A>("getSheetIndex"), InstanceMethod<&Workbook::B>("save"),\n'
        'InstanceMethod<&Workbook::E>("getMemoryUsage"), InstanceAccessor<&Workbook::C>("disposed", nullptr),\n'
        'StaticMethod<&Workbook::D>("load")\n'
    )
    NODE_ADDON = 'exports.Set("getVersion", f);\nexports.Set("Workbook", w);\nexports.Set("__setErrorClass", g);\n'
    WORKBOOK = (
        "export interface Workbook {\n  readonly disposed: boolean;\n"
        "  getSheetIndex(name: string): number | null;\n  save(format?: WorkbookFormat): Uint8Array;\n"
        "  getMemoryUsage(): number;\n"
        "  setValue(sheet: number, row: number, col: number, value: number): void;\n"
        "  sheet(nameOrIndex: string | number): Sheet;\n%s}\n"
        "export interface WorkbookCtor {\n  load(bytes: Uint8Array): Workbook;\n}\n"
    )
    ENTRY = (
        "export interface Formulon {\n  Workbook: WorkbookCtor;\n  getVersion(): string;\n"
        "  mergeFunctionMetadata: typeof mergeFunctionMetadata;\n%s}\n"
    )
    COMMON = (
        "export const FORMULON_MEMBERS = Object.freeze(['Workbook', 'mergeFunctionMetadata']);\n"
        "export function mergeFunctionMetadata(base) { return base; }\n"
        "export function installWorkbookExtensions(W) {\n  W.prototype.setValue = f;\n  W.prototype.sheet = g;\n}\n"
    )

    def setUp(self) -> None:
        self._tmp = TemporaryDirectory()
        self.root = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def write(self, rel: str, text: str) -> Path:
        path = self.root / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
        return path

    def tree(self, *, wasm_cpp=None, node_class=None, node_addon=None, wasm_dts=None, node_dts=None,
             common=None, native_index=None, entry=None, ledger=None) -> dict:  # fmt: skip
        entry_text = entry if entry is not None else "export { FormulonError } from './common.js';\n"
        native_dts = node_dts if node_dts is not None else self.WORKBOOK % ""
        return {
            "REPO_ROOT": self.root,
            # The fixture entry point ships one pure-JS helper.
            "NODE_PURE_JS_FREE_FUNCTIONS": {"mergeFunctionMetadata"},
            "LEDGER_PATH": self.write("names.yaml", ledger) if ledger is not None else self.root / "absent.yaml",
            "WASM_BINDINGS_CPP": self.write("wasm.cpp", wasm_cpp if wasm_cpp is not None else self.WASM_CPP),
            "WASM_DTS": self.write(
                "wasm.d.ts", wasm_dts if wasm_dts is not None else self.WORKBOOK % "" + self.ENTRY % ""
            ),
            "NODE_WORKBOOK_CLASS_CC": self.write("class.cc", node_class if node_class is not None else self.NODE_CLASS),
            "NODE_ADDON_CC": self.write("addon.cc", node_addon if node_addon is not None else self.NODE_ADDON),
            "NODE_DTS": self.write("node.d.ts", native_dts + self.ENTRY % ""),
            "NPM_COMMON_MJS": self.write("packages/npm/common.mjs", common if common is not None else self.COMMON),
            "NODE_INDEX_MJS": self.write(
                "packages/npm-native/index.mjs",
                native_index
                if native_index is not None
                else "import * as common from '../npm/common.mjs';\n"
                "common.installWorkbookExtensions(native.Workbook);\n"
                "const formulon = common.buildFormulon(native, common);\n",
            ),
            "NPM_ENTRY_MJS": (
                self.write("packages/npm/index.mjs", entry_text),
                self.write("packages/npm/threads.mjs", entry_text),
            ),
        }  # fmt: skip

    def run_check(self, check, **kwargs) -> list[str]:
        patches = [patch.object(declaration_checks, name, value) for name, value in self.tree(**kwargs).items()]
        for p in patches:
            p.start()
        try:
            return check()
        finally:
            for p in reversed(patches):
                p.stop()

    def test_a_matching_wasm_tree_is_clean(self) -> None:
        self.assertEqual(self.run_check(declaration_checks.check_dts_wasm), [])

    def test_wasm_pure_js_members_and_accessor_are_checked(self) -> None:
        problems = self.run_check(
            declaration_checks.check_dts_wasm,
            wasm_cpp=self.WASM_CPP.replace('  .property("disposed", &W::c)\n', "").replace(
                'function("getVersion"', 'function("getBuild"'
            ),
            wasm_dts=self.WORKBOOK.replace("readonly disposed: boolean;", "isValid(): boolean;") % "" + self.ENTRY % "",
        )  # fmt: skip
        text = "\n".join(problems)
        self.assertIn("Workbook / Sheet instance methods mismatch", text)
        self.assertIn("free functions mismatch", text)
        self.assertIn("getBuild", text)
        self.assertIn("Workbook accessors", text)

    def test_a_matching_node_tree_is_clean(self) -> None:
        self.assertEqual(self.run_check(declaration_checks.check_dts_node), [])

    def test_node_requires_the_internal_export_and_the_formulon_surface(self) -> None:
        problems = self.run_check(
            declaration_checks.check_dts_node,
            node_addon=self.NODE_ADDON.replace(
                'exports.Set("__setErrorClass", g);\n', 'exports.Set("evalFormula", h);\n'
            ),
            node_class=self.NODE_CLASS.replace('InstanceAccessor<&Workbook::C>("disposed", nullptr),', ""),
        )  # fmt: skip
        text = "\n".join(problems)
        self.assertIn("does not export the internal ['__setErrorClass']", text)
        self.assertIn("free functions mismatch", text)
        self.assertIn("evalFormula", text)
        self.assertIn("Workbook accessors", text)

    def test_a_matching_helper_tree_is_clean(self) -> None:
        self.assertEqual(self.run_check(declaration_checks.check_pure_js_helpers), [])

    def test_helpers_must_come_from_common_and_extensions_must_be_installed(self) -> None:
        problems = self.run_check(
            declaration_checks.check_pure_js_helpers,
            common=self.COMMON.replace("W.prototype.sheet = g;\n", ""),
            native_index="export function mergeFunctionMetadata(base) { return base; }\n",
            entry="export const X = 1;\n",
        )
        text = "\n".join(problems)
        self.assertIn("must import ../npm/common.mjs as a namespace", text)
        self.assertIn("does not add `Workbook.sheet`", text)
        self.assertIn("must take `installWorkbookExtensions` from ../npm/common.mjs", text)
        self.assertIn("does not re-export from './common.js'", text)

    LEDGER = (
        "schema: 1\ncollections:\n"
        "  table: { key: index, access: list, plural: tables, ops: [table.get, table.get_count] }\n"
        "entries:\n"
        "  table.get: { kind: op, scope: workbook, c: fm_workbook_get_table, js: getTable, surfaces: [c, wasm] }\n"
        "  table.get_count: { kind: op, scope: workbook, c: fm_workbook_get_table_count, js: getTableCount,"
        " surfaces: [c, wasm] }\n"
        "  zoom.get: { kind: op, scope: sheet, c: fm_sheet_get_zoom, js: getZoom, surfaces: [c, wasm],"
        " drop: { js: getZoomLevel, closes: R23.js } }\n"
    )

    def test_raw_only_registrations_and_sheet_methods_are_matched(self) -> None:
        cpp = self.WASM_CPP.replace(
            '.function("save", &W::b)\n',
            '.function("save", &W::b)\n  .function("getTableCount", &W::t)\n  .function("getTable", &W::u)\n'
            '  .function("getZoomLevel", &W::v)\n  .function("getZoom", &W::z)\n',
        )
        sheet = (
            "export interface Sheet {\n  readonly name: string;\n  readonly index: number;\n  getZoom(): number;\n}\n"
        )
        wasm_dts = self.WORKBOOK % "  listTables(): readonly TableEntry[];\n" + sheet + self.ENTRY % ""
        self.assertEqual(self.run_check(declaration_checks.check_dts_wasm, wasm_cpp=cpp, wasm_dts=wasm_dts,
                                        ledger=self.LEDGER), [])  # fmt: skip
        problems = self.run_check(declaration_checks.check_dts_wasm, wasm_cpp=cpp, wasm_dts=wasm_dts)
        text = "\n".join(problems)
        self.assertIn("getTableCount", text)
        self.assertIn("listTables", text)
        problems = self.run_check(
            declaration_checks.check_dts_wasm,
            wasm_cpp=cpp,
            wasm_dts=wasm_dts.replace("  getZoom(): number;\n", ""),
            ledger=self.LEDGER,
        )
        self.assertIn("getZoom", "\n".join(problems))

    def test_facade_built_ops_need_no_registration(self) -> None:
        ledger = (
            self.LEDGER
            + "  binding.get_values: { kind: op, scope: sheet, js: getValues, surfaces: [wasm], facade: handwritten }\n"
            "  fonts.list: { kind: list, scope: workbook, from: [table.get_count, table.get], js: listFonts,"
            " surfaces: [wasm] }\n"
            "  sheet.add: { kind: op, scope: workbook, c: fm_workbook_add_sheet, js: addSheet, surfaces: [c, wasm],"
            " facade: handwritten }\n"
        )
        sheet = "export interface Sheet {\n  getValues(): Value[][];\n}\n"
        members = "  listFonts(): readonly Font[];\n  addSheet(name: string): Sheet;\n"
        wasm_dts = self.WORKBOOK % members + sheet + self.ENTRY % ""
        text = "\n".join(self.run_check(declaration_checks.check_dts_wasm, wasm_dts=wasm_dts, ledger=ledger))
        self.assertNotIn("getValues", text)
        self.assertNotIn("listFonts", text)
        self.assertIn("addSheet", text)

    def test_ooxml_layers_and_facade_members_are_matched(self) -> None:
        cpp = self.WASM_CPP.replace(
            '.function("save", &W::b)\n',
            '.function("save", &W::b)\n  .function("addFont", &W::f)\n  .function("getCellXfIndex", &W::g)\n',
        )
        ooxml = (
            "export interface WorkbookOoxml {\n  addFont(font: Font): number;\n}\n"
            "export interface SheetOoxml {\n  getCellXfIndex(a1: string): number;\n}\n"
        )
        members = "  readonly ooxml: WorkbookOoxml;\n"
        entry = self.ENTRY % "  evaluateTypedFormula(formula: string): Value;\n"
        wasm_dts = self.WORKBOOK % members + ooxml + entry
        self.assertEqual(self.run_check(declaration_checks.check_dts_wasm, wasm_cpp=cpp, wasm_dts=wasm_dts), [])
        problems = self.run_check(
            declaration_checks.check_dts_wasm,
            wasm_cpp=cpp,
            wasm_dts=wasm_dts.replace("  addFont(font: Font): number;\n", "").replace(
                "evaluateTypedFormula", "evaluateTaggedFormula"
            ),
        )
        text = "\n".join(problems)
        self.assertIn("addFont", text)
        self.assertIn("evaluateTaggedFormula", text)

    def test_guarded_module_functions_count_as_free_functions(self) -> None:
        node_class = (
            self.NODE_CLASS + 'SetGuardedModuleFunction<&CompareText, 0x4ULL, 0x0ULL>(env, exports, "compareText");\n'
        )
        node_dts = self.WORKBOOK % "  getMemoryUsage(): number;\n"
        node_dts += self.ENTRY % "  compareText(a: string, b: string): number;\n"
        problems = self.run_check(declaration_checks.check_dts_node, node_class=node_class, node_dts=node_dts)
        self.assertEqual(problems, [])
        problems = self.run_check(declaration_checks.check_dts_node, node_dts=node_dts)
        self.assertIn("compareText", "\n".join(problems))

    FACADE_COMMON = (
        "export const FORMULON_MEMBERS = Object.freeze(['Workbook', 'mergeFunctionMetadata']);\n"
        "const RUNTIME_MEMBERS = new Set(['evalFormula']);\n"
        "export function mergeFunctionMetadata(base) { return base; }\n"
        "function makeFacade(raw) {\n  class Sheet { setValue() {} }\n  class Workbook { sheet() {} }\n"
        "  return Workbook;\n}\n"
        "export function buildFormulon(runtime, shared) {\n  const Workbook = makeFacade(runtime.Workbook);\n}\n"
    )

    def test_facade_wiring_replaces_the_extension_installer(self) -> None:
        native = (
            "import * as common from '../npm/common.mjs';\nconst formulon = common.buildFormulon(native, common);\n"
        )
        clean = self.run_check(declaration_checks.check_pure_js_helpers, common=self.FACADE_COMMON, native_index=native)
        self.assertEqual(clean, [])
        problems = self.run_check(
            declaration_checks.check_pure_js_helpers,
            common=self.FACADE_COMMON.replace("['evalFormula']", "['Workbook', 'evalFormula']").replace(
                "makeFacade(runtime.Workbook)", "runtime.Workbook"
            )
            + "export function installWorkbookExtensions(W) {}\n",
            native_index="import * as common from '../npm/common.mjs';\n"
            "common.installWorkbookExtensions(native.Workbook);\nconst formulon = common.buildFormulon(native, common);\n",
        )
        text = "\n".join(problems)
        self.assertIn("still `installWorkbookExtensions`", text)
        self.assertIn("does not call `makeFacade`", text)
        self.assertIn("still copies the raw `Workbook`", text)
        self.assertIn("index.mjs still calls `installWorkbookExtensions`", text)

    def test_enum_tables_follow_the_ledger_and_skip_absent_members(self) -> None:
        header = self.write(
            "h.h", "typedef enum { FM_BLANKS_ABSENT = -1, FM_BLANKS_GAP = 0, FM_BLANKS_ZERO = 1 } fm_blanks_t;\n"
        )
        ledger = self.write(
            "names.yaml",
            "schema: 1\ncollections: {}\nentries:\n"
            "  enum.blanks: { kind: enum, c: fm_blanks_t, js: Blanks, surfaces: [c, wasm] }\n",
        )
        dts = "export const Blanks: Readonly<{\n  Gap: 0;\n  Zero: 1;\n}>;\n"
        with (
            patch.object(declaration_checks, "CAPI_HEADER", header),
            patch.object(declaration_checks, "LEDGER_PATH", ledger),
        ):
            sources = declaration_checks._enum_sources(dts)
        self.assertEqual(sources["Blanks"][1], "fm_blanks_t")
        c_values = enum_decl._extract_c_typedef_enum(header.read_text(), "fm_blanks_t")
        self.assertEqual(declaration_checks._present(c_values), enum_decl._extract_ts_enum(dts, "Blanks"))
        self.assertNotEqual(c_values, enum_decl._extract_ts_enum(dts, "Blanks"))

    def test_a_patch_only_member_is_left_out_of_the_table(self) -> None:
        header = self.write(
            "h.h",
            "typedef enum { FM_CHART_TITLE_NONE = 0, FM_CHART_TITLE_TEXT = 1, FM_CHART_TITLE_REF = 2 } "
            "fm_chart_title_kind_t;\n",
        )
        ledger = self.write(
            "names.yaml",
            "schema: 1\ncollections: {}\nentries:\n"
            "  enum.chart_title_kind:\n"
            "    { kind: enum, c: fm_chart_title_kind_t, js: ChartTitleKind, surfaces: [c, wasm] }\n",
        )

        def problems(table: str) -> list:
            dts = self.write("w.d.ts", "export const ChartTitleKind: Readonly<{\n" + table + "}>;\n")
            with (
                patch.object(declaration_checks, "CAPI_HEADER", header),
                patch.object(declaration_checks, "LEDGER_PATH", ledger),
                patch.object(declaration_checks, "WASM_DTS", dts),
                patch.object(declaration_checks, "REPO_ROOT", dts.parent),
                patch.object(declaration_checks, "_DTS_ENUM_SOURCES", {}),
                patch.object(declaration_checks, "_check_js_constant_tables", lambda _text: []),
                patch.object(declaration_checks, "check_string_tables", lambda *_args: []),
            ):
                return declaration_checks.check_dts_enums()

        self.assertEqual(problems("  Text: 1;\n  Ref: 2;\n"), [])
        self.assertTrue(any("patch-only" in p for p in problems("  None: 0;\n  Text: 1;\n  Ref: 2;\n")))
        self.assertTrue(any("has ordinals" in p for p in problems("  Text: 1;\n")))

    def test_a_member_missing_from_common_members_or_the_interface_is_reported(self) -> None:
        problems = self.run_check(
            declaration_checks.check_pure_js_helpers,
            common=self.COMMON.replace(", 'mergeFunctionMetadata'", ""),
        )
        self.assertTrue(any("in FORMULON_MEMBERS" in p for p in problems), problems)

        problems = self.run_check(
            declaration_checks.check_pure_js_helpers,
            wasm_dts=self.WORKBOOK % "" + "export interface Formulon {\n  mergeFunctionMetadata: number;\n}\n",
        )
        self.assertTrue(any("does not declare `Formulon.mergeFunctionMetadata`" in p for p in problems), problems)

    def test_a_helper_missing_from_the_formulon_interface_is_reported(self) -> None:
        problems = self.run_check(
            declaration_checks.check_pure_js_helpers,
            wasm_dts=self.WORKBOOK % "" + "export interface Formulon {\n  getVersion(): string;\n}\n",
        )
        self.assertTrue(any("does not declare `Formulon.mergeFunctionMetadata`" in p for p in problems), problems)


class DocTokenTest(unittest.TestCase):
    """The doc-comment half of the identifier check, on throwaway declaration text and docstrings."""

    DTS = (
        "export interface CellRect { readonly firstRow: number }\n"
        "export const ValueKind: Readonly<{ Error: 4 }>;\n"
        "export interface WorkbookCtor { load(bytes: Uint8Array): Workbook }\n"
        "export interface Workbook {\n"
        "  /** Saves; see `save()`, `Workbook.getSheetIndex` and `Workbook.load`, `ValueKind.Error`, `firstRow`,\n"
        "   *  `sheetIndex`, `Uint8Array.from` and the `TABLE()` function. */\n"
        "  save(sheetIndex: number): Uint8Array;\n"
        "  getSheetIndex(name: string): number | null;\n"
        "}\n"
    )

    @staticmethod
    def dts_problems(text: str) -> list[str]:
        problems: list[str] = []
        readme_identifiers.check_dts_doc_tokens("fixture.d.ts", text, problems)
        return problems

    def test_resolving_tokens_are_clean(self) -> None:
        self.assertEqual(self.dts_problems(self.DTS), [])

    def test_stale_tokens_are_reported(self) -> None:
        text = self.DTS.replace(
            "`TABLE()` function.",
            "`isValid()`, `Workbook.delete`, `Range.Height`, `pivotSetColFieldOrder` and `ValueKind.Gone`.",
        )
        problems = "\n".join(self.dts_problems(text))
        stale = ("`isValid()`", "`Workbook.delete`", "`Range.Height`", "`pivotSetColFieldOrder`", "`ValueKind.Gone`")
        for token in stale:
            self.assertIn(token, problems)

    def test_listed_external_names_resolve_and_unlisted_ones_do_not(self) -> None:
        listed = self.DTS.replace("`TABLE()` function.", "`cellIs`, `phoneticPr` and `Range.Width`.")
        self.assertEqual(self.dts_problems(listed), [])
        unlisted = self.DTS.replace("`TABLE()` function.", "`cellIsNot` and `Range.Height`.")
        problems = "\n".join(self.dts_problems(unlisted))
        self.assertIn("`cellIsNot`", problems)
        self.assertIn("`Range.Height`", problems)

    def test_enum_members_resolve_against_the_exported_table(self) -> None:
        good = self.DTS.replace("`TABLE()` function.", "`ValueKind.Error`.")
        self.assertEqual(self.dts_problems(good), [])
        bad = self.DTS.replace("`TABLE()` function.", "`ValueKind.Missing`.")
        self.assertIn("`ValueKind.Missing` has no member `Missing`", "\n".join(self.dts_problems(bad)))

    def test_c_names_are_reported_unless_allowed(self) -> None:
        text = self.DTS.replace("Saves;", "Saves through fm_workbook_save, which returns kOk;")
        problems = "\n".join(self.dts_problems(text))
        self.assertIn("`fm_workbook_save`", problems)
        self.assertIn("`kOk`", problems)
        with patch.object(readme_identifiers, "C_TOKEN_DOC_ALLOWLIST", frozenset({"fm_workbook_save", "kOk"})):
            self.assertEqual(self.dts_problems(text), [])

    def test_fenced_examples_are_not_judged(self) -> None:
        text = self.DTS.replace("Saves;", "Saves.\n   * ```ts\n   * `oldName()` and `fm_x`\n   * ```\n   * ")
        self.assertEqual(self.dts_problems(text), [])

    PY = SimpleNamespace(
        classes={"Workbook": {"save", "close"}},
        module={"Workbook", "load_default"},
    )

    def py_problems(self, doc: str) -> list[str]:
        py = SimpleNamespace(classes=self.PY.classes, module=self.PY.module, docstrings=[("_workbook.py:1", doc)])
        problems: list[str] = []
        readme_identifiers.check_py_doc_tokens(py, problems)
        return problems

    def test_python_calls_resolve_to_public_names(self) -> None:
        good = "Use ``save()``, ``Workbook.close()``, ``load_default()``, ``float()`` or ``NOW()``."
        self.assertEqual(self.py_problems(good), [])
        bad = "Use ``recalculate()``, ``Workbook.delete()`` and ``save(x)``."
        problems = "\n".join(self.py_problems(bad))
        self.assertIn("``recalculate()``", problems)
        self.assertIn("``Workbook.delete()``", problems)
        self.assertNotIn("save(x)", problems)

    def test_python_docstrings_do_not_name_c_identifiers(self) -> None:
        problems = "\n".join(self.py_problems("Wraps fm_workbook_save and returns kOk."))
        self.assertIn("`fm_workbook_save`", problems)
        self.assertIn("`kOk`", problems)
        self.assertEqual(self.py_problems("Plain prose in 10 kB."), [])


class SheetSurfaceMigrationTest(unittest.TestCase):
    """The non-ledger checks read the `Sheet` class next to `Workbook`."""

    def setUp(self) -> None:
        self._tmp = TemporaryDirectory()
        self.root = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def write(self, name: str, text: str) -> Path:
        path = self.root / name
        path.write_text(text, encoding="utf-8")
        return path

    def test_readme_js_sheet_variables_are_checked_against_sheet(self) -> None:
        dts = self.write(
            "f.d.ts",
            "export interface Workbook {\n  sheet(key: string): Sheet;\n  listSheets(): readonly Sheet[];\n}\n"
            "export interface Sheet {\n  getValue(row: number, col: number): Value;\n}\n"
            "export interface WorkbookCtor {\n  createDefault(): Workbook;\n}\n"
            "export interface Formulon {\n  Workbook: WorkbookCtor;\n}\n"
            "export declare function createFormulon(): Promise<Formulon>;\n",
        )
        surface = readme_identifiers._JsSurface(dts)
        surfaces = {"@libraz/formulon": surface, "@libraz/formulon-native": surface}
        block = (
            "import createFormulon from '@libraz/formulon';\nconst f = await createFormulon();\n"
            "const wb = f.Workbook.createDefault();\nconst ws = wb.sheet('A');\nws.getValue(0, 0);\nws.setZoom(2);\n"
            "for (const s of wb.listSheets()) { s.getValue(0, 0); s.getZoom(); }\nwb.sheet('A').getFormula(0, 0);\n"
        )
        problems: list[str] = []
        readme_identifiers._check_js_block("R.md", 1, block, surfaces, problems)
        self.assertEqual(len(problems), 3, problems)
        self.assertTrue(all("is not a Sheet method" in p for p in problems), problems)

    def test_readme_python_sheet_variables_are_checked_against_sheet(self) -> None:
        py = SimpleNamespace(
            module={"Workbook"},
            classes={"Workbook": {"create_default", "sheet", "list_sheets"}, "Sheet": {"get_value", "name"}},
        )
        block = (
            "from formulon import Workbook\nwb = Workbook.create_default()\nws = wb[0]\nws.get_value(0, 0)\n"
            "ws.set_zoom(2)\nfor s in wb:\n    print(s.name, s.get_zoom())\nwb.sheet('A').get_formula(0, 0)\n"
        )
        problems: list[str] = []
        readme_identifiers._check_py_block("R.md", 1, block, py, problems)
        self.assertEqual(len(problems), 3, problems)
        self.assertTrue(all("is not a Sheet member" in p for p in problems), problems)

    def test_shared_shapes_compare_the_sheet_interface(self) -> None:
        sheet = "export interface Sheet {\n  getZoom(): number;\n}\n"
        base = "export interface Workbook {\n  save(): Uint8Array;\n}\nexport interface WorkbookCtor {\n}\n"
        wasm = self.write("w.d.ts", base + sheet)
        node = self.write("n.d.ts", base + sheet.replace("number", "string"))
        with (
            patch.object(shape_checks, "WASM_DTS", wasm),
            patch.object(shape_checks, "NODE_DTS", node),
            patch.object(shape_checks, "REPO_ROOT", self.root),
        ):
            problems = shape_checks.check_dts_shared_shapes()
        self.assertTrue(any("Sheet.getZoom" in p for p in problems), problems)

    def test_inline_decoder_may_live_on_sheet(self) -> None:
        import ast

        source = python_checks.PYTHON_SHEET.read_text(encoding="utf-8")
        tree = ast.parse(source)
        sheet = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == "Sheet")
        method = next(n for n in sheet.body if getattr(n, "name", "") == "get_pagination")
        moved = ast.get_source_segment(source, method)
        sheet_py = self.write("_sheet.py", source)
        with patch.object(python_checks, "PYTHON_SHEET", sheet_py):
            self.assertEqual(python_checks.check_python_inline_structs(), [])
            sheet_py.write_text(
                source.replace(moved, "def get_pagination_moved(self):\n        pass"), encoding="utf-8"
            )
            problems = python_checks.check_python_inline_structs()
        self.assertTrue(any("Sheet.get_pagination" in p for p in problems), problems)


class NativeDtsGeneratorTest(unittest.TestCase):
    WASM = (
        '//\n// WASM header.\n//\n\n/// <reference lib="esnext.disposable" />\n\n'
        "/**\n * @example\n * import createFormulon from '@libraz/formulon';\n */\n"
        "export const Kind: Readonly<{\n  A: 0;\n}>;\n\n"
        "export interface Workbook {\n  disposed: boolean;\n}\n\n"
        "export interface Formulon {\n"
        "  /** Factory. */\n  Workbook: WorkbookCtor;\n\n"
        "  /** Library version. */\n  getVersion(): string;\n"
        "}\n"
        "export default function createFormulon(): Promise<Formulon>;\n"
    )
    DELTA = "// delta\nexport interface Workbook {\n  /** Bytes. */\n  getMemoryUsage(): number;\n}\n"

    def setUp(self) -> None:
        self.gen = _load_gen_native_dts()
        self.out = self.gen.generate(self.WASM, self.DELTA)

    def test_wasm_header_is_replaced_and_directives_stay(self) -> None:
        self.assertNotIn("WASM header", self.out)
        self.assertIn('/// <reference lib="esnext.disposable" />', self.out)

    def test_the_usage_example_imports_the_native_package(self) -> None:
        self.assertIn("import createFormulon from '@libraz/formulon-native';", self.out)
        self.assertNotIn("from '@libraz/formulon';", self.out)

    def test_the_entry_point_is_shared_with_the_wasm_declarations(self) -> None:
        self.assertIn("export interface Formulon {", self.out)
        self.assertIn("export default function createFormulon(): Promise<Formulon>;", self.out)

    def test_delta_members_are_spliced_into_the_workbook_interface(self) -> None:
        interface = self.out[self.out.index("export interface Workbook {") :]
        interface = interface[: interface.index("\n}\n")]
        self.assertIn("disposed: boolean;", interface)
        self.assertIn("/** Bytes. */\n  getMemoryUsage(): number;", interface)

    def test_an_empty_delta_leaves_the_workbook_interface_unchanged(self) -> None:
        for delta in ("export interface Workbook {}\n", "export type Workbook = {};\n"):
            out = self.gen.generate(self.WASM, delta)
            self.assertIn("export interface Workbook {\n  disposed: boolean;\n}\n", out)

    def test_a_missing_workbook_interface_is_an_error(self) -> None:
        with self.assertRaises(self.gen.GenError):
            self.gen.generate("export interface Formulon {\n}\n", self.DELTA)


_VALUE_H = (
    "enum class ErrorCode : std::uint16_t {\n  Null = 0,\n  Div0,\n  // Excel 2007+.\n  GettingData,\n  NA,\n};\n"
    "inline constexpr ErrorCodeInfo kErrorTable[] = {\n"
    '    {"#NULL!", 0},          // ErrorCode::Null\n'
    '    {"#DIV/0!", 7},         // ErrorCode::Div0\n'
    '    {"#GETTING_DATA", 43},  // ErrorCode::GettingData\n'
    '    {"#N/A", 42},           // ErrorCode::NA\n'
    "};\n"
)
_ERROR_DTS = "export const ErrorCode: Readonly<{\n  Null: 0;\n  Div0: 1;\n  GettingData: 2;\n  NA: 3;\n}>;\n"
_ERROR_TABLE_JS = (
    "export const ErrorDisplayName = Object.freeze({\n  [ErrorCode.Null]: '#NULL!',\n  [ErrorCode.Div0]: '#DIV/0!',\n"
    "  [ErrorCode.GettingData]: '#GETTING_DATA',\n  [ErrorCode.NA]: '#N/A',\n});\n"
)


class StringTableTest(unittest.TestCase):
    DTS = _ERROR_DTS + "export const ErrorDisplayName: Readonly<Record<ErrorCode, string>>;\n"
    COMMON = "export const ErrorCode = Object.freeze({ Null: 0, Div0: 1, GettingData: 2, NA: 3 });\n" + _ERROR_TABLE_JS

    def test_a_table_matching_the_source_is_clean(self) -> None:
        self.assertEqual(declaration_checks.check_string_tables(self.COMMON, self.DTS, _VALUE_H), [])

    def test_string_tables_are_told_apart_from_ordinal_tables(self) -> None:
        self.assertEqual(declaration_checks.string_table_names(self.COMMON), {"ErrorDisplayName"})

    def test_order_and_values_follow_the_source_table(self) -> None:
        swapped = self.COMMON.replace("'#NULL!'", "'#X!'")
        problems = declaration_checks.check_string_tables(swapped, self.DTS, _VALUE_H)
        self.assertTrue(any("ErrorDisplayName" in p and "#X!" in p for p in problems), problems)
        reordered = self.COMMON.replace("  [ErrorCode.Null]: '#NULL!',\n", "").replace(
            "  [ErrorCode.NA]: '#N/A',\n", "  [ErrorCode.NA]: '#N/A',\n  [ErrorCode.Null]: '#NULL!',\n"
        )
        self.assertTrue(declaration_checks.check_string_tables(reordered, self.DTS, _VALUE_H))

    def test_a_declaration_and_its_runtime_table_come_together(self) -> None:
        problems = declaration_checks.check_string_tables(self.COMMON.replace(_ERROR_TABLE_JS, ""), self.DTS, _VALUE_H)
        self.assertTrue(any("does not export" in p for p in problems), problems)
        problems = declaration_checks.check_string_tables(self.COMMON, _ERROR_DTS, _VALUE_H)
        self.assertTrue(any("does not declare" in p for p in problems), problems)


class ErrorDisplayNamesGeneratorTest(unittest.TestCase):
    def setUp(self) -> None:
        self.gen = _load_codegen("gen_error_display_names")
        self.rows = self.gen.parse_error_table(_VALUE_H)

    def region(self, prefix: str) -> str:
        return (
            f"head\n{prefix} <error-display-names:begin> generated by tools/codegen/gen_error_display_names.py\n"
            f"{prefix} <error-display-names:end>\ntail\n"
        )

    def test_the_table_is_read_in_enum_order(self) -> None:
        self.assertEqual(
            self.rows, [("Null", "#NULL!"), ("Div0", "#DIV/0!"), ("GettingData", "#GETTING_DATA"), ("NA", "#N/A")]
        )

    def test_a_row_naming_another_enumerator_is_an_error(self) -> None:
        with self.assertRaises(self.gen.GenError):
            self.gen.parse_error_table(_VALUE_H.replace("// ErrorCode::Div0", "// ErrorCode::Ref"))

    def test_empty_regions_are_filled(self) -> None:
        js = self.gen.fill(self.region("//"), "js", self.rows, "common.mjs")
        self.assertIn(_ERROR_TABLE_JS, js)
        self.assertTrue(
            js.startswith("head\n// <error-display-names:begin>")
            and js.endswith("// <error-display-names:end>\ntail\n")
        )
        dts = self.gen.fill(self.region("//"), "dts", self.rows, "formulon.d.ts")
        self.assertIn("export const ErrorDisplayName: Readonly<Record<ErrorCode, string>>;\n", dts)
        self.assertEqual(declaration_checks.check_string_tables(js, _ERROR_DTS + dts, _VALUE_H), [])
        py = self.gen.fill(self.region("#"), "py", self.rows, "_workbook.py")
        self.assertIn('_ERROR_DISPLAY_NAMES = {\n    "NULL": "#NULL!",\n    "DIV0": "#DIV/0!",\n', py)
        self.assertIn('    "GETTING_DATA": "#GETTING_DATA",\n    "NA": "#N/A",\n}\n', py)

    def test_filling_is_idempotent(self) -> None:
        once = self.gen.fill(self.region("//"), "js", self.rows, "common.mjs")
        self.assertEqual(self.gen.fill(once, "js", self.rows, "common.mjs"), once)

    def test_a_missing_region_is_an_error_and_is_not_created(self) -> None:
        with self.assertRaises(self.gen.GenError) as caught:
            self.gen.fill("no region here\n", "js", self.rows, "common.mjs")
        self.assertIn("common.mjs", str(caught.exception))
        with self.assertRaises(self.gen.GenError):
            self.gen.fill("// <error-display-names:begin>\n", "js", self.rows, "common.mjs")


class FacadeAddressTablesTest(unittest.TestCase):
    LEDGER = {
        "schema": 1,
        "collections": {},
        "entries": {
            "value.set": {"kind": "op", "scope": "sheet", "js": "setValue", "py": "set_value",
                          "surfaces": ["wasm", "native", "py"], "cell_args": [[0, 1]]},
            "cf.evaluate": {"kind": "op", "scope": "sheet", "js": "evaluateCf", "py": "evaluate_cf",
                            "surfaces": ["wasm", "native", "py"], "cell_args": [[0, 1], [2, 3]]},
            "merge.add": {"kind": "op", "scope": "sheet", "js": "addMerge", "py": "add_merge",
                          "surfaces": ["wasm", "native", "py"], "rect_args": [0]},
            "old.set": {"kind": "op", "scope": "sheet", "js": "setOld", "surfaces": ["wasm"],
                        "drop": {"js": "setOld", "closes": "R5.list"}, "cell_args": [[0, 1]]},
            "table.remove": {"kind": "op", "scope": "sheet", "js": "removeTable", "py": "remove_table",
                             "surfaces": ["wasm", "native", "py"], "ref_args": [0, "table"]},
            "pivot_fields.list": {"kind": "list", "scope": "sheet", "js": "listPivotFields",
                                  "py": "list_pivot_fields", "surfaces": ["wasm", "native", "py"],
                                  "ref_args": [0, "pivot"]},
        },
    }  # fmt: skip

    def setUp(self) -> None:
        self.gen = _load_codegen("gen_facade_tables")
        self.ledger = api_conventions.Ledger(self.LEDGER)

    def test_public_names_map_to_positions(self) -> None:
        scope = self.gen.facade_scope
        self.assertEqual(
            scope.cell_args(self.ledger, "js"), {"Sheet.evaluateCf": [[0, 1], [2, 3]], "Sheet.setValue": [[0, 1]]}
        )
        self.assertEqual(scope.rect_args(self.ledger, "py"), {"Sheet.add_merge": [0]})
        self.assertEqual(
            scope.ref_args(self.ledger, "py"),
            {"Sheet.list_pivot_fields": [0, "pivot"], "Sheet.remove_table": [0, "table"]},
        )

    def test_regions_carry_the_tables(self) -> None:
        js = self.gen.address_tables("js", self.ledger)
        self.assertIn(
            "const CELL_ARGS = Object.freeze({\n  'Sheet.evaluateCf': [\n    [0, 1],\n    [2, 3],\n  ],\n"
            "  'Sheet.setValue': [[0, 1]],\n});\n",
            js,
        )
        self.assertIn("const RECT_ARGS = Object.freeze({\n  'Sheet.addMerge': [0],\n});\n", js)
        py = self.gen.address_tables("py", self.ledger)
        self.assertIn(
            'CELL_ARGS = {\n    "Sheet.evaluate_cf": ((0, 1), (2, 3)),\n    "Sheet.set_value": ((0, 1),),\n}\n', py
        )
        self.assertIn('RECT_ARGS = {\n    "Sheet.add_merge": (0,),\n}\n', py)
        self.assertIn(
            "const REF_ARGS = Object.freeze({\n  'Sheet.listPivotFields': [0, 'pivot'],\n"
            "  'Sheet.removeTable': [0, 'table'],\n});\n",
            js,
        )
        self.assertNotIn("REF_ARGS", py)

    HANDLES = {
        "schema": 1,
        "collections": {
            "pivot": {"key": "index", "access": "handles", "handle_class": "Pivot", "plural": "pivots"},
            "drawing_object": {"key": "natural", "by": "object_id", "access": "handles", "handle_class": "Drawing",
                               "plural": "drawings"},
            "image": {"key": "natural", "by": "object_id", "access": "handles", "handle_class": "Picture",
                      "handle_base": "Drawing", "plural": "pictures"},
            "chart": {"key": "natural", "by": "object_id", "access": "handles", "handle_class": "Chart",
                      "handle_base": "Drawing", "plural": "charts"},
        },
        "entries": {
            "column_width.set": {"kind": "op", "scope": "sheet", "js": "setColumnWidth", "py": "set_column_width",
                                 "surfaces": ["wasm", "py"], "span_args": [[0, "col", "span"]]},
            "pivot_anchor.set": {"kind": "op", "scope": "sheet", "c": "fm_sheet_set_pivot_anchor", "js": "setPivotAnchor", "py": "set_pivot_anchor",
                                 "surfaces": ["wasm", "py"], "cell_args": [[0, 1]],
                                 "handle": {"class": "Pivot", "supplies": ["sheet", "pivot"]}},
            "pivot_name.get": {"kind": "op", "scope": "sheet", "c": "fm_sheet_get_pivot_name", "js": "getPivotName", "py": "get_pivot_name",
                               "surfaces": ["wasm", "py"],
                               "handle": {"class": "Pivot", "js": "name", "py": "name", "property": True,
                                          "supplies": ["sheet", "pivot"]}},
            "drawing_name.set": {"kind": "op", "scope": "sheet", "c": "fm_sheet_set_drawing_name", "js": "setDrawingName", "py": "set_drawing_name",
                                 "surfaces": ["wasm", "py"],
                                 "handle": {"class": ["Picture", "Chart"], "js": "rename", "py": "rename",
                                            "supplies": ["sheet", "object_id"]}},
            "binding.pivot_get_properties": {"kind": "op", "scope": "sheet", "surfaces": ["wasm", "py"],
                                             "facade": "handwritten", "exception": "x",
                                             "handle": {"class": "Pivot", "js": "getProperties",
                                                        "py": "get_properties", "supplies": ["sheet", "pivot"]}},
            "pivot_runtime_id.get": {"kind": "op", "scope": "sheet", "js": "getPivotRuntimeId",
                                     "py": "get_pivot_runtime_id", "surfaces": ["wasm", "py"], "internal": True},
        },
    }  # fmt: skip

    def test_handle_methods_are_keyed_by_class(self) -> None:
        scope = self.gen.facade_scope
        ledger = api_conventions.Ledger(self.HANDLES)
        self.assertEqual(scope.cell_args(ledger, "js"), {"Pivot.setAnchor": [[0, 1]]})
        self.assertEqual(scope.span_args(ledger, "py"), {"Sheet.set_column_width": [[0, "col", "span"]]})
        self.assertEqual(
            scope.handle_forwarders(ledger, "js"),
            {
                "Chart.rename": "setDrawingName",
                "Picture.rename": "setDrawingName",
                "Pivot.getProperties": None,
                "Pivot.setAnchor": "setPivotAnchor",
            },
        )
        self.assertEqual(scope.handle_properties(ledger, "py"), {"Pivot.name": "get_pivot_name"})
        self.assertEqual(scope.methods(ledger, "sheet", "js"), ["setColumnWidth"])
        self.assertEqual(scope.public_name(self.HANDLES["entries"]["pivot_runtime_id.get"], "js"), None)
        self.assertEqual(scope.raw_name(self.HANDLES["entries"]["pivot_runtime_id.get"], "js"), "getPivotRuntimeId")
        py = self.gen.address_tables("py", ledger, facade=False)
        self.assertIn('"Pivot.set_anchor": ((0, 1),)', py)
        self.assertNotIn("Sheet.", py)

    def test_node_scope_labels_a_shared_raw_with_the_common_base(self) -> None:
        scope = self.gen.facade_scope
        ledger = api_conventions.Ledger(self.HANDLES)
        self.assertEqual(
            scope.handle_methods(ledger),
            [
                ("getPivotName", "Pivot.name", 2),
                ("setDrawingName", "Drawing.rename", 2),
                ("setPivotAnchor", "Pivot.setAnchor", 2),
            ],
        )

    def test_node_scope_inc_declares_both_tables(self) -> None:
        text = self.gen.node_scope_inc()
        self.assertIn("constexpr const char* kSheetMethods[] = {\n", text)
        self.assertIn("struct HandleMethod {\n  const char* raw;\n  const char* label;\n  int supplied;\n};\n", text)
        self.assertIn('    {"addPivotField", "Pivot.addField", 2},\n', text)
        self.assertIn('    {"setDrawingName", "Drawing.rename", 2},\n', text)
        guards = _load_codegen("gen_wasm_positional_guards")
        self.assertIn("addPivotField", guards._sheet_methods({"addPivotField", "nope"}, text))
        self.assertEqual(guards._handle_methods({"addPivotField"}, text), {"addPivotField": ["Pivot.addField", 2]})
        self.assertNotIn("evalFormula", guards.FREE_FUNCTION_GUARDS)
        self.assertIn("evaluateFormula", guards.FREE_FUNCTION_GUARDS)

    def test_only_node_scope_writes_the_inc_alone(self) -> None:
        self.assertEqual([path for path, _ in self.gen.outputs("node-scope")], [self.gen.NODE_SCOPE_INC])

    def test_patch_keys_are_the_patch_type_and_the_getter_record(self) -> None:
        dts = (
            "export type Patch<T> = { [K in keyof T]?: T[K] | null };\n"
            "export interface Margins { left: number; right: number; readonly leftStated: boolean; }\n"
            "export type MarginsPatch = Patch<Pick<Margins, 'left' | 'right'>>;\n"
            "export interface Sheet {\n"
            "  setMargins(table: TableRef, patch: MarginsPatch | null): void;\n"
            "  setWide(patch: Patch<{ " + " ".join(f"k{i:02}: number;" for i in range(30)) + " }>): void;\n"
            "}\n"
        )
        ledger = api_conventions.Ledger(
            {
                "schema": 1,
                "collections": {},
                "entries": {
                    "margins.set": {"kind": "op", "scope": "sheet", "js": "setMargins", "surfaces": ["wasm"],
                                    "patch": True, "patch_record": "Margins"},
                    "wide.set": {"kind": "op", "scope": "sheet", "js": "setWide", "surfaces": ["wasm"],
                                 "patch": True, "patch_record": "none", "patch_reason": "x"},
                },
            }
        )  # fmt: skip
        keys = self.gen.facade_scope.patch_keys(ledger, dts)
        self.assertEqual(keys["Sheet.setMargins"], [1, ["left", "leftStated", "right"]])
        table = self.gen.patch_table(ledger, dts)
        self.assertIn("  'Sheet.setMargins': [1, ['left', 'leftStated', 'right']],\n", table)
        self.assertIn("  'Sheet.setWide': [\n    0,\n    [\n      'k00',\n", table)
        self.assertTrue(all(len(line) <= 120 for line in table.splitlines()))

    def test_patch_nested_keys_are_each_level_and_its_record(self) -> None:
        dts = (
            "export interface Font { readonly bold: boolean; readonly family: number | null; }\n"
            "export interface BorderSide { readonly style: number; readonly color: Color | null; }\n"
            "export interface Border { readonly left: BorderSide; readonly diagonalUp: boolean; }\n"
            "export interface CellFormat { readonly font: Font; readonly border: Border; }\n"
            "export interface FontPatch { readonly bold?: boolean; readonly color?: ColorInput; }\n"
            "export interface BorderSidePatch { readonly style?: number; readonly color?: ColorInput; }\n"
            "export interface BorderPatch { readonly left?: BorderSidePatch; readonly diagonalUp?: boolean; }\n"
            "export interface CellFormatPatch { readonly font?: FontPatch; readonly border?: BorderPatch; }\n"
            "export interface Sheet {\n  setCellFormat(range: CellRect | string, patch: CellFormatPatch): void;\n}\n"
        )
        nested = {"font": "Font", "border": "Border", "border.left": "BorderSide"}
        ledger = api_conventions.Ledger(
            {
                "schema": 1,
                "collections": {},
                "entries": {
                    "binding.set_cell_format": {"kind": "op", "scope": "sheet", "js": "setCellFormat",
                                                "surfaces": ["wasm"], "patch": True, "patch_record": "CellFormat",
                                                "patch_nested": nested},
                },
            }
        )  # fmt: skip
        keys = self.gen.facade_scope.patch_nested_keys(ledger, dts)
        self.assertEqual(
            keys,
            {
                "Sheet.setCellFormat": {
                    "border": ["diagonalUp", "left"],
                    "border.left": ["color", "style"],
                    "font": ["bold", "color", "family"],
                }
            },
        )
        table = self.gen.patch_nested_table(ledger, dts)
        self.assertIn(
            "  'Sheet.setCellFormat': {\n    border: ['diagonalUp', 'left'],\n", table.replace("'border'", "border")
        )

    def test_the_json_is_sorted_and_per_language(self) -> None:
        import json

        text = self.gen.address_args_json(self.ledger)
        self.assertTrue(text.endswith("\n"))
        data = json.loads(text)
        self.assertEqual(data["js"]["cell_args"], {"Sheet.evaluateCf": [[0, 1], [2, 3]], "Sheet.setValue": [[0, 1]]})
        self.assertEqual(data["py"]["rect_args"], {"Sheet.add_merge": [0]})
        self.assertEqual(data["py"]["span_args"], {})
        self.assertEqual(text, self.gen.address_args_json(self.ledger))
        self.assertEqual(list(data), ["js", "py"])


class LedgerAddressArgsTest(unittest.TestCase):
    DTS = (
        "export interface Sheet {\n"
        "  setValue(row: number, col: number, value: Value): void;\n"
        "  setValue(a1: string, value: Value): void;\n"
        "  evaluateCf(row: number, col: number, anchorRow: number, anchorCol: number, f: string): Value;\n"
        "  addMerge(range: CellRect | string): void;\n"
        "  setPrintArea(ranges: readonly CellRect[] | null): void;\n"
        "  insertRows(row: number, count: number): void;\n"
        "  setFreeze(freezeRows: number, freezeCols: number): void;\n"
        "  removeTable(table: TableRef): void;\n"
        "  listPivotFields(pivot: PivotRef): PivotFieldInfo[];\n"
        "}\n"
    )

    @staticmethod
    def entries(**fields: dict) -> dict:
        base = {
            "value.set": {"kind": "op", "js": "setValue", "cell_args": [[0, 1]]},
            "cf.evaluate": {"kind": "op", "js": "evaluateCf", "cell_args": [[0, 1], [2, 3]]},
            "merge.add": {"kind": "op", "js": "addMerge", "rect_args": [0]},
            "print_area.set": {"kind": "op", "js": "setPrintArea", "rect_args": [0]},
            "rows.insert": {"kind": "op", "js": "insertRows", "span_args": [[0, "row", "start_count"]]},
            "freeze.set": {"kind": "op", "js": "setFreeze"},
            "table.remove": {"kind": "op", "js": "removeTable", "ref_args": [0, "table"]},
            "pivot_fields.list": {"kind": "list", "js": "listPivotFields", "ref_args": [0, "pivot"]},
        }
        for key, changes in fields.items():
            base[key.replace("__", ".")] = {**base.get(key.replace("__", "."), {}), **changes}
        return base

    def test_the_declarations_agree(self) -> None:
        validate = LedgerRenameChainTest._validate_names()
        self.assertEqual(validate.check_address_args(self.entries(), self.DTS), [])

    def test_a_missing_or_stale_position_is_reported(self) -> None:
        validate = LedgerRenameChainTest._validate_names()
        entries = self.entries()
        del entries["cf.evaluate"]["cell_args"]
        entries["rows.insert"]["cell_args"] = [[0, 1]]
        entries["merge.add"]["rect_args"] = [1]
        del entries["pivot_fields.list"]["ref_args"]
        entries["table.remove"]["ref_args"] = [0, "pivot"]
        problems = validate.check_address_args(entries, self.DTS)
        self.assertEqual(len(problems), 5, problems)
        self.assertTrue(any("pivot_fields.list" in p and "[0, 'pivot']" in p for p in problems), problems)
        self.assertTrue(any("cf.evaluate" in p and "[[0, 1], [2, 3]]" in p for p in problems), problems)

    def test_a_span_shaped_op_without_span_args_is_a_leak(self) -> None:
        validate = LedgerAddressArgsTest._validate_names_module()
        dts = self.DTS.replace("}\n", "  setColumnWidth(first: number, last: number, width: number): void;\n}\n")
        entries = self.entries()
        del entries["rows.insert"]["span_args"]
        entries["column_width.set"] = {"kind": "op", "js": "setColumnWidth"}
        problems = validate.check_address_args(entries, dts)
        self.assertEqual(len(problems), 2, problems)
        self.assertTrue(any("rows.insert" in p and "'row', 'start_count'" in p for p in problems), problems)
        self.assertTrue(any("column_width.set" in p and "'col', 'span'" in p for p in problems), problems)
        entries["column_width.set"]["span_args"] = [[0, "col", "span"]]
        entries["rows.insert"]["span_args"] = [[0, "row", "start_count"]]
        self.assertEqual(validate.check_address_args(entries, dts), [])

    def test_a_handle_op_is_read_from_its_class_in_the_handle_frame(self) -> None:
        validate = LedgerAddressArgsTest._validate_names_module()
        dts = self.DTS + (
            "export interface Pivot {\n"
            "  setAnchor(row: number, col: number): void;\n"
            "  addField(spec: PivotFieldInput): number;\n"
            "}\n"
        )
        entries = self.entries()
        handle = {"class": "Pivot", "supplies": ["sheet", "pivot"]}
        entries["pivot_anchor.set"] = {"kind": "op", "js": "setPivotAnchor", "cell_args": [[0, 1]], "handle": handle}
        entries["pivot_field.add"] = {"kind": "op", "js": "addPivotField", "handle": {**handle, "js": "addField"}}
        problems = validate.check_address_args(entries, dts)
        self.assertTrue(any("pivot_anchor.set" in p and "interface Pivot" in p for p in problems), problems)
        entries["pivot_anchor.set"]["handle"] = {**handle, "js": "setAnchor"}
        self.assertEqual(validate.check_address_args(entries, dts), [])
        entries["pivot_anchor.set"]["cell_args"] = [[1, 2]]
        self.assertEqual(len(validate.check_address_args(entries, dts)), 1)

    @staticmethod
    def _validate_names_module():
        return LedgerRenameChainTest._validate_names()

    def test_schema_shapes(self) -> None:
        validate = LedgerRenameChainTest._validate_names()

        def problems(**fields: object) -> list:
            entry = {"kind": "op", "scope": "sheet", "js": "setValue", "surfaces": ["wasm"], "exception": "x", **fields}
            return validate.check_schema({"schema": 1, "collections": {}, "entries": {"value.set": entry}})

        self.assertEqual(problems(cell_args=[[0, 1]], rect_args=[2]), [])
        self.assertEqual(problems(patch=True, patch_record="ChartLegend"), [])
        self.assertEqual(problems(patch=True, patch_record="none", patch_reason="no record getter"), [])
        self.assertTrue(problems(cell_args=[[0]]))
        self.assertTrue(problems(cell_args=[[-1, 0]]))
        self.assertTrue(problems(rect_args=["a"]))
        self.assertEqual(problems(ref_args=[0, "table"]), [])
        self.assertTrue(problems(ref_args=[0, "chart"]))
        self.assertTrue(problems(ref_args=[[0, "table"]]))
        self.assertTrue(problems(patch=True))
        self.assertTrue(problems(patch=False, patch_record="X"))
        self.assertTrue(problems(patch=True, patch_record="none"))
        self.assertTrue(problems(patch=True, patch_record="X", patch_reason="y"))
        self.assertEqual(problems(span_args=[[0, "col", "span"]]), [])
        self.assertEqual(problems(span_args=[[0, "row", "start_count"]]), [])
        self.assertTrue(problems(span_args=[[0, "sheet", "span"]]))
        self.assertTrue(problems(span_args=[[0, "row", "range"]]))
        self.assertTrue(problems(span_args=[0, "row", "span"]))
        nested = {"font": "Font", "border": "Border", "border.left": "BorderSide"}
        self.assertEqual(problems(patch=True, patch_record="CellFormat", patch_nested=nested), [])
        self.assertTrue(problems(patch_nested=nested))
        self.assertTrue(problems(patch=True, patch_record="CellFormat", patch_nested={"border.left": "BorderSide"}))
        self.assertTrue(problems(patch=True, patch_record="CellFormat", patch_nested={"font": 1}))


if __name__ == "__main__":
    unittest.main()
