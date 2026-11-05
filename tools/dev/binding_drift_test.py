#!/usr/bin/env python3
"""Regression tests for the JavaScript binding-drift seams."""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from tools.dev.binding_drift import abi_checks, api_conventions, declaration_checks, readme_checks, ts_decl


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

    def test_helper_requires_canonical_import_and_explicit_reexport(self) -> None:
        _, _, native = self._fixture(
            "export function mergeFunctionMetadata(base) { return base; }\n",
            "export function mergeFunctionMetadata(base) { return base; }\n",
        )
        self.assertIsNone(ts_decl._js_common_reexported_function_source(native, "mergeFunctionMetadata"))

        native.write_text(
            "import { mergeFunctionMetadata } from '../npm/common.mjs';\n",
            encoding="utf-8",
        )
        self.assertIsNone(ts_decl._js_common_reexported_function_source(native, "mergeFunctionMetadata"))


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
                    patch.object(declaration_checks, "_NODE_STAGED_COPIES", copies),
                    patch.object(declaration_checks, "_SKIPPED", []),
                ):
                    problems = declaration_checks.check_staged_dist()
                self.assertEqual(len(problems), 1)
                self.assertIn(missing.name, problems[0])

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


def _conventions(header: str, ledger: dict | None = None) -> api_conventions.Context:
    return api_conventions.Context(header, api_conventions.Ledger(ledger))


def _rule_hits(ctx: api_conventions.Context, rule_id: str) -> list[tuple[str, str]]:
    rules = {rule_id: api_conventions.all_rules()[rule_id]}
    return [(v.op_key, v.discriminator) for v in api_conventions.collect(ctx, rules)]


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

    def test_r8_c_judges_the_ledger_target_name(self) -> None:
        header = "FM_API const char* fm_version_string(void);\n"
        ledger = {
            "entries": {"version.get": {"kind": "free", "c": "fm_get_version", "was": {"c": "fm_version_string"}}}
        }
        self.assertEqual(_rule_hits(_conventions(header), "R8.c"), [("fm_version_string", "-")])
        self.assertEqual(_rule_hits(_conventions(header, ledger), "R8.c"), [])

    def test_r10_c_doc(self) -> None:
        decl = "FM_API fm_status_t fm_workbook_get_title(const fm_workbook_t* wb, const char** out_title);\n"
        self.assert_rule(
            "R10.c.doc",
            "/**\n * @brief Reads the title.\n */\n" + decl,
            "/**\n * @brief Reads the title.\n *\n * `*out_title` is a model-backed view.\n */\n" + decl,
            [("fm_workbook_get_title", "-")],
        )

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
            "R10.c.doc",
            "R11.axis",
            "R11.rect",
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


if __name__ == "__main__":
    unittest.main()
