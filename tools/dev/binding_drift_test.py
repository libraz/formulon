#!/usr/bin/env python3
"""Regression tests for the JavaScript binding-drift seams."""

from __future__ import annotations

import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from tools.dev.binding_drift import declaration_checks, ts_decl


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


if __name__ == "__main__":
    unittest.main()
