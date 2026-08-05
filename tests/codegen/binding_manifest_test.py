#!/usr/bin/env python3
"""Snapshot test for the binding-glue generator.

Invokes `tools/codegen/gen_bindings.py --check` in dry-run mode and
asserts it exits cleanly. The snapshots ARE the checked-in generated
files (`src/{c_api,wasm,node_addon}/generated/`); this test fails if
those files drift out of sync with `tools/codegen/binding_manifest.yaml`.

Additionally exercises three invariants on the manifest itself:
  - Every declared entry maps to a recognised body kind.
  - `js_method` is unique across the manifest (collisions would surface
    as embind / N-API registration errors at compile time, but failing
    here is far more helpful).
  - Each entry covers all three bindings (no partial migrations).

Run directly:
    python3 -m unittest tests.codegen.binding_manifest_test
"""

from __future__ import annotations

import re
import subprocess
import sys
import unittest
from pathlib import Path

try:
    import yaml
except ImportError:  # pragma: no cover - exercised with `python -S` below.
    yaml = None

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
GEN_SCRIPT = REPO_ROOT / "tools" / "codegen" / "gen_bindings.py"
MANIFEST = REPO_ROOT / "tools" / "codegen" / "binding_manifest.yaml"
C_HEADER = REPO_ROOT / "src" / "c_api" / "formulon_c.h"

# Mirrors gen_bindings.py::_C_HEADER_DECL_RE.
_C_HEADER_DECL_RE = re.compile(r"\bFM_API\b[^;{]*?\b(fm_[A-Za-z0-9_]+)\s*\(")

# Body kinds the generator understands. Kept in sync with
# `tools/codegen/gen_bindings.py::VALID_BODY_KINDS`.
VALID_BODY_KINDS = {
    "direct_size_t",
    "status_uint32_out",
    "status_sizet_out_with_sheet",
}

# Required fields per entry; mirrors gen_bindings.py::REQUIRED_FIELDS.
REQUIRED_FIELDS = (
    "name",
    "area",
    "body",
    "accessor",
    "js_method",
    "embind_cpp",
    "node_cpp",
)


class BindingManifestTest(unittest.TestCase):
    """Verifies the generator + manifest are consistent and complete."""

    @classmethod
    def setUpClass(cls) -> None:
        if yaml is None:
            raise unittest.SkipTest(
                "PyYAML is not available; skipping binding codegen drift check. "
                "Install with `pip install pyyaml` to enable."
            )
        cls.entries = cls._load_manifest_entries()

    @staticmethod
    def _load_manifest_entries() -> list[dict]:
        raw = yaml.safe_load(MANIFEST.read_text())
        if not isinstance(raw, dict):
            raise AssertionError("manifest top level must be a mapping")
        funcs = raw.get("functions") or []
        if not isinstance(funcs, list):
            raise AssertionError("manifest `functions` must be a list")
        return funcs

    # ---- Schema validation -----------------------------------------------

    def test_manifest_has_required_fields(self) -> None:
        for i, e in enumerate(self.entries):
            for field in REQUIRED_FIELDS:
                with self.subTest(entry=i, field=field):
                    self.assertIn(field, e, f"entry {i} missing field `{field}`")

    def test_body_kinds_are_recognised(self) -> None:
        for e in self.entries:
            with self.subTest(name=e.get("name")):
                self.assertIn(e["body"], VALID_BODY_KINDS)

    def test_js_methods_are_unique(self) -> None:
        seen: dict[str, str] = {}
        for e in self.entries:
            js = e["js_method"]
            if js in seen:
                self.fail(f"duplicate js_method `{js}` for {seen[js]} vs {e['name']}")
            seen[js] = e["name"]

    def test_c_abi_names_are_unique(self) -> None:
        seen = set()
        for e in self.entries:
            name = e["name"]
            self.assertNotIn(name, seen, f"duplicate C ABI name `{name}`")
            seen.add(name)

    def test_minimum_population(self) -> None:
        # Phase D acceptance criteria: the manifest must carry enough
        # entries to justify the codegen infrastructure.
        self.assertGreaterEqual(
            len(self.entries),
            6,
            f"manifest carries only {len(self.entries)} entries; "
            "below the 6-entry threshold the codegen infra is not worth it.",
        )

    # ---- Header coverage -------------------------------------------------

    def test_every_manifest_entry_is_declared_in_header(self) -> None:
        # The codegen emits only function bodies; the public `fm_*`
        # declarations are hand-written in `formulon_c.h`. A manifest entry
        # with no matching declaration there would build neither binding,
        # yet the generated-vs-checked-in snapshot check cannot see it. Pin
        # the coverage so a manifest/header divergence fails loudly.
        declared = set(_C_HEADER_DECL_RE.findall(C_HEADER.read_text()))
        for e in self.entries:
            with self.subTest(name=e["name"]):
                self.assertIn(
                    e["name"],
                    declared,
                    f"manifest entry `{e['name']}` has no FM_API declaration in {C_HEADER}",
                )

    def test_header_regex_matches_a_known_symbol(self) -> None:
        # Guard against the regex silently matching nothing (which would
        # make the coverage test vacuously pass): the header must expose at
        # least the entries the manifest references.
        declared = set(_C_HEADER_DECL_RE.findall(C_HEADER.read_text()))
        self.assertIn("fm_workbook_sheet_count", declared)

    # ---- Snapshot (generator vs checked-in output) -----------------------

    def test_generator_runs_clean(self) -> None:
        # `--check` exits 0 if the on-disk generated files match the
        # output the manifest would produce; non-zero means drift.
        result = subprocess.run(
            [sys.executable, str(GEN_SCRIPT), "--check"],
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            check=False,
        )
        if result.returncode != 0:
            self.fail(
                "generator drift detected. Run `python3 tools/codegen/gen_bindings.py` "
                "and commit the result.\n"
                f"stderr:\n{result.stderr}\nstdout:\n{result.stdout}"
            )

    def test_generated_files_exist(self) -> None:
        # Sanity: each manifest area must have a corresponding generated
        # file under each of the three binding trees.
        areas = sorted({e["area"] for e in self.entries})
        for area in areas:
            with self.subTest(area=area):
                self.assertTrue(
                    (REPO_ROOT / "src" / "c_api" / "generated" / f"{area}.cpp").exists(),
                    f"missing c_api/generated/{area}.cpp",
                )
                self.assertTrue(
                    (REPO_ROOT / "src" / "wasm" / "generated" / f"{area}.cpp").exists(),
                    f"missing wasm/generated/{area}.cpp",
                )
                self.assertTrue(
                    (REPO_ROOT / "src" / "node_addon" / "generated" / f"{area}.cc").exists(),
                    f"missing node_addon/generated/{area}.cc",
                )


if __name__ == "__main__":
    unittest.main()
