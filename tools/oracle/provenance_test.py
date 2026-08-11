#!/usr/bin/env python3
"""Unit tests for oracle target status/provenance policy."""

from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from tools.oracle.provenance import active_variant_names, cf_active, validate_targets
from tools.oracle.workbook_oracle_gen import _display_path, _workbook_primary


class ProvenancePolicyTest(unittest.TestCase):
    def test_workbook_primary_missing_from_manifest_fails_closed(self) -> None:
        with self.assertRaisesRegex(RuntimeError, "tracks.workbook.primary"):
            _workbook_primary({})

    def test_external_staging_path_is_renderable(self) -> None:
        self.assertIn("formulon", _display_path(Path("/tmp/formulon-m365-golden/example.golden.json")))

    def test_wanted_target_is_not_active_even_with_old_goldens(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            target_dir = root / "tests/oracle/variants/win/golden"
            target_dir.mkdir(parents=True)
            (target_dir.parent / "PROVENANCE.json").write_text(
                json.dumps(
                    {"target": "win", "status": "wanted", "classification": "reference-only", "verified": False}
                ),
                encoding="utf-8",
            )
            doc = {
                "primary": "mac",
                "targets": {
                    "mac": {"status": "primary", "environment_md": "tests/oracle/ENVIRONMENT.md"},
                    "win": {
                        "status": "wanted",
                        "environment_md": "tests/oracle/variants/win/ENVIRONMENT.md",
                    },
                },
            }
            self.assertEqual(active_variant_names(doc, root), [])
            self.assertEqual(validate_targets(doc, root), [])

    def test_reference_capture_cannot_claim_verified_product(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            env = root / "tests/oracle/variants/win/ENVIRONMENT.md"
            env.parent.mkdir(parents=True)
            (env.parent / "PROVENANCE.json").write_text(
                json.dumps(
                    {"status": "scaffolded", "classification": "active", "verified": True, "active_ctest": True}
                ),
                encoding="utf-8",
            )
            doc = {
                "primary": "mac",
                "targets": {
                    "mac": {"status": "primary"},
                    "win": {"status": "wanted", "environment_md": str(env.relative_to(root))},
                },
            }
            errors = validate_targets(doc, root)
            self.assertTrue(any("reference-only" in error for error in errors))

    def test_active_scaffold_requires_verified_provenance(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            env = root / "tests/oracle/variants/win/ENVIRONMENT.md"
            env.parent.mkdir(parents=True)
            golden_dir = env.parent / "golden"
            golden_dir.mkdir()
            golden_path = golden_dir / "suite.golden.json"
            golden_path.write_text(
                json.dumps(
                    {
                        "suite": "suite",
                        "kind": "workbook",
                        "environment": {
                            "excel_version": "16.0 (Build 18025)",
                            "excel_locale": "ja-JP",
                            "capture_id": "capture-1",
                        },
                        "cases": [{"id": "case", "expect": {}, "spec": {}}],
                    }
                ),
                encoding="utf-8",
            )
            (env.parent / "PROVENANCE.json").write_text(
                json.dumps(
                    {
                        "status": "scaffolded",
                        "classification": "active",
                        "verified": True,
                        "active_ctest": True,
                        "target": "win",
                        "product": "16.0 (Build 18025)",
                        "locale": "ja-JP",
                        "m365_sentinel": "ARRAYTOTEXT(1) == text '1'",
                        "capture_id": "capture-1",
                        "all_suites_same_capture": True,
                        "required_suites": ["suite"],
                        "captured_suites": ["suite"],
                        "suite_inventory": [
                            {
                                "suite": "suite",
                                "case_count": 1,
                                "sha256": hashlib.sha256(golden_path.read_bytes()).hexdigest(),
                                "capture_id": "capture-1",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )
            doc = {
                "primary": "mac",
                "targets": {
                    "mac": {"status": "primary"},
                    "win": {
                        "status": "scaffolded",
                        "locale": "ja-JP",
                        "output_dir": "tests/oracle/variants/win/golden",
                        "environment_md": str(env.relative_to(root)),
                    },
                },
            }
            self.assertEqual(active_variant_names(doc, root), ["win"])
            self.assertEqual(validate_targets(doc, root), [])
            golden_path.unlink()
            self.assertEqual(active_variant_names(doc, root), [])

    def test_cf_primary_requires_complete_hashed_capture(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            golden_dir = root / "tests/oracle/golden_cf"
            golden_dir.mkdir(parents=True)
            golden_path = golden_dir / "cf_smoke.golden.json"
            golden_path.write_text(
                json.dumps(
                    {
                        "name": "cf_smoke",
                        "cases": [{"id": "case", "cells": [{"row": 0, "col": 0, "matches": []}]}],
                    }
                ),
                encoding="utf-8",
            )
            marker = {
                "track": "cf",
                "target": "mac",
                "status": "primary",
                "classification": "active",
                "verified": True,
                "active_ctest": True,
                "locale": "ja-JP",
                "product": "Microsoft Excel for Mac 16.111.3",
                "m365_sentinel": "ARRAYTOTEXT(1) == text '1'",
                "capture_id": "capture-1",
                "required_suites": ["cf_smoke"],
                "captured_suites": ["cf_smoke"],
                "case_count": 1,
                "all_suites_same_capture": True,
                "suite_inventory": [
                    {
                        "suite": "cf_smoke",
                        "case_count": 1,
                        "sha256": hashlib.sha256(golden_path.read_bytes()).hexdigest(),
                        "capture_id": "capture-1",
                    }
                ],
            }
            (golden_dir / "PROVENANCE.json").write_text(json.dumps(marker), encoding="utf-8")
            doc = {
                "primary": "mac",
                "tracks": {"cf": {"primary": "mac", "variants": []}},
                "targets": {
                    "mac": {
                        "status": "primary",
                        "driver": "macos_excel",
                        "runs_on": ["Darwin"],
                        "locale": "ja-JP",
                    }
                },
            }
            self.assertTrue(cf_active(doc, root))
            self.assertEqual(validate_targets(doc, root), [])
            golden_path.write_text(golden_path.read_text(encoding="utf-8") + "\n", encoding="utf-8")
            self.assertFalse(cf_active(doc, root))


if __name__ == "__main__":
    unittest.main()
