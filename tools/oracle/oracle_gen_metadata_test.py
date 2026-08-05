#!/usr/bin/env python3
"""Focused regression tests for variant divergence metadata injection."""

from __future__ import annotations

import json
import unittest
from pathlib import Path

from tools.oracle import case_schema, oracle_gen

REPO_ROOT = Path(__file__).resolve().parents[2]
VARIANT_DIVERGENCE = REPO_ROOT / "tests/oracle/variants/win-365-ja_JP/divergence.yaml"


class OracleGeneratorMetadataTest(unittest.TestCase):
    def setUp(self) -> None:
        source_suites = case_schema.discover_suites(REPO_ROOT / "tests/oracle/cases")
        self.suites = oracle_gen._apply_divergence_metadata(
            source_suites,
            oracle_gen._load_divergence_metadata(VARIANT_DIVERGENCE, "win-365-ja_JP"),
        )
        self.by_name = {suite.name: suite for _, suite in self.suites}

    def test_suite_compare_modes_cover_every_case(self) -> None:
        for name in ("financial", "financial_accrual"):
            self.assertTrue(self.by_name[name].cases)
            self.assertTrue(all(case.compare_mode == "numeric_text" for case in self.by_name[name].cases))
        self.assertTrue(self.by_name["complex"].cases)
        self.assertTrue(all(case.compare_mode == "complex_text" for case in self.by_name["complex"].cases))

    def test_specific_overrides_are_limited_to_documented_cases(self) -> None:
        probes = {case.id: case for case in self.by_name["ironcalc_drift_probes"].cases}
        expected = {
            "tan_neg_three_pi_over_two",
            "tan_neg_pi_over_two",
            "tan_pi_over_two",
            "tan_three_pi_over_two",
            "misc_cot_pi_minus_eps",
            "misc_csc_pi_minus_eps",
            "misc_sec_halfpi_minus_eps",
            "misc_sec_halfpi_plus_eps",
        }
        self.assertEqual({case_id for case_id, case in probes.items() if case.tolerance is not None}, expected)
        self.assertTrue(all(probes[case_id].tolerance.rel == 0.001 for case_id in expected))
        date_case = next(
            case for case in self.by_name["datevalue_timevalue"].cases if case.id == "datevalue_agrees_with_date"
        )
        self.assertEqual(date_case.compare_mode, "datevalue_roundtrip_readback")

    def test_variant_goldens_match_generated_metadata(self) -> None:
        for name in ("financial", "financial_accrual", "complex", "datevalue_timevalue", "ironcalc_drift_probes"):
            expected = {
                case.id: {
                    "tolerance": case.tolerance.to_dict() if case.tolerance is not None else None,
                    "compare_mode": case.compare_mode,
                }
                for case in self.by_name[name].cases
            }
            golden_path = REPO_ROOT / "tests/oracle/variants/win-365-ja_JP/golden" / f"{name}.golden.json"
            golden = json.loads(golden_path.read_text(encoding="utf-8"))
            actual = {
                case["id"]: {
                    "tolerance": case.get("tolerance"),
                    "compare_mode": case.get("compare_mode"),
                }
                for case in golden["cases"]
            }
            self.assertEqual(actual, expected, name)

    def test_primary_suite_skip_expands_to_every_case(self) -> None:
        source_suites = case_schema.discover_suites(REPO_ROOT / "tests/oracle/cases")
        skips = oracle_gen._load_divergence_skips(
            REPO_ROOT / "tests/divergence.yaml",
            "mac-365-ja_JP",
            suites=source_suites,
        )
        filterxml = next(suite for _, suite in source_suites if suite.name == "filterxml")
        self.assertTrue(filterxml.cases)
        self.assertTrue(all(case.id in skips for case in filterxml.cases))


if __name__ == "__main__":
    unittest.main()
