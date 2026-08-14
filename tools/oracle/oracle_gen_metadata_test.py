#!/usr/bin/env python3
"""Focused regression tests for variant divergence metadata injection."""

from __future__ import annotations

import json
import tempfile
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

    def test_formula_case_merges_round_trip_to_golden(self) -> None:
        suite = self.by_name["spill_collision"]
        source_case = next(case for case in suite.cases if case.id == "spill_blocked_by_merged_anchor")
        self.assertEqual(source_case.merges, ["Z1:AA1"])

        golden = json.loads((REPO_ROOT / "tests/oracle/golden/spill_collision.golden.json").read_text(encoding="utf-8"))
        generated = next(record for record in golden["cases"] if record["id"] == source_case.id)
        self.assertEqual(generated.get("merges"), source_case.merges)


class FormulaCellSchemaTest(unittest.TestCase):
    """`formula_cell` validation and its path into the driver / golden."""

    def _load(self, body: str) -> case_schema.Suite:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "placement.yaml"
            path.write_text(body, encoding="utf-8")
            return case_schema.load_suite(path)

    def _suite_yaml(self, formula_cell: str) -> str:
        return f'suite: placement\ncases:\n  - id: c\n    formula: "=A:A"\n    formula_cell: {formula_cell}\n'

    def test_accepted_addresses_are_upper_cased(self) -> None:
        for raw, expected in (("aa5", "AA5"), ("Z1", "Z1"), ("XFD1048576", "XFD1048576")):
            with self.subTest(raw=raw):
                suite = self._load(self._suite_yaml(raw))
                self.assertEqual(suite.cases[0].formula_cell, expected)

    def test_rejected_addresses_name_the_case(self) -> None:
        for raw in ("'$Z$1'", "'Sheet2!Z1'", "'Z1:AA5'", "'Z'", "'1'", "XFE1", "Z1048577", '""'):
            with self.subTest(raw=raw), self.assertRaisesRegex(ValueError, "case 'c'"):
                self._load(self._suite_yaml(raw))

    def test_default_placement_is_absent_from_case_input_and_golden(self) -> None:
        case = case_schema.Case(id="c", formula="=1")
        self.assertNotIn("formula_cell", oracle_gen._case_input(case))

    def test_declared_placement_reaches_the_driver_input(self) -> None:
        case = case_schema.Case(id="c", formula="=A:A", formula_cell="AA5")
        self.assertEqual(oracle_gen._case_input(case)["formula_cell"], "AA5")

    def test_placement_round_trips_to_the_golden(self) -> None:
        suite = next(
            suite
            for _, suite in case_schema.discover_suites(REPO_ROOT / "tests/oracle/cases")
            if suite.name == "whole_axis_spill"
        )
        declared = {case.id: case.formula_cell for case in suite.cases}
        self.assertTrue(declared)
        self.assertTrue(all(cell for cell in declared.values()))

        golden = json.loads(
            (REPO_ROOT / "tests/oracle/golden/whole_axis_spill.golden.json").read_text(encoding="utf-8")
        )
        self.assertEqual({record["id"]: record.get("formula_cell") for record in golden["cases"]}, declared)

    def test_shape_capture_requires_samples_and_rejects_stray_ones(self) -> None:
        with self.assertRaisesRegex(ValueError, "requires a non-empty samples list"):
            self._load('suite: p\ncases:\n  - id: c\n    formula: "=A:A"\n    capture: shape\n')
        with self.assertRaisesRegex(ValueError, "samples requires capture: shape"):
            self._load('suite: p\ncases:\n  - id: c\n    formula: "=A:A"\n    samples: [Z1]\n')
        with self.assertRaisesRegex(ValueError, "unknown capture mode"):
            self._load('suite: p\ncases:\n  - id: c\n    formula: "=A:A"\n    capture: everything\n')
        with self.assertRaisesRegex(ValueError, "duplicate sample address"):
            self._load('suite: p\ncases:\n  - id: c\n    formula: "=A:A"\n    capture: shape\n    samples: [Z1, z1]\n')
        with self.assertRaisesRegex(ValueError, "at most"):
            addresses = ", ".join(f"Z{index + 1}" for index in range(case_schema.MAX_SHAPE_SAMPLES + 1))
            self._load(
                f'suite: p\ncases:\n  - id: c\n    formula: "=A:A"\n    capture: shape\n    samples: [{addresses}]\n'
            )

    def test_shape_capture_reaches_the_driver_input(self) -> None:
        suite = self._load(
            'suite: p\ncases:\n  - id: c\n    formula: "=A:A"\n    capture: shape\n    samples: [z1, aa5]\n'
        )
        case = suite.cases[0]
        self.assertEqual((case.capture, case.samples), ("shape", ["Z1", "AA5"]))
        driver_input = oracle_gen._case_input(case)
        self.assertEqual(driver_input["capture"], "shape")
        self.assertEqual(driver_input["samples"], ["Z1", "AA5"])

    def test_shape_captured_goldens_carry_shape_and_samples(self) -> None:
        golden = json.loads(
            (REPO_ROOT / "tests/oracle/golden/whole_axis_spill.golden.json").read_text(encoding="utf-8")
        )
        shaped = [record for record in golden["cases"] if record.get("capture") == "shape"]
        self.assertTrue(shaped)
        for record in shaped:
            expect = record.get("expect")
            self.assertIsNotNone(expect, record["id"])
            self.assertEqual(expect["kind"], "array_shape", record["id"])
            self.assertEqual(len(expect["shape"]), 2, record["id"])
            self.assertTrue(expect["samples"], record["id"])

    def test_suites_without_a_placement_keep_the_field_out_of_the_golden(self) -> None:
        golden = json.loads(
            (REPO_ROOT / "tests/oracle/golden/implicit_intersection.golden.json").read_text(encoding="utf-8")
        )
        self.assertTrue(all("formula_cell" not in record for record in golden["cases"]))


if __name__ == "__main__":
    unittest.main()
