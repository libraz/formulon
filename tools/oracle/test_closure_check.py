#!/usr/bin/env python3
"""Focused tests for closure filtering, aliases, and condition contracts."""

from __future__ import annotations

import contextlib
import io
import json
import unittest
from pathlib import Path
from typing import Dict, List, Optional, Set
from unittest import mock

from tools.oracle import closure_check as closure


def _context(
    *,
    behaviors: Optional[List[Dict]] = None,
    divergence_entries: Optional[List[Dict]] = None,
    catalog: Optional[Set[str]] = None,
    implemented: Optional[Set[str]] = None,
) -> closure.CoverageContext:
    catalog = catalog if catalog is not None else {"FN", "CANON", "ALIAS"}
    implemented = implemented if implemented is not None else set(catalog)
    return closure.CoverageContext(
        catalog=catalog,
        implemented=implemented,
        yaml_hits={fn: {Path(f"{fn.lower()}.yaml")} for fn in catalog},
        ironcalc_hits={},
        native_hits={},
        behaviors=behaviors or [],
        divergence_entries=divergence_entries or [],
        drift_result=closure.ConditionResult("behavior_drift", True, "drift=0"),
    )


class ConditionWrapperTests(unittest.TestCase):
    def test_public_condition_wrappers_keep_function_call_matching(self) -> None:
        cases = [
            {"id": "direct", "formula": "=FN(1)"},
            {"id": "spaced", "formula": "=fn (2)"},
            {"id": "nested", "formula": "=SUM(FN(3))"},
            {"id": "prefix", "formula": "=AFN(4)"},
            {"id": "suffix", "formula": "=FN_EXTRA(5)"},
        ]
        self.assertEqual([c["id"] for c in closure.cases_for_function("FN", cases)], ["direct", "spaced", "nested"])

    def test_no_cases_and_no_golden_keep_their_existing_condition_results(self) -> None:
        family = closure.Family("fn", ["FN"], None, 1, "")
        ctx = _context(behaviors=[{"function": "FN", "aspect": "shape", "name": "scalar"}])

        result = closure.evaluate_function("FN", family, [], set(), ctx)

        self.assertEqual(
            [(condition.name, condition.ok) for condition in result.conditions],
            [
                ("behaviors_declared", True),
                ("cases_present", False),
                ("golden_present", False),
                ("divergence_documented", True),
                ("not_in_pilot", True),
                ("behavior_drift", True),
            ],
        )

    def test_divergence_condition_uses_only_matching_cases(self) -> None:
        cases = [{"id": "fn_case", "formula": "=FN(1)"}, {"id": "other", "formula": "=OTHER(1)"}]
        complete = [{"id": "fn_case", "reason": "known", "prefer": "formulon", "last_verified_excel_version": "16.112"}]
        incomplete = [{"id": "fn_case", "reason": "known"}]

        self.assertTrue(closure.check_condition_4_divergence("FN", cases, []).ok)
        self.assertTrue(closure.check_condition_4_divergence("FN", cases, complete).ok)
        self.assertFalse(closure.check_condition_4_divergence("FN", cases, incomplete).ok)


class PrefilteringTests(unittest.TestCase):
    def test_evaluate_function_filters_behaviors_and_cases_once(self) -> None:
        family = closure.Family("fn", ["FN"], None, 1, "")
        cases = [{"id": "fn_case", "formula": "=FN(1)"}]
        ctx = _context(behaviors=[{"function": "FN", "aspect": "shape", "name": "scalar", "probe": "FN(1)"}])

        with (
            mock.patch.object(
                closure, "behaviors_for_function", wraps=closure.behaviors_for_function
            ) as behavior_filter,
            mock.patch.object(closure, "cases_for_function", wraps=closure.cases_for_function) as case_filter,
        ):
            result = closure.evaluate_function("FN", family, cases, {"fn_case"}, ctx)

        self.assertTrue(result.closed)
        self.assertEqual(behavior_filter.call_count, 1)
        self.assertEqual(case_filter.call_count, 1)

    def test_alias_reuses_canonical_conditions_and_preserves_alias_name(self) -> None:
        family = closure.Family("aliases", ["CANON", "ALIAS"], None, 1, "", aliases={"ALIAS": "CANON"})
        cases = [{"id": "canon_case", "formula": "=CANON(1)"}]
        ctx = _context(
            behaviors=[{"function": "CANON", "aspect": "shape", "name": "scalar", "probe": "CANON(1)"}],
        )

        canonical = closure.evaluate_function("CANON", family, cases, {"canon_case"}, ctx)
        alias = closure.evaluate_function("ALIAS", family, cases, {"canon_case"}, ctx)

        self.assertTrue(canonical.closed)
        self.assertTrue(alias.closed)
        self.assertEqual(alias.alias_of, "CANON")
        self.assertEqual(alias.conditions, canonical.conditions)

    def test_json_report_keeps_alias_and_condition_fields(self) -> None:
        family = closure.Family("aliases", ["CANON", "ALIAS"], None, 1, "", aliases={"ALIAS": "CANON"})
        cases = [{"id": "canon_case", "formula": "=CANON(1)"}]
        ctx = _context(
            behaviors=[{"function": "CANON", "aspect": "shape", "name": "scalar", "probe": "CANON(1)"}],
        )
        out = io.StringIO()
        with (
            mock.patch.object(closure, "_need_yaml"),
            mock.patch.object(
                closure, "load_families", return_value=({"aliases": family}, {"CANON": family, "ALIAS": family})
            ),
            mock.patch.object(closure, "build_coverage_context", return_value=ctx),
            mock.patch.object(closure, "load_cases_for_family", return_value=cases),
            mock.patch.object(closure, "load_golden_ids_for_family", return_value={"canon_case"}),
            contextlib.redirect_stdout(out),
        ):
            status = closure.main(["family=aliases", "--json"])

        payload = json.loads(out.getvalue())
        self.assertEqual(status, 0)
        self.assertTrue(payload["drift_ok"])
        self.assertEqual([item["fn"] for item in payload["closures"]], ["CANON", "ALIAS"])
        self.assertIsNone(payload["closures"][0]["alias_of"])
        self.assertEqual(payload["closures"][1]["alias_of"], "CANON")
        self.assertEqual(payload["closures"][1]["conditions"], payload["closures"][0]["conditions"])


if __name__ == "__main__":
    unittest.main()
