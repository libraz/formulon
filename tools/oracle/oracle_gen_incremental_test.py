#!/usr/bin/env python3
"""Tests for the generator's incremental capture: which expectations it keeps."""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from tools.oracle import case_schema, oracle_gen

ENV = {"excel_version": "16.113.3", "excel_locale": "en-US", "date1904": False, "iterative": False}


def _suite(cases: list) -> case_schema.Suite:
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "s.yaml"
        path.write_text(
            json.dumps({"suite": "s", "description": "d", "locale": "ja-JP", "cases": cases}),
            encoding="utf-8",
        )
        return case_schema.load_suite(path)


def _golden(cases: list, env: dict = ENV) -> dict:
    return {"suite": "s", "environment": dict(env, generated_at="2026-10-08T00:00:00Z"), "cases": cases}


class ReusableExpectationsTest(unittest.TestCase):
    def test_keeps_unchanged_cases_and_drops_new_or_edited_ones(self) -> None:
        suite = _suite(
            [
                {"id": "same", "formula": "=1+1"},
                {"id": "edited", "formula": "=2+2"},
                {"id": "new", "formula": "=3+3"},
                {"id": "setup", "setup": {"A1": 5}, "formula": "=A1"},
            ]
        )
        golden = _golden(
            [
                {"id": "same", "formula": "=1+1", "setup": {}, "expect": {"kind": "number", "value": 2.0}},
                {"id": "edited", "formula": "=2+3", "setup": {}, "expect": {"kind": "number", "value": 5.0}},
                {"id": "setup", "formula": "=A1", "setup": {"A1": 6}, "expect": {"kind": "number", "value": 6.0}},
            ]
        )
        kept = oracle_gen._reusable_expectations(suite, golden, ENV)
        self.assertEqual(kept, {"same": {"kind": "number", "value": 2.0}})

    def test_a_skipped_record_is_not_reused(self) -> None:
        suite = _suite([{"id": "a", "formula": "=1"}])
        golden = _golden([{"id": "a", "formula": "=1", "setup": {}, "skipped": "driver skip: x"}])
        self.assertEqual(oracle_gen._reusable_expectations(suite, golden, ENV), {})

    def test_a_different_capture_environment_reuses_nothing(self) -> None:
        suite = _suite([{"id": "a", "formula": "=1"}])
        record = {"id": "a", "formula": "=1", "setup": {}, "expect": {"kind": "number", "value": 1.0}}
        for key, value in (("excel_version", "16.112"), ("excel_locale", "ja-JP"), ("date1904", True)):
            with self.subTest(key=key):
                golden = _golden([record], env=dict(ENV, **{key: value}))
                self.assertEqual(oracle_gen._reusable_expectations(suite, golden, ENV), {})

    def test_shape_captures_are_always_rerun(self) -> None:
        suite = _suite([{"id": "a", "formula": "=SEQUENCE(9)", "capture": "shape", "samples": ["Z1"]}])
        golden = _golden(
            [
                {
                    "id": "a",
                    "formula": "=SEQUENCE(9)",
                    "setup": {},
                    "capture": "shape",
                    "expect": {"kind": "array_shape", "shape": [9, 1], "samples": {"Z1": 1.0}},
                }
            ]
        )
        self.assertEqual(oracle_gen._reusable_expectations(suite, golden, ENV), {})

    def test_written_golden_merges_kept_and_captured_cases_in_yaml_order(self) -> None:
        suite = _suite([{"id": "kept", "formula": "=1"}, {"id": "fresh", "formula": "=2"}])
        fresh = oracle_gen.CaseResult(id="fresh", kind="number", value=2.0)
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "s.golden.json"
            oracle_gen._write_golden(
                out,
                suite,
                dict(ENV, generated_at="now"),
                [fresh],
                reused={"kept": {"kind": "number", "value": 1.0}},
            )
            doc = json.loads(out.read_text(encoding="utf-8"))
        self.assertEqual([c["id"] for c in doc["cases"]], ["kept", "fresh"])
        self.assertEqual([c["expect"]["value"] for c in doc["cases"]], [1.0, 2.0])


if __name__ == "__main__":
    unittest.main()
