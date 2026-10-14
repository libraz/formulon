#!/usr/bin/env python3
"""Unit tests for the chart-track case / golden validator."""

from __future__ import annotations

import copy
import json
import tempfile
import unittest
from pathlib import Path

import yaml

from tools.oracle import chart_case_schema as schema
from tools.oracle.chart_oracle_gen import compare_models, package_chart_parts

_BUILT = {
    "id": "built",
    "sheets": [{"name": "Sheet1", "cells": {"A1": "Cat", "B2": 1, "C2": {"value": 0.5, "format": "0%"}}}],
    "charts": [{"id": "c1", "kind": "column", "series": [{"name": "Sheet1!$B$1", "val": "Sheet1!$B$2:$B$5"}]}],
    "ops": [{"op": "insert_rows", "sheet": "Sheet1", "at": 2}],
}
_EDIT = {
    "id": "edited",
    "fixture": "chart_column.xlsx",
    "target": {"sheet": "Sheet1", "chart": "col_clustered"},
    "edits": [{"op": "legend", "position": "t"}, {"op": "series", "series": "all", "labels": None}],
}


def _suite(*cases: dict, views: list | None = None) -> dict:
    return {"suite": "s", "views": views or ["refs"], "cases": [copy.deepcopy(c) for c in cases]}


class CaseShapeTest(unittest.TestCase):
    def test_committed_cases_validate(self) -> None:
        self.assertGreater(schema.validate_all(schema.CASES_DIR, schema.GOLDEN_DIR), 0)

    def test_built_and_edit_cases_pass(self) -> None:
        self.assertEqual(schema.validate_suite(_suite(_BUILT, _EDIT), "s"), ["built", "edited"])

    def _rejects(self, suite: dict, fragment: str) -> None:
        with self.assertRaises(schema.ValidationError) as ctx:
            schema.validate_suite(suite, "s")
        self.assertIn(fragment, str(ctx.exception))

    def test_suite_must_match_stem(self) -> None:
        suite = _suite(_BUILT)
        suite["suite"] = "other"
        self._rejects(suite, "file stem")

    def test_unknown_view(self) -> None:
        self._rejects(_suite(_BUILT, views=["pixels"]), "views")

    def test_duplicate_case_id(self) -> None:
        self._rejects(_suite(_BUILT, _BUILT), "duplicate id")

    def test_unknown_chart_kind(self) -> None:
        case = copy.deepcopy(_BUILT)
        case["charts"][0]["kind"] = "radar"
        self._rejects(_suite(case), "kind")

    def test_scatter_series_needs_y(self) -> None:
        case = copy.deepcopy(_BUILT)
        case["charts"][0]["kind"] = "scatter"
        case["charts"][0]["series"] = [{"x": "Sheet1!$A$2:$A$5", "val": "Sheet1!$B$2:$B$5"}]
        self._rejects(_suite(case), "unknown keys")

    def test_chart_needs_exactly_one_of_series_and_source(self) -> None:
        case = copy.deepcopy(_BUILT)
        case["charts"][0]["source"] = "A1:B5"
        self._rejects(_suite(case), "exactly one")

    def test_bad_cell_address(self) -> None:
        case = copy.deepcopy(_BUILT)
        case["sheets"][0]["cells"]["a1"] = 1
        self._rejects(_suite(case), "bad address")

    def test_rows_are_one_based(self) -> None:
        case = copy.deepcopy(_BUILT)
        case["ops"][0]["at"] = 0
        self._rejects(_suite(case), "1-based")

    def test_unknown_edit(self) -> None:
        case = copy.deepcopy(_EDIT)
        case["edits"].append({"op": "rotate"})
        self._rejects(_suite(case), "op")

    def test_edit_colors_are_rgb(self) -> None:
        case = copy.deepcopy(_EDIT)
        case["edits"] = [{"op": "series", "series": 0, "fill": {"color": "red"}}]
        self._rejects(_suite(case), "RRGGBB")

    def test_edit_case_needs_an_existing_fixture(self) -> None:
        case = copy.deepcopy(_EDIT)
        case["fixture"] = "no_such_file.xlsx"
        self._rejects(_suite(case), "is not under")

    def test_edit_case_cannot_build_sheets(self) -> None:
        case = copy.deepcopy(_EDIT)
        case["sheets"] = _BUILT["sheets"]
        self._rejects(_suite(case), "edit case")

    def test_ids_are_unique_across_suites(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            for stem in ("a", "b"):
                doc = _suite(_BUILT)
                doc["suite"] = stem
                Path(tmp, f"{stem}.yaml").write_text(yaml.safe_dump(doc), encoding="utf-8")
            with self.assertRaises(schema.ValidationError) as ctx:
                schema.validate_all(Path(tmp), None)
            self.assertIn("also defined", str(ctx.exception))


class GoldenTest(unittest.TestCase):
    def _golden(self, suite: dict) -> dict:
        return {
            "name": "s",
            "views": suite["views"],
            "environment": {"excel_version": "x"},
            "cases": [
                {
                    "id": c["id"],
                    "spec": json.loads(json.dumps(schema.case_spec(c))),
                    "expect": {"charts": [{"refs": {}}]},
                }
                for c in suite["cases"]
            ],
        }

    def test_matching_golden_passes(self) -> None:
        suite = _suite(_BUILT, _EDIT)
        self.assertEqual(schema.validate_golden(self._golden(suite), suite, "s"), ["built", "edited"])

    def test_spec_drift_is_stale(self) -> None:
        suite = _suite(_BUILT)
        golden = self._golden(suite)
        suite["cases"][0]["ops"][0]["at"] = 3
        with self.assertRaises(schema.ValidationError) as ctx:
            schema.validate_golden(golden, suite, "s")
        self.assertIn("capture the suite again", str(ctx.exception))

    def test_uncaptured_case(self) -> None:
        suite = _suite(_BUILT, _EDIT)
        golden = self._golden(suite)
        golden["cases"].pop()
        with self.assertRaises(schema.ValidationError) as ctx:
            schema.validate_golden(golden, suite, "s")
        self.assertIn("not captured", str(ctx.exception))

    def test_skipped_case_has_no_expect(self) -> None:
        suite = _suite(_BUILT)
        golden = self._golden(suite)
        golden["cases"][0]["skipped"] = "reason"
        with self.assertRaises(schema.ValidationError) as ctx:
            schema.validate_golden(golden, suite, "s")
        self.assertIn("no 'expect'", str(ctx.exception))
        del golden["cases"][0]["expect"]
        self.assertEqual(schema.validate_golden(golden, suite, "s"), ["built"])

    def test_golden_must_carry_every_view(self) -> None:
        suite = _suite(_BUILT, views=["refs", "cache"])
        with self.assertRaises(schema.ValidationError) as ctx:
            schema.validate_golden(self._golden(suite), suite, "s")
        self.assertIn("missing views", str(ctx.exception))


class OpenCheckLogicTest(unittest.TestCase):
    """The parts of the open check that need no Excel."""

    _CHART = {
        "sheet": "Sheet1",
        "name": "c1",
        "shape": [{"type": "bar", "series": 2}],
        "refs": {"series": [{"val": "Sheet1!$B$2:$B$5"}]},
        "format": {"legend": {"position": "t"}},
    }

    def test_unchanged_models_pass(self) -> None:
        model = {"charts": [self._CHART]}
        self.assertEqual(compare_models(model, copy.deepcopy(model), check_format=True), [])

    def test_each_check_reports_its_letter(self) -> None:
        before = {"charts": [self._CHART]}
        after = copy.deepcopy(before)
        after["charts"][0]["refs"]["series"][0]["val"] = "Sheet1!$B$3:$B$6"
        after["charts"][0]["shape"][0]["series"] = 1
        after["charts"][0]["format"]["legend"]["position"] = "b"
        problems = compare_models(before, after, check_format=True)
        self.assertEqual([p[:3] for p in problems], ["(b)", "(c)", "(e)"])
        self.assertIn(".legend.position", problems[2])
        self.assertEqual(len(compare_models(before, after, check_format=False)), 2)

    def test_lost_chart(self) -> None:
        self.assertEqual(
            compare_models({"charts": [self._CHART]}, {"charts": []}, check_format=True), ["chart count 1 -> 0"]
        )

    def test_package_walk_counts_fixture_charts(self) -> None:
        fixtures = schema.FIXTURES_DIR
        self.assertEqual(package_chart_parts(fixtures / "chart_column.xlsx"), (4, []))
        self.assertEqual(package_chart_parts(fixtures / "chart_chartsheet.xlsx")[1], [])

    def test_package_walk_reports_missing_part(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            import zipfile

            src = schema.FIXTURES_DIR / "chart_column.xlsx"
            dst = Path(tmp, "broken.xlsx")
            with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w") as zo:
                for info in zi.infolist():
                    if info.filename != "xl/charts/chart2.xml":
                        zo.writestr(info, zi.read(info.filename))
            self.assertEqual(package_chart_parts(dst), (4, ["xl/charts/chart2.xml"]))


if __name__ == "__main__":
    unittest.main()
