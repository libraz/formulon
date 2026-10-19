"""Regression tests for the nanobench completeness gate."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[2]
CHECKER = ROOT / "tools" / "bench" / "check_regression.py"


def bench_doc(*scenarios: tuple[str, float, float | None]) -> dict[str, Any]:
    results: list[dict[str, Any]] = []
    for name, median, error in scenarios:
        entry: dict[str, Any] = {"name": name, "median(elapsed)": median}
        if error is not None:
            entry["medianAbsolutePercentError(elapsed)"] = error
        results.append(entry)
    return {"results": results}


class CheckRegressionTest(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.results = self.root / "results"
        self.results.mkdir()
        self.runner = self.root / "runner.py"
        self.baseline = self.root / "baseline.json"

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def write_baseline(self, data: dict[str, Any]) -> bytes:
        encoded = json.dumps(data, indent=2).encode("utf-8") + b"\n"
        self.baseline.write_bytes(encoded)
        return encoded

    def write_result(self, stem: str, data: dict[str, Any]) -> bytes:
        encoded = json.dumps(data, indent=2).encode("utf-8") + b"\n"
        (self.results / f"{stem}.json").write_bytes(encoded)
        return encoded

    def write_runner(self, passes: list[dict[str, dict[str, Any]]]) -> None:
        """Create a small real runner whose output can vary between invocations."""
        serialized_passes = json.dumps(passes)
        source = f"""#!{sys.executable}
import json
from pathlib import Path

results_dir = Path({str(self.results)!r})
counter_path = Path({str(self.root / "runner-count.txt")!r})
passes = json.loads({serialized_passes!r})
run_index = int(counter_path.read_text()) if counter_path.exists() else 0
counter_path.write_text(str(run_index + 1))
for stem, document in passes[min(run_index, len(passes) - 1)].items():
    (results_dir / (stem + '.json')).write_text(json.dumps(document), encoding='utf-8')
"""
        self.runner.write_text(source, encoding="utf-8")
        self.runner.chmod(0o755)

    def run_checker(self, *options: str) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [
                sys.executable,
                str(CHECKER),
                "--runner",
                str(self.runner),
                "--results",
                str(self.results),
                "--baseline",
                str(self.baseline),
                *options,
            ],
            cwd=ROOT,
            capture_output=True,
            text=True,
            check=False,
        )

    def test_missing_baseline_bench_file_is_incomplete(self) -> None:
        self.write_runner([{"bench_present": bench_doc(("works", 1.0, None))}])
        self.write_baseline({"bench_present": {"works": 1.0}, "bench_missing": {"works": 1.0}})

        result = self.run_checker("--runs", "1")

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("missing", result.stderr.lower())

    def test_missing_baseline_scenario_is_incomplete(self) -> None:
        self.write_runner([{"bench_a": bench_doc(("present", 1.0, None))}])
        self.write_baseline({"bench_a": {"present": 1.0, "missing": 1.0}})

        result = self.run_checker("--runs", "1")

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("missing", result.stderr.lower())

    def test_empty_scenario_file_is_incomplete(self) -> None:
        self.write_result("bench_empty", {"results": []})
        self.write_baseline({"bench_empty": {"expected": 1.0}})

        result = self.run_checker("--skip-runner")

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("empty", result.stderr.lower())

    def test_real_run_removes_stale_bench_json_and_preserves_unrelated_json(self) -> None:
        self.write_result("bench_stale", bench_doc(("old", 0.1, None)))
        unrelated = self.results / "mac_probes.json"
        unrelated_bytes = b'{"probe": true}\n'
        unrelated.write_bytes(unrelated_bytes)
        self.write_runner([{"bench_current": bench_doc(("current", 1.0, None))}])
        self.write_baseline({"bench_current": {"current": 1.0}, "bench_stale": {"old": 0.1}})

        result = self.run_checker("--runs", "1")

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("missing", result.stderr.lower())
        self.assertFalse((self.results / "bench_stale.json").exists())
        self.assertEqual(unrelated.read_bytes(), unrelated_bytes)

    def test_alternating_partial_scenarios_cannot_be_union_reduced(self) -> None:
        self.write_runner(
            [
                {"bench_a": bench_doc(("first", 1.0, None))},
                {"bench_a": bench_doc(("second", 1.0, None))},
            ]
        )
        self.write_baseline({"bench_a": {"first": 1.0, "second": 1.0}})

        result = self.run_checker("--runs", "2")

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("coverage", result.stderr.lower())

    def test_added_bench_file_between_runs_is_incomplete(self) -> None:
        self.write_runner(
            [
                {"bench_a": bench_doc(("scenario", 1.0, None))},
                {
                    "bench_a": bench_doc(("scenario", 1.0, None)),
                    "bench_b": bench_doc(("scenario", 1.0, None)),
                },
            ]
        )
        self.write_baseline({"bench_a": {"scenario": 1.0}, "bench_b": {"scenario": 1.0}})

        result = self.run_checker("--runs", "2")

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("coverage", result.stderr.lower())

    def test_regenerate_baseline_leaves_file_unchanged_on_incomplete_runs(self) -> None:
        original = self.write_baseline({"bench_a": {"scenario": 99.0}})
        self.write_runner(
            [
                {"bench_a": bench_doc(("first", 1.0, None))},
                {"bench_a": bench_doc(("second", 1.0, None))},
            ]
        )

        result = self.run_checker("--runs", "2", "--regenerate-baseline")

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertEqual(self.baseline.read_bytes(), original)

    def test_minimum_of_consistent_runs_keeps_its_error_sample(self) -> None:
        self.write_runner(
            [
                {"bench_a": bench_doc(("scenario", 1.0, 0.4))},
                {"bench_a": bench_doc(("scenario", 1.2, 0.0))},
            ]
        )
        self.write_baseline({"bench_a": {"scenario": 0.9}})

        result = self.run_checker(
            "--runs",
            "2",
            "--threshold",
            "0.05",
            "--err-mult",
            "1",
            "--small-sec",
            "0",
        )

        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("actual=1.0000e+00s", result.stdout)
        self.assertIn("err=40.0%", result.stdout)

    def test_new_actual_cases_are_reported_but_allowed(self) -> None:
        self.write_runner(
            [
                {
                    "bench_a": bench_doc(("known", 1.0, None), ("new_scenario", 1.0, None)),
                    "bench_new": bench_doc(("new_case", 1.0, None)),
                }
            ]
        )
        self.write_baseline({"bench_a": {"known": 1.0}})

        result = self.run_checker("--runs", "1")

        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("NEW scenario", result.stdout)
        self.assertIn("NEW bench file", result.stdout)

    def test_skip_runner_preserves_existing_result_files(self) -> None:
        bench_bytes = self.write_result("bench_a", bench_doc(("scenario", 1.0, None)))
        unrelated = self.results / "mac_probes.json"
        unrelated_bytes = b'{"probe": true}\n'
        unrelated.write_bytes(unrelated_bytes)
        self.write_baseline({"bench_a": {"scenario": 1.0}})

        result = self.run_checker("--skip-runner")

        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertEqual((self.results / "bench_a.json").read_bytes(), bench_bytes)
        self.assertEqual(unrelated.read_bytes(), unrelated_bytes)

    def test_regression_returns_one(self) -> None:
        self.write_result("bench_a", bench_doc(("scenario", 2.0, 0.0)))
        self.write_baseline({"bench_a": {"scenario": 1.0}})

        result = self.run_checker(
            "--skip-runner",
            "--threshold",
            "0",
            "--err-mult",
            "0",
            "--small-sec",
            "0",
            "--small-widen",
            "0",
        )

        self.assertEqual(result.returncode, 1, result.stdout + result.stderr)

    def test_malformed_result_json_returns_three(self) -> None:
        (self.results / "bench_a.json").write_text("{", encoding="utf-8")
        self.write_baseline({"bench_a": {"scenario": 1.0}})

        result = self.run_checker("--skip-runner")

        self.assertEqual(result.returncode, 3, result.stdout + result.stderr)

    def test_runner_failure_exit_code_is_propagated(self) -> None:
        self.runner.write_text(f"#!{sys.executable}\nraise SystemExit(7)\n", encoding="utf-8")
        self.runner.chmod(0o755)

        result = self.run_checker("--runs", "1")

        self.assertEqual(result.returncode, 7, result.stdout + result.stderr)

    def test_baseline_in_results_namespace_is_rejected_before_runner(self) -> None:
        self.baseline = self.results / "bench_baseline.json"
        original = self.write_baseline({"bench_a": {"scenario": 1.0}})
        self.write_runner([{"bench_a": bench_doc(("scenario", 1.0, None))}])

        result = self.run_checker("--runs", "1")

        self.assertEqual(result.returncode, 3, result.stdout + result.stderr)
        self.assertIn("baseline", result.stderr.lower())
        self.assertEqual(self.baseline.read_bytes(), original)
        self.assertFalse((self.root / "runner-count.txt").exists())

    def test_baseline_symlink_in_results_is_rejected_in_skip_runner_mode(self) -> None:
        target = self.root / "baseline-target.json"
        target.write_bytes(json.dumps({"bench_a": {"scenario": 1.0}}).encode("utf-8"))
        alias = self.results / "bench_baseline_alias.json"
        try:
            alias.symlink_to(target)
        except (NotImplementedError, OSError) as exc:
            self.skipTest(f"symlinks are unavailable: {exc}")
        self.baseline = alias

        result = self.run_checker("--skip-runner")

        self.assertEqual(result.returncode, 3, result.stdout + result.stderr)
        self.assertTrue(alias.is_symlink())
        self.assertEqual(target.read_bytes(), json.dumps({"bench_a": {"scenario": 1.0}}).encode("utf-8"))

    def test_external_baseline_alias_to_results_is_rejected_before_regeneration(self) -> None:
        target = self.results / "bench_baseline.json"
        self.baseline = target
        original = self.write_baseline({"bench_a": {"scenario": 1.0}})
        alias = self.root / "baseline-alias.json"
        try:
            alias.symlink_to(target)
        except (NotImplementedError, OSError) as exc:
            self.skipTest(f"symlinks are unavailable: {exc}")
        self.baseline = alias
        self.write_runner([{"bench_a": bench_doc(("scenario", 1.0, None))}])

        result = self.run_checker("--runs", "1", "--regenerate-baseline")

        self.assertEqual(result.returncode, 3, result.stdout + result.stderr)
        self.assertEqual(target.read_bytes(), original)
        self.assertFalse((self.root / "runner-count.txt").exists())


if __name__ == "__main__":
    unittest.main()
