#!/usr/bin/env python3
"""Regression gate for the Formulon nanobench microbenchmark suite.

Runs the bench driver (one combined shell script that invokes every
`bench_*` executable in turn, writing per-bench nanobench JSON output
under ``${CMAKE_BINARY_DIR}/bench_results/``), parses the resulting
JSON files, and compares the median wall-clock time of every named
result against ``tools/bench/baseline.json``. Any benchmark whose
median exceeds the baseline by more than the per-bench tolerance fails
the script with a non-zero exit code.

Default tolerance is 20% (``--threshold 0.20``). Tighten when CI
runners stabilise. The first run on a new machine should regenerate
the baseline rather than fail the gate; pass ``--regenerate-baseline``
to write the just-measured numbers back to the baseline file. Note
that this is a destructive write — review the diff before committing.

Usage (typical invocation from ctest):

    check_regression.py
        --runner   build/tests/bench/run_bench_all.sh
        --results  build/bench_results/
        --baseline tools/bench/baseline.json

Exit codes:
    0   all benches within tolerance
    1   one or more benches regressed
    2   missing baseline / driver / results
    3   argument or parse error
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from typing import Any

# nanobench JSON shape (v4.3.x):
#
# {
#   "results": [
#     {
#       "title": "...",
#       "name":  "scenario name",
#       "unit":  "eval",
#       "median(elapsed)": 1.23e-6,
#       ...
#     },
#     ...
#   ]
# }
#
# We index baseline + actual by `(file_stem, scenario_name)` so a future
# bench that ships multiple scenarios under the same JSON file does not
# collide.

# Default per-benchmark regression tolerance. 20% accounts for noise on
# developer laptops and shared CI runners. Tighten once CI stabilises.
DEFAULT_THRESHOLD = 0.20


def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--runner", required=True, help="Path to the generated run_bench_all.sh script.")
    p.add_argument("--results", required=True, help="Directory that the runner writes JSON output to.")
    p.add_argument("--baseline", required=True, help="Path to the committed baseline JSON.")
    p.add_argument(
        "--threshold",
        type=float,
        default=DEFAULT_THRESHOLD,
        help=f"Regression tolerance as a fraction (default {DEFAULT_THRESHOLD:.2f}).",
    )
    p.add_argument(
        "--regenerate-baseline",
        action="store_true",
        help="After running, overwrite the baseline JSON with fresh measurements.",
    )
    p.add_argument(
        "--skip-runner",
        action="store_true",
        help="Skip invoking the runner; use the existing JSON files in --results.",
    )
    return p.parse_args(argv)


def run_driver(runner: Path, results: Path) -> None:
    """Invokes the bench driver script. The driver is the configure-time
    generated `run_bench_all.sh` that streams every bench's stdout/stderr
    to the parent terminal — we want that output visible from ctest so
    the user sees per-bench wall-clock timings even when the gate passes.
    """
    if not runner.exists():
        print(f"check_regression: runner not found: {runner}", file=sys.stderr)
        sys.exit(2)
    print(f"check_regression: invoking {runner}", flush=True)
    proc = subprocess.run([str(runner)], cwd=runner.parent, check=False)
    if proc.returncode != 0:
        print(f"check_regression: runner exited with code {proc.returncode}", file=sys.stderr)
        sys.exit(proc.returncode)
    if not results.exists():
        print(f"check_regression: results directory not produced: {results}", file=sys.stderr)
        sys.exit(2)


def load_results(results_dir: Path) -> dict[str, dict[str, float]]:
    """Loads every `*.json` file under `results_dir` and indexes the
    median elapsed times by `(stem, scenario_name)`. Returns the dict.
    """
    out: dict[str, dict[str, float]] = {}
    for path in sorted(results_dir.glob("*.json")):
        try:
            doc = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            print(f"check_regression: failed to parse {path}: {exc}", file=sys.stderr)
            sys.exit(3)
        scenarios: dict[str, float] = {}
        for entry in doc.get("results", []):
            name = entry.get("name") or entry.get("title")
            if not name:
                continue
            # nanobench reports both the median and a full sample list;
            # we use the median because it is most stable across runs.
            median = entry.get("median(elapsed)")
            if median is None:
                # Fallback: derive from `totalTime` / `numIterations` if
                # the json template did not emit median(elapsed).
                total = entry.get("totalTime")
                iters = entry.get("numIterations") or 1
                if total is not None and iters:
                    median = float(total) / float(iters)
            if median is None:
                continue
            scenarios[name] = float(median)
        out[path.stem] = scenarios
    return out


def load_baseline(path: Path) -> dict[str, dict[str, float]]:
    if not path.exists():
        print(
            "check_regression: baseline not found at "
            f"{path}. Run with --regenerate-baseline to create it.",
            file=sys.stderr,
        )
        sys.exit(2)
    try:
        # `baseline.json` may carry a leading `_comment` field describing
        # how to regenerate it; strip non-bench keys.
        doc = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"check_regression: failed to parse baseline {path}: {exc}", file=sys.stderr)
        sys.exit(3)
    if not isinstance(doc, dict):
        print(f"check_regression: baseline {path} is not a JSON object", file=sys.stderr)
        sys.exit(3)
    out: dict[str, dict[str, float]] = {}
    for stem, scenarios in doc.items():
        if stem.startswith("_"):
            continue
        if not isinstance(scenarios, dict):
            continue
        out[stem] = {}
        for name, entry in scenarios.items():
            if isinstance(entry, dict):
                if "median" in entry:
                    out[stem][name] = float(entry["median"])
            elif isinstance(entry, (int, float)):
                out[stem][name] = float(entry)
    return out


def write_baseline(path: Path, results: dict[str, dict[str, float]]) -> None:
    payload: dict[str, Any] = {
        "_comment": (
            "Formulon bench baseline. Regenerate via "
            "`./build/tests/bench/run_bench_all.sh && "
            "tools/bench/check_regression.py --runner build/tests/bench/run_bench_all.sh "
            "--results build/bench_results --baseline tools/bench/baseline.json "
            "--regenerate-baseline --skip-runner`. Per-bench scenario keys are nanobench "
            "`name` fields; values are median wall-clock seconds. Threshold for the gate "
            "is 20% wall-clock regression."
        ),
    }
    for stem, scenarios in sorted(results.items()):
        payload[stem] = {name: {"median": median} for name, median in sorted(scenarios.items())}
    path.write_text(json.dumps(payload, indent=2, sort_keys=False) + "\n", encoding="utf-8")
    print(f"check_regression: wrote baseline to {path}", flush=True)


def diff(actual: dict[str, dict[str, float]], baseline: dict[str, dict[str, float]],
         threshold: float) -> list[str]:
    """Returns a list of human-readable regression descriptions. Empty
    list means every measured scenario is within tolerance.
    """
    regressions: list[str] = []
    for stem, scenarios in sorted(actual.items()):
        base_scenarios = baseline.get(stem)
        if base_scenarios is None:
            print(f"check_regression: NEW bench file (no baseline): {stem}", flush=True)
            continue
        for name, median in sorted(scenarios.items()):
            base = base_scenarios.get(name)
            if base is None:
                print(f"check_regression: NEW scenario (no baseline): {stem} :: {name}", flush=True)
                continue
            if base <= 0.0:
                continue
            ratio = median / base
            delta_pct = (ratio - 1.0) * 100.0
            status = "OK"
            if ratio > 1.0 + threshold:
                status = "REGRESSED"
                regressions.append(
                    f"{stem} :: {name} -> "
                    f"baseline {base:.6e}s, actual {median:.6e}s, "
                    f"+{delta_pct:.1f}% (threshold +{threshold * 100.0:.1f}%)"
                )
            print(
                f"check_regression: {status:9s} {stem} :: {name}  "
                f"baseline={base:.4e}s actual={median:.4e}s delta={delta_pct:+.1f}%",
                flush=True,
            )
    return regressions


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    runner = Path(args.runner).resolve()
    results_dir = Path(args.results).resolve()
    baseline_path = Path(args.baseline).resolve()

    if not args.skip_runner:
        run_driver(runner, results_dir)
    if not results_dir.exists():
        print(f"check_regression: results directory missing: {results_dir}", file=sys.stderr)
        return 2

    actual = load_results(results_dir)
    if not actual:
        print(f"check_regression: no JSON files found in {results_dir}", file=sys.stderr)
        return 2

    if args.regenerate_baseline:
        write_baseline(baseline_path, actual)
        return 0

    baseline = load_baseline(baseline_path)
    regressions = diff(actual, baseline, args.threshold)
    if regressions:
        print("check_regression: REGRESSIONS DETECTED", file=sys.stderr)
        for line in regressions:
            print(f"  {line}", file=sys.stderr)
        return 1
    print("check_regression: all benches within tolerance", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
