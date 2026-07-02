#!/usr/bin/env python3
"""Regression gate for the Formulon nanobench microbenchmark suite.

Runs the bench driver (one combined shell script that invokes every
`bench_*` executable in turn, writing per-bench nanobench JSON output
under ``${CMAKE_BINARY_DIR}/bench_results/``), parses the resulting
JSON files, and compares each named scenario's wall-clock time against
``tools/bench/baseline.json``. A scenario fails the gate when its
measured time exceeds the baseline by more than a per-scenario
tolerance.

Noise robustness
----------------
Wall-clock microbenchmarking on a shared developer laptop is noisy:
thermal throttling, OS scheduling, and neighbouring processes all add
latency to whatever epoch happens to run while they are active. That
noise is *one-sided* — external interference only ever makes a run
slower, never faster — so a single unlucky invocation can push a
scenario tens or hundreds of percent over a fixed threshold even when
nothing in the code changed. nanobench's own reported measurement
error (``medianAbsolutePercentError``) does not help here: it captures
*within-run* epoch-to-epoch spread, not the *between-run* thermal /
scheduling shifts that dominate the flakiness (and it is exactly 0 for
the single-epoch slow benches, which are the most vulnerable).

This gate therefore defends against noise on three axes:

1. **Minimum of N runs** (``--runs``, default 3). The driver is invoked
   N times and, per scenario, the *minimum* median across runs is kept
   before comparison. Because external noise is one-sided, the minimum
   converges toward the true achievable time as N grows, while a genuine
   regression — which is slower on *every* run — survives the minimum
   untouched. This is the load-bearing mechanism.
2. **Measurement-error widening**. The effective tolerance for a scenario
   is widened by ``--err-mult`` times nanobench's own reported relative
   error, so scenarios nanobench itself flags as internally noisy get
   proportionally more slack.
3. **Small-duration floor**. Sub-``--small-sec`` scenarios (micro-ops
   where OS timer resolution and scheduling quanta dominate the relative
   noise floor regardless of iteration count) get an extra ``--small-widen``
   of tolerance.

Effective per-scenario tolerance = ``threshold``
    ``+ err-mult * reported_error``
    ``+ (small-widen if median < small-sec else 0)``.

Default base tolerance is 20% (``--threshold 0.20``). The first run on a
new machine should regenerate the baseline rather than fail the gate;
pass ``--regenerate-baseline`` to write the just-measured numbers (the
same min-of-N reduction used for comparison) back to the baseline file.
Note that this is a destructive write — review the diff before
committing, ideally captured on an otherwise-idle machine.

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
from typing import Any, NamedTuple

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

# Default base per-scenario regression tolerance. 20% accounts for the
# residual noise that survives the min-of-N reduction on a busy laptop.
DEFAULT_THRESHOLD = 0.20

# Number of times the driver is invoked; the per-scenario minimum median
# across these runs is what gets compared. 3 is a good cost/robustness
# balance: a full driver pass is ~7s wall, so 3 passes keep the gate
# comfortably under its ctest timeout while giving the minimum enough
# samples to skip past a single unlucky, load-contended run.
DEFAULT_RUNS = 3

# Effective tolerance is widened by this multiple of nanobench's own
# reported relative measurement error (medianAbsolutePercentError,
# expressed as a fraction). A scenario nanobench flags as 5% noisy thus
# earns +15% of extra slack at the default multiple.
DEFAULT_ERR_MULT = 3.0

# Scenarios whose (min-of-N) median is below this many seconds are
# treated as micro-ops: at these absolute durations the OS timer
# resolution and scheduling quantum dominate the relative-noise floor
# regardless of nanobench's iteration count, so they earn a flat extra
# slice of tolerance on top of the error-based widening.
DEFAULT_SMALL_SEC = 1e-5
DEFAULT_SMALL_WIDEN = 0.10


class Sample(NamedTuple):
    """One scenario's reduced measurement.

    ``median`` is the wall-clock median in seconds; ``err`` is
    nanobench's ``medianAbsolutePercentError(elapsed)`` for the run that
    produced this median, expressed as a fraction (0.01 == 1%). ``err``
    is 0.0 for single-epoch benches, which report no internal spread.
    """

    median: float
    err: float


def parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--runner", required=True, help="Path to the generated run_bench_all.sh script.")
    p.add_argument("--results", required=True, help="Directory that the runner writes JSON output to.")
    p.add_argument("--baseline", required=True, help="Path to the committed baseline JSON.")
    p.add_argument(
        "--threshold",
        type=float,
        default=DEFAULT_THRESHOLD,
        help=f"Base regression tolerance as a fraction (default {DEFAULT_THRESHOLD:.2f}).",
    )
    p.add_argument(
        "--runs",
        type=int,
        default=DEFAULT_RUNS,
        help=(
            "Number of driver invocations; the per-scenario minimum median across "
            f"runs is compared (default {DEFAULT_RUNS}). --runs 1 restores the old "
            "single-shot behaviour. Ignored with --skip-runner (the existing JSON "
            "files are treated as a single run)."
        ),
    )
    p.add_argument(
        "--err-mult",
        type=float,
        default=DEFAULT_ERR_MULT,
        help=(
            "Widen a scenario's tolerance by this multiple of nanobench's reported "
            f"relative measurement error (default {DEFAULT_ERR_MULT:.1f})."
        ),
    )
    p.add_argument(
        "--small-sec",
        type=float,
        default=DEFAULT_SMALL_SEC,
        help=(
            "Scenarios with a median below this many seconds get --small-widen extra "
            f"tolerance (default {DEFAULT_SMALL_SEC:g})."
        ),
    )
    p.add_argument(
        "--small-widen",
        type=float,
        default=DEFAULT_SMALL_WIDEN,
        help=(
            "Extra tolerance fraction added to sub-'--small-sec' micro-ops "
            f"(default {DEFAULT_SMALL_WIDEN:.2f})."
        ),
    )
    p.add_argument(
        "--regenerate-baseline",
        action="store_true",
        help="After running, overwrite the baseline JSON with the min-of-N measurements.",
    )
    p.add_argument(
        "--skip-runner",
        action="store_true",
        help="Skip invoking the runner; use the existing JSON files in --results as one run.",
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


def load_results(results_dir: Path) -> dict[str, dict[str, Sample]]:
    """Loads every `bench_*.json` file under `results_dir` and indexes
    the median elapsed time (plus nanobench's reported relative error)
    by `(stem, scenario_name)`. Returns the dict.

    The glob is restricted to the `bench_` prefix because the bench
    results directory is a shared scratch location (oracle harness
    drops `mac_probes.json` / `probes_full.json` here too); a broader
    `*.json` glob would crash when those non-nanobench documents are
    parsed.
    """
    out: dict[str, dict[str, Sample]] = {}
    for path in sorted(results_dir.glob("bench_*.json")):
        try:
            doc = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            print(f"check_regression: failed to parse {path}: {exc}", file=sys.stderr)
            sys.exit(3)
        scenarios: dict[str, Sample] = {}
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
            # medianAbsolutePercentError(elapsed) is nanobench's within-run
            # relative spread across epochs, already a fraction (0.01 == 1%).
            # It is 0 for single-epoch benches, which is fine — those lean on
            # the min-of-N reduction instead.
            err = entry.get("medianAbsolutePercentError(elapsed)")
            scenarios[name] = Sample(float(median), float(err) if err is not None else 0.0)
        out[path.stem] = scenarios
    return out


def reduce_runs(runs: list[dict[str, dict[str, Sample]]]) -> dict[str, dict[str, Sample]]:
    """Collapses N per-run result maps into one by taking, per scenario,
    the sample with the minimum median. External noise only ever adds
    latency, so the minimum across runs is the best estimate of the true
    achievable time; the error field carried through is the one reported
    by whichever run produced that minimum.
    """
    if not runs:
        return {}
    reduced: dict[str, dict[str, Sample]] = {}
    for run in runs:
        for stem, scenarios in run.items():
            dst = reduced.setdefault(stem, {})
            for name, sample in scenarios.items():
                best = dst.get(name)
                if best is None or sample.median < best.median:
                    dst[name] = sample
    return reduced


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


def write_baseline(path: Path, results: dict[str, dict[str, Sample]]) -> None:
    payload: dict[str, Any] = {
        "_comment": (
            "Formulon bench baseline. Regenerate via "
            "`tools/bench/check_regression.py --runner build/tests/bench/run_bench_all.sh "
            "--results build/bench_results --baseline tools/bench/baseline.json "
            "--regenerate-baseline`. Per-bench scenario keys are nanobench `name` fields; "
            "values are the minimum median wall-clock seconds across the --runs driver "
            "invocations. The gate compares the same min-of-N reduction against these and "
            "fails on a >20% regression (widened per scenario by nanobench's reported error "
            "and by a micro-op floor)."
        ),
    }
    for stem, scenarios in sorted(results.items()):
        payload[stem] = {name: {"median": sample.median} for name, sample in sorted(scenarios.items())}
    path.write_text(json.dumps(payload, indent=2, sort_keys=False) + "\n", encoding="utf-8")
    print(f"check_regression: wrote baseline to {path}", flush=True)


def effective_threshold(sample: Sample, base_threshold: float, err_mult: float,
                        small_sec: float, small_widen: float) -> float:
    """Per-scenario tolerance: the base threshold widened by nanobench's
    reported within-run error and, for micro-ops, by a fixed floor.
    """
    tol = base_threshold + err_mult * sample.err
    if sample.median < small_sec:
        tol += small_widen
    return tol


def diff(actual: dict[str, dict[str, Sample]], baseline: dict[str, dict[str, float]],
         base_threshold: float, err_mult: float, small_sec: float,
         small_widen: float) -> list[str]:
    """Returns a list of human-readable regression descriptions. Empty
    list means every measured scenario is within its effective tolerance.
    """
    regressions: list[str] = []
    for stem, scenarios in sorted(actual.items()):
        base_scenarios = baseline.get(stem)
        if base_scenarios is None:
            print(f"check_regression: NEW bench file (no baseline): {stem}", flush=True)
            continue
        for name, sample in sorted(scenarios.items()):
            base = base_scenarios.get(name)
            if base is None:
                print(f"check_regression: NEW scenario (no baseline): {stem} :: {name}", flush=True)
                continue
            if base <= 0.0:
                continue
            median = sample.median
            ratio = median / base
            delta_pct = (ratio - 1.0) * 100.0
            tol = effective_threshold(sample, base_threshold, err_mult, small_sec, small_widen)
            status = "OK"
            if ratio > 1.0 + tol:
                status = "REGRESSED"
                regressions.append(
                    f"{stem} :: {name} -> "
                    f"baseline {base:.6e}s, actual {median:.6e}s, "
                    f"+{delta_pct:.1f}% (tolerance +{tol * 100.0:.1f}%)"
                )
            print(
                f"check_regression: {status:9s} {stem} :: {name}  "
                f"baseline={base:.4e}s actual={median:.4e}s delta={delta_pct:+.1f}% "
                f"tol={tol * 100.0:.1f}% (err={sample.err * 100.0:.1f}%)",
                flush=True,
            )
    return regressions


def collect_runs(runner: Path, results_dir: Path, runs: int, skip_runner: bool) -> dict[str, dict[str, Sample]]:
    """Invokes the driver ``runs`` times (or reads the existing JSON once
    when ``skip_runner`` is set) and returns the per-scenario min-of-N
    reduction. Exits the process on a missing driver / results dir.
    """
    if skip_runner:
        if not results_dir.exists():
            print(f"check_regression: results directory missing: {results_dir}", file=sys.stderr)
            sys.exit(2)
        return reduce_runs([load_results(results_dir)])

    passes = max(1, runs)
    collected: list[dict[str, dict[str, Sample]]] = []
    for i in range(passes):
        print(f"check_regression: driver run {i + 1}/{passes}", flush=True)
        run_driver(runner, results_dir)
        # The driver overwrites the JSON files in place, so snapshot each
        # run's parse before the next invocation clobbers it.
        collected.append(load_results(results_dir))
    return reduce_runs(collected)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    runner = Path(args.runner).resolve()
    results_dir = Path(args.results).resolve()
    baseline_path = Path(args.baseline).resolve()

    actual = collect_runs(runner, results_dir, args.runs, args.skip_runner)
    if not actual:
        print(f"check_regression: no JSON files found in {results_dir}", file=sys.stderr)
        return 2

    if args.regenerate_baseline:
        write_baseline(baseline_path, actual)
        return 0

    baseline = load_baseline(baseline_path)
    regressions = diff(actual, baseline, args.threshold, args.err_mult, args.small_sec, args.small_widen)
    if regressions:
        print("check_regression: REGRESSIONS DETECTED", file=sys.stderr)
        for line in regressions:
            print(f"  {line}", file=sys.stderr)
        return 1
    print("check_regression: all benches within tolerance", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
