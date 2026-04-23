#!/usr/bin/env python3
"""Generates golden JSON from YAML cases by driving Mac Excel via xlwings.

Usage:
    python3 tools/oracle/oracle_gen.py [--suite NAME ...] [--cases-dir P]
                                       [--golden-dir P]

Default paths match the repository layout:

    cases:  tests/oracle/cases/*.yaml
    golden: tests/oracle/golden/*.golden.json

Each YAML under `cases/` is evaluated independently; the corresponding
`<suite>.golden.json` is rewritten from scratch. Failures in one suite do
not abort the rest — unless `--strict` is passed, in which case the first
error is fatal.

macOS + Excel 365 only. The generator calls `driver.ExcelOracle`, which
refuses to start on any other platform.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import sys
from pathlib import Path
from typing import Dict, List, Optional

# Local imports — accept both `python3 tools/oracle/oracle_gen.py` (no
# package) and `python3 -m tools.oracle.oracle_gen` (package-style).
try:  # pragma: no cover - trivial fallback
    from tools.oracle import case_schema
    from tools.oracle.driver import ExcelOracle, CaseResult, EnvironmentInfo
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import case_schema  # type: ignore
    from driver import ExcelOracle, CaseResult, EnvironmentInfo  # type: ignore


REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_CASES_DIR = REPO_ROOT / "tests" / "oracle" / "cases"
DEFAULT_GOLDEN_DIR = REPO_ROOT / "tests" / "oracle" / "golden"
DEFAULT_ENV_FILE = REPO_ROOT / "tests" / "oracle" / "ENVIRONMENT.md"
DEFAULT_DIVERGENCE = REPO_ROOT / "tests" / "divergence.yaml"


def _result_to_json(result: CaseResult) -> Dict[str, object]:
    """Shapes a `CaseResult` into the on-wire `expect` JSON object."""

    if result.kind == "blank":
        return {"kind": "blank"}
    if result.kind == "number":
        return {"kind": "number", "value": result.value}
    if result.kind == "bool":
        return {"kind": "bool", "value": result.value}
    if result.kind == "text":
        return {"kind": "text", "value": result.value}
    if result.kind == "error":
        return {"kind": "error", "code": result.error_code or "#UNKNOWN!"}
    if result.kind == "array":
        # Arrays are emitted but the C++ verifier doesn't yet understand
        # them — they currently fail with "unknown expect kind: array",
        # which is the desired outcome until spill tests are wired.
        return {
            "kind": "array",
            "value": result.value,
            "shape": result.array_shape,
        }
    return {"kind": result.kind}


def _env_to_json(env: EnvironmentInfo, iso_now: str) -> Dict[str, object]:
    return {
        "excel_version": env.excel_version,
        "excel_locale": env.excel_locale,
        "date1904": env.date1904,
        "iterative": env.iterative,
        "generated_at": iso_now,
    }


def _case_input(case: case_schema.Case) -> Dict[str, object]:
    return {
        "id": case.id,
        "formula": case.formula,
        "setup": case.setup,
    }


def _write_golden(
    out_path: Path,
    suite: case_schema.Suite,
    env_json: Dict[str, object],
    results: List[CaseResult],
    skipped: Optional[Dict[str, str]] = None,
) -> None:
    skipped = skipped or {}
    cases_out: List[Dict[str, object]] = []
    by_id = {r.id: r for r in results}
    for c in suite.cases:
        record: Dict[str, object] = {
            "id": c.id,
            "formula": c.formula,
            "setup": c.setup,
        }
        if c.id in skipped:
            record["skipped"] = skipped[c.id]
            cases_out.append(record)
            continue
        r = by_id.get(c.id)
        if r is None:
            record["skipped"] = "no result captured"
            cases_out.append(record)
            continue
        record["expect"] = _result_to_json(r)
        if c.tolerance is not None:
            record["tolerance"] = c.tolerance.to_dict()
        cases_out.append(record)

    doc = {
        "suite": suite.name,
        "description": suite.description,
        "environment": env_json,
        "tolerance": suite.tolerance.to_dict(),
        "cases": cases_out,
    }
    out_path.parent.mkdir(parents=True, exist_ok=True)
    # Python's json.dumps is deterministic under sort_keys=False; we keep
    # the declared field order so goldens diff cleanly on regeneration.
    out_path.write_text(
        json.dumps(doc, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def _load_divergence_skips(path: Path) -> Dict[str, str]:
    """Loads case-id -> reason from `tests/divergence.yaml` entries whose
    `mode` is `skip-oracle`."""

    if not path.exists():
        return {}
    try:
        import yaml  # type: ignore

        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except Exception:
        return {}
    out: Dict[str, str] = {}
    entries = doc.get("entries") or []
    if not isinstance(entries, list):
        return {}
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        cid = entry.get("id")
        mode = entry.get("mode", "tolerance")
        if mode == "skip-oracle" and isinstance(cid, str):
            out[cid] = str(entry.get("reason", "divergence.yaml skip-oracle"))
    return out


def _write_environment_md(path: Path, env: EnvironmentInfo, iso_now: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    body = (
        "# Oracle Environment\n\n"
        "This file records the Excel version and locale last used to\n"
        "regenerate `tests/oracle/golden/`. Reviewers should watch this\n"
        "file on oracle-gen PRs to catch version-driven divergences early.\n\n"
        f"- **Excel version**: `{env.excel_version}`\n"
        f"- **Excel locale**: `{env.excel_locale}`\n"
        f"- **date1904**: `{env.date1904}`\n"
        f"- **iterative**: `{env.iterative}`\n"
        f"- **generated_at**: `{iso_now}`\n"
    )
    path.write_text(body, encoding="utf-8")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--suite",
        action="append",
        default=None,
        metavar="NAME",
        help="Run only the named suite(s); defaults to all YAML files.",
    )
    parser.add_argument(
        "--cases-dir", type=Path, default=DEFAULT_CASES_DIR,
        help="Directory of *.yaml case files.",
    )
    parser.add_argument(
        "--golden-dir", type=Path, default=DEFAULT_GOLDEN_DIR,
        help="Directory to write *.golden.json files to.",
    )
    parser.add_argument(
        "--divergence", type=Path, default=DEFAULT_DIVERGENCE,
        help="YAML listing cases to skip / widen; see tests/divergence.yaml.",
    )
    parser.add_argument(
        "--strict", action="store_true",
        help="Abort on first suite failure instead of continuing.",
    )
    parser.add_argument(
        "--visible", action="store_true",
        help="Show the Excel window during generation (debug aid).",
    )
    args = parser.parse_args(argv)

    suites = case_schema.discover_suites(args.cases_dir)
    if args.suite:
        wanted = set(args.suite)
        suites = [(p, s) for (p, s) in suites if s.name in wanted]
    if not suites:
        print(f"oracle-gen: no YAML suites found in {args.cases_dir}")
        return 0

    skips = _load_divergence_skips(args.divergence)

    iso_now = _dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

    with ExcelOracle(visible=args.visible) as oracle:
        env = oracle.probe_environment()
        env_json = _env_to_json(env, iso_now)

        exit_code = 0
        for path, suite in suites:
            print(f"[oracle-gen] {suite.name}  ({len(suite.cases)} cases)")
            try:
                # Carry divergence.yaml skips into the per-run map. The
                # driver still gets the full case list so row numbering
                # stays aligned; we'll drop the skipped ids on write.
                runnable = [
                    c for c in suite.cases if c.id not in skips
                ]
                case_inputs = [_case_input(c) for c in runnable]
                env_copy = EnvironmentInfo(
                    excel_version=env.excel_version,
                    excel_locale=env.excel_locale,
                    date1904=suite.options.get("date1904", False),
                    iterative=suite.options.get("iterative", False),
                )
                this_env_json = _env_to_json(env_copy, iso_now)
                results = oracle.run_suite(
                    suite.name,
                    case_inputs,
                    date1904=env_copy.date1904,
                    iterative=env_copy.iterative,
                )
                out_path = args.golden_dir / f"{suite.name}.golden.json"
                _write_golden(
                    out_path,
                    suite,
                    this_env_json,
                    results,
                    skipped={c.id: skips[c.id] for c in suite.cases if c.id in skips},
                )
                print(f"  -> {out_path.relative_to(REPO_ROOT)}")
            except Exception as exc:
                exit_code = 1
                print(f"  ! failed: {exc}")
                if args.strict:
                    return exit_code

        _write_environment_md(DEFAULT_ENV_FILE, env, iso_now)
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
