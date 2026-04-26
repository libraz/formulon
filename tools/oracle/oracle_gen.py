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
    from tools.oracle.drivers import select_driver
    from tools.oracle.drivers.base import CaseResult, EnvironmentInfo
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import case_schema  # type: ignore
    from drivers import select_driver  # type: ignore
    from drivers.base import CaseResult, EnvironmentInfo  # type: ignore


REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_CASES_DIR = REPO_ROOT / "tests" / "oracle" / "cases"
DEFAULT_GOLDEN_DIR = REPO_ROOT / "tests" / "oracle" / "golden"
DEFAULT_ENV_FILE = REPO_ROOT / "tests" / "oracle" / "ENVIRONMENT.md"
DEFAULT_DIVERGENCE = REPO_ROOT / "tests" / "divergence.yaml"
DEFAULT_TARGETS_FILE = Path(__file__).resolve().parent / "targets.yaml"


def _load_targets(path: Path) -> Dict[str, object]:
    """Loads `targets.yaml`. Raises RuntimeError on any read / parse error.

    The file is required for `--target` resolution; oracle_gen will not
    silently fall back to hard-coded paths so that a typo in `--target`
    surfaces immediately rather than being papered over.
    """

    if not path.exists():
        raise RuntimeError(f"oracle targets file not found: {path}")
    try:
        import yaml  # type: ignore

        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except Exception as exc:
        raise RuntimeError(f"failed to parse {path}: {exc}") from exc
    if not isinstance(doc, dict):
        raise RuntimeError(f"{path} root must be a mapping")
    targets = doc.get("targets")
    if not isinstance(targets, dict) or not targets:
        raise RuntimeError(f"{path} has no `targets:` mapping")
    return doc


def _resolve_target(
    targets_doc: Dict[str, object], name: Optional[str]
) -> Dict[str, object]:
    """Returns the target record for `name` (or the primary if None)."""

    targets = targets_doc.get("targets") or {}
    if name is None:
        name = targets_doc.get("primary")  # type: ignore[assignment]
        if not isinstance(name, str):
            raise RuntimeError("targets.yaml is missing a `primary:` entry")
    if name not in targets:  # type: ignore[operator]
        avail = ", ".join(sorted(targets.keys()))  # type: ignore[union-attr]
        raise RuntimeError(f"unknown oracle target: {name!r} (available: {avail})")
    record = targets[name]  # type: ignore[index]
    if not isinstance(record, dict):
        raise RuntimeError(f"target {name!r} must be a mapping")
    record = dict(record)
    record["_name"] = name
    return record


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
        if c.compare_mode is not None and c.compare_mode != "exact":
            record["compare_mode"] = c.compare_mode
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


def _load_divergence_skips(path: Path, target_name: str) -> Dict[str, str]:
    """Loads case-id -> reason from divergence YAML entries whose mode is
    `skip-oracle` AND whose `applies_to` either includes `target_name` or
    is absent (default = applies to all targets).

    File-level read / parse failures fall back to an empty dict so a
    completely missing or unreadable divergence file does not abort
    generation. Entry-level type errors (e.g. `applies_to` is not a list)
    are surfaced as `RuntimeError` so typos are caught by the caller
    rather than silently masking entries.
    """

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
        if mode != "skip-oracle" or not isinstance(cid, str):
            continue
        if "applies_to" in entry and entry["applies_to"] is not None:
            applies = entry["applies_to"]
            if not isinstance(applies, list) or not all(
                isinstance(x, str) for x in applies
            ):
                raise RuntimeError(
                    f"{path}: entry {cid!r} has invalid `applies_to`: "
                    f"expected list of strings, got {applies!r}"
                )
            if target_name not in applies:
                continue
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
        "--target",
        default=None,
        metavar="NAME",
        help=(
            "Oracle target from tools/oracle/targets.yaml; defaults to the "
            "`primary:` entry (mac-365-ja_JP). The target supplies "
            "`output_dir` (-> --golden-dir) and `environment_md` unless "
            "those are explicitly overridden on the command line."
        ),
    )
    parser.add_argument(
        "--targets-file",
        type=Path,
        default=DEFAULT_TARGETS_FILE,
        help="Path to targets.yaml (rarely needs overriding).",
    )
    parser.add_argument(
        "--cases-dir", type=Path, default=DEFAULT_CASES_DIR,
        help="Directory of *.yaml case files.",
    )
    parser.add_argument(
        "--golden-dir", type=Path, default=None,
        help=(
            "Directory to write *.golden.json files to. Overrides the "
            "selected target's `output_dir`."
        ),
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

    # Resolve target metadata. Errors here are fatal — we'd rather refuse
    # to start than write goldens to a stale path on a typo.
    try:
        targets_doc = _load_targets(args.targets_file)
        target = _resolve_target(targets_doc, args.target)
    except RuntimeError as exc:
        print(f"oracle-gen: {exc}", file=sys.stderr)
        return 2

    # Per-target paths, but always honour explicit CLI overrides.
    target_output = target.get("output_dir")
    if args.golden_dir is not None:
        golden_dir = args.golden_dir
    elif isinstance(target_output, str) and target_output:
        golden_dir = REPO_ROOT / target_output
    else:
        golden_dir = DEFAULT_GOLDEN_DIR

    target_env_md = target.get("environment_md")
    if isinstance(target_env_md, str) and target_env_md:
        env_md_path = REPO_ROOT / target_env_md
    else:
        env_md_path = DEFAULT_ENV_FILE

    suites = case_schema.discover_suites(args.cases_dir)
    if args.suite:
        wanted = set(args.suite)
        suites = [(p, s) for (p, s) in suites if s.name in wanted]
    if not suites:
        print(f"oracle-gen: no YAML suites found in {args.cases_dir}")
        return 0

    try:
        skips = _load_divergence_skips(args.divergence, target["_name"])

        # Variants may declare an extra `divergence:` path in targets.yaml,
        # interpreted relative to the repo root. Entries there override the
        # primary file on key collision because they're more specific.
        target_div = target.get("divergence")
        if isinstance(target_div, str) and target_div:
            variant_div_path = REPO_ROOT / target_div
            if variant_div_path.exists():
                variant_skips = _load_divergence_skips(
                    variant_div_path, target["_name"]
                )
                skips.update(variant_skips)
    except RuntimeError as exc:
        print(f"oracle-gen: {exc}", file=sys.stderr)
        return 2

    iso_now = _dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

    # Driver factory errors (wrong host OS, missing config) are fatal --
    # we'd rather refuse to start than dump a confusing traceback halfway
    # through. The factory's RuntimeError already carries an actionable
    # message; just forward it as a regular CLI error.
    try:
        oracle_cm = select_driver(target, visible=args.visible)
    except RuntimeError as exc:
        print(f"oracle-gen: {exc}", file=sys.stderr)
        return 2

    with oracle_cm as oracle:
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
                out_path = golden_dir / f"{suite.name}.golden.json"
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

        _write_environment_md(env_md_path, env, iso_now)
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
