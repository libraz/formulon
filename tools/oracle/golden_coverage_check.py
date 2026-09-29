#!/usr/bin/env python3
"""Fail when the primary formula-oracle goldens do not cover every case.

Also fails on the reverse gap: a golden entry (case ID, or an entire
`<suite>.golden.json` file) with no corresponding `tests/oracle/cases/`
definition is a structural orphan. It was never generated from a
maintained case, so nothing re-derives it on the next `oracle-gen` run,
and per the release gate its cases must not be counted toward the
primary-oracle pass-rate denominator (see CLAUDE.md's release gate: the
denominator is "cases the primary oracle actually produced a value
for," which presupposes a real case definition behind it).

A third check covers the opposite direction of the same problem: a golden
case whose own `skipped` field is baked in with no `tests/divergence.yaml`
entry selecting it. `oracle_gen.py` bakes a `skipped` field from two
sources -- a registry `skip-oracle` entry, or a driver exception raised
during capture (a timeout, a bridge failure) -- and only the first is
backed by a reason and a `cause` a reviewer can see. An unbacked skip is a
permanent `GTEST_SKIP` that never shows up in `divergence_check.py`'s
per-cause tally, which is the release gate's sole accounting of what is
skipped and why.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import yaml

try:  # pragma: no cover - trivial fallback
    from tools.oracle import case_schema
    from tools.oracle.divergence_check import is_pending_stamp
    from tools.oracle.oracle_gen import DEFAULT_DIVERGENCE, _load_divergence_skips
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import case_schema  # type: ignore
    from divergence_check import is_pending_stamp  # type: ignore
    from oracle_gen import DEFAULT_DIVERGENCE, _load_divergence_skips  # type: ignore

REPO_ROOT = Path(__file__).resolve().parents[2]
CASES_DIR = REPO_ROOT / "tests/oracle/cases"
GOLDEN_DIR = REPO_ROOT / "tests/oracle/golden"
# The target `tests/oracle/golden/` goldens are captured against -- see
# `tools/oracle/targets.yaml`'s global `primary:`. Hardcoded rather than
# read from targets.yaml because this script is scoped to that one
# directory specifically, the same way GOLDEN_DIR is.
PRIMARY_TARGET = "mac-365-ja_JP"


def load_case_catalog() -> tuple[set[tuple[str, str]], set[str]]:
    """Returns ((suite, case ID) pairs, suite names) declared under tests/oracle/cases/.

    A case ID is documented as unique only *within* its suite (see
    `tests/oracle/README.md`), and 32 IDs are already reused across two
    suites -- so coverage has to be tracked per (suite, id) pair. A bare
    id set would let one suite's golden stand in for another suite's case
    of the same id and hide a real coverage gap.
    """

    pairs: set[tuple[str, str]] = set()
    suites: set[str] = set()
    for path in sorted(CASES_DIR.glob("*.yaml")):
        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        suite = doc.get("suite") if isinstance(doc.get("suite"), str) else None
        if suite:
            suites.add(suite)
        for case in doc.get("cases") or []:
            if suite and isinstance(case, dict) and isinstance(case.get("id"), str):
                pairs.add((suite, case["id"]))
    return pairs, suites


def load_golden_ids() -> tuple[set[tuple[str, str]], list[str]]:
    pairs: set[tuple[str, str]] = set()
    errors: list[str] = []
    case_pairs, suite_names = load_case_catalog()
    suites = case_schema.discover_suites(CASES_DIR) if CASES_DIR.is_dir() else []
    registry_skips = _load_divergence_skips(DEFAULT_DIVERGENCE, PRIMARY_TARGET, suites=suites)
    paths = sorted(GOLDEN_DIR.glob("*.golden.json"))
    if not paths:
        return pairs, [f"no primary golden files in {GOLDEN_DIR}"]
    for path in paths:
        try:
            doc: Any = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            errors.append(f"{path}: cannot parse JSON: {exc}")
            continue
        cases = doc.get("cases") if isinstance(doc, dict) else None
        if not isinstance(cases, list) or not cases:
            errors.append(f"{path}: cases must be a non-empty list")
            continue

        # Orphan-file check: an entire golden file whose `suite` has no
        # matching tests/oracle/cases/<suite>.yaml is never regenerated
        # from a maintained case definition (this is exactly how a
        # hand-seeded bootstrap file went unnoticed -- the old check only
        # looked for missing case IDs, never for a golden file with no
        # backing suite at all).
        suite = doc.get("suite") if isinstance(doc, dict) else None
        if isinstance(suite, str) and suite and suite not in suite_names:
            errors.append(f"{path}: suite {suite!r} has no tests/oracle/cases/{suite}.yaml")

        # Version-format check: a golden not produced by a real, verified
        # Excel capture (a hand-seeded placeholder, or the non-evidence
        # bare "16.0" Office-major stamp) must not silently look like
        # primary-oracle evidence. Delegated to divergence_check rather
        # than re-stated as a local regex -- the local copy claimed to
        # mirror that allowlist while in fact accepting "16.0".
        version = doc.get("environment", {}).get("excel_version") if isinstance(doc, dict) else None
        if is_pending_stamp(version):
            errors.append(f"{path}: environment.excel_version {version!r} is not a verified Microsoft 365 build stamp")

        for index, case in enumerate(cases):
            case_id = case.get("id") if isinstance(case, dict) else None
            if not isinstance(case_id, str) or not case_id:
                errors.append(f"{path}: cases[{index}] has no non-empty id")
                continue
            if not isinstance(suite, str) or not suite:
                continue
            pairs.add((suite, case_id))
            if (suite, case_id) not in case_pairs:
                errors.append(f"{path}: case {case_id!r} has no matching tests/oracle/cases/{suite}.yaml definition")
            if isinstance(case, dict) and "skipped" in case and case_id not in registry_skips:
                errors.append(
                    f"{path}: case {case_id!r} has a baked 'skipped' field with no tests/divergence.yaml "
                    "skip-oracle entry selecting it -- add one with a `cause`, or the release gate's "
                    "per-cause tally silently omits this permanently-skipped case"
                )
    return pairs, errors


def main() -> int:
    case_pairs, _suite_names = load_case_catalog()
    golden_pairs, errors = load_golden_ids()
    missing = sorted(case_pairs - golden_pairs)
    if missing:
        rendered = ", ".join(f"{suite}:{case_id}" for suite, case_id in missing)
        errors.append(f"missing {len(missing)} case(s): {rendered}")
    for error in errors:
        print(f"FAIL {error}", file=sys.stderr)
    print(f"oracle golden coverage: {len(golden_pairs)}/{len(case_pairs)} case IDs across primary goldens")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
