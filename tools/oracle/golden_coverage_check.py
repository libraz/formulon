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
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import yaml

try:  # pragma: no cover - trivial fallback
    from tools.oracle.divergence_check import is_pending_stamp
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from divergence_check import is_pending_stamp  # type: ignore

REPO_ROOT = Path(__file__).resolve().parents[2]
CASES_DIR = REPO_ROOT / "tests/oracle/cases"
GOLDEN_DIR = REPO_ROOT / "tests/oracle/golden"


def load_case_catalog() -> tuple[set[str], set[str]]:
    """Returns (case IDs, suite names) declared under tests/oracle/cases/."""

    ids: set[str] = set()
    suites: set[str] = set()
    for path in sorted(CASES_DIR.glob("*.yaml")):
        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        if isinstance(doc.get("suite"), str):
            suites.add(doc["suite"])
        for case in doc.get("cases") or []:
            if isinstance(case, dict) and isinstance(case.get("id"), str):
                ids.add(case["id"])
    return ids, suites


def load_golden_ids() -> tuple[set[str], list[str]]:
    ids: set[str] = set()
    errors: list[str] = []
    case_ids, suite_names = load_case_catalog()
    paths = sorted(GOLDEN_DIR.glob("*.golden.json"))
    if not paths:
        return ids, [f"no primary golden files in {GOLDEN_DIR}"]
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
            ids.add(case_id)
            if case_id not in case_ids:
                errors.append(f"{path}: case {case_id!r} has no matching tests/oracle/cases/ definition")
    return ids, errors


def main() -> int:
    case_ids, _suite_names = load_case_catalog()
    golden_ids, errors = load_golden_ids()
    missing = sorted(case_ids - golden_ids)
    if missing:
        errors.append(f"missing {len(missing)} case(s): {', '.join(missing)}")
    for error in errors:
        print(f"FAIL {error}", file=sys.stderr)
    print(f"oracle golden coverage: {len(golden_ids)}/{len(case_ids)} case IDs across primary goldens")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
