#!/usr/bin/env python3
"""Fail when the primary formula-oracle goldens do not cover every case."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import yaml

REPO_ROOT = Path(__file__).resolve().parents[2]
CASES_DIR = REPO_ROOT / "tests/oracle/cases"
GOLDEN_DIR = REPO_ROOT / "tests/oracle/golden"


def load_case_ids() -> set[str]:
    ids: set[str] = set()
    for path in sorted(CASES_DIR.glob("*.yaml")):
        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        for case in doc.get("cases") or []:
            if isinstance(case, dict) and isinstance(case.get("id"), str):
                ids.add(case["id"])
    return ids


def load_golden_ids() -> tuple[set[str], list[str]]:
    ids: set[str] = set()
    errors: list[str] = []
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
        for index, case in enumerate(cases):
            case_id = case.get("id") if isinstance(case, dict) else None
            if not isinstance(case_id, str) or not case_id:
                errors.append(f"{path}: cases[{index}] has no non-empty id")
                continue
            ids.add(case_id)
    return ids, errors


def main() -> int:
    case_ids = load_case_ids()
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
