#!/usr/bin/env python3
"""Validate the provenance and review state of ``tests/divergence.yaml``.

Each divergence record must be backed by an oracle case, an explicit alias
to a current oracle case, or a non-oracle scope with repository-local
evidence.  The latter is for intentionally non-oracle surfaces such as C
API shape choices and binary-fixture limitations; it prevents a free-form
``id`` from silently becoming a permanent skip.

By default structural problems fail and entries awaiting a fresh Excel
probe are reported as warnings.  ``--strict`` also fails for those pending
records, making it suitable for a reprobe-completion gate.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any

import yaml

REPO_ROOT = Path(__file__).resolve().parents[2]
FORMULA_CASES_DIR = REPO_ROOT / "tests" / "oracle" / "cases"
WORKBOOK_CASES_DIR = REPO_ROOT / "tests" / "oracle" / "cases_wb"
DEFAULT_DIVERGENCE = REPO_ROOT / "tests" / "divergence.yaml"
NON_ORACLE_SCOPES = {
    "api-contract",
    "environment",
    "fixture",
    "deferred-feature",
    "unit-test",
}


def load_case_catalog() -> tuple[set[str], set[str]]:
    """Load the raw case IDs and formula-suite names used by both tracks."""

    ids: set[str] = set()
    suites: set[str] = set()
    for path in sorted(FORMULA_CASES_DIR.glob("*.yaml")):
        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        if isinstance(doc.get("suite"), str):
            suites.add(doc["suite"])
        for case in doc.get("cases") or []:
            if isinstance(case, dict) and isinstance(case.get("id"), str):
                ids.add(case["id"])
    for path in sorted(WORKBOOK_CASES_DIR.glob("*.case.json")):
        doc = json.loads(path.read_text(encoding="utf-8"))
        for case in doc.get("cases") or []:
            if isinstance(case, dict) and isinstance(case.get("id"), str):
                ids.add(case["id"])
    return ids, suites


def load_case_ids() -> set[str]:
    """Compatibility helper for callers that only need individual IDs."""

    return load_case_catalog()[0]


_PLACEHOLDER_VERSION_RE = re.compile(r"16\.xx\.x", re.IGNORECASE)
_M365_BUILD_RE = re.compile(r"16\.\d+(\.\d+)?")


def is_pending_stamp(value: Any) -> bool:
    """Return whether the supplied stamp says it still needs live Excel.

    A verified stamp must be a bare Microsoft 365 build string, e.g.
    ``"16.111.2"`` or ``"16.112"``. Everything else counts as pending,
    including three specific non-evidence shapes CONTRIBUTING.md and
    tests/oracle/variants/win-365-ja_JP/ENVIRONMENT.md call out by name:

    - The literal ``"16.0"``. ``Application.Version`` reports this same
      bare major.0 string for every Office SKU from 2016 through 365 --
      it is not a build number and does not distinguish an Office 2019
      capture (which CONTRIBUTING.md forbids merging) from a genuine
      Microsoft 365 one.
    - The doc-template placeholder shape ``"16.xx.x (Build ...)"``.
    - Anything that isn't a bare ``16.<build>[.<patch>]`` string, e.g. a
      prose wrapper (``"Excel 365 (Mac, ja-JP, 16.111.2)"``) or a stamp
      that recorded a capture date instead of a build number
      (``"Excel 365 (Mac, ja-JP, 2026-07)"``). Both are semantically
      empty for stale-detection purposes even though they parse as
      non-empty strings.
    """

    if not isinstance(value, str) or not value.strip():
        return True
    normalized = value.strip()
    lowered = normalized.lower()
    if "unverified" in lowered or "needs live" in lowered or lowered == "unknown":
        return True
    if _PLACEHOLDER_VERSION_RE.search(normalized):
        return True
    if normalized == "16.0":
        return True
    return _M365_BUILD_RE.fullmatch(normalized) is None


def validate(path: Path, *, strict: bool) -> int:
    try:
        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except (OSError, yaml.YAMLError) as exc:
        print(f"FAIL {path}: cannot parse YAML: {exc}", file=sys.stderr)
        return 1
    entries = doc.get("entries") if isinstance(doc, dict) else None
    if not isinstance(entries, list):
        print(f"FAIL {path}: top-level entries must be a list", file=sys.stderr)
        return 1

    try:
        case_ids, suite_names = load_case_catalog()
    except (OSError, json.JSONDecodeError, yaml.YAMLError) as exc:
        print(f"FAIL case discovery: {exc}", file=sys.stderr)
        return 1

    errors: list[str] = []
    pending: list[str] = []
    documented = 0
    aliases = 0
    seen: set[str] = set()
    for index, entry in enumerate(entries):
        where = f"entries[{index}]"
        if not isinstance(entry, dict):
            errors.append(f"{where}: entry must be a mapping")
            continue
        selector_fields = [name for name in ("id", "ids", "suite", "suites") if name in entry]
        if len(selector_fields) != 1:
            errors.append(f"{where}: needs exactly one of id, ids, suite, or suites")
            continue
        selector = selector_fields[0]
        raw_selector = entry[selector]
        if selector in {"id", "suite"}:
            values = [raw_selector] if isinstance(raw_selector, str) and raw_selector else []
        else:
            values = raw_selector if isinstance(raw_selector, list) else []
        if not values or not all(isinstance(value, str) and value for value in values):
            errors.append(f"{where}: {selector} must contain non-empty strings")
            continue
        labels = [f"{selector}:{value}" for value in values]
        for label in labels:
            if label in seen:
                errors.append(f"{where}: duplicate selector {label!r}")
            seen.add(label)
        case_id = ", ".join(values)
        if not isinstance(entry.get("reason"), str) or not entry["reason"].strip():
            errors.append(f"{case_id}: missing non-empty reason")
        if not isinstance(entry.get("prefer"), str) or not entry["prefer"].strip():
            errors.append(f"{case_id}: missing non-empty prefer")

        alias = entry.get("case_alias")
        scope = entry.get("scope")
        if selector in {"suite", "suites"}:
            unknown = [value for value in values if value not in suite_names]
            if unknown:
                errors.append(f"{case_id}: unknown oracle suite(s): {', '.join(unknown)}")
            if alias is not None or scope is not None:
                errors.append(f"{case_id}: suite selector must not also set case_alias or scope")
        elif selector == "ids":
            unknown = [value for value in values if value not in case_ids]
            if unknown:
                errors.append(f"{case_id}: unknown oracle case ID(s): {', '.join(unknown)}")
            if alias is not None or scope is not None:
                errors.append(f"{case_id}: ids selector must not also set case_alias or scope")
        elif case_id in case_ids:
            if alias is not None or scope is not None:
                errors.append(f"{case_id}: direct oracle case must not also set case_alias or scope")
        elif alias is not None:
            if not isinstance(alias, str) or alias not in case_ids:
                errors.append(f"{case_id}: case_alias must name an existing oracle case")
            else:
                aliases += 1
        elif scope in NON_ORACLE_SCOPES:
            evidence = entry.get("evidence")
            if (
                not isinstance(evidence, list)
                or not evidence
                or not all(isinstance(item, str) and item for item in evidence)
            ):
                errors.append(f"{case_id}: scope {scope!r} requires a non-empty evidence list")
            else:
                missing = [item for item in evidence if not (REPO_ROOT / item).is_file()]
                if missing:
                    errors.append(f"{case_id}: evidence does not exist: {', '.join(missing)}")
                else:
                    documented += 1
        else:
            errors.append(
                f"{case_id}: orphan id; add case_alias or a scope from {sorted(NON_ORACLE_SCOPES)} with evidence"
            )

        if is_pending_stamp(entry.get("last_verified_excel_version")):
            pending.append(case_id)

    for message in errors:
        print(f"FAIL {message}", file=sys.stderr)
    for case_id in pending:
        print(f"PENDING {case_id}: needs a verified Excel version stamp", file=sys.stderr)

    print(
        f"divergence records: {len(entries)} entries, {len(case_ids)} oracle case IDs, "
        f"{aliases} aliases, {documented} documented non-oracle entries, {len(pending)} pending reprobes"
    )
    if errors or (strict and pending):
        return 1
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=DEFAULT_DIVERGENCE)
    parser.add_argument("--strict", action="store_true", help="also fail when an entry needs a live Excel reprobe")
    args = parser.parse_args()
    return validate(args.input, strict=args.strict)


if __name__ == "__main__":
    raise SystemExit(main())
