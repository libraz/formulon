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
VARIANTS_DIR = REPO_ROOT / "tests" / "oracle" / "variants"
NON_ORACLE_SCOPES = {
    "api-contract",
    "environment",
    "fixture",
    "deferred-feature",
    "unit-test",
}

# Why a `mode: skip-oracle` entry is not verified against Excel. Required,
# because the release gate excludes skipped cases from the pass-rate
# denominator on the grounds that the case "can never pass" -- and that is
# true of exactly one of these causes. The others are debts of different
# lifetimes, and the tally below is what keeps them from ageing into
# permanent exemptions on a reason string nobody re-reads.
#
#   excel-no-value          Excel produces no comparable value at all: a
#                           volatile / non-reproducible result, or a formula
#                           it rejects at entry. Nothing we build will make
#                           this case comparable. This is the only cause the
#                           gate's justification actually covers.
#   harness-cannot-capture  Excel answers; our bridge cannot record the
#                           answer (xlwings writes "" as blank, an error
#                           reads back as blank, a spill is larger than the
#                           capture ceiling). Fixable on our side; treat the
#                           count as a bug queue.
#   accepted-divergence     Both engines answer and we deliberately differ,
#                           preferring ours (an Excel quirk we don't
#                           reproduce, a stubbed external, a documented API
#                           shape). Long-lived by design.
#   engine-gap              Excel answers and Formulon is wrong or not there
#                           yet. Must come back: the entry is deleted when
#                           the engine work lands.
#   non-identifiable        Both engines answer and the quantity itself has
#                           no unique value to agree on: a smoothing weight
#                           the data cannot pin down once the fit saturates,
#                           an F statistic whose residual is zero in the
#                           limit. Neither answer is the wrong one, and no
#                           engine work closes it -- which is what separates
#                           this from `engine-gap`. It is not
#                           `accepted-divergence` either, because we did not
#                           choose to differ; the mathematics did.
#   unclassified            The entry's own text does not determine which of
#                           the others applies -- typically because it states
#                           two different blockers, or states none in these
#                           terms. Deliberately not a judgement call: a
#                           guessed cause is worse than an admitted gap,
#                           because nothing re-examines a confident label.
#                           The tally reports this count separately so it
#                           can be worked down against the entries' sources.
SKIP_CAUSES = {
    "excel-no-value",
    "harness-cannot-capture",
    "accepted-divergence",
    "engine-gap",
    "non-identifiable",
    "unclassified",
}


def load_case_catalog() -> tuple[set[str], dict[str, int]]:
    """Load the raw case IDs and the per-suite case counts used by both tracks.

    The suite counts exist so a `suite:`-selected skip can be tallied in
    cases rather than entries: one such entry removes a whole suite from
    the pass-rate denominator.
    """

    ids: set[str] = set()
    suites: dict[str, int] = {}
    for path in sorted(FORMULA_CASES_DIR.glob("*.yaml")):
        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        cases = [case for case in doc.get("cases") or [] if isinstance(case, dict) and isinstance(case.get("id"), str)]
        if isinstance(doc.get("suite"), str):
            suites[doc["suite"]] = len(cases)
        for case in cases:
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


def divergence_files() -> list[Path]:
    """Every divergence registry in the tree: the primary plus each variant."""

    return [DEFAULT_DIVERGENCE, *sorted(VARIANTS_DIR.glob("*/divergence.yaml"))]


def _display_path(path: Path) -> str:
    """Repo-relative path when possible, absolute otherwise (temp files)."""

    try:
        return str(path.relative_to(REPO_ROOT))
    except ValueError:
        return str(path)


def _prefer_by_case(path: Path) -> dict[str, str]:
    """Maps case ID -> `prefer` for the entries that name individual cases.

    Suite selectors are skipped: resolving one to its members would need
    the case catalog, and the defect this feeds is case-level.
    """

    out: dict[str, str] = {}
    try:
        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except (OSError, yaml.YAMLError):
        return out
    for entry in doc.get("entries") or []:
        if not isinstance(entry, dict):
            continue
        prefer = entry.get("prefer")
        if not isinstance(prefer, str) or not prefer.strip():
            continue
        if isinstance(entry.get("id"), str):
            values = [entry["id"]]
        elif isinstance(entry.get("ids"), list):
            values = [value for value in entry["ids"] if isinstance(value, str)]
        else:
            continue
        for case_id in values:
            out[case_id] = prefer.strip()
    return out


def cross_file_prefer_conflicts(paths: list[Path] | None = None) -> list[str]:
    """Reports cases whose `prefer` verdict differs between registries.

    The primary file and a variant file can each carry an entry for the
    same case -- that is normal, because a divergence can be real on both
    targets. What is not tenable is the two disagreeing about which
    implementation we trust: the same divergence recorded with opposite
    verdicts means at least one is wrong. This was found by hand once (four
    GROUPBY / PIVOTBY cases, `prefer: mac-excel-365` against
    `prefer: formulon`, both citing the same rejected formula), so it is
    checked mechanically now.
    """

    by_file = {path: _prefer_by_case(path) for path in (paths or divergence_files()) if path.is_file()}
    conflicts: list[str] = []
    case_ids = {case_id for mapping in by_file.values() for case_id in mapping}
    for case_id in sorted(case_ids):
        verdicts = {_display_path(path): mapping[case_id] for path, mapping in by_file.items() if case_id in mapping}
        if len(set(verdicts.values())) > 1:
            rendered = ", ".join(f"{where} says {verdict!r}" for where, verdict in sorted(verdicts.items()))
            conflicts.append(f"{case_id}: contradictory `prefer` across registries -- {rendered}")
    return conflicts


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
        case_ids, suite_case_counts = load_case_catalog()
    except (OSError, json.JSONDecodeError, yaml.YAMLError) as exc:
        print(f"FAIL case discovery: {exc}", file=sys.stderr)
        return 1

    errors: list[str] = []
    pending: list[str] = []
    documented = 0
    aliases = 0
    seen: set[str] = set()
    # Skipped *cases*, not entries: an entry selecting `ids: [a, b, c]`
    # removes three cases from the denominator, and the tally is only
    # meaningful in the same unit the pass rate is counted in.
    cause_counts: dict[str, int] = {cause: 0 for cause in sorted(SKIP_CAUSES)}
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
        # How many real oracle cases this entry removes from the run. An
        # alias / non-oracle-scope entry documents something that is not an
        # oracle case at all, so it removes none and must not inflate the
        # skip tally below.
        selected_cases = 0
        if selector in {"suite", "suites"}:
            unknown = [value for value in values if value not in suite_case_counts]
            if unknown:
                errors.append(f"{case_id}: unknown oracle suite(s): {', '.join(unknown)}")
            if alias is not None or scope is not None:
                errors.append(f"{case_id}: suite selector must not also set case_alias or scope")
            selected_cases = sum(suite_case_counts.get(value, 0) for value in values)
        elif selector == "ids":
            unknown = [value for value in values if value not in case_ids]
            if unknown:
                errors.append(f"{case_id}: unknown oracle case ID(s): {', '.join(unknown)}")
            if alias is not None or scope is not None:
                errors.append(f"{case_id}: ids selector must not also set case_alias or scope")
            selected_cases = sum(1 for value in values if value in case_ids)
        elif case_id in case_ids:
            if alias is not None or scope is not None:
                errors.append(f"{case_id}: direct oracle case must not also set case_alias or scope")
            selected_cases = 1
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

        if entry.get("mode") == "skip-oracle":
            cause = entry.get("cause")
            if cause not in SKIP_CAUSES:
                errors.append(
                    f"{case_id}: skip-oracle needs `cause` from {sorted(SKIP_CAUSES)}; got {cause!r}. "
                    "Only excel-no-value may leave the release-gate denominator."
                )
            else:
                cause_counts[cause] += selected_cases

        if is_pending_stamp(entry.get("last_verified_excel_version")):
            pending.append(case_id)

    # Cross-registry consistency is a property of the tree rather than of
    # one file, so it is reported on whichever registry is being checked.
    errors.extend(cross_file_prefer_conflicts())

    for message in errors:
        print(f"FAIL {message}", file=sys.stderr)
    for case_id in pending:
        print(f"PENDING {case_id}: needs a verified Excel version stamp", file=sys.stderr)

    print(
        f"divergence records: {len(entries)} entries, {len(case_ids)} oracle case IDs, "
        f"{aliases} aliases, {documented} documented non-oracle entries, {len(pending)} pending reprobes"
    )
    # The pass-rate denominator is decided here and nowhere else in the
    # tree: no tool computes the release gate's 99.5%, so this tally is the
    # only artefact that says what the skipped cases actually are.
    skipped_cases = sum(cause_counts.values())
    print(f"skipped cases by cause: {skipped_cases} total")
    notes = {
        "excel-no-value": " (outside the release-gate denominator)",
        "non-identifiable": " (both engines answer; the quantity has no unique value)",
        "unclassified": " (entry text does not determine the cause; needs a source review)",
    }
    for cause in sorted(cause_counts):
        print(f"  {cause}: {cause_counts[cause]}{notes.get(cause, '')}")
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
