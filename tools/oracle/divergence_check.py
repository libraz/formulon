#!/usr/bin/env python3
"""Validate the provenance and review state of the skip registries.

Each divergence record must be backed by an oracle case, an explicit alias
to a current oracle case, or a non-oracle scope with repository-local
evidence.  The latter is for intentionally non-oracle surfaces such as C
API shape choices and binary-fixture limitations; it prevents a free-form
``id`` from silently becoming a permanent skip.

By default structural problems fail, and entries awaiting a fresh Excel
probe or riding on a build older than ``MIN_VERIFIED_BUILD`` are reported
as warnings.  ``--strict`` also fails for those pending and stale records,
making it suitable for a reprobe-completion gate.

``tests/ironcalc_divergence.yaml`` is a second registry with its own
schema, and ``--input`` accepts it too: its entries name imported IronCalc
cases rather than Formulon oracle cases, and their evidence is a Mac-side
probe golden rather than an Excel version stamp.  Both registries are
checked the same way in the ways that matter -- every entry resolves to a
live case, carries a machine-checkable reason for not running, and is
tallied in cases so the removed population is visible.
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
IRONCALC_DIVERGENCE = REPO_ROOT / "tests" / "ironcalc_divergence.yaml"
GOLDEN_DIR = REPO_ROOT / "tests" / "oracle" / "golden"
IRONCALC_GOLDEN_DIR = GOLDEN_DIR / "ironcalc"
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


def load_observations(golden_dirs: "list[Path] | None" = None) -> dict[str, str]:
    """Load case-id -> where Excel's answer for a still-skipped case is recorded.

    A capture drives skipped cases whose stamp is still pending and parks
    the result under `observed` (see `oracle_gen._load_divergence_reprobes`).
    Surfacing that here is what turns a pending entry from "someone should
    look at this one day" into "the evidence is on disk, go adjudicate it" --
    otherwise the observation sits in a golden nobody re-reads.
    """

    out: dict[str, str] = {}
    if golden_dirs is None:
        golden_dirs = [REPO_ROOT / "tests/oracle/golden", REPO_ROOT / "tests/oracle/golden_wb"]
    for golden_dir in golden_dirs:
        for path in sorted(golden_dir.glob("*.golden.json")):
            try:
                doc = json.loads(path.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            for case in doc.get("cases") or []:
                if not isinstance(case, dict) or not isinstance(case.get("id"), str):
                    continue
                if "observed" in case:
                    out[case["id"]] = f"{_display_path(path)} (observed)"
                elif "observed_error" in case:
                    out[case["id"]] = f"{_display_path(path)} (probe failed: {case['observed_error']})"
    return out


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

# Oldest build whose observation still counts as current evidence, keyed
# by the release train the stamp comes from. An entry stamped below its
# train's floor is stale: Excel may have changed under it since, and
# nothing re-reads the reason string. Raise a value here once a reprobe
# pass has moved the entries it covers.
#
# The two trains do not share a number line, so one number cannot gate
# both: macOS Microsoft 365 ships `16.<build>[.<patch>]` (16.112), while
# Windows Microsoft 365 ships `16.0.<build>` (16.0.20228) because
# `Application.Version` there is pinned at the Office major. A stamp is
# read as Windows when its second component is 0 -- macOS 365 has no
# 16.0 build, and the bare `16.0` string is rejected outright by
# `is_pending_stamp` as naming no particular Excel.
MIN_VERIFIED_BUILD: dict[str, tuple[int, ...]] = {
    "mac": (16, 110),
    "windows": (16, 0, 20228),
}


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


def _parse_build(value: str) -> "tuple[str, tuple[int, ...]] | None":
    """Split a verified stamp into its release train and numeric build.

    Returns ``None`` for anything `is_pending_stamp` would reject, so the
    two never disagree about which strings carry a comparable build.
    """

    if is_pending_stamp(value):
        return None
    numbers = tuple(int(part) for part in value.strip().split("."))
    train = "windows" if len(numbers) > 1 and numbers[1] == 0 else "mac"
    return train, numbers


def is_stale_stamp(value: Any) -> bool:
    """Return whether a verified stamp predates its train's floor.

    A pending stamp is not stale -- it carries no build to compare, and
    `is_pending_stamp` already reports it. Keeping the two separate is
    what lets a capture be promoted on an older-but-real build while the
    divergence registry still demands a recent one.
    """

    if not isinstance(value, str):
        return False
    parsed = _parse_build(value)
    if parsed is None:
        return False
    train, numbers = parsed
    return numbers < MIN_VERIFIED_BUILD[train]


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
    stale: list[str] = []
    documented = 0
    aliases = 0
    seen: set[str] = set()
    # Skipped *cases*, not entries: an entry selecting `ids: [a, b, c]`
    # removes three cases from the denominator, and the tally is only
    # meaningful in the same unit the pass rate is counted in.
    cause_counts: dict[str, int] = {cause: 0 for cause in sorted(SKIP_CAUSES)}
    # Cases whose skip is not currently backed by a build we accept. They
    # stay out of the per-cause tally -- a skip on a stamp Excel has moved
    # past is not evidence of anything -- and are reported on their own
    # line so the population is not quietly smaller than it looks.
    unverified_cases = 0
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

        stamp = entry.get("last_verified_excel_version")
        stamp_pending = is_pending_stamp(stamp)
        stamp_stale = is_stale_stamp(stamp)

        if entry.get("mode") == "skip-oracle":
            cause = entry.get("cause")
            if cause not in SKIP_CAUSES:
                errors.append(
                    f"{case_id}: skip-oracle needs `cause` from {sorted(SKIP_CAUSES)}; got {cause!r}. "
                    "Only excel-no-value may leave the release-gate denominator."
                )
            elif stamp_pending or stamp_stale:
                unverified_cases += selected_cases
            else:
                cause_counts[cause] += selected_cases

        if stamp_pending:
            pending.append(case_id)
        elif stamp_stale:
            stale.append(f"{case_id}: {stamp}")

    # Cross-registry consistency is a property of the tree rather than of
    # one file, so it is reported on whichever registry is being checked.
    errors.extend(cross_file_prefer_conflicts())

    for message in errors:
        print(f"FAIL {message}", file=sys.stderr)
    observations = load_observations()
    for case_id in pending:
        note = observations.get(case_id)
        detail = f"; Excel answered, see {note}" if note else ""
        print(f"PENDING {case_id}: needs a verified Excel version stamp{detail}", file=sys.stderr)
    floors = ", ".join(
        f"{train} {'.'.join(str(part) for part in build)}" for train, build in MIN_VERIFIED_BUILD.items()
    )
    for detail in stale:
        print(f"STALE {detail}: older than the accepted build floor ({floors})", file=sys.stderr)

    print(
        f"divergence records: {len(entries)} entries, {len(case_ids)} oracle case IDs, "
        f"{aliases} aliases, {documented} documented non-oracle entries, "
        f"{len(pending)} pending reprobes, {len(stale)} stale stamps"
    )
    # The pass-rate denominator is decided here and nowhere else in the
    # tree: no tool computes the release gate's 99.5%, so this tally is the
    # only artefact that says what the skipped cases actually are.
    skipped_cases = sum(cause_counts.values())
    print(f"skipped cases by cause: {skipped_cases} total (entries on an accepted build only)")
    notes = {
        "excel-no-value": " (outside the release-gate denominator)",
        "non-identifiable": " (both engines answer; the quantity has no unique value)",
        "unclassified": " (entry text does not determine the cause; needs a source review)",
    }
    for cause in sorted(cause_counts):
        print(f"  {cause}: {cause_counts[cause]}{notes.get(cause, '')}")
    print(f"  not tallied: {unverified_cases} (skip rides on a pending or stale Excel stamp)")
    if errors or (strict and (pending or stale)):
        return 1
    return 0


# Why an IronCalc-imported case is skipped. The registry's own header
# requires every entry to cite a Mac-side probe golden as evidence that
# the skip records an IronCalc divergence rather than a Formulon bug --
# but two populations cannot cite one, and lumping them in with the rest
# is what made the requirement unenforceable:
#
#   mac-probe          Formulon matches Mac Excel and IronCalc's cached
#                      value is the outlier. Evidence is the probe golden,
#                      so this cause is derived from the citation rather
#                      than declared: an entry that names a probe and sets
#                      no `cause` is counted here.
#   importer-flatten   The imported case is not the case Excel evaluated.
#                      The importer flattens each fixture into a
#                      single-sheet, all-rows-visible, static-cell setup,
#                      which drops live spill anchors, workbook geometry
#                      (SHEET / SHEETS), auto-filter visibility (SUBTOTAL
#                      9 / 109) and post-2019 error codes. No probe can
#                      adjudicate it because the golden cannot express the
#                      state the divergence is about.
#   float-precision    The two answers differ only at the precision limit
#                      of f64 -- an accumulation-order ULP, or a cached
#                      value drifting from the arithmetically exact
#                      result. The correct value is fixed by arithmetic,
#                      not by what Excel displays, so a probe adds nothing.
IRONCALC_CAUSES = {
    "mac-probe",
    "importer-flatten",
    "float-precision",
}

IRONCALC_PREFER = {"mac", "ironcalc"}

_FIRST_NOTED_RE = re.compile(r"\d{4}-\d{2}-\d{2}")
# `Probe: <name>` / `Probes: <a>, <b>` inside a reason string, the form
# every pre-existing entry uses. Names never contain a sentence period,
# so the citation ends at the first one.
_PROSE_PROBE_RE = re.compile(r"probes?:\s*([^.]+)", re.IGNORECASE)


def load_ironcalc_catalog(golden_dir: Path | None = None) -> tuple[set[str], int]:
    """Load the imported IronCalc case IDs and the golden count.

    The ID is the `<suite>.<addr>` composite the parameterized runner
    prints, which is also the key the importer matches skips against --
    so an entry that does not appear here matches nothing at all.
    """

    ids: set[str] = set()
    goldens = 0
    for path in sorted((golden_dir or IRONCALC_GOLDEN_DIR).glob("*.golden.json")):
        try:
            doc = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        suite = doc.get("suite")
        if not isinstance(suite, str):
            continue
        goldens += 1
        for case in doc.get("cases") or []:
            if isinstance(case, dict) and isinstance(case.get("id"), str):
                ids.add(f"{suite}.{case['id']}")
    return ids, goldens


def load_probe_names(golden_dir: Path | None = None) -> set[str]:
    """Names an entry may cite as probe evidence: file stem, suite or case.

    The IronCalc goldens are excluded even though they live under the
    same tree: they are imported from the same fixtures the skips are
    about, so citing one would be evidence of nothing.
    """

    root = golden_dir or GOLDEN_DIR
    names: set[str] = set()
    for path in sorted(root.rglob("*.golden.json")):
        if IRONCALC_GOLDEN_DIR in path.parents:
            continue
        names.add(path.name[: -len(".golden.json")])
        try:
            doc = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        if isinstance(doc.get("suite"), str):
            names.add(doc["suite"])
        for case in doc.get("cases") or []:
            if isinstance(case, dict) and isinstance(case.get("id"), str):
                names.add(case["id"])
    return names


def cited_probes(entry: dict) -> list[str]:
    """Probe goldens an entry cites, from the `probe` key or its reason."""

    raw = entry.get("probe")
    if isinstance(raw, str):
        candidates = [raw]
    elif isinstance(raw, list):
        candidates = [item for item in raw if isinstance(item, str)]
    else:
        match = _PROSE_PROBE_RE.search(entry.get("reason") or "")
        if not match:
            return []
        # Drop the parenthetical case list -- it names cases inside the
        # cited golden, which the golden itself already accounts for.
        candidates = re.split(r",| and ", re.sub(r"\([^)]*\)", " ", match.group(1)))
    names = []
    for candidate in candidates:
        name = candidate.strip().strip(".,;")
        if name.endswith(".golden.json"):
            name = name[: -len(".golden.json")]
        if name:
            names.append(name)
    return names


def validate_ironcalc(path: Path, *, golden_dir: Path | None = None, probe_dir: Path | None = None) -> int:
    """Validate `tests/ironcalc_divergence.yaml` against the imported corpus.

    Every entry names one imported case, so entries and removed cases are
    the same unit here; the tally is still printed against the corpus size
    because that is the denominator the skips come out of.
    """

    try:
        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except (OSError, yaml.YAMLError) as exc:
        print(f"FAIL {path}: cannot parse YAML: {exc}", file=sys.stderr)
        return 1
    entries = doc.get("entries") if isinstance(doc, dict) else None
    if not isinstance(entries, list):
        print(f"FAIL {path}: top-level entries must be a list", file=sys.stderr)
        return 1

    case_ids, goldens = load_ironcalc_catalog(golden_dir)
    probe_names = load_probe_names(probe_dir)

    errors: list[str] = []
    seen: set[str] = set()
    cause_counts: dict[str, int] = {cause: 0 for cause in sorted(IRONCALC_CAUSES)}
    for index, entry in enumerate(entries):
        where = f"entries[{index}]"
        if not isinstance(entry, dict):
            errors.append(f"{where}: entry must be a mapping")
            continue
        case_id = entry.get("id")
        if not isinstance(case_id, str) or not case_id:
            errors.append(f"{where}: needs a non-empty id")
            continue
        if case_id in seen:
            errors.append(f"{case_id}: duplicate id")
        seen.add(case_id)
        if case_id not in case_ids:
            errors.append(f"{case_id}: orphan id; no imported IronCalc case carries it")
        mode = entry.get("mode", "skip-oracle")
        if mode != "skip-oracle":
            errors.append(f"{case_id}: mode must be skip-oracle; got {mode!r}")
        if not isinstance(entry.get("reason"), str) or not entry["reason"].strip():
            errors.append(f"{case_id}: missing non-empty reason")
        prefer = entry.get("prefer")
        if prefer not in IRONCALC_PREFER:
            errors.append(f"{case_id}: prefer must be one of {sorted(IRONCALC_PREFER)}; got {prefer!r}")
        first_noted = entry.get("first_noted")
        if not _FIRST_NOTED_RE.fullmatch(str(first_noted)):
            errors.append(f"{case_id}: first_noted must be a YYYY-MM-DD date; got {first_noted!r}")

        probes = cited_probes(entry)
        unknown = [name for name in probes if name not in probe_names]
        if unknown:
            errors.append(f"{case_id}: probe evidence names no golden under tests/oracle/golden: {', '.join(unknown)}")
        cause = entry.get("cause")
        if cause is None:
            # Derived, not guessed: the citation is the evidence, and an
            # entry without one has to say which of the causes that
            # cannot cite a probe applies.
            cause = "mac-probe" if probes and not unknown else None
            if cause is None:
                errors.append(
                    f"{case_id}: cites no probe golden, so it needs a `cause` from "
                    f"{sorted(IRONCALC_CAUSES - {'mac-probe'})}"
                )
        elif cause not in IRONCALC_CAUSES:
            errors.append(f"{case_id}: cause must be one of {sorted(IRONCALC_CAUSES)}; got {cause!r}")
            cause = None
        elif cause == "mac-probe" and not probes:
            errors.append(f"{case_id}: cause mac-probe requires a probe citation")
        if cause in cause_counts:
            cause_counts[cause] += 1

    for message in errors:
        print(f"FAIL {message}", file=sys.stderr)

    print(
        f"ironcalc divergence records: {len(entries)} entries over {len(case_ids)} imported cases in {goldens} goldens"
    )
    print(f"skipped cases by cause: {sum(cause_counts.values())} total")
    for cause in sorted(cause_counts):
        print(f"  {cause}: {cause_counts[cause]}")
    return 1 if errors else 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=DEFAULT_DIVERGENCE)
    parser.add_argument("--strict", action="store_true", help="also fail when an entry needs a live Excel reprobe")
    args = parser.parse_args()
    if args.input.resolve() == IRONCALC_DIVERGENCE:
        return validate_ironcalc(args.input)
    return validate(args.input, strict=args.strict)


if __name__ == "__main__":
    raise SystemExit(main())
