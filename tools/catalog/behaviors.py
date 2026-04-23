#!/usr/bin/env python3
"""Formulon behaviour-vocabulary drift reporter.

The canonical `tools/catalog/functions.txt` enumerates Excel function
*names*; this tool operates one level deeper, on function *sub-behaviours*
(TEXT format codes, DATEVALUE ja-JP era strings, NUMBERVALUE locale
separators, CHAR / CODE CP932 ranges, SEARCH / FIND wildcards, argument
shapes). The `tests/oracle/cases/*.yaml` corpus is the sole source of
truth for behaviour; this file adds a declared-intent layer on top and
checks that declared `expected` matches observed oracle coverage.

Usage::

    tools/catalog/behaviors.py              # grouped, ANSI-coloured report
    tools/catalog/behaviors.py --check      # exit 1 on drift, 0 otherwise
    tools/catalog/behaviors.py --missing    # list behaviours observed missing
    tools/catalog/behaviors.py --expected diverged

The `--check` mode is wired into ctest as the `BehaviorDrift` test. If
PyYAML is not importable (CI sandbox, fresh macOS), the script prints a
warning to stderr and exits 0 — mirroring the soft-import pattern in
`tools/oracle/oracle_gen.py::_load_divergence_skips` so that ctest stays
green without a hard dependency.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple


REPO_ROOT = Path(__file__).resolve().parents[2]
BEHAVIORS_PATH = REPO_ROOT / "tools" / "catalog" / "behaviors.yaml"
CASES_DIR = REPO_ROOT / "tests" / "oracle" / "cases"
DIVERGENCE_PATH = REPO_ROOT / "tests" / "divergence.yaml"

VALID_EXPECTED = ("impl", "diverged", "missing")


# ---- ANSI helpers --------------------------------------------------------

def _supports_color() -> bool:
    if os.environ.get("NO_COLOR"):
        return False
    return sys.stdout.isatty()


_ANSI = _supports_color()


def _c(code: str, text: str) -> str:
    if not _ANSI:
        return text
    return f"\033[{code}m{text}\033[0m"


def green(s: str) -> str:
    return _c("32", s)


def red(s: str) -> str:
    return _c("31", s)


def yellow(s: str) -> str:
    return _c("33", s)


def bold(s: str) -> str:
    return _c("1", s)


def dim(s: str) -> str:
    return _c("2", s)


# ---- Data classes --------------------------------------------------------

@dataclass(frozen=True)
class Behavior:
    """A single declared sub-behaviour entry."""

    group_function: str
    group_aspect: str
    name: str
    probe: Optional[str]
    probe_regex: Optional[str]
    expected: str  # 'impl' | 'diverged' | 'missing'


@dataclass
class Group:
    function: str
    aspect: str
    behaviors: List[Behavior]


@dataclass(frozen=True)
class Case:
    """A single oracle case. `formula` is the raw `=...` string; `id` is
    the case identifier cross-referenced by `tests/divergence.yaml`."""

    id: str
    formula: str


@dataclass
class Observation:
    """Per-behaviour observation derived from the case corpus."""

    behavior: Behavior
    status: str  # 'impl' | 'diverged' | 'missing'
    n_match: int
    n_diverged: int
    n_live: int


# ---- Soft YAML import ----------------------------------------------------

def _try_import_yaml():
    """Returns the `yaml` module or None. Mirrors the soft-import style of
    `tools/oracle/oracle_gen.py::_load_divergence_skips`: we would rather
    ship a slightly reduced `behaviors.py` than add PyYAML to the hard
    dependency set of the native test harness."""
    try:
        import yaml  # type: ignore

        return yaml
    except Exception:
        return None


# ---- Loaders -------------------------------------------------------------

def load_behaviors(path: Path) -> List[Group]:
    """Parses `behaviors.yaml`. Raises FileNotFoundError if the file is
    missing; raises ValueError on malformed entries."""
    yaml = _try_import_yaml()
    if yaml is None:
        raise RuntimeError("PyYAML not available")
    if not path.exists():
        raise FileNotFoundError(str(path))
    doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    raw_groups = doc.get("groups") or []
    if not isinstance(raw_groups, list):
        raise ValueError("behaviors.yaml: `groups` must be a list")
    groups: List[Group] = []
    for raw_group in raw_groups:
        if not isinstance(raw_group, dict):
            raise ValueError("behaviors.yaml: group entries must be mappings")
        function = str(raw_group.get("function", "")).strip()
        aspect = str(raw_group.get("aspect", "")).strip()
        raw_behaviors = raw_group.get("behaviors") or []
        if not isinstance(raw_behaviors, list):
            raise ValueError(f"behaviors.yaml: group {function!r} has non-list behaviors")
        behaviors: List[Behavior] = []
        for raw in raw_behaviors:
            if not isinstance(raw, dict):
                raise ValueError(f"behaviors.yaml: behaviour in {function!r} is not a mapping")
            name = str(raw.get("name", "")).strip()
            probe = raw.get("probe")
            probe_regex = raw.get("probe_regex")
            expected = str(raw.get("expected", "")).strip()
            if not name:
                raise ValueError(f"behaviors.yaml: behaviour in {function!r} missing `name`")
            if (probe is None) == (probe_regex is None):
                raise ValueError(
                    f"behaviors.yaml: behaviour {name!r} must have exactly one of "
                    "`probe` or `probe_regex`"
                )
            if expected not in VALID_EXPECTED:
                raise ValueError(
                    f"behaviors.yaml: behaviour {name!r} has invalid `expected={expected!r}`;"
                    f" must be one of {VALID_EXPECTED}"
                )
            behaviors.append(
                Behavior(
                    group_function=function,
                    group_aspect=aspect,
                    name=name,
                    probe=str(probe) if probe is not None else None,
                    probe_regex=str(probe_regex) if probe_regex is not None else None,
                    expected=expected,
                )
            )
        groups.append(Group(function=function, aspect=aspect, behaviors=behaviors))
    return groups


def load_divergence(path: Path) -> Dict[str, str]:
    """Loads `case_id -> reason` for entries whose `mode == skip-oracle`.
    Mirrors `tools/oracle/oracle_gen.py::_load_divergence_skips` including
    the "missing file / bad YAML -> empty dict" tolerance."""
    if not path.exists():
        return {}
    yaml = _try_import_yaml()
    if yaml is None:
        return {}
    try:
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


def load_oracle_cases(cases_dir: Path) -> List[Case]:
    """Scans every YAML under `cases_dir` and extracts `(id, formula)`
    pairs. Silently tolerates cases without a `formula` field (spill /
    setup-only cases) by skipping them; tolerates missing dir by
    returning an empty list."""
    if not cases_dir.exists():
        return []
    yaml = _try_import_yaml()
    if yaml is None:
        return []
    cases: List[Case] = []
    for yaml_path in sorted(cases_dir.glob("*.yaml")):
        try:
            doc = yaml.safe_load(yaml_path.read_text(encoding="utf-8")) or {}
        except Exception:
            continue
        raw_cases = doc.get("cases") or []
        if not isinstance(raw_cases, list):
            continue
        for raw in raw_cases:
            if not isinstance(raw, dict):
                continue
            cid = raw.get("id")
            formula = raw.get("formula")
            if not isinstance(cid, str) or not isinstance(formula, str):
                continue
            cases.append(Case(id=cid, formula=formula))
    return cases


# ---- Probing -------------------------------------------------------------

def _match(behavior: Behavior, formula: str) -> bool:
    if behavior.probe is not None:
        return behavior.probe in formula
    # probe_regex path.
    assert behavior.probe_regex is not None  # guaranteed by load_behaviors
    return re.search(behavior.probe_regex, formula) is not None


def observe(
    behavior: Behavior, cases: Sequence[Case], divergence: Dict[str, str]
) -> Observation:
    """Counts how many cases match the behaviour's probe, how many of
    those are skip-oracle divergences, and derives the observed status:

        impl      if at least one case matches AND is not a divergence.
        diverged  if all matching cases are divergences (and at least one
                  matched).
        missing   if no cases match.
    """
    n_match = 0
    n_diverged = 0
    for case in cases:
        if _match(behavior, case.formula):
            n_match += 1
            if case.id in divergence:
                n_diverged += 1
    n_live = n_match - n_diverged
    if n_match == 0:
        status = "missing"
    elif n_live > 0:
        status = "impl"
    else:
        status = "diverged"
    return Observation(
        behavior=behavior,
        status=status,
        n_match=n_match,
        n_diverged=n_diverged,
        n_live=n_live,
    )


# ---- Reporting -----------------------------------------------------------

def _status_label(status: str) -> str:
    if status == "impl":
        return green("impl")
    if status == "diverged":
        return yellow("diverged")
    return red("missing")


def _drift_marker(obs: Observation) -> str:
    if obs.status == obs.behavior.expected:
        return " "
    return red("!")


def print_report(
    groups: Sequence[Group],
    observations: Dict[Tuple[str, str], Observation],
    expected_filter: Optional[str] = None,
) -> None:
    total = 0
    counts = {"impl": 0, "diverged": 0, "missing": 0}
    drift = 0

    for group in groups:
        visible = []
        for beh in group.behaviors:
            obs = observations[(beh.group_function, beh.name)]
            if expected_filter and beh.expected != expected_filter:
                continue
            visible.append((beh, obs))
            total += 1
            counts[obs.status] += 1
            if obs.status != beh.expected:
                drift += 1
        if not visible:
            continue
        title = f"## {group.function} — {group.aspect}"
        print(bold(title))
        for beh, obs in visible:
            marker = _drift_marker(obs)
            probe_repr = beh.probe if beh.probe is not None else f"re:{beh.probe_regex}"
            line = (
                f"  {marker} {beh.name:<44} "
                f"expected={beh.expected:<9} observed={_status_label(obs.status):<18} "
                f"(match={obs.n_match}, live={obs.n_live}, diverged={obs.n_diverged}) "
                f"{dim(probe_repr)}"
            )
            print(line)
        print()

    summary = (
        f"Total {total}: {green(str(counts['impl']) + ' impl')}, "
        f"{yellow(str(counts['diverged']) + ' diverged')}, "
        f"{red(str(counts['missing']) + ' missing')}. "
        f"Drift: {red(str(drift)) if drift else green('0')}"
    )
    print(bold(summary))


def print_missing(
    groups: Sequence[Group],
    observations: Dict[Tuple[str, str], Observation],
) -> None:
    """Print one line per behaviour whose *observed* status is `missing`.
    Pipe-friendly: `function :: aspect :: name`."""
    for group in groups:
        for beh in group.behaviors:
            obs = observations[(beh.group_function, beh.name)]
            if obs.status == "missing":
                print(f"{group.function} :: {group.aspect} :: {beh.name}")


def check_drift(
    groups: Sequence[Group],
    observations: Dict[Tuple[str, str], Observation],
) -> int:
    """Returns 0 on no-drift; prints and returns 1 on drift. The printed
    diff is meant to be copy-pasted into a commit body so reviewers can
    audit the declared-expected changes."""
    drifted: List[Observation] = []
    for group in groups:
        for beh in group.behaviors:
            obs = observations[(beh.group_function, beh.name)]
            if obs.status != beh.expected:
                drifted.append(obs)
    if not drifted:
        return 0
    print(bold(red(f"behaviour drift: {len(drifted)} entries")), file=sys.stderr)
    for obs in drifted:
        beh = obs.behavior
        probe_repr = beh.probe if beh.probe is not None else f"re:{beh.probe_regex}"
        print(
            f"  {beh.group_function} :: {beh.name}: "
            f"expected={beh.expected}, observed={obs.status} "
            f"(match={obs.n_match}, live={obs.n_live}, diverged={obs.n_diverged}) "
            f"[{probe_repr}]",
            file=sys.stderr,
        )
    return 1


# ---- main ----------------------------------------------------------------

def _compute_observations(
    groups: Sequence[Group], cases: Sequence[Case], divergence: Dict[str, str]
) -> Dict[Tuple[str, str], Observation]:
    out: Dict[Tuple[str, str], Observation] = {}
    for group in groups:
        for beh in group.behaviors:
            out[(beh.group_function, beh.name)] = observe(beh, cases, divergence)
    return out


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--check",
        action="store_true",
        help="Exit 1 if declared expected drifts from observed status. "
        "Used by the BehaviorDrift ctest entry.",
    )
    parser.add_argument(
        "--missing",
        action="store_true",
        help="Print one line per behaviour whose observed status is `missing`.",
    )
    parser.add_argument(
        "--report",
        action="store_true",
        help="Print the grouped human-readable report (default when no "
        "other mode is selected).",
    )
    parser.add_argument(
        "--expected",
        metavar="STATUS",
        choices=VALID_EXPECTED,
        default=None,
        help="Filter the --report output to behaviours whose declared "
        f"`expected` equals STATUS (one of {VALID_EXPECTED}).",
    )
    args = parser.parse_args(argv)

    yaml = _try_import_yaml()
    if yaml is None:
        # Soft-fail: keep CI green on hosts without PyYAML. The check mode
        # still exits 0 so ctest passes.
        print(
            "behaviors.py: PyYAML not available; skipping drift check. "
            "Install with `pip install pyyaml` to enable.",
            file=sys.stderr,
        )
        return 0

    try:
        groups = load_behaviors(BEHAVIORS_PATH)
    except FileNotFoundError:
        print(f"behaviors.py: {BEHAVIORS_PATH} not found", file=sys.stderr)
        return 2
    except (ValueError, RuntimeError) as exc:
        print(f"behaviors.py: {exc}", file=sys.stderr)
        return 2

    cases = load_oracle_cases(CASES_DIR)
    divergence = load_divergence(DIVERGENCE_PATH)
    observations = _compute_observations(groups, cases, divergence)

    if args.check:
        return check_drift(groups, observations)
    if args.missing:
        print_missing(groups, observations)
        return 0
    # Default: --report.
    print_report(groups, observations, expected_filter=args.expected)
    return 0


if __name__ == "__main__":
    sys.exit(main())
