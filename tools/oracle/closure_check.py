#!/usr/bin/env python3
"""Oracle-closure verifier for Formulon.

Encodes the machine-readable definition of "this function has been
closed against the Mac Excel 365 oracle". A function is *closed* when
all six conditions hold:

  1. behaviors.yaml declares >= min_perspectives_per_member entries that
     belong to a `function: FN` group.
  2. tests/oracle/cases/*.yaml contains >= 1 case whose formula calls
     FN, covering each declared behavior probe.
  3. tests/oracle/golden/*.json contains a corresponding entry for
     every case under (2). (Existence only; numeric verification is the
     oracle-verify step's job.)
  4. Every divergence.yaml entry whose `id` matches a case from (2)
     carries `reason`, `prefer`, and `last_verified_excel_version`.
  5. coverage_gap.py reports FN as not in pilot_candidates.
  6. behaviors.py --check exits cleanly (no observed-vs-expected drift).

Conditions 1-4 are evaluated per-function; condition 5 is per-function
and condition 6 is global.

Usage
-----

    closure_check.py LAMBDA               # check one function
    closure_check.py family=lambda        # check one family
    closure_check.py --status LAMBDA      # show state, never fail
    closure_check.py --report             # all functions + summary
    closure_check.py --audit-family lambda

Exit codes:
    0 = all targets closed (or --status / --report; never fails)
    1 = at least one closure condition failed

PyYAML required; resolve via the oracle venv:
    tools/oracle/.venv/bin/python tools/oracle/closure_check.py ...
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
CATALOG_DIR = REPO_ROOT / "tools" / "catalog"
ORACLE_DIR = REPO_ROOT / "tools" / "oracle"
CASES_DIR = REPO_ROOT / "tests" / "oracle" / "cases"
GOLDEN_DIR = REPO_ROOT / "tests" / "oracle" / "golden"
BEHAVIORS_PATH = CATALOG_DIR / "behaviors.yaml"
FAMILIES_PATH = CATALOG_DIR / "families.yaml"
DIVERGENCE_PATH = REPO_ROOT / "tests" / "divergence.yaml"

DEFAULT_MIN_PERSPECTIVES = 5

sys.path.insert(0, str(CATALOG_DIR))
sys.path.insert(0, str(ORACLE_DIR))
import coverage_gap  # type: ignore  # noqa: E402
import status as catalog_status  # type: ignore  # noqa: E402

try:
    import yaml as _yaml  # type: ignore
except ImportError:
    _yaml = None


def _need_yaml() -> None:
    if _yaml is None:
        sys.stderr.write(
            "closure_check: PyYAML not available. Run via the oracle venv:\n"
            "  tools/oracle/.venv/bin/python tools/oracle/closure_check.py ...\n"
        )
        sys.exit(2)


@dataclass
class Family:
    name: str
    members: List[str]
    category_file: Optional[Path]
    min_perspectives_per_member: int
    cluster_reason: str
    # Members Excel treats as an alternate spelling of another member, as
    # `{alias: canonical}`. An alias cannot hold oracle cases of its own —
    # Excel rewrites it to the canonical name before it ever stores or
    # evaluates the formula — so its closure is the canonical member's.
    aliases: Dict[str, str] = field(default_factory=dict)


@dataclass
class ConditionResult:
    name: str
    ok: bool
    detail: str = ""


@dataclass
class FunctionClosure:
    fn: str
    family: Family
    conditions: List[ConditionResult] = field(default_factory=list)
    # Set when `fn` is an alias: the canonical member whose conditions these
    # are.
    alias_of: Optional[str] = None

    @property
    def closed(self) -> bool:
        return all(c.ok for c in self.conditions)


@dataclass
class CoverageContext:
    catalog: Set[str]
    implemented: Set[str]
    yaml_hits: Dict[str, Set[Path]]
    ironcalc_hits: Dict[str, Set[Path]]
    native_hits: Dict[str, Set[Path]]
    behaviors: List[Dict]
    divergence_entries: List[Dict]
    drift_result: ConditionResult


def load_families() -> Tuple[Dict[str, Family], Dict[str, Family]]:
    _need_yaml()
    if not FAMILIES_PATH.exists():
        return {}, {}
    raw = _yaml.safe_load(FAMILIES_PATH.read_text(encoding="utf-8")) or {}
    by_name: Dict[str, Family] = {}
    by_member: Dict[str, Family] = {}
    for entry in raw.get("families", []):
        cf = entry.get("category_file")
        if cf:
            cf_path = Path(cf)
            if not cf_path.is_absolute():
                cf_path = REPO_ROOT / cf_path
        else:
            cf_path = None
        fam = Family(
            name=entry["name"],
            members=list(entry.get("members", [])),
            category_file=cf_path,
            min_perspectives_per_member=int(entry.get("min_perspectives_per_member", DEFAULT_MIN_PERSPECTIVES)),
            cluster_reason=entry.get("cluster_reason", ""),
            aliases=dict(entry.get("aliases") or {}),
        )
        by_name[fam.name] = fam
        for m in fam.members:
            by_member[m] = fam
    return by_name, by_member


def singleton_family(fn: str) -> Family:
    return Family(
        name=fn.lower(),
        members=[fn],
        category_file=CASES_DIR / f"{fn.lower()}.yaml",
        min_perspectives_per_member=DEFAULT_MIN_PERSPECTIVES,
        cluster_reason="",
    )


def resolve_family(target: str, by_name: Dict[str, Family], by_member: Dict[str, Family]) -> Family:
    # Function-name lookup (by_member) takes precedence over family-name
    # lookup (by_name). Some function names collide with family names — e.g.
    # the INFO function vs. the `info` family of IS* / type-introspection
    # functions — and a bare `INFO` argument should be interpreted as the
    # function. The explicit `family=NAME` selector in main() routes the
    # other way for callers that want the family.
    fn = target.upper()
    if fn in by_member:
        return by_member[fn]
    if target.lower() in by_name:
        return by_name[target.lower()]
    return singleton_family(fn)


def load_behaviors() -> List[Dict]:
    _need_yaml()
    if not BEHAVIORS_PATH.exists():
        return []
    raw = _yaml.safe_load(BEHAVIORS_PATH.read_text(encoding="utf-8")) or {}
    out: List[Dict] = []
    for group in raw.get("groups", []):
        gfn = group.get("function", "")
        gaspect = group.get("aspect", "")
        for b in group.get("behaviors", []):
            out.append(
                {
                    "function": gfn,
                    "aspect": gaspect,
                    "name": b.get("name", ""),
                    "probe": b.get("probe"),
                    "probe_regex": b.get("probe_regex"),
                    "expected": b.get("expected", "missing"),
                }
            )
    return out


def load_divergence() -> List[Dict]:
    _need_yaml()
    if not DIVERGENCE_PATH.exists():
        return []
    raw = _yaml.safe_load(DIVERGENCE_PATH.read_text(encoding="utf-8")) or {}
    return list(raw.get("entries", []))


def load_cases_for_family(family: Family) -> List[Dict]:
    _need_yaml()
    files: List[Path] = []
    if family.category_file and family.category_file.exists():
        files = [family.category_file]
    else:
        files = sorted(CASES_DIR.glob("*.yaml"))
    cases: List[Dict] = []
    for fp in files:
        try:
            raw = _yaml.safe_load(fp.read_text(encoding="utf-8")) or {}
        except (_yaml.YAMLError, OSError):
            continue
        for c in raw.get("cases", []) or []:
            cases.append(c)
    return cases


def load_golden_ids_for_family(family: Family) -> Set[str]:
    out: Set[str] = set()
    candidates: List[Path] = []
    if family.category_file:
        candidate = GOLDEN_DIR / (family.category_file.stem + ".golden.json")
        if candidate.exists():
            candidates = [candidate]
    if not candidates:
        candidates = sorted(GOLDEN_DIR.glob("*.golden.json"))
    for fp in candidates:
        try:
            data = json.loads(fp.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            continue
        for c in data.get("cases", []) or []:
            cid = c.get("id")
            if cid:
                out.add(cid)
    return out


def fn_call_re(fn: str) -> re.Pattern:
    return re.compile(rf"(?<![A-Za-z0-9_]){re.escape(fn)}\s*\(", re.IGNORECASE)


def behaviors_for_function(fn: str, all_behaviors: List[Dict]) -> List[Dict]:
    """A behavior is 'for FN' only when its group's `function:` field
    matches FN. Probe-based ownership is not used because formulas often
    nest functions (e.g. BYROW(..., LAMBDA(...))) and cross-attributing
    behaviors inflates the count for the inner function."""
    return [b for b in all_behaviors if b["function"].upper() == fn.upper()]


def cases_for_function(fn: str, cases: List[Dict]) -> List[Dict]:
    pat = fn_call_re(fn)
    return [c for c in cases if pat.search(c.get("formula", ""))]


def behavior_probe_match_count(behavior: Dict, cases: List[Dict]) -> int:
    n = 0
    if behavior.get("probe"):
        needle = behavior["probe"]
        for c in cases:
            if needle in c.get("formula", ""):
                n += 1
    elif behavior.get("probe_regex"):
        try:
            rx = re.compile(behavior["probe_regex"])
        except re.error:
            return 0
        for c in cases:
            if rx.search(c.get("formula", "")):
                n += 1
    return n


def check_condition_1_behaviors(fn: str, all_behaviors: List[Dict], min_required: int) -> ConditionResult:
    matched = behaviors_for_function(fn, all_behaviors)
    ok = len(matched) >= min_required
    detail = f"{len(matched)} declared (need >= {min_required})"
    return ConditionResult("behaviors_declared", ok, detail)


def check_condition_2_cases(fn: str, fn_behaviors: List[Dict], cases: List[Dict]) -> ConditionResult:
    fn_cases = cases_for_function(fn, cases)
    if not fn_cases:
        return ConditionResult("cases_present", False, "no oracle case calls this function")
    if not fn_behaviors:
        return ConditionResult("cases_cover_behaviors", False, "no behaviors declared (see condition 1)")
    uncovered: List[str] = []
    for b in fn_behaviors:
        if behavior_probe_match_count(b, fn_cases) == 0:
            uncovered.append(b["name"] or "<unnamed>")
    ok = len(uncovered) == 0
    if ok:
        detail = f"{len(fn_cases)} cases cover all {len(fn_behaviors)} behaviors"
    else:
        detail = (
            f"{len(uncovered)}/{len(fn_behaviors)} behaviors uncovered: "
            + ", ".join(uncovered[:3])
            + ("..." if len(uncovered) > 3 else "")
        )
    return ConditionResult("cases_cover_behaviors", ok, detail)


def check_condition_3_golden(fn: str, cases: List[Dict], golden_ids: Set[str]) -> ConditionResult:
    fn_cases = cases_for_function(fn, cases)
    if not fn_cases:
        return ConditionResult("golden_present", False, "no cases to verify against golden")
    missing = [c["id"] for c in fn_cases if c.get("id") not in golden_ids]
    ok = len(missing) == 0
    if ok:
        detail = f"{len(fn_cases)} cases all have golden entries"
    else:
        detail = (
            f"{len(missing)}/{len(fn_cases)} cases missing golden: "
            + ", ".join(missing[:3])
            + ("..." if len(missing) > 3 else "")
        )
    return ConditionResult("golden_present", ok, detail)


def check_condition_4_divergence(fn: str, cases: List[Dict], divergence_entries: List[Dict]) -> ConditionResult:
    fn_case_ids = {c.get("id") for c in cases_for_function(fn, cases)}
    div_for_fn = [e for e in divergence_entries if e.get("id") in fn_case_ids]
    if not div_for_fn:
        return ConditionResult("divergence_documented", True, "no divergences")
    incomplete: List[str] = []
    for e in div_for_fn:
        missing_fields = []
        if not e.get("reason"):
            missing_fields.append("reason")
        if e.get("mode") != "skip-oracle":
            if not e.get("prefer"):
                missing_fields.append("prefer")
            if not e.get("last_verified_excel_version"):
                missing_fields.append("last_verified_excel_version")
        if missing_fields:
            incomplete.append(f"{e.get('id')}({', '.join(missing_fields)})")
    ok = len(incomplete) == 0
    if ok:
        detail = f"{len(div_for_fn)} divergences all documented"
    else:
        detail = f"{len(incomplete)}/{len(div_for_fn)} divergences incomplete: " + ", ".join(incomplete[:3])
    return ConditionResult("divergence_documented", ok, detail)


def check_condition_5_coverage_gap(fn: str, ctx: CoverageContext) -> ConditionResult:
    if fn not in ctx.catalog:
        return ConditionResult("not_in_pilot", False, f"{fn} not in functions.txt catalog")
    if fn not in ctx.implemented:
        return ConditionResult("not_in_pilot", False, f"{fn} not implemented")
    sources: List[str] = []
    if fn in ctx.yaml_hits:
        sources.append("yaml")
    if fn in ctx.ironcalc_hits:
        sources.append("ironcalc")
    if fn in ctx.native_hits:
        sources.append("native")
    has_oracle = bool(sources)
    detail = f"oracle source(s): {','.join(sources) or 'NONE'}"
    return ConditionResult("not_in_pilot", has_oracle, detail)


def check_condition_6_behavior_drift() -> ConditionResult:
    behaviors_py = CATALOG_DIR / "behaviors.py"
    if not behaviors_py.exists():
        return ConditionResult("behavior_drift", True, "behaviors.py not present, skipped")
    proc = subprocess.run(
        [sys.executable, str(behaviors_py), "--check"],
        capture_output=True,
        text=True,
    )
    ok = proc.returncode == 0
    if ok:
        detail = "drift=0"
    else:
        last = (proc.stdout.strip().splitlines() or ["drift detected"])[-1]
        detail = last
    return ConditionResult("behavior_drift", ok, detail)


def build_coverage_context() -> CoverageContext:
    _, catalog = catalog_status.load_catalog(catalog_status.CATALOG_PATH)
    return CoverageContext(
        catalog=catalog,
        implemented=catalog_status.scan_implemented(REPO_ROOT),
        yaml_hits=coverage_gap.scan_dir(coverage_gap.ORACLE_CASES_DIR, "*.yaml"),
        ironcalc_hits=coverage_gap.scan_dir(coverage_gap.IRONCALC_GOLDEN_DIR, "*.json"),
        native_hits=coverage_gap.scan_native_golden(coverage_gap.ORACLE_GOLDEN_DIR),
        behaviors=load_behaviors(),
        divergence_entries=load_divergence(),
        drift_result=check_condition_6_behavior_drift(),
    )


def evaluate_function(
    fn: str,
    family: Family,
    cases: List[Dict],
    golden_ids: Set[str],
    ctx: CoverageContext,
) -> FunctionClosure:
    alias_target = family.aliases.get(fn)
    if alias_target is not None:
        # The canonical member is evaluated here rather than its declaration
        # being taken on trust, so an alias cannot close while the function it
        # defers to is open.
        if alias_target == fn or alias_target in family.aliases or alias_target not in family.members:
            detail = f"alias target {alias_target} is not a non-alias member of family {family.name}"
            return FunctionClosure(
                fn=fn,
                family=family,
                conditions=[ConditionResult("alias_resolves", False, detail)],
                alias_of=alias_target,
            )
        target = evaluate_function(alias_target, family, cases, golden_ids, ctx)
        return FunctionClosure(fn=fn, family=family, conditions=target.conditions, alias_of=alias_target)

    fn_behaviors = behaviors_for_function(fn, ctx.behaviors)
    fc = FunctionClosure(fn=fn, family=family)
    fc.conditions.append(check_condition_1_behaviors(fn, ctx.behaviors, family.min_perspectives_per_member))
    fc.conditions.append(check_condition_2_cases(fn, fn_behaviors, cases))
    fc.conditions.append(check_condition_3_golden(fn, cases, golden_ids))
    fc.conditions.append(check_condition_4_divergence(fn, cases, ctx.divergence_entries))
    fc.conditions.append(check_condition_5_coverage_gap(fn, ctx))
    fc.conditions.append(ctx.drift_result)
    return fc


def audit_family(family: Family, all_behaviors: List[Dict]) -> List[Tuple[str, str, str]]:
    warnings: List[Tuple[str, str, str]] = []
    by_member: Dict[str, Set[str]] = {}
    aspect_origin: Dict[str, str] = {}
    for member in family.members:
        # An alias declares no behaviors of its own; auditing it against the
        # family's aspect union would report every one of them as missing.
        if member in family.aliases:
            continue
        aspects = {b["aspect"] for b in behaviors_for_function(member, all_behaviors) if b["aspect"]}
        by_member[member] = aspects
        for a in aspects:
            aspect_origin.setdefault(a, member)
    union = set().union(*by_member.values()) if by_member else set()
    for member, has in by_member.items():
        for missing in sorted(union - has):
            warnings.append((member, missing, aspect_origin[missing]))
    return warnings


GREEN = "\033[32m"
RED = "\033[31m"
YELLOW = "\033[33m"
RESET = "\033[0m"


def _isatty() -> bool:
    return sys.stdout.isatty()


def _colour(text: str, code: str) -> str:
    return f"{code}{text}{RESET}" if _isatty() else text


def print_function_closure(fc: FunctionClosure, verbose: bool = True) -> None:
    status = _colour("CLOSED", GREEN) if fc.closed else _colour("OPEN", RED)
    suffix = f"  (alias of {fc.alias_of})" if fc.alias_of else ""
    print(f"{fc.fn:20}  family={fc.family.name:24}  {status}{suffix}")
    if verbose or not fc.closed:
        for c in fc.conditions:
            mark = _colour("OK", GREEN) if c.ok else _colour("FAIL", RED)
            print(f"  [{mark:>4}]  {c.name:28}  {c.detail}")


def print_family_audit(family: Family, warnings: List[Tuple[str, str, str]]) -> None:
    if not warnings:
        print(_colour(f"family={family.name}: cross-member audit clean", GREEN))
        return
    print(_colour(f"family={family.name}: {len(warnings)} cross-member perspective gaps", YELLOW))
    for member, missing_aspect, origin in warnings:
        print(f"  {member:16}  missing aspect '{missing_aspect}' (declared on {origin})")


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("target", nargs="?")
    parser.add_argument("--status", action="store_true")
    parser.add_argument("--report", action="store_true")
    parser.add_argument("--audit-family", metavar="NAME")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args(argv)

    _need_yaml()
    by_name, by_member = load_families()

    if args.audit_family:
        if args.audit_family not in by_name:
            print(f"unknown family: {args.audit_family}", file=sys.stderr)
            return 2
        warnings = audit_family(by_name[args.audit_family], load_behaviors())
        print_family_audit(by_name[args.audit_family], warnings)
        return 0

    ctx = build_coverage_context()

    if args.report:
        targets = sorted(ctx.catalog & ctx.implemented)
    elif args.target:
        if args.target.startswith("family="):
            fname = args.target.split("=", 1)[1]
            if fname not in by_name:
                print(f"unknown family: {fname}", file=sys.stderr)
                return 2
            targets = list(by_name[fname].members)
        else:
            targets = [args.target.upper()]
    else:
        parser.print_help()
        return 2

    closures: List[FunctionClosure] = []
    family_cache: Dict[str, Tuple[List[Dict], Set[str]]] = {}
    for fn in targets:
        family = resolve_family(fn, by_name, by_member)
        if family.name not in family_cache:
            family_cache[family.name] = (
                load_cases_for_family(family),
                load_golden_ids_for_family(family),
            )
        cases, golden_ids = family_cache[family.name]
        fc = evaluate_function(fn, family, cases, golden_ids, ctx)
        closures.append(fc)

    if args.json:
        out = {
            "drift_ok": ctx.drift_result.ok,
            "closures": [
                {
                    "fn": fc.fn,
                    "family": fc.family.name,
                    "closed": fc.closed,
                    "alias_of": fc.alias_of,
                    "conditions": [{"name": c.name, "ok": c.ok, "detail": c.detail} for c in fc.conditions],
                }
                for fc in closures
            ],
        }
        json.dump(out, sys.stdout, indent=2)
        sys.stdout.write("\n")
        return 0 if (args.status or args.report or all(fc.closed for fc in closures)) else 1

    closed = sum(1 for fc in closures if fc.closed)
    total = len(closures)
    print(f"closure_check: {closed}/{total} closed, behavior-drift={'ok' if ctx.drift_result.ok else 'FAIL'}")
    print()
    verbose = not args.report
    for fc in closures:
        if args.report and fc.closed:
            print_function_closure(fc, verbose=False)
        else:
            print_function_closure(fc, verbose=verbose)
            print()

    if args.status or args.report:
        return 0
    return 0 if all(fc.closed for fc in closures) else 1


if __name__ == "__main__":
    sys.exit(main())
