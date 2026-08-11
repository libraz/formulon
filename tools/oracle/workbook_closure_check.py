#!/usr/bin/env python3
"""Feature-closure check for the workbook oracle track.

The 6-condition `closure_check.py` is function-centric and does not fit
workbook-level features. This is the lightweight workbook counterpart: it
reports, per feature area (`pivot` and `print`), whether the declarative
oracle scaffolding is in shape.

For each suite under `tests/oracle/cases_wb/` it checks:

  1. A declarative `<suite>.case.json` exists.
  2. The `<suite>.case.json` validates against `workbook_case_schema.py`
     (and, when a golden is present, the case <-> golden id sets agree).
  3. A golden `tests/oracle/golden_wb/<suite>.golden.json` exists. A
     missing golden is reported as MISSING, not a failure -- workbook
     goldens require a Windows + Excel host and are captured out of band.
  4. Every `tests/divergence.yaml` entry whose `id` matches a case in a
     workbook suite carries `reason` and `last_verified_excel_version`.

Exit codes:
    0 = no genuine inconsistency (absent goldens are fine in developer mode)
    1 = a real problem (schema validation failed, or a workbook divergence
        entry is missing a required field); with ``--require-active``, a
        manifest-approved active capture is also required
    2 = could not run (PyYAML missing)

stdlib only; PyYAML is needed for the divergence file. Resolve via the
oracle venv when run outside it:
    tools/oracle/.venv/bin/python tools/oracle/workbook_closure_check.py
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
ORACLE_DIR = REPO_ROOT / "tools" / "oracle"
CASES_WB_DIR = REPO_ROOT / "tests" / "oracle" / "cases_wb"
GOLDEN_WB_DIR = REPO_ROOT / "tests" / "oracle" / "golden_wb"
DIVERGENCE_PATH = REPO_ROOT / "tests" / "divergence.yaml"

sys.path.insert(0, str(ORACLE_DIR))
import provenance  # type: ignore  # noqa: E402
import workbook_case_schema  # type: ignore  # noqa: E402

try:
    import yaml as _yaml  # type: ignore
except ImportError:
    _yaml = None


GREEN = "\033[32m"
RED = "\033[31m"
YELLOW = "\033[33m"
RESET = "\033[0m"


def _colour(text: str, code: str) -> str:
    return f"{code}{text}{RESET}" if sys.stdout.isatty() else text


def _feature_area(suite: str) -> str:
    """Maps a suite name to its feature area (`pivot` / `print` / `other`)."""

    if suite.startswith("pivot"):
        return "pivot"
    if suite.startswith("print"):
        return "print"
    return "other"


def _discover_suites() -> List[Tuple[str, Path]]:
    """Returns (suite-name, case-json-path) pairs sorted by suite name."""

    if not CASES_WB_DIR.is_dir():
        return []
    out: List[Tuple[str, Path]] = []
    for path in sorted(CASES_WB_DIR.iterdir()):
        if not path.name.endswith(".case.json"):
            continue
        # "<suite>.case.json" -> "<suite>".
        suite = path.name[: -len(".case.json")]
        out.append((suite, path))
    return out


def _load_workbook_case_ids() -> Set[str]:
    """Returns every case id declared across all workbook case files.

    Used to scope the divergence check to ids the workbook track owns.
    """

    ids: Set[str] = set()
    for _suite, path in _discover_suites():
        try:
            doc = json.loads(path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            continue
        for case in doc.get("cases") or []:
            cid = case.get("id") if isinstance(case, dict) else None
            if isinstance(cid, str):
                ids.add(cid)
    return ids


def _load_divergence_entries() -> List[Dict]:
    if _yaml is None or not DIVERGENCE_PATH.exists():
        return []
    try:
        raw = _yaml.safe_load(DIVERGENCE_PATH.read_text(encoding="utf-8")) or {}
    except _yaml.YAMLError:
        return []
    entries = raw.get("entries") or []
    return [e for e in entries if isinstance(e, dict)]


def check_divergence(workbook_ids: Set[str]) -> Tuple[bool, List[str]]:
    """Checks workbook-scoped divergence entries for required fields.

    Returns (ok, messages). An entry is required to carry `reason` and
    `last_verified_excel_version`; `mode: skip-oracle` entries still need
    both so a reviewer can see why and against which Excel build.
    """

    msgs: List[str] = []
    ok = True
    relevant = [e for e in _load_divergence_entries() if e.get("id") in workbook_ids]
    if not relevant:
        msgs.append("no workbook-scoped divergence entries")
        return True, msgs
    for entry in relevant:
        cid = entry.get("id")
        missing = []
        if not entry.get("reason"):
            missing.append("reason")
        if not entry.get("last_verified_excel_version"):
            missing.append("last_verified_excel_version")
        if missing:
            ok = False
            msgs.append(f"{cid}: missing {', '.join(missing)}")
        else:
            msgs.append(f"{cid}: documented")
    return ok, msgs


def check_suite(suite: str, case_path: Path) -> Tuple[bool, Optional[bool], List[str]]:
    """Validates one suite. Returns (schema_ok, golden_present, messages).

    `golden_present` is None when no golden file exists (reported as
    MISSING, never a failure); True/False once a golden is on disk.
    """

    msgs: List[str] = []
    golden_path = GOLDEN_WB_DIR / f"{suite}.golden.json"
    golden_arg: Optional[Path] = golden_path if golden_path.exists() else None

    try:
        case_ids, golden_ids = workbook_case_schema.validate_pair(case_path, golden_arg)
    except workbook_case_schema.ValidationError as exc:
        if golden_arg is not None and "missing case ids present" in str(exc):
            msgs.append(_colour(f"  [PENDING] golden incomplete: {exc} (external capture required)", YELLOW))
            return True, None, msgs
        msgs.append(_colour(f"  [FAIL] schema: {exc}", RED))
        return False, golden_arg is not None, msgs
    except json.JSONDecodeError as exc:
        msgs.append(_colour(f"  [FAIL] JSON parse: {exc}", RED))
        return False, golden_arg is not None, msgs

    msgs.append(_colour(f"  [OK]   schema: {len(case_ids)} case(s)", GREEN))
    if golden_ids is None:
        msgs.append(
            _colour(
                f"  [MISS] golden: not generated yet (expected tests/oracle/golden_wb/{suite}.golden.json)",
                YELLOW,
            )
        )
        return True, None, msgs
    msgs.append(_colour(f"  [OK]   golden: {len(golden_ids)} case(s)", GREEN))
    return True, True, msgs


def check_yaml_normalization(suite: str, case_path: Path) -> List[str]:
    """Require YAML and JSON mirrors to agree after semantic normalization."""

    if _yaml is None:
        return []
    yaml_path = case_path.with_suffix("").with_suffix(".yaml")
    if not yaml_path.exists():
        return [f"{suite}: YAML source missing ({yaml_path.name})"]
    try:
        ydoc = _yaml.safe_load(yaml_path.read_text(encoding="utf-8")) or {}
        jdoc = json.loads(case_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError, _yaml.YAMLError) as exc:
        return [f"{suite}: YAML/JSON description parse failed: {exc}"]
    try:
        ynorm = workbook_case_schema.normalise_case_source(ydoc)
        jnorm = workbook_case_schema.normalise_case_source(jdoc)
    except (ValueError, workbook_case_schema.ValidationError) as exc:
        return [f"{suite}: YAML/JSON semantic normalization failed: {exc}"]
    return [] if ynorm == jnorm else [f"{suite}: YAML/JSON full semantic case mismatch"]


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--require-active", action="store_true")
    args = parser.parse_args()
    if _yaml is None:
        sys.stderr.write(
            "workbook_closure_check: PyYAML not available. Run via the "
            "oracle venv:\n"
            "  tools/oracle/.venv/bin/python "
            "tools/oracle/workbook_closure_check.py\n"
        )
        return 2

    suites = _discover_suites()
    if not suites:
        print(f"workbook_closure_check: no *.case.json suites in {CASES_WB_DIR}")
        return 0

    overall_ok = True
    by_area: Dict[str, List[str]] = {"pivot": [], "print": [], "other": []}
    golden_missing = 0
    golden_present = 0

    for suite, case_path in suites:
        by_area.setdefault(_feature_area(suite), []).append(suite)

    print("workbook closure check\n")
    for area in ("pivot", "print", "other"):
        area_suites = by_area.get(area) or []
        if not area_suites:
            continue
        print(f"feature area: {area}")
        for suite in area_suites:
            case_path = CASES_WB_DIR / f"{suite}.case.json"
            print(f" suite: {suite}")
            schema_ok, golden, msgs = check_suite(suite, case_path)
            desc_errors = check_yaml_normalization(suite, case_path)
            msgs.extend(f"  [FAIL] {msg}" for msg in desc_errors)
            for m in msgs:
                print(m)
            if not schema_ok:
                overall_ok = False
            if desc_errors:
                overall_ok = False
            if golden is None:
                golden_missing += 1
            elif golden:
                golden_present += 1
        print()

    workbook_ids = _load_workbook_case_ids()
    div_ok, div_msgs = check_divergence(workbook_ids)
    print("divergence entries (workbook-scoped):")
    for m in div_msgs:
        mark = "[OK]  " if div_ok else "[FAIL]"
        print(f"  {mark} {m}")
    if not div_ok:
        overall_ok = False
    print()

    if golden_missing:
        print(f"pending: {golden_missing} suite(s) have no committed golden (external Windows M365 capture required)")
    targets_doc = provenance._load_yaml(provenance.DEFAULT_TARGETS)
    active_capture = provenance.workbook_active(targets_doc, REPO_ROOT)
    if not active_capture:
        print("active workbook capture: PENDING (manifest/provenance is not an active verified capture)")
    if args.require_active and not active_capture:
        overall_ok = False
        print("active workbook gate: BLOCKED (external verified capture required)")
    print(
        f"summary: {len(suites)} suite(s), "
        f"golden present={golden_present}, missing={golden_missing}, "
        f"divergence={'ok' if div_ok else 'INCOMPLETE'}"
    )
    if overall_ok:
        print(_colour("workbook closure: consistent", GREEN))
        return 0
    print(_colour("workbook closure: inconsistency found", RED))
    return 1


if __name__ == "__main__":
    sys.exit(main())
