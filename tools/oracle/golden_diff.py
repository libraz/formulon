#!/usr/bin/env python3
"""Compare two sets of formula-oracle goldens case by case.

Each side is a golden directory, or ``<git-ref>:<dir>`` to read the
directory as committed at that ref. Cases are matched by suite file and
case id; a numeric pair counts as equal within the newer side's suite
tolerance. Typical uses:

    golden_diff.py HEAD:tests/oracle/targets/mac-365-ja_JP/golden tests/oracle/targets/mac-365-ja_JP/golden
    golden_diff.py tests/oracle/targets/mac-365-ja_JP/golden tests/oracle/targets/mac-365-en_US/golden --summary

Only the ``expect`` / ``skipped`` state is compared; environment stamps
are ignored. Exit status is 0 whatever the differences are.
"""

from __future__ import annotations

import argparse
import json
import math
import subprocess
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]


def _load_side(spec: str) -> Dict[str, Dict[str, Any]]:
    """Returns {suite: golden document} for a directory or ``ref:dir`` spec."""

    out: Dict[str, Dict[str, Any]] = {}
    path = Path(spec)
    if ":" in spec and not path.exists():
        ref, rel = spec.split(":", 1)
        names = subprocess.run(
            ["git", "ls-tree", "--name-only", f"{ref}:{rel}"],
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            check=True,
        ).stdout.split()
        for name in names:
            if name.endswith(".golden.json"):
                text = subprocess.run(
                    ["git", "show", f"{ref}:{rel.rstrip('/')}/{name}"],
                    cwd=REPO_ROOT,
                    capture_output=True,
                    text=True,
                    check=True,
                ).stdout
                out[name[: -len(".golden.json")]] = json.loads(text)
        return out
    for file in sorted(path.glob("*.golden.json")):
        out[file.name[: -len(".golden.json")]] = json.loads(file.read_text(encoding="utf-8"))
    return out


def _state(case: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """The comparable part of a case: its expect, or a skipped marker."""

    if "expect" in case:
        return case["expect"]
    if "skipped" in case:
        return {"kind": "skipped"}
    return None


def _close(a: Any, b: Any, tol: Dict[str, float]) -> bool:
    if isinstance(a, bool) or isinstance(b, bool) or not isinstance(a, (int, float)) or not isinstance(b, (int, float)):
        return a == b
    if math.isnan(a) or math.isnan(b):
        return False
    diff = abs(a - b)
    return diff <= tol.get("abs", 0.0) or diff <= tol.get("rel", 0.0) * max(abs(a), abs(b))


def same_state(a: Optional[Dict[str, Any]], b: Optional[Dict[str, Any]], tol: Dict[str, float]) -> bool:
    """True when two case states agree, numbers within ``tol``."""

    if a is None or b is None:
        return a is b
    if a.get("kind") != b.get("kind") or a.get("shape") != b.get("shape") or a.get("code") != b.get("code"):
        return False
    va, vb = a.get("value"), b.get("value")
    if isinstance(va, list) and isinstance(vb, list):
        return len(va) == len(vb) and all(_close(x, y, tol) for x, y in zip(va, vb))
    return _close(va, vb, tol)


def _render(state: Optional[Dict[str, Any]]) -> str:
    if state is None:
        return "-"
    if state.get("kind") == "skipped":
        return "skipped"
    body = state.get("code", state.get("value"))
    text = json.dumps(body, ensure_ascii=False)
    if len(text) > 70:
        text = text[:67] + "..."
    shape = f" {state['shape']}" if state.get("shape") else ""
    return f"{state.get('kind')} {text}{shape}"


def diff_sides(
    old: Dict[str, Dict[str, Any]],
    new: Dict[str, Dict[str, Any]],
    suites: Optional[List[str]] = None,
    one_sided: bool = False,
) -> List[Tuple[str, str, str, Optional[Dict[str, Any]], Optional[Dict[str, Any]]]]:
    """Returns (suite, case_id, change, old_state, new_state) for every differing case.

    ``change`` is one of ``changed``, ``added``, ``removed``, ``skipped`` (a
    value became a skip) or ``unskipped``. Only suites on both sides are
    compared unless ``one_sided`` is set, which diffs the rest against an
    empty suite.
    """

    rows = []
    names = set(old) | set(new) if one_sided else set(old) & set(new)
    for suite in sorted(names):
        if suites and suite not in suites:
            continue
        o_doc, n_doc = old.get(suite, {}), new.get(suite, {})
        tol = (n_doc or o_doc).get("tolerance") or {}
        o_cases = {c["id"]: c for c in o_doc.get("cases", [])}
        n_cases = {c["id"]: c for c in n_doc.get("cases", [])}
        for case_id in list(o_cases) + [k for k in n_cases if k not in o_cases]:
            o_state = _state(o_cases[case_id]) if case_id in o_cases else None
            n_state = _state(n_cases[case_id]) if case_id in n_cases else None
            if case_id not in n_cases:
                change = "removed"
            elif case_id not in o_cases:
                change = "added"
            else:
                case_tol = n_cases[case_id].get("tolerance") or tol
                if same_state(o_state, n_state, case_tol):
                    continue
                o_skip = (o_state or {}).get("kind") == "skipped"
                n_skip = (n_state or {}).get("kind") == "skipped"
                change = "skipped" if n_skip and not o_skip else "unskipped" if o_skip and not n_skip else "changed"
            rows.append((suite, case_id, change, o_state, n_state))
    return rows


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("old", help="golden directory, or <git-ref>:<dir>")
    parser.add_argument("new", help="golden directory, or <git-ref>:<dir>")
    parser.add_argument("--suite", action="append", help="limit to a suite (repeatable)")
    parser.add_argument("--summary", action="store_true", help="print per-suite counts instead of cases")
    parser.add_argument("--change", action="append", help="only show these change kinds (repeatable)")
    parser.add_argument("--no-added", action="store_true", help="hide cases present only on the new side")
    parser.add_argument("--one-sided", action="store_true", help="also diff suites present on one side only")
    args = parser.parse_args(argv)

    rows = diff_sides(_load_side(args.old), _load_side(args.new), args.suite, args.one_sided)
    if args.change:
        rows = [r for r in rows if r[2] in args.change]
    if args.no_added:
        rows = [r for r in rows if r[2] != "added"]
    if args.summary:
        per_suite = Counter(r[0] for r in rows)
        kinds = Counter(r[2] for r in rows)
        for suite, count in per_suite.most_common():
            print(f"{count:6d}  {suite}")
        print(f"{len(rows):6d}  total  " + "  ".join(f"{k}={v}" for k, v in sorted(kinds.items())))
        return 0
    for suite, case_id, change, o_state, n_state in rows:
        print(f"{suite}.{case_id}: {change}: {_render(o_state)} => {_render(n_state)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
