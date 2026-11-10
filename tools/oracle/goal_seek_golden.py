#!/usr/bin/env python3
"""Golden capture and check for the Goal Seek oracle family.

Cases live in ``tests/oracle/cases_wb/goal_seek/cases.yaml``: a function of A1
in B1, a goal, a start value and Application.MaxChange. The golden
(``goal_seek.golden.json`` beside it) records, per case, Excel's GoalSeek
return value and the final A1 / B1 as IEEE-754 hex, read back bit-exactly
(the cell is split into two integers with ``INT(x*2^k)`` so no decimal
rendering is involved). A case whose formula is an error at the start value
is refused by Excel and recorded as ``"refused"``.

Commands:
  capture                       drive Mac Excel (xlwings) and write the golden
  check                         golden ids == cases.yaml ids, values well formed
  run --apply <exe> <case_id>   run the engine and compare; exit 77 (CTest skip)
                                for a pending case

The engine executable takes ``<formula> <goal> <x0> <max_change>`` and prints
one JSON object ``{"result": "converged"|"failed"|"refused", "x": hex, "y": hex}``.
"""

from __future__ import annotations

import argparse
import fnmatch
import json
import math
import struct
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List

import yaml

CASES_DIR = Path(__file__).resolve().parents[2] / "tests" / "oracle" / "cases_wb" / "goal_seek"
CASES_YAML = CASES_DIR / "cases.yaml"
GOLDEN = CASES_DIR / "goal_seek.golden.json"
SKIP = 77


def to_hex(value: float) -> str:
    return struct.pack(">d", value).hex()


def from_hex(text: str) -> float:
    return struct.unpack(">d", bytes.fromhex(text))[0]


def load_cases() -> List[Dict[str, Any]]:
    spec = yaml.safe_load(CASES_YAML.read_text(encoding="utf-8"))
    pending = spec.get("pending", [])
    cases = []

    def add(fn: str, x0: float, max_change: float) -> None:
        case_id = f"{fn}:{x0:g}" + ("" if max_change == 0.001 else f":mc{max_change:g}")
        cases.append(
            {
                "id": case_id,
                "formula": spec["functions"][fn],
                "goal": spec["goals"][fn],
                "x0": float(x0),
                "max_change": max_change,
                "pending": any(fnmatch.fnmatch(f"{fn}:{x0:g}", p) for p in pending),
            }
        )

    for fn in spec["functions"]:
        for x0 in spec["starts"]:
            add(fn, x0, 0.001)
    for variant in spec.get("max_change_variants", []):
        for mc in variant["max_change"]:
            add(variant["function"], variant["x0"], mc)
    return cases


def _exact(sheet: Any, addr: str) -> Any:
    """Read a numeric cell bit-exactly through two integer helper cells."""
    v = sheet.range(addr).value
    if not isinstance(v, (int, float)) or v == 0 or not math.isfinite(v):
        return v
    scale = 52 - math.floor(math.log2(abs(v))) - 26
    sheet.range("N1").formula = f"=INT({addr}*2^{scale})"
    sheet.range("N2").formula = f"=({addr}*2^{scale}-N1)*2^26"
    hi, lo = sheet.range("N1").value, sheet.range("N2").value
    sheet.range("N1:N2").clear_contents()
    return (hi + lo / 2**26) / 2**scale


def capture(_: argparse.Namespace) -> int:
    import xlwings as xw

    app = xw.apps.active
    app.display_alerts = False
    for book in list(app.books):
        book.close()
    book = app.books.add()
    sheet = book.sheets[0]
    results = {}
    try:
        app.api.iteration.set(False)
        for case in load_cases():
            app.api.max_change.set(case["max_change"])
            sheet.range("A1").value = case["x0"]
            sheet.range("B1").formula = case["formula"]
            try:
                flag = sheet.range("B1").api.goal_seek(goal=case["goal"], changing_cell=sheet.range("A1").api)
                result = "converged" if flag is True else "failed"
            except Exception:  # Excel refuses when B1 is an error at x0.
                result = "refused"
            x, y = _exact(sheet, "A1"), _exact(sheet, "B1")
            entry: Dict[str, Any] = {"result": result, "x": to_hex(float(x))}
            if isinstance(y, (int, float)):
                entry["y"] = to_hex(float(y))
            results[case["id"]] = entry
    finally:
        app.api.max_change.set(0.001)
        book.close()
        app.display_alerts = True
    doc = {
        "suite": "goal_seek",
        "environment": {"target": "mac-365-ja_JP", "excel_version": str(app.version)},
        "cases": results,
    }
    GOLDEN.write_text(json.dumps(doc, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(f"wrote {len(results)} cases to {GOLDEN}")
    return 0


def check(_: argparse.Namespace) -> int:
    golden = json.loads(GOLDEN.read_text(encoding="utf-8"))["cases"]
    ids = [c["id"] for c in load_cases()]
    problems = [f"missing golden: {i}" for i in ids if i not in golden]
    problems += [f"golden without case: {i}" for i in golden if i not in ids]
    for case_id, entry in golden.items():
        if entry.get("result") not in ("converged", "failed", "refused"):
            problems.append(f"{case_id}: bad result {entry.get('result')!r}")
        for key in ("x", "y"):
            if key in entry and len(entry[key]) != 16:
                problems.append(f"{case_id}: {key} is not a 64-bit hex")
    for p in problems:
        print(p)
    print(f"{len(ids) - len(problems)}/{len(ids)} ok" if not problems else f"{len(problems)} problem(s)")
    return 1 if problems else 0


def run(args: argparse.Namespace) -> int:
    case = next((c for c in load_cases() if c["id"] == args.case_id), None)
    if case is None:
        print(f"unknown case {args.case_id}")
        return 1
    if case["pending"]:
        print(f"{args.case_id}: pending")
        return SKIP
    expected = json.loads(GOLDEN.read_text(encoding="utf-8"))["cases"][args.case_id]
    proc = subprocess.run(
        [args.apply, case["formula"], repr(float(case["goal"])), repr(case["x0"]), repr(case["max_change"])],
        capture_output=True,
        text=True,
        check=False,
    )
    if proc.returncode != 0:
        print(proc.stdout + proc.stderr)
        return 1
    actual = json.loads(proc.stdout)
    diffs = [
        f"{key}: expected {expected.get(key)} got {actual.get(key)}"
        for key in ("result", "x", "y")
        if expected.get(key) != actual.get(key)
    ]
    for d in diffs:
        print(f"{args.case_id}: {d}")
    return 1 if diffs else 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)
    sub.add_parser("capture").set_defaults(func=capture)
    sub.add_parser("check").set_defaults(func=check)
    p_run = sub.add_parser("run")
    p_run.add_argument("--apply", required=True)
    p_run.add_argument("case_id")
    p_run.set_defaults(func=run)
    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
