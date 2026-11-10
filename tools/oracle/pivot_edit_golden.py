#!/usr/bin/env python3
"""Golden projection for the pivot-edit oracle family.

Each case under ``tests/oracle/cases_wb/pivot_edit/<case>/`` holds two
Excel-saved workbooks: ``before.xlsx`` (PivotTable1 on Sheet2 in its starting
state) and ``after.xlsx`` (the same workbook after Excel applied the edit).
``cases.yaml`` describes each edit and carries ``ops``, the same edit through
the engine's pivot C API. ``golden.json`` is ``meta`` + the pivot-refresh
projection of ``after.xlsx`` (see pivot_refresh_golden.py), so a match is
bit-for-bit on the modelled pivot XML and rendered cells.

``engine_gaps.json`` lists the cases the engine does not match yet, with the
same contract as the pivot-refresh family: a listed case that still differs
is a CTest skip, a listed case that matches fails.

Commands:
  write [<case>...]             write golden.json from after.xlsx
  check                         goldens match after.xlsx, cover cases.yaml; gaps name cases
  run --apply <exe> <case_dir>  apply the ops to before.xlsx and compare
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, List

import yaml

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from tools.oracle import pivot_refresh_golden as prg  # noqa: E402

CASES_DIR = Path(__file__).resolve().parents[2] / "tests" / "oracle" / "cases_wb" / "pivot_edit"
CASES_YAML = CASES_DIR / "cases.yaml"
ENGINE_GAPS = CASES_DIR / "engine_gaps.json"
EXCEL_VERSION = "16.114.4"


def load_cases() -> Dict[str, Any]:
    return yaml.safe_load(CASES_YAML.read_text(encoding="utf-8"))


def _projection(case_dir: Path, xlsx: Path) -> Dict[str, Any]:
    return json.loads(json.dumps(prg.projection(xlsx, cells_from=case_dir / "before.xlsx")))


def write(names: List[str]) -> int:
    doc = load_cases()
    for name in names or list(doc["cases"]):
        case_dir = CASES_DIR / name
        spec = doc["cases"][name]
        meta = {
            "case": name,
            "target": doc["target"],
            "excel_version": EXCEL_VERSION,
            "description": spec["edit"],
        }
        golden = {"meta": meta, "projection": _projection(case_dir, case_dir / "after.xlsx")}
        path = case_dir / "golden.json"
        path.write_text(json.dumps(golden, ensure_ascii=False, indent=1, sort_keys=True) + "\n", encoding="utf-8")
        print(path)
    return 0


def check() -> int:
    doc = load_cases()
    names = list(doc["cases"])
    problems: List[str] = []
    for name in names:
        case_dir = CASES_DIR / name
        if not (case_dir / "golden.json").is_file():
            problems.append(f"{name}: no golden.json")
            continue
        if not doc["cases"][name].get("ops"):
            problems.append(f"{name}: no ops")
        golden = json.loads((case_dir / "golden.json").read_text(encoding="utf-8"))
        diffs = prg.diff_projection(golden["projection"], _projection(case_dir, case_dir / "after.xlsx"))
        if golden["meta"].get("case") != name:
            diffs.append(f"meta.case {golden['meta'].get('case')!r} != {name!r}")
        problems += [f"{name}: {d}" for d in diffs[:10]]
    problems += prg.engine_gap_problems(prg.load_engine_gaps(ENGINE_GAPS), names)
    for line in problems:
        print(f"FAIL {line}")
    print(f"{len(names)} pivot-edit cases, {len(problems)} problem(s)")
    return 1 if problems or not names else 0


def run(apply_exe: Path, case_dir: Path) -> int:
    name = case_dir.name
    spec = load_cases()["cases"][name]
    golden = json.loads((case_dir / "golden.json").read_text(encoding="utf-8"))
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "engine.xlsx"
        cmd = [str(apply_exe), str(case_dir / "before.xlsx"), str(out), "--profile", golden["meta"]["target"]]
        for op in spec["ops"]:
            cmd += ["--op", op]
        proc = subprocess.run(cmd, capture_output=True, text=True, check=False)
        if proc.returncode != 0:
            failures = [f"FAIL {name}: engine exit {proc.returncode}", proc.stdout + proc.stderr]
        else:
            failures = prg.diff_projection(golden["projection"], _projection(case_dir, out))
    code, lines = prg.verdict(name, failures, prg.load_engine_gaps(ENGINE_GAPS))
    for line in lines:
        print(line)
    return code


def main(argv: List[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("write")
    p.add_argument("cases", nargs="*")
    sub.add_parser("check")
    p = sub.add_parser("run")
    p.add_argument("--apply", required=True, type=Path)
    p.add_argument("case_dir", type=Path)
    args = parser.parse_args(argv)
    if args.cmd == "write":
        return write(args.cases)
    if args.cmd == "check":
        return check()
    return run(args.apply, args.case_dir)


if __name__ == "__main__":
    sys.exit(main())
