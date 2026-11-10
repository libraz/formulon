#!/usr/bin/env python3
"""Golden projection for the pivot-refresh oracle family.

Each case under ``tests/oracle/cases_wb/pivot_refresh/<case>/`` holds two
Excel-saved workbooks and the golden derived from them:

  before.xlsx   source already edited, pivot cache still stale
  after.xlsx    the same workbook after Excel's Refresh (or after reopening,
                for ``trigger: open``)
  golden.json   ``meta`` + ``projection(after.xlsx)``

``engine_gaps.json`` beside the cases lists the cases the engine does not
match yet, each with a reason. ``run`` reports a listed case that still
differs as a CTest skip and fails a listed case that matches, so an entry
leaves the list together with the engine fix.

The projection is the comparison unit: every pivot cache definition and its
records, every pivot table definition, the rendered cells of each pivot
(over the union of its before and after locations, so leftovers count) and
every formula cell's cached value. Attributes Excel stamps on every save
(``refreshedDate``, ``refreshedBy``, version stamps, ``xr:uid``, ``cacheId``)
and ``extLst`` blocks are dropped; everything else is kept, so a match is
bit-for-bit on the modelled XML.

Commands:
  extract <xlsx> [--cells-from <before.xlsx>]   print the projection
  write <case_dir> --target T --excel-version V [--trigger refresh|open]
        [--expect ok|error] [--description TEXT]
  check [<case_dir>...]                         golden.json == projection(after.xlsx),
                                                engine_gaps.json names existing cases
  compare <case_dir> <engine.xlsx>              engine output vs golden; exit 1 on diff
  run --apply <exe> <case_dir>                  run the engine on before.xlsx and compare;
                                                exit 77 (CTest skip) for a `pending` case
                                                or a listed engine gap that still differs
"""

from __future__ import annotations

import argparse
import json
import posixpath
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree as ET

CASES_DIR = Path(__file__).resolve().parents[2] / "tests" / "oracle" / "cases_wb" / "pivot_refresh"
ENGINE_GAPS = CASES_DIR / "engine_gaps.json"
SKIP = 77

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# Root attributes Excel rewrites on every save; never part of the comparison.
_VOLATILE_ATTRS = {
    "refreshedBy",
    "refreshedDate",
    "refreshedDateIso",
    "createdVersion",
    "refreshedVersion",
    "minRefreshableVersion",
    "updatedVersion",
    "cacheId",
    "uid",
    "id",
    "Ignorable",
}
_DROPPED_CHILDREN = {"extLst"}


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _canon(elem: ET.Element, root: bool = False) -> List[Any]:
    """[tag, {attrs}, [children]] with namespaces stripped and volatile parts dropped."""
    attrs = {}
    for key, value in elem.attrib.items():
        name = _local(key)
        if name in ("uid",) or (root and name in _VOLATILE_ATTRS):
            continue
        attrs[name] = value
    children = [_canon(child) for child in elem if _local(child.tag) not in _DROPPED_CHILDREN]
    return [_local(elem.tag), attrs, children]


class _Package:
    def __init__(self, path: Path) -> None:
        self.zip = zipfile.ZipFile(path)
        self.names = set(self.zip.namelist())

    def xml(self, part: str) -> ET.Element:
        return ET.fromstring(self.zip.read(part))

    def rels(self, part: str) -> Dict[str, Tuple[str, str]]:
        """rId -> (type suffix, absolute part name) for `part`."""
        folder, base = posixpath.split(part)
        rels_part = posixpath.join(folder, "_rels", base + ".rels")
        if rels_part not in self.names:
            return {}
        out = {}
        for rel in self.xml(rels_part):
            target = rel.get("Target", "")
            if rel.get("TargetMode") == "External":
                continue
            resolved = (
                target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
            )
            out[rel.get("Id")] = (rel.get("Type", "").rsplit("/", 1)[-1], resolved)
        return out


def _shared_strings(pkg: _Package, wb_rels: Dict[str, Tuple[str, str]]) -> List[str]:
    for kind, part in wb_rels.values():
        if kind == "sharedStrings":
            out = []
            for si in pkg.xml(part):
                out.append("".join(t.text or "" for t in si.iter(f"{{{_MAIN_NS}}}t") if _not_phonetic(si, t)))
            return out
    return []


def _not_phonetic(si: ET.Element, t: ET.Element) -> bool:
    for rph in si.iter(f"{{{_MAIN_NS}}}rPh"):
        if t in list(rph):
            return False
    return True


_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _parse_ref(ref: str) -> Tuple[int, int, int, int]:
    first, _, last = ref.replace("$", "").partition(":")
    last = last or first
    m1, m2 = _CELL_RE.match(first), _CELL_RE.match(last)
    return (int(m1.group(2)), _col_num(m1.group(1)), int(m2.group(2)), _col_num(m2.group(1)))


def _cell_value(c: ET.Element, sst: List[str]) -> List[Any]:
    t = c.get("t", "n")
    v = c.find(f"{{{_MAIN_NS}}}v")
    if t == "inlineStr":
        is_ = c.find(f"{{{_MAIN_NS}}}is")
        text = "".join(x.text or "" for x in is_.iter(f"{{{_MAIN_NS}}}t")) if is_ is not None else ""
        return ["s", text]
    if v is None:
        return ["blank", None]
    if t == "s":
        return ["s", sst[int(v.text)]]
    if t == "str":
        return ["s", v.text or ""]
    if t == "b":
        return ["b", v.text == "1"]
    if t == "e":
        return ["e", v.text]
    return ["n", float(v.text)]


def _workbook(pkg: _Package):
    wb_part = "xl/workbook.xml"
    wb = pkg.xml(wb_part)
    wb_rels = pkg.rels(wb_part)
    sheets = []
    for s in wb.iter(f"{{{_MAIN_NS}}}sheet"):
        rid = s.get(f"{{{_REL_NS}}}id")
        sheets.append((s.get("name"), wb_rels[rid][1]))
    caches = []
    for pc in wb.iter(f"{{{_MAIN_NS}}}pivotCache"):
        caches.append(wb_rels[pc.get(f"{{{_REL_NS}}}id")][1])
    return wb_rels, sheets, caches


def _pivot_locations(path: Path) -> Dict[Tuple[str, str], str]:
    """(sheet, pivot name) -> location ref."""
    pkg = _Package(path)
    _, sheets, _ = _workbook(pkg)
    out = {}
    for sheet_name, sheet_part in sheets:
        for kind, part in pkg.rels(sheet_part).values():
            if kind == "pivotTable":
                root = pkg.xml(part)
                loc = root.find(f"{{{_MAIN_NS}}}location")
                out[(sheet_name, root.get("name"))] = loc.get("ref") if loc is not None else ""
    return out


def projection(path: Path, cells_from: Optional[Path] = None) -> Dict[str, Any]:
    pkg = _Package(path)
    wb_rels, sheets, cache_parts = _workbook(pkg)
    sst = _shared_strings(pkg, wb_rels)

    caches = []
    for part in cache_parts:
        definition = pkg.xml(part)
        records = None
        for kind, rpart in pkg.rels(part).values():
            if kind == "pivotCacheRecords":
                records = [[[_local(v.tag), v.get("v")] for v in r] for r in pkg.xml(rpart)]
        caches.append({"definition": _canon(definition, root=True), "records": records})

    extra_locations = _pivot_locations(cells_from) if cells_from else {}
    pivots = []
    sheet_xml = {}
    for sheet_name, sheet_part in sheets:
        for kind, part in sorted(pkg.rels(sheet_part).values(), key=lambda kv: kv[1]):
            if kind != "pivotTable":
                continue
            root = pkg.xml(part)
            cache_index = None
            for rkind, rpart in pkg.rels(part).values():
                if rkind == "pivotCacheDefinition" and rpart in cache_parts:
                    cache_index = cache_parts.index(rpart)
            loc = root.find(f"{{{_MAIN_NS}}}location")
            refs = [loc.get("ref")] if loc is not None else []
            prior = extra_locations.get((sheet_name, root.get("name")))
            if prior:
                refs.append(prior)
            if sheet_part not in sheet_xml:
                sheet_xml[sheet_part] = pkg.xml(sheet_part)
            pivots.append(
                {
                    "sheet": sheet_name,
                    "cache": cache_index,
                    "definition": _canon(root, root=True),
                    "cells": _cells_in(sheet_xml[sheet_part], refs, sst),
                }
            )

    formulas = {}
    for sheet_name, sheet_part in sheets:
        root = sheet_xml[sheet_part] if sheet_part in sheet_xml else pkg.xml(sheet_part)
        for c in root.iter(f"{{{_MAIN_NS}}}c"):
            if c.find(f"{{{_MAIN_NS}}}f") is not None:
                formulas[f"{sheet_name}!{c.get('r')}"] = _cell_value(c, sst)
    return {"caches": caches, "pivots": pivots, "formula_cells": formulas}


def _cells_in(sheet: ET.Element, refs: List[str], sst: List[str]) -> Dict[str, Any]:
    boxes = [_parse_ref(r) for r in refs if r]
    if not boxes:
        return {}
    r1 = min(b[0] for b in boxes)
    c1 = min(b[1] for b in boxes)
    r2 = max(b[2] for b in boxes)
    c2 = max(b[3] for b in boxes)
    out = {}
    for c in sheet.iter(f"{{{_MAIN_NS}}}c"):
        m = _CELL_RE.match(c.get("r", ""))
        if not m:
            continue
        row, col = int(m.group(2)), _col_num(m.group(1))
        if r1 <= row <= r2 and c1 <= col <= c2:
            value = _cell_value(c, sst)
            if value[0] != "blank":
                out[c.get("r")] = value
    return out


def write_golden(case_dir: Path, meta: Dict[str, Any]) -> Path:
    golden = {
        "meta": meta,
        "projection": projection(case_dir / "after.xlsx", cells_from=case_dir / "before.xlsx"),
    }
    out = case_dir / "golden.json"
    out.write_text(json.dumps(golden, ensure_ascii=False, indent=1, sort_keys=True) + "\n", encoding="utf-8")
    return out


def _diff(a: Any, b: Any, path: str, out: List[str], limit: int = 40) -> None:
    if len(out) >= limit:
        return
    if type(a) is not type(b):
        out.append(f"{path}: {json.dumps(a, ensure_ascii=False)[:200]} != {json.dumps(b, ensure_ascii=False)[:200]}")
    elif isinstance(a, dict):
        for key in sorted(set(a) | set(b)):
            if key not in a or key not in b:
                out.append(f"{path}.{key}: present only in {'golden' if key in a else 'actual'}")
            else:
                _diff(a[key], b[key], f"{path}.{key}", out, limit)
    elif isinstance(a, list):
        if len(a) != len(b):
            out.append(f"{path}: length {len(a)} != {len(b)}")
        for i, (x, y) in enumerate(zip(a, b)):
            _diff(x, y, f"{path}[{i}]", out, limit)
    elif a != b:
        out.append(f"{path}: {json.dumps(a, ensure_ascii=False)[:200]} != {json.dumps(b, ensure_ascii=False)[:200]}")


def diff_projection(golden: Dict[str, Any], actual: Dict[str, Any]) -> List[str]:
    out: List[str] = []
    _diff(golden, actual, "", out)
    return out


def load_engine_gaps(path: Path = ENGINE_GAPS) -> Dict[str, str]:
    if not path.is_file():
        return {}
    return json.loads(path.read_text(encoding="utf-8"))["gaps"]


def engine_gap_problems(gaps: Dict[str, str], case_names: List[str]) -> List[str]:
    problems = [f"engine gap {name!r} names no case" for name in gaps if name not in case_names]
    problems += [f"engine gap {name!r} has no reason" for name, reason in gaps.items() if not str(reason).strip()]
    return problems


def verdict(case: str, failures: List[str], gaps: Dict[str, str]) -> Tuple[int, List[str]]:
    """Maps a case's comparison failures to an exit code and report lines."""
    if case in gaps:
        if failures:
            return SKIP, [f"ENGINE-GAP {case}: {gaps[case]}", *failures]
        return 1, [f"FAIL {case}: listed in engine_gaps.json but now matches; remove the entry"]
    return (1 if failures else 0), failures


def _case_dirs(args: List[str]) -> List[Path]:
    if args:
        return [Path(a) for a in args]
    return sorted(p for p in CASES_DIR.iterdir() if (p / "golden.json").exists())


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("extract")
    p.add_argument("xlsx", type=Path)
    p.add_argument("--cells-from", type=Path)
    p = sub.add_parser("write")
    p.add_argument("case_dir", type=Path)
    p.add_argument("--target", required=True)
    p.add_argument("--excel-version", required=True)
    p.add_argument("--trigger", choices=("refresh", "open"), default="refresh")
    p.add_argument("--expect", choices=("ok", "error"), default="ok")
    p.add_argument("--description", default="")
    p = sub.add_parser("check")
    p.add_argument("case_dirs", nargs="*")
    p = sub.add_parser("compare")
    p.add_argument("case_dir", type=Path)
    p.add_argument("engine_xlsx", type=Path)
    p = sub.add_parser("run")
    p.add_argument("--apply", required=True, type=Path)
    p.add_argument("case_dir", type=Path)
    args = parser.parse_args(argv)

    if args.cmd == "extract":
        json.dump(projection(args.xlsx, args.cells_from), sys.stdout, ensure_ascii=False, indent=1, sort_keys=True)
        print()
        return 0
    if args.cmd == "write":
        meta = {
            "case": args.case_dir.name,
            "target": args.target,
            "excel_version": args.excel_version,
            "trigger": args.trigger,
            "expect": args.expect,
            "description": args.description,
        }
        print(write_golden(args.case_dir, meta))
        return 0
    if args.cmd == "check":
        failed = 0
        dirs = _case_dirs(args.case_dirs)
        for case_dir in dirs:
            golden = json.loads((case_dir / "golden.json").read_text(encoding="utf-8"))
            actual = projection(case_dir / "after.xlsx", cells_from=case_dir / "before.xlsx")
            diffs = diff_projection(golden["projection"], json.loads(json.dumps(actual)))
            if golden["meta"].get("case") != case_dir.name:
                diffs.append(f"meta.case {golden['meta'].get('case')!r} != directory {case_dir.name!r}")
            if diffs:
                failed += 1
                print(f"FAIL {case_dir.name}", *diffs[:10], sep="\n  ")
        print(f"{len(dirs) - failed}/{len(dirs)} pivot-refresh goldens match their after.xlsx")
        all_cases = [d.name for d in _case_dirs([])]
        gap_problems = engine_gap_problems(load_engine_gaps(), all_cases)
        for line in gap_problems:
            print(f"FAIL {line}")
        return 1 if failed or gap_problems or not dirs else 0
    if args.cmd == "compare":
        golden = json.loads((args.case_dir / "golden.json").read_text(encoding="utf-8"))
        actual = projection(args.engine_xlsx, cells_from=args.case_dir / "before.xlsx")
        diffs = diff_projection(golden["projection"], json.loads(json.dumps(actual)))
        for line in diffs:
            print(line)
        return 1 if diffs else 0
    if args.cmd == "run":
        return _run(args.apply, args.case_dir)
    return 2


def _run(apply_exe: Path, case_dir: Path) -> int:
    golden = json.loads((case_dir / "golden.json").read_text(encoding="utf-8"))
    meta = golden["meta"]
    if meta.get("pending"):
        print(f"SKIP {case_dir.name}: pending -- {meta['pending']}")
        return SKIP
    failures: List[str] = []
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / "engine.xlsx"
        cmd = [str(apply_exe), str(case_dir / "before.xlsx"), str(out), "--profile", meta["target"]]
        if meta.get("trigger") == "open":
            cmd.append("--open")
        proc = subprocess.run(cmd, capture_output=True, text=True)
        expected_code = 3 if meta.get("expect") == "error" else 0
        if proc.returncode != expected_code:
            failures = [
                f"FAIL {case_dir.name}: engine exit {proc.returncode}, expected {expected_code}",
                proc.stdout + proc.stderr,
            ]
        else:
            actual = projection(out, cells_from=case_dir / "before.xlsx")
            failures = diff_projection(golden["projection"], json.loads(json.dumps(actual)))
    code, lines = verdict(case_dir.name, failures, load_engine_gaps())
    for line in lines:
        print(line)
    return code


if __name__ == "__main__":
    sys.exit(main())
