#!/usr/bin/env python3
"""Coverage-gap analyzer for the Formulon oracle test suites.

Cross-references the catalog at `tools/catalog/functions.txt` against three
notions of "is there a test for this function?":

  * implementation status   - re-used from `tools/catalog/status.py`
  * native Formulon yaml    - `tests/oracle/cases/*.yaml`
  * golden JSON corpora     - `tests/oracle/targets/mac-365-ja_JP/golden/**/*.json`
                              (the imported ironcalc fixtures live under
                              `tests/oracle/external/ironcalc/golden/` and are
                              counted separately from native goldens)

The default report lists *pilot candidates*: functions that are implemented
but have no oracle coverage in any of the three sources. Those are the safe
targets for the next round of `make oracle-gen` work — Mac Excel 365 (ja-JP)
acts as the authoritative oracle for fresh yaml suites we author in
`tests/oracle/cases/`.

`--shapes` adds a second layer: for every implemented function it reads the
yaml cases that call it and tags each call's arguments as array (an array
constant or a range), error (an error literal, NA(), or a cell set up to hold
an error), blank (a referenced cell the setup leaves empty) or omitted (an
empty argument slot), then ranks functions by how many of array / error /
blank no case exercises. That layer needs PyYAML (`make oracle-setup`); the
rest is stdlib only. Works on CPython 3.8+.

Usage:
    tools/oracle/coverage_gap.py                 # pilot candidates table
    tools/oracle/coverage_gap.py --all           # full per-function table
    tools/oracle/coverage_gap.py --json          # machine-readable
    tools/oracle/coverage_gap.py --missing-impl  # functions with no impl AND no oracle
    tools/oracle/coverage_gap.py --shapes        # functions ranked by untested argument shapes
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Set

REPO_ROOT = Path(__file__).resolve().parents[2]
CATALOG_DIR = REPO_ROOT / "tools" / "catalog"
ORACLE_CASES_DIR = REPO_ROOT / "tests" / "oracle" / "cases"
ORACLE_GOLDEN_DIR = REPO_ROOT / "tests" / "oracle" / "targets" / "mac-365-ja_JP" / "golden"
IRONCALC_GOLDEN_DIR = REPO_ROOT / "tests" / "oracle" / "external" / "ironcalc" / "golden"

sys.path.insert(0, str(CATALOG_DIR))
import status as catalog_status  # type: ignore  # noqa: E402

# Function-call detection. A name is uppercase ASCII, may contain digits,
# `_`, and `.` (for namespaced names like `FORECAST.ETS.SEASONALITY`), and
# is immediately followed by `(`. The look-behind rejects matches that
# would just be the tail of a longer identifier (`MyABS(` -> reject), but
# allows a leading `.` so OOXML-style namespace prefixes like
# `_xlfn._xlws.SORT(` still surface the call.
FUNCTION_CALL_RE = re.compile(r"(?<![A-Za-z0-9_])([A-Z][A-Z0-9_.]*)\s*\(")


# Built-ins Excel accepts only as operators: files store the internal name,
# but Mac Excel rejects it typed at entry, so cases spell the operator. Both
# sides must be reference characters so prose like `e.g.:` never matches.
_REF_BEFORE = r"(?<=[A-Za-z0-9$)\]])"
_REF_AFTER = r"(?=[A-Za-z0-9$])"
OPERATOR_FORMS: Dict[str, re.Pattern] = {
    "_TRO_ALL": re.compile(_REF_BEFORE + r"\.:\." + _REF_AFTER),
    "_TRO_LEADING": re.compile(_REF_BEFORE + r"\.:" + _REF_AFTER),
    "_TRO_TRAILING": re.compile(_REF_BEFORE + r":\." + _REF_AFTER),
}


def function_pattern(fn: str) -> re.Pattern:
    """Pattern matching a use of `fn`: its operator form, else a `FN(` call."""
    form = OPERATOR_FORMS.get(fn.upper())
    if form is not None:
        return form
    return re.compile(rf"(?<![A-Za-z0-9_]){re.escape(fn)}\s*\(", re.IGNORECASE)


def extract_calls(text: str) -> Set[str]:
    """Returns every uppercase name called with `(` in `text`, plus operator-form built-ins."""
    found = {m.group(1) for m in FUNCTION_CALL_RE.finditer(text)}
    found.update(fn for fn, pat in OPERATOR_FORMS.items() if pat.search(text))
    return found


def scan_dir(directory: Path, glob: str, recursive: bool = False) -> Dict[str, Set[Path]]:
    """Greps every file matching `glob` under `directory` for function calls.

    Returns a dict mapping function-name -> set of file paths that reference
    it. Scanning the entire file (rather than only `formula:` keys) keeps
    the regex robust against multi-line formulas, golden-JSON pretty-print
    line wrapping, and any other surrounding noise; over-counting coverage
    is preferred over missing coverage that exists.
    """
    out: Dict[str, Set[Path]] = {}
    if not directory.is_dir():
        return out
    iterator: Iterable[Path] = directory.rglob(glob) if recursive else directory.glob(glob)
    for path in sorted(iterator):
        text = path.read_text(encoding="utf-8", errors="replace")
        for fn in extract_calls(text):
            out.setdefault(fn, set()).add(path)
    return out


def scan_native_golden(directory: Path) -> Dict[str, Set[Path]]:
    """Scan only native target formula goldens; the external corpus is counted separately."""
    out: Dict[str, Set[Path]] = {}
    if not directory.is_dir():
        return out
    for path in sorted(directory.glob("*.json")):
        text = path.read_text(encoding="utf-8", errors="replace")
        for fn in extract_calls(text):
            out.setdefault(fn, set()).add(path)
    return out


# Argument shapes a function's cases should exercise, and those the ranking counts.
SHAPES = ("array", "error", "blank", "omitted")
RANKED_SHAPES = ("array", "error", "blank")
# Functions whose value depends on the environment or the clock, or that take no
# argument; an argument-shape case cannot be pinned for them.
SHAPE_EXEMPT = frozenset(
    "COPILOT CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE "
    "DETECTLANGUAGE IMAGE NOW PY RAND RANDARRAY RANDBETWEEN RTD STOCKHISTORY TODAY TRANSLATE WEBSERVICE "
    "FALSE TRUE NA PI".split()
)

_ERROR_LITERAL_RE = re.compile(
    r"#(?:N/A|VALUE!|REF!|DIV/0!|NUM!|NAME\?|NULL!|SPILL!|CALC!)|(?<![A-Za-z0-9_.])NA\(\)", re.I
)
_RANGE_RE = re.compile(
    r"\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+|(?<![A-Za-z0-9_])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![A-Za-z0-9_(])|\b\d+:\d+\b"
)
# A single-cell reference: not part of a longer name, not a call (`LOG10(`), not a range end.
_CELL_RE = re.compile(r"(?<![A-Za-z0-9_.:$])\$?([A-Z]{1,3})\$?(\d+)(?![A-Za-z0-9_(:])")
_STRING_RE = re.compile(r'"(?:[^"]|"")*"')


def call_arguments(formula: str, fn: str) -> List[List[str]]:
    """The argument lists of every call of `fn` in `formula`, split at top-level commas."""
    calls: List[List[str]] = []
    if fn.upper() in OPERATOR_FORMS:
        return calls
    for m in function_pattern(fn).finditer(formula):
        args, cur, depth, quoted = [], "", 1, False
        for ch in formula[m.end() :]:
            if ch == '"':
                quoted = not quoted
            elif not quoted and ch in "({":
                depth += 1
            elif not quoted and ch in ")}":
                depth -= 1
                if depth == 0:
                    break
            elif not quoted and depth == 1 and ch == ",":
                args.append(cur.strip())
                cur = ""
                continue
            cur += ch
        args.append(cur.strip())
        calls.append(args)
    return calls


def argument_shapes(args: Sequence[str], setup: Dict[str, object]) -> Set[str]:
    """Tags one call's arguments with the shapes in `SHAPES`."""
    cells = {str(k).upper().replace("$", ""): v for k, v in setup.items()}
    tags: Set[str] = set()
    if args != [""] and "" in args:
        tags.add("omitted")
    for arg in args:
        bare = _STRING_RE.sub('""', arg)
        if "{" in bare or _RANGE_RE.search(bare):
            tags.add("array")
        if _ERROR_LITERAL_RE.search(bare):
            tags.add("error")
        for col, row in _CELL_RE.findall(_RANGE_RE.sub("", bare)):
            value = cells.get(col + row)
            if value is None or value == "":
                tags.add("blank")
            elif isinstance(value, str) and (value.startswith("#") or _ERROR_LITERAL_RE.match(value.lstrip("="))):
                tags.add("error")
    return tags


def classify_shapes(cases_dir: Path, functions: Iterable[str]) -> Dict[str, Dict[str, object]]:
    """Per function: how many yaml case calls it has and the union of their argument shapes."""
    if str(REPO_ROOT) not in sys.path:  # run as a script: the repository root holds tools/yaml_load.py
        sys.path.insert(0, str(REPO_ROOT))
    from tools.yaml_load import load_yaml  # PyYAML is needed only here

    cases = []
    for path in sorted(cases_dir.glob("*.yaml")):
        doc = load_yaml(path.read_text(encoding="utf-8")) or {}
        for case in doc.get("cases") or []:
            formula = case.get("formula")
            if isinstance(formula, str):
                cases.append((formula, case.get("setup") or {}))
    out: Dict[str, Dict[str, object]] = {}
    for fn in functions:
        calls, shapes = 0, set()
        for formula, setup in cases:
            for args in call_arguments(formula, fn):
                calls += 1
                shapes |= argument_shapes(args, setup)
        out[fn] = {"calls": calls, "shapes": sorted(shapes)}
    return out


def shape_gaps(classified: Dict[str, Dict[str, object]]) -> List[Dict[str, object]]:
    """Functions that miss a ranked shape, most missing first, exempt and uncalled ones left out."""
    rows = []
    for fn, info in classified.items():
        if fn in SHAPE_EXEMPT or fn.startswith("_") or not info["calls"]:
            continue
        missing = [s for s in RANKED_SHAPES if s not in info["shapes"]]
        if missing:
            rows.append({"name": fn, "calls": info["calls"], "missing": missing})
    rows.sort(key=lambda r: (-len(r["missing"]), r["name"]))
    return rows


def build_rows(
    catalog: Set[str],
    implemented: Set[str],
    yaml_hits: Dict[str, Set[Path]],
    ironcalc_hits: Dict[str, Set[Path]],
    native_golden_hits: Dict[str, Set[Path]],
) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    for fn in sorted(catalog):
        impl = fn in implemented
        in_yaml = fn in yaml_hits
        in_ironcalc = fn in ironcalc_hits
        in_native_golden = fn in native_golden_hits
        any_oracle = in_yaml or in_ironcalc or in_native_golden
        rows.append(
            {
                "name": fn,
                "implemented": impl,
                "yaml": in_yaml,
                "ironcalc": in_ironcalc,
                "native_golden": in_native_golden,
                "any_oracle": any_oracle,
                "pilot_candidate": impl and not any_oracle,
            }
        )
    return rows


def print_table(rows: List[Dict[str, object]], title: str) -> None:
    print(f"{title}: {len(rows)}")
    print()
    header = f"{'name':32}  impl  yaml  ironcalc  native"
    print(header)
    print("-" * len(header))
    for r in rows:
        flag = lambda key: "Y" if r[key] else "."  # noqa: E731
        print(
            f"{str(r['name']):32}  {flag('implemented'):4}  "
            f"{flag('yaml'):4}  {flag('ironcalc'):8}  {flag('native_golden'):6}"
        )


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--all",
        action="store_true",
        help="Print the full per-function table (default: pilot candidates only).",
    )
    parser.add_argument(
        "--missing-impl",
        action="store_true",
        help="Filter for entries that are not implemented AND have no oracle coverage.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit machine-readable JSON instead of a text table.",
    )
    parser.add_argument(
        "--shapes",
        action="store_true",
        help="Rank implemented functions by the argument shapes (array, error, blank) no yaml case exercises.",
    )
    args = parser.parse_args(argv)

    _, catalog = catalog_status.load_catalog(catalog_status.CATALOG_PATH)
    implemented = catalog_status.scan_implemented(REPO_ROOT)

    yaml_hits = scan_dir(ORACLE_CASES_DIR, "*.yaml")
    ironcalc_hits = scan_dir(IRONCALC_GOLDEN_DIR, "*.json")
    native_golden_hits = scan_native_golden(ORACLE_GOLDEN_DIR)

    rows = build_rows(catalog, implemented, yaml_hits, ironcalc_hits, native_golden_hits)

    if args.shapes:
        gaps = shape_gaps(classify_shapes(ORACLE_CASES_DIR, sorted(implemented & catalog)))
        if args.json:
            json.dump({"shape_gaps": gaps}, sys.stdout, indent=2)
            sys.stdout.write("\n")
            return 0
        print(f"Functions missing an argument shape: {len(gaps)}")
        print()
        for r in gaps:
            print(f"{str(r['name']):32}  calls {r['calls']:4}  missing {', '.join(r['missing'])}")
        return 0

    if args.json:
        summary = {
            "catalog_size": len(catalog),
            "implemented": sum(1 for r in rows if r["implemented"]),
            "with_oracle": sum(1 for r in rows if r["any_oracle"]),
            "pilot_candidates": [r["name"] for r in rows if r["pilot_candidate"]],
            "missing_impl_no_oracle": [r["name"] for r in rows if not r["implemented"] and not r["any_oracle"]],
            "rows": rows,
        }
        json.dump(summary, sys.stdout, indent=2, default=list)
        sys.stdout.write("\n")
        return 0

    if args.all:
        print_table(rows, "Coverage gap (all functions)")
        return 0

    if args.missing_impl:
        filtered = [r for r in rows if not r["implemented"] and not r["any_oracle"]]
        print_table(filtered, "Functions with no implementation AND no oracle")
        return 0

    pilots = [r for r in rows if r["pilot_candidate"]]
    summary_line = (
        f"Catalog: {len(catalog)}, implemented: {sum(1 for r in rows if r['implemented'])}, "
        f"with-oracle: {sum(1 for r in rows if r['any_oracle'])}, "
        f"pilot-candidates: {len(pilots)}"
    )
    print(summary_line)
    print()
    print_table(pilots, "Pilot candidates (implemented, no oracle)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
