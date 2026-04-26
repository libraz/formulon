#!/usr/bin/env python3
"""Coverage-gap analyzer for the Formulon oracle test suites.

Cross-references the catalog at `tools/catalog/functions.txt` against three
notions of "is there a test for this function?":

  * implementation status   - re-used from `tools/catalog/status.py`
  * native Formulon yaml    - `tests/oracle/cases/*.yaml`
  * golden JSON corpora     - `tests/oracle/golden/**/*.json` (the imported
                              ironcalc fixtures live under
                              `tests/oracle/golden/ironcalc/` and are
                              counted separately from native goldens)

The default report lists *pilot candidates*: functions that are implemented
but have no oracle coverage in any of the three sources. Those are the safe
targets for the next round of `make oracle-gen` work — Mac Excel 365 (ja-JP)
acts as the authoritative oracle for fresh yaml suites we author in
`tests/oracle/cases/`.

Stdlib only. Works on CPython 3.8+.

Usage:
    tools/oracle/coverage_gap.py                 # pilot candidates table
    tools/oracle/coverage_gap.py --all           # full per-function table
    tools/oracle/coverage_gap.py --json          # machine-readable
    tools/oracle/coverage_gap.py --missing-impl  # functions with no impl AND no oracle
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
ORACLE_GOLDEN_DIR = REPO_ROOT / "tests" / "oracle" / "golden"
IRONCALC_GOLDEN_DIR = ORACLE_GOLDEN_DIR / "ironcalc"

sys.path.insert(0, str(CATALOG_DIR))
import status as catalog_status  # type: ignore  # noqa: E402

# Function-call detection. A name is uppercase ASCII, may contain digits,
# `_`, and `.` (for namespaced names like `FORECAST.ETS.SEASONALITY`), and
# is immediately followed by `(`. The look-behind rejects matches that
# would just be the tail of a longer identifier (`MyABS(` -> reject), but
# allows a leading `.` so OOXML-style namespace prefixes like
# `_xlfn._xlws.SORT(` still surface the call.
FUNCTION_CALL_RE = re.compile(r"(?<![A-Za-z0-9_])([A-Z][A-Z0-9_.]*)\s*\(")


def extract_calls(text: str) -> Set[str]:
    """Returns every uppercase name immediately followed by `(` in `text`."""
    return {m.group(1) for m in FUNCTION_CALL_RE.finditer(text)}


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
    """Scans `golden/*.json` only — recursing would re-count `golden/ironcalc/`
    JSONs which are surfaced separately as the ironcalc bucket."""
    out: Dict[str, Set[Path]] = {}
    if not directory.is_dir():
        return out
    for path in sorted(directory.glob("*.json")):
        text = path.read_text(encoding="utf-8", errors="replace")
        for fn in extract_calls(text):
            out.setdefault(fn, set()).add(path)
    return out


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
    args = parser.parse_args(argv)

    _, catalog = catalog_status.load_catalog(catalog_status.CATALOG_PATH)
    implemented = catalog_status.scan_implemented(REPO_ROOT)

    yaml_hits = scan_dir(ORACLE_CASES_DIR, "*.yaml")
    ironcalc_hits = scan_dir(IRONCALC_GOLDEN_DIR, "*.json")
    native_golden_hits = scan_native_golden(ORACLE_GOLDEN_DIR)

    rows = build_rows(catalog, implemented, yaml_hits, ironcalc_hits, native_golden_hits)

    if args.json:
        summary = {
            "catalog_size": len(catalog),
            "implemented": sum(1 for r in rows if r["implemented"]),
            "with_oracle": sum(1 for r in rows if r["any_oracle"]),
            "pilot_candidates": [r["name"] for r in rows if r["pilot_candidate"]],
            "missing_impl_no_oracle": [
                r["name"] for r in rows if not r["implemented"] and not r["any_oracle"]
            ],
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
