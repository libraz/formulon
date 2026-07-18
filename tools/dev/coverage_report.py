#!/usr/bin/env python3
"""Coverage report (local diagnostic).

Parses an lcov ``.info`` tracefile, slices the per-file line-coverage
records into the named "areas" defined below, prints a fixed-width
table plus the list of zero-coverage files (the actually useful
signal), and -- by default -- exits 0 regardless of whether any area
hit its target.

This is *not* a CI gate. Coverage is too slow to gate development on,
and a fixed numeric threshold is a known anti-pattern (it incentivises
"call once with no assertion" tests and breaks under benign refactors).
The targets below are aspirational; the table prints them so you can
see how far you are, but reaching them is not enforced. Pass
``--strict`` to opt into a non-zero exit when an area is below target
(used only for ad-hoc local enforcement, never in CI).

================  =============================================  =========
Area              Source roots                                   Target
================  =============================================  =========
util/value/eval   ``src/utils/**``, ``src/value.{h,cpp}``,            95%
                  ``src/eval/**`` excluding the function-family
                  directories listed under ``functions``
functions         ``src/eval/builtins/**``,                            98%
                  ``src/eval/lookups/**``,
                  ``src/eval/conditional_aggregates.{h,cpp}``,
                  ``src/eval/special_forms_lazy.{h,cpp}``,
                  ``src/eval/criteria.{h,cpp}``,
                  ``src/eval/text_ops.{h,cpp}``,
                  ``src/eval/date_time.{h,cpp}``,
                  ``src/eval/coerce.{h,cpp}``,
                  ``src/eval/utf8_length.{h,cpp}``,
                  ``src/eval/range_args.{h,cpp}``,
                  ``src/eval/structured_ref.{h,cpp}``
io                ``src/io/**``                                       90%
================  =============================================  =========

Anything else (parser, c_api, cli, wasm, cf, pivot, ...) is reported
under "(informational)" with no target attached. Files that match
no area at all are silently skipped.

Order matters: a file matching multiple areas is assigned to the
**first** matching area in ``AREAS``. ``functions`` is therefore
listed before ``util/value/eval`` so that ``src/eval/builtins/*.cpp``
slices correctly into ``functions``.

Usage::

    coverage_report.py <coverage.info> [--strict] [--json]

Exit codes:

* 0 -- report rendered (default; targets are not enforced)
* 1 -- ``--strict`` was set and at least one gated area was below target
* 2 -- argument error / missing or unparseable tracefile
"""

from __future__ import annotations

import argparse
import fnmatch
import json
import os
import sys
from typing import Iterable, List, Optional, Tuple

AREAS: List[Tuple[str, List[str], Optional[float]]] = [
    (
        "functions",
        [
            "src/eval/builtins/*",
            "src/eval/builtins/**/*",
            "src/eval/lookups/*",
            "src/eval/lookups/**/*",
            "src/eval/conditional_aggregates.h",
            "src/eval/conditional_aggregates.cpp",
            "src/eval/special_forms_lazy.h",
            "src/eval/special_forms_lazy.cpp",
            "src/eval/criteria.h",
            "src/eval/criteria.cpp",
            "src/eval/text_ops.h",
            "src/eval/text_ops.cpp",
            "src/eval/date_time.h",
            "src/eval/date_time.cpp",
            "src/eval/coerce.h",
            "src/eval/coerce.cpp",
            "src/eval/utf8_length.h",
            "src/eval/utf8_length.cpp",
            "src/eval/range_args.h",
            "src/eval/range_args.cpp",
            "src/eval/structured_ref.h",
            "src/eval/structured_ref.cpp",
        ],
        98.0,
    ),
    (
        "util/value/eval",
        [
            "src/utils/*",
            "src/utils/**/*",
            "src/value.h",
            "src/value.cpp",
            "src/eval/*",
            "src/eval/**/*",
        ],
        95.0,
    ),
    (
        "io",
        [
            "src/io/*",
            "src/io/**/*",
        ],
        90.0,
    ),
    ("parser", ["src/parser/*", "src/parser/**/*"], None),
    ("c_api", ["src/c_api/*", "src/c_api/**/*"], None),
    ("cli", ["src/cli/*", "src/cli/**/*"], None),
    ("wasm", ["src/wasm/*", "src/wasm/**/*"], None),
    ("cf", ["src/cf/*", "src/cf/**/*"], None),
    ("pivot", ["src/pivot/*", "src/pivot/**/*"], None),
]


def _normalize_path(abs_path: str) -> str:
    """Return the path tail starting at the first ``src/`` segment."""
    norm = os.path.normpath(abs_path).replace("\\", "/")
    idx = norm.find("/src/")
    if idx >= 0:
        return norm[idx + 1 :]
    if norm.startswith("src/"):
        return norm
    return norm


def parse_tracefile(path: str) -> List[Tuple[str, int, int]]:
    """Parse an lcov ``.info`` file.

    Returns a list of ``(normalized_path, lines_found, lines_hit)``
    tuples, one per ``end_of_record`` block. Records missing ``LF`` /
    ``LH`` are treated as ``(0, 0)``.
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(f"tracefile not found: {path}")

    records: List[Tuple[str, int, int]] = []
    cur_sf: Optional[str] = None
    cur_lf = 0
    cur_lh = 0

    with open(path, "r", encoding="utf-8", errors="replace") as fh:
        for raw in fh:
            line = raw.strip()
            if not line:
                continue
            if line.startswith("SF:"):
                cur_sf = line[3:]
                cur_lf = 0
                cur_lh = 0
            elif line.startswith("LF:"):
                try:
                    cur_lf = int(line[3:])
                except ValueError:
                    cur_lf = 0
            elif line.startswith("LH:"):
                try:
                    cur_lh = int(line[3:])
                except ValueError:
                    cur_lh = 0
            elif line == "end_of_record":
                if cur_sf is not None:
                    records.append((_normalize_path(cur_sf), cur_lf, cur_lh))
                cur_sf = None
                cur_lf = 0
                cur_lh = 0

    return records


def _match_area(path_tail: str) -> Optional[str]:
    for name, patterns, _target in AREAS:
        for pat in patterns:
            if fnmatch.fnmatch(path_tail, pat):
                return name
    return None


def slice_records(
    records: Iterable[Tuple[str, int, int]],
) -> "dict[str, Tuple[int, int]]":
    totals: "dict[str, Tuple[int, int]]" = {}
    for tail, lf, lh in records:
        area = _match_area(tail)
        if area is None:
            continue
        prev_lf, prev_lh = totals.get(area, (0, 0))
        totals[area] = (prev_lf + lf, prev_lh + lh)
    return totals


def zero_coverage_files(
    records: Iterable[Tuple[str, int, int]],
) -> List[Tuple[str, int]]:
    """Return ``(path, lines_found)`` for every file with LH == 0 and LF > 0.

    These are the genuinely useful signal: source files that no test
    in the fast suite ever touches. Sorted by lines_found descending so
    the largest untested files surface first.
    """
    out: List[Tuple[str, int]] = []
    for tail, lf, lh in records:
        if lf > 0 and lh == 0 and _match_area(tail) is not None:
            out.append((tail, lf))
    out.sort(key=lambda r: -r[1])
    return out


def _pct(hit: int, total: int) -> float:
    if total <= 0:
        return 0.0
    return (hit / total) * 100.0


def _fmt_status(pct: float, target: Optional[float]) -> str:
    if target is None:
        return "-"
    return "ok" if pct + 1e-9 >= target else "below"


def render_table(
    totals: "dict[str, Tuple[int, int]]",
    zeros: List[Tuple[str, int]],
) -> Tuple[str, bool]:
    """Render the report.

    Returns ``(rendered_text, any_below_target)``.
    """
    header = f"{'Area':<17}  {'Lines hit / total':<22}  {'Coverage':<10}  {'Target':<10}  Status"
    sep = "-" * 17 + "  " + "-" * 22 + "  " + "-" * 10 + "  " + "-" * 10 + "  " + "-" * 8

    lines: List[str] = [
        "Coverage report (informational; --strict opts into enforcement)",
        "",
        header,
        sep,
    ]

    gated_lf_total = 0
    gated_lh_total = 0
    any_below = False

    for name, _patterns, target in AREAS:
        if target is None:
            continue
        lf, lh = totals.get(name, (0, 0))
        pct = _pct(lh, lf)
        status = _fmt_status(pct, target)
        if status == "below":
            any_below = True
        gated_lf_total += lf
        gated_lh_total += lh
        lines.append(f"{name:<17}  {lh:>9} / {lf:<10}  {pct:>8.2f}%  {target:>8.2f}%  {status}")

    info_rows: List[str] = []
    saw_informational = False
    for name, _patterns, target in AREAS:
        if target is not None:
            continue
        lf, lh = totals.get(name, (0, 0))
        if lf == 0:
            continue
        pct = _pct(lh, lf)
        info_rows.append(f"{name:<17}  {lh:>9} / {lf:<10}  {pct:>8.2f}%  {'-':>9}   {'-':<8}")
        saw_informational = True

    if saw_informational:
        lines.append("(informational, no target)")
        lines.extend(info_rows)

    lines.append(sep)
    total_pct = _pct(gated_lh_total, gated_lf_total)
    lines.append(f"{'TOTAL (gated)':<17}  {gated_lh_total:>9} / {gated_lf_total:<10}  {total_pct:>8.2f}%")

    if zeros:
        lines.append("")
        lines.append(f"Zero-coverage files ({len(zeros)} total, no fast-suite test touches them):")
        # Cap the printed list to keep CI logs sane; full data is in
        # the json sidecar.
        for path, lf in zeros[:30]:
            lines.append(f"  {lf:>5} lines  {path}")
        if len(zeros) > 30:
            lines.append(f"  ... and {len(zeros) - 30} more (see coverage.json)")

    return "\n".join(lines) + "\n", any_below


def render_json(
    totals: "dict[str, Tuple[int, int]]",
    zeros: List[Tuple[str, int]],
) -> str:
    payload = {
        "areas": [],
        "gated_total": {"lines_found": 0, "lines_hit": 0, "coverage_pct": 0.0},
        "zero_coverage_files": [{"path": path, "lines_found": lf} for path, lf in zeros],
    }
    gated_lf = 0
    gated_lh = 0
    for name, _patterns, target in AREAS:
        lf, lh = totals.get(name, (0, 0))
        pct = _pct(lh, lf)
        gated = target is not None
        if gated:
            gated_lf += lf
            gated_lh += lh
        payload["areas"].append(
            {
                "name": name,
                "lines_found": lf,
                "lines_hit": lh,
                "coverage_pct": round(pct, 4),
                "target_pct": target,
                "gated": gated,
                "at_or_above_target": (pct + 1e-9 >= target) if gated else None,
            }
        )
    payload["gated_total"] = {
        "lines_found": gated_lf,
        "lines_hit": gated_lh,
        "coverage_pct": round(_pct(gated_lh, gated_lf), 4),
    }
    return json.dumps(payload, indent=2)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description="Slice an lcov tracefile into Formulon coverage areas "
        "and print a per-area report. Local diagnostic only -- not a "
        "CI gate. Pass --strict to opt into a non-zero exit when an "
        "area is below its CLAUDE.md target.",
    )
    parser.add_argument(
        "tracefile",
        help="Path to the lcov .info tracefile (post-filter).",
    )
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit 1 when any gated area is below its target. "
        "Off by default; coverage targets are aspirational, not gates.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit per-area numbers + zero-coverage list as JSON instead of the text report.",
    )
    args = parser.parse_args(argv)

    try:
        records = parse_tracefile(args.tracefile)
    except FileNotFoundError as exc:
        print(f"coverage_report: {exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"coverage_report: failed to read {args.tracefile}: {exc}", file=sys.stderr)
        return 2

    if not records:
        print(
            f"coverage_report: tracefile {args.tracefile} contained no end_of_record blocks; nothing to report",
            file=sys.stderr,
        )
        return 2

    totals = slice_records(records)
    zeros = zero_coverage_files(records)

    if args.json:
        sys.stdout.write(render_json(totals, zeros))
        sys.stdout.write("\n")
        return 0

    text, any_below = render_table(totals, zeros)
    sys.stdout.write(text)

    if any_below and args.strict:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
