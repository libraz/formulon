#!/usr/bin/env python3
"""Generates conditional-formatting golden JSON from CF cases by driving
Mac Excel via xlwings + openpyxl.

Usage:
    python3 tools/oracle/cf_oracle_gen.py [--suite NAME ...]
                                          [--cases-dir PATH]
                                          [--golden-dir PATH]

For each `<cases-dir>/<suite>.case.json`, the generator builds one xlsx
per case via openpyxl (cell data + CF rules), opens it under a hidden
Excel.app instance, recalculates, then captures
``DisplayFormat.Interior`` per cell in the case range. The captured fill
colours are matched back to the rule descriptors to produce a
``<suite>.golden.json`` file shaped like the hand-authored
self-baselines under ``tests/oracle/golden_cf/``.

Supported rule types (matches the smoke suite):

  - ``CellIs``, ``Top10``, ``AboveAverage``, ``ContainsText`` ->
    ``DifferentialFormat`` match. The rule's ``dxf`` fill is set to a
    deterministic fingerprint colour so the generator can map the
    captured ``DisplayFormat`` colour back to the originating rule's
    declared ``dxf_id``. Adding another DifferentialFormat-bearing rule
    type only requires extending ``_build_workbook`` to emit it via
    openpyxl; the capture path is shared.
  - ``ColorScale`` -> ``ColorScale`` match. The captured fill is the
    colour Excel resolved by interpolating between stops.

Not yet captured (intentional gaps):

  - ``DataBar``, ``IconSet`` cannot be captured through
    ``DisplayFormat`` alone (Excel does not expose the rendered bar
    length or the icon bucket via the public AppleScript surface), so
    they will likely stay deferred even after the other rule types
    land.
  - ``NotContainsText``, ``BeginsWith``, ``EndsWith``,
    ``ContainsBlanks`` / ``NotContainsBlanks``, ``ContainsErrors`` /
    ``NotContainsErrors``, ``TimePeriod``, ``DuplicateValues``,
    ``UniqueValues``, ``Expression`` — all DifferentialFormat-bearing.
    They reuse the same dxf-fingerprint trick; they are deferred until
    the smoke suite gains cases that exercise them.

macOS + Excel 365 only. The generator refuses to start on any other
platform.
"""

from __future__ import annotations

import argparse
import json
import platform
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_CASES_DIR = REPO_ROOT / "tests" / "oracle" / "cases_cf"
DEFAULT_GOLDEN_DIR = REPO_ROOT / "tests" / "oracle" / "golden_cf"


def _ensure_darwin() -> None:
    if platform.system() != "Darwin":
        raise RuntimeError(
            "cf_oracle_gen is macOS-only (xlwings drives Excel.app). "
            "Current platform: " + platform.system()
        )


def _addr(row: int, col: int) -> str:
    """Returns A1-style address from 0-based row/col."""
    name = ""
    c = col + 1
    while c > 0:
        c, r = divmod(c - 1, 26)
        name = chr(ord("A") + r) + name
    return f"{name}{row + 1}"


def _hex_for_dxf_id(idx: int) -> Tuple[int, int, int]:
    """Deterministic RGB fingerprint for ``dxf_id``.

    The three channels are derived via prime multipliers so that adjacent
    ids occupy distant points in colour space and the inverse mapping
    (captured colour -> dxf_id) stays unambiguous within +/- 2 channel
    tolerance.
    """
    r = (idx * 53 + 17) & 0xFF
    g = (idx * 89 + 41) & 0xFF
    b = (idx * 137 + 67) & 0xFF
    return (r, g, b)


def _hex6(rgb: Tuple[int, int, int]) -> str:
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def _channel_close(actual: List[int], expected: Tuple[int, int, int], tol: int = 2) -> bool:
    if len(actual) < 3:
        return False
    return all(abs(int(actual[i]) - expected[i]) <= tol for i in range(3))


def _build_workbook(case: Dict[str, Any], path: Path) -> List[Dict[str, Any]]:
    """Builds the .xlsx for a single case and returns rule descriptors.

    A descriptor records the inputs the capture stage needs to match a
    cell's resolved fill against an originating rule. Each descriptor
    carries the originating rule's ``priority`` and ``type``; for
    ``CellIs`` rules it also carries the deterministic dxf fingerprint
    colour (so the captured fill can be reverse-mapped to the rule's
    declared ``dxf_id``).
    """
    import openpyxl
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, Rule
    from openpyxl.styles import PatternFill
    from openpyxl.styles.differential import DifferentialStyle

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for entry in case["sheet"]:
        addr = _addr(entry["row"], entry["col"])
        kind = entry["kind"]
        v = entry.get("value")
        if kind == "number":
            ws[addr] = float(v)
        elif kind == "bool":
            ws[addr] = bool(v)
        elif kind == "text":
            ws[addr] = str(v)
        elif kind == "blank":
            pass
        elif kind == "error":
            triggers = {
                "#DIV/0!": "=1/0",
                "#NAME?": "=NONEXISTENT_FUNC()",
                "#VALUE!": '=VALUE("x")',
                "#NUM!": "=SQRT(-1)",
                "#N/A": "=NA()",
                "#REF!": "=OFFSET(A1,-1,-1)",
            }
            ws[addr] = triggers.get(v or "#VALUE!", '=VALUE("x")')
        else:
            raise NotImplementedError(f"sheet cell kind not supported: {kind}")

    descriptors: List[Dict[str, Any]] = []
    for block in case["cf_blocks"]:
        sqref_str = " ".join(
            f"{_addr(r['first_row'], r['first_col'])}:{_addr(r['last_row'], r['last_col'])}"
            for r in block["sqref"]
        )
        sqref_blocks = [
            {
                "first_row": int(r["first_row"]),
                "first_col": int(r["first_col"]),
                "last_row": int(r["last_row"]),
                "last_col": int(r["last_col"]),
            }
            for r in block["sqref"]
        ]
        for rule_idx, rule in enumerate(block["rules"]):
            rtype = rule["type"]
            priority = int(rule["priority"])
            if rtype == "CellIs":
                op_map = {
                    "LessThan": "lessThan",
                    "LessThanOrEqual": "lessThanOrEqual",
                    "Equal": "equal",
                    "NotEqual": "notEqual",
                    "GreaterThan": "greaterThan",
                    "GreaterThanOrEqual": "greaterThanOrEqual",
                    "Between": "between",
                    "NotBetween": "notBetween",
                }
                op = op_map[rule["operator"]]
                dxf_id = int(rule.get("dxf_id", rule_idx))
                fingerprint = _hex_for_dxf_id(dxf_id)
                fill_hex = "FF" + _hex6(fingerprint)
                fill = PatternFill(
                    start_color=fill_hex,
                    end_color=fill_hex,
                    fill_type="solid",
                )
                dxf = DifferentialStyle(fill=fill)
                formulas = [rule["formula1"]]
                if "formula2" in rule:
                    formulas.append(rule["formula2"])
                cir = CellIsRule(operator=op, formula=formulas, stopIfTrue=False)
                cir.dxf = dxf
                cir.priority = priority
                ws.conditional_formatting.add(sqref_str, cir)
                descriptors.append(
                    {
                        "type": rtype,
                        "priority": priority,
                        "sqref": sqref_blocks,
                        "dxf_id": dxf_id,
                        "dxf_fingerprint": fingerprint,
                    }
                )
            elif rtype in ("Top10", "AboveAverage", "ContainsText"):
                # All three are DifferentialFormat-bearing rules; the
                # capture path uses the same dxf-fingerprint reverse
                # mapping as `CellIs`. Only the openpyxl `Rule` payload
                # differs.
                dxf_id = int(rule.get("dxf_id", rule_idx))
                fingerprint = _hex_for_dxf_id(dxf_id)
                fill_hex = "FF" + _hex6(fingerprint)
                fill = PatternFill(
                    start_color=fill_hex,
                    end_color=fill_hex,
                    fill_type="solid",
                )
                dxf = DifferentialStyle(fill=fill)
                if rtype == "Top10":
                    rank = int(rule.get("rank", 10))
                    percent = bool(rule.get("percent", False))
                    bottom = bool(rule.get("bottom", False))
                    opx_rule = Rule(
                        type="top10",
                        rank=rank,
                        percent=percent,
                        bottom=bottom,
                        dxf=dxf,
                        priority=priority,
                        stopIfTrue=False,
                    )
                elif rtype == "AboveAverage":
                    above_average = bool(rule.get("above_average", True))
                    equal_average = bool(rule.get("equal_average", False))
                    std_dev = rule.get("std_dev")
                    rule_kwargs: Dict[str, Any] = {
                        "type": "aboveAverage",
                        "aboveAverage": above_average,
                        "equalAverage": equal_average,
                        "dxf": dxf,
                        "priority": priority,
                        "stopIfTrue": False,
                    }
                    if std_dev is not None:
                        rule_kwargs["stdDev"] = int(std_dev)
                    opx_rule = Rule(**rule_kwargs)
                else:  # ContainsText
                    text = rule.get("text", "")
                    if not text:
                        raise ValueError(
                            "ContainsText rule requires a `text` field"
                        )
                    # Excel synthesises a SEARCH() formula at write time;
                    # openpyxl wants us to provide it explicitly. Use the
                    # case range's first cell as the row anchor — Excel
                    # rewrites the row on apply just like a relative ref.
                    first_sqref = block["sqref"][0]
                    anchor = _addr(first_sqref["first_row"], first_sqref["first_col"])
                    formula = [f'NOT(ISERROR(SEARCH("{text}",{anchor})))']
                    opx_rule = Rule(
                        type="containsText",
                        operator="containsText",
                        text=text,
                        formula=formula,
                        dxf=dxf,
                        priority=priority,
                        stopIfTrue=False,
                    )
                ws.conditional_formatting.add(sqref_str, opx_rule)
                descriptors.append(
                    {
                        "type": rtype,
                        "priority": priority,
                        "sqref": sqref_blocks,
                        "dxf_id": dxf_id,
                        "dxf_fingerprint": fingerprint,
                    }
                )
            elif rtype == "ColorScale":
                spec = rule["color_scale"]
                thresholds = spec["thresholds"]
                colors = spec["colors"]
                type_map = {
                    "Min": "min",
                    "Max": "max",
                    "Number": "num",
                    "Percent": "percent",
                    "Percentile": "percentile",
                    "Formula": "formula",
                }
                ths_types = [type_map[t["type"]] for t in thresholds]
                ths_vals: List[Optional[Any]] = [t.get("value") for t in thresholds]
                colors_hex = [
                    f"FF{c['r']:02X}{c['g']:02X}{c['b']:02X}" for c in colors
                ]
                if len(thresholds) == 3:
                    csr = ColorScaleRule(
                        start_type=ths_types[0],
                        start_value=ths_vals[0],
                        start_color=colors_hex[0],
                        mid_type=ths_types[1],
                        mid_value=ths_vals[1],
                        mid_color=colors_hex[1],
                        end_type=ths_types[2],
                        end_value=ths_vals[2],
                        end_color=colors_hex[2],
                    )
                elif len(thresholds) == 2:
                    csr = ColorScaleRule(
                        start_type=ths_types[0],
                        start_value=ths_vals[0],
                        start_color=colors_hex[0],
                        end_type=ths_types[1],
                        end_value=ths_vals[1],
                        end_color=colors_hex[1],
                    )
                else:
                    raise NotImplementedError(
                        f"ColorScale with {len(thresholds)} stops not supported"
                    )
                csr.priority = priority
                ws.conditional_formatting.add(sqref_str, csr)
                descriptors.append(
                    {
                        "type": rtype,
                        "priority": priority,
                        "sqref": sqref_blocks,
                    }
                )
            else:
                raise NotImplementedError(
                    f"CF rule type not yet supported by oracle gen: {rtype}"
                )

    wb.save(path)
    return descriptors


def _cell_in_block(row: int, col: int, sqref: List[Dict[str, int]]) -> bool:
    for r in sqref:
        if (
            r["first_row"] <= row <= r["last_row"]
            and r["first_col"] <= col <= r["last_col"]
        ):
            return True
    return False


def _resolve_color_list(value: Any) -> Optional[List[int]]:
    """Resolves an appscript-color reference to ``[r, g, b]`` or returns
    ``None`` if the reference cannot be coerced.
    """
    if value is None:
        return None
    if callable(value):
        try:
            value = value()
        except Exception:
            return None
    if isinstance(value, list) and len(value) >= 3:
        try:
            return [int(value[0]), int(value[1]), int(value[2])]
        except (TypeError, ValueError):
            return None
    return None


def _resolve_pattern(value: Any) -> Optional[str]:
    if value is None:
        return None
    if callable(value):
        try:
            value = value()
        except Exception:
            return None
    return str(value)


def _classify_cell(
    cell, row: int, col: int, descriptors: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """Reads ``DisplayFormat.Interior`` for ``cell`` and emits one or
    more ``CFMatch``-shaped dicts.

    Logic:
      1. If pattern is ``pattern_none`` -> no CF rule fired -> return [].
      2. Walk descriptors in priority order. A ``CellIs`` descriptor whose
         declared sqref covers ``(row, col)`` and whose dxf fingerprint
         matches the captured colour (within +/- 2 per channel) emits
         a ``DifferentialFormat`` match and short-circuits — the dxf
         path is exclusive.
      3. Otherwise, the first ``ColorScale`` descriptor whose sqref
         covers the cell emits a ``ColorScale`` match with the captured
         colour. Multiple overlapping ColorScale rules are not handled;
         this matches the smoke suite's authoring rule.
    """
    try:
        interior = cell.api.display_format.interior_object
    except Exception:
        return []

    pattern = _resolve_pattern(interior.pattern)
    if pattern is None or pattern.endswith("pattern_none"):
        return []

    color_list = _resolve_color_list(interior.color)
    if color_list is None:
        return []

    sorted_descriptors = sorted(descriptors, key=lambda d: d["priority"])
    for desc in sorted_descriptors:
        if not _cell_in_block(row, col, desc["sqref"]):
            continue
        if "dxf_fingerprint" in desc:
            # All DifferentialFormat-bearing rule types share the same
            # capture: if the cell's resolved fill matches the rule's
            # dxf fingerprint, the rule fired and we emit a single
            # DifferentialFormat match referencing the declared dxf_id.
            if _channel_close(color_list, desc["dxf_fingerprint"], tol=2):
                return [
                    {
                        "kind": "DifferentialFormat",
                        "priority": desc["priority"],
                        "dxf_id": desc["dxf_id"],
                    }
                ]
        elif desc["type"] == "ColorScale":
            return [
                {
                    "kind": "ColorScale",
                    "priority": desc["priority"],
                    "color": {
                        "r": int(color_list[0]),
                        "g": int(color_list[1]),
                        "b": int(color_list[2]),
                        "a": 255,
                    },
                }
            ]
    return []


def _capture_via_excel(
    xlsx: Path, case: Dict[str, Any], descriptors: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """Opens ``xlsx`` in a hidden Excel.app, recalculates, then walks
    every cell in ``case['range']`` and decodes the resolved
    ``DisplayFormat`` against ``descriptors``.
    """
    import xlwings as xw  # type: ignore

    rng_spec = case["range"]
    r1 = int(rng_spec["first_row"])
    c1 = int(rng_spec["first_col"])
    r2 = int(rng_spec["last_row"])
    c2 = int(rng_spec["last_col"])

    app = xw.App(visible=False, add_book=False)
    app.calculation = "manual"
    app.screen_updating = False
    app.display_alerts = False
    try:
        wb = app.books.open(str(xlsx))
        try:
            app.calculate()
            sht = wb.sheets[0]
            cells: List[Dict[str, Any]] = []
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    addr = _addr(r, c)
                    cell = sht.range(addr)
                    matches = _classify_cell(cell, r, c, descriptors)
                    if matches:
                        cells.append({"row": r, "col": c, "matches": matches})
            return cells
        finally:
            try:
                wb.close()
            except Exception:
                pass
    finally:
        try:
            app.quit()
        except Exception:
            pass


def _process_suite(case_path: Path, golden_path: Path) -> None:
    doc = json.loads(case_path.read_text(encoding="utf-8"))
    out_cases: List[Dict[str, Any]] = []
    with tempfile.TemporaryDirectory(prefix="formulon-cf-oracle-") as tmp:
        tmpdir = Path(tmp)
        for case in doc["cases"]:
            xlsx = tmpdir / f"{case['id']}.xlsx"
            descriptors = _build_workbook(case, xlsx)
            cells = _capture_via_excel(xlsx, case, descriptors)
            out_cases.append({"id": case["id"], "cells": cells})

    out_doc = {
        "name": doc["name"],
        "description": (
            "Excel-actual golden captured by tools/oracle/cf_oracle_gen.py "
            "(macOS + Excel 365, ja-JP)."
        ),
        "cases": out_cases,
    }
    golden_path.parent.mkdir(parents=True, exist_ok=True)
    golden_path.write_text(
        json.dumps(out_doc, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def main() -> int:
    _ensure_darwin()

    try:
        import xlwings  # noqa: F401
        import openpyxl  # noqa: F401
    except ImportError as exc:
        print(
            f"[cf-oracle-gen] missing dependency: {exc}; run `make oracle-setup`",
            file=sys.stderr,
        )
        return 2

    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--suite",
        action="append",
        default=None,
        help="suite name (without `.case.json` suffix); repeat for several",
    )
    parser.add_argument(
        "--cases-dir",
        type=Path,
        default=DEFAULT_CASES_DIR,
        help=f"directory of CF case JSON files (default: {DEFAULT_CASES_DIR})",
    )
    parser.add_argument(
        "--golden-dir",
        type=Path,
        default=DEFAULT_GOLDEN_DIR,
        help=f"directory to write golden JSON (default: {DEFAULT_GOLDEN_DIR})",
    )
    args = parser.parse_args()

    cases_dir: Path = args.cases_dir
    golden_dir: Path = args.golden_dir

    files = sorted(cases_dir.glob("*.case.json"))
    if args.suite:
        wanted = set(args.suite)
        files = [f for f in files if f.name.removesuffix(".case.json") in wanted]
    if not files:
        print(f"[cf-oracle-gen] no .case.json files under {cases_dir}", file=sys.stderr)
        return 0

    for case_path in files:
        suite = case_path.name.removesuffix(".case.json")
        golden_path = golden_dir / f"{suite}.golden.json"
        print(f"[cf-oracle-gen] {case_path.name} -> {golden_path.name}")
        _process_suite(case_path, golden_path)
        print(f"[cf-oracle-gen]   wrote {golden_path}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
