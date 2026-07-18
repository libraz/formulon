#!/usr/bin/env python3
"""CF case + golden JSON validator.

The conditional-formatting (CF) oracle pipeline keeps three on-disk
artefacts in sync per suite:

  * tests/oracle/cases_cf/<suite>.yaml         — human-authored source
  * tests/oracle/cases_cf/<suite>.case.json    — JSON mirror consumed by
                                                 the future Mac driver
  * tests/oracle/golden_cf/<suite>.golden.json — committed Formulon-
                                                 self-baseline (until
                                                 the Mac driver lands)

This script validates the *case* JSON and the *golden* JSON: structural
shape, required fields, enum values, and cross-file consistency (every
case id in the case file must appear in the golden, and vice versa).

The C++ smoke test (`tests/oracle/cf_oracle_test.cpp`) does *not* read
either file: it pins the same scenarios as hardcoded gtest fixtures so
the gtest binary needs neither YAML nor JSON. This validator is the
guard rail that flags drift between the on-disk files and any future
machine-generated artefacts.

Usage:
    python3 tools/oracle/cf_case_schema.py \\
        tests/oracle/cases_cf/cf_smoke.case.json \\
        tests/oracle/golden_cf/cf_smoke.golden.json

Exits 0 on success, 1 on validation failure (with a path-and-field
location printed to stderr).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, List, Set, Tuple

# Mirrors `enum class formulon::cf::RuleType` in src/cf/cf_types.h.
_RULE_TYPES: Set[str] = {
    "Expression",
    "CellIs",
    "ColorScale",
    "DataBar",
    "IconSet",
    "Top10",
    "AboveAverage",
    "ContainsText",
    "NotContainsText",
    "BeginsWith",
    "EndsWith",
    "ContainsBlanks",
    "NotContainsBlanks",
    "ContainsErrors",
    "NotContainsErrors",
    "TimePeriod",
    "DuplicateValues",
    "UniqueValues",
}

# Mirrors `enum class formulon::cf::CellIsOperator`.
_CELL_IS_OPERATORS: Set[str] = {
    "LessThan",
    "LessThanOrEqual",
    "Equal",
    "NotEqual",
    "GreaterThanOrEqual",
    "GreaterThan",
    "Between",
    "NotBetween",
}

# Mirrors `enum class formulon::cf::CFMatchKind`.
_MATCH_KINDS: Set[str] = {
    "DifferentialFormat",
    "ColorScale",
    "DataBar",
    "IconSet",
}

# Recognised cell-value kinds for the case's `sheet` array. Matches the
# subset of `enum class ValueKind` that the CF smoke harness exercises.
_VALUE_KINDS: Set[str] = {"blank", "number", "bool", "text", "error"}


class ValidationError(Exception):
    """Validation failure with a path-and-field hint."""


def _at(*parts: Any) -> str:
    """Render a JSON-pointer-ish location for an error message."""
    return "/" + "/".join(str(p) for p in parts) if parts else "/"


def _require_keys(obj: Any, keys: List[str], where: str) -> None:
    if not isinstance(obj, dict):
        raise ValidationError(f"{where}: expected mapping, got {type(obj).__name__}")
    missing = [k for k in keys if k not in obj]
    if missing:
        raise ValidationError(f"{where}: missing required keys {missing}")


def _validate_color(obj: Any, where: str) -> None:
    _require_keys(obj, ["r", "g", "b", "a"], where)
    for ch in ("r", "g", "b", "a"):
        v = obj[ch]
        if not isinstance(v, int) or v < 0 or v > 255:
            raise ValidationError(f"{where}/{ch}: expected int in [0, 255], got {v!r}")


def _validate_range(obj: Any, where: str) -> None:
    _require_keys(obj, ["first_row", "first_col", "last_row", "last_col"], where)
    for k in ("first_row", "first_col", "last_row", "last_col"):
        if not isinstance(obj[k], int) or obj[k] < 0:
            raise ValidationError(f"{where}/{k}: expected non-negative int, got {obj[k]!r}")
    if obj["last_row"] < obj["first_row"] or obj["last_col"] < obj["first_col"]:
        raise ValidationError(f"{where}: last_row/last_col must be >= first_row/first_col")


def _validate_rule(rule: Any, where: str) -> None:
    _require_keys(rule, ["type", "priority"], where)
    rtype = rule["type"]
    if rtype not in _RULE_TYPES:
        raise ValidationError(f"{where}/type: unknown rule type {rtype!r}; expected one of {sorted(_RULE_TYPES)}")
    if not isinstance(rule["priority"], int):
        raise ValidationError(f"{where}/priority: expected int, got {type(rule['priority']).__name__}")

    if "operator" in rule and rule["operator"] not in _CELL_IS_OPERATORS:
        raise ValidationError(
            f"{where}/operator: unknown CellIs operator {rule['operator']!r}; "
            f"expected one of {sorted(_CELL_IS_OPERATORS)}"
        )
    if "dxf_id" in rule and not isinstance(rule["dxf_id"], int):
        raise ValidationError(f"{where}/dxf_id: expected int")

    if "color_scale" in rule:
        cs = rule["color_scale"]
        _require_keys(cs, ["thresholds", "colors"], f"{where}/color_scale")
        thresholds = cs["thresholds"]
        colors = cs["colors"]
        if not isinstance(thresholds, list) or not isinstance(colors, list):
            raise ValidationError(f"{where}/color_scale: thresholds and colors must be lists")
        if len(thresholds) != len(colors):
            raise ValidationError(
                f"{where}/color_scale: thresholds ({len(thresholds)}) and "
                f"colors ({len(colors)}) must have matching length"
            )
        for i, t in enumerate(thresholds):
            _require_keys(t, ["type"], f"{where}/color_scale/thresholds/{i}")
        for i, c in enumerate(colors):
            _validate_color(c, f"{where}/color_scale/colors/{i}")

    if "data_bar" in rule:
        db = rule["data_bar"]
        _require_keys(db, ["min", "max", "fill"], f"{where}/data_bar")
        for stop in ("min", "max"):
            _require_keys(db[stop], ["type"], f"{where}/data_bar/{stop}")
        _validate_color(db["fill"], f"{where}/data_bar/fill")
        for k in ("min_length_pct", "max_length_pct"):
            if k in db and not isinstance(db[k], int):
                raise ValidationError(f"{where}/data_bar/{k}: expected int")

    if "icon_set" in rule:
        ic = rule["icon_set"]
        _require_keys(ic, ["name", "thresholds"], f"{where}/icon_set")
        if not isinstance(ic["name"], str) or not ic["name"]:
            raise ValidationError(f"{where}/icon_set/name: expected non-empty string")
        thresholds = ic["thresholds"]
        if not isinstance(thresholds, list) or not thresholds:
            raise ValidationError(f"{where}/icon_set/thresholds: expected non-empty list")
        for i, t in enumerate(thresholds):
            _require_keys(t, ["type", "value"], f"{where}/icon_set/thresholds/{i}")


def _validate_cf_block(block: Any, where: str) -> None:
    _require_keys(block, ["sqref", "rules"], where)
    sqref = block["sqref"]
    if not isinstance(sqref, list) or not sqref:
        raise ValidationError(f"{where}/sqref: expected non-empty list")
    for i, sub in enumerate(sqref):
        _validate_range(sub, f"{where}/sqref/{i}")
    rules = block["rules"]
    if not isinstance(rules, list) or not rules:
        raise ValidationError(f"{where}/rules: expected non-empty list")
    for i, r in enumerate(rules):
        _validate_rule(r, f"{where}/rules/{i}")


def _validate_sheet_cell(cell: Any, where: str) -> None:
    _require_keys(cell, ["row", "col", "kind"], where)
    if not isinstance(cell["row"], int) or cell["row"] < 0:
        raise ValidationError(f"{where}/row: expected non-negative int")
    if not isinstance(cell["col"], int) or cell["col"] < 0:
        raise ValidationError(f"{where}/col: expected non-negative int")
    kind = cell["kind"]
    if kind not in _VALUE_KINDS:
        raise ValidationError(f"{where}/kind: unknown value kind {kind!r}; expected one of {sorted(_VALUE_KINDS)}")
    if kind != "blank" and "value" not in cell:
        raise ValidationError(f"{where}: kind={kind!r} requires a 'value' field")


def _validate_case_entry(case: Any, where: str) -> str:
    _require_keys(case, ["id", "sheet", "cf_blocks", "range"], where)
    cid = case["id"]
    if not isinstance(cid, str) or not cid:
        raise ValidationError(f"{where}/id: expected non-empty string")

    sheet = case["sheet"]
    if not isinstance(sheet, list):
        raise ValidationError(f"{where}/sheet: expected list")
    for i, c in enumerate(sheet):
        _validate_sheet_cell(c, f"{where}/sheet/{i}")

    blocks = case["cf_blocks"]
    if not isinstance(blocks, list) or not blocks:
        raise ValidationError(f"{where}/cf_blocks: expected non-empty list")
    for i, b in enumerate(blocks):
        _validate_cf_block(b, f"{where}/cf_blocks/{i}")

    _validate_range(case["range"], f"{where}/range")
    return cid


def validate_case_json(doc: Any) -> List[str]:
    """Returns the list of case ids declared in the case file."""
    _require_keys(doc, ["name", "cases"], _at())
    if not isinstance(doc["name"], str) or not doc["name"]:
        raise ValidationError(f"{_at('name')}: expected non-empty string")
    cases = doc["cases"]
    if not isinstance(cases, list) or not cases:
        raise ValidationError(f"{_at('cases')}: expected non-empty list")
    ids: List[str] = []
    seen: Set[str] = set()
    for i, case in enumerate(cases):
        cid = _validate_case_entry(case, _at("cases", i))
        if cid in seen:
            raise ValidationError(f"{_at('cases', i, 'id')}: duplicate id {cid!r}")
        seen.add(cid)
        ids.append(cid)
    return ids


def _validate_match(match: Any, where: str) -> None:
    _require_keys(match, ["kind", "priority"], where)
    kind = match["kind"]
    if kind not in _MATCH_KINDS:
        raise ValidationError(f"{where}/kind: unknown match kind {kind!r}; expected one of {sorted(_MATCH_KINDS)}")
    if not isinstance(match["priority"], int):
        raise ValidationError(f"{where}/priority: expected int")

    if kind == "DifferentialFormat":
        if "dxf_id" not in match or not isinstance(match["dxf_id"], int):
            raise ValidationError(f"{where}/dxf_id: required int for DifferentialFormat")
    elif kind == "ColorScale":
        if "color" not in match:
            raise ValidationError(f"{where}/color: required object for ColorScale")
        _validate_color(match["color"], f"{where}/color")
    elif kind == "DataBar":
        for k in ("bar_length_pct", "bar_axis_position_pct"):
            if k not in match or not isinstance(match[k], (int, float)):
                raise ValidationError(f"{where}/{k}: required number for DataBar")
        if "is_negative" not in match or not isinstance(match["is_negative"], bool):
            raise ValidationError(f"{where}/is_negative: required bool for DataBar")
        if "fill" not in match:
            raise ValidationError(f"{where}/fill: required object for DataBar")
        _validate_color(match["fill"], f"{where}/fill")
    elif kind == "IconSet":
        if "icon_set_name" not in match or not isinstance(match["icon_set_name"], str):
            raise ValidationError(f"{where}/icon_set_name: required string for IconSet")
        if "icon_index" not in match or not isinstance(match["icon_index"], int):
            raise ValidationError(f"{where}/icon_index: required int for IconSet")


def _validate_golden_cell(cell: Any, where: str) -> None:
    _require_keys(cell, ["row", "col", "matches"], where)
    if not isinstance(cell["row"], int) or cell["row"] < 0:
        raise ValidationError(f"{where}/row: expected non-negative int")
    if not isinstance(cell["col"], int) or cell["col"] < 0:
        raise ValidationError(f"{where}/col: expected non-negative int")
    matches = cell["matches"]
    if not isinstance(matches, list) or not matches:
        raise ValidationError(f"{where}/matches: expected non-empty list")
    for i, m in enumerate(matches):
        _validate_match(m, f"{where}/matches/{i}")


def validate_golden_json(doc: Any) -> List[str]:
    """Returns the list of case ids declared in the golden file."""
    _require_keys(doc, ["name", "cases"], _at())
    if not isinstance(doc["name"], str) or not doc["name"]:
        raise ValidationError(f"{_at('name')}: expected non-empty string")
    cases = doc["cases"]
    if not isinstance(cases, list) or not cases:
        raise ValidationError(f"{_at('cases')}: expected non-empty list")
    ids: List[str] = []
    seen: Set[str] = set()
    for i, case in enumerate(cases):
        where = _at("cases", i)
        _require_keys(case, ["id", "cells"], where)
        cid = case["id"]
        if not isinstance(cid, str) or not cid:
            raise ValidationError(f"{where}/id: expected non-empty string")
        if cid in seen:
            raise ValidationError(f"{where}/id: duplicate id {cid!r}")
        seen.add(cid)
        cells = case["cells"]
        if not isinstance(cells, list) or not cells:
            raise ValidationError(f"{where}/cells: expected non-empty list")
        for j, c in enumerate(cells):
            _validate_golden_cell(c, _at("cases", i, "cells", j))
        ids.append(cid)
    return ids


def _load_json(path: Path) -> Any:
    if not path.exists():
        raise ValidationError(f"{path}: file does not exist")
    with path.open("r", encoding="utf-8") as f:
        return json.load(f)


def validate_pair(case_path: Path, golden_path: Path) -> Tuple[List[str], List[str]]:
    """Validate both files and assert their case ids are consistent."""
    case_doc = _load_json(case_path)
    golden_doc = _load_json(golden_path)

    case_ids = validate_case_json(case_doc)
    golden_ids = validate_golden_json(golden_doc)

    case_set = set(case_ids)
    golden_set = set(golden_ids)
    missing_in_golden = sorted(case_set - golden_set)
    missing_in_case = sorted(golden_set - case_set)
    if missing_in_golden:
        raise ValidationError(f"{golden_path}: missing case ids present in {case_path.name}: {missing_in_golden}")
    if missing_in_case:
        raise ValidationError(f"{case_path}: missing case ids present in {golden_path.name}: {missing_in_case}")
    return case_ids, golden_ids


def _build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="Validate a CF case JSON file and its golden JSON file.",
    )
    p.add_argument(
        "case_json",
        type=Path,
        help="Path to tests/oracle/cases_cf/<suite>.case.json",
    )
    p.add_argument(
        "golden_json",
        type=Path,
        help="Path to tests/oracle/golden_cf/<suite>.golden.json",
    )
    return p


def main(argv: List[str]) -> int:
    args = _build_arg_parser().parse_args(argv)
    try:
        case_ids, _ = validate_pair(args.case_json, args.golden_json)
    except ValidationError as exc:
        print(f"validation error: {exc}", file=sys.stderr)
        return 1
    except json.JSONDecodeError as exc:
        print(f"JSON parse error: {exc}", file=sys.stderr)
        return 1
    print(f"OK: {args.case_json} <-> {args.golden_json} ({len(case_ids)} case(s): {', '.join(case_ids)})")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
