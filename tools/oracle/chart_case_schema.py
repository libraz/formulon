#!/usr/bin/env python3
"""Chart-track case and golden validator.

Case sources live in ``tests/oracle/cases_chart/<suite>.yaml``; the capture
(``chart_oracle_gen.py``) copies each case's spec into
``golden_chart/<suite>.golden.json`` next to what Excel saved, and the C++
verifier reads only the golden. This module checks both halves: the case
shape the C++ builder understands, and that every golden still carries the
spec its case file states (a case edited after capture is stale until it is
captured again).

Case shapes:

* built case -- ``sheets`` (+ optional ``names``), ``charts``, optional
  ``ops``. A chart either lists ``series`` (SERIES-formula references) or
  gives a ``source`` range for creation.
* edit case -- ``fixture`` (an Excel-saved file under
  ``tests/fixtures/excel``), ``target`` ``{sheet, chart}`` and ``edits``.

Usage::

    python3 tools/oracle/chart_case_schema.py [--cases-dir DIR] [--golden-dir DIR]

Exits 0 when everything validates, 1 otherwise.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional

import yaml

REPO_ROOT = Path(__file__).resolve().parents[2]
CASES_DIR = REPO_ROOT / "tests" / "oracle" / "cases_chart"
GOLDEN_DIR = REPO_ROOT / "tests" / "oracle" / "targets" / "mac-365-ja_JP" / "golden_chart"
FIXTURES_DIR = REPO_ROOT / "tests" / "fixtures" / "excel"

VIEWS = {"shape", "refs", "cache", "format"}
KINDS = {"column", "bar", "line", "area", "pie", "scatter", "doughnut"}
GROUPINGS = {"clustered", "stacked", "percentStacked", "standard"}
DIRECTIONS = {"auto", "rows", "columns"}
BLANKS = {"gap", "zero", "span"}
OPS = {"insert_rows", "delete_rows", "insert_cols", "delete_cols", "rename_sheet", "delete_sheet"}
EDITS = {"title", "legend", "series", "point", "axis", "plot", "filter", "display", "type"}
LEGEND_POSITIONS = {"none", "b", "t", "l", "r", "tr"}
TITLE_KINDS = {"none", "text", "auto"}
DASHES = {"solid", "dash", "dot", "dashDot"}
MARKERS = {"none", "auto", "circle", "square", "diamond", "triangle"}
LABEL_POSITIONS = {"bestFit", "ctr", "inBase", "inEnd", "outEnd", "t", "b", "l", "r"}
CROSSES = {"autoZero", "min", "max", "at"}
SPEC_KEYS = ("sheets", "names", "charts", "ops", "fixture", "target", "edits")
CASE_KEYS = {"id", "description", "factors", *SPEC_KEYS}

_A1 = re.compile(r"^[A-Z]{1,3}[1-9][0-9]*$")
_COLUMN = re.compile(r"^[A-Z]{1,3}$")
_RGB = re.compile(r"^[0-9A-F]{6}$")


class ValidationError(Exception):
    """A case or golden that does not match the schema."""


def _fail(where: str, message: str) -> None:
    raise ValidationError(f"{where}: {message}")


def _mapping(value: Any, where: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        _fail(where, "expected a mapping")
    return value


def _list(value: Any, where: str, *, non_empty: bool = True) -> List[Any]:
    if not isinstance(value, list) or (non_empty and not value):
        _fail(where, "expected a non-empty list" if non_empty else "expected a list")
    return value


def _string(value: Any, where: str) -> str:
    if not isinstance(value, str) or not value:
        _fail(where, "expected a non-empty string")
    return value


def _number(value: Any, where: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        _fail(where, "expected a number")
    return float(value)


def _bool(value: Any, where: str) -> bool:
    if not isinstance(value, bool):
        _fail(where, "expected a boolean")
    return value


def _enum(value: Any, allowed: set, where: str) -> str:
    if value not in allowed:
        _fail(where, f"expected one of {sorted(allowed)}, got {value!r}")
    return value


def _known_keys(obj: Mapping[str, Any], allowed: set, where: str) -> None:
    unknown = sorted(set(obj) - allowed)
    if unknown:
        _fail(where, f"unknown keys {unknown}")


def _cell(value: Any, where: str) -> None:
    if value is None or isinstance(value, (bool, int, float, str)):
        return
    obj = _mapping(value, where)
    _known_keys(obj, {"value", "format", "format_id"}, where)
    if "format" in obj:
        _string(obj["format"], f"{where}/format")
    if "format_id" in obj and (not isinstance(obj["format_id"], int) or not 0 <= obj["format_id"] < 164):
        _fail(f"{where}/format_id", "expected a built-in format id 0..163")
    if "value" in obj and not (obj["value"] is None or isinstance(obj["value"], (bool, int, float, str))):
        _fail(f"{where}/value", "expected a scalar")


def _sheets(value: Any, where: str) -> List[str]:
    names: List[str] = []
    for i, sheet in enumerate(_list(value, where)):
        at = f"{where}[{i}]"
        obj = _mapping(sheet, at)
        _known_keys(obj, {"name", "cells", "hidden_rows", "hidden_cols", "hidden"}, at)
        name = _string(obj.get("name"), f"{at}/name")
        if name in names:
            _fail(f"{at}/name", f"duplicate sheet {name!r}")
        names.append(name)
        for addr, cell in _mapping(obj.get("cells", {}), f"{at}/cells").items():
            if not isinstance(addr, str) or not _A1.match(addr):
                _fail(f"{at}/cells", f"bad address {addr!r}")
            _cell(cell, f"{at}/cells/{addr}")
        for row in _list(obj.get("hidden_rows", []), f"{at}/hidden_rows", non_empty=False):
            if isinstance(row, bool) or not isinstance(row, int) or row < 1:
                _fail(f"{at}/hidden_rows", "rows are 1-based integers")
        for col in _list(obj.get("hidden_cols", []), f"{at}/hidden_cols", non_empty=False):
            if not isinstance(col, str) or not _COLUMN.match(col):
                _fail(f"{at}/hidden_cols", "columns are letters")
        if "hidden" in obj:
            _bool(obj["hidden"], f"{at}/hidden")
    return names


def _chart(chart: Any, where: str, sheets: List[str]) -> str:
    obj = _mapping(chart, where)
    _known_keys(
        obj,
        {"id", "kind", "grouping", "host", "series", "source", "direction", "plot_vis_only", "blanks_as", "title_ref"},
        where,
    )
    cid = _string(obj.get("id"), f"{where}/id")
    kind = _enum(obj.get("kind"), KINDS, f"{where}/kind")
    if "grouping" in obj:
        _enum(obj["grouping"], GROUPINGS, f"{where}/grouping")
    if "host" in obj and obj["host"] not in sheets:
        _fail(f"{where}/host", f"unknown sheet {obj['host']!r}")
    if ("series" in obj) == ("source" in obj):
        _fail(where, "a chart gives exactly one of 'series' and 'source'")
    if "source" in obj:
        _string(obj["source"], f"{where}/source")
    if "direction" in obj:
        _enum(obj["direction"], DIRECTIONS, f"{where}/direction")
    if "series" in obj:
        if "direction" in obj:
            _fail(f"{where}/direction", "only a 'source' chart has a direction")
        value_key = "y" if kind == "scatter" else "val"
        allowed = {"name", "x", "y"} if kind == "scatter" else {"name", "cat", "val"}
        for i, series in enumerate(_list(obj["series"], f"{where}/series")):
            at = f"{where}/series[{i}]"
            s = _mapping(series, at)
            _known_keys(s, allowed, at)
            _string(s.get(value_key), f"{at}/{value_key}")
    if "plot_vis_only" in obj:
        _bool(obj["plot_vis_only"], f"{where}/plot_vis_only")
    if "blanks_as" in obj:
        _enum(obj["blanks_as"], BLANKS, f"{where}/blanks_as")
    if "title_ref" in obj:
        _string(obj["title_ref"], f"{where}/title_ref")
    return cid


def _op(op: Any, where: str) -> None:
    obj = _mapping(op, where)
    kind = _enum(obj.get("op"), OPS, f"{where}/op")
    _string(obj.get("sheet"), f"{where}/sheet")
    if kind == "rename_sheet":
        _known_keys(obj, {"op", "sheet", "to"}, where)
        _string(obj.get("to"), f"{where}/to")
    elif kind == "delete_sheet":
        _known_keys(obj, {"op", "sheet"}, where)
    else:
        _known_keys(obj, {"op", "sheet", "at", "count"}, where)
        at = obj.get("at")
        if kind.endswith("_rows"):
            if isinstance(at, bool) or not isinstance(at, int) or at < 1:
                _fail(f"{where}/at", "rows are 1-based integers")
        elif not isinstance(at, str) or not _COLUMN.match(at):
            _fail(f"{where}/at", "columns are letters")
        if "count" in obj and (isinstance(obj["count"], bool) or not isinstance(obj["count"], int) or obj["count"] < 1):
            _fail(f"{where}/count", "expected a positive integer")


def _rgb(value: Any, where: str) -> None:
    if not isinstance(value, str) or not _RGB.match(value):
        _fail(where, "colors are RRGGBB strings")


def _fill(value: Any, where: str) -> None:
    obj = _mapping(value, where)
    _known_keys(obj, {"color", "transparency"}, where)
    if obj.get("color") is not None:
        _rgb(obj["color"], f"{where}/color")
    if obj.get("transparency") is not None:
        _number(obj["transparency"], f"{where}/transparency")


def _line(value: Any, where: str) -> None:
    obj = _mapping(value, where)
    _known_keys(obj, {"color", "transparency", "width_pt", "dash", "none"}, where)
    if obj.get("color") is not None:
        _rgb(obj["color"], f"{where}/color")
    for key in ("transparency", "width_pt"):
        if obj.get(key) is not None:
            _number(obj[key], f"{where}/{key}")
    if obj.get("dash") is not None:
        _enum(obj["dash"], DASHES, f"{where}/dash")
    if "none" in obj:
        _bool(obj["none"], f"{where}/none")


def _labels(value: Any, where: str) -> None:
    if value is None:
        return
    obj = _mapping(value, where)
    _known_keys(
        obj,
        {"val", "cat", "ser", "pct", "legend_key", "separator", "position", "number_format", "font_color", "font_size"},
        where,
    )
    for key in ("val", "cat", "ser", "pct", "legend_key"):
        if key in obj:
            _bool(obj[key], f"{where}/{key}")
    if obj.get("position") is not None:
        _enum(obj["position"], LABEL_POSITIONS, f"{where}/position")
    if obj.get("font_color") is not None:
        _rgb(obj["font_color"], f"{where}/font_color")
    if obj.get("font_size") is not None:
        _number(obj["font_size"], f"{where}/font_size")


def _edit(edit: Any, where: str) -> None:
    obj = _mapping(edit, where)
    op = _enum(obj.get("op"), EDITS, f"{where}/op")
    if op == "title":
        _known_keys(obj, {"op", "kind", "text"}, where)
        kind = _enum(obj.get("kind"), TITLE_KINDS, f"{where}/kind")
        if kind == "text":
            _string(obj.get("text"), f"{where}/text")
    elif op == "legend":
        _known_keys(obj, {"op", "position", "overlay", "hidden"}, where)
        _enum(obj.get("position"), LEGEND_POSITIONS, f"{where}/position")
        if "overlay" in obj:
            _bool(obj["overlay"], f"{where}/overlay")
    elif op == "series":
        _known_keys(obj, {"op", "series", "fill", "line", "marker", "labels"}, where)
        series = obj.get("series")
        if series != "all" and (isinstance(series, bool) or not isinstance(series, int)):
            _fail(f"{where}/series", "expected a series idx or 'all'")
        if "fill" in obj:
            _fill(obj["fill"], f"{where}/fill")
        if "line" in obj:
            _line(obj["line"], f"{where}/line")
        if obj.get("marker") is not None:
            marker = _mapping(obj["marker"], f"{where}/marker")
            _known_keys(marker, {"symbol", "size", "fill", "line"}, f"{where}/marker")
            if "symbol" in marker:
                _enum(marker["symbol"], MARKERS, f"{where}/marker/symbol")
            if "fill" in marker:
                _fill(marker["fill"], f"{where}/marker/fill")
            if "line" in marker:
                _line(marker["line"], f"{where}/marker/line")
        if "labels" in obj:
            _labels(obj["labels"], f"{where}/labels")
    elif op == "point":
        _known_keys(obj, {"op", "series", "point", "fill", "line"}, where)
        for key in ("series", "point"):
            if isinstance(obj.get(key), bool) or not isinstance(obj.get(key), int):
                _fail(f"{where}/{key}", "expected an integer")
        if "fill" in obj:
            _fill(obj["fill"], f"{where}/fill")
        if "line" in obj:
            _line(obj["line"], f"{where}/line")
    elif op == "axis":
        _known_keys(
            obj,
            {
                "op",
                "axis",
                "visible",
                "title",
                "min",
                "max",
                "major_unit",
                "number_format",
                "major_gridlines",
                "reversed",
                "log_base",
                "crosses",
                "crosses_at",
                "tick_label_pos",
            },
            where,
        )
        _enum(obj.get("axis"), {"value", "category"}, f"{where}/axis")
        for key in ("min", "max", "major_unit", "log_base", "crosses_at"):
            if obj.get(key) is not None:
                _number(obj[key], f"{where}/{key}")
        if "crosses" in obj:
            _enum(obj["crosses"], CROSSES, f"{where}/crosses")
        if obj.get("tick_label_pos") is not None:
            _enum(obj["tick_label_pos"], {"nextTo", "low", "high", "none"}, f"{where}/tick_label_pos")
    elif op == "plot":
        _known_keys(obj, {"op", "gap_width", "overlap", "first_slice_angle", "vary_colors", "hole_size"}, where)
        if len(obj) == 1:
            _fail(where, "a plot edit changes at least one field")
    elif op == "filter":
        _known_keys(obj, {"op", "hidden_series", "hidden_categories"}, where)
    elif op == "display":
        _known_keys(obj, {"op", "blanks_as", "plot_vis_only"}, where)
        if "blanks_as" in obj:
            _enum(obj["blanks_as"], BLANKS, f"{where}/blanks_as")
        if "plot_vis_only" in obj:
            _bool(obj["plot_vis_only"], f"{where}/plot_vis_only")
    elif op == "type":
        _known_keys(obj, {"op", "kind", "grouping"}, where)
        _enum(obj.get("kind"), KINDS, f"{where}/kind")
        _enum(obj.get("grouping"), GROUPINGS, f"{where}/grouping")


def case_spec(case: Mapping[str, Any]) -> Dict[str, Any]:
    """The part of a case the builder reads; what the golden carries as ``spec``."""

    return {key: case[key] for key in SPEC_KEYS if key in case}


def validate_case(case: Any, where: str, *, fixtures_dir: Path = FIXTURES_DIR) -> str:
    obj = _mapping(case, where)
    _known_keys(obj, CASE_KEYS, where)
    cid = _string(obj.get("id"), f"{where}/id")
    where = f"{where}({cid})"
    if "fixture" in obj:
        for key in ("sheets", "names", "charts", "ops"):
            if key in obj:
                _fail(where, f"an edit case has no {key!r}")
        fixture = _string(obj["fixture"], f"{where}/fixture")
        if not (fixtures_dir / fixture).is_file():
            _fail(f"{where}/fixture", f"{fixture} is not under {fixtures_dir}")
        target = _mapping(obj.get("target"), f"{where}/target")
        _known_keys(target, {"sheet", "chart"}, f"{where}/target")
        _string(target.get("sheet"), f"{where}/target/sheet")
        _string(target.get("chart"), f"{where}/target/chart")
        for i, edit in enumerate(_list(obj.get("edits"), f"{where}/edits")):
            _edit(edit, f"{where}/edits[{i}]")
        return cid
    if "target" in obj or "edits" in obj:
        _fail(where, "'target' and 'edits' need a 'fixture'")
    sheets = _sheets(obj.get("sheets"), f"{where}/sheets")
    for i, name in enumerate(_list(obj.get("names", []), f"{where}/names", non_empty=False)):
        n = _mapping(name, f"{where}/names[{i}]")
        _known_keys(n, {"name", "refers_to", "scope"}, f"{where}/names[{i}]")
        _string(n.get("name"), f"{where}/names[{i}]/name")
        _string(n.get("refers_to"), f"{where}/names[{i}]/refers_to")
        if "scope" in n and n["scope"] not in sheets:
            _fail(f"{where}/names[{i}]/scope", f"unknown sheet {n['scope']!r}")
    ids: List[str] = []
    for i, chart in enumerate(_list(obj.get("charts"), f"{where}/charts")):
        chart_id = _chart(chart, f"{where}/charts[{i}]", sheets)
        if chart_id in ids:
            _fail(f"{where}/charts[{i}]/id", f"duplicate chart id {chart_id!r}")
        ids.append(chart_id)
    for i, op in enumerate(_list(obj.get("ops", []), f"{where}/ops", non_empty=False)):
        _op(op, f"{where}/ops[{i}]")
    return cid


def validate_suite(doc: Any, stem: str, *, fixtures_dir: Path = FIXTURES_DIR) -> List[str]:
    """Validates one case file and returns its case ids."""

    obj = _mapping(doc, stem)
    _known_keys(obj, {"suite", "description", "views", "capture_formulon", "templates", "cases"}, stem)
    if obj.get("suite") != stem:
        _fail(stem, f"suite must equal the file stem, got {obj.get('suite')!r}")
    views = _list(obj.get("views"), f"{stem}/views")
    for view in views:
        _enum(view, VIEWS, f"{stem}/views")
    if "capture_formulon" in obj:
        _bool(obj["capture_formulon"], f"{stem}/capture_formulon")
    ids: List[str] = []
    for i, case in enumerate(_list(obj.get("cases"), f"{stem}/cases")):
        cid = validate_case(case, f"{stem}/cases[{i}]", fixtures_dir=fixtures_dir)
        if cid in ids:
            _fail(f"{stem}/cases[{i}]", f"duplicate id {cid!r}")
        ids.append(cid)
    return ids


def load_suites(cases_dir: Path = CASES_DIR) -> Dict[str, Dict[str, Any]]:
    """Every chart suite by name, parsed but not validated."""

    out: Dict[str, Dict[str, Any]] = {}
    for path in sorted(cases_dir.glob("*.yaml")):
        out[path.stem] = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return out


def validate_golden(golden: Any, suite: Mapping[str, Any], stem: str) -> List[str]:
    """Checks a golden against its case file: same views, same ids, same specs."""

    where = f"{stem}.golden.json"
    obj = _mapping(golden, where)
    if obj.get("name") != stem:
        _fail(where, f"name must be {stem!r}")
    if obj.get("views") != suite.get("views"):
        _fail(where, "views differ from the case file")
    _mapping(obj.get("environment"), f"{where}/environment")
    by_id = {case["id"]: case for case in suite["cases"]}
    seen: List[str] = []
    for i, case in enumerate(_list(obj.get("cases"), f"{where}/cases")):
        at = f"{where}/cases[{i}]"
        entry = _mapping(case, at)
        cid = _string(entry.get("id"), f"{at}/id")
        if cid not in by_id:
            _fail(at, f"{cid!r} is not in {stem}.yaml")
        if cid in seen:
            _fail(at, f"duplicate id {cid!r}")
        seen.append(cid)
        # Compare through JSON so YAML-only spellings (1.0 vs 1) cannot differ.
        if json.loads(json.dumps(case_spec(by_id[cid]))) != entry.get("spec"):
            _fail(at, f"spec of {cid!r} differs from {stem}.yaml; capture the suite again")
        if "skipped" in entry:
            # A skip-oracle case records the registry's reason in place of `expect`.
            _string(entry["skipped"], f"{at}/skipped")
            if "expect" in entry:
                _fail(at, "a skipped case carries no 'expect'")
            continue
        charts = _list(
            _mapping(entry.get("expect"), f"{at}/expect").get("charts"), f"{at}/expect/charts", non_empty=False
        )
        for j, chart in enumerate(charts):
            missing = [view for view in suite["views"] if view not in _mapping(chart, f"{at}/expect/charts[{j}]")]
            if missing:
                _fail(f"{at}/expect/charts[{j}]", f"missing views {missing}")
    missing_cases = sorted(set(by_id) - set(seen))
    if missing_cases:
        _fail(where, f"cases not captured: {missing_cases}")
    return seen


def validate_all(
    cases_dir: Path = CASES_DIR, golden_dir: Optional[Path] = GOLDEN_DIR, *, fixtures_dir: Path = FIXTURES_DIR
) -> int:
    """Validates every suite and every golden present; returns the case count."""

    suites = load_suites(cases_dir)
    if not suites:
        raise ValidationError(f"{cases_dir}: no chart suites")
    owners: Dict[str, str] = {}
    total = 0
    for stem, doc in suites.items():
        for cid in validate_suite(doc, stem, fixtures_dir=fixtures_dir):
            if cid in owners:
                raise ValidationError(f"{stem}: id {cid!r} also defined by {owners[cid]}")
            owners[cid] = stem
            total += 1
    if golden_dir is not None and golden_dir.is_dir():
        for path in sorted(golden_dir.glob("*.golden.json")):
            stem = path.name[: -len(".golden.json")]
            if stem not in suites:
                raise ValidationError(f"{path.name}: no case file {stem}.yaml")
            validate_golden(json.loads(path.read_text(encoding="utf-8")), suites[stem], stem)
    return total


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Validate chart-track case files and goldens.")
    parser.add_argument("--cases-dir", type=Path, default=CASES_DIR)
    parser.add_argument("--golden-dir", type=Path, default=GOLDEN_DIR)
    args = parser.parse_args(argv)
    try:
        total = validate_all(args.cases_dir, args.golden_dir)
    except (ValidationError, json.JSONDecodeError, yaml.YAMLError) as exc:
        print(f"chart-case-schema: {exc}", file=sys.stderr)
        return 1
    print(f"OK: {total} chart case(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
