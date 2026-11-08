#!/usr/bin/env python3
"""Validates the shared case tables under `tools/api/`.

Checks `parse_error_cases.json`, `date_value_cases.json`, `roundtrip_cases.json`,
`roundtrip_pending.json` and `to_native_cases.json` for shape, so a malformed row fails
here instead of in one of the five suites that read the tables:

* parse errors: unique ids, a known entry, `formula` xor `formula_repeat`, every code a
  `ParseErrorCode` member (src/parser/parse_error.h plus the entry codes 29..36), a status
  equal to the one the code maps to, start <= end, and line / column agreeing with start;
* date values: unique ids, a known direction and input kind, an expectation per surface;
* round trips: unique ids, ledger ops that exist, and pending rows naming a real row;
* native values: every row carries a `native`, and a call that exists.

Exit 0 when every table is well formed. `--self-test` runs the validators over malformed
rows and fails unless each is rejected.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Callable, Dict, List

TOOLS_API = Path(__file__).resolve().parent
REPO_ROOT = TOOLS_API.parent.parent
PARSE_ERROR_H = REPO_ROOT / "src" / "parser" / "parse_error.h"

# ParseErrorCode values appended after the parser's own enum; the header takes over as it adds them.
ENTRY_CODES = {
    "EntryWrongArity": 29,
    "EntryReferenceRequired": 30,
    "EntryQualifiedBuiltin": 31,
    "EntryRefusedWhitespace": 32,
    "EntryParamShapedName": 33,
    "EntryDoubleEquals": 34,
    "EntryBomNotSupported": 35,
    "StructuredRefInvalid": 36,
}
# Code -> status. Every other code is 1000.
STATUS_OF = {
    "LexerUnterminatedString": 1001,
    "LexerInvalidNumberLiteral": 1003,
    "LexerInvalidReference": 1004,
    "InvalidReference": 1004,
    "InvalidRange": 1004,
    "LexerUnterminatedSheetQuote": 1005,
    "UnbalancedBrackets": 1006,
    "UnbalancedBraces": 1006,
    "ExpectedCloseParen": 1006,
    "NestedFormulaTooDeep": 1007,
    "ArrayRowMismatch": 1009,
    "ExpectedCommaOrSemiInArray": 1009,
    "StructuredRefInvalid": 1010,
    "EntryBomNotSupported": 1013,
    "TooManyErrors": 1014,
}
DEFAULT_STATUS = 1000
PARSE_ENTRIES = {
    "set_formula",
    "set_array_formula",
    "evaluate_formula",
    "evaluate_formula_array",
    "evaluate_cf_formula",
    "defined_name",
    "cf_rule_add",
    "cf_rule_update",
    "validation_add",
    "validation_update",
    "table_column_formula",
    "pivot_calculated_field",
    "pivot_calculated_item",
    "parse_reference",
    "cli_eval",
}
RECORD_ENTRIES = {
    "defined_name",
    "cf_rule_add",
    "cf_rule_update",
    "validation_add",
    "validation_update",
    "table_column_formula",
    "pivot_calculated_field",
    "pivot_calculated_item",
}
BASES = {"measured", "design"}
SURFACES = {"js", "py"}
BINDING_SURFACES = {"wasm", "native", "py"}
WRITE_KINDS = {"date", "datetime", "time", "aware_datetime", "timedelta", "invalid_date", "decimal"}
EFFECTIVE_FORMATS = {"general", "custom", "row_style", "column_style"}
NATIVE_CALLS = {"evaluate_formula", "evaluate_formula_array", "get_value", "get_values"}
ROUNDTRIP_KINDS = {"patch", "readd"}


def parse_error_codes() -> Dict[str, int]:
    """ParseErrorCode members (name -> value) from the header, plus the entry codes it does not carry yet."""
    text = PARSE_ERROR_H.read_text(encoding="utf-8")
    body = re.search(r"enum class ParseErrorCode[^{]*\{(.*?)\n\};", text, re.S)
    if body is None:
        raise SystemExit(f"check_case_tables: no ParseErrorCode enum in {PARSE_ERROR_H}")
    codes: Dict[str, int] = {}
    value = 0
    for line in body.group(1).splitlines():
        line = line.split("//", 1)[0].strip()
        for name in [p.strip() for p in line.split(",") if p.strip()]:
            member = re.fullmatch(r"(\w+)(?:\s*=\s*(\d+))?", name)
            if member is None:
                continue
            if member.group(2) is not None:
                value = int(member.group(2))
            codes[member.group(1)] = value
            value += 1
    for name, number in ENTRY_CODES.items():
        if codes.setdefault(name, number) != number:
            raise SystemExit(f"check_case_tables: {name} is {codes[name]} in the header, {number} in the design")
    return codes


def _int(v) -> bool:
    return isinstance(v, int) and not isinstance(v, bool)


def _unique(rows: List[dict], table: str) -> List[str]:
    ids = [r.get("id") for r in rows]
    problems = [f"{table}: row {i} has no string id" for i, v in enumerate(ids) if not isinstance(v, str) or not v]
    problems += [
        f"{table}: duplicate id {v}" for v in sorted({v for v in ids if isinstance(v, str) and ids.count(v) > 1})
    ]
    return problems


def _utf16_len(text: str) -> int:
    return len(text.encode("utf-16-le")) // 2


def check_parse_errors(doc: dict, codes: Dict[str, int]) -> List[str]:
    rows = doc.get("cases")
    if not isinstance(rows, list) or not rows:
        return ["parse_error_cases: `cases` is empty"]
    problems = _unique(rows, "parse_error_cases")
    for row in rows:
        w = f"parse_error_cases[{row.get('id')}]"
        if row.get("entry") not in PARSE_ENTRIES:
            problems.append(f"{w}: unknown entry {row.get('entry')!r}")
        if ("formula" in row) == ("formula_repeat" in row):
            problems.append(f"{w}: exactly one of formula / formula_repeat")
        text = None
        if isinstance(row.get("formula"), str):
            text = row["formula"]
        elif isinstance(row.get("formula_repeat"), list):
            segments = row["formula_repeat"]
            if not all(
                isinstance(s, list) and len(s) == 2 and isinstance(s[0], str) and _int(s[1]) and s[1] > 0
                for s in segments
            ):
                problems.append(f"{w}: formula_repeat is a list of [text, count>0]")
            else:
                text = "".join(t * n for t, n in segments)
        elif "formula" in row:
            problems.append(f"{w}: formula is a string")
        if row.get("basis") not in BASES:
            problems.append(f"{w}: basis is one of {sorted(BASES)}")
        if (row.get("entry") in RECORD_ENTRIES) != ("field" in row):
            problems.append(f"{w}: `field` belongs to the record entries, and to them only")
        names = row.get("codes")
        if not isinstance(names, list) or not names or not all(isinstance(n, str) for n in names):
            problems.append(f"{w}: codes is a non-empty list of names")
            continue
        unknown = [n for n in names if n not in codes]
        if unknown:
            problems.append(f"{w}: not ParseErrorCode members: {unknown}")
        status = row.get("status")
        if not _int(status):
            problems.append(f"{w}: status is an integer")
        else:
            wanted = {STATUS_OF.get(n, DEFAULT_STATUS) for n in names}
            if wanted != {status}:
                problems.append(f"{w}: status {status} does not match its codes' status {sorted(wanted)}")
        start, end, line, column = (row.get(k) for k in ("start", "end", "line", "column"))
        if start is None or end is None:
            if not (start is None and end is None and line is None and column is None):
                problems.append(f"{w}: start / end / line / column are all null or all set")
        else:
            if not (_int(start) and _int(end) and _int(line) and _int(column)) or not 0 <= start <= end:
                problems.append(f"{w}: start <= end, all integers")
            elif text is not None:
                units = _utf16_len(text)
                if end > units:
                    problems.append(f"{w}: end {end} is past the {units}-unit text")
                else:
                    prefix = text.encode("utf-16-le")[: start * 2].decode("utf-16-le", "surrogatepass")
                    want_line = 1 + prefix.count("\n")
                    want_col = start - (_utf16_len(prefix[: prefix.rfind("\n") + 1]) if "\n" in prefix else 0) + 1
                    if (line, column) != (want_line, want_col):
                        problems.append(
                            f"{w}: line/column {line}:{column} disagree with start {start} ({want_line}:{want_col})"
                        )
        if row.get("entry") == "parse_reference" and (
            status != 1004 or start != 0 or (text is not None and end != _utf16_len(text))
        ):
            problems.append(f"{w}: parse_reference fails with 1004 over 0..length")
        if row.get("entry") == "cli_eval":
            if row.get("exit") != 65:
                problems.append(f"{w}: cli_eval exits 65")
            if "stderr_suffix" not in row:
                problems.append(f"{w}: cli_eval rows carry stderr_suffix (null when not asserted)")
        elif "exit" in row or "stderr_suffix" in row:
            problems.append(f"{w}: exit / stderr_suffix belong to cli_eval rows")
    return problems


def _expect_ok(expect, w: str, direction: str) -> List[str]:
    if not isinstance(expect, dict):
        return [f"{w}: expect entry is an object"]
    if "error" in expect:
        if not (isinstance(expect["error"], str) and expect.get("layer") in (1, 2)) or set(expect) != {
            "error",
            "layer",
        }:
            return [f"{w}: an error expectation is {{error, layer: 1|2}}"]
        return []
    if direction == "write":
        if set(expect) != {"serial", "format", "format_id"} or not isinstance(expect["serial"], (int, float)):
            return [f"{w}: a write expectation is {{serial, format, format_id}} or {{error, layer}}"]
        if (expect["format"] is None) != (expect["format_id"] is None):
            return [f"{w}: format and format_id are both null or both set"]
        return []
    if "value" in expect and expect["value"] is None and len(expect) == 1:
        return []
    date = expect.get("date")
    sub = {"millisecond", "microsecond"} & set(date or {})
    if not (
        isinstance(date, dict)
        and len(expect) == 1
        and len(sub) == 1
        and {"year", "month", "day", "hour", "minute", "second"} <= set(date)
    ):
        return [
            f"{w}: a read expectation is {{date: {{year..second, millisecond|microsecond}}}}, {{value: null}} or {{error, layer}}"
        ]
    return []


def check_date_values(doc: dict) -> List[str]:
    rows = doc.get("cases")
    if not isinstance(rows, list) or not rows:
        return ["date_value_cases: `cases` is empty"]
    problems = _unique(rows, "date_value_cases")
    for row in rows:
        w = f"date_value_cases[{row.get('id')}]"
        direction = row.get("direction")
        if direction not in ("write", "read"):
            problems.append(f"{w}: direction is write | read")
            continue
        if not isinstance(row.get("date1904"), bool):
            problems.append(f"{w}: date1904 is a bool")
        inp = row.get("input")
        if not isinstance(inp, dict):
            problems.append(f"{w}: input is an object")
            continue
        if direction == "write":
            if inp.get("kind") not in WRITE_KINDS:
                problems.append(f"{w}: input.kind is one of {sorted(WRITE_KINDS)}")
            if row.get("effective_format") not in EFFECTIVE_FORMATS:
                problems.append(f"{w}: effective_format is one of {sorted(EFFECTIVE_FORMATS)}")
        else:
            if ("serial" in inp) == ("cell" in inp) or len(inp) != 1:
                problems.append(f"{w}: a read input is exactly {{serial}} or {{cell}}")
            elif "cell" in inp and inp["cell"] not in ("blank", "text", "bool", "error"):
                problems.append(f"{w}: input.cell is blank | text | bool | error")
            elif "serial" in inp and not (
                isinstance(inp["serial"], (int, float)) or inp["serial"] in ("NaN", "Infinity", "-Infinity")
            ):
                problems.append(f"{w}: input.serial is a number or NaN / Infinity / -Infinity")
        expect = row.get("expect")
        only = row.get("only", sorted(SURFACES))
        if not isinstance(only, list) or not set(only) <= SURFACES or not only:
            problems.append(f"{w}: only lists surfaces from {sorted(SURFACES)}")
            continue
        if not isinstance(expect, dict) or set(expect) != set(only):
            problems.append(f"{w}: expect has exactly one entry per running surface {sorted(only)}")
            continue
        for surface in only:
            problems += _expect_ok(expect[surface], f"{w}.{surface}", direction)
    return problems


def check_roundtrip(doc: dict, pending: dict, ledger_ids) -> List[str]:
    rows = doc.get("cases")
    if not isinstance(rows, list) or not rows:
        return ["roundtrip_cases: `cases` is empty"]
    problems = _unique(rows, "roundtrip_cases")
    for row in rows:
        w = f"roundtrip_cases[{row.get('id')}]"
        if row.get("kind") not in ROUNDTRIP_KINDS:
            problems.append(f"{w}: kind is one of {sorted(ROUNDTRIP_KINDS)}")
        if row.get("scope") not in ("sheet", "workbook"):
            problems.append(f"{w}: scope is sheet | workbook")
        if not isinstance(row.get("record"), str) or not row["record"]:
            problems.append(f"{w}: record names the record type")
        surfaces = row.get("surfaces")
        if not isinstance(surfaces, list) or not surfaces or not set(surfaces) <= BINDING_SURFACES:
            problems.append(f"{w}: surfaces lists binding surfaces from {sorted(BINDING_SURFACES)}")
        ops = [row.get("get"), row.get("set")]
        for op in ops:
            if not (isinstance(op, dict) and all(isinstance(op.get(k), str) and op[k] for k in ("op", "js", "py"))):
                problems.append(f"{w}: get / set are {{op, js, py}}")
            elif ledger_ids is not None and op["op"] not in ledger_ids:
                problems.append(f"{w}: `{op['op']}` is not in names.yaml")
        create = row.get("create")
        if not isinstance(create, list) or not all(isinstance(c, str) for c in create):
            problems.append(f"{w}: create is a list of op ids")
        elif ledger_ids is not None:
            problems += [f"{w}: `{c}` is not in names.yaml" for c in create if c not in ledger_ids]
    ids = {r.get("id") for r in rows}
    pend = pending.get("pending")
    if not isinstance(pend, list):
        return problems + ["roundtrip_pending: `pending` is a list"]
    seen = set()
    for item in pend:
        w = f"roundtrip_pending[{item.get('id') if isinstance(item, dict) else item}]"
        if not isinstance(item, dict) or item.get("id") not in ids:
            problems.append(f"{w}: id names no row of roundtrip_cases.json")
            continue
        if item["id"] in seen:
            problems.append(f"{w}: listed twice")
        seen.add(item["id"])
        for key in ("reason", "unit"):
            if not isinstance(item.get(key), str) or not item[key].strip():
                problems.append(f"{w}: {key} is a non-empty string")
        surfaces = item.get("surfaces")
        if not isinstance(surfaces, list) or not surfaces or not set(surfaces) <= BINDING_SURFACES:
            problems.append(f"{w}: surfaces lists binding surfaces from {sorted(BINDING_SURFACES)}")
    return problems


def _native_ok(value) -> bool:
    if value is None or isinstance(value, (bool, int, float, str)):
        return True
    if isinstance(value, dict):
        return (set(value) == {"error"} and isinstance(value["error"], str)) or value == {"passthrough": True}
    if isinstance(value, list):
        return bool(value) and all(
            isinstance(r, list) and r and len(r) == len(value[0]) and all(_native_ok(c) for c in r) for r in value
        )
    return False


def check_to_native(doc: dict) -> List[str]:
    problems: List[str] = []
    cases = doc.get("cases")
    if not isinstance(cases, list) or not cases:
        problems.append("to_native_cases: `cases` is empty")
    else:
        for i, row in enumerate(cases):
            w = f"to_native_cases.cases[{i}]"
            if ("formula" in row) == ("cell" in row):
                problems.append(f"{w}: exactly one of formula / cell")
            if "native" not in row or not _native_ok(row["native"]):
                problems.append(
                    f"{w}: native is a scalar, null, {{error}}, {{passthrough}} or a rectangular list of rows"
                )
    facade = doc.get("facade_cases")
    if not isinstance(facade, list) or not facade:
        return problems + ["to_native_cases: `facade_cases` is empty"]
    problems += _unique(facade, "to_native_cases.facade_cases")
    for row in facade:
        w = f"to_native_cases.facade_cases[{row.get('id')}]"
        call = row.get("call")
        if call not in NATIVE_CALLS:
            problems.append(f"{w}: call is one of {sorted(NATIVE_CALLS)}")
            continue
        source = [k for k in ("formula", "cell", "range") if k in row]
        wanted = ["formula"] if call.startswith("evaluate") else ["cell"] if call == "get_value" else ["range"]
        if source != wanted or not isinstance(row[wanted[0]], str):
            problems.append(f"{w}: {call} takes {wanted[0]}")
        if "native" not in row or not _native_ok(row["native"]):
            problems.append(f"{w}: native is a scalar, null, {{error}} or a rectangular list of rows")
        elif call in ("evaluate_formula_array", "get_values") and not isinstance(row["native"], list):
            problems.append(f"{w}: {call} always returns a two-dimensional list")
    return problems


def _ledger_ids():
    try:
        import yaml  # type: ignore
    except ImportError:
        return None
    return set(yaml.safe_load((TOOLS_API / "names.yaml").read_text(encoding="utf-8"))["entries"])


def _load(name: str) -> dict:
    return json.loads((TOOLS_API / name).read_text(encoding="utf-8"))


def check_all() -> List[str]:
    return (
        check_parse_errors(_load("parse_error_cases.json"), parse_error_codes())
        + check_date_values(_load("date_value_cases.json"))
        + check_roundtrip(_load("roundtrip_cases.json"), _load("roundtrip_pending.json"), _ledger_ids())
        + check_to_native(_load("to_native_cases.json"))
    )


def self_test() -> List[str]:
    """Malformed rows each validator must reject; returns the ones it let through."""
    codes = parse_error_codes()
    good_parse = {
        "id": "x",
        "entry": "set_formula",
        "formula": "=(1+2",
        "codes": ["ExpectedCloseParen"],
        "status": 1006,
        "start": None,
        "end": None,
        "line": None,
        "column": None,
        "basis": "design",
    }
    good_date = {
        "id": "x",
        "direction": "read",
        "date1904": False,
        "input": {"serial": 1.0},
        "expect": {
            "js": {"date": {"year": 1900, "month": 1, "day": 1, "hour": 0, "minute": 0, "second": 0, "millisecond": 0}},
            "py": {"date": {"year": 1900, "month": 1, "day": 1, "hour": 0, "minute": 0, "second": 0, "microsecond": 0}},
        },
    }
    good_rt = {
        "id": "x",
        "kind": "patch",
        "scope": "sheet",
        "record": "R",
        "get": {"op": "a.get", "js": "a", "py": "a"},
        "set": {"op": "a.set", "js": "b", "py": "b"},
        "create": [],
        "surfaces": ["py"],
    }
    good_native = {
        "cases": [{"formula": "=1", "native": 1}],
        "facade_cases": [{"id": "x", "call": "evaluate_formula", "formula": "=1", "native": 1}],
    }
    pending_ok = {"pending": []}

    def with_(base: dict, **kw) -> dict:
        out = dict(base)
        out.update(kw)
        return out

    def parse(row):
        return check_parse_errors({"cases": [row]}, codes)

    def date(row):
        return check_date_values({"cases": [row]})

    def rt(row, pending=pending_ok):
        return check_roundtrip({"cases": [row]}, pending, {"a.get", "a.set"})

    # The well-formed rows must pass, or the negatives below prove nothing.
    baseline = parse(good_parse) + date(good_date) + rt(good_rt) + check_to_native(good_native)
    if baseline:
        return [f"well-formed fixture rejected: {baseline}"]
    negatives: Dict[str, Callable[[], List[str]]] = {
        "unknown code": lambda: parse(with_(good_parse, codes=["NoSuchCode"])),
        "status disagrees with code": lambda: parse(with_(good_parse, status=1001)),
        "unknown entry": lambda: parse(with_(good_parse, entry="nope")),
        "both formula and formula_repeat": lambda: parse(with_(good_parse, formula_repeat=[["=", 1]])),
        "start past end": lambda: parse(with_(good_parse, start=5, end=2, line=1, column=6)),
        "line/column off": lambda: parse(with_(good_parse, start=2, end=3, line=1, column=9)),
        "end past text": lambda: parse(with_(good_parse, start=0, end=99, line=1, column=1)),
        "partial position": lambda: parse(with_(good_parse, start=1)),
        "bad basis": lambda: parse(with_(good_parse, basis="guess")),
        "field on a plain entry": lambda: parse(with_(good_parse, field="formula1")),
        "duplicate id": lambda: check_parse_errors({"cases": [good_parse, good_parse]}, codes),
        "date: bad direction": lambda: date(with_(good_date, direction="sideways")),
        "date: surface missing from expect": lambda: date(with_(good_date, expect={"js": good_date["expect"]["js"]})),
        "date: error without layer": lambda: date(
            with_(good_date, expect={"js": {"error": "X"}, "py": {"error": "X"}})
        ),
        "round trip: op not in ledger": lambda: rt(with_(good_rt, set={"op": "z.set", "js": "b", "py": "b"})),
        "round trip: bad surface": lambda: rt(with_(good_rt, surfaces=["java"])),
        "pending names no row": lambda: rt(
            good_rt, {"pending": [{"id": "ghost", "surfaces": ["py"], "reason": "r", "unit": "U"}]}
        ),
        "pending without reason": lambda: rt(
            good_rt, {"pending": [{"id": "x", "surfaces": ["py"], "reason": "", "unit": "U"}]}
        ),
        "native: ragged grid": lambda: check_to_native(
            {
                "cases": good_native["cases"],
                "facade_cases": [{"id": "x", "call": "evaluate_formula", "formula": "=1", "native": [[1, 2], [3]]}],
            }
        ),
        "native: array call returning a scalar": lambda: check_to_native(
            {
                "cases": good_native["cases"],
                "facade_cases": [{"id": "x", "call": "evaluate_formula_array", "formula": "=1", "native": 1}],
            }
        ),
    }
    return [name for name, run in negatives.items() if not run()]


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--self-test", action="store_true", help="only prove the validators reject malformed rows")
    args = parser.parse_args(argv)
    leaked = self_test()
    if leaked:
        print("check_case_tables self-test: malformed rows accepted: " + "; ".join(leaked), file=sys.stderr)
        return 1
    if args.self_test:
        return 0
    problems = check_all()
    for p in problems:
        print(p, file=sys.stderr)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
