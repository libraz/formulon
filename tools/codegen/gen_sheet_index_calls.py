#!/usr/bin/env python3
"""Generate the call table of every sheet-scoped C ABI function for the C tests.

Reads ``src/c_api/formulon_c.h`` and the ``scope: sheet`` entries of
``tools/api/names.yaml`` and writes one table entry per function: a call that
passes the workbook handle, a caller-chosen ``sheet_index`` and a value for
every other parameter. Scalars and by-value records are value-initialised;
each pointer parameter gets its own zeroed slot, so a ``const char*`` reads
``""`` and an input record is all zeroes. The test may replace any pointer
argument with ``NULL``; each ``const char*`` slot is flagged so the test can
fill it with other text, and each single ``const fm_cell_rect_t*`` slot so it
can pass another rectangle. An owned out-handle (``fm_X_t**`` or ``uint8_t**``)
gets a release step through the ``*_destroy`` function taking its pointee.

Each single input record (``const fm_X_t*`` that is not an array with a count)
also gets a poison step that makes every field of it invalid: text reads as
malformed UTF-8, integers sit at their type's maximum, doubles are NaN, and
nested records are poisoned in turn. Pointer fields, ``*count`` fields, padding
arrays and union members keep their zeroes. One ``Poison`` overload per record
type precedes the table.

The table is a build artifact consumed by
``tests/c_api/formulon_c_sheet_index_status_test.cpp``.

Modes:
  --output PATH   write the table to PATH (rewritten only when it changes)
  --check         exit 1 when the table cannot be generated (a ledger entry
                  without a declaration, a function whose first two
                  parameters are not the handle and ``size_t sheet_index``,
                  or an owned out-handle without a destroy function)
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Dict, List, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT))

from tools.dev.binding_drift.api_conventions import (  # noqa: E402
    CELL_RECT,
    CFunction,
    CParam,
    CTypedef,
    load_ledger,
    parse_header,
)
from tools.dev.binding_drift.surface_files import CAPI_HEADER  # noqa: E402

WORKBOOK_TYPES = ("fm_workbook_t*", "const fm_workbook_t*")
INTEGER_TYPES = frozenset(
    {"uint8_t", "uint16_t", "uint32_t", "uint64_t", "int8_t", "int16_t", "int32_t", "int64_t", "size_t"}
)
_COMMENT_RE = re.compile(r"/\*.*?\*/|//[^\n]*", re.S)
_STRUCT_BODY_RE = re.compile(r"typedef\s+struct\s*\{(.*?)\}\s*(fm_[A-Za-z0-9_]+)\s*;", re.S)
_UNION_RE = re.compile(r"union\s*\{(.*?)\}\s*[A-Za-z_][A-Za-z0-9_]*\s*;", re.S)
_ARRAY_MEMBER_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_]*)\s*\[")
_IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")


def unpoisonable_members(header: str) -> Dict[str, set]:
    """Per record type, the members a poison step must leave alone: arrays and union members."""
    skipped: Dict[str, set] = {}
    for match in _STRUCT_BODY_RE.finditer(header):
        body = _COMMENT_RE.sub(" ", match.group(1))
        names = set(_ARRAY_MEMBER_RE.findall(body))
        for union in _UNION_RE.finditer(body):
            names.update(_IDENT_RE.findall(union.group(1)))
        skipped[match.group(2)] = names
    return skipped


def is_array_param(params: Tuple[CParam, ...], i: int) -> bool:
    """Whether parameter `i` is a pointer whose element count is the next parameter."""
    follower = params[i + 1] if i + 1 < len(params) else None
    return follower is not None and "*" not in follower.ctype and follower.name.endswith("count")


class Poisoner:
    """Collects the ``Poison`` overload of every record type an input record reaches."""

    def __init__(self, typedefs: List[CTypedef], skipped: Dict[str, set]) -> None:
        self.records = {t.name: t for t in typedefs if t.kind == "struct"}
        self.skipped = skipped
        self.bodies: Dict[str, List[str]] = {}
        self.order: List[str] = []

    def add(self, name: str) -> bool:
        """Emit the overload for record `name` (and the records it nests); False if it is not a record."""
        record = self.records.get(name)
        if record is None:
            return False
        if name in self.bodies:
            return True
        self.bodies[name] = []
        lines = []
        for field in record.fields:
            if field.name == "struct_size" or field.name in self.skipped.get(name, set()):
                continue
            target = f"r.{field.name}"
            if field.ctype == "const char*":
                lines.append(f'{target} = "\\xFF";')
            elif "*" in field.ctype or field.name.endswith("count"):
                continue
            elif field.ctype in INTEGER_TYPES:
                lines.append(f"{target} = std::numeric_limits<{field.ctype}>::max();")
            elif field.ctype == "double":
                lines.append(f"{target} = std::numeric_limits<double>::quiet_NaN();")
            elif self.add(field.ctype):
                lines.append(f"Poison({target});")
        self.bodies[name] = lines
        self.order.append(name)
        return True

    def render(self) -> List[str]:
        out = [f"void Poison({name}& r);" for name in self.order]
        for name in self.order:
            body = " ".join(self.bodies[name]) if self.bodies[name] else "(void)r;"
            out.append(f"void Poison({name}& r) {{ {body} }}")
        return out


def sheet_functions() -> Tuple[List[CFunction], Dict[str, str], List[str], Poisoner]:
    """The sheet-scoped declarations in header order, the destroy function per pointee, problems,
    and the poisoner over the header's record types."""
    header = CAPI_HEADER.read_text(encoding="utf-8")
    functions, typedefs = parse_header(header)
    ledger = load_ledger()
    problems = list(ledger.problems)
    by_name = {fn.name: fn for fn in functions}
    destroyers = {
        fn.params[0].ctype: fn.name
        for fn in functions
        if fn.name.endswith("_destroy") and fn.ret == "void" and len(fn.params) == 1
    }
    wanted = set()
    for key, entry in ledger.entries.items():
        if not isinstance(entry, dict) or entry.get("scope") != "sheet":
            continue
        name = entry.get("c")
        if not isinstance(name, str):
            continue
        # A dropped C function is gone from the header by design.
        if isinstance(entry.get("drop"), dict) and entry["drop"].get("c") == name:
            continue
        if name not in by_name:
            problems.append(f"gen_sheet_index_calls: `{key}` names {name}, which formulon_c.h does not declare")
            continue
        wanted.add(name)
    selected = [fn for fn in functions if fn.name in wanted]
    for fn in selected:
        if len(fn.params) < 2 or fn.params[0].ctype not in WORKBOOK_TYPES or fn.params[1] != ("size_t", "sheet_index"):
            problems.append(
                f"gen_sheet_index_calls: {fn.name} does not start with (workbook handle, size_t sheet_index)"
            )
    return selected, destroyers, problems, Poisoner(typedefs, unpoisonable_members(header))


def render(functions: List[CFunction], destroyers: Dict[str, str], problems: List[str], poisoner: Poisoner) -> str:
    entries = []
    for fn in functions:
        args = ["wb", "sheet_index"]
        releases = []
        poisons = []
        slot = 0
        text_mask = 0
        rect_mask = 0
        params = fn.params[2:]
        for i, param in enumerate(params):
            if "*" not in param.ctype:
                args.append("{}")
                continue
            pointee = param.ctype[:-1]
            args.append(f"slots.arg<{pointee}, {slot}>()")
            if param.ctype == "const char*":
                text_mask |= 1 << slot
            if param.ctype == f"const {CELL_RECT}*" and not is_array_param(params, i):
                rect_mask |= 1 << slot
            if (
                param.ctype.startswith("const ")
                and param.pointer_depth == 1
                and not is_array_param(params, i)
                and poisoner.add(param.base_type)
            ):
                poisons.append(f"Poison(*slots.at<{param.base_type}, {slot}>());")
            if pointee.endswith("*") and not pointee.startswith("const "):
                destroy = destroyers.get(pointee)
                if destroy is None:
                    problems.append(
                        f"gen_sheet_index_calls: {fn.name}({param.name}) returns {pointee} with no destroy function"
                    )
                else:
                    releases.append(f"{destroy}(*slots.at<{pointee}, {slot}>());")
            slot += 1
        slots = "slots" if slot else "/*slots*/"
        call = (
            f"[](fm_workbook_t* wb, std::size_t sheet_index, ZeroSlots& {slots}) -> fm_status_t "
            f"{{ return {fn.name}({', '.join(args)}); }}"
        )
        release = f"[](ZeroSlots& slots) {{ {' '.join(releases)} }}" if releases else "nullptr"
        poison = f"[](ZeroSlots& slots) {{ {' '.join(poisons)} }}" if poisons else "nullptr"
        entries.append(f'    {{"{fn.name}", {slot}, {text_mask}U, {rect_mask}U, {call}, {release}, {poison}}},')
    lines = [
        "// Generated by tools/codegen/gen_sheet_index_calls.py from src/c_api/formulon_c.h",
        "// and tools/api/names.yaml; do not edit.",
        *poisoner.render(),
        "const SheetIndexCall kSheetIndexCalls[] = {",
        *entries,
        "};",
    ]
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--output", type=Path, help="path of the generated table")
    mode.add_argument("--check", action="store_true", help="fail when the table cannot be generated")
    args = parser.parse_args()

    functions, destroyers, problems, poisoner = sheet_functions()
    content = render(functions, destroyers, problems, poisoner)
    if not functions:
        problems.append("gen_sheet_index_calls: the ledger names no sheet-scoped C function")
    for problem in problems:
        print(problem, file=sys.stderr)
    if problems:
        return 1
    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        if not args.output.is_file() or args.output.read_text(encoding="utf-8") != content:
            args.output.write_text(content, encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
