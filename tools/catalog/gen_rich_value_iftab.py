#!/usr/bin/env python3
"""Generates the (function, argument position) table of propagated rich `#VALUE!` errors.

A cell whose formula hands on a rich `#VALUE!` from another cell is saved by Excel with one of three rich
values, decided by the function at the formula's root and the argument that carried the error:
`{errorType 2, iftab <function id>, propagated 1}`, a plain error (no rich value), or `{2, propagated 1}`.
The measured sweep is the source of truth; no rule reproduces it.

Inputs (tools/catalog/data/):
  rich_value_iftab_sweep.json           positions 1-2 of every catalog function
  rich_value_iftab_sweep_pos3plus.json  positions 3-6 (3-9 for ODDF*/ODDL*) plus LET/LAMBDA forms
  rich_value_iftab_variadic_tail.json   positions past the sweep cap, checked against the folded table

Outputs:
  src/io/generated/rich_value_iftab.inc        the table the writer reads
  tests/unit/io/rich_value_iftab_sweep_rows.inc  every measured sweep formula with Excel's shown value

Usage:
  tools/catalog/gen_rich_value_iftab.py           write both outputs
  tools/catalog/gen_rich_value_iftab.py --check   exit 1 when an output is stale

Stdlib only.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = REPO_ROOT / "tools" / "catalog" / "data"
SWEEP_PATHS = [DATA_DIR / "rich_value_iftab_sweep.json", DATA_DIR / "rich_value_iftab_sweep_pos3plus.json"]
TAIL_PATH = DATA_DIR / "rich_value_iftab_variadic_tail.json"
CATALOG_PATH = REPO_ROOT / "tools" / "catalog" / "functions.txt"
FUNC_ID_TABLE_PATH = REPO_ROOT / "src" / "io" / "xlsb" / "func_id_table.cpp"
OUT_INC = REPO_ROOT / "src" / "io" / "generated" / "rich_value_iftab.inc"
OUT_ROWS = REPO_ROOT / "tests" / "unit" / "io" / "rich_value_iftab_sweep_rows.inc"

# Results that state what Excel stored for the position.
MEASURED = {"iftab", "none", "plain-error", "not-error", "other-rich"}
# Results that state nothing about a position: the entry form was refused, the function is random or
# network-bound, or it takes no arguments.
DOCUMENTED_SKIPS = {"refused", "skipped", "no-args", "invalid-form", "beyond-arity", "beyond-cap", "ptg"}
# The highest position the sweep tried, per function.
SWEEP_CAP = 6
SWEEP_CAP_ODD = 9

# Variadic functions with a table entry: (first position of the repeating group, group size). A position past
# the last measured group folds back into it.
VARIADIC: Dict[str, Tuple[int, int]] = {
    "AGGREGATE": (3, 1),
    "AVERAGEIFS": (2, 2),
    "CHOOSE": (2, 1),
    "CHOOSECOLS": (2, 1),
    "CHOOSEROWS": (2, 1),
    "CONCATENATE": (1, 1),
    "COUNTIFS": (1, 2),
    "GETPIVOTDATA": (3, 2),
    "IFS": (1, 2),
    "IMPRODUCT": (1, 1),
    "IMSUM": (1, 1),
    "MAP": (1, 1),
    "MAXIFS": (2, 2),
    "MINIFS": (2, 2),
    "MODE": (1, 1),
    "MODE.MULT": (1, 1),
    "MODE.SNGL": (1, 1),
    "NPV": (2, 1),
    "SUBTOTAL": (2, 1),
    "SUMIFS": (2, 2),
    "SUMPRODUCT": (1, 1),
    "SWITCH": (2, 2),
    "TEXTJOIN": (3, 1),
}
# Functions the sweep took to its cap whose arity is fixed nonetheless.
FIXED_AT_CAP = {
    "BETA.DIST",
    "ENCODEURL",
    "FILTERXML",
    "FORECAST.ETS",
    "FORECAST.ETS.CONFINT",
    "FORECAST.ETS.STAT",
    "GROUPBY",
    "IPMT",
    "JIS",
    "PIVOTBY",
    "PPMT",
    "RATE",
    "TEXTAFTER",
    "TEXTBEFORE",
    "TEXTSPLIT",
    "VDB",
    "XLOOKUP",
}

# Ids of functions with no measured id: they never carry iftab, so the id is only a key.
SYNTHETIC_ID_BASE = 0x8000
MAX_POSITION = 16


def fail(message: str) -> None:
    print(f"gen_rich_value_iftab: {message}", file=sys.stderr)
    sys.exit(1)


def load_catalog() -> List[str]:
    names = []
    for line in CATALOG_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            names.append(line)
    return names


def load_func_id_table() -> Dict[str, Tuple[int, bool]]:
    """Name -> (id, variadic) from the XLSB function-id table."""
    text = FUNC_ID_TABLE_PATH.read_text(encoding="utf-8")
    rows: Dict[str, Tuple[int, bool]] = {}
    for m in re.finditer(r'\{(\d+), "([A-Z0-9._]+)", (\d+), (\w+), (true|false)\}', text):
        rows[m.group(2)] = (int(m.group(1)), m.group(4) == "kVariadicMax")
    if not rows:
        fail(f"no rows parsed from {FUNC_ID_TABLE_PATH}")
    return rows


def load_records(path: Path) -> list:
    return json.loads(path.read_text(encoding="utf-8"))


class Entry:
    def __init__(self, name: str) -> None:
        self.name = name
        self.id: Optional[int] = None
        self.iftab_mask = 0
        self.plain_mask = 0
        self.max_measured = 0
        self.at_cap = False
        self.tail_start = 0
        self.group = 0
        self.results: Dict[int, str] = {}

    def has_bits(self) -> bool:
        return self.iftab_mask != 0 or self.plain_mask != 0

    def classify(self, position: int) -> str:
        """`iftab`, `plain` or `default` for a 1-based position, folding a variadic tail."""
        if self.group != 0 and position >= self.tail_start + self.group:
            position = self.tail_start + (position - self.tail_start) % self.group
        if position < 1 or position > MAX_POSITION:
            return "default"
        bit = 1 << (position - 1)
        if self.iftab_mask & bit:
            return "iftab"
        if self.plain_mask & bit:
            return "plain"
        return "default"


def record_class(record: dict) -> str:
    if record["result"] == "iftab":
        return "iftab"
    if record["result"] == "plain-error" and record.get("shown") == "#VALUE!":
        return "plain"
    return "default"


def build_entries(records: list, func_ids: Dict[str, Tuple[int, bool]]) -> Dict[str, Entry]:
    entries: Dict[str, Entry] = {}
    for r in records:
        name = r["fn"]
        e = entries.setdefault(name, Entry(name))
        measured_id = r.get("ftab_or_internal_id")
        if measured_id is not None:
            if e.id is not None and e.id != measured_id:
                fail(f"{name}: two ids {e.id} and {measured_id}")
            e.id = measured_id
        result = r["result"]
        if result not in MEASURED and result not in DOCUMENTED_SKIPS:
            fail(f"{name}: unknown result {result!r}")
        if result == "beyond-cap":
            e.at_cap = True
        pos = r.get("argpos")
        if not isinstance(pos, int) or result not in MEASURED:
            continue
        if pos > MAX_POSITION:
            fail(f"{name}: position {pos} past the mask width")
        if pos in e.results and e.results[pos] != record_class(r):
            fail(f"{name}: position {pos} measured twice with different results")
        e.results[pos] = record_class(r)
        e.max_measured = max(e.max_measured, pos)
        cap = SWEEP_CAP_ODD if name.startswith(("ODDF", "ODDL")) else SWEEP_CAP
        if pos >= cap:
            e.at_cap = True
        bit = 1 << (pos - 1)
        if result == "iftab":
            if r.get("iftab") != measured_id:
                fail(f"{name}: iftab {r.get('iftab')} differs from the function id {measured_id}")
            e.iftab_mask |= bit
        elif record_class(r) == "plain":
            e.plain_mask |= bit
    for name, e in entries.items():
        if name in func_ids:
            ftab_id = func_ids[name][0]
            if e.id is not None and e.id != ftab_id:
                fail(f"{name}: measured id {e.id} differs from the ftab id {ftab_id}")
            e.id = ftab_id
        if e.iftab_mask and e.id is None:
            fail(f"{name}: iftab measured without an id")
    return entries


def assign_groups(entries: Dict[str, Entry], func_ids: Dict[str, Tuple[int, bool]]) -> None:
    for name, (first, group) in VARIADIC.items():
        if name not in entries:
            fail(f"variadic list names {name}, which the sweep does not carry")
        if name in func_ids and not func_ids[name][1]:
            fail(f"variadic list names {name}, which the function-id table gives a fixed arity")
    for name, e in entries.items():
        if not e.has_bits():
            continue
        if name in VARIADIC:
            first, group = VARIADIC[name]
            complete = (e.max_measured - first + 1) // group
            if complete <= 0:
                fail(f"{name}: no complete repeating group measured")
            e.tail_start = first + (complete - 1) * group
            e.group = group
            continue
        variadic_in_table = name in func_ids and func_ids[name][1]
        if variadic_in_table or (name not in func_ids and e.at_cap and name not in FIXED_AT_CAP):
            fail(f"{name}: variadic or measured to the cap, but in neither VARIADIC nor FIXED_AT_CAP")


def check_tail(entries: Dict[str, Entry], tail: list) -> None:
    for r in tail:
        e = entries.get(r["fn"])
        if e is None:
            fail(f"tail record names {r['fn']}, which the sweep does not carry")
        got = e.classify(r["argpos"]) if e.has_bits() else "default"
        want = record_class(r)
        if got != want:
            fail(f"tail {r['fn']} position {r['argpos']} ({r.get('form')}): table says {got}, Excel stored {want}")


def assign_synthetic_ids(entries: Dict[str, Entry]) -> None:
    next_id = SYNTHETIC_ID_BASE
    for name in sorted(entries):
        e = entries[name]
        if e.has_bits() and e.id is None:
            e.id = next_id
            next_id += 1


def c_string(text: str) -> str:
    return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'


def render_inc(entries: Dict[str, Entry], func_ids: Dict[str, Tuple[int, bool]]) -> str:
    rows = sorted((e for e in entries.values() if e.has_bits()), key=lambda e: e.id)
    ids = [e.id for e in rows]
    if len(ids) != len(set(ids)):
        fail("duplicate ids in the table")
    if any(e.id > 0xFFFF for e in rows):
        fail("an id does not fit 16 bits")
    first_only = [e for e in rows if (e.iftab_mask, e.plain_mask, e.tail_start, e.group) == (1, 0, 0, 0)]
    full = [e for e in rows if e not in first_only]
    names = sorted((e for e in rows if e.name not in func_ids), key=lambda e: e.name)
    offsets = []
    at = 0
    for e in names:
        offsets.append(at)
        at += len(e.name.encode("utf-8")) + 1
    if not names or at > 0xFFFF:
        fail("the name blob is empty or does not fit 16-bit offsets")
    out = [
        "// Generated by tools/catalog/gen_rich_value_iftab.py from tools/catalog/data/rich_value_iftab_*.json;",
        "// do not edit.",
        "",
        "// Bit p-1 of a mask is argument position p; a position in neither mask is {errorType 2, propagated 1}.",
        "// A variadic row folds a position at or past tail_start + group back into its last measured group. Ids at",
        "// or above 0x8000 are keys only: those functions never carry iftab.",
        "",
        "// Ids whose row is {iftab_mask 0x0001, plain_mask 0, tail_start 0, group 0}, sorted.",
        "constexpr std::uint16_t kRichValueIftabFirstOnlyIds[] = {",
    ]
    for e in first_only:
        out.append(f"    {e.id}U,  // {e.name}")
    out.append("};")
    out.append("")
    out.append("// Every other function with a measured iftab or plain position, sorted by id.")
    out.append("constexpr RichValueIftabRow kRichValueIftabRows[] = {")
    for e in full:
        out.append(
            f"    {{{e.id}U, 0x{e.iftab_mask:04X}U, 0x{e.plain_mask:04X}U, {e.tail_start}U, {e.group}U}},  // {e.name}"
        )
    out.append("};")
    out.append("")
    out.append("// Functions of either table the XLSB function-id table does not name, sorted by name: NUL-joined")
    out.append("// names, each one's offset into them, and its id.")
    out.append("constexpr char kRichValueIftabNameText[] =")
    out.extend(f'    {c_string(e.name)[:-1]}\\0"' for e in names)
    out[-1] += ";"
    out.append("constexpr std::uint16_t kRichValueIftabNameOffsets[] = {")
    for e, offset in zip(names, offsets):
        out.append(f"    {offset}U,  // {e.name}")
    out.append("};")
    out.append("constexpr std::uint16_t kRichValueIftabNameIds[] = {")
    for e in names:
        out.append(f"    {e.id}U,  // {e.name}")
    out.append("};")
    return "\n".join(out) + "\n"


def cell_row(record: dict) -> int:
    m = re.search(r'<c r="[A-Z]+(\d+)"', record.get("raw") or "")
    return int(m.group(1)) if m else 10


def render_rows(records: list) -> str:
    out = [
        "// Generated by tools/catalog/gen_rich_value_iftab.py from tools/catalog/data/rich_value_iftab_*.json;",
        "// do not edit.",
        "",
        "// Every measured sweep formula: the function, the 1-based argument position holding C5, the 1-based row of",
        "// the column-E cell Excel evaluated it in, the entered formula and the value Excel showed.",
        "constexpr IftabSweepRow kIftabSweepRows[] = {",
    ]
    for r in records:
        pos = r.get("argpos")
        if not isinstance(pos, int) or r["result"] not in MEASURED or not r.get("formula"):
            continue
        shown = r.get("shown") or ""
        out.append(f"    {{{c_string(r['fn'])}, {pos}U, {cell_row(r)}U, {c_string(r['formula'])}, {c_string(shown)}}},")
    out.append("};")
    return "\n".join(out) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--check", action="store_true", help="exit 1 when an output is stale")
    args = parser.parse_args()

    records: list = []
    for path in SWEEP_PATHS:
        records.extend(load_records(path))
    func_ids = load_func_id_table()
    swept = {r["fn"] for r in records}
    missing = [name for name in load_catalog() if name not in swept]
    if missing:
        fail("catalog functions missing from the sweep: " + ", ".join(missing))

    entries = build_entries(records, func_ids)
    assign_groups(entries, func_ids)
    check_tail(entries, load_records(TAIL_PATH))
    assign_synthetic_ids(entries)

    outputs = {OUT_INC: render_inc(entries, func_ids), OUT_ROWS: render_rows(records)}
    stale = [path for path, text in outputs.items() if not path.exists() or path.read_text(encoding="utf-8") != text]
    if args.check:
        for path in stale:
            print(f"stale: {path.relative_to(REPO_ROOT)}", file=sys.stderr)
        return 1 if stale else 0
    for path in stale:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(outputs[path], encoding="utf-8")
        print(f"wrote {path.relative_to(REPO_ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
