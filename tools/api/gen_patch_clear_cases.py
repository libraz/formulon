#!/usr/bin/env python3
"""Generate the shared patch clear-case table from ``tools/api/names.yaml``.

Every record or typedef entry with ``clearable: [field, ...]`` yields one row
of ``tools/api/patch_clear_cases.json``, keyed by the record's JS type name:

    {"<Record>": {"clearable": [...], "non_clearable": [...]}}

``clearable`` is the ledger's list; ``non_clearable`` every other field of the
record's interface in ``src/wasm/formulon.d.ts``. Field names are snake_case
(the C and Python spelling; JS takes the camelCase of each). The binding tests
read the table to check that clearing a ``non_clearable`` field (``null``,
``None``, ``*_op = CLEAR``) is refused and clearing a ``clearable`` one is not.

Modes:
  (default)  print the table
  --write    write tools/api/patch_clear_cases.json
  --check    exit 1 when the file differs from the table, or is missing
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Sequence

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT / "tools" / "dev"))

from binding_drift.api_conventions_bindings import JsSurface  # noqa: E402

LEDGER_PATH = REPO_ROOT / "tools" / "api" / "names.yaml"
WASM_DTS = REPO_ROOT / "src" / "wasm" / "formulon.d.ts"
OUTPUT = REPO_ROOT / "tools" / "api" / "patch_clear_cases.json"


def load_ledger(path: Path = LEDGER_PATH) -> dict:
    import yaml  # type: ignore

    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or data.get("schema") != 1 or not isinstance(data.get("entries"), dict):
        raise SystemExit(f"gen_patch_clear_cases: {path} is not a `schema: 1` ledger")
    return data


def snake_case(name: str) -> str:
    return re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", name).lower()


def cases(data: dict, dts_text: str) -> Dict[str, Dict[str, List[str]]]:
    """`{record: {clearable, non_clearable}}` for every entry with `clearable`."""
    decls = JsSurface(dts_text).decls
    out: Dict[str, Dict[str, List[str]]] = {}
    for key, entry in sorted(data["entries"].items()):
        if not isinstance(entry, dict) or "clearable" not in entry:
            continue
        name = entry.get("js")
        decl = decls.get(name) if isinstance(name, str) else None
        if decl is None or decl.kind != "interface":
            raise SystemExit(f"gen_patch_clear_cases: `{key}` has no `js` interface in {WASM_DTS.name}")
        fields = [snake_case(m.name) for m in decl.members if not m.is_method]
        clearable = list(entry["clearable"])
        unknown = sorted(set(clearable) - set(fields))
        if unknown:
            raise SystemExit(f"gen_patch_clear_cases: `{key}` lists fields {name} lacks: {', '.join(unknown)}")
        out[name] = {
            "clearable": sorted(clearable),
            "non_clearable": sorted(f for f in dict.fromkeys(fields) if f not in clearable),
        }
    return out


def render(table: Dict[str, Dict[str, List[str]]]) -> str:
    return json.dumps(table, indent=2, sort_keys=True) + "\n"


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--ledger", type=Path, default=LEDGER_PATH)
    parser.add_argument("--dts", type=Path, default=WASM_DTS)
    parser.add_argument("--output", type=Path, default=OUTPUT)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--write", action="store_true", help="write the clear-case table")
    mode.add_argument("--check", action="store_true", help="fail when the clear-case table is stale")
    args = parser.parse_args(argv)
    text = render(cases(load_ledger(args.ledger), args.dts.read_text(encoding="utf-8")))
    if args.write:
        args.output.write_text(text, encoding="utf-8")
        return 0
    if args.check:
        current = args.output.read_text(encoding="utf-8") if args.output.is_file() else None
        if current != text:
            print(f"gen_patch_clear_cases: {args.output.name} differs from the ledger; run with --write")
            return 1
        print(f"gen_patch_clear_cases: {args.output.name} matches the ledger")
        return 0
    sys.stdout.write(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
