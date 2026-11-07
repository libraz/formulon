#!/usr/bin/env python3
"""Rewrite empty initialisers of C ABI input records to FM_STRUCT_INIT(T).

The record set is every struct in the C header whose first member is
`uint32_t struct_size`; frozen value types carry no such member and are
therefore never touched. Empty initialisers (`T v{};`, `T v = {};`, `T{}`) are
rewritten. Non-empty brace initialisers and array initialisers are only
listed, because their values shift by one field once `struct_size` leads.

Usage: struct_init_rewrite.py [--check | --write] [--header H] PATH...
Without --check/--write the pending rewrites and manual items are listed.
Exit status: 1 for --check when rewrites are pending, 2 on usage errors.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

DEFAULT_HEADER = Path(__file__).resolve().parents[2] / "src" / "c_api" / "formulon_c.h"
SOURCE_SUFFIXES = {".c", ".cc", ".cpp", ".cxx", ".h", ".hpp", ".mm"}

_COMMENT_OR_STRING = re.compile(r"//[^\n]*|/\*.*?\*/|\"(?:\\.|[^\"\\\n])*\"|'(?:\\.|[^'\\\n])*'", re.DOTALL)
_STRUCT_OPEN = re.compile(r"typedef\s+struct\s*\w*\s*\{")
_FIRST_MEMBER = re.compile(r"\s*uint32_t\s+struct_size\s*;")
_TYPEDEF_NAME = re.compile(r"\s*(\w+)\s*;")
_TYPE_TOKEN = re.compile(r"\bfm_\w+_t\b")


@dataclass(frozen=True)
class Manual:
    """A non-empty initialiser that needs a human."""

    line: int
    kind: str
    text: str


def mask(text: str) -> str:
    """Blank comments and string literals, preserving offsets and newlines."""
    return _COMMENT_OR_STRING.sub(lambda m: re.sub(r"[^\n]", " ", m.group(0)), text)


def record_types(header_text: str) -> set[str]:
    """Return the struct typedef names whose first member is struct_size."""
    masked = mask(header_text)
    found: set[str] = set()
    for m in _STRUCT_OPEN.finditer(masked):
        if not _FIRST_MEMBER.match(masked, m.end()):
            continue
        close = _match_brace(masked, m.end() - 1)
        name = _TYPEDEF_NAME.match(masked, close + 1) if close >= 0 else None
        if name:
            found.add(name.group(1))
    return found


def _match_brace(masked: str, open_pos: int) -> int:
    depth = 0
    for i in range(open_pos, len(masked)):
        if masked[i] == "{":
            depth += 1
        elif masked[i] == "}":
            depth -= 1
            if depth == 0:
                return i
    return -1


def scan(text: str, types: set[str]) -> tuple[str, list[Manual]]:
    """Return (rewritten text, manual items) for one source file."""
    masked = mask(text)
    edits: list[tuple[int, int, str]] = []
    manual: list[Manual] = []
    pos = 0
    while True:
        m = _TYPE_TOKEN.search(masked, pos)
        if not m:
            break
        pos = m.end()
        t = m.group(0)
        if t not in types:
            continue
        # Temporary: T{...}
        tail = re.compile(r"\s*\{").match(masked, m.end())
        if tail and masked[m.end() : tail.end()].strip() == "{" and not masked[m.start() - 1 : m.start()].isalnum():
            open_pos = tail.end() - 1
            close = _match_brace(masked, open_pos)
            if close < 0:
                continue
            if masked[open_pos + 1 : close].strip() == "":
                edits.append((m.start(), close + 1, f"{t} FM_STRUCT_INIT({t})"))
            else:
                manual.append(_manual(text, masked, m.start(), close, "temporary"))
            pos = close + 1
            continue
        # Declaration: T name [array] [=] {...}
        d = re.compile(r"(\s+)(\w+)(\s*\[[^\]]*\])?\s*(=\s*)?\{").match(masked, m.end())
        if not d:
            continue
        name, array = d.group(2), d.group(3)
        open_pos = d.end() - 1
        close = _match_brace(masked, open_pos)
        if close < 0:
            continue
        empty = masked[open_pos + 1 : close].strip() == ""
        if array:
            manual.append(_manual(text, masked, m.start(), close, "array"))
        elif empty:
            edits.append((d.start(2), close + 1, f"{name} = FM_STRUCT_INIT({t})"))
        else:
            manual.append(
                _manual(
                    text,
                    masked,
                    m.start(),
                    close,
                    "designated" if re.match(r"\s*\.", masked[open_pos + 1 :]) else "positional",
                )
            )
        pos = close + 1
    out = text
    for start, end, repl in sorted(edits, reverse=True):
        out = out[:start] + repl + out[end:]
    return out, manual


def _manual(text: str, masked: str, start: int, close: int, kind: str) -> Manual:
    line = text.count("\n", 0, start) + 1
    snippet = " ".join(text[start : close + 1].split())
    if len(snippet) > 100:
        snippet = snippet[:97] + "..."
    return Manual(line, kind, snippet)


def iter_files(paths: list[Path]):
    for p in paths:
        if p.is_dir():
            yield from sorted(f for f in p.rglob("*") if f.is_file() and f.suffix in SOURCE_SUFFIXES)
        else:
            yield p


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("paths", nargs="+", type=Path)
    ap.add_argument("--header", type=Path, default=DEFAULT_HEADER, help="C header defining the records")
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("--check", action="store_true", help="exit 1 if rewrites are pending")
    mode.add_argument("--write", action="store_true", help="apply the rewrites in place")
    args = ap.parse_args(argv)

    types = record_types(args.header.read_text(encoding="utf-8"))
    if not types:
        print(f"error: no struct_size records found in {args.header}", file=sys.stderr)
        return 2

    pending = 0
    manual_total = 0
    for f in iter_files(args.paths):
        text = f.read_text(encoding="utf-8")
        new, manual = scan(text, types)
        changed = sum(1 for _ in re.finditer(r"FM_STRUCT_INIT\(", new)) - sum(
            1 for _ in re.finditer(r"FM_STRUCT_INIT\(", text)
        )
        if new != text:
            pending += changed
            if args.write:
                f.write_text(new, encoding="utf-8")
            else:
                print(f"{f}: {changed} rewrite(s) pending")
        for item in manual:
            manual_total += 1
            print(f"{f}:{item.line}: manual ({item.kind}): {item.text}")
    print(f"{'rewritten' if args.write else 'pending'}: {pending}; manual: {manual_total}", file=sys.stderr)
    return 1 if args.check and pending else 0


if __name__ == "__main__":
    sys.exit(main())
