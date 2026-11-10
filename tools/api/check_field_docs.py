#!/usr/bin/env python3
"""Check that every field of a public record carries its own doc comment.

The public records are the exported interfaces of ``src/wasm/formulon.d.ts`` and
``packages/npm-native/index.d.ts``, the public ``@dataclass`` classes of
``packages/python/formulon/*.py`` and the ``typedef struct`` records of
``src/c_api/formulon_c.h``. A field is documented when a doc comment sits on the
line directly above it or on its own line: ``/** ... */`` in a declaration file,
``#:`` in Python, ``/** ... */`` or a trailing ``/**< ... */`` / ``/* ... */`` in
the C header. A C record's ``struct_size`` is the versioning field every record
starts with and is not counted.

Exit 0 when every field is documented; otherwise print ``path:line: Record.field``
per undocumented field and exit 1. ``--json`` prints every field instead, with its
surface, record, name, line and doc text (or ``null``).
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterator, List, Optional

REPO_ROOT = Path(__file__).resolve().parents[2]
DTS_FILES = ("src/wasm/formulon.d.ts", "packages/npm-native/index.d.ts")
PYTHON_DIR = "packages/python/formulon"
C_HEADER = "src/c_api/formulon_c.h"


@dataclass
class Field:
    surface: str
    path: str
    line: int
    record: str
    name: str
    doc: Optional[str]


def _jsdoc_above(lines: List[str], index: int) -> Optional[str]:
    """The text of the `/** ... */` block ending on the line above `index`, or None."""
    end = index - 1
    if end < 0 or not lines[end].rstrip().endswith("*/"):
        return None
    start = end
    while start >= 0 and "/*" not in lines[start]:
        start -= 1
    if start < 0 or "/**" not in lines[start]:
        return None
    text = " ".join(lines[start : end + 1])
    text = re.sub(r"/\*\*<?|\*/", " ", text)
    text = re.sub(r"^\s*\*\s?", " ", text, flags=re.M)
    return re.sub(r"\s*\*\s+", " ", re.sub(r"\s+", " ", text)).strip()


def _inline_comment(line: str, marker: str) -> Optional[str]:
    """The comment that follows the declaration on `line`, or None."""
    at = line.find(marker)
    if at < 0:
        return None
    return re.sub(r"\*/\s*$", "", line[at + len(marker) :]).strip(" *<")


def dts_fields(rel: str) -> Iterator[Field]:
    lines = (REPO_ROOT / rel).read_text(encoding="utf-8").splitlines()
    record: Optional[str] = None
    depth = parens = 0
    for index, line in enumerate(lines):
        if record is None:
            match = re.match(r"^export interface (\w+)\b.*\{\s*$", line)
            if match:
                record, depth = match.group(1), 1
            continue
        code = re.sub(r"/\*.*?\*/|//.*$", "", line)
        if depth == 1 and parens == 0 and not line.lstrip().startswith(("*", "/*")):
            match = re.match(r"^\s+(?:readonly\s+)?([A-Za-z_$][\w$]*)\??\s*:", code)
            if match:
                doc = _inline_comment(line, "/**") or _jsdoc_above(lines, index)
                yield Field("dts", rel, index + 1, record, match.group(1), doc)
        depth += code.count("{") - code.count("}")
        parens += code.count("(") - code.count(")")
        if depth <= 0:
            record, parens = None, 0


def python_fields() -> Iterator[Field]:
    for path in sorted((REPO_ROOT / PYTHON_DIR).glob("*.py")):
        rel = str(path.relative_to(REPO_ROOT))
        lines = path.read_text(encoding="utf-8").splitlines()
        for index, line in enumerate(lines):
            match = re.match(r"^class ([A-Za-z]\w*)\b", line)
            if not match or not any(lines[i].startswith("@dataclass") for i in range(max(0, index - 3), index)):
                continue
            record = match.group(1)
            in_string = False
            for body in range(index + 1, len(lines)):
                text = lines[body]
                if text and not text.startswith((" ", "\t")):
                    break
                if re.match(r"^    (def|class|@)\b", text):
                    break
                quotes = text.count('"""') + text.count("'''")
                if in_string or quotes:
                    in_string = in_string != (quotes % 2 == 1)
                    continue
                field = re.match(r"^    ([a-z]\w*)\s*:\s*(?!ClassVar)", text)
                if not field:
                    continue
                above = lines[body - 1].strip()
                doc = _inline_comment(text, "#:") or (above[2:].strip() if above.startswith("#:") else None)
                yield Field("py", rel, body + 1, record, field.group(1), doc)


def c_fields() -> Iterator[Field]:
    lines = (REPO_ROOT / C_HEADER).read_text(encoding="utf-8").splitlines()
    index = 0
    while index < len(lines):
        if not re.match(r"^typedef struct\b[^;]*\{\s*$", lines[index]):
            index += 1
            continue
        depth = 0
        members = []
        end = index
        while end < len(lines):
            code = re.sub(r"/\*.*?\*/|//.*$", "", lines[end])
            if depth == 1 and end > index and not lines[end].lstrip().startswith(("*", "/*")):
                match = re.match(r"^\s+[\w\s\*]+?[\s\*]([a-z_]\w*)\s*(?:\[[^\]]*\])?\s*;", code)
                if match:
                    members.append((end, match.group(1)))
            depth += code.count("{") - code.count("}")
            if depth == 0 and end > index:
                break
            end += 1
        name = re.search(r"\}\s*(fm_\w+)\s*;", lines[end])
        if name:
            for at, member in members:
                if member == "struct_size":
                    continue
                doc = _inline_comment(lines[at], "/*") or _jsdoc_above(lines, at)
                yield Field("c", C_HEADER, at + 1, name.group(1), member, doc)
        index = end + 1


def all_fields() -> List[Field]:
    fields: List[Field] = []
    for rel in DTS_FILES:
        fields += dts_fields(rel)
    fields += python_fields()
    fields += c_fields()
    return fields


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--json", action="store_true", help="print every field as JSON")
    args = parser.parse_args(argv)
    fields = all_fields()
    if args.json:
        json.dump([asdict(f) for f in fields], sys.stdout, indent=1)
        return 0
    missing = [f for f in fields if not f.doc]
    for f in missing:
        print(f"{f.path}:{f.line}: {f.record}.{f.name}")
    return 1 if missing else 0


if __name__ == "__main__":
    sys.exit(main())
