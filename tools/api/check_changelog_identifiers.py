#!/usr/bin/env python3
"""Check that the API identifiers the current CHANGELOG release names exist.

Every identifier inside a code span of the ``[Unreleased]`` section and the
newest released section must be one of:

* a current public name: a token of ``src/wasm/formulon.d.ts``,
  ``packages/npm-native/index.d.ts``, the JavaScript entry modules, the Python
  package or ``src/c_api/formulon_c.h``;
* a name the previous release shipped, from the committed snapshot
  ``tools/api/released/<version>.txt`` (``--snapshot <tag>`` writes it from the
  tagged tree, so the check needs no git history);
* a name in :data:`NOT_API` (not a binding identifier at all).

A cell of a table column that lists former names (header ``Before``,
``Removed ...``, ``... meant`` or ``... (former name)``) must name a previous-release
identifier. So a release note never names an identifier no release shipped. Formula text
(``=...``), XML (``<...>``), paths, glob patterns (``*``), flags (``--x``) and
``_suffix`` shorthand are not checked. Only identifiers that look like API
names are: dotted or called (``a.b``, ``f()``), camelCase, or containing ``_``;
an all-caps token without ``_`` is an Excel function name.

Exit 0 when every identifier resolves; otherwise print ``line: identifier``
per miss and exit 1.
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Set, Tuple

REPO_ROOT = Path(__file__).resolve().parents[2]
CHANGELOG = REPO_ROOT / "CHANGELOG.md"
SURFACE_FILES = (
    "src/wasm/formulon.d.ts",
    "packages/npm-native/index.d.ts",
    "packages/npm/common.mjs",
    "packages/npm-native/index.mjs",
    "src/c_api/formulon_c.h",
)
PYTHON_PACKAGE = "packages/python/formulon"

RELEASED_DIR = REPO_ROOT / "tools" / "api" / "released"
# The files a release's public names are read from; `error.h` holds the status names `fm_status_string` returned.
SNAPSHOT_FILES = SURFACE_FILES + ("packages/npm/index.mjs", "src/utils/error.h")

# Identifiers that are not names of a binding, by what they name.
NOT_API: Dict[str, Set[str]] = {
    "JavaScript built-ins and the embind runtime methods": {
        "JSON",
        "stringify",
        "getFullYear",
        "getSeconds",
        "clone",
        "isDeleted",
    },
    "OOXML elements and attributes": {
        "lumMod",
        "lumOff",
        "containsDate",
        "containsNumber",
        "minDate",
        "maxDate",
        "minValue",
        "maxValue",
        "firstDataCol",
        "firstHeaderRow",
    },
    "Excel function arguments": {"total_depth", "relative_to", "if_not_found"},
    "command-line JSON keys": {"definedNames", "passthroughParts", "defined_names", "passthrough_parts"},
}

IDENT = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
CHAIN = re.compile(r"[A-Za-z_][\w.]*(?:\(\))?")
SPAN = re.compile(r"`([^`]+)`")
FORMER_HEADER = re.compile(r"^(Before|Removed\b.*|.*\bmeant|.*\(former name\))$")


def release_lines(text: str) -> Iterator[Tuple[int, str]]:
    """``(line number, line)`` of ``[Unreleased]`` and the newest released section."""
    released = 0
    for number, line in enumerate(text.splitlines(), start=1):
        if line.startswith("## ["):
            if "Unreleased" not in line:
                released += 1
            if released > 1:
                return
        yield number, line


def checked_text(lines: Iterable[Tuple[int, str]]) -> Iterator[Tuple[int, str, bool]]:
    """``(line number, text, former)`` pieces: table cells carry their column's kind."""
    former_cols: Set[int] = set()
    in_table = False
    for number, line in lines:
        if not line.startswith("|"):
            in_table = False
            yield number, line, False
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if not in_table:
            in_table = True
            former_cols = {i for i, c in enumerate(cells) if FORMER_HEADER.match(c)}
            continue
        if set(line.strip()) <= set("|-: "):
            continue
        for i, cell in enumerate(cells):
            yield number, cell, i in former_cols


def span_identifiers(span: str) -> Iterator[str]:
    """The API-shaped identifiers of one code span."""
    s = span.strip()
    if not s or s[0] in "=<-#\"'" or "/" in s or "*" in s or (" " in s and "(" not in s):
        return
    s = re.sub(r"'[^']*'|\"[^\"]*\"", "", s)
    for chain in CHAIN.findall(s):
        parts = chain.rstrip("()").split(".")
        dotted = len(parts) > 1 or chain.endswith("()")
        for token in parts:
            if not token or token.startswith("_") or token.endswith("_"):
                continue
            if token.isupper() and "_" not in token:
                continue
            if dotted or re.search(r"[a-z][A-Z]|_", token):
                yield token


def previous_release(text: str) -> str:
    """The version of the newest released section that is not the current one."""
    versions = re.findall(r"^## \[(\d+\.\d+\.\d+)\]", text, flags=re.M)
    return versions[1]


def released_names(version: str) -> Set[str]:
    lines = (RELEASED_DIR / f"{version}.txt").read_text(encoding="utf-8").splitlines()
    return {line for line in lines if line and not line.startswith("#")}


def write_snapshot(tag: str) -> None:
    """Write the identifier set of the tagged tree to ``tools/api/released/<version>.txt``."""
    paths = list(SNAPSHOT_FILES)
    listing = subprocess.run(
        ["git", "ls-tree", "--name-only", f"{tag}:{PYTHON_PACKAGE}"],
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
        text=True,
    )
    paths += [f"{PYTHON_PACKAGE}/{n}" for n in listing.stdout.split() if n.endswith(".py")]
    names: Set[str] = set()
    for path in paths:
        shown = subprocess.run(["git", "show", f"{tag}:{path}"], cwd=REPO_ROOT, capture_output=True, text=True)
        if shown.returncode == 0:
            names.update(IDENT.findall(shown.stdout))
    RELEASED_DIR.mkdir(exist_ok=True)
    header = f"# Generated by tools/api/check_changelog_identifiers.py --snapshot {tag}; do not edit.\n"
    body = "\n".join(sorted(names)) + "\n"
    (RELEASED_DIR / f"{tag.lstrip('v')}.txt").write_text(header + body, encoding="utf-8")


def current_names() -> Set[str]:
    names: Set[str] = set()
    files = [REPO_ROOT / f for f in SURFACE_FILES] + sorted((REPO_ROOT / PYTHON_PACKAGE).glob("*.py"))
    for path in files:
        names.update(IDENT.findall(path.read_text(encoding="utf-8")))
    return names


def misses(text: str, current: Set[str], released: Set[str]) -> List[Tuple[int, str, bool]]:
    """``(line, identifier, former column)`` of every identifier that does not resolve."""
    out: List[Tuple[int, str, bool]] = []
    for number, piece, is_former in checked_text(release_lines(text)):
        allowed = released if is_former else current | released
        for span in SPAN.findall(piece):
            out.extend((number, t, is_former) for t in span_identifiers(span) if t not in allowed)
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--snapshot", metavar="TAG", help="write tools/api/released/<version>.txt from a release tag")
    args = parser.parse_args()
    if args.snapshot:
        write_snapshot(args.snapshot)
        return 0
    text = CHANGELOG.read_text(encoding="utf-8")
    current = current_names()
    released = released_names(previous_release(text))
    for group in NOT_API.values():
        current |= group
        released |= group
    found = misses(text, current, released)
    for number, token, is_former in found:
        what = "no previous-release name" if is_former else "neither a current nor a previous-release name"
        print(f"CHANGELOG.md:{number}: `{token}` is {what}")
    return 1 if found else 0


if __name__ == "__main__":
    sys.exit(main())
