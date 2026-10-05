"""Parsing of TypeScript declaration files and the JS constant tables behind them."""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Optional, Set

from .surface_files import _read


def _extract_braced_block(text: str, start: int) -> str:
    """Returns the substring between `start` (just after an opening `{`)
    and its matching closing `}`, using a simple depth counter.

    Safe for this codebase's `.d.ts` files: the only braces inside an
    interface body are either the one being matched or self-contained
    `{ ... }` pairs inside single-line JSDoc comments (e.g. `` `{ status,
    index }` ``), which do not change the net depth.
    """
    depth = 1
    i = start
    while i < len(text) and depth > 0:
        if text[i] == "{":
            depth += 1
        elif text[i] == "}":
            depth -= 1
        i += 1
    return text[start : i - 1]


def _find_interface_body(text: str, name: str, source: Path) -> str:
    match = re.search(r"export interface %s\s*\{" % re.escape(name), text)
    if not match:
        print(f"check_binding_drift: interface {name!r} not found in {source}", file=sys.stderr)
        sys.exit(2)
    return _extract_braced_block(text, match.end())


# Matches a method signature's leading `name(` at the start of a line
# (after indentation); this also catches multi-line signatures whose
# parameter list wraps, since only the opening token is needed. Plain
# data fields (`name: Type;`) do not match -- no `(` follows the name.
_TS_METHOD_RE = re.compile(r"^\s*([A-Za-z_$][A-Za-z0-9_$]*)\s*\(", re.MULTILINE)


def _extract_ts_methods(body: str) -> Set[str]:
    return set(_TS_METHOD_RE.findall(body))


# Top-level members of an exported interface are indented exactly two spaces;
# anything deeper belongs to a nested object type and is compared as part of
# its parent's field text, not on its own.
_TS_MEMBER_NAME_RE = re.compile(r"^ {2}([A-Za-z_$][A-Za-z0-9_$]*)\s*\(", re.MULTILINE)
_TS_MEMBER_FIELD_RE = re.compile(r"^ {2}(?:readonly\s+)?([A-Za-z_$][A-Za-z0-9_$]*)(\??)\s*:", re.MULTILINE)


def _ts_return_types(body: str) -> dict[str, str]:
    """Maps each method in an interface body to its declared return type.

    The parameter list is skipped by bracket balancing so a wrapped or
    generic signature is read the same as a single-line one; whitespace in
    the return type is collapsed so formatting differences do not register
    as drift.
    """
    out: dict[str, str] = {}
    for match in _TS_MEMBER_NAME_RE.finditer(body):
        open_paren = body.index("(", match.start())
        depth = 0
        close_paren = -1
        for i in range(open_paren, len(body)):
            char = body[i]
            if char in "(<[":
                depth += 1
            elif char in ")>]":
                depth -= 1
                if depth == 0 and char == ")":
                    close_paren = i
                    break
        if close_paren < 0:
            continue
        rest = body[close_paren + 1 :]
        if ";" not in rest:
            continue
        declared = rest[: rest.index(";")].strip()
        if declared.startswith(":"):
            declared = declared[1:]
        out[match.group(1)] = " ".join(declared.split())
    return out


def _ts_interface_declaration(text: str, name: str, source: Path) -> tuple[str, str]:
    """Returns an interface's `extends` clause (may be empty) and its body.

    Unlike `_find_interface_body` this tolerates `extends`, which matters
    here: a record type that inherits its fields declares none of them
    locally, so the base list is part of the shape being compared.
    """
    match = re.search(r"^export interface %s\b([^{]*)\{" % re.escape(name), text, re.MULTILINE)
    if not match:
        print(f"check_binding_drift: interface {name!r} not found in {source}", file=sys.stderr)
        sys.exit(2)
    return " ".join(match.group(1).split()), _extract_braced_block(text, match.end())


def _ts_field_set(body: str) -> Set[str]:
    """Field names of an interface body, each suffixed with `?` when optional.

    Optionality is part of the compared shape on purpose: "present only on
    success" is how both declaration files spell a payload key the runtime
    drops, so a required declaration on one side and an optional one on the
    other is a real disagreement about what a caller may dereference.
    """
    return {name + optional for name, optional in _TS_MEMBER_FIELD_RE.findall(body)}


# ---------------------------------------------------------------------------
# Check 2d: pure-JS helpers shipped by both npm packages.
#
# A helper with no native entry point behind it has to be written out once per
# package, because the two packages share no module. Nothing then holds the
# copies together, and nothing notices when one package documents a helper it
# does not ship -- which is how the WASM declaration file came to send its
# readers to a package that is not published. Both halves are checked here:
# the helper is exported and declared by both packages, and the two
# implementations are the same text.
# ---------------------------------------------------------------------------


def _js_exported_function_source(text: str, name: str) -> Optional[str]:
    """Whitespace-normalised body of `export function <name>(...) { ... }`."""
    match = re.search(r"^export function %s\s*\(" % re.escape(name), text, re.MULTILINE)
    if not match:
        return None
    body_open = text.index("{", match.end())
    return " ".join(_extract_braced_block(text, body_open + 1).split())


# Frozen ordinal tables are the only way a JS consumer can name a value
# that crosses the boundary as a plain number, so both published ESM entry
# points have to carry the same ones. The WASM `.d.ts` above is already
# pinned to the C/C++ enums, which makes it the single source the two
# `index.mjs` files are measured against -- comparing them only to each
# other would let a shared mistake pass.
_JS_TABLE_RE = re.compile(r"^export const (\w+) = Object\.freeze\(\{(.*?)\}\);", re.DOTALL | re.MULTILINE)
_JS_SCALAR_RE = re.compile(r"^export const (\w+) = (-?\d+);", re.MULTILINE)
_JS_MEMBER_RE = re.compile(r"(\w+)\s*:\s*(-?\d+)")


def _parse_js_constants(path: Path) -> dict:
    """Reads an `index.mjs`'s exported ordinal tables and scalar constants."""
    text = _read(path)
    out: dict = {}
    for match in _JS_TABLE_RE.finditer(text):
        out[match.group(1)] = {name: int(value) for name, value in _JS_MEMBER_RE.findall(match.group(2))}
    for match in _JS_SCALAR_RE.finditer(text):
        out[match.group(1)] = int(match.group(2))
    return out
