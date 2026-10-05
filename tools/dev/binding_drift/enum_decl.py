"""Enumerator extraction from C, C++ and TypeScript enum bodies."""

from __future__ import annotations

import re
from typing import List, Optional

# ---------------------------------------------------------------------------
# Check 5: src/wasm/formulon.d.ts `export enum` ordinals <-> their C/C++
#          source enum. embind never registers these as a real `enum_<T>`
#          (that would cost WASM size for values only ever crossing the
#          boundary as plain numbers); the `.d.ts` copy is therefore the
#          sole place the JS-visible ordinals live, and nothing enforces
#          that it still matches the source enum after a reorder / insert.
# ---------------------------------------------------------------------------

_ENUM_BODY_COMMENT_RE = re.compile(r"//[^\n]*|/\*.*?\*/", re.DOTALL)
_ENUM_MEMBER_RE = re.compile(r"([A-Za-z_][A-Za-z0-9_]*)\s*(?:=\s*(-?(?:0[xX][0-9a-fA-F]+|\d+)))?\s*$")


def _parse_enum_body(body: str) -> Optional[List[int]]:
    """Parses a brace-delimited enumerator list into its ordinal values.

    Handles both explicit `Name = N` members and implicit sequential
    values (each unset member is one more than the previous value, C/C++/TS
    enum semantics). Returns `None` if a member's initializer is not a
    literal integer (e.g. it references another constant) -- the caller
    should skip comparison in that case rather than mis-flag a drift.
    """
    body = _ENUM_BODY_COMMENT_RE.sub("", body)
    values: List[int] = []
    next_value = 0
    for part in body.split(","):
        part = part.strip()
        if not part:
            continue
        match = _ENUM_MEMBER_RE.match(part)
        if not match:
            return None
        literal = match.group(2)
        if literal is not None:
            next_value = int(literal, 0)
        values.append(next_value)
        next_value += 1
    return values


def _extract_c_typedef_enum(text: str, type_name: str) -> Optional[List[int]]:
    # `[^}]*` (rather than a non-greedy `.*?`) deliberately cannot cross a
    # `}` boundary, so a search starting at an *earlier*, unrelated
    # `typedef enum { ... }` block cannot skip past its own close brace
    # and accidentally splice in a later block that happens to close with
    # this `type_name`.
    match = re.search(r"typedef enum\s*\{([^}]*)\}\s*" + re.escape(type_name) + r"\s*;", text, re.DOTALL)
    if not match:
        return None
    return _parse_enum_body(match.group(1))


def _cpp_enum_class_body(text: str, enum_name: str) -> Optional[str]:
    """The brace-delimited enumerator list of a scoped enum, or `None`."""
    match = re.search(
        r"enum class\s+" + re.escape(enum_name) + r"\s*(?::\s*[\w:]+)?\s*\{([^}]*)\}\s*;", text, re.DOTALL
    )
    return match.group(1) if match else None


def _extract_cpp_enum_class(text: str, enum_name: str) -> Optional[List[int]]:
    body = _cpp_enum_class_body(text, enum_name)
    if body is None:
        return None
    return _parse_enum_body(body)


def _extract_ts_enum(text: str, enum_name: str) -> Optional[List[int]]:
    match = re.search(r"export (?:const )?enum\s+" + re.escape(enum_name) + r"\s*\{([^}]*)\}", text, re.DOTALL)
    if not match:
        return None
    return _parse_enum_body(match.group(1))


def _parse_ts_named_enum(text: str, name: str) -> Optional[dict]:
    """Reads a `.d.ts` ordinal table in either idiom used across the two
    declaration files: a TS `export enum` or an `export const` whose type
    is a `Readonly<{...}>` literal."""
    match = re.search(r"export enum\s+" + re.escape(name) + r"\s*\{([^}]*)\}", text, re.DOTALL)
    if match is None:
        match = re.search(r"export const\s+" + re.escape(name) + r"\s*:\s*Readonly<\{([^}]*)\}>", text, re.DOTALL)
    if match is None:
        return None
    body = _ENUM_BODY_COMMENT_RE.sub("", match.group(1))
    members: dict = {}
    next_value = 0
    for part in body.replace(";", ",").split(","):
        part = part.strip()
        if not part:
            continue
        member = re.match(r"([A-Za-z_]\w*)\s*(?:[:=]\s*(-?\d+))?$", part)
        if member is None:
            return None
        if member.group(2) is not None:
            next_value = int(member.group(2))
        members[member.group(1)] = next_value
        next_value += 1
    return members


def _cpp_enum_class_values(text: str, enum_name: str) -> Optional[dict[str, int]]:
    """Parses a scoped enum into a name -> value map.

    The value-only sibling `_extract_cpp_enum_class` compares ordinal
    sequences; this one is for the checks that need to look a member up by
    name. Implicit members take one more than the previous value, as in C++.
    Returns `None` when the enum is absent or a member's initializer is not a
    literal integer, so the caller reports the parse failure rather than a
    false drift.
    """
    raw_body = _cpp_enum_class_body(text, enum_name)
    if raw_body is None:
        return None
    values: dict[str, int] = {}
    next_value = 0
    for part in _ENUM_BODY_COMMENT_RE.sub("", raw_body).split(","):
        part = part.strip()
        if not part:
            continue
        member = _ENUM_MEMBER_RE.match(part)
        if not member:
            return None
        literal = member.group(2)
        if literal is not None:
            next_value = int(literal, 0)
        values[member.group(1)] = next_value
        next_value += 1
    return values
