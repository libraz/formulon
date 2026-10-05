"""Parsing of the C ABI header: wasm32 struct layouts, function declarations and struct fields."""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import List, NamedTuple, Set

# ---------------------------------------------------------------------------
# Check 1b: Python WASM32 POD layouts <-> C ABI header.
# ---------------------------------------------------------------------------


def _wasm32_layout(fields: List[tuple[str, int, int]]) -> tuple[dict[str, int], int, int]:
    """Lays `fields` out in sequence, returning `(offsets, size, alignment)`."""
    offsets: dict[str, int] = {}
    offset = 0
    max_align = 1
    for name, size, align in fields:
        offset = (offset + align - 1) // align * align
        offsets[name] = offset
        offset += size
        max_align = max(max_align, align)
    return offsets, (offset + max_align - 1) // max_align * max_align, max_align


# wasm32 (ILP32) size and alignment for the leaf types the C ABI structs are
# built from. These are properties of the target ABI rather than anything the
# header states, which is why they are the only layout facts tabulated here:
# every type that has a declaration to read -- struct, union, array -- is
# measured from that declaration instead of quoted. A literal for an embedded
# struct would keep asserting its old size after the struct changed, which is
# the one thing this check exists to catch.
_WASM32_SCALARS = {
    "uint8_t": (1, 1),
    "uint16_t": (2, 2),
    "uint32_t": (4, 4),
    "int32_t": (4, 4),
    "uint64_t": (8, 8),
    "int64_t": (8, 8),
    "double": (8, 8),
    "size_t": (4, 4),
}
_WASM32_POINTER = (4, 4)
# A C enum takes `int` representation. The header pins each enumerator's
# ordinal but says nothing about the storage width, so this is not derivable
# either; *which* names are enums is read off the header.
_WASM32_ENUM = (4, 4)

_C_STRUCT_RE = re.compile(r"typedef\s+struct\s*\{(.*?)\}\s*(fm_[A-Za-z0-9_]+)\s*;", re.S)
_C_ENUM_RE = re.compile(r"typedef\s+enum\s*\{[^}]*\}\s*([A-Za-z_][A-Za-z0-9_]*)\s*;")
# An anonymous `union { ... } name;` member, matched before the enclosing body
# is split on `;` so the union's own members are not mistaken for the parent's.
_C_UNION_MEMBER_RE = re.compile(r"union\s*\{(.*?)\}\s*([A-Za-z_][A-Za-z0-9_]*)\s*;", re.S)
_C_DECLARATION_RE = re.compile(r"(.+?)\s+([A-Za-z_][A-Za-z0-9_]*)(?:\[(\d+)\])?")


class _LayoutError(Exception):
    """A C declaration the wasm32 layout resolver cannot measure."""


class _Member(NamedTuple):
    """One measured struct member. `ctype` is empty for an anonymous union."""

    name: str
    ctype: str
    size: int
    align: int


class _Wasm32Layouts:
    """Measures the C ABI's structs against the wasm32 ABI, from the header.

    Struct members are resolved recursively, so an embedded POD is measured
    rather than quoted: changing a field of `fm_color_spec` moves every struct
    that embeds it, and the Python side is checked against the new offsets.
    """

    def __init__(self, header: str) -> None:
        self._structs = {name: body for body, name in _C_STRUCT_RE.findall(header)}
        self._enums = set(_C_ENUM_RE.findall(header))
        self._members: dict[str, List[_Member]] = {}
        self._pending: Set[str] = set()

    def is_struct(self, ctype: str) -> bool:
        return ctype in self._structs

    def members(self, struct_name: str) -> List[_Member]:
        """Members of `struct_name` in declaration order, each measured."""
        cached = self._members.get(struct_name)
        if cached is not None:
            return cached
        body = self._structs.get(struct_name)
        if body is None:
            raise _LayoutError(f"{struct_name} missing from C header")
        if struct_name in self._pending:
            raise _LayoutError(f"{struct_name} embeds itself")
        self._pending.add(struct_name)
        try:
            measured = self._parse_body(struct_name, body)
        finally:
            self._pending.discard(struct_name)
        self._members[struct_name] = measured
        return measured

    def extent(self, ctype: str) -> tuple[int, int]:
        """`(size, alignment)` of one member's declared type."""
        if "*" in ctype:
            return _WASM32_POINTER
        ctype = ctype.replace("const ", "").strip()
        if ctype in self._structs:
            _, size, align = _wasm32_layout([(m.name, m.size, m.align) for m in self.members(ctype)])
            return size, align
        if ctype in _WASM32_SCALARS:
            return _WASM32_SCALARS[ctype]
        if ctype in self._enums:
            return _WASM32_ENUM
        raise _LayoutError(f"unknown field type {ctype!r}")

    def _parse_body(self, owner: str, body: str) -> List[_Member]:
        measured: List[_Member] = []
        position = 0
        for union in _C_UNION_MEMBER_RE.finditer(body):
            measured.extend(self._parse_declarations(owner, body[position : union.start()]))
            size, align = self._union_extent(owner, union.group(1))
            measured.append(_Member(union.group(2), "", size, align))
            position = union.end()
        measured.extend(self._parse_declarations(owner, body[position:]))
        return measured

    def _parse_declarations(self, owner: str, text: str) -> List[_Member]:
        measured: List[_Member] = []
        for declaration in text.split(";"):
            declaration = " ".join(declaration.split())
            if not declaration:
                continue
            match = _C_DECLARATION_RE.fullmatch(declaration)
            if not match:
                raise _LayoutError(f"cannot parse {owner}: {declaration!r}")
            ctype, name, count_text = match.groups()
            size, align = self.extent(ctype)
            measured.append(_Member(name, ctype.replace("const ", "").strip(), size * int(count_text or "1"), align))
        return measured

    def _union_extent(self, owner: str, body: str) -> tuple[int, int]:
        """A union overlays its members: widest size, strictest alignment."""
        size = 0
        align = 1
        for member in self._parse_declarations(owner, body):
            size = max(size, member.size)
            align = max(align, member.align)
        return (size + align - 1) // align * align, align


# ---------------------------------------------------------------------------
# Check 1c: Python `LIB.fm_*` call sites <-> the C header's declarations.
#
# `python-exports` matches call *names* only. Nothing compares what a call
# passes against what the entry point declares, so a parameter added to an
# existing `fm_*` function -- the shape a 1.0-frozen ABI is most likely to
# grow -- leaves the Python side calling the old signature and is caught only
# if a runtime binding test happens to exercise that path.
# ---------------------------------------------------------------------------

_C_FUNCTION_DECL_RE = re.compile(
    r"FM_API\s+([A-Za-z_][A-Za-z0-9_ *]*?)\b(fm_[A-Za-z0-9_]+)\s*\(([^;]*?)\)\s*;",
    re.S,
)


def _split_c_params(text: str) -> List[str]:
    """Splits a parameter list on its top-level commas."""
    parts: List[str] = []
    depth = 0
    current = ""
    for char in text:
        if char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        if char == "," and depth == 0:
            parts.append(current.strip())
            current = ""
        else:
            current += char
    parts.append(current.strip())
    return parts


def _c_function_params(header: str) -> dict[str, List[str]]:
    """Maps each `FM_API` entry point to its declared parameter types.

    The parameter *name* is stripped, so the values are types in declaration
    order (`["fm_workbook_t*", "size_t", "fm_value_t*"]`). Every parameter in
    this header is named, so dropping the last whitespace-delimited token is
    unambiguous.
    """
    return {name: params for name, (_, params) in c_abi_declarations(header).items()}


def c_abi_declarations(header: str) -> dict[str, tuple[str, List[str]]]:
    """Maps each `FM_API` entry point to its `(return type, parameter types)`.

    Public because `gen_c_abi_baseline.py` writes the released-surface file
    with it: generating the baseline through the same parser the check reads
    it back with keeps a formatting difference from masquerading as a break.
    """
    declarations: dict[str, tuple[str, List[str]]] = {}
    for ret, name, params in _C_FUNCTION_DECL_RE.findall(header):
        ret = " ".join(ret.split())
        params = " ".join(params.split())
        if params in ("", "void"):
            declarations[name] = (ret, [])
            continue
        declarations[name] = (ret, [" ".join(param.split()[:-1]) for param in _split_c_params(params)])
    return declarations


def _c_struct_fields(header: str, struct_name: str, source: Path) -> List[str]:
    blocks = {name: body for body, name in _C_STRUCT_RE.findall(header)}
    body = blocks.get(struct_name)
    if body is None:
        print(f"check_binding_drift: struct {struct_name!r} not found in {source}", file=sys.stderr)
        sys.exit(2)
    fields: List[str] = []
    for declaration in body.split(";"):
        declaration = " ".join(declaration.split())
        if not declaration:
            continue
        match = re.fullmatch(r"(.+?)\s+([A-Za-z_][A-Za-z0-9_]*)(?:\[(\d+)\])?", declaration)
        if match and not match.group(2).startswith("_pad"):
            fields.append(match.group(2))
    return fields
