"""Every binding read of an `fm_tristate_t` field decodes it explicitly, never by truthiness.

`FM_TRISTATE_FALSE` is 1, so `bool(field)`, `field != 0` or `if field:` reads FALSE as true. The fields are the
`int32_t` struct members src/c_api/formulon_c.h marks `/* fm_tristate_t */`. A name that is a plain `int32_t` in
another struct is checked only where the enclosing function names a struct holding it as a tristate (`fm_x_t`, or
`S.X` in Python) and none holding it plain. A read passes when it is compared with a
`FM_TRISTATE_*` / `TRISTATE_*` constant, handed to a tristate helper, taken by address, or assigned on (a write).

The marker is the header's single convention for an enum held in an `int32_t`, so it is checked too: a field the C
implementation (src/c_api) assigns a tristate to, compares with `FM_TRISTATE_*` or decodes with `tristate_to_bool`
must be marked in some struct.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

from .surface_files import REPO_ROOT

C_HEADER = REPO_ROOT / "src" / "c_api" / "formulon_c.h"
_C_IMPL_DIRS = (REPO_ROOT / "src" / "c_api", REPO_ROOT / "src" / "c_api" / "parts")
_GLUE_DIRS = (REPO_ROOT / "src" / "wasm" / "parts", REPO_ROOT / "src" / "node_addon" / "parts")
_PYTHON_DIR = REPO_ROOT / "packages" / "python" / "formulon"
# Helpers that decode or carry a tristate.
_HELPERS = ("tristate_flag", "JsTristate", "js_tristate", "TriToJs", "tristate", "_tri")
_DECODED = re.compile(r"\s*(?:==|!=)\s*(?:S\.)?(?:FM_)?TRISTATE_")
_WRITE = re.compile(r"\s*=(?!=)")


_STRUCT = re.compile(r"typedef\s+struct(?:\s+\w+)?\s*\{(.*?)\}\s*(fm_\w+)\s*;", re.S)
_INT_FIELD = re.compile(r"int32_t\s+([a-z_][a-z0-9_]*)\s*;(\s*/\*\s*fm_tristate_t)?")
_PY_STRUCT = re.compile(r"""\b([A-Z][A-Z0-9_]*)\s*=\s*Struct\(\s*["'](fm_\w+)["']""")


def struct_fields(header: str) -> Dict[str, Dict[str, bool]]:
    """`{struct: {int32 field: is tristate}}` for every struct of the header."""
    return {
        name: {m.group(1): m.group(2) is not None for m in _INT_FIELD.finditer(body)}
        for body, name in _STRUCT.findall(header)
    }


def tristate_fields(header: str) -> Set[str]:
    """Names declared only as `fm_tristate_t` fields."""
    structs = struct_fields(header).values()
    tristate = {f for fields in structs for f, tri in fields.items() if tri}
    plain = {f for fields in structs for f, tri in fields.items() if not tri}
    return tristate - plain


_IMPL_TRISTATE = (
    re.compile(r"(?:\.|->)([a-z_][a-z0-9_]*)\s*=\s*[^;=]*(?:\bFM_TRISTATE_|\btristate_of\()"),
    re.compile(r"(?:\.|->)([a-z_][a-z0-9_]*)\s*[!=]=\s*FM_TRISTATE_"),
    re.compile(r"FM_TRISTATE_[A-Z]+\s*[!=]=\s*[\w.>-]*?(?:\.|->)([a-z_][a-z0-9_]*)\b"),
    re.compile(r"\btristate_to_bool\(\s*[\w.>-]*?(?:\.|->)([a-z_][a-z0-9_]*)\s*\)"),
)


def unmarked_tristate_fields(header: str, sources: Iterable[Tuple[str, str]]) -> List[str]:
    """`file:line: field` for each C-implementation tristate use of a field no struct marks `fm_tristate_t`."""
    marked = {f for fields in struct_fields(header).values() for f, tri in fields.items() if tri}
    problems: List[str] = []
    for name, text in sources:
        for number, line in enumerate(text.splitlines(), 1):
            code = line.split("//", 1)[0]
            for pattern in _IMPL_TRISTATE:
                for field in pattern.findall(code):
                    if field not in marked:
                        problems.append(f"{name}:{number}: {field}")
    return problems


def _scopes(lines: List[str], python: bool) -> List[Tuple[int, int]]:
    """`(start, end)` line spans of the top-level functions (Python: every `def`)."""
    start_re = re.compile(r"^\s*(?:async\s+)?def\s") if python else re.compile(r"^[A-Za-z_]")
    starts = [i for i, line in enumerate(lines) if start_re.match(line)]
    return [(s, (starts[k + 1] if k + 1 < len(starts) else len(lines))) for k, s in enumerate(starts)]


def _scope_structs(text: str, python: bool, py_structs: Dict[str, str]) -> List[Set[str]]:
    """Per line, the C struct names its enclosing function mentions."""
    lines = text.splitlines()
    per_line: List[Set[str]] = [set() for _ in lines]
    for start, end in _scopes(lines, python):
        body = "\n".join(lines[start:end])
        if python:
            names = {py_structs[n] for n in re.findall(r"\bS\.([A-Z][A-Z0-9_]*)\b", body) if n in py_structs}
        else:
            names = set()
        names |= set(re.findall(r"\b(fm_\w+_t)\b", body))
        for i in range(start, end):
            per_line[i] = names
    return per_line


def _passes(line: str, start: int, end: int, field: str) -> bool:
    if _DECODED.match(line[end:]) or _WRITE.match(line[end:]):
        return True
    before = line[:start].rstrip()
    # Taken by address (`&p.allow_blank`), or copied into the same tristate field (`brk.manual = span.manual`).
    if re.search(r"&\s*[\w.>-]*$", before) or re.search(rf"[\w.>-]+(?:\.|->){field}\s*=\s*[\w.>-]*$", before):
        return True
    # `x == FM_TRISTATE_TRUE` written constant-first.
    if re.search(r"(?:FM_)?TRISTATE_[A-Z]+\s*(?:==|!=)\s*$", before):
        return True
    return any(re.search(rf"\b{h}\(\s*$", before) or re.search(rf"\b{h}\([^()]*$", before) for h in _HELPERS)


def scan(
    text: str,
    fields: Set[str],
    python: bool,
    structs: Optional[Dict[str, Dict[str, bool]]] = None,
    py_structs: Optional[Dict[str, str]] = None,
) -> Iterable[Tuple[int, str]]:
    """`(line, text)` of each truthiness read of a tristate field.

    `fields` are checked everywhere. With `structs`, a name that is a tristate in some structs and plain in others is
    checked only where the enclosing function names a struct holding it as a tristate and none holding it plain.
    """
    structs = structs or {}
    ambiguous = {f for s in structs.values() for f, tri in s.items() if tri} - fields
    names = "|".join(sorted(fields | ambiguous))
    if not names:
        return
    pattern = re.compile(rf"""\[\s*["']({names})["']\s*\]""" if python else rf"(?:\.|->)({names})\b")
    in_scope = _scope_structs(text, python, py_structs or {}) if ambiguous else []
    for number, line in enumerate(text.splitlines(), 1):
        code = line.split("#", 1)[0] if python else line.split("//", 1)[0]
        if not python:
            # A C++ string literal ("cfvo.gte") names a path, not a read.
            code = re.sub(r'"(?:[^"\\]|\\.)*"', lambda m: '"' + " " * (len(m.group(0)) - 2) + '"', code)
        for match in pattern.finditer(code):
            field = match.group(1)
            if field in ambiguous:
                holders = [structs[s][field] for s in in_scope[number - 1] if field in structs.get(s, {})]
                if not holders or not all(holders):
                    continue
            if not _passes(code, match.start(), match.end(), field):
                yield number, line.strip()


def check_tristate_reads(header: Optional[str] = None, roots: Optional[List[Tuple[Path, bool]]] = None) -> List[str]:
    """Each tristate field read in the bindings decodes explicitly."""
    header_is_default = header is None
    header = C_HEADER.read_text(encoding="utf-8") if header is None else header
    fields = tristate_fields(header)
    structs = struct_fields(header)
    if roots is None:
        roots = [(d, False) for d in _GLUE_DIRS] + [(_PYTHON_DIR, True)]
    problems: List[str] = []
    if header_is_default:
        impl = [
            (str(path.relative_to(REPO_ROOT)), path.read_text(encoding="utf-8"))
            for d in _C_IMPL_DIRS
            for path in sorted(d.glob("*.c*"))
        ]
        problems += [
            f"tristate-reads: {where} holds an fm_tristate_t but no struct field in formulon_c.h is marked "
            "`/* fm_tristate_t */`"
            for where in unmarked_tristate_fields(header, impl)
        ]
    for root, python in roots:
        layouts = root / "_structs.py"
        py_structs = (
            dict(_PY_STRUCT.findall(layouts.read_text(encoding="utf-8"))) if python and layouts.exists() else {}
        )
        for path in sorted(root.glob("*.py" if python else "*.c*")):
            for number, line in scan(path.read_text(encoding="utf-8"), fields, python, structs, py_structs):
                rel = path.relative_to(REPO_ROOT) if path.is_relative_to(REPO_ROOT) else path
                problems.append(f"tristate-reads: {rel}:{number}: `{line}` reads an fm_tristate_t by truthiness")
    return problems
