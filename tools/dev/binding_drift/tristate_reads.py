"""Every binding read of an `fm_tristate_t` field decodes it explicitly, never by truthiness.

`FM_TRISTATE_FALSE` is 1, so `bool(field)`, `field != 0` or `if field:` reads FALSE as true. The fields are the
`int32_t` members src/c_api/formulon_c.h marks `/* fm_tristate_t */` (a name that is a plain `int32_t` elsewhere in
the header is left out, since the read cannot be told apart by name). A read passes when it is compared with a
`FM_TRISTATE_*` / `TRISTATE_*` constant, handed to a tristate helper, taken by address, or assigned on (a write).
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable, List, Optional, Set, Tuple

from .surface_files import REPO_ROOT

C_HEADER = REPO_ROOT / "src" / "c_api" / "formulon_c.h"
_GLUE_DIRS = (REPO_ROOT / "src" / "wasm" / "parts", REPO_ROOT / "src" / "node_addon" / "parts")
_PYTHON_DIR = REPO_ROOT / "packages" / "python" / "formulon"
# Helpers that decode or carry a tristate.
_HELPERS = ("tristate_flag", "JsTristate", "js_tristate", "tristate", "_tri")
_DECODED = re.compile(r"\s*(?:==|!=)\s*(?:S\.)?(?:FM_)?TRISTATE_")
_WRITE = re.compile(r"\s*=(?!=)")


def tristate_fields(header: str) -> Set[str]:
    """Names declared only as `fm_tristate_t` fields."""
    tristate = set(re.findall(r"int32_t\s+([a-z_][a-z0-9_]*)\s*;\s*/\*\s*fm_tristate_t", header))
    plain = set(re.findall(r"int32_t\s+([a-z_][a-z0-9_]*)\s*;(?!\s*/\*\s*fm_tristate_t)", header))
    return tristate - plain


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


def scan(text: str, fields: Set[str], python: bool) -> Iterable[Tuple[int, str]]:
    """`(line, text)` of each truthiness read of a tristate field."""
    if not fields:
        return
    names = "|".join(sorted(fields))
    pattern = re.compile(rf"""\[\s*["']({names})["']\s*\]""" if python else rf"(?:\.|->)({names})\b")
    for number, line in enumerate(text.splitlines(), 1):
        code = line.split("#", 1)[0] if python else line.split("//", 1)[0]
        if not python:
            # A C++ string literal ("cfvo.gte") names a path, not a read.
            code = re.sub(r'"(?:[^"\\]|\\.)*"', lambda m: '"' + " " * (len(m.group(0)) - 2) + '"', code)
        for match in pattern.finditer(code):
            if not _passes(code, match.start(), match.end(), match.group(1)):
                yield number, line.strip()


def check_tristate_reads(header: Optional[str] = None, roots: Optional[List[Tuple[Path, bool]]] = None) -> List[str]:
    """Each tristate field read in the bindings decodes explicitly."""
    fields = tristate_fields(C_HEADER.read_text(encoding="utf-8") if header is None else header)
    if roots is None:
        roots = [(d, False) for d in _GLUE_DIRS] + [(_PYTHON_DIR, True)]
    problems: List[str] = []
    for root, python in roots:
        for path in sorted(root.glob("*.py" if python else "*.c*")):
            for number, line in scan(path.read_text(encoding="utf-8"), fields, python):
                rel = path.relative_to(REPO_ROOT) if path.is_relative_to(REPO_ROOT) else path
                problems.append(f"tristate-reads: {rel}:{number}: `{line}` reads an fm_tristate_t by truthiness")
    return problems
