"""Engine errors a call reports lead with the C parameter or field they are about: `<field> <constraint>`.

The bindings label an engine error `<Class.method> <message>` and turn a leading C path into the public path, so
a message that leads with a prose word (`invalid name`) or an operation prefix (`insert_cells: ...`) would
surface as a fake path or a C spelling. This check walks every literal an engine error is built from under src/
(bindings excluded) and requires the first token to be a C parameter or struct field the header declares; a
message about no argument at all leads with a capital (`Normal style cannot be removed`). Errors that describe
the content of a file being read or written (`kIo*`, `kSecurity*`, out of memory) keep their
`<part or record>: <problem>` form and are not checked.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable, List, Optional, Set, Tuple

from .surface_files import REPO_ROOT as ROOT

C_HEADER = ROOT / "src" / "c_api" / "formulon_c.h"
_SOURCE_DIR = ROOT / "src"
# Binding layers word their own messages; the engine corpus is everything else.
_EXCLUDED_PARTS = frozenset({"wasm", "node_addon"})
_CODES = r"(?:FormulonErrorCode::)?k(?!Io[A-Z]|Security[A-Z]|OutOfMemory\b)[A-Z][A-Za-z0-9]*"
# Helpers whose first argument is the message of an argument or state error.
_WRAPPERS = ("bad_arg", "invalid", "Invalid", "invalid_table", "refuse_write", "invalid_argument", "refuse")
_LITERAL = r'"((?:[^"\\]|\\.)*)"'
_DIRECT = re.compile(r"make_error\(\s*" + _CODES + r"\s*,\s*(?:std::string\(\s*)?" + _LITERAL)
_OP_PREFIX = re.compile(r"make_error\(\s*" + _CODES + r'\s*,\s*std::string\(\s*\w+\s*\)\s*\+\s*"\s*:')
_WRAPPED = re.compile(r"(?<![\w.:>])(?:" + "|".join(_WRAPPERS) + r")\(\s*" + _LITERAL)
_PATH = re.compile(r"[a-z][a-z0-9_]*(?:\[\d+\])*(?:\.[a-z][a-z0-9_]*(?:\[\d+\])*)*")


def c_field_names(header: str) -> Set[str]:
    """Every identifier the header declares as a function parameter or a struct field."""
    text = re.sub(r"/\*.*?\*/", "", header, flags=re.S)
    text = re.sub(r"//[^\n]*", "", text)
    return set(re.findall(r"\b([a-z][a-z0-9_]*)\s*(?:\[[^\]]*\])?\s*[;,)]", text))


def message_problem(message: str, fields: Set[str]) -> Optional[str]:
    """Why `message` breaks the grammar, or `None` when it leads with a known field or with a capital."""
    token = message.split(" ", 1)[0]
    if not token or token[0].isupper():
        return None
    if token.endswith(":"):
        return "leads with an operation prefix"
    if not _PATH.fullmatch(token):
        return "does not lead with a field"
    root = re.split(r"[.\[]", token)[0]
    return None if root in fields else f"leads with `{root}`, which is no C parameter or field"


def scan(text: str, fields: Set[str]) -> Iterable[Tuple[int, str, str]]:
    """`(line, message, problem)` for each grammar violation in one source text."""
    for match in _OP_PREFIX.finditer(text):
        yield text.count("\n", 0, match.start()) + 1, match.group(0), "leads with an operation prefix"
    for pattern in (_DIRECT, _WRAPPED):
        for match in pattern.finditer(text):
            problem = message_problem(match.group(1), fields)
            if problem is not None:
                yield text.count("\n", 0, match.start()) + 1, match.group(1), problem


def check_engine_messages(source_dir: Path = _SOURCE_DIR, header: Optional[str] = None) -> List[str]:
    """Each engine argument-error message under `source_dir` leads with a C parameter or field."""
    fields = c_field_names(C_HEADER.read_text(encoding="utf-8") if header is None else header)
    problems: List[str] = []
    for path in sorted(source_dir.rglob("*")):
        if path.suffix not in (".cpp", ".h") or _EXCLUDED_PARTS.intersection(path.relative_to(source_dir).parts):
            continue
        text = path.read_text(encoding="utf-8", errors="replace")
        for line, message, problem in scan(text, fields):
            rel = path.relative_to(ROOT) if path.is_relative_to(ROOT) else path
            problems.append(f"engine-messages: {rel}:{line}: `{message}` {problem}")
    return problems
