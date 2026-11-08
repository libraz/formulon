"""Drift check: the C `fm_parse_error_code_t` mirrors `parser::ParseErrorCode`."""

from __future__ import annotations

import re
from typing import Dict, List, Optional

from .enum_decl import _ENUM_BODY_COMMENT_RE, _ENUM_MEMBER_RE, _cpp_enum_class_values
from .surface_files import CAPI_HEADER, REPO_ROOT, _read

PARSE_ERROR_H = REPO_ROOT / "src" / "parser" / "parse_error.h"

_C_TYPE = "fm_parse_error_code_t"
_C_PREFIX = "FM_PARSE_ERROR_CODE_"


def _c_typedef_enum_values(text: str, type_name: str) -> Optional[Dict[str, int]]:
    """Name -> value map of a C `typedef enum { ... } type_name;`, or `None`."""
    match = re.search(r"typedef enum\s*\{([^}]*)\}\s*" + re.escape(type_name) + r"\s*;", text, re.DOTALL)
    if match is None:
        return None
    values: Dict[str, int] = {}
    next_value = 0
    for part in _ENUM_BODY_COMMENT_RE.sub("", match.group(1)).split(","):
        part = part.strip()
        if not part:
            continue
        member = _ENUM_MEMBER_RE.match(part)
        if member is None:
            return None
        if member.group(2) is not None:
            next_value = int(member.group(2), 0)
        values[member.group(1)] = next_value
        next_value += 1
    return values


def c_member_name(cpp_name: str) -> str:
    """`ExpectedRParenOrComma` -> `FM_PARSE_ERROR_CODE_EXPECTED_RPAREN_OR_COMMA`."""
    return _C_PREFIX + re.sub(r"(?<=[a-z0-9])(?=[A-Z])", "_", cpp_name).upper()


def compare_parse_error_codes(cpp: Dict[str, int], c: Dict[str, int]) -> List[str]:
    """Every C++ member has a C member of the derived name and the same value, and nothing else."""
    problems: List[str] = []
    expected = {c_member_name(name): value for name, value in cpp.items()}
    for name, value in sorted(expected.items(), key=lambda item: item[1]):
        if name not in c:
            problems.append(f"parse-error-codes: {_C_TYPE} lacks {name} (= {value})")
        elif c[name] != value:
            problems.append(f"parse-error-codes: {name} is {c[name]} in {_C_TYPE}, {value} in parser::ParseErrorCode")
    for name in sorted(set(c) - set(expected)):
        problems.append(f"parse-error-codes: {name} = {c[name]} has no parser::ParseErrorCode member")
    return problems


def check_parse_error_codes() -> List[str]:
    """`fm_parse_error_code_t` in the C header matches `parser::ParseErrorCode` by name and value."""
    # The member comments quote formulas such as `={1,2;3}`, whose braces would end the body early.
    cpp = _cpp_enum_class_values(_ENUM_BODY_COMMENT_RE.sub("", _read(PARSE_ERROR_H)), "ParseErrorCode")
    if cpp is None:
        return [f"parse-error-codes: cannot parse enum class ParseErrorCode in {PARSE_ERROR_H}"]
    c = _c_typedef_enum_values(_read(CAPI_HEADER), _C_TYPE)
    if c is None:
        return [f"parse-error-codes: cannot parse {_C_TYPE} in {CAPI_HEADER}"]
    return compare_parse_error_codes(cpp, c)
