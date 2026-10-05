"""Drift checks on the released C ABI baseline and the header error-code documentation."""

from __future__ import annotations

import re
from pathlib import Path
from typing import List

from .c_header import c_abi_declarations
from .enum_decl import _cpp_enum_class_values
from .surface_files import C_ABI_BASELINE, C_ABI_BREAKS, CAPI_HEADER, ERROR_H, REPO_ROOT, _read


def _parse_abi_manifest(path: Path) -> List[str]:
    """Non-comment, non-blank lines of a baseline / ledger file."""
    lines = []
    for raw in _read(path).splitlines():
        line = raw.strip()
        if line and not line.startswith("#"):
            lines.append(line)
    return lines


_ABI_BREAK_RE = re.compile(r"^(removed|signature|retyped)\s+(fm_[A-Za-z0-9_]+)\s*::\s*(\S.*)$")
_ABI_BASELINE_RE = re.compile(r"^(.+?)\s+(fm_[A-Za-z0-9_]+)\((.*)\)$")


def check_abi_baseline() -> List[str]:
    """The released C ABI surface is still declared, or its loss is recorded.

    A deletion or rename of a base entry point is invisible to every other
    check here: five of the eight base/`_ex` families have no in-repo caller
    at all, so the C API tests recompile against the new header and pass. The
    only party that notices is a third-party consumer compiled against the
    released header, which no test can stand in for -- hence a pinned copy of
    that header's surface plus an explicit ledger of what was broken on
    purpose.

    Signature comparison is by declared parameter *type*, so it catches an
    added, dropped or retyped parameter. It does NOT catch a by-value struct
    that keeps its name and changes size; that is a calling convention change
    the `static_assert(sizeof(...))` tripwires in the C ABI tests pin instead.
    """
    problems: List[str] = []

    baseline: dict[str, tuple[str, List[str]]] = {}
    for line in _parse_abi_manifest(C_ABI_BASELINE):
        match = _ABI_BASELINE_RE.match(line)
        if match is None:
            problems.append(f"abi-baseline: unparsable baseline line: {line!r}")
            continue
        ret, name, params = match.groups()
        parsed = [] if params.strip() in ("", "void") else [p.strip() for p in params.split(",")]
        baseline[name] = (ret.strip(), parsed)

    ledger: dict[str, str] = {}
    for line in _parse_abi_manifest(C_ABI_BREAKS):
        match = _ABI_BREAK_RE.match(line)
        if match is None:
            problems.append(f"abi-baseline: unparsable ledger line: {line!r}")
            continue
        kind, name, _reason = match.groups()
        if name in ledger:
            problems.append(f"abi-baseline: {name} is listed twice in {C_ABI_BREAKS.name}")
        ledger[name] = kind

    current = c_abi_declarations(_read(CAPI_HEADER))

    for name, declared in sorted(baseline.items()):
        recorded = ledger.get(name)
        if name not in current:
            if recorded != "removed":
                problems.append(
                    f"abi-baseline: {name} shipped in the released header but is gone from "
                    f"{CAPI_HEADER.relative_to(REPO_ROOT)}. A consumer built against the release "
                    f"loses it at link time. Record the break as `removed {name} :: <reason>` in "
                    f"{C_ABI_BREAKS.relative_to(REPO_ROOT)} and carry it into the CHANGELOG's "
                    "Removed section, or restore the entry point."
                )
            continue
        if current[name] != declared:
            if recorded not in ("signature", "retyped"):
                problems.append(
                    f"abi-baseline: {name} changed signature since the release.\n"
                    f"  released: {declared[0]} {name}({', '.join(declared[1]) or 'void'})\n"
                    f"  current:  {current[name][0]} {name}({', '.join(current[name][1]) or 'void'})\n"
                    f"  Record it in {C_ABI_BREAKS.relative_to(REPO_ROOT)} as `signature` (breaks a "
                    "stale caller) or `retyped` (invisible to a C caller and to the calling "
                    "convention), with the reason."
                )
            continue
        # Declared and unchanged: any ledger entry naming it is stale.
        if recorded is not None:
            problems.append(
                f"abi-baseline: {C_ABI_BREAKS.relative_to(REPO_ROOT)} records {name} as "
                f"`{recorded}`, but it is declared unchanged from the release. Drop the line -- a "
                "ledger that outlives its break stops meaning anything."
            )

    for name in sorted(set(ledger) - set(baseline)):
        problems.append(
            f"abi-baseline: {C_ABI_BREAKS.relative_to(REPO_ROOT)} records {name}, which was never "
            f"in the released surface. Only an entry point a consumer could have compiled against "
            "can be broken; drop the line."
        )

    return problems


# A doc comment quotes a status code as `` `kName` (N) ``, and the number may
# wrap onto the next comment line, so the separator spans a `*` continuation
# marker as well as plain spaces.
_HEADER_ERROR_CODE_RE = re.compile(r"`?(k[A-Z][A-Za-z0-9_]*)`?(?:[ \t]|\n[ \t]*\*)*\((\d+)\)")


def check_header_error_codes() -> List[str]:
    """Every numeric status code quoted in the C header matches its enumerator.

    `fm_status_t` is a bare `int32_t`, so the numbers the header's doc
    comments print next to an enumerator name are the whole contract a
    binding author reading only the header can act on. A comparison written
    against a stale number never matches, and mis-classifies whatever code
    actually holds that value.
    """
    problems: List[str] = []

    codes = _cpp_enum_class_values(_read(ERROR_H), "FormulonErrorCode")
    if codes is None:
        return [
            f"header-error-codes: could not parse `FormulonErrorCode` out of {ERROR_H.relative_to(REPO_ROOT)}; "
            "the enum moved or gained a non-literal initializer"
        ]

    header = _read(CAPI_HEADER)
    for match in _HEADER_ERROR_CODE_RE.finditer(header):
        name, quoted = match.group(1), match.group(2)
        line_number = header.count("\n", 0, match.start()) + 1
        site = f"{CAPI_HEADER.name}:{line_number}"
        if name not in codes:
            problems.append(
                f"header-error-codes: {site} documents `{name}` ({quoted}), which names no enumerator of "
                f"`FormulonErrorCode` in {ERROR_H.relative_to(REPO_ROOT)}"
            )
            continue
        if int(quoted) != codes[name]:
            problems.append(
                f"header-error-codes: {site} documents `{name}` as ({quoted}), but "
                f"{ERROR_H.relative_to(REPO_ROOT)} defines it as {codes[name]}"
            )

    return problems
