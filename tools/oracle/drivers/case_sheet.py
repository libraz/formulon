"""Case-to-worksheet primitives shared by the Mac and Windows drivers.

Sheet naming and addressing, merge application, error-cell triggers, and the
print-block helpers that do not touch a platform bridge.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional


def split_sheet_qualified_addr(key: str) -> "tuple[Optional[str], str]":
    """Splits a setup-key like ``"Sheet2!A1"`` into ``(sheet, a1)``.

    Returns ``(None, key)`` when the key is a bare A1 address (no ``!``).
    Single-quoted sheet names are unquoted: ``"'Sheet One'!A1"`` ->
    ``("Sheet One", "A1")``. Escaped quotes (``''`` inside a quoted
    name) collapse to a single quote per Excel's convention. The split
    is on the LAST ``!`` so a future stray ``!`` inside a quoted sheet
    name does not confuse it.
    """

    if "!" not in key:
        return None, key
    bang = key.rfind("!")
    sheet_part = key[:bang]
    addr_part = key[bang + 1 :]
    if sheet_part.startswith("'") and sheet_part.endswith("'") and len(sheet_part) >= 2:
        sheet_part = sheet_part[1:-1].replace("''", "'")
    return sheet_part, addr_part


def get_or_add_sheet(wb, name: str):
    """Returns the sheet whose display name matches ``name`` (case-insensitive),
    adding it at the end if absent.
    """

    target = name.casefold()
    for sht in wb.sheets:
        if sht.name.casefold() == target:
            return sht
    return wb.sheets.add(name=name, after=wb.sheets[len(wb.sheets) - 1])


def apply_merges(sht, merges: List[str]) -> None:
    """Apply case-declared inclusive A1 merge ranges on ``sht``."""

    for ref in merges:
        sht.range(ref).merge()


SHEET_FORBIDDEN = set("\\/?*[]:")


def sanitize_sheet_name(name: str) -> str:
    """Strips characters Excel disallows in sheet names and trims length."""

    cleaned = "".join("_" if c in SHEET_FORBIDDEN else c for c in name)
    return cleaned[:24] or "case"


ERROR_TRIGGERS = {
    "#DIV/0!": "=1/0",
    "#NAME?": "=NONEXISTENT_FUNC()",
    "#VALUE!": '=VALUE("x")',
    "#NUM!": "=SQRT(-1)",
    "#N/A": "=NA()",
    "#REF!": "=OFFSET(A1,-1,-1)",
    "#NULL!": "=A1 B1",
}


def error_trigger(code: str) -> str:
    return ERROR_TRIGGERS.get(code, '=VALUE("x")')


# Excel `XlPageOrientation` constants.
_XL_PORTRAIT = 1
_XL_LANDSCAPE = 2

PRINT_ORIENTATIONS = {
    "portrait": _XL_PORTRAIT,
    "landscape": _XL_LANDSCAPE,
}


def resolve_print_sheet(wb, print_spec: Dict[str, Any]):
    """Returns the worksheet the `print` block names, or raises."""

    sheet_name = print_spec.get("sheet")
    if not isinstance(sheet_name, str) or not sheet_name:
        raise RuntimeError("print block missing string 'sheet'")
    target = sheet_name.casefold()
    for sht in wb.sheets:
        if sht.name.casefold() == target:
            return sht
    raise RuntimeError(f"print 'sheet' names an unknown sheet {sheet_name!r}")


def normalise_print_area(area: str) -> str:
    """Strips ``$`` anchors and ``Sheet!`` qualifiers from a print area.

    Excel reports ``PrintArea`` fully qualified and anchored
    (``Sheet1!$A$1:$H$80``); the C++ engine compares against a bare
    ``A1:H80`` form, so both sides normalise the same way.
    """

    out_parts: List[str] = []
    for part in area.split(","):
        token = part.strip()
        if not token:
            continue
        bang = token.rfind("!")
        if bang != -1:
            token = token[bang + 1 :]
        token = token.replace("$", "")
        out_parts.append(token)
    return ",".join(out_parts)
