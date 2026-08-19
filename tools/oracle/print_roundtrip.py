#!/usr/bin/env python3
"""Authors the xlsx a print round-trip case is captured against.

The workbook oracle's existing `print` block builds a workbook inside
Excel with `books.add()` and reads back how Excel paginates it. That
answers "how does Excel lay this out", not "does Excel read what we
wrote" -- Formulon's bytes never reach Excel on that path.

A `roundtrip` case closes that gap. This module is the first half: it
drives Formulon's own print-authoring API from the case spec and hands
back the saved xlsx. The Windows driver is the second half; it opens
those bytes with `books.open` and reports what Excel makes of them.

The two halves are deliberately split across processes. Capture runs on
a Windows host (or a WSL2 bridge to one), and shipping the fixture as
bytes in the request payload keeps the Formulon dependency on the side
that already has the repo, rather than requiring a wheel install next to
Excel.

Usage:
    python3 tools/oracle/print_roundtrip.py \\
        tests/oracle/cases_wb/<suite>.case.json --id <case_id> \\
        --out /tmp/<case_id>.xlsx

The module needs `formulon` importable. From a source checkout that is
`PYTHONPATH=packages/python`; `--python-path` prepends a directory for
callers that would rather not set the environment variable.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# A1 cell reference, split into its column letters and row digits.
_A1_RE = re.compile(r"^([A-Za-z]{1,3})([0-9]{1,7})$")


class RoundtripSpecError(Exception):
    """The `roundtrip` block does not describe a workbook we can author."""


def _column_index(letters: str) -> int:
    """Converts A1 column letters to a zero-based index."""

    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    return index - 1


def _split_a1(addr: str) -> Tuple[int, int]:
    """Converts an A1 address to zero-based `(row, col)`."""

    match = _A1_RE.match(addr.strip())
    if match is None:
        raise RoundtripSpecError(f"not an A1 cell address: {addr!r}")
    return int(match.group(2)) - 1, _column_index(match.group(1))


def _write_cell(wb: Any, sheet: int, addr: str, rec: Dict[str, Any]) -> None:
    """Writes one normalised `{kind, value}` record through the binding.

    Mirrors the COM-side `_write_cell` in the Windows driver so a case's
    `sheets` block means the same thing on both halves of the round trip.
    """

    row, col = _split_a1(addr)
    kind = rec.get("kind")
    if kind == "blank":
        wb.set_blank(sheet, row, col)
    elif kind == "number":
        wb.set_number(sheet, row, col, float(rec["value"]))
    elif kind == "bool":
        wb.set_bool(sheet, row, col, bool(rec["value"]))
    elif kind == "text":
        wb.set_text(sheet, row, col, str(rec["value"]))
    elif kind == "formula":
        wb.set_formula(sheet, row, col, str(rec["formula"]))
    else:
        # `error` is deliberately unsupported: the COM side reaches an
        # error value by evaluating a trigger formula, and a round-trip
        # case has no reason to need one. Rejecting is better than
        # silently writing something that is not what the case says.
        raise RoundtripSpecError(f"cell kind {kind!r} is not supported in a roundtrip case")


def _apply_sheets(wb: Any, case: Dict[str, Any]) -> Dict[str, int]:
    """Materialises the `sheets` block, returning name -> sheet index."""

    sheets = case.get("sheets") or {}
    indices: Dict[str, int] = {}
    for position, (name, cells) in enumerate(sheets.items()):
        if position > 0:
            wb.add_sheet(name)
        else:
            wb.rename_sheet(0, name)
        indices[name] = position
        for addr, rec in (cells or {}).items():
            _write_cell(wb, position, addr, rec)
    if not indices:
        indices[wb.sheet_name(0)] = 0
    return indices


def _apply_layout(wb: Any, case: Dict[str, Any]) -> None:
    """Applies `column_widths` / `row_heights` to the first sheet.

    Scoped to sheet 0 for the same reason the COM builder is: the case
    schema carries one layout map per case, not one per sheet.
    """

    for col_key, width in (case.get("column_widths") or {}).items():
        col = _column_index(str(col_key))
        wb.set_column_width(0, col, col, float(width))
    for row_key, height in (case.get("row_heights") or {}).items():
        wb.set_row_height(0, int(row_key) - 1, float(height))


def _apply_print_settings(wb: Any, sheet: int, spec: Dict[str, Any]) -> None:
    """Drives the print-authoring API from the `roundtrip` block.

    Every sub-block is optional and every field inside one is optional:
    the setters take keyword-only patches, so a field the case omits is
    left at whatever the workbook already had. That mirrors what a caller
    authoring a report does, and it keeps a case's golden scoped to the
    settings the case actually states.
    """

    page_setup = spec.get("page_setup")
    if isinstance(page_setup, dict):
        wb.set_page_setup(
            sheet,
            orientation=page_setup.get("orientation"),
            paper_size=page_setup.get("paper_size"),
            scale=page_setup.get("scale"),
            fit_to_width=page_setup.get("fit_to_width"),
            fit_to_height=page_setup.get("fit_to_height"),
            fit_to_page=page_setup.get("fit_to_page"),
        )

    margins = spec.get("page_margins")
    if isinstance(margins, dict):
        wb.set_page_margins(
            sheet,
            left=margins.get("left"),
            right=margins.get("right"),
            top=margins.get("top"),
            bottom=margins.get("bottom"),
            header=margins.get("header"),
            footer=margins.get("footer"),
        )

    options = spec.get("print_options")
    if isinstance(options, dict):
        wb.set_print_options(
            sheet,
            grid_lines=options.get("grid_lines"),
            headings=options.get("headings"),
            horizontal_centered=options.get("horizontal_centered"),
            vertical_centered=options.get("vertical_centered"),
        )

    header_footer = spec.get("header_footer")
    if isinstance(header_footer, dict):
        wb.set_header_footer(
            sheet,
            odd_header=header_footer.get("odd_header"),
            odd_footer=header_footer.get("odd_footer"),
            even_header=header_footer.get("even_header"),
            even_footer=header_footer.get("even_footer"),
            first_header=header_footer.get("first_header"),
            first_footer=header_footer.get("first_footer"),
            different_odd_even=header_footer.get("different_odd_even"),
            different_first=header_footer.get("different_first"),
            scale_with_doc=header_footer.get("scale_with_doc"),
            align_with_margins=header_footer.get("align_with_margins"),
        )

    print_area = spec.get("print_area")
    if isinstance(print_area, str) and print_area:
        wb.set_print_area(sheet, print_area)

    titles = spec.get("print_titles")
    if isinstance(titles, dict):
        wb.set_print_titles(
            sheet,
            titles.get("repeat_rows", "") or "",
            titles.get("repeat_cols", "") or "",
        )

    for row in spec.get("row_breaks") or []:
        # Case rows are 1-based, matching every other A1-shaped field in
        # the schema; the API takes a zero-based row.
        wb.add_row_break(sheet, int(row) - 1)
    for col_key in spec.get("col_breaks") or []:
        wb.add_col_break(sheet, _column_index(str(col_key)))


def author_case(case: Dict[str, Any], *, module: Any) -> bytes:
    """Builds one round-trip case's workbook and returns the xlsx bytes.

    `module` is the imported `formulon` package, passed in rather than
    imported here so a caller can supply a wheel install or a source
    checkout without this module deciding which.
    """

    spec = case.get("roundtrip")
    if not isinstance(spec, dict):
        raise RoundtripSpecError(f"case {case.get('id')!r} carries no roundtrip block")

    # `create_default` rather than the bare constructor: it is the factory
    # a caller reaches for, so the fixture carries the same seeded style
    # table a real authored report would.
    with module.Workbook.create_default() as wb:
        sheet_indices = _apply_sheets(wb, case)
        _apply_layout(wb, case)

        sheet_key = spec.get("sheet", 0)
        if isinstance(sheet_key, str):
            if sheet_key not in sheet_indices:
                raise RoundtripSpecError(f"roundtrip sheet {sheet_key!r} is not one of {sorted(sheet_indices)}")
            sheet = sheet_indices[sheet_key]
        else:
            sheet = int(sheet_key)

        _apply_print_settings(wb, sheet, spec)
        return wb.save_as(module.WorkbookFormat.XLSX)


def fixture_payload(case: Dict[str, Any], *, module: Any) -> Dict[str, Any]:
    """Authors a case and wraps it for the driver request payload.

    The digest travels with the bytes so a golden can record which
    artefact Excel actually read. A round-trip golden that cannot name
    its input is not evidence about the writer that produced it.
    """

    data = author_case(case, module=module)
    return {
        "xlsx_base64": base64.b64encode(data).decode("ascii"),
        "xlsx_sha256": hashlib.sha256(data).hexdigest(),
        "xlsx_bytes": len(data),
    }


def load_case(case_path: Path, case_id: str) -> Dict[str, Any]:
    """Reads one case out of a workbook case JSON file."""

    doc = json.loads(case_path.read_text(encoding="utf-8"))
    for case in doc.get("cases") or []:
        if case.get("id") == case_id:
            return case
    raise RoundtripSpecError(f"case id {case_id!r} not found in {case_path}")


def roundtrip_case_ids(doc: Dict[str, Any]) -> List[str]:
    """Lists the ids in a case document that carry a `roundtrip` block."""

    return [str(case.get("id")) for case in doc.get("cases") or [] if isinstance(case.get("roundtrip"), dict)]


def _import_formulon(extra_path: Optional[str]) -> Any:
    if extra_path:
        sys.path.insert(0, extra_path)
    try:
        import formulon  # noqa: PLC0415 - deferred so --python-path can take effect
    except ImportError as exc:  # pragma: no cover - environment dependent
        raise RoundtripSpecError(
            "formulon is not importable; from a source checkout run with "
            "PYTHONPATH=packages/python or pass --python-path packages/python"
        ) from exc
    return formulon


def _main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("case_json", type=Path, help="tests/oracle/cases_wb/<suite>.case.json")
    parser.add_argument("--id", required=True, help="case id inside that file")
    parser.add_argument("--out", type=Path, help="write the xlsx here (default: print the digest only)")
    parser.add_argument("--python-path", help="directory to prepend to sys.path before importing formulon")
    args = parser.parse_args(argv)

    try:
        module = _import_formulon(args.python_path)
        case = load_case(args.case_json, args.id)
        data = author_case(case, module=module)
    except RoundtripSpecError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    digest = hashlib.sha256(data).hexdigest()
    if args.out is not None:
        args.out.write_bytes(data)
        print(f"{args.out}: {len(data)} bytes, sha256 {digest}")
    else:
        print(f"{len(data)} bytes, sha256 {digest}")
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
