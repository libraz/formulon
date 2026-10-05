"""PageSetup readers: what Excel resolved for a sheet's print settings.

Each reader maps an unavailable COM property to ``None`` so the golden records
what was observed rather than aborting the capture.
"""

from __future__ import annotations

from typing import Any, Dict, List, Tuple

from .case_sheet import PRINT_ORIENTATIONS, normalise_print_area
from .windows_com import com_bool, com_int, com_scalar, com_text

_PRINT_ORIENTATION_NAMES = {value: name for name, value in PRINT_ORIENTATIONS.items()}

# Excel COM reports `PageSetup.LeftMargin` etc. in points (72 pt = 1
# inch). The Formulon `PageMargins` struct works in inches, so we
# divide on read.
_POINTS_PER_INCH = 72.0

# A whole-column print area would otherwise issue 16,384 COM round trips
# for a diagnostic nobody asserts.
GEOMETRY_TRACK_CAP = 64


def split_a1(addr: str) -> Tuple[int, int]:
    """Splits a bare A1 address into 1-based (row, column) COM indices."""

    addr = addr.replace("$", "")
    i = 0
    while i < len(addr) and addr[i].isalpha():
        i += 1
    col = 0
    for ch in addr[:i]:
        col = col * 26 + (ord(ch.upper()) - ord("A") + 1)
    return int(addr[i:]), col


def read_orientation_value(page_setup) -> Any:
    """Returns the post-apply Orientation as `"portrait"` / `"landscape"`.

    Load-bearing rather than cosmetic: the pagination read can answer from
    a page geometry the case never asked for (a portrait case coming back
    with the preceding landscape case's breaks), and without the applied
    orientation beside the breaks that is indistinguishable from Excel
    genuinely disagreeing with the C++ paginator. An unreadable or
    unrecognised value is reported as-is rather than guessed at.
    """

    try:
        value = int(page_setup.Orientation)
    except Exception:
        return None
    return _PRINT_ORIENTATION_NAMES.get(value, value)


def read_paper_value(page_setup) -> Any:
    """Returns the post-apply `PaperSize` as the `XlPaperSize` int.

    The companion to the orientation read: paper and orientation together
    are what fix the page the breaks were measured on.
    """

    try:
        return int(page_setup.PaperSize)
    except Exception:
        return None


def read_zoom_value(page_setup) -> Any:
    """Returns the post-apply Zoom: int percent, or `False` when Fit-active.

    Excel's COM ``PageSetup.Zoom`` is ``False`` when "Fit to" pagination
    is engaged, else an integer percent (10..400). Boolean is checked
    before numeric because Python's ``isinstance(False, int)`` is True.
    """

    try:
        val = page_setup.Zoom
    except Exception:
        return None
    if isinstance(val, bool):
        return False if val is False else True
    if isinstance(val, (int, float)):
        return int(val)
    return None


def read_fit_value(page_setup, name: str) -> Any:
    """Returns the post-apply ``FitToPagesWide/Tall``: int or False (auto).

    Excel returns ``False`` for the "auto" / unset axis and an integer
    (1..32767) when constrained. Same bool-before-int ordering applies
    as for Zoom.
    """

    try:
        val = getattr(page_setup, name)
    except Exception:
        return None
    if isinstance(val, bool):
        return False
    if isinstance(val, (int, float)):
        return int(val)
    return None


_MARGIN_ATTRS = (
    ("left", "LeftMargin"),
    ("right", "RightMargin"),
    ("top", "TopMargin"),
    ("bottom", "BottomMargin"),
    ("header", "HeaderMargin"),
    ("footer", "FooterMargin"),
)


def read_margins(page_setup) -> Dict[str, Any]:
    """Returns the post-apply page margins in inches.

    Excel's COM `PageSetup.{Left,Right,Top,Bottom,Header,Footer}Margin`
    are points; we divide by 72 so the golden surfaces inches (matching
    Formulon's `PageMargins` struct and OOXML's <pageMargins> tag, which
    are both inch-denominated). The body-height calibration hunt in the
    print_matrix follow-up needs these values to distinguish "Excel
    applied the OOXML defaults" from "Excel applied a workbook-template
    preset that we never asked for".
    """

    out: Dict[str, Any] = {}
    for key, attr in _MARGIN_ATTRS:
        try:
            val = getattr(page_setup, attr)
        except Exception:
            out[key] = None
            continue
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            out[key] = round(float(val) / _POINTS_PER_INCH, 6)
        else:
            out[key] = None
    return out


# `XlPageBreak`. A manual break authored into the file must read back as
# manual; one that degraded to automatic means Excel discarded the
# `<rowBreaks man="1">` we wrote and re-derived the break itself.
_XL_PAGE_BREAK_MANUAL = -4135

# The header/footer sections Excel exposes per page class, in the order
# OOXML concatenates them into one string (`&L...&C...&R...`).
_HEADER_FOOTER_POSITIONS = ("Left", "Center", "Right")

# `PageSetup.<name>` for the odd/primary pages, and the `Page` object
# carrying the same three positions for the even and first-page classes.
_HEADER_FOOTER_CLASSES = (
    ("odd", None),
    ("even", "EvenPage"),
    ("first", "FirstPage"),
)


def _read_header_footer(page_setup) -> Dict[str, Any]:
    """Reads the six header/footer sections as Excel reports them.

    OOXML stores one string per section with `&L` / `&C` / `&R` markers
    inside it; COM splits the same content across three properties. The
    golden records the split form, because that is what Excel actually
    parsed the string into -- collapsing it back would hide a case where
    Excel put our text in the wrong third.
    """

    out: Dict[str, Any] = {
        "different_odd_even": com_bool(page_setup, "OddAndEvenPagesHeaderFooter"),
        "different_first": com_bool(page_setup, "DifferentFirstPageHeaderFooter"),
        "scale_with_doc": com_bool(page_setup, "ScaleWithDocHeaderFooter"),
        "align_with_margins": com_bool(page_setup, "AlignMarginsHeaderFooter"),
    }
    for class_name, page_attr in _HEADER_FOOTER_CLASSES:
        owner = page_setup if page_attr is None else com_scalar(page_setup, page_attr)
        for band in ("Header", "Footer"):
            for position in _HEADER_FOOTER_POSITIONS:
                key = f"{class_name}_{band.lower()}_{position.lower()}"
                if owner is None:
                    out[key] = None
                    continue
                if page_attr is None:
                    # The odd/primary sections are plain strings on
                    # `PageSetup` itself.
                    out[key] = com_text(owner, f"{position}{band}")
                    continue
                section = com_scalar(owner, f"{position}{band}")
                out[key] = None if section is None else com_text(section, "Text")
    return out


def _read_manual_breaks(ws) -> Dict[str, List[int]]:
    """Returns the manual row / column breaks as zero-based indices.

    Automatic breaks are filtered out: an authored break that Excel
    re-derived rather than honoured is the failure this case exists to
    catch, and keeping both kinds in one list would let a coincidental
    automatic break at the same position mask it.
    """

    rows: List[int] = []
    try:
        for i in range(1, int(ws.api.HPageBreaks.Count) + 1):
            brk = ws.api.HPageBreaks(i)
            if com_int(brk, "Type") == _XL_PAGE_BREAK_MANUAL:
                rows.append(int(brk.Location.Row) - 1)
    except Exception:
        pass
    cols: List[int] = []
    try:
        for i in range(1, int(ws.api.VPageBreaks.Count) + 1):
            brk = ws.api.VPageBreaks(i)
            if com_int(brk, "Type") == _XL_PAGE_BREAK_MANUAL:
                cols.append(int(brk.Location.Column) - 1)
    except Exception:
        pass
    rows.sort()
    cols.sort()
    return {"manual_row_breaks": rows, "manual_col_breaks": cols}


def read_roundtrip(sht) -> Dict[str, Any]:
    """Reads every print setting Excel resolved from an opened workbook."""

    page_setup = sht.api.PageSetup
    print_area = com_text(page_setup, "PrintArea") or ""
    title_rows = com_text(page_setup, "PrintTitleRows") or ""
    title_cols = com_text(page_setup, "PrintTitleColumns") or ""

    observed: Dict[str, Any] = {
        "page_setup": {
            "paper_size": com_int(page_setup, "PaperSize"),
            "orientation": com_int(page_setup, "Orientation"),
            "zoom": read_zoom_value(page_setup),
            "fit_to_pages_wide": read_fit_value(page_setup, "FitToPagesWide"),
            "fit_to_pages_tall": read_fit_value(page_setup, "FitToPagesTall"),
        },
        "page_margins": read_margins(page_setup),
        "print_options": {
            "grid_lines": com_bool(page_setup, "PrintGridlines"),
            "headings": com_bool(page_setup, "PrintHeadings"),
            "horizontal_centered": com_bool(page_setup, "CenterHorizontally"),
            "vertical_centered": com_bool(page_setup, "CenterVertically"),
        },
        "header_footer": _read_header_footer(page_setup),
        "print_area": normalise_print_area(print_area),
        "print_title_rows": normalise_print_area(title_rows),
        "print_title_cols": normalise_print_area(title_cols),
    }
    observed.update(_read_manual_breaks(sht))
    return observed
