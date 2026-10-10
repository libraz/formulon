"""Turns an xlwings cell or spill observation into a :class:`CaseResult`.

The Mac and Windows drivers read results the same way and differ only in two
platform primitives, which every entry point takes as callables:

  - ``displayed_text(cell)`` -- the cell's rendered string via the bridge.
  - ``app_evaluate(app)`` / ``spill_shape(app, anchor, *, max_cells)`` -- the
    ``Application.Evaluate`` adapter and the spill-shape probe built on it.
"""

from __future__ import annotations

import datetime as _dt
from typing import Any, Callable, Dict, List, Optional

from ._locale import normalise_error_token
from .base import (
    _ERR_DISPLAY_NAMES,
    CaseResult,
    _datetime_to_serial,
    case_formula_cell,
    case_shape_samples,
    is_empty_text_result,
)

DisplayedText = Callable[[Any], Optional[str]]
Evaluate = Callable[[str], Any]

# ERROR.TYPE result -> canonical error name.
_ERROR_TYPE_NAMES = {
    1: "#NULL!",
    2: "#DIV/0!",
    3: "#VALUE!",
    4: "#REF!",
    5: "#NAME?",
    6: "#NUM!",
    7: "#N/A",
    8: "#GETTING_DATA",
    9: "#SPILL!",
    10: "#CONNECT!",
    11: "#BLOCKED!",
    12: "#UNKNOWN!",
    13: "#FIELD!",
    14: "#CALC!",
}


def _error_name_from_excel(evaluate: Optional[Evaluate], cell) -> Optional[str]:
    """Asks Excel for the cell's error kind via ``ERROR.TYPE``; None if it cannot say."""

    if evaluate is None:
        return None
    try:
        code = evaluate(f"ERROR.TYPE({cell.get_address(external=True)})")
    except Exception:
        return None
    if isinstance(code, bool) or not isinstance(code, (int, float)):
        return None
    return _ERROR_TYPE_NAMES.get(int(code))


def error_display_from_cell(cell, displayed_text: DisplayedText, evaluate: Optional[Evaluate] = None) -> Optional[str]:
    """Returns the tokenised Excel error name for `cell`, or None.

    Walks four progressively weaker signals:
      1. ``xlwings.utils.CVErr`` -- ideal, but only surfaces on some
         Excel / xlwings build pairs.
      2. ``cell.value`` already a ``'#DIV/0!'``-style string.
      3. The displayed text (COM ``.Text``) matches a known error.
      4. ``cell.value`` is ``None`` AND the displayed text nonetheless
         starts with ``#``. This is the fallback path where the Python
         layer has coerced the error into ``None``.

    A ``#`` text that none of these resolve (a localized name missing from the
    table) is asked of Excel through ``ERROR.TYPE``; if that fails too, an
    error cell must not be recorded as blank, so this raises.
    """

    raw = cell.value
    try:
        from xlwings.utils import CVErr  # type: ignore

        if isinstance(raw, CVErr):
            s = str(raw)
            if s in _ERR_DISPLAY_NAMES:
                return s
    except Exception:  # pragma: no cover - older xlwings without CVErr
        pass

    if isinstance(raw, str):
        if raw in _ERR_DISPLAY_NAMES:
            return raw
        # ja-JP / de-DE / fr-FR builds occasionally surface the localized
        # token directly through cell.value (when the bridge has already
        # decoded the CVErr to a string). Normalise here so the golden
        # JSON always carries the canonical English form.
        canon = normalise_error_token(raw)
        if canon is not None:
            return canon

    text = displayed_text(cell)
    if text in _ERR_DISPLAY_NAMES:
        return text
    # Range.Text is locale-bound: de-DE returns "#WERT!" for #VALUE!,
    # fr-FR returns "#VALEUR!", and so on. Normalise through the shared
    # localisation map before falling back to the prefix heuristic.
    if text:
        canon = normalise_error_token(text)
        if canon is not None:
            return canon
    if text and text.startswith("#") and (text.endswith("!") or text.endswith("?") or text == "#N/A"):
        for name in _ERR_DISPLAY_NAMES:
            if text == name:
                return name
    # A value too wide for its column renders as a run of '#' (a long localized
    # TRUE such as VERDADEIRO); the bridge still read the value itself.
    if text and set(text) == {"#"} and raw is not None and not isinstance(raw, str):
        return None
    if text and text.startswith("#"):
        name = _error_name_from_excel(evaluate, cell)
        if name is not None:
            return name
        # A non-empty string value that is not an error is plain text.
        if not (isinstance(raw, str) and raw):
            raise RuntimeError(f"unresolvable error text {text!r}; refusing to record the cell as blank")
    return None


def classify_value(cell, evaluate, displayed_text: DisplayedText) -> CaseResult:
    """Converts an xlwings cell observation into a CaseResult.

    The ``cell.value`` read happens once up front; subsequent checks may
    consult the COM ``.Text`` fallback for edge cases (error cells,
    pre-1900 serials) where the Python-side value is lossy. ``evaluate`` is
    the Application.Evaluate adapter that resolves an empty read-back into
    ``""`` or blank; without it (``None``) an empty read-back stays blank.
    """

    err = error_display_from_cell(cell, displayed_text, evaluate)
    if err is not None:
        return CaseResult(id="", kind="error", error_code=err)

    v = cell.value
    # The classifier has no workbook epoch metadata, so datetimes use the
    # shared 1900-system conversion; 1904 cases should expose date values
    # through a numeric/boolean formula such as N(DATE()).
    if isinstance(v, _dt.datetime):
        return CaseResult(id="", kind="number", value=_datetime_to_serial(v))
    if isinstance(v, _dt.date):  # pragma: no cover - xlwings mostly returns datetime
        combined = _dt.datetime(v.year, v.month, v.day)
        return CaseResult(id="", kind="number", value=_datetime_to_serial(combined))

    if isinstance(v, bool):
        return CaseResult(id="", kind="bool", value=bool(v))
    if isinstance(v, (int, float)):
        return CaseResult(id="", kind="number", value=float(v))

    if v is None or v == "":
        # Pre-1900 serials come back as None (no datetime representation);
        # a displayed-text parse was never implemented, so they stay blank.
        text = displayed_text(cell)
        if text and text.strip():
            pass
        if evaluate is not None and is_empty_text_result(evaluate, cell):
            return CaseResult(id="", kind="text", value="")
        return CaseResult(id="", kind="blank")

    if isinstance(v, str):
        return CaseResult(id="", kind="text", value=v)

    if isinstance(v, list):
        rows = len(v)
        cols = 0
        if rows > 0 and isinstance(v[0], list):
            cols = len(v[0])
            flat = [array_cell_from_scalar(classify_python_scalar(item)) for row in v for item in row]
        else:
            cols = rows
            rows = 1
            flat = [array_cell_from_scalar(classify_python_scalar(item)) for item in v]
        return CaseResult(
            id="",
            kind="array",
            value=flat,
            array_shape=[rows, cols],
        )
    return CaseResult(id="", kind="text", value=str(v))


def classify_python_scalar(v: Any) -> CaseResult:
    """Classifies a scalar value already extracted from an xlwings array."""

    if isinstance(v, _dt.datetime):
        return CaseResult(id="", kind="number", value=_datetime_to_serial(v))
    if isinstance(v, _dt.date):  # pragma: no cover - xlwings mostly returns datetime
        combined = _dt.datetime(v.year, v.month, v.day)
        return CaseResult(id="", kind="number", value=_datetime_to_serial(combined))
    if isinstance(v, bool):
        return CaseResult(id="", kind="bool", value=bool(v))
    if isinstance(v, (int, float)):
        return CaseResult(id="", kind="number", value=float(v))
    if v is None or v == "":
        return CaseResult(id="", kind="blank")
    if isinstance(v, str):
        canon = normalise_error_token(v)
        if canon is not None:
            return CaseResult(id="", kind="error", error_code=canon)
        return CaseResult(id="", kind="text", value=v)
    return CaseResult(id="", kind="text", value=str(v))


def array_cell_from_scalar(result: CaseResult) -> Any:
    if result.kind == "blank":
        return None
    if result.kind in {"number", "bool", "text"}:
        return result.value
    if result.kind == "error":
        return {"kind": "error", "code": result.error_code or "#UNKNOWN!"}
    return {"kind": result.kind, "value": result.value}


def classify_shape_result(
    app,
    sht,
    anchor_addr: str,
    samples: List[str],
    *,
    displayed_text: DisplayedText,
    app_evaluate: Callable[[Any], Callable[[str], Any]],
    spill_shape: Callable[..., tuple[int, int]],
) -> CaseResult:
    """Records a spill as its shape plus the case's declared sample cells.

    For results too large to materialise cell by cell. The shape comes from
    the same non-invasive probe the cell walk uses; only the listed cells
    are read.
    """

    rows, cols = spill_shape(app, sht.range(anchor_addr), max_cells=None)
    evaluate = app_evaluate(app)
    values = {
        addr: array_cell_from_scalar(classify_value(sht.range(addr), evaluate, displayed_text)) for addr in samples
    }
    return CaseResult(id="", kind="array_shape", value=values, array_shape=[rows, cols])


def classify_case_result(
    app,
    sht,
    case: Dict[str, Any],
    *,
    shape_result: Callable[[Any, Any, str, List[str]], CaseResult],
    result_cell: Callable[[Any, Any, str], CaseResult],
) -> CaseResult:
    """Reads one case's result in whichever capture mode it declares."""

    anchor_addr = case_formula_cell(case)
    samples = case_shape_samples(case)
    if samples is not None:
        return shape_result(app, sht, anchor_addr, samples)
    return result_cell(app, sht, anchor_addr)


def classify_result_cell(
    app,
    sht,
    anchor_addr: str,
    *,
    displayed_text: DisplayedText,
    app_evaluate: Callable[[Any], Callable[[str], Any]],
    spill_shape: Callable[..., tuple[int, int]],
) -> CaseResult:
    """Classifies the anchor scalar or the full dynamic spill if present."""

    shape = spill_shape(app, sht.range(anchor_addr))
    rows, cols = shape
    evaluate = app_evaluate(app)
    if rows == 1 and cols == 1:
        return classify_value(sht.range(anchor_addr), evaluate, displayed_text)
    anchor = sht.range(anchor_addr)
    flat: List[Any] = []
    for r in range(rows):
        for c in range(cols):
            flat.append(array_cell_from_scalar(classify_value(anchor.offset(r, c), evaluate, displayed_text)))
    return CaseResult(id="", kind="array", value=flat, array_shape=[rows, cols])
