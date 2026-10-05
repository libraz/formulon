"""Windows COM access primitives: displayed text, raw-cell adapter, error text.

``xlwings`` on Windows is a thin layer over pywin32; these helpers cover the
parts of the object model it does not wrap and the way COM failures surface.
"""

from __future__ import annotations

from typing import Any, List, Optional


def cell_displayed_text(cell) -> Optional[str]:
    """Returns the rendered string of a cell via the COM bridge.

    On Windows, ``cell.api.Text`` is the canonical property that yields
    the displayed string (including ``'#DIV/0!'`` for errors whose
    Python-side ``.value`` has been coerced to ``None``).

    Why we don't try fallback property names here:
      Excel's IDispatch is case-insensitive, so ``text`` resolves to the
      same DISPID as ``Text`` and adds no coverage; ``DisplayValue`` is
      a Chart.Axis property and is not present on Range; ``string_value``
      is an xlwings *high-level* Range attribute, not a COM property.
      Worse, an earlier defensive version of this function iterated over
      a list that included the lowercase ``text``, and on Excel 16.0 /
      M365 a getattr for the lowercase form triggered an apparent
      infinite retry inside ``win32com``'s ``COMRetryObjectWrapper`` --
      cells whose ``Text`` legitimately resolved to ``""`` (e.g. the
      result of ``=ARRAYTOTEXT("")``) made the loop fall through to the
      lowercase try, which never returned and burned Excel CPU at ~60%
      indefinitely. We saw the run hang for 30+ minutes before
      diagnosis. Sticking to the canonical ``Text`` property avoids the
      whole pitfall.
    """

    try:
        api = cell.api
    except Exception:
        return None
    try:
        val = api.Text
    except Exception:
        return None
    if isinstance(val, str):
        return val or None
    return None


class CellAdapter:
    """Adapts a raw COM ``Range`` to the small surface ``classify_value``
    expects (a ``.value`` attribute plus the COM passthrough used by the
    error / displayed-text fallbacks).
    """

    def __init__(self, com_cell) -> None:
        self.api = com_cell

    @property
    def value(self) -> Any:
        return self.api.Value


def format_com_error(exc: BaseException) -> str:
    """Returns a one-line summary of a pywin32 / xlwings exception.

    Includes the COM HRESULT and Excel's localised description when
    present, so per-target divergence triage can match on the actual
    failure (e.g. ``COM -2147352567: 例外が発生しました。``) rather
    than a generic Python traceback. Falls back to ``repr(exc)`` for
    non-COM exceptions.
    """

    try:
        args = getattr(exc, "args", ()) or ()
        if args and isinstance(args[0], int):
            hresult = args[0]
            descr = ""
            if len(args) >= 2 and isinstance(args[1], str):
                descr = args[1]
            # pywin32 com_error.args = (hresult, source, excepinfo, argerr)
            # excepinfo = (wcode, source, description, helpfile, helpcontext, scode)
            # When the HRESULT is the generic DISP_E_EXCEPTION the real
            # Excel error message lives in excepinfo[2], so pull it out.
            extras: List[str] = []
            if len(args) >= 3 and args[2] is not None:
                info = args[2]
                if isinstance(info, tuple) and len(info) >= 3:
                    src = info[1] if isinstance(info[1], str) else ""
                    info_descr = info[2] if isinstance(info[2], str) else ""
                    scode = info[5] if len(info) >= 6 else None
                    if info_descr.strip():
                        extras.append(info_descr.strip())
                    if src.strip():
                        extras.append(f"source={src.strip()}")
                    if isinstance(scode, int):
                        extras.append(f"scode=0x{scode & 0xFFFFFFFF:08X}")
            if extras:
                descr = (descr + " -- " if descr else "") + "; ".join(extras)
            return f"COM {hresult}: {descr}".strip().rstrip(":")
    except Exception:
        pass
    return f"{type(exc).__name__}: {exc}"[:200]


def com_scalar(owner, attr: str) -> Any:
    """Reads one COM property, mapping an unavailable one to ``None``.

    A property Excel does not expose on this host must not abort the
    capture: the golden records `null` and the comparison skips that
    field, which is honest about what was observed.
    """

    try:
        return getattr(owner, attr)
    except Exception:
        return None


def com_bool(owner, attr: str) -> Optional[bool]:
    value = com_scalar(owner, attr)
    return bool(value) if isinstance(value, bool) else None


def com_int(owner, attr: str) -> Optional[int]:
    value = com_scalar(owner, attr)
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return int(value)


def com_text(owner, attr: str) -> Optional[str]:
    value = com_scalar(owner, attr)
    return value if isinstance(value, str) else None
