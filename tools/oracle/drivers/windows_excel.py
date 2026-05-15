#!/usr/bin/env python3
"""Windows Excel 365 oracle driver.

Drives Excel.exe through xlwings (which on Windows is a thin wrapper over
pywin32 / win32com COM automation) to evaluate Formulon oracle cases under
controlled options (manual calc / iterative off / 1904 off). The driver is
intentionally small -- the caller (``oracle_gen.py``) handles case loading,
divergence filtering, and golden JSON emission.

Windows-only. ``xlwings.App()`` here uses COM, so:

  - Excel must be installed and Office must be activated (an unactivated
    Excel refuses to run COM automation against a fresh workbook).
  - The Python interpreter must be running on Windows; calling this from
    WSL2 will fail in :func:`_ensure_windows`. WSL2 callers should use
    :class:`tools.oracle.drivers.wsl_bridge.WSLBridgeOracle`, which
    invokes us as a Windows-side subprocess.
  - The first launch may pop a "trust this app" dialog; after a single
    interactive grant subsequent runs go through silently.

## Options we pin

- ``calculation = 'manual'`` -- we call ``app.calculate()`` after every
  batch ourselves, so formulas never race partial input.
- ``screen_updating = False`` -- large batches are 10-20x faster without
  it.
- ``display_alerts = False`` -- suppresses dialogs that would otherwise
  block the COM thread.
- ``date_1904 = False`` per workbook (set after ``books.add()``).
- ``enable_iterative_calculation = False`` on the app. Iterative cases
  flip this locally and restore it.

## Batch layout

Mirrors :mod:`tools.oracle.drivers.macos_excel`: one worksheet per case,
the case formula at ``Z1``, setup cells written verbatim at their absolute
A1 addresses. A single ``app.calculate()`` resolves the whole batch.

## Wire-protocol entrypoints

Running ``python -m tools.oracle.drivers.windows_excel --input X.json
--output Y.json`` reads a command (``probe_environment`` or ``run_suite``)
from the input JSON and writes the result vector to the output JSON.
This is the legacy single-shot path; each invocation pays a full Excel
cold-start.

Running ``python -m tools.oracle.drivers.windows_excel --serve`` opens
Excel once, prints a ``ready`` line to stdout, then loops reading
newline-delimited JSON requests from stdin and writing newline-delimited
JSON responses to stdout. The bridge uses this mode so all suites in a
generation run share a single Excel instance -- Excel cold-start (5-15s
in our measurements) dominates total runtime when generating 90+ suites,
and amortising it once over the whole run cuts wall-clock by ~10x.
"""

from __future__ import annotations

import datetime as _dt
import platform
from pathlib import Path
from typing import Any, Dict, List, Optional

from .base import (
    CaseResult,
    EnvironmentInfo,
    OracleDriver,
    _datetime_to_serial,
    _ERR_DISPLAY_NAMES,
)
from ._locale import detect_locale_from_app, normalise_error_token

try:
    import xlwings as xw  # type: ignore
except ImportError as exc:  # pragma: no cover - handled in oracle_gen.py
    raise RuntimeError(
        "xlwings is not installed; run `make oracle-setup` first"
    ) from exc


def _ensure_windows() -> None:
    """Refuses to start unless the host OS is Windows.

    The COM automation path is Windows-only; on Mac it would silently
    fall back to AppleScript, which has different attribute spellings
    and would mis-classify error cells. Refuse loudly instead.
    """

    if platform.system() != "Windows":
        raise RuntimeError(
            "windows_excel driver is Windows-only (uses COM automation). "
            "Current platform: " + platform.system()
            + ". Use wsl_bridge from WSL2, or macos_excel on Darwin."
        )


def _cell_displayed_text(cell) -> Optional[str]:
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


def _error_display_from_cell(cell) -> Optional[str]:
    """Returns the tokenised Excel error name for `cell`, or None.

    Walks four progressively weaker signals:
      1. ``xlwings.utils.CVErr`` -- ideal, but only surfaces on some
         Excel / xlwings build pairs.
      2. ``cell.value`` already a ``'#DIV/0!'``-style string.
      3. The displayed text (COM ``.Text``) matches a known error.
      4. ``cell.value`` is ``None`` AND the displayed text nonetheless
         starts with ``#``. This is the fallback path where the Python
         layer has coerced the error into ``None``.
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

    text = _cell_displayed_text(cell)
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
    return None


def _classify_value(cell) -> CaseResult:
    """Converts an xlwings cell observation into a CaseResult.

    The ``cell.value`` read happens once up front; subsequent checks may
    consult the COM ``.Text`` fallback for edge cases (error cells,
    pre-1900 serials) where the Python-side value is lossy.
    """

    err = _error_display_from_cell(cell)
    if err is not None:
        return CaseResult(id="", kind="error", error_code=err)

    v = cell.value
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
        text = _cell_displayed_text(cell)
        if text and text.strip():
            pass
        return CaseResult(id="", kind="blank")

    if isinstance(v, str):
        return CaseResult(id="", kind="text", value=v)

    if isinstance(v, list):
        rows = len(v)
        cols = 0
        if rows > 0 and isinstance(v[0], list):
            cols = len(v[0])
            flat = [
                _array_cell_from_scalar(_classify_python_scalar(item))
                for row in v
                for item in row
            ]
        else:
            cols = rows
            rows = 1
            flat = [_array_cell_from_scalar(_classify_python_scalar(item)) for item in v]
        return CaseResult(
            id="",
            kind="array",
            value=flat,
            array_shape=[rows, cols],
        )
    return CaseResult(id="", kind="text", value=str(v))


def _classify_python_scalar(v: Any) -> CaseResult:
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


def _array_cell_from_scalar(result: CaseResult) -> Any:
    if result.kind == "blank":
        return None
    if result.kind in {"number", "bool", "text"}:
        return result.value
    if result.kind == "error":
        return {"kind": "error", "code": result.error_code or "#UNKNOWN!"}
    return {"kind": result.kind, "value": result.value}


def _write_spill_shape_probe(sht, anchor_addr: str = "Z1") -> bool:
    """Writes helper formulas that report the dynamic spill shape, if any."""

    rows_cell = sht.range("XFD1")
    cols_cell = sht.range("XFD2")
    try:
        rows_cell.formula2 = f"=ROWS({anchor_addr}#)"
    except Exception:
        try:
            rows_cell.formula = f"=ROWS({anchor_addr}#)"
        except Exception:
            return False
    try:
        cols_cell.formula2 = f"=COLUMNS({anchor_addr}#)"
    except Exception:
        try:
            cols_cell.formula = f"=COLUMNS({anchor_addr}#)"
        except Exception:
            return False
    return True


def _read_spill_shape_probe(sht) -> Optional["tuple[int, int]"]:
    rows = _classify_value(sht.range("XFD1"))
    cols = _classify_value(sht.range("XFD2"))
    if not rows.kind == cols.kind == "number":
        return None
    row_count = int(rows.value)
    col_count = int(cols.value)
    if row_count <= 0 or col_count <= 0:
        return None
    return row_count, col_count


def _classify_result_cell(sht, anchor_addr: str = "Z1") -> CaseResult:
    """Classifies the anchor scalar or the full dynamic spill if present."""

    shape = _read_spill_shape_probe(sht)
    if shape is None:
        return _classify_value(sht.range(anchor_addr))
    rows, cols = shape
    if rows == 1 and cols == 1:
        return _classify_value(sht.range(anchor_addr))
    anchor = sht.range(anchor_addr)
    flat: List[Any] = []
    for r in range(rows):
        for c in range(cols):
            flat.append(_array_cell_from_scalar(_classify_value(anchor.offset(r, c))))
    return CaseResult(id="", kind="array", value=flat, array_shape=[rows, cols])


class WindowsExcelOracle(OracleDriver):
    """Thin wrapper over a single hidden Excel.exe instance (COM).

    The instance is reused across suites so we pay COM startup latency
    only once. Call :meth:`close` (or use as a context manager) to tear
    down. Refuses to construct on non-Windows hosts.
    """

    def __init__(self, visible: bool = False) -> None:
        _ensure_windows()
        self._app = xw.App(visible=visible, add_book=False)
        # Several Application properties are workbook-gated on Excel
        # 16.0 / Office 365: assigning to them while no book is open
        # either raises a generic COM error (Calculation) or makes the
        # COM bridge hang indefinitely (EnableIterativeCalculation; we
        # observed 200+ CPU seconds of busy wait). We therefore set
        # only the workbook-independent properties here and re-apply
        # the gated ones inside run_suite() once books.add() succeeds.
        self._app.screen_updating = False
        self._app.display_alerts = False

    def __enter__(self) -> "WindowsExcelOracle":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def close(self) -> None:
        try:
            self._app.quit()
        except Exception:
            pass

    def probe_environment(self) -> EnvironmentInfo:
        """Returns a snapshot of Excel's version / locale for logging.

        ``app.version`` is a string property on recent xlwings on
        Windows; the fallback path queries the raw COM ``Version`` /
        ``Build`` properties. The locale is detected via
        ``Application.International(xlCountryCode)`` (see
        :func:`tools.oracle.drivers._locale.detect_locale_from_app`); a
        probe failure or unmapped country code yields an empty string,
        which the generator surfaces as a missing locale in
        ENVIRONMENT.md so the operator can investigate.
        """

        version = ""
        try:
            version = str(self._app.version) if self._app.version else ""
        except Exception:
            version = ""
        if not version:
            try:
                api = self._app.api
                v = api.Version() if callable(getattr(api, "Version", None)) else getattr(api, "Version", "")
                b = api.Build() if callable(getattr(api, "Build", None)) else getattr(api, "Build", "")
                version = str(v).strip()
                if b and str(b) not in version:
                    version = f"{version} (Build {b})"
            except Exception:
                pass
        locale = detect_locale_from_app(self._app) or ""
        return EnvironmentInfo(
            excel_version=version.strip(),
            excel_locale=locale,
            date1904=False,
            iterative=False,
        )

    def run_suite(
        self,
        suite_name: str,
        cases: List[Dict[str, Any]],
        *,
        date1904: bool = False,
        iterative: bool = False,
    ) -> List[CaseResult]:
        """Evaluates every case in `cases` in a fresh workbook and returns
        the observed results.

        Each ``cases[i]`` is expected to be a dict with keys ``id``,
        ``formula``, and ``setup`` (mapping A1 -> ``{kind, value, ...}``
        record). The driver does not normalise; it trusts
        ``case_schema._normalise_value`` upstream.
        """

        # Cross-sheet setup ("Sheet2!A1") cannot be shared across cases
        # in one workbook -- a later case's write to the same external
        # sheet would clobber an earlier case's snapshot, and Excel
        # would happily evaluate both formulas against the last value.
        # Detect that at the suite level and isolate each such case in
        # its own workbook. Suites without cross-sheet setup keep the
        # historical fast path: one shared workbook, one sheet per case.
        cross_sheet = any(
            any("!" in addr for addr in (case.get("setup") or {}))
            for case in cases
        )
        if cross_sheet:
            return self._run_suite_per_case_workbook(
                suite_name, cases, date1904=date1904, iterative=iterative
            )

        wb = self._app.books.add()
        try:
            # Application.Calculation can only be assigned with a workbook
            # open; the __init__ attempt is suppressed for that reason, so
            # apply it here once books.add() has succeeded.
            try:
                self._app.calculation = "manual"
            except Exception:
                pass
            try:
                wb.api.Date1904 = date1904
            except Exception:
                try:
                    wb.api.date1904 = date1904
                except Exception:
                    pass
            prior_iter = None
            # The VBA / COM property for "Enable iterative calculation" is
            # `Application.Iteration` (a Boolean), NOT
            # `EnableIterativeCalculation` — that name does not exist on
            # the COM Application interface despite being the GUI label.
            # Probed against Excel 16.0 ja-JP: reading
            # `app.api.EnableIterativeCalculation` raises a "name unknown"
            # COM error which the previous outer try/except silently
            # swallowed, leaving iteration off and circular-formula cases
            # stuck at their first evaluation. We DO NOT fall back to the
            # lowercase alias (`iteration`) because xlwings'
            # COMRetryObjectWrapper retries unknown lowercase names
            # indefinitely on Windows (same hang we hit with `.text`).
            try:
                prior_iter = bool(self._app.api.Iteration)
                self._app.api.Iteration = iterative
            except Exception:
                prior_iter = None

            first_sheet = wb.sheets[0]
            case_sheets: List[object] = []
            # Per-case write failures are recorded here so the rest of
            # the batch keeps moving. Without isolation, a single
            # rejected formula (e.g. dynamic-array syntax that this
            # build's COM bridge refuses on both Formula and Formula2)
            # raises out of the loop and aborts the whole suite -- 90+
            # adjacent cases that would have evaluated cleanly are then
            # silently lost.
            write_errors: Dict[str, str] = {}
            for i, case in enumerate(cases):
                if i == 0:
                    sht = first_sheet
                else:
                    sht = wb.sheets.add(after=case_sheets[-1])
                safe_id = _sanitize_sheet_name(case["id"])
                sht.name = f"c{i + 1:03d}_{safe_id}"[:31]
                case_sheets.append(sht)

                try:
                    setup = case.get("setup") or {}
                    for addr, rec in setup.items():
                        _write_cell(sht, addr, rec)
                    result_cell = sht.range("Z1")
                    try:
                        result_cell.number_format = "General"
                    except Exception:
                        pass
                    # Prefer Formula2 (Excel 2019+ dynamic-array
                    # semantics), but fall back to Formula on builds
                    # whose COM bridge rejects Formula2 with a generic
                    # COM error. The fallback adds an implicit @ to
                    # dynamic-array formulas; tests whose oracle
                    # behavior depends on spill output should be
                    # guarded by tests/divergence.yaml.
                    try:
                        result_cell.formula2 = case["formula"]
                    except Exception:
                        result_cell.formula = case["formula"]
                except Exception as exc:
                    write_errors[case["id"]] = _format_com_error(exc)

            self._app.calculate()
            for case, sht in zip(cases, case_sheets):
                if case["id"] not in write_errors:
                    _write_spill_shape_probe(sht)
            self._app.calculate()

            out: List[CaseResult] = []
            for case, sht in zip(cases, case_sheets):
                if case["id"] in write_errors:
                    out.append(CaseResult(
                        id=case["id"],
                        kind="skipped",
                        value=write_errors[case["id"]],
                    ))
                    continue
                try:
                    result = _classify_result_cell(sht)
                    result.id = case["id"]
                    out.append(result)
                except Exception as exc:
                    out.append(CaseResult(
                        id=case["id"],
                        kind="skipped",
                        value=_format_com_error(exc),
                    ))
            return out
        finally:
            try:
                if prior_iter is not None:
                    self._app.api.Iteration = prior_iter
            except Exception:
                pass
            try:
                wb.close()
            except Exception:
                pass

    def _run_suite_per_case_workbook(
        self,
        suite_name: str,
        cases: List[Dict[str, Any]],
        *,
        date1904: bool,
        iterative: bool,
    ) -> List[CaseResult]:
        """Per-case-workbook runner for cross-sheet-setup suites.

        Each case gets its own ``books.add()`` so the sheets created from
        ``"Sheet2!A1"`` setup keys cannot leak between cases. The cost is
        one workbook open/close per case (sub-second on a warm
        ``Application``); only suites that actually need this pay it.
        """

        try:
            self._app.calculation = "manual"
        except Exception:
            pass
        prior_iter = None
        try:
            prior_iter = bool(self._app.api.Iteration)
            self._app.api.Iteration = iterative
        except Exception:
            prior_iter = None
        out: List[CaseResult] = []
        try:
            for case in cases:
                wb = self._app.books.add()
                try:
                    try:
                        wb.api.Date1904 = date1904
                    except Exception:
                        try:
                            wb.api.date1904 = date1904
                        except Exception:
                            pass
                    sht = wb.sheets[0]
                    write_error: Optional[str] = None
                    try:
                        setup = case.get("setup") or {}
                        for addr, rec in setup.items():
                            sheet_name, bare_addr = _split_sheet_qualified_addr(addr)
                            target_sht = sht if sheet_name is None else _get_or_add_sheet(wb, sheet_name)
                            _write_cell(target_sht, bare_addr, rec)
                        result_cell = sht.range("Z1")
                        try:
                            result_cell.number_format = "General"
                        except Exception:
                            pass
                        try:
                            result_cell.formula2 = case["formula"]
                        except Exception:
                            result_cell.formula = case["formula"]
                    except Exception as exc:
                        write_error = _format_com_error(exc)
                    self._app.calculate()
                    if write_error is not None:
                        out.append(CaseResult(id=case["id"], kind="skipped", value=write_error))
                        continue
                    try:
                        _write_spill_shape_probe(sht)
                        self._app.calculate()
                        result = _classify_result_cell(sht)
                        result.id = case["id"]
                        out.append(result)
                    except Exception as exc:
                        out.append(CaseResult(
                            id=case["id"], kind="skipped", value=_format_com_error(exc),
                        ))
                finally:
                    try:
                        wb.close()
                    except Exception:
                        pass
            return out
        finally:
            try:
                if prior_iter is not None:
                    self._app.api.Iteration = prior_iter
            except Exception:
                pass


def _split_sheet_qualified_addr(key: str) -> "tuple[Optional[str], str]":
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
    addr_part = key[bang + 1:]
    if sheet_part.startswith("'") and sheet_part.endswith("'") and len(sheet_part) >= 2:
        sheet_part = sheet_part[1:-1].replace("''", "'")
    return sheet_part, addr_part


def _get_or_add_sheet(wb, name: str):
    """Returns the sheet whose display name matches ``name`` (case-insensitive),
    adding it at the end if absent. The xlwings ``books.Sheets`` collection
    is itself case-insensitive on lookup, but we surface that behavior
    explicitly here for clarity.
    """

    target = name.casefold()
    for sht in wb.sheets:
        if sht.name.casefold() == target:
            return sht
    return wb.sheets.add(name=name, after=wb.sheets[len(wb.sheets) - 1])


def _write_cell(sht, addr: str, rec: Dict[str, Any]) -> None:
    """Writes one normalised {kind, value} record to ``sht!addr``."""

    kind = rec.get("kind")
    rng = sht.range(addr)
    if kind == "blank":
        rng.clear_contents()
        return
    if kind == "number":
        rng.value = float(rec["value"])
        return
    if kind == "bool":
        rng.value = bool(rec["value"])
        return
    if kind == "text":
        rng.value = str(rec["value"])
        return
    if kind == "formula":
        try:
            rng.formula2 = rec["formula"]
        except Exception:
            rng.formula = rec["formula"]
        return
    if kind == "error":
        trigger = _error_trigger(rec.get("code", "#VALUE!"))
        try:
            rng.formula2 = trigger
        except Exception:
            rng.formula = trigger
        return
    raise ValueError(f"unknown cell kind: {kind}")


def _format_com_error(exc: BaseException) -> str:
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
            return f"COM {hresult}: {descr}".strip().rstrip(":")
    except Exception:
        pass
    return f"{type(exc).__name__}: {exc}"[:200]


_SHEET_FORBIDDEN = set("\\/?*[]:")


def _sanitize_sheet_name(name: str) -> str:
    """Strips characters Excel disallows in sheet names and trims length."""

    cleaned = "".join("_" if c in _SHEET_FORBIDDEN else c for c in name)
    return cleaned[:24] or "case"


_ERROR_TRIGGERS = {
    "#DIV/0!": "=1/0",
    "#NAME?": "=NONEXISTENT_FUNC()",
    "#VALUE!": '=VALUE("x")',
    "#NUM!": "=SQRT(-1)",
    "#N/A": "=NA()",
    "#REF!": "=OFFSET(A1,-1,-1)",
    "#NULL!": "=A1 B1",
}


def _error_trigger(code: str) -> str:
    return _ERROR_TRIGGERS.get(code, '=VALUE("x")')


def _env_to_json(env: EnvironmentInfo) -> Dict[str, Any]:
    return {
        "excel_version": env.excel_version,
        "excel_locale": env.excel_locale,
        "date1904": env.date1904,
        "iterative": env.iterative,
    }


def _dispatch(drv: "WindowsExcelOracle", env_json: Dict[str, Any], payload: Dict[str, Any]) -> Dict[str, Any]:
    """Routes one decoded request to the driver and returns the JSON response."""

    if payload.get("version") != 1:
        raise RuntimeError(f"unsupported wire version: {payload.get('version')}")
    cmd = payload.get("command")
    if cmd == "probe_environment":
        return {"version": 1, "environment": env_json, "results": []}
    if cmd == "run_suite":
        results = drv.run_suite(
            payload["suite_name"],
            payload["cases"],
            date1904=bool(payload.get("date1904", False)),
            iterative=bool(payload.get("iterative", False)),
        )
        results_json = [
            {
                "id": r.id,
                "kind": r.kind,
                "value": r.value,
                "error_code": r.error_code,
                "array_shape": r.array_shape,
            }
            for r in results
        ]
        return {"version": 1, "environment": env_json, "results": results_json}
    raise RuntimeError(f"unknown command: {cmd!r}")


def _serve(visible: bool) -> int:
    """Long-lived server mode for the WSL bridge.

    Opens Excel once, writes a single ``ready`` line carrying the cached
    environment, then loops on stdin reading one JSON request per line
    and writing one JSON response per line. The ``shutdown`` command
    closes Excel and exits cleanly; EOF on stdin is treated as a hard
    shutdown to handle a parent that died without sending the explicit
    message.

    All protocol I/O is utf-8. The caller is expected to launch the
    interpreter with ``-X utf8=1`` (the bridge does so unconditionally)
    so sys.stdin/stdout encoding stays sane across Windows code pages.
    """

    import json
    import sys

    sys.stdout.reconfigure(encoding="utf-8", newline="\n")
    sys.stdin.reconfigure(encoding="utf-8")

    with WindowsExcelOracle(visible=visible) as drv:
        env_json = _env_to_json(drv.probe_environment())
        sys.stdout.write(
            json.dumps(
                {"version": 1, "type": "ready", "environment": env_json},
                ensure_ascii=False,
            )
            + "\n"
        )
        sys.stdout.flush()

        for line in sys.stdin:
            line = line.strip()
            if not line:
                continue
            payload = json.loads(line)
            if payload.get("command") == "shutdown":
                break
            try:
                resp = _dispatch(drv, env_json, payload)
            except Exception as exc:
                resp = {
                    "version": 1,
                    "type": "error",
                    "error": f"{type(exc).__name__}: {exc}",
                }
            sys.stdout.write(json.dumps(resp, ensure_ascii=False) + "\n")
            sys.stdout.flush()
    return 0


def _main(argv: Optional[List[str]] = None) -> int:
    """Wire-protocol entrypoint for the WSL bridge.

    Two modes:

      ``--input X.json --output Y.json``
        Legacy single-shot: read one command, execute, write one
        response, exit. Pays one Excel cold-start per call.

      ``--serve``
        Long-lived server: open Excel once, write ``ready``, then loop
        on stdin reading newline-delimited JSON requests and writing
        newline-delimited JSON responses. Exits on the ``shutdown``
        command or EOF.

    Supported commands in both modes:
      - ``probe_environment`` -- returns the cached environment block.
      - ``run_suite`` -- evaluates the supplied case batch and returns
        a vector of normalised :class:`CaseResult` records.

    Wire-format version is pinned to 1; any future schema change must
    bump it both here and in the WSL bridge.
    """

    import argparse
    import json

    parser = argparse.ArgumentParser(description="windows_excel wire driver")
    parser.add_argument("--input", type=Path,
                        help="single-shot mode: read one JSON command from this path")
    parser.add_argument("--output", type=Path,
                        help="single-shot mode: write one JSON response to this path")
    parser.add_argument("--serve", action="store_true",
                        help="long-lived mode: serve newline-delimited JSON over stdio")
    parser.add_argument("--visible", action="store_true",
                        help="show the Excel window (debug aid for --serve)")
    args = parser.parse_args(argv)

    if args.serve:
        if args.input or args.output:
            raise SystemExit("--serve cannot be combined with --input/--output")
        return _serve(visible=args.visible)

    if not (args.input and args.output):
        raise SystemExit("either --serve or both --input and --output are required")

    payload = json.loads(args.input.read_text(encoding="utf-8"))
    visible = bool(payload.get("visible", False))
    with WindowsExcelOracle(visible=visible) as drv:
        env_json = _env_to_json(drv.probe_environment())
        out = _dispatch(drv, env_json, payload)
    args.output.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    return 0


if __name__ == "__main__":
    import sys

    sys.exit(_main())
