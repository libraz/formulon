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
the case formula at ``Z1`` (or at the case's ``formula_cell`` override),
setup cells written verbatim at their absolute A1 addresses. A single
``app.calculate()`` resolves the whole batch.

## Wire-protocol entrypoints

Running ``python -m tools.oracle.drivers.windows_excel --input X.json
--output Y.json`` reads a command (``probe_environment``, ``run_suite``,
or ``run_workbook_case``) from the input JSON and writes the response
JSON to the output path.
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

import base64
import datetime as _dt
import os
import platform
import shutil
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from ._locale import detect_locale_from_app, normalise_error_token
from .base import (
    _ERR_DISPLAY_NAMES,
    DEFAULT_FORMULA_CELL,
    MAX_CAPTURE_CELLS,
    CaseResult,
    EnvironmentInfo,
    OracleDriver,
    _datetime_to_serial,
    case_formula_cell,
    case_shape_samples,
    probe_spill_shape,
)

try:
    import xlwings as xw  # type: ignore
except ImportError as exc:  # pragma: no cover - handled in oracle_gen.py
    xw = None  # type: ignore[assignment]
    _XLWINGS_IMPORT_ERROR: Optional[ImportError] = exc
else:
    _XLWINGS_IMPORT_ERROR = None


def _ensure_windows() -> None:
    """Refuses to start unless the host OS is Windows.

    The COM automation path is Windows-only; on Mac it would silently
    fall back to AppleScript, which has different attribute spellings
    and would mis-classify error cells. Refuse loudly instead.
    """

    if platform.system() != "Windows":
        raise RuntimeError(
            "windows_excel driver is Windows-only (uses COM automation). "
            "Current platform: " + platform.system() + ". Use wsl_bridge from WSL2, or macos_excel on Darwin."
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
            flat = [_array_cell_from_scalar(_classify_python_scalar(item)) for row in v for item in row]
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


def _evaluate_spill_shape(app, anchor, *, max_cells: Optional[int] = MAX_CAPTURE_CELLS) -> tuple[int, int]:
    """Windows adapter for the shared COM Application.Evaluate probe."""

    return probe_spill_shape(lambda expression: app.api.Evaluate(expression), anchor, max_cells=max_cells)


def _classify_shape_result(app, sht, anchor_addr: str, samples: List[str]) -> CaseResult:
    """Records a spill as its shape plus the case's declared sample cells.

    Mirrors :func:`tools.oracle.drivers.macos_excel._classify_shape_result`.
    """

    rows, cols = _evaluate_spill_shape(app, sht.range(anchor_addr), max_cells=None)
    values = {addr: _array_cell_from_scalar(_classify_value(sht.range(addr))) for addr in samples}
    return CaseResult(id="", kind="array_shape", value=values, array_shape=[rows, cols])


def _classify_case_result(app, sht, case: Dict[str, Any]) -> CaseResult:
    """Reads one case's result in whichever capture mode it declares."""

    anchor_addr = case_formula_cell(case)
    samples = case_shape_samples(case)
    if samples is not None:
        return _classify_shape_result(app, sht, anchor_addr, samples)
    return _classify_result_cell(app, sht, anchor_addr)


def _classify_result_cell(app, sht, anchor_addr: str = DEFAULT_FORMULA_CELL) -> CaseResult:
    """Classifies the anchor scalar or the full dynamic spill if present."""

    shape = _evaluate_spill_shape(app, sht.range(anchor_addr))
    rows, cols = shape
    if rows == 1 and cols == 1:
        return _classify_value(sht.range(anchor_addr))
    anchor = sht.range(anchor_addr)
    flat: List[Any] = []
    for r in range(rows):
        for c in range(cols):
            flat.append(_array_cell_from_scalar(_classify_value(anchor.offset(r, c))))
    return CaseResult(id="", kind="array", value=flat, array_shape=[rows, cols])


# Names the printer whose driver metrics Excel paginates against; see
# `WindowsExcelOracle._ensure_printer_pinned`. Excel addresses a device as
# "<name> on <port>:" and the port is assigned per machine, so a bare name
# is tried against the ports Windows hands out first.
_PRINTER_ENV_VAR = "FORMULON_EXCEL_PRINTER"
_PRINTER_PORTS = ("Ne00:", "Ne01:", "Ne02:", "Ne03:", "Ne04:", "Ne05:", "PORTPROMPT:")
# Shipped on every supported Windows build and, unlike a physical device,
# imposes no hardware unprintable margin -- so the printable area is the
# paper minus the margins the case asked for, which is also the model the
# C++ paginator implements.
_DEFAULT_PRINTER = "Microsoft Print to PDF"


class WindowsExcelOracle(OracleDriver):
    """Thin wrapper over a single hidden Excel.exe instance (COM).

    The instance is reused across suites so we pay COM startup latency
    only once. Call :meth:`close` (or use as a context manager) to tear
    down. Refuses to construct on non-Windows hosts.
    """

    def __init__(self, visible: bool = False) -> None:
        _ensure_windows()
        if xw is None:
            raise RuntimeError("xlwings is not installed; run `make oracle-setup` first") from _XLWINGS_IMPORT_ERROR
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
        self._printer_pinned = False
        self._printer_warned = False

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

        On Windows ``Application.Version`` is the Office *major* version --
        the literal ``"16.0"`` for every SKU from Office 2016 through
        Microsoft 365. On its own it identifies nothing, which is exactly
        the shape of stamp that let an Office 2019 capture pass for a
        Microsoft 365 one; ``tools/oracle/divergence_check.py`` rejects it
        by name. ``Application.Build`` carries the discriminating part, so
        it is always appended (2019 builds are 10xxx, Microsoft 365 18xxx),
        producing a ``16.0.<build>`` stamp that survives that check --
        three components, never more, because that check accepts at most a
        ``16.<n>.<n>`` shape.

        The locale is detected via
        ``Application.International(xlCountryCode)`` (see
        :func:`tools.oracle.drivers._locale.detect_locale_from_app`); a
        probe failure or unmapped country code yields an empty string,
        which the generator surfaces as a missing locale in
        ENVIRONMENT.md so the operator can investigate.
        """

        version = ""
        try:
            version = str(self._app.version).strip() if self._app.version else ""
        except Exception:
            version = ""
        build = ""
        try:
            api = self._app.api
            if not version:
                v = api.Version() if callable(getattr(api, "Version", None)) else getattr(api, "Version", "")
                version = str(v).strip()
            b = api.Build() if callable(getattr(api, "Build", None)) else getattr(api, "Build", "")
            build = str(b).strip()
        except Exception:
            build = ""
        # `Build` is normally the bare build number ("18025"), but the
        # channels disagree: this ja-JP Microsoft 365 install reports
        # "20228.0" (build plus a sub-build that does not even match the
        # EXCEL.EXE file version), and others report a whole dotted version
        # there. Only the leading component discriminates a Microsoft 365
        # install from an Office 2019 one, so a whole dotted Build replaces
        # the bare version rather than doubling its prefix, and the result
        # is cut to three components -- appending a sub-build would stamp a
        # `16.0.<build>.<sub>` that `is_pending_stamp` refuses outright.
        if build:
            if not version:
                stamp = build
            elif build.startswith(version):
                stamp = build
            elif build.split(".")[0] in version.split("."):
                stamp = version
            else:
                stamp = f"{version}.{build.split('.')[0]}"
            version = ".".join(stamp.split(".")[:3])
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
        cross_sheet = any(any("!" in addr for addr in (case.get("setup") or {})) for case in cases)
        if cross_sheet:
            return self._run_suite_per_case_workbook(suite_name, cases, date1904=date1904, iterative=iterative)

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
                    _apply_merges(sht, case.get("merges") or [])
                    setup = case.get("setup") or {}
                    for addr, rec in setup.items():
                        _write_cell(sht, addr, rec)
                    result_cell = sht.range(case_formula_cell(case))
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

            out: List[CaseResult] = []
            for case, sht in zip(cases, case_sheets):
                if case["id"] in write_errors:
                    out.append(
                        CaseResult(
                            id=case["id"],
                            kind="skipped",
                            value=write_errors[case["id"]],
                        )
                    )
                    continue
                try:
                    result = _classify_case_result(self._app, sht, case)
                    result.id = case["id"]
                    out.append(result)
                except Exception as exc:
                    out.append(
                        CaseResult(
                            id=case["id"],
                            kind="skipped",
                            value=_format_com_error(exc),
                        )
                    )
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
                        _apply_merges(sht, case.get("merges") or [])
                        for addr, rec in setup.items():
                            sheet_name, bare_addr = _split_sheet_qualified_addr(addr)
                            target_sht = sht if sheet_name is None else _get_or_add_sheet(wb, sheet_name)
                            _write_cell(target_sht, bare_addr, rec)
                        result_cell = sht.range(case_formula_cell(case))
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
                        result = _classify_case_result(self._app, sht, case)
                        result.id = case["id"]
                        out.append(result)
                    except Exception as exc:
                        out.append(
                            CaseResult(
                                id=case["id"],
                                kind="skipped",
                                value=_format_com_error(exc),
                            )
                        )
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

    # -----------------------------------------------------------------------
    # Workbook oracle track (pivot tables)
    # -----------------------------------------------------------------------

    def run_workbook_case(self, case: Dict[str, Any]) -> Dict[str, Any]:
        """Builds the declarative workbook via COM and reads it back.

        ``case`` is a declarative workbook spec (see
        ``tools/oracle/workbook_case_schema.py``). The case is dispatched
        on which feature block it carries -- ``pivot`` or ``print``. The
        returned mapping is the golden ``expect`` block, either:

            { "pivot": { "anchor": "Sheet2!A1", "rows": N, "cols": M,
                         "grid": [ {"r":0,"c":0,"value":{...}}, ... ] } }

        or:

            { "print": { "print_area": "A1:H80",
                         "h_breaks": [..], "v_breaks": [..], "pages": N } }

        Raises ``RuntimeError`` (catchable -- the generator marks the
        suite failed) when the case carries neither block or the COM
        automation fails.
        """

        pivot_spec = case.get("pivot")
        print_spec = case.get("print")
        roundtrip_spec = case.get("roundtrip")
        if isinstance(roundtrip_spec, dict):
            return self._run_roundtrip_case(case, roundtrip_spec)
        if isinstance(print_spec, dict):
            return self._run_print_case(case, print_spec)
        if isinstance(pivot_spec, dict):
            return self._run_pivot_case(case, pivot_spec)
        # The case schema marks both feature blocks optional (see
        # tests/oracle/cases_wb/README.md), so a no-feature case --
        # typically a schema smoke -- yields an empty expect block
        # rather than a per-case failure.
        return {}

    def _run_pivot_case(self, case: Dict[str, Any], pivot_spec: Dict[str, Any]) -> Dict[str, Any]:
        """Builds and reads back a PivotTable case via COM."""

        wb = self._app.books.add()
        try:
            self._build_workbook_sheets(wb, case)
            pivot = _build_pivot_table(wb, pivot_spec)
            grid, rows, cols = _read_pivot_grid(wb, pivot.TableRange2)
            expect: Dict[str, Any] = {
                "pivot": {
                    "anchor": pivot_spec.get("anchor", ""),
                    "rows": rows,
                    "cols": cols,
                    "grid": grid,
                }
            }
            probes = pivot_spec.get("formula_probes")
            if probes:
                expect["formula_probes"] = _run_formula_probes(wb, probes)
            return expect
        except Exception as exc:
            raise RuntimeError(
                f"pivot automation failed for case {case.get('id')!r}: {_format_com_error(exc)}"
            ) from exc
        finally:
            try:
                wb.close()
            except Exception:
                pass

    def _run_print_case(self, case: Dict[str, Any], print_spec: Dict[str, Any]) -> Dict[str, Any]:
        """Applies print settings via COM and reads back the pagination.

        Excel computes automatic page breaks itself; this reads the
        resolved ``PrintArea``, every ``HPageBreaks`` / ``VPageBreaks``
        location, and the physical page count.
        """

        wb = self._app.books.add()
        self._ensure_printer_pinned()
        try:
            self._build_workbook_sheets(wb, case)
            return {"print": _apply_and_read_print(wb, print_spec)}
        except Exception as exc:
            raise RuntimeError(
                f"print automation failed for case {case.get('id')!r}: {_format_com_error(exc)}"
            ) from exc
        finally:
            try:
                wb.close()
            except Exception:
                pass

    def _run_roundtrip_case(self, case: Dict[str, Any], spec: Dict[str, Any]) -> Dict[str, Any]:
        """Opens a Formulon-written xlsx and reads back what Excel made of it.

        This is the only capture path where Formulon's own bytes reach
        Excel. Every other workbook case builds the workbook with
        ``books.add()`` and COM calls, which measures how Excel behaves
        rather than whether Excel understands what we wrote.

        The fixture arrives base64-encoded in the case block -- the
        generator authors it on the repo side (see
        ``tools/oracle/print_roundtrip.py``), so the Excel host needs no
        Formulon install and the WSL bridge needs no path translation.

        Note what this does NOT capture: whether Excel showed a repair
        dialog. Automation runs with alerts suppressed, so a repaired file
        opens silently and reads back exactly like a healthy one. That
        judgement stays with the mechanical checks on our side (child-
        element order, relationship resolution) and a one-off manual open.
        """

        encoded = spec.get("xlsx_base64")
        if not isinstance(encoded, str) or not encoded:
            raise RuntimeError(
                f"roundtrip case {case.get('id')!r} carries no authored fixture; "
                "the generator attaches `xlsx_base64` before dispatching"
            )
        self._ensure_printer_pinned()

        tmp_dir = tempfile.mkdtemp(prefix="formulon-roundtrip-")
        # Excel keys several caches off the file name, and a name it has
        # seen in this session can hand back a stale book; the case id
        # keeps each fixture distinct within a suite run.
        fixture = Path(tmp_dir) / f"{_sanitize_sheet_name(str(case.get('id') or 'case'))}.xlsx"
        fixture.write_bytes(base64.b64decode(encoded))

        wb = None
        try:
            wb = self._app.books.open(str(fixture))
            sheet_name = spec.get("sheet")
            sht = wb.sheets[sheet_name] if isinstance(sheet_name, str) and sheet_name else wb.sheets[0]
            observed = _read_roundtrip(sht)
            observed["xlsx_sha256"] = spec.get("xlsx_sha256")
            observed["xlsx_bytes"] = spec.get("xlsx_bytes")
            return {"roundtrip": observed}
        except Exception as exc:
            raise RuntimeError(
                f"roundtrip automation failed for case {case.get('id')!r}: {_format_com_error(exc)}"
            ) from exc
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass
            shutil.rmtree(tmp_dir, ignore_errors=True)

    def _ensure_printer_pinned(self) -> None:
        """Points Excel at the printer named by ``FORMULON_EXCEL_PRINTER``.

        Excel does not paginate against the paper size alone: automatic
        breaks are computed from the *active printer driver's* metrics, so
        a device with hardware unprintable margins shrinks the usable area
        and a network device that renegotiates its capabilities can move
        the breaks between two otherwise identical captures. Pinning a
        fixed device -- "Microsoft Print to PDF" has no hardware margins
        and is present on every Windows install -- is what makes a capture
        reproducible.

        The value may name the device alone or carry Excel's ``on <port>:``
        suffix; the port differs per machine, so a bare name is tried
        against the usual ports. Leaving the variable unset keeps whatever
        the host defaults to.

        ``ActivePrinter`` is workbook-gated in the same way as the
        ``Calculation`` / ``EnableIterativeCalculation`` properties
        `__init__` documents: assigning to it while no book is open raises
        "cannot set the ActivePrinter property of the Application class",
        no matter how the device is spelled. Callers must therefore pin
        *after* `books.add()`, and the success flag is set only once a
        device actually took -- a failed attempt that latched the flag is
        what silently paginated every print suite against the host's
        default network device.
        """

        if self._printer_pinned:
            return
        requested = os.environ.get(_PRINTER_ENV_VAR, "").strip()
        explicit = bool(requested)
        if not explicit:
            requested = _DEFAULT_PRINTER
        candidates = [requested] if " on " in requested else [f"{requested} on {port}" for port in _PRINTER_PORTS]
        for candidate in candidates:
            try:
                self._app.api.ActivePrinter = candidate
            except Exception:
                continue
            self._printer_pinned = True
            return
        if explicit:
            raise RuntimeError(
                f"{_PRINTER_ENV_VAR}={requested!r} names no printer Excel would accept "
                f"(tried: {', '.join(candidates)}); pagination would run against "
                f"{self._app.api.ActivePrinter!r} instead"
            )
        # An unset variable is a preference, not an instruction: capture
        # against whatever the host defaults to, but say so, because the
        # goldens then carry that device's metrics. The flag stays unset
        # so a later case retries -- only a device that took pins the run.
        if not self._printer_warned:
            self._printer_warned = True
            print(
                f"  ! {_DEFAULT_PRINTER} is unavailable; paginating against "
                f"{self._app.api.ActivePrinter!r}, whose driver metrics the goldens will carry",
                file=sys.stderr,
            )

    def _build_workbook_sheets(self, wb, case: Dict[str, Any]) -> None:
        """Materialises the declarative ``sheets`` block into ``wb``.

        Also applies the optional ``column_widths`` / ``row_heights``
        layout maps. Cell values follow the normalised ``{kind, value}``
        shape ``_write_cell`` already understands.
        """

        _pin_normal_font(wb)

        sheets = case.get("sheets") or {}
        for sheet_name, cells in sheets.items():
            sht = _get_or_add_sheet(wb, sheet_name)
            for addr, rec in (cells or {}).items():
                _write_cell(sht, addr, rec)

        # Sizes are applied against the first sheet; workbook cases that
        # need a per-sheet layout map can extend the schema later.
        widths = case.get("column_widths") or {}
        for col_key, width in widths.items():
            _apply_axis_size(wb.sheets[0], "column", str(col_key), float(width))
        heights = case.get("row_heights") or {}
        for row_key, height in heights.items():
            _apply_axis_size(wb.sheets[0], "row", str(row_key), float(height))

        # Hiding is read back and asserted rather than attempted quietly: a
        # column that silently failed to hide yields a golden that looks
        # perfectly ordinary while answering a different question than the
        # case asks.
        for col_key in case.get("hidden_columns") or []:
            target = wb.sheets[0].range(f"{col_key}1").api.EntireColumn
            target.Hidden = True
            if not target.Hidden:
                raise RuntimeError(f"column {col_key!r} did not hide; refusing to capture a case it does not match")
        for row_key in case.get("hidden_rows") or []:
            target = wb.sheets[0].range(f"A{row_key}").api.EntireRow
            target.Hidden = True
            if not target.Hidden:
                raise RuntimeError(f"row {row_key!r} did not hide; refusing to capture a case it does not match")


_NORMAL_FONT_NAME = "Calibri"
_NORMAL_FONT_SIZE = 11.0


def _pin_normal_font(wb) -> None:
    """Pins the Normal style to Calibri 11 before any content is written.

    A column width is stated in *characters*, which is a font-relative
    unit: it resolves through the Normal style font's maximum digit width.
    A ja-JP Excel opens new workbooks in the locale's UI font, so a case
    saying `column_widths: {"A:H": 30}` produced a physically different
    sheet on the capture host than `src/print/pagination.cpp` builds from
    the same number -- the engine's width model is the documented Calibri
    11 one (MDW 7px, the 8.43-characters == 64-pixels default). The two
    sides were resolving the same unit against different fonts, so every
    width-derived page break was compared across incompatible geometry.

    Pinning the font here makes the case self-describing rather than
    dependent on the capture host's locale defaults. Excel resizes existing
    columns when the Normal font changes, so this must run before any width
    is applied.
    """

    style = wb.api.Styles("Normal")
    style.Font.Name = _NORMAL_FONT_NAME
    style.Font.Size = _NORMAL_FONT_SIZE
    applied = str(style.Font.Name)
    if applied != _NORMAL_FONT_NAME:
        raise RuntimeError(
            f"Normal style font is {applied!r}, not {_NORMAL_FONT_NAME!r}; character-unit column "
            "widths would not be comparable with the engine's width model"
        )


def _apply_axis_size(sheet, axis: str, key: str, size: float) -> None:
    """Sets one column width / row height and asserts Excel took it.

    The key is a label or a span -- ``"A"``, ``"A:H"``, ``"3"``, ``"3:5"``
    -- which is what ``Columns`` / ``Rows`` accept. The address form this
    used to build (``f"{key}1"``) cannot express a span at all: a case
    declaring ``"A:H"`` produced the invalid reference ``"A:H1"``, threw,
    and was swallowed, so the capture ran at Excel's default width and the
    golden answered a different question than the case asks. The read-back
    is what keeps that failure from looking like an ordinary golden; the
    tolerance covers Excel snapping a size to its own grid.
    """

    axis_api = sheet.api.Columns if axis == "column" else sheet.api.Rows
    prop = "ColumnWidth" if axis == "column" else "RowHeight"
    target = axis_api(key)
    setattr(target, prop, size)
    applied = getattr(target, prop)
    # A span whose members disagree reads back as None; that cannot happen
    # here because every member was just assigned the same size.
    if applied is None or abs(float(applied) - size) > 1.0:
        raise RuntimeError(
            f"{axis} {key!r} kept {applied!r} instead of the requested {size}; "
            "refusing to capture a case it does not match"
        )


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
    addr_part = key[bang + 1 :]
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


def _apply_merges(sht, merges: List[str]) -> None:
    """Apply case-declared inclusive A1 merge ranges on ``sht``."""

    for ref in merges:
        sht.range(ref).merge()


# Excel `XlConsolidationFunction` constants, keyed by the declarative
# `agg` name. Mirrors the `Aggregation` enum the C++ builder accepts.
# Numeric values are the documented Office automation constants so the
# driver does not depend on the `win32com` constant cache being warm.
_XL_CONSOLIDATION = {
    "Sum": -4157,
    "Count": -4112,
    "Average": -4106,
    "Max": -4136,
    "Min": -4139,
    "Product": -4149,
    "CountNumbers": -4113,
    "StdDev": -4155,
    "StdDevP": -4156,
    "Var": -4164,
    "VarP": -4165,
}

# Excel `XlPivotFieldOrientation` constants.
_XL_ORIENT_ROW = 1
_XL_ORIENT_COLUMN = 2
_XL_ORIENT_PAGE = 3
_XL_ORIENT_DATA = 4

# Excel `XlPivotTableSourceType.xlDatabase`.
_XL_DATABASE = 1

# Excel `XlLayoutRowType` values for `RowAxisLayout`.
_XL_COMPACT_ROW = 0
_XL_TABULAR_ROW = 1
_XL_OUTLINE_ROW = 2

_PIVOT_LAYOUT_MODES = {
    "Compact": _XL_COMPACT_ROW,
    "Tabular": _XL_TABULAR_ROW,
    "Outline": _XL_OUTLINE_ROW,
}


def _build_pivot_table(wb, pivot_spec: Dict[str, Any]):
    """Creates a PivotTable in ``wb`` from a declarative ``pivot`` block.

    Returns the materialised ``PivotTable`` COM object. The caller reads
    ``TableRange2`` for the rendered grid and may then run post-build
    formula probes. The ``pivot`` block shape is documented on
    ``tests/oracle/workbook_builder.h``.
    """

    source = pivot_spec.get("source")
    anchor = pivot_spec.get("anchor")
    if not isinstance(source, str) or not source:
        raise RuntimeError("pivot block missing string 'source'")
    if not isinstance(anchor, str) or not anchor:
        raise RuntimeError("pivot block missing string 'anchor'")

    # Resolve the declarative anchor "Report!A1" to the destination
    # cell's COM Range, since CreatePivotTable accepts a Range here and
    # the string form has the same fully-qualified requirement as
    # SourceData. We also drive the source resolution through a Range
    # object so PivotCaches.Create does not have to parse the address
    # itself.
    src_sheet_name, src_addr = _split_sheet_qualified_addr(source)
    anchor_sheet_name, anchor_addr = _split_sheet_qualified_addr(anchor)
    if src_sheet_name is None:
        raise RuntimeError(f"pivot source must be sheet-qualified (e.g. 'Data!A1:C13'); got {source!r}")
    if anchor_sheet_name is None:
        raise RuntimeError(f"pivot anchor must be sheet-qualified (e.g. 'Report!A1'); got {anchor!r}")

    def _find_sheet(name: str):
        for sht in wb.sheets:
            if sht.name.casefold() == name.casefold():
                return sht
        return None

    src_sheet = _find_sheet(src_sheet_name)
    anchor_sheet = _find_sheet(anchor_sheet_name)
    if src_sheet is None:
        raise RuntimeError(f"pivot source references unknown sheet {src_sheet_name!r}")
    if anchor_sheet is None:
        raise RuntimeError(f"pivot anchor references unknown sheet {anchor_sheet_name!r}")

    source_range_api = src_sheet.range(src_addr).api
    anchor_range_api = anchor_sheet.range(anchor_addr).api

    # Activate the anchor sheet so CreatePivotTable's destination is on
    # the active sheet -- some Excel builds reject creating a pivot
    # whose destination is on an inactive sheet with E_INVALIDARG.
    try:
        anchor_sheet.activate()
    except Exception:
        pass

    api = wb.api
    cache = api.PivotCaches().Create(_XL_DATABASE, source_range_api)
    pivot = cache.CreatePivotTable(anchor_range_api, "FormulonPivot")

    for field_name in pivot_spec.get("row_fields") or []:
        pivot.PivotFields(field_name).Orientation = _XL_ORIENT_ROW
    for field_name in pivot_spec.get("col_fields") or []:
        pivot.PivotFields(field_name).Orientation = _XL_ORIENT_COLUMN
    for field_name in pivot_spec.get("page_fields") or []:
        pivot.PivotFields(field_name).Orientation = _XL_ORIENT_PAGE

    for data_field in pivot_spec.get("data_fields") or []:
        field_name = data_field.get("field")
        agg = data_field.get("agg", "Sum")
        consolidation = _XL_CONSOLIDATION.get(agg)
        if consolidation is None:
            raise RuntimeError(f"unknown aggregation {agg!r}")
        field = pivot.PivotFields(field_name)
        field.Orientation = _XL_ORIENT_DATA
        field.Function = consolidation

    # Manual item filters: hide the named items on their field.
    for filter_spec in pivot_spec.get("filters") or []:
        field_name = filter_spec.get("field")
        for item_name in filter_spec.get("hide") or []:
            try:
                pivot.PivotFields(field_name).PivotItems(item_name).Visible = False
            except Exception:
                pass

    layout = pivot_spec.get("layout")
    if isinstance(layout, str) and layout in _PIVOT_LAYOUT_MODES:
        try:
            pivot.RowAxisLayout(_PIVOT_LAYOUT_MODES[layout])
        except Exception:
            pass

    grand = pivot_spec.get("grand_totals") or {}
    if isinstance(grand, dict):
        if "rows" in grand:
            pivot.RowGrand = bool(grand["rows"])
        if "cols" in grand:
            pivot.ColumnGrand = bool(grand["cols"])

    pivot.RefreshTable()
    return pivot


def _run_formula_probes(wb, probes: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Write and read post-build formulas against a materialised PivotTable.

    Formula probes are intentionally evaluated only after ``RefreshTable``.
    A rendered PivotTable grid cannot exercise GETPIVOTDATA's page/data-axis
    routing, whereas a formula cell can.  The result shape matches the
    scalar grid records: ``{"kind": ..., "value": ...}`` or
    ``{"kind": "error", "code": "#REF!"}``.
    """

    if not isinstance(probes, list) or not probes:
        return []
    out: List[Dict[str, Any]] = []
    for probe in probes:
        if not isinstance(probe, dict):
            raise RuntimeError("pivot formula_probes entries must be objects")
        probe_id = probe.get("id")
        cell_ref = probe.get("cell")
        formula = probe.get("formula")
        if not isinstance(probe_id, str) or not probe_id:
            raise RuntimeError("pivot formula probe missing string 'id'")
        if not isinstance(cell_ref, str) or "!" not in cell_ref:
            raise RuntimeError(f"formula probe {probe_id!r} needs sheet-qualified cell")
        if not isinstance(formula, str) or not formula:
            raise RuntimeError(f"formula probe {probe_id!r} missing string 'formula'")
        sheet_name, bare_addr = _split_sheet_qualified_addr(cell_ref)
        target = None
        for sht in wb.sheets:
            if sht.name.casefold() == sheet_name.casefold():
                target = sht
                break
        if target is None:
            raise RuntimeError(f"formula probe {probe_id!r} references unknown sheet {sheet_name!r}")
        result_cell = target.range(bare_addr)
        try:
            result_cell.formula2 = formula
        except Exception:
            result_cell.formula = formula

    try:
        wb.app.calculate()
    except Exception as exc:
        # Do not silently emit stale values from a prior probe. A formula
        # result such as #REF! is a normal cell value and does not raise here;
        # a COM calculation failure is a driver failure and must be visible.
        raise RuntimeError(f"Excel failed to calculate formula probes: {_format_com_error(exc)}") from exc

    for probe in probes:
        sheet_name, bare_addr = _split_sheet_qualified_addr(probe["cell"])
        target = next(sht for sht in wb.sheets if sht.name.casefold() == sheet_name.casefold())
        result = _classify_value(_CellAdapter(target.range(bare_addr).api))
        out.append({"id": probe["id"], "cell": probe["cell"], "result": _grid_value_record(result)})
    return out


# Excel `XlPageOrientation` constants.
_XL_PORTRAIT = 1
_XL_LANDSCAPE = 2

_PRINT_ORIENTATIONS = {
    "portrait": _XL_PORTRAIT,
    "landscape": _XL_LANDSCAPE,
}

_PRINT_ORIENTATION_NAMES = {value: name for name, value in _PRINT_ORIENTATIONS.items()}

# `GET.DOCUMENT(50)` is the Excel-4 macro that returns the page count;
# used as a fallback when `PageSetup.Pages.Count` is unavailable.
_GET_DOCUMENT_PAGE_COUNT = 50

# `XlWindowView.xlPageBreakPreview` -- the view mode that forces Excel
# to actually paginate, so subsequent reads of HPageBreaks /
# VPageBreaks / Pages.Count reflect the print layout rather than the
# editing layout.
_XL_PAGE_BREAK_PREVIEW = 2

# PageBreakPreview's "factory default" window zoom. The Application
# remembers the last PBP zoom across workbooks; pinning to 60 makes
# the break read independent of which case came before.
_XL_PAGE_BREAK_PREVIEW_ZOOM = 60

# Excel COM reports `PageSetup.LeftMargin` etc. in points (72 pt = 1
# inch). The Formulon `PageMargins` struct works in inches, so we
# divide on read.
_POINTS_PER_INCH = 72.0


# How many times a case's pagination is read before the capture gives up
# on it settling. Two consecutive agreeing reads end the loop, so a stable
# case costs two reads and one wait.
_BREAK_READ_ATTEMPTS = 6

# How long to leave Excel alone between two pagination reads.
_LAYOUT_SETTLE_SECONDS = 0.3


def _settle_layout(ws, wb) -> None:
    """Nudges Excel into finishing its pagination pass before a read.

    Recalculates, then flips ``ScreenUpdating`` off and on: Excel uses
    that transition as a trigger to re-run the layout, without which the
    break collections can answer from the previous case's last pass.

    The wait afterwards is load-bearing rather than superstition: Excel
    lays a page out on its own schedule, so a tight read-again loop can
    take two consecutive readings from the same not-yet-recomputed
    layout. `tall_and_wide_table` alternated between its own portrait
    pagination and the preceding landscape case's until the reads were
    spaced out.
    """

    try:
        ws.Calculate()
    except Exception:
        pass
    try:
        wb.app.api.ScreenUpdating = False
        wb.app.api.ScreenUpdating = True
    except Exception:
        pass
    time.sleep(_LAYOUT_SETTLE_SECONDS)


def _read_pagination(ws, page_setup, wb) -> Tuple[List[int], List[int], int]:
    """Reads ``(h_breaks, v_breaks, pages)`` as Excel currently reports it.

    Breaks come back as the 0-based row / column indices the C++ engine
    reports; ``Pages.Count`` falls back to the XLM ``GET.DOCUMENT`` page
    count on hosts where the property is unavailable. Any read failure
    yields an empty axis rather than raising -- the caller's settle loop
    is what decides whether a reading is trustworthy.
    """

    h_breaks: List[int] = []
    try:
        for i in range(1, int(ws.HPageBreaks.Count) + 1):
            # `.Location.Row` is the 1-based row the break precedes.
            h_breaks.append(int(ws.HPageBreaks(i).Location.Row) - 1)
    except Exception:
        pass
    v_breaks: List[int] = []
    try:
        for i in range(1, int(ws.VPageBreaks.Count) + 1):
            v_breaks.append(int(ws.VPageBreaks(i).Location.Column) - 1)
    except Exception:
        pass
    h_breaks.sort()
    v_breaks.sort()

    pages = 0
    try:
        pages = int(page_setup.Pages.Count)
    except Exception:
        try:
            pages = int(wb.app.api.ExecuteExcel4Macro(f"GET.DOCUMENT({_GET_DOCUMENT_PAGE_COUNT})"))
        except Exception:
            pages = 0
    return h_breaks, v_breaks, pages


def _read_pagination_settled(ws, page_setup, wb) -> Tuple[List[int], List[int], int]:
    """Reads the pagination and keeps only a value that reproduces.

    Every nudge the caller performs -- PageBreakPreview, the pinned PBP
    zoom, recalc, the ScreenUpdating flip -- is a way of asking Excel to
    finish paginating before the read, and none of them is a guarantee: a
    stale read still surfaced intermittently, with
    `print_titles_repeat_rows` returning the preceding landscape case's
    break at full-suite position while every solo capture returned its
    own. So the read itself decides. The triple is read repeatedly with a
    settle between attempts and only a reading two consecutive reads
    agree on is returned; the first read doubles as the dummy touch that
    populates Excel's cache. A case that never settles raises, because a
    golden nobody can reproduce is worse than a missing one.
    """

    reading: Optional[Tuple[List[int], List[int], int]] = None
    for _attempt in range(_BREAK_READ_ATTEMPTS):
        candidate = _read_pagination(ws, page_setup, wb)
        if candidate == reading:
            return candidate
        reading = candidate
        _settle_layout(ws, wb)
    raise RuntimeError(
        f"page breaks did not settle after {_BREAK_READ_ATTEMPTS} reads (last: {reading}); "
        "refusing to capture a golden that does not reproduce"
    )


def _resolve_print_sheet(wb, print_spec: Dict[str, Any]):
    """Returns the worksheet the `print` block names, or raises."""

    sheet_name = print_spec.get("sheet")
    if not isinstance(sheet_name, str) or not sheet_name:
        raise RuntimeError("print block missing string 'sheet'")
    target = sheet_name.casefold()
    for sht in wb.sheets:
        if sht.name.casefold() == target:
            return sht
    raise RuntimeError(f"print 'sheet' names an unknown sheet {sheet_name!r}")


def _apply_and_read_print(wb, print_spec: Dict[str, Any]) -> Dict[str, Any]:
    """Applies the declarative `print` block to ``wb`` and reads the result.

    Returns the golden ``expect.print`` mapping:

        { "print_area": "A1:H80", "h_breaks": [..0-based rows..],
          "v_breaks": [..0-based cols..], "pages": N }

    Page break locations come straight from Excel's automatic
    pagination (``HPageBreaks`` / ``VPageBreaks``); the 1-based COM
    row / column indices are converted to the 0-based indices the C++
    engine reports.
    """

    sht = _resolve_print_sheet(wb, print_spec)
    ws = sht.api
    page_setup = ws.PageSetup

    # State-leakage defense: clear any automatic / manual breaks the Excel
    # Application instance may have cached from a prior case, and pin the
    # editing-view break overlay off so it cannot interact with the
    # PageBreakPreview read below. Each case opens a fresh workbook, but
    # the Application is shared across the suite -- without these the
    # round-2 capture showed identical-spec cases producing different
    # `pages` / `v_breaks` depending on suite position (see the round-3
    # handoff). Done BEFORE manual_breaks application so a case-spec
    # manual break is not wiped by this reset.
    try:
        ws.ResetAllPageBreaks()
    except Exception:
        pass
    try:
        wb.app.api.DisplayPageBreaks = False
    except Exception:
        pass

    # --- print area ----------------------------------------------------------
    print_area = print_spec.get("print_area")
    if isinstance(print_area, str) and print_area:
        page_setup.PrintArea = print_area

    # --- print titles --------------------------------------------------------
    titles = print_spec.get("print_titles")
    if isinstance(titles, dict):
        rows = titles.get("rows")
        cols = titles.get("cols")
        if isinstance(rows, str) and rows:
            page_setup.PrintTitleRows = f"{sht.name}!{rows}"
        if isinstance(cols, str) and cols:
            page_setup.PrintTitleColumns = f"{sht.name}!{cols}"

    # --- page setup ----------------------------------------------------------
    setup = print_spec.get("page_setup")
    if isinstance(setup, dict):
        orientation = setup.get("orientation")
        if orientation in _PRINT_ORIENTATIONS:
            page_setup.Orientation = _PRINT_ORIENTATIONS[orientation]
        if setup.get("paper") is not None:
            try:
                page_setup.PaperSize = int(setup["paper"])
            except Exception:
                pass
        # Pre-baseline: clear FitToPages and pin Zoom=100 before applying
        # the caller's intent. Excel's workbook template default is
        # `FitToPages=ON` in some installs, which would contaminate a
        # Zoom-only or Fit-OFF case where the caller never sets the Fit
        # axis explicitly. See tests/oracle/cases_wb/print_matrix.* --
        # the per-block "pin everything else" guarantee relies on this.
        try:
            page_setup.FitToPagesWide = False
        except Exception:
            pass
        try:
            page_setup.FitToPagesTall = False
        except Exception:
            pass
        try:
            page_setup.Zoom = 100
        except Exception:
            pass
        fit_w = int(setup.get("fit_to_width") or 0)
        fit_h = int(setup.get("fit_to_height") or 0)
        if fit_w or fit_h:
            try:
                page_setup.Zoom = False
            except Exception:
                pass
            try:
                page_setup.FitToPagesWide = fit_w if fit_w else False
            except Exception:
                pass
            try:
                page_setup.FitToPagesTall = fit_h if fit_h else False
            except Exception:
                pass
        elif setup.get("scale") is not None:
            try:
                page_setup.Zoom = int(setup["scale"])
            except Exception:
                pass

    # --- manual breaks -------------------------------------------------------
    breaks = print_spec.get("manual_breaks")
    if isinstance(breaks, dict):
        for row1 in breaks.get("rows") or []:
            try:
                ws.Rows(int(row1)).PageBreak = 1  # xlPageBreakManual
            except Exception:
                pass
        for col_letter in breaks.get("cols") or []:
            try:
                ws.Columns(str(col_letter)).PageBreak = 1
            except Exception:
                pass

    # --- read back -----------------------------------------------------------
    # Touch ResetAllPageBreaks-free; reading the collections forces Excel
    # to recompute the automatic layout.
    resolved_area = ""
    try:
        resolved_area = str(page_setup.PrintArea or "")
    except Exception:
        resolved_area = ""
    resolved_area = _normalise_print_area(resolved_area)

    # Force a true pagination pass by activating the sheet and switching
    # the window into PageBreakPreview view. Without this, at low Zoom
    # (e.g. 25%, 50%) `PageSetup.Pages.Count` returns inflated values that
    # do NOT reconcile with `HPageBreaks` / `VPageBreaks` -- the count is
    # influenced by the display zoom rather than the printable layout.
    # PageBreakPreview makes Excel actually paginate before we read, so
    # Pages.Count matches (H+1)*(V+1). Restored afterwards.
    #
    # Additionally pin the PBP-mode window zoom to its default (60). PBP
    # remembers the last zoom used by the *Application*, not the workbook,
    # so a prior case that left PBP at a different zoom would feed back
    # into the next case's break read. Round 2 showed identical-spec
    # cases producing different `v_breaks` depending on suite position;
    # pinning the zoom is the load-bearing piece of this fix.
    prior_view = None
    try:
        sht.activate()
        active_window = wb.app.api.ActiveWindow
        prior_view = active_window.View
        active_window.View = _XL_PAGE_BREAK_PREVIEW
        try:
            active_window.Zoom = _XL_PAGE_BREAK_PREVIEW_ZOOM
        except Exception:
            pass
    except Exception:
        prior_view = None

    try:
        # Force a layout settle before reading: recalc, then flip
        # ScreenUpdating off-then-on. Excel uses the ScreenUpdating
        # transition as a trigger to re-run the pagination layout; without
        # it the HPageBreaks / VPageBreaks collections can return stale
        # counts from the prior case's last layout pass.
        _settle_layout(ws, wb)

        h_breaks, v_breaks, pages = _read_pagination_settled(ws, page_setup, wb)
    finally:
        if prior_view is not None:
            try:
                wb.app.api.ActiveWindow.View = prior_view
            except Exception:
                pass

    # Round-trip read what Excel actually applied. Without this, a case
    # where Excel disagreed with the spec (e.g. FitToPages override of
    # Zoom, or workbook-template margin overrides) is indistinguishable
    # from a true engine bug. Load-bearing diagnostic for the print_matrix
    # follow-up matrix (see print_matrix.FOLLOWUP.HANDOFF.md, removed
    # once the second round of goldens landed).
    applied: Dict[str, Any] = {
        "orientation": _read_orientation_value(page_setup),
        "paper": _read_paper_value(page_setup),
        "zoom": _read_zoom_value(page_setup),
        "fit_to_width": _read_fit_value(page_setup, "FitToPagesWide"),
        "fit_to_height": _read_fit_value(page_setup, "FitToPagesTall"),
        "margins": _read_margins(page_setup),
    }

    return {
        "print_area": resolved_area,
        "h_breaks": h_breaks,
        "v_breaks": v_breaks,
        "pages": pages,
        "applied_page_setup": applied,
        "applied_geometry": _read_applied_geometry(sht, resolved_area),
    }


def _read_applied_geometry(sht, resolved_area: str) -> Dict[str, Any]:
    """Records the physical sizes Excel resolved, in points.

    Break positions alone cannot say *why* a page broke where it did: they
    only bound the track sizes that produced them, so a model that
    disagrees can only be diagnosed by solving inequalities across the
    whole suite -- and an under-determined system sends you back for
    another capture. These are the numbers `src/print/pagination.cpp`
    predicts (`ColumnCharsToPoints`, `RowHeightPoints`,
    `compute_printable_area`), read directly, so a disagreement is
    attributable offline from one capture instead of another Excel
    session.

    Best-effort: a probe failure omits its key rather than failing the
    case, since this is diagnostic and never asserted.
    """

    out: Dict[str, Any] = {}
    ws = sht.api

    def points(getter) -> Any:
        try:
            value = getter()
            return round(float(value), 4) if value is not None else None
        except Exception:
            return None

    # Per-track sizes over the resolved print area's first rectangle. The
    # engine's model is per-track, so a summed width would hide which
    # track it disagrees about.
    first_rect = resolved_area.split(",")[0] if resolved_area else ""
    if ":" in first_rect:
        try:
            start, end = first_rect.split(":")
            first_row, first_col = _split_a1(start)
            last_row, last_col = _split_a1(end)
        except Exception:
            first_row = first_col = last_row = last_col = None
        if first_col is not None and last_col - first_col < _GEOMETRY_TRACK_CAP:
            out["column_widths_pt"] = [
                points(lambda c=col: ws.Columns(c).Width) for col in range(first_col, last_col + 1)
            ]
        if first_row is not None and last_row - first_row < _GEOMETRY_TRACK_CAP:
            out["row_heights_pt"] = [points(lambda r=row: ws.Rows(r).Height) for row in range(first_row, last_row + 1)]

    # The device the breaks were measured on. Excel paginates against the
    # active printer driver's metrics, not the paper size alone, so a
    # golden captured on another device is not comparable -- recording the
    # name is what lets that be seen from the golden instead of inferred
    # from a break that moved. (`PageSetup.PageWidth` / `PageHeight` stood
    # here and always read as null: they are Word properties, absent from
    # Excel's object model, so the printable-body diagnostic they were
    # meant to provide never existed. The applied orientation and paper in
    # `applied_page_setup` carry that separation instead.)
    try:
        out["printer"] = str(sht.book.app.api.ActivePrinter)
    except Exception:
        pass
    # The Normal style font is what makes a character-unit width physical;
    # recording it is what says whether _pin_normal_font took effect.
    try:
        style = sht.book.api.Styles("Normal").Font
        out["normal_font"] = f"{style.Name} {float(style.Size):g}"
    except Exception:
        pass
    return out


# A whole-column print area would otherwise issue 16,384 COM round trips
# for a diagnostic nobody asserts.
_GEOMETRY_TRACK_CAP = 64


def _split_a1(addr: str) -> Tuple[int, int]:
    """Splits a bare A1 address into 1-based (row, column) COM indices."""

    addr = addr.replace("$", "")
    i = 0
    while i < len(addr) and addr[i].isalpha():
        i += 1
    col = 0
    for ch in addr[:i]:
        col = col * 26 + (ord(ch.upper()) - ord("A") + 1)
    return int(addr[i:]), col


def _read_orientation_value(page_setup) -> Any:
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


def _read_paper_value(page_setup) -> Any:
    """Returns the post-apply `PaperSize` as the `XlPaperSize` int.

    The companion to the orientation read: paper and orientation together
    are what fix the page the breaks were measured on.
    """

    try:
        return int(page_setup.PaperSize)
    except Exception:
        return None


def _read_zoom_value(page_setup) -> Any:
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


def _read_fit_value(page_setup, name: str) -> Any:
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


def _read_margins(page_setup) -> Dict[str, Any]:
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


def _com_scalar(owner, attr: str) -> Any:
    """Reads one COM property, mapping an unavailable one to ``None``.

    A property Excel does not expose on this host must not abort the
    capture: the golden records `null` and the comparison skips that
    field, which is honest about what was observed.
    """

    try:
        return getattr(owner, attr)
    except Exception:
        return None


def _com_bool(owner, attr: str) -> Optional[bool]:
    value = _com_scalar(owner, attr)
    return bool(value) if isinstance(value, bool) else None


def _com_int(owner, attr: str) -> Optional[int]:
    value = _com_scalar(owner, attr)
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    return int(value)


def _com_text(owner, attr: str) -> Optional[str]:
    value = _com_scalar(owner, attr)
    return value if isinstance(value, str) else None


def _read_header_footer(page_setup) -> Dict[str, Any]:
    """Reads the six header/footer sections as Excel reports them.

    OOXML stores one string per section with `&L` / `&C` / `&R` markers
    inside it; COM splits the same content across three properties. The
    golden records the split form, because that is what Excel actually
    parsed the string into -- collapsing it back would hide a case where
    Excel put our text in the wrong third.
    """

    out: Dict[str, Any] = {
        "different_odd_even": _com_bool(page_setup, "OddAndEvenPagesHeaderFooter"),
        "different_first": _com_bool(page_setup, "DifferentFirstPageHeaderFooter"),
        "scale_with_doc": _com_bool(page_setup, "ScaleWithDocHeaderFooter"),
        "align_with_margins": _com_bool(page_setup, "AlignMarginsHeaderFooter"),
    }
    for class_name, page_attr in _HEADER_FOOTER_CLASSES:
        owner = page_setup if page_attr is None else _com_scalar(page_setup, page_attr)
        for band in ("Header", "Footer"):
            for position in _HEADER_FOOTER_POSITIONS:
                key = f"{class_name}_{band.lower()}_{position.lower()}"
                if owner is None:
                    out[key] = None
                    continue
                if page_attr is None:
                    # The odd/primary sections are plain strings on
                    # `PageSetup` itself.
                    out[key] = _com_text(owner, f"{position}{band}")
                    continue
                section = _com_scalar(owner, f"{position}{band}")
                out[key] = None if section is None else _com_text(section, "Text")
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
            if _com_int(brk, "Type") == _XL_PAGE_BREAK_MANUAL:
                rows.append(int(brk.Location.Row) - 1)
    except Exception:
        pass
    cols: List[int] = []
    try:
        for i in range(1, int(ws.api.VPageBreaks.Count) + 1):
            brk = ws.api.VPageBreaks(i)
            if _com_int(brk, "Type") == _XL_PAGE_BREAK_MANUAL:
                cols.append(int(brk.Location.Column) - 1)
    except Exception:
        pass
    rows.sort()
    cols.sort()
    return {"manual_row_breaks": rows, "manual_col_breaks": cols}


def _read_roundtrip(sht) -> Dict[str, Any]:
    """Reads every print setting Excel resolved from an opened workbook."""

    page_setup = sht.api.PageSetup
    print_area = _com_text(page_setup, "PrintArea") or ""
    title_rows = _com_text(page_setup, "PrintTitleRows") or ""
    title_cols = _com_text(page_setup, "PrintTitleColumns") or ""

    observed: Dict[str, Any] = {
        "page_setup": {
            "paper_size": _com_int(page_setup, "PaperSize"),
            "orientation": _com_int(page_setup, "Orientation"),
            "zoom": _read_zoom_value(page_setup),
            "fit_to_pages_wide": _read_fit_value(page_setup, "FitToPagesWide"),
            "fit_to_pages_tall": _read_fit_value(page_setup, "FitToPagesTall"),
        },
        "page_margins": _read_margins(page_setup),
        "print_options": {
            "grid_lines": _com_bool(page_setup, "PrintGridlines"),
            "headings": _com_bool(page_setup, "PrintHeadings"),
            "horizontal_centered": _com_bool(page_setup, "CenterHorizontally"),
            "vertical_centered": _com_bool(page_setup, "CenterVertically"),
        },
        "header_footer": _read_header_footer(page_setup),
        "print_area": _normalise_print_area(print_area),
        "print_title_rows": _normalise_print_area(title_rows),
        "print_title_cols": _normalise_print_area(title_cols),
    }
    observed.update(_read_manual_breaks(sht))
    return observed


def _normalise_print_area(area: str) -> str:
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


def _read_pivot_grid(wb, table_range) -> "tuple[List[Dict[str, Any]], int, int]":
    """Reads every cell of ``table_range`` into an anchor-relative grid.

    Returns ``(grid, rows, cols)`` where ``grid`` is a list of
    ``{"r", "c", "value"}`` records (``r`` / ``c`` are 0-based offsets
    from the pivot anchor) and ``value`` is the normalised ``{kind,
    value}`` record produced by ``_classify_value``.
    """

    rows = int(table_range.Rows.Count)
    cols = int(table_range.Columns.Count)
    top = int(table_range.Row)
    left = int(table_range.Column)
    sht = table_range.Worksheet

    grid: List[Dict[str, Any]] = []
    for r in range(rows):
        for c in range(cols):
            com_cell = sht.Cells(top + r, left + c)
            result = _classify_value(_CellAdapter(com_cell))
            grid.append(
                {
                    "r": r,
                    "c": c,
                    "value": _grid_value_record(result),
                }
            )
    return grid, rows, cols


def _grid_value_record(result: CaseResult) -> Dict[str, Any]:
    """Shapes a ``CaseResult`` into the golden grid's ``{kind, value}``."""

    if result.kind == "blank":
        return {"kind": "blank"}
    if result.kind == "error":
        return {"kind": "error", "code": result.error_code or "#UNKNOWN!"}
    return {"kind": result.kind, "value": result.value}


class _CellAdapter:
    """Adapts a raw COM ``Range`` to the small surface ``_classify_value``
    expects (a ``.value`` attribute plus the COM passthrough used by the
    error / displayed-text fallbacks).
    """

    def __init__(self, com_cell) -> None:
        self.api = com_cell

    @property
    def value(self) -> Any:
        return self.api.Value


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
    if cmd == "run_workbook_case":
        expect = drv.run_workbook_case(payload["case"])
        return {"version": 1, "environment": env_json, "expect": expect}
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
      - ``run_workbook_case`` -- builds the declarative workbook spec
        carried by ``case`` and returns the observed pivot/print
        ``expect`` mapping.

    Wire-format version is pinned to 1; any future schema change must
    bump it both here and in the WSL bridge.
    """

    import argparse
    import json

    parser = argparse.ArgumentParser(description="windows_excel wire driver")
    parser.add_argument("--input", type=Path, help="single-shot mode: read one JSON command from this path")
    parser.add_argument("--output", type=Path, help="single-shot mode: write one JSON response to this path")
    parser.add_argument("--serve", action="store_true", help="long-lived mode: serve newline-delimited JSON over stdio")
    parser.add_argument("--visible", action="store_true", help="show the Excel window (debug aid for --serve)")
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
