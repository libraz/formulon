#!/usr/bin/env python3
"""Mac Excel 365 oracle driver.

Drives Excel.app through xlwings to evaluate Formulon oracle cases under
controlled options (manual calc / iterative off / per-workbook date system). The driver is
intentionally small — the caller (`oracle_gen.py`) handles case loading,
divergence filtering, and golden JSON emission.

macOS only. `xlwings.App()` is a thin wrapper around AppleEvents, so Excel
must be installed and the Terminal / iTerm / IDE that hosts Python must
have Automation permission for "System Events" + "Microsoft Excel".

## Options we pin

- `calculation = 'manual'` — we call `app.calculate()` after every batch
  ourselves, so formulas never race partial input.
- `screen_updating = False` — large batches are 10-20x faster without it.
- `display_alerts = False` — suppresses dialogs that would otherwise block
  the AppleEvent thread.
- `date_1904` is set and read back per workbook through the Mac
  AppleScript reference immediately after `books.add()`.
- `iteration = False` on the app. Iterative cases flip this locally and
  restore it.

## Batch layout

Each suite is written to one workbook with one worksheet per case. Setup
cells are written at their absolute A1 address (as given in the YAML), and
the formula under test is written to Z1 on that case's worksheet — or to
the case's `formula_cell` address when it declares one, which is how a
case probes placement-dependent behaviour such as whether a whole-column
spill still fits below the formula row.

After writing, a single `app.calculate()` resolves the whole sheet; results
are then read in a second pass and packaged into `CaseResult` records.
"""

from __future__ import annotations

import datetime as _dt
import platform
from typing import Any, Dict, List, Optional

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
except ImportError as exc:  # pragma: no cover - exercised on minimal CI hosts
    xw = None  # type: ignore[assignment]
    _XLWINGS_IMPORT_ERROR: Optional[ImportError] = exc
else:
    _XLWINGS_IMPORT_ERROR = None


def _ensure_darwin() -> None:
    if platform.system() != "Darwin":
        raise RuntimeError(
            "oracle-gen is macOS-only (xlwings drives Excel.app). Current platform: " + platform.system()
        )


def _get_iteration(app) -> bool:
    """Read Mac Excel's application-level iterative-calculation setting."""

    return bool(app.api.iteration.get())


def _set_iteration(app, enabled: bool) -> None:
    """Set Mac Excel's application-level iterative-calculation setting."""

    app.api.iteration.set(enabled)


def _set_date1904(wb, enabled: bool) -> None:
    """Set and verify a workbook's Mac Excel 1904 date-system flag.

    On macOS, ``wb.api.date_1904`` is an AppleScript reference rather than a
    normal Python attribute.  Assigning to it directly can leave the
    workbook's date system unchanged without raising, so use the reference's
    explicit ``set`` / ``get`` operations and reject a mismatched readback.
    Exceptions from either AppleScript operation deliberately propagate.
    """

    expected = bool(enabled)
    date_1904 = wb.api.date_1904
    date_1904.set(expected)
    actual = date_1904.get()
    if actual != expected:
        raise RuntimeError(f"Excel rejected workbook date_1904={expected!r}: read back {actual!r}")


class _FormulaRetentionError(RuntimeError):
    """Raised when Excel does not retain an assigned formula as written."""


def _formula_readback(cell) -> Optional[str]:
    """Returns a non-empty formula readback from ``cell``, if available.

    ``formula2`` is the preferred Mac Excel property, but older xlwings /
    AppleScript combinations may expose only ``formula``.  A formula readback
    is intentionally not compared with the input: Excel is allowed to
    canonicalise formulas while retaining them.
    """

    for attr in ("formula2", "formula"):
        try:
            value = getattr(cell, attr)
        except Exception:
            continue
        try:
            value = value() if callable(value) else value
        except Exception:
            continue
        if isinstance(value, str) and value.strip():
            return value
    return None


def _assign_formula(cell, formula: Any, *, context: str) -> None:
    """Assigns a formula and rejects Excel's silent non-retention failure modes.

    Assignment exceptions deliberately propagate unchanged.  Two distinct Mac
    failure modes are caught by reading the formula back:

    * Excel accepts the AppleEvent but leaves the cell empty.
    * The AppleEvent setter keeps only a prefix of a long formula.  The cut
      lands wherever the byte budget runs out, so a chain such as
      ``=A1+A2+...+A150`` comes back as ``=A1+...+A66`` -- still a valid
      formula, still evaluating, and therefore indistinguishable from a real
      Excel answer to anything downstream.  Capturing that as a golden would
      pin an observation Excel never made.

    A readback that merely differs from the input is fine: Excel is allowed to
    canonicalise a formula it retains in full.  Only a readback that is a
    strict prefix of the input is treated as truncation, because
    canonicalisation rewrites tokens rather than dropping the tail.

    The assignment goes through ``formula2`` and not ``formula_local``.
    That is not a preference: the localised route is **inert on this
    bridge**.  Assigning to ``formula_local`` leaves the cell blank and
    raises nothing, for every formula including a trivial control -- a
    third silent-non-retention mode, and the one this function cannot
    catch because it never gets a chance to read anything back.  It
    matters because ``formula_local`` is the route that carries localised
    function names, so it is the natural place to look when a function
    appears to be missing.  A blank read from it is a fact about the
    bridge and not an observation about Excel; anything investigating an
    apparently-unavailable function has to use ``formula2`` plus an
    in-Excel probe such as ``ERROR.TYPE(f(...))``.  Verified on Excel
    16.112 while establishing that FILTERXML is genuinely absent on Mac.
    """

    cell.formula2 = formula
    if not (isinstance(formula, str) and formula and formula.startswith("=")):
        return
    readback = _formula_readback(cell)
    if readback is None:
        raise _FormulaRetentionError(
            f"Excel silently rejected formula for {context}: formula={formula!r}; formula2/formula readback was empty"
        )
    if len(readback) < len(formula) and formula.startswith(readback):
        raise _FormulaRetentionError(
            f"Excel retained only a prefix of the formula for {context}: "
            f"sent {len(formula)} characters, read back {len(readback)}. "
            f"formula={formula!r}; readback={readback!r}. "
            "Assign a formula this long through a saved workbook instead of the AppleEvent setter."
        )


def _cell_displayed_text(cell) -> Optional[str]:
    """Returns the rendered string of a cell via the AppleScript bridge.

    On Mac, `cell.api.string_value` is the property that yields the
    displayed string (including `'#DIV/0!'` for errors whose Python-side
    `.value` has been coerced to `None`). We try a small list of names
    because the bridge attribute spelling varies across Excel releases;
    `string_value` is the canonical one on 16.x.
    """

    try:
        api = cell.api
    except Exception:
        return None
    for attr in ("string_value", "text", "displayed_value", "formatted_text"):
        try:
            obj = getattr(api, attr)
        except Exception:
            continue
        try:
            val = obj() if callable(obj) else obj
        except Exception:
            continue
        if isinstance(val, str) and val:
            return val
    return None


def _error_display_from_cell(cell) -> Optional[str]:
    """Returns the tokenised Excel error name for `cell`, or None.

    Walks four progressively weaker signals:
      1. `xlwings.utils.CVErr` — ideal, but only surfaces on some Excel /
         xlwings build pairs.
      2. `cell.value` already a '#DIV/0!'-style string.
      3. The displayed text (AppleScript `.text`) matches a known error.
      4. `cell.value` is `None` AND the displayed text nonetheless starts
         with `#`. This is the fallback Mac path where the Python layer
         has coerced the error into `None`.
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
        canon = normalise_error_token(raw)
        if canon is not None:
            return canon

    text = _cell_displayed_text(cell)
    if text in _ERR_DISPLAY_NAMES:
        return text
    # AppleScript `string_value` is locale-bound: de-DE returns "#WERT!",
    # fr-FR "#VALEUR!", and so on. Normalise through the shared
    # localisation map before falling back to the prefix heuristic.
    if text:
        canon = normalise_error_token(text)
        if canon is not None:
            return canon
    # Last-ditch: match by prefix (`#DIV/0!...` in an ex-format-localised
    # build, for example). All Excel errors start with `#` and end with
    # `!` or `?`; we don't want to catch text that happens to start with
    # '#'.
    if text and text.startswith("#") and (text.endswith("!") or text.endswith("?") or text == "#N/A"):
        for name in _ERR_DISPLAY_NAMES:
            if text == name:
                return name
    return None


def _classify_value(cell) -> CaseResult:
    """Converts an xlwings cell observation into a CaseResult.

    The `cell.value` read happens once up front; subsequent checks may
    consult the AppleScript `.text` fallback for Mac edge cases (error
    cells, pre-1900 serials) where the Python-side value is lossy.
    """

    err = _error_display_from_cell(cell)
    if err is not None:
        return CaseResult(id="", kind="error", error_code=err)

    v = cell.value
    # Datetime must come before bool/int because `datetime` is distinct
    # from `int` in Python, but xlwings may return a `date` (no time) for
    # pure DATE() results. Both use the shared 1900-system conversion so the
    # golden carries a number kind (matching Formulon's Value::Number). The
    # classifier has no workbook epoch metadata; 1904 cases should therefore
    # expose date values through a numeric/boolean formula such as N(DATE()).
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
        # Mac Excel can return None for pre-1900 serials (Python can't
        # represent them as datetime). Inspect the displayed text and try
        # to parse it as an integer serial before falling back to blank.
        text = _cell_displayed_text(cell)
        if text and text.strip():
            # Trivial numeric fallback: some ja-JP date formats render
            # as "1900/2/29" even though cell.value is None. Bail to
            # blank here; the YAML case should be rewritten to force a
            # non-date display if a serial is needed.
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
    """Mac adapter for the shared Application.Evaluate shape probe."""

    return probe_spill_shape(lambda expression: app.api.evaluate(name=expression), anchor, max_cells=max_cells)


def _classify_shape_result(app, sht, anchor_addr: str, samples: List[str]) -> CaseResult:
    """Records a spill as its shape plus the case's declared sample cells.

    For results too large to materialise cell by cell. The shape comes from
    the same non-invasive probe the cell walk uses; only the listed cells
    are read.
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


class ExcelOracle(OracleDriver):
    """Thin wrapper over a single hidden Excel.app instance.

    The instance is reused across suites so we pay AppleEvents startup
    latency only once. Call :meth:`close` (or use as a context manager) to
    tear down.
    """

    def __init__(self, visible: bool = False) -> None:
        _ensure_darwin()
        if xw is None:
            raise RuntimeError("xlwings is not installed; run `make oracle-setup` first") from _XLWINGS_IMPORT_ERROR
        self._app = xw.App(visible=visible, add_book=False)
        self._app.calculation = "manual"
        self._app.screen_updating = False
        self._app.display_alerts = False
        # The iterative-calc setting is an app-level flag on Mac Excel.
        # Restore to False at construction time; per-suite overrides flip
        # and restore as needed.
        try:
            _set_iteration(self._app, False)
        except Exception:
            pass

    def __enter__(self) -> "ExcelOracle":
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

        Prefer xlwings's top-level properties (`app.version`) over the raw
        `.api` applescript proxies — on Mac the proxies return reference
        objects whose `__repr__` leaks the applescript call expression.
        """

        version = ""
        # `App.version` is a Mac-safe string property on recent xlwings.
        try:
            version = str(self._app.version) if self._app.version else ""
        except Exception:
            version = ""
        # `App.api.build` is a property call on Mac appscript; invoking it
        # as a no-arg callable coerces the reference to its value.
        if not version:
            try:
                api = self._app.api
                v = api.version() if callable(api.version) else api.version
                b = api.build() if callable(api.build) else api.build
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

        Each `cases[i]` is expected to be a dict with keys `id`, `formula`,
        and `setup` (mapping A1 -> `{kind, value, ...}` record). The
        driver does not normalise; it trusts `case_schema._normalise_value`
        upstream.
        """

        # Cross-sheet setup ("Sheet2!A1") must be isolated per case --
        # see the windows_excel.py counterpart for the rationale.
        cross_sheet = any(any("!" in addr for addr in (case.get("setup") or {})) for case in cases)
        if cross_sheet:
            return self._run_suite_per_case_workbook(suite_name, cases, date1904=date1904, iterative=iterative)

        wb = self._app.books.add()
        try:
            # Pin per-workbook options before any formula touches volatile
            # state (NOW/TODAY, date arithmetic).
            _set_date1904(wb, date1904)
            prior_iter = None
            try:
                prior_iter = _get_iteration(self._app)
                _set_iteration(self._app, iterative)
            except Exception:
                prior_iter = None

            # One worksheet per case. Previously the whole suite shared a
            # single sheet and later cases' setup writes would overwrite
            # earlier cases' cells — e.g. case N writing A2="x" clobbered
            # case M's A2=true, silently changing A1:A3's content by the
            # time `calculate()` ran. Per-sheet isolation costs a few
            # extra AppleEvents per case but makes the batch correct.
            first_sheet = wb.sheets[0]
            case_sheets: List[object] = []
            write_errors: Dict[str, str] = {}
            for i, case in enumerate(cases):
                if i == 0:
                    sht = first_sheet
                else:
                    sht = wb.sheets.add(after=case_sheets[-1])
                # Sheet names are capped at 31 chars and forbid a handful
                # of punctuation. Prefix with the ordinal so duplicates
                # across cases (same case id in two suites) are safe, and
                # clip the id to the remaining budget.
                safe_id = _sanitize_sheet_name(case["id"])
                sht.name = f"c{i + 1:03d}_{safe_id}"[:31]
                case_sheets.append(sht)

                try:
                    _apply_merges(sht, case.get("merges") or [])
                    setup = case.get("setup") or {}
                    for addr, rec in setup.items():
                        _write_cell(
                            sht,
                            addr,
                            rec,
                            context=f"case {case['id']!r} setup {addr!r}",
                        )
                    result_cell = sht.range(case_formula_cell(case))
                    # Pin the result cell to General format. Otherwise Excel
                    # auto-formats DATE()/TIME() results as m/d/yyyy and
                    # xlwings hands us a Python datetime (or None for the
                    # 1900-02-29 ghost day) instead of the raw serial we
                    # want to capture.
                    try:
                        result_cell.number_format = "General"
                    except Exception:
                        pass
                    _assign_formula(
                        result_cell,
                        case["formula"],
                        context=f"case {case['id']!r} result",
                    )
                except Exception as exc:
                    write_errors[case["id"]] = _format_mac_error(exc)

            # Single recalc across all sheets once every formula is written.
            self._app.calculate()

            out: List[CaseResult] = []
            for case, sht in zip(cases, case_sheets):
                if case["id"] in write_errors:
                    out.append(CaseResult(id=case["id"], kind="skipped", value=write_errors[case["id"]]))
                    continue
                try:
                    result = _classify_case_result(self._app, sht, case)
                    result.id = case["id"]
                    out.append(result)
                except Exception as exc:
                    out.append(CaseResult(id=case["id"], kind="skipped", value=_format_mac_error(exc)))
            return out
        finally:
            try:
                if prior_iter is not None:
                    _set_iteration(self._app, prior_iter)
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
        """Per-case-workbook runner for cross-sheet-setup suites."""

        prior_iter = None
        try:
            prior_iter = _get_iteration(self._app)
            _set_iteration(self._app, iterative)
        except Exception:
            prior_iter = None
        out: List[CaseResult] = []
        try:
            for case in cases:
                wb = self._app.books.add()
                try:
                    # Pin the workbook date system before writing any setup or
                    # formula. _set_date1904 intentionally propagates
                    # AppleScript failures instead of marking them skipped.
                    _set_date1904(wb, date1904)
                    try:
                        sht = wb.sheets[0]
                        setup = case.get("setup") or {}
                        _apply_merges(sht, case.get("merges") or [])
                        for addr, rec in setup.items():
                            sheet_name, bare_addr = _split_sheet_qualified_addr(addr)
                            target_sht = sht if sheet_name is None else _get_or_add_sheet(wb, sheet_name)
                            _write_cell(
                                target_sht,
                                bare_addr,
                                rec,
                                context=f"case {case['id']!r} setup {addr!r}",
                            )
                        result_cell = sht.range(case_formula_cell(case))
                        try:
                            result_cell.number_format = "General"
                        except Exception:
                            pass
                        _assign_formula(
                            result_cell,
                            case["formula"],
                            context=f"case {case['id']!r} result",
                        )
                    except Exception as exc:
                        out.append(
                            CaseResult(
                                id=case["id"],
                                kind="skipped",
                                value=_format_mac_error(exc),
                            )
                        )
                        continue

                    try:
                        self._app.calculate()
                        result = _classify_case_result(self._app, sht, case)
                        result.id = case["id"]
                        out.append(result)
                    except Exception as exc:
                        out.append(
                            CaseResult(
                                id=case["id"],
                                kind="skipped",
                                value=_format_mac_error(exc),
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
                    _set_iteration(self._app, prior_iter)
            except Exception:
                pass

    # -----------------------------------------------------------------------
    # Workbook oracle track (pivot tables) -- best-effort variant
    # -----------------------------------------------------------------------
    #
    # The workbook oracle's primary is Windows Excel: reliable PivotTable
    # automation needs the Windows COM object model. macOS Excel exposes
    # pivot objects through the xlwings ``.api`` AppleScript bridge, but
    # several pivot operations (RowAxisLayout, per-item visibility) are
    # either missing or unstable there. This implementation drives the
    # subset that works and raises a clear, *catchable* RuntimeError for
    # anything it cannot reproduce so the generator marks the case
    # skipped rather than emitting a wrong golden. Mac is a variant, not
    # a gate, so a partial implementation is acceptable here.

    def run_workbook_case(self, case: Dict[str, Any]) -> Dict[str, Any]:
        """Best-effort workbook-feature build via the macOS object model.

        Returns the golden ``expect`` block on success; raises a
        ``RuntimeError`` (catchable) when the feature cannot be driven on
        Mac so the case is marked skipped. Mac is a workbook-track
        variant, not a gate, so a partial implementation is acceptable.
        """

        print_spec = case.get("print")
        if isinstance(print_spec, dict):
            return self._run_print_case(case, print_spec)

        pivot_spec = case.get("pivot")
        if not isinstance(pivot_spec, dict):
            # Both feature blocks are optional in the case schema (see
            # tests/oracle/cases_wb/README.md); a no-feature case --
            # typically a schema smoke -- yields an empty expect block.
            return {}
        # The macOS object model does not expose a stable PivotCaches API
        # comparable to Windows COM. Rather than emit a divergent golden,
        # surface a clear skip so the operator runs the Windows primary
        # for pivot goldens.
        raise RuntimeError(
            "macOS Excel cannot reliably drive PivotTable automation; "
            "generate pivot goldens on the Windows primary target "
            f"(case {case.get('id')!r})"
        )

    def _run_print_case(self, case: Dict[str, Any], print_spec: Dict[str, Any]) -> Dict[str, Any]:
        """Best-effort print build via the macOS xlwings `.api` bridge.

        Mac Excel does expose `PageSetup` and `HPageBreaks` /
        `VPageBreaks` through the AppleScript object model, so the print
        path is driven directly. Any operation the bridge rejects raises
        a clear, catchable ``RuntimeError`` so the case is marked
        skipped rather than emitting a wrong golden.
        """

        wb = self._app.books.add()
        try:
            self._build_workbook_sheets(wb, case)
            return {"print": _apply_and_read_print(wb, print_spec)}
        except RuntimeError:
            raise
        except Exception as exc:
            raise RuntimeError(
                f"print automation failed for case {case.get('id')!r}: {_format_mac_error(exc)}"
            ) from exc
        finally:
            try:
                wb.close()
            except Exception:
                pass

    def _build_workbook_sheets(self, wb, case: Dict[str, Any]) -> None:
        """Materialises the declarative ``sheets`` block plus the optional
        ``column_widths`` / ``row_heights`` layout maps into ``wb``.
        """

        sheets = case.get("sheets") or {}
        for sheet_name, cells in sheets.items():
            sht = _get_or_add_sheet(wb, sheet_name)
            for addr, rec in (cells or {}).items():
                _write_cell(
                    sht,
                    addr,
                    rec,
                    context=f"workbook case {case.get('id')!r} {sheet_name}!{addr!s}",
                )

        widths = case.get("column_widths") or {}
        for col_key, width in widths.items():
            try:
                wb.sheets[0].range(f"{col_key}1").column_width = float(width)
            except Exception:
                pass
        heights = case.get("row_heights") or {}
        for row_key, height in heights.items():
            try:
                wb.sheets[0].range(f"A{row_key}").row_height = float(height)
            except Exception:
                pass


def _split_sheet_qualified_addr(key: str) -> "tuple[Optional[str], str]":
    """Splits ``"Sheet2!A1"`` into ``("Sheet2", "A1")``; returns ``(None, key)``
    for bare A1 keys. Mirrors the Windows driver helper -- see
    ``windows_excel._split_sheet_qualified_addr`` for the quoting rules.
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
    adding it at the end if absent.
    """

    target = name.casefold()
    for sht in wb.sheets:
        if sht.name.casefold() == target:
            return sht
    return wb.sheets.add(name=name, after=wb.sheets[len(wb.sheets) - 1])


# Excel `XlPageOrientation` constants.
_XL_PORTRAIT = 1
_XL_LANDSCAPE = 2

_PRINT_ORIENTATIONS = {
    "portrait": _XL_PORTRAIT,
    "landscape": _XL_LANDSCAPE,
}


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


def _normalise_print_area(area: str) -> str:
    """Strips ``$`` anchors and ``Sheet!`` qualifiers from a print area.

    Mirrors the Windows driver helper -- the C++ engine compares against
    a bare ``A1:H80`` form, so both drivers normalise identically.
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


def _apply_and_read_print(wb, print_spec: Dict[str, Any]) -> Dict[str, Any]:
    """Applies the declarative `print` block to ``wb`` and reads it back.

    Returns the golden ``expect.print`` mapping. Mac Excel's object
    model exposes ``PageSetup`` and the page-break collections through
    the xlwings AppleScript bridge; the 1-based COM-style row / column
    indices are converted to the 0-based indices the C++ engine reports.
    Any unsupported operation surfaces as a catchable ``RuntimeError``.
    """

    sht = _resolve_print_sheet(wb, print_spec)
    ws = sht.api
    page_setup = ws.PageSetup

    print_area = print_spec.get("print_area")
    if isinstance(print_area, str) and print_area:
        page_setup.PrintArea = print_area

    titles = print_spec.get("print_titles")
    if isinstance(titles, dict):
        rows = titles.get("rows")
        cols = titles.get("cols")
        if isinstance(rows, str) and rows:
            page_setup.PrintTitleRows = f"{sht.name}!{rows}"
        if isinstance(cols, str) and cols:
            page_setup.PrintTitleColumns = f"{sht.name}!{cols}"

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
        fit_w = int(setup.get("fit_to_width") or 0)
        fit_h = int(setup.get("fit_to_height") or 0)
        if fit_w or fit_h:
            page_setup.Zoom = False
            page_setup.FitToPagesWide = fit_w if fit_w else False
            page_setup.FitToPagesTall = fit_h if fit_h else False
        elif setup.get("scale") is not None:
            page_setup.Zoom = int(setup["scale"])

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

    resolved_area = ""
    try:
        resolved_area = _normalise_print_area(str(page_setup.PrintArea or ""))
    except Exception:
        resolved_area = ""

    h_breaks: List[int] = []
    v_breaks: List[int] = []
    try:
        for i in range(1, int(ws.HPageBreaks.Count) + 1):
            h_breaks.append(int(ws.HPageBreaks(i).Location.Row) - 1)
        for i in range(1, int(ws.VPageBreaks.Count) + 1):
            v_breaks.append(int(ws.VPageBreaks(i).Location.Column) - 1)
    except Exception as exc:
        raise RuntimeError(f"macOS Excel could not read page-break collections: {_format_mac_error(exc)}") from exc
    h_breaks.sort()
    v_breaks.sort()

    pages = 0
    try:
        pages = int(page_setup.Pages.Count)
    except Exception:
        pages = 0

    return {
        "print_area": resolved_area,
        "h_breaks": h_breaks,
        "v_breaks": v_breaks,
        "pages": pages,
    }


def _write_cell(sht, addr: str, rec: Dict[str, Any], *, context: str = "setup cell") -> None:
    """Writes one normalised {kind, value} record to `sht!addr`."""

    kind = rec.get("kind")
    rng = sht.range(addr)
    if kind == "blank":
        rng.clear_contents()
        return
    if kind == "number":
        rng.value = float(rec["value"])
        return
    if kind == "bool":
        # Writing a Python bool lands as TRUE/FALSE in Excel.
        rng.value = bool(rec["value"])
        return
    if kind == "text":
        # Prepend an apostrophe is tempting but changes the cell's stored
        # value; leave it to Excel's default string handling.
        rng.value = str(rec["value"])
        return
    if kind == "formula":
        _assign_formula(rng, rec["formula"], context=context)
        return
    if kind == "error":
        # Seed errors via a formula since Excel has no API for "set this
        # cell to #DIV/0! without a formula". Use the canonical trigger
        # for each error code.
        trigger = _error_trigger(rec.get("code", "#VALUE!"))
        _assign_formula(rng, trigger, context=f"{context} error trigger")
        return
    raise ValueError(f"unknown cell kind: {kind}")


def _apply_merges(sht, merges: List[str]) -> None:
    """Apply case-declared inclusive A1 merge ranges on ``sht``."""

    for ref in merges:
        sht.range(ref).merge()


def _format_mac_error(exc: BaseException) -> str:
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
    "#REF!": "=OFFSET(A1,-1,-1)",  # relative moves out of range
    "#NULL!": "=A1 B1",  # intersection with no overlap
}


def _error_trigger(code: str) -> str:
    return _ERROR_TRIGGERS.get(code, '=VALUE("x")')
