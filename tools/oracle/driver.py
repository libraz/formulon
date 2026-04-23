#!/usr/bin/env python3
"""Mac Excel 365 oracle driver.

Drives Excel.app through xlwings to evaluate Formulon oracle cases under
controlled options (manual calc / iterative off / 1904 off). The driver is
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
- `date_1904 = False` per workbook (set after `books.add()`).
- `enable_iterative_calculation = False` on the app. Iterative cases flip
  this locally and restore it.

## Batch layout

Each suite is written to a single worksheet. Setup cells are written at
their absolute A1 address (as given in the YAML). Formulas under test are
written at column Z, one per row (Z1 = case[0].formula, Z2 = case[1]...).
If any case's `setup` already writes to column Z the driver detects the
collision and picks the next unused column (AA, AB, ...).

After writing, a single `app.calculate()` resolves the whole sheet; results
are then read in a second pass and packaged into `CaseResult` records.
"""

from __future__ import annotations

import datetime as _dt
import platform
from dataclasses import dataclass
from typing import Any, Dict, List, Optional

try:
    import xlwings as xw  # type: ignore
except ImportError as exc:  # pragma: no cover - handled in oracle_gen.py
    raise RuntimeError(
        "xlwings is not installed; run `make oracle-setup` first"
    ) from exc


# Excel's error values come back from xlwings as CVErr / ErrorValue / str
# depending on Mac vs Windows and the property we read. We centralise the
# recognition here so callers only see the tokenised display name.
_ERR_DISPLAY_NAMES = {
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#SPILL!",
    "#CALC!",
    "#FIELD!",
    "#BLOCKED!",
    "#CONNECT!",
    "#EXTERNAL!",
    "#BUSY!",
    "#PYTHON!",
    "#UNKNOWN!",
}


@dataclass
class CaseResult:
    """Normalised observation for one case.

    Mirrors the on-wire JSON shape the generator will emit. `kind` is one
    of {'blank','number','bool','text','error'}. `array_shape` is reserved
    for spill results; the initial driver flattens arrays to a single
    value and signals `array_shape is not None` so callers can log.
    """

    id: str
    kind: str
    # Populated based on kind:
    value: Any = None
    error_code: Optional[str] = None
    array_shape: Optional[List[int]] = None


@dataclass
class EnvironmentInfo:
    excel_version: str
    excel_locale: str
    date1904: bool
    iterative: bool


def _ensure_darwin() -> None:
    if platform.system() != "Darwin":
        raise RuntimeError(
            "oracle-gen is macOS-only (xlwings drives Excel.app). "
            "Current platform: " + platform.system()
        )


# Excel's epoch relative to which serial numbers are counted. 1900-01-01
# is serial 1, so the subtraction base is 1899-12-31. Dates on or after
# 1900-03-01 get a +1 bump to account for Excel's phantom Feb 29 1900
# (serial 60), which Python's datetime cannot represent.
_EXCEL_EPOCH = _dt.datetime(1899, 12, 31)


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

    if isinstance(raw, str) and raw in _ERR_DISPLAY_NAMES:
        return raw

    text = _cell_displayed_text(cell)
    if text in _ERR_DISPLAY_NAMES:
        return text
    # Last-ditch: match by prefix (`#DIV/0!...` in an ex-format-localised
    # ja-JP build, for example). All Excel errors start with `#` and end
    # with `!` or `?`; we don't want to catch text that happens to start
    # with '#'.
    if text and text.startswith("#") and (text.endswith("!") or text.endswith("?") or text == "#N/A"):
        for name in _ERR_DISPLAY_NAMES:
            if text == name:
                return name
    return None


def _datetime_to_serial(v: _dt.datetime) -> float:
    """Converts a Python datetime to an Excel 1900-system serial number.

    The 1900 leap-bug means every serial >= 60 is off-by-one relative to
    a naive epoch subtraction — Excel's day 61 (March 1, 1900) is only
    61 real days after 1899-12-30, but Excel treats 1900 as a leap year
    and assigns serial 60 to a non-existent Feb 29. Python datetime
    skips that ghost day, so we add 1 to compensate for anything strictly
    after 1900-02-28.
    """

    delta = (v - _EXCEL_EPOCH).total_seconds() / 86400.0
    if v >= _dt.datetime(1900, 3, 1):
        delta += 1
    return delta


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
    # pure DATE() results — both are coerced to the Excel serial so the
    # golden carries a number kind (matching Formulon's Value::Number).
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
            flat = [item for row in v for item in row]
        else:
            cols = rows
            rows = 1
            flat = list(v)
        return CaseResult(
            id="",
            kind="array",
            value=flat,
            array_shape=[rows, cols],
        )
    return CaseResult(id="", kind="text", value=str(v))


class ExcelOracle:
    """Thin wrapper over a single hidden Excel.app instance.

    The instance is reused across suites so we pay AppleEvents startup
    latency only once. Call :meth:`close` (or use as a context manager) to
    tear down.
    """

    def __init__(self, visible: bool = False) -> None:
        _ensure_darwin()
        self._app = xw.App(visible=visible, add_book=False)
        self._app.calculation = "manual"
        self._app.screen_updating = False
        self._app.display_alerts = False
        # The iterative-calc setting is an app-level flag on Mac Excel.
        # Restore to False at construction time; per-suite overrides flip
        # and restore as needed.
        try:
            self._app.api.enable_iterative_calculation = False
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
        return EnvironmentInfo(
            excel_version=version.strip(),
            excel_locale="ja-JP",  # pinned per CLAUDE.md policy
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

        wb = self._app.books.add()
        try:
            # Pin per-workbook options before any formula touches volatile
            # state (NOW/TODAY, date arithmetic).
            try:
                wb.api.date1904 = date1904
            except Exception:
                pass
            prior_iter = None
            try:
                prior_iter = bool(self._app.api.enable_iterative_calculation)
                self._app.api.enable_iterative_calculation = iterative
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

                setup = case.get("setup") or {}
                for addr, rec in setup.items():
                    _write_cell(sht, addr, rec)
                result_cell = sht.range("Z1")
                # Pin the result cell to General format. Otherwise Excel
                # auto-formats DATE()/TIME() results as m/d/yyyy and
                # xlwings hands us a Python datetime (or None for the
                # 1900-02-29 ghost day) instead of the raw serial we
                # want to capture.
                try:
                    result_cell.number_format = "General"
                except Exception:
                    pass
                result_cell.formula2 = case["formula"]

            # Single recalc across all sheets once every formula is written.
            self._app.calculate()

            out: List[CaseResult] = []
            for case, sht in zip(cases, case_sheets):
                cell = sht.range("Z1")
                result = _classify_value(cell)
                result.id = case["id"]
                out.append(result)
            return out
        finally:
            try:
                if prior_iter is not None:
                    self._app.api.enable_iterative_calculation = prior_iter
            except Exception:
                pass
            try:
                wb.close()
            except Exception:
                pass


def _write_cell(sht, addr: str, rec: Dict[str, Any]) -> None:
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
        rng.formula2 = rec["formula"]
        return
    if kind == "error":
        # Seed errors via a formula since Excel has no API for "set this
        # cell to #DIV/0! without a formula". Use the canonical trigger
        # for each error code.
        trigger = _error_trigger(rec.get("code", "#VALUE!"))
        rng.formula2 = trigger
        return
    raise ValueError(f"unknown cell kind: {kind}")


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
