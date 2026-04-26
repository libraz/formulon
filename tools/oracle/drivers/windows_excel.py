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

## Wire-protocol entrypoint

Running ``python -m tools.oracle.drivers.windows_excel --input X.json
--output Y.json`` reads a command (``probe_environment`` or ``run_suite``)
from the input JSON and writes the result vector to the output JSON. This
is the bridge that :class:`tools.oracle.drivers.wsl_bridge.WSLBridgeOracle`
calls into when the orchestration host is WSL2.
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
    Python-side ``.value`` has been coerced to ``None``). We try a small
    list of names because the bridge attribute spelling varies across
    Excel releases; ``Text`` is the canonical one on 16.x COM, with the
    other names kept for forward / backward compatibility with the
    underlying xlwings build.
    """

    try:
        api = cell.api
    except Exception:
        return None
    for attr in ("Text", "DisplayValue", "string_value", "text"):
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

    if isinstance(raw, str) and raw in _ERR_DISPLAY_NAMES:
        return raw

    text = _cell_displayed_text(cell)
    if text in _ERR_DISPLAY_NAMES:
        return text
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


class WindowsExcelOracle(OracleDriver):
    """Thin wrapper over a single hidden Excel.exe instance (COM).

    The instance is reused across suites so we pay COM startup latency
    only once. Call :meth:`close` (or use as a context manager) to tear
    down. Refuses to construct on non-Windows hosts.
    """

    def __init__(self, visible: bool = False) -> None:
        _ensure_windows()
        self._app = xw.App(visible=visible, add_book=False)
        self._app.calculation = "manual"
        self._app.screen_updating = False
        self._app.display_alerts = False
        try:
            self._app.api.EnableIterativeCalculation = False
        except Exception:
            # Older COM proxies expose the same property under the Mac
            # snake_case spelling; tolerate either rather than failing
            # the whole batch on a property name mismatch.
            try:
                self._app.api.enable_iterative_calculation = False
            except Exception:
                pass

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
        ``Build`` properties. The locale is hard-coded to ``ja-JP`` for
        parity with the Mac driver -- we expect the operator to run a
        ja-JP-localised Office install when targeting ``win-365-ja_JP``.
        Future work: cross-check ``Application.International`` against
        the targets.yaml ``locale`` field and warn on mismatch.
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

        Each ``cases[i]`` is expected to be a dict with keys ``id``,
        ``formula``, and ``setup`` (mapping A1 -> ``{kind, value, ...}``
        record). The driver does not normalise; it trusts
        ``case_schema._normalise_value`` upstream.
        """

        wb = self._app.books.add()
        try:
            try:
                wb.api.Date1904 = date1904
            except Exception:
                try:
                    wb.api.date1904 = date1904
                except Exception:
                    pass
            prior_iter = None
            try:
                prior_iter = bool(self._app.api.EnableIterativeCalculation)
                self._app.api.EnableIterativeCalculation = iterative
            except Exception:
                try:
                    prior_iter = bool(self._app.api.enable_iterative_calculation)
                    self._app.api.enable_iterative_calculation = iterative
                except Exception:
                    prior_iter = None

            first_sheet = wb.sheets[0]
            case_sheets: List[object] = []
            for i, case in enumerate(cases):
                if i == 0:
                    sht = first_sheet
                else:
                    sht = wb.sheets.add(after=case_sheets[-1])
                safe_id = _sanitize_sheet_name(case["id"])
                sht.name = f"c{i + 1:03d}_{safe_id}"[:31]
                case_sheets.append(sht)

                setup = case.get("setup") or {}
                for addr, rec in setup.items():
                    _write_cell(sht, addr, rec)
                result_cell = sht.range("Z1")
                try:
                    result_cell.number_format = "General"
                except Exception:
                    pass
                result_cell.formula2 = case["formula"]

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
                    try:
                        self._app.api.EnableIterativeCalculation = prior_iter
                    except Exception:
                        self._app.api.enable_iterative_calculation = prior_iter
            except Exception:
                pass
            try:
                wb.close()
            except Exception:
                pass


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
        rng.formula2 = rec["formula"]
        return
    if kind == "error":
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
    "#REF!": "=OFFSET(A1,-1,-1)",
    "#NULL!": "=A1 B1",
}


def _error_trigger(code: str) -> str:
    return _ERROR_TRIGGERS.get(code, '=VALUE("x")')


def _main(argv: Optional[List[str]] = None) -> int:
    """Wire-protocol entrypoint for the WSL bridge.

    Reads a JSON command from ``--input``, writes results to
    ``--output``. This lets the WSL-side orchestrator invoke us via
    subprocess without sharing Python state.

    Supported commands:
      - ``probe_environment`` -- returns the environment block only.
      - ``run_suite`` -- evaluates the supplied case batch and returns
        a vector of normalised :class:`CaseResult` records.

    Wire-format version is pinned to 1; any future schema change must
    bump it both here and in the WSL bridge.
    """

    import argparse
    import json

    parser = argparse.ArgumentParser(description="windows_excel wire driver")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args(argv)

    payload = json.loads(args.input.read_text(encoding="utf-8"))
    if payload.get("version") != 1:
        raise RuntimeError(f"unsupported wire version: {payload.get('version')}")

    cmd = payload.get("command")
    visible = bool(payload.get("visible", False))

    with WindowsExcelOracle(visible=visible) as drv:
        env = drv.probe_environment()
        env_json = {
            "excel_version": env.excel_version,
            "excel_locale": env.excel_locale,
            "date1904": env.date1904,
            "iterative": env.iterative,
        }
        if cmd == "probe_environment":
            out = {"version": 1, "environment": env_json, "results": []}
        elif cmd == "run_suite":
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
            out = {"version": 1, "environment": env_json, "results": results_json}
        else:
            raise RuntimeError(f"unknown command: {cmd!r}")

    args.output.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
    return 0


if __name__ == "__main__":
    import sys

    sys.exit(_main())
