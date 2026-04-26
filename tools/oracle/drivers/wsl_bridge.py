#!/usr/bin/env python3
"""WSL2 -> Windows Excel bridge driver.

Implements :class:`OracleDriver` by invoking ``python.exe -m
tools.oracle.drivers.windows_excel`` as a Windows-side subprocess. JSON
files in ``/tmp`` (translated to Windows paths via ``wslpath -w``) carry
the case batch and result vectors across the WSL/Win boundary.

Used automatically when :func:`tools.oracle.drivers.select_driver` sees:

  - ``target.driver == 'windows_excel'`` AND host is WSL2.

Native Windows hosts use :class:`WindowsExcelOracle` directly; native
macOS hosts have no use for this module.

## Why subprocess instead of in-process COM?

WSL2's Linux Python cannot load the Windows pywin32 COM bridge -- the
two ABIs are incompatible. The only way to drive Excel from a WSL2
shell is to invoke Windows-side ``python.exe`` (which can be mounted as
``/mnt/c/...``) and shuttle data over a serializable channel. JSON
files are slower than a pipe but cleanly survive the encoding boundary
(WSL writes utf-8 directly to /tmp; Windows-side Python reads the same
bytes via the translated path with ``encoding='utf-8'`` pinned).

## Configuration

  - ``target['win_python']``: absolute Windows-side path to
    ``python.exe``. Required. The ``make oracle-setup`` preflight is
    expected to discover a usable interpreter and write it back to
    ``targets.yaml``; until that lands the field must be filled in
    manually.
"""

from __future__ import annotations

import json
import platform
import subprocess
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional

from .base import CaseResult, EnvironmentInfo, OracleDriver


def _ensure_wsl2() -> None:
    """Refuses to start unless the host is a WSL2 kernel.

    Plain Linux without the Microsoft kernel marker would have no
    ``python.exe`` to invoke and no ``wslpath`` to translate paths.
    Surfacing the failure here keeps error messages local to the
    bridge instead of bubbling up as a confused subprocess error.
    """

    if platform.system() != "Linux":
        raise RuntimeError(
            f"WSL bridge requires Linux/WSL2 host, got {platform.system()}"
        )
    try:
        proc = Path("/proc/version").read_text(encoding="utf-8").lower()
    except OSError as exc:
        raise RuntimeError(
            "cannot read /proc/version (WSL2 detection failed)"
        ) from exc
    if "microsoft" not in proc:
        raise RuntimeError(
            "WSL bridge requires WSL2 (no 'microsoft' in /proc/version)"
        )


def _to_windows_path(p: Path) -> str:
    """Translates a WSL path to a Windows path via ``wslpath -w``.

    Raises if ``wslpath`` is missing (only present on WSL) or returns
    an empty string (which would silently break the subprocess
    invocation).
    """

    out = subprocess.check_output(["wslpath", "-w", str(p)], text=True).strip()
    if not out:
        raise RuntimeError(f"wslpath returned empty for {p}")
    return out


class WSLBridgeOracle(OracleDriver):
    """Drives a Windows-side :class:`WindowsExcelOracle` through subprocess.

    Lifetime is bounded by the surrounding ``with`` block: ``__enter__``
    creates a private ``/tmp`` directory for input/output JSON files, and
    ``__exit__`` cleans it up. Each :meth:`probe_environment` /
    :meth:`run_suite` call writes a fresh ``input.json`` and reads the
    corresponding ``output.json``, so callers do not need to coordinate
    file naming.
    """

    def __init__(self, *, win_python: str, visible: bool = False) -> None:
        _ensure_wsl2()
        if not win_python:
            raise RuntimeError(
                "WSL bridge requires `win_python` in targets.yaml; "
                "run `make oracle-setup` to discover it."
            )
        self._win_python = win_python
        self._visible = visible
        self._tmpdir: Optional[tempfile.TemporaryDirectory] = None

    def __enter__(self) -> "WSLBridgeOracle":
        self._tmpdir = tempfile.TemporaryDirectory(prefix="formulon-oracle-")
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._tmpdir is not None:
            self._tmpdir.cleanup()
            self._tmpdir = None

    def _invoke(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """Runs one wire-protocol round-trip against the Windows driver.

        The payload is serialized to ``input.json``, the Windows-side
        Python is invoked with translated Windows paths, and the
        response is parsed from ``output.json``. Subprocess failure
        (non-zero return code or missing output file) is surfaced as a
        ``RuntimeError`` carrying both stdout and stderr so the
        operator can diagnose Office activation / COM issues.
        """

        assert self._tmpdir is not None, "use as context manager"
        tmp = Path(self._tmpdir.name)
        in_path = tmp / "input.json"
        out_path = tmp / "output.json"
        in_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        cmd = [
            self._win_python,
            "-m", "tools.oracle.drivers.windows_excel",
            "--input", _to_windows_path(in_path),
            "--output", _to_windows_path(out_path),
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(
                f"windows_excel subprocess failed (rc={result.returncode}):\n"
                f"stdout: {result.stdout}\nstderr: {result.stderr}"
            )
        if not out_path.exists():
            raise RuntimeError(
                f"windows_excel did not produce output (stdout: {result.stdout})"
            )
        return json.loads(out_path.read_text(encoding="utf-8"))

    def probe_environment(self) -> EnvironmentInfo:
        out = self._invoke(
            {"version": 1, "command": "probe_environment", "visible": self._visible}
        )
        env = out["environment"]
        return EnvironmentInfo(
            excel_version=env.get("excel_version", ""),
            excel_locale=env.get("excel_locale", ""),
            date1904=bool(env.get("date1904", False)),
            iterative=bool(env.get("iterative", False)),
        )

    def run_suite(
        self,
        suite_name: str,
        cases: List[Dict[str, Any]],
        *,
        date1904: bool = False,
        iterative: bool = False,
    ) -> List[CaseResult]:
        out = self._invoke(
            {
                "version": 1,
                "command": "run_suite",
                "suite_name": suite_name,
                "visible": self._visible,
                "date1904": date1904,
                "iterative": iterative,
                "cases": cases,
            }
        )
        return [
            CaseResult(
                id=r["id"],
                kind=r["kind"],
                value=r.get("value"),
                error_code=r.get("error_code"),
                array_shape=r.get("array_shape"),
            )
            for r in out["results"]
        ]
