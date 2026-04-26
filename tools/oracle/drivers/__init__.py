"""Oracle driver factory + re-exports for convenience.

This package hosts the concrete oracle backends:

  - :mod:`tools.oracle.drivers.macos_excel` -- xlwings + AppleScript on
    Darwin. Drives the canonical Mac Excel 365 / ja-JP reference.
  - :mod:`tools.oracle.drivers.windows_excel` -- xlwings + COM on
    Windows. Drives the variant Windows Excel 365 / ja-JP reference.
  - :mod:`tools.oracle.drivers.wsl_bridge` -- WSL2 -> Windows Excel
    subprocess bridge. Used automatically when WSL2 hosts a
    ``windows_excel`` target.

Use :func:`select_driver` to pick the right one for the current host
based on a ``targets.yaml`` record. Direct imports of concrete modules
are also fine for callers that already know which platform they are on
(e.g. tests, the legacy :mod:`tools.oracle.driver` shim).
"""

from __future__ import annotations

import platform
from pathlib import Path
from typing import Any, Dict

from .base import CaseResult, EnvironmentInfo, OracleDriver

__all__ = ["CaseResult", "EnvironmentInfo", "OracleDriver", "select_driver"]


def _is_wsl2() -> bool:
    """Heuristic: WSL2 always exposes ``microsoft`` in ``/proc/version``."""

    if platform.system() != "Linux":
        return False
    try:
        return "microsoft" in Path("/proc/version").read_text(encoding="utf-8").lower()
    except OSError:
        return False


def select_driver(target: Dict[str, Any], *, visible: bool = False) -> OracleDriver:
    """Returns an :class:`OracleDriver` instance for `target` on the current host.

    Resolution table:

      ``target.driver == 'macos_excel'``  -> :class:`MacExcelOracle`
        (Darwin only)
      ``target.driver == 'windows_excel'`` on Windows host
        -> :class:`WindowsExcelOracle` (direct COM)
      ``target.driver == 'windows_excel'`` on WSL2 host
        -> :class:`WSLBridgeOracle` (subprocess to Windows-side Python)
      ``target.driver == 'wsl_bridge'`` (explicit)
        -> :class:`WSLBridgeOracle`

    Raises ``RuntimeError`` when the host can't reach the requested
    driver. Callers should treat that as a hard CLI error (exit
    non-zero) rather than silently falling back; misrouting goldens to
    the wrong target dir is much harder to detect after the fact.
    """

    driver_name = target.get("driver")
    host = platform.system()

    if driver_name == "macos_excel":
        if host != "Darwin":
            raise RuntimeError(
                f"target driver 'macos_excel' requires Darwin host, got {host}"
            )
        from .macos_excel import ExcelOracle as MacExcelOracle

        return MacExcelOracle(visible=visible)

    if driver_name == "windows_excel":
        if host == "Windows":
            from .windows_excel import WindowsExcelOracle

            return WindowsExcelOracle(visible=visible)
        if _is_wsl2():
            from .wsl_bridge import WSLBridgeOracle

            return WSLBridgeOracle(
                win_python=str(target.get("win_python") or ""),
                visible=visible,
            )
        raise RuntimeError(
            f"target driver 'windows_excel' needs Windows or WSL2 host, got {host}"
        )

    if driver_name == "wsl_bridge":
        if not _is_wsl2():
            raise RuntimeError("'wsl_bridge' driver requires WSL2 host")
        from .wsl_bridge import WSLBridgeOracle

        return WSLBridgeOracle(
            win_python=str(target.get("win_python") or ""),
            visible=visible,
        )

    raise RuntimeError(f"unknown driver: {driver_name!r}")
