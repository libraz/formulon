#!/usr/bin/env python3
"""Backward-compat re-export shim for the legacy `driver` module.

New code should import from `tools.oracle.drivers.macos_excel` (Mac
Excel concrete driver) or `tools.oracle.drivers.base` (driver-agnostic
types) directly. This shim only exists so historical entry points
(`from tools.oracle.driver import ExcelOracle, CaseResult,
EnvironmentInfo`) keep working while callers migrate.

The dual try/except mirrors the convention used elsewhere under
`tools/oracle/`: the file may be imported either as
`tools.oracle.driver` (package-style, e.g. via `python -m`) or as the
top-level `driver` module after a `sys.path` insert
(script-style, e.g. via `python tools/oracle/oracle_gen.py`).
"""

from __future__ import annotations

try:  # pragma: no cover - trivial fallback
    from tools.oracle.drivers.base import CaseResult, EnvironmentInfo
    from tools.oracle.drivers.macos_excel import ExcelOracle
except ImportError:  # pragma: no cover
    import sys
    from pathlib import Path

    # When loaded script-style, the parent of this file (tools/oracle/) is
    # already on sys.path but the repo-root namespace package "tools" may
    # not be. Push the local "drivers/" directory directly so the relative
    # imports inside macos_excel still resolve via the same sibling path.
    _here = Path(__file__).resolve().parent
    sys.path.insert(0, str(_here))
    from drivers.base import CaseResult, EnvironmentInfo  # type: ignore
    from drivers.macos_excel import ExcelOracle  # type: ignore

__all__ = ["CaseResult", "EnvironmentInfo", "ExcelOracle"]
