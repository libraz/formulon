# Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
"""Formulon -- Excel 365 calculation engine, Python binding.

Public API:
  * :func:`eval_formula` -- one-shot formula evaluation.
  * :func:`library_version` -- query the underlying ``libformulon`` build.
  * :class:`Workbook` -- full workbook lifecycle.
  * :class:`Value`, :class:`ValueKind` -- cell value POD.
  * :class:`FormulonError` -- host-side failure exception.

Cell-level Excel errors (``#DIV/0!`` and friends) surface as ``Value``
instances with ``kind == ValueKind.ERROR``; only host-side failures
(NULL handle, parse errors during ``Workbook.load``, OOM) raise
:class:`FormulonError`.
"""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError
from importlib.metadata import version as _dist_version
from pathlib import Path

from ._c import LIB, ValueKind, decode_cstr
from .metadata import (
    FunctionMetadataEntry,
    FunctionMetadataLocalized,
    FunctionMetadataProvider,
    MergedFunctionMetadata,
    merge_function_metadata,
)
from .workbook import (
    CalcMode,
    Cell,
    CellNode,
    CellStyle,
    CellXf,
    CfCellResult,
    CfMatch,
    ColumnLayout,
    Comment,
    ConditionalFormat,
    ConditionalFormatInput,
    DataValidation,
    DataValidationInput,
    DefinedName,
    ExternalLink,
    FillRecord,
    FontRecord,
    FormulonError,
    FunctionMetadata,
    Hyperlink,
    MergeRange,
    PassthroughPart,
    PivotAggregation,
    PivotAxis,
    PivotCalendar,
    PivotCell,
    PivotCellKind,
    PivotDataFieldSpec,
    PivotDateGrouping,
    PivotFieldSpec,
    PivotFilterSpec,
    PivotFilterType,
    PivotFilterValueKind,
    PivotLayout,
    PivotShowValuesAs,
    RowLayout,
    SheetProtection,
    SheetView,
    SpillInfo,
    Table,
    Value,
    Workbook,
    WorkbookFormat,
)

def _resolve_version() -> str:
    """Resolve the package version from its single source of truth.

    The canonical version lives in ``[project].version`` of
    ``packages/python/pyproject.toml`` (the same value a release bumps).
    For an installed wheel that value is copied into the dist metadata, so
    :func:`importlib.metadata.version` returns it without re-reading the
    project file. When running from an uninstalled source tree there is no
    dist metadata; fall back to parsing ``pyproject.toml`` directly, then
    to the C-ABI build version as a last resort. The number is never
    hardcoded here, so a release bumps exactly one place.
    """
    try:
        return _dist_version("formulon")
    except PackageNotFoundError:
        pass

    # Source-tree fallback: read [project].version from the sibling
    # pyproject.toml. `tomllib` is stdlib from 3.11; older interpreters
    # fall back to a minimal line scan to avoid a hard `tomli` dependency.
    pyproject = Path(__file__).resolve().parent.parent / "pyproject.toml"
    try:
        raw = pyproject.read_bytes()
    except OSError:
        raw = b""
    if raw:
        try:
            import tomllib

            parsed = tomllib.loads(raw.decode("utf-8"))
            project_version = parsed.get("project", {}).get("version")
            if isinstance(project_version, str) and project_version:
                return project_version
        except ModuleNotFoundError:
            for line in raw.decode("utf-8").splitlines():
                stripped = line.strip()
                if stripped.startswith("version") and "=" in stripped:
                    candidate = stripped.split("=", 1)[1].strip().strip("\"'")
                    if candidate:
                        return candidate
                    break

    # Last resort: the version baked into the loaded WASM module.
    return decode_cstr(LIB.fm_version_string())


__version__ = _resolve_version()

__all__ = [
    "CalcMode",
    "Cell",
    "CellNode",
    "CellStyle",
    "CellXf",
    "CfCellResult",
    "CfMatch",
    "ColumnLayout",
    "Comment",
    "ConditionalFormat",
    "ConditionalFormatInput",
    "DataValidation",
    "DataValidationInput",
    "DefinedName",
    "ExternalLink",
    "FillRecord",
    "FontRecord",
    "FormulonError",
    "FunctionMetadata",
    "FunctionMetadataEntry",
    "FunctionMetadataLocalized",
    "FunctionMetadataProvider",
    "Hyperlink",
    "MergeRange",
    "MergedFunctionMetadata",
    "PassthroughPart",
    "PivotAggregation",
    "PivotAxis",
    "PivotCalendar",
    "PivotCell",
    "PivotCellKind",
    "PivotDataFieldSpec",
    "PivotDateGrouping",
    "PivotFieldSpec",
    "PivotFilterSpec",
    "PivotFilterType",
    "PivotFilterValueKind",
    "PivotLayout",
    "PivotShowValuesAs",
    "RowLayout",
    "SheetProtection",
    "SheetView",
    "SpillInfo",
    "Table",
    "Value",
    "ValueKind",
    "Workbook",
    "WorkbookFormat",
    "__version__",
    "eval_formula",
    "library_version",
    "merge_function_metadata",
    "version_string",
]


def library_version() -> str:
    """Return the version string baked into the loaded WASM module.

    Returns:
      The result of ``fm_version_string()``. Always non-empty.
    """
    return decode_cstr(LIB.fm_version_string())


# Backward-compat alias mirroring the npm binding's ``versionString`` name.
version_string = library_version


def eval_formula(formula: str) -> Value:
    """Evaluate a single formula against a fresh, default workbook.

    The formula is written to ``Sheet1!A1``, the workbook is recalculated,
    and the resulting cell value is returned.

    Cell-level Excel errors (``#DIV/0!``, ``#VALUE!``, etc.) surface as a
    :class:`Value` with ``kind == ValueKind.ERROR``; only host-side
    failures (parser crash, OOM, NULL handle) raise :class:`FormulonError`.

    Args:
      formula: the formula text, with or without a leading ``=``.

    Returns:
      The :class:`Value` cached at ``Sheet1!A1`` after recalc.
    """
    with Workbook.create_default() as wb:
        wb.set_formula(0, 0, 0, formula)
        wb.recalc()
        return wb.get_value(0, 0, 0)
