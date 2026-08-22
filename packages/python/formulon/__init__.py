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
    CfColor,
    CfMatch,
    CfValueObject,
    CivilTime,
    ColorScale,
    ColorSpec,
    ColumnLayout,
    Comment,
    CommentEntry,
    ConditionalFormat,
    ConditionalFormatInput,
    DataBar,
    DataValidation,
    DataValidationInput,
    DefinedName,
    DifferentialFormat,
    ErrorCode,
    ExternalLink,
    ExternalLinkKind,
    FillRecord,
    FontRecord,
    FormulonError,
    FunctionMetadata,
    Hyperlink,
    IconSet,
    IterativeSettings,
    LogLevel,
    MergeRange,
    PageBreak,
    PageMargins,
    PageSetup,
    PaginationResult,
    PassthroughPart,
    PhoneticRun,
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
    PivotReportLayout,
    PivotShowValuesAs,
    PivotWorksheetSource,
    ReadDiagnostics,
    RowLayout,
    SaveDiagnostics,
    SheetProtection,
    SheetView,
    SheetVisibility,
    SpillInfo,
    StyleBatchIndices,
    Table,
    Value,
    Workbook,
    WorkbookFormat,
    _check,
    _sint,
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
    "CfColor",
    "CfMatch",
    "CfValueObject",
    "CivilTime",
    "ColorSpec",
    "ColumnLayout",
    "Comment",
    "CommentEntry",
    "ConditionalFormat",
    "ConditionalFormatInput",
    "ColorScale",
    "DataBar",
    "DataValidation",
    "DataValidationInput",
    "DefinedName",
    "DifferentialFormat",
    "ErrorCode",
    "ExternalLink",
    "ExternalLinkKind",
    "FillRecord",
    "LogLevel",
    "FontRecord",
    "FormulonError",
    "FunctionMetadata",
    "FunctionMetadataEntry",
    "FunctionMetadataLocalized",
    "FunctionMetadataProvider",
    "Hyperlink",
    "IterativeSettings",
    "IconSet",
    "MergeRange",
    "MergedFunctionMetadata",
    "PageBreak",
    "PageMargins",
    "PageSetup",
    "PaginationResult",
    "PassthroughPart",
    "PhoneticRun",
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
    "PivotReportLayout",
    "PivotShowValuesAs",
    "PivotWorksheetSource",
    "ReadDiagnostics",
    "RowLayout",
    "SaveDiagnostics",
    "SheetProtection",
    "SheetView",
    "SheetVisibility",
    "SpillInfo",
    "StyleBatchIndices",
    "Table",
    "Value",
    "ValueKind",
    "Workbook",
    "WorkbookFormat",
    "__version__",
    "error_display_name",
    "eval_formula",
    "library_version",
    "set_log_min_level",
    "merge_function_metadata",
    "version_string",
]


def library_version() -> str:
    """Return the version string baked into the loaded WASM module.

    Returns:
      The result of ``fm_version_string()``. Always non-empty.
    """
    return decode_cstr(LIB.fm_version_string())


def error_display_name(error_code: int) -> str:
    """Return an Excel literal such as ``"#DIV/0!"`` for an error code.

    Unknown numeric values return ``"#UNKNOWN!"``.
    """
    return decode_cstr(LIB.fm_error_display_name(_sint(error_code, "error")))


# Backward-compat alias mirroring the npm binding's ``versionString`` name.
version_string = library_version


def set_log_min_level(level: int) -> None:
    """Set the engine's minimum structured-log severity.

    This is **process-wide** state, not per :class:`Workbook`, so it is a
    module-level function rather than a method.

    The default is :attr:`LogLevel.OFF`, under which the engine writes
    nothing: an embedded library must not write to the host's stderr
    unless the host asks it to. Raising the threshold to
    :attr:`LogLevel.WARN` or below makes the engine emit one JSON record
    per line on stderr, including the per-cell XLSB downgrade diagnostics.

    Args:
      level: a :class:`LogLevel` ordinal.

    Raises:
      FormulonError: when ``level`` is outside :class:`LogLevel`.
    """
    _check(LIB.fm_set_log_min_level(_sint(level, "level")), "fm_set_log_min_level")


def eval_formula(formula: str) -> Value:
    """Evaluate a single formula against a fresh, default workbook.

    The formula is evaluated read-only as if entered at ``Sheet1!A1``;
    the temporary workbook is not mutated or recalculated.

    Cell-level Excel errors (``#DIV/0!``, ``#VALUE!``, etc.) surface as a
    :class:`Value` with ``kind == ValueKind.ERROR``; only host-side
    failures (parser crash, OOM, NULL handle) raise :class:`FormulonError`.

    Args:
      formula: the formula text, with or without a leading ``=``.

    Returns:
      The scalar :class:`Value` at the formula's top-left result cell.
    """
    with Workbook.create_default() as wb:
        return wb.evaluate_formula_array(0, 0, 0, formula)[0][0]
