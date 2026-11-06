"""Formulon -- Excel 365 calculation engine, Python binding.

Public API:
  * :func:`eval_formula` -- one-shot formula evaluation.
  * :func:`get_version` -- query the version of the bundled engine.
  * :class:`Workbook` -- full workbook lifecycle.
  * :class:`Sheet` -- A1-addressed view of one sheet.
  * :class:`Value`, :class:`ValueKind` -- cell value.
  * :class:`ExcelError` -- an Excel cell error as a Python value.
  * :class:`FormulonError` and its subclasses -- host-side failures.

Cell-level Excel errors (``#DIV/0!`` and friends) surface as ``Value``
instances with ``kind == ValueKind.ERROR`` (``ExcelError`` through
:meth:`Value.to_python`); only host-side failures (a closed workbook, an
unreadable archive in ``Workbook.load``, out of memory) raise
:class:`FormulonError`.
"""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError
from importlib.metadata import version as _dist_version
from pathlib import Path

from ._c import (
    LIB,
    IndexOutOfRangeError,
    InvalidArgumentError,
    NotFoundError,
    ValueKind,
    WorkbookClosedError,
    decode_cstr,
)
from ._status import StatusCode
from .metadata import (
    FunctionMetadataEntry,
    FunctionMetadataLocalized,
    FunctionMetadataProvider,
    MergedFunctionMetadata,
    merge_function_metadata,
)
from .workbook import (
    PIVOT_VALUES_FIELD,
    AnchorEditAs,
    AnchorKind,
    ArrayFormula,
    AutoFilter,
    AutoFilterColumn,
    AutoFilterDateGroup,
    AutoFilterEval,
    AutoFilterSortCondition,
    AutoFilterSortState,
    CalcMode,
    Cell,
    CellNode,
    CellRect,
    CellsInRange,
    CellStyle,
    CellXf,
    CfCell,
    CfColor,
    CfIcon,
    CfMatch,
    Cfvo,
    Chart,
    ChartAxis,
    ChartAxisCrosses,
    ChartAxisKind,
    ChartAxisPatch,
    ChartAxisTitleText,
    ChartBlanksAs,
    ChartCategories,
    ChartData,
    ChartDataSeries,
    ChartDataSource,
    ChartDisplayPatch,
    ChartFill,
    ChartFillPatch,
    ChartGrouping,
    ChartInsertKind,
    ChartLabelPosition,
    ChartLabels,
    ChartLabelsPatch,
    ChartLegend,
    ChartLegendPosition,
    ChartLine,
    ChartLineDash,
    ChartLinePatch,
    ChartMarker,
    ChartMarkerPatch,
    ChartMarkerSymbol,
    ChartPlot,
    ChartPlotPatch,
    ChartPlotType,
    ChartPoint,
    ChartPointPatch,
    ChartPointState,
    ChartPreviewOverridesInput,
    ChartSeries,
    ChartSeriesDirection,
    ChartSeriesPatch,
    ChartSeriesRanges,
    ChartSource,
    ChartTitle,
    ChartTitleKind,
    ChartUnresolved,
    CivilTime,
    CollationClasses,
    CollationStrength,
    Color,
    ColorContext,
    ColorResolution,
    ColorScale,
    ColumnLayout,
    Comment,
    ConditionalFormat,
    ConditionalFormatInput,
    DataBar,
    DataTable,
    DataTableInputCell,
    DataValidation,
    DataValidationInput,
    DateTimeGrouping,
    DefinedName,
    DifferentialFormat,
    DisplayStatus,
    DisplayText,
    DrawingObject,
    DrawingObjectKind,
    DynamicFilterType,
    EffectiveStyle,
    EffectiveStyleSource,
    ErrorCode,
    ErrorName,
    ExcelError,
    ExternalLink,
    ExternalLinkKind,
    Fill,
    FilterKind,
    FilterOperator,
    Font,
    FormulonError,
    FunctionMetadata,
    GeometryMode,
    GradientFill,
    GradientStop,
    Hyperlink,
    IconSet,
    ImageFormat,
    ImageInfo,
    IterativeSettings,
    LocaleFacts,
    LogLevel,
    MarginsPt,
    Mention,
    PageBreak,
    PageLayout,
    PageMargins,
    PageSetup,
    Pagination,
    Paper,
    ParsedReference,
    PassthroughPart,
    Person,
    PhoneticProperties,
    PhoneticPropertiesInput,
    PhoneticRun,
    PivotAggregation,
    PivotAxis,
    PivotCacheOptions,
    PivotCalculatedItem,
    PivotCalendar,
    PivotCell,
    PivotCellKind,
    PivotDataFieldInput,
    PivotDatePeriod,
    PivotField,
    PivotFieldInput,
    PivotFilterInput,
    PivotFilterType,
    PivotFilterValueKind,
    PivotGrandTotals,
    PivotItem,
    PivotLayout,
    PivotOptions,
    PivotReportLayout,
    PivotShowValuesAs,
    PivotSortOrder,
    PivotStyle,
    PivotWorksheetSource,
    PrintTitles,
    ProtectedRange,
    ReadDiagnostics,
    RectPt,
    RefMode,
    ResolvedColor,
    RowLayout,
    SaveDiagnostics,
    Shape,
    ShapePaint,
    Sheet,
    SheetComment,
    SheetFormatDefaults,
    SheetPrintTitles,
    SheetProtection,
    SheetRect,
    SheetSelection,
    SheetView,
    SheetVisibility,
    Slicer,
    SlicerCrossFilter,
    SlicerSortOrder,
    SortBy,
    SortMethod,
    Sparkline,
    SparklineAxisType,
    SparklineEmptyCells,
    SparklineGroup,
    SparklineType,
    Spill,
    StyleBatchIndices,
    Table,
    TableColumn,
    TableStyle,
    TableStyleDefaults,
    TableStyleElement,
    Theme,
    ThemeFonts,
    ThemeScriptFont,
    ThemeSource,
    ThreadedComment,
    ValidationErrorStyle,
    ValidationOutcome,
    Value,
    WidthModel,
    Workbook,
    WorkbookFormat,
    WorkbookProtection,
    _check,
    _sint,
    canonicalize_formula,
    canonicalize_function_name,
    compare_text,
    get_collation_classes,
    get_function_metadata,
    get_locale_facts,
    list_function_names,
    list_locale_error_names,
    localize_formula,
    localize_function_name,
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
    return decode_cstr(LIB.fm_get_version())


__version__ = _resolve_version()

__all__ = [
    "AutoFilterDateGroup",
    "AutoFilterColumn",
    "AutoFilterSortCondition",
    "AutoFilterSortState",
    "AutoFilter",
    "AutoFilterEval",
    "Mention",
    "Person",
    "ThreadedComment",
    "ValidationOutcome",
    "CalcMode",
    "RefMode",
    "Cell",
    "CellsInRange",
    "CellNode",
    "CellStyle",
    "CellXf",
    "CfCell",
    "CfColor",
    "CfIcon",
    "CfMatch",
    "Cfvo",
    "CivilTime",
    "CollationClasses",
    "CollationStrength",
    "Color",
    "ColumnLayout",
    "Comment",
    "SheetComment",
    "ConditionalFormat",
    "ConditionalFormatInput",
    "ColorScale",
    "DataBar",
    "DataValidation",
    "DataValidationInput",
    "DefinedName",
    "DifferentialFormat",
    "ErrorCode",
    "ErrorName",
    "ExcelError",
    "ExternalLink",
    "ExternalLinkKind",
    "Fill",
    "LogLevel",
    "Font",
    "GradientFill",
    "GradientStop",
    "FormulonError",
    "IndexOutOfRangeError",
    "InvalidArgumentError",
    "NotFoundError",
    "WorkbookClosedError",
    "StatusCode",
    "FunctionMetadata",
    "FunctionMetadataEntry",
    "FunctionMetadataLocalized",
    "FunctionMetadataProvider",
    "Hyperlink",
    "IterativeSettings",
    "PivotGrandTotals",
    "SheetPrintTitles",
    "SheetSelection",
    "TableStyleDefaults",
    "IconSet",
    "LocaleFacts",
    "CellRect",
    "MergedFunctionMetadata",
    "PageBreak",
    "PageMargins",
    "PageSetup",
    "Pagination",
    "PassthroughPart",
    "PhoneticProperties",
    "PhoneticPropertiesInput",
    "PhoneticRun",
    "PivotAggregation",
    "PivotAxis",
    "PivotCacheOptions",
    "PivotCalculatedItem",
    "PivotCalendar",
    "PivotCell",
    "PivotCellKind",
    "PivotDataFieldInput",
    "PivotDatePeriod",
    "PivotField",
    "PivotFieldInput",
    "PivotFilterInput",
    "PivotFilterType",
    "PivotFilterValueKind",
    "PIVOT_VALUES_FIELD",
    "PivotItem",
    "PivotStyle",
    "PivotLayout",
    "PivotOptions",
    "PivotReportLayout",
    "PivotShowValuesAs",
    "PivotSortOrder",
    "PivotWorksheetSource",
    "ProtectedRange",
    "ReadDiagnostics",
    "RowLayout",
    "SaveDiagnostics",
    "SheetProtection",
    "SheetRect",
    "Sheet",
    "SheetView",
    "Shape",
    "ShapePaint",
    "SheetVisibility",
    "DataTable",
    "DataTableInputCell",
    "Slicer",
    "SlicerCrossFilter",
    "SlicerSortOrder",
    "Sparkline",
    "SparklineAxisType",
    "SparklineEmptyCells",
    "SparklineGroup",
    "SparklineType",
    "Spill",
    "ParsedReference",
    "ArrayFormula",
    "StyleBatchIndices",
    "Table",
    "TableColumn",
    "TableStyle",
    "TableStyleElement",
    "Value",
    "ValueKind",
    "DisplayStatus",
    "DisplayText",
    "GeometryMode",
    "MarginsPt",
    "PageLayout",
    "Paper",
    "PrintTitles",
    "RectPt",
    "SheetFormatDefaults",
    "ColorContext",
    "AnchorEditAs",
    "AnchorKind",
    "ChartAxisCrosses",
    "ChartAxis",
    "ChartAxisKind",
    "ChartAxisPatch",
    "ChartDisplayPatch",
    "ChartAxisTitleText",
    "ChartBlanksAs",
    "ChartCategories",
    "ChartData",
    "ChartDataSeries",
    "ChartDataSource",
    "ChartFill",
    "ChartFillPatch",
    "ChartGrouping",
    "Chart",
    "ChartInsertKind",
    "ChartLabelPosition",
    "ChartLabels",
    "ChartLabelsPatch",
    "ChartLegend",
    "ChartLegendPosition",
    "ChartLineDash",
    "ChartLine",
    "ChartLinePatch",
    "ChartMarker",
    "ChartMarkerPatch",
    "ChartMarkerSymbol",
    "ChartPlot",
    "ChartPlotPatch",
    "ChartPlotType",
    "ChartPoint",
    "ChartPointPatch",
    "ChartPointState",
    "ChartPreviewOverridesInput",
    "ChartSeriesDirection",
    "ChartSeries",
    "ChartSeriesPatch",
    "ChartSeriesRanges",
    "ChartSource",
    "ChartTitle",
    "ChartTitleKind",
    "ChartUnresolved",
    "DrawingObject",
    "DrawingObjectKind",
    "ImageFormat",
    "ImageInfo",
    "FilterKind",
    "FilterOperator",
    "DynamicFilterType",
    "SortBy",
    "SortMethod",
    "DateTimeGrouping",
    "ValidationErrorStyle",
    "ColorResolution",
    "EffectiveStyle",
    "ResolvedColor",
    "EffectiveStyleSource",
    "Theme",
    "ThemeFonts",
    "ThemeScriptFont",
    "ThemeSource",
    "WidthModel",
    "Workbook",
    "WorkbookProtection",
    "WorkbookFormat",
    "__version__",
    "get_error_display_name",
    "eval_formula",
    "get_version",
    "set_log_min_level",
    "merge_function_metadata",
    "canonicalize_formula",
    "canonicalize_function_name",
    "compare_text",
    "get_collation_classes",
    "get_function_metadata",
    "get_locale_facts",
    "list_function_names",
    "list_locale_error_names",
    "localize_formula",
    "localize_function_name",
]


def get_version() -> str:
    """Return the version string of the bundled engine.

    Always non-empty.
    """
    return decode_cstr(LIB.fm_get_version())


def get_error_display_name(error_code: ErrorCode) -> str:
    """Return an Excel literal such as ``"#DIV/0!"`` for an error code.

    Unknown numeric values return ``"#UNKNOWN!"``.
    """
    return decode_cstr(LIB.fm_get_error_display_name(_sint(error_code, "error")))


def set_log_min_level(level: LogLevel) -> None:
    """Set the engine's minimum structured-log severity.

    This is **process-wide** state, not per :class:`Workbook`, so it is a
    module-level function rather than a method.

    The default is :attr:`LogLevel.OFF`, under which the engine writes
    nothing: an embedded library must not write to the host's stderr
    unless the host asks it to. Raising the threshold to
    :attr:`LogLevel.WARN` or below makes the engine emit one JSON record
    per line on stderr, including the per-cell XLSB downgrade diagnostics.

    Args:
      level: a :class:`LogLevel` member.

    Raises:
      FormulonError: when ``level`` is outside :class:`LogLevel`.
    """
    _check(LIB.fm_set_log_min_level(_sint(level, "level")))


def eval_formula(formula: str) -> Value:
    """Evaluate a single formula against a fresh, default workbook.

    The formula is evaluated read-only as if entered at ``Sheet1!A1``;
    the temporary workbook is not mutated or recalculated.

    Cell-level Excel errors (``#DIV/0!``, ``#VALUE!``, etc.) surface as a
    :class:`Value` with ``kind == ValueKind.ERROR``; only host-side
    failures (a parser crash, out of memory) raise :class:`FormulonError`.

    Args:
      formula: the formula text, with or without a leading ``=``.

    Returns:
      The scalar :class:`Value` at the formula's top-left result cell.
    """
    with Workbook.create_default() as wb:
        return wb._evaluate_formula_top_left(0, 0, 0, formula)
