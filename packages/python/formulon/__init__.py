"""Formulon -- Excel 365 calculation engine, Python binding.

Public API:
  * :func:`eval_formula` -- one-shot formula evaluation.
  * :func:`get_version` -- query the version of the bundled engine.
  * :class:`Workbook` -- full workbook lifecycle.
  * :class:`Sheet` -- every operation on one sheet of a workbook.
  * :class:`Value`, :class:`ValueKind` -- cell value.
  * :class:`ExcelError` -- an Excel cell error as a Python value;
    :attr:`ErrorCode.display_name` spells a code as Excel does (``#DIV/0!``).
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
)
from ._c import decode_cstr as _decode_cstr
from ._status import StatusCode
from .metadata import (
    FunctionMetadataEntry,
    FunctionMetadataLocalized,
    FunctionMetadataProvider,
    MergedFunctionMetadata,
    merge_function_metadata,
)
from .sheet import Sheet
from .workbook import (
    PIVOT_SHOW_AS_BASE_NEXT,
    PIVOT_SHOW_AS_BASE_PREVIOUS,
    PIVOT_VALUES_FIELD,
    UNCHANGED,
    AnchorEditAs,
    AnchorKind,
    ArrayFormula,
    AutoFilter,
    AutoFilterColumn,
    AutoFilterDateGroup,
    AutoFilterEval,
    AutoFilterSortCondition,
    AutoFilterSortState,
    Border,
    BorderSide,
    CalcMode,
    Cell,
    CellNode,
    CellRect,
    CellsInRange,
    CellStyle,
    CellXf,
    CfCell,
    CfColor,
    CfDataBarAxisPosition,
    CfDataBarDirection,
    CfIcon,
    CfMatch,
    CfMatchKind,
    Cfvo,
    Chart,
    ChartAxis,
    ChartAxisCrosses,
    ChartAxisKind,
    ChartAxisPatch,
    ChartAxisPosition,
    ChartAxisTitleText,
    ChartBarDirection,
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
    CollationClasses,
    CollationStrength,
    Color,
    ColorContext,
    ColorKind,
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
    DefinedNameInput,
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
    ExcelProfileId,
    ExternalLink,
    ExternalLinkKind,
    Fill,
    FilterKind,
    FilterOperator,
    Font,
    FormulonError,
    FunctionAvailability,
    FunctionMetadata,
    GeometryMode,
    GradientFill,
    GradientStop,
    Hyperlink,
    IconSet,
    ImageFormat,
    ImageInfo,
    IterativeSettings,
    LocaleDateOrder,
    LocaleFacts,
    LogLevel,
    MarginsPt,
    Mention,
    PageBreak,
    PageLayout,
    PageMargins,
    PageOrientation,
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
    PivotCacheField,
    PivotCacheOptions,
    PivotCacheSharedItem,
    PivotCalculatedItem,
    PivotCalendar,
    PivotCell,
    PivotCellKind,
    PivotDataField,
    PivotDataFieldInput,
    PivotDatePeriod,
    PivotField,
    PivotFieldInput,
    PivotFilter,
    PivotFilterInput,
    PivotFilterType,
    PivotFilterValueKind,
    PivotGrandTotals,
    PivotInput,
    PivotItem,
    PivotLayout,
    PivotOptions,
    PivotReportLayout,
    PivotShowValuesAs,
    PivotSortOrder,
    PivotStyle,
    PivotSummary,
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
    SheetComment,
    SheetFormatDefaults,
    SheetProtection,
    SheetRect,
    SheetRef,
    SheetSelection,
    SheetView,
    SheetViewMode,
    SheetVisibility,
    Slicer,
    SlicerCrossFilter,
    SlicerSortOrder,
    SortBy,
    SortMethod,
    Span,
    Sparkline,
    SparklineAxisType,
    SparklineEmptyCells,
    SparklineGroup,
    SparklineType,
    Spill,
    Table,
    TableColumn,
    TableStyle,
    TableStyleDefaults,
    TableStyleElement,
    ThemeColors,
    ThemeColorsInput,
    ThemeFonts,
    ThemeScriptFont,
    ThemeScriptFontInput,
    ThemeSource,
    ThreadedComment,
    ThreadedCommentInput,
    ValidationErrorStyle,
    ValidationOutcome,
    Value,
    WidthModel,
    Workbook,
    WorkbookFormat,
    WorkbookProtection,
    canonicalize_formula,
    canonicalize_function_name,
    compare_text,
    get_collation_classes,
    get_function_metadata,
    get_image_info,
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
    return _decode_cstr(LIB.fm_get_version())


__version__ = _resolve_version()

__all__ = [
    "__version__",
    "AnchorEditAs",
    "AnchorKind",
    "ArrayFormula",
    "AutoFilter",
    "AutoFilterColumn",
    "AutoFilterDateGroup",
    "AutoFilterEval",
    "AutoFilterSortCondition",
    "AutoFilterSortState",
    "Border",
    "BorderSide",
    "CalcMode",
    "canonicalize_formula",
    "canonicalize_function_name",
    "Cell",
    "CellNode",
    "CellRect",
    "CellsInRange",
    "CellStyle",
    "CellXf",
    "CfCell",
    "CfColor",
    "CfDataBarAxisPosition",
    "CfDataBarDirection",
    "CfIcon",
    "CfMatch",
    "CfMatchKind",
    "Cfvo",
    "Chart",
    "ChartAxis",
    "ChartAxisCrosses",
    "ChartAxisKind",
    "ChartAxisPatch",
    "ChartAxisPosition",
    "ChartAxisTitleText",
    "ChartBarDirection",
    "ChartBlanksAs",
    "ChartCategories",
    "ChartData",
    "ChartDataSeries",
    "ChartDataSource",
    "ChartDisplayPatch",
    "ChartFill",
    "ChartFillPatch",
    "ChartGrouping",
    "ChartInsertKind",
    "ChartLabelPosition",
    "ChartLabels",
    "ChartLabelsPatch",
    "ChartLegend",
    "ChartLegendPosition",
    "ChartLine",
    "ChartLineDash",
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
    "ChartSeries",
    "ChartSeriesDirection",
    "ChartSeriesPatch",
    "ChartSeriesRanges",
    "ChartSource",
    "ChartTitle",
    "ChartTitleKind",
    "ChartUnresolved",
    "CollationClasses",
    "CollationStrength",
    "Color",
    "ColorContext",
    "ColorKind",
    "ColorResolution",
    "ColorScale",
    "ColumnLayout",
    "Comment",
    "compare_text",
    "ConditionalFormat",
    "ConditionalFormatInput",
    "DataBar",
    "DataTable",
    "DataTableInputCell",
    "DataValidation",
    "DataValidationInput",
    "DateTimeGrouping",
    "DefinedName",
    "DefinedNameInput",
    "DifferentialFormat",
    "DisplayStatus",
    "DisplayText",
    "DrawingObject",
    "DrawingObjectKind",
    "DynamicFilterType",
    "EffectiveStyle",
    "EffectiveStyleSource",
    "ErrorCode",
    "ErrorName",
    "eval_formula",
    "ExcelError",
    "ExcelProfileId",
    "ExternalLink",
    "ExternalLinkKind",
    "Fill",
    "FilterKind",
    "FilterOperator",
    "Font",
    "FormulonError",
    "FunctionAvailability",
    "FunctionMetadata",
    "FunctionMetadataEntry",
    "FunctionMetadataLocalized",
    "FunctionMetadataProvider",
    "GeometryMode",
    "get_collation_classes",
    "get_function_metadata",
    "get_image_info",
    "get_locale_facts",
    "get_status_name",
    "get_version",
    "GradientFill",
    "GradientStop",
    "Hyperlink",
    "IconSet",
    "ImageFormat",
    "ImageInfo",
    "IndexOutOfRangeError",
    "InvalidArgumentError",
    "IterativeSettings",
    "list_function_names",
    "list_locale_error_names",
    "LocaleDateOrder",
    "LocaleFacts",
    "localize_formula",
    "localize_function_name",
    "LogLevel",
    "MarginsPt",
    "Mention",
    "merge_function_metadata",
    "MergedFunctionMetadata",
    "NotFoundError",
    "PageBreak",
    "PageLayout",
    "PageMargins",
    "PageOrientation",
    "PageSetup",
    "Pagination",
    "Paper",
    "ParsedReference",
    "PassthroughPart",
    "Person",
    "PhoneticProperties",
    "PhoneticPropertiesInput",
    "PhoneticRun",
    "PIVOT_SHOW_AS_BASE_NEXT",
    "PIVOT_SHOW_AS_BASE_PREVIOUS",
    "PIVOT_VALUES_FIELD",
    "PivotAggregation",
    "PivotAxis",
    "PivotCacheField",
    "PivotCacheOptions",
    "PivotCacheSharedItem",
    "PivotCalculatedItem",
    "PivotCalendar",
    "PivotCell",
    "PivotCellKind",
    "PivotDataField",
    "PivotDataFieldInput",
    "PivotDatePeriod",
    "PivotField",
    "PivotFieldInput",
    "PivotFilter",
    "PivotFilterInput",
    "PivotFilterType",
    "PivotFilterValueKind",
    "PivotGrandTotals",
    "PivotInput",
    "PivotItem",
    "PivotLayout",
    "PivotOptions",
    "PivotReportLayout",
    "PivotShowValuesAs",
    "PivotSortOrder",
    "PivotStyle",
    "PivotSummary",
    "PivotWorksheetSource",
    "PrintTitles",
    "ProtectedRange",
    "ReadDiagnostics",
    "RectPt",
    "RefMode",
    "ResolvedColor",
    "RowLayout",
    "SaveDiagnostics",
    "set_log_min_level",
    "Shape",
    "ShapePaint",
    "Sheet",
    "SheetComment",
    "SheetFormatDefaults",
    "SheetProtection",
    "SheetRect",
    "SheetRef",
    "SheetSelection",
    "SheetView",
    "SheetViewMode",
    "SheetVisibility",
    "Slicer",
    "SlicerCrossFilter",
    "SlicerSortOrder",
    "SortBy",
    "SortMethod",
    "Span",
    "Sparkline",
    "SparklineAxisType",
    "SparklineEmptyCells",
    "SparklineGroup",
    "SparklineType",
    "Spill",
    "StatusCode",
    "Table",
    "TableColumn",
    "TableStyle",
    "TableStyleDefaults",
    "TableStyleElement",
    "ThemeColors",
    "ThemeColorsInput",
    "ThemeFonts",
    "ThemeScriptFont",
    "ThemeScriptFontInput",
    "ThemeSource",
    "ThreadedComment",
    "ThreadedCommentInput",
    "UNCHANGED",
    "ValidationErrorStyle",
    "ValidationOutcome",
    "Value",
    "ValueKind",
    "WidthModel",
    "Workbook",
    "WorkbookClosedError",
    "WorkbookFormat",
    "WorkbookProtection",
]


def get_version() -> str:
    """Return the version string of the bundled engine.

    Always non-empty.
    """
    return _decode_cstr(LIB.fm_get_version())


def get_status_name(status: StatusCode) -> str:
    """Return the name of a status code without the ``FM_STATUS_`` prefix (``"INVALID_ARGUMENT"``).

    An unknown code returns ``"UNRECOGNIZED_STATUS"``.
    """
    from ._c import _sint

    return _decode_cstr(LIB.fm_get_status_name(_sint(status, "status")))


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
    from ._c import _sint
    from .workbook import _check

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
