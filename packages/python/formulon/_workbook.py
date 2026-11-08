"""Pythonic wrapper around the Formulon engine.

Public surface:
  * :class:`Workbook` -- context manager owning one engine workbook.
  * :class:`Sheet` -- one sheet of a workbook (defined in :mod:`formulon._sheet`).
  * :class:`Value` -- frozen dataclass holding a cell value.
  * :class:`ExcelError` -- an Excel cell error as a Python value.
  * :class:`FormulonError` -- exception raised on host-side failures.
  * :class:`RangeCell`, :class:`DefinedName`, :class:`TableProperties` -- iteration tuples.

Reads return native values (:data:`NativeValue`): a cell-level Excel error
(``#DIV/0!``, ``#VALUE!``, ...) is an :class:`ExcelError` value, not an
exception. Host-side problems (a closed workbook, an unreadable archive in
``Workbook.load``, out of memory) raise :class:`FormulonError`.

Sheet, row, column, and count arguments are 32-bit in the engine. A value
that does not fit the parameter it is bound for raises ``ValueError``
before the call reaches the engine, where it would otherwise be
truncated modulo 2**32 into a different, valid-looking coordinate.

The engine is a WebAssembly module (``formulon_capi.wasm``) run by
``wasmtime``; see :mod:`formulon._c` for the low-level memory and call
plumbing.
"""

from __future__ import annotations

import datetime
import enum
import math
import numbers
import operator
import os
import struct
from dataclasses import dataclass, field
from enum import IntEnum, IntFlag
from functools import wraps
from inspect import isgeneratorfunction
from typing import (
    Any,
    BinaryIO,
    Dict,
    Iterator,
    List,
    Literal,
    Mapping,
    Optional,
    Sequence,
    Tuple,
    TypedDict,
    Union,
    cast,
    overload,
)
from xml.etree import ElementTree

from . import _structs as S
from ._c import LIB, FormulaSyntaxError, FormulonError, InvalidFileError, ValueKind, _sint, _uint, fm_value_t_size
from ._status import StatusCode

__all__ = [
    "Alignment",
    "AlignmentPatch",
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
    "BorderPatch",
    "BorderSide",
    "BorderSidePatch",
    "BorderStyle",
    "CalcMode",
    "canonicalize_formula",
    "canonicalize_function_name",
    "CELL_STYLE_BUILTIN_ID_NONE",
    "CellAddress",
    "CellFormat",
    "CellNode",
    "CellProtection",
    "CellProtectionPatch",
    "CellRect",
    "CellsInRange",
    "CellStyle",
    "CellXf",
    "CellXfInput",
    "CfCell",
    "CfDataBarAxisPosition",
    "CfDataBarDirection",
    "CfIcon",
    "CfMatch",
    "CfMatchKind",
    "CfRuleType",
    "CfValueObjectType",
    "Cfvo",
    "ChartAxis",
    "ChartAxisCrosses",
    "ChartAxisKind",
    "ChartAxisPatch",
    "ChartAxisPosition",
    "ChartAxisTitleText",
    "ChartBarDirection",
    "BlanksAs",
    "ChartCategories",
    "ChartData",
    "ChartDataSeries",
    "ChartDataSource",
    "ChartDisplayPatch",
    "ChartDomain",
    "ChartExtendedLayout",
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
    "ChartProperties",
    "ChartRole",
    "ChartSeries",
    "ChartSeriesDirection",
    "ChartSeriesPatch",
    "ChartSeriesRanges",
    "ChartSource",
    "ChartSourceKind",
    "ChartTickLabelPosition",
    "ChartTitle",
    "ChartTitleKind",
    "ChartUnresolved",
    "CollationClasses",
    "CollationStrength",
    "Color",
    "ColorContext",
    "ColorInput",
    "ColorKind",
    "ColorResolution",
    "ColorScale",
    "ColumnBreak",
    "ColumnLayout",
    "Comment",
    "compare_text",
    "ConditionalFormat",
    "ConditionalFormatInput",
    "DataBar",
    "DataTable",
    "DataTableInput",
    "DataTableInputCell",
    "DataValidation",
    "DataValidationInput",
    "DatePeriod",
    "DateTimeGrouping",
    "DefinedName",
    "DefinedNameInput",
    "DisplayStatus",
    "DisplayText",
    "DrawingAnchor",
    "DrawingObjectKind",
    "DrawingProperties",
    "Dxf",
    "DxfInput",
    "DynamicFilterType",
    "EffectiveStyle",
    "EffectiveStyleSource",
    "enum_name",
    "ErrorCode",
    "ErrorName",
    "ExcelError",
    "ExternalLink",
    "ExternalLinkKind",
    "Fill",
    "FillPatch",
    "FillPattern",
    "FilterKind",
    "ComparisonOperator",
    "Font",
    "FontFamily",
    "FontPatch",
    "FontScheme",
    "FontVerticalAlign",
    "FunctionAvailability",
    "FunctionMetadata",
    "GeometryMode",
    "get_collation_classes",
    "get_function_metadata",
    "get_image_info",
    "get_locale_facts",
    "GradientFill",
    "GradientStop",
    "GradientType",
    "HeaderFooter",
    "HorizontalAlign",
    "Hyperlink",
    "HyperlinkInput",
    "IconSet",
    "IconSetName",
    "ImageFormat",
    "ImageInfo",
    "InvalidFileError",
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
    "NativeValue",
    "PageOrder",
    "PaperSize",
    "PhoneticAlignment",
    "PhoneticType",
    "PictureInput",
    "PivotStyleOptions",
    "RowBreak",
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
    "PhoneticRun",
    "PictureProperties",
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
    "PivotFilterKind",
    "PivotGrandTotals",
    "PivotInput",
    "PivotItem",
    "PivotLayout",
    "PivotOptions",
    "PivotProperties",
    "PivotReportLayout",
    "PivotShowValuesAs",
    "PivotSortOrder",
    "PivotWorksheetSource",
    "PrintOptions",
    "PrintTitles",
    "ProfileId",
    "ProtectedRange",
    "RangeCell",
    "ReadDiagnostics",
    "ReadingOrder",
    "RectPt",
    "RefMode",
    "ResolvedColor",
    "RowLayout",
    "SaveDiagnostics",
    "ShapeInput",
    "ShapePaint",
    "ShapeProperties",
    "SheetComment",
    "SheetFormatDefaults",
    "SheetProtection",
    "SheetProtectionInput",
    "SheetRect",
    "SheetRef",
    "SheetSelection",
    "SheetView",
    "SheetViewMode",
    "SheetVisibility",
    "SlicerCrossFilter",
    "SlicerInput",
    "SlicerProperties",
    "SlicerSortOrder",
    "SortBy",
    "SortMethod",
    "Span",
    "Sparkline",
    "SparklineAxisType",
    "SparklineGroup",
    "SparklineType",
    "Spill",
    "TableColumn",
    "TableInput",
    "TableProperties",
    "TableStyle",
    "TableStyleDefaults",
    "TableStyleElement",
    "TableStyleElementType",
    "TableStyleOptions",
    "TableTotalsRowFunction",
    "ThemeColors",
    "ThemeFonts",
    "ThemeScriptFont",
    "ThemeSource",
    "ThreadedComment",
    "ThreadedCommentInput",
    "to_native",
    "UNCHANGED",
    "Unchanged",
    "UnderlineStyle",
    "ValidationErrorStyle",
    "ValidationOutcome",
    "ValidationType",
    "Value",
    "VerticalAlign",
    "WidthModel",
    "WidthPlatform",
    "Workbook",
    "WorkbookFormat",
    "WorkbookProtection",
]


# ---------------------------------------------------------------------------
# Error type
# ---------------------------------------------------------------------------

# `formulon::FormulonErrorCode` ordinals. The C ABI intentionally exposes
# status codes as integers, so bindings retain these matching stable values.
_STATUS_NOT_FOUND = 6

# `fm_cell_range_next_cursor` reports UINT64_MAX when no further cell exists.
_CELL_RANGE_NO_CURSOR = (1 << 64) - 1


class _UnchangedType(enum.Enum):
    """The type of :data:`UNCHANGED`; an enum so copies and pickles keep the one instance."""

    UNCHANGED = "UNCHANGED"

    def __repr__(self) -> str:
        return "formulon.UNCHANGED"


#: Default of a partial-update keyword argument: the field keeps its current value.
UNCHANGED: Literal[_UnchangedType.UNCHANGED] = _UnchangedType.UNCHANGED

#: Annotation type of :data:`UNCHANGED`; write a partial-update argument as ``Union[T, Unchanged]``.
Unchanged = Literal[_UnchangedType.UNCHANGED]


#: A sheet argument: a :class:`Sheet` of the same workbook, a 0-based sheet index, or a sheet name.
SheetRef = Union["Sheet", int, str]

#: Excel profile id accepted by `Workbook.set_profile_id`.
ProfileId = Literal[
    "mac-365-ja_JP",
    "win-365-ja_JP",
    "mac-365-en_US",
    "win-365-en_US",
    "mac-365-de_DE",
    "win-365-de_DE",
    "mac-365-fr_FR",
    "win-365-fr_FR",
    "mac-365-zh_CN",
    "win-365-zh_CN",
    "mac-365-ko_KR",
    "win-365-ko_KR",
    "mac-365-th_TH",
    "win-365-th_TH",
    "mac-365-es_ES",
    "win-365-es_ES",
    "mac-365-es_MX",
    "win-365-es_MX",
    "mac-365-pt_BR",
    "win-365-pt_BR",
    "mac-365-ru_RU",
    "win-365-ru_RU",
    "mac-365-zh_TW",
    "win-365-zh_TW",
    "mac-365-it_IT",
    "win-365-it_IT",
    "mac-365-nl_NL",
    "win-365-nl_NL",
]

# `fm_date_order_t` ordinals, indexed by the C value.

# `FormulonError` itself lives in `formulon._c` (imported above): its
# constructor needs to run from inside that module's WASM call trampoline,
# on a `wasmtime.Trap`, which this module cannot depend on without a cycle.


@dataclass(frozen=True)
class SaveDiagnostics:
    """Bytes and loss counters returned by ``save_with_diagnostics``.

    A field means the same thing whichever container was written; where
    the underlying event exists in only one container, the other simply
    never raises it. Coverage is partial by design -- an all-zero result
    means none of these losses occurred, not that nothing was logged.

    Summing the fields does **not** give "how many things did I lose":
    ``dropped_part_count`` and ``dropped_relationship_count`` both rise
    for one dropped part.
    """

    bytes: bytes
    #: Formula cells emitted as cached literals. Zero for XLSX.
    downgraded_formula_count: int
    #: Sheet features not lowered to records. Zero for XLSX.
    deferred_feature_count: int
    #: Passthrough parts dropped for a path collision. Both containers.
    dropped_part_count: int
    #: Relationships dropped because their target part is gone. Both
    #: containers. Rises with ``dropped_part_count`` for one dropped part.
    dropped_relationship_count: int
    #: Tables emitted under a writer-assigned id. Zero for XLSB.
    renumbered_part_count: int


@dataclass(frozen=True)
class ReadDiagnostics:
    """Loss and recovery counters captured while loading a workbook.

    Same contract as :class:`SaveDiagnostics`: one meaning per field
    across both containers, partial coverage by design.
    """

    #: Formula cells whose stored formula could not be decoded. XLSB only.
    undecoded_formula_count: int
    #: Defined names skipped for the same reason. XLSB only.
    undecoded_defined_name_count: int
    #: Package parts whose content type could not be resolved. XLSB only;
    #: the XLSX reader keeps unmodelled parts as passthrough.
    undecoded_part_count: int
    #: Presentation-overlay entries dropped for an unusable reference.
    #: XLSX only.
    skipped_feature_count: int
    #: Workbook parts with an unrecognised content type. Max 1. XLSX only.
    unknown_content_type_count: int


def _check(status: int) -> None:
    """Raise :class:`FormulonError` (or its subclass for ``status``) if ``status`` is non-zero."""
    if status != 0:
        raise FormulonError(status)


# `_uint` / `_sint` live in `formulon._c` (imported above) alongside
# `FormulonError`: they validate a value right at the WASM call boundary
# that module owns, and `_c.py`'s own `fm_get_status_name` call needs one too.

# ---------------------------------------------------------------------------
# Value POD
# ---------------------------------------------------------------------------


# The one payload field each kind carries; BLANK, REF and LAMBDA carry none.
_VALUE_PAYLOAD_FIELD = {
    ValueKind.NUMBER: "number",
    ValueKind.BOOL: "boolean",
    ValueKind.TEXT: "text",
    ValueKind.ERROR: "error_code",
    ValueKind.ARRAY: "array",
}


@dataclass(frozen=True)
class ExcelError:
    """An Excel cell error such as ``#DIV/0!``.

    Native reads return one for an error value, and :meth:`Sheet.set_value`
    stores one as a cell error. Its string form is the en-US display literal.
    """

    code: ErrorCode

    def __post_init__(self) -> None:
        if isinstance(self.code, bool):
            raise TypeError("formulon: ExcelError.code must be an ErrorCode, got bool")
        object.__setattr__(self, "code", ErrorCode(_sint(self.code, "code")))

    def __str__(self) -> str:
        return self.code.display_name


#: A value as plain Python data: ``None`` for blank, ``float``, ``bool``,
#: ``str``, :class:`ExcelError`, or a row-major list of rows of them for an
#: array result.
NativeValue = Union[None, float, bool, str, ExcelError, List[List["NativeValue"]]]


@dataclass(frozen=True)
class Value:
    """A cell or formula value tagged with its kind (the ``*_typed_*`` reads return it).

    Exactly one of ``number`` / ``boolean`` / ``text`` / ``error_code`` /
    ``array`` is populated according to ``kind``; ``array`` holds the rows of
    an array result, each a tuple of scalar values of equal length. The
    constructor enforces the invariant: ``kind`` is coerced to
    :class:`ValueKind`, a missing or extra payload raises ``ValueError`` and a
    payload of the wrong type raises ``TypeError``.
    """

    kind: ValueKind
    number: Optional[float] = None
    boolean: Optional[bool] = None
    text: Optional[str] = None
    error_code: Optional[ErrorCode] = None
    array: Optional[Tuple[Tuple["Value", ...], ...]] = None

    def __post_init__(self) -> None:
        kind = self.kind
        if not isinstance(kind, ValueKind):
            kind = ValueKind(_sint(kind, "kind"))
            object.__setattr__(self, "kind", kind)
        required = _VALUE_PAYLOAD_FIELD.get(kind)
        for name in _VALUE_PAYLOAD_FIELD.values():
            if name != required and getattr(self, name) is not None:
                raise ValueError(f"formulon: Value(kind={kind.name}) takes no {name}")
        if required is None:
            return
        payload = getattr(self, required)
        if payload is None:
            raise ValueError(f"formulon: Value(kind={kind.name}) requires {required}")
        if kind == ValueKind.NUMBER:
            if isinstance(payload, bool) or not isinstance(payload, numbers.Real):
                raise TypeError(f"formulon: Value.number must be a real number, got {type(payload).__name__}")
            object.__setattr__(self, "number", float(payload))
        elif kind == ValueKind.BOOL:
            if not isinstance(payload, bool):
                raise TypeError(f"formulon: Value.boolean must be a bool, got {type(payload).__name__}")
        elif kind == ValueKind.TEXT:
            if not isinstance(payload, str):
                raise TypeError(f"formulon: Value.text must be a str, got {type(payload).__name__}")
        elif kind == ValueKind.ERROR:
            object.__setattr__(self, "error_code", ErrorCode(_sint(payload, "error_code")))
        else:
            object.__setattr__(self, "array", _array_payload(payload))

    def __repr__(self) -> str:
        kind = self.kind
        field = _VALUE_PAYLOAD_FIELD.get(kind)
        if field is None:
            return f"Value({kind.name})"
        payload = getattr(self, field)
        if kind == ValueKind.ERROR:
            shown = payload.name
        elif kind == ValueKind.ARRAY:
            shown = f"{len(payload)}x{len(payload[0])}"
        elif kind == ValueKind.TEXT:
            shown = repr(payload if len(payload) <= 32 else payload[:32] + "...")
        else:
            shown = repr(payload)
        return f"Value({kind.name}, {shown})"

    @classmethod
    def _from_wasm(cls, ptr: int) -> "Value":
        """Decode the value struct at engine-memory offset ``ptr``.

        Layout: int32 kind | 4 pad | union (8 bytes). ARRAY / REF / LAMBDA
        carry no payload.
        """
        buf = LIB.read_bytes(ptr, fm_value_t_size)
        kind_int = struct.unpack_from("<i", buf, 0)[0]
        kind = ValueKind(kind_int)
        if kind is ValueKind.NUMBER:
            number = struct.unpack_from("<d", buf, 8)[0]
            return cls(kind=kind, number=number)
        if kind is ValueKind.BOOL:
            boolean = struct.unpack_from("<i", buf, 8)[0]
            return cls(kind=kind, boolean=bool(boolean))
        if kind is ValueKind.TEXT:
            text_ptr = struct.unpack_from("<I", buf, 8)[0]
            return cls(kind=kind, text=LIB.read_cstr(text_ptr))
        if kind is ValueKind.ERROR:
            error_code = struct.unpack_from("<i", buf, 8)[0]
            return cls(kind=kind, error_code=ErrorCode(error_code))
        # BLANK / REF / LAMBDA: no payload. A value slot never holds an array.
        return cls(kind=kind)


def to_native(value: Value) -> NativeValue:
    """Convert a tagged :class:`Value` to plain Python data (:data:`NativeValue`).

    - ``BLANK`` -> ``None``
    - ``NUMBER`` -> ``float``
    - ``BOOL`` -> ``bool``
    - ``TEXT`` -> ``str``
    - ``ERROR`` -> :class:`ExcelError` (Excel cell errors are values,
      not Python exceptions)
    - ``ARRAY`` -> a list of rows, each a list of the cells' native values
    """
    kind = value.kind
    if kind is ValueKind.BLANK:
        return None
    if kind is ValueKind.NUMBER:
        return value.number
    if kind is ValueKind.BOOL:
        return value.boolean
    if kind is ValueKind.TEXT:
        return value.text
    if kind is ValueKind.ERROR:
        return ExcelError(cast(ErrorCode, value.error_code))
    assert kind is ValueKind.ARRAY, f"a {kind.name} value never reaches a public read"
    return [[to_native(cell) for cell in row] for row in cast(Tuple[Tuple[Value, ...], ...], value.array)]


def _array_payload(rows: object) -> Tuple[Tuple[Value, ...], ...]:
    """``rows`` as a non-empty rectangle of scalar :class:`Value` cells, rows and cells as tuples."""
    if isinstance(rows, (str, bytes)) or not isinstance(rows, (list, tuple)):
        raise TypeError(f"formulon: Value.array must be a sequence of rows, got {type(rows).__name__}")
    grid = tuple(tuple(row) if isinstance(row, (list, tuple)) else row for row in rows)
    if not grid or not all(isinstance(row, tuple) and row for row in grid):
        raise ValueError("formulon: Value.array must hold at least one non-empty row")
    if any(len(row) != len(grid[0]) for row in grid):
        raise ValueError("formulon: Value.array rows must all have the same length")
    for row in grid:
        for cell in row:
            if not isinstance(cell, Value) or cell.kind is ValueKind.ARRAY:
                raise TypeError(f"formulon: Value.array cells must be scalar Values, got {cell!r}")
    return grid


def _folded(grid: List[List[Value]]) -> Value:
    """A 1x1 result as its one cell, any other shape as an ``ARRAY`` value."""
    if len(grid) == 1 and len(grid[0]) == 1:
        return grid[0][0]
    return Value(kind=ValueKind.ARRAY, array=tuple(tuple(row) for row in grid))


# ---------------------------------------------------------------------------
# Iteration tuples
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class RangeCell:
    """One written coordinate, as :meth:`Sheet.list_cells` returns it.

    ``value`` is the computed value as native data and ``kind`` its
    :class:`ValueKind`; ``formula`` is ``None`` for a constant.
    """

    row: int
    col: int
    formula: Optional[str]
    value: NativeValue
    kind: ValueKind


@dataclass(frozen=True)
class DefinedName:
    """One defined name, as ``Workbook.list_defined_names()`` and ``Workbook.get_defined_name()`` return it."""

    name: str
    #: The formula text, with a leading ``=``.
    formula: str
    #: The 0-based index of the scope sheet; ``None`` for workbook scope.
    local_sheet_id: Optional[int]
    #: The note Excel shows in the Name Manager; ``None`` when none.
    comment: Optional[str] = None
    #: Whether the name is hidden from the Name Manager.
    hidden: bool = False


@dataclass
class DefinedNameInput:
    """A defined name :meth:`Workbook.set_defined_name` writes.

    ``formula`` takes the text with or without a leading ``=``; ``scope`` is
    the sheet the name is local to, ``None`` for workbook scope.
    """

    name: str
    formula: str
    scope: Optional[SheetRef] = None
    comment: Optional[str] = None
    hidden: bool = False


class TableTotalsRowFunction(IntEnum):
    """The totals-row function of a table column, in ``ST_TotalsRowFunction`` order."""

    NONE = 0
    SUM = 1
    MIN = 2
    MAX = 3
    AVERAGE = 4
    COUNT = 5
    COUNT_NUMBERS = 6
    STD_DEV = 7
    VAR = 8
    CUSTOM = 9


@dataclass(frozen=True)
class TableColumn:
    """One column of :attr:`TableProperties.columns`.

    ``totals_row_function`` is the column's :class:`TableTotalsRowFunction`
    (``NONE`` when absent); ``totals_row_formula`` is the ``custom``
    function's formula.
    """

    name: str
    totals_row_function: TableTotalsRowFunction = TableTotalsRowFunction.NONE
    totals_row_label: str = ""
    totals_row_formula: str = ""

    def __post_init__(self) -> None:
        object.__setattr__(self, "totals_row_function", _enum(TableTotalsRowFunction, self.totals_row_function))


@dataclass(frozen=True)
class TableProperties:
    """One table, as :meth:`Table.get_properties` returns it.

    ``name`` is the one table name, the name formulas and :meth:`Workbook.table`
    use. ``header_row`` and ``totals_row`` say whether the table has those
    rows. The style is read through :meth:`Table.get_style_options`.
    """

    name: str
    #: The table's sheet and range.
    range: SheetRect
    header_row: bool = True
    totals_row: bool = False
    columns: List[TableColumn] = field(default_factory=list)


@dataclass(frozen=True)
class TableStyleOptions:
    """The style element of a table (``<tableStyleInfo>``), as :meth:`Table.get_style_options` returns it.

    ``style_name`` is ``None`` for an element without a named style and may
    name a custom style (:meth:`Workbook.get_table_style`).
    :meth:`Table.update_style_options` takes these fields as keyword arguments.
    """

    style_name: Optional[str] = None
    show_first_column: bool = False
    show_last_column: bool = False
    show_row_stripes: bool = False
    show_column_stripes: bool = False


@dataclass
class TableInput:
    """A table :meth:`Sheet.add_table` creates.

    ``range`` is a :class:`CellRect` or an unqualified A1 range on the sheet.
    Omitted ``columns`` take their names from the header cells, an empty one
    named as Excel does (``Column1``, ...). Given, ``columns`` holds one
    non-empty, table-unique name per column of the range: a non-empty header
    cell must already hold that name, and an empty one receives it.
    ``style_name`` ``None`` applies the workbook's table-style default
    (:meth:`Workbook.get_table_style_defaults`), else ``TableStyleMedium2``;
    ``""`` creates the table without a style element. A styled table starts
    with row stripes on and the other toggles off; a ``show_*`` that is not
    ``None`` overrides that.
    """

    range: Union[CellRect, str]
    name: str
    columns: Optional[Sequence[str]] = None
    style_name: Optional[str] = None
    show_first_column: Optional[bool] = None
    show_last_column: Optional[bool] = None
    show_row_stripes: Optional[bool] = None
    show_column_stripes: Optional[bool] = None
    header_row: bool = True
    totals_row: bool = False


@dataclass(frozen=True)
class TableStyleElement:
    """One region of a custom table style.

    ``type`` is the :class:`TableStyleElementType` region; ``size`` is the
    stripe band size; ``dxf_index`` indexes the differential formats,
    ``None`` for a region listed without formatting.
    """

    type: TableStyleElementType
    size: int = 1
    dxf_index: Optional[int] = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "type", _enum(TableStyleElementType, self.type))


@dataclass(frozen=True)
class TableStyle:
    """A custom table / PivotTable style from ``<tableStyles>``."""

    name: str
    pivot: bool = True
    table: bool = True
    elements: List[TableStyleElement] = field(default_factory=list)


@dataclass(frozen=True)
class PassthroughPart:
    """One row of ``WorkbookOoxml.list_passthroughs()``."""

    path: str


# ---------------------------------------------------------------------------
# Enumerations (mirror the C ABI ordinals)
# ---------------------------------------------------------------------------


class TableStyleElementType(IntEnum):
    """The region a :class:`TableStyleElement` formats, in ``ST_TableStyleType`` order."""

    WHOLE_TABLE = 0
    HEADER_ROW = 1
    TOTAL_ROW = 2
    FIRST_COLUMN = 3
    LAST_COLUMN = 4
    FIRST_ROW_STRIPE = 5
    SECOND_ROW_STRIPE = 6
    FIRST_COLUMN_STRIPE = 7
    SECOND_COLUMN_STRIPE = 8
    FIRST_HEADER_CELL = 9
    LAST_HEADER_CELL = 10
    FIRST_TOTAL_CELL = 11
    LAST_TOTAL_CELL = 12
    FIRST_SUBTOTAL_COLUMN = 13
    SECOND_SUBTOTAL_COLUMN = 14
    THIRD_SUBTOTAL_COLUMN = 15
    FIRST_SUBTOTAL_ROW = 16
    SECOND_SUBTOTAL_ROW = 17
    THIRD_SUBTOTAL_ROW = 18
    BLANK_ROW = 19
    FIRST_COLUMN_SUBHEADING = 20
    SECOND_COLUMN_SUBHEADING = 21
    THIRD_COLUMN_SUBHEADING = 22
    FIRST_ROW_SUBHEADING = 23
    SECOND_ROW_SUBHEADING = 24
    THIRD_ROW_SUBHEADING = 25
    PAGE_FIELD_LABELS = 26
    PAGE_FIELD_VALUES = 27


class PhoneticType(IntEnum):
    """Kana form a phonetic guide generates readings in (``<phoneticPr type>``).

    Zero is the state of an absent ``<phoneticPr>``, not the schema default.
    """

    HALFWIDTH_KATAKANA = 0
    FULLWIDTH_KATAKANA = 1
    HIRAGANA = 2
    NO_CONVERSION = 3


class PhoneticAlignment(IntEnum):
    """How a phonetic guide's kana is distributed over the characters it covers (``<phoneticPr alignment>``).

    Zero is the state of an absent ``<phoneticPr>``, not the schema default.
    """

    NO_CONTROL = 0
    LEFT = 1
    CENTER = 2
    DISTRIBUTED = 3


class PageOrder(IntEnum):
    """The order a worksheet's pages are numbered in (``ST_PageOrder``)."""

    DOWN_THEN_OVER = 0
    OVER_THEN_DOWN = 1


class PaperSize:
    """Named OOXML ``paperSize`` codes the engine knows.

    Not an enum: the code set is open, so ``PageSetup.paper_size`` stays an
    ``int`` and an unknown code reads back unchanged.
    """

    LETTER = 1
    LEGAL = 5
    A3 = 8
    A4 = 9
    A5 = 11
    B4 = 12
    B5 = 13
    A6 = 70


class ChartTickLabelPosition(IntEnum):
    """Where an axis draws its tick labels (``ST_TickLblPos``); ``NONE`` draws no labels.

    An axis without the element reads ``None``.
    """

    NEXT_TO = 0
    HIGH = 1
    LOW = 2
    NONE = 3


class ChartDomain(IntEnum):
    """What a chart's category axis carries: category labels or X values."""

    CATEGORY = 0
    XY = 1


class WidthPlatform(IntEnum):
    """The platform whose metrics a :class:`WidthModel` uses."""

    WIN = 0
    MAC = 1


class GeometryMode(IntEnum):
    """Which Windows column-width figure a geometry call uses."""

    DISPLAY = 0
    PRINT = 1


class CollationStrength(IntEnum):
    """Text fold selected for a collation call.

    The strength selects the text fold only; type boundaries never depend on it.
    """

    #: The fold of ``=``, criteria, SEARCH and approximate lookup.
    COMPARE = 0
    #: The fold of SORT, UNIQUE, XLOOKUP, GROUPBY and pivots.
    DYNAMIC = 1
    #: The fold of MATCH 0 and VLOOKUP FALSE.
    CLASSIC_EXACT = 2


@dataclass(frozen=True)
class CollationClasses:
    """Result of :func:`get_collation_classes`.

    ``class_ids`` holds one id per input value, dense from 0 in ascending
    SORT order; ``class_count`` is the number of distinct classes.
    """

    class_ids: List[int]
    class_count: int


class DisplayStatus(IntEnum):
    """Outcome of a display-text rendering."""

    OK = 0
    OVERFLOW = 1
    INVALID_FORMAT = 2


class FilterKind(IntEnum):
    """Criterion of an AutoFilter column."""

    NONE = 0
    VALUES = 1
    CUSTOM = 2
    TOP10 = 3
    DYNAMIC = 4
    COLOR = 5
    ICON = 6


class ComparisonOperator(IntEnum):
    """Comparison of a ``CELL_IS`` conditional-format rule, a data-validation rule or an AutoFilter condition.

    The order is that of the ``ST_DataValidationOperator`` schema type. An
    AutoFilter condition takes only the six non-range members.
    """

    BETWEEN = 0
    NOT_BETWEEN = 1
    EQUAL = 2
    NOT_EQUAL = 3
    LESS_THAN = 4
    LESS_THAN_OR_EQUAL = 5
    GREATER_THAN = 6
    GREATER_THAN_OR_EQUAL = 7


class DynamicFilterType(IntEnum):
    """Dynamic AutoFilter type, in `ST_DynamicFilterType` schema order."""

    NULL = 0
    ABOVE_AVERAGE = 1
    BELOW_AVERAGE = 2
    TOMORROW = 3
    TODAY = 4
    YESTERDAY = 5
    NEXT_WEEK = 6
    THIS_WEEK = 7
    LAST_WEEK = 8
    NEXT_MONTH = 9
    THIS_MONTH = 10
    LAST_MONTH = 11
    NEXT_QUARTER = 12
    THIS_QUARTER = 13
    LAST_QUARTER = 14
    NEXT_YEAR = 15
    THIS_YEAR = 16
    LAST_YEAR = 17
    YEAR_TO_DATE = 18
    Q1 = 19
    Q2 = 20
    Q3 = 21
    Q4 = 22
    M1 = 23
    M2 = 24
    M3 = 25
    M4 = 26
    M5 = 27
    M6 = 28
    M7 = 29
    M8 = 30
    M9 = 31
    M10 = 32
    M11 = 33
    M12 = 34


class SortBy(IntEnum):
    """What a sort condition sorts by."""

    VALUE = 0
    CELL_COLOR = 1
    FONT_COLOR = 2
    ICON = 3


class SortMethod(IntEnum):
    """Sort method of an AutoFilter sort state."""

    NONE = 0
    PIN_YIN = 1
    STROKE = 2


class DateTimeGrouping(IntEnum):
    """Granularity of an AutoFilter date-group item."""

    YEAR = 0
    MONTH = 1
    DAY = 2
    HOUR = 3
    MINUTE = 4
    SECOND = 5


class ValidationErrorStyle(IntEnum):
    """Data-validation error style."""

    STOP = 0
    WARNING = 1
    INFORMATION = 2


class ValidationType(IntEnum):
    """What a data-validation rule checks."""

    NONE = 0
    WHOLE = 1
    DECIMAL = 2
    LIST = 3
    DATE = 4
    TIME = 5
    TEXT_LENGTH = 6
    CUSTOM = 7


class CfRuleType(IntEnum):
    """The kind of a conditional-format rule."""

    EXPRESSION = 0
    CELL_IS = 1
    COLOR_SCALE = 2
    DATA_BAR = 3
    ICON_SET = 4
    TOP10 = 5
    ABOVE_AVERAGE = 6
    CONTAINS_TEXT = 7
    NOT_CONTAINS_TEXT = 8
    BEGINS_WITH = 9
    ENDS_WITH = 10
    CONTAINS_BLANKS = 11
    NOT_CONTAINS_BLANKS = 12
    CONTAINS_ERRORS = 13
    NOT_CONTAINS_ERRORS = 14
    TIME_PERIOD = 15
    DUPLICATE_VALUES = 16
    UNIQUE_VALUES = 17


class DatePeriod(IntEnum):
    """The period of a ``TIME_PERIOD`` conditional-format rule, in ``ST_TimePeriod`` schema order."""

    TODAY = 0
    YESTERDAY = 1
    TOMORROW = 2
    LAST_7_DAYS = 3
    THIS_MONTH = 4
    LAST_MONTH = 5
    NEXT_MONTH = 6
    THIS_WEEK = 7
    LAST_WEEK = 8
    NEXT_WEEK = 9


class CfValueObjectType(IntEnum):
    """How a conditional-format threshold reads its value."""

    NUMBER = 0
    PERCENT = 1
    PERCENTILE = 2
    MIN = 3
    MAX = 4
    FORMULA = 5
    AUTO_MIN = 6
    AUTO_MAX = 7


class IconSetName(IntEnum):
    """An icon set; the digit is the level count."""

    THREE_ARROWS = 0
    THREE_ARROWS_GRAY = 1
    THREE_FLAGS = 2
    THREE_TRAFFIC_LIGHTS_1 = 3
    THREE_TRAFFIC_LIGHTS_2 = 4
    THREE_SIGNS = 5
    THREE_SYMBOLS = 6
    THREE_SYMBOLS_2 = 7
    FOUR_ARROWS = 8
    FOUR_ARROWS_GRAY = 9
    FOUR_RED_TO_BLACK = 10
    FOUR_RATING = 11
    FOUR_TRAFFIC_LIGHTS = 12
    FIVE_ARROWS = 13
    FIVE_ARROWS_GRAY = 14
    FIVE_RATING = 15
    FIVE_QUARTERS = 16
    #: The sets below exist only as levels of a custom icon set (:class:`CfIcon`).
    THREE_STARS = 17
    THREE_TRIANGLES = 18
    FIVE_BOXES = 19
    NO_ICONS = 20


def enum_name(enum_type: "type[enum.Enum]", value: int) -> Optional[str]:
    """The member name ``value`` has in ``enum_type``, or ``None`` when no member has it.

    Example::

        enum_name(CfRuleType, 3)  # "DATA_BAR"
    """
    if not (isinstance(enum_type, type) and issubclass(enum_type, enum.Enum)):
        raise TypeError(f"formulon: enum_type must be an enum class, got {enum_type!r}")
    try:
        return enum_type(value).name
    except ValueError:
        return None


class ColorContext(IntEnum):
    """Where a colour is used; decides what auto means."""

    FONT = 0
    FILL_FOREGROUND = 1
    FILL_BACKGROUND = 2
    BORDER = 3


class ImageFormat(IntEnum):
    """Raster format of an image."""

    UNKNOWN = 0
    PNG = 1
    JPEG = 2
    GIF = 3
    BMP = 4


class DrawingObjectKind(IntEnum):
    """What a drawing anchor holds."""

    PICTURE = 0
    SHAPE = 1
    CHART = 2
    GROUP = 3
    CONNECTOR = 4
    GRAPHIC_FRAME = 5
    OTHER = 6


class AnchorKind(IntEnum):
    """Anchor element type of a drawing object."""

    ONE_CELL = 0
    TWO_CELL = 1
    ABSOLUTE = 2


class AnchorEditAs(IntEnum):
    """How a drawing object follows row and column edits."""

    ONE_CELL = 0
    TWO_CELL = 1
    ABSOLUTE = 2


@dataclass
class DrawingAnchor:
    """The placement of a picture or chart frame (``set_anchor``, ``add_picture``, ``add_chart``).

    ``row`` / ``col`` / ``row_offset_emu`` / ``col_offset_emu`` are the top-left marker of the anchor
    box; ``width_emu`` / ``height_emu`` the object's own unrotated size. A zero size is the picture's
    pixel size times 9525 on insert, Excel's default frame for a chart, and keeps the current size on
    ``set_anchor``. ``edit_as`` applies to a two-cell anchor only. ``DrawingAnchor()`` is a one-cell
    anchor at A1.
    """

    anchor_kind: AnchorKind = AnchorKind.ONE_CELL
    edit_as: AnchorEditAs = AnchorEditAs.TWO_CELL
    row: int = 0
    col: int = 0
    row_offset_emu: int = 0
    col_offset_emu: int = 0
    width_emu: int = 0
    height_emu: int = 0


class ShapePaint(IntEnum):
    """How a shape's fill or outline is specified; a shape without the element reads ``None``."""

    SOLID = 0
    NONE = 1
    OTHER = 2


class SlicerSortOrder(IntEnum):
    """Item sort order of a table slicer."""

    ASCENDING = 0
    DESCENDING = 1


class SlicerCrossFilter(IntEnum):
    """How a table slicer shows items the other filters exclude."""

    SHOW_ITEMS_WITH_DATA_AT_TOP = 0
    SHOW_ITEMS_WITH_NO_DATA = 1
    NONE = 2


class SparklineType(IntEnum):
    """Sparkline group type."""

    LINE = 0
    COLUMN = 1
    STACKED = 2


class SparklineAxisType(IntEnum):
    """Sparkline axis scaling."""

    INDIVIDUAL = 0
    GROUP = 1
    CUSTOM = 2


class ColorResolution(IntEnum):
    """How a colour was resolved."""

    EXACT = 0
    DEFAULT_THEME = 1
    INDEX_OUT_OF_RANGE = 2
    THEME_UNPARSEABLE = 3
    AUTO_CONTEXT = 4


class ThemeSource(IntEnum):
    """Where the reported theme colours came from."""

    PART = 0
    DEFAULT = 1
    UNPARSEABLE = 2


class EffectiveStyleSource(IntEnum):
    """Which level supplied a cell's effective xf."""

    CELL = 0
    ROW = 1
    COLUMN = 2
    DEFAULT = 3


class CalcMode(IntEnum):
    """Workbook-level calc mode (``<calcPr calcMode>``).

    - ``AUTO`` (the default): a value read and a save first calculate the
      dirty cells.
    - ``MANUAL``: nothing is calculated until :meth:`Workbook.recalc`.
    - ``AUTO_NO_TABLE``: as ``AUTO``, but the What-If data tables wait for
      :meth:`Workbook.recalc_data_tables`.

    A volatile function (``NOW()``, ``RAND()``) is evaluated again when a change
    or :meth:`Workbook.recalc` recalculates it, not on every read.
    """

    AUTO = 0
    MANUAL = 1
    AUTO_NO_TABLE = 2


class RefMode(IntEnum):
    """Reference notation formulas are shown in (``<calcPr refMode>``)."""

    A1 = 0
    R1C1 = 1


class SheetViewMode(IntEnum):
    """The sheet's view: normal, page break preview or page layout."""

    NORMAL = 0
    PAGE_BREAK_PREVIEW = 1
    PAGE_LAYOUT = 2


class PageOrientation(IntEnum):
    """The printed page's orientation."""

    PORTRAIT = 0
    LANDSCAPE = 1


class SheetVisibility(IntEnum):
    """Sheet tab visibility, mirroring OOXML ``<sheet state>``.

    ``VERY_HIDDEN`` differs from ``HIDDEN`` in that Excel leaves such a
    sheet out of its "Unhide" dialog, which is how a workbook keeps a
    settings or lookup sheet out of a user's reach.
    """

    VISIBLE = 0
    HIDDEN = 1
    VERY_HIDDEN = 2


class LogLevel(IntEnum):
    """Minimum severity for the engine's structured log stream.

    ``OFF`` discards every record and is the default: an embedded library
    must not write to the host's stderr unless the host asks it to.
    """

    DEBUG = 0
    INFO = 1
    WARN = 2
    ERROR = 3
    OFF = 4


class ErrorCode(IntEnum):
    """Excel cell-error ordinals (mirror of ``formulon::ErrorCode``).

    These are the values carried by :attr:`Value.error_code` and accepted
    by :meth:`Sheet.set_error`.
    """

    NULL = 0
    DIV0 = 1
    VALUE = 2
    REF = 3
    NAME = 4
    NUM = 5
    NA = 6
    GETTING_DATA = 7
    SPILL = 8
    CALC = 9
    FIELD = 10
    BLOCKED = 11
    CONNECT = 12
    EXTERNAL = 13
    BUSY = 14
    PYTHON = 15
    UNKNOWN = 16

    @property
    def display_name(self) -> str:
        """The error as Excel spells it in a cell (``#DIV/0!``)."""
        return _ERROR_DISPLAY_NAMES[self.name]


# Display literal per `ErrorCode` member name, generated from `kErrorTable` in src/value.h.
# <error-display-names:begin>
_ERROR_DISPLAY_NAMES = {
    "NULL": "#NULL!",
    "DIV0": "#DIV/0!",
    "VALUE": "#VALUE!",
    "REF": "#REF!",
    "NAME": "#NAME?",
    "NUM": "#NUM!",
    "NA": "#N/A",
    "GETTING_DATA": "#GETTING_DATA",
    "SPILL": "#SPILL!",
    "CALC": "#CALC!",
    "FIELD": "#FIELD!",
    "BLOCKED": "#BLOCKED!",
    "CONNECT": "#CONNECT!",
    "EXTERNAL": "#EXTERNAL!",
    "BUSY": "#BUSY!",
    "PYTHON": "#PYTHON!",
    "UNKNOWN": "#UNKNOWN!",
}
# <error-display-names:end>


class LocaleDateOrder(IntEnum):
    """The order a locale reads the day, month and year of a date in."""

    MDY = 0
    YMD = 1
    DMY = 2


class FunctionAvailability(IntEnum):
    """How far a registered function is implemented."""

    IMPLEMENTED = 0
    IMPLEMENTED_UNVERIFIED = 1
    ENVIRONMENT_BOUND = 2
    UNAVAILABLE_STUB = 3


class ColorKind(IntEnum):
    """How an OOXML ``<color>`` states its value; ``NONE`` has no selector and is written as ``rgb``."""

    NONE = 0
    RGB = 1
    THEME = 2
    INDEXED = 3
    AUTO = 4


class HorizontalAlign(IntEnum):
    """OOXML ``<alignment horizontal>`` of a cell format."""

    GENERAL = 0
    LEFT = 1
    CENTER = 2
    RIGHT = 3
    FILL = 4
    JUSTIFY = 5
    CENTER_CONTINUOUS = 6
    DISTRIBUTED = 7


class VerticalAlign(IntEnum):
    """OOXML ``<alignment vertical>`` of a cell format; ``BOTTOM`` is the default."""

    TOP = 0
    CENTER = 1
    BOTTOM = 2
    JUSTIFY = 3
    DISTRIBUTED = 4


class ReadingOrder(IntEnum):
    """OOXML ``<alignment readingOrder>``: the text direction of a cell."""

    CONTEXT = 0
    LEFT_TO_RIGHT = 1
    RIGHT_TO_LEFT = 2


class UnderlineStyle(IntEnum):
    """OOXML ``<u>`` of a font."""

    NONE = 0
    SINGLE = 1
    DOUBLE = 2
    SINGLE_ACCOUNTING = 3
    DOUBLE_ACCOUNTING = 4


class FontVerticalAlign(IntEnum):
    """OOXML ``<vertAlign>`` of a font."""

    BASELINE = 0
    SUPERSCRIPT = 1
    SUBSCRIPT = 2


class FontScheme(IntEnum):
    """OOXML ``<scheme>`` of a font: its link to the theme fonts; ``NONE`` when absent."""

    NONE = 0
    MAJOR = 1
    MINOR = 2


class FontFamily(IntEnum):
    """OOXML ``<family>`` font-family class."""

    NOT_APPLICABLE = 0
    ROMAN = 1
    SWISS = 2
    MODERN = 3
    SCRIPT = 4
    DECORATIVE = 5


class FillPattern(IntEnum):
    """OOXML ``<patternFill patternType>``."""

    NONE = 0
    SOLID = 1
    MEDIUM_GRAY = 2
    DARK_GRAY = 3
    LIGHT_GRAY = 4
    DARK_HORIZONTAL = 5
    DARK_VERTICAL = 6
    DARK_DOWN = 7
    DARK_UP = 8
    DARK_GRID = 9
    DARK_TRELLIS = 10
    LIGHT_HORIZONTAL = 11
    LIGHT_VERTICAL = 12
    LIGHT_DOWN = 13
    LIGHT_UP = 14
    LIGHT_GRID = 15
    LIGHT_TRELLIS = 16
    GRAY125 = 17
    GRAY0625 = 18


class GradientType(IntEnum):
    """OOXML ``<gradientFill type>``."""

    LINEAR = 0
    PATH = 1


class BorderStyle(IntEnum):
    """OOXML ``style`` of a border side."""

    NONE = 0
    THIN = 1
    MEDIUM = 2
    DASHED = 3
    DOTTED = 4
    THICK = 5
    DOUBLE = 6
    HAIR = 7
    MEDIUM_DASHED = 8
    DASH_DOT = 9
    MEDIUM_DASH_DOT = 10
    DASH_DOT_DOT = 11
    MEDIUM_DASH_DOT_DOT = 12
    SLANT_DASH_DOT = 13


class CfMatchKind(IntEnum):
    """Which kind of conditional-format rule a :class:`CfMatch` reports."""

    DIFFERENTIAL_FORMAT = 0
    COLOR_SCALE = 1
    DATA_BAR = 2
    ICON_SET = 3


class CfDataBarDirection(IntEnum):
    """The direction a data bar grows in."""

    CONTEXT = 0
    LEFT_TO_RIGHT = 1
    RIGHT_TO_LEFT = 2


class CfDataBarAxisPosition(IntEnum):
    """Where a data bar's axis sits."""

    AUTOMATIC = 0
    MIDDLE = 1
    NONE = 2


class ExternalLinkKind(IntEnum):
    """External-link kind carried by :attr:`ExternalLink.kind`."""

    UNKNOWN = 0
    EXTERNAL_BOOK = 1
    OLE = 2
    DDE = 3


class PivotAxis(IntEnum):
    """PivotTable axis for a field."""

    ROW = 0
    COL = 1
    VALUE = 2
    PAGE = 3


class WorkbookFormat(IntEnum):
    """Container format selector for :meth:`Workbook.save`.

    ``UNKNOWN`` is not a valid save target.
    """

    UNKNOWN = 0
    XLSX = 1
    XLSB = 2


class PivotAggregation(IntEnum):
    """Aggregation function for a value-axis field."""

    SUM = 0
    COUNT = 1
    AVERAGE = 2
    MAX = 3
    MIN = 4
    PRODUCT = 5
    COUNT_NUMBERS = 6
    STDDEV = 7
    STDDEVP = 8
    VAR = 9
    VARP = 10


class PivotShowValuesAs(IntEnum):
    """Show-values-as derivation applied to a data-field aggregate."""

    NORMAL = 0
    PERCENT_OF_ROW = 1
    PERCENT_OF_COL = 2
    PERCENT_OF_TOTAL = 3
    RUNNING_TOTAL_IN_ROW = 4
    RUNNING_TOTAL_IN_COL = 5
    INDEX = 6
    DIFFERENCE_FROM = 7
    PERCENT_DIFFERENCE_FROM = 8
    PERCENT_OF_PARENT_ROW = 9
    PERCENT_OF_PARENT_COL = 10
    PERCENT_OF_PARENT = 11
    PERCENT_OF = 12
    PERCENT_OF_RUNNING_TOTAL = 13
    RANK_ASCENDING = 14
    RANK_DESCENDING = 15


class PivotFilterKind(IntEnum):
    """Filter type of a pivot filter: every OOXML pivot filter type.

    The values are the C ABI's filter-type ordinals.
    """

    #: The top / bottom N filter (``count``, ``percent`` or ``sum`` by the operand).
    VALUE_TOP_COUNT = 0
    VALUE_GREATER_THAN = 1
    VALUE_BETWEEN = 2
    LABEL_CONTAINS = 3
    LABEL_BEGINS_WITH = 4
    #: Dates between two dates (``dateBetween``).
    DATE_BETWEEN = 5
    VALUE_EQUAL = 6
    VALUE_NOT_EQUAL = 7
    VALUE_GREATER_THAN_OR_EQUAL = 8
    VALUE_LESS_THAN = 9
    VALUE_LESS_THAN_OR_EQUAL = 10
    VALUE_NOT_BETWEEN = 11
    LABEL_EQUAL = 12
    LABEL_NOT_EQUAL = 13
    LABEL_NOT_BEGINS_WITH = 14
    LABEL_ENDS_WITH = 15
    LABEL_NOT_ENDS_WITH = 16
    LABEL_NOT_CONTAINS = 17
    LABEL_GREATER_THAN = 18
    LABEL_GREATER_THAN_OR_EQUAL = 19
    LABEL_LESS_THAN = 20
    LABEL_LESS_THAN_OR_EQUAL = 21
    LABEL_BETWEEN = 22
    LABEL_NOT_BETWEEN = 23
    VALUE_BOTTOM_COUNT = 24
    VALUE_TOP_PERCENT = 25
    VALUE_BOTTOM_PERCENT = 26
    VALUE_TOP_SUM = 27
    VALUE_BOTTOM_SUM = 28
    DATE_EQUAL = 29
    DATE_NOT_EQUAL = 30
    DATE_OLDER_THAN = 31
    DATE_OLDER_THAN_OR_EQUAL = 32
    DATE_NEWER_THAN = 33
    DATE_NEWER_THAN_OR_EQUAL = 34
    DATE_NOT_BETWEEN = 35
    TOMORROW = 36
    TODAY = 37
    YESTERDAY = 38
    NEXT_WEEK = 39
    THIS_WEEK = 40
    LAST_WEEK = 41
    NEXT_MONTH = 42
    THIS_MONTH = 43
    LAST_MONTH = 44
    NEXT_QUARTER = 45
    THIS_QUARTER = 46
    LAST_QUARTER = 47
    NEXT_YEAR = 48
    THIS_YEAR = 49
    LAST_YEAR = 50
    YEAR_TO_DATE = 51
    Q1 = 52
    Q2 = 53
    Q3 = 54
    Q4 = 55
    M1 = 56
    M2 = 57
    M3 = 58
    M4 = 59
    M5 = 60
    M6 = 61
    M7 = 62
    M8 = 63
    M9 = 64
    M10 = 65
    M11 = 66
    M12 = 67


class PivotDatePeriod(IntFlag):
    """Calendar periods of a pivot date group, combined into the ``periods``
    mask of :meth:`Pivot.set_field_date_group`."""

    SECONDS = 1
    MINUTES = 2
    HOURS = 4
    DAYS = 8
    MONTHS = 16
    QUARTERS = 32
    YEARS = 64


class PivotCalendar(IntEnum):
    """Calendar a pivot date group's years follow."""

    GREGORIAN = 0
    JAPANESE = 1


class PivotCellKind(IntEnum):
    """Discriminator for a projected PivotTable layout cell."""

    HEADER = 0
    ROW_LABEL = 1
    COLUMN_LABEL = 2
    DATA = 3
    ROW_SUBTOTAL = 4
    COLUMN_SUBTOTAL = 5
    GRAND_TOTAL = 6
    BLANK = 7


class PivotReportLayout(IntEnum):
    """Pivot report layout setting (distinct from :class:`PivotLayout`)."""

    COMPACT = 0
    TABULAR = 1
    OUTLINE = 2


class PivotSortOrder(IntEnum):
    """Sibling item order of a pivot field.

    ``MANUAL`` keeps items in the order added with
    :meth:`Pivot.add_field_item` and takes no ``by_field`` key.
    """

    MANUAL = 0
    ASCENDING = 1
    DESCENDING = 2


# Sentinel value for ``PivotDataFieldInput.show_as_base_item`` meaning
# "(previous)" / "(next)" (mirrors the C ABI constants).
PIVOT_SHOW_AS_BASE_PREVIOUS = 1048828
PIVOT_SHOW_AS_BASE_NEXT = 1048829

# The "Σ Values" level in a pivot row / column field order, as read by
# ``Pivot.get_row_field_order`` / ``Pivot.get_column_field_order`` and accepted
# by ``Pivot.set_row_field_order`` / ``Pivot.set_column_field_order`` (``FM_PIVOT_VALUES_FIELD`` reinterpreted as a signed value).
PIVOT_VALUES_FIELD = -2
_PIVOT_VALUES_FIELD_C = 0xFFFFFFFE
_PIVOT_NO_DATA_FIELD = 0xFFFFFFFF

# Sentinel for ``CellStyle.builtin_id`` meaning "custom (no builtinId)".
CELL_STYLE_BUILTIN_ID_NONE = S.CELL_STYLE_BUILTIN_ID_NONE


# ---------------------------------------------------------------------------
# Structured value / input dataclasses
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class CellRect:
    """Cell rectangle: 0-based, inclusive corners ``(first_row, first_col)`` through ``(last_row, last_col)``."""

    first_row: int
    first_col: int
    last_row: int
    last_col: int


@dataclass(frozen=True)
class CellAddress:
    """One cell: 0-based ``row`` and ``col``. A record field of this type also takes an A1 string (``"C3"``)."""

    row: int
    col: int


@dataclass(frozen=True)
class SheetRect:
    """A cell rectangle on sheet index ``sheet``."""

    sheet: int
    rect: CellRect


@dataclass(frozen=True)
class Hyperlink:
    """A sheet hyperlink covering the cells of ``rect``; an attribute the link does not state is ``None``."""

    rect: CellRect
    target: Optional[str]
    location: Optional[str]
    display: Optional[str]
    tooltip: Optional[str]


@dataclass
class HyperlinkInput:
    """A hyperlink :meth:`Sheet.set_hyperlink` / :meth:`Sheet.set_hyperlink_range` set.

    ``target`` ``None`` is a link inside the workbook and needs ``location``
    (its destination); a ``None`` ``display`` or ``tooltip`` is the same as
    omitting it. A :class:`Hyperlink` read back is accepted in its place.
    """

    target: Optional[str]
    location: Optional[str] = None
    display: Optional[str] = None
    tooltip: Optional[str] = None


@dataclass(frozen=True)
class Comment:
    """A cell comment."""

    author: str
    text: str


@dataclass(frozen=True)
class SheetComment:
    """One enumerated sheet comment, including its zero-based anchor."""

    row: int
    col: int
    author: str
    text: str


@dataclass(frozen=True)
class DataValidation:
    """A sheet ``<dataValidation>`` rule (getter result); a text the rule does not state is ``None``."""

    ranges: List[CellRect]
    type: ValidationType
    operator: ComparisonOperator
    error_style: ValidationErrorStyle
    allow_blank: bool
    show_input_message: bool
    show_error_message: bool
    show_in_cell_dropdown: bool
    formula1: Optional[str]
    formula2: Optional[str]
    error_title: Optional[str]
    error_message: Optional[str]
    prompt_title: Optional[str]
    prompt_message: Optional[str]

    def __post_init__(self) -> None:
        object.__setattr__(self, "type", _enum(ValidationType, self.type))
        object.__setattr__(self, "operator", _enum(ComparisonOperator, self.operator))


@dataclass
class DataValidationInput:
    """Argument shape for :meth:`Sheet.add_validation`.

    Every field except ``type`` and ``ranges`` (at least one rectangle) has a
    default. ``allow_blank``, ``show_input_message``, ``show_error_message``
    and ``show_in_cell_dropdown`` default to ``True`` as in Excel's Data
    Validation dialog; the texts default to empty.
    """

    type: ValidationType
    #: Rectangles, unqualified A1 ranges, or one A1 string.
    ranges: Union[List[Union[CellRect, str]], str]
    operator: ComparisonOperator = ComparisonOperator.BETWEEN
    error_style: ValidationErrorStyle = ValidationErrorStyle.STOP
    allow_blank: bool = True
    show_input_message: bool = True
    show_error_message: bool = True
    show_in_cell_dropdown: bool = True
    formula1: str = ""
    formula2: str = ""
    error_title: str = ""
    error_message: str = ""
    prompt_title: str = ""
    prompt_message: str = ""


# `dxf_index` value meaning "no differential format".
_FILTER_DXF_INDEX_NONE = 0xFFFFFFFF


@dataclass
class AutoFilterDateGroup:
    """One ``<dateGroupItem>`` of a values filter; finer fields than ``grouping`` are zero."""

    year: int = 0
    month: int = 0
    day: int = 0
    hour: int = 0
    minute: int = 0
    second: int = 0
    grouping: DateTimeGrouping = DateTimeGrouping.YEAR


@dataclass
class AutoFilterColumn:
    """One ``<filterColumn>``; only the fields of the criterion selected by ``kind`` are read."""

    #: The column's 0-based offset from the filter range's first column.
    column_offset: int
    hidden_button: bool = False
    show_button: bool = True
    kind: FilterKind = FilterKind.NONE
    filter_blank: bool = False
    values: List[str] = field(default_factory=list)
    date_groups: List[AutoFilterDateGroup] = field(default_factory=list)
    custom_and: bool = False
    custom_count: int = 0
    #: A used condition states its operator; ``None`` (the default) in a used slot is rejected on write,
    #: and an unused slot reads as ``None``.
    operator1: Optional[ComparisonOperator] = None
    value1: str = ""
    operator2: Optional[ComparisonOperator] = None
    value2: str = ""
    top: bool = True
    percent: bool = False
    #: The top-10 count or percent (``val``).
    top_value: float = 0.0
    #: The top-10 ``filterVal``; ``None`` when absent.
    filter_value: Optional[float] = None
    dynamic_type: DynamicFilterType = DynamicFilterType.NULL
    #: The dynamic filter's ``val`` / ``maxVal``; ``None`` when absent.
    dynamic_value: Optional[float] = None
    dynamic_max_value: Optional[float] = None
    value_iso: str = ""
    max_value_iso: str = ""
    dxf_index: int = _FILTER_DXF_INDEX_NONE
    cell_color: bool = False
    icon_set: int = 0
    #: ``None`` when the icon filter names no icon.
    icon_index: Optional[int] = None

    def __post_init__(self) -> None:
        for name in ("operator1", "operator2"):
            value = getattr(self, name)
            object.__setattr__(self, name, None if value is None else _enum(ComparisonOperator, value))


@dataclass
class AutoFilterSortCondition:
    """One ``<sortCondition>`` of an AutoFilter sort state."""

    ref: CellRect
    descending: bool = False
    sort_by: SortBy = SortBy.VALUE
    custom_list: str = ""
    #: ``None`` when the condition names no differential format.
    dxf_index: Optional[int] = None
    icon_set: int = 0
    #: ``None`` when the condition names no icon.
    icon_index: Optional[int] = None


@dataclass
class AutoFilterSortState:
    """The ``<sortState>`` of an AutoFilter; ``sort_method`` is a :class:`SortMethod`."""

    ref: CellRect
    column_sort: bool = False
    case_sensitive: bool = False
    sort_method: SortMethod = SortMethod.NONE
    conditions: List[AutoFilterSortCondition] = field(default_factory=list)


@dataclass(frozen=True)
class AutoFilter:
    """A sheet or table AutoFilter; ``range`` includes the header row, ``sort`` may be ``None``."""

    range: CellRect
    columns: List[AutoFilterColumn] = field(default_factory=list)
    sort: Optional[AutoFilterSortState] = None


@dataclass(frozen=True)
class DisplayText:
    """The text Excel displays for a value and the rendering outcome (``display_status``)."""

    text: str
    display_status: DisplayStatus


@dataclass(frozen=True)
class AutoFilterEval:
    """Rows an AutoFilter shows: body row ``first_row + i`` matches when ``match[i]`` is true."""

    first_row: int
    match: List[bool]


@dataclass(frozen=True)
class CellsInRange:
    """One page of populated cells; ``next_cursor`` is ``None`` once no further cell exists."""

    cells: List[RangeCell]
    next_cursor: Optional[int]


@dataclass(frozen=True)
class ValidationOutcome:
    """Result of :meth:`Sheet.validate_value`.

    ``rule_index`` and ``error_style`` are ``None`` (and ``valid`` ``True``)
    when no rule covers the cell; otherwise they describe the applied rule,
    ``rule_index`` indexing :meth:`Sheet.list_validations`.
    """

    valid: bool
    rule_index: Optional[int]
    error_style: Optional[ValidationErrorStyle]


@dataclass(frozen=True)
class Mention:
    """One ``@person`` mention; ``start`` and ``length`` are UTF-16 code units."""

    person_id: str
    mention_id: str
    start: int
    length: int


@dataclass(frozen=True)
class ThreadedComment:
    """A threaded comment: a thread opener, or a reply to the thread ``parent_id`` names."""

    id: str
    row: int
    col: int
    person_id: str
    #: The local wall-clock time it was created; whole seconds. ``None`` when the comment states none.
    created: Optional[datetime.datetime]
    text: str
    parent_id: Optional[str]
    done: bool
    mentions: List[Mention]


@dataclass
class ThreadedCommentInput:
    """A threaded comment :meth:`Sheet.add_threaded_comment` adds.

    ``parent_id`` ``None`` opens a new thread at ``(row, col)``; otherwise
    the comment replies to that thread, taking its anchor and ignoring
    ``row``, ``col`` and ``done``. ``created`` is a ``datetime`` (a naive
    local wall clock; an aware one is a ValueError) or a ``date`` (midnight);
    fractions of a second are dropped, and ``None`` states no time.
    """

    id: str
    person_id: str
    text: str
    row: int = 0
    col: int = 0
    created: Optional[Union[datetime.datetime, datetime.date]] = None
    parent_id: Optional[str] = None
    done: bool = False
    mentions: Sequence[Mention] = ()


@dataclass(frozen=True)
class Person:
    """One entry of the workbook's person list."""

    id: str
    display_name: str
    user_id: str = ""
    provider_id: str = ""


@dataclass(frozen=True)
class ImageInfo:
    """Format and pixel size of an image."""

    format: ImageFormat
    width_px: int
    height_px: int


@dataclass(frozen=True)
class PictureProperties:
    """A picture's image format, pixel size and a copy of its media bytes (:meth:`Picture.get_properties`)."""

    format: ImageFormat
    width_px: int
    height_px: int
    bytes: bytes


@dataclass(frozen=True)
class DrawingProperties:
    """One anchored object of a sheet's drawing; sizes and offsets are in EMU."""

    object_id: int
    kind: DrawingObjectKind
    anchor_kind: AnchorKind
    edit_as: AnchorEditAs
    from_row: int
    from_col: int
    from_row_offset_emu: int
    from_col_offset_emu: int
    to_row: int
    to_col: int
    to_row_offset_emu: int
    to_col_offset_emu: int
    width_emu: int
    height_emu: int
    format: ImageFormat
    name: str
    description: str


@dataclass(frozen=True)
class ShapeProperties:
    """A shape, connector or group of a sheet's drawing; sizes and offsets are in EMU.

    The anchor fields read as in :class:`DrawingProperties`. ``fill_color`` and
    ``line_color`` are RGB or theme colours when the matching kind is ``ShapePaint.SOLID``
    and ``None`` otherwise; a kind is ``None`` when the shape states no fill or
    outline element.
    ``rotation_deg`` is in degrees and ``line_width_pt`` in points (``0`` without a
    width); ``preset_geometry`` is empty for custom geometry or a group.
    """

    object_id: int
    kind: DrawingObjectKind
    anchor_kind: AnchorKind
    edit_as: AnchorEditAs
    from_row: int
    from_col: int
    from_row_offset_emu: int
    from_col_offset_emu: int
    to_row: int
    to_col: int
    to_row_offset_emu: int
    to_col_offset_emu: int
    width_emu: int
    height_emu: int
    name: str
    preset_geometry: str
    text: str
    text_box: bool
    fill_kind: Optional[ShapePaint]
    fill_color: Optional[Color]
    line_kind: Optional[ShapePaint]
    line_color: Optional[Color]
    line_width_pt: float
    rotation_deg: float
    hidden: bool


@dataclass
class PictureInput:
    """Name, description and placement of a picture :meth:`Sheet.add_picture` inserts.

    An empty ``name`` becomes ``Picture <id>``; ``anchor`` ``None`` is a
    one-cell anchor at A1.
    """

    name: str = ""
    description: str = ""
    anchor: Optional[DrawingAnchor] = None


@dataclass
class ShapeInput:
    """A shape :meth:`Sheet.add_shape` appends; a ``None`` field is not applied.

    ``preset_geometry`` defaults to ``rect`` and ``name`` to ``Shape <id>``; a
    ``preset_geometry`` naming a line or connector geometry makes a connector,
    which holds no text. A two-cell anchor's size is its box (``width_emu`` /
    ``height_emu`` are ignored); an absolute anchor's position is
    ``from_col_offset_emu`` (x) and ``from_row_offset_emu`` (y). ``fill_kind``
    and ``line_kind`` must be ``SOLID`` (with an RGB or theme ``fill_color`` /
    ``line_color``) or ``NONE``; a colour needs its kind. ``line_width_pt`` is
    in points and ``rotation_deg`` in degrees.
    """

    anchor_kind: AnchorKind
    edit_as: AnchorEditAs = AnchorEditAs.TWO_CELL
    from_row: int = 0
    from_col: int = 0
    from_row_offset_emu: int = 0
    from_col_offset_emu: int = 0
    to_row: int = 0
    to_col: int = 0
    to_row_offset_emu: int = 0
    to_col_offset_emu: int = 0
    width_emu: int = 0
    height_emu: int = 0
    name: Optional[str] = None
    preset_geometry: Optional[str] = None
    text: Optional[str] = None
    text_box: Optional[bool] = None
    fill_kind: Optional[ShapePaint] = None
    fill_color: Optional[ColorInput] = None
    line_kind: Optional[ShapePaint] = None
    line_color: Optional[ColorInput] = None
    line_width_pt: Optional[float] = None
    rotation_deg: Optional[float] = None
    hidden: Optional[bool] = None


@dataclass(frozen=True)
class SlicerProperties:
    """A table slicer of a worksheet; sizes and offsets are in EMU.

    ``object_id`` is the slicer's frame in the sheet's drawing (``0`` without
    one); the anchor fields read as in :class:`DrawingProperties`, and ``to_*`` is
    zero unless the anchor is two-cell. ``cache_name``, ``table_name`` and
    ``column_name`` identify the slicer cache and the table column it filters;
    the selection itself lives in the table's AutoFilter. ``style`` is
    ``None`` without a style.
    """

    object_id: int
    name: str
    caption: str
    cache_name: str
    table_name: str
    column_name: str
    style: Optional[str]
    column_count: int
    row_height_emu: int
    sort_order: SlicerSortOrder
    cross_filter: SlicerCrossFilter
    anchor_kind: AnchorKind
    from_row: int
    from_col: int
    from_row_offset_emu: int
    from_col_offset_emu: int
    to_row: int
    to_col: int
    to_row_offset_emu: int
    to_col_offset_emu: int


@dataclass
class SlicerInput:
    """A slicer :meth:`Sheet.add_slicer` adds; a ``None`` field is not applied.

    The frame is the two-cell anchor ``from_*`` .. ``to_*`` (0-based cells
    plus EMU offsets). The name and caption default to the column name (the
    name is made unique), ``row_height_emu`` to 241300 and ``column_count``
    to 1.
    """

    from_row: int
    from_col: int
    to_row: int
    to_col: int
    from_row_offset_emu: int = 0
    from_col_offset_emu: int = 0
    to_row_offset_emu: int = 0
    to_col_offset_emu: int = 0
    name: Optional[str] = None
    caption: Optional[str] = None
    style: Optional[str] = None
    column_count: Optional[int] = None
    row_height_emu: Optional[int] = None
    sort_order: Optional[SlicerSortOrder] = None
    cross_filter: Optional[SlicerCrossFilter] = None


@dataclass
class DataTableInputCell:
    """An input cell of a data table; ``deleted`` marks one removed by a row or column delete."""

    row: int
    col: int
    deleted: bool


@dataclass(frozen=True)
class DataTable:
    """A What-If data table (``TABLE()``): its result range and input cells, all 0-based.

    The row input cell receives the values of the row above the range, the column input cell those of the column
    left of it. A one-variable table has one input (``row_oriented`` says which, the other is ``None``), a
    two-variable table both. A deleted input makes the table evaluate to ``#REF!``.
    """

    #: The result range.
    rect: CellRect
    two_dimensional: bool
    row_oriented: bool
    row_input: Optional[DataTableInputCell]
    column_input: Optional[DataTableInputCell]


@dataclass
class DataTableInput:
    """A data table :meth:`Sheet.add_data_table` creates.

    ``rect`` is the result range, a :class:`CellRect` or an unqualified A1
    range. ``row_input`` / ``column_input`` are the input cells that receive the
    values of the row above / the column left of the range, each a
    :class:`DataTableInputCell`, a ``(row, col)`` pair or one A1 cell; give
    one for a one-variable table, both for a two-variable one.
    """

    rect: Union[CellRect, str]
    row_input: Union[DataTableInputCell, Tuple[int, int], str, None] = None
    column_input: Union[DataTableInputCell, Tuple[int, int], str, None] = None


@dataclass(frozen=True)
class Sparkline:
    """One sparkline: the data ``formula`` and its single A1 ``location``.

    ``formula`` is a sheet-qualified range, with a leading ``=`` on read and optional on input.
    """

    formula: str
    location: str


class BlanksAs(IntEnum):
    """How blank cells plot (a chart or a sparkline group); a chart without the element reads ``None``."""

    ZERO = 0
    GAP = 1
    SPAN = 2


@dataclass(frozen=True)
class SparklineGroup:
    """A sparkline group of a sheet's worksheet extension list.

    A colour with ``kind=0`` is absent; ``manual_min`` / ``manual_max`` are
    ``None`` when the axis bound is not set. ``date_axis_ref`` is the group's
    date range (leading ``=`` on read, optional on input), ``None`` when absent.
    """

    type: SparklineType = SparklineType.LINE
    color_series: Color = field(default_factory=lambda: Color())
    color_negative: Color = field(default_factory=lambda: Color())
    color_axis: Color = field(default_factory=lambda: Color())
    color_markers: Color = field(default_factory=lambda: Color())
    color_first: Color = field(default_factory=lambda: Color())
    color_last: Color = field(default_factory=lambda: Color())
    color_high: Color = field(default_factory=lambda: Color())
    color_low: Color = field(default_factory=lambda: Color())
    markers: bool = False
    high: bool = False
    low: bool = False
    first: bool = False
    last: bool = False
    negative: bool = False
    display_x_axis: bool = False
    display_hidden: bool = False
    min_axis_type: SparklineAxisType = SparklineAxisType.INDIVIDUAL
    max_axis_type: SparklineAxisType = SparklineAxisType.INDIVIDUAL
    manual_min: Optional[float] = None
    manual_max: Optional[float] = None
    date_axis: bool = False
    date_axis_ref: Optional[str] = None
    line_weight: float = 0.75
    empty_cells_as: BlanksAs = BlanksAs.GAP
    sparklines: List[Sparkline] = field(default_factory=list)

    def __post_init__(self) -> None:
        object.__setattr__(self, "empty_cells_as", _enum(BlanksAs, self.empty_cells_as))


class ChartPlotType(IntEnum):
    """Plot group kind."""

    BAR = 0
    BAR_3D = 1
    LINE = 2
    LINE_3D = 3
    AREA = 4
    AREA_3D = 5
    PIE = 6
    PIE_3D = 7
    DOUGHNUT = 8
    OF_PIE = 9
    SCATTER = 10
    BUBBLE = 11
    RADAR = 12
    STOCK = 13
    SURFACE = 14
    SURFACE_3D = 15
    #: An extended (``cx:``) chart; :attr:`ChartPlot.extended_layout` names its layout.
    EXTENDED = 16


class ChartExtendedLayout(IntEnum):
    """Layout of an extended chart's plot (``cx:series/@layoutId``)."""

    BOX_WHISKER = 0
    CLUSTERED_COLUMN = 1
    FUNNEL = 2
    PARETO_LINE = 3
    REGION_MAP = 4
    SUNBURST = 5
    TREEMAP = 6
    WATERFALL = 7


class ChartInsertKind(IntEnum):
    """Kinds :meth:`Sheet.add_chart` and :meth:`Chart.set_type` produce."""

    COLUMN = 0
    BAR = 1
    LINE = 2
    AREA = 3
    PIE = 4
    SCATTER = 5
    DOUGHNUT = 6


class ChartGrouping(IntEnum):
    """Series grouping of a bar, line or area plot."""

    CLUSTERED = 0
    STACKED = 1
    PERCENT_STACKED = 2
    STANDARD = 3


class ChartPointState(IntEnum):
    """State of a resolved data point."""

    VALUE = 0
    GAP = 1
    CONNECT = 2


class ChartAxisKind(IntEnum):
    """Axis element type."""

    CATEGORY = 0
    VALUE = 1
    DATE = 2
    SERIES = 3


class ChartAxisCrosses(IntEnum):
    """Where an axis crosses its partner; an axis without the element reads ``None``."""

    AUTO_ZERO = 0
    MIN = 1
    MAX = 2
    AT = 3


class ChartTitleKind(IntEnum):
    """What a chart or axis title holds; a chart or axis without a title reads ``None``."""

    TEXT = 1
    REF = 2
    AUTO = 3


class ChartSeriesDirection(IntEnum):
    """Direction series are split from a source range."""

    AUTO = 0
    COLUMNS = 1
    ROWS = 2


class ChartLegendPosition(IntEnum):
    """Legend placement; a chart without a legend reads ``None``."""

    RIGHT = 0
    BOTTOM = 1
    TOP_RIGHT = 2
    LEFT = 3
    TOP = 4


class ChartAxisPosition(IntEnum):
    """The side of the plot area an axis sits on."""

    BOTTOM = 0
    TOP = 1
    LEFT = 2
    RIGHT = 3


class ChartBarDirection(IntEnum):
    """Whether a bar-family plot draws vertical columns or horizontal bars."""

    COLUMN = 0
    BAR = 1


class ChartLineDash(IntEnum):
    """Outline dash style."""

    SOLID = 0
    DASH = 1
    DOT = 2
    DASH_DOT = 3
    OTHER = 4


class ChartMarkerSymbol(IntEnum):
    """Series marker symbol; ``PICTURE`` is read-only and a patch carrying it is rejected."""

    NONE = 0
    AUTO = 1
    CIRCLE = 2
    SQUARE = 3
    DIAMOND = 4
    TRIANGLE = 5
    X = 6
    STAR = 7
    PLUS = 8
    DOT = 9
    DASH = 10
    PICTURE = 11


class ChartLabelPosition(IntEnum):
    """Data label placement."""

    BEST_FIT = 0
    CENTER = 1
    INSIDE_BASE = 2
    INSIDE_END = 3
    OUTSIDE_END = 4
    ABOVE = 5
    BELOW = 6
    LEFT = 7
    RIGHT = 8


class ChartRole(IntEnum):
    """A chart input: the six series roles, the chart title and an axis title.

    The six series roles also index the per-series inputs (:class:`ChartSeriesRanges`).
    """

    NAME = 0
    CATEGORIES = 1
    VALUES = 2
    X = 3
    Y = 4
    SIZE = 5
    TITLE = 6
    AXIS_TITLE = 7


class ChartSourceKind(IntEnum):
    """What a chart series input holds."""

    NONE = 0
    REF = 1
    LITERAL = 2


@dataclass(frozen=True)
class ChartDataSource:
    """One series input: ``kind`` is ``NONE``, ``REF`` (``formula`` set) or ``LITERAL`` (``literal`` set)."""

    kind: ChartSourceKind
    formula: str = ""
    literal: List[Union[float, str]] = field(default_factory=list)


@dataclass(frozen=True)
class ChartTitle:
    """A chart or axis title; ``text`` is the flattened rich text (``TEXT``), ``formula`` the reference (``REF``)."""

    kind: ChartTitleKind
    text: str = ""
    formula: str = ""


@dataclass(frozen=True)
class ChartFill:
    """Solid fill; ``color`` is an RGB or theme :class:`Color` and ``transparency_pct`` a percentage, each ``None`` when absent."""

    color: Optional[Color] = None
    transparency_pct: Optional[float] = None


@dataclass(frozen=True)
class ChartLine:
    """Outline; ``color`` is an RGB or theme :class:`Color` (``None`` when absent) and ``none`` an explicit no-line."""

    color: Optional[Color] = None
    transparency_pct: Optional[float] = None
    width_pt: Optional[float] = None
    dash: Optional[ChartLineDash] = None
    none: bool = False


@dataclass(frozen=True)
class ChartMarker:
    """Series or point marker; ``size_pt`` is in points."""

    symbol: ChartMarkerSymbol
    size_pt: Optional[int]
    fill: ChartFill
    line: ChartLine


@dataclass(frozen=True)
class ChartLabels:
    """Data labels of a series."""

    show_value: bool
    show_category: bool
    show_series: bool
    show_percent: bool
    show_legend_key: bool
    separator: Optional[str]
    position: Optional[ChartLabelPosition]
    number_format: Optional[str]
    source_linked: bool
    font_color: Optional[Color]
    font_size_pt: Optional[float]
    deleted: bool


@dataclass(frozen=True)
class ChartPoint:
    """One formatted data point (``c:dPt``)."""

    index: int
    fill: ChartFill
    line: ChartLine
    marker: Optional[ChartMarker]


@dataclass(frozen=True)
class ChartSeries:
    """One series of a plot group; ``index`` is its ``c:idx``."""

    index: int
    order: int
    filtered: bool
    name: ChartDataSource
    categories: ChartDataSource
    values: ChartDataSource
    x: ChartDataSource
    y: ChartDataSource
    size: ChartDataSource
    fill: ChartFill
    line: ChartLine
    marker: Optional[ChartMarker]
    smooth: Optional[bool]
    points: List[ChartPoint]
    labels: Optional[ChartLabels]


@dataclass(frozen=True)
class ChartPlot:
    """One plot group; ``bar_direction`` is ``None`` outside the bar family or when the file omits it.

    ``extended_layout`` is ``None`` unless ``type`` is ``EXTENDED``. ``gap_width_pct`` is a percentage of
    the bar width and ``first_slice_angle_deg`` is in degrees.
    """

    type: ChartPlotType
    extended_layout: Optional[ChartExtendedLayout]
    bar_direction: Optional[ChartBarDirection]
    grouping: Optional[ChartGrouping]
    vary_colors: bool
    gap_width_pct: Optional[int]
    overlap: Optional[int]
    hole_size: Optional[int]
    first_slice_angle_deg: Optional[int]
    axis_ids: List[int]
    series: List[ChartSeries]


@dataclass(frozen=True)
class ChartAxis:
    """One axis; ``crosses_at`` is ``None`` unless ``crosses`` is ``AT``, and ``title`` is ``None`` without a title."""

    id: int
    kind: ChartAxisKind
    position: ChartAxisPosition
    visible: bool
    cross_axis: int
    crosses: Optional[ChartAxisCrosses]
    crosses_at: Optional[float]
    min: Optional[float]
    max: Optional[float]
    major_unit: Optional[float]
    minor_unit: Optional[float]
    log_base: Optional[float]
    reversed: bool
    number_format: Optional[str]
    source_linked: bool
    title: Optional[ChartTitle]
    major_gridlines: bool
    minor_gridlines: bool
    tick_label_position: Optional[ChartTickLabelPosition]


@dataclass(frozen=True)
class ChartLegend:
    """Legend state; ``hidden_series`` lists the ``c:idx`` of series whose legend entry is hidden."""

    position: ChartLegendPosition
    overlay: bool
    hidden_series: List[int] = field(default_factory=list)


@dataclass(frozen=True)
class ChartSource:
    """The rectangle every series was split from and how it was split."""

    range: str
    direction: ChartSeriesDirection
    header_row: bool
    header_column: bool


@dataclass(frozen=True)
class ChartProperties:
    """The typed view of a chart (:meth:`Chart.get_properties`).

    ``title`` and ``legend`` are ``None`` when the chart has none, ``blanks_as``
    when the file omits it, and ``source`` when the series share no rectangle.
    """

    object_id: int
    name: str
    description: str
    is_extended: bool
    is_pivot_chart: bool
    on_chartsheet: bool
    title: Optional[ChartTitle]
    auto_title_deleted: Optional[bool]
    legend: Optional[ChartLegend]
    blanks_as: Optional[BlanksAs]
    plot_visible_only: Optional[bool]
    date1904: bool
    style: Optional[int]
    plots: List[ChartPlot]
    axes: List[ChartAxis]
    hidden_categories: List[int]
    source: Optional[ChartSource]
    #: ``c16r3:dataDisplayOptions16/dispNaAsBlank``; ``None`` when the element is absent.
    na_as_blank: Optional[bool]


@dataclass(frozen=True)
class ChartSeriesRanges:
    """The resolved rectangles of each series input."""

    name: List[SheetRect] = field(default_factory=list)
    categories: List[SheetRect] = field(default_factory=list)
    values: List[SheetRect] = field(default_factory=list)
    x: List[SheetRect] = field(default_factory=list)
    y: List[SheetRect] = field(default_factory=list)
    size: List[SheetRect] = field(default_factory=list)


@dataclass(frozen=True)
class ChartDataSeries:
    """One resolved series; ``values`` and ``states`` (:class:`ChartPointState`) are parallel.

    ``values`` is the y value in the XY domain; ``x`` is ``None`` outside it
    and ``sizes`` is ``None`` unless the series is a bubble series.
    ``format_code`` is the number format of the values; ``x_format_code`` is
    that of the x values, empty outside the XY domain.
    """

    plot_index: int
    index: int
    name: Optional[str]
    format_code: str
    x_format_code: str
    filtered: bool
    values: List[float]
    states: List[ChartPointState]
    x: Optional[List[float]]
    sizes: Optional[List[float]]
    ranges: ChartSeriesRanges


@dataclass(frozen=True)
class ChartCategories:
    """Resolved categories; ``values`` is set for numeric (date) categories and ``generated`` when the labels are 1..n."""

    labels: List[str]
    levels: List[List[str]]
    values: Optional[List[float]]
    format_code: str
    generated: bool


@dataclass(frozen=True)
class ChartAxisTitleText:
    """Resolved title text of one axis."""

    axis_id: int
    text: str


@dataclass(frozen=True)
class ChartUnresolved:
    """An input that did not resolve; its points read as gaps.

    ``source`` is the role of the input; ``axis_id`` names the axis of an ``AXIS_TITLE``.
    """

    series_index: int
    source: ChartRole
    axis_id: int


@dataclass
class ChartData:
    """A chart's series resolved against the current cell values (:meth:`Chart.get_data`)."""

    domain: ChartDomain
    point_count: int
    categories: ChartCategories
    series: List[ChartDataSeries]
    title_text: Optional[str]
    axis_title_texts: List[ChartAxisTitleText]
    unresolved: List[ChartUnresolved]


class ChartFillPatch(TypedDict, total=False):
    """Fill patch: an omitted key is kept, ``None`` removes the element; a colour is an RGB or theme ``ColorInput`` (indexed and automatic are rejected)."""

    color: Optional[ColorInput]
    transparency_pct: Optional[float]


class ChartLinePatch(TypedDict, total=False):
    """Outline patch; ``none`` is a bool and rejects ``None``; a colour is an RGB or theme ``ColorInput`` (indexed and automatic are rejected)."""

    color: Optional[ColorInput]
    transparency_pct: Optional[float]
    width_pt: Optional[float]
    dash: Optional[ChartLineDash]
    none: Optional[bool]


class ChartMarkerPatch(TypedDict, total=False):
    """Marker patch; ``symbol`` rejects ``None``."""

    symbol: Optional[ChartMarkerSymbol]
    size_pt: Optional[float]
    fill: Optional[ChartFillPatch]
    line: Optional[ChartLinePatch]


class ChartLabelsPatch(TypedDict, total=False):
    """Data-label patch; the ``show_*`` flags are bools and reject ``None``."""

    show_value: Optional[bool]
    show_category: Optional[bool]
    show_series: Optional[bool]
    show_percent: Optional[bool]
    show_legend_key: Optional[bool]
    separator: Optional[str]
    position: Optional[ChartLabelPosition]
    number_format: Optional[str]
    font_color: Optional[ColorInput]
    font_size_pt: Optional[float]


class ChartSeriesPatch(TypedDict, total=False):
    """Series patch.

    A reference is its new text or a :class:`ChartDataSource` with a reference; ``""`` removes ``name``,
    ``categories`` or ``x`` (an empty ``values``, ``y`` or ``size`` is rejected). ``None`` on ``marker``
    or ``labels`` removes it.
    """

    name: Union[ChartDataSource, str, None]
    categories: Union[ChartDataSource, str, None]
    values: Union[ChartDataSource, str, None]
    x: Union[ChartDataSource, str, None]
    y: Union[ChartDataSource, str, None]
    size: Union[ChartDataSource, str, None]
    fill: Optional[ChartFillPatch]
    line: Optional[ChartLinePatch]
    marker: Optional[ChartMarkerPatch]
    smooth: Optional[bool]
    labels: Optional[ChartLabelsPatch]


@dataclass
class ChartSeriesInput:
    """A new series for :meth:`Chart.add_series`.

    ``values`` is required (``y`` for a scatter plot); every other field is optional and reads as the
    same key of :class:`ChartSeriesPatch`.
    """

    values: Union[ChartDataSource, str, None] = None
    name: Union[ChartDataSource, str, None] = None
    categories: Union[ChartDataSource, str, None] = None
    x: Union[ChartDataSource, str, None] = None
    y: Union[ChartDataSource, str, None] = None
    size: Union[ChartDataSource, str, None] = None
    fill: Optional[ChartFillPatch] = None
    line: Optional[ChartLinePatch] = None
    marker: Optional[ChartMarkerPatch] = None
    smooth: Optional[bool] = None
    labels: Optional[ChartLabelsPatch] = None


class ChartPointPatch(TypedDict, total=False):
    """Data-point patch; an empty patch (or ``None``) removes the point's formatting."""

    fill: Optional[ChartFillPatch]
    line: Optional[ChartLinePatch]
    marker: Optional[ChartMarkerPatch]


class ChartPlotPatch(TypedDict, total=False):
    """Plot-group patch; a field the plot type does not have is rejected. ``hole_size`` is 10..90 (doughnut only)."""

    gap_width_pct: Optional[int]
    overlap: Optional[int]
    first_slice_angle_deg: Optional[int]
    vary_colors: Optional[bool]
    hole_size: Optional[int]


class ChartAxisPatch(TypedDict, total=False):
    """Axis patch.

    ``title`` takes the forms of :meth:`Chart.set_title`: text, or a
    :class:`ChartTitle` (``AUTO`` leaves the text to the viewer).
    ``None`` on ``title`` removes the axis title, on
    ``log_base`` makes the axis linear and on ``tick_label_position``
    (a :class:`ChartTickLabelPosition`) restores the default.
    ``crosses_at`` is read when ``crosses`` is ``AT``.
    """

    visible: Optional[bool]
    title: Optional[Union[str, ChartTitle]]
    min: Optional[float]
    max: Optional[float]
    major_unit: Optional[float]
    number_format: Optional[str]
    major_gridlines: Optional[bool]
    reversed: Optional[bool]
    log_base: Optional[float]
    crosses: Optional[ChartAxisCrosses]
    crosses_at: Optional[float]
    tick_label_position: Optional[ChartTickLabelPosition]


class ChartDisplayPatch(TypedDict, total=False):
    """Display patch for :meth:`Chart.update_display`; an omitted key is kept, ``None`` removes the element."""

    blanks_as: Optional[BlanksAs]
    plot_visible_only: Optional[bool]
    na_as_blank: Optional[bool]


@dataclass(frozen=True)
class WorkbookProtection:
    """Workbook ``<workbookProtection>``. Hash fields are stored as given; the host computes them.

    A text field the element does not state is ``None``.
    """

    lock_structure: bool = False
    lock_windows: bool = False
    lock_revision: bool = False
    #: The legacy 16-bit password hash (hex, ``workbookPassword``).
    workbook_legacy_password_hash: Optional[str] = None
    workbook_algorithm_name: Optional[str] = None
    workbook_hash_value: Optional[str] = None
    workbook_salt_value: Optional[str] = None
    workbook_spin_count: int = 0
    #: The legacy 16-bit password hash (hex, ``revisionsPassword``).
    revisions_legacy_password_hash: Optional[str] = None
    revisions_algorithm_name: Optional[str] = None
    revisions_hash_value: Optional[str] = None
    revisions_salt_value: Optional[str] = None
    revisions_spin_count: int = 0


@dataclass(frozen=True)
class ProtectedRange:
    """A sheet ``<protectedRange>`` covering the rectangles of ``ranges``."""

    name: str
    ranges: List[CellRect]
    #: The legacy 16-bit password hash (hex, ``password``); ``None`` when absent.
    legacy_password_hash: Optional[str] = None
    algorithm_name: str = ""
    hash_value: str = ""
    salt_value: str = ""
    spin_count: int = 0
    security_descriptor: str = ""


@dataclass(frozen=True)
class SheetProtection:
    """The protection of a protected sheet (:meth:`Sheet.get_sheet_protection`).

    Each ``allow_*`` flag permits one action on the protected sheet. The
    defaults are those of Excel's Protect Sheet dialog: selecting locked and
    unlocked cells is allowed, every other action is not. A text field the
    sheet does not state is ``None``.
    """

    algorithm_name: Optional[str] = None
    hash_value: Optional[str] = None
    salt_value: Optional[str] = None
    spin_count: int = 0
    #: The legacy 16-bit password hash (hex, ``password``).
    legacy_password_hash: Optional[str] = None
    allow_edit_objects: bool = False
    allow_edit_scenarios: bool = False
    allow_format_cells: bool = False
    allow_format_columns: bool = False
    allow_format_rows: bool = False
    allow_insert_columns: bool = False
    allow_insert_rows: bool = False
    allow_insert_hyperlinks: bool = False
    allow_delete_columns: bool = False
    allow_delete_rows: bool = False
    allow_select_locked_cells: bool = True
    allow_select_unlocked_cells: bool = True
    allow_sort: bool = False
    allow_auto_filter: bool = False
    allow_pivot_tables: bool = False


@dataclass
class SheetProtectionInput:
    """The protection :meth:`Sheet.set_sheet_protection` applies; the fields and defaults of :class:`SheetProtection`."""

    algorithm_name: Optional[str] = None
    hash_value: Optional[str] = None
    salt_value: Optional[str] = None
    spin_count: int = 0
    legacy_password_hash: Optional[str] = None
    allow_edit_objects: bool = False
    allow_edit_scenarios: bool = False
    allow_format_cells: bool = False
    allow_format_columns: bool = False
    allow_format_rows: bool = False
    allow_insert_columns: bool = False
    allow_insert_rows: bool = False
    allow_insert_hyperlinks: bool = False
    allow_delete_columns: bool = False
    allow_delete_rows: bool = False
    allow_select_locked_cells: bool = True
    allow_select_unlocked_cells: bool = True
    allow_sort: bool = False
    allow_auto_filter: bool = False
    allow_pivot_tables: bool = False


@dataclass(frozen=True)
class SheetView:
    """Per-sheet view: zoom, frozen-pane counts, tab visibility, and the
    display / orientation flags mirrored from OOXML ``<sheetView>``.
    """

    #: The zoom as a percentage (100 is 100%).
    zoom_pct: int
    freeze_rows: int
    freeze_cols: int
    visibility: SheetVisibility = SheetVisibility.VISIBLE
    show_grid_lines: bool = True
    show_row_column_headers: bool = True
    show_zeros: bool = True
    right_to_left: bool = False
    tab_selected: bool = False
    view_mode: SheetViewMode = SheetViewMode.NORMAL
    top_left_row: int = 0
    top_left_col: int = 0
    pane_top_left_row: int = 0
    pane_top_left_col: int = 0
    #: The split position of a split (unfrozen) pane, in twips.
    split_x_twips: float = 0.0
    split_y_twips: float = 0.0
    show_formulas: bool = False
    show_outline_symbols: bool = True


@dataclass(frozen=True)
class ColumnLayout:
    """A per-column-range layout override (inclusive ``[first, last]``).

    ``width_chars`` is the explicit width in characters (the raw source bit or
    a legacy non-zero width), ``None`` when the span states none; an explicit
    zero stays distinguishable from an absent width. The column XF is read
    through :meth:`SheetOoxml.get_column_xf_index`.
    """

    first: int
    last: int
    width_chars: Optional[float]
    hidden: bool
    outline_level: int
    #: ``collapsed="1"``: the outline group this span summarises is collapsed.
    collapsed: bool = False


@dataclass(frozen=True)
class RowLayout:
    """A per-row layout override; the row XF is read through :meth:`SheetOoxml.get_row_xf_index`."""

    row: int
    #: The stored height in points; ``None`` when the row stores none.
    height_pt: Optional[float]
    hidden: bool
    outline_level: int
    #: ``customHeight="1"``: an explicit override rather than a cached auto height.
    custom_height: bool = False
    #: ``collapsed="1"``: the outline group this row summarises is collapsed.
    collapsed: bool = False


@dataclass(frozen=True)
class IterativeSettings:
    """Iterative-calculation settings read back by ``get_iterative``."""

    enabled: bool
    max_iterations: int
    max_change: float


@dataclass(frozen=True)
class SheetSelection:
    """The active pane's selection read back by ``get_selection``."""

    active_row: int
    active_col: int
    ranges: List[CellRect] = field(default_factory=list)


@dataclass(frozen=True)
class TableStyleDefaults:
    """Default table and pivot style names read back by ``get_table_style_defaults``; ``None`` when unset."""

    table_style: Optional[str]
    pivot_style: Optional[str]


@dataclass(frozen=True)
class PivotGrandTotals:
    """Grand-total flags read back by :meth:`Pivot.get_grand_totals`."""

    rows: bool
    columns: bool


@dataclass(frozen=True)
class Pagination:
    """Resolved physical pagination for one worksheet.

    ``print_area`` is the sheet's declared ``_xlnm.Print_Area`` as
    :class:`CellRect` rectangles. It is empty when the sheet declares no print area and is
    not backfilled with the used range; pagination itself still falls back
    to the used range, so ``page_count`` can be non-zero while this is
    empty. The break lists hold the breaks in effect, ascending, each before
    the zero-based row or column a new physical page begins with.
    """

    page_count: int
    print_area: List[CellRect]
    row_breaks: List["RowBreak"]
    column_breaks: List["ColumnBreak"]
    paper: "Paper"
    margins: "MarginsPt"
    printable: "RectPt"
    #: Effective print scale as a percentage (100 is 100%).
    scale_pct: float
    page_order: PageOrder
    print_titles: "PrintTitles"
    pages: List["PageLayout"]

    def __post_init__(self) -> None:
        object.__setattr__(self, "page_order", _enum(PageOrder, self.page_order))


@dataclass(frozen=True)
class RectPt:
    """A rectangle in points."""

    x: float
    y: float
    width_pt: float
    height_pt: float


@dataclass(frozen=True)
class Paper:
    """The paper a sheet prints on, after the orientation swap."""

    width_pt: float
    height_pt: float
    landscape: bool
    #: False when ``paperSize`` was unrecognised and A4 substituted.
    known: bool


@dataclass(frozen=True)
class MarginsPt:
    """Page margins in inches."""

    left_in: float
    right_in: float
    top_in: float
    bottom_in: float
    header_in: float
    footer_in: float


@dataclass(frozen=True)
class Span:
    """A 0-based, inclusive run of rows or columns."""

    first: int
    last: int


@dataclass(frozen=True)
class PrintTitles:
    """The rows and columns repeated on every printed page; ``None`` for an axis that repeats nothing."""

    rows: Optional[Span]
    columns: Optional[Span]


@dataclass(frozen=True)
class PageLayout:
    """One physical page: the cell block it carries and where it sits."""

    area_index: int
    #: The page's cell block.
    rect: CellRect
    origin_x_pt: float
    origin_y_pt: float
    width_pt: float
    height_pt: float


@dataclass(frozen=True)
class SheetFormatDefaults:
    """Sheet default column / row metrics (``<sheetFormatPr>``).

    ``default_col_width_chars`` is a stored width (characters with the column
    padding included) and ``default_row_height_pt`` in points; each is ``None``
    when the sheet stores none. ``Sheet.update_format_defaults`` takes these
    fields as keyword arguments: ``None`` clears the stored value and
    :data:`UNCHANGED` (the default) keeps it.
    ``effective_default_col_width_chars`` / ``effective_default_row_height_pt``
    are read-only values the sheet lays out with: the stated default, else the
    column Excel derives from ``base_col_width_chars`` and the Normal font
    (Calibri 11, base 8: 9.140625, shown as 8.43) and 102/7 pt; the row height is 0 when
    ``zero_height`` hides rows by default. They are ignored when setting.
    """

    default_col_width_chars: Optional[float] = None
    default_row_height_pt: Optional[float] = None
    base_col_width_chars: float = 8.0
    effective_default_col_width_chars: float = 0.0
    effective_default_row_height_pt: float = 0.0
    #: ``customHeight``: ``default_row_height_pt`` is an explicit override rather
    #: than a cached figure Excel recomputes from the Normal font. ``None``
    #: keeps the stored flag when setting.
    custom_height: Optional[bool] = None
    #: ``zeroHeight``: rows are hidden unless they say otherwise. ``None``
    #: keeps the stored flag when setting.
    zero_height: Optional[bool] = None


@dataclass(frozen=True)
class WidthModel:
    """Column-width conversion model for the workbook's Normal font.

    ``points_per_char`` and ``padding_pt`` are per displayed character:
    ``chars * points_per_char + padding_pt``; stored widths go through
    ``column_chars_to_pt``.
    """

    points_per_char: float
    padding_pt: float
    normal_font_size_pt: float
    #: False when the Normal font is not a measured calibration point.
    calibrated: bool
    normal_font_name: str
    platform: WidthPlatform

    def __post_init__(self) -> None:
        object.__setattr__(self, "platform", _enum(WidthPlatform, self.platform))


@dataclass(frozen=True)
class RowBreak:
    """One row break (``<rowBreaks><brk>``).

    The break sits before the zero-based row ``row`` and covers the columns
    ``first..last`` (OOXML ``min`` / ``max``). ``manual`` is False for a
    break Excel computed rather than one a user placed. As the argument of
    :meth:`Sheet.set_row_break`, ``row`` is ignored and ``first == last == 0``
    covers the whole row.
    """

    row: int
    first: int = 0
    last: int = 0
    manual: bool = True


@dataclass(frozen=True)
class ColumnBreak:
    """One column break (``<colBreaks><brk>``): before the zero-based column ``col``, covering the rows ``first..last``.

    ``manual`` and the argument form of :meth:`Sheet.set_column_break` are as in :class:`RowBreak`.
    """

    col: int
    first: int = 0
    last: int = 0
    manual: bool = True


@dataclass(frozen=True)
class PageSetup:
    """Effective ``<pageSetup>`` values, with per-attribute presence.

    Every value field carries the setting in force. The ``*_stated`` flags
    say whether the XML actually declares it, which is the only way to tell
    ``scale="100"`` from an absent attribute that defaults to 100.
    ``scale_pct`` is the print scale as a percentage.

    ``fit_to_page`` lives in ``<sheetPr><pageSetUpPr>`` rather than
    ``<pageSetup>``; it is reported here because callers reason about it as
    part of one page-setup decision.
    """

    #: ``None`` when the sheet states no orientation (Excel prints portrait).
    orientation: Optional[PageOrientation]
    paper_size: int
    scale_pct: int
    fit_to_width: int
    fit_to_height: int
    fit_to_page: bool
    orientation_stated: bool
    paper_size_stated: bool
    scale_pct_stated: bool
    fit_to_width_stated: bool
    fit_to_height_stated: bool
    fit_to_page_stated: bool


@dataclass(frozen=True)
class PageMargins:
    """Effective ``<pageMargins>`` values in inches, with presence flags."""

    left_in: float
    right_in: float
    top_in: float
    bottom_in: float
    header_in: float
    footer_in: float
    left_in_stated: bool
    right_in_stated: bool
    top_in_stated: bool
    bottom_in_stated: bool
    header_in_stated: bool
    footer_in_stated: bool


@dataclass(frozen=True)
class PrintOptions:
    """Effective ``<printOptions>`` values; an attribute the sheet does not state reads false with its flag false."""

    grid_lines: bool
    headings: bool
    horizontal_centered: bool
    vertical_centered: bool
    grid_lines_stated: bool
    headings_stated: bool
    horizontal_centered_stated: bool
    vertical_centered_stated: bool


@dataclass(frozen=True)
class HeaderFooter:
    """Effective ``<headerFooter>`` values.

    A section the sheet does not have is ``None``; the text is decoded as
    :meth:`Sheet.update_header_footer` takes it. A flag the sheet does not state
    reads its schema default (``different_odd_even`` / ``different_first``
    false, ``scale_with_doc`` / ``align_with_margins`` true) with its
    ``*_stated`` flag false.
    """

    odd_header: Optional[str]
    odd_footer: Optional[str]
    even_header: Optional[str]
    even_footer: Optional[str]
    first_header: Optional[str]
    first_footer: Optional[str]
    different_odd_even: bool
    different_first: bool
    scale_with_doc: bool
    align_with_margins: bool
    different_odd_even_stated: bool
    different_first_stated: bool
    scale_with_doc_stated: bool
    align_with_margins_stated: bool


@dataclass(frozen=True)
class CfMatch:
    """A resolved conditional-format match for one cell.

    Active fields depend on ``kind``; the rest carry default-zero values.
    """

    kind: CfMatchKind
    priority: int
    dxf_index_engaged: bool
    dxf_index: int
    color: Color
    bar_length_pct: float
    bar_axis_position_pct: float
    bar_is_negative: bool
    bar_fill: Color
    bar_border_engaged: bool
    bar_border: Color
    bar_gradient: bool
    icon_set_name: IconSetName
    icon_index: int
    bar_direction: CfDataBarDirection = CfDataBarDirection.CONTEXT

    def __post_init__(self) -> None:
        object.__setattr__(self, "icon_set_name", _enum(IconSetName, self.icon_set_name))


@dataclass(frozen=True)
class CfCell:
    """One cell's CF matches inside a viewport-range evaluation."""

    row: int
    col: int
    matches: List[CfMatch]


@dataclass(frozen=True)
class Cfvo:
    """One conditional-format threshold."""

    type: CfValueObjectType
    value: Optional[str] = None
    gte: bool = True

    def __post_init__(self) -> None:
        object.__setattr__(self, "type", _enum(CfValueObjectType, self.type))


@dataclass(frozen=True)
class ColorScale:
    """``<colorScale>`` payload: one colour per threshold (a ``ColorInput`` on input)."""

    thresholds: List[Cfvo]
    colors: List[Color]


@dataclass(frozen=True)
class DataBar:
    """``<dataBar>`` payload of a conditional-format rule.

    The fields after ``max_length_pct`` live in the ``x14`` extension
    rather than the legacy element. ``None`` means "do not override the
    model default" -- gradient fill on, automatic axis, negative fill
    equal to ``fill``, no border, and a black axis. Reading a rule back
    always populates them, so ``add_conditional_format`` of a value
    returned by :meth:`Sheet.list_cf_rules` reproduces the rule. Each colour
    takes a ``ColorInput`` on input.
    """

    minimum: Cfvo
    maximum: Cfvo
    fill: Color
    show_value: bool = True
    min_length_pct: int = 10
    max_length_pct: int = 90
    gradient: Optional[bool] = None
    axis_position: Optional[CfDataBarAxisPosition] = None
    negative_fill: Optional[Color] = None
    border: Optional[Color] = None
    negative_border: Optional[Color] = None
    axis_color: Optional[Color] = None
    #: The edge the bar grows from.
    direction: CfDataBarDirection = CfDataBarDirection.CONTEXT


@dataclass(frozen=True)
class IconSet:
    """``<iconSet>`` payload of a conditional-format rule.

    ``percent`` is round-trip only: it is preserved across load and save
    but never consulted during evaluation. Each threshold in
    ``thresholds`` carries its own ``type``, and that type is what
    interprets it.

    ``floor`` is the lower bound of the lowest icon's bucket; a cell below
    it gets no icon. ``None`` means Excel's default, ``percent 0``; reading
    a rule back always populates it.
    """

    name: IconSetName
    thresholds: List[Cfvo]
    reverse: bool = False
    show_value: bool = True
    percent: bool = True
    floor: Optional[Cfvo] = None
    #: Custom icon set: the icon drawn for each level, empty when the set
    #: draws its own icons.
    icons: List["CfIcon"] = field(default_factory=list)

    def __post_init__(self) -> None:
        object.__setattr__(self, "name", _enum(IconSetName, self.name))


@dataclass(frozen=True)
class CfIcon:
    """One level of a custom icon set.

    ``icon_set`` is the :class:`IconSetName` the icon comes from
    (``NO_ICONS`` for a level that draws nothing); ``icon_index`` is the
    zero-based icon within it.
    """

    icon_set: IconSetName
    icon_index: int = 0

    def __post_init__(self) -> None:
        object.__setattr__(self, "icon_set", _enum(IconSetName, self.icon_set))


@dataclass(frozen=True)
class ConditionalFormat:
    """A conditional-format rule (getter result, flattened priority order)."""

    id: str
    type: CfRuleType
    priority: int
    stop_if_true: bool
    ranges: List[CellRect]
    dxf_index: Optional[int]
    formula1: str
    formula2: str
    operator: Optional[ComparisonOperator]
    rank: Optional[int]
    percent: bool
    bottom: bool
    above_average: bool
    equal_average: bool
    std_dev: Optional[float]
    text: str
    time_period: Optional[DatePeriod]
    color_scale: Optional[ColorScale]
    data_bar: Optional[DataBar]
    icon_set: Optional[IconSet]

    def __post_init__(self) -> None:
        object.__setattr__(self, "type", _enum(CfRuleType, self.type))
        if self.operator is not None:
            object.__setattr__(self, "operator", _enum(ComparisonOperator, self.operator))
        if self.time_period is not None:
            object.__setattr__(self, "time_period", _enum(DatePeriod, self.time_period))


@dataclass
class ConditionalFormatInput:
    """Argument shape for :meth:`Sheet.add_cf_rule`.

    Set the corresponding visual payload for color-scale, data-bar, or
    icon-set rules. The C ABI validates each payload's cardinality.
    """

    #: Rectangles, unqualified A1 ranges, or one A1 string.
    ranges: Union[List[Union[CellRect, str]], str]
    type: CfRuleType
    priority: int = 0
    stop_if_true: bool = False
    id: str = ""
    dxf_index: Optional[int] = None
    formula1: str = ""
    formula2: str = ""
    operator: Optional[ComparisonOperator] = None
    rank: Optional[int] = None
    percent: bool = False
    bottom: bool = False
    above_average: bool = False
    equal_average: bool = False
    std_dev: Optional[float] = None
    text: str = ""
    time_period: Optional[DatePeriod] = None
    color_scale: Optional[ColorScale] = None
    data_bar: Optional[DataBar] = None
    icon_set: Optional[IconSet] = None


@dataclass(frozen=True)
class CellNode:
    """A workbook-wide cell coordinate (precedents / dependents trace)."""

    sheet: int
    row: int
    col: int


@dataclass(frozen=True)
class Spill:
    """Dynamic-array spill region info for a cell."""

    engaged: bool
    #: ``True`` when the cell anchors a spill that could not fill its footprint.
    blocked: bool
    #: The region including its anchor (for a blocked anchor, the footprint it attempted); ``None`` when the
    #: cell is neither part of a spill nor a blocked anchor.
    rect: Optional[CellRect]


@dataclass(frozen=True)
class ParsedReference:
    """A parsed reference: the rectangle ``rect`` and its sheet qualifier (``None`` when unqualified)."""

    sheet_name: Optional[str]
    rect: CellRect


@dataclass(frozen=True)
class ArrayFormula:
    """A legacy array (CSE) formula: the block it fills and its stored text."""

    #: The block; its top-left cell holds the formula.
    rect: CellRect
    formula: str


@dataclass(frozen=True)
class FunctionMetadata:
    """Catalog metadata for a single Formulon function."""

    name: str
    min_arity: int
    #: ``None`` denotes an unbounded variadic.
    max_arity: Optional[int]
    availability: FunctionAvailability
    signature: Optional[str]
    description: Optional[str]


@dataclass(frozen=True)
class ErrorName:
    """One error value's spelling under a profile.

    ``canonical`` is the en-invariant text and ``localized`` the profile's
    Excel spelling. ``measured`` is ``False`` when the localized spelling
    is assumed rather than captured from Excel.
    """

    canonical: str
    localized: str
    measured: bool


@dataclass(frozen=True)
class LocaleFacts:
    """Scalar locale facts of one :data:`ProfileId`.

    ``measured`` is ``True`` for the profiles captured from Excel (every
    ``mac-*`` id and ``win-365-ja_JP``) and ``False`` for the estimated
    ``win-*`` ids. Every scalar field is captured in each Mac locale, so
    ``measured`` covers them all; only ``error_names`` carries per-entry
    flags.
    ``error_names`` is indexed by error-code ordinal.
    """

    decimal_separator: str
    group_separator: str
    list_separator: str
    array_column_separator: str
    array_row_separator: str
    true_name: str
    false_name: str
    date_order: LocaleDateOrder
    currency_symbol: str
    #: ``True`` when the currency symbol follows the number.
    currency_suffix: bool
    #: ``True`` when one space separates the symbol and the number.
    currency_space: bool
    currency_default_decimals: int
    measured: bool
    #: ``True`` when the collation order of the profile was measured in Excel
    #: (every ``mac-*`` id); ``False`` for the estimated ``win-*`` ids, even
    #: ``win-365-ja_JP``, whose ``measured`` is ``True``.
    collation_measured: bool
    error_names: List[ErrorName]


@dataclass(frozen=True)
class CellXf:
    """A resolved ``<xf>`` style record.

    An alignment or protection attribute is ``None`` when the record does
    not state it; the ``<alignment>`` child exists when any alignment
    attribute is stated, and the ``<protection>`` child when ``locked`` or
    ``hidden`` is (without it a cell is locked and not hidden).
    """

    font_index: int
    fill_index: int
    border_index: int
    num_fmt_id: int
    horizontal: Optional[HorizontalAlign] = None
    vertical: Optional[VerticalAlign] = None
    wrap_text: Optional[bool] = None
    justify_last_line: Optional[bool] = None
    #: The ``<cellStyleXfs>`` entry the format inherits from.
    style_xf_index: int = 0
    text_rotation: Optional[int] = None
    indent: Optional[int] = None
    relative_indent: Optional[int] = None
    shrink_to_fit: Optional[bool] = None
    reading_order: Optional[ReadingOrder] = None
    apply_number_format: bool = False
    apply_font: bool = False
    apply_fill: bool = False
    apply_border: bool = False
    apply_alignment: bool = False
    apply_protection: bool = False
    quote_prefix: bool = False
    locked: Optional[bool] = None
    hidden: Optional[bool] = None


@dataclass
class CellXfInput:
    """An ``<xf>`` record :meth:`WorkbookOoxml.add_cell_xf` and :meth:`WorkbookOoxml.add_cell_style_xf` write.

    A ``has_*`` flag left ``None`` is inferred: ``has_horizontal``,
    ``has_vertical``, ``has_wrap_text`` and ``has_alignment`` read true
    (their values are always supplied), ``has_justify_last_line`` follows
    ``justify_last_line``, and ``has_protection`` is true when ``locked`` or
    ``hidden`` is given. An optional attribute left ``None`` is not written;
    without ``<protection>`` a cell is locked and not hidden. ``style_xf_index``
    names the ``<cellStyleXfs>`` entry the format inherits from
    (``add_cell_xf`` only).
    """

    font_index: int
    fill_index: int
    border_index: int
    num_fmt_id: int
    horizontal: HorizontalAlign
    vertical: VerticalAlign
    wrap_text: bool
    justify_last_line: Optional[bool] = None
    has_alignment: Optional[bool] = None
    has_horizontal: Optional[bool] = None
    has_vertical: Optional[bool] = None
    has_wrap_text: Optional[bool] = None
    has_justify_last_line: Optional[bool] = None
    apply_number_format: bool = False
    apply_font: bool = False
    apply_fill: bool = False
    apply_border: bool = False
    apply_alignment: bool = False
    apply_protection: bool = False
    quote_prefix: bool = False
    has_protection: Optional[bool] = None
    locked: Optional[bool] = None
    hidden: Optional[bool] = None
    #: 0..180, or 255 for vertical text.
    text_rotation: Optional[int] = None
    #: 0..255.
    indent: Optional[int] = None
    relative_indent: Optional[int] = None
    shrink_to_fit: Optional[bool] = None
    reading_order: Optional[ReadingOrder] = None
    style_xf_index: int = 0


@dataclass(frozen=True)
class Color:
    """A colour as an OOXML ``<color>`` element states it.

    ``kind`` selects the meaningful field: ``argb`` (``0xAARRGGBB``) for
    ``ColorKind.RGB``, ``theme`` and ``tint`` for ``ColorKind.THEME``,
    ``indexed`` for ``ColorKind.INDEXED``; ``ColorKind.NONE`` is no colour.
    Records read from a file keep the specification as written, so a theme
    colour stays a theme reference. An input that takes a colour also takes
    ``"#RRGGBB"`` or ``"#AARRGGBB"`` (:data:`ColorInput`).
    """

    kind: ColorKind = ColorKind.NONE
    argb: int = 0
    theme: int = 0
    tint: float = 0.0
    indexed: int = 0


@dataclass(frozen=True)
class PhoneticRun:
    """One ``<rPh>`` block: ``text`` reads the half-open span ``[start, end)``.

    Offsets are UTF-16 code units, which is how Excel indexes string
    positions. A whole-cell reading is the single run
    ``PhoneticRun(0, <length of the cell text>, kana)``.
    """

    start: int = 0
    end: int = 0
    text: str = ""


@dataclass(frozen=True)
class PhoneticProperties:
    """Result of :meth:`Sheet.get_cell_phonetic_properties`; the fields are those of ``set_cell_phonetic_properties``.

    ``present`` is true when the properties are explicit (a ``<phoneticPr>``
    element or :meth:`Sheet.set_cell_phonetic_properties`, an all-zero triple
    included), false for the state an absent element resolves to.
    """

    font_index: int = 0
    type: PhoneticType = PhoneticType.HALFWIDTH_KATAKANA
    alignment: PhoneticAlignment = PhoneticAlignment.NO_CONTROL
    present: bool = False

    def __post_init__(self) -> None:
        object.__setattr__(self, "type", _enum(PhoneticType, self.type))
        object.__setattr__(self, "alignment", _enum(PhoneticAlignment, self.alignment))


@dataclass(frozen=True)
class Font:
    """A font record (``add_font`` input / ``get_font`` result).

    A toggle (``bold``, ``italic``, ``strike``, ``outline``, ``shadow``,
    ``condense``, ``extend``) is ``None`` without its OOXML element: on a
    differential font an absent ``<b>`` means "leave the source formatting
    unchanged" while ``False`` (``<b val="0"/>``) means "switch bold off".
    ``size_pt`` (points), ``family`` and ``charset`` are ``None`` without their element;
    a ``None`` size leaves the size unchanged on a differential font, and
    every cell-font entry point writes a size. ``color`` is ``ColorKind.NONE``
    without a ``<color>`` element and takes a ``ColorInput`` on input.
    """

    name: str = ""
    size_pt: Optional[float] = 11.0
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    strike: Optional[bool] = None
    outline: Optional[bool] = None
    shadow: Optional[bool] = None
    condense: Optional[bool] = None
    extend: Optional[bool] = None
    underline: UnderlineStyle = UnderlineStyle.NONE
    vert_align: FontVerticalAlign = FontVerticalAlign.BASELINE
    family: Optional[FontFamily] = None
    charset: Optional[int] = None
    #: ``<scheme>`` theme link; ``FontScheme.NONE`` when absent.
    scheme: FontScheme = FontScheme.NONE
    color: Color = field(default_factory=Color)


@dataclass(frozen=True)
class GradientStop:
    """One ``<stop>`` of a gradient fill: a colour at ``position`` (0..1)."""

    position: float = 0.0
    color: Color = field(default_factory=Color)


@dataclass
class GradientFill:
    """A ``<gradientFill>``.

    ``degree`` is a linear gradient's angle; ``left`` / ``right`` / ``top`` /
    ``bottom`` (0..1) locate a path gradient's focus.
    """

    type: GradientType = GradientType.LINEAR
    degree: float = 0.0
    left: float = 0.0
    right: float = 0.0
    top: float = 0.0
    bottom: float = 0.0
    stops: List[GradientStop] = field(default_factory=list)


@dataclass(frozen=True)
class Fill:
    """A fill record.

    ``foreground`` and ``background`` are the pattern's colours, each a
    ``ColorInput`` on input. A ``gradient`` makes the record a
    ``<gradientFill>`` and the pattern fields are ignored.
    """

    pattern: FillPattern = FillPattern.NONE
    foreground: Color = field(default_factory=Color)
    background: Color = field(default_factory=Color)
    gradient: Optional[GradientFill] = None


@dataclass(frozen=True)
class BorderSide:
    """One side of a border: its style and colour (a ``ColorInput`` on input)."""

    style: BorderStyle = BorderStyle.NONE
    color: Color = field(default_factory=Color)


@dataclass(frozen=True)
class Border:
    """A border record of the style table: four sides, the diagonal and its directions."""

    left: BorderSide = field(default_factory=BorderSide)
    right: BorderSide = field(default_factory=BorderSide)
    top: BorderSide = field(default_factory=BorderSide)
    bottom: BorderSide = field(default_factory=BorderSide)
    diagonal: BorderSide = field(default_factory=BorderSide)
    diagonal_up: bool = False
    diagonal_down: bool = False


_BORDER_SIDES = ("left", "right", "top", "bottom", "diagonal")


@dataclass(frozen=True)
class Alignment:
    """How a cell's text is placed (OOXML ``<alignment>``).

    A format without the element reads ``GENERAL`` / ``BOTTOM``, every flag
    false and every number 0, with ``ReadingOrder.CONTEXT``.
    """

    horizontal: HorizontalAlign = HorizontalAlign.GENERAL
    vertical: VerticalAlign = VerticalAlign.BOTTOM
    wrap_text: bool = False
    shrink_to_fit: bool = False
    #: Indent level, 0 to 255.
    indent: int = 0
    #: Degrees, 0 to 180, or 255 for vertical text.
    text_rotation: int = 0
    reading_order: ReadingOrder = ReadingOrder.CONTEXT
    #: Effective with distributed horizontal alignment.
    justify_last_line: bool = False
    relative_indent: int = 0


@dataclass(frozen=True)
class CellProtection:
    """Whether a cell is locked and its formula hidden while the sheet is protected (OOXML ``<protection>``)."""

    locked: bool = True
    hidden: bool = False


@dataclass(frozen=True)
class CellFormat:
    """The format a cell shows, as :meth:`Sheet.get_cell_format` reads it.

    Read from the cell's effective XF (cell, then row, then column, then the
    default); conditional formatting is not applied. ``source`` is the layer it
    came from; the XF index is :meth:`SheetOoxml.get_cell_xf_index`.
    ``number_format`` is the format code, ``"General"`` for General;
    ``alignment`` and ``protection`` hold the defaults when the XF has no such
    element.
    """

    source: EffectiveStyleSource
    number_format: str
    font: Font
    fill: Fill
    border: Border
    alignment: Alignment
    protection: CellProtection


#: A colour in a format patch: ``"#RRGGBB"`` or ``"#AARRGGBB"`` (either case;
#: alpha defaults to ``FF``), or a :class:`Color`.
ColorInput = Union[str, Color]


class FontPatch(TypedDict, total=False):
    """The font keys of :meth:`Sheet.update_cell_format`; an omitted key keeps its value.

    ``None`` clears ``size_pt``, ``bold``, ``italic`` or ``strike`` (the element
    is removed and reads back ``None``).
    """

    name: str
    #: Points; finite and not negative.
    size_pt: Optional[float]
    bold: Optional[bool]
    italic: Optional[bool]
    strike: Optional[bool]
    underline: UnderlineStyle
    vert_align: FontVerticalAlign
    color: ColorInput


class FillPatch(TypedDict, total=False):
    """The fill keys of :meth:`Sheet.update_cell_format`.

    Setting ``foreground`` or ``background`` on a fill without a pattern makes
    it ``FillPattern.SOLID`` unless ``pattern`` is given; any fill key drops a
    gradient.
    """

    pattern: FillPattern
    foreground: ColorInput
    background: ColorInput


class BorderSidePatch(TypedDict, total=False):
    """One border side of a :class:`BorderPatch`."""

    style: BorderStyle
    color: ColorInput


class BorderPatch(TypedDict, total=False):
    """The border keys of :meth:`Sheet.update_cell_format`; each side merges into the current side."""

    left: BorderSidePatch
    right: BorderSidePatch
    top: BorderSidePatch
    bottom: BorderSidePatch
    diagonal: BorderSidePatch
    diagonal_up: bool
    diagonal_down: bool


class AlignmentPatch(TypedDict, total=False):
    """The alignment keys of :meth:`Sheet.update_cell_format`."""

    horizontal: HorizontalAlign
    vertical: VerticalAlign
    wrap_text: bool
    shrink_to_fit: bool
    #: 0 to 255.
    indent: int
    #: 0 to 180, or 255.
    text_rotation: int
    reading_order: ReadingOrder
    justify_last_line: bool
    relative_indent: int


class CellProtectionPatch(TypedDict, total=False):
    """The protection keys of :meth:`Sheet.update_cell_format`."""

    locked: bool
    hidden: bool


@dataclass(frozen=True)
class Dxf:
    """A differential format (``<dxf>``), the formatting a conditional format or table style applies.

    The record :meth:`WorkbookOoxml.get_dxf` returns. A part is ``None`` when
    the format does not set it. ``number_format`` is the format code; the
    workbook assigns its number-format id.
    """

    font: Optional[Font] = None
    fill: Optional[Fill] = None
    border: Optional[Border] = None
    number_format: Optional[str] = None
    alignment: Optional[Alignment] = None
    protection: Optional[CellProtection] = None


@dataclass
class DxfInput:
    """The differential format :meth:`WorkbookOoxml.add_dxf` takes; the parts and defaults of :class:`Dxf`."""

    font: Optional[Font] = None
    fill: Optional[Fill] = None
    border: Optional[Border] = None
    number_format: Optional[str] = None
    alignment: Optional[Alignment] = None
    protection: Optional[CellProtection] = None


def _decode_color(ptr: int) -> Color:
    return _color_from(S.COLOR_SPEC.unpack(LIB, ptr))


def _color_fields(spec: Color) -> Dict[str, object]:
    """The colour-spec record fields of ``spec``."""
    return {
        "kind": _uint(spec.kind, "kind", 8),
        "argb": _uint(spec.argb, "argb"),
        "theme": _uint(spec.theme, "theme"),
        "tint": float(spec.tint),
        "indexed": _uint(spec.indexed, "indexed"),
    }


def _color_from(d: Mapping[str, Any]) -> Color:
    """A :class:`Color` from decoded colour-spec record fields."""
    return Color(
        kind=_enum(ColorKind, d["kind"]), argb=d["argb"], theme=d["theme"], tint=d["tint"], indexed=d["indexed"]
    )


def _optional_color_from(d: Mapping[str, Any]) -> Optional[Color]:
    """:func:`_color_from`, ``None`` for an absent colour (``ColorKind.NONE``)."""
    return None if d["kind"] == ColorKind.NONE else _color_from(d)


def _pack_color(ptr: int, spec: ColorInput, label: str = "color") -> None:
    """Write the ``ColorInput`` ``spec`` as the colour-spec record at ``ptr``."""
    S.COLOR_SPEC.pack(LIB, ptr, _color_fields(color_input(spec, label)))


# The font toggles whose element may be absent, each paired with a ``has_<name>`` C flag.
_FONT_TOGGLES = ("bold", "italic", "strike", "outline", "shadow", "condense", "extend")


def _decode_font(ptr: int) -> Font:
    d = S.FONT_RECORD.unpack(LIB, ptr)
    color = _decode_color(ptr + S.FONT_RECORD.offsets["color"][1])
    return Font(
        name=LIB.read_cstr(d["name"]),
        size_pt=d["size_pt"] if d["has_size_pt"] else None,
        **{name: bool(d[name]) if d[f"has_{name}"] else None for name in _FONT_TOGGLES},
        underline=_enum(UnderlineStyle, d["underline"]),
        vert_align=_enum(FontVerticalAlign, d["vert_align"]),
        family=_enum(FontFamily, d["family"]) if d["has_family"] else None,
        charset=d["charset"] if d["has_charset"] else None,
        scheme=_enum(FontScheme, d["scheme"]),
        color=color,
    )


def _pack_font(ptr: int, record: Font, scope: _ArgScope) -> None:
    fields: Dict[str, object] = {
        "size_pt": 0.0 if record.size_pt is None else float(record.size_pt),
        "has_size_pt": 0 if record.size_pt is None else 1,
        "underline": _uint(record.underline, "underline", 8),
        "vert_align": _uint(record.vert_align, "vert_align", 8),
        "has_family": 0 if record.family is None else 1,
        "family": _uint(record.family or 0, "family", 8),
        "has_charset": 0 if record.charset is None else 1,
        "charset": _uint(record.charset or 0, "charset", 8),
        "scheme": _uint(record.scheme, "scheme", 8),
    }
    for name in _FONT_TOGGLES:
        value = getattr(record, name)
        fields[name] = 1 if value else 0
        fields[f"has_{name}"] = 0 if value is None else 1
    S.FONT_RECORD.pack(LIB, ptr, fields)
    _pack_color(ptr + S.FONT_RECORD.offsets["color"][1], record.color, "Font.color")
    S.write_str_field(LIB, ptr, S.FONT_RECORD, "name", record.name, scope)


def _decode_fill(ptr: int) -> Fill:
    d = S.FILL_RECORD.unpack(LIB, ptr)
    gradient = None
    if d["has_gradient"]:
        stop_size = S.GRADIENT_STOP.size
        stops = []
        for i in range(d["gradient_stop_count"]):
            stop_ptr = d["gradient_stops"] + i * stop_size
            stop = S.GRADIENT_STOP.unpack(LIB, stop_ptr)
            stop_color = _decode_color(stop_ptr + S.GRADIENT_STOP.offsets["color"][1])
            stops.append(GradientStop(position=stop["position"], color=stop_color))
        gradient = GradientFill(
            type=_enum(GradientType, d["gradient_type"]),
            degree=d["gradient_degree"],
            left=d["gradient_left"],
            right=d["gradient_right"],
            top=d["gradient_top"],
            bottom=d["gradient_bottom"],
            stops=stops,
        )
    return Fill(
        pattern=_enum(FillPattern, d["pattern"]),
        foreground=_decode_color(ptr + S.FILL_RECORD.offsets["foreground"][1]),
        background=_decode_color(ptr + S.FILL_RECORD.offsets["background"][1]),
        gradient=gradient,
    )


def _pack_fill(ptr: int, record: Fill, scope: _ArgScope) -> None:
    """Write ``record`` into the fill record struct at ``ptr``; the gradient
    stops are allocated through ``scope``."""
    values: Dict[str, object] = {
        "pattern": _uint(record.pattern, "pattern", 8),
    }
    gradient = record.gradient
    stops_ptr = 0
    if gradient is not None:
        stop_size = S.GRADIENT_STOP.size
        if gradient.stops:
            stops_ptr = scope.alloc(stop_size * len(gradient.stops))
            LIB.write_bytes(stops_ptr, b"\x00" * (stop_size * len(gradient.stops)))
            for i, stop in enumerate(gradient.stops):
                stop_ptr = stops_ptr + i * stop_size
                S.GRADIENT_STOP.pack(LIB, stop_ptr, {"position": float(stop.position)})
                _pack_color(stop_ptr + S.GRADIENT_STOP.offsets["color"][1], stop.color, f"GradientStop[{i}].color")
        values.update(
            has_gradient=1,
            gradient_type=_sint(gradient.type, "gradient.type"),
            gradient_degree=float(gradient.degree),
            gradient_left=float(gradient.left),
            gradient_right=float(gradient.right),
            gradient_top=float(gradient.top),
            gradient_bottom=float(gradient.bottom),
            gradient_stops=stops_ptr,
            gradient_stop_count=len(gradient.stops),
        )
    S.FILL_RECORD.pack(LIB, ptr, values)
    _pack_color(ptr + S.FILL_RECORD.offsets["foreground"][1], record.foreground, "Fill.foreground")
    _pack_color(ptr + S.FILL_RECORD.offsets["background"][1], record.background, "Fill.background")


def _decode_border_record(ptr: int) -> Border:
    sides: Dict[str, BorderSide] = {}
    for side in _BORDER_SIDES:
        side_ptr = ptr + S.BORDER_RECORD.offsets[side][1]
        d = S.BORDER_SIDE.unpack(LIB, side_ptr)
        side_color = _decode_color(side_ptr + S.BORDER_SIDE.offsets["color"][1])
        sides[side] = BorderSide(style=_enum(BorderStyle, d["style"]), color=side_color)
    outer = S.BORDER_RECORD.unpack(LIB, ptr)
    return Border(**sides, diagonal_up=bool(outer["diagonal_up"]), diagonal_down=bool(outer["diagonal_down"]))


def _pack_border_record(ptr: int, border: Border) -> None:
    if not isinstance(border, Border):
        raise TypeError(f"formulon: expected a Border, got {type(border).__name__}")
    S.BORDER_RECORD.pack(
        LIB,
        ptr,
        {"diagonal_up": 1 if border.diagonal_up else 0, "diagonal_down": 1 if border.diagonal_down else 0},
    )
    for side in _BORDER_SIDES:
        spec: BorderSide = getattr(border, side)
        side_ptr = ptr + S.BORDER_RECORD.offsets[side][1]
        S.BORDER_SIDE.pack(LIB, side_ptr, {"style": _uint(spec.style, "style", 8)})
        _pack_color(side_ptr + S.BORDER_SIDE.offsets["color"][1], spec.color, f"Border.{side}.color")


# OOXML tokens of the alignment enums, in ordinal order.
_HORIZONTAL_TOKENS = ("general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed")
_VERTICAL_TOKENS = ("top", "center", "bottom", "justify", "distributed")
# (Alignment field, OOXML attribute) of the boolean and integer alignment attributes.
_ALIGNMENT_FLAGS = (
    ("wrap_text", "wrapText"),
    ("shrink_to_fit", "shrinkToFit"),
    ("justify_last_line", "justifyLastLine"),
)
_ALIGNMENT_NUMBERS = (
    ("indent", "indent"),
    ("text_rotation", "textRotation"),
    ("relative_indent", "relativeIndent"),
    ("reading_order", "readingOrder"),
)


def _xml_flag(text: str) -> bool:
    return text in ("1", "true")


def _xml_element(text: str, tag: str) -> Optional[Any]:
    """The ``<tag>`` element a dxf child fragment holds; ``None`` for an empty fragment."""
    if not text:
        return None
    element = ElementTree.fromstring(text)
    return element if element.tag.rsplit("}", 1)[-1] == tag else None


def _alignment_from_xml(text: str) -> Optional[Alignment]:
    """Decode a dxf's ``<alignment>`` fragment; an absent attribute reads its default."""
    element = _xml_element(text, "alignment")
    if element is None:
        return None
    attrs = {key.rsplit("}", 1)[-1]: value for key, value in element.attrib.items()}
    values: Dict[str, Any] = {}
    if attrs.get("horizontal") in _HORIZONTAL_TOKENS:
        values["horizontal"] = HorizontalAlign(_HORIZONTAL_TOKENS.index(attrs["horizontal"]))
    if attrs.get("vertical") in _VERTICAL_TOKENS:
        values["vertical"] = VerticalAlign(_VERTICAL_TOKENS.index(attrs["vertical"]))
    for name, attr in _ALIGNMENT_FLAGS:
        if attr in attrs:
            values[name] = _xml_flag(attrs[attr])
    for name, attr in _ALIGNMENT_NUMBERS:
        if attr in attrs:
            values[name] = int(attrs[attr])
    if "reading_order" in values:
        values["reading_order"] = _enum(ReadingOrder, values["reading_order"])
    return Alignment(**values)


def _alignment_to_xml(alignment: Optional[Alignment]) -> Optional[str]:
    """Encode an :class:`Alignment` as a dxf ``<alignment>`` fragment, stating the attributes off their default."""
    if alignment is None:
        return None
    if not isinstance(alignment, Alignment):
        raise TypeError(f"formulon: Dxf.alignment must be an Alignment, got {type(alignment).__name__}")
    default = Alignment()
    attrs: List[str] = []
    if alignment.horizontal != default.horizontal:
        attrs.append(f'horizontal="{_HORIZONTAL_TOKENS[HorizontalAlign(alignment.horizontal)]}"')
    if alignment.vertical != default.vertical:
        attrs.append(f'vertical="{_VERTICAL_TOKENS[VerticalAlign(alignment.vertical)]}"')
    for name, attr in _ALIGNMENT_NUMBERS:
        value = getattr(alignment, name)
        if value != getattr(default, name):
            attrs.append(f'{attr}="{_sint(int(value), name)}"')
    for name, attr in _ALIGNMENT_FLAGS:
        if getattr(alignment, name):
            attrs.append(f'{attr}="1"')
    return "<alignment" + "".join(" " + a for a in attrs) + "/>"


def _protection_from_xml(text: str) -> Optional[CellProtection]:
    """Decode a dxf's ``<protection>`` fragment; an absent attribute reads its default."""
    element = _xml_element(text, "protection")
    if element is None:
        return None
    attrs = {key.rsplit("}", 1)[-1]: value for key, value in element.attrib.items()}
    default = CellProtection()
    return CellProtection(
        locked=_xml_flag(attrs["locked"]) if "locked" in attrs else default.locked,
        hidden=_xml_flag(attrs["hidden"]) if "hidden" in attrs else default.hidden,
    )


def _protection_to_xml(protection: Optional[CellProtection]) -> Optional[str]:
    """Encode a :class:`CellProtection` as a dxf ``<protection>`` fragment."""
    if protection is None:
        return None
    if not isinstance(protection, CellProtection):
        raise TypeError(f"formulon: Dxf.protection must be a CellProtection, got {type(protection).__name__}")
    return f'<protection locked="{1 if protection.locked else 0}" hidden="{1 if protection.hidden else 0}"/>'


@dataclass(frozen=True)
class CellStyle:
    """A named cell style (``<cellStyle>`` entry); ``style_xf_index`` is its ``<cellStyleXfs>`` entry."""

    name: str
    style_xf_index: int
    builtin_id: int = CELL_STYLE_BUILTIN_ID_NONE
    i_level: int = 0
    hidden: bool = False
    custom_builtin: bool = False


@dataclass(frozen=True)
class ThemeFonts:
    """The theme's major (heading) and minor (body) typefaces."""

    major_latin: str = ""
    major_east_asian: str = ""
    minor_latin: str = ""
    minor_east_asian: str = ""


@dataclass(frozen=True)
class ThemeScriptFont:
    """One ``a:font`` entry of the major (``FontScheme.MAJOR``, heading) or minor (``FontScheme.MINOR``, body) theme font set.

    ``script`` is an ISO 15924 code (``Jpan``) and ``typeface`` its font.
    """

    scheme: FontScheme
    script: str
    typeface: str


@dataclass(frozen=True)
class ThemeColors:
    """The theme's 12 scheme colours as ``0xFFRRGGBB`` values in ``a:clrScheme`` order.

    ``source`` says whether they come from the theme part or are the default
    theme's (the part is absent or unparseable). This is not the order of a
    colour's theme index, which swaps the first two pairs (0 = lt1, 1 = dk1,
    2 = lt2, 3 = dk2).
    """

    source: ThemeSource
    colors: List[int]


@dataclass(frozen=True)
class ResolvedColor:
    """A colour resolved to ``0xAARRGGBB`` and how it was resolved."""

    argb: int
    resolution: ColorResolution


@dataclass(frozen=True)
class EffectiveStyle:
    """The formatting a cell shows, read from its selected xf record.

    ``borders`` holds left, right, top, bottom, diagonal.
    """

    xf_index: int
    source: EffectiveStyleSource
    font_index: int
    fill_index: int
    border_index: int
    font_color: ResolvedColor
    fill_foreground: ResolvedColor
    fill_background: ResolvedColor
    borders: List[ResolvedColor]
    protection: CellProtection
    #: The number format code.
    number_format: str


@dataclass(frozen=True)
class ExternalLink:
    """An external-link record (``<externalReferences>`` entry)."""

    index: int
    rel_id: str
    part_path: str
    target: str
    target_external: bool
    kind: ExternalLinkKind


@dataclass(frozen=True)
class PivotCell:
    """One projected PivotTable layout cell."""

    row: int
    col: int
    #: A date item label reads as its serial number; ``number_format`` tells it is a date.
    value: NativeValue
    kind: PivotCellKind
    depth: int
    field_name: str
    number_format: str


@dataclass(frozen=True)
class PivotLayout:
    """The projected rectangular layout of a PivotTable."""

    #: The rectangle the layout covers in absolute sheet coordinates; ``None`` when it spans no cell.
    rect: Optional[CellRect]
    cells: List[PivotCell]


@dataclass(frozen=True)
class PivotWorksheetSource:
    """Optional worksheet-source metadata for a pivot cache."""

    ref: Optional[str] = None
    sheet: Optional[str] = None
    name: Optional[str] = None


@dataclass
class PivotInput:
    """A pivot table :meth:`Sheet.add_pivot` creates: its name, its cache and its top-left cell."""

    name: str
    cache_id: int
    #: The top-left cell, or its A1 text (``"C3"``).
    anchor: Union[CellAddress, str]


@dataclass(frozen=True)
class PivotProperties:
    """A pivot table's name and the id of its cache, as :meth:`Pivot.get_properties` returns them."""

    name: str
    cache_id: int


@dataclass(frozen=True)
class PivotCacheField:
    """One pivot-cache field; ``formula`` (with a leading ``=``) is ``None`` unless the field is calculated."""

    name: str
    formula: Optional[str]


@dataclass(frozen=True)
class PivotCacheSharedItem:
    """One shared item of a pivot-cache field; a date item reads as its serial number and has ``is_date`` set."""

    value: NativeValue
    is_date: bool


@dataclass
class PivotFieldInput:
    """Argument shape for :meth:`Pivot.add_field`.

    ``subtotal_top`` defaults to ``True``, Excel's default: a group's subtotal
    sits on its own heading row. ``False`` heads the group with a label-only
    row and closes it with an ``"<item> Total"`` row.
    """

    source_name: str
    custom_name: str = ""
    axis: PivotAxis = PivotAxis.ROW
    subtotal_top: bool = True
    number_format: str = ""


@dataclass(frozen=True)
class PivotField:
    """A pivot field's definition, as read by :meth:`Pivot.list_fields`.

    ``axis`` is ``None`` for a field that sits on no axis. ``cache_field_index``
    is the bound pivot-cache field (pivot field ``i`` binds to cache field
    ``i``), ``None`` when the cache does not cover the field. For a date-grouped
    field, ``date_periods`` holds the periods of its group fields (empty when
    ungrouped) and ``date_calendar`` their calendar; ``group_base_index`` is
    the field a group field groups, ``None`` for any other field. A text the
    field does not state is ``None``.
    ``subtotal_functions`` lists the custom subtotal functions in ascending
    aggregation order; ``default_subtotal`` is false whenever it is non-empty,
    since custom functions replace the automatic subtotal.
    """

    source_name: str
    custom_name: Optional[str]
    axis: Optional[PivotAxis]
    subtotal_top: bool
    number_format: Optional[str]
    cache_field_index: Optional[int]
    sort_order: PivotSortOrder
    sort_by_field: Optional[str]
    date_periods: PivotDatePeriod
    date_calendar: PivotCalendar
    group_base_index: Optional[int]
    default_subtotal: bool
    subtotal_functions: List[PivotAggregation]
    insert_blank_row: bool = False
    fill_down_labels: bool = False
    show_all: bool = False


@dataclass(frozen=True)
class PivotItem:
    """One pivot-field item, as read by :meth:`Pivot.get_field_item`.

    ``name`` is the item's own label (``None`` for an item bound only by cache
    index); ``cache_index`` indexes the bound cache field's shared items and
    is ``None`` for an item bound by label only. ``caption`` is the text shown
    instead of the item's value (``None`` when it has none); ``show_detail``
    is false for a collapsed item.
    """

    name: Optional[str]
    cache_index: Optional[int]
    visible: bool
    caption: Optional[str] = None
    show_detail: bool = True


@dataclass(frozen=True)
class PivotStyleOptions:
    """A pivot table's style, as read by :meth:`Pivot.get_style_options`.

    ``style_name`` is ``None`` when the style element carries no style name.
    :meth:`Pivot.update_style_options` takes these fields as keyword arguments.
    """

    style_name: Optional[str]
    show_row_headers: bool
    show_column_headers: bool
    show_row_stripes: bool
    show_column_stripes: bool
    show_last_column: bool


@dataclass(frozen=True)
class PivotOptions:
    """The PivotTable Options of one pivot table, as read by :meth:`Pivot.get_options`.

    A read fills every member; a string option reads ``None`` when unset.
    :meth:`Pivot.update_options` takes these fields as keyword arguments.
    A ``missing_caption`` of exactly ``-`` is written as the number 0.
    ``indent`` runs 0-126 and ``page_wrap`` 0-255 (0 does not wrap).
    """

    show_error: bool
    error_caption: Optional[str]
    show_missing: bool
    missing_caption: Optional[str]
    show_headers: bool
    show_drill: bool
    show_data_tips: bool
    grid_drop_zones: bool
    show_empty_row: bool
    show_empty_column: bool
    show_data_drop_down: bool
    print_drill: bool
    field_print_titles: bool
    item_print_titles: bool
    merge_item: bool
    indent: int
    page_over_then_down: bool
    page_wrap: int
    multiple_field_filters: bool
    preserve_formatting: bool
    custom_list_sort: bool
    alt_text: Optional[str]
    alt_text_summary: Optional[str]
    grand_total_caption: Optional[str]
    auto_formatting: bool
    enable_drill: bool


@dataclass(frozen=True)
class PivotCacheOptions:
    """The options of one pivot cache, as read by :meth:`PivotCache.get_options`.

    ``missing_items_limit`` is ``-1`` for automatic, ``0`` for none, or up to 1048576.
    """

    save_data: bool
    missing_items_limit: int
    refresh_on_load: bool


@dataclass(frozen=True)
class PivotCalculatedItem:
    """One calculated item, as read by :meth:`PivotCache.list_calculated_items`.

    ``formula`` has a leading ``=``; ``scoped`` is true for an entry read from
    a file that applies the formula only where the item meets items of other fields.
    """

    field_index: int
    name: str
    formula: str
    scoped: bool


@dataclass
class PivotDataFieldInput:
    """Argument shape for :meth:`Pivot.add_data_field` / :meth:`Pivot.set_data_field`.

    An empty or ``None`` ``name`` asks the engine to name the data field from
    the workbook profile's pivot wording (``"Sum of Sales"`` under en-US,
    ``"合計 / Sales"`` under ja-JP); that needs ``field_index`` to be in range.
    ``show_as_base_field`` / ``show_as_base_item`` are ``None`` when unset.
    """

    name: Optional[str]
    field_index: int
    aggregation: PivotAggregation = PivotAggregation.SUM
    number_format: Optional[str] = None
    show_as: PivotShowValuesAs = PivotShowValuesAs.NORMAL
    show_as_base_field: Optional[int] = None
    show_as_base_item: Optional[int] = None


@dataclass
class PivotFilterInput:
    """Argument shape for :meth:`Pivot.add_filter`.

    ``value`` is the operand: a number for a value filter (an integer count
    for the top / bottom count types) or a date filter, a ``str`` for a label
    filter, ``None`` for a relative-period type. ``value_high`` is the upper
    bound of a range type (a ``str`` for ``LABEL_BETWEEN`` / ``LABEL_NOT_BETWEEN``),
    ``None`` otherwise. A ``bool`` is rejected.
    """

    axis: PivotAxis
    field_name: str
    type: PivotFilterKind
    value: Union[int, float, str, None] = None
    value_high: Union[int, float, str, None] = None
    data_field_index: int = 0


def _nullable_enum(value: Optional[int], name: str) -> int:
    """The C ordinal of an enum argument whose absence is ``None``; the C ``-1`` stays internal."""
    if value is None:
        return -1
    ordinal = _sint(value, name)
    if ordinal < 0:
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT,
            _diagnostic_override=(f"{name} {ordinal} is not a member; pass None for none", f"{name}={ordinal}"),
        )
    return ordinal


def _local_sheet_id(value: Optional[int]) -> int:
    """The C ``local_sheet_id``: ``-1`` for workbook scope (``None``), else the sheet index."""
    if value is None:
        return -1
    index = _uint(value, "local_sheet_id", bits=64)
    if index > 0x7FFFFFFF:
        raise FormulonError(
            StatusCode.SHEET_INDEX_OUT_OF_RANGE,
            _diagnostic_override=(f"local_sheet_id {index} is out of range", f"local_sheet_id={index}"),
        )
    return index


def _print_titles_from(d: Dict[str, int]) -> PrintTitles:
    """Decode the C print-titles record; an axis whose flag is clear reads ``None``."""
    return PrintTitles(
        rows=Span(d["first_row"], d["last_row"]) if d["has_rows"] else None,
        columns=Span(d["first_col"], d["last_col"]) if d["has_columns"] else None,
    )


@dataclass(frozen=True)
class PivotDataField:
    """One data field of a pivot, as :meth:`Pivot.list_data_fields` reads it.

    The fields read as in :class:`PivotDataFieldInput`; ``number_format``,
    ``show_as_base_field`` and ``show_as_base_item`` are ``None`` when unset.
    """

    name: str
    field_index: int
    aggregation: PivotAggregation
    number_format: Optional[str]
    show_as: PivotShowValuesAs
    show_as_base_field: Optional[int]
    show_as_base_item: Optional[int]


@dataclass(frozen=True)
class PivotFilter:
    """One stored filter of a pivot, as :meth:`Pivot.list_filters` reads it.

    The fields read as in :class:`PivotFilterInput`; ``value`` and
    ``value_high`` are ``None`` when the filter has no such operand.
    """

    axis: PivotAxis
    field_name: str
    type: PivotFilterKind
    value: Optional[Union[float, str]]
    value_high: Optional[Union[float, str]]
    data_field_index: int


# The C operand-kind discriminator of a pivot filter operand.
_OPERAND_NONE = -1
_OPERAND_INT = 0
_OPERAND_DOUBLE = 1
_OPERAND_TEXT = 2
_COUNT_FILTER_KINDS = (PivotFilterKind.VALUE_TOP_COUNT, PivotFilterKind.VALUE_BOTTOM_COUNT)


def _operand_fields(value: object, name: str, *, count_filter: bool) -> Dict[str, object]:
    """The ``<name>_kind`` / ``_int`` / ``_double`` / text slots one filter operand fills."""
    if value is None:
        return {name + "_kind": _OPERAND_NONE}
    if isinstance(value, str):
        return {name + "_kind": _OPERAND_TEXT, name + "_text": value}
    if isinstance(value, bool) or not isinstance(value, numbers.Real):
        raise TypeError(f"formulon: {name} must be a number, a str or None, got {type(value).__name__}")
    if count_filter and math.isfinite(value) and value == int(value) and abs(value) < (1 << 31):
        return {name + "_kind": _OPERAND_INT, name + "_int": int(value)}
    return {name + "_kind": _OPERAND_DOUBLE, name + "_double": float(value)}


def _encode_pivot_filter_operands(spec: PivotFilterInput) -> Dict[str, object]:
    """The operand slots of the C filter spec for ``spec.value`` / ``spec.value_high``.

    A text operand lands in ``value_text`` / ``value_high_text``; those two keys are
    returned only when used, so the caller writes each as a string field or leaves it ``NULL``.
    """
    count_filter = spec.type in _COUNT_FILTER_KINDS
    return {
        **_operand_fields(spec.value, "value", count_filter=count_filter),
        **_operand_fields(spec.value_high, "value_high", count_filter=False),
    }


def _decode_pivot_filter_operand(d: Dict[str, object], name: str) -> Optional[Union[float, str]]:
    """One filter operand of a decoded C filter spec; ``None`` when the slot is unset."""
    kind = d[name + "_kind"]
    if kind == _OPERAND_INT:
        return float(d[name + "_int"])
    if kind == _OPERAND_DOUBLE:
        return d[name + "_double"]
    if kind == _OPERAND_TEXT:
        return LIB.read_cstr(d[name + "_text"])
    return None


def _cell_xf_fields(record: CellXfInput) -> Dict[str, object]:
    """Project a :class:`CellXfInput` onto the cell XF struct field map.

    Shared by the direct ``<xf>`` writer and the named-style ``<xf>``
    writer, which pack the same struct, so the presence-flag defaulting
    lives in one place.
    """
    if not isinstance(record, CellXfInput):
        raise TypeError(f"formulon: expected a CellXfInput, got {type(record).__name__}")

    def stated(flag: Optional[bool], inferred: bool) -> int:  # noqa: FBT001 -- private presence-flag encoder
        return 1 if (inferred if flag is None else flag) else 0

    return {
        "font_index": _uint(record.font_index, "font_index"),
        "fill_index": _uint(record.fill_index, "fill_index"),
        "border_index": _uint(record.border_index, "border_index"),
        "num_fmt_id": _uint(record.num_fmt_id, "num_fmt_id", 16),
        "horizontal": _uint(record.horizontal, "horizontal", 8),
        "has_horizontal": stated(record.has_horizontal, True),
        "vertical": _uint(record.vertical, "vertical", 8),
        "has_vertical": stated(record.has_vertical, True),
        "wrap_text": 1 if record.wrap_text else 0,
        "has_wrap_text": stated(record.has_wrap_text, True),
        "justify_last_line": 1 if record.justify_last_line else 0,
        "has_justify_last_line": stated(record.has_justify_last_line, record.justify_last_line is not None),
        "has_alignment": stated(record.has_alignment, True),
        "style_xf_index": _uint(record.style_xf_index, "style_xf_index"),
        "has_text_rotation": 1 if record.text_rotation is not None else 0,
        "text_rotation": _uint(record.text_rotation or 0, "text_rotation"),
        "has_indent": 1 if record.indent is not None else 0,
        "indent": _uint(record.indent or 0, "indent"),
        "has_relative_indent": 1 if record.relative_indent is not None else 0,
        "relative_indent": _sint(record.relative_indent or 0, "relative_indent"),
        "has_shrink_to_fit": 1 if record.shrink_to_fit is not None else 0,
        "shrink_to_fit": 1 if record.shrink_to_fit else 0,
        "has_reading_order": 1 if record.reading_order is not None else 0,
        "reading_order": _sint(record.reading_order or 0, "reading_order"),
        "apply_number_format": 1 if record.apply_number_format else 0,
        "apply_font": 1 if record.apply_font else 0,
        "apply_fill": 1 if record.apply_fill else 0,
        "apply_border": 1 if record.apply_border else 0,
        "apply_alignment": 1 if record.apply_alignment else 0,
        "apply_protection": 1 if record.apply_protection else 0,
        "quote_prefix": 1 if record.quote_prefix else 0,
        "has_protection": stated(record.has_protection, record.locked is not None or record.hidden is not None),
        "locked": 0 if record.locked is False else 1,
        "hidden": 1 if record.hidden else 0,
    }


# ---------------------------------------------------------------------------
# Workbook handle
# ---------------------------------------------------------------------------


def _tristate_default(flag: Optional[bool]) -> int:  # noqa: FBT001 -- private tri-state encoder
    """Map an optional flag to ``fm_tristate_t``: ``None`` is ``DEFAULT`` (Excel's default), then ``FALSE`` / ``TRUE``."""
    if flag is None:
        return 0
    return 2 if flag else 1


def _tristate(flag: Optional[bool]) -> int:  # noqa: FBT001 -- private tri-state encoder
    """Map an optional flag to the ABI's preserve / clear / set encoding.

    ``None`` becomes ``-1`` (leave the stored flag alone); ``False`` and
    ``True`` become ``0`` and ``1``.
    """
    if flag is None:
        return -1
    return 1 if flag else 0


class _ArgScope:
    """Owns every engine-memory allocation made to marshal one call's arguments.

    Enter it before the first allocation; every block allocated through it,
    packers included, is freed when the ``with`` block exits, normally or by
    an exception. It is also the one place a Python string becomes a C
    string pointer: ``None`` is the null pointer and ``""`` a real empty string,
    except through :meth:`nonempty_cstr`, for a slot whose C contract reads
    an empty string as absent.
    """

    __slots__ = ("_ptrs",)

    def __init__(self) -> None:
        self._ptrs: List[int] = []

    def __enter__(self) -> "_ArgScope":
        return self

    def __exit__(self, *exc_info: object) -> None:
        ptrs, self._ptrs = self._ptrs, []
        for ptr in ptrs:
            LIB.free(ptr)

    def _take(self, size: int) -> int:
        ptr = LIB.alloc(size)
        if ptr:
            self._ptrs.append(ptr)
        return ptr

    def alloc(self, size: int) -> int:
        """Allocate ``size`` zeroed bytes; 0 when ``size`` is 0."""
        ptr = self._take(size)
        if ptr:
            LIB.write_bytes(ptr, b"\x00" * size)
        return ptr

    def out_ptr(self) -> int:
        """Allocate a zeroed 4-byte slot for an out-i32 / out-pointer."""
        return self.alloc(4)

    def out_u64(self) -> int:
        """Allocate a zeroed 8-byte slot for an out-uint64."""
        return self.alloc(8)

    def struct(self, layout: S.Struct) -> int:
        """Allocate one zeroed ``layout`` struct; a record carries its ``struct_size``."""
        ptr = self.alloc(layout.size)
        if layout.is_record:
            LIB.write_bytes(ptr, layout.blank())
        return ptr

    def struct_array(self, layout: S.Struct, count: int) -> int:
        """Allocate a zeroed contiguous array of ``count`` ``layout`` structs."""
        ptr = self.alloc(layout.size * count)
        if ptr and layout.is_record:
            LIB.write_bytes(ptr, bytes(layout.blank()) * count)
        return ptr

    def value_array(self, count: int) -> int:
        """Allocate a zeroed contiguous array of ``count`` value-struct slots."""
        return self.alloc(fm_value_t_size * count)

    def buffer(self, data: bytes) -> int:
        """Copy ``data`` into engine memory; 0 when it is empty."""
        ptr = self._take(len(data))
        if ptr:
            LIB.write_bytes(ptr, data)
        return ptr

    def cstr(self, value: str) -> int:
        """Copy ``value`` as a NUL-terminated UTF-8 string, ``""`` included.

        A lone surrogate becomes U+FFFD and a surrogate pair its code point,
        as the JavaScript bindings convert a UTF-16 string.
        """
        if not isinstance(value, str):
            raise TypeError(f"expected str, got {type(value).__name__}")
        try:
            encoded = value.encode("utf-8")
        except UnicodeEncodeError:
            utf16 = value.encode("utf-16-le", "surrogatepass")
            encoded = utf16.decode("utf-16-le", "replace").encode("utf-8")
        return self.buffer(encoded + b"\x00")

    def opt_cstr(self, value: Optional[str]) -> int:
        """The null pointer for ``None``; otherwise :meth:`cstr`."""
        return 0 if value is None else self.cstr(value)

    def nonempty_cstr(self, value: Optional[str]) -> int:
        """The null pointer for ``None`` and ``""``; otherwise :meth:`cstr`."""
        return 0 if value is None or value == "" else self.cstr(value)


def _read_image_info(ptr: int) -> ImageInfo:
    """Decode an image info struct."""
    d = S.IMAGE_INFO.unpack(LIB, ptr)
    return ImageInfo(format=ImageFormat(d["format"]), width_px=d["width_px"], height_px=d["height_px"])


def _read_count(fn, *args) -> int:
    """Call a ``(... , out_count*)`` ABI function and return the count."""
    with _ArgScope() as scope:
        out = scope.out_ptr()
        _check(fn(*args, out))
        return LIB.read_u32(out)


def _read_cstr_out(fn, *args) -> str:
    """Call a ``(..., out_cstr*)`` ABI function and decode its result."""
    with _ArgScope() as scope:
        out = scope.out_ptr()
        _check(fn(*args, out))
        return LIB.read_cstr(LIB.read_u32(out))


def _read_optional_cstr_out(fn, *args) -> Optional[str]:
    """Call a ``(..., out_present*, out_cstr*)`` ABI function; ``None`` when nothing is present."""
    with _ArgScope() as scope:
        present = scope.out_ptr()
        out = scope.out_ptr()
        _check(fn(*args, present, out))
        return LIB.read_cstr(LIB.read_u32(out)) if LIB.read_u32(present) else None


def _finite(value: Any, name: str) -> float:
    """``value`` as a finite ``float``; ``TypeError`` for a non-number, ``InvalidArgumentError`` for NaN / infinity."""
    if isinstance(value, bool) or not isinstance(value, numbers.Real):
        raise TypeError(f"formulon: {name} must be a number, got {type(value).__name__}")
    x = float(value)
    if not math.isfinite(x):
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT, _diagnostic_override=(f"{name} must be finite", f"{name}={x!r}")
        )
    return x


_SHAPE_FIELDS = {
    "name": 0x1,
    "preset_geometry": 0x2,
    "anchor": 0x4,
    "text": 0x8,
    "text_box": 0x10,
    "fill": 0x20,
    "line": 0x40,
    "line_width": 0x80,
    "rotation": 0x100,
    "hidden": 0x200,
}


def _pack_shape(
    scope: _ArgScope,
    name,
    preset_geometry,
    text,
    text_box,
    anchor_kind,
    edit_as,
    from_row,
    from_col,
    from_row_offset_emu,
    from_col_offset_emu,
    to_row,
    to_col,
    to_row_offset_emu,
    to_col_offset_emu,
    width_emu,
    height_emu,
    fill_kind,
    fill_color,
    line_kind,
    line_color,
    line_width_pt,
    rotation_deg,
    hidden,
) -> int:
    """Build a shape input struct; a ``None`` argument is a field left out of the mask.

    ``fill_color`` / ``line_color`` are the ``ColorInput`` of a solid kind, ``None`` for any other;
    a colour without its kind raises ``TypeError``. ``line_width_pt`` is in points and
    ``rotation_deg`` in degrees.
    """
    for kind_name, kind, color_name, color in (
        ("fill_kind", fill_kind, "fill_color", fill_color),
        ("line_kind", line_kind, "line_color", line_color),
    ):
        if kind is None and color is not None:
            raise TypeError(f"formulon: {color_name} needs {kind_name}")
    ptr = scope.struct_array(S.SHAPE, 1)
    mask = 0
    values: Dict[str, Any] = {}
    if text_box is not None:
        mask |= _SHAPE_FIELDS["text_box"]
        values["text_box"] = 1 if text_box else 0
    if anchor_kind is not None:
        mask |= _SHAPE_FIELDS["anchor"]
        values.update(
            anchor_kind=_sint(anchor_kind, "anchor_kind"),
            edit_as=_sint(edit_as, "edit_as"),
            from_row=_uint(from_row, "from_row"),
            from_col=_uint(from_col, "from_col"),
            from_row_offset_emu=_sint(from_row_offset_emu, "from_row_offset_emu", 64),
            from_col_offset_emu=_sint(from_col_offset_emu, "from_col_offset_emu", 64),
            to_row=_uint(to_row, "to_row"),
            to_col=_uint(to_col, "to_col"),
            to_row_offset_emu=_sint(to_row_offset_emu, "to_row_offset_emu", 64),
            to_col_offset_emu=_sint(to_col_offset_emu, "to_col_offset_emu", 64),
            width_emu=_sint(width_emu, "width_emu", 64),
            height_emu=_sint(height_emu, "height_emu", 64),
        )
    if fill_kind is not None:
        mask |= _SHAPE_FIELDS["fill"]
        values.update(fill_kind=_sint(fill_kind, "fill_kind"))
        values["fill_color"] = _color_fields(optional_color_input(fill_color, "fill_color"))
    if line_kind is not None:
        mask |= _SHAPE_FIELDS["line"]
        values.update(line_kind=_sint(line_kind, "line_kind"))
        values["line_color"] = _color_fields(optional_color_input(line_color, "line_color"))
    if line_width_pt is not None:
        mask |= _SHAPE_FIELDS["line_width"]
        values["line_width_pt"] = _finite(line_width_pt, "line_width_pt")
    if rotation_deg is not None:
        mask |= _SHAPE_FIELDS["rotation"]
        values["rotation_deg"] = _finite(rotation_deg, "rotation_deg")
    if hidden is not None:
        mask |= _SHAPE_FIELDS["hidden"]
        values["hidden"] = 1 if hidden else 0
    for key, text_value in (("name", name), ("preset_geometry", preset_geometry), ("text", text)):
        if text_value is not None:
            mask |= _SHAPE_FIELDS[key]
            S.write_str_field(LIB, ptr, S.SHAPE, key, text_value, scope, allow_empty=True)
    values["fields"] = mask
    # `write_str_field` patched the pointer slots in place; pack the scalars around them.
    raw = bytearray(LIB.read_bytes(ptr, S.SHAPE.size))
    S.SHAPE.pack_into(raw, 0, values)
    LIB.write_bytes(ptr, bytes(raw))
    return ptr


_SLICER_FIELDS = {
    "name": 0x1,
    "caption": 0x2,
    "style": 0x4,
    "column_count": 0x8,
    "row_height_emu": 0x10,
    "anchor": 0x20,
    "sort_order": 0x40,
    "cross_filter": 0x80,
}


def _pack_slicer(
    scope: _ArgScope,
    name,
    caption,
    style,
    column_count,
    row_height_emu,
    anchor,
    sort_order,
    cross_filter,
) -> int:
    """Build a slicer input struct; a ``None`` argument is a field left out of the mask.

    ``anchor`` is ``None`` or ``(from_row, from_col, to_row, to_col)`` followed by
    ``(from_row_offset_emu, from_col_offset_emu, to_row_offset_emu, to_col_offset_emu)``.
    """
    ptr = scope.struct_array(S.SLICER, 1)
    mask = 0
    values: Dict[str, int] = {}
    if column_count is not None:
        mask |= _SLICER_FIELDS["column_count"]
        values["column_count"] = _uint(column_count, "column_count")
    if row_height_emu is not None:
        mask |= _SLICER_FIELDS["row_height_emu"]
        values["row_height_emu"] = _sint(row_height_emu, "row_height_emu", 64)
    if anchor is not None:
        mask |= _SLICER_FIELDS["anchor"]
        corner, offsets = tuple(anchor[:4]), tuple(anchor[4:]) or (0, 0, 0, 0)
        if len(corner) != 4 or len(offsets) != 4:
            raise TypeError("anchor must be (from_row, from_col, to_row, to_col) plus optional EMU offsets")
        values.update(
            from_row=_uint(corner[0], "from_row"),
            from_col=_uint(corner[1], "from_col"),
            to_row=_uint(corner[2], "to_row"),
            to_col=_uint(corner[3], "to_col"),
            from_row_offset_emu=_sint(offsets[0], "from_row_offset_emu", 64),
            from_col_offset_emu=_sint(offsets[1], "from_col_offset_emu", 64),
            to_row_offset_emu=_sint(offsets[2], "to_row_offset_emu", 64),
            to_col_offset_emu=_sint(offsets[3], "to_col_offset_emu", 64),
        )
    if sort_order is not None:
        mask |= _SLICER_FIELDS["sort_order"]
        values["sort_order"] = _sint(sort_order, "sort_order")
    if cross_filter is not None:
        mask |= _SLICER_FIELDS["cross_filter"]
        values["cross_filter"] = _sint(cross_filter, "cross_filter")
    for key, text_value in (("name", name), ("caption", caption), ("style", style)):
        if text_value is not None:
            mask |= _SLICER_FIELDS[key]
            S.write_str_field(LIB, ptr, S.SLICER, key, text_value, scope, allow_empty=True)
    values["fields"] = mask
    # `write_str_field` patched the pointer slots in place; pack the scalars around them.
    raw = bytearray(LIB.read_bytes(ptr, S.SLICER.size))
    S.SLICER.pack_into(raw, 0, values)
    LIB.write_bytes(ptr, bytes(raw))
    return ptr


_SPARKLINE_COLORS = (
    "color_series",
    "color_negative",
    "color_axis",
    "color_markers",
    "color_first",
    "color_last",
    "color_high",
    "color_low",
)


def _pack_sparkline_group(scope: _ArgScope, group: SparklineGroup) -> int:
    """Build a sparkline group input struct from ``group`` (``sparklines`` is passed separately)."""
    ptr = scope.struct_array(S.SPARKLINE_GROUP, 1)
    S.SPARKLINE_GROUP.pack(
        LIB,
        ptr,
        {
            "type": _sint(group.type, "type"),
            "markers": 1 if group.markers else 0,
            "high": 1 if group.high else 0,
            "low": 1 if group.low else 0,
            "first": 1 if group.first else 0,
            "last": 1 if group.last else 0,
            "negative": 1 if group.negative else 0,
            "display_x_axis": 1 if group.display_x_axis else 0,
            "display_hidden": 1 if group.display_hidden else 0,
            "min_axis_type": _sint(group.min_axis_type, "min_axis_type"),
            "max_axis_type": _sint(group.max_axis_type, "max_axis_type"),
            "has_manual_min": 0 if group.manual_min is None else 1,
            "has_manual_max": 0 if group.manual_max is None else 1,
            "manual_min": float(group.manual_min or 0.0),
            "manual_max": float(group.manual_max or 0.0),
            "date_axis": 1 if group.date_axis else 0,
            "line_weight": float(group.line_weight),
            "empty_cells_as": _sint(group.empty_cells_as, "empty_cells_as"),
        },
    )
    for key in _SPARKLINE_COLORS:
        _pack_color(ptr + S.SPARKLINE_GROUP.offsets[key][1], getattr(group, key), f"SparklineGroup.{key}")
    S.write_str_field(LIB, ptr, S.SPARKLINE_GROUP, "date_axis_ref", group.date_axis_ref, scope)
    return ptr


def _pack_sparklines(scope: _ArgScope, sparklines: Sequence[Sparkline]) -> int:
    """Build a sparkline struct array; an empty list still gets a valid block (the core rejects a zero count)."""
    ptr = scope.struct_array(S.SPARKLINE, max(len(sparklines), 1))
    for i, item in enumerate(sparklines):
        base = ptr + i * S.SPARKLINE.size
        S.write_str_field(LIB, base, S.SPARKLINE, "formula", item.formula, scope)
        S.write_str_field(LIB, base, S.SPARKLINE, "location", item.location, scope)
    return ptr


def _image_anchor_fields(
    anchor_kind,
    edit_as,
    row,
    col,
    row_offset_emu,
    col_offset_emu,
    width_emu,
    height_emu,
) -> Dict[str, int]:
    """Convert the shared image-anchor fields to their ABI representation."""
    return {
        "anchor_kind": _sint(anchor_kind, "anchor_kind"),
        "edit_as": _sint(edit_as, "edit_as"),
        "row": _uint(row, "row"),
        "col": _uint(col, "col"),
        "row_offset_emu": _sint(row_offset_emu, "row_offset_emu", 64),
        "col_offset_emu": _sint(col_offset_emu, "col_offset_emu", 64),
        "width_emu": _sint(width_emu, "width_emu", 64),
        "height_emu": _sint(height_emu, "height_emu", 64),
    }


# ---------------------------------------------------------------------------
# Chart marshalling
# ---------------------------------------------------------------------------

# `fm_chart_patch_op_t`: a zeroed patch field keeps, so only CLEAR and SET are written.
_PATCH_CLEAR = 1
_PATCH_SET = 2

# The C "no title" / "no legend" members; the binding reports them as a ``None`` record.
_CHART_TITLE_NONE = 0
_CHART_LEGEND_NONE = -1
# The series roles of `fm_chart_role_t` in order, which is also the order of `fm_chart_data_series_t.range_count`.
_CHART_ROLES = ("name", "categories", "values", "x", "y", "size")

# (patch key, C field, scalar kind) per patch struct.
_FILL_PATCH_SPEC = (("color", "color", "color"), ("transparency_pct", "transparency_pct", "f64"))
_LINE_PATCH_SPEC = (
    ("color", "color", "color"),
    ("transparency_pct", "transparency_pct", "f64"),
    ("width_pt", "width_pt", "f64"),
    ("dash", "dash", "i32"),
    ("none", "none", "bool"),
)
_MARKER_PATCH_SPEC = (("symbol", "symbol", "i32"), ("size_pt", "size_pt", "f64"))
_LABELS_PATCH_SPEC = (
    ("show_value", "show_value", "bool"),
    ("show_category", "show_category", "bool"),
    ("show_series", "show_series", "bool"),
    ("show_percent", "show_percent", "bool"),
    ("show_legend_key", "show_legend_key", "bool"),
    ("separator", "separator", "str"),
    ("position", "position", "i32"),
    ("number_format", "number_format", "str"),
    ("font_color", "font_color", "color"),
    ("font_size_pt", "font_size_pt", "f64"),
)
_PLOT_PATCH_SPEC = (
    ("gap_width_pct", "gap_width_pct", "u32"),
    ("overlap", "overlap", "i32"),
    ("first_slice_angle_deg", "first_slice_angle_deg", "u32"),
    ("vary_colors", "vary_colors", "bool"),
    ("hole_size", "hole_size", "u32"),
)
_AXIS_PATCH_SPEC = (
    ("visible", "visible", "bool"),
    ("title", "title", "title"),
    ("min", "min", "f64"),
    ("max", "max", "f64"),
    ("major_unit", "major_unit", "f64"),
    ("number_format", "number_format", "str"),
    ("major_gridlines", "major_gridlines", "bool"),
    ("reversed", "reversed", "bool"),
    ("log_base", "log_base", "f64"),
    ("crosses", "crosses", "i32"),
    ("tick_label_position", "tick_label_position", "i32"),
)
# Patch fields whose `_op` companion drops the unit suffix of the value field.
_PATCH_OP_PREFIX = {"font_size_pt": "font_size"}
_PREVIEW_OVERRIDE_SPEC = (
    ("direction", "direction", "i32"),
    ("header_row", "header_row", "bool"),
    ("header_column", "header_column", "bool"),
    ("blanks_as", "blanks_as", "i32"),
    ("plot_visible_only", "plot_visible_only", "bool"),
    ("na_as_blank", "na_as_blank", "bool"),
)
_DISPLAY_PATCH_SPEC = (
    ("blanks_as", "blanks_as", "i32"),
    ("plot_visible_only", "plot_visible_only", "bool"),
    ("na_as_blank", "na_as_blank", "bool"),
)
_SERIES_REF_KEYS = ("name", "categories", "values", "x", "y", "size")


def _enum(cls, value: int):
    """Wrap ``value`` in ``cls``, keeping the raw integer for a member the binding does not know."""
    try:
        return cls(value)
    except ValueError:
        return value


def _enum_or_none(cls, value: int):
    """``None`` for the C ``-1`` "absent" member, else :func:`_enum`."""
    return None if value == -1 else _enum(cls, value)


def _nan_none(value: float) -> Optional[float]:
    return None if value != value else value


def _tri(value: int) -> Optional[bool]:
    """Decode a chart tri-state (absent / false / true)."""
    return None if value == 0 else value == 2


def _opt_cstr(ptr: int) -> Optional[str]:
    return LIB.read_cstr(ptr) if ptr else None


def _read_struct(fn, layout: S.Struct, *args) -> Dict[str, object]:
    """Call a ``(..., out_struct*)`` ABI function and decode the struct it fills."""
    with _ArgScope() as scope:
        ptr = scope.struct(layout)
        _check(fn(*args, ptr))
        return layout.unpack(LIB, ptr)


def _read_chart_title(d: Dict[str, object]) -> Optional[ChartTitle]:
    if d["kind"] == _CHART_TITLE_NONE:
        return None
    return ChartTitle(
        kind=_enum(ChartTitleKind, d["kind"]), text=LIB.read_cstr(d["text"]), formula=LIB.read_cstr(d["formula"])
    )


def _read_chart_fill(d: Dict[str, object]) -> ChartFill:
    return ChartFill(color=_optional_color_from(d["color"]), transparency_pct=_nan_none(d["transparency_pct"]))


def _read_chart_line(d: Dict[str, object]) -> ChartLine:
    return ChartLine(
        color=_optional_color_from(d["color"]),
        transparency_pct=_nan_none(d["transparency_pct"]),
        width_pt=_nan_none(d["width_pt"]),
        dash=_enum(ChartLineDash, d["dash"]) if d["has_dash"] else None,
        none=bool(d["none"]),
    )


def _read_chart_marker(d: Dict[str, object]) -> ChartMarker:
    return ChartMarker(
        symbol=_enum(ChartMarkerSymbol, d["symbol"]),
        size_pt=d["size_pt"] if d["has_size_pt"] else None,
        fill=_read_chart_fill(d["fill"]),
        line=_read_chart_line(d["line"]),
    )


def _read_chart_labels(d: Dict[str, object]) -> ChartLabels:
    return ChartLabels(
        show_value=bool(d["show_value"]),
        show_category=bool(d["show_category"]),
        show_series=bool(d["show_series"]),
        show_percent=bool(d["show_percent"]),
        show_legend_key=bool(d["show_legend_key"]),
        separator=_opt_cstr(d["separator"]),
        position=_enum(ChartLabelPosition, d["position"]) if d["has_position"] else None,
        number_format=_opt_cstr(d["number_format"]),
        source_linked=bool(d["source_linked"]),
        font_color=_optional_color_from(d["font_color"]),
        font_size_pt=_nan_none(d["font_size_pt"]),
        deleted=bool(d["deleted"]),
    )


def _read_chart_data_source(chart: int, plot: int, series: int, role: int, d: Dict[str, object]) -> ChartDataSource:
    kind = _enum(ChartSourceKind, d["kind"])
    literal: List[Union[float, str]] = []
    if kind == ChartSourceKind.LITERAL:
        for i in range(d["literal_count"]):
            point = _read_struct(LIB.fm_chart_get_series_literal_at, S.CHART_LITERAL, chart, plot, series, role, i)
            literal.append(LIB.read_cstr(point["text"]) if point["is_text"] else point["number"])
    return ChartDataSource(kind=kind, formula=LIB.read_cstr(d["formula"]), literal=literal)


def _read_chart_series(chart: int, plot: int, series: int) -> ChartSeries:
    d = _read_struct(LIB.fm_chart_get_series_at, S.CHART_SERIES, chart, plot, series)
    sources = {key: _read_chart_data_source(chart, plot, series, role, d[key]) for role, key in enumerate(_CHART_ROLES)}
    points: List[ChartPoint] = []
    for i in range(d["point_count"]):
        p = _read_struct(LIB.fm_chart_get_series_point_at, S.CHART_POINT, chart, plot, series, i)
        points.append(
            ChartPoint(
                index=p["idx"],
                fill=_read_chart_fill(p["fill"]),
                line=_read_chart_line(p["line"]),
                marker=_read_chart_marker(p["marker"]) if p["has_marker"] else None,
            )
        )
    return ChartSeries(
        index=d["idx"],
        order=d["order"],
        filtered=bool(d["filtered"]),
        fill=_read_chart_fill(d["fill"]),
        line=_read_chart_line(d["line"]),
        marker=_read_chart_marker(d["marker"]) if d["has_marker"] else None,
        smooth=_tri(d["smooth"]),
        points=points,
        labels=_read_chart_labels(d["labels"]) if d["has_labels"] else None,
        **sources,
    )


def _read_chart_plot(chart: int, plot: int) -> ChartPlot:
    d = _read_struct(LIB.fm_chart_get_plot_at, S.CHART_PLOT, chart, plot)
    return ChartPlot(
        type=_enum(ChartPlotType, d["type"]),
        extended_layout=_enum_or_none(ChartExtendedLayout, d["extended_layout"]),
        bar_direction=_enum_or_none(ChartBarDirection, d["bar_direction"]),
        grouping=_enum(ChartGrouping, d["grouping"]) if d["has_grouping"] else None,
        vary_colors=_tri(d["vary_colors"]) is True,
        gap_width_pct=d["gap_width_pct"] if d["has_gap_width_pct"] else None,
        overlap=d["overlap"] if d["has_overlap"] else None,
        hole_size=d["hole_size"] if d["has_hole_size"] else None,
        first_slice_angle_deg=d["first_slice_angle_deg"] if d["has_first_slice_angle_deg"] else None,
        axis_ids=[
            _read_count(LIB.fm_chart_get_plot_axis_id_at, chart, _uint(plot, "plot"), _uint(i, "idx"))
            for i in range(d["axis_id_count"])
        ],
        series=[_read_chart_series(chart, plot, i) for i in range(d["series_count"])],
    )


def _read_chart_axis(chart: int, idx: int) -> ChartAxis:
    d = _read_struct(LIB.fm_chart_get_axis_at, S.CHART_AXIS, chart, idx)
    return ChartAxis(
        id=d["id"],
        kind=_enum(ChartAxisKind, d["kind"]),
        position=_enum(ChartAxisPosition, d["position"]),
        visible=bool(d["visible"]),
        cross_axis=d["cross_axis"],
        crosses=_enum_or_none(ChartAxisCrosses, d["crosses"]),
        crosses_at=_nan_none(d["crosses_at"]),
        min=_nan_none(d["min"]),
        max=_nan_none(d["max"]),
        major_unit=_nan_none(d["major_unit"]),
        minor_unit=_nan_none(d["minor_unit"]),
        log_base=_nan_none(d["log_base"]),
        reversed=bool(d["reversed"]),
        number_format=_opt_cstr(d["number_format"]),
        source_linked=bool(d["source_linked"]),
        title=_read_chart_title(d["title"]),
        major_gridlines=bool(d["major_gridlines"]),
        minor_gridlines=bool(d["minor_gridlines"]),
        tick_label_position=_enum_or_none(ChartTickLabelPosition, d["tick_label_position"]),
    )


def _read_chart(chart: int) -> ChartProperties:
    """Decode an open chart handle."""
    d = _read_struct(LIB.fm_chart_get_info, S.CHART_INFO, chart)
    source = None
    if d["source_range"]:
        source = ChartSource(
            range=LIB.read_cstr(d["source_range"]),
            direction=_enum(ChartSeriesDirection, d["source_direction"]),
            header_row=bool(d["source_header_row"]),
            header_column=bool(d["source_header_column"]),
        )
    return ChartProperties(
        object_id=d["object_id"],
        name=LIB.read_cstr(d["name"]),
        description=LIB.read_cstr(d["description"]),
        is_extended=bool(d["is_extended"]),
        is_pivot_chart=bool(d["is_pivot_chart"]),
        on_chartsheet=bool(d["on_chartsheet"]),
        title=_read_chart_title(d["title"]),
        auto_title_deleted=_tri(d["auto_title_deleted"]),
        legend=None
        if d["legend_position"] == _CHART_LEGEND_NONE
        else ChartLegend(
            position=_enum(ChartLegendPosition, d["legend_position"]),
            overlay=bool(d["legend_overlay"]),
            hidden_series=[
                _read_count(LIB.fm_chart_get_legend_hidden_at, chart, _uint(i, "idx"))
                for i in range(d["legend_hidden_count"])
            ],
        ),
        blanks_as=_enum_or_none(BlanksAs, d["blanks_as"]),
        plot_visible_only=_tri(d["plot_visible_only"]),
        na_as_blank=_tri(d["na_as_blank"]),
        date1904=bool(d["date1904"]),
        style=d["style"] if d["has_style"] else None,
        plots=[_read_chart_plot(chart, i) for i in range(d["plot_count"])],
        axes=[_read_chart_axis(chart, i) for i in range(d["axis_count"])],
        hidden_categories=[
            _read_count(LIB.fm_chart_get_hidden_category_at, chart, _uint(i, "idx"))
            for i in range(d["hidden_category_count"])
        ],
        source=source,
    )


def _read_chart_data_array(ptr: int, count: int, fmt: str, width: int) -> Optional[List]:
    """Copy ``count`` values of struct format ``fmt`` out of a handle-owned array; ``None`` when there is no array."""
    if not ptr:
        return None
    if count == 0:
        return []
    return list(struct.unpack(f"<{count}{fmt}", LIB.read_bytes(ptr, count * width)))


def _read_chart_data(data: int) -> ChartData:
    """Decode an open chart-data handle."""
    info = _read_struct(LIB.fm_chart_data_get_info, S.CHART_DATA_INFO, data)
    labels: List[str] = []
    values: List[float] = []
    with _ArgScope() as scope:
        label_out = scope.out_ptr()
        value_out = scope.alloc(8)
        for i in range(info["category_label_count"]):
            _check(LIB.fm_chart_data_get_category_at(data, _uint(i, "idx"), label_out, value_out))
            labels.append(LIB.read_cstr(LIB.read_u32(label_out)))
            values.append(LIB.read_f64(value_out))
        levels: List[List[str]] = []
        for level in range(info["category_level_count"]):
            size = _read_count(LIB.fm_chart_data_get_category_level_count, data, _uint(level, "level"))
            names: List[str] = []
            for i in range(size):
                _check(LIB.fm_chart_data_get_category_level_at(data, _uint(level, "level"), _uint(i, "idx"), label_out))
                names.append(LIB.read_cstr(LIB.read_u32(label_out)))
            levels.append(names)

    series: List[ChartDataSeries] = []
    for i in range(info["series_count"]):
        s = _read_struct(LIB.fm_chart_data_get_series_at, S.CHART_DATA_SERIES, data, i)
        points = _read_struct(LIB.fm_chart_data_get_series_points, S.CHART_DATA_POINTS, data, i)
        n = points["count"]
        ranges = ChartSeriesRanges()
        for role, key in enumerate(_CHART_ROLES):
            for j in range(s["range_count"][role]):
                r = _read_struct(LIB.fm_chart_data_get_range_at, S.SHEET_RECT, data, i, role, j)
                getattr(ranges, key).append(SheetRect(sheet=r["sheet_index"], rect=_cell_rect_from(r["rect"])))
        series.append(
            ChartDataSeries(
                plot_index=s["plot_index"],
                index=s["idx"],
                name=_opt_cstr(s["name"]),
                format_code=LIB.read_cstr(s["format_code"]),
                x_format_code=LIB.read_cstr(s["x_format_code"]),
                filtered=bool(s["filtered"]),
                values=_read_chart_data_array(points["values"], n, "d", 8) or [],
                states=[_enum(ChartPointState, st) for st in _read_chart_data_array(points["states"], n, "B", 1) or []],
                x=_read_chart_data_array(points["x"], n, "d", 8),
                sizes=_read_chart_data_array(points["sizes"], n, "d", 8) if s["has_sizes"] else None,
                ranges=ranges,
            )
        )
    return ChartData(
        domain=ChartDomain.XY if info["is_xy"] else ChartDomain.CATEGORY,
        point_count=info["point_count"],
        categories=ChartCategories(
            labels=labels,
            levels=levels,
            values=values if info["category_value_count"] else None,
            format_code=LIB.read_cstr(info["category_format_code"]),
            generated=bool(info["categories_generated"]),
        ),
        series=series,
        title_text=_opt_cstr(info["title_text"]),
        axis_title_texts=[
            ChartAxisTitleText(axis_id=t["axis_id"], text=LIB.read_cstr(t["text"]))
            for t in (
                _read_struct(LIB.fm_chart_data_get_axis_title_at, S.CHART_AXIS_TITLE_TEXT, data, i)
                for i in range(info["axis_title_count"])
            )
        ],
        unresolved=[
            ChartUnresolved(series_index=u["series_index"], source=_enum(ChartRole, u["source"]), axis_id=u["axis_id"])
            for u in (
                _read_struct(LIB.fm_chart_data_get_unresolved_at, S.CHART_UNRESOLVED, data, i)
                for i in range(info["unresolved_count"])
            )
        ],
    )


def _title_fields(title: Union[str, ChartTitle], scope: _ArgScope) -> Dict[str, object]:
    """Encode a title given as its text or as a :class:`ChartTitle`."""
    if isinstance(title, ChartTitle):
        return {
            "kind": _sint(title.kind, "kind"),
            "text": scope.cstr(title.text),
            "formula": scope.cstr(title.formula),
        }
    return {"kind": 1, "text": scope.cstr(title)}


def _patch_scalar(kind: str, value, key: str, scope: _ArgScope):
    if kind == "u32":
        return _uint(value, key)
    if kind == "color":
        return _color_fields(color_input(value, key))
    if kind == "i32":
        return _sint(value, key)
    if kind == "bool":
        return 1 if value else 0
    if kind == "f64":
        return float(value)
    if kind == "title":
        return _title_fields(value, scope)
    return scope.cstr(value)


def _check_patch_keys(patch: Mapping, spec: Sequence, extra: Sequence[str], what: str) -> None:
    unknown = set(patch) - {key for key, _, _ in spec} - set(extra)
    if unknown:
        raise TypeError(f"formulon: unknown {what} key(s): {sorted(unknown)}")


def _encode_patch(
    spec: Sequence, patch: Mapping, extra: Sequence[str], what: str, scope: _ArgScope
) -> Dict[str, object]:
    """Encode the ``(op, value)`` pair of every ``spec`` key present in ``patch``.

    An absent key stays zero (keep), ``None`` writes the clear op and any other
    value the set op beside it.
    """
    _check_patch_keys(patch, spec, extra, what)
    out: Dict[str, object] = {}
    for key, field_name, kind in spec:
        if key not in patch:
            continue
        value = patch[key]
        op_field = _PATCH_OP_PREFIX.get(field_name, field_name) + "_op"
        if value is None:
            out[op_field] = _PATCH_CLEAR
        else:
            out[op_field] = _PATCH_SET
            out[field_name] = _patch_scalar(kind, value, key, scope)
    return out


def _encode_fill_patch(patch: Optional[Mapping], scope: _ArgScope) -> Dict[str, object]:
    return _encode_patch(_FILL_PATCH_SPEC, patch or {}, (), "fill patch", scope)


def _encode_line_patch(patch: Optional[Mapping], scope: _ArgScope) -> Dict[str, object]:
    return _encode_patch(_LINE_PATCH_SPEC, patch or {}, (), "line patch", scope)


def _encode_marker_patch(patch: Mapping, scope: _ArgScope) -> Dict[str, object]:
    out = _encode_patch(_MARKER_PATCH_SPEC, patch, ("fill", "line"), "marker patch", scope)
    out["fill"] = _encode_fill_patch(patch.get("fill"), scope)
    out["line"] = _encode_line_patch(patch.get("line"), scope)
    return out


def _encode_labels_patch(patch: Mapping, scope: _ArgScope) -> Dict[str, object]:
    return _encode_patch(_LABELS_PATCH_SPEC, patch, (), "labels patch", scope)


def _encode_optional_patch(out: Dict[str, object], key: str, patch: Mapping, encode, scope: _ArgScope) -> None:
    """Encode a nested patch that ``None`` removes (``<key>_op`` beside the nested struct)."""
    if key not in patch:
        return
    if patch[key] is None:
        out[key + "_op"] = _PATCH_CLEAR
    else:
        out[key + "_op"] = _PATCH_SET
        out[key] = encode(patch[key], scope)


def _source_text(value: object, key: str) -> object:
    """The reference text of a series input; a :class:`ChartDataSource` reads as its formula."""
    if not isinstance(value, ChartDataSource):
        return value
    if value.kind == ChartSourceKind.LITERAL:
        raise TypeError(f"formulon: {key} cannot be set from a literal ChartDataSource")
    return value.formula


def _encode_series_patch(patch: Mapping, scope: _ArgScope) -> Dict[str, object]:
    unknown = set(patch) - set(_SERIES_REF_KEYS) - {"fill", "line", "marker", "smooth", "labels"}
    if unknown:
        raise TypeError(f"formulon: unknown series patch key(s): {sorted(unknown)}")
    out: Dict[str, object] = {
        key: scope.cstr(_source_text(patch[key], key)) for key in _SERIES_REF_KEYS if key in patch
    }
    out["fill"] = _encode_fill_patch(patch.get("fill"), scope)
    out["line"] = _encode_line_patch(patch.get("line"), scope)
    _encode_optional_patch(out, "marker", patch, _encode_marker_patch, scope)
    _encode_optional_patch(out, "labels", patch, _encode_labels_patch, scope)
    if "smooth" in patch:
        out["smooth_op"] = _PATCH_CLEAR if patch["smooth"] is None else _PATCH_SET
        out["smooth"] = 1 if patch["smooth"] else 0
    return out


def _encode_point_patch(patch: Optional[Mapping], scope: _ArgScope) -> Dict[str, object]:
    patch = patch or {}
    unknown = set(patch) - {"fill", "line", "marker"}
    if unknown:
        raise TypeError(f"formulon: unknown point patch key(s): {sorted(unknown)}")
    out: Dict[str, object] = {
        "fill": _encode_fill_patch(patch.get("fill"), scope),
        "line": _encode_line_patch(patch.get("line"), scope),
    }
    _encode_optional_patch(out, "marker", patch, _encode_marker_patch, scope)
    return out


def _encode_axis_patch(patch: Mapping, scope: _ArgScope) -> Dict[str, object]:
    out = _encode_patch(_AXIS_PATCH_SPEC, patch, ("crosses_at",), "axis patch", scope)
    if "crosses_at" in patch:
        out["crosses_at"] = float(patch["crosses_at"])
    return out


def _header_mode(flag: Optional[bool]) -> int:  # noqa: FBT001 -- private tri-state encoder
    """Map an optional header flag to ``fm_tristate_t``: ``None`` is Excel's detection."""
    return _tristate_default(flag)


def _id_array(values: Sequence[int], name: str, scope: _ArgScope) -> int:
    """Allocate ``values`` as a ``uint32_t`` array (0 when empty)."""
    if not values:
        return 0
    ptr = scope.alloc(4 * len(values))
    LIB.write_bytes(ptr, struct.pack(f"<{len(values)}I", *(_uint(v, name) for v in values)))
    return ptr


def _read_owned_buffer(fn, *args) -> bytes:
    """Call a save-style ABI function and copy its owned byte buffer."""
    with _ArgScope() as scope:
        data_ptr = 0
        try:
            out_ptr_ptr = scope.out_ptr()
            out_len_ptr = scope.out_ptr()
            _check(fn(*args, out_ptr_ptr, out_len_ptr))
            data_ptr = LIB.read_u32(out_ptr_ptr)
            data_len = LIB.read_u32(out_len_ptr)
            if data_len == 0 or data_ptr == 0:
                return b""
            return LIB.read_bytes(data_ptr, data_len)
        finally:
            if data_ptr:
                LIB.fm_buffer_destroy(data_ptr)


def _pack_cell_rect_array(ranges: Sequence[CellRect], scope: _ArgScope) -> int:
    """Pack a list of ``CellRect`` into a contiguous engine-memory array.

    Returns the array pointer (0 when empty), allocated through ``scope``.
    """
    if not ranges:
        return 0
    size = S.CELL_RECT.size
    ptr = scope.alloc(size * len(ranges))
    for i, r in enumerate(ranges):
        _pack_cell_rect_at(ptr + i * size, r)
    return ptr


def _decode_cell_rect_array(ptr: int, count: int) -> List[CellRect]:
    """Decode the ``count`` contiguous cell rect structs at ``ptr``."""
    size = S.CELL_RECT.size
    return [_decode_cell_rect_at(ptr + i * size) for i in range(count)]


def _check_value(value: Value) -> Value:
    """Reject anything but a :class:`Value`, whose constructor validated its payload."""
    if not isinstance(value, Value):
        raise TypeError(f"formulon: expected a Value, got {type(value).__name__}")
    return value


def _write_value_input(value: Value, value_ptr: int, scope: _ArgScope) -> None:
    """Write ``value`` as a value struct at ``value_ptr``.

    ``value`` has passed :func:`_check_value`, so its payload matches its
    kind. A text payload's buffer is allocated through ``scope``.
    """
    payload = b"\x00" * 8
    if value.kind == ValueKind.NUMBER:
        payload = struct.pack("<d", value.number)
    elif value.kind == ValueKind.BOOL:
        payload = struct.pack("<i", 1 if value.boolean else 0) + b"\x00" * 4
    elif value.kind == ValueKind.ERROR:
        payload = struct.pack("<i", value.error_code) + b"\x00" * 4
    elif value.kind == ValueKind.TEXT:
        payload = struct.pack("<I", scope.cstr(value.text)) + b"\x00" * 4
    LIB.write_bytes(value_ptr, struct.pack("<i", value.kind) + b"\x00" * 4 + payload)


def _pack_mentions(mentions: Sequence[Mention], scope: _ArgScope) -> int:
    """Pack ``mentions`` into a contiguous mention struct array (0 when empty)."""
    if not mentions:
        return 0
    base = scope.struct_array(S.MENTION, len(mentions))
    for i, m in enumerate(mentions):
        ptr = base + i * S.MENTION.size
        S.MENTION.pack(LIB, ptr, {"start": _uint(m.start, "start"), "length": _uint(m.length, "length")})
        S.write_str_field(LIB, ptr, S.MENTION, "person_id", m.person_id, scope)
        S.write_str_field(LIB, ptr, S.MENTION, "mention_id", m.mention_id, scope)
    return base


def _write_mentions(struct_ptr: int, layout: S.Struct, mentions: Sequence[Mention], scope: _ArgScope) -> None:
    """Store a mention array and its count into the ``mentions`` / ``mention_count`` fields."""
    LIB.write_bytes(struct_ptr + layout.offsets["mentions"][1], struct.pack("<I", _pack_mentions(mentions, scope)))
    LIB.write_bytes(struct_ptr + layout.offsets["mention_count"][1], struct.pack("<I", len(mentions)))


def _decode_mentions(base: int, count: int) -> List[Mention]:
    out: List[Mention] = []
    for i in range(count):
        d = S.MENTION.unpack(LIB, base + i * S.MENTION.size)
        out.append(
            Mention(
                person_id=LIB.read_cstr(d["person_id"]),
                mention_id=LIB.read_cstr(d["mention_id"]),
                start=d["start"],
                length=d["length"],
            )
        )
    return out


def _require_cell_rect(r: object) -> CellRect:
    if not isinstance(r, CellRect):
        raise TypeError(f"formulon: expected a CellRect, got {type(r).__name__}")
    return r


def _cell_rect_fields(r: CellRect) -> Dict[str, int]:
    """The cell rect struct field values of ``r``; anything but a :class:`CellRect` is a ``TypeError``."""
    r = _require_cell_rect(r)
    return {
        "first_row": _uint(r.first_row, "first_row"),
        "first_col": _uint(r.first_col, "first_col"),
        "last_row": _uint(r.last_row, "last_row"),
        "last_col": _uint(r.last_col, "last_col"),
    }


def _cell_rect_from(d: Dict[str, int]) -> CellRect:
    return CellRect(d["first_row"], d["first_col"], d["last_row"], d["last_col"])


def _hyperlink_from(d: Dict[str, Any]) -> Hyperlink:
    return Hyperlink(
        rect=_cell_rect_from(d["rect"]),
        target=LIB.read_cstr(d["target"]) or None,
        location=LIB.read_cstr(d["location"]) or None,
        display=LIB.read_cstr(d["display"]) or None,
        tooltip=LIB.read_cstr(d["tooltip"]) or None,
    )


def _parse_created(text: str) -> Optional[datetime.datetime]:
    """The ``datetime`` a threaded comment's ``YYYY-MM-DDTHH:MM:SS[.ss]`` stamp names, hundredths dropped."""
    if not text:
        return None
    try:
        return datetime.datetime.strptime(text[:19], "%Y-%m-%dT%H:%M:%S")
    except ValueError:
        return None


def _format_created(value: Optional[Union[datetime.datetime, datetime.date]]) -> str:
    """The C stamp of a threaded comment's ``created`` input; ``""`` for ``None``."""
    if value is None:
        return ""
    return civil_datetime(value, "created").strftime("%Y-%m-%dT%H:%M:%S") + ".00"


def _threaded_comment_from(d: Dict[str, Any]) -> ThreadedComment:
    return ThreadedComment(
        id=LIB.read_cstr(d["id"]),
        row=d["row"],
        col=d["col"],
        person_id=LIB.read_cstr(d["person_id"]),
        created=_parse_created(LIB.read_cstr(d["created"])),
        text=LIB.read_cstr(d["text"]),
        parent_id=LIB.read_cstr(d["parent_id"]) or None,
        done=bool(d["done"]),
        mentions=_decode_mentions(d["mentions"], d["mention_count"]),
    )


def _drawing_object_from(d: Dict[str, Any]) -> DrawingProperties:
    return DrawingProperties(
        object_id=d["object_id"],
        kind=_enum(DrawingObjectKind, d["kind"]),
        anchor_kind=_enum(AnchorKind, d["anchor_kind"]),
        edit_as=_enum(AnchorEditAs, d["edit_as"]),
        from_row=d["from_row"],
        from_col=d["from_col"],
        from_row_offset_emu=d["from_row_offset_emu"],
        from_col_offset_emu=d["from_col_offset_emu"],
        to_row=d["to_row"],
        to_col=d["to_col"],
        to_row_offset_emu=d["to_row_offset_emu"],
        to_col_offset_emu=d["to_col_offset_emu"],
        width_emu=d["width_emu"],
        height_emu=d["height_emu"],
        format=_enum(ImageFormat, d["image_format"]),
        name=LIB.read_cstr(d["name"]),
        description=LIB.read_cstr(d["description"]),
    )


def _shape_from(d: Dict[str, Any]) -> ShapeProperties:
    return ShapeProperties(
        object_id=d["object_id"],
        kind=_enum(DrawingObjectKind, d["kind"]),
        anchor_kind=_enum(AnchorKind, d["anchor_kind"]),
        edit_as=_enum(AnchorEditAs, d["edit_as"]),
        from_row=d["from_row"],
        from_col=d["from_col"],
        from_row_offset_emu=d["from_row_offset_emu"],
        from_col_offset_emu=d["from_col_offset_emu"],
        to_row=d["to_row"],
        to_col=d["to_col"],
        to_row_offset_emu=d["to_row_offset_emu"],
        to_col_offset_emu=d["to_col_offset_emu"],
        width_emu=d["width_emu"],
        height_emu=d["height_emu"],
        name=LIB.read_cstr(d["name"]),
        preset_geometry=LIB.read_cstr(d["preset_geometry"]),
        text=LIB.read_cstr(d["text"]),
        text_box=bool(d["text_box"]),
        fill_kind=_enum_or_none(ShapePaint, d["fill_kind"]),
        fill_color=_optional_color_from(d["fill_color"]),
        line_kind=_enum_or_none(ShapePaint, d["line_kind"]),
        line_color=_optional_color_from(d["line_color"]),
        line_width_pt=d["line_width_pt"],
        rotation_deg=d["rotation_deg"],
        hidden=bool(d["hidden"]),
    )


def _slicer_from(d: Dict[str, Any]) -> SlicerProperties:
    return SlicerProperties(
        object_id=d["object_id"],
        name=LIB.read_cstr(d["name"]),
        caption=LIB.read_cstr(d["caption"]),
        cache_name=LIB.read_cstr(d["cache_name"]),
        table_name=LIB.read_cstr(d["table_name"]),
        column_name=LIB.read_cstr(d["column_name"]),
        style=LIB.read_cstr(d["style"]) or None,
        column_count=d["column_count"],
        row_height_emu=d["row_height_emu"],
        sort_order=_enum(SlicerSortOrder, d["sort_order"]),
        cross_filter=_enum(SlicerCrossFilter, d["cross_filter"]),
        anchor_kind=_enum(AnchorKind, d["anchor_kind"]),
        from_row=d["from_row"],
        from_col=d["from_col"],
        from_row_offset_emu=d["from_row_offset_emu"],
        from_col_offset_emu=d["from_col_offset_emu"],
        to_row=d["to_row"],
        to_col=d["to_col"],
        to_row_offset_emu=d["to_row_offset_emu"],
        to_col_offset_emu=d["to_col_offset_emu"],
    )


def _person_from(d: Dict[str, Any]) -> Person:
    return Person(
        id=LIB.read_cstr(d["id"]),
        display_name=LIB.read_cstr(d["display_name"]),
        user_id=LIB.read_cstr(d["user_id"]),
        provider_id=LIB.read_cstr(d["provider_id"]),
    )


def _cell_style_from(d: Dict[str, Any]) -> CellStyle:
    return CellStyle(
        name=LIB.read_cstr(d["name"]),
        style_xf_index=d["style_xf_index"],
        builtin_id=d["builtin_id"],
        i_level=d["i_level"],
        hidden=bool(d["hidden"]),
        custom_builtin=bool(d["custom_builtin"]),
    )


def _table_style_from(d: Dict[str, Any]) -> TableStyle:
    elements = []
    for i in range(d["element_count"]):
        e = S.TABLE_STYLE_ELEMENT.unpack(LIB, d["elements"] + i * S.TABLE_STYLE_ELEMENT.size)
        elements.append(
            TableStyleElement(
                type=e["type"],
                size=e["size"],
                dxf_index=e["dxf_index"] if e["has_dxf_index"] else None,
            )
        )
    return TableStyle(
        name=LIB.read_cstr(d["name"]),
        pivot=d["pivot"] != 0,
        table=d["table"] != 0,
        elements=list(elements),
    )


def _pack_cell_rect_at(ptr: int, r: CellRect) -> None:
    r = _require_cell_rect(r)
    S.CELL_RECT.pack(
        LIB,
        ptr,
        {
            "first_row": _uint(r.first_row, "first_row"),
            "first_col": _uint(r.first_col, "first_col"),
            "last_row": _uint(r.last_row, "last_row"),
            "last_col": _uint(r.last_col, "last_col"),
        },
    )


def _decode_cell_rect_at(ptr: int) -> CellRect:
    return _cell_rect_from(S.CELL_RECT.unpack(LIB, ptr))


_FILTER_COLUMN_BOOLS = (
    "hidden_button",
    "show_button",
    "filter_blank",
    "custom_and",
    "top",
    "percent",
    "cell_color",
)
# (record field, C field) pairs; the C struct keeps the OOXML attribute spellings.
_FILTER_COLUMN_INTS = (
    ("kind", "kind"),
    ("custom_count", "custom_count"),
    ("operator1", "op1"),
    ("operator2", "op2"),
    ("dynamic_type", "dynamic_type"),
    ("icon_set", "icon_set"),
)
_FILTER_COLUMN_DOUBLES = (("top_value", "top_val"),)
# Optional members: (record field, C field, the ``has_*`` flag that states it).
_FILTER_COLUMN_OPTIONAL = (
    ("filter_value", "filter_val", "has_filter_val"),
    ("dynamic_value", "dyn_val", "has_dyn_val"),
    ("dynamic_max_value", "dyn_max_val", "has_dyn_max_val"),
    ("icon_index", "icon_index", "has_icon_index"),
)
_FILTER_COLUMN_STRINGS = (
    ("value1", "val1"),
    ("value2", "val2"),
    ("value_iso", "val_iso"),
    ("max_value_iso", "max_val_iso"),
)


def _pack_filter_column(ptr: int, col: AutoFilterColumn, scope: _ArgScope) -> None:
    fields: Dict[str, object] = {
        "column_offset": _uint(col.column_offset, "column_offset"),
        "dxf_index": _uint(col.dxf_index, "dxf_index"),
    }
    for name in _FILTER_COLUMN_BOOLS:
        fields[name] = 1 if getattr(col, name) else 0
    for name, c_name in _FILTER_COLUMN_INTS:
        value = getattr(col, name)
        # An omitted operator packs as BETWEEN, which C rejects in a used condition slot.
        fields[c_name] = _sint(0 if value is None else value, name)
    for name, c_name in _FILTER_COLUMN_DOUBLES:
        fields[c_name] = float(getattr(col, name))
    for name, c_name, flag in _FILTER_COLUMN_OPTIONAL:
        value = getattr(col, name)
        fields[flag] = 0 if value is None else 1
        if value is None:
            continue
        if name == "icon_index":
            fields[c_name] = _sint(value, name)
        else:
            fields[c_name] = float(value)
    fields["value_count"] = _uint(len(col.values), "value_count")
    fields["date_group_count"] = _uint(len(col.date_groups), "date_group_count")
    S.FILTER_COLUMN.pack(LIB, ptr, fields)
    for name, c_name in _FILTER_COLUMN_STRINGS:
        S.write_str_field(LIB, ptr, S.FILTER_COLUMN, c_name, getattr(col, name), scope)
    if col.values:
        arr = scope.alloc(4 * len(col.values))
        for i, text in enumerate(col.values):
            text_ptr = scope.cstr(text)
            LIB.write_bytes(arr + 4 * i, struct.pack("<I", text_ptr))
        LIB.write_bytes(ptr + S.FILTER_COLUMN.offsets["values"][1], struct.pack("<I", arr))
    if col.date_groups:
        base = scope.struct_array(S.DATE_GROUP_ITEM, len(col.date_groups))
        for i, g in enumerate(col.date_groups):
            S.DATE_GROUP_ITEM.pack(
                LIB,
                base + i * S.DATE_GROUP_ITEM.size,
                {
                    "year": _uint(g.year, "year", 16),
                    "month": _uint(g.month, "month", 8),
                    "day": _uint(g.day, "day", 8),
                    "hour": _uint(g.hour, "hour", 8),
                    "minute": _uint(g.minute, "minute", 8),
                    "second": _uint(g.second, "second", 8),
                    "grouping": _uint(g.grouping, "grouping", 8),
                },
            )
        LIB.write_bytes(ptr + S.FILTER_COLUMN.offsets["date_groups"][1], struct.pack("<I", base))


def _pack_sort_condition(ptr: int, cond: AutoFilterSortCondition, scope: _ArgScope) -> None:
    S.SORT_CONDITION.pack(
        LIB,
        ptr,
        {
            "descending": 1 if cond.descending else 0,
            "sort_by": _sint(cond.sort_by, "sort_by"),
            "dxf_index": _uint(cond.dxf_index or 0, "dxf_index"),
            "has_dxf_index": 0 if cond.dxf_index is None else 1,
            "icon_set": _sint(cond.icon_set, "icon_set"),
            "icon_index": _sint(cond.icon_index or 0, "icon_index"),
            "has_icon_index": 0 if cond.icon_index is None else 1,
        },
    )
    _pack_cell_rect_at(ptr + S.SORT_CONDITION.offsets["ref"][1], cond.ref)
    S.write_str_field(LIB, ptr, S.SORT_CONDITION, "custom_list", cond.custom_list, scope)


def _pack_auto_filter(ptr: int, af: AutoFilter, scope: _ArgScope) -> None:
    """Fill the AutoFilter struct at ``ptr``; every nested buffer is allocated through ``scope``."""
    S.AUTO_FILTER.pack(
        LIB,
        ptr,
        {
            "column_count": _uint(len(af.columns), "column_count"),
        },
    )
    off = S.AUTO_FILTER.offsets
    _pack_cell_rect_at(ptr + off["range"][1], af.range)
    if af.columns:
        base = scope.struct_array(S.FILTER_COLUMN, len(af.columns))
        for i, col in enumerate(af.columns):
            _pack_filter_column(base + i * S.FILTER_COLUMN.size, col, scope)
        LIB.write_bytes(ptr + off["columns"][1], struct.pack("<I", base))
    sort = af.sort
    if sort is None:
        return
    fields = {
        "has_sort": 1,
        "column_sort": 1 if sort.column_sort else 0,
        "case_sensitive": 1 if sort.case_sensitive else 0,
        "sort_method": _sint(sort.sort_method, "sort_method"),
        "condition_count": _uint(len(sort.conditions), "condition_count"),
    }
    for name, value in fields.items():
        LIB.write_bytes(ptr + off[name][1], struct.pack("<i" if name != "condition_count" else "<I", value))
    _pack_cell_rect_at(ptr + off["sort_ref"][1], sort.ref)
    if sort.conditions:
        base = scope.struct_array(S.SORT_CONDITION, len(sort.conditions))
        for i, cond in enumerate(sort.conditions):
            _pack_sort_condition(base + i * S.SORT_CONDITION.size, cond, scope)
        LIB.write_bytes(ptr + off["conditions"][1], struct.pack("<I", base))


def _decode_filter_column(ptr: int) -> AutoFilterColumn:
    d = S.FILTER_COLUMN.unpack(LIB, ptr)
    values: List[str] = []
    if d["value_count"]:
        raw = LIB.read_bytes(d["values"], 4 * d["value_count"])
        values = [LIB.read_cstr(p) for p in struct.unpack(f"<{d['value_count']}I", raw)]
    groups: List[AutoFilterDateGroup] = []
    for i in range(d["date_group_count"]):
        g = S.DATE_GROUP_ITEM.unpack(LIB, d["date_groups"] + i * S.DATE_GROUP_ITEM.size)
        groups.append(AutoFilterDateGroup(**g))
    kwargs: Dict[str, object] = {name: bool(d[name]) for name in _FILTER_COLUMN_BOOLS}
    kwargs.update({name: d[c_name] for name, c_name in _FILTER_COLUMN_INTS + _FILTER_COLUMN_DOUBLES})
    used = d["custom_count"] if d["kind"] == FilterKind.CUSTOM else 0
    for slot, name in enumerate(("operator1", "operator2"), start=1):
        if slot > used:
            kwargs[name] = None
    kwargs.update({name: d[c_name] if d[flag] else None for name, c_name, flag in _FILTER_COLUMN_OPTIONAL})
    kwargs.update({name: LIB.read_cstr(d[c_name]) for name, c_name in _FILTER_COLUMN_STRINGS})
    return AutoFilterColumn(
        column_offset=d["column_offset"], dxf_index=d["dxf_index"], values=values, date_groups=groups, **kwargs
    )  # type: ignore[arg-type]


def _decode_sort_condition(ptr: int) -> AutoFilterSortCondition:
    d = S.SORT_CONDITION.unpack(LIB, ptr)
    return AutoFilterSortCondition(
        ref=_decode_cell_rect_at(ptr + S.SORT_CONDITION.offsets["ref"][1]),
        descending=bool(d["descending"]),
        sort_by=d["sort_by"],
        custom_list=LIB.read_cstr(d["custom_list"]),
        dxf_index=d["dxf_index"] if d["has_dxf_index"] else None,
        icon_set=d["icon_set"],
        icon_index=d["icon_index"] if d["has_icon_index"] else None,
    )


def _decode_auto_filter(ptr: int) -> AutoFilter:
    d = S.AUTO_FILTER.unpack(LIB, ptr)
    off = S.AUTO_FILTER.offsets
    sort = None
    if d["has_sort"]:
        sort = AutoFilterSortState(
            ref=_decode_cell_rect_at(ptr + off["sort_ref"][1]),
            column_sort=bool(d["column_sort"]),
            case_sensitive=bool(d["case_sensitive"]),
            sort_method=d["sort_method"],
            conditions=[
                _decode_sort_condition(d["conditions"] + i * S.SORT_CONDITION.size) for i in range(d["condition_count"])
            ],
        )
    return AutoFilter(
        range=_decode_cell_rect_at(ptr + off["range"][1]),
        columns=[_decode_filter_column(d["columns"] + i * S.FILTER_COLUMN.size) for i in range(d["column_count"])],
        sort=sort,
    )


def _pack_phonetic_run_array(runs: Sequence[PhoneticRun], scope: _ArgScope) -> int:
    """Pack a list of ``PhoneticRun`` into a contiguous engine-memory array.

    Returns the array pointer (0 when empty). Every buffer allocated here --
    the array and each run's kana string -- goes through ``scope``.
    """
    if not runs:
        return 0
    size = S.PHONETIC_RUN.size
    ptr = scope.alloc(size * len(runs))
    for i, run in enumerate(runs):
        slot = ptr + i * size
        S.PHONETIC_RUN.pack(LIB, slot, {"sb": _uint(run.start, "start"), "eb": _uint(run.end, "end")})
        S.write_str_field(LIB, slot, S.PHONETIC_RUN, "text", run.text, scope, allow_empty=True)
    return ptr


# The container each recognised save-path extension names.
_SAVE_EXTENSIONS = {".xlsx": WorkbookFormat.XLSX, ".xlsb": WorkbookFormat.XLSB}


def _save_format(dest: object, format: Optional[WorkbookFormat]) -> WorkbookFormat:
    """The container a save writes: ``format``, else the one a path's extension names.

    Bytes and streams default to XLSX; a path with any other extension needs ``format``.
    """
    if not (dest is None or isinstance(dest, (str, os.PathLike)) or callable(getattr(dest, "write", None))):
        raise TypeError(f"Workbook.save: expected a path or a binary stream, got {type(dest).__name__}")
    implied = None
    if isinstance(dest, (str, os.PathLike)):
        implied = _SAVE_EXTENSIONS.get(os.path.splitext(os.fspath(dest))[1].lower())
    if format is None:
        if implied is not None:
            return implied
        if isinstance(dest, (str, os.PathLike)):
            raise ValueError(
                f"formulon: cannot infer the workbook format from {os.fspath(dest)!r}; "
                "use a .xlsx or .xlsb extension or pass format=WorkbookFormat.XLSX or WorkbookFormat.XLSB"
            )
        return WorkbookFormat.XLSX
    fmt = WorkbookFormat(format)
    if implied is not None and fmt != implied:
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT,
            _diagnostic_override=(f"format {fmt.name} contradicts the extension of {os.fspath(dest)!r}", ""),
        )
    return fmt


def _write_dest(dest: object, data: bytes) -> None:
    """Write ``data`` to a path or a binary stream; nothing for ``None``."""
    if dest is None:
        return
    if isinstance(dest, (str, os.PathLike)):
        with open(dest, "wb") as f:
            f.write(data)
    else:
        cast(BinaryIO, dest).write(data)


#: A pivot-cache value: a number, text, a boolean, a blank (``None``) or an Excel error.
_PivotCacheValue = Union[None, float, str, bool, ExcelError, datetime.date]


def _pivot_cache_value_kind(value: object) -> str:
    """Which typed C writer stores ``value``: ``number``, ``date``, ``text``, ``bool``, ``blank`` or ``error``.

    A ``date`` or ``datetime`` is a date; a number is a number.
    """
    if value is None:
        return "blank"
    if isinstance(value, bool):
        return "bool"
    if isinstance(value, datetime.date):
        return "date"
    if isinstance(value, numbers.Real):
        return "number"
    if isinstance(value, str):
        return "text"
    if isinstance(value, ExcelError):
        return "error"
    raise TypeError(
        "formulon: a pivot cache value must be a number, a str, a bool, None, an ExcelError, "
        f"a date or a datetime, got {type(value).__name__}"
    )


class Workbook:
    """An Excel workbook held by the engine.

    Construction:
      Use one of :meth:`create_default`, :meth:`create_for_profile`,
      :meth:`create_empty`, or :meth:`load`. Calling ``Workbook()``
      directly raises ``TypeError``.

    Sheets:
      The workbook holds the workbook-wide operations; everything that
      targets one sheet lives on its :class:`Sheet` handle. :meth:`sheet`
      and ``wb[key]`` take a sheet name, a 0-based index or a handle;
      ``len(wb)`` is the sheet count and iterating yields the handles in
      sheet order.

    Lookups:
      A handle or result entry point (:meth:`sheet`, ``table``, ``pivot``,
      ``pivot_cache``, the ``drawing`` family, ``get_chart_data``) raises
      :class:`NotFoundError` when nothing matches and
      :class:`IndexOutOfRangeError` for an index past the end. A getter keyed
      by a natural name (:meth:`get_defined_name`) returns ``None`` instead.

    Recalculation:
      Under the default ``CalcMode.AUTO`` a value read and a save first
      calculate whatever changed, so a read never returns a stale value.
      Under ``CalcMode.MANUAL`` a read returns the value as of the last
      :meth:`recalc`; :meth:`has_dirty_cells` reports whether a change has
      happened since.

    Lifetime:
      Always release via :meth:`close`, ideally through ``with`` syntax::

          with Workbook.create_default() as wb:
              wb[0].set_value("A1", 2)
              wb[0].set_formula("A2", "=A1*10")
              print(wb[0].get_value("A2"))  # 20.0

      :meth:`close` is idempotent and :attr:`closed` reports it. Any
      other call on a closed workbook raises :class:`WorkbookClosedError`.
      The destructor calls ``close`` as a defensive net, but relying on
      garbage-collection ordering is unsafe; prefer the context-manager
      form.

    Threading:
      Each ``Workbook`` instance is single-threaded. Every engine call is
      serialised process-wide by an internal lock, so cross-thread misuse
      cannot corrupt engine state -- it just blocks. The wheel recalculates
      serially; it has no parallel-recalc entry point. There is no
      iterative-progress callback; cap the work with
      ``update_iterative(max_iterations=...)``.
    """

    __slots__ = ("_handle", "_sheets", "_tables", "_pivot_caches", "_ooxml")

    def __init__(self) -> None:
        raise TypeError("use Workbook.create_default(), create_for_profile() or load()")

    @classmethod
    def _new(cls) -> "Workbook":
        # Factories fill the handle in; 0 is the closed state.
        wb = object.__new__(cls)
        wb._handle = 0
        wb._sheets = None
        # One handle per table runtime id and per cache id; an entry outlives its object.
        wb._tables = {}
        wb._pivot_caches = {}
        return wb

    # -- Factories ---------------------------------------------------------
    @classmethod
    def create_default(cls) -> "Workbook":
        """Build a default workbook (a single sheet named ``Sheet1``).

        Example::

            with Workbook.create_default() as wb:
                wb[0].set_number("A1", 42.0)
        """
        wb = cls._new()
        with _ArgScope() as scope:
            out_ptr = scope.out_ptr()
            status = LIB.fm_workbook_create_default(out_ptr)
            _check(status)
            wb._handle = LIB.read_u32(out_ptr)
        return wb

    @classmethod
    def create_for_profile(cls, profile_id: ProfileId) -> "Workbook":
        """Build a new workbook with the defaults of an Excel profile.

        ``profile_id`` has the form ``{mac,win}-365-<locale>`` (see
        :meth:`set_profile_id`). The workbook carries that profile as
        its :meth:`get_profile_id` and takes the profile's locale defaults:
        body font (name, charset, minor theme scheme), phonetic font (font 1,
        ja/ko/zh only), Normal style name, vertical alignment, first sheet
        name, row height, margins and phonetic settings. ``mac-*`` values are
        measured; ``win-*`` ids are estimates (body size 11, no row height or
        base column width set). Unlike :meth:`create_default` no theme part is
        written, so the workbook reads the default Office theme. Changing
        font 0 later (:meth:`set_default_font`) leaves font 1 untouched.
        Unknown ids raise :class:`InvalidArgumentError`.
        """
        wb = cls._new()
        with _ArgScope() as scope:
            pid_ptr = scope.cstr(profile_id)
            out_ptr = scope.out_ptr()
            status = LIB.fm_workbook_create_for_profile(pid_ptr, out_ptr)
            _check(status)
            wb._handle = LIB.read_u32(out_ptr)
        return wb

    @classmethod
    def create_empty(cls) -> "Workbook":
        """Build an empty workbook (no sheets).

        Saving an empty workbook fails just as Excel rejects sheet-less
        archives; callers are expected to add at least one sheet before
        :meth:`save`.

        The style table is seeded with Excel's reserved defaults exactly as
        :meth:`create_default` seeds it, so the first index :meth:`WorkbookOoxml.add_font`
        / :meth:`WorkbookOoxml.add_fill` / :meth:`WorkbookOoxml.add_border` hands back is non-zero.
        """
        wb = cls._new()
        with _ArgScope() as scope:
            out_ptr = scope.out_ptr()
            status = LIB.fm_workbook_create_empty(out_ptr)
            _check(status)
            wb._handle = LIB.read_u32(out_ptr)
        return wb

    @classmethod
    def load(cls, source: Union[bytes, bytearray, memoryview, str, "os.PathLike[str]", BinaryIO]) -> "Workbook":
        """Load a workbook from bytes, a file path or a binary stream.

        Auto-detects ``.xlsx`` vs ``.xlsb`` from the package contents; a
        path's extension is not consulted.

        Args:
          source: the archive as ``bytes`` / ``bytearray`` / ``memoryview``;
            a ``str`` or :class:`os.PathLike` naming a file; or an object
            with a ``read`` method returning bytes (a file opened in
            ``"rb"`` mode, :class:`io.BytesIO`).

        Raises:
          TypeError: when ``source`` is none of the above, or a stream's
            ``read`` does not return bytes.
          OSError: when the file cannot be opened or read
            (``FileNotFoundError`` for a missing path).
          InvalidArgumentError: when the input is empty.
          FormulonError: when the input cannot be parsed as a valid
            OOXML or XLSB archive.

        Example::

            with Workbook.load("book.xlsx") as wb:
                print(wb[0].get_value("A1"))
        """
        if isinstance(source, (bytes, bytearray, memoryview)):
            buf = bytes(source)
        elif isinstance(source, (str, os.PathLike)):
            with open(source, "rb") as f:
                buf = f.read()
        elif callable(getattr(source, "read", None)):
            data = source.read()
            if not isinstance(data, (bytes, bytearray, memoryview)):
                raise TypeError(f"Workbook.load: the stream returned {type(data).__name__}, expected bytes")
            buf = bytes(data)
        else:
            raise TypeError(f"Workbook.load: expected bytes, a path or a binary stream, got {type(source).__name__}")
        wb = cls._new()
        with _ArgScope() as scope:
            # An empty input still passes a real pointer: the engine reports
            # it as empty rather than as a missing buffer.
            data_ptr = scope.buffer(buf) if buf else scope.alloc(1)
            out_ptr = scope.out_ptr()
            status = LIB.fm_workbook_load(data_ptr, _uint(len(buf), "len"), out_ptr)
            _check(status)
            wb._handle = LIB.read_u32(out_ptr)
        return wb

    # -- Lifecycle ---------------------------------------------------------
    def close(self) -> None:
        """Release the underlying handle. Idempotent."""
        if self._handle:
            LIB.fm_workbook_destroy(self._handle)
            self._handle = 0

    def __enter__(self) -> "Workbook":
        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        self.close()

    def __del__(self) -> None:  # pragma: no cover -- GC ordering best-effort
        try:
            self.close()
        except Exception:  # noqa: BLE001 -- destructors must not raise
            pass

    @property
    def closed(self) -> bool:
        """``True`` once :meth:`close` has released the workbook."""
        return self._handle == 0

    def get_memory_usage(self) -> int:
        """Return an estimate of the workbook's heap footprint, in bytes.

        Intended for a host that pools workbooks and needs a size to
        evict on. The figure covers the cell store, the shared-string
        storage, the passthrough part payloads and the workbook-level
        metadata. It **excludes** allocator overhead, the dependency
        graph and the style tables, so it is an estimate for relative
        pressure rather than the process's retained heap; a
        formula-heavy workbook undercounts by the size of its graph.

        The walk visits every materialised cell, so call it at coarse
        boundaries (after a load or a recalc) rather than per edit.
        """
        h = self._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_workbook_get_memory_usage(h, out))
            return LIB.read_u32(out)

    def _require(self) -> int:
        if not self._handle:
            raise FormulonError(
                StatusCode.BINDING_INVALID_HANDLE,
                _diagnostic_override=("the workbook is closed", ""),
            )
        return self._handle

    def _check_live(self) -> int:
        """Re-validate the handle at a generator resume point.

        The iterator methods below hoist their C ABI handle out of the
        per-item loop, and a `yield` in that loop hands control back to
        the caller -- who can call `close()` before asking for the next
        item. Reading `self._handle` under `LIB.call_lock` (rather than
        `_require`'s plain attribute read) also closes the same window
        against a concurrent `close()` from another thread.
        """
        with LIB.call_lock:
            if not self._handle:
                raise FormulonError(
                    StatusCode.BINDING_INVALID_HANDLE,
                    _diagnostic_override=("the workbook is closed", ""),
                )
            return self._handle

    # -- Sheets ------------------------------------------------------------
    def _sheet_count(self) -> int:
        h = self._require()
        return _read_count(LIB.fm_workbook_get_sheet_count, h)

    def _sheet_name(self, index: int) -> str:
        h = self._require()
        return _read_cstr_out(LIB.fm_workbook_get_sheet_name, h, _uint(index, "sheet_index"))

    def _find_sheet(self, name: str) -> Optional[int]:
        """Index of the sheet named ``name`` (case-insensitive, as sheet names are unique), or ``None``."""
        if not isinstance(name, str):
            raise TypeError(f"formulon: name must be a str, got {type(name).__name__}")
        h = self._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            present = scope.out_ptr()
            _check(LIB.fm_workbook_get_sheet_index(h, scope.cstr(name), out, present))
            if not LIB.read_i32(present):
                return None
            return LIB.read_u32(out)

    def _sheet_handles(self) -> List[Sheet]:
        """The live handle list, built from the engine on first use and kept in sheet order after that."""
        from ._sheet import Sheet

        self._require()
        if self._sheets is None:
            self._sheets = [Sheet(self, index, self._sheet_name(index)) for index in range(self._sheet_count())]
        return self._sheets

    def _insert_sheet_handle(self, handles: List[Sheet], index: int) -> Sheet:
        """Track a sheet the engine inserted at ``index``; ``handles`` is the list read before the change."""
        from ._sheet import Sheet

        sheet = Sheet(self, index, self._sheet_name(index))
        handles.insert(index, sheet)
        self._renumber_sheets(handles)
        return sheet

    @staticmethod
    def _renumber_sheets(handles: List[Sheet]) -> None:
        for position, sheet in enumerate(handles):
            sheet._position = position

    def sheet(self, key: SheetRef) -> Sheet:
        """Return the handle of one sheet.

        Every call for the same sheet returns the same :class:`Sheet` object.

        Args:
          key: a sheet name (compared case-insensitively, as sheet names are
            unique), a 0-based sheet index, or a :class:`Sheet` of this
            workbook.

        Raises:
          NotFoundError: when no sheet has the name ``key``, or ``key`` is a
            removed sheet.
          IndexOutOfRangeError: when ``key`` is an index ``>=`` the sheet count.
          ValueError: when ``key`` is a negative index or does not fit 64 bits.
          TypeError: when ``key`` is none of the above, or a :class:`Sheet` of
            another workbook.

        Example::

            ws = wb.sheet("Sheet1")
            ws.set_text("A1", "Total")
        """
        from ._sheet import Sheet

        self._require()
        if isinstance(key, Sheet):
            if key._workbook is not self:
                raise TypeError("formulon: the sheet belongs to another workbook")
            key._index()
            return key
        if isinstance(key, str):
            index = self._find_sheet(key)
            if index is None:
                raise FormulonError(StatusCode.NOT_FOUND, _diagnostic_override=(f"no sheet is named {key!r}", ""))
            return self._sheet_handles()[index]
        if isinstance(key, bool):
            raise TypeError("formulon: sheet key must be a str, an int or a Sheet, got bool")
        try:
            index = operator.index(key)
        except TypeError:
            raise TypeError(f"formulon: sheet key must be a str, an int or a Sheet, got {type(key).__name__}") from None
        index = _uint(index, "sheet_index", bits=64)
        handles = self._sheet_handles()
        count = len(handles)
        if index >= count:
            raise FormulonError(
                StatusCode.SHEET_INDEX_OUT_OF_RANGE,
                _diagnostic_override=(
                    f"sheet index {index} is out of range ({count} {'sheet' if count == 1 else 'sheets'})",
                    f"sheet_index={index} sheet_count={count}",
                ),
            )
        return handles[index]

    def list_sheets(self) -> List[Sheet]:
        """Every sheet's handle, in sheet order, as a new list."""
        return list(self._sheet_handles())

    def __getitem__(self, key: SheetRef) -> Sheet:
        """``wb[key]`` is :meth:`sheet`."""
        return self.sheet(key)

    def __len__(self) -> int:
        """The sheet count."""
        return len(self._sheet_handles())

    def __iter__(self) -> Iterator[Sheet]:
        """Every sheet's handle, in sheet order."""
        return iter(self.list_sheets())

    def __contains__(self, key: object) -> bool:
        """``True`` when ``key`` is a live :class:`Sheet` of this workbook, names a sheet, or is a valid index."""
        from ._sheet import Sheet

        if isinstance(key, Sheet):
            return key._workbook is self and key._position is not None and not self.closed
        if isinstance(key, str):
            return self._find_sheet(key) is not None
        if isinstance(key, bool) or not isinstance(key, numbers.Integral):
            return False
        return 0 <= int(key) < len(self._sheet_handles())

    def __repr__(self) -> str:
        if self.closed:
            return "Workbook(closed)"
        return f"Workbook(sheets={len(self._sheet_handles())})"

    def active_sheet(self) -> Sheet:
        """The active sheet (``activeTab`` of the first workbook view).

        Without workbook views (an XLSB load) the first tab-selected sheet;
        the first sheet when the workbook records neither. Removing or hiding
        the active sheet activates the next visible sheet, else the previous one.
        """
        h = self._require()
        handles = self._sheet_handles()
        return handles[_read_count(LIB.fm_workbook_get_active_sheet, h)]

    def set_active_sheet(self, sheet: SheetRef) -> None:
        """Mark ``sheet`` active so :meth:`save` records it.

        Creates a minimal ``<bookViews>`` when the workbook has none; other
        ``<workbookView>`` attributes survive, and ``firstSheet`` is lowered
        to the sheet's index when larger. The sheet becomes the only
        tab-selected one. A hidden sheet raises.
        """
        index = self.sheet(sheet).index
        h = self._require()
        _check(LIB.fm_workbook_set_active_sheet(h, _uint(index, "sheet_index")))

    def add_sheet(self, name: str) -> Sheet:
        """Append a new sheet with the given UTF-8 display name and return its handle.

        Raises:
          FormulonError: ``StatusCode.INVALID_SHEET_NAME`` when the name is
            empty, longer than 31 UTF-16 units, contains one of ``:``,
            backslash, ``/``, ``?``, ``*``, ``[`` or ``]``, begins or ends
            with an apostrophe, is ``History`` in any case, or collides
            case-insensitively (Unicode simple case folding) with an existing
            sheet; ``StatusCode.SHEET_COUNT_LIMIT_EXCEEDED`` at the maximum
            sheet count.
        """
        h = self._require()
        handles = self._sheet_handles()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            out = scope.out_ptr()
            _check(LIB.fm_workbook_add_sheet(h, name_ptr, out))
            index = LIB.read_u32(out)
        return self._insert_sheet_handle(handles, index)

    def _evaluate_typed_formula(self, sheet: int, row: int, col: int, formula: str) -> Value:
        """Evaluate as :meth:`Sheet.evaluate_typed_formula`: a 1x1 result is its cell, any other an ``ARRAY``."""
        if not isinstance(formula, str):
            raise TypeError(f"formulon: formula must be a str, got {type(formula).__name__}")
        return _folded(self._evaluate_formula_cells(sheet, row, col, formula, top_left_only=False))

    def _evaluate_formula_cells(
        self, sheet: int, row: int, col: int, formula: str, *, top_left_only: bool
    ) -> List[List[Value]]:
        """Evaluate ``formula`` into an owned value array and read its cells (only the first when asked)."""
        h = self._require()
        with _ArgScope() as scope:
            formula_ptr = scope.cstr(formula)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_evaluate_formula_array(
                    h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), formula_ptr, out
                )
            )
            array = LIB.read_u32(out)
            try:
                rows_ptr = scope.out_ptr()
                cols_ptr = scope.out_ptr()
                _check(LIB.fm_value_array_get_dimensions(array, rows_ptr, cols_ptr))
                rows = 1 if top_left_only else LIB.read_u32(rows_ptr)
                cols = 1 if top_left_only else LIB.read_u32(cols_ptr)
                stride = LIB.read_u32(cols_ptr)
                value_ptr = scope.alloc(fm_value_t_size)
                grid: List[List[Value]] = []
                for r in range(rows):
                    out_row: List[Value] = []
                    for c in range(cols):
                        _check(LIB.fm_value_array_get_value_at(array, _uint(r * stride + c, "index"), value_ptr))
                        out_row.append(Value._from_wasm(value_ptr))
                    grid.append(out_row)
                return grid
            finally:
                LIB.fm_value_array_destroy(array)

    # -- Recalc ------------------------------------------------------------
    def recalc(self) -> None:
        """Calculate every dirty cell now.

        Under ``CalcMode.AUTO`` reads already do this; call it under
        ``CalcMode.MANUAL``, or to re-evaluate volatile functions. The wheel
        always recalculates serially.

        Example::

            wb.set_calc_mode(CalcMode.MANUAL)
            wb[0].set_formula("B1", "=A1*2")
            wb.recalc()
        """
        h = self._require()
        status = LIB.fm_workbook_recalc(h)
        _check(status)

    def has_dirty_cells(self) -> bool:
        """Whether any cell is dirty: a change has left its value uncomputed since the last :meth:`recalc`.

        Edits to cells, defined names, the profile or the file name, row /
        column inserts and deletes, and sheet removal make cells dirty;
        volatile functions do not, and neither does a freshly created
        workbook. A loaded workbook with formulas is dirty until its first
        full recalc; :meth:`Sheet.recalc_viewport` does not clean it, and a
        failed recalc leaves it dirty. Under ``CalcMode.AUTO`` a value read
        calculates first, so this is mostly of use under ``CalcMode.MANUAL``.
        """
        h = self._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_workbook_has_dirty_cells(h, out))
            return LIB.read_i32(out) != 0

    def update_iterative(
        self,
        *,
        enabled: Union[bool, Unchanged] = UNCHANGED,
        max_iterations: Union[int, Unchanged] = UNCHANGED,
        max_change: Union[float, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the iterative-calculation settings; an omitted argument keeps its setting.

        ``max_iterations`` of 0 is treated as 1 and values above 32767
        (Excel's limit) as 32767; :meth:`get_iterative` reports the clamped
        value.

        Raises:
          TypeError: when ``enabled`` is not a ``bool`` or ``max_iterations``
            is not an integer.
          ValueError: when ``max_iterations`` or ``max_change`` is negative.
          InvalidArgumentError: when ``max_change`` is NaN or infinite.
        """
        # A negative field keeps its setting.
        fields: Dict[str, object] = {"enabled": _sint(-1, "enabled"), "max_iterations": _sint(-1, "max_iterations")}
        fields["max_change"] = -1.0
        if enabled is not UNCHANGED:
            if not isinstance(enabled, bool):
                raise TypeError(f"formulon: enabled must be a bool, got {type(enabled).__name__}")
            fields["enabled"] = 1 if enabled else 0
        if max_iterations is not UNCHANGED:
            if (
                isinstance(max_iterations, bool)
                or not isinstance(max_iterations, numbers.Real)
                or not math.isfinite(max_iterations)
                or max_iterations != int(max_iterations)
            ):
                raise TypeError(f"formulon: max_iterations must be an integer, got {max_iterations!r}")
            fields["max_iterations"] = _uint(max_iterations, "max_iterations", bits=31)
        if max_change is not UNCHANGED:
            if isinstance(max_change, bool) or not isinstance(max_change, numbers.Real):
                raise TypeError(f"formulon: max_change must be a number, got {type(max_change).__name__}")
            if max_change < 0:
                raise ValueError(f"formulon: max_change must be non-negative, got {max_change!r}")
            fields["max_change"] = float(max_change)
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.ITERATIVE_SETTINGS_PATCH)
            S.ITERATIVE_SETTINGS_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_workbook_update_iterative(h, ptr))

    def get_iterative(self) -> IterativeSettings:
        """Read back the iterative-calculation settings.

        The counterpart of :meth:`update_iterative`; the cap and threshold are
        the stored values even when ``enabled`` is ``False``.
        """
        h = self._require()
        with _ArgScope() as scope:
            enabled_ptr = scope.out_ptr()
            iterations_ptr = scope.out_ptr()
            change_ptr = scope.alloc(8)
            LIB.write_bytes(change_ptr, b"\x00" * 8)
            _check(LIB.fm_workbook_get_iterative(h, enabled_ptr, iterations_ptr, change_ptr))
            return IterativeSettings(
                enabled=bool(LIB.read_i32(enabled_ptr)),
                max_iterations=LIB.read_u32(iterations_ptr),
                max_change=LIB.read_f64(change_ptr),
            )

    # -- Save --------------------------------------------------------------
    @overload
    def save(self, dest: None = None, *, format: Optional[WorkbookFormat] = None) -> bytes: ...

    @overload
    def save(
        self, dest: Union[str, "os.PathLike[str]", BinaryIO], *, format: Optional[WorkbookFormat] = None
    ) -> None: ...

    def save(
        self,
        dest: Union[None, str, "os.PathLike[str]", BinaryIO] = None,
        *,
        format: Optional[WorkbookFormat] = None,
    ) -> Optional[bytes]:
        """Serialise the workbook as ``.xlsx`` or ``.xlsb``.

        Args:
          dest: where to write. Omitted or ``None``, the archive is returned
            as ``bytes``. A ``str`` or :class:`os.PathLike` names a file to
            create or overwrite; an object with a ``write`` method (a file
            opened in ``"wb"`` mode, :class:`io.BytesIO`) receives the bytes.
          format: the container. Omitted, a path ending in ``.xlsx`` or
            ``.xlsb`` (any case) writes that container, bytes and streams
            get XLSX, and any other path extension is an error. ``format``
            decides only the container: ``.xlsm`` with ``WorkbookFormat.XLSX``
            writes a plain XLSX package.

        Returns:
          The archive bytes when ``dest`` is omitted, otherwise ``None``.

        Raises:
          ValueError: when ``format`` is not a :class:`WorkbookFormat` value,
            or is omitted for a path that ends in neither ``.xlsx`` nor ``.xlsb``.
          InvalidArgumentError: for ``WorkbookFormat.UNKNOWN``, or a ``format``
            that contradicts a ``.xlsx`` / ``.xlsb`` extension of ``dest``.
          TypeError: when ``dest`` is none of the above.
          OSError: when writing ``dest`` fails.

        Example::

            wb.save("out.xlsx")
            data = wb.save()  # bytes
        """
        fmt = _save_format(dest, format)
        h = self._require()
        data = _read_owned_buffer(LIB.fm_workbook_save_as, h, _sint(fmt, "format"))
        _write_dest(dest, data)
        return data if dest is None else None

    def save_with_diagnostics(
        self,
        dest: Union[None, str, "os.PathLike[str]", BinaryIO] = None,
        *,
        format: Optional[WorkbookFormat] = None,
    ) -> SaveDiagnostics:
        """Serialise the workbook as :meth:`save` does and return what the save cost.

        ``dest`` and ``format`` follow :meth:`save`; ``SaveDiagnostics.bytes``
        holds the archive whether or not it was also written to ``dest``.

        Raises:
          ValueError: when ``format`` is not a :class:`WorkbookFormat` value,
            or is omitted for a path that ends in neither ``.xlsx`` nor ``.xlsb``.
          InvalidArgumentError: for ``WorkbookFormat.UNKNOWN``, or a ``format``
            that contradicts a ``.xlsx`` / ``.xlsb`` extension of ``dest``.
          TypeError: when ``dest`` is none of the accepted kinds.
          OSError: when writing ``dest`` fails.
        """
        fmt = _save_format(dest, format)
        h = self._require()
        with _ArgScope() as scope:
            out_ptr_ptr = scope.out_ptr()
            out_len_ptr = scope.out_ptr()
            diag_ptr = scope.struct(S.SAVE_DIAGNOSTICS)
            status = LIB.fm_workbook_save_with_diagnostics(h, _sint(fmt, "format"), out_ptr_ptr, out_len_ptr, diag_ptr)
            _check(status)
            data_ptr = LIB.read_u32(out_ptr_ptr)
            data_len = LIB.read_u32(out_len_ptr)
            try:
                data = b"" if data_len == 0 or data_ptr == 0 else LIB.read_bytes(data_ptr, data_len)
            finally:
                if data_ptr:
                    LIB.fm_buffer_destroy(data_ptr)
            d = S.SAVE_DIAGNOSTICS.unpack(LIB, diag_ptr)
        _write_dest(dest, data)
        return SaveDiagnostics(
            bytes=data,
            downgraded_formula_count=d["downgraded_formula_count"],
            deferred_feature_count=d["deferred_feature_count"],
            dropped_part_count=d["dropped_part_count"],
            dropped_relationship_count=d["dropped_relationship_count"],
            renumbered_part_count=d["renumbered_part_count"],
        )

    def get_diagnostics(self) -> ReadDiagnostics:
        """Return the loss counters captured when this workbook was loaded."""
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.READ_DIAGNOSTICS)
            _check(LIB.fm_workbook_get_diagnostics(h, ptr))
            d = S.READ_DIAGNOSTICS.unpack(LIB, ptr)
            return ReadDiagnostics(
                undecoded_formula_count=d["undecoded_formula_count"],
                undecoded_defined_name_count=d["undecoded_defined_name_count"],
                undecoded_part_count=d["undecoded_part_count"],
                skipped_feature_count=d["skipped_feature_count"],
                unknown_content_type_count=d["unknown_content_type_count"],
            )

    def list_defined_names(self) -> List[DefinedName]:
        """Return every defined name in declaration order."""
        h = self._require()
        n = _read_count(LIB.fm_workbook_get_defined_name_count, h)
        # Scratch is hoisted out of the loop: each
        # LIB.alloc/free is a locked WASM call, so per-item slots would
        # make the walk cost O(n) allocator round-trips on top of the
        # unavoidable O(n) ABI reads.
        with _ArgScope() as scope:
            ptr = scope.struct(S.DEFINED_NAME)
            out: List[DefinedName] = []
            for i in range(n):
                S.zero_struct(LIB, S.DEFINED_NAME, ptr)
                _check(LIB.fm_workbook_get_defined_name_at(h, _uint(i, "idx"), ptr))
                d = S.DEFINED_NAME.unpack(LIB, ptr)
                local_sheet_id = d["local_sheet_id"]
                out.append(
                    DefinedName(
                        name=LIB.read_cstr(d["name"]),
                        formula=LIB.read_cstr(d["formula"]),
                        local_sheet_id=None if local_sheet_id < 0 else local_sheet_id,
                        comment=LIB.read_cstr(d["comment"]) or None,
                        hidden=bool(d["hidden"]),
                    )
                )
        return out

    def table(self, name: str) -> Table:
        """Return the handle of the table named ``name`` (compared case-insensitively), on any sheet.

        Table names are unique across the workbook, which is why this lookup
        is workbook-wide; :meth:`Sheet.table` addresses a sheet's tables by
        index as well. Every call for the same table returns the same
        :class:`Table` object.

        Raises:
          TypeError: when ``name`` is not a ``str``.
          NotFoundError: when no table has the name ``name``.
        """
        if not isinstance(name, str):
            raise TypeError(f"formulon: table name must be a str, got {type(name).__name__}")
        h = self._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            sheet_ptr = scope.out_ptr()
            index_ptr = scope.out_ptr()
            present = scope.out_ptr()
            _check(LIB.fm_workbook_find_table(h, name_ptr, sheet_ptr, index_ptr, present))
            if LIB.read_i32(present) == 0:
                raise FormulonError(
                    StatusCode.NOT_FOUND, _diagnostic_override=(f'no table is named "{name}"', f"name={name}")
                )
            sheet, index = LIB.read_u32(sheet_ptr), LIB.read_u32(index_ptr)
        return self._table_handle(sheet, index)

    def _table_count(self, sheet_index: int) -> int:
        """Return the number of tables on sheet ``sheet_index``."""
        h = self._require()
        return _read_count(LIB.fm_sheet_get_table_count, h, _uint(sheet_index, "sheet_index"))

    def _table_handle(self, sheet_index: int, table_index: int) -> Table:
        """The handle of table ``table_index`` on ``sheet_index``, created on first use and kept for the workbook's life."""
        from ._table import Table

        h = self._require()
        runtime_id = _read_count(
            LIB.fm_sheet_get_table_runtime_id, h, _uint(sheet_index, "sheet_index"), _uint(table_index, "table_index")
        )
        handle = self._tables.get(runtime_id)
        if handle is None:
            handle = Table(self, runtime_id)
            self._tables[runtime_id] = handle
        return handle

    def _add_table(self, sheet_index: int, table: TableInput) -> Table:
        """Create the table ``table`` describes on sheet ``sheet_index`` and return its handle.

        The body of :meth:`Sheet.add_table`; see there for the contract.
        """
        if not isinstance(table, TableInput):
            raise TypeError(f"formulon: expected a TableInput, got {type(table).__name__}")
        rect = cell_rect(self, table.range, "TableInput.range")
        if isinstance(table.name, str) and self._names_a_cell(table.name):
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _diagnostic_override=(f'invalid table name: "{table.name}" is a cell address', f"name={table.name}"),
            )
        h = self._require()
        sheet = _uint(sheet_index, "sheet_index")
        with _ArgScope() as scope:
            rect_ptr = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rect_ptr, rect)
            name_ptr = scope.nonempty_cstr(table.name)
            style_ptr = scope.opt_cstr(table.style_name)
            names = [str(column) for column in table.columns or ()]
            name_ptrs = [scope.nonempty_cstr(column) for column in names]
            array_ptr = scope.alloc(4 * len(name_ptrs)) if name_ptrs else 0
            if array_ptr:
                LIB.write_bytes(array_ptr, b"".join(struct.pack("<I", p) for p in name_ptrs))
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_add_table(
                    h,
                    sheet,
                    rect_ptr,
                    name_ptr,
                    array_ptr,
                    _uint(len(name_ptrs), "column_count"),
                    style_ptr,
                    1 if table.header_row else 0,
                    1 if table.totals_row else 0,
                    out,
                )
            )
            index = LIB.read_u32(out)
        handle = self._table_handle(sheet_index, index)
        toggles = {
            name: getattr(table, name)
            for name in ("show_first_column", "show_last_column", "show_row_stripes", "show_column_stripes")
            if getattr(table, name) is not None
        }
        if toggles:
            try:
                with _ArgScope() as scope:
                    ptr = scope.struct(S.TABLE_STYLE_OPTIONS_PATCH)
                    values: Dict[str, int] = {}
                    for name, value in toggles.items():
                        values[name + "_op"] = _PATCH_SET
                        values[name] = 1 if value else 0
                    S.TABLE_STYLE_OPTIONS_PATCH.pack(LIB, ptr, values)
                    _check(LIB.fm_sheet_update_table_style_options(h, sheet, _uint(index, "table_index"), ptr))
            except BaseException:
                _check(LIB.fm_sheet_remove_table(h, sheet, _uint(index, "table_index")))
                raise
        return handle

    def _names_a_cell(self, text: str) -> bool:
        """Whether ``text`` reads as one A1 cell (``"T2"``), which a table name may not be."""
        try:
            ref = self.parse_reference(text)
        except FormulaSyntaxError:
            return False
        rect = ref.rect
        return ref.sheet_name is None and rect.first_row == rect.last_row and rect.first_col == rect.last_col

    # -- Sheet structure ---------------------------------------------------
    def copy_sheet(
        self, source: SheetRef, insert_before: Optional[int] = None, new_name: Optional[str] = None
    ) -> Sheet:
        """Copy the worksheet ``source`` and return the copy's handle.

        The copy goes before the sheet now at ``insert_before``; omitted or
        ``None`` (or the sheet count) appends it, so ``copy_sheet(src, None,
        "Copy")`` names an appended copy. Without ``new_name`` it is named as Excel does
        (``Src (2)``); an explicit name follows the :meth:`add_sheet` rules.
        The active sheet does not change.

        Raises:
          FormulonError: ``StatusCode.SHEET_INDEX_OUT_OF_RANGE``, ``StatusCode.INVALID_SHEET_NAME``,
            ``StatusCode.SHEET_COUNT_LIMIT_EXCEEDED``, ``StatusCode.PRECONDITION_FAILED`` (very hidden
            source) or ``StatusCode.NOT_IMPLEMENTED`` (content the copy cannot model).
        """
        source_index = self.sheet(source).index
        h = self._require()
        handles = self._sheet_handles()
        with _ArgScope() as scope:
            name_ptr = 0 if new_name is None else scope.cstr(new_name)
            index = _read_count(
                LIB.fm_workbook_copy_sheet,
                h,
                _uint(source_index, "source"),
                _uint(len(handles) if insert_before is None else insert_before, "insert_before"),
                name_ptr,
            )
        return self._insert_sheet_handle(handles, index)

    def remove_sheet(self, sheet: SheetRef) -> None:
        """Remove ``sheet``; its handle raises :class:`NotFoundError` from then on.

        Raises:
          FormulonError: with ``StatusCode.CANNOT_REMOVE_LAST_SHEET`` when it is the
            workbook's only sheet, and with ``StatusCode.INVALID_ARGUMENT`` when it is
            the only visible one; a workbook keeps at least one visible sheet.
        """
        removed = self.sheet(sheet)
        index = removed.index
        h = self._require()
        handles = self._sheet_handles()
        _check(LIB.fm_workbook_remove_sheet(h, _uint(index, "sheet_index")))
        handles.pop(index)
        removed._position = None
        self._renumber_sheets(handles)

    # -- Defined names -----------------------------------------------------
    def set_defined_name(self, defined_name: DefinedNameInput) -> None:
        """Set or append a defined name in workbook or sheet scope.

        An existing name in the same scope is replaced as a whole: an
        omitted ``comment`` or ``hidden`` goes back to its default.

        Raises:
          InvalidArgumentError: when ``formula`` is empty (remove a name with
            :meth:`remove_defined_name`).
        """
        local_sheet_id = _local_sheet_id(None if defined_name.scope is None else self.sheet(defined_name.scope).index)
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.DEFINED_NAME_INPUT)
            S.DEFINED_NAME_INPUT.pack(
                LIB,
                ptr,
                {
                    "name": scope.cstr(defined_name.name),
                    "formula": scope.cstr(defined_name.formula),
                    "local_sheet_id": _sint(local_sheet_id, "local_sheet_id"),
                    "comment": scope.nonempty_cstr(defined_name.comment),
                    "hidden": 1 if defined_name.hidden else 0,
                },
            )
            _check(LIB.fm_workbook_set_defined_name(h, ptr))

    # -- Row / column structural edits -------------------------------------
    #
    # The four edits below move every structure the engine models: cells,
    # formulas, merges, conditional-format ranges, validations, hyperlinks,
    # tables, print areas, manual breaks and the auto-filter range. The
    # ranges and formulas inside the worksheet ``<extLst>`` (the ``x14``
    # conditional-formatting block, sparkline groups, ``x14`` validations)
    # move with them, as Excel moves them, and so do the cells named by the
    # other worksheet content kept verbatim for a lossless save: protected
    # ranges, scenarios, a sheet sort state, cell watches, ignored-error
    # ranges and web-publish ranges.
    #
    # They do not move the anchor of an embedded OLE object or form control,
    # also kept verbatim: it keeps its pre-edit cell, and nothing reports it.
    # An insert also does not copy the formatting of the row or column before
    # it onto the new one, which Excel does -- a sparkline included.
    def get_defined_name(self, name: str, scope: Optional[SheetRef] = None) -> Optional[DefinedName]:
        """Return the defined name ``name`` (ASCII case-insensitive) in ``scope``, or ``None``.

        ``scope`` is ``None`` for workbook scope, else the sheet the name is local to.
        """
        sheet = None if scope is None else self.sheet(scope).index
        h = self._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            present = scope.out_ptr()
            stored_ptr = scope.out_ptr()
            formula_ptr = scope.out_ptr()
            comment_ptr = scope.out_ptr()
            hidden_ptr = scope.out_ptr()
            _check(
                LIB.fm_workbook_get_defined_name(
                    h,
                    name_ptr,
                    _sint(_local_sheet_id(sheet), "local_sheet_id"),
                    present,
                    stored_ptr,
                    formula_ptr,
                    comment_ptr,
                    hidden_ptr,
                )
            )
            if not LIB.read_u32(present):
                return None
            return DefinedName(
                name=LIB.read_cstr(LIB.read_u32(stored_ptr)),
                formula=LIB.read_cstr(LIB.read_u32(formula_ptr)),
                local_sheet_id=sheet,
                comment=LIB.read_cstr(LIB.read_u32(comment_ptr)) or None,
                hidden=bool(LIB.read_i32(hidden_ptr)),
            )

    def remove_defined_name(self, name: str, scope: Optional[SheetRef] = None) -> None:
        """Remove the defined name ``name`` in ``scope`` (``None`` for workbook scope, else a sheet).

        Raises:
          NotFoundError: when the scope has no such name.
        """
        sheet = None if scope is None else self.sheet(scope).index
        h = self._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            _check(LIB.fm_workbook_remove_defined_name(h, name_ptr, _sint(_local_sheet_id(sheet), "local_sheet_id")))

    # -- Calc policy / behaviour profile -----------------------------------
    def get_calc_mode(self) -> CalcMode:
        """Return the workbook's calc mode (see :class:`CalcMode`)."""
        h = self._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_workbook_get_calc_mode(h, out))
            return CalcMode(LIB.read_i32(out))

    def set_calc_mode(self, mode: CalcMode) -> None:
        """Set the workbook's calc mode (see :class:`CalcMode`).

        To read the values a loaded file has cached, set ``CalcMode.MANUAL``
        right after :meth:`load`; to keep them through a save as well, also
        :meth:`set_calc_on_save` ``False``.
        """
        h = self._require()
        _check(LIB.fm_workbook_set_calc_mode(h, _sint(mode, "mode")))

    def _read_flag(self, fn) -> bool:
        h = self._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(fn(h, out))
            return bool(LIB.read_i32(out))

    def get_full_precision(self) -> bool:
        """Return ``<calcPr fullPrecision>``; false is "set precision as displayed" (metadata only)."""
        return self._read_flag(LIB.fm_workbook_get_full_precision)

    def set_full_precision(self, full_precision: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set ``<calcPr fullPrecision>``."""
        h = self._require()
        _check(LIB.fm_workbook_set_full_precision(h, 1 if full_precision else 0))

    def get_calc_on_save(self) -> bool:
        """Return ``<calcPr calcOnSave>`` (default true)."""
        return self._read_flag(LIB.fm_workbook_get_calc_on_save)

    def set_calc_on_save(self, calc_on_save: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set ``<calcPr calcOnSave>``."""
        h = self._require()
        _check(LIB.fm_workbook_set_calc_on_save(h, 1 if calc_on_save else 0))

    def get_ref_mode(self) -> RefMode:
        """Return ``<calcPr refMode>``: the reference notation formulas are shown in."""
        h = self._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_workbook_get_ref_mode(h, out))
            return RefMode(LIB.read_i32(out))

    def set_ref_mode(self, mode: RefMode) -> None:
        """Set ``<calcPr refMode>``."""
        h = self._require()
        _check(LIB.fm_workbook_set_ref_mode(h, _sint(mode, "mode")))

    def get_date1904(self) -> bool:
        """Return whether the workbook uses the 1904 date system."""
        return self._read_flag(LIB.fm_workbook_get_date1904)

    def set_date1904(self, date1904: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Switch the date system. Stored serials keep their numbers, so every date shifts by 1462 days."""
        h = self._require()
        _check(LIB.fm_workbook_set_date1904(h, 1 if date1904 else 0))

    def get_pinned_now(self) -> Optional[datetime.datetime]:
        """Return the pinned wall-clock reading as a naive ``datetime``, or
        ``None`` when the workbook follows the host clock (the default)."""
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CIVIL_TIME)
            flag = scope.out_ptr()
            _check(LIB.fm_workbook_get_pinned_now(h, ptr, flag))
            if LIB.read_i32(flag) == 0:
                return None
            d = S.CIVIL_TIME.unpack(LIB, ptr)
            return datetime.datetime(d["year"], d["month"], d["day"], d["hour"], d["minute"], d["second"])

    def set_pinned_now(self, t: Optional[Union[datetime.datetime, datetime.date]]) -> None:
        """Pin every clock-dependent result to the local time ``t``; ``None`` releases the pin.

        A ``date`` pins midnight and a naive ``datetime`` is read as the local
        wall clock; an aware ``datetime`` is a ``ValueError``.

        ``NOW()``, ``TODAY()`` and the pivot relative-period filters ("this
        month", "year to date", ...) otherwise each read the clock
        independently, which makes a recalc internally inconsistent across a
        midnight boundary and makes any such result untestable. Fractions of
        a second are dropped.

        This is model state, not file state: nothing in the OOXML package
        records it, so :meth:`save` drops it. Setting or clearing the pin
        marks every formula cell dirty, as a profile or iterative change
        does; in ``CalcMode.AUTO`` the next read recomputes, in
        ``CalcMode.MANUAL`` :meth:`recalc` does.

        Raises:
          TypeError: when ``t`` is not a ``datetime.date`` or ``datetime.datetime``.
          InvalidArgumentError: unless the year is in ``[1900, 9999]``.
        """
        h = self._require()
        if t is None:
            _check(LIB.fm_workbook_set_pinned_now(h, 0))
            return
        t = civil_datetime(t, "the pinned time")
        with _ArgScope() as scope:
            ptr = scope.struct(S.CIVIL_TIME)
            S.CIVIL_TIME.pack(
                LIB,
                ptr,
                {
                    "year": _sint(t.year, "year"),
                    "month": _sint(t.month, "month"),
                    "day": _sint(t.day, "day"),
                    "hour": _sint(t.hour, "hour"),
                    "minute": _sint(t.minute, "minute"),
                    "second": _sint(t.second, "second"),
                },
            )
            _check(LIB.fm_workbook_set_pinned_now(h, ptr))

    def get_profile_id(self) -> ProfileId:
        """Return the id of the workbook's Excel profile.

        New workbooks default to ``win-365-en_US``. Select another host or
        locale with :meth:`set_profile_id`.
        """
        h = self._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_workbook_get_profile_id(h, out))
            return cast(ProfileId, LIB.read_cstr(LIB.read_u32(out)))

    def set_profile_id(self, profile_id: ProfileId) -> None:
        """Set the workbook's Excel profile by id.

        New workbooks use ``win-365-en_US`` by default. Ids have the form
        ``{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}``.
        A change marks every formula cell dirty.
        """
        h = self._require()
        with _ArgScope() as scope:
            pid_ptr = scope.cstr(profile_id)
            _check(LIB.fm_workbook_set_profile_id(h, pid_ptr))

    def get_file_name(self) -> Optional[str]:
        """Return the file name the workbook is shown under, or ``None`` when unset.

        While set, every formula getter spells a reference to the workbook's
        own defined name (stored ``[0]!Name``) as ``Book.xlsx!Name``, quoted
        when needed. Loading from bytes leaves it unset.
        """
        h = self._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_workbook_get_file_name(h, out))
            return LIB.read_cstr(LIB.read_u32(out)) or None

    def set_file_name(self, name: Optional[str]) -> None:
        """Set :meth:`get_file_name`, extension included and without a directory; ``None`` (or ``""``) unsets it.

        Entered formulas then read ``Book.xlsx!Name`` and ``[Book.xlsx]Sheet1!A1``
        as this workbook and create no external link. A changed name marks
        every formula cell dirty.

        Raises:
          FormulonError: with ``StatusCode.INVALID_ARGUMENT`` when the name lacks a
            workbook extension, holds one of ``[ ] : / \\ ? *``, is longer
            than 218 bytes, or equals an external link's book name.
        """
        h = self._require()
        with _ArgScope() as scope:
            _check(LIB.fm_workbook_set_file_name(h, scope.cstr(name or "")))

    # -- Protection records ------------------------------------------------
    _WORKBOOK_PROTECTION_TEXT = (
        "workbook_legacy_password_hash",
        "workbook_algorithm_name",
        "workbook_hash_value",
        "workbook_salt_value",
        "revisions_legacy_password_hash",
        "revisions_algorithm_name",
        "revisions_hash_value",
        "revisions_salt_value",
    )

    def get_workbook_protection(self) -> WorkbookProtection:
        """Read the workbook's ``<workbookProtection>``; without the element every flag is false and every count zero."""
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.WORKBOOK_PROTECTION)
            _check(LIB.fm_workbook_get_workbook_protection(h, ptr))
            d = S.WORKBOOK_PROTECTION.unpack(LIB, ptr)
            return WorkbookProtection(
                lock_structure=bool(d["lock_structure"]),
                lock_windows=bool(d["lock_windows"]),
                lock_revision=bool(d["lock_revision"]),
                workbook_spin_count=d["workbook_spin_count"],
                revisions_spin_count=d["revisions_spin_count"],
                **{fld: LIB.read_cstr(d[fld]) or None for fld in self._WORKBOOK_PROTECTION_TEXT},
            )

    def set_workbook_protection(self, protection: Optional[WorkbookProtection]) -> None:
        """Replace the modelled ``<workbookProtection>`` attributes.

        ``None``, like an all-default record, removes the element.
        """
        if protection is None:
            protection = WorkbookProtection()
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.WORKBOOK_PROTECTION)
            S.WORKBOOK_PROTECTION.pack(
                LIB,
                ptr,
                {
                    "lock_structure": 1 if protection.lock_structure else 0,
                    "lock_windows": 1 if protection.lock_windows else 0,
                    "lock_revision": 1 if protection.lock_revision else 0,
                    "workbook_spin_count": _uint(protection.workbook_spin_count, "workbook_spin_count"),
                    "revisions_spin_count": _uint(protection.revisions_spin_count, "revisions_spin_count"),
                },
            )
            for fld in self._WORKBOOK_PROTECTION_TEXT:
                S.write_str_field(LIB, ptr, S.WORKBOOK_PROTECTION, fld, getattr(protection, fld), scope)
            _check(LIB.fm_workbook_set_workbook_protection(h, ptr))

    # -- AutoFilter ----------------------------------------------------------
    #
    # ``indices`` are the leading ``size_t`` arguments that address the filter:
    # ``(sheet_index,)`` for a sheet filter, ``(sheet_index, table_index)`` for
    # a table's.
    def _read_auto_filter(self, fn, *indices: int) -> Optional[AutoFilter]:
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.AUTO_FILTER)
            present = scope.out_ptr()
            _check(fn(h, *(_uint(i, "index") for i in indices), ptr, present))
            if LIB.read_u32(present) == 0:
                return None
            return _decode_auto_filter(ptr)

    def _write_auto_filter(self, fn, auto_filter: Optional[AutoFilter], *indices: int) -> None:
        """Replace the filter; ``None`` removes it (a no-op without one)."""
        h = self._require()
        with _ArgScope() as scope:
            ptr = 0
            if auto_filter is not None:
                ptr = scope.struct_array(S.AUTO_FILTER, 1)
                _pack_auto_filter(ptr, auto_filter, scope)
            _check(fn(h, *(_uint(i, "index") for i in indices), ptr))

    def _evaluate_auto_filter(self, fn, *indices: int) -> "tuple[int, List[bool]]":
        h = self._require()
        args = tuple(_uint(i, "index") for i in indices)
        with _ArgScope() as scope:
            count_ptr = scope.out_ptr()
            first_ptr = scope.out_ptr()
            buf = 0
            # Size the buffer first, then fetch.
            _check(fn(h, *args, 0, 0, count_ptr, first_ptr))
            total = LIB.read_u32(count_ptr)
            if total == 0:
                return LIB.read_u32(first_ptr), []
            buf = scope.alloc(total)
            _check(fn(h, *args, buf, total, count_ptr, first_ptr))
            flags = LIB.read_bytes(buf, total)
            return LIB.read_u32(first_ptr), [b != 0 for b in flags]

    def _lookup(self, fn, leading: tuple, layout, decode):
        """Run the by-key getter ``fn(h, *leading, out_present, out)``; ``None`` when the key is absent.

        A ``str`` in ``leading`` is passed as a C string.
        """
        h = self._require()
        with _ArgScope() as scope:
            args = [scope.cstr(a) if isinstance(a, str) else a for a in leading]
            present = scope.out_ptr()
            ptr = scope.struct(layout)
            _check(fn(h, *args, present, ptr))
            if not LIB.read_u32(present):
                return None
            return decode(layout.unpack(LIB, ptr))

    def recalc_data_tables(self) -> None:
        """Recompute every data table now, whatever the calc mode (the explicit calculation of ``AUTO_NO_TABLE``)."""
        _check(LIB.fm_workbook_recalc_data_tables(self._require()))

    def list_people(self) -> List[Person]:
        """Return the workbook's person list in file order."""
        h = self._require()
        n = _read_count(LIB.fm_workbook_get_person_count, h)
        with _ArgScope() as scope:
            ptr = scope.struct(S.PERSON)
            out: List[Person] = []
            for i in range(n):
                S.zero_struct(LIB, S.PERSON, ptr)
                _check(LIB.fm_workbook_get_person_at(h, _uint(i, "idx"), ptr))
                d = S.PERSON.unpack(LIB, ptr)
                out.append(_person_from(d))
        return out

    def add_person(self, person: Person) -> None:
        """Append a person to the workbook's person list."""
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.PERSON, 1)
            for fld in ("id", "display_name", "user_id", "provider_id"):
                S.write_str_field(LIB, ptr, S.PERSON, fld, getattr(person, fld), scope)
            _check(LIB.fm_workbook_add_person(h, ptr))

    def get_person(self, person_id: str) -> Optional[Person]:
        """Return the person ``person_id``, or ``None`` when the list has none."""
        return self._lookup(LIB.fm_workbook_get_person, (person_id,), S.PERSON, _person_from)

    def remove_person(self, person_id: str) -> None:
        """Remove the person ``person_id`` (refused while a comment or mention still names them)."""
        h = self._require()
        with _ArgScope() as scope:
            id_ptr = scope.cstr(person_id)
            _check(LIB.fm_workbook_remove_person(h, id_ptr))

    def format_value(self, value: Value, format_code: str = "") -> DisplayText:
        """Render ``value`` under the number format ``format_code`` (empty means General)."""
        h = self._require()
        _check_value(value)
        with _ArgScope() as scope:
            value_ptr = scope.alloc(fm_value_t_size)
            text_ptr = scope.out_ptr()
            status_ptr = scope.out_ptr()
            _write_value_input(value, value_ptr, scope)
            fmt_ptr = scope.cstr(format_code)
            _check(LIB.fm_workbook_format_value(h, value_ptr, fmt_ptr, text_ptr, status_ptr))
            return DisplayText(LIB.read_cstr(LIB.read_u32(text_ptr)), DisplayStatus(LIB.read_i32(status_ptr)))

    # -- Print settings: raw XML -------------------------------------------
    #
    # The five worksheet print elements are stored verbatim and re-emitted
    # on save, so these getters return what the file says rather than a
    # re-serialisation of a partial model. An absent element reads as "".
    #
    # Setting the empty string removes the element and restores the
    # structured defaults it fed; any non-empty value must be exactly one
    # well-formed element with the matching root name.

    # -- Print settings: area and titles -----------------------------------

    # -- Print settings: manual page breaks --------------------------------

    # -- Print settings: typed patch setters -------------------------------
    #
    # Each keyword defaults to ``None`` meaning "leave this attribute as the
    # file has it". That is how the C ``*_engaged`` flags are expressed in
    # Python: a caller states the two or three attributes it cares about and
    # every other one, including attributes the engine does not model, stays
    # exactly where the file put it.

    @staticmethod
    def _decode_cell_xf(ptr: int) -> CellXf:
        d = S.CELL_XF.unpack(LIB, ptr)
        return CellXf(
            font_index=d["font_index"],
            fill_index=d["fill_index"],
            border_index=d["border_index"],
            num_fmt_id=d["num_fmt_id"],
            horizontal=_enum(HorizontalAlign, d["horizontal"]) if d["has_horizontal"] else None,
            vertical=_enum(VerticalAlign, d["vertical"]) if d["has_vertical"] else None,
            wrap_text=bool(d["wrap_text"]) if d["has_wrap_text"] else None,
            justify_last_line=bool(d["justify_last_line"]) if d["has_justify_last_line"] else None,
            style_xf_index=d["style_xf_index"],
            text_rotation=d["text_rotation"] if d["has_text_rotation"] else None,
            indent=d["indent"] if d["has_indent"] else None,
            relative_indent=d["relative_indent"] if d["has_relative_indent"] else None,
            shrink_to_fit=bool(d["shrink_to_fit"]) if d["has_shrink_to_fit"] else None,
            reading_order=_enum(ReadingOrder, d["reading_order"]) if d["has_reading_order"] else None,
            apply_number_format=bool(d["apply_number_format"]),
            apply_font=bool(d["apply_font"]),
            apply_fill=bool(d["apply_fill"]),
            apply_border=bool(d["apply_border"]),
            apply_alignment=bool(d["apply_alignment"]),
            apply_protection=bool(d["apply_protection"]),
            quote_prefix=bool(d["quote_prefix"]),
            locked=bool(d["locked"]) if d["has_protection"] else None,
            hidden=bool(d["hidden"]) if d["has_protection"] else None,
        )

    def _get_style_record(self, fn, layout: S.Struct, index: int, index_field: str, decode):
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(layout)
            _check(fn(h, _uint(index, index_field), ptr))
            return decode(ptr)

    def _style_count(self, fn) -> int:
        h = self._require()
        return _read_count(fn, h)

    def set_default_font(self, record: Font) -> None:
        """Declare the workbook's default font (font 0).

        Font 0 is what an unstyled cell resolves to. A new workbook seeds it
        with Excel's Calibri 11, and :meth:`WorkbookOoxml.add_font` can only append beside
        it, so this is the way a ja-JP host declares e.g. ``游ゴシック`` for
        the cells it never styles explicitly. Read it back with
        :meth:`get_default_font`.
        """
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.FONT_RECORD)
            _pack_font(ptr, record, scope)
            _check(LIB.fm_workbook_set_default_font(h, ptr))

    def get_default_font(self) -> Font:
        """Return the workbook's default font (font 0), the font an unstyled cell resolves to."""
        return self.ooxml.get_font(0)  # type: ignore[attr-defined]

    def list_cell_styles(self) -> List[CellStyle]:
        """Return the named cell styles in file order."""
        h = self._require()
        n = _read_count(LIB.fm_workbook_get_cell_style_count, h)
        with _ArgScope() as scope:
            ptr = scope.struct(S.CELL_STYLE_RECORD)
            out: List[CellStyle] = []
            for i in range(n):
                S.zero_struct(LIB, S.CELL_STYLE_RECORD, ptr)
                _check(LIB.fm_workbook_get_cell_style_at(h, _uint(i, "index"), ptr))
                out.append(_cell_style_from(S.CELL_STYLE_RECORD.unpack(LIB, ptr)))
        return out

    def get_cell_style(self, name: str) -> Optional[CellStyle]:
        """Return the named cell style ``name`` (matched exactly), or ``None``."""
        return self._lookup(LIB.fm_workbook_get_cell_style, (name,), S.CELL_STYLE_RECORD, _cell_style_from)

    def set_cell_style(self, record: CellStyle) -> None:
        """Add or replace the named cell style ``record.name``.

        ``record.style_xf_index`` must reference an existing named-style xf --
        register one with :meth:`WorkbookOoxml.add_cell_style_xf` first.
        ``builtin_id`` is an Excel built-in style ordinal (0..53); leave it at
        ``CELL_STYLE_BUILTIN_ID_NONE`` for a custom style.
        """
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CELL_STYLE_RECORD)
            S.CELL_STYLE_RECORD.pack(
                LIB,
                ptr,
                {
                    "style_xf_index": _uint(record.style_xf_index, "style_xf_index"),
                    "builtin_id": _uint(record.builtin_id, "builtin_id"),
                    "i_level": _uint(record.i_level, "i_level"),
                    "hidden": 1 if record.hidden else 0,
                    "custom_builtin": 1 if record.custom_builtin else 0,
                },
            )
            S.write_str_field(LIB, ptr, S.CELL_STYLE_RECORD, "name", record.name, scope, allow_empty=True)
            _check(LIB.fm_workbook_set_cell_style(h, ptr))

    def remove_cell_style(self, name: str) -> None:
        """Remove the named cell style ``name``.

        Cell xfs inheriting from it are repointed at Normal; no xf index
        shifts. The Normal style and unknown names raise ``FormulonError``.
        """
        h = self._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            _check(LIB.fm_workbook_remove_cell_style(h, name_ptr))

    def _add_xf_record(self, record: CellXfInput, fn) -> int:
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CELL_XF)
            out = scope.out_ptr()
            S.CELL_XF.pack(LIB, ptr, _cell_xf_fields(record))
            _check(fn(h, ptr, out))
            return LIB.read_u32(out)

    def list_table_styles(self) -> List[TableStyle]:
        """Return the custom table / PivotTable styles in file order."""
        h = self._require()
        n = _read_count(LIB.fm_workbook_get_table_style_count, h)
        with _ArgScope() as scope:
            ptr = scope.struct(S.TABLE_STYLE)
            out: List[TableStyle] = []
            for i in range(n):
                S.zero_struct(LIB, S.TABLE_STYLE, ptr)
                _check(LIB.fm_workbook_get_table_style_at(h, _uint(i, "index"), ptr))
                out.append(_table_style_from(S.TABLE_STYLE.unpack(LIB, ptr)))
        return out

    def get_table_style(self, name: str) -> Optional[TableStyle]:
        """Return the custom table style ``name`` (case-insensitive), or ``None``."""
        return self._lookup(LIB.fm_workbook_get_table_style, (name,), S.TABLE_STYLE, _table_style_from)

    def set_table_style(self, style: TableStyle) -> None:
        """Add a custom table style, or replace the one with the same name.

        The name may not be a built-in style name; each element type may
        appear once, ``size`` must be at least 1 and ``dxf_index`` must name an
        existing differential format (:meth:`WorkbookOoxml.add_dxf`).
        """
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.TABLE_STYLE)
            elements_ptr = 0
            if style.elements:
                elements_ptr = scope.struct_array(S.TABLE_STYLE_ELEMENT, len(style.elements))
                for i, element in enumerate(style.elements):
                    S.TABLE_STYLE_ELEMENT.pack(
                        LIB,
                        elements_ptr + i * S.TABLE_STYLE_ELEMENT.size,
                        {
                            "type": _sint(element.type, "type"),
                            "size": _uint(element.size, "size"),
                            "has_dxf_index": 0 if element.dxf_index is None else 1,
                            "dxf_index": _uint(0 if element.dxf_index is None else element.dxf_index, "dxf_index"),
                        },
                    )
            name_ptr = scope.cstr(style.name)
            S.TABLE_STYLE.pack(
                LIB,
                ptr,
                {
                    "name": name_ptr,
                    "pivot": 1 if style.pivot else 0,
                    "table": 1 if style.table else 0,
                    "elements": elements_ptr,
                    "element_count": _uint(len(style.elements), "element_count"),
                },
            )
            _check(LIB.fm_workbook_set_table_style(h, ptr))

    def remove_table_style(self, name: str) -> None:
        """Remove the custom table style ``name``.

        A default naming it reverts to the built-in default; tables and
        pivots naming it keep the name.
        """
        h = self._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            _check(LIB.fm_workbook_remove_table_style(h, name_ptr))

    def get_table_style_defaults(self) -> TableStyleDefaults:
        """Return the default table and pivot style names; a name is ``None`` when unset."""
        h = self._require()
        with _ArgScope() as scope:
            table_ptr = scope.out_ptr()
            pivot_ptr = scope.out_ptr()
            _check(LIB.fm_workbook_get_table_style_defaults(h, table_ptr, pivot_ptr))
            return TableStyleDefaults(
                table_style=LIB.read_cstr(LIB.read_u32(table_ptr)) or None,
                pivot_style=LIB.read_cstr(LIB.read_u32(pivot_ptr)) or None,
            )

    def update_table_style_defaults(
        self,
        *,
        table_style: Union[str, Unchanged, None] = UNCHANGED,
        pivot_style: Union[str, Unchanged, None] = UNCHANGED,
    ) -> None:
        """Set the default style names; an omitted argument keeps its name and ``None`` (or ``""``) removes it."""
        h = self._require()
        with _ArgScope() as scope:
            # The C call keeps a NULL name and removes an empty one.
            ptrs = [0 if value is UNCHANGED else scope.cstr(value or "") for value in (table_style, pivot_style)]
            _check(LIB.fm_workbook_update_table_style_defaults(h, ptrs[0], ptrs[1]))

    def get_theme_fonts(self) -> ThemeFonts:
        """Read the theme's major (heading) and minor (body) typefaces.

        The default theme is reported when the part is absent or unparseable;
        :attr:`ThemeColors.source` of :meth:`get_theme_colors` tells the cases apart.
        """
        h = self._require()
        with _ArgScope() as scope:
            fonts_ptr = scope.struct(S.THEME_FONTS)
            _check(LIB.fm_workbook_get_theme_fonts(h, fonts_ptr))
            f = S.THEME_FONTS.unpack(LIB, fonts_ptr)
            return ThemeFonts(
                major_latin=LIB.read_cstr(f["major_latin"]),
                major_east_asian=LIB.read_cstr(f["major_east_asian"]),
                minor_latin=LIB.read_cstr(f["minor_latin"]),
                minor_east_asian=LIB.read_cstr(f["minor_east_asian"]),
            )

    def get_theme_colors(self) -> ThemeColors:
        """Read the theme's 12 scheme colours and where they came from.

        The default theme is reported when the part is absent or unparseable;
        ``source`` tells the cases apart.
        """
        h = self._require()
        with _ArgScope() as scope:
            colors_ptr = scope.struct(S.THEME_COLORS)
            source_ptr = scope.out_ptr()
            _check(LIB.fm_workbook_get_theme_colors(h, colors_ptr, source_ptr))
            return ThemeColors(
                source=ThemeSource(LIB.read_i32(source_ptr)),
                colors=list(struct.unpack("<12I", LIB.read_bytes(colors_ptr + S.THEME_COLORS.offsets["argb"][1], 48))),
            )

    def set_theme_colors(self, *, colors: Sequence[int]) -> None:
        """Replace the theme's 12 scheme colours (``a:clrScheme`` order; alpha ignored)."""
        h = self._require()
        values = [_uint(c, "colors") for c in colors]
        if len(values) != 12:
            raise ValueError(f"formulon: colors must hold 12 entries, got {len(values)}")
        with _ArgScope() as scope:
            ptr = scope.struct(S.THEME_COLORS)
            LIB.write_bytes(ptr + S.THEME_COLORS.offsets["argb"][1], struct.pack("<12I", *values))
            _check(LIB.fm_workbook_set_theme_colors(h, ptr))

    def set_theme_fonts(self, fonts: ThemeFonts) -> None:
        """Replace the theme's major and minor typefaces."""
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.THEME_FONTS)
            for fname in ("major_latin", "major_east_asian", "minor_latin", "minor_east_asian"):
                S.write_str_field(LIB, ptr, S.THEME_FONTS, fname, getattr(fonts, fname), scope, allow_empty=True)
            _check(LIB.fm_workbook_set_theme_fonts(h, ptr))

    def get_theme_script_font(self, *, scheme: FontScheme, script: str) -> Optional[ThemeScriptFont]:
        """Return the entry for ``script`` in the major (``FontScheme.MAJOR``, heading) or minor (``FontScheme.MINOR``,
        body) font set, or ``None``.

        Raises:
          FormulonError: ``StatusCode.INVALID_ARGUMENT`` for ``FontScheme.NONE`` or any other value.
        """
        h = self._require()
        with _ArgScope() as scope:
            script_ptr = scope.cstr(script)
            present = scope.out_ptr()
            face_ptr = scope.out_ptr()
            _check(LIB.fm_workbook_get_theme_script_font(h, _sint(scheme, "scheme"), script_ptr, present, face_ptr))
            if not LIB.read_u32(present):
                return None
            return ThemeScriptFont(
                scheme=_enum(FontScheme, scheme), script=script, typeface=LIB.read_cstr(LIB.read_u32(face_ptr))
            )

    def set_theme_script_font(self, *, scheme: FontScheme, script: str, typeface: str) -> None:
        """Set the typeface of one script in the major (heading) or minor (body) font set.

        ``scheme`` picks the heading set (``FontScheme.MAJOR``) or the body set
        (``FontScheme.MINOR``); ``FontScheme.NONE`` raises. ``script`` is an ISO 15924 code
        of exactly four ASCII letters (``Jpan``, ``Hang``); ``typeface`` must be non-empty. Other
        scripts are unchanged and an unlisted script is appended. Setting
        ``Jpan`` also changes ``ThemeFonts.*_east_asian``. A workbook without a
        theme part first gets one carrying the full default set. Invalid
        arguments raise ``FormulonError``.
        """
        h = self._require()
        with _ArgScope() as scope:
            script_ptr = scope.cstr(script)
            face_ptr = scope.cstr(typeface)
            _check(LIB.fm_workbook_set_theme_script_font(h, _sint(scheme, "scheme"), script_ptr, face_ptr))

    def reset_theme(self) -> None:
        """Restore the default Office theme by removing the theme part.

        :meth:`get_theme_colors` then reports ``ThemeSource.DEFAULT``. Excel adds a
        theme part again when it saves the file.
        """
        h = self._require()
        _check(LIB.fm_workbook_reset_theme(h))

    def resolve_color(self, spec: Color, context: ColorContext = ColorContext.FONT) -> ResolvedColor:
        """Resolve ``spec`` to the ``0xAARRGGBB`` Excel renders.

        ``context`` decides what an automatic or absent colour means (black
        for font and border, white for fills).
        """
        h = self._require()
        with _ArgScope() as scope:
            spec_ptr = scope.struct(S.COLOR_SPEC)
            argb_ptr = scope.out_ptr()
            res_ptr = scope.out_ptr()
            _pack_color(spec_ptr, spec)
            _check(LIB.fm_workbook_resolve_color(h, spec_ptr, _sint(context, "context"), argb_ptr, res_ptr))
            return ResolvedColor(LIB.read_u32(argb_ptr), ColorResolution(LIB.read_i32(res_ptr)))

    # -- Pivot caches ------------------------------------------------------
    def list_pivot_caches(self) -> List[PivotCache]:
        """The handles of the pivot caches the workbook owns."""
        h = self._require()
        count = _read_count(LIB.fm_workbook_get_pivot_cache_count, h)
        return [
            self._pivot_cache_handle(_read_count(LIB.fm_workbook_get_pivot_cache_id_at, h, _uint(index, "idx")))
            for index in range(count)
        ]

    def pivot_cache(self, cache_id: Union[int, PivotCache]) -> PivotCache:
        """Return the handle of the pivot cache ``cache_id``.

        Every call for the same cache id returns the same :class:`PivotCache` object.

        Raises:
          NotFoundError: when the workbook has no cache with that id.
          InvalidArgumentError: when ``cache_id`` is a :class:`PivotCache` of another workbook.
        """
        wanted = _uint(self._pivot_cache_id(cache_id), "cache_id")
        h = self._require()
        for index in range(_read_count(LIB.fm_workbook_get_pivot_cache_count, h)):
            if _read_count(LIB.fm_workbook_get_pivot_cache_id_at, h, _uint(index, "idx")) == wanted:
                return self._pivot_cache_handle(wanted)
        raise FormulonError(
            StatusCode.NOT_FOUND,
            _diagnostic_override=(f"the workbook has no pivot cache {wanted}", f"cache_id={wanted}"),
        )

    def add_pivot_cache(self, requested_id: int = 0) -> PivotCache:
        """Create a new empty pivot cache and return its handle.

        Its id is ``requested_id`` when non-zero, else one the engine assigns.
        A new cache has no worksheet source, and a workbook containing one
        cannot be saved: call :meth:`PivotCache.set_worksheet_source` before
        saving.
        """
        h = self._require()
        return self._pivot_cache_handle(
            _read_count(LIB.fm_workbook_add_pivot_cache, h, _uint(requested_id, "requested_id"))
        )

    def _pivot_cache_handle(self, cache_id: int) -> PivotCache:
        """The handle of cache id ``cache_id``, whether or not the workbook has that cache now."""
        from ._pivot_cache import PivotCache

        handle = self._pivot_caches.get(cache_id)
        if handle is None:
            handle = PivotCache(self, cache_id)
            self._pivot_caches[cache_id] = handle
        return handle

    def _pivot_cache_id(self, cache: Union[int, PivotCache]) -> int:
        """The cache id a cache id or a :class:`PivotCache` of this workbook names."""
        from ._pivot_cache import PivotCache

        if isinstance(cache, PivotCache):
            if cache.workbook is not self:
                raise FormulonError(
                    StatusCode.INVALID_ARGUMENT,
                    _diagnostic_override=("the pivot cache belongs to another workbook", ""),
                )
            return cache.cache_id
        if isinstance(cache, bool):
            raise TypeError("formulon: cache_id must be an int or a PivotCache, got bool")
        try:
            return operator.index(cache)
        except TypeError:
            raise TypeError(f"formulon: cache_id must be an int or a PivotCache, got {type(cache).__name__}") from None

    def remove_pivot_cache(self, cache_id: Union[int, PivotCache]) -> None:
        """Remove the pivot cache ``cache_id``.

        Its handle raises :class:`NotFoundError` until a cache with the same id
        is added. Fails while a pivot table still references the cache.

        Raises:
          InvalidArgumentError: when ``cache_id`` is a :class:`PivotCache` of another workbook.
        """
        wanted = _uint(self._pivot_cache_id(cache_id), "cache_id")
        h = self._require()
        _check(LIB.fm_workbook_remove_pivot_cache(h, wanted))

    # -- Reference text ----------------------------------------------------
    def parse_reference(
        self,
        text: str,
        *,
        r1c1: bool = False,
        base_row: Optional[int] = None,
        base_col: Optional[int] = None,
    ) -> ParsedReference:
        """Parse ``text`` as ``INDIRECT`` reads a reference.

        ``base_row`` and ``base_col`` anchor R1C1 relative axes and are given together or not at all.

        Raises:
          FormulaSyntaxError: with ``StatusCode.PARSER_INVALID_REFERENCE`` when ``text`` is not a reference.
        """
        if (base_row is None) != (base_col is None):
            raise ValueError("formulon: base_row and base_col must be given together")
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.REFERENCE)
            _check(
                LIB.fm_workbook_parse_reference(
                    h,
                    scope.cstr(text),
                    _sint(1 if r1c1 else 0, "r1c1"),
                    _sint(0 if base_row is None else 1, "has_base"),
                    _uint(base_row or 0, "base_row"),
                    _uint(base_col or 0, "base_col"),
                    ptr,
                )
            )
            d = S.REFERENCE.unpack(LIB, ptr)
            return ParsedReference(
                sheet_name=LIB.read_cstr(d["sheet_name"]) if d["has_sheet"] else None,
                rect=_cell_rect_from(d["rect"]),
            )

    def format_reference(
        self,
        rect: CellRect,
        *,
        sheet_name: Optional[str] = None,
        absolute: bool = False,
    ) -> str:
        """Format ``rect`` as A1 reference text (``B:D``, ``2:11``, ``'My Sheet'!A1``).

        Raises:
          FormulonError: with ``StatusCode.INVALID_ARGUMENT`` for a coordinate outside the grid or a reversed rectangle.
        """
        h = self._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            _check(
                LIB.fm_workbook_format_reference(
                    h,
                    scope.cstr(sheet_name) if sheet_name else 0,
                    rp,
                    _sint(1 if absolute else 0, "absolute"),
                    out,
                )
            )
            return LIB.read_cstr(LIB.read_u32(out))

    # -- External links ----------------------------------------------------
    def list_external_links(self) -> List[ExternalLink]:
        """Every external-link record of the workbook, in order."""
        return [self._external_link_at(i) for i in range(self._external_link_count())]

    def _external_link_count(self) -> int:
        """Return the number of external-link records on the workbook."""
        h = self._require()
        return _read_count(
            LIB.fm_workbook_get_external_link_count,
            h,
        )

    def _external_link_at(self, index: int) -> ExternalLink:
        """Read the ``index``-th external-link record."""
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.EXTERNAL_LINK_RECORD)
            _check(LIB.fm_workbook_get_external_link(h, _uint(index, "index"), ptr))
            d = S.EXTERNAL_LINK_RECORD.unpack(LIB, ptr)
            return ExternalLink(
                index=d["index"],
                rel_id=LIB.read_cstr(d["rel_id"]),
                part_path=LIB.read_cstr(d["part_path"]),
                target=LIB.read_cstr(d["target"]),
                target_external=bool(d["target_external"]),
                kind=ExternalLinkKind(d["kind"]),
            )


def _lock_workbook_method(method):
    """Serialize one complete public operation on the shared engine instance.

    A normal method can make several ABI calls while it copies borrowed
    pointers into Python objects.  Locking only those individual calls leaves
    the handle and borrowed memory exposed between calls.  Generators need a
    slightly different shape: hold the lock through one ``next`` call so the
    item is decoded before it is yielded, then release it while the consumer
    examines the item.  Closing the underlying generator is also serialized
    so its scratch buffers are released before another operation starts.
    """

    if isgeneratorfunction(method):

        @wraps(method)
        def locked_generator(*args, **kwargs):
            iterator = method(*args, **kwargs)
            try:
                while True:
                    with LIB.call_lock:
                        try:
                            item = next(iterator)
                        except StopIteration:
                            return
                    yield item
            finally:
                with LIB.call_lock:
                    close = getattr(iterator, "close", None)
                    if close is not None:
                        close()

        return locked_generator

    if getattr(method, "__name__", None) == "__del__":

        @wraps(method)
        def locked_finalizer(*args, **kwargs):
            # Keep the destructor's no-raise contract even when interpreter
            # teardown has already cleared LIB. During normal execution the
            # same RLock also covers the nested close() call.
            try:
                with LIB.call_lock:
                    return method(*args, **kwargs)
            except Exception:  # noqa: BLE001 - finalizers must not raise
                return None

        return locked_finalizer

    @wraps(method)
    def locked(*args, **kwargs):
        with LIB.call_lock:
            return method(*args, **kwargs)

    return locked


def _install_workbook_locks(cls: type) -> None:
    """Wrap every operation of ``cls`` (``Workbook`` or ``Sheet``) while preserving its descriptors."""

    for name, descriptor in list(cls.__dict__.items()):
        if isinstance(descriptor, property):
            setattr(
                cls,
                name,
                property(
                    _lock_workbook_method(descriptor.fget) if descriptor.fget is not None else None,
                    _lock_workbook_method(descriptor.fset) if descriptor.fset is not None else None,
                    _lock_workbook_method(descriptor.fdel) if descriptor.fdel is not None else None,
                    descriptor.__doc__,
                ),
            )
        elif isinstance(descriptor, classmethod):
            setattr(cls, name, classmethod(_lock_workbook_method(descriptor.__func__)))
        elif isinstance(descriptor, staticmethod):
            setattr(cls, name, staticmethod(_lock_workbook_method(descriptor.__func__)))
        elif callable(descriptor):
            setattr(cls, name, _lock_workbook_method(descriptor))


# <ooxml-namespace>
class WorkbookOoxml:
    """The OOXML-level layer of a workbook, :attr:`Workbook.ooxml`.

    The style tables a cell XF indexes (fonts, fills, borders, number
    formats, ``<cellXfs>``, ``<cellStyleXfs>``, DXFs) and the passthrough
    parts. Most code uses :meth:`Sheet.update_cell_format` and
    :meth:`Sheet.get_cell_format` instead.
    """

    __slots__ = ("_workbook",)

    def __init__(self, workbook: Workbook) -> None:
        self._workbook = workbook

    def __repr__(self) -> str:
        return f"{self._workbook!r}.ooxml"

    def list_passthroughs(self) -> List[PassthroughPart]:
        """Return the passthrough OOXML part paths."""
        h = self._workbook._require()
        n = _read_count(LIB.fm_workbook_get_passthrough_count, h)
        # Scratch hoisted out of the loop; see list_defined_names.
        with _ArgScope() as scope:
            path_ptr = scope.out_ptr()
            out: List[PassthroughPart] = []
            for i in range(n):
                status = LIB.fm_workbook_get_passthrough_at(h, _uint(i, "idx"), path_ptr)
                _check(status)
                out.append(PassthroughPart(path=LIB.read_cstr(LIB.read_u32(path_ptr))))
        return out

    def has_passthrough(self, path: str) -> bool:
        """Whether a passthrough part sits at ``path`` (no leading ``/``)."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            path_ptr = scope.cstr(path)
            present = scope.out_ptr()
            _check(LIB.fm_workbook_has_passthrough(h, path_ptr, present))
            return LIB.read_u32(present) != 0

    def get_cell_xf(self, xf_index: int) -> CellXf:
        """Return the resolved ``<xf>`` record at ``xf_index``."""
        return self._workbook._get_style_record(
            LIB.fm_workbook_get_cell_xf,
            S.CELL_XF,
            xf_index,
            "xf_index",
            self._workbook._decode_cell_xf,
        )

    def get_font(self, font_index: int) -> Font:
        """Return the font record at ``font_index``."""
        return self._workbook._get_style_record(
            LIB.fm_workbook_get_font, S.FONT_RECORD, font_index, "font_index", _decode_font
        )

    def get_fill(self, fill_index: int) -> Fill:
        """Return the fill record at ``fill_index``."""
        return self._workbook._get_style_record(
            LIB.fm_workbook_get_fill, S.FILL_RECORD, fill_index, "fill_index", _decode_fill
        )

    def get_border(self, border_index: int) -> Border:
        """Return the border record at ``border_index``."""
        return self._workbook._get_style_record(
            LIB.fm_workbook_get_border,
            S.BORDER_RECORD,
            border_index,
            "border_index",
            _decode_border_record,
        )

    def get_dxf(self, dxf_index: int) -> Dxf:
        """Read one differential format by its index (a conditional format's ``dxf_index``)."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.DXF_RECORD)
            _check(LIB.fm_workbook_get_dxf(h, _uint(dxf_index, "dxf_index"), ptr))
            d = S.DXF_RECORD.unpack(LIB, ptr)
            offsets = S.DXF_RECORD.offsets
            font = _decode_font(ptr + offsets["font"][1]) if d["font_engaged"] else None
            fill = _decode_fill(ptr + offsets["fill"][1]) if d["fill_engaged"] else None
            return Dxf(
                font=font,
                fill=fill,
                border=_decode_border_record(ptr + offsets["border"][1]) if d["border_engaged"] else None,
                number_format=_opt_cstr(d["number_format"]),
                alignment=_alignment_from_xml(LIB.read_cstr(d["alignment_xml"])),
                protection=_protection_from_xml(LIB.read_cstr(d["protection_xml"])),
            )

    def get_number_format(self, num_fmt_id: int) -> Optional[str]:
        """Return the effective format code for ``num_fmt_id`` in the workbook profile's vocabulary.

        Built-in id 14 is ``m/d/yy`` under en-US and ``yyyy/m/d`` under ja-JP. ``None`` when no format maps to it.
        """
        h = self._workbook._require()
        return _read_optional_cstr_out(
            LIB.fm_workbook_get_number_format,
            h,
            _uint(num_fmt_id, "num_fmt_id", 16),
        )

    def list_fonts(self) -> List[Font]:
        """Every font record, in index order."""
        return [self.get_font(i) for i in range(self._workbook._style_count(LIB.fm_workbook_get_font_count))]

    def list_fills(self) -> List[Fill]:
        """Every fill record, in index order."""
        return [self.get_fill(i) for i in range(self._workbook._style_count(LIB.fm_workbook_get_fill_count))]

    def list_borders(self) -> List[Border]:
        """Every border record, in index order."""
        return [self.get_border(i) for i in range(self._workbook._style_count(LIB.fm_workbook_get_border_count))]

    def list_cell_xfs(self) -> List[CellXf]:
        """Every ``<xf>`` record of ``<cellXfs>``, in index order."""
        return [self.get_cell_xf(i) for i in range(self._workbook._style_count(LIB.fm_workbook_get_cell_xf_count))]

    def list_cell_style_xfs(self) -> List[CellXf]:
        """Every ``<cellStyleXfs>`` record, in index order."""
        count = self._workbook._style_count(LIB.fm_workbook_get_cell_style_xf_count)
        return [self.get_cell_style_xf(i) for i in range(count)]

    def list_dxfs(self) -> List[Dxf]:
        """Every differential format, in index order."""
        return [self.get_dxf(i) for i in range(self._workbook._style_count(LIB.fm_workbook_get_dxf_count))]

    def add_font(self, record: Font) -> int:
        """Add (dedup) a font record; return its index."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.FONT_RECORD)
            out = scope.out_ptr()
            _pack_font(ptr, record, scope)
            _check(LIB.fm_workbook_add_font(h, ptr, out))
            return LIB.read_u32(out)

    def set_font(self, font_index: int, record: Font) -> None:
        """Overwrite an existing font slot in place.

        Every ``<xf>`` naming ``font_index`` restyles at once, so this is a
        bulk change rather than a local edit; :meth:`add_font` is the way to
        introduce a new appearance. The index must already exist.
        """
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.FONT_RECORD)
            _pack_font(ptr, record, scope)
            _check(LIB.fm_workbook_set_font(h, _uint(font_index, "font_index"), ptr))

    def add_fill(self, record: Fill) -> int:
        """Add (dedup) a fill record; return its index."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.FILL_RECORD)
            out = scope.out_ptr()
            _pack_fill(ptr, record, scope)
            _check(LIB.fm_workbook_add_fill(h, ptr, out))
            return LIB.read_u32(out)

    def add_border(self, record: Border) -> int:
        """Add (dedup) a border record; return its index."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.BORDER_RECORD)
            out = scope.out_ptr()
            _pack_border_record(ptr, record)
            _check(LIB.fm_workbook_add_border(h, ptr, out))
            return LIB.read_u32(out)

    def add_number_format(self, format_code: str) -> int:
        """Add (dedup) a number-format code in the workbook profile's vocabulary; return its resolved id.

        A built-in id whose profile code equals ``format_code`` wins (``m/d/yy`` is 14 under en-US); the ECMA-376
        spelling ``mm-dd-yy`` becomes a custom format.
        """
        h = self._workbook._require()
        with _ArgScope() as scope:
            code_ptr = scope.cstr(format_code)
            out = scope.out_ptr()
            _check(LIB.fm_workbook_add_number_format(h, code_ptr, out))
            return LIB.read_u32(out)

    def add_cell_xf(self, record: CellXfInput) -> int:
        """Add (dedup) an ``<xf>`` record; return its index.

        Out-of-range font, fill or border indices and an unregistered
        ``num_fmt_id`` raise :class:`InvalidArgumentError` rather than growing
        the style tables.
        """
        return self._workbook._add_xf_record(record, LIB.fm_workbook_add_cell_xf)

    def add_cell_style_xf(self, record: CellXfInput) -> int:
        """Add (dedup) a named-style ``<xf>``; return its ``<cellStyleXfs>`` id.

        The counterpart of :meth:`get_cell_style_xf`. Named-style xfs are
        the records a ``<cellStyle>`` points at; :meth:`add_cell_xf` writes
        the per-cell ``<cellXfs>`` pool instead.
        """
        return self._workbook._add_xf_record(record, LIB.fm_workbook_add_cell_style_xf)

    def add_dxf(self, record: Union[DxfInput, Dxf]) -> int:
        """Add (dedup) a differential format and return its index.

        Accepts a :class:`DxfInput` or the :class:`Dxf` :meth:`get_dxf` returns.
        A ``number_format`` is registered as :meth:`add_number_format` does.
        """
        if not isinstance(record, (DxfInput, Dxf)):
            raise TypeError(f"formulon: expected a DxfInput, got {type(record).__name__}")
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.DXF_RECORD)
            out = scope.out_ptr()
            S.DXF_RECORD.pack(
                LIB,
                ptr,
                {
                    "font_engaged": 1 if record.font is not None else 0,
                    "fill_engaged": 1 if record.fill is not None else 0,
                    "border_engaged": 1 if record.border is not None else 0,
                },
            )
            offsets = S.DXF_RECORD.offsets
            if record.font is not None:
                _pack_font(ptr + offsets["font"][1], record.font, scope)
            if record.fill is not None:
                _pack_fill(ptr + offsets["fill"][1], record.fill, scope)
            if record.border is not None:
                _pack_border_record(ptr + offsets["border"][1], record.border)
            S.write_str_field(LIB, ptr, S.DXF_RECORD, "number_format", record.number_format, scope)
            S.write_str_field(LIB, ptr, S.DXF_RECORD, "alignment_xml", _alignment_to_xml(record.alignment), scope)
            S.write_str_field(LIB, ptr, S.DXF_RECORD, "protection_xml", _protection_to_xml(record.protection), scope)
            _check(LIB.fm_workbook_add_dxf(h, ptr, out))
            return LIB.read_u32(out)

    def get_cell_style_xf(self, cell_style_xf_index: int) -> CellXf:
        """Return the named-style xf record at ``cell_style_xf_index``."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CELL_XF)
            _check(LIB.fm_workbook_get_cell_style_xf(h, _uint(cell_style_xf_index, "cell_style_xf_index"), ptr))
            return self._workbook._decode_cell_xf(ptr)


def _workbook_ooxml(self: Workbook) -> WorkbookOoxml:
    """The OOXML-level layer of the workbook: style tables, named styles and passthrough parts. Always the same object."""
    layer = getattr(self, "_ooxml", None)
    if layer is None:
        layer = WorkbookOoxml(self)
        self._ooxml = layer
    return layer


Workbook.ooxml = property(_workbook_ooxml)  # type: ignore[attr-defined]
_install_workbook_locks(WorkbookOoxml)
# </ooxml-namespace>


_install_workbook_locks(Workbook)


# -- Images (process-wide) ---------------------------------------------------
@_lock_workbook_method
def get_image_info(data: Union[bytes, bytearray, memoryview]) -> ImageInfo:
    """Sniff the format and pixel size of PNG, JPEG, GIF or BMP bytes from the header alone.

    Raises:
      FormulonError: with ``StatusCode.IO_IMAGE_UNSUPPORTED`` when the bytes are
        not a recognised image, the header is truncated, or a dimension is zero.
    """
    raw = bytes(data)
    with _ArgScope() as scope:
        buf = scope.buffer(raw)
        info = scope.struct(S.IMAGE_INFO)
        _check(LIB.fm_get_image_info(buf, _uint(len(raw), "len"), info))
        return _read_image_info(info)


# -- Function catalog and locale (process-wide) -----------------------------
@_lock_workbook_method
def get_function_metadata(name: str) -> Optional[FunctionMetadata]:
    """Return metadata for ``name`` or ``None`` when unknown.

    The metadata is locale-invariant.

    Args:
      name: canonical function name, matched case-insensitively.

    Returns:
      The metadata, or ``None`` when ``name`` matches no registered
      function.
    """
    with _ArgScope() as scope:
        name_ptr = scope.cstr(name)
        present = scope.out_ptr()
        ptr = scope.struct(S.FUNCTION_METADATA)
        _check(LIB.fm_get_function_metadata(name_ptr, present, ptr))
        if not LIB.read_u32(present):
            return None
        d = S.FUNCTION_METADATA.unpack(LIB, ptr)
        sig = LIB.read_cstr(d["signature_template"]) if d["signature_template"] else None
        desc = LIB.read_cstr(d["description"]) if d["description"] else None
        # 0xFFFFFFFF is the unbounded / unknown-arity sentinel; expose it
        # as None so callers do not treat it as a concrete upper bound.
        raw_max = d["max_arity"]
        max_arity = None if raw_max == 0xFFFFFFFF else raw_max
        return FunctionMetadata(
            name=LIB.read_cstr(d["canonical_name"]),
            min_arity=d["min_arity"],
            max_arity=max_arity,
            availability=_enum(FunctionAvailability, d["availability"]),
            signature=sig,
            description=desc,
        )


def _map_profile_text(fn_name: str, text: str, profile_id: str) -> str:
    """Call a ``(text, profile_id) -> text`` catalog entry point."""
    with _ArgScope() as scope:
        text_ptr = scope.cstr(text)
        pid_ptr = scope.cstr(profile_id)
        return _read_cstr_out(getattr(LIB, fn_name), text_ptr, pid_ptr)


@_lock_workbook_method
def list_function_names() -> List[str]:
    """Return the canonical names of every registered Formulon function, sorted."""
    return [
        _read_cstr_out(LIB.fm_get_function_name_at, _uint(index, "idx"))
        for index in range(_read_count(LIB.fm_get_function_count))
    ]


@_lock_workbook_method
def localize_function_name(canonical_name: str, profile_id: ProfileId) -> str:
    """Return the display name of ``canonical_name`` under ``profile_id``."""
    return _map_profile_text("fm_localize_function_name", canonical_name, profile_id)


@_lock_workbook_method
def canonicalize_function_name(localized_name: str, profile_id: ProfileId) -> str:
    """Return the canonical English name for ``localized_name`` under ``profile_id``."""
    return _map_profile_text("fm_canonicalize_function_name", localized_name, profile_id)


@_lock_workbook_method
def localize_formula(formula: str, profile_id: ProfileId) -> str:
    """Rewrite stored (en-invariant) formula text into ``profile_id``'s spelling.

    Function names, list / array separators, decimal points, booleans
    and error names take the locale form; strings, quoted sheet names,
    external qualifiers and structured-reference brackets are kept
    verbatim. A pure text rewrite: the grammar is not validated.
    """
    return _map_profile_text("fm_localize_formula", formula, profile_id)


@_lock_workbook_method
def canonicalize_formula(formula: str, profile_id: ProfileId) -> str:
    """Rewrite ``profile_id``-spelled formula text into the en-invariant form."""
    return _map_profile_text("fm_canonicalize_formula", formula, profile_id)


def _read_locale_error_names(scope: _ArgScope, pid_ptr: int) -> List[ErrorName]:
    canonical_out = scope.out_ptr()
    localized_out = scope.out_ptr()
    measured_out = scope.out_ptr()
    errors = []
    for index in range(_read_count(LIB.fm_get_locale_error_name_count)):
        _check(
            LIB.fm_get_locale_error_name_at(
                pid_ptr, _uint(index, "error_index"), canonical_out, localized_out, measured_out
            )
        )
        errors.append(
            ErrorName(
                canonical=LIB.read_cstr(LIB.read_u32(canonical_out)),
                localized=LIB.read_cstr(LIB.read_u32(localized_out)),
                measured=LIB.read_u32(measured_out) != 0,
            )
        )
    return errors


@_lock_workbook_method
def list_locale_error_names(profile_id: ProfileId) -> List[ErrorName]:
    """Return the error names of ``profile_id`` in error-code order.

    Raises:
      FormulonError: when ``profile_id`` is not a documented profile.
    """
    with _ArgScope() as scope:
        return _read_locale_error_names(scope, scope.cstr(profile_id))


@_lock_workbook_method
def get_locale_facts(profile_id: ProfileId) -> LocaleFacts:
    """Return the locale facts (separators, names, currency, error names) of ``profile_id``.

    Raises:
      FormulonError: when ``profile_id`` is not a documented profile.
    """
    with _ArgScope() as scope:
        pid_ptr = scope.cstr(profile_id)
        ptr = scope.struct(S.LOCALE_FACTS)
        _check(LIB.fm_get_locale_facts(pid_ptr, ptr))
        d = S.LOCALE_FACTS.unpack(LIB, ptr)
        errors = _read_locale_error_names(scope, pid_ptr)
        return LocaleFacts(
            decimal_separator=LIB.read_cstr(d["decimal_separator"]),
            group_separator=LIB.read_cstr(d["group_separator"]),
            list_separator=LIB.read_cstr(d["list_separator"]),
            array_column_separator=LIB.read_cstr(d["array_column_separator"]),
            array_row_separator=LIB.read_cstr(d["array_row_separator"]),
            true_name=LIB.read_cstr(d["true_name"]),
            false_name=LIB.read_cstr(d["false_name"]),
            date_order=LocaleDateOrder(d["date_order"]),
            currency_symbol=LIB.read_cstr(d["currency_symbol"]),
            currency_suffix=d["currency_suffix"] != 0,
            currency_space=d["currency_space"] != 0,
            currency_default_decimals=d["currency_default_decimals"],
            measured=d["measured"] != 0,
            collation_measured=d["collation_measured"] != 0,
            error_names=errors,
        )


@_lock_workbook_method
def compare_text(a: str, b: str, strength: CollationStrength, profile_id: ProfileId) -> int:
    """Compare two texts under the collation of ``profile_id``.

    Returns -1, 0 or 1. This is the text fold only, not ``=`` on two
    cells. ``strength`` selects which fold applies.

    Raises:
      FormulonError: on an unknown ``profile_id`` or an out-of-range ``strength``.
    """
    with _ArgScope() as scope:
        a_ptr = scope.cstr(a)
        b_ptr = scope.cstr(b)
        pid_ptr = scope.cstr(profile_id)
        out = scope.out_ptr()
        _check(LIB.fm_compare_text(a_ptr, b_ptr, _sint(strength, "strength"), pid_ptr, out))
        return struct.unpack("<i", LIB.read_bytes(out, 4))[0]


@_lock_workbook_method
def get_collation_classes(
    values: Sequence[Value], strength: CollationStrength, profile_id: ProfileId
) -> CollationClasses:
    """Classify ``values`` into collation-ordered equivalence classes.

    Two values share an id only when they have the same kind and compare
    equal under the ``strength`` text fold (``-0`` equals ``0``; blanks
    are always equal). Ids are dense from 0 in ascending SORT order:
    numbers < text < booleans < errors < blanks. ``strength`` selects the
    text fold only; type boundaries never depend on it.

    Ids are valid only within one call. Usage:

    * Dedup or checklist order: take the first position of each id and
      order by id.
    * Sort: stable-sort by id; reverse for descending, keeping blanks last.
    * Re-applying a saved selection: classify the saved values plus the
      column values in one call and select the matching ids.

    Raises:
      FormulonError: on an unknown ``profile_id``, an out-of-range
        ``strength``, a non-finite number or an out-of-range error code.
    """
    values = [_check_value(v) for v in values]
    n = len(values)
    with _ArgScope() as scope:
        pid_ptr = scope.cstr(profile_id)
        count_out = scope.alloc(4)
        base = scope.value_array(n)
        for i, v in enumerate(values):
            _write_value_input(v, base + i * fm_value_t_size, scope)
        ids = 0
        if n:
            ids = scope.alloc(4 * n)
        _check(
            LIB.fm_get_collation_classes(base, _uint(n, "count"), _sint(strength, "strength"), pid_ptr, ids, count_out)
        )
        class_ids = list(struct.unpack(f"<{n}I", LIB.read_bytes(ids, 4 * n))) if n else []
        return CollationClasses(class_ids=class_ids, class_count=LIB.read_u32(count_out))


# Imported last: these modules import this one, and the annotations above name their classes, which
# `typing.get_type_hints` resolves from this module's globals.
from ._inputs import cell_rect, civil_datetime, color_input, optional_color_input  # noqa: E402
from ._pivot_cache import PivotCache  # noqa: E402
from ._sheet import Sheet  # noqa: E402
from ._table import Table  # noqa: E402
