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

import builtins
import dataclasses
import datetime
import enum
import json
import math
import numbers
import operator
import os
import re
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

from . import _handle_tables
from . import _structs as S
from ._c import (
    LIB,
    FormulaSyntaxError,
    FormulonError,
    InvalidFileError,
    ValueKind,
    _argument_message,
    _NoCopy,
    _past_count,
    _public_call,
    _require_bool,
    _require_number,
    _sint,
    _uint,
    fm_value_t_size,
)
from ._status import StatusCode

__all__ = [
    "Alignment",
    "AlignmentPatch",
    "AnchorEditAs",
    "AnchorKind",
    "ArrayFormula",
    "AutoFilter",
    "AutoFilterColumn",
    "AutoFilterCondition",
    "AutoFilterDateGroup",
    "AutoFilterEval",
    "AutoFilterInput",
    "AutoFilterSortCondition",
    "AutoFilterSortState",
    "Border",
    "BorderPatch",
    "BorderSide",
    "BorderSidePatch",
    "BorderStyle",
    "CalcMode",
    "CompatibilityVersion",
    "canonicalize_formula",
    "canonicalize_function_name",
    "CellAddress",
    "CellFormat",
    "CellNode",
    "CellProtection",
    "CellProtectionPatch",
    "CellRect",
    "CellShift",
    "CellsInRange",
    "CellStyle",
    "CellXf",
    "CellXfInput",
    "CfAboveAverageRuleInput",
    "CfCell",
    "CfCellIsRuleInput",
    "CfColorScale",
    "CfColorScaleInput",
    "CfColorScaleRuleInput",
    "CfDataBar",
    "CfDataBarAxisPosition",
    "CfDataBarDirection",
    "CfDataBarInput",
    "CfDataBarRuleInput",
    "CfExpressionRuleInput",
    "CfIcon",
    "CfIconSet",
    "CfIconSetInput",
    "CfIconSetRuleInput",
    "CfMatch",
    "CfMatchBar",
    "CfMatchIcon",
    "CfMatchKind",
    "CfPlainRuleInput",
    "CfRuleType",
    "CfRuleId",
    "CfRuleRef",
    "CfTextRuleInput",
    "CfTimePeriodRuleInput",
    "CfTop10RuleInput",
    "CfValueObjectType",
    "CfValueObject",
    "ChartAxis",
    "ChartAxisCrosses",
    "ChartAxisKey",
    "ChartAxisKind",
    "ChartAxisPatch",
    "ChartAxisPosition",
    "ChartAxisRef",
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
    "ChartFilter",
    "ChartGrouping",
    "ChartInput",
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
    "ChartSourceInput",
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
    "ColumnBreak",
    "ColumnLayout",
    "Comment",
    "compare_text",
    "ConditionalFormat",
    "ConditionalFormatInput",
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
    "DrawingAnchorInput",
    "DrawingObjectKind",
    "DrawingProperties",
    "Dxf",
    "DxfInput",
    "DynamicFilterType",
    "EffectiveBorders",
    "EffectiveStyle",
    "EffectiveStyleSource",
    "enum_name",
    "ErrorCode",
    "LocaleErrorName",
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
    "IconSetName",
    "ImageFormat",
    "ImageInfo",
    "InvalidFileError",
    "IterativeSettings",
    "list_function_names",
    "list_profile_ids",
    "LocaleDateOrder",
    "LocaleFacts",
    "localize_formula",
    "localize_function_name",
    "LogLevel",
    "PaginationMargins",
    "Mention",
    "MentionInput",
    "NativeValue",
    "PageOrder",
    "PaperSize",
    "PhoneticAlignment",
    "PhoneticType",
    "DrawingPlacementInput",
    "PivotDateGroup",
    "PivotDateGroupInput",
    "PivotNumericGroup",
    "PivotNumericGroupInput",
    "PivotStyleOptions",
    "PivotWorksheetSourceInput",
    "PrintTitlesInput",
    "ProtectedRangeInput",
    "ProtectedRangeRef",
    "PageBreakInput",
    "RowBreak",
    "PaginationPage",
    "PageMargins",
    "PageOrientation",
    "PageSetup",
    "Pagination",
    "PaginationPaper",
    "ParsedReference",
    "NumberFormat",
    "Person",
    "PersonInput",
    "PhoneticProperties",
    "PhoneticRun",
    "PictureProperties",
    "PIVOT_VALUES_FIELD",
    "PivotAggregation",
    "PivotAxis",
    "PivotCacheField",
    "PivotCacheRef",
    "PivotCacheOptions",
    "PivotCacheValue",
    "PivotCalculatedItem",
    "PivotCalendar",
    "PivotCell",
    "PivotCellKind",
    "PivotDataField",
    "PivotDataFieldInput",
    "PivotDatePeriod",
    "PivotField",
    "PivotFieldRef",
    "PivotFilter",
    "PivotFilterInput",
    "PivotFilterKind",
    "PivotGrandTotals",
    "PivotInput",
    "PivotItem",
    "PivotLayout",
    "PivotOptions",
    "PivotProperties",
    "PivotRef",
    "PivotReportLayout",
    "PivotShowAsBase",
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
    "ShapeProperties",
    "SheetComment",
    "SheetFormatDefaults",
    "SheetProtection",
    "SheetProtectionInput",
    "SheetRect",
    "SheetRef",
    "SheetSelection",
    "SheetSelectionInput",
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
    "SparklineGroupId",
    "SparklineGroupInput",
    "SparklineGroupRef",
    "SparklineInput",
    "SparklineType",
    "Spill",
    "SpillCause",
    "TableColumn",
    "TableColumnPatch",
    "TableInput",
    "TableProperties",
    "TableRef",
    "TableStyle",
    "TableStyleDefaults",
    "TableStyleElement",
    "TableStyleElementType",
    "TableStyleOptions",
    "TableTotalsRowFunction",
    "ThemeColors",
    "ThemeColorsInput",
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
    "ValidationId",
    "ValidationOutcome",
    "ValidationRef",
    "ValidationType",
    "Value",
    "VerticalAlign",
    "WidthModel",
    "WidthPlatform",
    "Workbook",
    "WorkbookFormat",
    "WorkbookLocation",
    "WorkbookLocationInput",
    "WorkbookProtection",
    "WorkbookProtectionInput",
    "XfCell",
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

#: The opaque id of a data-validation rule, stable while the workbook is open and assigned again on load.
ValidationId = str

#: A data-validation rule: its position in :meth:`Sheet.list_validations` or its :data:`ValidationId`.
ValidationRef = Union[int, ValidationId]

#: The opaque id of a sparkline group, stable while the workbook is open and assigned again on load.
SparklineGroupId = str

#: A sparkline group: its position in :meth:`Sheet.list_sparkline_groups` or its :data:`SparklineGroupId`.
SparklineGroupRef = Union[int, SparklineGroupId]

#: A pivot cache: its ``cache_id`` or its :class:`PivotCache` handle.
PivotCacheRef = Union[int, "PivotCache"]

#: A pivot field: its index (the cache field index) or its source name, matched case-insensitively.
PivotFieldRef = Union[int, str]

#: A table: its index on the sheet, its name (matched case-insensitively) or its :class:`Table` handle.
TableRef = Union[int, str, "Table"]

#: A pivot table: its index on the sheet, its name (the first of several) or its :class:`Pivot` handle.
PivotRef = Union[int, str, "Pivot"]

#: The id of a conditional-format rule: a brace-wrapped GUID, stable while the workbook is open and
#: returned by :meth:`Sheet.add_cf_rule`.
CfRuleId = str

#: A conditional-format rule: its flattened position or its :data:`CfRuleId`; a ``str`` is always an id.
CfRuleRef = Union[int, CfRuleId]

#: A protected range: its position or its name; a ``str`` is always a name.
ProtectedRangeRef = Union[int, str]

#: Excel profile id accepted by ``Workbook.set_profile_id``.
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
    """

    #: The written workbook.
    bytes: bytes
    #: The number of formula cells emitted as cached literals. Always 0 for XLSX.
    downgraded_formula_count: int
    #: The number of sheet features not lowered to records. Always 0 for XLSX.
    deferred_feature_count: int
    #: The number of passthrough parts dropped for colliding with a generated path, or for two entries claiming one
    #: path. Applies to both containers.
    dropped_part_count: int
    #: The number of relationships dropped because their target part is gone, at package, workbook or (XLSB) sheet
    #: scope. Rises with ``dropped_part_count`` for one dropped part, so do not add them.
    dropped_relationship_count: int
    #: The number of tables emitted under a writer-assigned id. Always 0 for XLSB.
    renumbered_part_count: int


@dataclass(frozen=True)
class ReadDiagnostics:
    """Loss and recovery counters captured while loading a workbook.

    Same contract as :class:`SaveDiagnostics`: one meaning per field
    across both containers, partial coverage by design.
    """

    #: The number of formula cells whose stored formula could not be decoded. XLSB only.
    undecoded_formula_count: int
    #: The number of defined names skipped because their stored formula could not be decoded. XLSB only.
    undecoded_defined_name_count: int
    #: The number of package parts whose content type could not be resolved, so they were not decoded. XLSB only; the
    #: XLSX reader keeps unmodelled parts as passthrough.
    undecoded_part_count: int
    #: The number of entries dropped on load: one per merge, hyperlink or data validation and one per
    #: conditional-format block with an unusable reference (XLSX only), and one per table column whose header dxf or
    #: header cell style was dropped, which Excel refuses to open (XLSX and XLSB), and one per XLSB table part dropped
    #: whole because its header slots could not be cleared.
    skipped_feature_count: int
    #: The number of workbook parts whose content type was unrecognised; at most 1. XLSX only.
    unknown_content_type_count: int


# C field names a record publishes under another name; the generated table lists the others.
_C_FIELD_RENAMES = _handle_tables.C_FIELD_NAMES
_PATH_TOKEN = re.compile(r"[a-z_][a-z0-9_]*(?:\.[a-z_][a-z0-9_]*|\[\d+\])*")
_IDENTIFIER = re.compile(r"[a-z_][a-z0-9_]*")


def _relabel_engine_error(error: FormulonError) -> FormulonError:
    """Label an engine error with the public call: ``<Class.method> <message>``.

    When the message's first token is a C path naming one of the call's arguments (a parameter, ``sheet_index``
    naming ``sheet``, a field of the call's one record argument, or a renamed C field), that token becomes the
    public path; a message that leads with any other word keeps its wording.
    """
    label, frame = _public_call()
    if frame is None:
        return error
    code = frame.f_code
    first = 1 if code.co_varnames[:1] in (("self",), ("cls",)) else 0
    names = code.co_varnames[first : code.co_argcount + code.co_kwonlyargcount]
    path, _, rest = error.message.partition(" ")
    public = _public_path(path, names, frame.f_locals) if rest and _PATH_TOKEN.fullmatch(path) else None
    error.message = f"{label} {public} {rest}" if public is not None else f"{label} {error.message}"
    error.args = (error.message,)
    return error


def _public_path(path: str, names: Sequence[str], frame_locals: Dict[str, Any]) -> Optional[str]:
    """The public spelling of C path ``path`` rooted at the argument it names, or ``None`` when it names none."""
    renamed = _IDENTIFIER.sub(lambda m: _C_FIELD_RENAMES.get(m.group(0), m.group(0)), path)
    c_root = _IDENTIFIER.match(path).group(0)  # type: ignore[union-attr]
    root = _IDENTIFIER.match(renamed).group(0)  # type: ignore[union-attr]
    tail = renamed[len(root) :]
    stem = root[: -len("_index")] if root.endswith("_index") else root
    for candidate in (root, stem):
        if candidate in names:
            return candidate + tail
    records = [
        n for n in names if dataclasses.is_dataclass(frame_locals.get(n)) and not isinstance(frame_locals.get(n), type)
    ]
    if len(records) == 1:
        fields = {f.name for f in dataclasses.fields(frame_locals[records[0]])}
        if root in fields or c_root in _C_FIELD_RENAMES:
            return f"{records[0]}.{renamed}"
    return renamed if c_root in _C_FIELD_RENAMES else None


def _check(status: int) -> None:
    """Raise :class:`FormulonError` (or its subclass for ``status``) if ``status`` is non-zero."""
    if status != 0:
        raise _relabel_engine_error(FormulonError._from_engine(status))


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

    #: The error, an ``ErrorCode`` value.
    code: ErrorCode

    def __post_init__(self) -> None:
        if isinstance(self.code, bool):
            raise TypeError("ExcelError.code must be an ErrorCode, got bool")
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

    Exactly one payload field is populated, according to ``kind``. The
    constructor enforces the invariant: ``kind`` is coerced to
    :class:`ValueKind`, a missing or extra payload raises ``ValueError`` and a
    payload of the wrong type raises ``TypeError``.
    """

    #: The ``ValueKind`` of the value; it selects which payload field is populated.
    kind: ValueKind
    #: The numeric payload; set only for the number kind.
    number: Optional[float] = None
    #: The logical payload; set only for the bool kind.
    boolean: Optional[bool] = None
    #: The text payload; set only for the text kind.
    text: Optional[str] = None
    #: The error payload, an ``ErrorCode``; set only for the error kind.
    error_code: Optional[ErrorCode] = None
    #: The rows of an array result, each a tuple of scalar values of equal length; set only for the array kind.
    array: Optional[Tuple[Tuple["Value", ...], ...]] = None

    def __post_init__(self) -> None:
        kind = self.kind
        if not isinstance(kind, ValueKind):
            kind = ValueKind(_sint(kind, "kind"))
            object.__setattr__(self, "kind", kind)
        required = _VALUE_PAYLOAD_FIELD.get(kind)
        for name in _VALUE_PAYLOAD_FIELD.values():
            if name != required and getattr(self, name) is not None:
                raise ValueError(f"Value(kind={kind.name}) takes no {name}")
        if required is None:
            return
        payload = getattr(self, required)
        if payload is None:
            raise ValueError(f"Value(kind={kind.name}) requires {required}")
        if kind == ValueKind.NUMBER:
            if isinstance(payload, bool) or not isinstance(payload, numbers.Real):
                raise TypeError(f"Value.number must be a real number, got {type(payload).__name__}")
            object.__setattr__(self, "number", float(payload))
        elif kind == ValueKind.BOOL:
            if not isinstance(payload, bool):
                raise TypeError(f"Value.boolean must be a bool, got {type(payload).__name__}")
        elif kind == ValueKind.TEXT:
            if not isinstance(payload, str):
                raise TypeError(f"Value.text must be a str, got {type(payload).__name__}")
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

    Raises:
      TypeError: when ``value`` is not a :class:`Value`.
    """
    kind = value_input(value, "value").kind
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
        raise TypeError(f"Value.array must be a sequence of rows, got {type(rows).__name__}")
    grid = tuple(tuple(row) if isinstance(row, (list, tuple)) else row for row in rows)
    if not grid or not all(isinstance(row, tuple) and row for row in grid):
        raise ValueError("Value.array must hold at least one non-empty row")
    if any(len(row) != len(grid[0]) for row in grid):
        raise ValueError("Value.array rows must all have the same length")
    for row in grid:
        for cell in row:
            if not isinstance(cell, Value) or cell.kind is ValueKind.ARRAY:
                raise TypeError(f"Value.array cells must be scalar Values, got {cell!r}")
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
class XfCell:
    """One stored cell and its own non-default cell XF, as :meth:`SheetOoxml.list_cell_xf_indices` returns it."""

    #: The 0-based row.
    row: int
    #: The 0-based column.
    col: int
    #: The cell's own cell-XF index into ``WorkbookOoxml.list_cell_xfs()``; never 0.
    xf_index: int


@dataclass(frozen=True)
class RangeCell:
    """One written coordinate, as :meth:`Sheet.list_cells` returns it."""

    #: The 0-based row.
    row: int
    #: The 0-based column.
    col: int
    #: The formula text as the sheet formula getter spells it; ``None`` for a constant.
    formula: Optional[str]
    #: The computed value as native data.
    value: NativeValue
    #: The ``ValueKind`` of ``value``.
    kind: ValueKind


@dataclass(frozen=True)
class DefinedName:
    """One defined name, as ``Workbook.list_defined_names()`` and ``Workbook.get_defined_name()`` return it."""

    #: The name as stored.
    name: str
    #: The formula text, spelled as the sheet formula getter spells formulas.
    formula: str
    #: ``None`` for workbook scope; otherwise the 0-based index of the sheet the name is local to.
    scope: Optional[int]
    #: The note Excel shows in the Name Manager; ``None`` when none.
    comment: Optional[str] = None
    #: ``True`` when the name is hidden from the Name Manager.
    hidden: bool = False


@dataclass
class DefinedNameInput:
    """A defined name :meth:`Workbook.set_defined_name` writes."""

    #: The name to define.
    name: str
    #: The formula text; a leading ``=`` is optional.
    formula: str
    #: The sheet the name is local to; omitted or ``None`` makes it workbook scope.
    scope: Optional[SheetRef] = None
    #: The Name Manager note; omitted or ``None`` means no comment.
    comment: Optional[str] = None
    #: ``True`` hides the name from the Name Manager; omitted means visible.
    hidden: bool = False


def _defined_name_from(d: Dict[str, Any]) -> DefinedName:
    """Build a :class:`DefinedName` from a decoded ``fm_defined_name_t``."""
    local_sheet_id = d["local_sheet_id"]
    return DefinedName(
        name=LIB.read_cstr(d["name"]),
        formula=LIB.read_cstr(d["formula"]),
        scope=None if local_sheet_id < 0 else local_sheet_id,
        comment=LIB.read_cstr(d["comment"]) or None,
        hidden=bool(d["hidden"]),
    )


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
    """One column of :attr:`TableProperties.columns`."""

    #: The column's name, the text of its header cell.
    name: str
    #: The function the column's totals-row cell applies; the none member when the column has no totals-row function.
    totals_row_function: TableTotalsRowFunction = TableTotalsRowFunction.NONE
    #: The column's totals-row label text; ``None`` when the column has none.
    totals_row_label: Optional[str] = None
    #: The formula of the ``custom`` totals-row function, leading ``=`` included; ``None`` for any other function.
    totals_row_formula: Optional[str] = None
    #: The column's calculated-column formula as it reads in the first data row, leading ``=`` included
    #: (``=[@Qty]*10``); ``None`` when the column has none.
    calculated_formula: Optional[str] = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "totals_row_function", _enum(TableTotalsRowFunction, self.totals_row_function))


class TableColumnPatch(TypedDict, total=False):
    """One element of :meth:`Table.update`'s ``columns``, a partial :class:`TableColumn`.

    An omitted key keeps the setting and ``None`` clears it (``totals_row_function=None`` is the none member). A
    :class:`TableColumn` read back is accepted in its place; its ``name`` is ignored. The result is checked as a
    whole: the custom function needs a non-empty ``totals_row_formula``, any other function drops the formula, and an
    empty string for a label or formula raises ``InvalidArgumentError`` (clear with ``None``).
    """

    #: The function the column's totals-row cell applies; ``None`` is the none member.
    totals_row_function: Optional[TableTotalsRowFunction]
    #: The totals-row label text.
    totals_row_label: Optional[str]
    #: The formula of the ``custom`` totals-row function, with or without a leading ``=``.
    totals_row_formula: Optional[str]
    #: The column's calculated-column formula, with or without a leading ``=``, as entered in the first data row
    #: (``[@Qty]*10`` binds to this table). The mark is the state: a formula that differs from the column's mark (after
    #: binding and normalizing) marks the column and fills every data cell, constants included, with it moved to its
    #: row; the same formula changes nothing and refills nothing; ``None`` removes the mark and leaves the cells; a
    #: table with no data row only takes the mark.
    calculated_formula: Optional[str]


@dataclass(frozen=True)
class TableProperties:
    """One table, as :meth:`Table.get_properties` returns it.

    The style is read through :meth:`Table.get_style_options`.
    """

    #: The table's one name: the name formulas and ``Workbook.table`` use, also written as the file's name attribute.
    name: str
    #: The table's sheet and range, header and totals rows included.
    range: SheetRect
    #: ``True`` when the table has a header row.
    header_row: bool = True
    #: ``True`` when the table has a totals row.
    totals_row: bool = False
    #: The table's columns, in sheet order.
    columns: List[TableColumn] = field(default_factory=list)


@dataclass(frozen=True)
class TableStyleOptions:
    """The style element of a table (``<tableStyleInfo>``), as :meth:`Table.get_style_options` returns it.

    :meth:`Table.update_style_options` takes these fields as keyword arguments.
    """

    #: The name of the applied table style; ``None`` for an element without a named style. It may name a custom style.
    style_name: Optional[str] = None
    #: Whether the first column is formatted specially.
    show_first_column: bool = False
    #: Whether the last column is formatted specially.
    show_last_column: bool = False
    #: Whether odd and even rows are banded.
    show_row_stripes: bool = False
    #: Whether odd and even columns are banded.
    show_column_stripes: bool = False


@dataclass
class TableInput:
    """A table :meth:`Sheet.add_table` creates."""

    #: The cells the table covers, header row included, as a ``CellRect`` or an unqualified A1 range on the receiving
    #: sheet.
    range: Union[CellRect, str]
    #: The new table's name; it must not be one another table already uses, nor be or start with a reference in the
    #: profile locale's R1C1 letters (de ``Z1S1``).
    name: str
    #: One name per column of the range; each non-empty header cell must already hold its name and an empty one receives
    #: it. When omitted the names come from the header cells, an empty one named as Excel does (``Column1``, ...). Blank
    #: and repeated names are numbered as :meth:`Sheet.add_table` describes.
    columns: Optional[Sequence[str]] = None
    #: The built-in or custom table style to apply; omitted or ``None`` is the workbook's table-style default
    #: (:meth:`Workbook.get_table_style_defaults`), else ``TableStyleMedium2``; an empty string is invalid. A
    #: custom style that does not apply to tables, and a built-in slicer style (``SlicerStyleLight1``, ...), raise
    #: :class:`InvalidArgumentError`; when the default is used and does not apply to tables, the call fails with
    #: ``StatusCode.PRECONDITION_FAILED``.
    style_name: Optional[str] = None
    #: ``False`` creates the table without a style element, which needs ``style_name`` omitted; default ``True``.
    styled: bool = True
    #: Whether the first column is emphasized; omitted takes the creation default, which is off.
    show_first_column: Optional[bool] = None
    #: Whether the last column is emphasized; omitted takes the creation default, which is off.
    show_last_column: Optional[bool] = None
    #: Whether odd rows are banded; omitted takes the creation default, which is on for a styled table.
    show_row_stripes: Optional[bool] = None
    #: Whether odd columns are banded; omitted takes the creation default, which is off.
    show_column_stripes: Optional[bool] = None
    #: ``True`` when the table has a header row; default ``True``.
    header_row: bool = True
    #: ``True`` when the table has a totals row; default ``False``.
    totals_row: bool = False


@dataclass(frozen=True)
class TableStyleElement:
    """One region of a custom table style."""

    #: The table region this element formats (a table style element type).
    type: TableStyleElementType
    #: The stripe band size in rows or columns, at least 1; 1 for a region that is not a stripe.
    size: int = 1
    #: Index of the differential format applied to the region; ``None`` for a region listed without formatting.
    dxf_index: Optional[int] = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "type", _enum(TableStyleElementType, self.type))


@dataclass(frozen=True)
class TableStyle:
    """A custom table / PivotTable style from ``<tableStyles>``."""

    #: The style name.
    name: str
    #: Whether PivotTables may use the style; a PivotTable given a style that excludes them is refused at entry.
    pivot: bool = True
    #: Whether tables may use the style; a table given a style that excludes them is refused at entry.
    table: bool = True
    #: The formatted regions of the style.
    elements: List[TableStyleElement] = field(default_factory=list)


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
    """Which column-width figure a geometry call uses.

    ``DISPLAY`` is the screen figure of the workbook profile's host: Windows ``Range.Width`` at 96 DPI, or Mac Excel's
    whole-point layout. ``PRINT`` is the DPI-stable Windows figure pagination breaks on, on every host. Row heights are
    the same in both modes.
    """

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
    """Result of :func:`get_collation_classes`."""

    #: One class id per input value, dense from 0 in ascending sort order.
    class_ids: List[int]
    #: The number of distinct classes.
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
        raise TypeError(f"enum_type must be an enum class, got {enum_type!r}")
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
    """Anchor element type of a drawing object.

    ``DEFAULT`` is input only and asks for the operation's default: one-cell for a picture, two-cell for a
    chart, shape or slicer. A read never returns it.
    """

    DEFAULT = 0
    ONE_CELL = 1
    TWO_CELL = 2
    ABSOLUTE = 3


class AnchorEditAs(IntEnum):
    """How a drawing object follows row and column edits.

    ``DEFAULT`` is input only and asks for the anchor's own kind; a read reports the stored value.
    """

    DEFAULT = 0
    ONE_CELL = 1
    TWO_CELL = 2
    ABSOLUTE = 3


@dataclass(frozen=True)
class DrawingAnchor:
    """The placement of a drawing object as read; sizes and offsets are in EMU.

    ``from_*`` is the top-left marker; ``to_*`` is the bottom-right marker of a two-cell anchor and zero otherwise.
    """

    #: The anchor kind: absolute, one-cell or two-cell.
    kind: AnchorKind
    #: How the object follows row and column edits (it moves and resizes, moves only, or stays put); written for a
    #: two-cell anchor only.
    edit_as: AnchorEditAs
    #: 0-based row of the top-left marker.
    from_row: int
    #: 0-based column of the top-left marker.
    from_col: int
    #: Vertical offset of the top-left corner within ``from_row``, in EMU; for an absolute anchor, the y position from
    #: A1.
    from_row_offset_emu: int
    #: Horizontal offset of the top-left corner within ``from_col``, in EMU; for an absolute anchor, the x position from
    #: A1.
    from_col_offset_emu: int
    #: 0-based row of the bottom-right marker of a two-cell anchor. 0 unless the anchor is two-cell.
    to_row: int
    #: 0-based column of the bottom-right marker of a two-cell anchor. 0 unless the anchor is two-cell.
    to_col: int
    #: Vertical offset of the bottom-right corner within ``to_row``, in EMU. 0 unless the anchor is two-cell.
    to_row_offset_emu: int
    #: Horizontal offset of the bottom-right corner within ``to_col``, in EMU. 0 unless the anchor is two-cell.
    to_col_offset_emu: int
    #: The object's own unrotated width in EMU; a two-cell anchor's width is its box.
    width_emu: int
    #: The object's own unrotated height in EMU; a two-cell anchor's height is its box.
    height_emu: int


@dataclass
class DrawingAnchorInput:
    """A drawing anchor to add or to merge into an object's; a ``None`` field is not applied.

    On add, an omitted ``kind`` is the operation's default (one-cell for a picture, two-cell for a chart,
    shape or slicer), an omitted ``edit_as`` the anchor's own kind and every other omitted field ``0``.
    A two-cell anchor takes its box from ``to_*`` (an all-zero ``to_*`` sizes the box from ``width_emu`` /
    ``height_emu``); a one-cell anchor takes the size. A zero size is the picture's pixel size times 9525
    and Excel's default frame for a chart; a shape or slicer needs a non-zero one. On update the fields
    merge into the current anchor: an omitted ``kind`` keeps the object's kind, a given one converts it,
    and a ``to_*`` field on a one-cell anchor or a size on a two-cell one is refused.
    """

    #: The anchor kind: absolute, one-cell or two-cell; an absolute kind is not added.
    kind: Optional[AnchorKind] = None
    #: How the object follows row and column edits (it moves and resizes, moves only, or stays put); written for a
    #: two-cell anchor only.
    edit_as: Optional[AnchorEditAs] = None
    #: 0-based row of the top-left marker.
    from_row: Optional[int] = None
    #: 0-based column of the top-left marker.
    from_col: Optional[int] = None
    #: Vertical offset of the top-left corner within ``from_row``, in EMU; for an absolute anchor, the y position from
    #: A1.
    from_row_offset_emu: Optional[int] = None
    #: Horizontal offset of the top-left corner within ``from_col``, in EMU; for an absolute anchor, the x position from
    #: A1.
    from_col_offset_emu: Optional[int] = None
    #: 0-based row of the bottom-right marker of a two-cell anchor. Not accepted on a one-cell anchor.
    to_row: Optional[int] = None
    #: 0-based column of the bottom-right marker of a two-cell anchor. Not accepted on a one-cell anchor.
    to_col: Optional[int] = None
    #: Vertical offset of the bottom-right corner within ``to_row``, in EMU. Not accepted on a one-cell anchor.
    to_row_offset_emu: Optional[int] = None
    #: Horizontal offset of the bottom-right corner within ``to_col``, in EMU. Not accepted on a one-cell anchor.
    to_col_offset_emu: Optional[int] = None
    #: The object's own unrotated width in EMU. Not accepted on a two-cell anchor, except as the size of an all-zero
    #: bottom-right marker box on add.
    width_emu: Optional[int] = None
    #: The object's own unrotated height in EMU. Not accepted on a two-cell anchor, except as the size of an all-zero
    #: bottom-right marker box on add.
    height_emu: Optional[int] = None


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


class CompatibilityVersion(IntEnum):
    """Excel's Compatibility Version of a workbook.

    - ``VERSION_1``: LEN, MID, FIND, SEARCH and REPLACE count UTF-16 units.
    - ``VERSION_2``: those functions count code points (a supplementary-plane
      character is one position).
    - ``VERSION_3``: adds nested arrays, which Formulon does not implement.

    The ordinals are the values stored in the file, so they start at 1. Any
    other stored number is also accepted as a value, but
    :meth:`Workbook.set_compatibility_version` accepts only the three members.
    """

    VERSION_1 = 1
    VERSION_2 = 2
    VERSION_3 = 3

    @classmethod
    def _missing_(cls, value: object) -> Optional["CompatibilityVersion"]:
        # A stamp a later Excel writes (4 and up) reads back as a pseudo-member with that number.
        if not isinstance(value, int) or isinstance(value, bool) or value < 0:
            return None
        member = int.__new__(cls, value)
        member._name_ = f"VERSION_{value}"
        member._value_ = value
        return cls._value2member_map_.setdefault(value, member)


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
    COLUMN = 1
    VALUE = 2
    PAGE = 3


class CellShift(IntEnum):
    """The direction :meth:`Sheet.insert_cells` / :meth:`Sheet.delete_cells` move the cells of the band."""

    DOWN = 0
    RIGHT = 1
    UP = 2
    LEFT = 3


class SpillCause(IntEnum):
    """What keeps a spill or a PivotTable report from being drawn, a :class:`Spill`'s ``cause``.

    ``OCCUPIED`` covers every obstruction other than a merge. A range running past the sheet edge is
    ``BEYOND_EDGE`` whatever lies inside the grid; otherwise the first obstruction in row-major order decides.
    """

    OCCUPIED = 0
    MERGED = 1
    BEYOND_EDGE = 2


class WorkbookFormat(IntEnum):
    """Container format selector for :meth:`Workbook.save`."""

    XLSX = 0
    XLSB = 1


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
    PERCENT_OF_COLUMN = 2
    PERCENT_OF_TOTAL = 3
    RUNNING_TOTAL_IN_ROW = 4
    RUNNING_TOTAL_IN_COLUMN = 5
    INDEX = 6
    DIFFERENCE_FROM = 7
    PERCENT_DIFFERENCE_FROM = 8
    PERCENT_OF_PARENT_ROW = 9
    PERCENT_OF_PARENT_COLUMN = 10
    PERCENT_OF_PARENT = 11
    PERCENT_OF = 12
    PERCENT_OF_RUNNING_TOTAL = 13
    RANK_ASCENDING = 14
    RANK_DESCENDING = 15


class PivotFilterKind(IntEnum):
    """The type of a pivot label, value or date filter (:meth:`Pivot.add_filter`): every OOXML pivot filter type.

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


class PivotShowAsBase(IntEnum):
    """Special values of ``PivotDataFieldInput.show_as_base_item``: the "(previous)" / "(next)" base items."""

    PREVIOUS = 1048828
    NEXT = 1048829


# The "Σ Values" level in a pivot row / column field order, as read by
# ``Pivot.get_row_field_order`` / ``Pivot.get_column_field_order`` and accepted
# by ``Pivot.set_row_field_order`` / ``Pivot.set_column_field_order`` (``FM_PIVOT_VALUES_FIELD`` reinterpreted as a signed value).
PIVOT_VALUES_FIELD = -2
_PIVOT_VALUES_FIELD_C = 0xFFFFFFFE
_PIVOT_NO_DATA_FIELD = 0xFFFFFFFF

# ---------------------------------------------------------------------------
# Structured value / input dataclasses
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class CellRect:
    """Cell rectangle: 0-based, inclusive corners ``(first_row, first_col)`` through ``(last_row, last_col)``."""

    #: The 0-based first row.
    first_row: int
    #: The 0-based first column.
    first_col: int
    #: The 0-based last row, inclusive.
    last_row: int
    #: The 0-based last column, inclusive.
    last_col: int


@dataclass(frozen=True)
class CellAddress:
    """One cell: 0-based ``row`` and ``col``. A record field of this type also takes an A1 string (``"C3"``)."""

    #: The 0-based row.
    row: int
    #: The 0-based column.
    col: int


@dataclass(frozen=True)
class SheetRect:
    """A cell rectangle on sheet index ``sheet``."""

    #: The 0-based index of the sheet.
    sheet: int
    #: The cell rectangle on that sheet.
    rect: CellRect


@dataclass(frozen=True)
class Hyperlink:
    """A sheet hyperlink covering the cells of ``range``; an attribute the link does not state is ``None``."""

    #: The cells the hyperlink covers.
    range: CellRect
    #: The absolute or relative target (URL, email address, file); ``None`` for an in-workbook link.
    target: Optional[str]
    #: The in-workbook destination; ``None`` for an external target.
    location: Optional[str]
    #: The display text override; ``None`` when the hyperlink has none.
    display: Optional[str]
    #: The tooltip text; ``None`` when the hyperlink has none.
    tooltip: Optional[str]


@dataclass
class HyperlinkInput:
    """A hyperlink :meth:`Sheet.set_hyperlink` and :meth:`Sheet.insert_hyperlink_at` take.

    A :class:`Hyperlink` read back is accepted in its place.
    """

    #: The cells the link covers, a ``CellRect`` or an A1 range (``"B2"`` for one cell). Required: it is the key
    #: :meth:`Sheet.set_hyperlink` replaces on.
    range: Union[CellRect, str]
    #: The absolute or relative target (URL, email address, file); ``None`` makes an in-workbook link, which needs
    #: ``location``.
    target: Optional[str]
    #: The in-workbook destination as an A1 reference; omitted or ``None`` for an external target.
    location: Optional[str] = None
    #: The display text override; omitted or ``None`` is none.
    display: Optional[str] = None
    #: The tooltip text; omitted or ``None`` is none.
    tooltip: Optional[str] = None


@dataclass
class CommentInput:
    """The note :meth:`Sheet.set_comment` writes.

    A :class:`Comment` read back is accepted in its place.
    """

    #: The comment's text.
    text: str
    #: The comment's author; omitted or ``None`` means none.
    author: Optional[str] = None


@dataclass(frozen=True)
class Comment:
    """A cell comment; an omitted ``author`` means none."""

    #: The comment's text.
    text: str
    #: The comment's author; ``None`` when none.
    author: Optional[str] = field(default=None, kw_only=True)


@dataclass(frozen=True)
class SheetComment:
    """One enumerated sheet comment, including its anchor cell."""

    #: The cell the comment is anchored to.
    cell: CellAddress
    #: The comment's author; ``None`` when none.
    author: Optional[str]
    #: The comment's text.
    text: str


@dataclass(frozen=True)
class DataValidation:
    """A sheet ``<dataValidation>`` rule (getter result); a text the rule does not state is ``None``."""

    #: The rule's id: stable while the workbook is open, across structural edits; assigned again on load and not saved.
    id: str
    #: The cell ranges the rule applies to.
    ranges: List[CellRect]
    #: The kind of value the rule allows.
    type: ValidationType
    #: The comparison the rule applies between the entered value and its formulas.
    operator: ComparisonOperator
    #: How an invalid entry is reported: stop, warning or information.
    error_style: ValidationErrorStyle
    #: A blank cell passes the rule.
    allow_blank: bool
    #: The prompt (``prompt_title`` and ``prompt_message``) shows when the cell is selected.
    show_input_message: bool
    #: The error alert (``error_title`` and ``error_message``) shows on an invalid entry.
    show_error_message: bool
    #: The in-cell dropdown arrow is shown (``list`` rules only).
    show_in_cell_dropdown: bool
    #: The first formula or value of the rule, leading ``=`` included; ``None`` when the rule states none.
    formula1: Optional[str]
    #: The second formula or value of the rule (the upper bound of a between rule), leading ``=`` included; ``None``
    #: when the rule states none.
    formula2: Optional[str]
    #: The title of the error alert; ``None`` when the rule has none.
    error_title: Optional[str]
    #: The text of the error alert; ``None`` when the rule has none.
    error_message: Optional[str]
    #: The title of the input prompt; ``None`` when the rule has none.
    prompt_title: Optional[str]
    #: The text of the input prompt; ``None`` when the rule has none.
    prompt_message: Optional[str]

    def __post_init__(self) -> None:
        object.__setattr__(self, "type", _enum(ValidationType, self.type))
        object.__setattr__(self, "operator", _enum(ComparisonOperator, self.operator))


@dataclass
class DataValidationInput:
    """Argument shape for :meth:`Sheet.add_validation`.

    Every field except ``type`` and ``ranges`` has a default; the texts default to empty.
    """

    #: The kind of value the rule allows.
    type: ValidationType
    #: The cell ranges the rule applies to: rectangles, unqualified A1 ranges, or one A1 string; at least one is
    #: required.
    ranges: Union[List[Union[CellRect, str]], str]
    #: The comparison the rule applies between the entered value and its formulas; default is between.
    operator: ComparisonOperator = ComparisonOperator.BETWEEN
    #: How an invalid entry is reported; default is stop.
    error_style: ValidationErrorStyle = ValidationErrorStyle.STOP
    #: A blank cell passes the rule; default ``True``.
    allow_blank: bool = True
    #: Show the prompt when the cell is selected; default ``True``.
    show_input_message: bool = True
    #: Show the error alert on an invalid entry; default ``True``.
    show_error_message: bool = True
    #: Show the in-cell dropdown arrow (``list`` rules only); default ``True``.
    show_in_cell_dropdown: bool = True
    #: The first formula or value of the rule; the leading ``=`` is optional and an empty string is none.
    formula1: str = ""
    #: The second formula or value of the rule (the upper bound of a between rule); the leading ``=`` is optional and an
    #: empty string is none.
    formula2: str = ""
    #: The title of the error alert; an empty string is none.
    error_title: str = ""
    #: The text of the error alert; an empty string is none.
    error_message: str = ""
    #: The title of the input prompt; an empty string is none.
    prompt_title: str = ""
    #: The text of the input prompt; an empty string is none.
    prompt_message: str = ""


@dataclass
class AutoFilterDateGroup:
    """One ``<dateGroupItem>`` of a values filter; finer fields than ``grouping`` are zero."""

    #: Calendar year.
    year: int = 0
    #: Month of the year, 1-12; 0 when the item is coarser than month.
    month: int = 0
    #: Day of the month, 1-31; 0 when the item is coarser than day.
    day: int = 0
    #: Hour of the day, 0-23; 0 when the item is coarser than hour.
    hour: int = 0
    #: Minute of the hour, 0-59; 0 when the item is coarser than minute.
    minute: int = 0
    #: Second of the minute, 0-59; 0 when the item is coarser than second.
    second: int = 0
    #: The granularity the item matches at; fields finer than it are zero.
    grouping: DateTimeGrouping = DateTimeGrouping.YEAR


@dataclass(frozen=True)
class AutoFilterCondition:
    """One custom condition of an AutoFilter column: a comparison ``operator`` and its ``value`` text."""

    #: The comparison applied to the cell; any comparison other than between and not-between.
    operator: ComparisonOperator
    #: The text the cell is compared with (``customFilter val``).
    value: str = ""

    def __post_init__(self) -> None:
        object.__setattr__(self, "operator", _enum(ComparisonOperator, self.operator))


@dataclass
class AutoFilterColumn:
    """One ``<filterColumn>``; only the fields of the criterion selected by ``kind`` are read."""

    #: The column's 0-based offset from the filter range's first column.
    column_offset: int
    #: ``True`` when the column's filter button is hidden.
    hidden_button: bool = False
    #: ``True`` when the column's filter button is shown; default ``True``.
    show_button: bool = True
    #: Which criterion the column filters on; only that criterion's fields are meaningful.
    kind: FilterKind = FilterKind.NONE
    #: ``True`` when the values filter also matches blank cells.
    filter_blank: bool = False
    #: The displayed texts the values filter keeps.
    values: List[str] = field(default_factory=list)
    #: The date-group items the values filter keeps.
    date_groups: List[AutoFilterDateGroup] = field(default_factory=list)
    #: ``True`` when the two custom conditions are joined by AND, ``False`` when by OR.
    custom_and: bool = False
    #: The one or two custom conditions; empty unless the kind is custom.
    conditions: List[AutoFilterCondition] = field(default_factory=list)
    #: ``True`` when the top-10 filter keeps the top items, ``False`` for the bottom ones.
    top: bool = True
    #: ``True`` when ``top_value`` is a percentage rather than an item count.
    percent: bool = False
    #: The top-10 filter's rank: how many items, or what percentage.
    top_value: float = 0.0
    #: The cell value Excel last stored as the top-10 threshold; ``None`` when absent.
    filter_value: Optional[float] = None
    #: Which dynamic criterion the column filters on, such as above average or this month.
    dynamic_type: DynamicFilterType = DynamicFilterType.NULL
    #: The lower bound Excel last stored for the dynamic filter; ``None`` when absent.
    dynamic_value: Optional[float] = None
    #: The upper bound Excel last stored for the dynamic filter; ``None`` when absent.
    dynamic_max_value: Optional[float] = None
    #: The ISO date text of the dynamic filter's lower bound, kept as written; ``None`` when absent.
    value_iso: Optional[str] = None
    #: The ISO date text of the dynamic filter's upper bound, kept as written; ``None`` when absent.
    max_value_iso: Optional[str] = None
    #: The index of the differential format whose fill or font colour the color filter matches; ``None`` when the column
    #: names none.
    dxf_index: Optional[int] = None
    #: ``True`` when the color filter matches the cell fill, ``False`` when it matches the font colour.
    cell_color: bool = False
    #: The icon set the icon filter's icon comes from.
    icon_set: int = 0
    #: The index of the icon the icon filter keeps; ``None`` when the filter names no icon.
    icon_index: Optional[int] = None


@dataclass
class AutoFilterSortCondition:
    """One ``<sortCondition>`` of an AutoFilter sort state."""

    #: The cells the condition sorts by, the key column of the sorted range.
    range: CellRect
    #: ``True`` sorts in descending order.
    descending: bool = False
    #: What the condition sorts by: the value, the cell colour, the font colour or the icon.
    sort_by: SortBy = SortBy.VALUE
    #: The custom sort list, as comma-separated items; ``None`` when the condition uses none.
    custom_list: Optional[str] = None
    #: The index of the differential format whose colour the condition sorts by; ``None`` when it names none.
    dxf_index: Optional[int] = None
    #: The icon set the condition's icon comes from.
    icon_set: int = 0
    #: The index of the icon the condition sorts by; ``None`` when it names no icon.
    icon_index: Optional[int] = None


@dataclass
class AutoFilterSortState:
    """The ``<sortState>`` of an AutoFilter; ``sort_method`` is a :class:`SortMethod`."""

    #: The cells the sort applies to.
    range: CellRect
    #: ``True`` when the sort is by columns (left to right) rather than by rows.
    column_sort: bool = False
    #: ``True`` when text comparison is case-sensitive.
    case_sensitive: bool = False
    #: How East Asian text is ordered: by default order, by pinyin or by stroke count.
    sort_method: SortMethod = SortMethod.NONE
    #: The sort conditions, in priority order.
    conditions: List[AutoFilterSortCondition] = field(default_factory=list)


@dataclass(frozen=True)
class AutoFilter:
    """A sheet or table AutoFilter; ``range`` includes the header row, ``sort`` may be ``None``."""

    #: The filtered range, header row included.
    range: CellRect
    #: The filter columns, in strictly ascending column-offset order.
    columns: List[AutoFilterColumn] = field(default_factory=list)
    #: The AutoFilter's sort state; ``None`` when it has none.
    sort: Optional[AutoFilterSortState] = None


@dataclass
class AutoFilterInput:
    """An AutoFilter :meth:`Sheet.set_auto_filter` writes.

    An :class:`AutoFilter` read back is accepted in its place.
    """

    #: The filtered range, header row included, as a ``CellRect`` or an unqualified A1 range.
    range: Union[CellRect, str]
    #: The filter columns, in strictly ascending column-offset order; omitted is empty.
    columns: Sequence[AutoFilterColumn] = ()
    #: The AutoFilter's sort state; omitted or ``None`` is none.
    sort: Optional[AutoFilterSortState] = None


@dataclass(frozen=True)
class DisplayText:
    """The text Excel displays for a value and the rendering outcome (``display_status``)."""

    #: The text Excel displays for the value.
    text: str
    #: The rendering outcome, a ``DisplayStatus`` value: ok, overflow or invalid format.
    display_status: DisplayStatus


@dataclass(frozen=True)
class AutoFilterEval:
    """Rows an AutoFilter shows: body row ``first_row + i`` matches when ``match[i]`` is true."""

    #: The sheet row of the first body row the result covers.
    first_row: int
    #: One entry per body row: ``True`` when body row ``first_row + i`` satisfies the criteria.
    match: List[bool]


@dataclass(frozen=True)
class CellsInRange:
    """One page of populated cells; ``next_cursor`` is ``None`` once no further cell exists."""

    #: The populated cells of this page.
    cells: List[RangeCell]
    #: The resume position for the next page; ``None`` when the range is exhausted.
    next_cursor: Optional[int]


@dataclass(frozen=True)
class ValidationOutcome:
    """Result of :meth:`Sheet.validate_value`."""

    #: ``True`` when the value passes the rule, or when no rule covers the cell.
    valid: bool
    #: The id of the rule that was applied; ``None`` when no rule covers the cell.
    rule_id: Optional[str]
    #: The applied rule's error style; ``None`` when no rule covers the cell.
    error_style: Optional[ValidationErrorStyle]


@dataclass(frozen=True)
class Mention:
    """One ``@person`` mention covering the half-open span ``[start, end)`` of the text, in code points."""

    #: The id of the mentioned person.
    person_id: str
    #: The mention's own id, a brace-wrapped GUID.
    mention_id: str
    #: The start of the mentioned span in the comment text, as a code point offset (inclusive).
    start: int
    #: The end of the mentioned span in the comment text, as a code point offset (exclusive).
    end: int


@dataclass
class MentionInput:
    """A mention of a new or updated threaded comment. A :class:`Mention` read back is accepted in its
    place.
    """

    #: The id of the mentioned person.
    person_id: str
    #: The start of the mentioned span in the comment text, as a code point offset (inclusive).
    start: int
    #: The end of the mentioned span in the comment text, as a code point offset (exclusive); it must not be less than
    #: ``start``.
    end: int
    #: The mention's own id, a brace-wrapped GUID; omitted, one is generated.
    mention_id: Optional[str] = None


@dataclass(frozen=True)
class ThreadedComment:
    """A threaded comment: a thread opener, or a reply to the thread ``parent_id`` names."""

    #: The comment's id, a brace-wrapped GUID stored in upper case.
    id: str
    #: The cell the thread is anchored to.
    cell: CellAddress
    #: The id of the person who wrote the comment.
    person_id: str
    #: The creation time as a naive UTC ``datetime``, in whole seconds; ``None`` when the comment states none.
    created: Optional[datetime.datetime]
    #: The comment's text.
    text: str
    #: The id of the thread's opening comment this comment replies to; ``None`` for an opening comment.
    parent_id: Optional[str]
    #: ``True`` when the thread is resolved; meaningful on an opening comment only.
    done: bool
    #: The people mentioned in the text.
    mentions: List[Mention]


@dataclass
class ThreadedCommentInput:
    """A threaded comment :meth:`Sheet.add_threaded_comment` adds.

    A reply takes the anchor of the thread it answers.
    """

    #: The id of the person who wrote the comment.
    person_id: str
    #: The comment's text.
    text: str
    #: The comment's id, a brace-wrapped GUID; omitted, one is generated.
    id: Optional[str] = None
    #: The cell a new thread is anchored to, a ``CellAddress`` or an A1 cell; omitted is A1. Ignored on a reply.
    cell: Union[CellAddress, str, None] = None
    #: The creation time, in UTC: a naive ``datetime`` is read as UTC, an aware one is converted to UTC and a ``date``
    #: is midnight UTC; fractions of a second are dropped. Stored as ``YYYY-MM-DDTHH:MM:SS.00`` without a zone
    #: designator. Omitted is the current UTC time, or, with the clock pinned by :meth:`Workbook.set_pinned_now`, the
    #: pinned time written as-is (the pin is a local wall clock and is not converted).
    created: Optional[Union[datetime.datetime, datetime.date]] = None
    #: The id of the thread's opening comment to reply to; omitted or ``None`` opens a new thread, and an empty string
    #: is rejected.
    parent_id: Optional[str] = None
    #: ``True`` marks a new thread as resolved; default ``False``. Ignored on a reply.
    done: bool = False
    #: The people mentioned in the text; omitted is none.
    mentions: Sequence[Union[MentionInput, Mention]] = ()


@dataclass(frozen=True)
class Person:
    """One entry of the workbook's person list."""

    #: The person's id.
    id: str
    #: The person's display name.
    display_name: str
    #: The user id; ``None`` when the person has none.
    user_id: Optional[str] = None
    #: The identity provider id; ``None`` when the person has none.
    provider_id: Optional[str] = None


@dataclass
class PersonInput:
    """A person :meth:`Workbook.add_person` appends. A :class:`Person` read back is accepted in its
    place.
    """

    #: The person's display name.
    display_name: str
    #: The id; omitted generates a brace-wrapped GUID.
    id: Optional[str] = None
    #: The user id; omitted or ``None`` means none.
    user_id: Optional[str] = None
    #: The identity provider id; omitted or ``None`` means none.
    provider_id: Optional[str] = None


@dataclass(frozen=True)
class ImageInfo:
    """Format and pixel size of an image."""

    #: The image's file format.
    format: ImageFormat
    #: Image width in pixels.
    width_px: int
    #: Image height in pixels.
    height_px: int


@dataclass(frozen=True)
class PictureProperties(ImageInfo):
    """A picture's :class:`ImageInfo` and a copy of its media bytes (:meth:`Picture.get_properties`)."""

    #: A copy of the picture's media bytes.
    bytes: bytes


@dataclass(frozen=True)
class DrawingProperties:
    """The fields every anchored object of a sheet's drawing has; sizes and offsets are in EMU. ``description``
    is ``None`` when the object has none."""

    #: The object's id in the sheet's drawing.
    object_id: int
    #: The object's kind: picture, shape, chart, group, connector, other graphic frame or other.
    kind: DrawingObjectKind
    #: The object's name.
    name: str
    #: The object's alternative text; ``None`` when it has none.
    description: Optional[str]
    #: ``True`` when the object is hidden.
    hidden: bool
    #: Where the object sits on the sheet.
    anchor: DrawingAnchor


@dataclass(frozen=True)
class ShapeProperties:
    """The shape fields of a shape, connector or group; the anchor, name, description and visibility are on
    :class:`DrawingProperties`.
    """

    #: The object's id in the sheet's drawing.
    object_id: int
    #: The object's kind: shape, connector or group.
    kind: DrawingObjectKind
    #: Name of the preset geometry; ``None`` for custom geometry or a group.
    preset_geometry: Optional[str]
    #: The shape's text, paragraphs joined with a newline.
    text: str
    #: ``True`` when the shape is a text box.
    text_box: bool
    #: The shape's own fill; ``None`` when it has none and draws its style's.
    fill: Optional[ChartFill]
    #: The shape's own outline; ``None`` when it has none and draws its style's.
    line: Optional[ChartLine]
    #: Clockwise rotation in degrees.
    rotation_deg: float


@dataclass
class DrawingPlacementInput:
    """Name, description, visibility and placement of a new drawing object (:meth:`Sheet.add_picture`, the
    placement of a :class:`ChartInput`; :class:`ShapeInput` and :class:`SlicerInput` extend it).
    """

    #: The object's name; empty or omitted takes the kind's default.
    name: str = ""
    #: The object's alternative text; empty or omitted means none.
    description: str = ""
    #: Whether the object is hidden; unset means shown.
    hidden: Optional[bool] = None
    #: Where the object is placed; unset takes the operation's default anchor kind at A1.
    anchor: Optional[DrawingAnchorInput] = None


@dataclass
class ShapeInput(DrawingPlacementInput):
    """A shape :meth:`Sheet.add_shape` appends; a ``None`` field is not applied.

    ``name`` defaults to ``Shape <id>``. The anchor is two-cell by default; an absolute anchor is not added.
    """

    #: Name of the preset geometry; defaults to rect, and a line or connector geometry makes a connector.
    preset_geometry: Optional[str] = None
    #: The shape's text, paragraphs joined with a newline; a connector holds none.
    text: Optional[str] = None
    #: Whether the shape is a text box.
    text_box: Optional[bool] = None
    #: The shape's own fill; unset draws the shape style's.
    fill: Optional[ChartFill] = None
    #: The shape's own outline; unset draws the shape style's.
    line: Optional[ChartLine] = None
    #: Clockwise rotation in degrees.
    rotation_deg: Optional[float] = None


@dataclass(frozen=True)
class SlicerProperties:
    """The slicer fields of a table slicer; the frame's anchor, name, description and visibility are on
    :class:`DrawingProperties`.

    The selection itself lives in the table's AutoFilter.
    """

    #: The id of the slicer's frame in the sheet's drawing.
    object_id: int
    #: The caption shown in the slicer header.
    caption: str
    #: The name of the slicer cache the slicer belongs to.
    cache_name: str
    #: The name of the table the slicer filters.
    table_name: str
    #: The name of the table column the slicer filters.
    column_name: str
    #: The name of the slicer style; ``None`` without a style.
    style_name: Optional[str]
    #: The number of columns the slicer lays its item buttons out in.
    column_count: int
    #: The height of each item row in points.
    row_height_pt: float
    #: The order the slicer lists its items in.
    sort_order: SlicerSortOrder
    #: How the slicer cross-filters its items (items with data listed first, items without data shown in place, or no
    #: cross-filtering).
    cross_filter: SlicerCrossFilter


@dataclass
class SlicerInput(DrawingPlacementInput):
    """A slicer :meth:`Sheet.add_slicer` adds; a ``None`` field is not applied.

    The frame is always two-cell (``anchor.kind`` ``None`` or ``TWO_CELL``) and needs a box or a size. The
    name and caption default to the column name (the name is made unique).
    """

    #: The caption shown in the slicer header; the column name when omitted.
    caption: Optional[str] = None
    #: The name of the slicer style; ``None`` creates the slicer without a style and an empty string is rejected.
    #: Omitted leaves the default.
    style_name: Optional[str] = None
    #: The number of columns the slicer lays its item buttons out in; 1 when omitted.
    column_count: Optional[int] = None
    #: The height of each item row in points, finite and not negative; 19 when omitted.
    row_height_pt: Optional[float] = None
    #: The order the slicer lists its items in; ascending when omitted.
    sort_order: Optional[SlicerSortOrder] = None
    #: How the slicer cross-filters its items (items with data listed first, items without data shown in place, or no
    #: cross-filtering); the default cross filter when omitted.
    cross_filter: Optional[SlicerCrossFilter] = None


@dataclass(frozen=True)
class DataTableInputCell:
    """An input cell of a data table; ``deleted`` marks one removed by a row or column delete."""

    #: The input cell's 0-based row.
    row: int
    #: The input cell's 0-based column.
    col: int
    #: ``True`` when a row or column delete removed the input cell; ignored on input.
    deleted: bool


@dataclass(frozen=True)
class DataTable:
    """A What-If data table (``TABLE()``): its result range and input cells, all 0-based.

    The row input cell receives the values of the row above the range, the column input cell those of the column
    left of it. A one-variable table has one input (``row_oriented`` says which, the other is ``None``), a
    two-variable table both. A deleted input makes the table evaluate to ``#REF!``.
    """

    #: The result range of the table, 0-based and inclusive.
    range: CellRect
    #: ``True`` for a two-variable table, which has both a row and a column input cell.
    two_dimensional: bool
    #: For a one-variable table, ``True`` when its input is the row input cell and ``False`` when it is the column input
    #: cell.
    row_oriented: bool
    #: The cell that receives the values of the row above the range; ``None`` when the table has none.
    row_input: Optional[DataTableInputCell]
    #: The cell that receives the values of the column left of the range; ``None`` when the table has none.
    column_input: Optional[DataTableInputCell]


@dataclass
class DataTableInput:
    """A data table :meth:`Sheet.add_data_table` creates.

    ``row_input`` / ``column_input`` each take a :class:`CellAddress`, a :class:`DataTableInputCell` read back or
    one A1 cell.
    """

    #: The result range of the table, a ``CellRect`` or an unqualified A1 range.
    range: Union[CellRect, str]
    #: The cell that receives the values of the row above the range; omitted or ``None`` for none. Give it alone or with
    #: ``column_input``.
    row_input: Union[CellAddress, DataTableInputCell, str, None] = None
    #: The cell that receives the values of the column left of the range; omitted or ``None`` for none. Give it alone or
    #: with ``row_input``.
    column_input: Union[CellAddress, DataTableInputCell, str, None] = None


@dataclass(frozen=True)
class Sparkline:
    """One sparkline: the data ``source`` and its single ``location`` cell."""

    #: The data range the sparkline plots; ``None`` when the stored range text does not parse, in which case that text
    #: is kept so the sparkline is written back unchanged.
    source: Optional[SheetRect]
    #: The single cell that shows the sparkline.
    location: CellAddress
    _stored_source: Optional[str] = field(default=None, init=False, repr=False, compare=False)


@dataclass
class SparklineInput:
    """One sparkline of an input. A :class:`Sparkline` read back is accepted in its place."""

    #: The data range, as a range record or sheet-qualified A1 text; unqualified text resolves on the receiving sheet.
    source: Union[SheetRect, str, None]
    #: The cell that shows the sparkline, as an address record or A1 text.
    location: Union[CellAddress, str]


class BlanksAs(IntEnum):
    """How blank cells plot (a chart or a sparkline group); a chart without the element reads ``None``."""

    ZERO = 0
    GAP = 1
    SPAN = 2


@dataclass(frozen=True)
class SparklineGroup:
    """A sparkline group of a sheet's worksheet extension list."""

    #: The group's id: an opaque string, unique on its sheet, stable while the workbook is open and not saved; a read
    #: always carries one.
    id: str = ""
    #: The sparkline type: line, column or stacked (win/loss).
    type: SparklineType = SparklineType.LINE
    #: Colour of the line or columns; ``None`` when the group states none.
    color_series: Optional[Color] = None
    #: Colour of negative points; ``None`` when the group states none.
    color_negative: Optional[Color] = None
    #: Colour of the horizontal axis; ``None`` when the group states none.
    color_axis: Optional[Color] = None
    #: Colour of the markers; ``None`` when the group states none.
    color_markers: Optional[Color] = None
    #: Colour of the first point; ``None`` when the group states none.
    color_first: Optional[Color] = None
    #: Colour of the last point; ``None`` when the group states none.
    color_last: Optional[Color] = None
    #: Colour of the highest point; ``None`` when the group states none.
    color_high: Optional[Color] = None
    #: Colour of the lowest point; ``None`` when the group states none.
    color_low: Optional[Color] = None
    #: ``True`` to show a marker at every point of a line sparkline.
    markers: bool = False
    #: ``True`` to highlight the highest point.
    high: bool = False
    #: ``True`` to highlight the lowest point.
    low: bool = False
    #: ``True`` to highlight the first point.
    first: bool = False
    #: ``True`` to highlight the last point.
    last: bool = False
    #: ``True`` to highlight negative points.
    negative: bool = False
    #: ``True`` to show the horizontal axis.
    display_x_axis: bool = False
    #: ``True`` to plot data in hidden rows and columns.
    display_hidden: bool = False
    #: ``True`` to plot the group right to left; the default is ``False``.
    right_to_left: bool = False
    #: How the vertical axis minimum is chosen: per sparkline, shared by the group, or the custom value.
    min_axis_type: SparklineAxisType = SparklineAxisType.INDIVIDUAL
    #: How the vertical axis maximum is chosen: per sparkline, shared by the group, or the custom value.
    max_axis_type: SparklineAxisType = SparklineAxisType.INDIVIDUAL
    #: The custom vertical axis minimum; ``None`` when not set.
    manual_min: Optional[float] = None
    #: The custom vertical axis maximum; ``None`` when not set.
    manual_max: Optional[float] = None
    #: ``True`` to space points by a date axis.
    date_axis: bool = False
    #: The group's date range; ``None`` when absent or when its stored text does not parse.
    date_axis_range: Optional[SheetRect] = None
    #: Line width in points; 0.75 for a new group.
    line_width_pt: float = 0.75
    #: How empty cells are plotted: as a gap, as zero, or by connecting the neighbours.
    empty_cells_as: BlanksAs = BlanksAs.GAP
    #: The sparklines of the group.
    sparklines: List[Sparkline] = field(default_factory=list)
    _stored_date_axis: Optional[str] = field(default=None, init=False, repr=False, compare=False)

    def __post_init__(self) -> None:
        object.__setattr__(self, "empty_cells_as", _enum(BlanksAs, self.empty_cells_as))


@dataclass
class SparklineGroupInput:
    """A sparkline group :meth:`Sheet.add_sparkline_group` adds: every :class:`SparklineGroup` field, each
    defaulting as a new line group does. A :class:`SparklineGroup` read back is accepted in its place.
    """

    #: The sparkline type: line, column or stacked (win/loss).
    type: SparklineType = SparklineType.LINE
    #: Colour of the line or columns; ``None`` for no colour.
    color_series: Optional[ColorInput] = None
    #: Colour of negative points; ``None`` for no colour.
    color_negative: Optional[ColorInput] = None
    #: Colour of the horizontal axis; ``None`` for no colour.
    color_axis: Optional[ColorInput] = None
    #: Colour of the markers; ``None`` for no colour.
    color_markers: Optional[ColorInput] = None
    #: Colour of the first point; ``None`` for no colour.
    color_first: Optional[ColorInput] = None
    #: Colour of the last point; ``None`` for no colour.
    color_last: Optional[ColorInput] = None
    #: Colour of the highest point; ``None`` for no colour.
    color_high: Optional[ColorInput] = None
    #: Colour of the lowest point; ``None`` for no colour.
    color_low: Optional[ColorInput] = None
    #: ``True`` to show a marker at every point of a line sparkline.
    markers: bool = False
    #: ``True`` to highlight the highest point.
    high: bool = False
    #: ``True`` to highlight the lowest point.
    low: bool = False
    #: ``True`` to highlight the first point.
    first: bool = False
    #: ``True`` to highlight the last point.
    last: bool = False
    #: ``True`` to highlight negative points.
    negative: bool = False
    #: ``True`` to show the horizontal axis.
    display_x_axis: bool = False
    #: ``True`` to plot data in hidden rows and columns.
    display_hidden: bool = False
    #: ``True`` to plot the group right to left; the default is ``False``.
    right_to_left: bool = False
    #: How the vertical axis minimum is chosen: per sparkline, shared by the group, or the custom value.
    min_axis_type: SparklineAxisType = SparklineAxisType.INDIVIDUAL
    #: How the vertical axis maximum is chosen: per sparkline, shared by the group, or the custom value.
    max_axis_type: SparklineAxisType = SparklineAxisType.INDIVIDUAL
    #: The custom vertical axis minimum; ``None`` when not set.
    manual_min: Optional[float] = None
    #: The custom vertical axis maximum; ``None`` when not set.
    manual_max: Optional[float] = None
    #: ``True`` to space points by a date axis.
    date_axis: bool = False
    #: The group's date range, as a range record or sheet-qualified A1 text; ``None`` for none.
    date_axis_range: Union[SheetRect, str, None] = None
    #: Line width in points; 0.75 for a new group.
    line_width_pt: float = 0.75
    #: How empty cells are plotted: as a gap, as zero, or by connecting the neighbours.
    empty_cells_as: BlanksAs = BlanksAs.GAP
    #: The sparklines of the group; an add needs at least one.
    sparklines: Sequence[Union[SparklineInput, Sparkline]] = ()

    def __post_init__(self) -> None:
        self.empty_cells_as = _enum(BlanksAs, self.empty_cells_as)


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

    #: Whether the input is empty, a range reference or literal points.
    kind: ChartSourceKind
    #: The reference text of a range-reference input; empty for another kind.
    formula: str = ""
    #: The literal points of a literal input, numbers or strings in order; empty for another kind.
    literal: List[Union[float, str]] = field(default_factory=list)


@dataclass(frozen=True)
class ChartTitle:
    """A chart or axis title.

    A ``TEXT`` title needs ``text``.
    """

    #: Whether the title is absent, automatic, literal text or a cell reference.
    kind: ChartTitleKind
    #: The flattened rich text of a text title; ``None`` for another kind.
    text: Optional[str] = None
    #: The reference of a reference title; ``None`` for another kind.
    formula: Optional[str] = None


@dataclass(frozen=True)
class ChartFill:
    """Solid fill of a chart element."""

    #: The solid fill colour, an RGB or theme colour; ``None`` when the element states none.
    color: Optional[Color] = None
    #: The fill transparency as a percentage (0..100); ``None`` when the colour has no alpha.
    transparency_pct: Optional[float] = None
    #: ``True`` for an explicit no-fill.
    none: bool = False
    #: ``True`` on a read when the stored paint cannot be stated by this record (a gradient, pattern or picture); read
    #: only.
    other: bool = False


@dataclass(frozen=True)
class ChartLine:
    """Outline of a chart element."""

    #: The outline colour, an RGB or theme colour; ``None`` when the element states none.
    color: Optional[Color] = None
    #: The outline transparency as a percentage (0..100); ``None`` when the colour has no alpha.
    transparency_pct: Optional[float] = None
    #: The outline width in points; ``None`` when the element states none.
    width_pt: Optional[float] = None
    #: The dash style of the outline; ``None`` when the element states none.
    dash: Optional[ChartLineDash] = None
    #: ``True`` for an explicit no-line.
    none: bool = False
    #: ``True`` on a read when the outline fill cannot be stated by this record (a gradient or pattern); read only.
    other: bool = False


@dataclass(frozen=True)
class ChartMarker:
    """Series or point marker; ``size_pt`` is in points."""

    #: The marker symbol.
    symbol: ChartMarkerSymbol
    #: The marker size in points; ``None`` when the marker states none.
    size_pt: Optional[int]
    #: The marker fill; ``None`` when the marker states none of its own.
    fill: ChartFill
    #: The marker outline; ``None`` when the marker states none of its own.
    line: ChartLine


@dataclass(frozen=True)
class ChartLabels:
    """Data labels of a series."""

    #: ``True`` when the label shows the point value.
    show_value: bool
    #: ``True`` when the label shows the category name.
    show_category: bool
    #: ``True`` when the label shows the series name.
    show_series: bool
    #: ``True`` when the label shows the percentage of the whole (pie-type plots).
    show_percent: bool
    #: ``True`` when the label shows the legend key.
    show_legend_key: bool
    #: The text placed between the label parts; ``None`` when the chart states none.
    separator: Optional[str]
    #: Where the label sits relative to its point; ``None`` when the chart states none.
    position: Optional[ChartLabelPosition]
    #: The number format code of the label; ``None`` when the chart states none.
    number_format: Optional[str]
    #: ``True`` when the label number format follows the source cells.
    source_linked: bool
    #: The label font colour; ``None`` when the chart states none.
    font_color: Optional[Color]
    #: The label font size in points; ``None`` when the chart states none.
    font_size_pt: Optional[float]
    #: ``True`` when the labels are deleted (c:delete) for the series.
    deleted: bool


@dataclass(frozen=True)
class ChartPoint:
    """One formatted data point (``c:dPt``)."""

    #: The 0-based index of the data point within its series.
    index: int
    #: The point fill; ``None`` when the point states none of its own.
    fill: Optional[ChartFill]
    #: The point outline; ``None`` when the point states none of its own.
    line: Optional[ChartLine]
    #: The point marker; ``None`` when the point states none of its own.
    marker: Optional[ChartMarker]


@dataclass(frozen=True)
class ChartSeries:
    """One series of a plot group; ``series_id`` is its ``c:idx``, the key the series methods take."""

    #: The series' ``c:idx``, the key the series update and remove operations take.
    series_id: int
    #: The plotting order of the series within the chart (``c:order``).
    order: int
    #: ``True`` when the chart filter hides the series.
    filtered: bool
    #: The input that supplies the series name.
    name: ChartDataSource
    #: The input that supplies the category labels.
    categories: ChartDataSource
    #: The input that supplies the values of a category plot.
    values: ChartDataSource
    #: The input that supplies the x values of a scatter or bubble series.
    x: ChartDataSource
    #: The input that supplies the y values of a scatter or bubble series.
    y: ChartDataSource
    #: The input that supplies the bubble sizes of a bubble series.
    size: ChartDataSource
    #: The series fill; ``None`` when the series states none of its own.
    fill: Optional[ChartFill]
    #: The series outline; ``None`` when the series states none of its own.
    line: Optional[ChartLine]
    #: The series marker; ``None`` when the series has none.
    marker: Optional[ChartMarker]
    #: ``True`` when a line or scatter series draws a smoothed line; ``None`` when the series states none.
    smooth: Optional[bool]
    #: The formatted data points of the series.
    points: List[ChartPoint]
    #: The data labels of the series; ``None`` when the series has none.
    labels: Optional[ChartLabels]


@dataclass(frozen=True)
class ChartPlot:
    """One plot group."""

    #: The plot type of the group.
    type: ChartPlotType
    #: The series layout of an extended plot; ``None`` for any other plot type.
    extended_layout: Optional[ChartExtendedLayout]
    #: Whether bars run vertically (column) or horizontally; ``None`` outside the bar family or when the file omits it.
    bar_direction: Optional[ChartBarDirection]
    #: How the series are stacked or grouped (standard, stacked, percent-stacked); ``None`` when the plot type has none.
    grouping: Optional[ChartGrouping]
    #: ``True`` when each point of a single-series plot takes its own colour.
    vary_colors: bool
    #: The gap between bar clusters as a percentage of the bar width; ``None`` when the plot type has none.
    gap_width_pct: Optional[int]
    #: The overlap between bars of a cluster as a percentage of the bar width; ``None`` when the plot type has none.
    overlap_pct: Optional[int]
    #: The doughnut hole size as a percentage of the doughnut diameter; ``None`` when the plot type has none.
    hole_size_pct: Optional[int]
    #: The angle in degrees of the first pie or doughnut slice; ``None`` when the plot type has none.
    first_slice_angle_deg: Optional[int]
    #: The ids of the axes the plot group uses.
    axis_ids: List[int]
    #: The series of the plot group, in order.
    series: List[ChartSeries]


@dataclass(frozen=True)
class ChartAxis:
    """One axis; ``crosses_at`` is ``None`` unless ``crosses`` is ``AT``, and ``title`` is ``None`` without a title."""

    #: The axis id (c:axId) that ``cross_axis`` and a plot group's axis ids refer to.
    id: int
    #: The kind of the axis (category, value, date or series).
    kind: ChartAxisKind
    #: The side of the plot area the axis sits on.
    position: ChartAxisPosition
    #: ``True`` when the axis is shown; ``False`` for a deleted axis (``c:delete``).
    visible: bool
    #: The id of the axis this axis crosses.
    cross_axis: int
    #: Where this axis crosses its cross axis (automatic, minimum, maximum or at ``crosses_at``); ``None`` when the file
    #: omits it.
    crosses: Optional[ChartAxisCrosses]
    #: The value on the cross axis where this axis crosses; ``None`` unless ``crosses`` is the at-value kind.
    crosses_at: Optional[float]
    #: The fixed minimum of the axis scale; ``None`` when it is automatic.
    min: Optional[float]
    #: The fixed maximum of the axis scale; ``None`` when it is automatic.
    max: Optional[float]
    #: The fixed distance between major ticks; ``None`` when it is automatic.
    major_unit: Optional[float]
    #: The fixed distance between minor ticks; ``None`` when it is automatic.
    minor_unit: Optional[float]
    #: The base of a logarithmic scale; ``None`` for a linear axis.
    log_base: Optional[float]
    #: ``True`` when the axis runs in reverse order.
    reversed: bool
    #: The number format code of the tick labels; ``None`` when the axis states none.
    number_format: Optional[str]
    #: ``True`` when the tick label number format follows the source cells.
    source_linked: bool
    #: The axis title; ``None`` when the axis has none.
    title: Optional[ChartTitle]
    #: ``True`` when major gridlines are drawn.
    major_gridlines: bool
    #: ``True`` when minor gridlines are drawn.
    minor_gridlines: bool
    #: Where the tick labels sit relative to the axis; ``None`` when the axis element is absent.
    tick_label_position: Optional[ChartTickLabelPosition]


@dataclass(frozen=True)
class ChartAxisKey:
    """An axis addressed by its kind within plot group ``plot`` (the first by default)."""

    #: The kind of axis to address.
    kind: ChartAxisKind
    #: The 0-based position of the plot group; 0 when omitted.
    plot: int = 0


#: An axis id, a :class:`ChartAxisKind` (the first plot group's axis of that kind) or a :class:`ChartAxisKey`.
ChartAxisRef = Union[int, ChartAxisKind, ChartAxisKey]


def _resolve_chart_axis(axis: ChartAxisRef, properties: ChartProperties) -> int:
    """The axis id ``axis`` names in the chart ``properties`` read.

    An integer is an axis id and is returned unchanged for the engine to check.

    Raises:
      TypeError: when ``axis`` is not an integer, a :class:`ChartAxisKind` or a :class:`ChartAxisKey`.
      IndexOutOfRangeError: when the key's ``plot`` is past the last plot group.
      NotFoundError: when the plot group has no axis of that kind.
    """
    if isinstance(axis, bool) or not isinstance(axis, (int, ChartAxisKey)):
        raise TypeError(f"axis must be an int, a ChartAxisKind or a ChartAxisKey, got {type(axis).__name__}")
    if isinstance(axis, int) and not isinstance(axis, ChartAxisKind):
        return axis
    key = axis if isinstance(axis, ChartAxisKey) else ChartAxisKey(axis)
    plot = _uint(key.plot, "plot", role="position")
    if plot >= len(properties.plots):
        raise FormulonError(StatusCode.INDEX_OUT_OF_RANGE, f"plot {plot} is past the last plot group", f"plot={plot}")
    kinds = {a.id: a.kind for a in properties.axes}
    for axis_id in properties.plots[plot].axis_ids:
        if kinds.get(axis_id) == key.kind:
            return axis_id
    raise FormulonError(StatusCode.NOT_FOUND, f"plot {plot} has no {key.kind!r} axis", f"plot={plot}")


@dataclass(frozen=True)
class ChartLegend:
    """Legend state; ``hidden_entries`` lists the ``c:idx`` of series whose legend entry is hidden."""

    #: Where the legend sits relative to the plot.
    position: ChartLegendPosition
    #: ``True`` when the legend overlays the plot area.
    overlay: bool
    #: The ``c:idx`` values of the series whose legend entry is hidden.
    hidden_entries: List[int] = field(default_factory=list)


@dataclass(frozen=True)
class ChartSource:
    """The rectangle every series was split from and how it was split."""

    #: The rectangle every series was split from.
    range: SheetRect
    #: Whether the series run along rows or columns of ``range``.
    direction: ChartSeriesDirection
    #: ``True`` when the first row of ``range`` holds series or category names rather than data.
    header_row: bool
    #: ``True`` when the first column of ``range`` holds series or category names rather than data.
    header_column: bool


@dataclass
class ChartSourceInput:
    """A chart source :meth:`Chart.set_source` takes. A :class:`ChartSource` read back is accepted in its
    place.
    """

    #: The source range, a rectangle or an A1 range (an unqualified one resolves on the chart's sheet).
    range: Union[SheetRect, str]
    #: Whether the series run along rows or columns of the range; automatic when omitted.
    direction: ChartSeriesDirection = ChartSeriesDirection.AUTO
    #: ``True`` when the first row holds names; ``None`` or omitted follows Excel's detection.
    header_row: Optional[bool] = None
    #: ``True`` when the first column holds names; ``None`` or omitted follows Excel's detection.
    header_column: Optional[bool] = None


@dataclass(frozen=True)
class ChartFilter:
    """The series and categories a chart filter hides, as ``c:idx`` / category indices; both empty when unfiltered."""

    #: The ``c:idx`` values of the series the filter hides.
    hidden_series: List[int] = field(default_factory=list)
    #: The 0-based indexes of the categories the filter hides.
    hidden_categories: List[int] = field(default_factory=list)


@dataclass
class ChartInput(DrawingPlacementInput):
    """A chart :meth:`Sheet.add_chart` creates; the name, description, visibility and ``anchor`` are those of
    :class:`DrawingPlacementInput`.

    The anchor reads as for :meth:`Sheet.add_picture`, except that a zero ``width_emu`` or ``height_emu`` means
    Excel's default 360 x 216 pt frame and an omitted anchor kind is two-cell.
    """

    #: The plot type of the new chart: bar, line, area, pie, doughnut or scatter.
    type: ChartPlotType = field(kw_only=True)
    #: The source range, a rectangle or an A1 range (an unqualified one resolves on the receiving sheet).
    range: Union[SheetRect, str] = field(kw_only=True)
    #: Whether bars run vertically (column) or horizontally; applies to bar plots only, column when omitted.
    bar_direction: Optional[ChartBarDirection] = None
    #: How the series are stacked or grouped; applies to bar, line and area plots.
    grouping: ChartGrouping = ChartGrouping.CLUSTERED
    #: Whether the series run along rows or columns of the source range; automatic when omitted.
    direction: ChartSeriesDirection = ChartSeriesDirection.AUTO
    #: ``True`` when the first row of the source holds names; ``None`` or omitted follows Excel's detection.
    header_row: Optional[bool] = None
    #: ``True`` when the first column of the source holds names; ``None`` or omitted follows Excel's detection.
    header_column: Optional[bool] = None
    #: The chart title; omitted is Excel's automatic title, ``None`` is no title, and empty text is rejected.
    title: Union[str, ChartTitle, None] = ChartTitle(ChartTitleKind.AUTO)
    #: The legend position; omitted is the bottom (Excel's default for a new chart), ``None`` is no legend.
    legend: Optional[ChartLegendPosition] = ChartLegendPosition.BOTTOM


@dataclass(frozen=True)
class ChartProperties:
    """The typed view of a chart (:meth:`Chart.get_properties`); the anchor, name and description are on
    :class:`DrawingProperties`.
    """

    #: The id of the chart's drawing object on its sheet.
    object_id: int
    #: ``True`` for an extended (``cx:``) chart.
    is_extended: bool
    #: ``True`` when the chart is bound to a pivot table.
    is_pivot_chart: bool
    #: ``True`` when the chart lives on a chartsheet rather than as an embedded object.
    on_chartsheet: bool
    #: The chart title; ``None`` when the chart has none.
    title: Optional[ChartTitle]
    #: ``True`` when the automatic title is deleted; ``None`` when the chart states none.
    auto_title_deleted: Optional[bool]
    #: The chart legend; ``None`` when the chart has none.
    legend: Optional[ChartLegend]
    #: How empty cells are plotted (gap, zero or span); ``None`` when the file omits it.
    blanks_as: Optional[BlanksAs]
    #: ``True`` when only visible cells are plotted; ``None`` when the chart states none.
    plot_visible_only: Optional[bool]
    #: ``True`` when the chart uses the 1904 date system.
    date1904: bool
    #: The chart style number (1..48); ``None`` when the chart states none.
    style: Optional[int]
    #: The plot groups of the chart.
    plots: List[ChartPlot]
    #: The axes of the chart.
    axes: List[ChartAxis]
    #: The series and categories the chart filter hides.
    filter: ChartFilter
    #: The range every series was split from; ``None`` when the series share no such rectangle.
    source: Optional[ChartSource]
    #: ``True`` when ``#N/A`` points draw as blank cells; ``None`` when the chart carries no setting.
    na_as_blank: Optional[bool]


@dataclass(frozen=True)
class ChartSeriesRanges:
    """The resolved rectangles of each series input."""

    #: The rectangles the series name reads.
    name: List[SheetRect] = field(default_factory=list)
    #: The rectangles the category labels read.
    categories: List[SheetRect] = field(default_factory=list)
    #: The rectangles the values read.
    values: List[SheetRect] = field(default_factory=list)
    #: The rectangles the x values read.
    x: List[SheetRect] = field(default_factory=list)
    #: The rectangles the y values read.
    y: List[SheetRect] = field(default_factory=list)
    #: The rectangles the bubble sizes read.
    size: List[SheetRect] = field(default_factory=list)


@dataclass(frozen=True)
class ChartDataSeries:
    """One resolved series; ``values`` and ``states`` (:class:`ChartPointState`) are parallel."""

    #: The 0-based position of the series' plot group.
    plot_index: int
    #: The series' ``c:idx``.
    series_id: int
    #: The resolved series name; ``None`` when the series has no name reference.
    name: Optional[str]
    #: The number format code of the values (the y values in the XY domain).
    number_format: str
    #: The number format code of the x values; ``None`` outside the XY domain.
    x_number_format: Optional[str]
    #: ``True`` when the chart filter hides the series.
    filtered: bool
    #: The resolved values, one per point (the y values in the XY domain).
    values: List[float]
    #: The state of each point, parallel to ``values``.
    states: List[ChartPointState]
    #: The resolved x values, one per point; ``None`` outside the XY domain.
    x: Optional[List[float]]
    #: The resolved bubble sizes, one per point; ``None`` unless the series is a bubble series.
    sizes: Optional[List[float]]
    #: The cell ranges each input of the series reads.
    ranges: ChartSeriesRanges


@dataclass(frozen=True)
class ChartCategories:
    """Resolved categories; ``values`` is set for numeric (date) categories and ``generated`` when the labels are 1..n."""

    #: The category labels, one per category.
    labels: List[str]
    #: The labels of the upper levels of a multi-level category axis, outermost first; the lowest level is ``labels``. A
    #: group label sits at the points where it is not blank and is empty elsewhere.
    levels: List[List[str]]
    #: The numeric (date) category values; ``None`` when the categories are not numeric.
    values: Optional[List[Optional[float]]]
    #: The number format code of the numeric category values; ``None`` when not numeric.
    number_format: Optional[str]
    #: ``True`` when the chart has no category reference and the labels are generated as 1..n.
    generated: bool


@dataclass(frozen=True)
class ChartAxisTitleText:
    """Resolved title text of one axis."""

    #: The id of the axis the title belongs to.
    axis_id: int
    #: The resolved title text.
    text: str


@dataclass(frozen=True)
class ChartUnresolved:
    """An input that did not resolve; its points read as gaps."""

    #: The ``c:idx`` of the series the input belongs to; 0 for a title.
    series_id: int
    #: The role of the input that did not resolve.
    role: ChartRole
    #: The id of the axis when the role is an axis title.
    axis_id: int


@dataclass
class ChartData:
    """A chart's series resolved against the current cell values (:meth:`Chart.get_data`)."""

    #: Whether the data is in the category domain or the XY domain.
    domain: ChartDomain
    #: The number of points per series.
    point_count: int
    #: The resolved category labels.
    categories: ChartCategories
    #: The resolved series.
    series: List[ChartDataSeries]
    #: The resolved chart title text; ``None`` when the chart has none.
    title_text: Optional[str]
    #: The resolved title text of each axis that has one.
    axis_title_texts: List[ChartAxisTitleText]
    #: The inputs that did not resolve; their points read as gaps.
    unresolved: List[ChartUnresolved]


class ChartFillPatch(TypedDict, total=False):
    """Fill patch: an omitted key is kept, ``None`` removes the element; a colour is an RGB or theme ``ColorInput`` (indexed and automatic are rejected)."""

    #: The solid fill colour, an RGB or theme colour. Omitted keeps it, ``None`` removes it.
    color: Optional[ColorInput]
    #: The fill transparency as a percentage, 0 to 100. Omitted keeps it, ``None`` removes it.
    transparency_pct: Optional[float]
    #: ``True`` writes an explicit no-fill and ``False`` removes it; omitted keeps it, ``None`` is rejected.
    none: Optional[bool]


class ChartLinePatch(TypedDict, total=False):
    """Outline patch; a colour is an RGB or theme ``ColorInput`` (indexed and automatic are rejected)."""

    #: The outline colour, an RGB or theme colour. Omitted keeps it, ``None`` removes it.
    color: Optional[ColorInput]
    #: The outline transparency as a percentage, 0 to 100. Omitted keeps it, ``None`` removes it.
    transparency_pct: Optional[float]
    #: The outline width in points. Omitted keeps it, ``None`` removes it.
    width_pt: Optional[float]
    #: The dash style of the outline. Omitted keeps it, ``None`` removes it.
    dash: Optional[ChartLineDash]
    #: ``True`` writes an explicit no-line and ``False`` removes it. Omitted keeps it; ``None`` is rejected.
    none: Optional[bool]


class ChartMarkerPatch(TypedDict, total=False):
    """Marker patch."""

    #: The marker symbol; the picture symbol is rejected. Omitted keeps it; ``None`` is rejected.
    symbol: Optional[ChartMarkerSymbol]
    #: The marker size in points. Omitted keeps it, ``None`` removes it.
    size_pt: Optional[float]
    #: The patch applied to the marker fill. Omitted keeps it, ``None`` removes it.
    fill: Optional[ChartFillPatch]
    #: The patch applied to the marker outline. Omitted keeps it, ``None`` removes it.
    line: Optional[ChartLinePatch]


class ChartLabelsPatch(TypedDict, total=False):
    """Data-label patch."""

    #: Whether the point value is shown. Omitted keeps it; ``None`` is rejected.
    show_value: Optional[bool]
    #: Whether the category name is shown. Omitted keeps it; ``None`` is rejected.
    show_category: Optional[bool]
    #: Whether the series name is shown. Omitted keeps it; ``None`` is rejected.
    show_series: Optional[bool]
    #: Whether the percentage of the whole is shown. Omitted keeps it; ``None`` is rejected.
    show_percent: Optional[bool]
    #: Whether the legend key is shown. Omitted keeps it; ``None`` is rejected.
    show_legend_key: Optional[bool]
    #: The text placed between the label parts. Omitted keeps it, ``None`` removes it.
    separator: Optional[str]
    #: Where the label sits relative to its point. Omitted keeps it, ``None`` removes it.
    position: Optional[ChartLabelPosition]
    #: The number format code of the label. Omitted keeps it, ``None`` removes it.
    number_format: Optional[str]
    #: The label font colour, an RGB or theme colour. Omitted keeps it, ``None`` removes it.
    font_color: Optional[ColorInput]
    #: The label font size in points. Omitted keeps it, ``None`` removes it.
    font_size_pt: Optional[float]
    #: Whether the label number format follows the source cells. Omitted keeps it; ``None`` is rejected.
    source_linked: Optional[bool]
    #: ``True`` hides the labels. Omitted keeps it; ``None`` is rejected.
    deleted: Optional[bool]


class ChartSeriesPatch(TypedDict, total=False):
    """Series patch."""

    #: The series name input, a reference or a data source; an empty reference or ``None`` removes it. Omitted keeps it.
    name: Union[ChartDataSource, str, None]
    #: The category labels input, a reference or a data source; an empty reference or ``None`` removes it. Omitted keeps
    #: it.
    categories: Union[ChartDataSource, str, None]
    #: The values input, a reference or a data source; an empty reference is rejected, and so is ``None``. Omitted keeps
    #: it.
    values: Union[ChartDataSource, str, None]
    #: The x values input, a reference or a data source; an empty reference or ``None`` removes it. Omitted keeps it.
    x: Union[ChartDataSource, str, None]
    #: The y values input, a reference or a data source; an empty reference is rejected, and so is ``None``. Omitted
    #: keeps it.
    y: Union[ChartDataSource, str, None]
    #: The bubble sizes input, a reference or a data source; an empty reference is rejected, and so is ``None``. Omitted
    #: keeps it.
    size: Union[ChartDataSource, str, None]
    #: The patch applied to the series fill. Omitted keeps it, ``None`` removes it.
    fill: Optional[ChartFillPatch]
    #: The patch applied to the series outline. Omitted keeps it, ``None`` removes it.
    line: Optional[ChartLinePatch]
    #: The patch applied to the series marker. Omitted keeps it, ``None`` removes it.
    marker: Optional[ChartMarkerPatch]
    #: Whether the line is drawn smoothed instead of as straight segments. Omitted keeps it; ``None`` is rejected.
    smooth: Optional[bool]
    #: The patch applied to the data labels. Omitted keeps it, ``None`` removes it.
    labels: Optional[ChartLabelsPatch]


@dataclass
class ChartSeriesInput:
    """A new series for :meth:`Chart.add_series`.

    Every field other than ``values`` and ``y`` is optional and reads as the same key of :class:`ChartSeriesPatch`.
    """

    #: The values input; required on a category plot, rejected on a scatter or bubble plot.
    values: Union[ChartDataSource, str, None] = None
    #: The series name input; omitted adds none.
    name: Union[ChartDataSource, str, None] = None
    #: The category labels input; omitted adds none.
    categories: Union[ChartDataSource, str, None] = None
    #: The x values input of a scatter or bubble series; omitted adds none.
    x: Union[ChartDataSource, str, None] = None
    #: The y values input; required on a scatter or bubble plot, rejected on a category plot.
    y: Union[ChartDataSource, str, None] = None
    #: The bubble sizes input of a bubble series; omitted adds none.
    size: Union[ChartDataSource, str, None] = None
    #: The series fill to set; omitted states none.
    fill: Optional[ChartFillPatch] = None
    #: The series outline to set; omitted states none.
    line: Optional[ChartLinePatch] = None
    #: The series marker to set; omitted states none.
    marker: Optional[ChartMarkerPatch] = None
    #: ``True`` to draw a smoothed line; omitted states none.
    smooth: Optional[bool] = None
    #: The data labels to set; omitted states none.
    labels: Optional[ChartLabelsPatch] = None


class ChartPointPatch(TypedDict, total=False):
    """Data-point patch; an omitted key is kept, so an empty patch changes nothing. ``None`` on a key removes it."""

    #: The patch applied to the point fill. Omitted keeps it, ``None`` removes it.
    fill: Optional[ChartFillPatch]
    #: The patch applied to the point outline. Omitted keeps it, ``None`` removes it.
    line: Optional[ChartLinePatch]
    #: The patch applied to the point marker. Omitted keeps it, ``None`` removes it.
    marker: Optional[ChartMarkerPatch]


class ChartPlotPatch(TypedDict, total=False):
    """Plot-group patch; a field the plot type does not have is rejected."""

    #: The gap between bar clusters as a percentage of the bar width. Omitted keeps it; ``None`` is rejected.
    gap_width_pct: Optional[int]
    #: The overlap between bars of a cluster as a percentage of the bar width. Omitted keeps it; ``None`` is rejected.
    overlap_pct: Optional[int]
    #: The angle in degrees of the first pie or doughnut slice. Omitted keeps it; ``None`` is rejected.
    first_slice_angle_deg: Optional[int]
    #: Whether each point of a single-series plot takes its own colour. Omitted keeps it; ``None`` is rejected.
    vary_colors: Optional[bool]
    #: The doughnut hole size as a percentage of the diameter, 10 to 90; doughnut plots only. Omitted keeps it; ``None``
    #: is rejected.
    hole_size_pct: Optional[int]


class ChartAxisPatch(TypedDict, total=False):
    """Axis patch."""

    #: Whether the axis is shown. Omitted keeps it. ``None`` is rejected.
    visible: Optional[bool]
    #: The axis title, as text or a title of the automatic kind (Excel's own title). ``None`` removes the title. Omitted
    #: keeps it.
    title: Optional[Union[str, ChartTitle]]
    #: The fixed minimum of the axis scale; ``None`` makes it automatic. Omitted keeps it.
    min: Optional[float]
    #: The fixed maximum of the axis scale; ``None`` makes it automatic. Omitted keeps it.
    max: Optional[float]
    #: The fixed distance between major ticks; ``None`` makes it automatic. Omitted keeps it.
    major_unit: Optional[float]
    #: The number format code of the tick labels. ``None`` removes it. Omitted keeps it.
    number_format: Optional[str]
    #: Whether major gridlines are drawn. Omitted keeps it. ``None`` is rejected.
    major_gridlines: Optional[bool]
    #: Whether the axis runs in reverse order. Omitted keeps it. ``None`` is rejected.
    reversed: Optional[bool]
    #: The base of a logarithmic scale; ``None`` makes the axis linear. Omitted keeps it.
    log_base: Optional[float]
    #: Where this axis crosses its cross axis (automatic, minimum, maximum or at ``crosses_at``). Omitted keeps it.
    #: ``None`` is rejected.
    crosses: Optional[ChartAxisCrosses]
    #: The value on the cross axis where this axis crosses; read when ``crosses`` is the at-value kind. Omitted keeps
    #: it. ``None`` is rejected.
    crosses_at: Optional[float]
    #: Where the tick labels sit relative to the axis; ``None`` restores the default. Omitted keeps it.
    tick_label_position: Optional[ChartTickLabelPosition]
    #: The fixed distance between minor ticks; ``None`` makes it automatic. Omitted keeps it.
    minor_unit: Optional[float]
    #: Whether minor gridlines are drawn. Omitted keeps it. ``None`` is rejected.
    minor_gridlines: Optional[bool]
    #: The side of the plot area the axis sits on. Omitted keeps it. ``None`` is rejected.
    position: Optional[ChartAxisPosition]
    #: Whether the tick label number format follows the source cells. Omitted keeps it. ``None`` is rejected.
    source_linked: Optional[bool]


class ChartDisplayPatch(TypedDict, total=False):
    """Display patch for :meth:`Chart.update_display`; an omitted key is kept, ``None`` removes the element."""

    #: How empty cells are plotted: as gaps, as zero or by spanning them. Omitted keeps it, ``None`` removes it.
    blanks_as: Optional[BlanksAs]
    #: Whether only visible cells are plotted. Omitted keeps it, ``None`` removes it.
    plot_visible_only: Optional[bool]
    #: Whether ``#N/A`` points are drawn as blank cells; ``False`` and ``None`` both remove the setting, as Excel does.
    #: Omitted keeps it.
    na_as_blank: Optional[bool]
    #: The chart style number, 1 to 48. Omitted keeps it, ``None`` removes it.
    style: Optional[int]


@dataclass(frozen=True)
class WorkbookProtection:
    """Workbook ``<workbookProtection>``. Hash fields are stored as given; the host computes them.

    A text field the element does not state is ``None``.
    """

    #: ``True`` when the workbook structure is locked, so sheets cannot be added, deleted, renamed, moved or hidden.
    lock_structure: bool = False
    #: ``True`` when the workbook windows are locked, so they cannot be moved or resized.
    lock_windows: bool = False
    #: ``True`` when the workbook is locked for revisions (shared-workbook tracking cannot be switched off).
    lock_revision: bool = False
    #: The legacy 16-bit hash of the structure-lock password, as the hex text Excel writes; ``None`` when absent.
    workbook_legacy_password_hash: Optional[str] = None
    #: The name of the hash algorithm of the structure-lock password; ``None`` when absent.
    workbook_algorithm_name: Optional[str] = None
    #: The hash of the structure-lock password, stored as given; ``None`` when absent.
    workbook_hash_value: Optional[str] = None
    #: The salt of the structure-lock password hash, stored as given; ``None`` when absent.
    workbook_salt_value: Optional[str] = None
    #: The number of hash iterations of the structure-lock password; 0 when absent.
    workbook_spin_count: int = 0
    #: The legacy 16-bit hash of the revision-lock password, as the hex text Excel writes; ``None`` when absent.
    revisions_legacy_password_hash: Optional[str] = None
    #: The name of the hash algorithm of the revision-lock password; ``None`` when absent.
    revisions_algorithm_name: Optional[str] = None
    #: The hash of the revision-lock password, stored as given; ``None`` when absent.
    revisions_hash_value: Optional[str] = None
    #: The salt of the revision-lock password hash, stored as given; ``None`` when absent.
    revisions_salt_value: Optional[str] = None
    #: The number of hash iterations of the revision-lock password; 0 when absent.
    revisions_spin_count: int = 0


@dataclass
class WorkbookProtectionInput:
    """The ``<workbookProtection>`` :meth:`Workbook.set_workbook_protection` writes; every field defaults as in
    :class:`WorkbookProtection`. A :class:`WorkbookProtection` read back is accepted in its place.
    """

    #: ``True`` locks the workbook structure; omitted is ``False``.
    lock_structure: bool = False
    #: ``True`` locks the workbook windows; omitted is ``False``.
    lock_windows: bool = False
    #: ``True`` locks the workbook for revisions; omitted is ``False``.
    lock_revision: bool = False
    #: The legacy 16-bit hash of the structure-lock password, as the hex text Excel writes; omitted or ``None`` is
    #: absent.
    workbook_legacy_password_hash: Optional[str] = None
    #: The name of the hash algorithm of the structure-lock password; omitted or ``None`` is absent.
    workbook_algorithm_name: Optional[str] = None
    #: The hash of the structure-lock password, stored as given; omitted or ``None`` is absent.
    workbook_hash_value: Optional[str] = None
    #: The salt of the structure-lock password hash, stored as given; omitted or ``None`` is absent.
    workbook_salt_value: Optional[str] = None
    #: The number of hash iterations of the structure-lock password; omitted is 0, which is absent.
    workbook_spin_count: int = 0
    #: The legacy 16-bit hash of the revision-lock password, as the hex text Excel writes; omitted or ``None`` is
    #: absent.
    revisions_legacy_password_hash: Optional[str] = None
    #: The name of the hash algorithm of the revision-lock password; omitted or ``None`` is absent.
    revisions_algorithm_name: Optional[str] = None
    #: The hash of the revision-lock password, stored as given; omitted or ``None`` is absent.
    revisions_hash_value: Optional[str] = None
    #: The salt of the revision-lock password hash, stored as given; omitted or ``None`` is absent.
    revisions_salt_value: Optional[str] = None
    #: The number of hash iterations of the revision-lock password; omitted is 0, which is absent.
    revisions_spin_count: int = 0


@dataclass(frozen=True)
class ProtectedRange:
    """A sheet ``<protectedRange>`` covering the rectangles of ``ranges``."""

    #: The range's name, unique among the sheet's protected ranges.
    name: str
    #: The rectangles the protected range covers.
    ranges: Union[Sequence[Union[CellRect, str]], str]
    #: The legacy 16-bit password hash, as the hex text Excel writes; ``None`` when absent.
    legacy_password_hash: Optional[str] = None
    #: The name of the hash algorithm that produced ``hash_value``; ``None`` when absent.
    algorithm_name: Optional[str] = None
    #: The password hash, stored as given; ``None`` when absent.
    hash_value: Optional[str] = None
    #: The salt the password hash was computed with, stored as given; ``None`` when absent.
    salt_value: Optional[str] = None
    #: The number of hash iterations; 0 when absent.
    spin_count: int = 0
    #: The security descriptor (SDDL text) naming who may edit the range without a password; ``None`` when absent.
    security_descriptor: Optional[str] = None


@dataclass
class ProtectedRangeInput:
    """A protected range :meth:`Sheet.add_protected_range` adds. A :class:`ProtectedRange` read back is
    accepted in its place.
    """

    #: The range's name; it must be non-empty and unused by the sheet's other protected ranges.
    name: str
    #: The rectangles the range covers: a sequence of ``CellRect``s and A1 ranges, or one A1 string; at least one is
    #: required.
    ranges: Union[Sequence[Union[CellRect, str]], str]
    #: The legacy 16-bit password hash, as the hex text Excel writes; omitted or ``None`` is absent.
    legacy_password_hash: Optional[str] = None
    #: The name of the hash algorithm that produced ``hash_value``; omitted or ``None`` is absent.
    algorithm_name: Optional[str] = None
    #: The password hash, stored as given; omitted or ``None`` is absent.
    hash_value: Optional[str] = None
    #: The salt the password hash was computed with, stored as given; omitted or ``None`` is absent.
    salt_value: Optional[str] = None
    #: The number of hash iterations; omitted is 0, which is absent.
    spin_count: int = 0
    #: The security descriptor (SDDL text) naming who may edit the range without a password; omitted or ``None`` is
    #: absent.
    security_descriptor: Optional[str] = None


@dataclass(frozen=True)
class WorkbookLocation:
    """Where a workbook is saved (:meth:`Workbook.get_location`).

    The location is metadata only: it is never read from or written to a file
    system. ``directory`` is absolute in one of three canonical forms: POSIX
    (``/a/b``, a backslash being a literal character, ``/`` alone the root),
    drive (``C:/a``, given as ``C:\\a``, ``C:/a`` or ``c:\\a\\``; the drive
    root ``C:/``) or UNC (``//server/share/a``, given as
    ``\\\\server\\share\\a`` or ``//server/share/a``). The trailing
    separator is stripped except at a root.
    """

    #: The folder in canonical form; ``None`` when the workbook has a name but no folder.
    directory: Optional[str]
    #: The file name, extension included.
    file_name: str


@dataclass
class WorkbookLocationInput:
    """The location :meth:`Workbook.set_location` takes; a :class:`WorkbookLocation` is accepted as is."""

    #: The file name, extension included; required.
    file_name: str
    #: The folder, absolute in one of the forms :class:`WorkbookLocation` lists; omitted or ``None`` is no folder.
    directory: Optional[str] = None


@dataclass(frozen=True)
class SheetProtection:
    """The protection of a protected sheet (:meth:`Sheet.get_sheet_protection`).

    Each ``allow_*`` flag permits one action on the protected sheet. The
    defaults are those of Excel's Protect Sheet dialog: selecting locked and
    unlocked cells is allowed, every other action is not. A text field the
    sheet does not state is ``None``.
    """

    #: The name of the hash algorithm that produced ``hash_value``; ``None`` when the sheet states none.
    algorithm_name: Optional[str] = None
    #: The password hash, stored as given; ``None`` when the sheet states none.
    hash_value: Optional[str] = None
    #: The salt the password hash was computed with, stored as given; ``None`` when the sheet states none.
    salt_value: Optional[str] = None
    #: The number of hash iterations; 0 when the sheet states none.
    spin_count: int = 0
    #: The legacy 16-bit password hash, as the hex text Excel writes; ``None`` when the sheet states none.
    legacy_password_hash: Optional[str] = None
    #: ``True`` when users may edit objects (shapes, charts, pictures) on the protected sheet; Excel's default
    #: ``False``.
    allow_edit_objects: bool = False
    #: ``True`` when users may edit scenarios on the protected sheet; Excel's default ``False``.
    allow_edit_scenarios: bool = False
    #: ``True`` when users may format cells on the protected sheet; Excel's default ``False``.
    allow_format_cells: bool = False
    #: ``True`` when users may format columns on the protected sheet; Excel's default ``False``.
    allow_format_columns: bool = False
    #: ``True`` when users may format rows on the protected sheet; Excel's default ``False``.
    allow_format_rows: bool = False
    #: ``True`` when users may insert columns on the protected sheet; Excel's default ``False``.
    allow_insert_columns: bool = False
    #: ``True`` when users may insert rows on the protected sheet; Excel's default ``False``.
    allow_insert_rows: bool = False
    #: ``True`` when users may insert hyperlinks on the protected sheet; Excel's default ``False``.
    allow_insert_hyperlinks: bool = False
    #: ``True`` when users may delete columns on the protected sheet; Excel's default ``False``.
    allow_delete_columns: bool = False
    #: ``True`` when users may delete rows on the protected sheet; Excel's default ``False``.
    allow_delete_rows: bool = False
    #: ``True`` when users may select locked cells on the protected sheet; Excel's default ``True``.
    allow_select_locked_cells: bool = True
    #: ``True`` when users may select unlocked cells on the protected sheet; Excel's default ``True``.
    allow_select_unlocked_cells: bool = True
    #: ``True`` when users may sort on the protected sheet; Excel's default ``False``.
    allow_sort: bool = False
    #: ``True`` when users may use AutoFilter on the protected sheet; Excel's default ``False``.
    allow_auto_filter: bool = False
    #: ``True`` when users may use PivotTables on the protected sheet; Excel's default ``False``.
    allow_pivot_tables: bool = False


@dataclass
class SheetProtectionInput:
    """The protection :meth:`Sheet.set_sheet_protection` applies; the fields and defaults of :class:`SheetProtection`."""

    #: The name of the hash algorithm that produced ``hash_value``; omitted or ``None`` is absent.
    algorithm_name: Optional[str] = None
    #: The password hash, stored as given; omitted or ``None`` is absent.
    hash_value: Optional[str] = None
    #: The salt the password hash was computed with, stored as given; omitted or ``None`` is absent.
    salt_value: Optional[str] = None
    #: The number of hash iterations; omitted is 0, which is absent.
    spin_count: int = 0
    #: The legacy 16-bit password hash, as the hex text Excel writes; omitted or ``None`` is absent.
    legacy_password_hash: Optional[str] = None
    #: ``True`` when users may edit objects (shapes, charts, pictures) on the protected sheet; default ``False``.
    allow_edit_objects: bool = False
    #: ``True`` when users may edit scenarios on the protected sheet; default ``False``.
    allow_edit_scenarios: bool = False
    #: ``True`` when users may format cells on the protected sheet; default ``False``.
    allow_format_cells: bool = False
    #: ``True`` when users may format columns on the protected sheet; default ``False``.
    allow_format_columns: bool = False
    #: ``True`` when users may format rows on the protected sheet; default ``False``.
    allow_format_rows: bool = False
    #: ``True`` when users may insert columns on the protected sheet; default ``False``.
    allow_insert_columns: bool = False
    #: ``True`` when users may insert rows on the protected sheet; default ``False``.
    allow_insert_rows: bool = False
    #: ``True`` when users may insert hyperlinks on the protected sheet; default ``False``.
    allow_insert_hyperlinks: bool = False
    #: ``True`` when users may delete columns on the protected sheet; default ``False``.
    allow_delete_columns: bool = False
    #: ``True`` when users may delete rows on the protected sheet; default ``False``.
    allow_delete_rows: bool = False
    #: ``True`` when users may select locked cells on the protected sheet; default ``True``.
    allow_select_locked_cells: bool = True
    #: ``True`` when users may select unlocked cells on the protected sheet; default ``True``.
    allow_select_unlocked_cells: bool = True
    #: ``True`` when users may sort on the protected sheet; default ``False``.
    allow_sort: bool = False
    #: ``True`` when users may use AutoFilter on the protected sheet; default ``False``.
    allow_auto_filter: bool = False
    #: ``True`` when users may use PivotTables on the protected sheet; default ``False``.
    allow_pivot_tables: bool = False


@dataclass(frozen=True)
class SheetView:
    """Per-sheet view: zoom, frozen-pane counts, and the display / orientation flags mirrored from OOXML
    ``<sheetView>``. A sheet's visibility is :meth:`Sheet.get_visibility`.
    """

    #: The zoom as a percentage, 10 to 400; 100 is 100%.
    zoom_pct: int
    #: The number of frozen rows at the top; 0 means no frozen rows.
    freeze_rows: int
    #: The number of frozen columns at the left; 0 means no frozen columns.
    freeze_cols: int
    #: ``True`` when the cell gridlines are shown; the default is ``True``.
    show_grid_lines: bool = True
    #: ``True`` when the row and column headers are shown; the default is ``True``.
    show_row_column_headers: bool = True
    #: ``True`` when cells holding zero show the 0; the default is ``True``.
    show_zeros: bool = True
    #: ``True`` when the sheet is laid out right to left; the default is ``False``.
    right_to_left: bool = False
    #: ``True`` when the sheet's tab is selected: ``False`` when a loaded file leaves it unset and on an added sheet,
    #: ``True`` on the first sheet of a new workbook. Plain metadata: it does not by itself make the sheet the active one.
    tab_selected: bool = False
    #: The view mode (normal, page break preview or page layout); the normal view when the attribute is absent.
    view_mode: SheetViewMode = SheetViewMode.NORMAL
    #: The scroll origin of the sheet view (the first visible cell, ``topLeftCell``); A1 by default.
    top_left: CellAddress = field(default_factory=lambda: CellAddress(0, 0))
    #: The first visible cell of a frozen or split view's lower-right pane (``<pane topLeftCell>``): the first cell past
    #: the frozen rows and columns until a host scrolls it. A1 without panes.
    pane_top_left: CellAddress = field(default_factory=lambda: CellAddress(0, 0))
    #: The horizontal position of an unfrozen split in twips; 0 means no split on that axis, and it is always 0 while
    #: the view is frozen.
    split_x_twips: float = 0.0
    #: The vertical position of an unfrozen split in twips; 0 means no split on that axis, and it is always 0 while the
    #: view is frozen.
    split_y_twips: float = 0.0
    #: ``True`` when cells show their formula text instead of the value; the default is ``False``.
    show_formulas: bool = False
    #: ``True`` when the outline gutter's group symbols are shown; the default is ``True``.
    show_outline_symbols: bool = True


@dataclass(frozen=True)
class ColumnLayout:
    """A per-column-range layout override (inclusive ``[first, last]``).

    The column XF is read through :meth:`SheetOoxml.get_column_xf_index`.
    """

    #: 0-based index of the first column of the span.
    first: int
    #: 0-based index of the last column of the span, inclusive.
    last: int
    #: The explicit width in characters of the Normal font (raw presence or a non-zero legacy width); ``None`` when the
    #: span states none, and an explicit zero stays distinguishable from an absent width.
    width_chars: Optional[float]
    #: ``True`` when the columns of the span are hidden.
    hidden: bool
    #: The outline level of the span; 0 means not grouped.
    outline_level: int
    #: ``True`` when the outline group this span summarises is collapsed (the ``collapsed`` attribute).
    collapsed: bool = False


@dataclass(frozen=True)
class RowLayout:
    """A per-row layout override; the row XF is read through :meth:`SheetOoxml.get_row_xf_index`."""

    #: 0-based index of the row.
    row: int
    #: The stored height in points; ``None`` when the row stores none.
    height_pt: Optional[float]
    #: ``True`` when the row is hidden.
    hidden: bool
    #: The outline level of the row; 0 means not grouped.
    outline_level: int
    #: ``True`` when the height is an explicit override (``customHeight="1"``) rather than an auto height that Excel
    #: recomputes on its own.
    custom_height: bool = False
    #: ``True`` when the outline group this row summarises is collapsed (the ``collapsed`` attribute).
    collapsed: bool = False


@dataclass(frozen=True)
class IterativeSettings:
    """Iterative-calculation settings read back by ``get_iterative``."""

    #: ``True`` when iterative calculation is on; off by default, when cyclic components surface ``#REF!``.
    enabled: bool
    #: The iteration cap, 1 to 32767; meaningful even while ``enabled`` is ``False``.
    max_iterations: int
    #: The absolute convergence threshold, finite and non-negative; meaningful even while ``enabled`` is ``False``.
    max_change: float


@dataclass(frozen=True)
class SheetSelection:
    """The active pane's selection read back by ``get_selection``."""

    #: The active cell.
    active: CellAddress
    #: The selected ranges.
    ranges: List[CellRect] = field(default_factory=list)


@dataclass
class SheetSelectionInput:
    """A selection :meth:`Sheet.set_selection` writes. A :class:`SheetSelection` read back is accepted
    in its place.
    """

    #: The active cell, as a cell address or an A1 string.
    active: Union[CellAddress, str]
    #: The selected ranges, as rectangles or A1 strings, or one A1 string; empty or omitted selects the active cell
    #: alone.
    ranges: Union[Sequence[Union[CellRect, str]], str] = ()


@dataclass(frozen=True)
class TableStyleDefaults:
    """Default table and pivot style names read back by ``get_table_style_defaults``; ``None`` when unset."""

    #: The style a new table takes when none is given; ``None`` when unset.
    table_style_name: Optional[str]
    #: The style a new PivotTable takes; ``None`` when unset.
    pivot_style_name: Optional[str]


@dataclass(frozen=True)
class PivotGrandTotals:
    """Grand-total flags read back by :meth:`Pivot.get_grand_totals`."""

    #: Whether the grand total band for the rows shows.
    rows: bool
    #: Whether the grand total band for the columns shows.
    columns: bool


@dataclass(frozen=True)
class Pagination:
    """Resolved physical pagination for one worksheet.

    The break lists hold the breaks in effect, ascending, each before the zero-based row or column a new
    physical page begins with.
    """

    #: The number of physical pages.
    page_count: int
    #: The sheet's declared ``_xlnm.Print_Area``; empty when the sheet declares none, and not backfilled with the used
    #: range. Pagination still falls back to the used range in that case, so ``page_count`` can be non-zero while this
    #: is empty.
    print_area: List[CellRect]
    #: The row breaks in effect, ascending by ``row``.
    row_breaks: List["RowBreak"]
    #: The column breaks in effect, ascending by ``col``.
    column_breaks: List["ColumnBreak"]
    #: The paper the layout ran on, after the orientation swap. Its ``known`` is false when the stored paper size was
    #: unrecognised and A4 was substituted.
    paper: "PaginationPaper"
    #: The page margins.
    margins: "PaginationMargins"
    #: The cell-body rectangle before print-title reservation, measured from the page's top-left corner.
    printable: "RectPt"
    #: The effective print scale as a percentage (100 is 100%).
    scale_pct: float
    #: The order in which pages are numbered over the print area, a ``PageOrder`` value.
    page_order: PageOrder
    #: The repeat rows and columns in effect; ``None`` when neither axis repeats.
    print_titles: Optional["PrintTitles"]
    #: The physical pages in print order.
    pages: List["PaginationPage"]

    def __post_init__(self) -> None:
        object.__setattr__(self, "page_order", _enum(PageOrder, self.page_order))


@dataclass(frozen=True)
class RectPt:
    """A rectangle in points."""

    #: The distance in points from the rectangle's origin to its left edge.
    x: float
    #: The distance in points from the rectangle's origin to its top edge.
    y: float
    #: The width in points.
    width_pt: float
    #: The height in points.
    height_pt: float


@dataclass(frozen=True)
class PaginationPaper:
    """The paper a sheet prints on, after the orientation swap."""

    #: The paper width in points, after the orientation swap.
    width_pt: float
    #: The paper height in points, after the orientation swap.
    height_pt: float
    #: ``True`` when the paper is laid out landscape.
    landscape: bool
    #: ``False`` when the stored paper size was unrecognised and A4 was substituted.
    known: bool


@dataclass(frozen=True)
class PaginationMargins:
    """Page margins in inches."""

    #: The left page margin in inches.
    left_in: float
    #: The right page margin in inches.
    right_in: float
    #: The top page margin in inches.
    top_in: float
    #: The bottom page margin in inches.
    bottom_in: float
    #: The distance in inches from the page top edge to the header band.
    header_in: float
    #: The distance in inches from the page bottom edge to the footer band.
    footer_in: float


@dataclass(frozen=True)
class Span:
    """A 0-based, inclusive run of rows or columns."""

    #: The 0-based first row or column.
    first: int
    #: The 0-based last row or column, inclusive.
    last: int


@dataclass(frozen=True)
class PrintTitles:
    """The rows and columns repeated on every printed page; ``None`` for an axis that repeats nothing."""

    #: The rows repeated at the top of every printed page; ``None`` when no rows repeat.
    rows: Optional[Span]
    #: The columns repeated at the left of every printed page; ``None`` when no columns repeat.
    columns: Optional[Span]


@dataclass
class PrintTitlesInput:
    """The print titles :meth:`Sheet.set_print_titles` writes. A :class:`PrintTitles` read back is accepted
    in its place.
    """

    #: The rows to repeat: a span, a row span string (such as ``'3'`` or ``'1:2'``), or ``None`` for no repeat.
    rows: Union[Span, str, None]
    #: The columns to repeat: a span, a column span string (such as ``'B'`` or ``'A:B'``), or ``None`` for no repeat.
    columns: Union[Span, str, None]


@dataclass(frozen=True)
class PaginationPage:
    """One physical page: the cell block it carries and where it sits."""

    #: 0-based index of the print area this page belongs to (0 for the used-range fallback).
    area_index: int
    #: The cell block the page carries.
    range: CellRect
    #: Horizontal position in points of the cell block's left edge on the paper.
    origin_x_pt: float
    #: Vertical position in points of the cell block's top edge on the paper.
    origin_y_pt: float
    #: Scaled width in points of the page's cell block, repeated title columns excluded.
    width_pt: float
    #: Scaled height in points of the page's cell block, repeated title rows excluded.
    height_pt: float


@dataclass(frozen=True)
class SheetFormatDefaults:
    """Sheet default column / row metrics (``<sheetFormatPr>``).

    ``Sheet.update_format_defaults`` takes these fields as keyword arguments: ``None`` clears the stored value
    and :data:`UNCHANGED` (the default) keeps it. The ``effective_*`` fields are ignored when setting.
    """

    #: The stored default column width in characters, column padding included; ``None`` when the sheet stores none.
    default_column_width_chars: Optional[float] = None
    #: The stored default row height in points; ``None`` when the sheet stores none.
    default_row_height_pt: Optional[float] = None
    #: The base column width in characters, from which Excel derives the default column width when none is stored;
    #: always present, the OOXML default is 8.
    base_column_width_chars: float = 8.0
    #: Stored width in characters the sheet lays columns out with: the stated default column width when set, else the
    #: width Excel derives from the base column width and the Normal font (Windows, Calibri 11, base 8: 9.140625, shown
    #: as 8.43; Mac, Calibri 11, base 10: 10.83203125, shown as 10). Under a Mac profile a sheet that stores no default
    #: row height (no ``<sheetFormatPr>``) derives it from base 10, although ``base_column_width_chars`` still reads 8.
    #: Read-only.
    effective_default_column_width_chars: float = 0.0
    #: Points the sheet lays rows out with: the stated default row height when set, else 102/7 (about 14.57); 0 when the
    #: sheet hides rows by default. Read-only.
    effective_default_row_height_pt: float = 0.0
    #: ``True`` when ``default_row_height_pt`` is an explicit override that stays as written, rather than a cached
    #: figure Excel recomputes from the Normal font; ``None`` keeps the stored flag when setting.
    custom_height: Optional[bool] = None
    #: ``True`` when rows are hidden by default unless they say otherwise; ``None`` keeps the stored flag when setting.
    zero_height: Optional[bool] = None


@dataclass(frozen=True)
class WidthModel:
    """Column-width conversion model for the workbook's Normal font.

    A width of ``chars`` characters is ``chars * points_per_char + padding_pt``; stored widths go through
    ``convert_column_chars_to_pt``, which snaps to whole screen pixels (whole points under a Mac profile, where a pixel
    is a point).
    """

    #: The points added per displayed character.
    points_per_char: float
    #: The constant padding in points.
    padding_pt: float
    #: The point size of the Normal font.
    normal_font_size_pt: float
    #: ``False`` when the Normal font is not a measured column-width calibration point of the platform and that
    #: platform's Calibri 11 figure was extrapolated.
    calibrated: bool
    #: The name of the Normal font the metrics were looked up for. Under a Mac profile a Normal font that follows the
    #: theme reports the theme's font.
    normal_font_name: str
    #: The ``WidthPlatform`` whose metrics the model uses. ``PRINT`` mode always reports Windows.
    platform: WidthPlatform

    def __post_init__(self) -> None:
        object.__setattr__(self, "platform", _enum(WidthPlatform, self.platform))


@dataclass(frozen=True)
class RowBreak:
    """One row break (``<rowBreaks><brk>``).

    The break covers the columns ``first..last`` (OOXML ``min`` / ``max``). As the argument of
    :meth:`Sheet.set_row_break`, ``row`` is ignored and ``first == last == 0``
    covers the whole row.
    """

    #: 0-based row the break sits before.
    row: int
    #: The first column the break covers (OOXML ``min``).
    first: int = 0
    #: The last column the break covers (OOXML ``max``).
    last: int = 0
    #: ``True`` for a break a user placed, ``False`` for one Excel computed.
    manual: bool = True


@dataclass
class PageBreakInput:
    """The break :meth:`Sheet.set_row_break` / :meth:`Sheet.set_column_break` write.

    It covers the cross-axis span ``first..last``; ``first == last == 0`` covers the whole row or column.
    """

    #: The first row or column the break covers on the perpendicular axis; 0 together with ``last`` covers the whole row
    #: or column.
    first: int = 0
    #: The last row or column the break covers on the perpendicular axis; 0 together with ``first`` covers the whole row
    #: or column.
    last: int = 0
    #: ``True`` for a break a user placed, ``False`` for one Excel would compute; defaults to ``True``.
    manual: bool = True


@dataclass(frozen=True)
class ColumnBreak:
    """One column break (``<colBreaks><brk>``), covering the rows ``first..last``.

    ``manual`` and the argument form of :meth:`Sheet.set_column_break` are as in :class:`RowBreak`.
    """

    #: 0-based column the break sits before.
    col: int
    #: The first row the break covers (OOXML ``min``).
    first: int = 0
    #: The last row the break covers (OOXML ``max``).
    last: int = 0
    #: ``True`` for a break a user placed, ``False`` for one Excel computed.
    manual: bool = True


@dataclass(frozen=True)
class PageSetup:
    """Effective ``<pageSetup>`` values, with per-attribute presence.

    Every value field carries the setting in force. The ``*_stated`` flags
    say whether the XML actually declares it, which is the only way to tell
    ``scale="100"`` from an absent attribute that defaults to 100.

    ``fit_to_page`` lives in ``<sheetPr><pageSetUpPr>`` rather than
    ``<pageSetup>``; it is reported here because callers reason about it as
    part of one page-setup decision.
    """

    #: The page orientation; ``None`` when the sheet states none (Excel prints portrait).
    orientation: Optional[PageOrientation]
    #: The OOXML paperSize code in force; an open set, so unknown codes round-trip.
    paper_size: int
    #: The print scale as a percentage, 10 to 400; 100 when the attribute is absent.
    scale_pct: int
    #: The number of pages across the fit-to-page mode targets.
    fit_to_width: int
    #: The number of pages down the fit-to-page mode targets.
    fit_to_height: int
    #: ``True`` when the fit-to-page mode is selected; read from ``<sheetPr><pageSetUpPr>``, not ``<pageSetup>``.
    fit_to_page: bool
    #: ``True`` when the XML states ``paper_size``; ``False`` means it reads a default.
    paper_size_stated: bool
    #: ``True`` when the XML states ``scale_pct``; ``False`` means it reads a default.
    scale_pct_stated: bool
    #: ``True`` when the XML states ``fit_to_width``; ``False`` means it reads a default.
    fit_to_width_stated: bool
    #: ``True`` when the XML states ``fit_to_height``; ``False`` means it reads a default.
    fit_to_height_stated: bool
    #: ``True`` when the XML states ``fit_to_page``; ``False`` means it reads a default.
    fit_to_page_stated: bool


@dataclass(frozen=True)
class PageMargins:
    """Effective ``<pageMargins>`` values in inches, with presence flags."""

    #: The left margin in inches.
    left_in: float
    #: The right margin in inches.
    right_in: float
    #: The top margin in inches.
    top_in: float
    #: The bottom margin in inches.
    bottom_in: float
    #: The distance in inches from the top edge of the paper to the header.
    header_in: float
    #: The distance in inches from the bottom edge of the paper to the footer.
    footer_in: float
    #: ``True`` when the XML states ``left_in``; ``False`` means it reads a default.
    left_in_stated: bool
    #: ``True`` when the XML states ``right_in``; ``False`` means it reads a default.
    right_in_stated: bool
    #: ``True`` when the XML states ``top_in``; ``False`` means it reads a default.
    top_in_stated: bool
    #: ``True`` when the XML states ``bottom_in``; ``False`` means it reads a default.
    bottom_in_stated: bool
    #: ``True`` when the XML states ``header_in``; ``False`` means it reads a default.
    header_in_stated: bool
    #: ``True`` when the XML states ``footer_in``; ``False`` means it reads a default.
    footer_in_stated: bool


@dataclass(frozen=True)
class PrintOptions:
    """Effective ``<printOptions>`` values; an attribute the sheet does not state reads false with its flag false."""

    #: ``True`` when the cell gridlines are printed.
    grid_lines: bool
    #: ``True`` when the row and column headings are printed.
    headings: bool
    #: ``True`` when the print area is centered horizontally on the page.
    horizontal_centered: bool
    #: ``True`` when the print area is centered vertically on the page.
    vertical_centered: bool
    #: ``True`` when the XML states ``grid_lines``; ``False`` means it reads a default.
    grid_lines_stated: bool
    #: ``True`` when the XML states ``headings``; ``False`` means it reads a default.
    headings_stated: bool
    #: ``True`` when the XML states ``horizontal_centered``; ``False`` means it reads a default.
    horizontal_centered_stated: bool
    #: ``True`` when the XML states ``vertical_centered``; ``False`` means it reads a default.
    vertical_centered_stated: bool


@dataclass(frozen=True)
class HeaderFooter:
    """Effective ``<headerFooter>`` values.

    A section the sheet does not have is ``None``; the text is decoded as
    :meth:`Sheet.update_header_footer` takes it. A flag the sheet does not state
    reads its schema default with its ``*_stated`` flag false.
    """

    #: The decoded text of the header of odd pages (every page unless ``different_odd_even`` or ``different_first``
    #: apply); ``None`` when the sheet has no such section.
    odd_header: Optional[str]
    #: The decoded text of the footer of odd pages (every page unless ``different_odd_even`` or ``different_first``
    #: apply); ``None`` when the sheet has no such section.
    odd_footer: Optional[str]
    #: The decoded text of the header of even pages, used when ``different_odd_even`` is set; ``None`` when the sheet
    #: has no such section.
    even_header: Optional[str]
    #: The decoded text of the footer of even pages, used when ``different_odd_even`` is set; ``None`` when the sheet
    #: has no such section.
    even_footer: Optional[str]
    #: The decoded text of the header of the first page, used when ``different_first`` is set; ``None`` when the sheet
    #: has no such section.
    first_header: Optional[str]
    #: The decoded text of the footer of the first page, used when ``different_first`` is set; ``None`` when the sheet
    #: has no such section.
    first_footer: Optional[str]
    #: ``True`` when odd and even pages use different headers and footers; reads ``False`` when unstated.
    different_odd_even: bool
    #: ``True`` when the first page uses its own header and footer; reads ``False`` when unstated.
    different_first: bool
    #: ``True`` when the header and footer scale with the document's print scale; reads ``True`` when unstated.
    scale_with_doc: bool
    #: ``True`` when the header and footer align with the page margins; reads ``True`` when unstated.
    align_with_margins: bool
    #: ``True`` when the XML states ``different_odd_even``; ``False`` means it reads a default.
    different_odd_even_stated: bool
    #: ``True`` when the XML states ``different_first``; ``False`` means it reads a default.
    different_first_stated: bool
    #: ``True`` when the XML states ``scale_with_doc``; ``False`` means it reads a default.
    scale_with_doc_stated: bool
    #: ``True`` when the XML states ``align_with_margins``; ``False`` means it reads a default.
    align_with_margins_stated: bool


@dataclass(frozen=True)
class CfMatchBar:
    """The data bar of a data-bar :class:`CfMatch`."""

    #: Length of the bar as a percentage (0 to 100) of the cell width.
    length_pct: float
    #: Position of the axis as a percentage of the cell width, measured from the left edge: 0 for an axis at the left
    #: edge, 50 for a centred axis.
    axis_position_pct: float
    #: ``True`` when the bar draws in the negative direction.
    is_negative: bool
    #: The fill colour of the bar, resolved to RGB.
    fill: Color
    #: The border colour, resolved to RGB; ``None`` when the bar draws no border.
    border: Optional[Color]
    #: ``True`` when the bar is filled with a gradient.
    gradient: bool
    #: The direction of the rule; the geometry is stated from the left edge and a host mirrors a right-to-left bar.
    direction: CfDataBarDirection = CfDataBarDirection.CONTEXT


@dataclass(frozen=True)
class CfMatchIcon:
    """The icon of an icon-set :class:`CfMatch`: its set and the 0-based icon index within it."""

    #: The icon set the icon comes from.
    icon_set: IconSetName
    #: The 0-based index of the icon within the set, after the rule's reverse setting has been applied.
    icon_index: int

    def __post_init__(self) -> None:
        object.__setattr__(self, "icon_set", _enum(IconSetName, self.icon_set))


@dataclass(frozen=True)
class CfMatch:
    """A resolved conditional-format match for one cell."""

    #: The kind of match; selects which of the other members carry data.
    kind: CfMatchKind
    #: The workbook-global priority of the matching rule; a smaller number ranks higher, and matches of a cell come in
    #: ascending order.
    priority: int
    #: The 0-based index of the differential format the rule applies; ``None`` when the rule applies no format. Set for
    #: a differential-format match.
    dxf_index: Optional[int] = None
    #: The interpolated fill colour of a colour-scale match, resolved to RGB; ``None`` for other kinds.
    color: Optional[Color] = None
    #: The data bar of a data-bar match; ``None`` for other kinds.
    bar: Optional[CfMatchBar] = None
    #: The icon of an icon-set match; ``None`` for other kinds.
    icon: Optional[CfMatchIcon] = None


@dataclass(frozen=True)
class CfCell:
    """One cell's CF matches inside a viewport-range evaluation."""

    #: The 0-based row of the cell.
    row: int
    #: The 0-based column of the cell.
    col: int
    #: The rules that match the cell, in ascending priority order, so the first entry is the highest priority.
    matches: List[CfMatch]


@dataclass(frozen=True)
class CfValueObject:
    """One conditional-format threshold."""

    #: The kind of threshold, which decides how ``value`` is read (a number, a percentage, a percentile, a formula, or
    #: the minimum or maximum of the data).
    type: CfValueObjectType
    #: The number, percentage or formula the ``type`` reads, with a leading ``=`` on a formula; ``None`` for a type that
    #: reads none (minimum, maximum).
    value: Optional[str] = None
    #: Whether the threshold is inclusive: ``True`` compares with ``>=``, ``False`` with ``>``.
    gte: bool = True

    def __post_init__(self) -> None:
        object.__setattr__(self, "type", _enum(CfValueObjectType, self.type))


@dataclass(frozen=True)
class CfColorScale:
    """``<colorScale>`` payload of a conditional-format rule as read: one colour per threshold."""

    #: The two or three thresholds of the scale, from lowest to highest.
    thresholds: List[CfValueObject]
    #: The colour at each threshold, in the same order and count as ``thresholds``.
    colors: List[Color]


@dataclass
class CfColorScaleInput:
    """``<colorScale>`` payload of a rule to add: two or three thresholds with as many colours."""

    #: The two or three thresholds of the scale, from lowest to highest.
    thresholds: List[CfValueObject]
    #: The colour at each threshold, in the same order and count as ``thresholds``.
    colors: List[ColorInput]


@dataclass(frozen=True)
class CfDataBar:
    """``<dataBar>`` payload of a conditional-format rule as read.

    The fields after ``max_length_pct`` live in the ``x14`` extension rather than the legacy element;
    ``None`` is a setting the rule does not state.
    """

    #: The threshold at which the bar is shortest.
    min: CfValueObject
    #: The threshold at which the bar is longest.
    max: CfValueObject
    #: The fill colour of the bar.
    fill: Color
    #: ``True`` (the default) shows the cell value beside the bar; ``False`` hides it.
    show_value: bool = True
    #: The shortest bar length as a percentage of the cell width; defaults to 10.
    min_length_pct: int = 10
    #: The longest bar length as a percentage of the cell width; defaults to 90.
    max_length_pct: int = 90
    #: ``True`` fills the bar with a gradient, ``False`` with a solid colour; ``None`` when the rule does not state it,
    #: which means a gradient.
    gradient: Optional[bool] = None
    #: Where the axis between negative and positive bars sits; ``None`` when the rule does not state it, which means
    #: automatic.
    axis_position: Optional[CfDataBarAxisPosition] = None
    #: The fill colour of negative bars; ``None`` when the rule does not state it, which means the same as ``fill``.
    negative_fill: Optional[Color] = None
    #: The border colour of positive bars; ``None`` when the rule states none, which means no border.
    border: Optional[Color] = None
    #: The border colour of negative bars; ``None`` when the rule states none, which means no negative border.
    negative_border: Optional[Color] = None
    #: The colour of the axis line; ``None`` when the rule states none, which means black.
    axis_color: Optional[Color] = None
    #: The edge the bar grows from.
    direction: CfDataBarDirection = CfDataBarDirection.CONTEXT


@dataclass
class CfDataBarInput:
    """``<dataBar>`` payload of a rule to add."""

    #: The threshold at which the bar is shortest.
    min: CfValueObject
    #: The threshold at which the bar is longest.
    max: CfValueObject
    #: The fill colour of the bar.
    fill: ColorInput
    #: ``True`` (the default) shows the cell value beside the bar; ``False`` hides it.
    show_value: bool = True
    #: The shortest bar length as a percentage of the cell width; defaults to 10.
    min_length_pct: int = 10
    #: The longest bar length as a percentage of the cell width; defaults to 90.
    max_length_pct: int = 90
    #: ``True`` fills the bar with a gradient, ``False`` with a solid colour; omitted or ``None`` keeps the default, a
    #: gradient.
    gradient: Optional[bool] = None
    #: Where the axis between negative and positive bars sits; omitted or ``None`` keeps the default, automatic.
    axis_position: Optional[CfDataBarAxisPosition] = None
    #: The fill colour of negative bars; omitted or ``None`` keeps the default, the same as ``fill``.
    negative_fill: Optional[ColorInput] = None
    #: The border colour of positive bars; omitted or ``None`` keeps the default, no border.
    border: Optional[ColorInput] = None
    #: The border colour of negative bars; omitted or ``None`` keeps the default, no negative border.
    negative_border: Optional[ColorInput] = None
    #: The colour of the axis line; omitted or ``None`` keeps the default, black.
    axis_color: Optional[ColorInput] = None
    #: The edge the bar grows from; defaults to the context direction (the sheet's reading direction).
    direction: CfDataBarDirection = CfDataBarDirection.CONTEXT


@dataclass(frozen=True)
class CfIconSet:
    """``<iconSet>`` payload of a conditional-format rule as read."""

    #: The icon set the rule draws from.
    name: IconSetName
    #: The thresholds between the icons: N-1 for an icon set of N icons; the lowest bound is ``floor``, not a threshold.
    #: A formula value carries its leading ``=`` on read.
    thresholds: List[CfValueObject]
    #: ``True`` draws the icons in reverse order; defaults to ``False``.
    reverse: bool = False
    #: ``True`` (the default) shows the cell value beside the icon; ``False`` hides it.
    show_value: bool = True
    #: Round-trip only: preserved across load and save but never consulted during evaluation, since each threshold's own
    #: type decides how it is read. Defaults to ``True``.
    percent: bool = True
    #: The lower bound of the lowest icon's bucket; a cell below it gets no icon. ``None`` when the rule states none.
    floor: Optional[CfValueObject] = None
    #: For a custom icon set, the icon drawn for each level; empty or ``None`` when the set draws its own icons.
    icons: List["CfIcon"] = field(default_factory=list)

    def __post_init__(self) -> None:
        object.__setattr__(self, "name", _enum(IconSetName, self.name))


@dataclass
class CfIconSetInput:
    """``<iconSet>`` payload of a rule to add."""

    #: The icon set the rule draws from.
    name: IconSetName
    #: The thresholds between the icons: exactly N-1 for an icon set of N icons (the lowest bound is ``floor``, not a
    #: threshold); any other count is rejected as an invalid argument. A formula value takes a leading ``=`` with or
    #: without it.
    thresholds: List[CfValueObject]
    #: ``True`` draws the icons in reverse order; defaults to ``False``.
    reverse: bool = False
    #: ``True`` (the default) shows the cell value beside the icon; ``False`` hides it.
    show_value: bool = True
    #: Round-trip only: preserved across load and save but never consulted during evaluation, since each threshold's own
    #: type decides how it is read. Defaults to ``True``.
    percent: bool = True
    #: The lower bound of the lowest icon's bucket; a cell below it gets no icon. Omitted or ``None`` keeps Excel's
    #: default, a percent threshold of 0.
    floor: Optional[CfValueObject] = None
    #: For a custom icon set, one icon per level (the set's icon count); omitted, empty or ``None`` uses the set's own
    #: icons. Evaluation still buckets against ``name``.
    icons: List["CfIcon"] = field(default_factory=list)

    def __post_init__(self) -> None:
        self.name = _enum(IconSetName, self.name)


@dataclass(frozen=True)
class CfIcon:
    """One level of a custom icon set."""

    #: The icon set the icon comes from; the no-icons member for a level that draws nothing.
    icon_set: IconSetName
    #: The 0-based index of the icon within ``icon_set``.
    icon_index: int = 0

    def __post_init__(self) -> None:
        object.__setattr__(self, "icon_set", _enum(IconSetName, self.icon_set))


@dataclass(frozen=True)
class ConditionalFormat:
    """A conditional-format rule (getter result, flattened priority order)."""

    #: The rule's identifier, unique within the workbook.
    id: str
    #: The kind of rule; selects which of the other members apply.
    type: CfRuleType
    #: The workbook-global priority of the rule; a smaller number ranks higher.
    priority: int
    #: ``True`` when lower-priority rules are not applied to a cell this rule matches.
    stop_if_true: bool
    #: The cell ranges the rule applies to.
    ranges: List[CellRect]
    #: The 0-based index of the differential format the rule applies; ``None`` when the rule applies no format.
    dxf_index: Optional[int]
    #: The first formula of the rule, leading ``=`` included; ``None`` when the rule has none. Used by expression and
    #: cell-is rules, and stored by the text and blank/error/duplicate kinds.
    formula1: Optional[str]
    #: The second formula of the rule, leading ``=`` included, for the between and not-between operators; ``None``
    #: otherwise.
    formula2: Optional[str]
    #: The comparison operator; ``None`` unless the rule is a cell-is rule.
    operator: Optional[ComparisonOperator]
    #: The number of values (or percent of values with ``percent``) a top/bottom rule selects; ``None`` unless the rule
    #: is a top/bottom rule.
    rank: Optional[int]
    #: ``True`` when ``rank`` is a percentage of the values rather than a count; ``None`` unless the rule is a
    #: top/bottom rule.
    percent: bool
    #: ``True`` when a top/bottom rule selects the lowest values; ``None`` unless the rule is a top/bottom rule.
    bottom: bool
    #: ``True`` for values above the average, ``False`` for below; ``None`` unless the rule is an above-average rule.
    above_average: bool
    #: ``True`` when values equal to the average also match; ``None`` unless the rule is an above-average rule.
    equal_average: bool
    #: The number of standard deviations from the average the rule requires; ``None`` for a plain comparison with the
    #: average, and unless the rule is an above-average rule.
    std_dev: Optional[float]
    #: The text a text rule tests for; ``None`` unless the rule is a contains, not-contains, begins-with or ends-with
    #: rule.
    text: Optional[str]
    #: The date period a time-period rule matches; ``None`` unless the rule is a time-period rule.
    time_period: Optional[DatePeriod]
    #: The colour-scale payload; ``None`` unless the rule is a colour-scale rule.
    color_scale: Optional[CfColorScale]
    #: The data-bar payload; ``None`` unless the rule is a data-bar rule.
    data_bar: Optional[CfDataBar]
    #: The icon-set payload; ``None`` unless the rule is an icon-set rule.
    icon_set: Optional[CfIconSet]

    def __post_init__(self) -> None:
        object.__setattr__(self, "type", _enum(CfRuleType, self.type))
        if self.operator is not None:
            object.__setattr__(self, "operator", _enum(ComparisonOperator, self.operator))
        if self.time_period is not None:
            object.__setattr__(self, "time_period", _enum(DatePeriod, self.time_period))


@dataclass(kw_only=True)
class _CfRuleInput:
    """The members every rule input shares."""

    #: Rectangles, unqualified A1 ranges, or one A1 string.
    ranges: Union[List[Union[CellRect, str]], str]
    #: The rule's priority; zero or less is one past the highest priority already on the sheet.
    priority: int = 0
    #: ``True`` stops the evaluation of lower-priority rules on a cell once this rule matches; defaults to ``False``.
    stop_if_true: bool = False
    #: A GUID in braces, hex digits in either case; omitted, empty or already used on the sheet, the engine generates an
    #: upper-case one, and anything else is invalid.
    id: str = ""


@dataclass(kw_only=True)
class _CfFormatRuleInput(_CfRuleInput):
    """A rule that applies a differential format."""

    #: The index into the workbook's differential formats of the format to apply; ``None`` applies no format.
    dxf_index: Optional[int] = None


@dataclass(kw_only=True)
class CfExpressionRuleInput(_CfFormatRuleInput):
    """``EXPRESSION``: formats a cell where ``formula1`` (leading ``=`` optional) is true."""

    #: The formula that decides whether the cell is formatted; a leading ``=`` is optional.
    formula1: str
    #: Discriminant: always the expression rule kind.
    type: CfRuleType = field(default=CfRuleType.EXPRESSION, init=False)


@dataclass(kw_only=True)
class CfCellIsRuleInput(_CfFormatRuleInput):
    """``CELL_IS``: compares the cell with ``formula1``, and ``formula2`` for ``BETWEEN`` / ``NOT_BETWEEN``."""

    #: How the cell is compared with ``formula1`` (and ``formula2`` for between and not-between).
    operator: ComparisonOperator
    #: The value or formula the cell is compared with; a leading ``=`` is optional.
    formula1: str
    #: The upper bound for the between and not-between operators; a leading ``=`` is optional. Omitted for other
    #: operators.
    formula2: Optional[str] = None
    #: Discriminant: always the cell-is rule kind.
    type: CfRuleType = field(default=CfRuleType.CELL_IS, init=False)


@dataclass(kw_only=True)
class CfColorScaleRuleInput(_CfRuleInput):
    """``COLOR_SCALE``: two or three thresholds with as many colours."""

    #: The colour-scale payload.
    color_scale: CfColorScaleInput
    #: Discriminant: always the colour-scale rule kind.
    type: CfRuleType = field(default=CfRuleType.COLOR_SCALE, init=False)


@dataclass(kw_only=True)
class CfDataBarRuleInput(_CfRuleInput):
    """``DATA_BAR``."""

    #: The data-bar payload.
    data_bar: CfDataBarInput
    #: Discriminant: always the data-bar rule kind.
    type: CfRuleType = field(default=CfRuleType.DATA_BAR, init=False)


@dataclass(kw_only=True)
class CfIconSetRuleInput(_CfRuleInput):
    """``ICON_SET``: the cell's bucket among the thresholds picks an icon of the set."""

    #: The icon-set payload.
    icon_set: CfIconSetInput
    #: Discriminant: always the icon-set rule kind.
    type: CfRuleType = field(default=CfRuleType.ICON_SET, init=False)


@dataclass(kw_only=True)
class CfTop10RuleInput(_CfFormatRuleInput):
    """``TOP10``: the top (or with ``bottom``, the bottom) ``rank`` values, or ``rank`` percent with ``percent``."""

    #: The number of values to select, or the percentage of values when ``percent`` is set.
    rank: Optional[int] = None
    #: ``True`` reads ``rank`` as a percentage of the values rather than a count; defaults to ``False``.
    percent: bool = False
    #: ``True`` selects the lowest values instead of the highest; defaults to ``False``.
    bottom: bool = False
    #: Discriminant: always the top/bottom rule kind.
    type: CfRuleType = field(default=CfRuleType.TOP10, init=False)


@dataclass(kw_only=True)
class CfAboveAverageRuleInput(_CfFormatRuleInput):
    """``ABOVE_AVERAGE``: values above (or with ``above_average=False``, below) the mean."""

    #: ``True`` (the default) matches values above the average, ``False`` values below it.
    above_average: bool = True
    #: ``True`` also matches values equal to the average; defaults to ``False``.
    equal_average: bool = False
    #: The number of standard deviations from the average a value must differ by; omitted compares with the average
    #: itself.
    std_dev: Optional[float] = None
    #: Discriminant: always the above-average rule kind.
    type: CfRuleType = field(default=CfRuleType.ABOVE_AVERAGE, init=False)


@dataclass(kw_only=True)
class CfTextRuleInput(_CfFormatRuleInput):
    """``CONTAINS_TEXT`` / ``NOT_CONTAINS_TEXT`` / ``BEGINS_WITH`` / ``ENDS_WITH``: a text test."""

    #: The text test: contains, not-contains, begins-with or ends-with.
    type: CfRuleType
    #: The text the cell is tested against.
    text: str
    #: The stored test formula Excel writes beside the rule; omitted or ``None`` derives it.
    formula1: Optional[str] = None


@dataclass(kw_only=True)
class CfTimePeriodRuleInput(_CfFormatRuleInput):
    """``TIME_PERIOD``: a date in ``time_period`` relative to today."""

    #: The date period, relative to today, that a cell's date must fall in.
    time_period: DatePeriod
    #: The stored test formula Excel writes beside the rule; omitted or ``None`` derives it.
    formula1: Optional[str] = None
    #: Discriminant: always the time-period rule kind.
    type: CfRuleType = field(default=CfRuleType.TIME_PERIOD, init=False)


@dataclass(kw_only=True)
class CfPlainRuleInput(_CfFormatRuleInput):
    """``CONTAINS_BLANKS`` / ``NOT_CONTAINS_BLANKS`` / ``CONTAINS_ERRORS`` / ``NOT_CONTAINS_ERRORS`` /
    ``DUPLICATE_VALUES`` / ``UNIQUE_VALUES``: a test with no operand.
    """

    #: The test: blanks, errors, duplicate values or unique values, or the negation of blanks or errors.
    type: CfRuleType
    #: The stored test formula; omitted or ``None`` derives it.
    formula1: Optional[str] = None


#: The argument of :meth:`Sheet.add_cf_rule`: one dataclass per rule type, each carrying the fields its kind uses.
ConditionalFormatInput = Union[
    CfExpressionRuleInput,
    CfCellIsRuleInput,
    CfColorScaleRuleInput,
    CfDataBarRuleInput,
    CfIconSetRuleInput,
    CfTop10RuleInput,
    CfAboveAverageRuleInput,
    CfTextRuleInput,
    CfTimePeriodRuleInput,
    CfPlainRuleInput,
]


@dataclass(frozen=True)
class CellNode:
    """A workbook-wide cell coordinate (precedents / dependents trace)."""

    #: The 0-based index of the sheet.
    sheet: int
    #: The 0-based row.
    row: int
    #: The 0-based column.
    col: int


@dataclass(frozen=True)
class Spill:
    """Dynamic-array spill region info for a cell, or the range of a PivotTable report."""

    #: The cell holding the dynamic-array formula, or the PivotTable's anchor cell (below the page-field rows that
    #: ``range`` includes).
    anchor: CellAddress
    #: ``True`` when the cell is the anchor of a spill refused with ``#SPILL!``, or of a blocked PivotTable report;
    #: the region is then the rectangle it attempted. ``False`` for a drawn spill or report.
    blocked: bool
    #: What blocks the region; ``None`` when ``blocked`` is ``False``.
    cause: Optional[SpillCause]
    #: The region including its anchor, whose top-left cell holds the dynamic-array formula or is a PivotTable's
    #: anchor; a report's region also takes the page-field rows above its anchor. A blocked region running past the
    #: sheet edge is clipped to the grid.
    range: CellRect


@dataclass(frozen=True)
class ParsedReference:
    """A parsed reference: the rectangle ``range`` and its sheet qualifier (``None`` when unqualified)."""

    #: The unquoted sheet qualifier; ``None`` when the text has none.
    sheet_name: Optional[str]
    #: The 0-based inclusive rectangle, with a whole column or row expanded to the grid edge.
    range: CellRect


@dataclass(frozen=True)
class ArrayFormula:
    """A legacy array (CSE) formula: the block it fills and its stored text.

    As the argument of :meth:`Sheet.set_array_formula`, ``range`` also takes an unqualified A1 range.
    """

    #: The block the formula fills; its top-left cell holds the formula.
    range: Union[CellRect, str]
    #: The formula text, with a leading ``=``; the ``=`` is optional on input.
    formula: str


@dataclass(frozen=True)
class FunctionMetadata:
    """Catalog metadata for a single Formulon function."""

    #: The canonical function name.
    name: str
    #: The minimum number of arguments written at entry.
    min_arity: int
    #: The maximum number of arguments written at entry; ``None`` for an open argument list.
    max_arity: Optional[int]
    #: How far the engine implements the function: a real implementation, host-bound, or an intentionally unavailable
    #: fixed-error stub.
    availability: FunctionAvailability
    #: The signature template; ``None`` until a host metadata document fills it in.
    signature: Optional[str]
    #: The human-readable description; ``None`` until a host metadata document fills it in.
    description: Optional[str]


@dataclass(frozen=True)
class LocaleErrorName:
    """One error value's spelling under a profile."""

    #: The canonical (en-invariant) error text.
    canonical: str
    #: The profile's Excel spelling of the error.
    localized: str
    #: ``False`` when the localized spelling is assumed rather than captured from Excel.
    measured: bool


@dataclass(frozen=True)
class LocaleFacts:
    """Scalar locale facts of one :data:`ProfileId`.

    Every scalar field is captured in each Mac locale, so ``measured`` covers them all; only ``error_names``
    carries per-entry flags.
    """

    #: The decimal separator.
    decimal_separator: str
    #: The digit group (thousands) separator.
    group_separator: str
    #: The separator between function arguments in formulas.
    list_separator: str
    #: The separator between columns in an array constant.
    array_column_separator: str
    #: The separator between rows in an array constant.
    array_row_separator: str
    #: The spelling of the logical value TRUE.
    true_name: str
    #: The spelling of the logical value FALSE.
    false_name: str
    #: The order of year, month and day in short dates, a ``LocaleDateOrder`` value.
    date_order: LocaleDateOrder
    #: The currency symbol.
    currency_symbol: str
    #: ``True`` when the currency symbol follows the number.
    currency_suffix: bool
    #: ``True`` when one space separates the currency symbol and the number.
    currency_space: bool
    #: The number of decimal places the currency format shows by default.
    currency_default_decimals: int
    #: ``True`` for an Excel-backed profile (every mac profile and win-365-ja_JP); ``False`` for an estimated one.
    measured: bool
    #: ``True`` when the text collation of the profile was captured from Excel (every mac profile); ``False`` for the
    #: estimated win profiles, including win-365-ja_JP, whose ``measured`` is ``True`` for the other facts.
    collation_measured: bool
    #: The spelling of each error value under the profile, indexed by error-code ordinal.
    error_names: List[LocaleErrorName]


@dataclass(frozen=True)
class CellXf:
    """A resolved ``<xf>`` style record.

    An alignment or protection attribute is ``None`` when the record does
    not state it; the ``<alignment>`` child exists when any alignment
    attribute is stated, and the ``<protection>`` child when ``locked`` or
    ``hidden`` is.
    """

    #: Index of the font in the workbook font table.
    font_index: int
    #: Index of the fill in the workbook fill table.
    fill_index: int
    #: Index of the border in the workbook border table.
    border_index: int
    #: The number format id: a built-in id or one declared in ``<numFmts>``.
    num_fmt_id: int
    #: The horizontal alignment. ``None`` when the attribute is absent.
    horizontal: Optional[HorizontalAlign] = None
    #: The vertical alignment; bottom when the attribute is absent. ``None`` when the attribute is absent.
    vertical: Optional[VerticalAlign] = None
    #: Whether text wraps within the cell. ``None`` when the attribute is absent.
    wrap_text: Optional[bool] = None
    #: Whether the last line is justified; effective with distributed horizontal alignment. ``None`` when the attribute
    #: is absent.
    justify_last_line: Optional[bool] = None
    #: Index into ``<cellStyleXfs>`` of the named style this format inherits from.
    style_xf_index: int = 0
    #: The text rotation in degrees, 0 to 180, or 255 for vertical text. ``None`` when the attribute is absent.
    text_rotation: Optional[int] = None
    #: The indent level, 0 to 255. ``None`` when the attribute is absent.
    indent: Optional[int] = None
    #: The signed indent relative to the parent style. ``None`` when the attribute is absent.
    relative_indent: Optional[int] = None
    #: Whether text shrinks to fit the cell width. ``None`` when the attribute is absent.
    shrink_to_fit: Optional[bool] = None
    #: The text reading order. ``None`` when the attribute is absent.
    reading_order: Optional[ReadingOrder] = None
    #: Whether this record's number format overrides the one of its parent named style.
    apply_number_format: bool = False
    #: Whether this record's font overrides the one of its parent named style.
    apply_font: bool = False
    #: Whether this record's fill overrides the one of its parent named style.
    apply_fill: bool = False
    #: Whether this record's border overrides the one of its parent named style.
    apply_border: bool = False
    #: Whether this record's alignment overrides the one of its parent named style.
    apply_alignment: bool = False
    #: Whether this record's protection overrides the one of its parent named style.
    apply_protection: bool = False
    #: Whether the cell text carries a leading apostrophe that marks it as text.
    quote_prefix: bool = False
    #: Whether cells are locked once the sheet is protected; ``None`` without a ``<protection>`` child, which means
    #: locked.
    locked: Optional[bool] = None
    #: Whether the formula is hidden once the sheet is protected; ``None`` without a ``<protection>`` child, which means
    #: not hidden.
    hidden: Optional[bool] = None


@dataclass
class CellXfInput:
    """An ``<xf>`` record :meth:`WorkbookOoxml.add_cell_xf` and :meth:`WorkbookOoxml.add_cell_style_xf` write.

    An optional attribute left ``None`` is not written: ``<alignment>`` is written when any alignment
    attribute is given and ``<protection>`` when ``locked`` or ``hidden`` is. A :class:`CellXf` read back is
    accepted in its place. ``style_xf_index`` applies to ``add_cell_xf`` only.
    """

    #: Index of the font in the workbook font table.
    font_index: int
    #: Index of the fill in the workbook fill table.
    fill_index: int
    #: Index of the border in the workbook border table.
    border_index: int
    #: The number format id: a built-in id or one declared in ``<numFmts>``.
    num_fmt_id: int
    #: The horizontal alignment; omitted or ``None`` is an absent attribute.
    horizontal: Optional[HorizontalAlign] = None
    #: The vertical alignment; omitted or ``None`` is an absent attribute (bottom).
    vertical: Optional[VerticalAlign] = None
    #: Whether text wraps within the cell; omitted or ``None`` is an absent attribute.
    wrap_text: Optional[bool] = None
    #: Whether the last line is justified, effective with distributed alignment; omitted or ``None`` is an absent
    #: attribute.
    justify_last_line: Optional[bool] = None
    #: Whether this record's number format overrides the one of its parent named style; omitted is ``False``.
    apply_number_format: bool = False
    #: Whether this record's font overrides the one of its parent named style; omitted is ``False``.
    apply_font: bool = False
    #: Whether this record's fill overrides the one of its parent named style; omitted is ``False``.
    apply_fill: bool = False
    #: Whether this record's border overrides the one of its parent named style; omitted is ``False``.
    apply_border: bool = False
    #: Whether this record's alignment overrides the one of its parent named style; omitted is ``False``.
    apply_alignment: bool = False
    #: Whether this record's protection overrides the one of its parent named style; omitted is ``False``.
    apply_protection: bool = False
    #: Whether the cell text carries a leading apostrophe that marks it as text; omitted is ``False``.
    quote_prefix: bool = False
    #: Whether cells are locked once the sheet is protected; omitted or ``None`` is an absent attribute, which means
    #: locked.
    locked: Optional[bool] = None
    #: Whether the formula is hidden once the sheet is protected; omitted or ``None`` is an absent attribute, which
    #: means not hidden.
    hidden: Optional[bool] = None
    #: The text rotation in degrees, 0 to 180, or 255 for vertical text; omitted or ``None`` when absent.
    text_rotation: Optional[int] = None
    #: The indent level, 0 to 255; omitted or ``None`` when absent.
    indent: Optional[int] = None
    #: The signed indent relative to the parent style; omitted or ``None`` when absent.
    relative_indent: Optional[int] = None
    #: Whether text shrinks to fit the cell width; omitted or ``None`` when absent, so an explicit ``False`` is kept.
    shrink_to_fit: Optional[bool] = None
    #: The text reading order; omitted or ``None`` when absent.
    reading_order: Optional[ReadingOrder] = None
    #: Index into ``<cellStyleXfs>`` of the named style this format inherits from; it must name an entry added as a
    #: named-style format. Defaults to 0.
    style_xf_index: int = 0


@dataclass(frozen=True)
class Color:
    """A colour as an OOXML ``<color>`` element states it.

    Records read from a file keep the specification as written, so a theme
    colour stays a theme reference. An input that takes a colour also takes
    ``"#RRGGBB"`` or ``"#AARRGGBB"`` (:data:`ColorInput`).
    """

    #: Which selector the colour carries; only the matching field below is meaningful.
    kind: ColorKind = ColorKind.NONE
    #: The colour as 0xAARRGGBB; meaningful when ``kind`` is RGB.
    argb: int = 0
    #: The theme colour index; meaningful when ``kind`` is theme.
    theme: int = 0
    #: The theme tint from -1 (darker) to 1 (lighter); meaningful when ``kind`` is theme.
    tint: float = 0.0
    #: The legacy palette index; meaningful when ``kind`` is indexed.
    indexed: int = 0


@dataclass(frozen=True)
class PhoneticRun:
    """One ``<rPh>`` block: ``text`` reads the half-open span ``[start, end)`` of the cell text.

    Offsets are code points into the cell's ``str``. A whole-cell reading is
    the single run ``PhoneticRun(0, len(<cell text>), kana)``.
    """

    #: The inclusive start of the span, in code points.
    start: int = 0
    #: The exclusive end of the span, in code points.
    end: int = 0
    #: The kana reading of the span.
    text: str = ""


@dataclass(frozen=True)
class PhoneticProperties:
    """Result of :meth:`Sheet.get_cell_phonetic_properties`; the fields are those of
    ``set_cell_phonetic_properties``.
    """

    #: The index into the workbook's font table of the ruby text.
    font_index: int = 0
    #: The kana form, a ``PhoneticType`` value.
    type: PhoneticType = PhoneticType.HALFWIDTH_KATAKANA
    #: How the kana is distributed over the characters it covers, a ``PhoneticAlignment`` value.
    alignment: PhoneticAlignment = PhoneticAlignment.NO_CONTROL
    #: ``True`` when the properties are explicit: a phoneticPr element or a set call, an all-zero triple included;
    #: ``False`` for the state an absent element resolves to. Ignored on input.
    stated: bool = False

    def __post_init__(self) -> None:
        object.__setattr__(self, "type", _enum(PhoneticType, self.type))
        object.__setattr__(self, "alignment", _enum(PhoneticAlignment, self.alignment))


@dataclass(frozen=True)
class Font:
    """A font record (``add_font`` input / ``get_font`` result)."""

    #: The typeface name; ``None`` when the font names none.
    name: Optional[str] = None
    #: The font size in points; ``None`` without a ``<sz>`` element, which on a differential font leaves the size
    #: unchanged (a workbook font always has one).
    size_pt: Optional[float] = None
    #: Whether the font is bold (``<b>``); ``None`` without the element, which on a differential font leaves bold
    #: unchanged, while ``False`` switches it off.
    bold: Optional[bool] = None
    #: Whether the font is italic (``<i>``), with the same ``None`` rule as ``bold``.
    italic: Optional[bool] = None
    #: Whether the text is struck through (``<strike>``), with the same ``None`` rule as ``bold``.
    strike: Optional[bool] = None
    #: Whether the glyphs are drawn as outlines (``<outline>``), with the same ``None`` rule as ``bold``.
    outline: Optional[bool] = None
    #: Whether the glyphs carry a shadow (``<shadow>``), with the same ``None`` rule as ``bold``.
    shadow: Optional[bool] = None
    #: Whether the glyphs are condensed (``<condense>``), with the same ``None`` rule as ``bold``.
    condense: Optional[bool] = None
    #: Whether the glyphs are extended (``<extend>``), with the same ``None`` rule as ``bold``.
    extend: Optional[bool] = None
    #: The underline style.
    underline: UnderlineStyle = UnderlineStyle.NONE
    #: The vertical alignment of the text run (baseline, superscript or subscript).
    vert_align: FontVerticalAlign = FontVerticalAlign.BASELINE
    #: The font-family class; ``None`` without a ``<family>`` element.
    family: Optional[FontFamily] = None
    #: The charset codepage id (for example 128 for Shift_JIS); ``None`` without a ``<charset>`` element.
    charset: Optional[int] = None
    #: The theme font set the font links to (major or minor); none when the font has no ``<scheme>`` element. Meaningful
    #: only on a workbook that carries a theme part.
    scheme: FontScheme = FontScheme.NONE
    #: The ``<color>`` element, a ``ColorInput`` (or ``None``, the absent colour) on input; the none kind without one,
    #: which renders as the automatic colour.
    color: Color = field(default_factory=Color)


@dataclass(frozen=True)
class GradientStop:
    """One ``<stop>`` of a gradient fill: a colour at ``position`` (0..1)."""

    #: The position of the stop along the gradient, 0..1.
    position: float = 0.0
    #: The colour at the stop.
    color: Color = field(default_factory=Color)


@dataclass(frozen=True)
class GradientFill:
    """A ``<gradientFill>``."""

    #: Whether the gradient is linear or path.
    type: GradientType = GradientType.LINEAR
    #: The angle of a linear gradient in degrees.
    angle_deg: float = 0.0
    #: The left edge (0..1) of a path gradient's focus.
    left: float = 0.0
    #: The right edge (0..1) of a path gradient's focus.
    right: float = 0.0
    #: The top edge (0..1) of a path gradient's focus.
    top: float = 0.0
    #: The bottom edge (0..1) of a path gradient's focus.
    bottom: float = 0.0
    #: The gradient's colour stops; writing more than 1024 raises ``SecurityError``.
    stops: List[GradientStop] = field(default_factory=list)


@dataclass(frozen=True)
class Fill:
    """A fill record."""

    #: The fill pattern (none, solid or one of the standard patterns).
    pattern: FillPattern = FillPattern.NONE
    #: The pattern's ``<fgColor>``, a ``ColorInput`` on input; the none kind when the element is absent.
    foreground: Color = field(default_factory=Color)
    #: The pattern's ``<bgColor>``, a ``ColorInput`` on input; the none kind when the element is absent.
    background: Color = field(default_factory=Color)
    #: The gradient of a ``<gradientFill>``, which makes the pattern fields ignored; ``None`` for a pattern fill.
    gradient: Optional[GradientFill] = None


@dataclass(frozen=True)
class BorderSide:
    """One side of a border: its style and colour (a ``ColorInput`` on input)."""

    #: The line style of the side (none, thin, medium, dashed, ...).
    style: BorderStyle = BorderStyle.NONE
    #: The side's ``<color>``; the none kind without one.
    color: Color = field(default_factory=Color)


@dataclass(frozen=True)
class Border:
    """A border record of the style table: four sides, the diagonal and its directions."""

    #: The left side.
    left: BorderSide = field(default_factory=BorderSide)
    #: The right side.
    right: BorderSide = field(default_factory=BorderSide)
    #: The top side.
    top: BorderSide = field(default_factory=BorderSide)
    #: The bottom side.
    bottom: BorderSide = field(default_factory=BorderSide)
    #: The line drawn diagonally; ``diagonal_up`` and ``diagonal_down`` choose its directions.
    diagonal: BorderSide = field(default_factory=BorderSide)
    #: Whether the diagonal runs from bottom left to top right.
    diagonal_up: bool = False
    #: Whether the diagonal runs from top left to bottom right.
    diagonal_down: bool = False


_BORDER_SIDES = ("left", "right", "top", "bottom", "diagonal")


@dataclass(frozen=True)
class Alignment:
    """How a cell's text is placed (OOXML ``<alignment>``).

    A format without the element reads ``GENERAL`` / ``BOTTOM``, every flag
    false and every number 0, with ``ReadingOrder.CONTEXT``.
    """

    #: The horizontal alignment.
    horizontal: HorizontalAlign = HorizontalAlign.GENERAL
    #: The vertical alignment.
    vertical: VerticalAlign = VerticalAlign.BOTTOM
    #: Whether text wraps within the cell.
    wrap_text: bool = False
    #: Whether text shrinks to fit the cell width.
    shrink_to_fit: bool = False
    #: The indent level, 0 to 255.
    indent: int = 0
    #: The text rotation in degrees, 0 to 180, or 255 for vertical text.
    text_rotation: int = 0
    #: The text reading order.
    reading_order: ReadingOrder = ReadingOrder.CONTEXT
    #: Whether the last line is justified; effective with distributed horizontal alignment.
    justify_last_line: bool = False
    #: The signed indent relative to the parent style.
    relative_indent: int = 0


@dataclass(frozen=True)
class CellProtection:
    """Whether a cell is locked and its formula hidden while the sheet is protected (OOXML ``<protection>``)."""

    #: ``True`` when the cell is locked, so it cannot be edited while its sheet is protected.
    locked: bool = True
    #: ``True`` when the cell's formula is hidden from the formula bar while its sheet is protected.
    hidden: bool = False


@dataclass(frozen=True)
class CellFormat:
    """The format a cell shows, as :meth:`Sheet.get_cell_format` reads it.

    Read from the cell's effective XF (cell, then row, then column, then the
    default); conditional formatting is not applied. The XF index is
    :meth:`SheetOoxml.get_cell_xf_index`.
    """

    #: Which layer (cell, row, column or default) the effective XF came from.
    source: EffectiveStyleSource
    #: The number format code; ``General`` for General.
    number_format: str
    #: The cell's font.
    font: Font
    #: The cell's fill.
    fill: Fill
    #: The cell's border.
    border: Border
    #: The XF's alignment; the defaults when it has none.
    alignment: Alignment
    #: The XF's protection; locked and not hidden when it has none.
    protection: CellProtection


#: A colour in a format patch: ``"#RRGGBB"`` or ``"#AARRGGBB"`` (either case;
#: alpha defaults to ``FF``), or a :class:`Color`.
ColorInput = Union[str, Color]


class FontPatch(TypedDict, total=False):
    """The font keys of :meth:`Sheet.update_cell_format`; an omitted key keeps its value."""

    #: The typeface name; omitted keeps it, ``None`` resets it to the Normal style's font name. A different name with
    #: ``scheme`` omitted also sets ``scheme`` to ``FontScheme.NONE``, as Excel unlinks a font from the theme when its
    #: name changes: a theme-linked font is drawn with the theme's face whatever name it stores.
    name: Optional[str]
    #: The font size in points, 1 to 409; omitted keeps it, ``None`` resets it to the Normal style's font size.
    size_pt: Optional[float]
    #: Whether the font is bold; omitted keeps it, ``None`` removes the ``<b>`` element.
    bold: Optional[bool]
    #: Whether the font is italic; omitted keeps it, ``None`` removes the ``<i>`` element.
    italic: Optional[bool]
    #: Whether the text is struck through; omitted keeps it, ``None`` removes the ``<strike>`` element.
    strike: Optional[bool]
    #: Whether the glyphs are drawn as outlines; omitted keeps it, ``None`` removes the ``<outline>`` element.
    outline: Optional[bool]
    #: Whether the glyphs carry a shadow; omitted keeps it, ``None`` removes the ``<shadow>`` element.
    shadow: Optional[bool]
    #: Whether the glyphs are condensed; omitted keeps it, ``None`` removes the ``<condense>`` element.
    condense: Optional[bool]
    #: Whether the glyphs are extended; omitted keeps it, ``None`` removes the ``<extend>`` element.
    extend: Optional[bool]
    #: The underline style; omitted keeps it.
    underline: UnderlineStyle
    #: The vertical alignment of the text run (baseline, superscript or subscript); omitted keeps it.
    vert_align: FontVerticalAlign
    #: The font-family class; omitted keeps it, ``None`` removes the ``<family>`` element.
    family: Optional[FontFamily]
    #: The charset codepage id; omitted keeps it, ``None`` removes the ``<charset>`` element. A name change keeps it:
    #: Excel derives it from the new face, so a caller that knows the face's charset states it.
    charset: Optional[int]
    #: The theme font set the font links to; omitted keeps it.
    scheme: FontScheme
    #: The font colour; omitted keeps it, ``None`` removes the ``<color>`` element.
    color: Optional[ColorInput]


class FillPatch(TypedDict, total=False):
    """The fill keys of :meth:`Sheet.update_cell_format`.

    Setting ``foreground`` or ``background`` on a fill without a pattern makes
    it ``FillPattern.SOLID`` unless ``pattern`` is given; any fill key drops a
    gradient.
    """

    #: The fill pattern; omitted keeps it.
    pattern: FillPattern
    #: The pattern's foreground colour; omitted keeps it. Setting it on a fill without a pattern makes the fill solid
    #: unless ``pattern`` is given.
    foreground: ColorInput
    #: The pattern's background colour; omitted keeps it. Setting it on a fill without a pattern makes the fill solid
    #: unless ``pattern`` is given.
    background: ColorInput
    #: A gradient to set, which replaces the pattern; omitted keeps it, ``None`` drops the gradient.
    gradient: Optional[GradientFill]


class BorderSidePatch(TypedDict, total=False):
    """One border side of a :class:`BorderPatch`."""

    #: The line style of the side; omitted keeps it.
    style: BorderStyle
    #: The colour of the side; omitted keeps it.
    color: ColorInput


class BorderPatch(TypedDict, total=False):
    """The border keys of :meth:`Sheet.update_cell_format`; each side merges into the current side."""

    #: Changes to the left side, merged into the current side; omitted keeps it.
    left: BorderSidePatch
    #: Changes to the right side, merged into the current side; omitted keeps it.
    right: BorderSidePatch
    #: Changes to the top side, merged into the current side; omitted keeps it.
    top: BorderSidePatch
    #: Changes to the bottom side, merged into the current side; omitted keeps it.
    bottom: BorderSidePatch
    #: Changes to the diagonal line, merged into the current line; omitted keeps it.
    diagonal: BorderSidePatch
    #: Whether the diagonal runs from bottom left to top right; omitted keeps it.
    diagonal_up: bool
    #: Whether the diagonal runs from top left to bottom right; omitted keeps it.
    diagonal_down: bool


class AlignmentPatch(TypedDict, total=False):
    """The alignment keys of :meth:`Sheet.update_cell_format`."""

    #: The horizontal alignment; omitted keeps it.
    horizontal: HorizontalAlign
    #: The vertical alignment; omitted keeps it.
    vertical: VerticalAlign
    #: Whether text wraps within the cell; omitted keeps it.
    wrap_text: bool
    #: Whether text shrinks to fit the cell width; omitted keeps it.
    shrink_to_fit: bool
    #: The indent level, 0 to 255; omitted keeps it.
    indent: int
    #: The text rotation in degrees, 0 to 180, or 255 for vertical text; omitted keeps it.
    text_rotation: int
    #: The text reading order; omitted keeps it.
    reading_order: ReadingOrder
    #: Whether the last line is justified; effective with distributed horizontal alignment. Omitted keeps it.
    justify_last_line: bool
    #: The signed indent relative to the parent style; omitted keeps it.
    relative_indent: int


class CellProtectionPatch(TypedDict, total=False):
    """The protection keys of :meth:`Sheet.update_cell_format`."""

    #: Whether cells are locked once the sheet is protected; omitted keeps it.
    locked: bool
    #: Whether the formula is hidden once the sheet is protected; omitted keeps it.
    hidden: bool


@dataclass(frozen=True)
class Dxf:
    """A differential format (``<dxf>``), the formatting a conditional format or table style applies.

    The record :meth:`WorkbookOoxml.get_dxf` returns.
    """

    #: The font part; ``None`` when the format sets none.
    font: Optional[Font] = None
    #: The fill part; ``None`` when the format sets none.
    fill: Optional[Fill] = None
    #: The border part; ``None`` when the format sets none.
    border: Optional[Border] = None
    #: The number format code; ``None`` when the format sets none. The format's id is assigned on save.
    number_format: Optional[str] = None
    #: The alignment part; ``None`` when the format sets none.
    alignment: Optional[Alignment] = None
    #: The protection part; ``None`` when the format sets none.
    protection: Optional[CellProtection] = None


@dataclass
class DxfInput:
    """The differential format :meth:`WorkbookOoxml.add_dxf` takes; the parts and defaults of :class:`Dxf`."""

    #: The font part; omitted or ``None`` is not set.
    font: Optional[Font] = None
    #: The fill part; omitted or ``None`` is not set.
    fill: Optional[Fill] = None
    #: The border part; omitted or ``None`` is not set.
    border: Optional[Border] = None
    #: A number format code; omitted or ``None`` is not set. The format's id is assigned on save.
    number_format: Optional[str] = None
    #: The alignment part; omitted or ``None`` is not set.
    alignment: Optional[Alignment] = None
    #: The protection part; omitted or ``None`` is not set.
    protection: Optional[CellProtection] = None


def _decode_color(ptr: int) -> Color:
    return _color_from(S.COLOR_SPEC.unpack(LIB, ptr))


def _decode_optional_color(ptr: int) -> Optional[Color]:
    """:func:`_decode_color`, ``None`` for an absent colour (``ColorKind.NONE``)."""
    return _optional_color_from(S.COLOR_SPEC.unpack(LIB, ptr))


def _color_fields(spec: Color) -> Dict[str, object]:
    """The colour-spec record fields of ``spec``."""
    return {
        "kind": _uint(spec.kind, "kind", 8),
        "argb": _uint(spec.argb, "argb"),
        "theme": _uint(spec.theme, "theme"),
        "tint": _require_number(spec.tint, "tint", finite=False),
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


def _pack_optional_color(ptr: int, spec: Optional[ColorInput], label: str = "color") -> None:
    """:func:`_pack_color`, with ``None`` writing the absent colour (``ColorKind.NONE``)."""
    _pack_color(ptr, Color() if spec is None else spec, label)


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
        "size_pt": 0.0 if record.size_pt is None else _require_number(record.size_pt, "size_pt", finite=False),
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
        fields[name] = 0 if value is None else int(_require_bool(value, name))
        fields[f"has_{name}"] = 0 if value is None else 1
    S.FONT_RECORD.pack(LIB, ptr, fields)
    _pack_optional_color(ptr + S.FONT_RECORD.offsets["color"][1], record.color, "Font.color")
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
            angle_deg=d["gradient_angle_deg"],
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
                S.GRADIENT_STOP.pack(
                    LIB, stop_ptr, {"position": _require_number(stop.position, "position", finite=False)}
                )
                _pack_color(stop_ptr + S.GRADIENT_STOP.offsets["color"][1], stop.color, f"GradientStop[{i}].color")
        values.update(
            has_gradient=1,
            gradient_type=_sint(gradient.type, "gradient.type"),
            gradient_angle_deg=_require_number(gradient.angle_deg, "angle_deg", finite=False),
            gradient_left=_require_number(gradient.left, "left", finite=False),
            gradient_right=_require_number(gradient.right, "right", finite=False),
            gradient_top=_require_number(gradient.top, "top", finite=False),
            gradient_bottom=_require_number(gradient.bottom, "bottom", finite=False),
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
        raise TypeError(f"expected a Border, got {type(border).__name__}")
    S.BORDER_RECORD.pack(
        LIB,
        ptr,
        {
            "diagonal_up": _require_bool(border.diagonal_up, "diagonal_up"),
            "diagonal_down": _require_bool(border.diagonal_down, "diagonal_down"),
        },
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
        raise TypeError(f"Dxf.alignment must be an Alignment, got {type(alignment).__name__}")
    default = Alignment()
    attrs: List[str] = []
    if alignment.horizontal != default.horizontal:
        attrs.append(f'horizontal="{_HORIZONTAL_TOKENS[HorizontalAlign(alignment.horizontal)]}"')
    if alignment.vertical != default.vertical:
        attrs.append(f'vertical="{_VERTICAL_TOKENS[VerticalAlign(alignment.vertical)]}"')
    for name, attr in _ALIGNMENT_NUMBERS:
        value = getattr(alignment, name)
        if value != getattr(default, name):
            attrs.append(f'{attr}="{_sint(value, name)}"')
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
        raise TypeError(f"Dxf.protection must be a CellProtection, got {type(protection).__name__}")
    locked = int(_require_bool(protection.locked, "locked"))
    hidden = int(_require_bool(protection.hidden, "hidden"))
    return f'<protection locked="{locked}" hidden="{hidden}"/>'


@dataclass(frozen=True)
class NumberFormat:
    """One declared number format: its id and format code (:meth:`WorkbookOoxml.list_number_formats`)."""

    #: The number format id.
    num_fmt_id: int
    #: The number format code.
    number_format: str


@dataclass(frozen=True)
class CellStyle:
    """A named cell style (``<cellStyle>`` entry); ``style_xf_index`` is its ``<cellStyleXfs>`` entry."""

    #: The display name of the style.
    name: str
    #: Index into the ``<cellStyleXfs>`` table of the format the style applies.
    style_xf_index: int = 0
    #: The built-in style ordinal (0..53); ``None`` for a custom style.
    builtin_id: Optional[int] = None
    #: The outline level (0..6) of the RowLevel_n / ColLevel_n styles; 0 for other styles.
    outline_level: int = 0
    #: Whether the style is hidden from the style gallery.
    hidden: bool = False
    #: Whether this is a built-in style the workbook has changed;
    #: :meth:`WorkbookOoxml.set_cell_style_xf` sets it on every built-in style naming the edited slot.
    custom_builtin: bool = False


@dataclass(frozen=True)
class ThemeFonts:
    """The theme's major (heading) and minor (body) typefaces."""

    #: The Latin typeface of the major (heading) font set.
    major_latin: str
    #: The East Asian typeface of the major (heading) font set: the ``a:ea`` face, else the Jpan script face.
    major_east_asian: str
    #: The Latin typeface of the minor (body) font set.
    minor_latin: str
    #: The East Asian typeface of the minor (body) font set: the ``a:ea`` face, else the Jpan script face.
    minor_east_asian: str


@dataclass(frozen=True)
class ThemeScriptFont:
    """One ``a:font`` entry of the major (``FontScheme.MAJOR``, heading) or minor (``FontScheme.MINOR``, body) theme font set."""

    #: The font set the face belongs to: major (heading) or minor (body).
    scheme: FontScheme
    #: The ISO 15924 script code, such as Jpan.
    script: str
    #: The typeface used for the script.
    typeface: str


@dataclass(frozen=True)
class ThemeColors:
    """The theme's 12 scheme colours as ``0xFFRRGGBB`` values in ``a:clrScheme`` order."""

    #: Where the theme came from: the theme part, or the default theme when the part is absent or unparseable.
    source: ThemeSource
    #: The 12 scheme colours as 0xFFRRGGBB in ``a:clrScheme`` order: dk1, lt1, dk2, lt2, accent1 to accent6, hlink,
    #: folHlink. This is not the order of a colour's theme index, which swaps the first two pairs (0 = lt1, 1 = dk1, 2 =
    #: lt2, 3 = dk2).
    colors: List[int]


@dataclass
class ThemeColorsInput:
    """The 12 scheme colours :meth:`Workbook.set_theme_colors` writes, as ``0xFFRRGGBB`` values in
    ``a:clrScheme`` order (alpha ignored). A :class:`ThemeColors` read back is accepted in its place; its
    ``source`` is ignored.
    """

    #: The 12 scheme colours in ``a:clrScheme`` order: dk1, lt1, dk2, lt2, accent1 to accent6, hlink, folHlink. Each is
    #: an RGB colour.
    colors: Sequence[int]


@dataclass(frozen=True)
class ResolvedColor:
    """A colour resolved to ``0xAARRGGBB`` and how it was resolved."""

    #: The resolved colour as 0xAARRGGBB.
    argb: int
    #: How the colour was obtained (exact, default theme, out-of-range index, unparseable theme or automatic).
    resolution: ColorResolution


@dataclass(frozen=True)
class EffectiveBorders:
    """The resolved colour of each border of an effective style; ``diagonal`` is the diagonal border."""

    #: The resolved colour of the left border.
    left: ResolvedColor
    #: The resolved colour of the right border.
    right: ResolvedColor
    #: The resolved colour of the top border.
    top: ResolvedColor
    #: The resolved colour of the bottom border.
    bottom: ResolvedColor
    #: The resolved colour of the diagonal border.
    diagonal: ResolvedColor


@dataclass(frozen=True)
class EffectiveStyle:
    """The formatting a cell shows, read from its selected xf record."""

    #: The index of the selected cell format record.
    xf_index: int
    #: The level that supplied the format record, an ``EffectiveStyleSource`` value.
    source: EffectiveStyleSource
    #: The index into the font table.
    font_index: int
    #: The index into the fill table.
    fill_index: int
    #: The index into the border table.
    border_index: int
    #: The resolved font color.
    font_color: ResolvedColor
    #: The resolved foreground (pattern) color of the fill.
    fill_foreground: ResolvedColor
    #: The resolved background color of the fill.
    fill_background: ResolvedColor
    #: The resolved color of each border.
    borders: EffectiveBorders
    #: The cell protection flags.
    protection: CellProtection
    #: The number format code.
    number_format: str


@dataclass(frozen=True)
class ExternalLink:
    """An external-link record (``<externalReferences>`` entry)."""

    #: The 1-based document order in the externalReferences list: the ``[n]`` of formula text, as in ``=[1]Sheet1!A1``.
    book_number: int
    #: The id of the link in the workbook relationships, such as ``rId3``.
    rel_id: str
    #: The resolved package-relative path of the body part, such as ``xl/externalLinks/externalLink1.xml``.
    part_path: str
    #: The URL or path of the remote workbook; ``None`` when the link's relationship names no target.
    target: Optional[str]
    #: ``True`` when the relationship's TargetMode is External.
    target_external: bool
    #: The kind of link, an ``ExternalLinkKind`` value.
    kind: ExternalLinkKind


@dataclass(frozen=True)
class PivotCell:
    """One projected PivotTable layout cell."""

    #: The absolute 0-based sheet row of the cell.
    row: int
    #: The absolute 0-based sheet column of the cell.
    col: int
    #: The cell's value; a date item label reads as its serial number, and ``number_format`` tells it is a date.
    value: NativeValue
    #: What the cell is in the layout: a header, a row or column label, a data value, a subtotal, a grand total or a
    #: blank.
    kind: PivotCellKind
    #: The nesting depth of a header cell; 0 for data cells.
    depth: int
    #: The name of the source field the cell belongs to; ``None`` when the cell has no source field.
    field_name: Optional[str]
    #: The data field's number format code of the cell; ``None`` when it has none.
    number_format: Optional[str]


@dataclass(frozen=True)
class PivotLayout:
    """The projected rectangular layout of a PivotTable."""

    #: The rectangle the layout covers in absolute sheet coordinates; ``None`` when it spans no cell.
    range: Optional[CellRect]
    #: The projected cells, sparse and in row-major order.
    cells: List[PivotCell]


@dataclass(frozen=True)
class PivotWorksheetSource:
    """A pivot cache's worksheet source, as :meth:`PivotCache.get_worksheet_source` reads it."""

    #: The source range with its sheet; ``None`` when the source has no reference or names a sheet the workbook does not
    #: have.
    range: Optional[SheetRect] = None
    #: The defined name or table the source reads; ``None`` when absent.
    name: Optional[str] = None


@dataclass
class PivotWorksheetSourceInput:
    """A pivot cache's worksheet source for :meth:`PivotCache.set_worksheet_source`.

    A qualified range or ``SheetRect`` whose sheet disagrees with ``sheet`` is an ``InvalidArgumentError``. A
    :class:`PivotWorksheetSource` read back is accepted in its place.
    """

    #: The source range, a rectangle or an A1 range; omitted leaves the attribute off. An unqualified A1 range takes its
    #: sheet from ``sheet``.
    range: Union[SheetRect, str, None] = None
    #: The sheet that qualifies an unqualified A1 ``range``; it may name a sheet the workbook does not have. Omitted
    #: leaves the attribute off.
    sheet: Optional[SheetRef] = None
    #: The defined name or table the source reads; omitted leaves the attribute off.
    name: Optional[str] = None


@dataclass
class PivotDateGroupInput:
    """A date grouping :meth:`Pivot.set_field_date_group` applies."""

    #: The date periods to group by, an OR of ``PivotDatePeriod`` bits.
    periods: PivotDatePeriod
    #: The calendar the periods follow; Gregorian by default.
    calendar: PivotCalendar = PivotCalendar.GREGORIAN
    #: The length in days of each group when grouping by days alone; 1 groups by single days.
    interval_days: int = 1
    #: The first bound of the grouping, a ``date`` or a ``datetime`` (a naive local wall clock); the field's first value
    #: when not given.
    start: Optional[Union[datetime.datetime, datetime.date]] = None
    #: The last bound of the grouping, a ``date`` or a ``datetime`` (a naive local wall clock); the field's last value
    #: when not given.
    end: Optional[Union[datetime.datetime, datetime.date]] = None
    #: ``True`` makes the first bound follow the field's data whatever ``start`` holds; ``False`` requires ``start``.
    auto_start: Optional[bool] = None
    #: ``True`` makes the last bound follow the field's data whatever ``end`` holds; ``False`` requires ``end``.
    auto_end: Optional[bool] = None


@dataclass(frozen=True)
class PivotDateGroup:
    """A date grouping as :meth:`Pivot.get_field_date_group` reads it.

    The members are those of :class:`PivotDateGroupInput`. A read record is accepted as it is by
    :meth:`Pivot.set_field_date_group`; a bound whose ``auto_*`` is true follows the data and its stored value is
    not written.
    """

    #: The ``PivotDatePeriod`` bits the field is grouped by.
    periods: PivotDatePeriod
    #: The calendar the periods follow.
    calendar: PivotCalendar
    #: The length in days of each group when grouping by days alone; 1 otherwise.
    interval_days: int
    #: The first bound as stored (a naive local wall clock), even where ``auto_start`` says it follows the data.
    start: datetime.datetime
    #: The last bound as stored (a naive local wall clock), even where ``auto_end`` says it follows the data.
    end: datetime.datetime
    #: Whether the first bound follows the field's smallest value.
    auto_start: bool
    #: Whether the last bound follows the field's largest value.
    auto_end: bool


@dataclass
class PivotNumericGroupInput:
    """A numeric grouping :meth:`Pivot.set_field_numeric_group` applies."""

    #: The width of each group.
    interval: float
    #: The first bound; the field's smallest value when not given.
    start: Optional[float] = None
    #: The last bound; the field's largest value when not given.
    end: Optional[float] = None
    #: ``True`` makes the first bound follow the field's data whatever ``start`` holds; ``False`` requires ``start``.
    auto_start: Optional[bool] = None
    #: ``True`` makes the last bound follow the field's data whatever ``end`` holds; ``False`` requires ``end``.
    auto_end: Optional[bool] = None


@dataclass(frozen=True)
class PivotNumericGroup:
    """A numeric grouping as :meth:`Pivot.get_field_numeric_group` reads it.

    A read record is accepted as it is by :meth:`Pivot.set_field_numeric_group`; a bound whose ``auto_*`` is true
    follows the data and its stored value is not written.
    """

    #: The width of each group.
    interval: float
    #: The first bound as stored, even where ``auto_start`` says it follows the data.
    start: float
    #: The last bound as stored, even where ``auto_end`` says it follows the data.
    end: float
    #: Whether the first bound follows the field's smallest value.
    auto_start: bool
    #: Whether the last bound follows the field's largest value.
    auto_end: bool


@dataclass
class PivotInput:
    """A pivot table :meth:`Sheet.add_pivot` creates: its name, its cache and its top-left cell."""

    #: The new pivot's name.
    name: str
    #: The id of the pivot cache the pivot draws from.
    cache_id: PivotCacheRef
    #: The top-left cell of the report, a cell address or its A1 text such as ``C3``.
    anchor: Union[CellAddress, str]


@dataclass(frozen=True)
class PivotProperties:
    """A pivot table's definition, as :meth:`Pivot.get_properties` returns it, composed of what the separate
    getters read. ``style_options`` is ``None`` when the definition has no style element."""

    #: The pivot's name.
    name: str
    #: The id of the pivot cache the pivot draws from.
    cache_id: int
    #: The report layout form: compact, outline or tabular.
    report_layout: PivotReportLayout
    #: Whether the row and column grand total bands show.
    grand_totals: PivotGrandTotals
    #: The table style of the pivot; ``None`` when the definition has no style element.
    style_options: Optional[PivotStyleOptions]


@dataclass(frozen=True)
class PivotCacheField:
    """One pivot-cache field.

    The ``contains_*`` flags are the field's ``<sharedItems>`` content flags, read-only: an attribute the file
    omits reads as its schema default, and a field with none of its own reads as the flags derived from its values.
    """

    #: The name of the field.
    name: str
    #: A calculated field's formula, leading ``=`` included; ``None`` for a source field.
    formula: Optional[str]
    #: Whether the field holds a number.
    contains_number: bool = False
    #: Whether every number the field holds is an integer.
    contains_integer: bool = False
    #: Whether the field holds text, a boolean or an error.
    contains_string: bool = False
    #: Whether the field holds a blank.
    contains_blank: bool = False
    #: Whether the field holds a date.
    contains_date: bool = False
    #: Whether the field mixes two or more of text, numbers and dates.
    contains_mixed_types: bool = False
    #: Whether the field holds text or a blank.
    contains_semi_mixed_types: bool = False
    #: Whether the field holds text or a number, something that is not a date.
    contains_non_date: bool = False


@dataclass(frozen=True)
class PivotCacheValue:
    """One pivot-cache value, a shared item or a record value; a date reads as its serial number and has ``is_date`` set."""

    #: The cache value; a blank is ``None`` and a date reads as its serial number, which ``is_date`` tells apart.
    value: NativeValue
    #: Whether the value is a date.
    is_date: bool


@dataclass(frozen=True)
class PivotField:
    """A pivot field's definition, as read by :meth:`Pivot.list_fields`.

    Read a grouping with :meth:`Pivot.get_field_date_group` or :meth:`Pivot.get_field_numeric_group`. A text the
    field does not state is ``None``.
    """

    #: The name of the bound cache field.
    source_name: str
    #: The report name of the field; ``None`` when the field has no report name.
    custom_name: Optional[str]
    #: The axis the field sits on; ``None`` when it sits on no axis.
    axis: Optional[PivotAxis]
    #: Whether subtotals show on the group's own heading row (true) or at the bottom (false).
    subtotal_top: bool
    #: The number format code of the field in the workbook profile's vocabulary; ``None`` for none or an unregistered
    #: id.
    number_format: Optional[str]
    #: The index of the bound pivot-cache field, ``None`` when the cache does not cover this field. Pivot field i binds
    #: to cache field i.
    cache_field_index: Optional[int]
    #: How the field's items are ordered.
    sort_order: PivotSortOrder
    #: The data field the items sort by; ``None`` sorts by label.
    sort_by_field: Optional[str]
    #: For a group field, the index of the field it groups; ``None`` otherwise.
    group_base_index: Optional[int]
    #: Whether the automatic subtotal shows; false whenever ``subtotal_functions`` is non-empty, since custom functions
    #: replace it.
    default_subtotal: bool
    #: The custom subtotal functions, in ascending aggregation order; empty when there are none.
    subtotal_functions: List[PivotAggregation]
    #: Whether a blank row follows each item block of the field.
    insert_blank_row: bool = False
    #: Whether the field's item labels repeat on every row of its block.
    fill_down_labels: bool = False
    #: Whether the field lists items without data.
    show_all: bool = False
    #: True only when every item and the field's default expand marker are expanded; a field without items reads true.
    show_detail: bool = True


@dataclass(frozen=True)
class PivotItem:
    """One pivot-field item, as read by :meth:`Pivot.get_field_item`."""

    #: The item's label: the one it was added with, else its shared item's value;
    #: ``None`` for the blank item or an item bound to nothing. A loaded item reads the same.
    name: Optional[str]
    #: The index into the bound cache field's shared items, ``None`` for an item bound by label only.
    cache_index: Optional[int]
    #: Whether the item is shown.
    visible: bool
    #: The caption shown instead of the item's value; ``None`` when it has none.
    caption: Optional[str] = None
    #: Whether the item is expanded; false for a collapsed item.
    show_detail: bool = True


@dataclass(frozen=True)
class PivotStyleOptions:
    """A pivot table's style, as read by :meth:`Pivot.get_style_options`.

    ``style_name`` is ``None`` when the style element carries no style name.
    :meth:`Pivot.update_style_options` takes these fields as keyword arguments.
    """

    #: The name of the table style; ``None`` when the style element carries no name.
    style_name: Optional[str]
    #: Whether the style formats the row headers.
    show_row_headers: bool
    #: Whether the style formats the column headers.
    show_column_headers: bool
    #: Whether the style bands the rows.
    show_row_stripes: bool
    #: Whether the style bands the columns.
    show_column_stripes: bool
    #: Whether the style formats the last column.
    show_last_column: bool


@dataclass(frozen=True)
class PivotOptions:
    """The PivotTable Options of one pivot table, as read by :meth:`Pivot.get_options`.

    A read fills every member; a string option reads ``None`` when unset.
    :meth:`Pivot.update_options` takes these fields as keyword arguments.
    """

    #: Whether error cells show ``error_caption``; the caption is kept while this is off.
    show_error: bool
    #: The text error cells show while ``show_error`` is on; ``None`` when unset.
    error_caption: Optional[str]
    #: Whether empty value cells show ``missing_caption``.
    show_missing: bool
    #: The text empty value cells show while ``show_missing`` is on; a caption of exactly ``-`` is written as the number
    #: 0. ``None`` when unset.
    missing_caption: Optional[str]
    #: Whether field captions and filter drop-downs show.
    show_headers: bool
    #: Whether the expand / collapse buttons show.
    show_drill: bool
    #: Whether contextual tooltips show.
    show_data_tips: bool
    #: Whether the classic layout with drop zones drawn on the grid is used.
    grid_drop_zones: bool
    #: Whether items without data show on rows; the layout follows each field's own show-all setting.
    show_empty_row: bool
    #: Whether items without data show on columns; the layout follows each field's own show-all setting.
    show_empty_column: bool
    #: Whether the field drop-down buttons show.
    show_data_drop_down: bool
    #: Whether the expand / collapse buttons are printed.
    print_drill: bool
    #: Whether field captions repeat as print titles on each printed page.
    field_print_titles: bool
    #: Whether row labels repeat on each printed page.
    item_print_titles: bool
    #: Whether item labels are merged and centered.
    merge_item: bool
    #: The indent per level of compact-form labels, 0 to 126 (the dialog shows this plus one).
    indent: int
    #: Whether report filter fields fill across before they go down.
    page_over_then_down: bool
    #: The number of report filter fields per column or row before wrapping, 0 to 255; 0 does not wrap.
    page_wrap: int
    #: Whether a field may carry more than one filter type.
    multiple_field_filters: bool
    #: Whether cell formatting is kept when the report updates.
    preserve_formatting: bool
    #: Whether custom lists take part in sorting.
    custom_list_sort: bool
    #: The alternative-text title of the PivotTable; ``None`` when unset.
    alt_text: Optional[str]
    #: The alternative-text description of the PivotTable; ``None`` when unset.
    alt_text_summary: Optional[str]
    #: The grand total label shown in place of the locale's word; ``None`` when unset.
    grand_total_caption: Optional[str]
    #: Whether column widths autofit when the report updates.
    auto_formatting: bool
    #: Whether Show Details is allowed.
    enable_drill: bool


@dataclass(frozen=True)
class PivotCacheOptions:
    """The options of one pivot cache, as read by :meth:`PivotCache.get_options`."""

    #: Whether the source records are saved with the file; off, the records part is left out and the reports keep their
    #: sheet cells.
    save_data: bool
    #: The number of items kept per field after they leave the source, 0 to 1048576; ``None`` is automatic.
    missing_items_limit: Optional[int]
    #: Whether the data refreshes when the file opens.
    refresh_on_load: bool


@dataclass(frozen=True)
class PivotCalculatedItem:
    """One calculated item, as read by :meth:`PivotCache.list_calculated_items`."""

    #: The index of the cache field the item belongs to.
    field_index: int
    #: The item's name.
    name: str
    #: The formula text, leading ``=`` included.
    formula: str
    #: Whether the entry was read from a file and applies its formula only where the item meets items of other fields.
    scoped: bool


@dataclass
class PivotDataFieldInput:
    """Argument shape for :meth:`Pivot.add_data_field` / :meth:`Pivot.set_data_field`."""

    #: The index of the source pivot field the data field aggregates; it must be in range.
    field_index: PivotFieldRef
    #: The display name of the data field, the key GETPIVOTDATA looks it up by. When ``None`` or empty the engine names
    #: it from the workbook profile's pivot wording, such as ``Sum of Sales`` (en-US) or ``合計 / Sales`` (ja-JP),
    #: numbered 2, 3, ... while the name repeats another data field's or a field's display name; that needs
    #: ``field_index`` to be in range.
    name: Optional[str] = None
    #: The aggregation function applied to the source field; ``PivotAggregation.SUM`` by default.
    aggregation: PivotAggregation = PivotAggregation.SUM
    #: The number format code of the data field's values; ``None`` or empty for none.
    number_format: Optional[str] = None
    #: How the value is derived from the data: as is, or as a difference, percentage, running total or rank against a
    #: base.
    show_as: PivotShowValuesAs = PivotShowValuesAs.NORMAL
    #: The index of the base pivot field a relative show-as mode compares against; ``None`` when unset. It must be in
    #: range.
    show_as_base_field: Optional[int] = None
    #: The base item of a relative show-as mode: an item index or a ``PivotShowAsBase`` member; ``None`` when unset.
    show_as_base_item: Optional[int] = None


@dataclass(kw_only=True)
class PivotFilterInput:
    """Argument shape for :meth:`Pivot.add_filter`."""

    #: The index of the pivot field the filter applies to; the filter takes that field's axis.
    field_index: PivotFieldRef
    #: The kind of the filter: a value, label, date or relative-period filter, and its comparison.
    kind: PivotFilterKind
    #: The operand: a number for a value or date filter (an integer count for the top and bottom count kinds), a string
    #: for a label filter, and ``None`` for a relative-period kind. For a range kind this is the lower bound. A ``bool``
    #: is rejected.
    value: Union[int, float, str, None] = None
    #: The upper bound of a range kind, a string for the label ranges; ``None`` for every other kind.
    value_high: Union[int, float, str, None] = None
    #: The index of the data field a value filter scores; label and date filters ignore it. 0 by default.
    data_field_index: int = 0


def _nullable_enum(value: Optional[int], name: str) -> int:
    """The C ordinal of an enum argument whose absence is ``None``; the C ``-1`` stays internal."""
    if value is None:
        return -1
    ordinal = _sint(value, name)
    if ordinal < 0:
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT,
            f"{name} {ordinal} is not a member; pass None for none",
            f"{name}={ordinal}",
        )
    return ordinal


def _local_sheet_id(value: Optional[int]) -> int:
    """The C ``local_sheet_id``: ``-1`` for workbook scope (``None``), else the sheet index."""
    if value is None:
        return -1
    index = _uint(value, "scope", bits=64, role="sheet")
    if index > 0x7FFFFFFF:
        raise FormulonError(
            StatusCode.SHEET_INDEX_OUT_OF_RANGE,
            f"scope {index} is out of range",
            f"scope={index}",
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

    The fields read as in :class:`PivotDataFieldInput`.
    """

    #: The display name of the data field, the key GETPIVOTDATA looks it up by.
    name: str
    #: The index of the pivot field the data field aggregates.
    field_index: int
    #: The aggregation function applied to the source field.
    aggregation: PivotAggregation
    #: The number format code of the data field's values; ``None`` for none.
    number_format: Optional[str]
    #: How the value is derived from the data: as is, or as a difference, percentage, running total or rank against a
    #: base.
    show_as: PivotShowValuesAs
    #: The index of the base pivot field a relative show-as mode compares against; ``None`` when unset.
    show_as_base_field: Optional[int]
    #: The base item of a relative show-as mode: an index into the base field's items, or the previous or next marker;
    #: ``None`` when unset.
    show_as_base_item: Optional[int]


@dataclass(frozen=True, kw_only=True)
class PivotFilter:
    """One stored filter of a pivot, as :meth:`Pivot.list_filters` reads it.

    The fields read as in :class:`PivotFilterInput`.
    """

    #: The index of the filtered pivot field.
    field_index: int
    #: The kind of the filter: a value, label, date or relative-period filter, and its comparison.
    kind: PivotFilterKind
    #: The operand, a number or a string; ``None`` when the filter has no operand. For a range kind this is the lower
    #: bound.
    value: Optional[Union[float, str]]
    #: The upper bound of a range kind; ``None`` when the filter has no upper bound.
    value_high: Optional[Union[float, str]]
    #: The index of the data field a value filter scores; label and date filters ignore it.
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
        raise TypeError(f"{name} must be a number, a str or None, got {type(value).__name__}")
    if count_filter and math.isfinite(value) and value == int(value) and abs(value) < (1 << 31):
        return {name + "_kind": _OPERAND_INT, name + "_int": int(value)}
    return {name + "_kind": _OPERAND_DOUBLE, name + "_double": float(value)}


def _encode_pivot_filter_operands(spec: PivotFilterInput) -> Dict[str, object]:
    """The operand slots of the C filter spec for ``spec.value`` / ``spec.value_high``.

    A text operand lands in ``value_text`` / ``value_high_text``; those two keys are
    returned only when used, so the caller writes each as a string field or leaves it ``NULL``.
    """
    count_filter = spec.kind in _COUNT_FILTER_KINDS
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


_PIVOT_FIELD_PATCH_SPEC = (
    ("custom_name", "custom_name", "str"),
    ("number_format", "number_format", "str"),
    ("default_subtotal", "default_subtotal", "bool"),
    ("insert_blank_row", "insert_blank_row", "bool"),
    ("fill_down_labels", "fill_down_labels", "bool"),
    ("show_all", "show_all", "bool"),
    ("subtotal_top", "subtotal_top", "bool"),
    ("sort_order", "sort_order", "i32"),
    ("sort_by_field", "sort_by_field", "str"),
    ("show_detail", "show_detail", "bool"),
)
_PIVOT_FIELD_PATCH_CLEARABLE = ("custom_name", "number_format", "sort_by_field")
_PIVOT_ITEM_PATCH_SPEC = (
    ("visible", "visible", "bool"),
    ("show_detail", "show_detail", "bool"),
    ("caption", "caption", "str"),
)
_PIVOT_ITEM_PATCH_CLEARABLE = ("caption",)
_PIVOT_STYLE_PATCH_SPEC = (
    ("style_name", "style_name", "str"),
    ("show_row_headers", "show_row_headers", "bool"),
    ("show_column_headers", "show_column_headers", "bool"),
    ("show_row_stripes", "show_row_stripes", "bool"),
    ("show_column_stripes", "show_column_stripes", "bool"),
    ("show_last_column", "show_last_column", "bool"),
)
_PIVOT_STYLE_PATCH_CLEARABLE = ("style_name",)


def _encode_clearable_patch(
    spec: Sequence, clearable: Sequence[str], patch: Mapping, what: str, scope: _ArgScope
) -> Dict[str, object]:
    """Encode a flat patch: a key left out keeps, ``None`` clears a ``clearable`` key, else is a ``TypeError``."""
    for key, value in patch.items():
        if value is None and key not in clearable:
            raise TypeError(f"{key} cannot be None in a {what}")
    return _encode_patch(spec, patch, (), what, scope)


def _pack_pivot_patch(
    layout: S.Struct, spec: Sequence, clearable: Sequence[str], patch: Mapping, what: str, scope: _ArgScope
) -> int:
    """Pack a pivot field, item or style patch into a fresh struct and return its pointer."""
    fields = _encode_clearable_patch(spec, clearable, patch, what, scope)
    ptr = scope.struct(layout)
    layout.pack(LIB, ptr, fields)
    return ptr


def _pack_pivot_field_patch(patch: Mapping, scope: _ArgScope) -> int:
    """An ``fm_pivot_field_patch_t`` for ``fm_sheet_update_pivot_field``."""
    return _pack_pivot_patch(
        S.PIVOT_FIELD_PATCH, _PIVOT_FIELD_PATCH_SPEC, _PIVOT_FIELD_PATCH_CLEARABLE, patch, "pivot field patch", scope
    )


def _pack_pivot_item_patch(patch: Mapping, scope: _ArgScope) -> int:
    """An ``fm_pivot_item_patch_t`` for ``fm_sheet_update_pivot_field_item``."""
    return _pack_pivot_patch(
        S.PIVOT_ITEM_PATCH, _PIVOT_ITEM_PATCH_SPEC, _PIVOT_ITEM_PATCH_CLEARABLE, patch, "pivot item patch", scope
    )


def _pack_pivot_style_options_patch(patch: Mapping, scope: _ArgScope) -> int:
    """An ``fm_pivot_style_options_patch_t`` for ``fm_sheet_update_pivot_style_options``."""
    return _pack_pivot_patch(
        S.PIVOT_STYLE_OPTIONS_PATCH,
        _PIVOT_STYLE_PATCH_SPEC,
        _PIVOT_STYLE_PATCH_CLEARABLE,
        patch,
        "pivot style options patch",
        scope,
    )


def _pack_pivot_spec(scope: _ArgScope, name: str, cache_id: int, anchor_row: int, anchor_col: int) -> int:
    """An ``fm_pivot_spec_t`` for ``fm_sheet_add_pivot``."""
    ptr = scope.struct(S.PIVOT_SPEC)
    S.PIVOT_SPEC.pack(
        LIB,
        ptr,
        {
            "name": scope.cstr(name),
            "cache_id": _uint(cache_id, "cache_id"),
            "anchor_row": _uint(anchor_row, "anchor_row", role="coordinate"),
            "anchor_col": _uint(anchor_col, "anchor_col", role="coordinate"),
        },
    )
    return ptr


def _pack_pivot_date_group(scope: _ArgScope, group: PivotDateGroupInput, *, date1904: bool) -> int:
    """An ``fm_pivot_date_group_t`` for ``fm_sheet_set_pivot_field_date_group``.

    ``group`` is a :class:`PivotDateGroupInput` or a :class:`PivotDateGroup` read back; ``date1904`` is the
    workbook's date system, which turns ``start`` / ``end`` into serials.
    """
    from ._inputs import civil_datetime, date_serial

    fields: Dict[str, object] = {
        "periods": _uint(group.periods, "periods"),
        "calendar": _sint(group.calendar, "calendar"),
        "interval_days": _uint(group.interval_days, "interval_days"),
    }
    for name, bound in (("start", group.start), ("end", group.end)):
        auto = getattr(group, "auto_" + name)
        if auto is not None:
            _require_bool(auto, "group.auto_" + name)
        if auto is False and bound is None:
            raise TypeError(_argument_message("group." + name, "is required"))
        if bound is not None and auto is not True:
            fields["auto_" + name] = _tristate_default(False)
            fields[name] = float(date_serial(civil_datetime(bound, name), date1904=date1904)[0])
        else:
            fields["auto_" + name] = _tristate_default(True)
    ptr = scope.struct(S.PIVOT_DATE_GROUP)
    S.PIVOT_DATE_GROUP.pack(LIB, ptr, fields)
    return ptr


def _pack_pivot_numeric_group(scope: _ArgScope, group: Union[PivotNumericGroupInput, PivotNumericGroup]) -> int:
    """An ``fm_pivot_numeric_group_t`` for ``fm_sheet_set_pivot_field_numeric_group``.

    ``group`` is a :class:`PivotNumericGroupInput` or a :class:`PivotNumericGroup` read back.
    """
    fields: Dict[str, object] = {"interval": _finite(group.interval, "group.interval")}
    for name in ("start", "end"):
        bound = getattr(group, name)
        auto = getattr(group, "auto_" + name)
        if auto is not None:
            _require_bool(auto, "group.auto_" + name)
        if auto is False and bound is None:
            raise TypeError(_argument_message("group." + name, "is required"))
        if bound is not None and auto is not True:
            fields["auto_" + name] = _tristate_default(False)
            fields[name] = _finite(bound, "group." + name)
        else:
            fields["auto_" + name] = _tristate_default(True)
    ptr = scope.struct(S.PIVOT_NUMERIC_GROUP)
    S.PIVOT_NUMERIC_GROUP.pack(LIB, ptr, fields)
    return ptr


def _pivot_worksheet_source_from(workbook: Workbook, d: Dict[str, Any]) -> PivotWorksheetSource:
    """Build a :class:`PivotWorksheetSource` from a decoded ``fm_pivot_worksheet_source_t``."""
    ref, sheet = LIB.read_cstr(d["ref"]), LIB.read_cstr(d["sheet"])
    source_range = None
    index = workbook._find_sheet(sheet) if sheet else None
    if ref and index is not None:
        try:
            source_range = SheetRect(sheet=index, rect=workbook.parse_reference(ref).range)
        except FormulonError:
            source_range = None
    return PivotWorksheetSource(range=source_range, name=LIB.read_cstr(d["name"]) or None)


def _pack_pivot_worksheet_source(
    scope: _ArgScope, workbook: Workbook, source: Union[PivotWorksheetSourceInput, PivotWorksheetSource]
) -> int:
    """An ``fm_pivot_worksheet_source_t`` for ``fm_workbook_set_pivot_cache_worksheet_source``.

    A ``None`` source clears the element: the caller passes a NULL pointer and does not call this.

    The range's sheet is written as a name and the A1 text unqualified, as the element stores them.
    """
    if not isinstance(source, (PivotWorksheetSourceInput, PivotWorksheetSource)):
        raise TypeError(f"expected a PivotWorksheetSourceInput, got {type(source).__name__}")
    given = getattr(source, "sheet", None)
    given_name = None if given is None else (given if isinstance(given, str) else workbook.sheet(given).name)
    value = source.range
    ref: Optional[str] = None
    sheet_name = given_name
    if isinstance(value, SheetRect):
        ref = workbook.format_reference(value.rect)
        sheet_name = workbook._sheet_name(value.sheet)
    elif isinstance(value, str):
        parsed = workbook.parse_reference(value)
        ref = workbook.format_reference(parsed.range)
        if parsed.sheet_name is not None:
            sheet_name = parsed.sheet_name
    elif value is not None:
        raise TypeError(f"range must be a SheetRect, an A1 string or None, got {type(value).__name__}")
    if given_name is not None and sheet_name is not None and given_name.casefold() != sheet_name.casefold():
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT,
            "the range names a different sheet than sheet",
            f"sheet={given_name}",
        )
    ptr = scope.struct(S.PIVOT_WORKSHEET_SOURCE)
    S.PIVOT_WORKSHEET_SOURCE.pack(
        LIB,
        ptr,
        {
            "ref": scope.nonempty_cstr(ref),
            "sheet": scope.nonempty_cstr(sheet_name),
            "name": scope.nonempty_cstr(source.name),
        },
    )
    return ptr


def _cf_match_from(d: Dict[str, Any]) -> CfMatch:
    """Build a :class:`CfMatch` from a decoded ``fm_cf_match_info_t``, keeping the members of its ``kind`` only."""
    kind = _enum(CfMatchKind, d["kind"])
    match = {"kind": kind, "priority": d["priority"]}
    if kind == CfMatchKind.DIFFERENTIAL_FORMAT:
        match["dxf_index"] = d["dxf_index"] if d["has_dxf_index"] else None
    elif kind == CfMatchKind.COLOR_SCALE:
        match["color"] = _color_from(d["color"])
    elif kind == CfMatchKind.DATA_BAR:
        match["bar"] = CfMatchBar(
            length_pct=d["bar_length_pct"],
            axis_position_pct=d["bar_axis_position_pct"],
            is_negative=bool(d["bar_is_negative"]),
            fill=_color_from(d["bar_fill"]),
            border=_optional_color_from(d["bar_border"]),
            gradient=bool(d["bar_gradient"]),
            direction=_enum(CfDataBarDirection, d["bar_direction"]),
        )
    else:
        match["icon"] = CfMatchIcon(icon_set=d["icon_set"], icon_index=d["icon_index"])
    return CfMatch(**match)  # type: ignore[arg-type]


def _cell_xf_fields(record: Union[CellXfInput, CellXf]) -> Dict[str, object]:
    """Project a :class:`CellXfInput` onto the cell XF struct field map.

    Shared by the direct ``<xf>`` writer and the named-style ``<xf>``
    writer, which pack the same struct, so the presence-flag defaulting
    lives in one place.
    """
    if not isinstance(record, (CellXfInput, CellXf)):
        raise TypeError(f"expected a CellXfInput, got {type(record).__name__}")

    return {
        "font_index": _uint(record.font_index, "font_index", role="position"),
        "fill_index": _uint(record.fill_index, "fill_index", role="position"),
        "border_index": _uint(record.border_index, "border_index", role="position"),
        "num_fmt_id": _uint(record.num_fmt_id, "num_fmt_id", 16),
        "horizontal": _uint(record.horizontal or 0, "horizontal", 8),
        "has_horizontal": 0 if record.horizontal is None else 1,
        "vertical": _uint(record.vertical or 0, "vertical", 8),
        "has_vertical": 0 if record.vertical is None else 1,
        "wrap_text": _require_bool(record.wrap_text, "wrap_text") if record.wrap_text is not None else 0,
        "has_wrap_text": 0 if record.wrap_text is None else 1,
        "justify_last_line": (
            _require_bool(record.justify_last_line, "justify_last_line") if record.justify_last_line is not None else 0
        ),
        "has_justify_last_line": 0 if record.justify_last_line is None else 1,
        "has_alignment": 1
        if any(
            value is not None
            for value in (
                record.horizontal,
                record.vertical,
                record.wrap_text,
                record.justify_last_line,
                record.text_rotation,
                record.indent,
                record.relative_indent,
                record.shrink_to_fit,
                record.reading_order,
            )
        )
        else 0,
        "style_xf_index": _uint(record.style_xf_index, "style_xf_index", role="position"),
        "has_text_rotation": 1 if record.text_rotation is not None else 0,
        "text_rotation": _uint(record.text_rotation or 0, "text_rotation"),
        "has_indent": 1 if record.indent is not None else 0,
        "indent": _uint(record.indent or 0, "indent"),
        "has_relative_indent": 1 if record.relative_indent is not None else 0,
        "relative_indent": _sint(record.relative_indent or 0, "relative_indent"),
        "has_shrink_to_fit": 1 if record.shrink_to_fit is not None else 0,
        "shrink_to_fit": (
            _require_bool(record.shrink_to_fit, "shrink_to_fit") if record.shrink_to_fit is not None else 0
        ),
        "has_reading_order": 1 if record.reading_order is not None else 0,
        "reading_order": _sint(record.reading_order or 0, "reading_order"),
        "apply_number_format": _require_bool(record.apply_number_format, "apply_number_format"),
        "apply_font": _require_bool(record.apply_font, "apply_font"),
        "apply_fill": _require_bool(record.apply_fill, "apply_fill"),
        "apply_border": _require_bool(record.apply_border, "apply_border"),
        "apply_alignment": _require_bool(record.apply_alignment, "apply_alignment"),
        "apply_protection": _require_bool(record.apply_protection, "apply_protection"),
        "quote_prefix": _require_bool(record.quote_prefix, "quote_prefix"),
        "has_protection": 1 if record.locked is not None or record.hidden is not None else 0,
        "locked": 0 if record.locked is False else 1,
        "hidden": _require_bool(record.hidden, "hidden") if record.hidden is not None else 0,
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
        """Copy ``data`` into engine memory; an empty ``data`` still gets a 1-byte block, so the pointer is never 0."""
        ptr = self._take(max(len(data), 1))
        if data:
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
        raise TypeError(_argument_message(name, "must be a number"))
    x = float(value)
    if not math.isfinite(x):
        raise FormulonError(StatusCode.INVALID_ARGUMENT, _argument_message(name, "must be finite"), f"{name}={x!r}")
    return x


_ANCHOR_FIELDS = (
    ("from_row", "coordinate"),
    ("from_col", "coordinate"),
    ("from_row_offset_emu", 64),
    ("from_col_offset_emu", 64),
    ("to_row", "coordinate"),
    ("to_col", "coordinate"),
    ("to_row_offset_emu", 64),
    ("to_col_offset_emu", 64),
    ("width_emu", 64),
    ("height_emu", 64),
)


def _anchor_input(anchor: Optional[DrawingAnchorInput]) -> DrawingAnchorInput:
    """``anchor`` itself, or the all-omitted anchor for ``None``; ``TypeError`` for any other type."""
    if anchor is None:
        return DrawingAnchorInput()
    if not isinstance(anchor, DrawingAnchorInput):
        raise TypeError(f"anchor must be a DrawingAnchorInput or None, got {type(anchor).__name__}")
    return anchor


def _anchor_fields(anchor: Optional[DrawingAnchorInput]) -> Dict[str, int]:
    """The ``fm_drawing_anchor_t`` fields of an anchor to add; an omitted field is ``0`` (the default kind)."""
    anchor = _anchor_input(anchor)
    fields: Dict[str, int] = {
        "kind": _sint(AnchorKind.DEFAULT if anchor.kind is None else anchor.kind, "kind"),
        "edit_as": _sint(AnchorEditAs.DEFAULT if anchor.edit_as is None else anchor.edit_as, "edit_as"),
    }
    for key, how in _ANCHOR_FIELDS:
        value = getattr(anchor, key)
        if value is None:
            continue
        fields[key] = _sint(value, key, 64) if how == 64 else _uint(value, key, role=how)
    return fields


def _anchor_patch_fields(anchor: Optional[DrawingAnchorInput]) -> Dict[str, int]:
    """The ``fm_drawing_anchor_patch_t`` fields of an anchor merge; an omitted field is left ``KEEP``."""
    anchor = _anchor_input(anchor)
    fields: Dict[str, int] = {}
    if anchor.kind is not None:
        fields["kind"] = _sint(anchor.kind, "kind")
    if anchor.edit_as is not None:
        fields["edit_as"] = _sint(anchor.edit_as, "edit_as")
    for key, how in _ANCHOR_FIELDS:
        value = getattr(anchor, key)
        if value is not None:
            fields[key] = _sint(value, key, 64) if how == 64 else _uint(value, key, role=how)
    for key in list(fields):
        fields[key + "_op"] = _PATCH_SET
    return fields


def _placement_fields(scope: _ArgScope, placement: Optional[DrawingPlacementInput]) -> Dict[str, Any]:
    """The ``fm_drawing_placement_spec_t`` fields of ``placement``; ``None`` is the all-default placement."""
    if placement is None:
        placement = DrawingPlacementInput()
    elif not isinstance(placement, DrawingPlacementInput):
        raise TypeError(f"placement must be a DrawingPlacementInput or None, got {type(placement).__name__}")
    fields: Dict[str, Any] = {
        "name": scope.nonempty_cstr(placement.name),
        "description": scope.nonempty_cstr(placement.description),
        "anchor": _anchor_fields(placement.anchor),
    }
    if placement.hidden is not None:
        fields["hidden"] = _tristate_default(_require_bool(placement.hidden, "hidden"))
    return fields


def _chart_fill_fields(fill: Optional[ChartFill], scope: _ArgScope) -> Dict[str, Any]:
    """The ``fm_chart_fill_t`` fields of ``fill``; ``None`` is the all-zero record (the style's fill)."""
    if fill is None:
        return {}
    if not isinstance(fill, ChartFill):
        raise TypeError(f"fill must be a ChartFill or None, got {type(fill).__name__}")
    fields: Dict[str, Any] = {"none": _require_bool(fill.none, "none"), "other": _require_bool(fill.other, "other")}
    if fill.color is not None:
        fields["color"] = _color_fields(color_input(fill.color, "fill.color"))
    if fill.transparency_pct is not None:
        fields["has_transparency_pct"] = 1
        fields["transparency_pct"] = _finite(fill.transparency_pct, "fill.transparency_pct")
    return fields


def _chart_line_fields(line: Optional[ChartLine], scope: _ArgScope) -> Dict[str, Any]:
    """The ``fm_chart_line_t`` fields of ``line``; ``None`` is the all-zero record (the style's outline)."""
    if line is None:
        return {}
    if not isinstance(line, ChartLine):
        raise TypeError(f"line must be a ChartLine or None, got {type(line).__name__}")
    fields: Dict[str, Any] = {"none": _require_bool(line.none, "none"), "other": _require_bool(line.other, "other")}
    if line.color is not None:
        fields["color"] = _color_fields(color_input(line.color, "line.color"))
    if line.transparency_pct is not None:
        fields["has_transparency_pct"] = 1
        fields["transparency_pct"] = _finite(line.transparency_pct, "line.transparency_pct")
    if line.width_pt is not None:
        fields["has_width_pt"] = 1
        fields["width_pt"] = _finite(line.width_pt, "line.width_pt")
    if line.dash is not None:
        fields["has_dash"] = 1
        fields["dash"] = _sint(line.dash, "line.dash")
    return fields


def _shape_paint_patch(name: str, value: object, scope: _ArgScope, *, clear: bool) -> Dict[str, object]:
    """The nested ``fill`` / ``line`` patch of a shape patch.

    ``clear`` sets every member to ``CLEAR``, which removes the shape's own element. A :class:`ChartFill` /
    :class:`ChartLine` sets each member it states and clears the ones it leaves ``None``; one that reads
    ``other`` states nothing and keeps the stored paint. A mapping is a :class:`ChartFillPatch` /
    :class:`ChartLinePatch`; ``None`` keeps everything.
    """
    spec = _FILL_PATCH_SPEC if name == "fill" else _LINE_PATCH_SPEC
    if clear:
        return {field_name + "_op": _PATCH_CLEAR for _, field_name, _ in spec}
    if value is None:
        return {}
    record_type = ChartFill if name == "fill" else ChartLine
    if isinstance(value, record_type):
        if value.other:
            return {}
        value = {key: getattr(value, key) for key, _, _ in spec}
    elif not isinstance(value, Mapping):
        raise TypeError(f"{name} must be a {record_type.__name__}, a patch or None, got {type(value).__name__}")
    return _encode_patch(spec, value, (), f"{name} patch", scope)


def _pack_shape_spec(scope: _ArgScope, spec: ShapeInput) -> int:
    """Build an ``fm_shape_spec_t`` from a :class:`ShapeInput`; a ``None`` field is left at its zero default."""
    if not isinstance(spec, ShapeInput):
        raise TypeError(f"input must be a ShapeInput, got {type(spec).__name__}")
    ptr = scope.struct(S.SHAPE_SPEC)
    values: Dict[str, Any] = {
        "placement": _placement_fields(scope, spec),
        "fill": _chart_fill_fields(spec.fill, scope),
        "line": _chart_line_fields(spec.line, scope),
    }
    if spec.rotation_deg is not None:
        values["rotation_deg"] = _finite(spec.rotation_deg, "rotation_deg")
    if spec.text_box is not None:
        values["text_box"] = _require_bool(spec.text_box, "text_box")
    values["preset_geometry"] = scope.nonempty_cstr(spec.preset_geometry)
    values["text"] = scope.nonempty_cstr(spec.text)
    S.SHAPE_SPEC.pack(LIB, ptr, values)
    return ptr


def _pack_shape_patch(
    scope: _ArgScope,
    preset_geometry,
    text,
    text_box,
    fill,
    line,
    rotation_deg,
    clear: Sequence[str] = (),
) -> int:
    """Build an ``fm_shape_patch_t``; a ``None`` argument is a field left ``KEEP``, a name in ``clear``
    (``"fill"`` or ``"line"``) removes the shape's own paint.

    ``fill`` / ``line`` are a :class:`ChartFill` / :class:`ChartLine` or a patch mapping (see
    :func:`_shape_paint_patch`); ``rotation_deg`` is in degrees.
    """
    ptr = scope.struct(S.SHAPE_PATCH)
    values: Dict[str, Any] = {
        "fill": _shape_paint_patch("fill", fill, scope, clear="fill" in clear),
        "line": _shape_paint_patch("line", line, scope, clear="line" in clear),
    }
    if rotation_deg is not None:
        values["rotation_deg_op"] = _PATCH_SET
        values["rotation_deg"] = _finite(rotation_deg, "rotation_deg")
    if text_box is not None:
        values["text_box_op"] = _PATCH_SET
        values["text_box"] = _require_bool(text_box, "text_box")
    if preset_geometry is not None:
        values["preset_geometry_op"] = _PATCH_SET
        values["preset_geometry"] = scope.cstr(preset_geometry)
    if text is not None:
        values["text_op"] = _PATCH_SET
        values["text"] = scope.cstr(text)
    S.SHAPE_PATCH.pack(LIB, ptr, values)
    return ptr


def _pack_slicer_spec(scope: _ArgScope, spec: SlicerInput) -> int:
    """Build an ``fm_slicer_spec_t`` from a :class:`SlicerInput`; a ``None`` field is left at its zero default."""
    if not isinstance(spec, SlicerInput):
        raise TypeError(f"input must be a SlicerInput, got {type(spec).__name__}")
    ptr = scope.struct(S.SLICER_SPEC)
    values: Dict[str, Any] = {"placement": _placement_fields(scope, spec)}
    if spec.column_count is not None:
        values["column_count"] = _uint(spec.column_count, "column_count")
    if spec.sort_order is not None:
        values["sort_order"] = _sint(spec.sort_order, "sort_order")
    if spec.cross_filter is not None:
        values["cross_filter"] = _sint(spec.cross_filter, "cross_filter")
    if spec.row_height_pt is not None:
        values["row_height_pt"] = _finite(spec.row_height_pt, "row_height_pt")
    values["caption"] = scope.nonempty_cstr(spec.caption)
    values["style_name"] = scope.opt_cstr(spec.style_name)
    S.SLICER_SPEC.pack(LIB, ptr, values)
    return ptr


def _pack_slicer_patch(
    scope: _ArgScope,
    caption,
    style_name,
    column_count,
    row_height_pt,
    sort_order,
    cross_filter,
    clear: Sequence[str] = (),
) -> int:
    """Build an ``fm_slicer_patch_t``; a ``None`` argument is a field left ``KEEP``, a name in ``clear`` is ``CLEAR``."""
    ptr = scope.struct(S.SLICER_PATCH)
    values: Dict[str, Any] = {}
    if column_count is not None:
        values["column_count"] = _uint(column_count, "column_count")
    if row_height_pt is not None:
        values["row_height_pt"] = _finite(row_height_pt, "row_height_pt")
    if sort_order is not None:
        values["sort_order"] = _sint(sort_order, "sort_order")
    if cross_filter is not None:
        values["cross_filter"] = _sint(cross_filter, "cross_filter")
    for key, text_value in (("caption", caption), ("style_name", style_name)):
        if text_value is not None:
            values[key] = scope.cstr(text_value)
    for key in list(values):
        values[key + "_op"] = _PATCH_SET
    for key in clear:
        values[key + "_op"] = _PATCH_CLEAR
    S.SLICER_PATCH.pack(LIB, ptr, values)
    return ptr


def _pack_drawing_patch(scope: _ArgScope, name, description, hidden, anchor, clear: Sequence[str] = ()) -> int:
    """Build an ``fm_drawing_patch_t``; a ``None`` argument is a field left ``KEEP``, a name in ``clear`` is ``CLEAR``.

    ``anchor`` is a :class:`DrawingAnchorInput` merged into the object's anchor.
    """
    ptr = scope.struct(S.DRAWING_PATCH)
    values: Dict[str, Any] = {}
    if name is not None:
        values["name"] = scope.cstr(name)
        values["name_op"] = _PATCH_SET
    if description is not None:
        values["description"] = scope.cstr(description)
        values["description_op"] = _PATCH_SET
    if hidden is not None:
        values["hidden"] = _require_bool(hidden, "hidden")
        values["hidden_op"] = _PATCH_SET
    for key in clear:
        values[key + "_op"] = _PATCH_CLEAR
    if anchor is not None:
        values["anchor"] = _anchor_patch_fields(anchor)
    S.DRAWING_PATCH.pack(LIB, ptr, values)
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


def _qualified_range_text(
    workbook: Workbook, value: Union[SheetRect, str, None], sheet: int, label: str, stored: Optional[str] = None
) -> Optional[str]:
    """The sheet-qualified A1 text of a ``SheetRect`` or A1 string for the engine; ``None`` is ``stored`` (or ``None``).

    An unqualified string resolves on sheet index ``sheet``.
    """
    if value is None:
        return stored
    from ._inputs import sheet_rect

    resolved = sheet_rect(workbook, value, sheet, label)
    return workbook.format_reference(resolved.rect, sheet_name=workbook._sheet_name(resolved.sheet))


def _sparkline_range_text(
    workbook: Workbook, value: Union[SheetRect, str, None], sheet: int, label: str, stored: Optional[str] = None
) -> Optional[str]:
    """``_qualified_range_text`` for a sparkline source or date range, except that text which is not one rectangle on
    a sheet of this workbook (a 3-D span, another workbook's range, a defined name, ``#REF!``) goes to the engine as
    written; the engine accepts or refuses it."""
    if isinstance(value, str):
        try:
            ref = workbook.parse_reference(value.lstrip("="))
        except FormulonError:
            return value
        if ref.sheet_name is not None and workbook._find_sheet(ref.sheet_name) is None:
            return value
    return _qualified_range_text(workbook, value, sheet, label, stored)


def _parse_qualified_range(workbook: Workbook, text: str, sheet: int) -> Optional[SheetRect]:
    """The :class:`SheetRect` a stored range text names (unqualified: on ``sheet``); ``None`` if it does not resolve."""
    try:
        ref = workbook.parse_reference(text.lstrip("="))
    except FormulonError:
        return None
    if ref.sheet_name is None:
        return SheetRect(sheet=sheet, rect=ref.range)
    index = workbook._find_sheet(ref.sheet_name)
    return None if index is None else SheetRect(sheet=index, rect=ref.range)


def _decode_sparkline(workbook: Workbook, sheet: int, d: Dict[str, Any]) -> Sparkline:
    """Build a :class:`Sparkline` from a decoded ``fm_sparkline_t``."""
    text = LIB.read_cstr(d["formula"])
    location = workbook.parse_reference(LIB.read_cstr(d["location"]))
    item = Sparkline(
        source=_parse_qualified_range(workbook, text, sheet),
        location=CellAddress(row=location.range.first_row, col=location.range.first_col),
    )
    if item.source is None:
        object.__setattr__(item, "_stored_source", text)
    return item


def _decode_sparkline_date_axis(workbook: Workbook, sheet: int, text: str) -> Tuple[Optional[SheetRect], Optional[str]]:
    """``(range, stored text)`` of a group's date axis; the text is kept only when it does not parse."""
    if not text:
        return None, None
    parsed = _parse_qualified_range(workbook, text, sheet)
    return parsed, (None if parsed is not None else text)


def _pack_sparkline_group(
    scope: _ArgScope, group: Union[SparklineGroup, SparklineGroupInput], workbook: Workbook, sheet: int
) -> int:
    """Build a sparkline group input struct from ``group`` (``sparklines`` is passed separately).

    ``sheet`` is the index of the receiving sheet, which resolves an unqualified ``date_axis_range``.
    """
    ptr = scope.struct_array(S.SPARKLINE_GROUP, 1)
    S.SPARKLINE_GROUP.pack(
        LIB,
        ptr,
        {
            "type": _sint(group.type, "type"),
            "markers": _require_bool(group.markers, "markers"),
            "high": _require_bool(group.high, "high"),
            "low": _require_bool(group.low, "low"),
            "first": _require_bool(group.first, "first"),
            "last": _require_bool(group.last, "last"),
            "negative": _require_bool(group.negative, "negative"),
            "display_x_axis": _require_bool(group.display_x_axis, "display_x_axis"),
            "display_hidden": _require_bool(group.display_hidden, "display_hidden"),
            "right_to_left": _require_bool(group.right_to_left, "right_to_left"),
            "min_axis_type": _sint(group.min_axis_type, "min_axis_type"),
            "max_axis_type": _sint(group.max_axis_type, "max_axis_type"),
            "has_manual_min": 0 if group.manual_min is None else 1,
            "has_manual_max": 0 if group.manual_max is None else 1,
            "manual_min": 0.0
            if group.manual_min is None
            else _require_number(group.manual_min, "manual_min", finite=False),
            "manual_max": 0.0
            if group.manual_max is None
            else _require_number(group.manual_max, "manual_max", finite=False),
            "date_axis": _require_bool(group.date_axis, "date_axis"),
            "line_width_pt": _require_number(group.line_width_pt, "line_width_pt", finite=False),
            "empty_cells_as": _sint(group.empty_cells_as, "empty_cells_as"),
        },
    )
    for key in _SPARKLINE_COLORS:
        _pack_optional_color(ptr + S.SPARKLINE_GROUP.offsets[key][1], getattr(group, key), f"SparklineGroup.{key}")
    date_axis = _sparkline_range_text(
        workbook,
        group.date_axis_range,
        sheet,
        "SparklineGroup.date_axis_range",
        getattr(group, "_stored_date_axis", None),
    )
    S.write_str_field(LIB, ptr, S.SPARKLINE_GROUP, "date_axis_ref", date_axis, scope)
    return ptr


def _pack_sparklines(
    scope: _ArgScope, sparklines: Sequence[Union[SparklineInput, Sparkline]], workbook: Workbook, sheet: int
) -> int:
    """Build a sparkline struct array; an empty list still gets a valid block (the core rejects a zero count).

    ``sheet`` is the index of the receiving sheet, which resolves an unqualified ``source``.
    """
    from ._inputs import _cell_address

    ptr = scope.struct_array(S.SPARKLINE, max(len(sparklines), 1))
    for i, item in enumerate(sparklines):
        base = ptr + i * S.SPARKLINE.size
        source = _sparkline_range_text(
            workbook, item.source, sheet, "Sparkline.source", getattr(item, "_stored_source", None)
        )
        if source is None:
            # No source is a sparkline whose source was deleted, stored as `#REF!` is.
            source = ""
        location = _cell_address(workbook, item.location, "Sparkline.location")
        location_text = workbook.format_reference(CellRect(location.row, location.col, location.row, location.col))
        S.write_str_field(LIB, base, S.SPARKLINE, "formula", source, scope)
        S.write_str_field(LIB, base, S.SPARKLINE, "location", location_text, scope)
    return ptr


# ---------------------------------------------------------------------------
# Chart marshalling
# ---------------------------------------------------------------------------

# `fm_patch_op_t`: a zeroed patch field keeps, so only CLEAR and SET are written.
_PATCH_CLEAR = 1
_PATCH_SET = 2

# The C "no title" / "no legend" members; the binding reports them as a ``None`` record.
_CHART_TITLE_NONE = 0
_CHART_LEGEND_NONE = -1
# The series roles of `fm_chart_role_t` in order, which is also the order of `fm_chart_data_series_info_t.range_count`.
_CHART_ROLES = ("name", "categories", "values", "x", "y", "size")

# (patch key, C field, scalar kind) per patch struct.
_FILL_PATCH_SPEC = (
    ("color", "color", "color"),
    ("transparency_pct", "transparency_pct", "f64"),
    ("none", "none", "bool"),
)
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
    ("source_linked", "source_linked", "bool"),
    ("deleted", "deleted", "bool"),
)
_PLOT_PATCH_SPEC = (
    ("gap_width_pct", "gap_width_pct", "u32"),
    ("overlap_pct", "overlap_pct", "i32"),
    ("first_slice_angle_deg", "first_slice_angle_deg", "u32"),
    ("vary_colors", "vary_colors", "bool"),
    ("hole_size_pct", "hole_size_pct", "u32"),
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
    ("minor_unit", "minor_unit", "f64"),
    ("minor_gridlines", "minor_gridlines", "bool"),
    ("position", "position", "i32"),
    ("source_linked", "source_linked", "bool"),
)
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
    ("style", "style", "i32"),
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


def _tri(value: int) -> Optional[bool]:
    """Decode a chart tri-state (absent / false / true)."""
    return None if value == S.TRISTATE_DEFAULT else S.tristate_flag(value)


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
        kind=_enum(ChartTitleKind, d["kind"]),
        text=LIB.read_optional_cstr(d["text"]),
        formula=LIB.read_optional_cstr(d["formula"]),
    )


def _read_chart_fill(d: Dict[str, object]) -> ChartFill:
    return ChartFill(
        color=_optional_color_from(d["color"]),
        transparency_pct=d["transparency_pct"] if d["has_transparency_pct"] else None,
        none=bool(d["none"]),
        other=bool(d["other"]),
    )


def _read_chart_line(d: Dict[str, object]) -> ChartLine:
    return ChartLine(
        color=_optional_color_from(d["color"]),
        transparency_pct=d["transparency_pct"] if d["has_transparency_pct"] else None,
        width_pt=d["width_pt"] if d["has_width_pt"] else None,
        dash=_enum(ChartLineDash, d["dash"]) if d["has_dash"] else None,
        none=bool(d["none"]),
        other=bool(d["other"]),
    )


def _read_chart_marker(d: Dict[str, object]) -> ChartMarker:
    return ChartMarker(
        symbol=_enum(ChartMarkerSymbol, d["symbol"]),
        size_pt=d["size_pt"] if d["has_size_pt"] else None,
        fill=_shape_paint_read(_read_chart_fill(d["fill"])),
        line=_shape_paint_read(_read_chart_line(d["line"])),
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
        font_size_pt=d["font_size_pt"] if d["has_font_size_pt"] else None,
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
                fill=_shape_paint_read(_read_chart_fill(p["fill"])),
                line=_shape_paint_read(_read_chart_line(p["line"])),
                marker=_read_chart_marker(p["marker"]) if p["has_marker"] else None,
            )
        )
    return ChartSeries(
        series_id=d["series_id"],
        order=d["order"],
        filtered=bool(d["filtered"]),
        fill=_shape_paint_read(_read_chart_fill(d["fill"])),
        line=_shape_paint_read(_read_chart_line(d["line"])),
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
        overlap_pct=d["overlap_pct"] if d["has_overlap_pct"] else None,
        hole_size_pct=d["hole_size_pct"] if d["has_hole_size_pct"] else None,
        first_slice_angle_deg=d["first_slice_angle_deg"] if d["has_first_slice_angle_deg"] else None,
        axis_ids=[
            _read_count(LIB.fm_chart_get_plot_axis_id_at, chart, _uint(plot, "plot"), _uint(i, "idx", role="position"))
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
        crosses_at=d["crosses_at"] if d["has_crosses_at"] else None,
        min=d["min"] if d["has_min"] else None,
        max=d["max"] if d["has_max"] else None,
        major_unit=d["major_unit"] if d["has_major_unit"] else None,
        minor_unit=d["minor_unit"] if d["has_minor_unit"] else None,
        log_base=d["log_base"] if d["has_log_base"] else None,
        reversed=bool(d["reversed"]),
        number_format=_opt_cstr(d["number_format"]),
        source_linked=bool(d["source_linked"]),
        title=_read_chart_title(d["title"]),
        major_gridlines=bool(d["major_gridlines"]),
        minor_gridlines=bool(d["minor_gridlines"]),
        tick_label_position=_enum_or_none(ChartTickLabelPosition, d["tick_label_position"]),
    )


def _read_chart(chart: int, workbook: Workbook, sheet: int) -> ChartProperties:
    """Decode an open chart handle; ``workbook`` and ``sheet`` resolve the source range text."""
    d = _read_struct(LIB.fm_chart_get_info, S.CHART_INFO, chart)
    source = None
    source_text = LIB.read_cstr(d["source_range"]) if d["source_range"] else ""
    source_range = _parse_qualified_range(workbook, source_text, sheet) if source_text else None
    if source_range is not None:
        source = ChartSource(
            range=source_range,
            direction=_enum(ChartSeriesDirection, d["source_direction"]),
            header_row=bool(d["source_header_row"]),
            header_column=bool(d["source_header_column"]),
        )
    plots = [_read_chart_plot(chart, i) for i in range(d["plot_count"])]
    return ChartProperties(
        object_id=d["object_id"],
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
            hidden_entries=[
                _read_count(LIB.fm_chart_get_legend_hidden_at, chart, _uint(i, "idx", role="position"))
                for i in range(d["legend_hidden_count"])
            ],
        ),
        blanks_as=_enum_or_none(BlanksAs, d["blanks_as"]),
        plot_visible_only=_tri(d["plot_visible_only"]),
        na_as_blank=_tri(d["na_as_blank"]),
        date1904=bool(d["date1904"]),
        style=d["style"] if d["has_style"] else None,
        plots=plots,
        axes=[_read_chart_axis(chart, i) for i in range(d["axis_count"])],
        filter=ChartFilter(
            hidden_series=[series.series_id for plot in plots for series in plot.series if series.filtered],
            hidden_categories=[
                _read_count(LIB.fm_chart_get_hidden_category_at, chart, _uint(i, "idx", role="position"))
                for i in range(d["hidden_category_count"])
            ],
        ),
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
    values: List[Optional[float]] = []
    with _ArgScope() as scope:
        label_out = scope.out_ptr()
        has_value_out = scope.out_ptr()
        value_out = scope.alloc(8)
        for i in range(info["category_label_count"]):
            _check(
                LIB.fm_chart_data_get_category_at(
                    data, _uint(i, "idx", role="position"), label_out, has_value_out, value_out
                )
            )
            labels.append(LIB.read_cstr(LIB.read_u32(label_out)))
            values.append(LIB.read_f64(value_out) if LIB.read_i32(has_value_out) else None)
        levels: List[List[str]] = []
        for level in range(info["category_level_count"]):
            size = _read_count(LIB.fm_chart_data_get_category_level_count, data, _uint(level, "level"))
            names: List[str] = []
            for i in range(size):
                _check(
                    LIB.fm_chart_data_get_category_level_at(
                        data, _uint(level, "level"), _uint(i, "idx", role="position"), label_out
                    )
                )
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
                series_id=s["series_id"],
                name=_opt_cstr(s["name"]),
                number_format=LIB.read_cstr(s["number_format"]),
                x_number_format=LIB.read_optional_cstr(s["x_number_format"]),
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
            number_format=LIB.read_optional_cstr(info["category_number_format"]),
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
            ChartUnresolved(series_id=u["series_id"], role=_enum(ChartRole, u["role"]), axis_id=u["axis_id"])
            for u in (
                _read_struct(LIB.fm_chart_data_get_unresolved_at, S.CHART_UNRESOLVED, data, i)
                for i in range(info["unresolved_count"])
            )
        ],
    )


def _title_fields(title: Union[str, ChartTitle], scope: _ArgScope) -> Dict[str, object]:
    """Encode a title given as its text or as a :class:`ChartTitle`."""
    if isinstance(title, ChartTitle):
        if title.kind == ChartTitleKind.TEXT and title.text is None:
            raise FormulonError(StatusCode.INVALID_ARGUMENT, "a TEXT title needs text", "field=text")
        return {
            "kind": _sint(title.kind, "kind"),
            "text": scope.opt_cstr(title.text),
            "formula": scope.opt_cstr(title.formula),
        }
    return {"kind": 1, "text": scope.cstr(title)}


def _patch_scalar(kind: str, value, key: str, scope: _ArgScope, path: str = ""):
    if path:
        key = f"{path}.{key}"
    if kind == "u32":
        return _uint(value, key)
    if kind == "color":
        return _color_fields(color_input(value, key))
    if kind == "i32":
        return _sint(value, key)
    if kind == "bool":
        return _require_bool(value, key)
    if kind == "f64":
        return _require_number(value, key, finite=False)
    if kind == "title":
        return _title_fields(value, scope)
    return scope.cstr(value)


def _check_patch_keys(patch: Mapping, spec: Sequence, extra: Sequence[str], what: str, path: str = "") -> None:
    unknown = set(patch) - {key for key, _, _ in spec} - set(extra)
    if unknown and path:
        raise TypeError(_argument_message(path, f"has no field {sorted(unknown)[0]}"))
    if unknown:
        raise TypeError(f"unknown {what} key(s): {sorted(unknown)}")


def _encode_patch(
    spec: Sequence, patch: Mapping, extra: Sequence[str], what: str, scope: _ArgScope, path: str = ""
) -> Dict[str, object]:
    """Encode the ``(op, value)`` pair of every ``spec`` key present in ``patch``.

    An absent key stays zero (keep), ``None`` writes the clear op and any other
    value the set op beside it. A non-empty ``path`` names the patch (an element of an array
    argument) in the messages of its refusals.
    """
    _check_patch_keys(patch, spec, extra, what, path)
    out: Dict[str, object] = {}
    for key, field_name, kind in spec:
        if key not in patch:
            continue
        value = patch[key]
        op_field = field_name + "_op"
        if value is None:
            out[op_field] = _PATCH_CLEAR
        else:
            out[op_field] = _PATCH_SET
            out[field_name] = _patch_scalar(kind, value, key, scope, path)
    return out


_TABLE_COLUMN_PATCH_SPEC = (
    ("totals_row_function", "totals_row_function", "i32"),
    ("totals_row_label", "totals_row_label", "str"),
    ("totals_row_formula", "totals_row_formula", "str"),
    ("calculated_formula", "calculated_formula", "str"),
)


def _encode_table_column_patch(
    entry: Union[TableColumnPatch, TableColumn], scope: _ArgScope, index: int
) -> Dict[str, object]:
    """Encode element ``index`` of ``Table.update``'s ``columns``: a patch mapping, or a read :class:`TableColumn`."""
    if isinstance(entry, TableColumn):
        entry = {key: getattr(entry, key) for key, _, _ in _TABLE_COLUMN_PATCH_SPEC}
    elif not isinstance(entry, Mapping):
        raise TypeError(_argument_message("columns", "must be a sequence of TableColumnPatch"))
    return _encode_patch(_TABLE_COLUMN_PATCH_SPEC, entry, ("name",), "column patch", scope, f"columns[{index}]")


def _encode_paint_patches(out: Dict[str, object], patch: Mapping, scope: _ArgScope) -> None:
    """Encode the nested ``fill`` / ``line`` of ``patch``: an absent key keeps the paint, ``None`` removes it,
    and a record or patch mapping applies as in :func:`_shape_paint_patch`."""
    for name in ("fill", "line"):
        out[name] = _shape_paint_patch(name, patch.get(name), scope, clear=name in patch and patch[name] is None)


def _encode_marker_patch(patch: Mapping, scope: _ArgScope) -> Dict[str, object]:
    out = _encode_patch(_MARKER_PATCH_SPEC, patch, ("fill", "line"), "marker patch", scope)
    _encode_paint_patches(out, patch, scope)
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


def _data_source_spec(value: object, key: str, scope: _ArgScope) -> Dict[str, object]:
    """Encode a series input as an ``fm_chart_data_source_spec_t``: reference text, or a :class:`ChartDataSource`."""
    if isinstance(value, str):
        return {"kind": ChartSourceKind.REF, "formula": scope.cstr(value)}
    if not isinstance(value, ChartDataSource):
        raise TypeError(f"{key} must be a str, a ChartDataSource or None, got {type(value).__name__}")
    kind = _enum(ChartSourceKind, value.kind)
    if kind == ChartSourceKind.REF:
        return {"kind": int(kind), "formula": scope.cstr(value.formula)}
    if kind == ChartSourceKind.NONE:
        return {"kind": int(kind)}
    points = list(value.literal)
    array = scope.struct_array(S.CHART_LITERAL, len(points))
    for i, point in enumerate(points):
        is_text = isinstance(point, str)
        if not is_text and (isinstance(point, bool) or not isinstance(point, numbers.Real)):
            raise TypeError(f"{key} literal points must be numbers or strings, got {type(point).__name__}")
        S.CHART_LITERAL.pack(
            LIB,
            array + i * S.CHART_LITERAL.size,
            {
                "idx": _uint(i, "idx", role="position"),
                "is_text": 1 if is_text else 0,
                "number": 0.0 if is_text else float(point),
                "text": scope.cstr(point) if is_text else 0,
            },
        )
    return {"kind": int(kind), "literal": array, "literal_count": len(points)}


def _encode_series_patch(patch: Mapping, scope: _ArgScope) -> Dict[str, object]:
    unknown = set(patch) - set(_SERIES_REF_KEYS) - {"fill", "line", "marker", "smooth", "labels"}
    if unknown:
        raise TypeError(f"unknown series patch key(s): {sorted(unknown)}")
    out: Dict[str, object] = {}
    for key in _SERIES_REF_KEYS:
        if key not in patch:
            continue
        if patch[key] is None:
            out[key + "_op"] = _PATCH_CLEAR
        else:
            out[key + "_op"] = _PATCH_SET
            out[key] = _data_source_spec(patch[key], key, scope)
    _encode_paint_patches(out, patch, scope)
    _encode_optional_patch(out, "marker", patch, _encode_marker_patch, scope)
    _encode_optional_patch(out, "labels", patch, _encode_labels_patch, scope)
    if "smooth" in patch:
        out["smooth_op"] = _PATCH_CLEAR if patch["smooth"] is None else _PATCH_SET
        out["smooth"] = 0 if patch["smooth"] is None else _require_bool(patch["smooth"], "smooth")
    return out


def _encode_point_patch(patch: Optional[Mapping], scope: _ArgScope) -> Dict[str, object]:
    """Encode a data-point patch; an empty patch keeps everything and ``None`` clears every member.

    A point whose override ends empty loses its ``c:dPt``, so ``None`` is how a point override is removed.
    """
    if patch is None:
        patch = {"fill": None, "line": None, "marker": None}
    unknown = set(patch) - {"fill", "line", "marker"}
    if unknown:
        raise TypeError(f"unknown point patch key(s): {sorted(unknown)}")
    out: Dict[str, object] = {}
    _encode_paint_patches(out, patch, scope)
    _encode_optional_patch(out, "marker", patch, _encode_marker_patch, scope)
    return out


def _encode_axis_patch(patch: Mapping, scope: _ArgScope) -> Dict[str, object]:
    out = _encode_patch(_AXIS_PATCH_SPEC, patch, ("crosses_at",), "axis patch", scope)
    if "crosses_at" in patch:
        out["crosses_at_op"] = _PATCH_SET
        out["crosses_at"] = _require_number(patch["crosses_at"], "crosses_at", finite=False)
    return out


def _header_mode(flag: Optional[bool]) -> int:  # noqa: FBT001 -- private tri-state encoder
    """Map an optional header flag to ``fm_tristate_t``: ``None`` is Excel's detection."""
    return _tristate_default(flag)


def _chart_range_text(workbook: Workbook, value: Union[SheetRect, str], sheet: int, label: str) -> str:
    """The sheet-qualified A1 text of a chart range; an unqualified string resolves on sheet index ``sheet``."""
    text = _qualified_range_text(workbook, value, sheet, label)
    if text is None:
        raise TypeError(f"{label} must be a SheetRect or an A1 string")
    return text


def _chart_source_fields(
    workbook: Workbook, sheet: int, source: Union[ChartSourceInput, ChartSource]
) -> Tuple[str, int, int, int]:
    """``(range text, direction, header_row, header_column)`` of a chart source, for ``fm_sheet_set_chart_source``.

    A :class:`ChartSource` read back passes as is; its headers are explicit.
    """
    if not isinstance(source, (ChartSourceInput, ChartSource)):
        raise TypeError(f"expected a ChartSourceInput, got {type(source).__name__}")
    return (
        _chart_range_text(workbook, source.range, sheet, "ChartSourceInput.range"),
        _sint(_enum(ChartSeriesDirection, source.direction), "direction"),
        _sint(_header_mode(source.header_row), "header_row"),
        _sint(_header_mode(source.header_column), "header_column"),
    )


def _encode_chart_insert(scope: _ArgScope, workbook: Workbook, sheet: int, insert: ChartInput) -> int:
    """Pack ``insert`` into an ``fm_chart_spec_t`` and return its pointer.

    The name, description, visibility and anchor of ``insert`` become the embedded
    ``fm_drawing_placement_spec_t``. ``bar_direction`` on a plot type other than ``BAR`` is a ``TypeError``.
    """
    if not isinstance(insert, ChartInput):
        raise TypeError(f"expected a ChartInput, got {type(insert).__name__}")
    plot_type = _enum(ChartPlotType, insert.type)
    if insert.bar_direction is not None and plot_type != ChartPlotType.BAR:
        raise TypeError(f"bar_direction applies to a BAR chart only, not {plot_type.name}")
    title = insert.title
    auto_title = isinstance(title, ChartTitle) and title.kind == ChartTitleKind.AUTO
    if title is None:
        title_fields = {"kind": _sint(_CHART_TITLE_NONE, "kind")}
    else:
        title_fields = _title_fields(title, scope)
    ptr = scope.struct(S.CHART_SPEC)
    S.CHART_SPEC.pack(
        LIB,
        ptr,
        {
            "type": _sint(plot_type, "type"),
            "bar_direction": _sint(
                ChartBarDirection.COLUMN if insert.bar_direction is None else insert.bar_direction, "bar_direction"
            ),
            "grouping": _sint(insert.grouping, "grouping"),
            "source": scope.cstr(_chart_range_text(workbook, insert.range, sheet, "ChartInput.range")),
            "direction": _sint(insert.direction, "direction"),
            "header_row": _sint(_header_mode(insert.header_row), "header_row"),
            "header_column": _sint(_header_mode(insert.header_column), "header_column"),
            "has_title": 0 if auto_title else 1,
            "title": {} if auto_title else title_fields,
            "legend": _sint(_nullable_enum(insert.legend, "legend"), "legend"),
            "placement": _placement_fields(scope, insert),
        },
    )
    return ptr


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


def _write_value_input(value: Value, value_ptr: int, scope: _ArgScope) -> None:
    """Write ``value`` as a value struct at ``value_ptr``.

    ``value`` is a :class:`Value`, so its payload matches its
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


def _pack_mentions(mentions: Sequence[Union[MentionInput, Mention]], scope: _ArgScope) -> int:
    """Pack ``mentions`` into a contiguous mention struct array (0 when empty), spans in code points.

    A ``None`` or empty ``mention_id`` is left ``NULL`` and takes an engine-generated id.
    """
    if not mentions:
        return 0
    base = scope.struct_array(S.MENTION, len(mentions))
    for i, m in enumerate(mentions):
        if not isinstance(m, (MentionInput, Mention)):
            raise TypeError(f"expected a MentionInput, got {type(m).__name__}")
        ptr = base + i * S.MENTION.size
        S.MENTION.pack(LIB, ptr, {"start_cp": _uint(m.start, "start"), "end_cp": _uint(m.end, "end")})
        S.write_str_field(LIB, ptr, S.MENTION, "person_id", m.person_id, scope)
        S.write_str_field(LIB, ptr, S.MENTION, "mention_id", m.mention_id, scope)
    return base


def _write_mentions(
    struct_ptr: int, layout: S.Struct, mentions: Sequence[Union[MentionInput, Mention]], scope: _ArgScope
) -> None:
    """Store the mention array and its count into ``mentions`` / ``mentions_count``."""
    packed = _pack_mentions(mentions, scope)
    LIB.write_bytes(struct_ptr + layout.offsets["mentions"][1], struct.pack("<I", packed))
    LIB.write_bytes(struct_ptr + layout.offsets["mentions_count"][1], struct.pack("<I", len(mentions)))


def _decode_mentions(base: int, count: int) -> List[Mention]:
    out: List[Mention] = []
    for i in range(count):
        d = S.MENTION.unpack(LIB, base + i * S.MENTION.size)
        out.append(
            Mention(
                person_id=LIB.read_cstr(d["person_id"]),
                mention_id=LIB.read_cstr(d["mention_id"]),
                start=d["start_cp"],
                end=d["end_cp"],
            )
        )
    return out


def _require_cell_rect(r: object) -> CellRect:
    if not isinstance(r, CellRect):
        raise TypeError(f"expected a CellRect, got {type(r).__name__}")
    return r


def _cell_rect_fields(r: CellRect) -> Dict[str, int]:
    """The cell rect struct field values of ``r``; anything but a :class:`CellRect` is a ``TypeError``."""
    r = _require_cell_rect(r)
    return {
        "first_row": _uint(r.first_row, "first_row", role="coordinate"),
        "first_col": _uint(r.first_col, "first_col", role="coordinate"),
        "last_row": _uint(r.last_row, "last_row", role="coordinate"),
        "last_col": _uint(r.last_col, "last_col", role="coordinate"),
    }


def _cell_rect_from(d: Dict[str, int]) -> CellRect:
    return CellRect(d["first_row"], d["first_col"], d["last_row"], d["last_col"])


def _spill_from(d: Dict[str, Any]) -> Spill:
    return Spill(
        anchor=CellAddress(row=d["anchor"]["row"], col=d["anchor"]["col"]),
        blocked=bool(d["blocked"]),
        cause=_enum_or_none(SpillCause, d["cause"]),
        range=_cell_rect_from(d["rect"]),
    )


def _hyperlink_from(d: Dict[str, Any]) -> Hyperlink:
    return Hyperlink(
        range=_cell_rect_from(d["rect"]),
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
    if isinstance(value, datetime.datetime) and value.tzinfo is not None:
        value = value.astimezone(datetime.timezone.utc).replace(tzinfo=None)
    return civil_datetime(value, "created").strftime("%Y-%m-%dT%H:%M:%S") + ".00"


def _threaded_comment_from(d: Dict[str, Any]) -> ThreadedComment:
    text = LIB.read_cstr(d["text"])
    return ThreadedComment(
        id=LIB.read_cstr(d["id"]),
        cell=CellAddress(row=d["row"], col=d["col"]),
        person_id=LIB.read_cstr(d["person_id"]),
        created=_parse_created(LIB.read_cstr(d["created"])),
        text=text,
        parent_id=LIB.read_cstr(d["parent_id"]) or None,
        done=bool(d["done"]),
        mentions=_decode_mentions(d["mentions"], d["mentions_count"]),
    )


def _drawing_anchor_from(d: Dict[str, Any]) -> DrawingAnchor:
    return DrawingAnchor(
        kind=_enum(AnchorKind, d["kind"]),
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
    )


def _drawing_object_from(d: Dict[str, Any]) -> DrawingProperties:
    return DrawingProperties(
        object_id=d["object_id"],
        kind=_enum(DrawingObjectKind, d["kind"]),
        name=LIB.read_cstr(d["name"]),
        description=LIB.read_cstr(d["description"]) or None,
        hidden=bool(d["hidden"]),
        anchor=_drawing_anchor_from(d["anchor"]),
    )


def _shape_from(d: Dict[str, Any]) -> ShapeProperties:
    return ShapeProperties(
        object_id=d["object_id"],
        kind=_enum(DrawingObjectKind, d["kind"]),
        preset_geometry=LIB.read_optional_cstr(d["preset_geometry"]),
        text=LIB.read_cstr(d["text"]),
        text_box=bool(d["text_box"]),
        fill=_shape_paint_read(_read_chart_fill(d["fill"])),
        line=_shape_paint_read(_read_chart_line(d["line"])),
        rotation_deg=d["rotation_deg"],
    )


def _shape_paint_read(record: Union[ChartFill, ChartLine]) -> Optional[Union[ChartFill, ChartLine]]:
    """``record``, or ``None`` when it states nothing (the object then draws its style's paint)."""
    stated = (
        record.none
        or record.other
        or record.color is not None
        or record.transparency_pct is not None
        or getattr(record, "width_pt", None) is not None
        or getattr(record, "dash", None) is not None
    )
    return record if stated else None


def _slicer_from(d: Dict[str, Any]) -> SlicerProperties:
    return SlicerProperties(
        object_id=d["object_id"],
        caption=LIB.read_cstr(d["caption"]),
        cache_name=LIB.read_cstr(d["cache_name"]),
        table_name=LIB.read_cstr(d["table_name"]),
        column_name=LIB.read_cstr(d["column_name"]),
        style_name=_opt_cstr(d["style_name"]),
        column_count=d["column_count"],
        row_height_pt=d["row_height_pt"],
        sort_order=_enum(SlicerSortOrder, d["sort_order"]),
        cross_filter=_enum(SlicerCrossFilter, d["cross_filter"]),
    )


def _person_from(d: Dict[str, Any]) -> Person:
    return Person(
        id=LIB.read_cstr(d["id"]),
        display_name=LIB.read_cstr(d["display_name"]),
        user_id=_opt_cstr(d["user_id"]),
        provider_id=_opt_cstr(d["provider_id"]),
    )


def _cell_style_from(d: Dict[str, Any]) -> CellStyle:
    return CellStyle(
        name=LIB.read_cstr(d["name"]),
        style_xf_index=d["style_xf_index"],
        builtin_id=d["builtin_id"] if d["has_builtin_id"] else None,
        outline_level=d["outline_level"],
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
        pivot=S.tristate_flag(d["pivot"]),
        table=S.tristate_flag(d["table"]),
        elements=list(elements),
    )


def _pack_cell_rect_at(ptr: int, r: CellRect) -> None:
    r = _require_cell_rect(r)
    S.CELL_RECT.pack(
        LIB,
        ptr,
        {
            "first_row": _uint(r.first_row, "first_row", role="coordinate"),
            "first_col": _uint(r.first_col, "first_col", role="coordinate"),
            "last_row": _uint(r.last_row, "last_row", role="coordinate"),
            "last_col": _uint(r.last_col, "last_col", role="coordinate"),
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
    ("value_iso", "val_iso"),
    ("max_value_iso", "max_val_iso"),
)


def _pack_filter_column(ptr: int, col: AutoFilterColumn, scope: _ArgScope) -> None:
    fields: Dict[str, object] = {
        "column_offset": _uint(col.column_offset, "column_offset"),
        "has_dxf_index": 0 if col.dxf_index is None else 1,
        "dxf_index": _uint(0 if col.dxf_index is None else col.dxf_index, "dxf_index", role="position"),
    }
    for name in _FILTER_COLUMN_BOOLS:
        fields[name] = int(_require_bool(getattr(col, name), name))
    for name, c_name in _FILTER_COLUMN_INTS:
        fields[c_name] = _sint(getattr(col, name), name)
    for name, c_name in _FILTER_COLUMN_DOUBLES:
        fields[c_name] = _require_number(getattr(col, name), name, finite=False)
    for name, c_name, flag in _FILTER_COLUMN_OPTIONAL:
        value = getattr(col, name)
        fields[flag] = 0 if value is None else 1
        if value is None:
            continue
        if name == "icon_index":
            fields[c_name] = _sint(value, name)
        else:
            fields[c_name] = _require_number(value, name, finite=False)
    fields["value_count"] = _uint(len(col.values), "value_count")
    fields["date_group_count"] = _uint(len(col.date_groups), "date_group_count")
    fields["condition_count"] = _uint(len(col.conditions), "condition_count")
    S.FILTER_COLUMN.pack(LIB, ptr, fields)
    for name, c_name in _FILTER_COLUMN_STRINGS:
        S.write_str_field(LIB, ptr, S.FILTER_COLUMN, c_name, getattr(col, name), scope)
    if col.values:
        arr = scope.alloc(4 * len(col.values))
        for i, text in enumerate(col.values):
            text_ptr = scope.cstr(text)
            LIB.write_bytes(arr + 4 * i, struct.pack("<I", text_ptr))
        LIB.write_bytes(ptr + S.FILTER_COLUMN.offsets["values"][1], struct.pack("<I", arr))
    if col.conditions:
        base = scope.struct_array(S.FILTER_CONDITION, len(col.conditions))
        for i, condition in enumerate(col.conditions):
            S.FILTER_CONDITION.pack(
                LIB,
                base + i * S.FILTER_CONDITION.size,
                {
                    "op": _sint(_enum(ComparisonOperator, condition.operator), "operator"),
                    "value": scope.cstr(condition.value),
                },
            )
        LIB.write_bytes(ptr + S.FILTER_COLUMN.offsets["conditions"][1], struct.pack("<I", base))
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
            "descending": _require_bool(cond.descending, "descending"),
            "sort_by": _sint(cond.sort_by, "sort_by"),
            "dxf_index": _uint(cond.dxf_index or 0, "dxf_index", role="position"),
            "has_dxf_index": 0 if cond.dxf_index is None else 1,
            "icon_set": _sint(cond.icon_set, "icon_set"),
            "icon_index": _sint(cond.icon_index or 0, "icon_index"),
            "has_icon_index": 0 if cond.icon_index is None else 1,
        },
    )
    _pack_cell_rect_at(ptr + S.SORT_CONDITION.offsets["ref"][1], cond.range)
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
        "column_sort": _require_bool(sort.column_sort, "column_sort"),
        "case_sensitive": _require_bool(sort.case_sensitive, "case_sensitive"),
        "sort_method": _sint(sort.sort_method, "sort_method"),
        "condition_count": _uint(len(sort.conditions), "condition_count"),
    }
    for name, value in fields.items():
        LIB.write_bytes(ptr + off[name][1], struct.pack("<i" if name != "condition_count" else "<I", value))
    _pack_cell_rect_at(ptr + off["sort_ref"][1], sort.range)
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
    conditions = [
        AutoFilterCondition(operator=_enum(ComparisonOperator, c["op"]), value=LIB.read_cstr(c["value"]))
        for c in (
            S.FILTER_CONDITION.unpack(LIB, d["conditions"] + i * S.FILTER_CONDITION.size)
            for i in range(d["condition_count"])
        )
    ]
    kwargs.update({name: d[c_name] if d[flag] else None for name, c_name, flag in _FILTER_COLUMN_OPTIONAL})
    kwargs.update({name: LIB.read_optional_cstr(d[c_name]) for name, c_name in _FILTER_COLUMN_STRINGS})
    return AutoFilterColumn(
        column_offset=d["column_offset"],
        dxf_index=d["dxf_index"] if d["has_dxf_index"] else None,
        values=values,
        date_groups=groups,
        conditions=conditions,
        **kwargs,
    )  # type: ignore[arg-type]


def _decode_sort_condition(ptr: int) -> AutoFilterSortCondition:
    d = S.SORT_CONDITION.unpack(LIB, ptr)
    return AutoFilterSortCondition(
        range=_decode_cell_rect_at(ptr + S.SORT_CONDITION.offsets["ref"][1]),
        descending=bool(d["descending"]),
        sort_by=d["sort_by"],
        custom_list=LIB.read_optional_cstr(d["custom_list"]),
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
            range=_decode_cell_rect_at(ptr + off["sort_ref"][1]),
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
    """Pack a list of ``PhoneticRun`` into a contiguous engine-memory array, spans in code points.

    Returns the array pointer (0 when empty). Every buffer allocated here --
    the array and each run's kana string -- goes through ``scope``.
    """
    if not runs:
        return 0
    size = S.PHONETIC_RUN.size
    ptr = scope.alloc(size * len(runs))
    for i, run in enumerate(runs):
        slot = ptr + i * size
        S.PHONETIC_RUN.pack(LIB, slot, {"start_cp": _uint(run.start, "start"), "end_cp": _uint(run.end, "end")})
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
                f"cannot infer the workbook format from {os.fspath(dest)!r}; "
                "use a .xlsx or .xlsb extension or pass format=WorkbookFormat.XLSX or WorkbookFormat.XLSB"
            )
        return WorkbookFormat.XLSX
    fmt = WorkbookFormat(format)
    if implied is not None and fmt != implied:
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT,
            f"format {fmt.name} contradicts the extension of {os.fspath(dest)!r}",
            "",
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
        "a pivot cache value must be a number, a str, a bool, None, an ExcelError, "
        f"a date or a datetime, got {type(value).__name__}"
    )


class Workbook(_NoCopy):
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

    Names:
      Every named collection (sheets, tables, table columns, defined names, pivots, protected ranges, cell
      styles, custom table styles) and the person ids share one identity fold. Writing a name (add, rename,
      copy) is refused when another member of the collection already holds it under the wide fold: the
      workbook collation, which ignores case, full/half width and hiragana/katakana. Reading a name
      resolves an exact match first, then a match under Unicode simple case folding, then the one member
      matching under the wide fold, if there is exactly one. A default sheet name (``add_sheet()`` without a
      name) is free under the same wide fold.

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

        Names in the loaded file resolve as the workbook's identity fold does: an exact match, then a
        case-only match, then a unique wide fold match. Two names of a collection that collide only under the
        wide fold (``ﾃﾞｰﾀ`` and ``データ``) load and stay; the file is not rejected or renamed. A threaded
        comment's ``dT`` is normalised on load: a trailing ``Z`` is dropped, and a numeric offset or a
        malformed stamp reads as no ``created`` (the attribute is omitted on save).

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
          InvalidFileError: when the input cannot be parsed as a valid
            OOXML or XLSB archive.
          SecurityError: when the file exceeds an engine ceiling; the error
            context's ``part=`` names it.

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

    @property
    def ooxml(self) -> WorkbookOoxml:
        """The OOXML-level layer of the workbook. Always the same object.

        It holds the style record tables by index (``list_fonts``, ``list_fills``, ``list_borders``,
        ``add_number_format``, ``get_number_format``, ``list_cell_style_xfs``, ``add_cell_style_xf``, ...) and
        the passthrough parts (``list_passthroughs``, ``has_passthrough``). Named styles
        (:meth:`list_cell_styles`, :meth:`get_cell_style`, :meth:`set_cell_style`, :meth:`remove_cell_style`)
        live on the workbook itself.
        """
        layer = getattr(self, "_ooxml", None)
        if layer is None:
            layer = WorkbookOoxml._new(self)
            self._ooxml = layer
        return layer

    def get_memory_usage(self) -> int:
        """Return an estimate of the workbook's heap footprint, in bytes.

        Intended for a host that pools workbooks and needs a size to
        evict on. The figure covers the cell store, the shared-string
        storage, the passthrough part payloads and the workbook-level
        metadata. It **excludes** allocator overhead, the dependency
        graph and the style tables, so it is an estimate for relative
        pressure rather than the process's retained heap; a
        formula-heavy workbook undercounts by the size of its graph.

        The walk visits every stored cell, so call it at coarse
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
                "the workbook is closed",
                "",
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
                    "the workbook is closed",
                    "",
                )
            return self._handle

    # -- Sheets ------------------------------------------------------------
    def _sheet_count(self) -> int:
        h = self._require()
        return _read_count(LIB.fm_workbook_get_sheet_count, h)

    def _sheet_name(self, index: int) -> str:
        h = self._require()
        return _read_cstr_out(LIB.fm_workbook_get_sheet_name, h, _uint(index, "sheet_index", role="sheet"))

    def _find_sheet(self, name: str) -> Optional[int]:
        """Index of the sheet ``name`` resolves to (exact, case-only, then unique wide fold), or ``None``."""
        if not isinstance(name, str):
            raise TypeError(f"name must be a str, got {type(name).__name__}")
        h = self._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            present = scope.out_ptr()
            _check(LIB.fm_workbook_get_sheet_index(h, scope.cstr(name), present, out))
            if not LIB.read_i32(present):
                return None
            return LIB.read_u32(out)

    def _sheet_handles(self) -> List[Sheet]:
        """The live handle list, built from the engine on first use and kept in sheet order after that."""
        from ._sheet import Sheet

        self._require()
        if self._sheets is None:
            self._sheets = [Sheet._new(self, index, self._sheet_name(index)) for index in range(self._sheet_count())]
        return self._sheets

    def _insert_sheet_handle(self, handles: List[Sheet], index: int) -> Sheet:
        """Track a sheet the engine inserted at ``index``; ``handles`` is the list read before the change."""
        from ._sheet import Sheet

        sheet = Sheet._new(self, index, self._sheet_name(index))
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
          key: a sheet name (resolved by the identity fold: exact, case-only,
            then unique wide fold), a 0-based sheet index, or a :class:`Sheet`
            of this workbook.

        Raises:
          NotFoundError: when no sheet has the name ``key``, or ``key`` is a
            removed sheet.
          IndexOutOfRangeError: when ``key`` is an index ``>=`` the sheet count.
          IndexOutOfRangeError: when ``key`` is a negative index (``StatusCode.SHEET_INDEX_OUT_OF_RANGE``).
          ValueError: when ``key`` does not fit 64 bits.
          InvalidArgumentError: when ``key`` is a :class:`Sheet` of another
            workbook.
          TypeError: when ``key`` is none of the above.

        Example::

            ws = wb.sheet("Sheet1")
            ws.set_text("A1", "Total")
        """
        from ._sheet import Sheet

        self._require()
        if isinstance(key, Sheet):
            if key._workbook is not self:
                raise FormulonError(StatusCode.INVALID_ARGUMENT, "the sheet belongs to another workbook", "")
            key._index()
            return key
        if isinstance(key, str):
            index = self._find_sheet(key)
            if index is None:
                raise FormulonError(
                    StatusCode.NOT_FOUND,
                    _argument_message("key", f"{json.dumps(key, ensure_ascii=False)} names no sheet"),
                    "",
                )
            return self._sheet_handles()[index]
        if isinstance(key, bool):
            raise TypeError("sheet key must be a str, an int or a Sheet, got bool")
        try:
            index = operator.index(key)
        except TypeError:
            raise TypeError(f"sheet key must be a str, an int or a Sheet, got {type(key).__name__}") from None
        index = _uint(index, "sheet_index", bits=64, role="sheet")
        handles = self._sheet_handles()
        count = len(handles)
        if index >= count:
            raise FormulonError(
                StatusCode.SHEET_INDEX_OUT_OF_RANGE,
                _argument_message("key", f"{index} {_past_count(count, 'sheet')}"),
                f"sheet_index={index} sheet_count={count}",
            )
        return handles[index]

    def list_sheets(self) -> List[Sheet]:
        """Every sheet's handle, in sheet order, as a new list."""
        return list(self._sheet_handles())

    def __getitem__(self, key: SheetRef) -> Sheet:
        """``wb[key]`` is :meth:`sheet`, except that a negative index counts from the end (``wb[-1]`` is the
        last sheet) as for a sequence."""
        if isinstance(key, numbers.Integral) and not isinstance(key, bool) and key < 0:
            key = int(key) + len(self)
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

        The first sheet without workbook views; tab selection is not
        consulted. A load follows Excel's open rule: a hidden active sheet is
        unhidden, and a very hidden one passes activation to the first visible
        sheet. Removing or hiding the active sheet activates the next visible
        sheet, else the previous one.
        """
        h = self._require()
        handles = self._sheet_handles()
        return handles[_read_count(LIB.fm_workbook_get_active_sheet, h)]

    def set_active_sheet(self, sheet: SheetRef) -> None:
        """Mark ``sheet`` active so :meth:`save` records it.

        Creates a minimal ``<bookViews>`` when the workbook has none; other
        ``<workbookView>`` attributes survive, and ``firstSheet`` is lowered
        to the sheet's index when larger. The sheet becomes the only
        tab-selected one.

        Raises:
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when ``sheet`` is hidden.
        """
        index = self.sheet(sheet).index
        h = self._require()
        _check(LIB.fm_workbook_set_active_sheet(h, _uint(index, "sheet_index", role="sheet")))

    def _insert_before_position(self, insert_before: Optional[SheetRef], count: int) -> int:
        """The position an ``insert_before`` argument names: ``None`` or ``count`` append, any other
        :data:`SheetRef` is the sheet the new position goes before."""
        if insert_before is None:
            return count
        if (
            isinstance(insert_before, numbers.Integral)
            and not isinstance(insert_before, bool)
            and int(insert_before) == count
        ):
            return count
        return self.sheet(insert_before).index

    def add_sheet(self, name: Optional[str] = None, *, insert_before: Optional[SheetRef] = None) -> Sheet:
        """Add a new sheet and return its handle.

        Without ``name`` the sheet is named ``<base><N>``: the profile's first sheet name without its trailing
        ``1`` and the smallest number ``N`` (from 1) such that no sheet holds ``<base>N`` under the wide fold of the
        identity fold; a new workbook's first sheet takes 1, so the next default is usually ``Sheet2`` or
        ``Tabelle2``. An empty name is rejected.

        The sheet goes before ``insert_before`` (a sheet, a name or a position); omitted, ``None`` or the sheet
        count appends it.

        Raises:
          AlreadyExistsError: when the name collides with an existing sheet under
            the wide fold of the identity fold (case, width and kana type).
          FormulonError: ``StatusCode.INVALID_SHEET_NAME`` when the name is
            empty, longer than 31 UTF-16 units, contains one of ``:``,
            backslash, ``/``, ``?``, ``*``, ``[`` or ``]``, begins or ends
            with an apostrophe, or is ``History`` in any case.
          SecurityError: at the maximum sheet count.
        """
        h = self._require()
        handles = self._sheet_handles()
        before = self._insert_before_position(insert_before, len(handles))
        with _ArgScope() as scope:
            name_ptr = 0 if name is None else scope.cstr(name)
            out = scope.out_ptr()
            _check(LIB.fm_workbook_add_sheet(h, name_ptr, _uint(before, "insert_before", role="sheet"), out))
            index = LIB.read_u32(out)
        return self._insert_sheet_handle(handles, index)

    def _evaluate_typed_formula(self, sheet: int, row: int, col: int, formula: str) -> Value:
        """Evaluate as :meth:`Sheet.evaluate_typed_formula`: a 1x1 result is its cell, any other an ``ARRAY``."""
        if not isinstance(formula, str):
            raise TypeError(f"formula must be a str, got {type(formula).__name__}")
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
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    formula_ptr,
                    out,
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
                        _check(
                            LIB.fm_value_array_get_value_at(
                                array, _uint(r * stride + c, "index", role="position"), value_ptr
                            )
                        )
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

        Raises:
          TypeError: when ``enabled`` is not a ``bool`` or ``max_iterations``
            is not an integer.
          InvalidArgumentError: when ``max_iterations`` is negative, 0 or above
            32767 (Excel's limit), or ``max_change`` is negative, NaN or infinite.
        """
        fields: Dict[str, object] = {}
        if enabled is not UNCHANGED:
            fields["enabled_op"] = _PATCH_SET
            fields["enabled"] = int(_require_bool(enabled, "enabled"))
        if max_iterations is not UNCHANGED:
            fields["max_iterations_op"] = _PATCH_SET
            fields["max_iterations"] = _uint(max_iterations, "max_iterations")
        if max_change is not UNCHANGED:
            fields["max_change_op"] = _PATCH_SET
            fields["max_change"] = _require_number(max_change, "max_change", finite=False)
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
        d = _read_struct(LIB.fm_workbook_get_iterative, S.ITERATIVE_SETTINGS, h)
        return IterativeSettings(
            enabled=bool(d["enabled"]),
            max_iterations=d["max_iterations"],
            max_change=d["max_change"],
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
          InvalidArgumentError: for a ``format`` that contradicts a ``.xlsx`` / ``.xlsb`` extension of ``dest``.
          IoError: when the write fails, including while a pivot uses a cache
            without a worksheet source (see :meth:`add_pivot_cache`); a cache
            no pivot or slicer uses is not written. An ``.xlsb`` save of a
            workbook holding a feature this writer cannot write to xlsb raises
            it with code ``StatusCode.IO_FEATURE_NOT_WRITABLE`` and writes
            nothing; the error's ``context`` names the feature as ``kind=`` and,
            for a sheet-scoped one, the sheet as ``sheet_index=``.
          TypeError: when ``dest`` is none of the above.
          OSError: when writing ``dest`` fails.

        Example::

            wb.save("out.xlsx")
            data = wb.save()  # bytes
        """
        fmt = _save_format(dest, format)
        h = self._require()
        data = _read_owned_buffer(LIB.fm_workbook_save, h, _sint(fmt, "format"))
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
          InvalidArgumentError: for a ``format`` that contradicts a ``.xlsx`` / ``.xlsb`` extension of ``dest``.
          IoError: in the cases :meth:`save` raises it.
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
                _check(LIB.fm_workbook_get_defined_name_at(h, _uint(i, "idx", role="position"), ptr))
                d = S.DEFINED_NAME.unpack(LIB, ptr)
                out.append(_defined_name_from(d))
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
            raise TypeError(f"table name must be a str, got {type(name).__name__}")
        h = self._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            sheet_ptr = scope.out_ptr()
            index_ptr = scope.out_ptr()
            present = scope.out_ptr()
            _check(LIB.fm_workbook_find_table(h, name_ptr, present, sheet_ptr, index_ptr))
            if LIB.read_i32(present) == 0:
                raise FormulonError(StatusCode.NOT_FOUND, f'no table is named "{name}"', f"name={name}")
            sheet, index = LIB.read_u32(sheet_ptr), LIB.read_u32(index_ptr)
        return self._table_handle(sheet, index)

    def _table_count(self, sheet_index: int) -> int:
        """Return the number of tables on sheet ``sheet_index``."""
        h = self._require()
        return _read_count(LIB.fm_sheet_get_table_count, h, _uint(sheet_index, "sheet_index", role="sheet"))

    def _table_handle(self, sheet_index: int, table_index: int) -> Table:
        """The handle of table ``table_index`` on ``sheet_index``, created on first use and kept for the workbook's life."""
        from ._table import Table

        h = self._require()
        runtime_id = _read_count(
            LIB.fm_sheet_get_table_runtime_id,
            h,
            _uint(sheet_index, "sheet_index", role="sheet"),
            _uint(table_index, "table_index", role="position"),
        )
        handle = self._tables.get(runtime_id)
        if handle is None:
            handle = Table._new(self, runtime_id)
            self._tables[runtime_id] = handle
        return handle

    def _add_table(self, sheet_index: int, table: TableInput) -> Table:
        """Create the table ``table`` describes on sheet ``sheet_index`` and return its handle.

        The body of :meth:`Sheet.add_table`; see there for the contract.
        """
        if not isinstance(table, TableInput):
            raise TypeError(f"expected a TableInput, got {type(table).__name__}")
        rect = cell_rect(self, table.range, "TableInput.range")
        if isinstance(table.name, str) and self._names_a_cell(table.name):
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                f'invalid table name: "{table.name}" is a cell address',
                f"name={table.name}",
            )
        h = self._require()
        sheet = _uint(sheet_index, "sheet_index", role="sheet")
        with _ArgScope() as scope:
            names = [str(column) for column in table.columns or ()]
            name_ptrs = [scope.nonempty_cstr(column) for column in names]
            array_ptr = scope.alloc(4 * len(name_ptrs)) if name_ptrs else 0
            if array_ptr:
                LIB.write_bytes(array_ptr, b"".join(struct.pack("<I", p) for p in name_ptrs))
            spec = scope.struct(S.TABLE_SPEC)
            S.TABLE_SPEC.pack(
                LIB,
                spec,
                {
                    "range": _cell_rect_fields(rect),
                    "name": scope.nonempty_cstr(table.name),
                    "column_names": array_ptr,
                    "column_count": _uint(len(name_ptrs), "column_count"),
                    "style_name": scope.opt_cstr(table.style_name),
                    "styled": _tristate_default(bool(table.styled)),
                    "header_row": _tristate_default(bool(table.header_row)),
                    "totals_row": _tristate_default(bool(table.totals_row)),
                    "show_first_column": _tristate_default(table.show_first_column),
                    "show_last_column": _tristate_default(table.show_last_column),
                    "show_row_stripes": _tristate_default(table.show_row_stripes),
                    "show_column_stripes": _tristate_default(table.show_column_stripes),
                },
            )
            out = scope.out_ptr()
            _check(LIB.fm_sheet_add_table(h, sheet, spec, out))
            index = LIB.read_u32(out)
        return self._table_handle(sheet_index, index)

    def _names_a_cell(self, text: str) -> bool:
        """Whether ``text`` reads as one A1 cell (``"T2"``), which a table name may not be."""
        try:
            ref = self.parse_reference(text)
        except FormulaSyntaxError:
            return False
        rect = ref.range
        return ref.sheet_name is None and rect.first_row == rect.last_row and rect.first_col == rect.last_col

    # -- Sheet structure ---------------------------------------------------
    def copy_sheet(
        self, source: SheetRef, *, insert_before: Optional[SheetRef] = None, new_name: Optional[str] = None
    ) -> Sheet:
        """Copy the worksheet ``source`` and return the copy's handle.

        The copy goes before ``insert_before`` (a sheet, a name or a position); omitted, ``None`` or the sheet
        count appends it. Without ``new_name`` it is named as Excel does
        (``Src (2)``); an explicit name follows the :meth:`add_sheet` rules.
        The active sheet does not change.

        Raises:
          AlreadyExistsError: when ``new_name`` collides with an existing sheet.
          SecurityError: when the workbook is already at the maximum sheet count.
          FormulonError: ``StatusCode.SHEET_INDEX_OUT_OF_RANGE``, ``StatusCode.INVALID_SHEET_NAME``,
            ``StatusCode.PRECONDITION_FAILED`` (very hidden source) or
            ``StatusCode.NOT_IMPLEMENTED`` (content the copy cannot model).
        """
        source_index = self.sheet(source).index
        h = self._require()
        handles = self._sheet_handles()
        before = self._insert_before_position(insert_before, len(handles))
        with _ArgScope() as scope:
            name_ptr = 0 if new_name is None else scope.cstr(new_name)
            index = _read_count(
                LIB.fm_workbook_copy_sheet,
                h,
                _uint(source_index, "source"),
                _uint(before, "insert_before", role="sheet"),
                name_ptr,
            )
        return self._insert_sheet_handle(handles, index)

    def remove_sheet(self, sheet: SheetRef) -> None:
        """Remove ``sheet``; its handle raises :class:`NotFoundError` from then on.

        Raises:
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when it is the
            workbook's only sheet or its only visible one; a workbook keeps at
            least one visible sheet.
        """
        removed = self.sheet(sheet)
        index = removed.index
        h = self._require()
        handles = self._sheet_handles()
        _check(LIB.fm_workbook_remove_sheet(h, _uint(index, "sheet_index", role="sheet")))
        handles.pop(index)
        removed._position = None
        self._renumber_sheets(handles)

    # -- Defined names -----------------------------------------------------
    def set_defined_name(self, defined_name: Union[DefinedNameInput, DefinedName]) -> None:
        """Set or append a defined name in workbook or sheet scope.

        A :class:`DefinedName` read back from :meth:`get_defined_name` is accepted.

        An existing name in the same scope is replaced as a whole: an
        omitted ``comment`` or ``hidden`` goes back to its default.

        Raises:
          InvalidArgumentError: when ``formula`` is empty (remove a name with
            :meth:`remove_defined_name`), or the name is malformed. A new name that is,
            or starts with, a reference in the profile locale's R1C1 letters (de ``Z1S1``) is
            refused; redefining an existing name stays allowed. The check uses the profile
            at the time of the call, so a loaded name is never re-checked.
          AlreadyExistsError: when the name equals a table's name.
          IndexOutOfRangeError: when ``scope`` is a sheet index past the sheet count.
        """
        local_sheet_id = _local_sheet_id(None if defined_name.scope is None else self.sheet(defined_name.scope).index)
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.DEFINED_NAME)
            S.DEFINED_NAME.pack(
                LIB,
                ptr,
                {
                    "name": scope.cstr(defined_name.name),
                    "formula": scope.cstr(defined_name.formula),
                    "local_sheet_id": _sint(local_sheet_id, "local_sheet_id"),
                    "comment": scope.nonempty_cstr(defined_name.comment),
                    "hidden": _require_bool(defined_name.hidden, "hidden"),
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
        """Return the defined name ``name`` (compared case-insensitively under Unicode simple case folding) in ``scope``, or ``None``.

        ``scope`` is ``None`` for workbook scope, else the sheet the name is local to.

        Raises:
          IndexOutOfRangeError: when ``scope`` is a sheet index past the sheet count.
        """
        sheet = None if scope is None else self.sheet(scope).index
        h = self._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            present = scope.out_ptr()
            out = scope.struct(S.DEFINED_NAME)
            _check(
                LIB.fm_workbook_get_defined_name(
                    h, name_ptr, _sint(_local_sheet_id(sheet), "local_sheet_id"), present, out
                )
            )
            if not LIB.read_u32(present):
                return None
            return _defined_name_from(S.DEFINED_NAME.unpack(LIB, out))

    def remove_defined_name(self, name: str, scope: Optional[SheetRef] = None) -> None:
        """Remove the defined name ``name`` in ``scope`` (``None`` for workbook scope, else a sheet).

        Raises:
          NotFoundError: when the scope has no such name.
          IndexOutOfRangeError: when ``scope`` is a sheet index past the sheet count.
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

    def get_compatibility_version(self) -> CompatibilityVersion:
        """Return the Compatibility Version the workbook evaluates under.

        It is the workbook's stored value. A new workbook reads ``VERSION_2``;
        a loaded workbook whose file carries none reads ``VERSION_1``, as Excel
        computes it. A stored ``VERSION_3`` evaluates as ``VERSION_2``. Values
        above 3, stamped by a later Excel, come back as a
        ``CompatibilityVersion`` pseudo-member with that number; they evaluate
        as ``VERSION_2`` and are written back unchanged.
        """
        h = self._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_workbook_get_compatibility_version(h, out))
            return CompatibilityVersion(LIB.read_i32(out))

    def set_compatibility_version(self, version: CompatibilityVersion) -> None:
        """Set the Compatibility Version (see :class:`CompatibilityVersion`).

        Setting the current value changes nothing. ``VERSION_3`` is stored and
        written as given and evaluates as ``VERSION_2``. A change marks every
        formula dirty only when it moves between ``VERSION_1`` and ``VERSION_2``
        or ``VERSION_3``; a change between ``VERSION_2`` and ``VERSION_3`` marks
        nothing. A value other than ``VERSION_1``, ``VERSION_2`` or
        ``VERSION_3``, including one read back from a later Excel's stamp,
        raises :class:`InvalidArgumentError` and leaves the workbook unchanged.
        """
        h = self._require()
        _check(LIB.fm_workbook_set_compatibility_version(h, _sint(version, "version")))

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
        _check(LIB.fm_workbook_set_full_precision(h, int(_require_bool(full_precision, "full_precision"))))

    def get_calc_on_save(self) -> bool:
        """Return ``<calcPr calcOnSave>`` (default true)."""
        return self._read_flag(LIB.fm_workbook_get_calc_on_save)

    def set_calc_on_save(self, calc_on_save: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set ``<calcPr calcOnSave>``."""
        h = self._require()
        _check(LIB.fm_workbook_set_calc_on_save(h, int(_require_bool(calc_on_save, "calc_on_save"))))

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
        _check(LIB.fm_workbook_set_date1904(h, int(_require_bool(date1904, "date1904"))))

    def get_pinned_now(self) -> Optional[datetime.datetime]:
        """Return the pinned wall-clock reading as a naive ``datetime``, or
        ``None`` when the workbook follows the host clock (the default)."""
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CIVIL_TIME)
            flag = scope.out_ptr()
            _check(LIB.fm_workbook_get_pinned_now(h, flag, ptr))
            if LIB.read_i32(flag) == 0:
                return None
            d = S.CIVIL_TIME.unpack(LIB, ptr)
            return datetime.datetime(d["year"], d["month"], d["day"], d["hour"], d["minute"], d["second"])

    def set_pinned_now(self, time: Optional[Union[datetime.datetime, datetime.date]]) -> None:
        """Pin every clock-dependent result to the local time ``time``; ``None`` releases the pin.

        A threaded comment stamped while the clock is pinned carries this time as written, without a
        UTC conversion, because the pin has no time zone.

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
          TypeError: when ``time`` is not a ``datetime.date`` or ``datetime.datetime``.
          InvalidArgumentError: unless the year is in ``[1900, 9999]``.
        """
        h = self._require()
        if time is None:
            _check(LIB.fm_workbook_set_pinned_now(h, 0))
            return
        pinned = civil_datetime(time, "the pinned time")
        with _ArgScope() as scope:
            ptr = scope.struct(S.CIVIL_TIME)
            S.CIVIL_TIME.pack(
                LIB,
                ptr,
                {
                    "year": _sint(pinned.year, "year"),
                    "month": _sint(pinned.month, "month"),
                    "day": _sint(pinned.day, "day"),
                    "hour": _sint(pinned.hour, "hour"),
                    "minute": _sint(pinned.minute, "minute"),
                    "second": _sint(pinned.second, "second"),
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

    def get_location(self) -> Optional[WorkbookLocation]:
        """Return where the workbook is saved; ``None`` when it has no location.

        ``None`` when the workbook has no location, as for a book that has not
        been saved. The location is metadata only: it is never read from or written to a
        file system. It drives ``CELL("filename")``, ``CELL("address")`` for
        other sheets, the ``[0]!Name`` spelling, the saved absolute folder
        path, and the relative form of local hyperlink targets in a saved file.
        ``INFO("directory")`` is the application's current directory and does
        not follow it. Loading (:func:`load`) leaves the location unset, and
        :meth:`save` does not set it.

        While a file name is set, every formula getter spells a reference to the
        workbook's own defined name (stored ``[0]!Name``) as ``Book.xlsx!Name``,
        quoted when needed.
        """
        h = self._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.WORKBOOK_LOCATION)
            _check(LIB.fm_workbook_get_location(h, present, ptr))
            if not LIB.read_i32(present):
                return None
            d = S.WORKBOOK_LOCATION.unpack(LIB, ptr)
            directory = LIB.read_cstr(d["directory"]) if d["directory"] else None
            return WorkbookLocation(directory=directory, file_name=LIB.read_cstr(d["file_name"]))

    def set_location(self, location: Union[WorkbookLocationInput, WorkbookLocation, None]) -> None:
        """Set :meth:`get_location`; ``None`` unsets it.

        Set the location to the destination before saving to a new place, as
        Excel's Save As does; relative targets in the saved file are computed
        against it. The directory is stored in its canonical form (see
        :class:`WorkbookLocation`), so ``set_location(get_location())`` changes
        nothing. Entered formulas then read ``Book.xlsx!Name`` and
        ``[Book.xlsx]Sheet1!A1`` as this workbook and create no external link.
        When a folder first appears, every relative local hyperlink target is
        resolved against it. Values that read the location (``CELL``,
        ``FORMULATEXT``) refresh at the next recalculation, as Excel's do
        after Save As; a save writes them as computed.

        Raises:
          TypeError: when ``location`` is neither a :class:`WorkbookLocationInput`,
            a :class:`WorkbookLocation` nor ``None``, ``file_name`` is not a string,
            or ``directory`` is neither a string nor ``None``.
          FormulonError: with ``StatusCode.INVALID_ARGUMENT``, changing nothing,
            when ``file_name`` is empty, lacks a workbook extension, holds one of
            ``[ ] : / \\ ? *``, is longer than 218 bytes, or equals an external
            link's book name; when ``directory`` is ``""`` (use ``None``), is not
            absolute in one of the three forms, holds ``[`` or ``]``, or has a
            ``.``, ``..`` or empty inner segment.
          SecurityError: when the directory and file name together are longer than
            1,024 bytes.
        """
        h = self._require()
        if location is None:
            _check(LIB.fm_workbook_set_location(h, 0))
            return
        if not isinstance(location, (WorkbookLocationInput, WorkbookLocation)):
            raise TypeError(
                _argument_message("location", "must be a WorkbookLocationInput, a WorkbookLocation or None")
            )
        if not isinstance(location.file_name, str):
            raise TypeError(_argument_message("location.file_name", "must be a string"))
        if location.directory is not None and not isinstance(location.directory, str):
            raise TypeError(_argument_message("location.directory", "must be a string or None"))
        with _ArgScope() as scope:
            ptr = scope.struct(S.WORKBOOK_LOCATION)
            S.write_str_field(LIB, ptr, S.WORKBOOK_LOCATION, "directory", location.directory, scope, allow_empty=True)
            S.write_str_field(LIB, ptr, S.WORKBOOK_LOCATION, "file_name", location.file_name, scope, allow_empty=True)
            _check(LIB.fm_workbook_set_location(h, ptr))

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

    def get_workbook_protection(self) -> Optional[WorkbookProtection]:
        """Read the workbook's ``<workbookProtection>``; ``None`` when the workbook has no such element.

        An element with every lock off is present and reads as a record.
        """
        h = self._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.WORKBOOK_PROTECTION)
            _check(LIB.fm_workbook_get_workbook_protection(h, present, ptr))
            if not LIB.read_u32(present):
                return None
            d = S.WORKBOOK_PROTECTION.unpack(LIB, ptr)
            return WorkbookProtection(
                lock_structure=bool(d["lock_structure"]),
                lock_windows=bool(d["lock_windows"]),
                lock_revision=bool(d["lock_revision"]),
                workbook_spin_count=d["workbook_spin_count"],
                revisions_spin_count=d["revisions_spin_count"],
                **{fld: LIB.read_cstr(d[fld]) or None for fld in self._WORKBOOK_PROTECTION_TEXT},
            )

    def set_workbook_protection(self, protection: Union[WorkbookProtectionInput, WorkbookProtection, None]) -> None:
        """Replace the modelled ``<workbookProtection>`` attributes.

        ``None`` removes the element (a no-op without one); a record with every lock off keeps an element. A
        :class:`WorkbookProtection` read back is accepted.
        """
        if protection is None:
            _check(LIB.fm_workbook_set_workbook_protection(self._require(), 0))
            return
        if not isinstance(protection, (WorkbookProtectionInput, WorkbookProtection)):
            raise TypeError(f"expected a WorkbookProtectionInput, got {type(protection).__name__}")
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.WORKBOOK_PROTECTION)
            S.WORKBOOK_PROTECTION.pack(
                LIB,
                ptr,
                {
                    "lock_structure": _require_bool(protection.lock_structure, "lock_structure"),
                    "lock_windows": _require_bool(protection.lock_windows, "lock_windows"),
                    "lock_revision": _require_bool(protection.lock_revision, "lock_revision"),
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
    def _write_auto_filter(self, fn, auto_filter: Optional[AutoFilter], *indices: int) -> None:
        """Replace the filter; ``None`` removes it (a no-op without one)."""
        h = self._require()
        with _ArgScope() as scope:
            ptr = 0
            if auto_filter is not None:
                ptr = scope.struct_array(S.AUTO_FILTER, 1)
                _pack_auto_filter(ptr, auto_filter, scope)
            _check(fn(h, *(_uint(i, "index", role="position") for i in indices), ptr))

    def _evaluate_auto_filter(self, fn, *indices: int) -> "tuple[int, List[bool]]":
        h = self._require()
        args = tuple(_uint(i, "index", role="position") for i in indices)
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
                _check(LIB.fm_workbook_get_person_at(h, _uint(i, "idx", role="position"), ptr))
                d = S.PERSON.unpack(LIB, ptr)
                out.append(_person_from(d))
        return out

    def add_person(self, person: Union[PersonInput, Person]) -> str:
        """Append a person to the workbook's person list and return the id as stored (upper case).

        An omitted or ``None`` ``id`` takes an engine-generated GUID.

        Raises:
          AlreadyExistsError: when the id is already listed.
          InvalidArgumentError: for an id that is not a GUID, or an empty ``display_name``.
        """
        if not isinstance(person, (PersonInput, Person)):
            raise TypeError(f"expected a PersonInput, got {type(person).__name__}")
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.PERSON, 1)
            for fld in ("id", "display_name", "user_id", "provider_id"):
                S.write_str_field(LIB, ptr, S.PERSON, fld, getattr(person, fld), scope)
            out = scope.out_ptr()
            _check(LIB.fm_workbook_add_person(h, ptr, out))
            return LIB.read_cstr(LIB.read_u32(out))

    def get_person(self, id: str) -> Optional[Person]:  # noqa: A002 -- named after the declaration
        """Return the person ``id``, or ``None`` when the list has none."""
        return self._lookup(LIB.fm_workbook_get_person, (id,), S.PERSON, _person_from)

    def remove_person(self, id: str) -> None:  # noqa: A002 -- named after the declaration
        """Remove the person ``id``.

        Raises:
          NotFoundError: when no person has the id.
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` while a comment or mention still names them.
        """
        h = self._require()
        with _ArgScope() as scope:
            id_ptr = scope.cstr(id)
            _check(LIB.fm_workbook_remove_person(h, id_ptr))

    def format_value(self, value: "CellValueInput", number_format: str) -> DisplayText:
        """Render ``value`` under the number format ``number_format`` (empty means General).

        ``value`` is any cell input: a native number, ``str`` or ``bool``, a ``date`` / ``datetime`` (read as a
        serial under the workbook's date system), an :class:`ExcelError` or a :class:`Value`.

        Raises:
          TypeError: when ``value`` is not a cell value (a ``list`` included).
        """
        from ._sheet import _value_of_input

        h = self._require()
        # The date system is read only for a date, so a value of the wrong type is refused before any call.
        date1904 = self.get_date1904() if isinstance(value, (datetime.date, datetime.time)) else False
        value = _value_of_input(value, date1904=date1904)
        with _ArgScope() as scope:
            value_ptr = scope.alloc(fm_value_t_size)
            text_ptr = scope.out_ptr()
            status_ptr = scope.out_ptr()
            _write_value_input(value, value_ptr, scope)
            fmt_ptr = scope.cstr(number_format)
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
    # file has it". That is how the C ``has_*`` flags are expressed in
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
            _check(fn(h, _uint(index, index_field, role="position"), ptr))
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

        A ``None`` size or name keeps the current one.
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
                _check(LIB.fm_workbook_get_cell_style_at(h, _uint(i, "index", role="position"), ptr))
                out.append(_cell_style_from(S.CELL_STYLE_RECORD.unpack(LIB, ptr)))
        return out

    def get_cell_style(self, name: str) -> Optional[CellStyle]:
        """Return the named cell style ``name``, or ``None``.

        Names resolve by name identity; a built-in style also answers to its name in the workbook
        profile's locale, and among exact or case-only matches the first style wins.
        """
        return self._lookup(LIB.fm_workbook_get_cell_style, (name,), S.CELL_STYLE_RECORD, _cell_style_from)

    def set_cell_style(self, record: CellStyle) -> None:
        """Add or replace the named cell style ``record.name``.

        ``record.style_xf_index`` must reference an existing named-style xf --
        register one with :meth:`WorkbookOoxml.add_cell_style_xf` first.
        ``builtin_id`` is an Excel built-in style ordinal (0..53); ``None`` marks a
        custom style. The name resolves as in :meth:`get_cell_style`. A custom style it
        resolves to is replaced whole; a built-in one keeps its ``builtin_id`` and
        ``outline_level``, even when ``record.builtin_id`` is ``None``, and takes the
        record's name, ``style_xf_index``, ``hidden`` and ``custom_builtin``.

        Raises:
          AlreadyExistsError: when the name resolves to a built-in style and the record
            carries a different ``builtin_id``, or collides with another style.
        """
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CELL_STYLE_RECORD)
            S.CELL_STYLE_RECORD.pack(
                LIB,
                ptr,
                {
                    "style_xf_index": _uint(record.style_xf_index, "style_xf_index", role="position"),
                    "has_builtin_id": 0 if record.builtin_id is None else 1,
                    "builtin_id": 0 if record.builtin_id is None else _uint(record.builtin_id, "builtin_id"),
                    "outline_level": _uint(record.outline_level, "outline_level"),
                    "hidden": _require_bool(record.hidden, "hidden") if record.hidden is not None else 0,
                    "custom_builtin": _require_bool(record.custom_builtin, "custom_builtin"),
                },
            )
            S.write_str_field(LIB, ptr, S.CELL_STYLE_RECORD, "name", record.name, scope, allow_empty=True)
            _check(LIB.fm_workbook_set_cell_style(h, ptr))

    def remove_cell_style(self, name: str) -> None:
        """Remove the named cell style ``name``.

        Cell xfs inheriting from it are repointed at Normal; no xf index
        shifts. The name resolves as in :meth:`get_cell_style`. A name no style has, or
        that two styles match only under the wide fold, raises ``NotFoundError`` and the
        Normal style ``InvalidArgumentError``.
        """
        h = self._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            _check(LIB.fm_workbook_remove_cell_style(h, name_ptr))

    def _add_xf_record(self, record: Union[CellXfInput, CellXf], fn) -> int:
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
                _check(LIB.fm_workbook_get_table_style_at(h, _uint(i, "index", role="position"), ptr))
                out.append(_table_style_from(S.TABLE_STYLE.unpack(LIB, ptr)))
        return out

    def get_table_style(self, name: str) -> Optional[TableStyle]:
        """Return the custom table style ``name`` (Unicode simple case folding), or ``None``."""
        return self._lookup(LIB.fm_workbook_get_table_style, (name,), S.TABLE_STYLE, _table_style_from)

    def set_table_style(self, style: TableStyle) -> None:
        """Add a custom table style, or replace the one with the same name.

        The name may not be a built-in style name (table, PivotTable or slicer); a loaded file's custom
        style with such a name is dropped. Each element type may
        appear once and ``size`` must be at least 1. A ``dxf_index`` that does not name an existing
        differential format (:meth:`WorkbookOoxml.add_dxf`) fails with ``StatusCode.INDEX_OUT_OF_RANGE``.

        Raises:
          FormulonError: with ``code == StatusCode.PRECONDITION_FAILED`` when replacing a style whose new
            ``table`` or ``pivot`` flag excludes a kind that a table, a PivotTable or the default style of that
            kind still uses; nothing changes.
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
                            "dxf_index": _uint(
                                0 if element.dxf_index is None else element.dxf_index, "dxf_index", role="position"
                            ),
                        },
                    )
            name_ptr = scope.cstr(style.name)
            S.TABLE_STYLE.pack(
                LIB,
                ptr,
                {
                    "name": name_ptr,
                    "pivot": _tristate_default(bool(style.pivot)),
                    "table": _tristate_default(bool(style.table)),
                    "elements": elements_ptr,
                    "element_count": _uint(len(style.elements), "element_count"),
                },
            )
            _check(LIB.fm_workbook_set_table_style(h, ptr))

    def remove_table_style(self, name: str) -> None:
        """Remove the custom table style ``name`` (Unicode simple case folding).

        Every table using it takes the default table style and every
        PivotTable using it takes the default PivotTable style. A default
        naming it first reverts to the built-in default
        (``TableStyleMedium2`` / ``PivotStyleLight16``).
        """
        h = self._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            _check(LIB.fm_workbook_remove_table_style(h, name_ptr))

    def get_table_style_defaults(self) -> TableStyleDefaults:
        """Return the default table and pivot style names; a name is ``None`` when unset."""
        d = _read_struct(LIB.fm_workbook_get_table_style_defaults, S.TABLE_STYLE_DEFAULTS, self._require())
        return TableStyleDefaults(
            table_style_name=_opt_cstr(d["table_style_name"]),
            pivot_style_name=_opt_cstr(d["pivot_style_name"]),
        )

    def update_table_style_defaults(
        self,
        *,
        table_style_name: Union[str, Unchanged, None] = UNCHANGED,
        pivot_style_name: Union[str, Unchanged, None] = UNCHANGED,
    ) -> None:
        """Set the default style names; an omitted argument keeps its name and ``None`` removes it.

        Raises:
          InvalidArgumentError: for an empty name, a custom style that does not apply to the kind (tables for
            ``table_style_name``, PivotTables for ``pivot_style_name``) or a built-in slicer style.
        """
        fields: Dict[str, object] = {}
        h = self._require()
        with _ArgScope() as scope:
            for key, value in (("table_style_name", table_style_name), ("pivot_style_name", pivot_style_name)):
                if value is UNCHANGED:
                    continue
                if value is None:
                    fields[key + "_op"] = _PATCH_CLEAR
                else:
                    fields[key + "_op"] = _PATCH_SET
                    fields[key] = scope.cstr(value)
            ptr = scope.struct(S.TABLE_STYLE_DEFAULTS_PATCH)
            S.TABLE_STYLE_DEFAULTS_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_workbook_update_table_style_defaults(h, ptr))

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

    def set_theme_colors(self, colors: Union[ThemeColorsInput, ThemeColors]) -> None:
        """Replace the theme's 12 scheme colours (``a:clrScheme`` order; alpha ignored).

        A :class:`ThemeColors` read back is accepted; its ``source`` is ignored.
        """
        if not isinstance(colors, (ThemeColorsInput, ThemeColors)):
            raise TypeError(f"expected a ThemeColorsInput, got {type(colors).__name__}")
        h = self._require()
        values = [_uint(c, "colors") for c in colors.colors]
        if len(values) != 12:
            raise ValueError(f"colors must hold 12 entries, got {len(values)}")
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

    def list_theme_script_fonts(self, scheme: FontScheme) -> List[ThemeScriptFont]:
        """Return the script typefaces of the major (``FontScheme.MAJOR``) or minor (``FontScheme.MINOR``) font
        set in file order; empty without a theme part.

        Raises:
          FormulonError: ``StatusCode.INVALID_ARGUMENT`` for ``FontScheme.NONE`` or any other value.
        """
        h = self._require()
        count = _read_count(LIB.fm_workbook_get_theme_script_font_count, h, _sint(scheme, "scheme"))
        fonts: List[ThemeScriptFont] = []
        with _ArgScope() as scope:
            script_ptr = scope.out_ptr()
            face_ptr = scope.out_ptr()
            for index in range(count):
                _check(
                    LIB.fm_workbook_get_theme_script_font_at(
                        h, _sint(scheme, "scheme"), _uint(index, "idx", role="position"), script_ptr, face_ptr
                    )
                )
                fonts.append(
                    ThemeScriptFont(
                        scheme=_enum(FontScheme, scheme),
                        script=LIB.read_cstr(LIB.read_u32(script_ptr)),
                        typeface=LIB.read_cstr(LIB.read_u32(face_ptr)),
                    )
                )
        return fonts

    def set_theme_script_font(self, font: ThemeScriptFont) -> None:
        """Set the typeface of one script in the major (heading) or minor (body) font set.

        ``font.scheme`` picks the heading set (``FontScheme.MAJOR``) or the body set
        (``FontScheme.MINOR``); ``FontScheme.NONE`` raises. ``script`` is an ISO 15924 code
        of exactly four ASCII letters (``Jpan``, ``Hang``); ``typeface`` must be non-empty. Other
        scripts are unchanged and an unlisted script is appended. Setting
        ``Jpan`` also changes ``ThemeFonts.*_east_asian``. A workbook without a
        theme part first gets one carrying the full default set. Invalid
        arguments raise ``FormulonError``.
        """
        if not isinstance(font, ThemeScriptFont):
            raise TypeError(f"expected a ThemeScriptFont, got {type(font).__name__}")
        h = self._require()
        with _ArgScope() as scope:
            script_ptr = scope.cstr(font.script)
            face_ptr = scope.cstr(font.typeface)
            _check(LIB.fm_workbook_set_theme_script_font(h, _sint(font.scheme, "scheme"), script_ptr, face_ptr))

    def reset_theme(self) -> None:
        """Restore the default Office theme by removing the theme part.

        :meth:`get_theme_colors` then reports ``ThemeSource.DEFAULT``. Excel adds a
        theme part again when it saves the file.
        """
        h = self._require()
        _check(LIB.fm_workbook_reset_theme(h))

    def apply_theme(self, bytes: Union[bytes, bytearray, memoryview]) -> None:
        """Apply a theme file (``.thmx``) as Excel's Page Layout > Themes > Browse does.

        The theme part is replaced together with the images the theme uses; the
        previous theme's images are dropped. A variant in the file is ignored
        (the default theme is applied). Fonts that follow the theme switch to the
        new theme's faces, and formulas are recalculated when the theme's font
        faces change. The call is atomic: on failure the workbook is unchanged.
        Windows Excel's behaviour is unmeasured.

        Raises:
          InvalidArgumentError: when ``bytes`` is empty.
          InvalidFileError: when ``bytes`` is not a zip archive
            (``StatusCode.IO_ZIP_CORRUPT``), has no theme or a missing target
            (``IO_RELATIONSHIP_BROKEN``), or holds a theme that cannot be parsed
            (``IO_XML_PARSE``).
          SecurityError: when the archive exceeds 64 MiB.
        """
        h = self._require()
        raw = builtins.bytes(bytes)
        with _ArgScope() as scope:
            # An empty input still passes a real pointer: the engine reports
            # it as empty rather than as a missing buffer.
            buf = scope.buffer(raw) if raw else scope.alloc(1)
            _check(LIB.fm_workbook_apply_theme(h, buf, _uint(len(raw), "len")))

    def save_theme(self) -> bytes:
        """Write the current theme as a theme file (``.thmx``) that :meth:`apply_theme` accepts.

        The package holds the theme part (the default Office theme when the
        workbook has none) and the images it uses, without a thumbnail.

        Raises:
          InvalidFileError: when the workbook's theme part cannot be parsed
            (``StatusCode.IO_XML_PARSE``).
        """
        h = self._require()
        return _read_owned_buffer(LIB.fm_workbook_save_theme, h)

    def resolve_color(self, spec: Color, context: ColorContext) -> ResolvedColor:
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
            self._pivot_cache_handle(
                _read_count(LIB.fm_workbook_get_pivot_cache_id_at, h, _uint(index, "idx", role="position"))
            )
            for index in range(count)
        ]

    def pivot_cache(self, cache_id: PivotCacheRef) -> PivotCache:
        """Return the handle of the pivot cache ``cache_id``.

        Every call for the same cache id returns the same :class:`PivotCache` object.

        Raises:
          NotFoundError: when the workbook has no cache with that id.
          InvalidArgumentError: when ``cache_id`` is a :class:`PivotCache` of another workbook.
        """
        wanted = _uint(self._pivot_cache_id(cache_id), "cache_id")
        h = self._require()
        for index in range(_read_count(LIB.fm_workbook_get_pivot_cache_count, h)):
            if _read_count(LIB.fm_workbook_get_pivot_cache_id_at, h, _uint(index, "idx", role="position")) == wanted:
                return self._pivot_cache_handle(wanted)
        raise FormulonError(
            StatusCode.NOT_FOUND,
            f"the workbook has no pivot cache {wanted}",
            f"cache_id={wanted}",
        )

    def add_pivot_cache(self, requested_id: Optional[int] = None) -> PivotCache:
        """Create a new empty pivot cache and return its handle.

        Its id is ``requested_id``, else one the engine assigns; a cache id is never 0.
        A new cache has no worksheet source: call
        :meth:`PivotCache.set_worksheet_source` before a pivot uses it. Saving
        fails while a pivot uses a cache without a worksheet source. A cache no
        pivot or slicer uses is not written.

        Raises:
          InvalidArgumentError: when ``requested_id`` is 0.
        """
        h = self._require()
        if requested_id is None:
            requested_id = 0
        elif _uint(requested_id, "requested_id") == 0:
            raise FormulonError(StatusCode.INVALID_ARGUMENT, "a pivot cache id is never 0", "requested_id=0")
        return self._pivot_cache_handle(
            _read_count(LIB.fm_workbook_add_pivot_cache, h, _uint(requested_id, "requested_id"))
        )

    def _pivot_cache_handle(self, cache_id: int) -> PivotCache:
        """The handle of cache id ``cache_id``, whether or not the workbook has that cache now."""
        from ._pivot_cache import PivotCache

        handle = self._pivot_caches.get(cache_id)
        if handle is None:
            handle = PivotCache._new(self, cache_id)
            self._pivot_caches[cache_id] = handle
        return handle

    def _pivot_cache_id(self, cache: PivotCacheRef) -> int:
        """The cache id a cache id or a :class:`PivotCache` of this workbook names."""
        from ._pivot_cache import PivotCache

        if isinstance(cache, PivotCache):
            if cache.workbook is not self:
                raise FormulonError(
                    StatusCode.INVALID_ARGUMENT,
                    "the pivot cache belongs to another workbook",
                    "",
                )
            return cache.cache_id
        if isinstance(cache, bool):
            raise TypeError("cache_id must be an int or a PivotCache, got bool")
        try:
            return operator.index(cache)
        except TypeError:
            raise TypeError(f"cache_id must be an int or a PivotCache, got {type(cache).__name__}") from None

    def remove_pivot_cache(self, cache_id: PivotCacheRef) -> None:
        """Remove the pivot cache ``cache_id``.

        Its handle raises :class:`NotFoundError` until a cache with the same id
        is added.

        Raises:
          InvalidArgumentError: when ``cache_id`` is a :class:`PivotCache` of another workbook.
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` while a pivot table still references the cache.
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
          InvalidArgumentError: when the base is outside the grid.
          TypeError: when ``r1c1`` is not a bool, or only one of ``base_row`` and ``base_col`` is given.
        """
        fields = _notation_fields(r1c1=r1c1, base_row=base_row, base_col=base_col)
        h = self._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.REFERENCE)
            spec = scope.struct(S.REFERENCE_NOTATION_SPEC)
            S.REFERENCE_NOTATION_SPEC.pack(LIB, spec, fields)
            _check(LIB.fm_workbook_parse_reference(h, scope.cstr(text), spec, ptr))
            d = S.REFERENCE.unpack(LIB, ptr)
            return ParsedReference(
                sheet_name=_opt_cstr(d["sheet_name"]),
                range=_cell_rect_from(d["rect"]),
            )

    def format_reference(
        self,
        range: Union[ParsedReference, CellRect, str],
        *,
        sheet_name: Optional[str] = None,
        absolute: bool = False,
    ) -> str:
        """Format ``range`` as A1 reference text (``B:D``, ``2:11``, ``'My Sheet'!A1``).

        ``range`` is a :class:`CellRect`, an unqualified A1 range, or a
        :class:`ParsedReference`, which supplies its own ``sheet_name``.

        Raises:
          TypeError: when ``sheet_name`` is given with a :class:`ParsedReference`.
          FormulonError: with ``StatusCode.INVALID_ARGUMENT`` for a coordinate outside the grid, a reversed rectangle
            or a sheet-qualified string.
        """
        if isinstance(range, ParsedReference):
            if sheet_name is not None:
                raise TypeError("sheet_name cannot be given with a ParsedReference")
            sheet_name, rect = range.sheet_name, range.range
        else:
            from ._inputs import cell_rect

            rect = cell_rect(self, range, "format_reference range")
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
                    int(_require_bool(absolute, "absolute")),
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
            _check(LIB.fm_workbook_get_external_link_at(h, _uint(index, "index", role="position"), ptr))
            d = S.EXTERNAL_LINK_RECORD.unpack(LIB, ptr)
            return ExternalLink(
                book_number=d["book_number"],
                rel_id=LIB.read_cstr(d["rel_id"]),
                part_path=LIB.read_cstr(d["part_path"]),
                target=LIB.read_optional_cstr(d["target"]),
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
class WorkbookOoxml(_NoCopy):
    """The OOXML-level layer of a workbook, :attr:`Workbook.ooxml`.

    The style tables a cell XF indexes (fonts, fills, borders, number
    formats, ``<cellXfs>``, ``<cellStyleXfs>``, DXFs) and the passthrough
    parts. Most code uses :meth:`Sheet.update_cell_format` and
    :meth:`Sheet.get_cell_format` instead.
    """

    __slots__ = ("_workbook",)

    def __init__(self, *args: object, **kwargs: object) -> None:
        raise TypeError("obtain the OOXML layer from Workbook.ooxml")

    @classmethod
    def _new(cls, workbook: Workbook) -> "WorkbookOoxml":
        layer = object.__new__(cls)
        layer._workbook = workbook
        return layer

    def __repr__(self) -> str:
        return f"{self._workbook!r}.ooxml"

    def list_passthroughs(self) -> List[str]:
        """Return the passthrough OOXML part paths."""
        h = self._workbook._require()
        n = _read_count(LIB.fm_workbook_get_passthrough_count, h)
        # Scratch hoisted out of the loop; see list_defined_names.
        with _ArgScope() as scope:
            path_ptr = scope.out_ptr()
            out: List[str] = []
            for i in range(n):
                status = LIB.fm_workbook_get_passthrough_at(h, _uint(i, "idx", role="position"), path_ptr)
                _check(status)
                out.append(LIB.read_cstr(LIB.read_u32(path_ptr)))
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
            LIB.fm_workbook_get_cell_xf_at,
            S.CELL_XF,
            xf_index,
            "xf_index",
            self._workbook._decode_cell_xf,
        )

    def get_font(self, font_index: int) -> Font:
        """Return the font record at ``font_index``."""
        return self._workbook._get_style_record(
            LIB.fm_workbook_get_font_at, S.FONT_RECORD, font_index, "font_index", _decode_font
        )

    def get_fill(self, fill_index: int) -> Fill:
        """Return the fill record at ``fill_index``."""
        return self._workbook._get_style_record(
            LIB.fm_workbook_get_fill_at, S.FILL_RECORD, fill_index, "fill_index", _decode_fill
        )

    def get_border(self, border_index: int) -> Border:
        """Return the border record at ``border_index``."""
        return self._workbook._get_style_record(
            LIB.fm_workbook_get_border_at,
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
            _check(LIB.fm_workbook_get_dxf_at(h, _uint(dxf_index, "dxf_index", role="position"), ptr))
            d = S.DXF_RECORD.unpack(LIB, ptr)
            offsets = S.DXF_RECORD.offsets
            font = _decode_font(ptr + offsets["font"][1]) if d["has_font"] else None
            fill = _decode_fill(ptr + offsets["fill"][1]) if d["has_fill"] else None
            return Dxf(
                font=font,
                fill=fill,
                border=_decode_border_record(ptr + offsets["border"][1]) if d["has_border"] else None,
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

    def list_number_formats(self) -> List[NumberFormat]:
        """Every declared number format in ascending ``num_fmt_id`` order: custom ids and stored overrides of
        built-in ids. Built-in ids without a stored code are not listed."""
        wb = self._workbook
        h = wb._require()
        n = wb._style_count(LIB.fm_workbook_get_number_format_count)
        with _ArgScope() as scope:
            ptr = scope.struct(S.NUMBER_FORMAT_INFO)
            out: List[NumberFormat] = []
            for i in range(n):
                S.zero_struct(LIB, S.NUMBER_FORMAT_INFO, ptr)
                _check(LIB.fm_workbook_get_number_format_at(h, _uint(i, "index", role="position"), ptr))
                d = S.NUMBER_FORMAT_INFO.unpack(LIB, ptr)
                out.append(NumberFormat(num_fmt_id=d["num_fmt_id"], number_format=LIB.read_cstr(d["number_format"])))
        return out

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
            _check(LIB.fm_workbook_set_font(h, _uint(font_index, "font_index", role="position"), ptr))

    def set_fill(self, fill_index: int, record: Fill) -> None:
        """Overwrite an existing fill slot in place, as :meth:`set_font` does for a font.

        Raises:
          SecurityError: when the fill has more than 1024 gradient stops.
        """
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.FILL_RECORD)
            _pack_fill(ptr, record, scope)
            _check(LIB.fm_workbook_set_fill(h, _uint(fill_index, "fill_index", role="position"), ptr))

    def add_fill(self, record: Fill) -> int:
        """Add (dedup) a fill record; return its index.

        Raises:
          SecurityError: when the fill has more than 1024 gradient stops.
        """
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.FILL_RECORD)
            out = scope.out_ptr()
            _pack_fill(ptr, record, scope)
            _check(LIB.fm_workbook_add_fill(h, ptr, out))
            return LIB.read_u32(out)

    def set_border(self, border_index: int, record: Border) -> None:
        """Overwrite an existing border slot in place, as :meth:`set_font` does for a font."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.BORDER_RECORD)
            _pack_border_record(ptr, record)
            _check(LIB.fm_workbook_set_border(h, _uint(border_index, "border_index", role="position"), ptr))

    def add_border(self, record: Border) -> int:
        """Add (dedup) a border record; return its index."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.BORDER_RECORD)
            out = scope.out_ptr()
            _pack_border_record(ptr, record)
            _check(LIB.fm_workbook_add_border(h, ptr, out))
            return LIB.read_u32(out)

    def add_number_format(self, number_format: str) -> int:
        """Add (dedup) a number-format code in the workbook profile's vocabulary; return its resolved id.

        A built-in id whose profile code equals ``number_format`` wins (``m/d/yy`` is 14 under en-US); the ECMA-376
        spelling ``mm-dd-yy`` becomes a custom format.
        """
        h = self._workbook._require()
        with _ArgScope() as scope:
            code_ptr = scope.cstr(number_format)
            out = scope.out_ptr()
            _check(LIB.fm_workbook_add_number_format(h, code_ptr, out))
            return LIB.read_u32(out)

    def add_cell_xf(self, record: Union[CellXfInput, CellXf]) -> int:
        """Add (dedup) an ``<xf>`` record; return its index.

        Out-of-range font, fill or border indices and an unregistered
        ``num_fmt_id`` raise :class:`InvalidArgumentError` rather than growing
        the style tables.
        """
        return self._workbook._add_xf_record(record, LIB.fm_workbook_add_cell_xf)

    def set_cell_xf(self, xf_index: int, record: Union[CellXfInput, CellXf]) -> None:
        """Overwrite an existing ``<cellXfs>`` slot in place, as :meth:`set_font` does for a font: every cell
        naming ``xf_index`` restyles at once. The index must already exist."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CELL_XF)
            S.CELL_XF.pack(LIB, ptr, _cell_xf_fields(record))
            _check(LIB.fm_workbook_set_cell_xf(h, _uint(xf_index, "xf_index", role="position"), ptr))

    def set_cell_style_xf(self, cell_style_xf_index: int, record: Union[CellXfInput, CellXf]) -> None:
        """Overwrite an existing ``<cellStyleXfs>`` slot in place, as :meth:`set_cell_xf` does for a cell xf.

        Every named style naming ``cell_style_xf_index`` takes the record, and when
        the record changes the slot each built-in style naming it gets
        ``custom_builtin``. The index must already exist.

        Raises:
          IndexOutOfRangeError: past the named-style xf count, or for a font, fill or border index past its table.
          InvalidArgumentError: for an unregistered ``num_fmt_id`` or an alignment value outside its range.
        """
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CELL_XF)
            S.CELL_XF.pack(LIB, ptr, _cell_xf_fields(record))
            _check(
                LIB.fm_workbook_set_cell_style_xf(
                    h, _uint(cell_style_xf_index, "cell_style_xf_index", role="position"), ptr
                )
            )

    def add_cell_style_xf(self, record: Union[CellXfInput, CellXf]) -> int:
        """Add (dedup) a named-style ``<xf>``; return its ``<cellStyleXfs>`` id.

        The counterpart of :meth:`get_cell_style_xf`. Named-style xfs are
        the records a ``<cellStyle>`` points at; :meth:`add_cell_xf` writes
        the per-cell ``<cellXfs>`` pool instead.
        """
        return self._workbook._add_xf_record(record, LIB.fm_workbook_add_cell_style_xf)

    def add_dxf(self, input: Union[DxfInput, Dxf]) -> int:  # noqa: A002 -- named after the declaration
        """Add (dedup) a differential format and return its index.

        Accepts a :class:`DxfInput` or the :class:`Dxf` :meth:`get_dxf` returns.
        A ``number_format`` is registered as :meth:`add_number_format` does. A differential format no reference
        uses is not saved; indices are stable only within the live workbook.

        Raises:
          SecurityError: when its fill has more than 1024 gradient stops.
        """
        if not isinstance(input, (DxfInput, Dxf)):
            raise TypeError(f"expected a DxfInput, got {type(input).__name__}")
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.DXF_RECORD)
            out = scope.out_ptr()
            _pack_dxf(ptr, input, scope)
            _check(LIB.fm_workbook_add_dxf(h, ptr, out))
            return LIB.read_u32(out)

    def set_dxf(self, dxf_index: int, input: Union[DxfInput, Dxf]) -> None:  # noqa: A002 -- named after the declaration
        """Overwrite an existing differential format in place, as :meth:`set_font` does for a font: every
        conditional format naming ``dxf_index`` restyles at once. Accepts what :meth:`add_dxf` accepts, and
        raises ``SecurityError`` for a fill with more than 1024 gradient stops."""
        if not isinstance(input, (DxfInput, Dxf)):
            raise TypeError(f"expected a DxfInput, got {type(input).__name__}")
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.DXF_RECORD)
            _pack_dxf(ptr, input, scope)
            _check(LIB.fm_workbook_set_dxf(h, _uint(dxf_index, "dxf_index", role="position"), ptr))

    def get_cell_style_xf(self, cell_style_xf_index: int) -> CellXf:
        """Return the named-style xf record at ``cell_style_xf_index``."""
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CELL_XF)
            _check(
                LIB.fm_workbook_get_cell_style_xf_at(
                    h, _uint(cell_style_xf_index, "cell_style_xf_index", role="position"), ptr
                )
            )
            return self._workbook._decode_cell_xf(ptr)


def _pack_dxf(ptr: int, record: Union[DxfInput, Dxf], scope: _ArgScope) -> None:
    """Pack ``record`` into the ``fm_dxf_t`` at ``ptr``."""
    S.DXF_RECORD.pack(
        LIB,
        ptr,
        {
            "has_font": 1 if record.font is not None else 0,
            "has_fill": 1 if record.fill is not None else 0,
            "has_border": 1 if record.border is not None else 0,
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


_install_workbook_locks(WorkbookOoxml)
# </ooxml-namespace>


_install_workbook_locks(Workbook)


# -- Images (process-wide) ---------------------------------------------------
@_lock_workbook_method
def get_image_info(
    bytes: Union[bytes, bytearray, memoryview],  # noqa: A002 -- named after the declaration
) -> ImageInfo:
    """Sniff the format and pixel size of PNG, JPEG, GIF or BMP bytes from the header alone.

    Raises:
      IoError: with ``StatusCode.IO_IMAGE_UNSUPPORTED`` when the bytes are
        not a recognised image, the header is truncated, or a dimension is zero.
    """
    raw = builtins.bytes(bytes)
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
        sig = LIB.read_optional_cstr(d["signature_template"])
        desc = LIB.read_optional_cstr(d["description"])
        # An unbounded variadic has no upper bound: expose it as None.
        max_arity = d["max_arity"] if d["has_max_arity"] else None
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


def _notation_fields(*, r1c1: bool, base_row: Optional[int], base_col: Optional[int]) -> Dict[str, int]:
    """The fields of an ``fm_reference_notation_spec_t`` from the keyword arguments.

    Raises:
      TypeError: when ``r1c1`` is not a boolean, or only one of ``base_row`` and ``base_col`` is given.
    """
    if (base_row is None) != (base_col is None):
        raise TypeError(_argument_message("base_row and base_col", "must be given together"))
    return {
        "r1c1": _tristate_default(_require_bool(r1c1, "r1c1")),
        "has_base": 0 if base_row is None else 1,
        "base_row": _uint(base_row or 0, "base_row", role="coordinate"),
        "base_col": _uint(base_col or 0, "base_col", role="coordinate"),
    }


def _map_formula_text(
    fn_name: str, formula: str, profile_id: str, *, r1c1: bool, base_row: Optional[int], base_col: Optional[int]
) -> str:
    """Call ``fm_localize_formula`` / ``fm_canonicalize_formula`` with the notation keyword arguments."""
    fields = _notation_fields(r1c1=r1c1, base_row=base_row, base_col=base_col)
    with _ArgScope() as scope:
        text_ptr = scope.cstr(formula)
        pid_ptr = scope.cstr(profile_id)
        spec = scope.struct(S.REFERENCE_NOTATION_SPEC)
        S.REFERENCE_NOTATION_SPEC.pack(LIB, spec, fields)
        return _read_cstr_out(getattr(LIB, fn_name), text_ptr, pid_ptr, spec)


@_lock_workbook_method
def list_function_names() -> List[str]:
    """Return the canonical names of every registered Formulon function, sorted."""
    return [
        _read_cstr_out(LIB.fm_get_function_name_at, _uint(index, "idx", role="position"))
        for index in range(_read_count(LIB.fm_get_function_count))
    ]


@_lock_workbook_method
def localize_function_name(canonical_name: str, profile_id: ProfileId) -> str:
    """Return the display name of ``canonical_name`` under ``profile_id``.

    Raises:
      NotFoundError: when ``canonical_name`` is not a registered function.
      InvalidArgumentError: when ``profile_id`` is not a documented profile.
    """
    return _map_profile_text("fm_localize_function_name", canonical_name, profile_id)


@_lock_workbook_method
def canonicalize_function_name(localized_name: str, profile_id: ProfileId) -> str:
    """Return the canonical English name for ``localized_name`` under ``profile_id``.

    Raises:
      NotFoundError: when ``localized_name`` matches no function under the profile.
      InvalidArgumentError: when ``profile_id`` is not a documented profile.
    """
    return _map_profile_text("fm_canonicalize_function_name", localized_name, profile_id)


@_lock_workbook_method
def localize_formula(
    formula: str,
    profile_id: ProfileId,
    *,
    r1c1: bool = False,
    base_row: Optional[int] = None,
    base_col: Optional[int] = None,
) -> str:
    """Rewrite stored (en-invariant) formula text into ``profile_id``'s spelling.

    Function names, list / array separators, decimal points, booleans
    and error names take the locale form; sheet qualifiers are requoted by
    the profile's display rule; strings, external qualifiers and
    structured-reference brackets are kept verbatim.

    With ``r1c1=True`` every reference is written in R1C1 with the profile's
    row / column letters and brackets (de ``Z(-1)S(1)``), relative references
    measured from ``base_row`` / ``base_col``, which are given together or not
    at all.

    Raises:
      FormulaSyntaxError: when ``formula`` does not parse, or with
        ``StatusCode.PARSER_INVALID_REFERENCE`` when an R1C1 relative reference has no base;
        ``parse_error`` locates the failure as code-point offsets into ``formula``.
      InvalidArgumentError: when ``profile_id`` is not a documented profile, or the base is outside the grid.
      TypeError: when ``r1c1`` is not a boolean, or only one of ``base_row`` and ``base_col`` is given.
    """
    return _map_formula_text(
        "fm_localize_formula", formula, profile_id, r1c1=r1c1, base_row=base_row, base_col=base_col
    )


@_lock_workbook_method
def canonicalize_formula(
    formula: str,
    profile_id: ProfileId,
    *,
    r1c1: bool = False,
    base_row: Optional[int] = None,
    base_col: Optional[int] = None,
) -> str:
    """Rewrite ``profile_id``-spelled formula text into the en-invariant form.

    With ``r1c1=True`` the profile's R1C1 references are read back into A1,
    relative references measured from ``base_row`` / ``base_col``.

    Raises:
      FormulaSyntaxError: when ``formula`` does not parse under the profile, or with
        ``StatusCode.PARSER_INVALID_REFERENCE`` when an R1C1 reference is malformed or relative with no
        base; ``parse_error`` locates the failure as code-point offsets into ``formula``, the text as passed.
      InvalidArgumentError: when ``profile_id`` is not a documented profile, or the base is outside the grid.
      TypeError: when ``r1c1`` is not a boolean, or only one of ``base_row`` and ``base_col`` is given.
    """
    return _map_formula_text(
        "fm_canonicalize_formula", formula, profile_id, r1c1=r1c1, base_row=base_row, base_col=base_col
    )


def _read_locale_error_names(scope: _ArgScope, pid_ptr: int) -> List[LocaleErrorName]:
    canonical_out = scope.out_ptr()
    localized_out = scope.out_ptr()
    measured_out = scope.out_ptr()
    errors = []
    for index in range(_read_count(LIB.fm_get_locale_error_name_count)):
        _check(
            LIB.fm_get_locale_error_name_at(
                pid_ptr, _uint(index, "error_index", role="position"), canonical_out, localized_out, measured_out
            )
        )
        errors.append(
            LocaleErrorName(
                canonical=LIB.read_cstr(LIB.read_u32(canonical_out)),
                localized=LIB.read_cstr(LIB.read_u32(localized_out)),
                measured=LIB.read_u32(measured_out) != 0,
            )
        )
    return errors


@_lock_workbook_method
def list_profile_ids() -> List[ProfileId]:
    """Return every supported Excel profile id (28), in the engine's order."""
    count = _read_count(LIB.fm_get_profile_id_count)
    return [
        cast(ProfileId, _read_cstr_out(LIB.fm_get_profile_id_at, _uint(index, "idx", role="position")))
        for index in range(count)
    ]


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
    * Feed it kind-preserving values (:meth:`Sheet.get_typed_values`), which
      keep a blank distinct from 0.

    Raises:
      TypeError: when ``values`` is not a sequence or an element is not a
        :class:`Value`.
      FormulonError: on an unknown ``profile_id``, an out-of-range
        ``strength``, a non-finite number or an out-of-range error code.

    Example::

        column = [v for row in ws.get_typed_values("A2:A100") for v in row]
        ids = formulon.get_collation_classes(column, formulon.CollationStrength.DYNAMIC, "win-365-en_US").class_ids
        seen = set()
        distinct = [v for v, i in zip(column, ids) if i not in seen and not seen.add(i)]
    """
    if isinstance(values, (str, bytes)) or not isinstance(values, Sequence):
        raise TypeError(_argument_message("values", "must be an array"))
    values = [value_input(v, f"values[{i}]") for i, v in enumerate(values)]
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
from ._inputs import cell_rect, civil_datetime, color_input, value_input  # noqa: E402
from ._pivot import Pivot  # noqa: E402
from ._pivot_cache import PivotCache  # noqa: E402
from ._sheet import CellValueInput, Sheet  # noqa: E402
from ._table import Table  # noqa: E402
