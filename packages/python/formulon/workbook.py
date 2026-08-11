"""Pythonic wrapper around the Formulon C ABI (WASM transport).

Public surface:
  * :class:`Workbook` -- context-manager around an ``fm_workbook_t*`` handle.
  * :class:`Value` -- frozen dataclass mirroring ``fm_value_t``.
  * :class:`FormulonError` -- exception raised on host-side failures.
  * :class:`Cell`, :class:`DefinedName`, :class:`Table` -- iteration tuples.

Cell-level Excel errors (``#DIV/0!``, ``#VALUE!``, etc.) are reported as
``Value`` instances with ``kind == ValueKind.ERROR``; only host-side
problems (NULL handle, parse failure during ``Workbook.load``, OOM)
raise :class:`FormulonError`.

The transport is WebAssembly (``formulon_capi.wasm``) loaded via
``wasmtime``; see :mod:`formulon._c` for the low-level memory and call
plumbing.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from enum import IntEnum
from typing import Dict, Iterator, List, NamedTuple, Optional, Sequence, Union

from . import _structs as S
from ._c import LIB, ValueKind, fm_value_t_size

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
    "ColumnLayout",
    "Comment",
    "ConditionalFormat",
    "ConditionalFormatInput",
    "ColorScale",
    "DataBar",
    "DataValidation",
    "DataValidationInput",
    "DefinedName",
    "DifferentialFormat",
    "ExternalLink",
    "FillRecord",
    "FontRecord",
    "FormulonError",
    "FunctionMetadata",
    "Hyperlink",
    "IconSet",
    "MergeRange",
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
    "PivotReportLayout",
    "PivotShowValuesAs",
    "PivotWorksheetSource",
    "RowLayout",
    "SaveDiagnostics",
    "SheetProtection",
    "SheetView",
    "SpillInfo",
    "Table",
    "Value",
    "ValueKind",
    "Workbook",
    "XlsbReadDiagnostics",
]


# ---------------------------------------------------------------------------
# Error type
# ---------------------------------------------------------------------------

# `formulon::FormulonErrorCode::kNotFound`. The C ABI intentionally exposes
# status codes as integers, so bindings retain this matching stable ordinal.
_STATUS_NOT_FOUND = 6


class FormulonError(Exception):
    """Raised for host-side failures of the Formulon C ABI.

    Attributes:
      status: numeric ``fm_status_t`` (matches
        ``formulon::FormulonErrorCode`` ordinals).
      status_name: symbolic name (e.g. ``"kInvalidArgument"``) returned
        by ``fm_status_string``.
      message: thread-local diagnostic message captured from
        ``fm_last_error_message`` at the time the failure was detected.
      context: optional context string captured from
        ``fm_last_error_context``.

    The message and context are copied out of the thread-local WASM
    buffers eagerly, so they remain valid even after subsequent WASM
    calls overwrite those buffers.
    """

    def __init__(self, status: int, *, op: str = "") -> None:
        self.status = int(status)
        self.status_name = LIB.read_cstr(LIB.fm_status_string(self.status))
        self.message, self.context = LIB.last_diagnostic(self.status)
        prefix = f"{op}: " if op else ""
        text = f"{prefix}{self.status_name} ({self.status})"
        if self.message:
            text += f": {self.message}"
        if self.context:
            text += f" [{self.context}]"
        super().__init__(text)


@dataclass(frozen=True)
class SaveDiagnostics:
    """Bytes and loss counters returned by ``save_ex_with_diagnostics``."""

    bytes: bytes
    downgraded_formula_count: int
    deferred_feature_count: int


@dataclass(frozen=True)
class XlsbReadDiagnostics:
    """Loss and recovery counters captured while loading an XLSB workbook."""

    undecoded_formula_count: int
    undecoded_defined_name_count: int
    dropped_part_count: int


def _check(status: int, op: str) -> None:
    """Raise :class:`FormulonError` if ``status`` is non-zero."""
    if status != 0:
        raise FormulonError(status, op=op)


# ---------------------------------------------------------------------------
# Value POD
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Value:
    """Pythonic mirror of ``fm_value_t``.

    Exactly one of ``number`` / ``boolean`` / ``text`` / ``error_code`` is
    populated according to ``kind``. The constructor enforces the
    invariant.
    """

    kind: ValueKind
    number: Optional[float] = None
    boolean: Optional[bool] = None
    text: Optional[str] = None
    error_code: Optional[int] = None

    @classmethod
    def _from_wasm(cls, ptr: int) -> "Value":
        """Decode an ``fm_value_t`` stored at WASM offset ``ptr``.

        Layout (wasm32): int32 kind | 4 pad | union (8 bytes).

        Reserved kinds (ARRAY / REF / LAMBDA) currently surface as bare
        ``Value(kind=...)`` with no payload; they are reserved in the C
        ABI for a later bundle.
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
            return cls(kind=kind, error_code=int(error_code))
        # BLANK / ARRAY / REF / LAMBDA: no scalar payload.
        return cls(kind=kind)

    def to_python(self) -> Union[None, float, bool, str, "Value"]:
        """Convert to a native Python type.

        - ``BLANK`` -> ``None``
        - ``NUMBER`` -> ``float``
        - ``BOOL`` -> ``bool``
        - ``TEXT`` -> ``str``
        - ``ERROR`` -> the :class:`Value` itself (Excel cell errors are
          values, not Python exceptions; callers inspect ``error_code``)
        - ``ARRAY`` / ``REF`` / ``LAMBDA`` -> the :class:`Value` itself
          (passthrough, payload not yet exposed by the C ABI)
        """
        if self.kind is ValueKind.BLANK:
            return None
        if self.kind is ValueKind.NUMBER:
            return self.number  # type: ignore[return-value]
        if self.kind is ValueKind.BOOL:
            return self.boolean  # type: ignore[return-value]
        if self.kind is ValueKind.TEXT:
            return self.text  # type: ignore[return-value]
        # ERROR / ARRAY / REF / LAMBDA: surface the wrapping Value.
        return self


# ---------------------------------------------------------------------------
# Iteration tuples
# ---------------------------------------------------------------------------


class Cell(NamedTuple):
    """One row of ``Workbook.iter_cells()``."""

    row: int
    col: int
    formula: Optional[str]
    value: Value


class DefinedName(NamedTuple):
    """One row of ``Workbook.iter_defined_names()``."""

    name: str
    formula: str
    local_sheet_id: int


class Table(NamedTuple):
    """One row of ``Workbook.iter_tables()``."""

    name: str
    display_name: str
    ref: str
    sheet_index: int


class PassthroughPart(NamedTuple):
    """One row of ``Workbook.iter_passthrough()``."""

    path: str


# ---------------------------------------------------------------------------
# Enumerations (mirror the C ABI ordinals)
# ---------------------------------------------------------------------------


class CalcMode(IntEnum):
    """Workbook-level calc mode (``<calcPr calcMode>`` policy).

    Plain round-trip metadata: the engine recalcs all dirty cells on every
    :meth:`Workbook.recalc` regardless of this value.
    """

    AUTO = 0
    MANUAL = 1
    AUTO_NO_TABLE = 2


class PivotAxis(IntEnum):
    """PivotTable axis for a field."""

    ROW = 0
    COL = 1
    VALUE = 2
    PAGE = 3


class WorkbookFormat(IntEnum):
    """Container format selector for :meth:`Workbook.save_ex`.

    Mirrors ``fm_workbook_format_t``. ``UNKNOWN`` is not a valid save
    target.
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


class PivotFilterType(IntEnum):
    """Filter type for an active (slicer-applied) pivot filter."""

    VALUE_TOP_10 = 0
    VALUE_GREATER_THAN = 1
    VALUE_BETWEEN = 2
    LABEL_CONTAINS = 3
    LABEL_BEGINS_WITH = 4
    LABEL_DATE = 5


class PivotFilterValueKind(IntEnum):
    """Discriminator for a pivot filter spec's variant payload."""

    NONE = -1
    INT = 0
    DOUBLE = 1
    TEXT = 2


class PivotDateGrouping(IntEnum):
    """Date-grouping granularity for a pivot field."""

    DAY = 0
    MONTH = 1
    QUARTER = 2
    YEAR = 3
    WEEK = 4
    HOUR = 5
    MINUTE = 6
    SECOND = 7


class PivotCalendar(IntEnum):
    """Calendar system used by pivot date grouping."""

    GREGORIAN = 0
    JAPANESE = 1


class PivotCellKind(IntEnum):
    """Discriminator for a projected PivotTable layout cell."""

    HEADER = 0
    ROW_LABEL = 1
    COL_LABEL = 2
    DATA = 3
    ROW_SUBTOTAL = 4
    COL_SUBTOTAL = 5
    GRAND_TOTAL = 6
    BLANK = 7


class PivotReportLayout(IntEnum):
    """Pivot report layout setting (distinct from :class:`PivotLayout`)."""

    COMPACT = 0
    TABULAR = 1
    OUTLINE = 2


# Sentinel value for ``PivotDataFieldSpec.show_as_base_item`` meaning
# "(previous)" / "(next)" (mirrors the C ABI constants).
PIVOT_SHOW_AS_BASE_PREVIOUS = 1048828
PIVOT_SHOW_AS_BASE_NEXT = 1048829

# Sentinel for ``CellStyle.builtin_id`` meaning "custom (no builtinId)".
CELL_STYLE_BUILTIN_ID_NONE = S.CELL_STYLE_BUILTIN_ID_NONE


# ---------------------------------------------------------------------------
# Structured value / input dataclasses
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class MergeRange:
    """Inclusive cell rectangle (0-based corners)."""

    first_row: int
    first_col: int
    last_row: int
    last_col: int


@dataclass(frozen=True)
class Hyperlink:
    """A sheet hyperlink anchored at ``(row, col)``."""

    row: int
    col: int
    target: str
    location: str
    display: str
    tooltip: str


@dataclass(frozen=True)
class Comment:
    """A cell comment."""

    author: str
    text: str


class CommentEntry(NamedTuple):
    """One enumerated sheet comment, including its zero-based anchor."""

    row: int
    col: int
    author: str
    text: str


@dataclass(frozen=True)
class DataValidation:
    """A sheet ``<dataValidation>`` rule (getter result)."""

    ranges: List[MergeRange]
    type: int
    op: int
    error_style: int
    allow_blank: bool
    show_input_message: bool
    show_error_message: bool
    show_dropdown: bool
    formula1: str
    formula2: str
    error_title: str
    error_message: str
    prompt_title: str
    prompt_message: str


@dataclass
class DataValidationInput:
    """Argument shape for :meth:`Workbook.add_validation`.

    Every field except ``type`` defaults to a benign zero/false/empty so
    callers populate only the parts they need. ``show_dropdown`` defaults
    to ``True`` to match the OOXML default (the in-cell dropdown arrow is
    shown for ``list`` validations unless explicitly suppressed).
    """

    type: int
    ranges: List[MergeRange] = field(default_factory=list)
    op: int = 0
    error_style: int = 0
    allow_blank: bool = False
    show_input_message: bool = False
    show_error_message: bool = False
    show_dropdown: bool = True
    formula1: str = ""
    formula2: str = ""
    error_title: str = ""
    error_message: str = ""
    prompt_title: str = ""
    prompt_message: str = ""


@dataclass
class SheetProtection:
    """Sheet ``<sheetProtection>`` flags. Booleans are real ``bool``."""

    enabled: bool = False
    algorithm_name: str = ""
    hash_value: str = ""
    salt_value: str = ""
    spin_count: int = 0
    legacy_password: str = ""
    sheet: bool = False
    objects: bool = False
    scenarios: bool = False
    format_cells: bool = False
    format_columns: bool = False
    format_rows: bool = False
    insert_columns: bool = False
    insert_rows: bool = False
    insert_hyperlinks: bool = False
    delete_columns: bool = False
    delete_rows: bool = False
    select_locked_cells: bool = False
    select_unlocked_cells: bool = False
    sort: bool = False
    auto_filter: bool = False
    pivot_tables: bool = False


@dataclass(frozen=True)
class SheetView:
    """Per-sheet view: zoom, frozen-pane counts, tab-hidden flag, and the
    display / orientation flags mirrored from OOXML ``<sheetView>``."""

    zoom_scale: int
    freeze_rows: int
    freeze_cols: int
    tab_hidden: bool
    show_grid_lines: bool = True
    show_row_col_headers: bool = True
    show_zeros: bool = True
    right_to_left: bool = False
    tab_selected: bool = False
    view_mode: str = ""


@dataclass(frozen=True)
class ColumnLayout:
    """A per-column-range layout override (inclusive ``[first, last]``)."""

    first: int
    last: int
    width: float
    hidden: bool
    outline_level: int


@dataclass(frozen=True)
class RowLayout:
    """A per-row layout override."""

    row: int
    height: float
    hidden: bool
    outline_level: int


@dataclass(frozen=True)
class PaginationResult:
    """Resolved physical pagination for one worksheet.

    ``print_area`` contains inclusive, zero-based ``(first_row, first_col,
    last_row, last_col)`` rectangles.  Break lists contain the zero-based row
    or column a new physical page begins before.
    """

    page_count: int
    print_area: List[tuple[int, int, int, int]]
    horizontal_breaks: List[int]
    vertical_breaks: List[int]


@dataclass(frozen=True)
class CfMatch:
    """A resolved conditional-format match for one cell.

    Active fields depend on ``kind`` (0 differential format, 1 color
    scale, 2 data bar, 3 icon set); the rest carry default-zero values.
    """

    kind: int
    priority: int
    dxf_id_engaged: bool
    dxf_id: int
    color: tuple  # (r, g, b, a)
    bar_length_pct: float
    bar_axis_position_pct: float
    bar_is_negative: bool
    bar_fill: tuple
    bar_border_engaged: bool
    bar_border: tuple
    bar_gradient: bool
    icon_set_name: int
    icon_index: int


@dataclass(frozen=True)
class CfCellResult:
    """One cell's CF matches inside a viewport-range evaluation."""

    row: int
    col: int
    matches: List[CfMatch]


@dataclass(frozen=True)
class CfColor:
    """RGBA color used by a visual conditional-format rule."""

    r: int
    g: int
    b: int
    a: int = 255


@dataclass(frozen=True)
class CfValueObject:
    """One conditional-format threshold (``fm_cfvo_t``)."""

    type: int
    value: Optional[str] = None
    gte: bool = True


@dataclass(frozen=True)
class ColorScale:
    thresholds: List[CfValueObject]
    colors: List[CfColor]


@dataclass(frozen=True)
class DataBar:
    minimum: CfValueObject
    maximum: CfValueObject
    fill: CfColor
    show_value: bool = True
    min_length_pct: int = 10
    max_length_pct: int = 90


@dataclass(frozen=True)
class IconSet:
    name: int
    thresholds: List[CfValueObject]
    reverse: bool = False
    show_value: bool = True
    percent: bool = True


@dataclass(frozen=True)
class ConditionalFormat:
    """A conditional-format rule (getter result, flattened priority order)."""

    id: str
    type: int
    priority: int
    stop_if_true: bool
    sqref: List[MergeRange]
    dxf_id: Optional[int]
    formula1: str
    formula2: str
    op: int
    rank: int
    percent: bool
    bottom: bool
    above_average: bool
    equal_average: bool
    std_dev: float
    text: str
    time_period: int
    color_scale: Optional[ColorScale]
    data_bar: Optional[DataBar]
    icon_set: Optional[IconSet]


@dataclass
class ConditionalFormatInput:
    """Argument shape for :meth:`Workbook.add_conditional_format`.

    Set the corresponding visual payload for color-scale, data-bar, or
    icon-set rules. The C ABI validates each payload's cardinality.
    """

    sqref: List[MergeRange]
    type: int
    priority: int = 0
    stop_if_true: bool = False
    id: str = ""
    dxf_id_engaged: bool = False
    dxf_id: int = 0
    formula1: str = ""
    formula2: str = ""
    op_engaged: bool = False
    op: int = 0
    rank_engaged: bool = False
    rank: int = 0
    percent: bool = False
    bottom: bool = False
    above_average: bool = False
    equal_average: bool = False
    std_dev_engaged: bool = False
    std_dev: float = 0.0
    text: str = ""
    time_period_engaged: bool = False
    time_period: int = 0
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
class SpillInfo:
    """Dynamic-array spill region info for a cell."""

    engaged: bool
    anchor_row: int
    anchor_col: int
    rows: int
    cols: int


@dataclass(frozen=True)
class FunctionMetadata:
    """Catalog metadata for a single Formulon function."""

    name: str
    min_arity: int
    #: ``None`` denotes an unbounded variadic or a lazy / special form
    #: whose upper arity is unknown.
    max_arity: Optional[int]
    availability: int
    signature_template: Optional[str]
    description: Optional[str]


def _cell_xf_has_alignment(record: "CellXf") -> bool:
    return (
        record.has_alignment is True
        or (record.has_horizontal_align if record.has_horizontal_align is not None else record.horizontal_align != 0)
        or (record.has_vertical_align if record.has_vertical_align is not None else record.vertical_align != 2)
        or (record.has_wrap_text if record.has_wrap_text is not None else record.wrap_text)
        or (record.has_justify_last_line if record.has_justify_last_line is not None else record.justify_last_line)
        or record.text_rotation is not None
        or record.indent is not None
        or record.relative_indent is not None
        or record.shrink_to_fit is not None
        or record.reading_order is not None
    )


@dataclass(frozen=True)
class CellXf:
    """A resolved ``<xf>`` style record."""

    font_index: int
    fill_index: int
    border_index: int
    num_fmt_id: int
    horizontal_align: int
    vertical_align: int
    wrap_text: bool
    #: ``None`` infers presence from explicitly supplied alignment values;
    #: ``False`` requests omission when all alignment values are defaults.
    #: Non-default values still imply that an alignment child is required.
    has_alignment: Optional[bool] = None
    justify_last_line: bool = False
    xf_id: int = 0
    text_rotation: Optional[int] = None
    indent: Optional[int] = None
    relative_indent: Optional[int] = None
    shrink_to_fit: Optional[bool] = None
    reading_order: Optional[int] = None
    has_horizontal_align: Optional[bool] = None
    has_vertical_align: Optional[bool] = None
    has_wrap_text: Optional[bool] = None
    has_justify_last_line: Optional[bool] = None


@dataclass
class FontRecord:
    """A font record (``add_font`` input / ``get_font`` result)."""

    name: str = ""
    size: float = 11.0
    color_argb: int = 0
    bold: bool = False
    italic: bool = False
    strike: bool = False
    underline: int = 0


@dataclass
class FillRecord:
    """A fill record."""

    pattern: int = 0
    fg_argb: int = 0
    bg_argb: int = 0


@dataclass
class DifferentialFormat:
    """An optional-style-fragment record used by conditional formats.

    ``border`` uses the same dictionary shape as :meth:`Workbook.add_border`.
    A number format is engaged when ``num_fmt_id`` is not ``None``.
    """

    font: Optional[FontRecord] = None
    fill: Optional[FillRecord] = None
    border: Optional[Dict[str, object]] = None
    num_fmt_id: Optional[int] = None
    num_fmt_code: str = ""


@dataclass(frozen=True)
class CellStyle:
    """A named cell style (``<cellStyle>`` entry)."""

    name: str
    xf_id: int
    builtin_id: int
    i_level: int
    hidden: bool
    custom_builtin: bool


@dataclass(frozen=True)
class ExternalLink:
    """An external-link record (``<externalReferences>`` entry)."""

    index: int
    rel_id: str
    part_path: str
    target: str
    target_external: bool
    kind: int


@dataclass(frozen=True)
class PivotCell:
    """One projected PivotTable layout cell."""

    row: int
    col: int
    value: Value
    kind: int
    depth: int
    field_name: str
    number_format: str


@dataclass(frozen=True)
class PivotLayout:
    """The projected rectangular layout of a PivotTable."""

    top: int
    left: int
    rows: int
    cols: int
    cells: List[PivotCell]


@dataclass(frozen=True)
class PivotWorksheetSource:
    """Optional worksheet-source metadata for a pivot cache."""

    ref: Optional[str] = None
    sheet: Optional[str] = None
    name: Optional[str] = None


@dataclass
class PivotFieldSpec:
    """Argument shape for :meth:`Workbook.pivot_field_add`."""

    source_name: str
    custom_name: str = ""
    axis: int = PivotAxis.ROW
    subtotal_top: bool = False
    number_format: str = ""


@dataclass
class PivotDataFieldSpec:
    """Argument shape for ``pivot_data_field_add`` / ``pivot_data_field_set``.

    ``show_as_base_field`` / ``show_as_base_item`` use ``-1`` for "unset".
    """

    name: str
    field_index: int
    aggregation: int = PivotAggregation.SUM
    number_format: str = ""
    show_as: int = PivotShowValuesAs.NORMAL
    show_as_base_field: int = -1
    show_as_base_item: int = -1


@dataclass
class PivotFilterSpec:
    """Argument shape for :meth:`Workbook.pivot_filter_add`."""

    axis: int
    field_name: str
    type: int
    value_kind: int = PivotFilterValueKind.NONE
    value_int: int = 0
    value_double: float = 0.0
    value_text: str = ""
    value_high_kind: int = PivotFilterValueKind.NONE
    value_high_int: int = 0
    value_high_double: float = 0.0
    data_field_index: int = 0


# ---------------------------------------------------------------------------
# Workbook handle
# ---------------------------------------------------------------------------


def _alloc_out_ptr() -> int:
    """Allocate a 4-byte WASM scratch slot for an out-i32 / out-ptr."""
    ptr = LIB.alloc(4)
    LIB.write_bytes(ptr, b"\x00\x00\x00\x00")
    return ptr


def _read_count(fn, *args) -> int:
    """Call a ``(... , out_count*)`` ABI function and return the count."""
    out = _alloc_out_ptr()
    try:
        _check(fn(*args, out), getattr(fn, "__name__", "count"))
        return LIB.read_u32(out)
    finally:
        LIB.free(out)


def _opt_str_ptr(value: Optional[str], owned: List[int]) -> int:
    """Allocate ``value`` as UTF-8 (NULL for empty/None); track for free."""
    if value is None or value == "":
        return 0
    ptr, _ = LIB.alloc_utf8(value)
    owned.append(ptr)
    return ptr


def _pack_merge_array(ranges: Sequence[MergeRange], owned: List[int]) -> int:
    """Pack a list of ``MergeRange`` into a contiguous WASM array.

    Returns the array pointer (0 when empty). The pointer is appended to
    ``owned`` for later release.
    """
    if not ranges:
        return 0
    size = S.MERGE_RANGE.size
    ptr = LIB.alloc(size * len(ranges))
    owned.append(ptr)
    for i, r in enumerate(ranges):
        S.MERGE_RANGE.pack(
            LIB,
            ptr + i * size,
            {
                "first_row": r.first_row,
                "first_col": r.first_col,
                "last_row": r.last_row,
                "last_col": r.last_col,
            },
        )
    return ptr


class Workbook:
    """Owns an ``fm_workbook_t*`` handle (a WASM i32 offset).

    Construction:
      Use one of :meth:`create_default`, :meth:`create_empty`, or
      :meth:`load`. Direct ``Workbook()`` instantiation is supported as a
      stub for subclassing but does not allocate a handle.

    Lifetime:
      Always release via :meth:`close`, ideally through ``with`` syntax::

          with Workbook.create_default() as wb:
              wb.set_number(0, 0, 0, 42.0)
              wb.recalc()
              print(wb.get_value(0, 0, 0).to_python())

      :meth:`close` is idempotent. The destructor calls ``close`` as a
      defensive net, but relying on garbage-collection ordering is
      unsafe; prefer the context-manager form.

    Threading:
      Each ``Workbook`` instance is single-threaded. The underlying
      wasmtime store is serialised process-wide by an internal lock in
      :mod:`formulon._c`, so cross-thread misuse cannot corrupt WASM
      state -- it just blocks. The parallel recalc scheduler degrades
      to serial under WASM (single-threaded build).
    """

    __slots__ = ("_handle",)

    def __init__(self) -> None:
        # 0 == NULL pointer. Use the factory methods to allocate.
        self._handle: int = 0

    # -- Factories ---------------------------------------------------------
    @classmethod
    def create_default(cls) -> "Workbook":
        """Build a default workbook (a single sheet named ``Sheet1``)."""
        wb = cls()
        out_ptr = _alloc_out_ptr()
        try:
            status = LIB.fm_workbook_create(out_ptr)
            _check(status, "fm_workbook_create")
            wb._handle = LIB.read_u32(out_ptr)
        finally:
            LIB.free(out_ptr)
        return wb

    @classmethod
    def create_empty(cls) -> "Workbook":
        """Build an empty workbook (no sheets).

        Saving an empty workbook fails just as Excel rejects sheet-less
        archives; callers are expected to add at least one sheet before
        :meth:`save`.
        """
        wb = cls()
        out_ptr = _alloc_out_ptr()
        try:
            status = LIB.fm_workbook_create_empty(out_ptr)
            _check(status, "fm_workbook_create_empty")
            wb._handle = LIB.read_u32(out_ptr)
        finally:
            LIB.free(out_ptr)
        return wb

    @classmethod
    def load(cls, data: Union[bytes, bytearray, memoryview]) -> "Workbook":
        """Load a workbook from an in-memory ``.xlsx`` byte buffer.

        Args:
          data: the raw OOXML archive bytes. Accepts any object that
            satisfies the buffer protocol.

        Raises:
          FormulonError: when the input cannot be parsed as a valid
            OOXML archive.
        """
        if not isinstance(data, (bytes, bytearray, memoryview)):
            raise TypeError(f"Workbook.load: expected bytes-like, got {type(data).__name__}")
        buf = bytes(data)
        wb = cls()
        data_ptr = LIB.alloc_bytes(buf) if len(buf) > 0 else 0
        out_ptr = _alloc_out_ptr()
        try:
            status = LIB.fm_workbook_load(data_ptr, len(buf), out_ptr)
            _check(status, "fm_workbook_load")
            wb._handle = LIB.read_u32(out_ptr)
        finally:
            LIB.free(out_ptr)
            if data_ptr:
                LIB.free(data_ptr)
        return wb

    # -- Lifecycle ---------------------------------------------------------
    def close(self) -> None:
        """Release the underlying handle. Idempotent."""
        if self._handle:
            LIB.fm_workbook_destroy(self._handle)
            self._handle = 0

    def __enter__(self) -> "Workbook":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:  # type: ignore[no-untyped-def]
        self.close()

    def __del__(self) -> None:  # pragma: no cover -- GC ordering best-effort
        try:
            self.close()
        except Exception:  # noqa: BLE001 -- destructors must not raise
            pass

    @property
    def is_valid(self) -> bool:
        """``True`` while the handle is still alive."""
        return self._handle != 0

    def _require(self) -> int:
        if not self.is_valid:
            raise FormulonError(
                # 7000-band: bindings / C API. 7000 is kBindingInvalidHandle
                # in src/utils/error.h.
                7000,
                op="Workbook (handle is NULL or already closed)",
            )
        return self._handle

    # -- Sheets ------------------------------------------------------------
    def sheet_count(self) -> int:
        """Number of sheets in the workbook."""
        h = self._require()
        return int(LIB.fm_workbook_sheet_count(h))

    def sheet_name(self, index: int) -> str:
        """Display name (UTF-8) of the sheet at ``index``."""
        h = self._require()
        out_ptr = _alloc_out_ptr()
        try:
            status = LIB.fm_workbook_sheet_name(h, index, out_ptr)
            _check(status, "fm_workbook_sheet_name")
            return LIB.read_cstr(LIB.read_u32(out_ptr))
        finally:
            LIB.free(out_ptr)

    def add_sheet(self, name: str) -> None:
        """Append a new sheet with the given UTF-8 display name."""
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(name)
        try:
            status = LIB.fm_workbook_add_sheet(h, name_ptr)
            _check(status, "fm_workbook_add_sheet")
        finally:
            LIB.free(name_ptr)

    # -- Cell mutation -----------------------------------------------------
    def set_number(self, sheet: int, row: int, col: int, value: float) -> None:
        h = self._require()
        status = LIB.fm_workbook_set_number(h, sheet, row, col, float(value))
        _check(status, "fm_workbook_set_number")

    def set_bool(self, sheet: int, row: int, col: int, value: bool) -> None:
        h = self._require()
        status = LIB.fm_workbook_set_bool(h, sheet, row, col, 1 if value else 0)
        _check(status, "fm_workbook_set_bool")

    def set_error(self, sheet: int, row: int, col: int, error_code: int) -> None:
        h = self._require()
        status = LIB.fm_workbook_set_error(h, sheet, row, col, int(error_code))
        _check(status, "fm_workbook_set_error")

    def set_text(self, sheet: int, row: int, col: int, value: str) -> None:
        h = self._require()
        text_ptr, _ = LIB.alloc_utf8(value)
        try:
            status = LIB.fm_workbook_set_text(h, sheet, row, col, text_ptr)
            _check(status, "fm_workbook_set_text")
        finally:
            LIB.free(text_ptr)

    def set_blank(self, sheet: int, row: int, col: int) -> None:
        h = self._require()
        status = LIB.fm_workbook_set_blank(h, sheet, row, col)
        _check(status, "fm_workbook_set_blank")

    def set_formula(self, sheet: int, row: int, col: int, formula: str) -> None:
        h = self._require()
        formula_ptr, _ = LIB.alloc_utf8(formula)
        try:
            status = LIB.fm_workbook_set_formula(h, sheet, row, col, formula_ptr)
            _check(status, "fm_workbook_set_formula")
        finally:
            LIB.free(formula_ptr)

    # -- Cell read ---------------------------------------------------------
    def get_value(self, sheet: int, row: int, col: int) -> Value:
        """Read the cached value at ``(sheet, row, col)``.

        Reflects the most recent recalc; mutations made after the last
        :meth:`recalc` are not visible to dependents until you call
        :meth:`recalc` again.
        """
        h = self._require()
        value_ptr = LIB.alloc(fm_value_t_size)
        try:
            status = LIB.fm_workbook_get_value(h, sheet, row, col, value_ptr)
            _check(status, "fm_workbook_get_value")
            return Value._from_wasm(value_ptr)
        finally:
            LIB.free(value_ptr)

    # -- Ad-hoc array evaluation ------------------------------------------
    def evaluate_cf_formula(
        self, sheet: int, row: int, col: int, anchor_row: int, anchor_col: int, formula: str
    ) -> bool:
        """Evaluate a conditional-format formula at one cell without mutation."""
        h = self._require()
        formula_ptr, _ = LIB.alloc_utf8(formula)
        value_ptr = LIB.alloc(fm_value_t_size)
        try:
            _check(
                LIB.fm_workbook_evaluate_cf_formula(
                    h, int(sheet), int(row), int(col), int(anchor_row), int(anchor_col), formula_ptr, value_ptr
                ),
                "fm_workbook_evaluate_cf_formula",
            )
            return Value._from_wasm(value_ptr).boolean is True
        finally:
            LIB.free(formula_ptr)
            LIB.free(value_ptr)

    def evaluate_formula_array(self, sheet: int, row: int, col: int, formula: str) -> List[List[Value]]:
        """Evaluate ``formula`` as if entered at ``(sheet, row, col)`` and
        return the whole multi-cell result, without mutating the workbook.

        Unlike a scalar evaluation (which reduces an array to its top-left
        element), a dynamic-array formula such as ``=SEQUENCE(2,3)`` yields
        the full ``rows x cols`` grid as a nested list in row-major order
        (``result[r][c]``); a scalar result such as ``=1+2`` is returned as
        a ``1 x 1`` grid (``[[Value(...)]]``).

        The evaluation is read-only: no cell value is written and no spill
        is committed. Cell-level Excel errors surface as
        :class:`Value` entries with ``kind == ValueKind.ERROR``; only
        host-side failures (NULL handle, out-of-range sheet) raise
        :class:`FormulonError`.

        Args:
          sheet: 0-based sheet index the formula is anchored on.
          row: 0-based anchor row (drives ``ROW()`` and relative refs).
          col: 0-based anchor column (drives ``COLUMN()``).
          formula: the formula text, with or without a leading ``=``.

        Returns:
          A ``rows x cols`` nested list of :class:`Value`.
        """
        h = self._require()
        formula_ptr, _ = LIB.alloc_utf8(formula)
        rows_ptr = _alloc_out_ptr()
        cols_ptr = _alloc_out_ptr()
        try:
            status = LIB.fm_workbook_evaluate_formula_array(h, sheet, row, col, formula_ptr, rows_ptr, cols_ptr)
            _check(status, "fm_workbook_evaluate_formula_array")
            rows = LIB.read_u32(rows_ptr)
            cols = LIB.read_u32(cols_ptr)
        finally:
            LIB.free(formula_ptr)
            LIB.free(rows_ptr)
            LIB.free(cols_ptr)

        # Second step: read each stashed cell back by its row-major index.
        # The stash on the handle stays valid until the next array
        # evaluation or mutation, so it survives the loop below.
        grid: List[List[Value]] = []
        value_ptr = LIB.alloc(fm_value_t_size)
        try:
            for r in range(rows):
                out_row: List[Value] = []
                for c in range(cols):
                    index = r * cols + c
                    status = LIB.fm_workbook_evaluate_formula_array_cell(h, index, value_ptr)
                    _check(status, "fm_workbook_evaluate_formula_array_cell")
                    out_row.append(Value._from_wasm(value_ptr))
                grid.append(out_row)
        finally:
            LIB.free(value_ptr)
        return grid

    # -- Recalc ------------------------------------------------------------
    def recalc(self) -> None:
        """Drive a full incremental recalc.

        Public bindings use the serial recalc contract. The internal C++
        parallel scheduler is not exposed by the C ABI, so native CLI and
        language bindings have the same recalculation semantics.
        """
        h = self._require()
        status = LIB.fm_workbook_recalc(h)
        _check(status, "fm_workbook_recalc")

    def set_iterative(self, enabled: bool, max_iterations: int, max_change: float) -> None:
        """Configure iterative calculation."""
        h = self._require()
        status = LIB.fm_workbook_set_iterative(h, 1 if enabled else 0, int(max_iterations), float(max_change))
        _check(status, "fm_workbook_set_iterative")

    # -- Save --------------------------------------------------------------
    def save(self) -> bytes:
        """Serialise the workbook to an in-memory ``.xlsx`` byte stream.

        The returned bytes are an independent copy; the underlying WASM
        buffer is freed before this method returns.
        """
        h = self._require()
        out_ptr_ptr = LIB.alloc(4)
        out_len_ptr = LIB.alloc(4)
        LIB.write_bytes(out_ptr_ptr, b"\x00\x00\x00\x00")
        LIB.write_bytes(out_len_ptr, b"\x00\x00\x00\x00")
        try:
            status = LIB.fm_workbook_save(h, out_ptr_ptr, out_len_ptr)
            _check(status, "fm_workbook_save")
            data_ptr = LIB.read_u32(out_ptr_ptr)
            data_len = LIB.read_u32(out_len_ptr)
            try:
                if data_len == 0 or data_ptr == 0:
                    return b""
                return LIB.read_bytes(data_ptr, data_len)
            finally:
                if data_ptr:
                    LIB.fm_buffer_free(data_ptr)
        finally:
            LIB.free(out_ptr_ptr)
            LIB.free(out_len_ptr)

    def save_ex(self, fmt: "WorkbookFormat | int") -> bytes:
        """Serialise the workbook to an in-memory byte stream in ``fmt``.

        ``WorkbookFormat.XLSX`` produces the same bytes as :meth:`save`;
        ``WorkbookFormat.XLSB`` produces an MS-XLSB package via the
        ``.xlsb`` writer. The returned bytes are an independent copy;
        the underlying WASM buffer is freed before this method returns.
        """
        h = self._require()
        out_ptr_ptr = LIB.alloc(4)
        out_len_ptr = LIB.alloc(4)
        LIB.write_bytes(out_ptr_ptr, b"\x00\x00\x00\x00")
        LIB.write_bytes(out_len_ptr, b"\x00\x00\x00\x00")
        try:
            status = LIB.fm_workbook_save_ex(h, int(fmt), out_ptr_ptr, out_len_ptr)
            _check(status, "fm_workbook_save_ex")
            data_ptr = LIB.read_u32(out_ptr_ptr)
            data_len = LIB.read_u32(out_len_ptr)
            try:
                if data_len == 0 or data_ptr == 0:
                    return b""
                return LIB.read_bytes(data_ptr, data_len)
            finally:
                if data_ptr:
                    LIB.fm_buffer_free(data_ptr)
        finally:
            LIB.free(out_ptr_ptr)
            LIB.free(out_len_ptr)

    def save_ex_with_diagnostics(self, fmt: "WorkbookFormat | int") -> SaveDiagnostics:
        """Serialise the workbook and return XLSB loss/defer counters.

        XLSX saves return zero counters. The returned bytes are an
        independent copy and the underlying WASM allocation is released
        before this method returns.
        """
        h = self._require()
        scratch: list[int] = []
        try:
            for _ in range(4):
                ptr = LIB.alloc(4)
                scratch.append(ptr)
                LIB.write_bytes(ptr, b"\x00\x00\x00\x00")
            out_ptr_ptr, out_len_ptr, downgraded_ptr, deferred_ptr = scratch
            status = LIB.fm_workbook_save_ex_with_diagnostics(
                h, int(fmt), out_ptr_ptr, out_len_ptr, downgraded_ptr, deferred_ptr
            )
            _check(status, "fm_workbook_save_ex_with_diagnostics")
            data_ptr = LIB.read_u32(out_ptr_ptr)
            data_len = LIB.read_u32(out_len_ptr)
            try:
                data = b"" if data_len == 0 or data_ptr == 0 else LIB.read_bytes(data_ptr, data_len)
            finally:
                if data_ptr:
                    LIB.fm_buffer_free(data_ptr)
            return SaveDiagnostics(
                bytes=data,
                downgraded_formula_count=LIB.read_u32(downgraded_ptr),
                deferred_feature_count=LIB.read_u32(deferred_ptr),
            )
        finally:
            for ptr in scratch:
                LIB.free(ptr)

    def xlsb_read_diagnostics(self) -> XlsbReadDiagnostics:
        """Return recovery and dropped-part counters captured on load."""
        h = self._require()
        scratch: list[int] = []
        try:
            for _ in range(3):
                ptr = LIB.alloc(4)
                scratch.append(ptr)
                LIB.write_bytes(ptr, b"\x00\x00\x00\x00")
            formula_ptr, name_ptr, dropped_ptr = scratch
            status = LIB.fm_workbook_xlsb_read_diagnostics_ex(h, formula_ptr, name_ptr, dropped_ptr)
            _check(status, "fm_workbook_xlsb_read_diagnostics_ex")
            return XlsbReadDiagnostics(
                undecoded_formula_count=LIB.read_u32(formula_ptr),
                undecoded_defined_name_count=LIB.read_u32(name_ptr),
                dropped_part_count=LIB.read_u32(dropped_ptr),
            )
        finally:
            for ptr in scratch:
                LIB.free(ptr)

    # -- Iteration ---------------------------------------------------------
    def iter_cells(self, sheet: int) -> Iterator[Cell]:
        """Iterate over every stored cell on ``sheet``.

        Iteration order is implementation-defined but stable for an
        unmutated workbook (sorted by ``(row, col)`` ascending per the C
        ABI). Borrowed text pointers are decoded eagerly per yielded
        item. Do not mutate the workbook while consuming this iterator: the
        C ABI invalidates a cell enumeration after a workbook mutation.
        """
        h = self._require()
        out_count = _alloc_out_ptr()
        try:
            status = LIB.fm_workbook_cell_count(h, sheet, out_count)
            _check(status, "fm_workbook_cell_count")
            n = LIB.read_u32(out_count)
        finally:
            LIB.free(out_count)
        # Scratch: row(4) + col(4) + formula_ptr(4) + value(16) = 28 bytes.
        # Allocate them as separate slots so the ABI's
        # individual out-parameter contract is respected.
        row_ptr = LIB.alloc(4)
        col_ptr = LIB.alloc(4)
        formula_ptr = LIB.alloc(4)
        value_ptr = LIB.alloc(fm_value_t_size)
        try:
            LIB.write_bytes(row_ptr, b"\x00\x00\x00\x00")
            LIB.write_bytes(col_ptr, b"\x00\x00\x00\x00")
            LIB.write_bytes(formula_ptr, b"\x00\x00\x00\x00")
            for i in range(n):
                status = LIB.fm_workbook_cell_at(h, sheet, i, row_ptr, col_ptr, formula_ptr, value_ptr)
                _check(status, "fm_workbook_cell_at")
                row = LIB.read_u32(row_ptr)
                col = LIB.read_u32(col_ptr)
                formula_addr = LIB.read_u32(formula_ptr)
                formula = LIB.read_cstr(formula_addr) if formula_addr else None
                value = Value._from_wasm(value_ptr)
                yield Cell(row=row, col=col, formula=formula, value=value)
        finally:
            LIB.free(row_ptr)
            LIB.free(col_ptr)
            LIB.free(formula_ptr)
            LIB.free(value_ptr)

    def iter_defined_names(self) -> Iterator[DefinedName]:
        """Iterate over every defined name in declaration order."""
        h = self._require()
        n = int(LIB.fm_workbook_defined_name_count(h))
        for i in range(n):
            name_ptr = LIB.alloc(4)
            formula_ptr = LIB.alloc(4)
            local_sheet_ptr = LIB.alloc(4)
            LIB.write_bytes(name_ptr, b"\x00\x00\x00\x00")
            LIB.write_bytes(formula_ptr, b"\x00\x00\x00\x00")
            LIB.write_bytes(local_sheet_ptr, b"\x00\x00\x00\x00")
            try:
                status = LIB.fm_workbook_defined_name_at_ex(h, i, name_ptr, formula_ptr, local_sheet_ptr)
                _check(status, "fm_workbook_defined_name_at_ex")
                name = LIB.read_cstr(LIB.read_u32(name_ptr))
                formula = LIB.read_cstr(LIB.read_u32(formula_ptr))
                local_sheet_id = LIB.read_i32(local_sheet_ptr)
            finally:
                LIB.free(name_ptr)
                LIB.free(formula_ptr)
                LIB.free(local_sheet_ptr)
            yield DefinedName(name=name, formula=formula, local_sheet_id=local_sheet_id)

    def iter_tables(self) -> Iterator[Table]:
        """Iterate over every table in declaration order."""
        h = self._require()
        n = int(LIB.fm_workbook_table_count(h))
        for i in range(n):
            name_ptr = LIB.alloc(4)
            display_ptr = LIB.alloc(4)
            ref_ptr = LIB.alloc(4)
            sheet_ptr = LIB.alloc(4)
            for p in (name_ptr, display_ptr, ref_ptr, sheet_ptr):
                LIB.write_bytes(p, b"\x00\x00\x00\x00")
            try:
                status = LIB.fm_workbook_table_at(h, i, name_ptr, display_ptr, ref_ptr, sheet_ptr)
                _check(status, "fm_workbook_table_at")
                name = LIB.read_cstr(LIB.read_u32(name_ptr))
                display = LIB.read_cstr(LIB.read_u32(display_ptr))
                ref = LIB.read_cstr(LIB.read_u32(ref_ptr))
                sheet = LIB.read_u32(sheet_ptr)
            finally:
                LIB.free(name_ptr)
                LIB.free(display_ptr)
                LIB.free(ref_ptr)
                LIB.free(sheet_ptr)
            yield Table(name=name, display_name=display, ref=ref, sheet_index=sheet)

    def iter_passthrough(self) -> Iterator[PassthroughPart]:
        """Iterate over passthrough OOXML part paths."""
        h = self._require()
        n = int(LIB.fm_workbook_passthrough_count(h))
        for i in range(n):
            path_ptr = LIB.alloc(4)
            LIB.write_bytes(path_ptr, b"\x00\x00\x00\x00")
            try:
                status = LIB.fm_workbook_passthrough_at(h, i, path_ptr)
                _check(status, "fm_workbook_passthrough_at")
                path = LIB.read_cstr(LIB.read_u32(path_ptr))
            finally:
                LIB.free(path_ptr)
            yield PassthroughPart(path=path)

    # -- Sheet structure ---------------------------------------------------
    def move_sheet(self, from_index: int, to_index: int) -> None:
        """Move the sheet at ``from_index`` to ``to_index`` (post-removal)."""
        h = self._require()
        _check(
            LIB.fm_workbook_move_sheet(h, int(from_index), int(to_index)),
            "fm_workbook_move_sheet",
        )

    def remove_sheet(self, index: int) -> None:
        """Remove the sheet at ``index``."""
        h = self._require()
        _check(LIB.fm_workbook_remove_sheet(h, int(index)), "fm_workbook_remove_sheet")

    def rename_sheet(self, index: int, new_name: str) -> None:
        """Rename the sheet at ``index`` to ``new_name``."""
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(new_name)
        try:
            _check(
                LIB.fm_workbook_rename_sheet(h, int(index), name_ptr),
                "fm_workbook_rename_sheet",
            )
        finally:
            LIB.free(name_ptr)

    # -- Defined names -----------------------------------------------------
    def set_defined_name(self, name: str, formula: str) -> None:
        """Set, append, or remove a workbook-scoped defined name.

        An empty ``formula`` removes the entry.
        """
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(name)
        formula_ptr, _ = LIB.alloc_utf8(formula)
        try:
            _check(
                LIB.fm_workbook_set_defined_name(h, name_ptr, formula_ptr),
                "fm_workbook_set_defined_name",
            )
        finally:
            LIB.free(name_ptr)
            LIB.free(formula_ptr)

    def set_defined_name_scoped(self, name: str, formula: str, local_sheet_id: int) -> None:
        """Set, append, or remove a defined name in workbook or sheet scope.

        Use ``local_sheet_id=-1`` for workbook scope, or a 0-based sheet
        index for sheet-local scope. An empty ``formula`` removes the
        matching entry in that scope.
        """
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(name)
        formula_ptr, _ = LIB.alloc_utf8(formula)
        try:
            _check(
                LIB.fm_workbook_set_defined_name_scoped(h, name_ptr, formula_ptr, int(local_sheet_id)),
                "fm_workbook_set_defined_name_scoped",
            )
        finally:
            LIB.free(name_ptr)
            LIB.free(formula_ptr)

    # -- Row / column structural edits -------------------------------------
    def insert_rows(self, sheet: int, row: int, count: int) -> None:
        """Insert ``count`` rows at ``row`` on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_workbook_insert_rows(h, int(sheet), int(row), int(count)),
            "fm_workbook_insert_rows",
        )

    def delete_rows(self, sheet: int, row: int, count: int) -> None:
        """Delete ``count`` rows starting at ``row`` on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_workbook_delete_rows(h, int(sheet), int(row), int(count)),
            "fm_workbook_delete_rows",
        )

    def insert_cols(self, sheet: int, col: int, count: int) -> None:
        """Insert ``count`` columns at ``col`` on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_workbook_insert_cols(h, int(sheet), int(col), int(count)),
            "fm_workbook_insert_cols",
        )

    def delete_cols(self, sheet: int, col: int, count: int) -> None:
        """Delete ``count`` columns starting at ``col`` on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_workbook_delete_cols(h, int(sheet), int(col), int(count)),
            "fm_workbook_delete_cols",
        )

    # -- Calc policy / behaviour profile -----------------------------------
    def calc_mode(self) -> CalcMode:
        """Return the workbook's calc mode (round-trip metadata)."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(LIB.fm_workbook_calc_mode(h, out), "fm_workbook_calc_mode")
            return CalcMode(LIB.read_i32(out))
        finally:
            LIB.free(out)

    def set_calc_mode(self, mode: int) -> None:
        """Set the workbook's calc mode."""
        h = self._require()
        _check(LIB.fm_workbook_set_calc_mode(h, int(mode)), "fm_workbook_set_calc_mode")

    def excel_profile_id(self) -> str:
        """Return the workbook's active Excel formula profile id."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_excel_profile_id(h, out),
                "fm_workbook_excel_profile_id",
            )
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(out)

    def set_excel_profile_id(self, profile_id: str) -> None:
        """Set the workbook's Excel formula profile by id."""
        h = self._require()
        pid_ptr, _ = LIB.alloc_utf8(profile_id)
        try:
            _check(
                LIB.fm_workbook_set_excel_profile_id(h, pid_ptr),
                "fm_workbook_set_excel_profile_id",
            )
        finally:
            LIB.free(pid_ptr)

    # -- Partial recalc ----------------------------------------------------
    #
    # Note: ``fm_workbook_set_iterative_progress`` (the per-sweep callback)
    # is intentionally NOT bound -- it takes a C function pointer that the
    # wasmtime-py host cannot synthesise into the module's function table.
    def partial_recalc(
        self,
        sheet: int,
        first_row: int,
        last_row: int,
        first_col: int,
        last_col: int,
    ) -> int:
        """Recalc the dependency closure for a viewport; return cell count."""
        h = self._require()
        vp = S.alloc_struct(LIB, S.VIEWPORT)
        out = _alloc_out_ptr()
        try:
            S.VIEWPORT.pack(
                LIB,
                vp,
                {
                    "sheet": int(sheet),
                    "first_row": int(first_row),
                    "last_row": int(last_row),
                    "first_col": int(first_col),
                    "last_col": int(last_col),
                },
            )
            _check(
                LIB.fm_workbook_partial_recalc(h, vp, out),
                "fm_workbook_partial_recalc",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(vp)
            LIB.free(out)

    # -- Lambda text -------------------------------------------------------
    def lambda_text_at(self, sheet: int, row: int, col: int) -> str:
        """Render the lambda cached at ``(sheet, row, col)`` as text."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_lambda_text_at(h, int(sheet), int(row), int(col), out),
                "fm_workbook_lambda_text_at",
            )
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(out)

    # -- Merges ------------------------------------------------------------
    def add_merge(self, sheet: int, merge: MergeRange) -> None:
        """Append a merge range to ``sheet``."""
        h = self._require()
        owned: List[int] = []
        ptr = _pack_merge_array([merge], owned)
        try:
            _check(LIB.fm_sheet_add_merge(h, int(sheet), ptr), "fm_sheet_add_merge")
        finally:
            for p in owned:
                LIB.free(p)

    def remove_merge(self, sheet: int, merge: MergeRange) -> None:
        """Remove every merge on ``sheet`` overlapping ``merge``."""
        h = self._require()
        owned: List[int] = []
        ptr = _pack_merge_array([merge], owned)
        try:
            _check(
                LIB.fm_sheet_remove_merge(h, int(sheet), ptr),
                "fm_sheet_remove_merge",
            )
        finally:
            for p in owned:
                LIB.free(p)

    def remove_merge_at(self, sheet: int, index: int) -> None:
        """Remove the merge at ``index`` on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_sheet_remove_merge_at(h, int(sheet), int(index)),
            "fm_sheet_remove_merge_at",
        )

    def clear_merges(self, sheet: int) -> None:
        """Drop every merge range on ``sheet``."""
        h = self._require()
        _check(LIB.fm_sheet_clear_merges(h, int(sheet)), "fm_sheet_clear_merges")

    def merge_count(self, sheet: int) -> int:
        """Return the number of merge ranges on ``sheet``."""
        h = self._require()
        return _read_count(LIB.fm_sheet_get_merge_count, h, int(sheet))

    def get_merges(self, sheet: int) -> List[MergeRange]:
        """Return every merge range on ``sheet``."""
        h = self._require()
        n = self.merge_count(sheet)
        out: List[MergeRange] = []
        for i in range(n):
            ptr = S.alloc_struct(LIB, S.MERGE_RANGE)
            try:
                _check(
                    LIB.fm_sheet_get_merge_at(h, int(sheet), i, ptr),
                    "fm_sheet_get_merge_at",
                )
                d = S.MERGE_RANGE.unpack(LIB, ptr)
                out.append(
                    MergeRange(
                        first_row=d["first_row"],
                        first_col=d["first_col"],
                        last_row=d["last_row"],
                        last_col=d["last_col"],
                    )
                )
            finally:
                LIB.free(ptr)
        return out

    # -- Hyperlinks --------------------------------------------------------
    def add_hyperlink(
        self,
        sheet: int,
        row: int,
        col: int,
        target: str,
        display: str = "",
        tooltip: str = "",
        location: str = "",
    ) -> None:
        """Append a hyperlink to ``sheet``."""
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.HYPERLINK)
        try:
            S.HYPERLINK.pack(LIB, ptr, {"row": int(row), "col": int(col)})
            S.write_str_field(LIB, ptr, S.HYPERLINK, "target", target, owned)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "location", location, owned)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "display", display, owned)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "tooltip", tooltip, owned)
            _check(LIB.fm_sheet_add_hyperlink(h, int(sheet), ptr), "fm_sheet_add_hyperlink")
        finally:
            LIB.free(ptr)
            for p in owned:
                LIB.free(p)

    def remove_hyperlink(self, sheet: int, row: int, col: int) -> None:
        """Remove every hyperlink anchored at ``(row, col)`` on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_sheet_remove_hyperlink(h, int(sheet), int(row), int(col)),
            "fm_sheet_remove_hyperlink",
        )

    def remove_hyperlink_at(self, sheet: int, index: int) -> None:
        """Remove the hyperlink at ``index`` on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_sheet_remove_hyperlink_at(h, int(sheet), int(index)),
            "fm_sheet_remove_hyperlink_at",
        )

    def clear_hyperlinks(self, sheet: int) -> None:
        """Drop every hyperlink on ``sheet``."""
        h = self._require()
        _check(LIB.fm_sheet_clear_hyperlinks(h, int(sheet)), "fm_sheet_clear_hyperlinks")

    def hyperlink_count(self, sheet: int) -> int:
        """Return the number of hyperlinks on ``sheet``."""
        h = self._require()
        return _read_count(LIB.fm_sheet_get_hyperlink_count, h, int(sheet))

    def get_hyperlinks(self, sheet: int) -> List[Hyperlink]:
        """Return every hyperlink on ``sheet``."""
        h = self._require()
        n = self.hyperlink_count(sheet)
        out: List[Hyperlink] = []
        for i in range(n):
            ptr = S.alloc_struct(LIB, S.HYPERLINK)
            try:
                _check(
                    LIB.fm_sheet_get_hyperlink_at(h, int(sheet), i, ptr),
                    "fm_sheet_get_hyperlink_at",
                )
                d = S.HYPERLINK.unpack(LIB, ptr)
                out.append(
                    Hyperlink(
                        row=d["row"],
                        col=d["col"],
                        target=LIB.read_cstr(d["target"]),
                        location=LIB.read_cstr(d["location"]),
                        display=LIB.read_cstr(d["display"]),
                        tooltip=LIB.read_cstr(d["tooltip"]),
                    )
                )
            finally:
                LIB.free(ptr)
        return out

    # -- Comments ----------------------------------------------------------
    def get_comment(self, sheet: int, row: int, col: int) -> Optional[Comment]:
        """Return the comment at ``(sheet, row, col)`` or ``None`` if absent."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.COMMENT)
        try:
            status = LIB.fm_sheet_get_comment_at(h, int(sheet), int(row), int(col), ptr)
            if status == _STATUS_NOT_FOUND:
                return None
            _check(status, "fm_sheet_get_comment_at")
            d = S.COMMENT.unpack(LIB, ptr)
            return Comment(
                author=LIB.read_cstr(d["author"]),
                text=LIB.read_cstr(d["text"]),
            )
        finally:
            LIB.free(ptr)

    def set_comment(self, sheet: int, row: int, col: int, author: str, text: str) -> None:
        """Set, replace, or (empty ``text``) remove a cell comment."""
        h = self._require()
        owned: List[int] = []
        author_ptr = _opt_str_ptr(author, owned)
        text_ptr = _opt_str_ptr(text, owned)
        try:
            _check(
                LIB.fm_sheet_set_comment(h, int(sheet), int(row), int(col), author_ptr, text_ptr),
                "fm_sheet_set_comment",
            )
        finally:
            for p in owned:
                LIB.free(p)

    def comment_count(self, sheet: int) -> int:
        """Return the number of comments on ``sheet``."""
        h = self._require()
        return _read_count(LIB.fm_sheet_get_comment_count, h, int(sheet))

    def get_comments(self, sheet: int) -> List[CommentEntry]:
        """Return every comment on ``sheet`` in storage order.

        Unlike :meth:`get_comment`, this discovers comments on otherwise
        empty cells without requiring callers to scan worksheet coordinates.
        """
        h = self._require()
        out: List[CommentEntry] = []
        for index in range(self.comment_count(sheet)):
            ptr = S.alloc_struct(LIB, S.COMMENT)
            try:
                _check(
                    LIB.fm_sheet_get_comment_at_index(h, int(sheet), index, ptr),
                    "fm_sheet_get_comment_at_index",
                )
                d = S.COMMENT.unpack(LIB, ptr)
                out.append(CommentEntry(d["row"], d["col"], LIB.read_cstr(d["author"]), LIB.read_cstr(d["text"])))
            finally:
                LIB.free(ptr)
        return out

    # -- Data validations --------------------------------------------------
    def validation_count(self, sheet: int) -> int:
        """Return the number of data-validation rules on ``sheet``."""
        h = self._require()
        return _read_count(LIB.fm_sheet_get_validation_count, h, int(sheet))

    def get_validation_at(self, sheet: int, index: int) -> DataValidation:
        """Read the ``index``-th data-validation rule on ``sheet``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.DATA_VALIDATION)
        try:
            _check(
                LIB.fm_sheet_get_validation_at(h, int(sheet), int(index), ptr),
                "fm_sheet_get_validation_at",
            )
            return self._decode_validation(ptr)
        finally:
            LIB.free(ptr)

    def get_validations(self, sheet: int) -> List[DataValidation]:
        """Return every data-validation rule on ``sheet``."""
        n = self.validation_count(sheet)
        return [self.get_validation_at(sheet, i) for i in range(n)]

    @staticmethod
    def _decode_validation(ptr: int) -> DataValidation:
        d = S.DATA_VALIDATION.unpack(LIB, ptr)
        ranges: List[MergeRange] = []
        base = d["ranges"]
        rsize = S.MERGE_RANGE.size
        for i in range(d["range_count"]):
            rd = S.MERGE_RANGE.unpack(LIB, base + i * rsize)
            ranges.append(
                MergeRange(
                    first_row=rd["first_row"],
                    first_col=rd["first_col"],
                    last_row=rd["last_row"],
                    last_col=rd["last_col"],
                )
            )
        return DataValidation(
            ranges=ranges,
            type=d["type"],
            op=d["op"],
            error_style=d["error_style"],
            allow_blank=bool(d["allow_blank"]),
            show_input_message=bool(d["show_input_message"]),
            show_error_message=bool(d["show_error_message"]),
            show_dropdown=bool(d["show_dropdown"]),
            formula1=LIB.read_cstr(d["formula1"]),
            formula2=LIB.read_cstr(d["formula2"]),
            error_title=LIB.read_cstr(d["error_title"]),
            error_message=LIB.read_cstr(d["error_message"]),
            prompt_title=LIB.read_cstr(d["prompt_title"]),
            prompt_message=LIB.read_cstr(d["prompt_message"]),
        )

    def add_validation(self, sheet: int, validation: DataValidationInput) -> None:
        """Append a data-validation rule to ``sheet``."""
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.DATA_VALIDATION)
        try:
            S.DATA_VALIDATION.pack(
                LIB,
                ptr,
                {
                    "type": int(validation.type),
                    "op": int(validation.op),
                    "error_style": int(validation.error_style),
                    "allow_blank": 1 if validation.allow_blank else 0,
                    "show_input_message": 1 if validation.show_input_message else 0,
                    "show_error_message": 1 if validation.show_error_message else 0,
                    "show_dropdown": 1 if validation.show_dropdown else 0,
                },
            )
            ranges_ptr = _pack_merge_array(validation.ranges, owned)
            ro = S.DATA_VALIDATION.offsets
            LIB.write_bytes(ptr + ro["ranges"][1], struct.pack("<I", ranges_ptr))
            LIB.write_bytes(
                ptr + ro["range_count"][1],
                struct.pack("<I", len(validation.ranges)),
            )
            for fld in (
                "formula1",
                "formula2",
                "error_title",
                "error_message",
                "prompt_title",
                "prompt_message",
            ):
                S.write_str_field(LIB, ptr, S.DATA_VALIDATION, fld, getattr(validation, fld), owned)
            _check(
                LIB.fm_sheet_add_validation(h, int(sheet), ptr),
                "fm_sheet_add_validation",
            )
        finally:
            LIB.free(ptr)
            for p in owned:
                LIB.free(p)

    def remove_validation_at(self, sheet: int, index: int) -> None:
        """Remove the validation rule at ``index`` on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_sheet_remove_validation_at(h, int(sheet), int(index)),
            "fm_sheet_remove_validation_at",
        )

    def clear_validations(self, sheet: int) -> None:
        """Drop every validation rule on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_sheet_clear_validations(h, int(sheet)),
            "fm_sheet_clear_validations",
        )

    # -- Sheet protection --------------------------------------------------
    _PROTECT_FLAGS = (
        "sheet",
        "objects",
        "scenarios",
        "format_cells",
        "format_columns",
        "format_rows",
        "insert_columns",
        "insert_rows",
        "insert_hyperlinks",
        "delete_columns",
        "delete_rows",
        "select_locked_cells",
        "select_unlocked_cells",
        "sort",
        "auto_filter",
        "pivot_tables",
    )

    def get_sheet_protection(self, sheet: int) -> SheetProtection:
        """Read the sheet's ``<sheetProtection>`` flags."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.SHEET_PROTECTION)
        try:
            _check(
                LIB.fm_sheet_get_protection(h, int(sheet), ptr),
                "fm_sheet_get_protection",
            )
            d = S.SHEET_PROTECTION.unpack(LIB, ptr)
            prot = SheetProtection(
                enabled=bool(d["enabled"]),
                algorithm_name=LIB.read_cstr(d["algorithm_name"]),
                hash_value=LIB.read_cstr(d["hash_value"]),
                salt_value=LIB.read_cstr(d["salt_value"]),
                spin_count=d["spin_count"],
                legacy_password=LIB.read_cstr(d["legacy_password"]),
            )
            for flag in self._PROTECT_FLAGS:
                setattr(prot, flag, bool(d[flag]))
            return prot
        finally:
            LIB.free(ptr)

    def set_sheet_protection(self, sheet: int, protection: SheetProtection) -> None:
        """Replace the sheet's ``<sheetProtection>`` flags wholesale."""
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.SHEET_PROTECTION)
        try:
            values: Dict[str, int] = {
                "enabled": 1 if protection.enabled else 0,
                "spin_count": int(protection.spin_count),
            }
            for flag in self._PROTECT_FLAGS:
                values[flag] = 1 if getattr(protection, flag) else 0
            S.SHEET_PROTECTION.pack(LIB, ptr, values)
            for fld in ("algorithm_name", "hash_value", "salt_value", "legacy_password"):
                S.write_str_field(LIB, ptr, S.SHEET_PROTECTION, fld, getattr(protection, fld), owned)
            _check(
                LIB.fm_sheet_set_protection(h, int(sheet), ptr),
                "fm_sheet_set_protection",
            )
        finally:
            LIB.free(ptr)
            for p in owned:
                LIB.free(p)

    # -- Sheet view / layout -----------------------------------------------
    def paginate(self, sheet: int) -> PaginationResult:
        """Resolve the worksheet's print area, page breaks, and page count.

        The result is a snapshot: later workbook or layout changes do not
        mutate the returned lists.
        """
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(LIB.fm_workbook_paginate(h, int(sheet), out), "fm_workbook_paginate")
            pagination = LIB.read_u32(out)
            if pagination == 0:
                raise FormulonError("fm_workbook_paginate returned a null result")
            try:
                page_count = int(LIB.fm_pagination_page_count(pagination))
                range_count = int(LIB.fm_pagination_print_area_count(pagination))
                break_count = int(LIB.fm_pagination_horizontal_break_count(pagination))
                vertical_break_count = int(LIB.fm_pagination_vertical_break_count(pagination))
                range_ptr = LIB.alloc(16)
                value_ptr = _alloc_out_ptr()
                try:
                    print_area = []
                    for i in range(range_count):
                        _check(
                            LIB.fm_pagination_print_area_at(pagination, i, range_ptr),
                            "fm_pagination_print_area_at",
                        )
                        print_area.append(struct.unpack("<IIII", LIB.read_bytes(range_ptr, 16)))
                    horizontal_breaks = []
                    for i in range(break_count):
                        _check(
                            LIB.fm_pagination_horizontal_break_at(pagination, i, value_ptr),
                            "fm_pagination_horizontal_break_at",
                        )
                        horizontal_breaks.append(LIB.read_u32(value_ptr))
                    vertical_breaks = []
                    for i in range(vertical_break_count):
                        _check(
                            LIB.fm_pagination_vertical_break_at(pagination, i, value_ptr),
                            "fm_pagination_vertical_break_at",
                        )
                        vertical_breaks.append(LIB.read_u32(value_ptr))
                finally:
                    LIB.free(range_ptr)
                    LIB.free(value_ptr)
                return PaginationResult(page_count, print_area, horizontal_breaks, vertical_breaks)
            finally:
                LIB.fm_pagination_destroy(pagination)
        finally:
            LIB.free(out)

    def get_sheet_view(self, sheet: int) -> SheetView:
        """Read the full per-sheet view (zoom, freeze, tab-hidden, and the
        display / orientation flags)."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.SHEET_VIEW_EX)
        try:
            _check(LIB.fm_sheet_get_view_ex(h, int(sheet), ptr), "fm_sheet_get_view_ex")
            d = S.SHEET_VIEW_EX.unpack(LIB, ptr)
            return SheetView(
                zoom_scale=d["zoom_scale"],
                freeze_rows=d["freeze_rows"],
                freeze_cols=d["freeze_cols"],
                tab_hidden=bool(d["tab_hidden"]),
                show_grid_lines=bool(d["show_grid_lines"]),
                show_row_col_headers=bool(d["show_row_col_headers"]),
                show_zeros=bool(d["show_zeros"]),
                right_to_left=bool(d["right_to_left"]),
                tab_selected=bool(d["tab_selected"]),
                view_mode=LIB.read_cstr(d["view_mode"]) or "",
            )
        finally:
            LIB.free(ptr)

    def set_sheet_zoom(self, sheet: int, zoom_scale: int) -> None:
        """Set the sheet zoom percentage (clamped to ``[10, 400]``)."""
        h = self._require()
        _check(LIB.fm_sheet_set_zoom(h, int(sheet), int(zoom_scale)), "fm_sheet_set_zoom")

    def set_sheet_freeze(self, sheet: int, freeze_rows: int, freeze_cols: int) -> None:
        """Set the frozen pane in ``(rows, cols)``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_freeze(h, int(sheet), int(freeze_rows), int(freeze_cols)),
            "fm_sheet_set_freeze",
        )

    def set_sheet_tab_hidden(self, sheet: int, hidden: bool) -> None:
        """Set the sheet tab's hidden flag."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_tab_hidden(h, int(sheet), 1 if hidden else 0),
            "fm_sheet_set_tab_hidden",
        )

    def set_sheet_show_grid_lines(self, sheet: int, show: bool) -> None:
        """Set the sheet's ``showGridLines`` flag."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_show_grid_lines(h, int(sheet), 1 if show else 0),
            "fm_sheet_set_show_grid_lines",
        )

    def set_sheet_show_row_col_headers(self, sheet: int, show: bool) -> None:
        """Set the sheet's ``showRowColHeaders`` flag."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_show_row_col_headers(h, int(sheet), 1 if show else 0),
            "fm_sheet_set_show_row_col_headers",
        )

    def set_sheet_show_zeros(self, sheet: int, show: bool) -> None:
        """Set the sheet's ``showZeros`` flag."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_show_zeros(h, int(sheet), 1 if show else 0),
            "fm_sheet_set_show_zeros",
        )

    def set_sheet_right_to_left(self, sheet: int, right_to_left: bool) -> None:
        """Set the sheet's ``rightToLeft`` flag."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_right_to_left(h, int(sheet), 1 if right_to_left else 0),
            "fm_sheet_set_right_to_left",
        )

    def set_sheet_tab_selected(self, sheet: int, selected: bool) -> None:
        """Set the sheet's ``tabSelected`` flag."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_tab_selected(h, int(sheet), 1 if selected else 0),
            "fm_sheet_set_tab_selected",
        )

    def set_sheet_view_mode(self, sheet: int, mode: str) -> None:
        """Set the sheet's ``<sheetView view="...">`` mode: ``""``,
        ``"pageBreakPreview"``, or ``"pageLayout"``. Stored verbatim.

        Unlike most string setters in this module, an empty ``mode`` is a
        meaningful value here (the OOXML-default "normal" view), not a
        request to pass ``NULL`` -- ``fm_sheet_set_view_mode`` rejects a
        `NULL` pointer outright, so this always allocates a real buffer.
        """
        h = self._require()
        mode_ptr, _ = LIB.alloc_utf8(mode)
        try:
            _check(
                LIB.fm_sheet_set_view_mode(h, int(sheet), mode_ptr),
                "fm_sheet_set_view_mode",
            )
        finally:
            LIB.free(mode_ptr)

    def get_sheet_columns(self, sheet: int) -> List[ColumnLayout]:
        """Return the column-layout overrides on ``sheet``."""
        h = self._require()
        n = _read_count(LIB.fm_sheet_get_column_count, h, int(sheet))
        out: List[ColumnLayout] = []
        for i in range(n):
            ptr = S.alloc_struct(LIB, S.COLUMN_LAYOUT)
            try:
                _check(
                    LIB.fm_sheet_get_column(h, int(sheet), i, ptr),
                    "fm_sheet_get_column",
                )
                d = S.COLUMN_LAYOUT.unpack(LIB, ptr)
                out.append(
                    ColumnLayout(
                        first=d["first"],
                        last=d["last"],
                        width=d["width"],
                        hidden=bool(d["hidden"]),
                        outline_level=d["outline_level"],
                    )
                )
            finally:
                LIB.free(ptr)
        return out

    def set_column_width(self, sheet: int, first: int, last: int, width: float) -> None:
        """Set the column width override on ``[first, last]``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_column_width(h, int(sheet), int(first), int(last), float(width)),
            "fm_sheet_set_column_width",
        )

    def set_column_hidden(self, sheet: int, first: int, last: int, hidden: bool) -> None:
        """Set the column hidden flag on ``[first, last]``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_column_hidden(h, int(sheet), int(first), int(last), 1 if hidden else 0),
            "fm_sheet_set_column_hidden",
        )

    def set_column_outline(self, sheet: int, first: int, last: int, level: int) -> None:
        """Set the column outline level on ``[first, last]``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_column_outline(h, int(sheet), int(first), int(last), int(level)),
            "fm_sheet_set_column_outline",
        )

    def get_sheet_row_overrides(self, sheet: int) -> List[RowLayout]:
        """Return the row-layout overrides on ``sheet``."""
        h = self._require()
        n = _read_count(LIB.fm_sheet_get_row_override_count, h, int(sheet))
        out: List[RowLayout] = []
        for i in range(n):
            ptr = S.alloc_struct(LIB, S.ROW_LAYOUT)
            try:
                _check(
                    LIB.fm_sheet_get_row_override(h, int(sheet), i, ptr),
                    "fm_sheet_get_row_override",
                )
                d = S.ROW_LAYOUT.unpack(LIB, ptr)
                out.append(
                    RowLayout(
                        row=d["row"],
                        height=d["height"],
                        hidden=bool(d["hidden"]),
                        outline_level=d["outline_level"],
                    )
                )
            finally:
                LIB.free(ptr)
        return out

    def set_row_height(self, sheet: int, row: int, height: float) -> None:
        """Set the row height override at ``row``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_row_height(h, int(sheet), int(row), float(height)),
            "fm_sheet_set_row_height",
        )

    def set_row_hidden(self, sheet: int, row: int, hidden: bool) -> None:
        """Set the row hidden flag at ``row``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_row_hidden(h, int(sheet), int(row), 1 if hidden else 0),
            "fm_sheet_set_row_hidden",
        )

    def set_row_outline(self, sheet: int, row: int, level: int) -> None:
        """Set the row outline level at ``row``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_row_outline(h, int(sheet), int(row), int(level)),
            "fm_sheet_set_row_outline",
        )

    # -- Conditional formatting --------------------------------------------
    @staticmethod
    def _decode_cf_color(d: Dict[str, int], prefix: str) -> tuple:
        return (
            d[prefix + "r"],
            d[prefix + "g"],
            d[prefix + "b"],
            d[prefix + "a"],
        )

    def evaluate_cf_range(
        self,
        sheet: int,
        first_row: int,
        first_col: int,
        last_row: int,
        last_col: int,
        today_serial: float = float("nan"),
    ) -> List[CfCellResult]:
        """Evaluate every CF block on ``sheet`` against an inclusive range."""
        h = self._require()
        out = _alloc_out_ptr()
        results = 0
        try:
            _check(
                LIB.fm_workbook_cf_evaluate_range(
                    h,
                    int(sheet),
                    int(first_row),
                    int(first_col),
                    int(last_row),
                    int(last_col),
                    float(today_serial),
                    out,
                ),
                "fm_workbook_cf_evaluate_range",
            )
            results = LIB.read_u32(out)
            cells: List[CfCellResult] = []
            n = LIB.fm_cf_results_cell_count(results)
            for ci in range(n):
                rp = _alloc_out_ptr()
                cp = _alloc_out_ptr()
                mcp = _alloc_out_ptr()
                try:
                    _check(
                        LIB.fm_cf_results_cell_at(results, ci, rp, cp, mcp),
                        "fm_cf_results_cell_at",
                    )
                    row = LIB.read_u32(rp)
                    col = LIB.read_u32(cp)
                    mcount = LIB.read_u32(mcp)
                finally:
                    LIB.free(rp)
                    LIB.free(cp)
                    LIB.free(mcp)
                matches: List[CfMatch] = []
                for mi in range(mcount):
                    mptr = S.alloc_struct(LIB, S.CF_MATCH)
                    try:
                        _check(
                            LIB.fm_cf_results_match_at(results, ci, mi, mptr),
                            "fm_cf_results_match_at",
                        )
                        d = S.CF_MATCH.unpack(LIB, mptr)
                        matches.append(
                            CfMatch(
                                kind=d["kind"],
                                priority=d["priority"],
                                dxf_id_engaged=bool(d["dxf_id_engaged"]),
                                dxf_id=d["dxf_id"],
                                color=self._decode_cf_color(d, "color_"),
                                bar_length_pct=d["bar_length_pct"],
                                bar_axis_position_pct=d["bar_axis_position_pct"],
                                bar_is_negative=bool(d["bar_is_negative"]),
                                bar_fill=self._decode_cf_color(d, "bar_fill_"),
                                bar_border_engaged=bool(d["bar_border_engaged"]),
                                bar_border=self._decode_cf_color(d, "bar_border_"),
                                bar_gradient=bool(d["bar_gradient"]),
                                icon_set_name=d["icon_set_name"],
                                icon_index=d["icon_index"],
                            )
                        )
                    finally:
                        LIB.free(mptr)
                cells.append(CfCellResult(row=row, col=col, matches=matches))
            return cells
        finally:
            if results:
                LIB.fm_cf_results_destroy(results)
            LIB.free(out)

    def cf_count(self, sheet: int) -> int:
        """Return the number of CF rules on ``sheet`` (flattened)."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(LIB.fm_sheet_cf_count(h, int(sheet), out), "fm_sheet_cf_count")
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    @staticmethod
    def _decode_cfvo(ptr: int) -> CfValueObject:
        d = S.CFVO.unpack(LIB, ptr)
        return CfValueObject(
            type=d["type"], value=LIB.read_cstr(d["value"]) if d["value"] else None, gte=bool(d["gte"])
        )

    @staticmethod
    def _decode_cf_color_at(ptr: int) -> CfColor:
        d = S.CF_COLOR.unpack(LIB, ptr)
        return CfColor(d["r"], d["g"], d["b"], d["a"])

    @staticmethod
    def _write_cfvo(ptr: int, value: CfValueObject, owned: List[int]) -> None:
        S.CFVO.pack(LIB, ptr, {"type": int(value.type), "gte": 1 if value.gte else 0})
        if value.value is not None:
            S.write_str_field(LIB, ptr, S.CFVO, "value", value.value, owned)

    @staticmethod
    def _write_cf_color(ptr: int, color: CfColor) -> None:
        S.CF_COLOR.pack(LIB, ptr, {"r": int(color.r), "g": int(color.g), "b": int(color.b), "a": int(color.a)})

    def get_conditional_format_at(self, sheet: int, index: int) -> ConditionalFormat:
        """Read the ``index``-th CF rule on ``sheet`` (flattened order)."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.CF_RULE)
        try:
            _check(
                LIB.fm_sheet_cf_get_at(h, int(sheet), int(index), ptr),
                "fm_sheet_cf_get_at",
            )
            d = S.CF_RULE.unpack(LIB, ptr)
            sqref: List[MergeRange] = []
            base = d["sqref"]
            rsize = S.CF_CELL_RANGE.size
            for i in range(d["sqref_count"]):
                rd = S.CF_CELL_RANGE.unpack(LIB, base + i * rsize)
                sqref.append(
                    MergeRange(
                        first_row=rd["first_row"],
                        first_col=rd["first_col"],
                        last_row=rd["last_row"],
                        last_col=rd["last_col"],
                    )
                )
            color_scale = None
            if d["color_scale_count"]:
                thresholds = [
                    self._decode_cfvo(d["color_scale_thresholds"] + i * S.CFVO.size)
                    for i in range(d["color_scale_count"])
                ]
                colors = [
                    self._decode_cf_color_at(d["color_scale_colors"] + i * S.CF_COLOR.size)
                    for i in range(d["color_scale_count"])
                ]
                color_scale = ColorScale(thresholds, colors)
            offsets = S.CF_RULE.offsets
            data_bar = None
            if d["data_bar_engaged"]:
                data_bar = DataBar(
                    self._decode_cfvo(ptr + offsets["data_bar_min"][1]),
                    self._decode_cfvo(ptr + offsets["data_bar_max"][1]),
                    self._decode_cf_color_at(ptr + offsets["data_bar_fill"][1]),
                    bool(d["data_bar_show_value"]),
                    d["data_bar_min_length_pct"],
                    d["data_bar_max_length_pct"],
                )
            icon_set = None
            if d["icon_set_engaged"]:
                icon_set = IconSet(
                    d["icon_set_name"],
                    [
                        self._decode_cfvo(d["icon_set_thresholds"] + i * S.CFVO.size)
                        for i in range(d["icon_set_threshold_count"])
                    ],
                    bool(d["icon_set_reverse"]),
                    bool(d["icon_set_show_value"]),
                    bool(d["icon_set_percent"]),
                )
            return ConditionalFormat(
                id=LIB.read_cstr(d["id"]),
                type=d["type"],
                priority=d["priority"],
                stop_if_true=bool(d["stop_if_true"]),
                sqref=sqref,
                dxf_id=d["dxf_id"] if d["dxf_id_engaged"] else None,
                formula1=LIB.read_cstr(d["formula1"]),
                formula2=LIB.read_cstr(d["formula2"]),
                op=d["op"],
                rank=d["rank"],
                percent=bool(d["percent"]),
                bottom=bool(d["bottom"]),
                above_average=bool(d["above_average"]),
                equal_average=bool(d["equal_average"]),
                std_dev=d["std_dev"],
                text=LIB.read_cstr(d["text"]),
                time_period=d["time_period"],
                color_scale=color_scale,
                data_bar=data_bar,
                icon_set=icon_set,
            )
        finally:
            LIB.free(ptr)

    def get_conditional_formats(self, sheet: int) -> List[ConditionalFormat]:
        """Return every CF rule on ``sheet`` in flattened priority order."""
        n = self.cf_count(sheet)
        return [self.get_conditional_format_at(sheet, i) for i in range(n)]

    def add_conditional_format(self, sheet: int, rule: ConditionalFormatInput) -> int:
        """Append a non-visual CF rule to ``sheet``; return its index."""
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.CF_RULE)
        out = _alloc_out_ptr()
        try:
            S.CF_RULE.pack(
                LIB,
                ptr,
                {
                    "type": int(rule.type),
                    "op": int(rule.op),
                    "time_period": int(rule.time_period),
                    "priority": int(rule.priority),
                    "stop_if_true": 1 if rule.stop_if_true else 0,
                    "dxf_id_engaged": 1 if rule.dxf_id_engaged else 0,
                    "dxf_id": int(rule.dxf_id),
                    "op_engaged": 1 if rule.op_engaged else 0,
                    "rank_engaged": 1 if rule.rank_engaged else 0,
                    "rank": int(rule.rank),
                    "percent": 1 if rule.percent else 0,
                    "bottom": 1 if rule.bottom else 0,
                    "above_average": 1 if rule.above_average else 0,
                    "equal_average": 1 if rule.equal_average else 0,
                    "std_dev_engaged": 1 if rule.std_dev_engaged else 0,
                    "std_dev": float(rule.std_dev),
                    "time_period_engaged": 1 if rule.time_period_engaged else 0,
                },
            )
            sqref_ptr = 0
            if rule.sqref:
                rsize = S.CF_CELL_RANGE.size
                sqref_ptr = LIB.alloc(rsize * len(rule.sqref))
                owned.append(sqref_ptr)
                for i, r in enumerate(rule.sqref):
                    S.CF_CELL_RANGE.pack(
                        LIB,
                        sqref_ptr + i * rsize,
                        {
                            "first_row": r.first_row,
                            "first_col": r.first_col,
                            "last_row": r.last_row,
                            "last_col": r.last_col,
                        },
                    )
            ro = S.CF_RULE.offsets
            LIB.write_bytes(ptr + ro["sqref"][1], struct.pack("<I", sqref_ptr))
            LIB.write_bytes(ptr + ro["sqref_count"][1], struct.pack("<I", len(rule.sqref)))
            S.write_str_field(LIB, ptr, S.CF_RULE, "id", rule.id, owned)
            S.write_str_field(LIB, ptr, S.CF_RULE, "formula1", rule.formula1, owned)
            S.write_str_field(LIB, ptr, S.CF_RULE, "formula2", rule.formula2, owned)
            S.write_str_field(LIB, ptr, S.CF_RULE, "text", rule.text, owned)
            if rule.color_scale is not None:
                count = len(rule.color_scale.thresholds)
                if count != len(rule.color_scale.colors):
                    raise ValueError("color_scale thresholds and colors must have equal length")
                thresholds = LIB.alloc(count * S.CFVO.size)
                colors = LIB.alloc(count * S.CF_COLOR.size)
                owned.extend((thresholds, colors))
                for i, value in enumerate(rule.color_scale.thresholds):
                    self._write_cfvo(thresholds + i * S.CFVO.size, value, owned)
                for i, color in enumerate(rule.color_scale.colors):
                    self._write_cf_color(colors + i * S.CF_COLOR.size, color)
                LIB.write_bytes(ptr + ro["color_scale_thresholds"][1], struct.pack("<I", thresholds))
                LIB.write_bytes(ptr + ro["color_scale_colors"][1], struct.pack("<I", colors))
                LIB.write_bytes(ptr + ro["color_scale_count"][1], struct.pack("<I", count))
            if rule.data_bar is not None:
                data_bar = rule.data_bar
                self._write_cfvo(ptr + ro["data_bar_min"][1], data_bar.minimum, owned)
                self._write_cfvo(ptr + ro["data_bar_max"][1], data_bar.maximum, owned)
                self._write_cf_color(ptr + ro["data_bar_fill"][1], data_bar.fill)
                LIB.write_bytes(ptr + ro["data_bar_engaged"][1], struct.pack("<i", 1))
                LIB.write_bytes(ptr + ro["data_bar_show_value"][1], struct.pack("<i", 1 if data_bar.show_value else 0))
                LIB.write_bytes(ptr + ro["data_bar_min_length_pct"][1], struct.pack("<B", int(data_bar.min_length_pct)))
                LIB.write_bytes(ptr + ro["data_bar_max_length_pct"][1], struct.pack("<B", int(data_bar.max_length_pct)))
            if rule.icon_set is not None:
                icon_set = rule.icon_set
                count = len(icon_set.thresholds)
                thresholds = LIB.alloc(count * S.CFVO.size)
                owned.append(thresholds)
                for i, value in enumerate(icon_set.thresholds):
                    self._write_cfvo(thresholds + i * S.CFVO.size, value, owned)
                LIB.write_bytes(ptr + ro["icon_set_thresholds"][1], struct.pack("<I", thresholds))
                LIB.write_bytes(ptr + ro["icon_set_engaged"][1], struct.pack("<i", 1))
                LIB.write_bytes(ptr + ro["icon_set_name"][1], struct.pack("<B", int(icon_set.name)))
                LIB.write_bytes(ptr + ro["icon_set_threshold_count"][1], struct.pack("<I", count))
                LIB.write_bytes(ptr + ro["icon_set_reverse"][1], struct.pack("<i", 1 if icon_set.reverse else 0))
                LIB.write_bytes(ptr + ro["icon_set_show_value"][1], struct.pack("<i", 1 if icon_set.show_value else 0))
                LIB.write_bytes(ptr + ro["icon_set_percent"][1], struct.pack("<i", 1 if icon_set.percent else 0))
            _check(
                LIB.fm_sheet_cf_add_rule(h, int(sheet), ptr, out),
                "fm_sheet_cf_add_rule",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)
            for p in owned:
                LIB.free(p)

    def remove_conditional_format_at(self, sheet: int, index: int) -> None:
        """Remove the CF rule at ``index`` on ``sheet`` (flattened order)."""
        h = self._require()
        _check(
            LIB.fm_sheet_cf_remove_at(h, int(sheet), int(index)),
            "fm_sheet_cf_remove_at",
        )

    def clear_conditional_formats(self, sheet: int) -> None:
        """Drop every CF block on ``sheet``."""
        h = self._require()
        _check(LIB.fm_sheet_cf_clear(h, int(sheet)), "fm_sheet_cf_clear")

    # -- Styles ------------------------------------------------------------
    def get_cell_xf_index(self, sheet: int, row: int, col: int) -> int:
        """Return the xf (style record) index attached to a cell."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_cell_get_xf_index(h, int(sheet), int(row), int(col), out),
                "fm_cell_get_xf_index",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    def set_cell_xf_index(self, sheet: int, row: int, col: int, xf_index: int) -> None:
        """Store ``xf_index`` on the cell at ``(row, col)``."""
        h = self._require()
        _check(
            LIB.fm_cell_set_xf_index(h, int(sheet), int(row), int(col), int(xf_index)),
            "fm_cell_set_xf_index",
        )

    @staticmethod
    def _decode_cell_xf(ptr: int) -> CellXf:
        d = S.CELL_XF_EX2.unpack(LIB, ptr)
        return CellXf(
            font_index=d["font_index"],
            fill_index=d["fill_index"],
            border_index=d["border_index"],
            num_fmt_id=d["num_fmt_id"],
            horizontal_align=d["horizontal_align"],
            vertical_align=d["vertical_align"],
            wrap_text=bool(d["wrap_text"]),
            has_alignment=bool(d["has_alignment"]),
            justify_last_line=bool(d["justify_last_line"]),
            xf_id=d["xf_id"],
            text_rotation=d["text_rotation"] if d["has_text_rotation"] else None,
            indent=d["indent"] if d["has_indent"] else None,
            relative_indent=d["relative_indent"] if d["has_relative_indent"] else None,
            shrink_to_fit=bool(d["shrink_to_fit"]) if d["has_shrink_to_fit"] else None,
            reading_order=d["reading_order"] if d["has_reading_order"] else None,
            has_horizontal_align=bool(d["has_horizontal_align"]),
            has_vertical_align=bool(d["has_vertical_align"]),
            has_wrap_text=bool(d["has_wrap_text"]),
            has_justify_last_line=bool(d["has_justify_last_line"]),
        )

    def get_cell_xf(self, xf_index: int) -> CellXf:
        """Return the resolved ``<xf>`` record at ``xf_index``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.CELL_XF_EX2)
        try:
            _check(
                LIB.fm_styles_get_cell_xf_ex2(h, int(xf_index), ptr),
                "fm_styles_get_cell_xf_ex2",
            )
            return self._decode_cell_xf(ptr)
        finally:
            LIB.free(ptr)

    def get_font(self, font_index: int) -> FontRecord:
        """Return the resolved font record at ``font_index``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.FONT_RECORD)
        try:
            _check(LIB.fm_styles_get_font(h, int(font_index), ptr), "fm_styles_get_font")
            d = S.FONT_RECORD.unpack(LIB, ptr)
            return FontRecord(
                name=LIB.read_cstr(d["name"]),
                size=d["size"],
                color_argb=d["color_argb"],
                bold=bool(d["bold"]),
                italic=bool(d["italic"]),
                strike=bool(d["strike"]),
                underline=d["underline"],
            )
        finally:
            LIB.free(ptr)

    def get_fill(self, fill_index: int) -> FillRecord:
        """Return the resolved fill record at ``fill_index``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.FILL_RECORD)
        try:
            _check(LIB.fm_styles_get_fill(h, int(fill_index), ptr), "fm_styles_get_fill")
            d = S.FILL_RECORD.unpack(LIB, ptr)
            return FillRecord(pattern=d["pattern"], fg_argb=d["fg_argb"], bg_argb=d["bg_argb"])
        finally:
            LIB.free(ptr)

    def get_border(self, border_index: int) -> Dict[str, object]:
        """Return the resolved border record at ``border_index``.

        Each side is a ``{"style", "color_argb"}`` dict; the result also
        carries ``diagonal_up`` / ``diagonal_down`` booleans.
        """
        h = self._require()
        ptr = S.alloc_struct(LIB, S.BORDER_RECORD)
        try:
            _check(
                LIB.fm_styles_get_border(h, int(border_index), ptr),
                "fm_styles_get_border",
            )
            d = S.BORDER_RECORD.unpack(LIB, ptr)
            sides = {}
            for side in ("left", "right", "top", "bottom", "diagonal"):
                sides[side] = {
                    "style": d[side + "_style"],
                    "color_argb": d[side + "_color_argb"],
                }
            sides["diagonal_up"] = bool(d["diagonal_up"])
            sides["diagonal_down"] = bool(d["diagonal_down"])
            return sides
        finally:
            LIB.free(ptr)

    @staticmethod
    def _decode_border(ptr: int) -> Dict[str, object]:
        d = S.BORDER_RECORD.unpack(LIB, ptr)
        sides: Dict[str, object] = {}
        for side in ("left", "right", "top", "bottom", "diagonal"):
            sides[side] = {
                "style": d[side + "_style"],
                "color_argb": d[side + "_color_argb"],
            }
        sides["diagonal_up"] = bool(d["diagonal_up"])
        sides["diagonal_down"] = bool(d["diagonal_down"])
        return sides

    def get_dxf(self, dxf_index: int) -> DifferentialFormat:
        """Read one differential format by its conditional-format ``dxfId``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.DXF_RECORD)
        try:
            _check(LIB.fm_styles_get_dxf(h, int(dxf_index), ptr), "fm_styles_get_dxf")
            d = S.DXF_RECORD.unpack(LIB, ptr)
            offsets = S.DXF_RECORD.offsets
            font = None
            if d["font_engaged"]:
                fd = S.FONT_RECORD.unpack(LIB, ptr + offsets["font"][1])
                font = FontRecord(
                    name=LIB.read_cstr(fd["name"]),
                    size=fd["size"],
                    color_argb=fd["color_argb"],
                    bold=bool(fd["bold"]),
                    italic=bool(fd["italic"]),
                    strike=bool(fd["strike"]),
                    underline=fd["underline"],
                )
            fill = None
            if d["fill_engaged"]:
                fill_d = S.FILL_RECORD.unpack(LIB, ptr + offsets["fill"][1])
                fill = FillRecord(fill_d["pattern"], fill_d["fg_argb"], fill_d["bg_argb"])
            return DifferentialFormat(
                font=font,
                fill=fill,
                border=self._decode_border(ptr + offsets["border"][1]) if d["border_engaged"] else None,
                num_fmt_id=d["num_fmt_id"] if d["num_fmt_engaged"] else None,
                num_fmt_code=LIB.read_cstr(d["num_fmt_code"]) if d["num_fmt_engaged"] else "",
            )
        finally:
            LIB.free(ptr)

    def get_num_fmt(self, num_fmt_id: int) -> str:
        """Return the format code registered for ``num_fmt_id``."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_styles_get_num_fmt_string(h, int(num_fmt_id), out),
                "fm_styles_get_num_fmt_string",
            )
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(out)

    def font_count(self) -> int:
        """Return the number of font records registered."""
        return self._style_count(LIB.fm_styles_get_font_count)

    def fill_count(self) -> int:
        """Return the number of fill records registered."""
        return self._style_count(LIB.fm_styles_get_fill_count)

    def border_count(self) -> int:
        """Return the number of border records registered."""
        return self._style_count(LIB.fm_styles_get_border_count)

    def cell_xf_count(self) -> int:
        """Return the number of ``<xf>`` records registered."""
        return self._style_count(LIB.fm_styles_get_cell_xf_count)

    def cell_style_count(self) -> int:
        """Return the number of named cell styles."""
        return self._style_count(LIB.fm_styles_get_cell_style_count)

    def cell_style_xf_count(self) -> int:
        """Return the number of ``<cellStyleXfs>`` records."""
        return self._style_count(LIB.fm_styles_get_cell_style_xf_count)

    def dxf_count(self) -> int:
        """Return the number of differential formats available to CF rules."""
        return self._style_count(LIB.fm_styles_get_dxf_count)

    def _style_count(self, fn) -> int:
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(fn(h, out), getattr(fn, "__name__", "style_count"))
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    def add_font(self, record: FontRecord) -> int:
        """Add (dedup) a font record; return its index."""
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.FONT_RECORD)
        out = _alloc_out_ptr()
        try:
            S.FONT_RECORD.pack(
                LIB,
                ptr,
                {
                    "size": float(record.size),
                    "color_argb": int(record.color_argb),
                    "bold": 1 if record.bold else 0,
                    "italic": 1 if record.italic else 0,
                    "strike": 1 if record.strike else 0,
                    "underline": int(record.underline),
                },
            )
            S.write_str_field(LIB, ptr, S.FONT_RECORD, "name", record.name, owned)
            _check(LIB.fm_styles_add_font(h, ptr, out), "fm_styles_add_font")
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)
            for p in owned:
                LIB.free(p)

    def add_fill(self, record: FillRecord) -> int:
        """Add (dedup) a fill record; return its index."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.FILL_RECORD)
        out = _alloc_out_ptr()
        try:
            S.FILL_RECORD.pack(
                LIB,
                ptr,
                {
                    "pattern": int(record.pattern),
                    "fg_argb": int(record.fg_argb),
                    "bg_argb": int(record.bg_argb),
                },
            )
            _check(LIB.fm_styles_add_fill(h, ptr, out), "fm_styles_add_fill")
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)

    def add_border(self, sides: Dict[str, object]) -> int:
        """Add (dedup) a border record; return its index.

        ``sides`` mirrors :meth:`get_border`'s shape: a dict with
        ``left`` / ``right`` / ``top`` / ``bottom`` / ``diagonal`` entries
        (each ``{"style", "color_argb"}``) plus optional ``diagonal_up`` /
        ``diagonal_down`` booleans.
        """
        h = self._require()
        ptr = S.alloc_struct(LIB, S.BORDER_RECORD)
        out = _alloc_out_ptr()
        try:
            values: Dict[str, int] = {
                "diagonal_up": 1 if sides.get("diagonal_up") else 0,
                "diagonal_down": 1 if sides.get("diagonal_down") else 0,
            }
            for side in ("left", "right", "top", "bottom", "diagonal"):
                spec = sides.get(side) or {}
                values[side + "_style"] = int(spec.get("style", 0))
                values[side + "_color_argb"] = int(spec.get("color_argb", 0))
            S.BORDER_RECORD.pack(LIB, ptr, values)
            _check(LIB.fm_styles_add_border(h, ptr, out), "fm_styles_add_border")
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)

    def add_num_fmt(self, format_code: str) -> int:
        """Add (dedup) a number-format code; return its resolved id."""
        h = self._require()
        code_ptr, _ = LIB.alloc_utf8(format_code)
        out = _alloc_out_ptr()
        try:
            _check(LIB.fm_styles_add_num_fmt(h, code_ptr, out), "fm_styles_add_num_fmt")
            return struct.unpack("<H", LIB.read_bytes(out, 2))[0]
        finally:
            LIB.free(code_ptr)
            LIB.free(out)

    def add_cell_xf(self, record: CellXf) -> int:
        """Add (dedup) an ``<xf>`` record; return its index."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.CELL_XF_EX2)
        out = _alloc_out_ptr()
        try:
            S.CELL_XF_EX2.pack(
                LIB,
                ptr,
                {
                    "font_index": int(record.font_index),
                    "fill_index": int(record.fill_index),
                    "border_index": int(record.border_index),
                    "num_fmt_id": int(record.num_fmt_id),
                    "horizontal_align": int(record.horizontal_align),
                    "vertical_align": int(record.vertical_align),
                    "wrap_text": 1 if record.wrap_text else 0,
                    "has_alignment": 1
                    if (record.has_alignment if record.has_alignment is not None else _cell_xf_has_alignment(record))
                    else 0,
                    "justify_last_line": 1 if record.justify_last_line else 0,
                    "xf_id": int(record.xf_id),
                    "has_text_rotation": 1 if record.text_rotation is not None else 0,
                    "text_rotation": int(record.text_rotation or 0),
                    "has_indent": 1 if record.indent is not None else 0,
                    "indent": int(record.indent or 0),
                    "has_relative_indent": 1 if record.relative_indent is not None else 0,
                    "relative_indent": int(record.relative_indent or 0),
                    "has_shrink_to_fit": 1 if record.shrink_to_fit is not None else 0,
                    "shrink_to_fit": 1 if record.shrink_to_fit else 0,
                    "has_reading_order": 1 if record.reading_order is not None else 0,
                    "reading_order": int(record.reading_order or 0),
                    "has_horizontal_align": 1
                    if (
                        record.has_horizontal_align
                        if record.has_horizontal_align is not None
                        else record.horizontal_align != 0
                    )
                    else 0,
                    "has_vertical_align": 1
                    if (
                        record.has_vertical_align
                        if record.has_vertical_align is not None
                        else record.vertical_align != 2
                    )
                    else 0,
                    "has_wrap_text": 1
                    if (record.has_wrap_text if record.has_wrap_text is not None else record.wrap_text)
                    else 0,
                    "has_justify_last_line": 1
                    if (
                        record.has_justify_last_line
                        if record.has_justify_last_line is not None
                        else record.justify_last_line
                    )
                    else 0,
                },
            )
            _check(LIB.fm_styles_add_cell_xf_ex2(h, ptr, out), "fm_styles_add_cell_xf_ex2")
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)

    def add_dxf(self, record: DifferentialFormat) -> int:
        """Add (dedup) a differential format and return its ``dxfId``."""
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.DXF_RECORD)
        out = _alloc_out_ptr()
        try:
            S.DXF_RECORD.pack(
                LIB,
                ptr,
                {
                    "font_engaged": 1 if record.font is not None else 0,
                    "fill_engaged": 1 if record.fill is not None else 0,
                    "border_engaged": 1 if record.border is not None else 0,
                    "num_fmt_engaged": 1 if record.num_fmt_id is not None else 0,
                    "num_fmt_id": int(record.num_fmt_id or 0),
                },
            )
            offsets = S.DXF_RECORD.offsets
            if record.font is not None:
                font_ptr = ptr + offsets["font"][1]
                S.FONT_RECORD.pack(
                    LIB,
                    font_ptr,
                    {
                        "size": float(record.font.size),
                        "color_argb": int(record.font.color_argb),
                        "bold": 1 if record.font.bold else 0,
                        "italic": 1 if record.font.italic else 0,
                        "strike": 1 if record.font.strike else 0,
                        "underline": int(record.font.underline),
                    },
                )
                S.write_str_field(LIB, font_ptr, S.FONT_RECORD, "name", record.font.name, owned)
            if record.fill is not None:
                S.FILL_RECORD.pack(
                    LIB,
                    ptr + offsets["fill"][1],
                    {
                        "pattern": int(record.fill.pattern),
                        "fg_argb": int(record.fill.fg_argb),
                        "bg_argb": int(record.fill.bg_argb),
                    },
                )
            if record.border is not None:
                values: Dict[str, int] = {
                    "diagonal_up": 1 if record.border.get("diagonal_up") else 0,
                    "diagonal_down": 1 if record.border.get("diagonal_down") else 0,
                }
                for side in ("left", "right", "top", "bottom", "diagonal"):
                    spec = record.border.get(side) or {}
                    values[side + "_style"] = int(spec.get("style", 0))
                    values[side + "_color_argb"] = int(spec.get("color_argb", 0))
                S.BORDER_RECORD.pack(LIB, ptr + offsets["border"][1], values)
            S.write_str_field(LIB, ptr, S.DXF_RECORD, "num_fmt_code", record.num_fmt_code, owned)
            _check(LIB.fm_styles_add_dxf(h, ptr, out), "fm_styles_add_dxf")
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)
            for owned_ptr in owned:
                LIB.free(owned_ptr)

    def get_cell_style(self, index: int) -> CellStyle:
        """Return the named cell style at ``index``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.CELL_STYLE_RECORD)
        try:
            _check(
                LIB.fm_styles_get_cell_style(h, int(index), ptr),
                "fm_styles_get_cell_style",
            )
            d = S.CELL_STYLE_RECORD.unpack(LIB, ptr)
            return CellStyle(
                name=LIB.read_cstr(d["name"]),
                xf_id=d["xf_id"],
                builtin_id=d["builtin_id"],
                i_level=d["i_level"],
                hidden=bool(d["hidden"]),
                custom_builtin=bool(d["custom_builtin"]),
            )
        finally:
            LIB.free(ptr)

    def get_cell_style_xf(self, index: int) -> CellXf:
        """Return the named-style xf record at ``index``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.CELL_XF_EX2)
        try:
            _check(
                LIB.fm_styles_get_cell_style_xf_ex2(h, int(index), ptr),
                "fm_styles_get_cell_style_xf_ex2",
            )
            return self._decode_cell_xf(ptr)
        finally:
            LIB.free(ptr)

    # -- Pivot layout projection -------------------------------------------
    def pivot_count(self, sheet: int) -> int:
        """Return the number of PivotTables anchored on ``sheet``."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_count(h, int(sheet), out),
                "fm_workbook_pivot_count",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    def pivot_layout(self, sheet: int, pivot_index: int) -> PivotLayout:
        """Evaluate and project a PivotTable into concrete grid cells."""
        h = self._require()
        out = _alloc_out_ptr()
        handle = 0
        try:
            _check(
                LIB.fm_workbook_pivot_layout(h, int(sheet), int(pivot_index), out),
                "fm_workbook_pivot_layout",
            )
            handle = LIB.read_u32(out)
            top_p = _alloc_out_ptr()
            left_p = _alloc_out_ptr()
            rows_p = _alloc_out_ptr()
            cols_p = _alloc_out_ptr()
            try:
                _check(
                    LIB.fm_pivot_cells_bounds(handle, top_p, left_p, rows_p, cols_p),
                    "fm_pivot_cells_bounds",
                )
                top = LIB.read_u32(top_p)
                left = LIB.read_u32(left_p)
                rows = LIB.read_u32(rows_p)
                cols = LIB.read_u32(cols_p)
            finally:
                LIB.free(top_p)
                LIB.free(left_p)
                LIB.free(rows_p)
                LIB.free(cols_p)
            cells: List[PivotCell] = []
            n = LIB.fm_pivot_cells_count(handle)
            for i in range(n):
                cptr = S.alloc_struct(LIB, S.PIVOT_CELL)
                try:
                    _check(
                        LIB.fm_pivot_cells_at(handle, i, cptr),
                        "fm_pivot_cells_at",
                    )
                    d = S.PIVOT_CELL.unpack(LIB, cptr)
                    value = Value._from_wasm(cptr + S.PIVOT_CELL_VALUE_OFFSET)
                    cells.append(
                        PivotCell(
                            row=d["row"],
                            col=d["col"],
                            value=value,
                            kind=d["kind"],
                            depth=d["depth"],
                            field_name=LIB.read_cstr(d["field_name"]),
                            number_format=LIB.read_cstr(d["number_format"]),
                        )
                    )
                finally:
                    LIB.free(cptr)
            return PivotLayout(top=top, left=left, rows=rows, cols=cols, cells=cells)
        finally:
            if handle:
                LIB.fm_pivot_cells_destroy(handle)
            LIB.free(out)

    # -- Pivot caches ------------------------------------------------------
    def pivot_cache_count(self) -> int:
        """Return the number of pivot caches owned by the workbook."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_cache_count(h, out),
                "fm_workbook_pivot_cache_count",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    def pivot_cache_id_at(self, index: int) -> int:
        """Return the cache id at flat ``index``."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_cache_id_at(h, int(index), out),
                "fm_workbook_pivot_cache_id_at",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    def pivot_cache_create(self, requested_id: int = 0) -> int:
        """Create a new empty pivot cache; return its assigned id."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_cache_create(h, int(requested_id), out),
                "fm_workbook_pivot_cache_create",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    def get_pivot_cache_worksheet_source(self, cache_id: int) -> Optional[PivotWorksheetSource]:
        """Return a cache's worksheet-source metadata, or ``None`` when absent."""
        h = self._require()
        present = _alloc_out_ptr()
        out_ref = _alloc_out_ptr()
        out_sheet = _alloc_out_ptr()
        out_name = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_cache_get_worksheet_source(
                    h, int(cache_id), present, out_ref, out_sheet, out_name
                ),
                "fm_workbook_pivot_cache_get_worksheet_source",
            )
            if not LIB.read_i32(present):
                return None
            return PivotWorksheetSource(
                LIB.read_cstr(LIB.read_u32(out_ref)) if LIB.read_u32(out_ref) else None,
                LIB.read_cstr(LIB.read_u32(out_sheet)) if LIB.read_u32(out_sheet) else None,
                LIB.read_cstr(LIB.read_u32(out_name)) if LIB.read_u32(out_name) else None,
            )
        finally:
            LIB.free(present)
            LIB.free(out_ref)
            LIB.free(out_sheet)
            LIB.free(out_name)

    def set_pivot_cache_worksheet_source(self, cache_id: int, source: Optional[PivotWorksheetSource]) -> None:
        """Set cache worksheet-source metadata; pass ``None`` to clear it."""
        h = self._require()
        owned: List[int] = []
        try:
            values = () if source is None else (source.ref, source.sheet, source.name)
            for value in values:
                if value is None:
                    owned.append(0)
                else:
                    pointer, _ = LIB.alloc_utf8(value)
                    owned.append(pointer)
            pointers = owned if source is not None else (0, 0, 0)
            _check(
                LIB.fm_workbook_pivot_cache_set_worksheet_source(
                    h, int(cache_id), 0 if source is None else 1, *pointers
                ),
                "fm_workbook_pivot_cache_set_worksheet_source",
            )
        finally:
            for pointer in owned:
                if pointer:
                    LIB.free(pointer)

    def pivot_cache_remove(self, cache_id: int) -> None:
        """Remove the pivot cache with id ``cache_id``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_remove(h, int(cache_id)),
            "fm_workbook_pivot_cache_remove",
        )

    def pivot_cache_field_count(self, cache_id: int) -> int:
        """Return the number of fields on the cache."""
        h = self._require()
        return _read_count(LIB.fm_workbook_pivot_cache_field_count, h, int(cache_id))

    def pivot_cache_field_name(self, cache_id: int, field_idx: int) -> str:
        """Read the name of cache field ``field_idx``."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_cache_field_name(h, int(cache_id), int(field_idx), out),
                "fm_workbook_pivot_cache_field_name",
            )
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(out)

    def pivot_cache_field_add(self, cache_id: int, name: str) -> int:
        """Append a field to the cache; return its index."""
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(name)
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_cache_field_add(h, int(cache_id), name_ptr, out),
                "fm_workbook_pivot_cache_field_add",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(name_ptr)
            LIB.free(out)

    def pivot_cache_field_clear(self, cache_id: int) -> None:
        """Drop every field (and record) from the cache."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_field_clear(h, int(cache_id)),
            "fm_workbook_pivot_cache_field_clear",
        )

    def pivot_cache_field_shared_item_count(self, cache_id: int, field_idx: int) -> int:
        """Return the shared-item count on cache field ``field_idx``."""
        h = self._require()
        return _read_count(
            LIB.fm_workbook_pivot_cache_field_shared_item_count,
            h,
            int(cache_id),
            int(field_idx),
        )

    def pivot_cache_field_add_shared_item_number(self, cache_id: int, field_idx: int, value: float) -> None:
        """Append a numeric shared item to a cache field."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_field_add_shared_item_number(h, int(cache_id), int(field_idx), float(value)),
            "fm_workbook_pivot_cache_field_add_shared_item_number",
        )

    def pivot_cache_field_add_shared_item_text(self, cache_id: int, field_idx: int, value: str) -> None:
        """Append a text shared item to a cache field."""
        h = self._require()
        vp, _ = LIB.alloc_utf8(value)
        try:
            _check(
                LIB.fm_workbook_pivot_cache_field_add_shared_item_text(h, int(cache_id), int(field_idx), vp),
                "fm_workbook_pivot_cache_field_add_shared_item_text",
            )
        finally:
            LIB.free(vp)

    def pivot_cache_field_add_shared_item_bool(self, cache_id: int, field_idx: int, value: bool) -> None:
        """Append a boolean shared item to a cache field."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_field_add_shared_item_bool(h, int(cache_id), int(field_idx), 1 if value else 0),
            "fm_workbook_pivot_cache_field_add_shared_item_bool",
        )

    def pivot_cache_field_add_shared_item_blank(self, cache_id: int, field_idx: int) -> None:
        """Append a blank shared item to a cache field."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_field_add_shared_item_blank(h, int(cache_id), int(field_idx)),
            "fm_workbook_pivot_cache_field_add_shared_item_blank",
        )

    def pivot_cache_field_add_shared_item_error(self, cache_id: int, field_idx: int, error_code: int) -> None:
        """Append an Excel error shared item to a cache field."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_field_add_shared_item_error(h, int(cache_id), int(field_idx), int(error_code)),
            "fm_workbook_pivot_cache_field_add_shared_item_error",
        )

    def pivot_cache_field_clear_shared_items(self, cache_id: int, field_idx: int) -> None:
        """Drop every shared item from a cache field."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_field_clear_shared_items(h, int(cache_id), int(field_idx)),
            "fm_workbook_pivot_cache_field_clear_shared_items",
        )

    def pivot_cache_record_count(self, cache_id: int) -> int:
        """Return the number of records on the cache."""
        h = self._require()
        return _read_count(LIB.fm_workbook_pivot_cache_record_count, h, int(cache_id))

    def pivot_cache_record_add(self, cache_id: int) -> int:
        """Append an empty record to the cache; return its index."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_cache_record_add(h, int(cache_id), out),
                "fm_workbook_pivot_cache_record_add",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    def pivot_cache_record_clear(self, cache_id: int) -> None:
        """Drop every record from the cache."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_record_clear(h, int(cache_id)),
            "fm_workbook_pivot_cache_record_clear",
        )

    def pivot_cache_record_set_number(self, cache_id: int, record_idx: int, field_idx: int, value: float) -> None:
        """Set cache cell ``(record_idx, field_idx)`` to a number."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_record_set_number(
                h, int(cache_id), int(record_idx), int(field_idx), float(value)
            ),
            "fm_workbook_pivot_cache_record_set_number",
        )

    def pivot_cache_record_set_text(self, cache_id: int, record_idx: int, field_idx: int, value: str) -> None:
        """Set cache cell ``(record_idx, field_idx)`` to text."""
        h = self._require()
        vp, _ = LIB.alloc_utf8(value)
        try:
            _check(
                LIB.fm_workbook_pivot_cache_record_set_text(h, int(cache_id), int(record_idx), int(field_idx), vp),
                "fm_workbook_pivot_cache_record_set_text",
            )
        finally:
            LIB.free(vp)

    def pivot_cache_record_set_bool(self, cache_id: int, record_idx: int, field_idx: int, value: bool) -> None:
        """Set cache cell ``(record_idx, field_idx)`` to a boolean."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_record_set_bool(
                h, int(cache_id), int(record_idx), int(field_idx), 1 if value else 0
            ),
            "fm_workbook_pivot_cache_record_set_bool",
        )

    def pivot_cache_record_set_blank(self, cache_id: int, record_idx: int, field_idx: int) -> None:
        """Set cache cell ``(record_idx, field_idx)`` to blank."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_record_set_blank(h, int(cache_id), int(record_idx), int(field_idx)),
            "fm_workbook_pivot_cache_record_set_blank",
        )

    def pivot_cache_record_set_error(self, cache_id: int, record_idx: int, field_idx: int, error_code: int) -> None:
        """Set cache cell ``(record_idx, field_idx)`` to an Excel error."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_record_set_error(
                h, int(cache_id), int(record_idx), int(field_idx), int(error_code)
            ),
            "fm_workbook_pivot_cache_record_set_error",
        )

    # -- Pivot tables ------------------------------------------------------
    def pivot_create(self, sheet: int, name: str, cache_id: int, anchor_row: int, anchor_col: int) -> int:
        """Create a new empty pivot table; return its flat index."""
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(name)
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_create(
                    h,
                    int(sheet),
                    name_ptr,
                    int(cache_id),
                    int(anchor_row),
                    int(anchor_col),
                    out,
                ),
                "fm_workbook_pivot_create",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(name_ptr)
            LIB.free(out)

    def get_pivot_report_layout(self, sheet: int, pivot_index: int) -> PivotReportLayout:
        """Return the pivot's compact, tabular, or outline report layout."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_get_layout(h, int(sheet), int(pivot_index), out),
                "fm_workbook_pivot_get_layout",
            )
            return PivotReportLayout(LIB.read_i32(out))
        finally:
            LIB.free(out)

    def set_pivot_report_layout(self, sheet: int, pivot_index: int, layout: PivotReportLayout) -> None:
        """Set the pivot's compact, tabular, or outline report layout."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_set_layout(h, int(sheet), int(pivot_index), int(layout)),
            "fm_workbook_pivot_set_layout",
        )

    def pivot_remove(self, sheet: int, pivot_index: int) -> None:
        """Remove the pivot table at ``pivot_index``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_remove(h, int(sheet), int(pivot_index)),
            "fm_workbook_pivot_remove",
        )

    def pivot_set_name(self, sheet: int, pivot_index: int, name: str) -> None:
        """Rename the pivot table."""
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(name)
        try:
            _check(
                LIB.fm_workbook_pivot_set_name(h, int(sheet), int(pivot_index), name_ptr),
                "fm_workbook_pivot_set_name",
            )
        finally:
            LIB.free(name_ptr)

    def pivot_set_anchor(
        self,
        sheet: int,
        pivot_index: int,
        anchor_row: int,
        anchor_col: int,
        span_rows: int,
        span_cols: int,
    ) -> None:
        """Update the pivot's anchor cell and span."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_set_anchor(
                h,
                int(sheet),
                int(pivot_index),
                int(anchor_row),
                int(anchor_col),
                int(span_rows),
                int(span_cols),
            ),
            "fm_workbook_pivot_set_anchor",
        )

    def pivot_set_grand_totals(self, sheet: int, pivot_index: int, rows_enabled: bool, cols_enabled: bool) -> None:
        """Toggle the row / column grand-total bands."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_set_grand_totals(
                h,
                int(sheet),
                int(pivot_index),
                1 if rows_enabled else 0,
                1 if cols_enabled else 0,
            ),
            "fm_workbook_pivot_set_grand_totals",
        )

    def pivot_field_count(self, sheet: int, pivot_index: int) -> int:
        """Return the number of fields configured on the pivot."""
        h = self._require()
        return _read_count(LIB.fm_workbook_pivot_field_count, h, int(sheet), int(pivot_index))

    def pivot_field_add(self, sheet: int, pivot_index: int, spec: PivotFieldSpec) -> int:
        """Append a field to the pivot; return its index."""
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.PIVOT_FIELD_SPEC)
        out = _alloc_out_ptr()
        try:
            S.PIVOT_FIELD_SPEC.pack(
                LIB,
                ptr,
                {
                    "axis": int(spec.axis),
                    "subtotal_top": 1 if spec.subtotal_top else 0,
                },
            )
            S.write_str_field(LIB, ptr, S.PIVOT_FIELD_SPEC, "source_name", spec.source_name, owned)
            S.write_str_field(LIB, ptr, S.PIVOT_FIELD_SPEC, "custom_name", spec.custom_name, owned)
            S.write_str_field(
                LIB,
                ptr,
                S.PIVOT_FIELD_SPEC,
                "number_format",
                spec.number_format,
                owned,
            )
            _check(
                LIB.fm_workbook_pivot_field_add(h, int(sheet), int(pivot_index), ptr, out),
                "fm_workbook_pivot_field_add",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)
            for p in owned:
                LIB.free(p)

    def pivot_field_clear(self, sheet: int, pivot_index: int) -> None:
        """Drop every field from the pivot."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_clear(h, int(sheet), int(pivot_index)),
            "fm_workbook_pivot_field_clear",
        )

    def pivot_field_set_axis(self, sheet: int, pivot_index: int, field_idx: int, axis: int) -> None:
        """Set the axis of pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_set_axis(h, int(sheet), int(pivot_index), int(field_idx), int(axis)),
            "fm_workbook_pivot_field_set_axis",
        )

    def pivot_field_set_sort(
        self,
        sheet: int,
        pivot_index: int,
        field_idx: int,
        ascending: bool,
        by_field: str = "",
    ) -> None:
        """Set the sort directive on pivot field ``field_idx``."""
        h = self._require()
        owned: List[int] = []
        by_ptr = _opt_str_ptr(by_field, owned)
        try:
            _check(
                LIB.fm_workbook_pivot_field_set_sort(
                    h,
                    int(sheet),
                    int(pivot_index),
                    int(field_idx),
                    1 if ascending else 0,
                    by_ptr,
                ),
                "fm_workbook_pivot_field_set_sort",
            )
        finally:
            for p in owned:
                LIB.free(p)

    def pivot_field_set_subtotal_top(self, sheet: int, pivot_index: int, field_idx: int, top: bool) -> None:
        """Set the ``subtotal_top`` flag on pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_set_subtotal_top(
                h, int(sheet), int(pivot_index), int(field_idx), 1 if top else 0
            ),
            "fm_workbook_pivot_field_set_subtotal_top",
        )

    def pivot_field_add_aggregation(self, sheet: int, pivot_index: int, field_idx: int, agg: int) -> None:
        """Append an aggregation to pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_add_aggregation(h, int(sheet), int(pivot_index), int(field_idx), int(agg)),
            "fm_workbook_pivot_field_add_aggregation",
        )

    def pivot_field_clear_aggregations(self, sheet: int, pivot_index: int, field_idx: int) -> None:
        """Drop every aggregation from pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_clear_aggregations(h, int(sheet), int(pivot_index), int(field_idx)),
            "fm_workbook_pivot_field_clear_aggregations",
        )

    def pivot_field_add_item(self, sheet: int, pivot_index: int, field_idx: int, name: str, visible: bool) -> None:
        """Append a manual-filter item to pivot field ``field_idx``."""
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(name)
        try:
            _check(
                LIB.fm_workbook_pivot_field_add_item(
                    h,
                    int(sheet),
                    int(pivot_index),
                    int(field_idx),
                    name_ptr,
                    1 if visible else 0,
                ),
                "fm_workbook_pivot_field_add_item",
            )
        finally:
            LIB.free(name_ptr)

    def pivot_field_clear_items(self, sheet: int, pivot_index: int, field_idx: int) -> None:
        """Drop every manual-filter item from pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_clear_items(h, int(sheet), int(pivot_index), int(field_idx)),
            "fm_workbook_pivot_field_clear_items",
        )

    def pivot_field_set_item_visible(
        self,
        sheet: int,
        pivot_index: int,
        field_idx: int,
        item_idx: int,
        visible: bool,
    ) -> None:
        """Toggle the visibility of item ``item_idx`` on field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_set_item_visible(
                h,
                int(sheet),
                int(pivot_index),
                int(field_idx),
                int(item_idx),
                1 if visible else 0,
            ),
            "fm_workbook_pivot_field_set_item_visible",
        )

    def pivot_field_add_subtotal_fn(self, sheet: int, pivot_index: int, field_idx: int, agg: int) -> None:
        """Append a subtotal-fn entry to pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_add_subtotal_fn(h, int(sheet), int(pivot_index), int(field_idx), int(agg)),
            "fm_workbook_pivot_field_add_subtotal_fn",
        )

    def pivot_field_clear_subtotal_fns(self, sheet: int, pivot_index: int, field_idx: int) -> None:
        """Drop every subtotal-fn entry from pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_clear_subtotal_fns(h, int(sheet), int(pivot_index), int(field_idx)),
            "fm_workbook_pivot_field_clear_subtotal_fns",
        )

    def pivot_field_set_date_group(
        self,
        sheet: int,
        pivot_index: int,
        field_idx: int,
        granularity: int,
        calendar: int,
        start_year: int = -1,
        end_year: int = -1,
    ) -> None:
        """Configure date-grouping on pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_set_date_group(
                h,
                int(sheet),
                int(pivot_index),
                int(field_idx),
                int(granularity),
                int(calendar),
                int(start_year),
                int(end_year),
            ),
            "fm_workbook_pivot_field_set_date_group",
        )

    def pivot_field_clear_date_group(self, sheet: int, pivot_index: int, field_idx: int) -> None:
        """Remove the date-grouping config from pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_clear_date_group(h, int(sheet), int(pivot_index), int(field_idx)),
            "fm_workbook_pivot_field_clear_date_group",
        )

    def pivot_field_set_number_format(self, sheet: int, pivot_index: int, field_idx: int, fmt: str) -> None:
        """Set the OOXML number-format string on pivot field ``field_idx``."""
        h = self._require()
        fmt_ptr, _ = LIB.alloc_utf8(fmt)
        try:
            _check(
                LIB.fm_workbook_pivot_field_set_number_format(h, int(sheet), int(pivot_index), int(field_idx), fmt_ptr),
                "fm_workbook_pivot_field_set_number_format",
            )
        finally:
            LIB.free(fmt_ptr)

    def pivot_set_row_field_order(self, sheet: int, pivot_index: int, indices: Sequence[int]) -> None:
        """Replace the row-axis field order with ``indices``."""
        self._pivot_set_field_order(LIB.fm_workbook_pivot_set_row_field_order, sheet, pivot_index, indices)

    def pivot_set_col_field_order(self, sheet: int, pivot_index: int, indices: Sequence[int]) -> None:
        """Replace the column-axis field order with ``indices``."""
        self._pivot_set_field_order(LIB.fm_workbook_pivot_set_col_field_order, sheet, pivot_index, indices)

    def _pivot_set_field_order(self, fn, sheet: int, pivot_index: int, indices: Sequence[int]) -> None:
        h = self._require()
        arr_ptr = 0
        try:
            if indices:
                arr_ptr = LIB.alloc(4 * len(indices))
                buf = b"".join(struct.pack("<I", int(i)) for i in indices)
                LIB.write_bytes(arr_ptr, buf)
            _check(
                fn(h, int(sheet), int(pivot_index), arr_ptr, len(indices)),
                getattr(fn, "__name__", "pivot_set_field_order"),
            )
        finally:
            if arr_ptr:
                LIB.free(arr_ptr)

    def pivot_data_field_count(self, sheet: int, pivot_index: int) -> int:
        """Return the number of ``<dataField>`` entries on the pivot."""
        h = self._require()
        return _read_count(LIB.fm_workbook_pivot_data_field_count, h, int(sheet), int(pivot_index))

    def pivot_data_field_add(self, sheet: int, pivot_index: int, spec: PivotDataFieldSpec) -> int:
        """Append a data-field entry; return its index."""
        h = self._require()
        out = _alloc_out_ptr()
        owned: List[int] = []
        ptr = self._pack_data_field_spec(spec, owned)
        try:
            _check(
                LIB.fm_workbook_pivot_data_field_add(h, int(sheet), int(pivot_index), ptr, out),
                "fm_workbook_pivot_data_field_add",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)
            for p in owned:
                LIB.free(p)

    def pivot_data_field_set(
        self,
        sheet: int,
        pivot_index: int,
        data_field_idx: int,
        spec: PivotDataFieldSpec,
    ) -> None:
        """Replace the data-field entry at ``data_field_idx`` in place."""
        h = self._require()
        owned: List[int] = []
        ptr = self._pack_data_field_spec(spec, owned)
        try:
            _check(
                LIB.fm_workbook_pivot_data_field_set(h, int(sheet), int(pivot_index), int(data_field_idx), ptr),
                "fm_workbook_pivot_data_field_set",
            )
        finally:
            LIB.free(ptr)
            for p in owned:
                LIB.free(p)

    @staticmethod
    def _pack_data_field_spec(spec: PivotDataFieldSpec, owned: List[int]) -> int:
        ptr = S.alloc_struct(LIB, S.PIVOT_DATA_FIELD_SPEC)
        S.PIVOT_DATA_FIELD_SPEC.pack(
            LIB,
            ptr,
            {
                "field_index": int(spec.field_index),
                "aggregation": int(spec.aggregation),
                "show_as": int(spec.show_as),
                "show_as_base_field": int(spec.show_as_base_field),
                "show_as_base_item": int(spec.show_as_base_item),
            },
        )
        S.write_str_field(LIB, ptr, S.PIVOT_DATA_FIELD_SPEC, "name", spec.name, owned)
        S.write_str_field(
            LIB,
            ptr,
            S.PIVOT_DATA_FIELD_SPEC,
            "number_format",
            spec.number_format,
            owned,
        )
        return ptr

    def pivot_data_field_clear(self, sheet: int, pivot_index: int) -> None:
        """Drop every data-field entry from the pivot."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_data_field_clear(h, int(sheet), int(pivot_index)),
            "fm_workbook_pivot_data_field_clear",
        )

    def pivot_filter_count(self, sheet: int, pivot_index: int) -> int:
        """Return the number of active filters on the pivot."""
        h = self._require()
        return _read_count(LIB.fm_workbook_pivot_filter_count, h, int(sheet), int(pivot_index))

    def pivot_filter_add(self, sheet: int, pivot_index: int, spec: PivotFilterSpec) -> None:
        """Append an active filter to the pivot."""
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.PIVOT_FILTER_SPEC_EX)
        try:
            S.PIVOT_FILTER_SPEC_EX.pack(
                LIB,
                ptr,
                {
                    "axis": int(spec.axis),
                    "type": int(spec.type),
                    "data_field_index": int(spec.data_field_index),
                    "value_kind": int(spec.value_kind),
                    "value_int": int(spec.value_int),
                    "value_double": float(spec.value_double),
                    "value_high_kind": int(spec.value_high_kind),
                    "value_high_int": int(spec.value_high_int),
                    "value_high_double": float(spec.value_high_double),
                },
            )
            S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC_EX, "field_name", spec.field_name, owned)
            S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC_EX, "value_text", spec.value_text, owned)
            _check(
                LIB.fm_workbook_pivot_filter_add_ex(h, int(sheet), int(pivot_index), ptr),
                "fm_workbook_pivot_filter_add_ex",
            )
        finally:
            LIB.free(ptr)
            for p in owned:
                LIB.free(p)

    def pivot_filter_clear(self, sheet: int, pivot_index: int) -> None:
        """Drop every active filter from the pivot."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_filter_clear(h, int(sheet), int(pivot_index)),
            "fm_workbook_pivot_filter_clear",
        )

    def pivot_filter_remove_at(self, sheet: int, pivot_index: int, filter_idx: int) -> None:
        """Remove the active filter at ``filter_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_filter_remove_at(h, int(sheet), int(pivot_index), int(filter_idx)),
            "fm_workbook_pivot_filter_remove_at",
        )

    # -- Dependency-graph trace --------------------------------------------
    def precedents(self, sheet: int, row: int, col: int, depth: int = 1) -> List[CellNode]:
        """Return the cells ``(sheet, row, col)`` reads (up to ``depth``)."""
        return self._trace(LIB.fm_workbook_precedents, sheet, row, col, depth)

    def dependents(self, sheet: int, row: int, col: int, depth: int = 1) -> List[CellNode]:
        """Return the cells that read ``(sheet, row, col)`` (up to ``depth``)."""
        return self._trace(LIB.fm_workbook_dependents, sheet, row, col, depth)

    def _trace(self, fn, sheet: int, row: int, col: int, depth: int) -> List[CellNode]:
        h = self._require()
        out = _alloc_out_ptr()
        handle = 0
        try:
            _check(
                fn(h, int(sheet), int(row), int(col), int(depth), out),
                getattr(fn, "__name__", "trace"),
            )
            handle = LIB.read_u32(out)
            nodes: List[CellNode] = []
            n = LIB.fm_cell_nodes_count(handle)
            for i in range(n):
                nptr = S.alloc_struct(LIB, S.CELL_NODE)
                try:
                    _check(LIB.fm_cell_nodes_at(handle, i, nptr), "fm_cell_nodes_at")
                    d = S.CELL_NODE.unpack(LIB, nptr)
                    nodes.append(CellNode(sheet=d["sheet"], row=d["row"], col=d["col"]))
                finally:
                    LIB.free(nptr)
            return nodes
        finally:
            if handle:
                LIB.fm_cell_nodes_destroy(handle)
            LIB.free(out)

    # -- Dynamic-array spill -----------------------------------------------
    def spill_info(self, sheet: int, row: int, col: int) -> SpillInfo:
        """Return dynamic-array spill info for ``(sheet, row, col)``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.SPILL_INFO)
        try:
            _check(
                LIB.fm_workbook_spill_info(h, int(sheet), int(row), int(col), ptr),
                "fm_workbook_spill_info",
            )
            d = S.SPILL_INFO.unpack(LIB, ptr)
            return SpillInfo(
                engaged=bool(d["engaged"]),
                anchor_row=d["anchor_row"],
                anchor_col=d["anchor_col"],
                rows=d["rows"],
                cols=d["cols"],
            )
        finally:
            LIB.free(ptr)

    # -- Function catalog (workbook-independent) ---------------------------
    @staticmethod
    def function_count() -> int:
        """Return the total number of registered Formulon functions."""
        return int(LIB.fm_function_count())

    @staticmethod
    def function_name_at(index: int) -> str:
        """Return the canonical name of the ``index``-th function."""
        out = _alloc_out_ptr()
        try:
            _check(LIB.fm_function_name_at(int(index), out), "fm_function_name_at")
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(out)

    @staticmethod
    def function_metadata(name: str, locale: int = 0) -> Optional[FunctionMetadata]:
        """Return metadata for ``name`` or ``None`` when unknown.

        ``locale`` is ``0`` for ``en-US`` and ``1`` for ``ja-JP``.
        """
        name_ptr, _ = LIB.alloc_utf8(name)
        ptr = S.alloc_struct(LIB, S.FUNCTION_METADATA)
        try:
            status = LIB.fm_function_metadata(name_ptr, int(locale), ptr)
            if status != 0:
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
                availability=d["availability"],
                signature_template=sig,
                description=desc,
            )
        finally:
            LIB.free(name_ptr)
            LIB.free(ptr)

    @staticmethod
    def localize_function_name(canonical_name: str, locale: int = 0) -> str:
        """Return the localized display name for ``canonical_name``."""
        name_ptr, _ = LIB.alloc_utf8(canonical_name)
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_function_localize(name_ptr, int(locale), out),
                "fm_function_localize",
            )
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(name_ptr)
            LIB.free(out)

    @staticmethod
    def canonicalize_function_name(localized_name: str, locale: int = 0) -> str:
        """Return the canonical English name for ``localized_name``."""
        name_ptr, _ = LIB.alloc_utf8(localized_name)
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_function_canonicalize(name_ptr, int(locale), out),
                "fm_function_canonicalize",
            )
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(name_ptr)
            LIB.free(out)

    # -- External links ----------------------------------------------------
    def external_link_count(self) -> int:
        """Return the number of external-link records on the workbook."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_external_link_count(h, out),
                "fm_workbook_external_link_count",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    def get_external_link_at(self, index: int) -> ExternalLink:
        """Read the ``index``-th external-link record."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.EXTERNAL_LINK_RECORD)
        try:
            _check(
                LIB.fm_workbook_external_link_at(h, int(index), ptr),
                "fm_workbook_external_link_at",
            )
            d = S.EXTERNAL_LINK_RECORD.unpack(LIB, ptr)
            return ExternalLink(
                index=d["index"],
                rel_id=LIB.read_cstr(d["rel_id"]),
                part_path=LIB.read_cstr(d["part_path"]),
                target=LIB.read_cstr(d["target"]),
                target_external=bool(d["target_external"]),
                kind=d["kind"],
            )
        finally:
            LIB.free(ptr)

    def get_external_links(self) -> List[ExternalLink]:
        """Return every external-link record in document order."""
        n = self.external_link_count()
        return [self.get_external_link_at(i) for i in range(n)]
