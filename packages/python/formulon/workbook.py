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

Sheet, row, column, and count arguments are 32-bit on the C ABI. A value
that does not fit the parameter it is bound for raises ``ValueError``
before the call crosses into WebAssembly, where it would otherwise be
truncated modulo 2**32 into a different, valid-looking coordinate.

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
    "ColorSpec",
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
    "ErrorCode",
    "ExternalLink",
    "ExternalLinkKind",
    "FillRecord",
    "FontRecord",
    "FormulonError",
    "FunctionMetadata",
    "Hyperlink",
    "IterativeSettings",
    "LogLevel",
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
    "ReadDiagnostics",
    "RowLayout",
    "SaveDiagnostics",
    "SheetProtection",
    "SheetView",
    "SheetVisibility",
    "SpillInfo",
    "Table",
    "Value",
    "ValueKind",
    "Workbook",
]


# ---------------------------------------------------------------------------
# Error type
# ---------------------------------------------------------------------------

# `formulon::FormulonErrorCode` ordinals. The C ABI intentionally exposes
# status codes as integers, so bindings retain these matching stable values.
_STATUS_INVALID_ARGUMENT = 2
_STATUS_NOT_FOUND = 6
# 7000-band: bindings / C API (src/utils/error.h).
_STATUS_BINDING_INVALID_HANDLE = 7000
_STATUS_BINDING_NULL_POINTER = 7001

# Inclusive bounds of `fm_locale_t` (0 = en-US, 1 = ja-JP).
_LOCALE_MIN = 0
_LOCALE_MAX = 1


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

    def __init__(
        self,
        status: int,
        *,
        op: str = "",
        _diagnostic_override: Optional[tuple[str, str]] = None,
    ) -> None:
        self.status = int(status)
        # Take the Python-side snapshot before calling fm_status_string.
        # The latter is a non-status pointer-returning export and must not
        # become a diagnostic read or overwrite the pending snapshot. An
        # override still drains a pending C diagnostic so it cannot leak into
        # a later exception on this Python thread.
        message, context = LIB.last_diagnostic(self.status)
        if _diagnostic_override is not None:
            message, context = _diagnostic_override
        self.status_name = LIB.read_cstr(LIB.fm_status_string(_sint(self.status, "status")))
        self.message, self.context = message, context
        prefix = f"{op}: " if op else ""
        text = f"{prefix}{self.status_name} ({self.status})"
        if self.message:
            text += f": {self.message}"
        if self.context:
            text += f" [{self.context}]"
        super().__init__(text)


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


def _check(status: int, op: str) -> None:
    """Raise :class:`FormulonError` if ``status`` is non-zero."""
    if status != 0:
        raise FormulonError(status, op=op)


def _uint(value: int, name: str, bits: int = 32) -> int:
    """Validate ``value`` as an unsigned WASM argument of ``bits`` width.

    Every integer the C ABI takes by value crosses into WebAssembly as an
    ``i32``, and ``wasmtime`` marshals a Python ``int`` into that slot
    modulo 2**32 without complaint. An index the caller never intended --
    ``2**32 + 5``, or a negative produced by an arithmetic slip -- would
    therefore arrive on the engine side as a perfectly plausible
    coordinate (``5``) and silently overwrite an unrelated cell. The
    engine's own grid bounds cannot catch that: they see only the wrapped
    value.

    The same applies to an integer packed into a struct the ABI reads
    through a pointer, with a different failure: ``struct.pack`` rejects
    the value outright and raises ``struct.error``, which is not the
    exception this API documents and not one a caller catches. Both routes
    go through here so the diagnosis is the same either way.

    Args:
      value: the caller-supplied index, count, or length.
      name: the C ABI parameter or struct field name, used in the message.
      bits: width of the destination (32 for ``uint32_t`` / ``size_t``,
        16 and 8 for the narrower struct fields, which overflow well
        before an ``i32`` does).

    Returns:
      ``value`` as an ``int``, unchanged.

    Raises:
      ValueError: when ``value`` is negative or does not fit ``bits``.
    """
    ivalue = int(value)
    if ivalue < 0 or ivalue >= (1 << bits):
        raise ValueError(f"formulon: {name} out of range for uint{bits}: {ivalue}")
    return ivalue


def _sint(value: int, name: str, bits: int = 32) -> int:
    """Validate ``value`` as a signed WASM argument of ``bits`` width.

    Signed counterpart of :func:`_uint`, for the ``int32_t`` parameters
    and struct fields (enum selectors, iteration caps, the ``-1``
    sentinels of the pivot date-group and show-as APIs) where a negative
    value is meaningful.

    Raises:
      ValueError: when ``value`` does not fit a signed ``bits``-wide int.
    """
    ivalue = int(value)
    limit = 1 << (bits - 1)
    if ivalue < -limit or ivalue >= limit:
        raise ValueError(f"formulon: {name} out of range for int{bits}: {ivalue}")
    return ivalue


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
    by :meth:`Workbook.set_error`.
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
    """Container format selector for :meth:`Workbook.save_as`.

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
    """A sheet hyperlink covering an inclusive cell rectangle."""

    row: int
    col: int
    last_row: int
    last_col: int
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
class CivilTime:
    """A wall-clock reading in local civil fields.

    Local calendar fields are carried rather than a timestamp so a pinned
    reading has no residual timezone interpretation: the same values
    reproduce the same results on any host, in any zone.
    """

    year: int
    month: int
    day: int
    hour: int = 0
    minute: int = 0
    second: int = 0


@dataclass(frozen=True)
class SheetView:
    """Per-sheet view: zoom, frozen-pane counts, tab visibility, and the
    display / orientation flags mirrored from OOXML ``<sheetView>``.

    ``visibility`` is the authoritative tab state; ``tab_hidden`` is its
    two-state view and is true for both hidden states, so code that knows
    only the bool sees a very-hidden sheet as hidden rather than visible.
    """

    zoom_scale: int
    freeze_rows: int
    freeze_cols: int
    tab_hidden: bool
    visibility: SheetVisibility = SheetVisibility.VISIBLE
    show_grid_lines: bool = True
    show_row_col_headers: bool = True
    show_zeros: bool = True
    right_to_left: bool = False
    tab_selected: bool = False
    view_mode: str = ""


@dataclass(frozen=True)
class ColumnLayout:
    """A per-column-range layout override (inclusive ``[first, last]``).

    ``has_width`` is logical presence: it is true for the raw source bit or
    a legacy non-zero width. An explicit zero remains distinguishable from an
    absent width.
    """

    first: int
    last: int
    width: float
    hidden: bool
    outline_level: int
    has_width: bool = False
    has_style: bool = False
    style_xf: int = 0


@dataclass(frozen=True)
class RowLayout:
    """A per-row layout override."""

    row: int
    height: float
    hidden: bool
    outline_level: int
    has_style: bool = False
    style_xf: int = 0


class IterativeSettings(NamedTuple):
    """Iterative-calculation settings read back by ``get_iterative``."""

    enabled: bool
    max_iterations: int
    max_change: float


@dataclass(frozen=True)
class PaginationResult:
    """Resolved physical pagination for one worksheet.

    ``print_area`` is the sheet's declared ``_xlnm.Print_Area`` as
    inclusive, zero-based ``(first_row, first_col, last_row, last_col)``
    rectangles. It is empty when the sheet declares no print area and is
    not backfilled with the used range; pagination itself still falls back
    to the used range, so ``page_count`` can be non-zero while this is
    empty. Break lists contain the zero-based row or column a new physical
    page begins before.
    """

    page_count: int
    print_area: List[tuple[int, int, int, int]]
    horizontal_breaks: List[int]
    vertical_breaks: List[int]


@dataclass(frozen=True)
class PageBreak:
    """One manual page break as the worksheet declares it.

    ``id`` is the zero-based row or column the break precedes; ``min`` /
    ``max`` bound the break on the perpendicular axis. ``manual`` is False
    for a break Excel computed and recorded rather than one a user placed.
    """

    id: int
    min: int
    max: int
    manual: bool


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

    orientation: int
    paper_size: int
    scale: int
    fit_to_width: int
    fit_to_height: int
    fit_to_page: bool
    orientation_stated: bool
    paper_size_stated: bool
    scale_stated: bool
    fit_to_width_stated: bool
    fit_to_height_stated: bool
    fit_to_page_stated: bool


@dataclass(frozen=True)
class PageMargins:
    """Effective ``<pageMargins>`` values in inches, with presence flags."""

    left: float
    right: float
    top: float
    bottom: float
    header: float
    footer: float
    left_stated: bool
    right_stated: bool
    top_stated: bool
    bottom_stated: bool
    header_stated: bool
    footer_stated: bool


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
    """``<dataBar>`` payload of a conditional-format rule.

    The fields after ``max_length_pct`` live in the ``x14`` extension
    rather than the legacy element. ``None`` means "do not override the
    model default" -- gradient fill on, automatic axis, negative fill
    equal to ``fill``, no border, and a black axis. Reading a rule back
    always populates them, so ``add_conditional_format`` of a value
    returned by ``get_conditional_format_at`` reproduces the rule.
    """

    minimum: CfValueObject
    maximum: CfValueObject
    fill: CfColor
    show_value: bool = True
    min_length_pct: int = 10
    max_length_pct: int = 90
    gradient: Optional[bool] = None
    axis_position: Optional[int] = None
    negative_fill: Optional[CfColor] = None
    border: Optional[CfColor] = None
    negative_border: Optional[CfColor] = None
    axis_color: Optional[CfColor] = None


@dataclass(frozen=True)
class IconSet:
    """``<iconSet>`` payload of a conditional-format rule.

    ``percent`` is round-trip only: it is preserved across load and save
    but never consulted during evaluation. Each threshold in
    ``thresholds`` carries its own ``type``, and that type is what
    interprets it.
    """

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
class ColorSpec:
    """How an OOXML ``<color>`` element expressed its value.

    Records read from a file carry the original specification here. When
    ``kind`` is non-zero, it is authoritative; the owning record's
    ``*_argb`` field is not an Excel-rendered value for theme/indexed/auto.
    It is literal RGB for ``kind=1`` or a compatibility fallback otherwise.
    ``kind`` is ``0=none``, ``1=rgb``, ``2=theme``, ``3=indexed``, ``4=auto``;
    the default ``kind=0`` makes the writer emit ``*_argb`` as ``rgb``.
    """

    kind: int = 0
    rgb: int = 0
    theme: int = 0
    tint: float = 0.0
    indexed: int = 0


@dataclass
class PhoneticRun:
    """One ``<rPh>`` block: ``text`` reads the span ``[sb, eb)``.

    Offsets are UTF-16 code units, which is how Excel indexes string
    positions. A whole-cell reading is the single run
    ``PhoneticRun(0, <length of the cell text>, kana)``.
    """

    sb: int = 0
    eb: int = 0
    text: str = ""


@dataclass
class FontRecord:
    """A font record (``add_font`` input / ``get_font`` result).

    The ``has_*`` flags distinguish an absent OOXML element from an
    explicit ``val="0"``: on a differential font an absent ``<b>`` means
    "leave the source formatting unchanged" while ``<b val="0"/>`` means
    "switch bold off". ``color`` is authoritative when its ``kind`` is
    non-zero; ``color_argb`` is literal RGB for ``kind=1`` and a
    compatibility fallback for theme/indexed/auto selectors, not a
    rendered colour.
    """

    name: str = ""
    size: float = 11.0
    color_argb: int = 0
    bold: bool = False
    italic: bool = False
    strike: bool = False
    has_bold: bool = False
    has_italic: bool = False
    has_strike: bool = False
    underline: int = 0
    #: 0=baseline, 1=superscript, 2=subscript.
    vert_align: int = 0
    has_family: bool = False
    family: int = 0
    has_charset: bool = False
    charset: int = 0
    color: ColorSpec = field(default_factory=ColorSpec)


@dataclass
class FillRecord:
    """A fill record.

    ``fg`` and ``bg`` carry the authoritative OOXML colour selectors. Their
    sibling ``*_argb`` fields are literal RGB for ``kind=1`` or compatibility
    fallbacks otherwise; ``kind=0`` writes the sibling as ``rgb``.
    """

    pattern: int = 0
    fg_argb: int = 0
    bg_argb: int = 0
    fg: ColorSpec = field(default_factory=ColorSpec)
    bg: ColorSpec = field(default_factory=ColorSpec)


@dataclass
class DifferentialFormat:
    """An optional-style-fragment record used by conditional formats.

    ``border`` uses the same dictionary shape as :meth:`Workbook.add_border`.
    A number format is engaged when ``num_fmt_id`` is not ``None``.
    ``alignment_xml`` and ``protection_xml`` carry serialized OOXML child
    fragments; an empty string means absent. Their semantic content survives
    round-tripping, while XML lexical formatting may normalize on load.
    """

    font: Optional[FontRecord] = None
    fill: Optional[FillRecord] = None
    border: Optional[Dict[str, object]] = None
    num_fmt_id: Optional[int] = None
    num_fmt_code: str = ""
    alignment_xml: str = ""
    protection_xml: str = ""


def _decode_color(ptr: int) -> ColorSpec:
    d = S.COLOR_SPEC.unpack(LIB, ptr)
    return ColorSpec(kind=d["kind"], rgb=d["rgb"], theme=d["theme"], tint=d["tint"], indexed=d["indexed"])


def _pack_color(ptr: int, spec: ColorSpec) -> None:
    S.COLOR_SPEC.pack(
        LIB,
        ptr,
        {
            "kind": _uint(spec.kind, "kind", 8),
            "rgb": _uint(spec.rgb, "rgb"),
            "theme": _uint(spec.theme, "theme"),
            "tint": float(spec.tint),
            "indexed": _uint(spec.indexed, "indexed"),
        },
    )


def _decode_font(ptr: int) -> FontRecord:
    d = S.FONT_RECORD.unpack(LIB, ptr)
    return FontRecord(
        name=LIB.read_cstr(d["name"]),
        size=d["size"],
        color_argb=d["color_argb"],
        bold=bool(d["bold"]),
        italic=bool(d["italic"]),
        strike=bool(d["strike"]),
        has_bold=bool(d["has_bold"]),
        has_italic=bool(d["has_italic"]),
        has_strike=bool(d["has_strike"]),
        underline=d["underline"],
        vert_align=d["vert_align"],
        has_family=bool(d["has_family"]),
        family=d["family"],
        has_charset=bool(d["has_charset"]),
        charset=d["charset"],
        color=_decode_color(ptr + S.FONT_RECORD.offsets["color"][1]),
    )


def _pack_font(ptr: int, record: FontRecord, owned: List[int]) -> None:
    S.FONT_RECORD.pack(
        LIB,
        ptr,
        {
            "size": float(record.size),
            "color_argb": _uint(record.color_argb, "color_argb"),
            "bold": 1 if record.bold else 0,
            "italic": 1 if record.italic else 0,
            "strike": 1 if record.strike else 0,
            "has_bold": 1 if record.has_bold else 0,
            "has_italic": 1 if record.has_italic else 0,
            "has_strike": 1 if record.has_strike else 0,
            "underline": _uint(record.underline, "underline", 8),
            "vert_align": _uint(record.vert_align, "vert_align", 8),
            "has_family": 1 if record.has_family else 0,
            "family": _uint(record.family, "family", 8),
            "has_charset": 1 if record.has_charset else 0,
            "charset": _uint(record.charset, "charset", 8),
        },
    )
    _pack_color(ptr + S.FONT_RECORD.offsets["color"][1], record.color)
    S.write_str_field(LIB, ptr, S.FONT_RECORD, "name", record.name, owned)


def _decode_fill(ptr: int) -> FillRecord:
    d = S.FILL_RECORD.unpack(LIB, ptr)
    return FillRecord(
        pattern=d["pattern"],
        fg_argb=d["fg_argb"],
        bg_argb=d["bg_argb"],
        fg=_decode_color(ptr + S.FILL_RECORD.offsets["fg"][1]),
        bg=_decode_color(ptr + S.FILL_RECORD.offsets["bg"][1]),
    )


def _pack_fill(ptr: int, record: FillRecord) -> None:
    S.FILL_RECORD.pack(
        LIB,
        ptr,
        {
            "pattern": _uint(record.pattern, "pattern", 8),
            "fg_argb": _uint(record.fg_argb, "fg_argb"),
            "bg_argb": _uint(record.bg_argb, "bg_argb"),
        },
    )
    _pack_color(ptr + S.FILL_RECORD.offsets["fg"][1], record.fg)
    _pack_color(ptr + S.FILL_RECORD.offsets["bg"][1], record.bg)


def _decode_border_record(ptr: int) -> Dict[str, object]:
    sides: Dict[str, object] = {}
    for side in ("left", "right", "top", "bottom", "diagonal"):
        side_ptr = ptr + S.BORDER_RECORD.offsets[side][1]
        d = S.BORDER_SIDE.unpack(LIB, side_ptr)
        sides[side] = {
            "style": d["style"],
            "color_argb": d["color_argb"],
            "color": _decode_color(side_ptr + S.BORDER_SIDE.offsets["color"][1]),
        }
    outer = S.BORDER_RECORD.unpack(LIB, ptr)
    sides["diagonal_up"] = bool(outer["diagonal_up"])
    sides["diagonal_down"] = bool(outer["diagonal_down"])
    return sides


def _pack_border_record(ptr: int, sides: Dict[str, object]) -> None:
    S.BORDER_RECORD.pack(
        LIB,
        ptr,
        {
            "diagonal_up": 1 if sides.get("diagonal_up") else 0,
            "diagonal_down": 1 if sides.get("diagonal_down") else 0,
        },
    )
    for side in ("left", "right", "top", "bottom", "diagonal"):
        spec = sides.get(side) or {}
        side_ptr = ptr + S.BORDER_RECORD.offsets[side][1]
        S.BORDER_SIDE.pack(
            LIB,
            side_ptr,
            {
                "style": _uint(spec.get("style", 0), "style", 8),
                "color_argb": _uint(spec.get("color_argb", 0), "color_argb"),
            },
        )
        _pack_color(side_ptr + S.BORDER_SIDE.offsets["color"][1], spec.get("color") or ColorSpec())


class StyleBatchIndices(NamedTuple):
    """Indices assigned by :meth:`Workbook.add_batch`, per style table.

    Each list is positionally aligned with the request list it came from
    and is empty when that list was omitted.
    """

    fonts: List[int]
    fills: List[int]
    borders: List[int]
    cell_xfs: List[int]
    num_fmts: List[int]


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
    kind: ExternalLinkKind


@dataclass(frozen=True)
class PivotCell:
    """One projected PivotTable layout cell."""

    row: int
    col: int
    value: Value
    kind: PivotCellKind
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
    axis: "PivotAxis | int" = PivotAxis.ROW
    subtotal_top: bool = False
    number_format: str = ""


@dataclass
class PivotDataFieldSpec:
    """Argument shape for ``pivot_data_field_add`` / ``pivot_data_field_set``.

    ``show_as_base_field`` / ``show_as_base_item`` use ``-1`` for "unset".
    """

    name: str
    field_index: int
    aggregation: "PivotAggregation | int" = PivotAggregation.SUM
    number_format: str = ""
    show_as: "PivotShowValuesAs | int" = PivotShowValuesAs.NORMAL
    show_as_base_field: int = -1
    show_as_base_item: int = -1


@dataclass
class PivotFilterSpec:
    """Argument shape for :meth:`Workbook.pivot_filter_add`."""

    axis: "PivotAxis | int"
    field_name: str
    type: "PivotFilterType | int"
    value_kind: "PivotFilterValueKind | int" = PivotFilterValueKind.NONE
    value_int: int = 0
    value_double: float = 0.0
    value_text: str = ""
    value_high_kind: "PivotFilterValueKind | int" = PivotFilterValueKind.NONE
    value_high_int: int = 0
    value_high_double: float = 0.0
    data_field_index: int = 0


def _cell_xf_fields(record: CellXf) -> Dict[str, object]:
    """Project a :class:`CellXf` onto the ``fm_cell_xf`` field map.

    Shared by the direct ``<xf>`` writer and the named-style ``<xf>``
    writer, which pack the same struct, so the presence-flag defaulting
    lives in one place.
    """
    return {
        "font_index": _uint(record.font_index, "font_index"),
        "fill_index": _uint(record.fill_index, "fill_index"),
        "border_index": _uint(record.border_index, "border_index"),
        "num_fmt_id": _uint(record.num_fmt_id, "num_fmt_id", 16),
        "horizontal_align": _uint(record.horizontal_align, "horizontal_align", 8),
        "vertical_align": _uint(record.vertical_align, "vertical_align", 8),
        "wrap_text": 1 if record.wrap_text else 0,
        "has_alignment": 1
        if (record.has_alignment if record.has_alignment is not None else _cell_xf_has_alignment(record))
        else 0,
        "justify_last_line": 1 if record.justify_last_line else 0,
        "xf_id": _uint(record.xf_id, "xf_id"),
        "has_text_rotation": 1 if record.text_rotation is not None else 0,
        "text_rotation": _uint(record.text_rotation or 0, "text_rotation"),
        "has_indent": 1 if record.indent is not None else 0,
        "indent": _uint(record.indent or 0, "indent"),
        "has_relative_indent": 1 if record.relative_indent is not None else 0,
        "relative_indent": _sint(record.relative_indent or 0, "relative_indent"),
        "has_shrink_to_fit": 1 if record.shrink_to_fit is not None else 0,
        "shrink_to_fit": 1 if record.shrink_to_fit else 0,
        "has_reading_order": 1 if record.reading_order is not None else 0,
        "reading_order": _uint(record.reading_order or 0, "reading_order"),
        "has_horizontal_align": 1
        if (record.has_horizontal_align if record.has_horizontal_align is not None else record.horizontal_align != 0)
        else 0,
        "has_vertical_align": 1
        if (record.has_vertical_align if record.has_vertical_align is not None else record.vertical_align != 2)
        else 0,
        "has_wrap_text": 1 if (record.has_wrap_text if record.has_wrap_text is not None else record.wrap_text) else 0,
        "has_justify_last_line": 1
        if (record.has_justify_last_line if record.has_justify_last_line is not None else record.justify_last_line)
        else 0,
    }


# ---------------------------------------------------------------------------
# Workbook handle
# ---------------------------------------------------------------------------


def _tristate(flag: Optional[bool]) -> int:
    """Map an optional flag to the ABI's preserve / clear / set encoding.

    ``None`` becomes ``-1`` (leave the stored flag alone); ``False`` and
    ``True`` become ``0`` and ``1``.
    """
    if flag is None:
        return -1
    return 1 if flag else 0


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
                "first_row": _uint(r.first_row, "first_row"),
                "first_col": _uint(r.first_col, "first_col"),
                "last_row": _uint(r.last_row, "last_row"),
                "last_col": _uint(r.last_col, "last_col"),
            },
        )
    return ptr


def _pack_phonetic_run_array(runs: Sequence[PhoneticRun], owned: List[int]) -> int:
    """Pack a list of ``PhoneticRun`` into a contiguous WASM array.

    Returns the array pointer (0 when empty). Every buffer allocated here --
    the array and each run's kana string -- is appended to ``owned`` for
    later release.
    """
    if not runs:
        return 0
    size = S.PHONETIC_RUN.size
    ptr = LIB.alloc(size * len(runs))
    owned.append(ptr)
    for i, run in enumerate(runs):
        slot = ptr + i * size
        S.PHONETIC_RUN.pack(LIB, slot, {"sb": _uint(run.sb, "sb"), "eb": _uint(run.eb, "eb")})
        S.write_str_field(LIB, slot, S.PHONETIC_RUN, "text", run.text, owned)
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
      state -- it just blocks. This no-pthread wheel deliberately exports
      only serial recalc; it has no public parallel-recalc entry point.
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

        The style table is seeded with Excel's reserved defaults exactly as
        :meth:`create_default` seeds it, so the first index :meth:`add_font`
        / :meth:`add_fill` / :meth:`add_border` hands back is non-zero.
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
            status = LIB.fm_workbook_load(data_ptr, _uint(len(buf), "len"), out_ptr)
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

    def memory_usage(self) -> int:
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
        out = _alloc_out_ptr()
        try:
            _check(LIB.fm_workbook_memory_usage(h, out), "fm_workbook_memory_usage")
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    def _require(self) -> int:
        if not self.is_valid:
            raise FormulonError(
                _STATUS_BINDING_INVALID_HANDLE,
                op="Workbook",
                _diagnostic_override=("handle is NULL or already closed", ""),
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
            status = LIB.fm_workbook_sheet_name(h, _uint(index, "index"), out_ptr)
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
        """Store the number ``value`` at ``(sheet, row, col)``.

        ``sheet``, ``row`` and ``col`` are all 0-based.
        """
        h = self._require()
        status = LIB.fm_workbook_set_number(
            h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), float(value)
        )
        _check(status, "fm_workbook_set_number")

    def set_bool(self, sheet: int, row: int, col: int, value: bool) -> None:
        """Store the boolean ``value`` at ``(sheet, row, col)``.

        ``sheet``, ``row`` and ``col`` are all 0-based.
        """
        h = self._require()
        status = LIB.fm_workbook_set_bool(
            h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), 1 if value else 0
        )
        _check(status, "fm_workbook_set_bool")

    def set_error(self, sheet: int, row: int, col: int, error_code: int) -> None:
        """Store an Excel cell error at ``(sheet, row, col)``.

        ``sheet``, ``row`` and ``col`` are all 0-based. ``error_code`` is
        a ``formulon::ErrorCode`` ordinal, not a display literal: ``0`` is
        ``#NULL!``, ``1`` ``#DIV/0!``, ``2`` ``#VALUE!``, ``3`` ``#REF!``,
        ``4`` ``#NAME?``, ``5`` ``#NUM!``, ``6`` ``#N/A``. The same
        ordinals come back out through :attr:`Value.error_code`.
        """
        h = self._require()
        status = LIB.fm_workbook_set_error(
            h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), _sint(error_code, "error")
        )
        _check(status, "fm_workbook_set_error")

    def set_text(self, sheet: int, row: int, col: int, value: str) -> None:
        """Store the string ``value`` at ``(sheet, row, col)``.

        ``sheet``, ``row`` and ``col`` are all 0-based. The text is stored
        verbatim; a leading ``=`` does not make it a formula (use
        :meth:`set_formula`).
        """
        h = self._require()
        text_ptr, _ = LIB.alloc_utf8(value)
        try:
            status = LIB.fm_workbook_set_text(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), text_ptr
            )
            _check(status, "fm_workbook_set_text")
        finally:
            LIB.free(text_ptr)

    def set_blank(self, sheet: int, row: int, col: int) -> None:
        """Clear the cell at ``(sheet, row, col)``.

        ``sheet``, ``row`` and ``col`` are all 0-based. Clearing drops the
        stored value or formula; cell formatting is unaffected.
        """
        h = self._require()
        status = LIB.fm_workbook_set_blank(h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"))
        _check(status, "fm_workbook_set_blank")

    def set_formula(self, sheet: int, row: int, col: int, formula: str) -> None:
        """Store the formula ``formula`` at ``(sheet, row, col)``.

        ``sheet``, ``row`` and ``col`` are all 0-based. The leading ``=``
        is optional. The cell keeps its previous cached value until the
        next :meth:`recalc`.
        """
        h = self._require()
        formula_ptr, _ = LIB.alloc_utf8(formula)
        try:
            status = LIB.fm_workbook_set_formula(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), formula_ptr
            )
            _check(status, "fm_workbook_set_formula")
        finally:
            LIB.free(formula_ptr)

    def set_phonetic(self, sheet: int, row: int, col: int, text: str) -> None:
        """Set the cell's phonetic guide (OOXML ``<rPh>``, furigana).

        ``sheet``, ``row`` and ``col`` are all 0-based. An empty ``text``
        removes the guide. The guide is display metadata attached to the
        cell's string; it does not affect evaluation.
        """
        h = self._require()
        text_ptr, _ = LIB.alloc_utf8(text)
        try:
            status = LIB.fm_workbook_set_cell_phonetic(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), text_ptr
            )
            _check(status, "fm_workbook_set_cell_phonetic")
        finally:
            LIB.free(text_ptr)

    def set_phonetic_runs(self, sheet: int, row: int, col: int, runs: Sequence[PhoneticRun]) -> None:
        """Set the cell's phonetic guide as one ``<rPh>`` block per run.

        ``sheet``, ``row`` and ``col`` are all 0-based. An empty ``runs``
        removes the guide. Unlike :meth:`set_phonetic`, which annotates the
        whole cell, this keeps the spans each reading covers -- the
        difference ``PHONETIC`` and a partially annotated cell depend on.

        The runs must be an ordered partition: each needs ``sb <= eb`` and
        must start at or after the previous run's ``eb``.
        """
        h = self._require()
        owned: List[int] = []
        ptr = _pack_phonetic_run_array(runs, owned)
        try:
            status = LIB.fm_workbook_set_cell_phonetic_runs(
                h,
                _uint(sheet, "sheet_index"),
                _uint(row, "row"),
                _uint(col, "col"),
                ptr,
                _uint(len(runs), "count"),
            )
            _check(status, "fm_workbook_set_cell_phonetic_runs")
        finally:
            for p in owned:
                LIB.free(p)

    def get_phonetic_runs(self, sheet: int, row: int, col: int) -> List[PhoneticRun]:
        """Return the cell's ``<rPh>`` blocks, spans included.

        ``sheet``, ``row`` and ``col`` are all 0-based. An unannotated cell
        yields an empty list. :meth:`get_phonetic` returns the same readings
        concatenated, without the spans.
        """
        h = self._require()
        count_out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_get_cell_phonetic_run_count(
                    h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), count_out
                ),
                "fm_workbook_get_cell_phonetic_run_count",
            )
            count = LIB.read_u32(count_out)
        finally:
            LIB.free(count_out)
        out: List[PhoneticRun] = []
        run_ptr = S.alloc_struct(LIB, S.PHONETIC_RUN)
        try:
            for i in range(count):
                S.zero_struct(LIB, S.PHONETIC_RUN, run_ptr)
                _check(
                    LIB.fm_workbook_get_cell_phonetic_run(
                        h,
                        _uint(sheet, "sheet_index"),
                        _uint(row, "row"),
                        _uint(col, "col"),
                        _uint(i, "run_index"),
                        run_ptr,
                    ),
                    "fm_workbook_get_cell_phonetic_run",
                )
                fields = S.PHONETIC_RUN.unpack(LIB, run_ptr)
                # Decoded before the next read: the text points into the
                # handle's scratch, which every read refreshes.
                out.append(
                    PhoneticRun(
                        sb=fields["sb"],
                        eb=fields["eb"],
                        text=LIB.read_cstr(fields["text"]) if fields["text"] else "",
                    )
                )
        finally:
            LIB.free(run_ptr)
        return out

    def get_phonetic(self, sheet: int, row: int, col: int) -> str:
        """Return the cell's phonetic guide, or ``""`` when it has none.

        ``sheet``, ``row`` and ``col`` are all 0-based.
        """
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_get_cell_phonetic(
                    h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), out
                ),
                "fm_workbook_get_cell_phonetic",
            )
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(out)

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
            status = LIB.fm_workbook_get_value(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), value_ptr
            )
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
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(row, "row"),
                    _uint(col, "col"),
                    _uint(anchor_row, "anchor_row"),
                    _uint(anchor_col, "anchor_col"),
                    formula_ptr,
                    value_ptr,
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
            status = LIB.fm_workbook_evaluate_formula_array(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), formula_ptr, rows_ptr, cols_ptr
            )
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
                    status = LIB.fm_workbook_evaluate_formula_array_cell(h, _uint(index, "index"), value_ptr)
                    _check(status, "fm_workbook_evaluate_formula_array_cell")
                    out_row.append(Value._from_wasm(value_ptr))
                grid.append(out_row)
        finally:
            LIB.free(value_ptr)
        return grid

    # -- Recalc ------------------------------------------------------------
    def recalc(self) -> None:
        """Drive a full incremental recalc.

        This no-pthread WASM wheel intentionally exposes only the serial
        recalc contract. Thread-capable native surfaces may opt in to the
        parallel scheduler, but this method always runs serially.
        """
        h = self._require()
        status = LIB.fm_workbook_recalc(h)
        _check(status, "fm_workbook_recalc")

    def set_iterative(self, enabled: bool, max_iterations: int, max_change: float) -> None:
        """Configure iterative calculation."""
        h = self._require()
        status = LIB.fm_workbook_set_iterative(
            h, 1 if enabled else 0, _sint(max_iterations, "max_iterations"), float(max_change)
        )
        _check(status, "fm_workbook_set_iterative")

    def set_iterative_enabled(self, enabled: bool) -> None:
        """Toggle iterative calculation, keeping the cap and threshold.

        The counterpart of Excel's enable checkbox, which is separate from
        its advanced iterative-calculation settings; :meth:`set_iterative`
        writes all three at once.
        """
        h = self._require()
        _check(
            LIB.fm_workbook_set_iterative_enabled(h, 1 if enabled else 0),
            "fm_workbook_set_iterative_enabled",
        )

    def get_iterative(self) -> IterativeSettings:
        """Read back the iterative-calculation settings.

        The counterpart of :meth:`set_iterative`; the cap and threshold are
        the stored values even when ``enabled`` is ``False``.
        """
        h = self._require()
        enabled_ptr = _alloc_out_ptr()
        iterations_ptr = _alloc_out_ptr()
        change_ptr = LIB.alloc(8)
        try:
            LIB.write_bytes(change_ptr, b"\x00" * 8)
            _check(
                LIB.fm_workbook_get_iterative(h, enabled_ptr, iterations_ptr, change_ptr),
                "fm_workbook_get_iterative",
            )
            return IterativeSettings(
                enabled=bool(LIB.read_i32(enabled_ptr)),
                max_iterations=LIB.read_u32(iterations_ptr),
                max_change=LIB.read_f64(change_ptr),
            )
        finally:
            LIB.free(enabled_ptr)
            LIB.free(iterations_ptr)
            LIB.free(change_ptr)

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

    def save_as(self, fmt: "WorkbookFormat | int") -> bytes:
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
            status = LIB.fm_workbook_save_as(h, _sint(fmt, "format"), out_ptr_ptr, out_len_ptr)
            _check(status, "fm_workbook_save_as")
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

    def save_with_diagnostics(self, fmt: "WorkbookFormat | int") -> SaveDiagnostics:
        """Serialise the workbook and return what the save cost.

        The returned bytes are an independent copy and the underlying
        WASM allocation is released before this method returns.
        """
        h = self._require()
        scratch: list[int] = []
        try:
            for _ in range(2):
                ptr = LIB.alloc(4)
                scratch.append(ptr)
                LIB.write_bytes(ptr, b"\x00\x00\x00\x00")
            out_ptr_ptr, out_len_ptr = scratch
            diag_ptr = S.alloc_struct(LIB, S.SAVE_DIAGNOSTICS)
            scratch.append(diag_ptr)
            status = LIB.fm_workbook_save_with_diagnostics(h, _sint(fmt, "format"), out_ptr_ptr, out_len_ptr, diag_ptr)
            _check(status, "fm_workbook_save_with_diagnostics")
            data_ptr = LIB.read_u32(out_ptr_ptr)
            data_len = LIB.read_u32(out_len_ptr)
            try:
                data = b"" if data_len == 0 or data_ptr == 0 else LIB.read_bytes(data_ptr, data_len)
            finally:
                if data_ptr:
                    LIB.fm_buffer_free(data_ptr)
            d = S.SAVE_DIAGNOSTICS.unpack(LIB, diag_ptr)
            return SaveDiagnostics(
                bytes=data,
                downgraded_formula_count=d["downgraded_formula_count"],
                deferred_feature_count=d["deferred_feature_count"],
                dropped_part_count=d["dropped_part_count"],
                dropped_relationship_count=d["dropped_relationship_count"],
                renumbered_part_count=d["renumbered_part_count"],
            )
        finally:
            for ptr in scratch:
                LIB.free(ptr)

    def read_diagnostics(self) -> ReadDiagnostics:
        """Return the loss counters captured when this workbook was loaded."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.READ_DIAGNOSTICS)
        try:
            _check(LIB.fm_workbook_read_diagnostics(h, ptr), "fm_workbook_read_diagnostics")
            d = S.READ_DIAGNOSTICS.unpack(LIB, ptr)
            return ReadDiagnostics(
                undecoded_formula_count=d["undecoded_formula_count"],
                undecoded_defined_name_count=d["undecoded_defined_name_count"],
                undecoded_part_count=d["undecoded_part_count"],
                skipped_feature_count=d["skipped_feature_count"],
                unknown_content_type_count=d["unknown_content_type_count"],
            )
        finally:
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
            status = LIB.fm_workbook_cell_count(h, _uint(sheet, "sheet_index"), out_count)
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
                status = LIB.fm_workbook_cell_at(
                    h, _uint(sheet, "sheet_index"), _uint(i, "idx"), row_ptr, col_ptr, formula_ptr, value_ptr
                )
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
        # Scratch is hoisted out of the loop (as in iter_cells): each
        # LIB.alloc/free is a locked WASM call, so per-item slots would
        # make the walk cost O(n) allocator round-trips on top of the
        # unavoidable O(n) ABI reads.
        name_ptr = _alloc_out_ptr()
        formula_ptr = _alloc_out_ptr()
        local_sheet_ptr = _alloc_out_ptr()
        try:
            for i in range(n):
                status = LIB.fm_workbook_defined_name_at(h, _uint(i, "idx"), name_ptr, formula_ptr, local_sheet_ptr)
                _check(status, "fm_workbook_defined_name_at")
                name = LIB.read_cstr(LIB.read_u32(name_ptr))
                formula = LIB.read_cstr(LIB.read_u32(formula_ptr))
                local_sheet_id = LIB.read_i32(local_sheet_ptr)
                yield DefinedName(name=name, formula=formula, local_sheet_id=local_sheet_id)
        finally:
            LIB.free(name_ptr)
            LIB.free(formula_ptr)
            LIB.free(local_sheet_ptr)

    def iter_tables(self) -> Iterator[Table]:
        """Iterate over every table in declaration order."""
        h = self._require()
        n = int(LIB.fm_workbook_table_count(h))
        # Scratch hoisted out of the loop; see iter_defined_names.
        name_ptr = _alloc_out_ptr()
        display_ptr = _alloc_out_ptr()
        ref_ptr = _alloc_out_ptr()
        sheet_ptr = _alloc_out_ptr()
        try:
            for i in range(n):
                status = LIB.fm_workbook_table_at(h, _uint(i, "idx"), name_ptr, display_ptr, ref_ptr, sheet_ptr)
                _check(status, "fm_workbook_table_at")
                name = LIB.read_cstr(LIB.read_u32(name_ptr))
                display = LIB.read_cstr(LIB.read_u32(display_ptr))
                ref = LIB.read_cstr(LIB.read_u32(ref_ptr))
                sheet = LIB.read_u32(sheet_ptr)
                yield Table(name=name, display_name=display, ref=ref, sheet_index=sheet)
        finally:
            LIB.free(name_ptr)
            LIB.free(display_ptr)
            LIB.free(ref_ptr)
            LIB.free(sheet_ptr)

    def table_create(
        self,
        sheet: int,
        ref: str,
        name: str,
        display_name: str,
        column_names: Sequence[str],
        style_name: str = "",
        header_row: bool = True,
        totals_row: bool = False,
    ) -> int:
        """Create a worksheet table and return its index.

        ``column_names`` must hold one non-empty, table-unique header per
        column in ``ref``, and its length must equal the width of ``ref``.
        An empty ``style_name`` omits ``<tableStyleInfo>``. With
        ``header_row`` set, the caller still owns the header cells: Excel
        expects the first row of ``ref`` to hold exactly these names.
        """
        h = self._require()
        owned: List[int] = []
        try:
            ref_ptr = _opt_str_ptr(ref, owned)
            name_ptr = _opt_str_ptr(name, owned)
            display_ptr = _opt_str_ptr(display_name, owned)
            style_ptr = _opt_str_ptr(style_name, owned)
            names = [str(column) for column in column_names]
            name_ptrs = [_opt_str_ptr(column, owned) for column in names]
            array_ptr = LIB.alloc(4 * len(name_ptrs)) if name_ptrs else 0
            if array_ptr:
                owned.append(array_ptr)
                LIB.write_bytes(array_ptr, b"".join(struct.pack("<I", p) for p in name_ptrs))
            out = _alloc_out_ptr()
            owned.append(out)
            _check(
                LIB.fm_workbook_table_create(
                    h,
                    _uint(sheet, "sheet_index"),
                    ref_ptr,
                    name_ptr,
                    display_ptr,
                    array_ptr,
                    _uint(len(name_ptrs), "column_count"),
                    style_ptr,
                    1 if header_row else 0,
                    1 if totals_row else 0,
                    out,
                ),
                "fm_workbook_table_create",
            )
            return LIB.read_u32(out)
        finally:
            for ptr in owned:
                LIB.free(ptr)

    def table_update(
        self,
        index: int,
        ref: str,
        style_name: Optional[str] = None,
        header_row: Optional[bool] = None,
        totals_row: Optional[bool] = None,
    ) -> None:
        """Replace a table's range and optional visual style.

        ``ref`` is required and must keep the width the table's column
        list already describes. ``style_name=None`` preserves the existing
        style payload; an empty string removes ``<tableStyleInfo>``.
        ``header_row=None`` / ``totals_row=None`` preserve those flags.
        """
        h = self._require()
        owned: List[int] = []
        try:
            ref_ptr = _opt_str_ptr(ref, owned)
            style_ptr = 0 if style_name is None else _opt_str_ptr(style_name, owned)
            _check(
                LIB.fm_workbook_table_update(
                    h,
                    _uint(index, "index"),
                    ref_ptr,
                    style_ptr,
                    _sint(_tristate(header_row), "header_row"),
                    _sint(_tristate(totals_row), "totals_row"),
                ),
                "fm_workbook_table_update",
            )
        finally:
            for ptr in owned:
                LIB.free(ptr)

    def table_remove(self, index: int) -> None:
        """Remove the table at ``index``."""
        h = self._require()
        _check(LIB.fm_workbook_table_remove(h, _uint(index, "index")), "fm_workbook_table_remove")

    def iter_passthrough(self) -> Iterator[PassthroughPart]:
        """Iterate over passthrough OOXML part paths."""
        h = self._require()
        n = int(LIB.fm_workbook_passthrough_count(h))
        # Scratch hoisted out of the loop; see iter_defined_names.
        path_ptr = _alloc_out_ptr()
        try:
            for i in range(n):
                status = LIB.fm_workbook_passthrough_at(h, _uint(i, "idx"), path_ptr)
                _check(status, "fm_workbook_passthrough_at")
                path = LIB.read_cstr(LIB.read_u32(path_ptr))
                yield PassthroughPart(path=path)
        finally:
            LIB.free(path_ptr)

    # -- Sheet structure ---------------------------------------------------
    def move_sheet(self, from_index: int, to_index: int) -> None:
        """Move the sheet at ``from_index`` to ``to_index`` (post-removal)."""
        h = self._require()
        _check(
            LIB.fm_workbook_move_sheet(h, _uint(from_index, "from_index"), _uint(to_index, "to_index")),
            "fm_workbook_move_sheet",
        )

    def remove_sheet(self, index: int) -> None:
        """Remove the sheet at ``index``."""
        h = self._require()
        _check(LIB.fm_workbook_remove_sheet(h, _uint(index, "index")), "fm_workbook_remove_sheet")

    def rename_sheet(self, index: int, new_name: str) -> None:
        """Rename the sheet at ``index`` to ``new_name``."""
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(new_name)
        try:
            _check(
                LIB.fm_workbook_rename_sheet(h, _uint(index, "index"), name_ptr),
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
                LIB.fm_workbook_set_defined_name_scoped(
                    h, name_ptr, formula_ptr, _sint(local_sheet_id, "local_sheet_id")
                ),
                "fm_workbook_set_defined_name_scoped",
            )
        finally:
            LIB.free(name_ptr)
            LIB.free(formula_ptr)

    # -- Row / column structural edits -------------------------------------
    #
    # The four edits below move every structure the engine models: cells,
    # formulas, merges, conditional-format ranges, validations, hyperlinks,
    # tables, print areas, manual breaks and the auto-filter range.
    #
    # They do not move coordinates held in worksheet content the engine
    # keeps byte-verbatim because it does not model it -- the worksheet
    # ``<extLst>`` (the ``x14`` conditional-formatting block behind DataBar
    # negative-fill / axis / gradient settings, sparkline groups, slicer
    # anchors) and any unmodelled ``<worksheet>`` child kept so a save does
    # not drop it. Those keep their pre-edit rectangles.
    #
    # What that looks like in Excel differs per extension: it drops an
    # ``x14`` entry whose range no longer matches the legacy rule it
    # extends, so an extended DataBar quietly reverts to its legacy
    # rendering, while a sparkline group keeps drawing and reads its source
    # range from the wrong cells. Neither raises an exception or a
    # diagnostic counter. Re-author those extensions after editing rows or
    # columns on a sheet that carries them.
    def insert_rows(self, sheet: int, row: int, count: int) -> None:
        """Insert ``count`` rows at ``row`` on ``sheet``.

        Coordinates inside verbatim-retained worksheet extensions
        (``<extLst>``, unmodelled ``<worksheet>`` children) are not
        remapped; see the note above this method group.
        """
        h = self._require()
        _check(
            LIB.fm_workbook_insert_rows(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(count, "count")),
            "fm_workbook_insert_rows",
        )

    def delete_rows(self, sheet: int, row: int, count: int) -> None:
        """Delete ``count`` rows starting at ``row`` on ``sheet``.

        Coordinates inside verbatim-retained worksheet extensions are not
        remapped; see the note above this method group.
        """
        h = self._require()
        _check(
            LIB.fm_workbook_delete_rows(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(count, "count")),
            "fm_workbook_delete_rows",
        )

    def insert_cols(self, sheet: int, col: int, count: int) -> None:
        """Insert ``count`` columns at ``col`` on ``sheet``.

        Coordinates inside verbatim-retained worksheet extensions are not
        remapped; see the note above this method group.
        """
        h = self._require()
        _check(
            LIB.fm_workbook_insert_cols(h, _uint(sheet, "sheet"), _uint(col, "col"), _uint(count, "count")),
            "fm_workbook_insert_cols",
        )

    def delete_cols(self, sheet: int, col: int, count: int) -> None:
        """Delete ``count`` columns starting at ``col`` on ``sheet``.

        Coordinates inside verbatim-retained worksheet extensions are not
        remapped; see the note above this method group.
        """
        h = self._require()
        _check(
            LIB.fm_workbook_delete_cols(h, _uint(sheet, "sheet"), _uint(col, "col"), _uint(count, "count")),
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

    def set_calc_mode(self, mode: "CalcMode | int") -> None:
        """Set the workbook's calc mode."""
        h = self._require()
        _check(LIB.fm_workbook_set_calc_mode(h, _sint(mode, "mode")), "fm_workbook_set_calc_mode")

    def pinned_now(self) -> "CivilTime | None":
        """Return the pinned wall-clock reading, or ``None`` when the
        workbook follows the host clock (the default)."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.CIVIL_TIME)
        flag = _alloc_out_ptr()
        try:
            _check(LIB.fm_workbook_pinned_now(h, ptr, flag), "fm_workbook_pinned_now")
            if LIB.read_i32(flag) == 0:
                return None
            d = S.CIVIL_TIME.unpack(LIB, ptr)
            return CivilTime(
                year=d["year"],
                month=d["month"],
                day=d["day"],
                hour=d["hour"],
                minute=d["minute"],
                second=d["second"],
            )
        finally:
            LIB.free(flag)
            LIB.free(ptr)

    def set_pinned_now(
        self,
        year: int,
        month: int,
        day: int,
        hour: int = 0,
        minute: int = 0,
        second: int = 0,
    ) -> None:
        """Pin every clock-dependent result to one instant.

        ``NOW()``, ``TODAY()`` and the pivot relative-period filters ("this
        month", "year to date", ...) otherwise each read the clock
        independently, which makes a recalc internally inconsistent across a
        midnight boundary and makes any such result untestable.

        This is model state, not file state: nothing in the OOXML package
        records it, so :meth:`save` drops it. Existing formula cells keep
        their cached values until the next :meth:`recalc`.

        Raises:
          FormulonError: with ``kInvalidArgument`` unless ``year`` is in
            ``[1900, 9999]``, ``month`` in ``[1, 12]``, ``day`` within that
            month's real length, ``hour`` in ``[0, 23]``, and ``minute`` /
            ``second`` in ``[0, 59]``. The pin is a calendar instant, not a
            normalising constructor: a month of 13 is rejected rather than
            rolled into the next year.
        """
        h = self._require()
        ptr = S.alloc_struct(LIB, S.CIVIL_TIME)
        try:
            S.CIVIL_TIME.pack(
                LIB,
                ptr,
                {
                    "year": _sint(year, "year"),
                    "month": _sint(month, "month"),
                    "day": _sint(day, "day"),
                    "hour": _sint(hour, "hour"),
                    "minute": _sint(minute, "minute"),
                    "second": _sint(second, "second"),
                },
            )
            _check(LIB.fm_workbook_set_pinned_now(h, ptr), "fm_workbook_set_pinned_now")
        finally:
            LIB.free(ptr)

    def clear_pinned_now(self) -> None:
        """Release the pin so clock-dependent results follow the host clock
        again. Clearing an unpinned workbook does nothing."""
        h = self._require()
        _check(LIB.fm_workbook_clear_pinned_now(h), "fm_workbook_clear_pinned_now")

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
                    "sheet": _uint(sheet, "sheet"),
                    "first_row": _uint(first_row, "first_row"),
                    "last_row": _uint(last_row, "last_row"),
                    "first_col": _uint(first_col, "first_col"),
                    "last_col": _uint(last_col, "last_col"),
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
                LIB.fm_workbook_lambda_text_at(
                    h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), out
                ),
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
            _check(LIB.fm_sheet_add_merge(h, _uint(sheet, "sheet"), ptr), "fm_sheet_add_merge")
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
                LIB.fm_sheet_remove_merge(h, _uint(sheet, "sheet"), ptr),
                "fm_sheet_remove_merge",
            )
        finally:
            for p in owned:
                LIB.free(p)

    def remove_merge_at(self, sheet: int, index: int) -> None:
        """Remove the merge at ``index`` on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_sheet_remove_merge_at(h, _uint(sheet, "sheet"), _uint(index, "index")),
            "fm_sheet_remove_merge_at",
        )

    def clear_merges(self, sheet: int) -> None:
        """Drop every merge range on ``sheet``."""
        h = self._require()
        _check(LIB.fm_sheet_clear_merges(h, _uint(sheet, "sheet")), "fm_sheet_clear_merges")

    def merge_count(self, sheet: int) -> int:
        """Return the number of merge ranges on ``sheet``."""
        h = self._require()
        return _read_count(LIB.fm_sheet_get_merge_count, h, _uint(sheet, "sheet"))

    def get_merges(self, sheet: int) -> List[MergeRange]:
        """Return every merge range on ``sheet``."""
        h = self._require()
        n = self.merge_count(sheet)
        out: List[MergeRange] = []
        # One scratch block for the whole walk; see iter_defined_names.
        ptr = S.alloc_struct(LIB, S.MERGE_RANGE)
        try:
            for i in range(n):
                S.zero_struct(LIB, S.MERGE_RANGE, ptr)
                _check(
                    LIB.fm_sheet_get_merge_at(h, _uint(sheet, "sheet"), _uint(i, "index"), ptr),
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
            S.HYPERLINK.pack(
                LIB,
                ptr,
                {
                    "row": _uint(row, "row"),
                    "col": _uint(col, "col"),
                    "last_row": _uint(row, "row"),
                    "last_col": _uint(col, "col"),
                },
            )
            S.write_str_field(LIB, ptr, S.HYPERLINK, "target", target, owned)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "location", location, owned)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "display", display, owned)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "tooltip", tooltip, owned)
            _check(LIB.fm_sheet_add_hyperlink(h, _uint(sheet, "sheet"), ptr), "fm_sheet_add_hyperlink")
        finally:
            LIB.free(ptr)
            for p in owned:
                LIB.free(p)

    def add_hyperlink_range(
        self,
        sheet: int,
        row: int,
        col: int,
        last_row: int,
        last_col: int,
        target: str,
        display: str = "",
        tooltip: str = "",
        location: str = "",
    ) -> None:
        """Append a hyperlink covering the inclusive cell rectangle."""
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.HYPERLINK)
        try:
            S.HYPERLINK.pack(
                LIB,
                ptr,
                {
                    "row": _uint(row, "row"),
                    "col": _uint(col, "col"),
                    "last_row": _uint(last_row, "last_row"),
                    "last_col": _uint(last_col, "last_col"),
                },
            )
            S.write_str_field(LIB, ptr, S.HYPERLINK, "target", target, owned)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "location", location, owned)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "display", display, owned)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "tooltip", tooltip, owned)
            _check(LIB.fm_sheet_add_hyperlink(h, _uint(sheet, "sheet"), ptr), "fm_sheet_add_hyperlink")
        finally:
            LIB.free(ptr)
            for p in owned:
                LIB.free(p)

    def remove_hyperlink(self, sheet: int, row: int, col: int) -> None:
        """Remove every hyperlink anchored at ``(row, col)`` on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_sheet_remove_hyperlink(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col")),
            "fm_sheet_remove_hyperlink",
        )

    def remove_hyperlink_at(self, sheet: int, index: int) -> None:
        """Remove the hyperlink at ``index`` on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_sheet_remove_hyperlink_at(h, _uint(sheet, "sheet"), _uint(index, "index")),
            "fm_sheet_remove_hyperlink_at",
        )

    def clear_hyperlinks(self, sheet: int) -> None:
        """Drop every hyperlink on ``sheet``."""
        h = self._require()
        _check(LIB.fm_sheet_clear_hyperlinks(h, _uint(sheet, "sheet")), "fm_sheet_clear_hyperlinks")

    def hyperlink_count(self, sheet: int) -> int:
        """Return the number of hyperlinks on ``sheet``."""
        h = self._require()
        return _read_count(LIB.fm_sheet_get_hyperlink_count, h, _uint(sheet, "sheet"))

    def get_hyperlinks(self, sheet: int) -> List[Hyperlink]:
        """Return every hyperlink on ``sheet``."""
        h = self._require()
        n = self.hyperlink_count(sheet)
        out: List[Hyperlink] = []
        # One scratch block for the whole walk; see iter_defined_names.
        ptr = S.alloc_struct(LIB, S.HYPERLINK)
        try:
            for i in range(n):
                S.zero_struct(LIB, S.HYPERLINK, ptr)
                _check(
                    LIB.fm_sheet_get_hyperlink_at(h, _uint(sheet, "sheet"), _uint(i, "index"), ptr),
                    "fm_sheet_get_hyperlink_at",
                )
                d = S.HYPERLINK.unpack(LIB, ptr)
                out.append(
                    Hyperlink(
                        row=d["row"],
                        col=d["col"],
                        last_row=d["last_row"],
                        last_col=d["last_col"],
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
            status = LIB.fm_sheet_get_comment_at(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), ptr)
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
                LIB.fm_sheet_set_comment(
                    h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), author_ptr, text_ptr
                ),
                "fm_sheet_set_comment",
            )
        finally:
            for p in owned:
                LIB.free(p)

    def comment_count(self, sheet: int) -> int:
        """Return the number of comments on ``sheet``."""
        h = self._require()
        return _read_count(LIB.fm_sheet_get_comment_count, h, _uint(sheet, "sheet"))

    def get_comments(self, sheet: int) -> List[CommentEntry]:
        """Return every comment on ``sheet`` in storage order.

        Unlike :meth:`get_comment`, this discovers comments on otherwise
        empty cells without requiring callers to scan worksheet coordinates.
        """
        h = self._require()
        out: List[CommentEntry] = []
        # One scratch block for the whole walk; see iter_defined_names.
        ptr = S.alloc_struct(LIB, S.COMMENT)
        try:
            for index in range(self.comment_count(sheet)):
                S.zero_struct(LIB, S.COMMENT, ptr)
                _check(
                    LIB.fm_sheet_get_comment_at_index(h, _uint(sheet, "sheet"), _uint(index, "index"), ptr),
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
        return _read_count(LIB.fm_sheet_get_validation_count, h, _uint(sheet, "sheet"))

    def get_validation_at(self, sheet: int, index: int) -> DataValidation:
        """Read the ``index``-th data-validation rule on ``sheet``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.DATA_VALIDATION)
        try:
            _check(
                LIB.fm_sheet_get_validation_at(h, _uint(sheet, "sheet"), _uint(index, "index"), ptr),
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
                    "type": _uint(validation.type, "type", 8),
                    "op": _uint(validation.op, "op", 8),
                    "error_style": _uint(validation.error_style, "error_style", 8),
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
                LIB.fm_sheet_add_validation(h, _uint(sheet, "sheet"), ptr),
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
            LIB.fm_sheet_remove_validation_at(h, _uint(sheet, "sheet"), _uint(index, "index")),
            "fm_sheet_remove_validation_at",
        )

    def clear_validations(self, sheet: int) -> None:
        """Drop every validation rule on ``sheet``."""
        h = self._require()
        _check(
            LIB.fm_sheet_clear_validations(h, _uint(sheet, "sheet")),
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
                LIB.fm_sheet_get_protection(h, _uint(sheet, "sheet_index"), ptr),
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
                "spin_count": _uint(protection.spin_count, "spin_count"),
            }
            for flag in self._PROTECT_FLAGS:
                values[flag] = 1 if getattr(protection, flag) else 0
            S.SHEET_PROTECTION.pack(LIB, ptr, values)
            for fld in ("algorithm_name", "hash_value", "salt_value", "legacy_password"):
                S.write_str_field(LIB, ptr, S.SHEET_PROTECTION, fld, getattr(protection, fld), owned)
            _check(
                LIB.fm_sheet_set_protection(h, _uint(sheet, "sheet_index"), ptr),
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
            _check(LIB.fm_workbook_paginate(h, _uint(sheet, "sheet_index"), out), "fm_workbook_paginate")
            pagination = LIB.read_u32(out)
            if pagination == 0:
                raise FormulonError(
                    _STATUS_BINDING_NULL_POINTER,
                    op="fm_workbook_paginate",
                    _diagnostic_override=("returned kOk with a null pagination handle", ""),
                )
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
                            LIB.fm_pagination_print_area_at(pagination, _uint(i, "index"), range_ptr),
                            "fm_pagination_print_area_at",
                        )
                        print_area.append(struct.unpack("<IIII", LIB.read_bytes(range_ptr, 16)))
                    horizontal_breaks = []
                    for i in range(break_count):
                        _check(
                            LIB.fm_pagination_horizontal_break_at(pagination, _uint(i, "index"), value_ptr),
                            "fm_pagination_horizontal_break_at",
                        )
                        horizontal_breaks.append(LIB.read_u32(value_ptr))
                    vertical_breaks = []
                    for i in range(vertical_break_count):
                        _check(
                            LIB.fm_pagination_vertical_break_at(pagination, _uint(i, "index"), value_ptr),
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
        ptr = S.alloc_struct(LIB, S.SHEET_VIEW)
        try:
            _check(LIB.fm_sheet_get_view(h, _uint(sheet, "sheet_index"), ptr), "fm_sheet_get_view")
            d = S.SHEET_VIEW.unpack(LIB, ptr)
            return SheetView(
                zoom_scale=d["zoom_scale"],
                freeze_rows=d["freeze_rows"],
                freeze_cols=d["freeze_cols"],
                tab_hidden=bool(d["tab_hidden"]),
                visibility=SheetVisibility(d["visibility"]),
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
        _check(
            LIB.fm_sheet_set_zoom(h, _uint(sheet, "sheet_index"), _uint(zoom_scale, "zoom_scale")), "fm_sheet_set_zoom"
        )

    def set_sheet_freeze(self, sheet: int, freeze_rows: int, freeze_cols: int) -> None:
        """Set the frozen pane in ``(rows, cols)``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_freeze(
                h, _uint(sheet, "sheet_index"), _uint(freeze_rows, "freeze_rows"), _uint(freeze_cols, "freeze_cols")
            ),
            "fm_sheet_set_freeze",
        )

    def set_sheet_tab_hidden(self, sheet: int, hidden: bool) -> None:
        """Set the sheet tab's hidden flag.

        The two-state view of :meth:`set_sheet_visibility`. ``True`` on an
        already very-hidden sheet leaves it very-hidden rather than
        demoting it, since "hidden" says nothing that sheet does not
        already satisfy; ``False`` shows it from either hidden state.
        """
        h = self._require()
        _check(
            LIB.fm_sheet_set_tab_hidden(h, _uint(sheet, "sheet_index"), 1 if hidden else 0),
            "fm_sheet_set_tab_hidden",
        )

    def set_sheet_visibility(self, sheet: int, visibility: "SheetVisibility | int") -> None:
        """Set the sheet tab's visibility to one of the three states.

        The only way to newly state :attr:`SheetVisibility.VERY_HIDDEN`, and
        the only way to demote a very-hidden sheet to plain hidden;
        :meth:`set_sheet_tab_hidden` can express neither.
        """
        h = self._require()
        _check(
            LIB.fm_sheet_set_visibility(h, _uint(sheet, "sheet_index"), _sint(visibility, "visibility")),
            "fm_sheet_set_visibility",
        )

    def set_sheet_show_grid_lines(self, sheet: int, show: bool) -> None:
        """Set the sheet's ``showGridLines`` flag."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_show_grid_lines(h, _uint(sheet, "sheet_index"), 1 if show else 0),
            "fm_sheet_set_show_grid_lines",
        )

    def set_sheet_show_row_col_headers(self, sheet: int, show: bool) -> None:
        """Set the sheet's ``showRowColHeaders`` flag."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_show_row_col_headers(h, _uint(sheet, "sheet_index"), 1 if show else 0),
            "fm_sheet_set_show_row_col_headers",
        )

    def set_sheet_show_zeros(self, sheet: int, show: bool) -> None:
        """Set the sheet's ``showZeros`` flag."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_show_zeros(h, _uint(sheet, "sheet_index"), 1 if show else 0),
            "fm_sheet_set_show_zeros",
        )

    def set_sheet_right_to_left(self, sheet: int, right_to_left: bool) -> None:
        """Set the sheet's ``rightToLeft`` flag."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_right_to_left(h, _uint(sheet, "sheet_index"), 1 if right_to_left else 0),
            "fm_sheet_set_right_to_left",
        )

    def set_sheet_tab_selected(self, sheet: int, selected: bool) -> None:
        """Set the sheet's ``tabSelected`` flag."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_tab_selected(h, _uint(sheet, "sheet_index"), 1 if selected else 0),
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
                LIB.fm_sheet_set_view_mode(h, _uint(sheet, "sheet_index"), mode_ptr),
                "fm_sheet_set_view_mode",
            )
        finally:
            LIB.free(mode_ptr)

    def get_auto_filter_xml(self, sheet: int) -> str:
        """Return the sheet's ``<autoFilter>`` XML fragment, or ``""``.

        The fragment carries the filter-column criteria, sort state and
        extension payload verbatim, so a host can preserve definitions it
        does not interpret.
        """
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_sheet_get_auto_filter_xml(h, _uint(sheet, "sheet_index"), out),
                "fm_sheet_get_auto_filter_xml",
            )
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(out)

    def set_auto_filter_xml(self, sheet: int, xml: str) -> None:
        """Replace the sheet's ``<autoFilter>`` XML fragment.

        An empty ``xml`` removes the AutoFilter. Non-empty input must be a
        complete ``<autoFilter ...>`` fragment; the engine preserves filter
        extensions verbatim rather than validating them in detail.
        """
        h = self._require()
        xml_ptr, _ = LIB.alloc_utf8(xml)
        try:
            _check(
                LIB.fm_sheet_set_auto_filter_xml(h, _uint(sheet, "sheet_index"), xml_ptr),
                "fm_sheet_set_auto_filter_xml",
            )
        finally:
            LIB.free(xml_ptr)

    # -- Print settings: raw XML -------------------------------------------
    #
    # The five worksheet print elements are stored verbatim and re-emitted
    # on save, so these getters return what the file says rather than a
    # re-serialisation of a partial model. An absent element reads as "".
    #
    # Setting the empty string removes the element and restores the
    # structured defaults it fed; any non-empty value must be exactly one
    # well-formed element with the matching root name.

    def _get_print_xml(self, sheet: int, export: str) -> str:
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(getattr(LIB, export)(h, _uint(sheet, "sheet_index"), out), export)
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(out)

    def _set_print_xml(self, sheet: int, export: str, xml: str) -> None:
        h = self._require()
        xml_ptr, _ = LIB.alloc_utf8(xml)
        try:
            _check(getattr(LIB, export)(h, _uint(sheet, "sheet_index"), xml_ptr), export)
        finally:
            LIB.free(xml_ptr)

    def get_page_setup_xml(self, sheet: int) -> str:
        """Return the sheet's ``<pageSetup>`` fragment, or ``""``."""
        return self._get_print_xml(sheet, "fm_sheet_get_page_setup_xml")

    def set_page_setup_xml(self, sheet: int, xml: str) -> None:
        """Replace the sheet's ``<pageSetup>`` fragment.

        Note that ``fitToPage`` is not a ``<pageSetup>`` attribute -- use
        :meth:`set_fit_to_page`. A fragment carrying ``r:id`` is rejected
        unless the sheet already has a printerSettings part, because the
        reference would otherwise dangle and Excel would repair the file.
        """
        self._set_print_xml(sheet, "fm_sheet_set_page_setup_xml", xml)

    def get_page_margins_xml(self, sheet: int) -> str:
        """Return the sheet's ``<pageMargins>`` fragment, or ``""``."""
        return self._get_print_xml(sheet, "fm_sheet_get_page_margins_xml")

    def set_page_margins_xml(self, sheet: int, xml: str) -> None:
        """Replace the sheet's ``<pageMargins>`` fragment."""
        self._set_print_xml(sheet, "fm_sheet_set_page_margins_xml", xml)

    def get_print_options_xml(self, sheet: int) -> str:
        """Return the sheet's ``<printOptions>`` fragment, or ``""``."""
        return self._get_print_xml(sheet, "fm_sheet_get_print_options_xml")

    def set_print_options_xml(self, sheet: int, xml: str) -> None:
        """Replace the sheet's ``<printOptions>`` fragment."""
        self._set_print_xml(sheet, "fm_sheet_set_print_options_xml", xml)

    def get_header_footer_xml(self, sheet: int) -> str:
        """Return the sheet's ``<headerFooter>`` fragment, or ``""``."""
        return self._get_print_xml(sheet, "fm_sheet_get_header_footer_xml")

    def set_header_footer_xml(self, sheet: int, xml: str) -> None:
        """Replace the sheet's ``<headerFooter>`` fragment.

        Header and footer codes are introduced by ``&``, which lives in XML
        element text and so must arrive escaped: Excel writes
        ``<oddHeader>&amp;C...</oddHeader>``. Pass decoded text to
        :meth:`set_header_footer` instead to let the engine escape it.
        """
        self._set_print_xml(sheet, "fm_sheet_set_header_footer_xml", xml)

    def get_sheet_pr_xml(self, sheet: int) -> str:
        """Return the sheet's ``<sheetPr>`` fragment, or ``""``.

        ``<sheetPr>`` is not print-only: it also carries the tab colour and
        the VBA code name, which is why toggling fit-to-page has its own
        entry point rather than going through this setter.
        """
        return self._get_print_xml(sheet, "fm_sheet_get_sheet_pr_xml")

    def set_sheet_pr_xml(self, sheet: int, xml: str) -> None:
        """Replace the sheet's ``<sheetPr>`` fragment."""
        self._set_print_xml(sheet, "fm_sheet_set_sheet_pr_xml", xml)

    def set_fit_to_page(self, sheet: int, enabled: bool) -> None:
        """Set or clear ``<sheetPr><pageSetUpPr fitToPage>``.

        Leaves every other ``<sheetPr>`` child and attribute alone.
        ``fitToPage`` only selects the mode; state the target with
        ``fit_to_width`` / ``fit_to_height`` on :meth:`set_page_setup`.
        """
        h = self._require()
        _check(
            LIB.fm_sheet_set_fit_to_page(h, _uint(sheet, "sheet_index"), 1 if enabled else 0),
            "fm_sheet_set_fit_to_page",
        )

    # -- Print settings: area and titles -----------------------------------

    def get_print_area(self, sheet: int) -> str:
        """Return the print area as comma-separated A1 ranges, or ``""``.

        A whole-axis area reads back expanded to explicit corners (``"A:D"``
        becomes ``"A1:D1048576"``), because this reports the rectangles the
        paginator resolves rather than the authored shape.
        """
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(LIB.fm_sheet_get_print_area(h, _uint(sheet, "sheet_index"), out), "fm_sheet_get_print_area")
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(out)

    def set_print_area(self, sheet: int, ranges_a1: str) -> None:
        """Write ``_xlnm.Print_Area`` from one or more A1 ranges.

        Accepts ``"A1:F8"``, ``"A1:B10,D5:E20"``, ``"A:D"`` or ``"1:50"``,
        with or without ``$``. Corners are absolutised and each area is
        qualified with this sheet's name. An empty string removes the name.
        """
        h = self._require()
        ranges_ptr, _ = LIB.alloc_utf8(ranges_a1)
        try:
            _check(
                LIB.fm_sheet_set_print_area(h, _uint(sheet, "sheet_index"), ranges_ptr),
                "fm_sheet_set_print_area",
            )
        finally:
            LIB.free(ranges_ptr)

    def get_print_titles(self, sheet: int) -> tuple[str, str]:
        """Return ``(repeat_rows, repeat_cols)`` as ``"1:2"`` / ``"A:A"``.

        Either is ``""`` when that axis is unset.
        """
        h = self._require()
        rows_out = _alloc_out_ptr()
        cols_out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_sheet_get_print_titles(h, _uint(sheet, "sheet_index"), rows_out, cols_out),
                "fm_sheet_get_print_titles",
            )
            return (LIB.read_cstr(LIB.read_u32(rows_out)), LIB.read_cstr(LIB.read_u32(cols_out)))
        finally:
            LIB.free(rows_out)
            LIB.free(cols_out)

    def set_print_titles(self, sheet: int, repeat_rows: str = "", repeat_cols: str = "") -> None:
        """Write ``_xlnm.Print_Titles`` from a row span and a column span.

        ``repeat_rows`` is a whole-row span (``"1:2"``), ``repeat_cols`` a
        whole-column span (``"A:A"``). Both empty removes the defined name.
        """
        h = self._require()
        rows_ptr, _ = LIB.alloc_utf8(repeat_rows)
        cols_ptr, _ = LIB.alloc_utf8(repeat_cols)
        try:
            _check(
                LIB.fm_sheet_set_print_titles(h, _uint(sheet, "sheet_index"), rows_ptr, cols_ptr),
                "fm_sheet_set_print_titles",
            )
        finally:
            LIB.free(rows_ptr)
            LIB.free(cols_ptr)

    # -- Print settings: manual page breaks --------------------------------

    def add_row_break(self, sheet: int, row: int, manual: bool = True) -> None:
        """Upsert a row break before ``row``, spanning the whole sheet."""
        h = self._require()
        _check(
            LIB.fm_sheet_add_row_break(h, _uint(sheet, "sheet_index"), _uint(row, "row"), 1 if manual else 0),
            "fm_sheet_add_row_break",
        )

    def add_col_break(self, sheet: int, col: int, manual: bool = True) -> None:
        """Upsert a column break before ``col``, spanning the whole sheet."""
        h = self._require()
        _check(
            LIB.fm_sheet_add_col_break(h, _uint(sheet, "sheet_index"), _uint(col, "col"), 1 if manual else 0),
            "fm_sheet_add_col_break",
        )

    def remove_row_break(self, sheet: int, row: int) -> None:
        """Remove the row break at ``row``. An absent break is a no-op."""
        h = self._require()
        _check(
            LIB.fm_sheet_remove_row_break(h, _uint(sheet, "sheet_index"), _uint(row, "row")),
            "fm_sheet_remove_row_break",
        )

    def remove_col_break(self, sheet: int, col: int) -> None:
        """Remove the column break at ``col``. An absent break is a no-op."""
        h = self._require()
        _check(
            LIB.fm_sheet_remove_col_break(h, _uint(sheet, "sheet_index"), _uint(col, "col")),
            "fm_sheet_remove_col_break",
        )

    def clear_breaks(self, sheet: int) -> None:
        """Remove every manual break on both axes."""
        h = self._require()
        _check(LIB.fm_sheet_clear_breaks(h, _uint(sheet, "sheet_index")), "fm_sheet_clear_breaks")

    def _read_breaks(self, sheet: int, count_export: str, at_export: str) -> List[PageBreak]:
        h = self._require()
        count = int(getattr(LIB, count_export)(h, _uint(sheet, "sheet_index")))
        out: List[PageBreak] = []
        ptr = S.alloc_struct(LIB, S.PAGE_BREAK)
        try:
            for i in range(count):
                _check(getattr(LIB, at_export)(h, _uint(sheet, "sheet_index"), _uint(i, "index"), ptr), at_export)
                d = S.PAGE_BREAK.unpack(LIB, ptr)
                out.append(PageBreak(id=d["id"], min=d["min"], max=d["max"], manual=bool(d["manual"])))
        finally:
            LIB.free(ptr)
        return out

    def get_row_breaks(self, sheet: int) -> List[PageBreak]:
        """Return the row breaks, ascending by row."""
        return self._read_breaks(sheet, "fm_sheet_row_break_count", "fm_sheet_row_break_at")

    def get_col_breaks(self, sheet: int) -> List[PageBreak]:
        """Return the column breaks, ascending by column."""
        return self._read_breaks(sheet, "fm_sheet_col_break_count", "fm_sheet_col_break_at")

    # -- Print settings: typed patch setters -------------------------------
    #
    # Each keyword defaults to ``None`` meaning "leave this attribute as the
    # file has it". That is how the C ``*_engaged`` flags are expressed in
    # Python: a caller states the two or three attributes it cares about and
    # every other one, including attributes the engine does not model, stays
    # exactly where the file put it.

    def set_page_setup(
        self,
        sheet: int,
        *,
        orientation: Optional[int] = None,
        paper_size: Optional[int] = None,
        scale: Optional[int] = None,
        fit_to_width: Optional[int] = None,
        fit_to_height: Optional[int] = None,
        fit_to_page: Optional[bool] = None,
    ) -> None:
        """Apply a partial ``<pageSetup>`` update.

        ``orientation`` is 0 (default) / 1 (portrait) / 2 (landscape);
        ``paper_size`` is the OOXML code (9 = A4). ``scale`` outside
        ``[10, 400]`` is rejected rather than clamped -- a mis-stated print
        scale lands on paper, so rounding it silently would hide the error.
        ``fit_to_page`` is routed to ``<sheetPr><pageSetUpPr>``.
        """
        h = self._require()
        ptr = S.alloc_struct(LIB, S.PAGE_SETUP)
        try:
            S.PAGE_SETUP.pack(
                LIB,
                ptr,
                {
                    "orientation_engaged": 0 if orientation is None else 1,
                    "orientation": _uint(orientation or 0, "orientation"),
                    "paper_size_engaged": 0 if paper_size is None else 1,
                    "paper_size": _uint(paper_size or 0, "paper_size"),
                    "scale_engaged": 0 if scale is None else 1,
                    "scale": _uint(scale or 0, "scale"),
                    "fit_to_width_engaged": 0 if fit_to_width is None else 1,
                    "fit_to_width": _uint(fit_to_width or 0, "fit_to_width"),
                    "fit_to_height_engaged": 0 if fit_to_height is None else 1,
                    "fit_to_height": _uint(fit_to_height or 0, "fit_to_height"),
                    "fit_to_page_engaged": 0 if fit_to_page is None else 1,
                    "fit_to_page": 1 if fit_to_page else 0,
                },
            )
            _check(LIB.fm_sheet_set_page_setup(h, _uint(sheet, "sheet_index"), ptr), "fm_sheet_set_page_setup")
        finally:
            LIB.free(ptr)

    def set_page_margins(
        self,
        sheet: int,
        *,
        left: Optional[float] = None,
        right: Optional[float] = None,
        top: Optional[float] = None,
        bottom: Optional[float] = None,
        header: Optional[float] = None,
        footer: Optional[float] = None,
    ) -> None:
        """Apply a partial ``<pageMargins>`` update, in inches.

        A negative, infinite or NaN margin is rejected: the paginator
        subtracts these from the paper, so such a value does not describe a
        printable body.
        """
        h = self._require()
        ptr = S.alloc_struct(LIB, S.PAGE_MARGINS)
        try:
            values: Dict[str, object] = {}
            for name, value in (
                ("left", left),
                ("right", right),
                ("top", top),
                ("bottom", bottom),
                ("header", header),
                ("footer", footer),
            ):
                values[f"{name}_engaged"] = 0 if value is None else 1
                values[name] = 0.0 if value is None else float(value)
            S.PAGE_MARGINS.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_set_page_margins(h, _uint(sheet, "sheet_index"), ptr), "fm_sheet_set_page_margins")
        finally:
            LIB.free(ptr)

    def set_print_options(
        self,
        sheet: int,
        *,
        grid_lines: Optional[bool] = None,
        headings: Optional[bool] = None,
        horizontal_centered: Optional[bool] = None,
        vertical_centered: Optional[bool] = None,
    ) -> None:
        """Apply a partial ``<printOptions>`` update."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.PRINT_OPTIONS)
        try:
            values: Dict[str, object] = {}
            for name, value in (
                ("grid_lines", grid_lines),
                ("headings", headings),
                ("horizontal_centered", horizontal_centered),
                ("vertical_centered", vertical_centered),
            ):
                values[f"{name}_engaged"] = 0 if value is None else 1
                values[name] = 1 if value else 0
            S.PRINT_OPTIONS.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_set_print_options(h, _uint(sheet, "sheet_index"), ptr), "fm_sheet_set_print_options")
        finally:
            LIB.free(ptr)

    def set_header_footer(
        self,
        sheet: int,
        *,
        odd_header: Optional[str] = None,
        odd_footer: Optional[str] = None,
        even_header: Optional[str] = None,
        even_footer: Optional[str] = None,
        first_header: Optional[str] = None,
        first_footer: Optional[str] = None,
        different_odd_even: Optional[bool] = None,
        different_first: Optional[bool] = None,
        scale_with_doc: Optional[bool] = None,
        align_with_margins: Optional[bool] = None,
    ) -> None:
        """Apply a partial ``<headerFooter>`` update.

        Each section is tri-state: ``None`` leaves it untouched, ``""``
        clears it, anything else replaces it. Pass decoded text -- Excel's
        formatting codes go in plainly (``"&C&\\"MS Gothic\\"Report &P/&N"``)
        and the engine applies the XML escaping.
        """
        h = self._require()
        ptr = S.alloc_struct(LIB, S.HEADER_FOOTER)
        owned: List[int] = []
        try:
            values: Dict[str, object] = {}
            for name, value in (
                ("different_odd_even", different_odd_even),
                ("different_first", different_first),
                ("scale_with_doc", scale_with_doc),
                ("align_with_margins", align_with_margins),
            ):
                values[f"{name}_engaged"] = 0 if value is None else 1
                values[name] = 1 if value else 0
            S.HEADER_FOOTER.pack(LIB, ptr, values)
            # Written after `pack`, which zeroes the whole block: a NULL
            # pointer is the "leave this section alone" signal, and an
            # empty string still needs a real (empty) buffer to mean
            # "clear it".
            for field, value in (
                ("odd_header", odd_header),
                ("odd_footer", odd_footer),
                ("even_header", even_header),
                ("even_footer", even_footer),
                ("first_header", first_header),
                ("first_footer", first_footer),
            ):
                if value is None:
                    continue
                text_ptr, _ = LIB.alloc_utf8(value)
                owned.append(text_ptr)
                _, off = S.HEADER_FOOTER.offsets[field]
                LIB.write_bytes(ptr + off, struct.pack("<I", text_ptr))
            _check(LIB.fm_sheet_set_header_footer(h, _uint(sheet, "sheet_index"), ptr), "fm_sheet_set_header_footer")
        finally:
            LIB.free(ptr)
            for owned_ptr in owned:
                LIB.free(owned_ptr)

    def get_page_setup(self, sheet: int) -> PageSetup:
        """Read the effective page setup and which attributes state it."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.PAGE_SETUP)
        try:
            _check(LIB.fm_sheet_get_page_setup(h, _uint(sheet, "sheet_index"), ptr), "fm_sheet_get_page_setup")
            d = S.PAGE_SETUP.unpack(LIB, ptr)
            return PageSetup(
                orientation=d["orientation"],
                paper_size=d["paper_size"],
                scale=d["scale"],
                fit_to_width=d["fit_to_width"],
                fit_to_height=d["fit_to_height"],
                fit_to_page=bool(d["fit_to_page"]),
                orientation_stated=bool(d["orientation_engaged"]),
                paper_size_stated=bool(d["paper_size_engaged"]),
                scale_stated=bool(d["scale_engaged"]),
                fit_to_width_stated=bool(d["fit_to_width_engaged"]),
                fit_to_height_stated=bool(d["fit_to_height_engaged"]),
                fit_to_page_stated=bool(d["fit_to_page_engaged"]),
            )
        finally:
            LIB.free(ptr)

    def get_page_margins(self, sheet: int) -> PageMargins:
        """Read the effective page margins in inches, with presence flags."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.PAGE_MARGINS)
        try:
            _check(LIB.fm_sheet_get_page_margins(h, _uint(sheet, "sheet_index"), ptr), "fm_sheet_get_page_margins")
            d = S.PAGE_MARGINS.unpack(LIB, ptr)
            return PageMargins(
                left=d["left"],
                right=d["right"],
                top=d["top"],
                bottom=d["bottom"],
                header=d["header"],
                footer=d["footer"],
                left_stated=bool(d["left_engaged"]),
                right_stated=bool(d["right_engaged"]),
                top_stated=bool(d["top_engaged"]),
                bottom_stated=bool(d["bottom_engaged"]),
                header_stated=bool(d["header_engaged"]),
                footer_stated=bool(d["footer_engaged"]),
            )
        finally:
            LIB.free(ptr)

    def get_sheet_columns(self, sheet: int) -> List[ColumnLayout]:
        """Return the column-layout overrides on ``sheet``."""
        h = self._require()
        n = _read_count(LIB.fm_sheet_get_column_count, h, _uint(sheet, "sheet_index"))
        out: List[ColumnLayout] = []
        # One scratch block for the whole walk; see iter_defined_names.
        ptr = S.alloc_struct(LIB, S.COLUMN_LAYOUT)
        try:
            for i in range(n):
                S.zero_struct(LIB, S.COLUMN_LAYOUT, ptr)
                _check(
                    LIB.fm_sheet_get_column(h, _uint(sheet, "sheet_index"), _uint(i, "idx"), ptr),
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
                        has_width=bool(d["has_width"]),
                        has_style=bool(d["has_style"]),
                        style_xf=d["style_xf"],
                    )
                )
        finally:
            LIB.free(ptr)
        return out

    def set_column_width(self, sheet: int, first: int, last: int, width: float) -> None:
        """Set the column width override on ``[first, last]``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_column_width(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), float(width)
            ),
            "fm_sheet_set_column_width",
        )

    def set_column_hidden(self, sheet: int, first: int, last: int, hidden: bool) -> None:
        """Set the column hidden flag on ``[first, last]``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_column_hidden(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), 1 if hidden else 0
            ),
            "fm_sheet_set_column_hidden",
        )

    def set_column_outline(self, sheet: int, first: int, last: int, level: int) -> None:
        """Set the column outline level on ``[first, last]``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_column_outline(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), _uint(level, "level", 8)
            ),
            "fm_sheet_set_column_outline",
        )

    def get_sheet_row_overrides(self, sheet: int) -> List[RowLayout]:
        """Return the row-layout overrides on ``sheet``."""
        h = self._require()
        n = _read_count(LIB.fm_sheet_get_row_override_count, h, _uint(sheet, "sheet_index"))
        out: List[RowLayout] = []
        # One scratch block for the whole walk; see iter_defined_names.
        ptr = S.alloc_struct(LIB, S.ROW_LAYOUT)
        try:
            for i in range(n):
                S.zero_struct(LIB, S.ROW_LAYOUT, ptr)
                _check(
                    LIB.fm_sheet_get_row_override(h, _uint(sheet, "sheet_index"), _uint(i, "idx"), ptr),
                    "fm_sheet_get_row_override",
                )
                d = S.ROW_LAYOUT.unpack(LIB, ptr)
                out.append(
                    RowLayout(
                        row=d["row"],
                        height=d["height"],
                        hidden=bool(d["hidden"]),
                        outline_level=d["outline_level"],
                        has_style=bool(d["has_style"]),
                        style_xf=d["style_xf"],
                    )
                )
        finally:
            LIB.free(ptr)
        return out

    def set_row_height(self, sheet: int, row: int, height: float) -> None:
        """Set the row height override at ``row``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_row_height(h, _uint(sheet, "sheet_index"), _uint(row, "row"), float(height)),
            "fm_sheet_set_row_height",
        )

    def set_row_hidden(self, sheet: int, row: int, hidden: bool) -> None:
        """Set the row hidden flag at ``row``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_row_hidden(h, _uint(sheet, "sheet_index"), _uint(row, "row"), 1 if hidden else 0),
            "fm_sheet_set_row_hidden",
        )

    def set_row_outline(self, sheet: int, row: int, level: int) -> None:
        """Set the row outline level at ``row``."""
        h = self._require()
        _check(
            LIB.fm_sheet_set_row_outline(h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(level, "level", 8)),
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
        """Evaluate every CF block on ``sheet`` against an inclusive range.

        Args:
          sheet: 0-based sheet index.
          first_row: 0-based inclusive first row.
          first_col: 0-based inclusive first column.
          last_row: 0-based inclusive last row.
          last_col: 0-based inclusive last column.
          today_serial: date basis pinned for ``timePeriod`` rules, as an
            Excel serial date. The default ``float("nan")`` disables them:
            no ``timePeriod`` rule can match, so pass a serial date to use
            them. ``0.0`` is *not* a disabling value -- it is the valid
            serial for 1899-12-30 and would evaluate those rules against
            that date.
        """
        h = self._require()
        out = _alloc_out_ptr()
        results = 0
        try:
            _check(
                LIB.fm_workbook_cf_evaluate_range(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(first_row, "first_row"),
                    _uint(first_col, "first_col"),
                    _uint(last_row, "last_row"),
                    _uint(last_col, "last_col"),
                    float(today_serial),
                    out,
                ),
                "fm_workbook_cf_evaluate_range",
            )
            results = LIB.read_u32(out)
            cells: List[CfCellResult] = []
            n = LIB.fm_cf_results_cell_count(results)
            # One fixed set of out slots for the whole walk: allocating
            # per cell would put a locked WASM malloc/free pair on every
            # item, which dominates the cost of a viewport-sized range.
            rp = _alloc_out_ptr()
            cp = _alloc_out_ptr()
            mcp = _alloc_out_ptr()
            mptr = S.alloc_struct(LIB, S.CF_MATCH)
            try:
                for ci in range(n):
                    _check(
                        LIB.fm_cf_results_cell_at(results, _uint(ci, "cell_idx"), rp, cp, mcp),
                        "fm_cf_results_cell_at",
                    )
                    row = LIB.read_u32(rp)
                    col = LIB.read_u32(cp)
                    mcount = LIB.read_u32(mcp)
                    matches: List[CfMatch] = []
                    for mi in range(mcount):
                        S.zero_struct(LIB, S.CF_MATCH, mptr)
                        _check(
                            LIB.fm_cf_results_match_at(results, _uint(ci, "cell_idx"), _uint(mi, "match_idx"), mptr),
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
                    cells.append(CfCellResult(row=row, col=col, matches=matches))
            finally:
                LIB.free(rp)
                LIB.free(cp)
                LIB.free(mcp)
                LIB.free(mptr)
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
            _check(LIB.fm_sheet_cf_count(h, _uint(sheet, "sheet_index"), out), "fm_sheet_cf_count")
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
        S.CFVO.pack(LIB, ptr, {"type": _uint(value.type, "type", 8), "gte": 1 if value.gte else 0})
        if value.value is not None:
            S.write_str_field(LIB, ptr, S.CFVO, "value", value.value, owned)

    @staticmethod
    def _write_cf_color(ptr: int, color: CfColor) -> None:
        S.CF_COLOR.pack(
            LIB,
            ptr,
            {
                "r": _uint(color.r, "r", 8),
                "g": _uint(color.g, "g", 8),
                "b": _uint(color.b, "b", 8),
                "a": _uint(color.a, "a", 8),
            },
        )

    def get_conditional_format_at(self, sheet: int, index: int) -> ConditionalFormat:
        """Read the ``index``-th CF rule on ``sheet`` (flattened order)."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.CF_RULE)
        try:
            _check(
                LIB.fm_sheet_cf_get_at(h, _uint(sheet, "sheet_index"), _uint(index, "idx"), ptr),
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
                    # x14 extension. Each field is decoded only when the C
                    # ABI reports its `*_engaged` flag, so `None` keeps its
                    # "no override" meaning on the way back in.
                    bool(d["data_bar_gradient"]) if d["data_bar_gradient_engaged"] else None,
                    d["data_bar_axis_position"] if d["data_bar_axis_position_engaged"] else None,
                    self._decode_cf_color_at(ptr + offsets["data_bar_negative_fill"][1])
                    if d["data_bar_negative_fill_engaged"]
                    else None,
                    self._decode_cf_color_at(ptr + offsets["data_bar_border"][1])
                    if d["data_bar_border_engaged"]
                    else None,
                    self._decode_cf_color_at(ptr + offsets["data_bar_negative_border"][1])
                    if d["data_bar_negative_border_engaged"]
                    else None,
                    self._decode_cf_color_at(ptr + offsets["data_bar_axis_color"][1])
                    if d["data_bar_axis_color_engaged"]
                    else None,
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
                    "type": _uint(rule.type, "type", 8),
                    "op": _uint(rule.op, "op", 8),
                    "time_period": _uint(rule.time_period, "time_period", 8),
                    "priority": _sint(rule.priority, "priority"),
                    "stop_if_true": 1 if rule.stop_if_true else 0,
                    "dxf_id_engaged": 1 if rule.dxf_id_engaged else 0,
                    "dxf_id": _uint(rule.dxf_id, "dxf_id"),
                    "op_engaged": 1 if rule.op_engaged else 0,
                    "rank_engaged": 1 if rule.rank_engaged else 0,
                    "rank": _sint(rule.rank, "rank"),
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
                            "first_row": _uint(r.first_row, "first_row"),
                            "first_col": _uint(r.first_col, "first_col"),
                            "last_row": _uint(r.last_row, "last_row"),
                            "last_col": _uint(r.last_col, "last_col"),
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
                # x14 extension. A `None` leaves the `*_engaged` flag clear,
                # which the C ABI reads as "keep the model default".
                if data_bar.gradient is not None:
                    LIB.write_bytes(ptr + ro["data_bar_gradient_engaged"][1], struct.pack("<i", 1))
                    LIB.write_bytes(ptr + ro["data_bar_gradient"][1], struct.pack("<i", 1 if data_bar.gradient else 0))
                if data_bar.axis_position is not None:
                    LIB.write_bytes(ptr + ro["data_bar_axis_position_engaged"][1], struct.pack("<i", 1))
                    LIB.write_bytes(
                        ptr + ro["data_bar_axis_position"][1], struct.pack("<B", int(data_bar.axis_position))
                    )
                for field, color in (
                    ("negative_fill", data_bar.negative_fill),
                    ("border", data_bar.border),
                    ("negative_border", data_bar.negative_border),
                    ("axis_color", data_bar.axis_color),
                ):
                    if color is not None:
                        LIB.write_bytes(ptr + ro[f"data_bar_{field}_engaged"][1], struct.pack("<i", 1))
                        self._write_cf_color(ptr + ro[f"data_bar_{field}"][1], color)
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
                LIB.fm_sheet_cf_add_rule(h, _uint(sheet, "sheet_index"), ptr, out),
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
            LIB.fm_sheet_cf_remove_at(h, _uint(sheet, "sheet_index"), _uint(index, "idx")),
            "fm_sheet_cf_remove_at",
        )

    def clear_conditional_formats(self, sheet: int) -> None:
        """Drop every CF block on ``sheet``."""
        h = self._require()
        _check(LIB.fm_sheet_cf_clear(h, _uint(sheet, "sheet_index")), "fm_sheet_cf_clear")

    # -- Styles ------------------------------------------------------------
    def get_cell_xf_index(self, sheet: int, row: int, col: int) -> int:
        """Return the xf (style record) index attached to a cell."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_cell_get_xf_index(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), out),
                "fm_cell_get_xf_index",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    def set_cell_xf_index(self, sheet: int, row: int, col: int, xf_index: int) -> None:
        """Store ``xf_index`` on the cell at ``(row, col)``."""
        h = self._require()
        _check(
            LIB.fm_cell_set_xf_index(
                h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), _uint(xf_index, "xf_index")
            ),
            "fm_cell_set_xf_index",
        )

    def set_range_xf_index(
        self, sheet: int, first_row: int, first_col: int, last_row: int, last_col: int, xf_index: int
    ) -> None:
        """Store ``xf_index`` on every cell in the inclusive rectangle.

        Cells that do not exist yet are materialised as blanks carrying only
        the style, so an empty ruled box renders. Ruling a report this way
        costs one call instead of one per cell.
        """
        h = self._require()
        _check(
            LIB.fm_sheet_set_range_xf_index(
                h,
                _uint(sheet, "sheet"),
                _uint(first_row, "first_row"),
                _uint(first_col, "first_col"),
                _uint(last_row, "last_row"),
                _uint(last_col, "last_col"),
                _uint(xf_index, "xf_index"),
            ),
            "fm_sheet_set_range_xf_index",
        )

    @staticmethod
    def _decode_cell_xf(ptr: int) -> CellXf:
        d = S.CELL_XF.unpack(LIB, ptr)
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
        ptr = S.alloc_struct(LIB, S.CELL_XF)
        try:
            _check(
                LIB.fm_styles_get_cell_xf(h, _uint(xf_index, "xf_index"), ptr),
                "fm_styles_get_cell_xf",
            )
            return self._decode_cell_xf(ptr)
        finally:
            LIB.free(ptr)

    def get_font(self, font_index: int) -> FontRecord:
        """Return the font record at ``font_index``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.FONT_RECORD)
        try:
            _check(LIB.fm_styles_get_font(h, _uint(font_index, "font_index"), ptr), "fm_styles_get_font")
            return _decode_font(ptr)
        finally:
            LIB.free(ptr)

    def get_fill(self, fill_index: int) -> FillRecord:
        """Return the fill record at ``fill_index``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.FILL_RECORD)
        try:
            _check(LIB.fm_styles_get_fill(h, _uint(fill_index, "fill_index"), ptr), "fm_styles_get_fill")
            return _decode_fill(ptr)
        finally:
            LIB.free(ptr)

    def get_border(self, border_index: int) -> Dict[str, object]:
        """Return the border record at ``border_index``.

        Each side is a ``{"style", "color_argb", "color"}`` dict, where
        ``color`` is the round-trip :class:`ColorSpec`; the result also
        carries ``diagonal_up`` / ``diagonal_down`` booleans.
        """
        h = self._require()
        ptr = S.alloc_struct(LIB, S.BORDER_RECORD)
        try:
            _check(
                LIB.fm_styles_get_border(h, _uint(border_index, "border_index"), ptr),
                "fm_styles_get_border",
            )
            return _decode_border_record(ptr)
        finally:
            LIB.free(ptr)

    def get_dxf(self, dxf_index: int) -> DifferentialFormat:
        """Read one differential format by its conditional-format ``dxfId``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.DXF_RECORD)
        try:
            _check(LIB.fm_styles_get_dxf(h, _uint(dxf_index, "dxf_index"), ptr), "fm_styles_get_dxf")
            d = S.DXF_RECORD.unpack(LIB, ptr)
            offsets = S.DXF_RECORD.offsets
            font = _decode_font(ptr + offsets["font"][1]) if d["font_engaged"] else None
            fill = _decode_fill(ptr + offsets["fill"][1]) if d["fill_engaged"] else None
            return DifferentialFormat(
                font=font,
                fill=fill,
                border=_decode_border_record(ptr + offsets["border"][1]) if d["border_engaged"] else None,
                num_fmt_id=d["num_fmt_id"] if d["num_fmt_engaged"] else None,
                num_fmt_code=LIB.read_cstr(d["num_fmt_code"]) if d["num_fmt_engaged"] else "",
                alignment_xml=LIB.read_cstr(d["alignment_xml"]),
                protection_xml=LIB.read_cstr(d["protection_xml"]),
            )
        finally:
            LIB.free(ptr)

    def get_num_fmt(self, num_fmt_id: int) -> str:
        """Return the format code registered for ``num_fmt_id``."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_styles_get_num_fmt_string(h, _uint(num_fmt_id, "num_fmt_id", 16), out),
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
            _pack_font(ptr, record, owned)
            _check(LIB.fm_styles_add_font(h, ptr, out), "fm_styles_add_font")
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)
            for p in owned:
                LIB.free(p)

    def set_font(self, font_index: int, record: FontRecord) -> None:
        """Overwrite an existing font slot in place.

        Every ``<xf>`` naming ``font_index`` restyles at once, so this is a
        bulk change rather than a local edit; :meth:`add_font` is the way to
        introduce a new appearance. The index must already exist.
        """
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.FONT_RECORD)
        try:
            _pack_font(ptr, record, owned)
            _check(
                LIB.fm_styles_set_font(h, _uint(font_index, "font_index"), ptr),
                "fm_styles_set_font",
            )
        finally:
            LIB.free(ptr)
            for p in owned:
                LIB.free(p)

    def set_default_font(self, record: FontRecord) -> None:
        """Declare the workbook's default font (font 0).

        Font 0 is what an unstyled cell resolves to. A new workbook seeds it
        with Excel's Calibri 11, and :meth:`add_font` can only append beside
        it, so this is the way a ja-JP host declares e.g. ``游ゴシック`` for
        the cells it never styles explicitly. Read it back with
        ``get_font(0)``.
        """
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.FONT_RECORD)
        try:
            _pack_font(ptr, record, owned)
            _check(LIB.fm_workbook_set_default_font(h, ptr), "fm_workbook_set_default_font")
        finally:
            LIB.free(ptr)
            for p in owned:
                LIB.free(p)

    def add_fill(self, record: FillRecord) -> int:
        """Add (dedup) a fill record; return its index."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.FILL_RECORD)
        out = _alloc_out_ptr()
        try:
            _pack_fill(ptr, record)
            _check(LIB.fm_styles_add_fill(h, ptr, out), "fm_styles_add_fill")
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)

    def add_border(self, sides: Dict[str, object]) -> int:
        """Add (dedup) a border record; return its index.

        ``sides`` mirrors :meth:`get_border`'s shape: a dict with
        ``left`` / ``right`` / ``top`` / ``bottom`` / ``diagonal`` entries
        (each ``{"style", "color_argb"}`` plus an optional ``color``
        :class:`ColorSpec`) plus optional ``diagonal_up`` /
        ``diagonal_down`` booleans.
        """
        h = self._require()
        ptr = S.alloc_struct(LIB, S.BORDER_RECORD)
        out = _alloc_out_ptr()
        try:
            _pack_border_record(ptr, sides)
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
        return self._add_xf_record(record, LIB.fm_styles_add_cell_xf, "fm_styles_add_cell_xf")

    def add_cell_style_xf(self, record: CellXf) -> int:
        """Add (dedup) a named-style ``<xf>``; return its ``<cellStyleXfs>`` id.

        The counterpart of :meth:`get_cell_style_xf`. Named-style xfs are
        the records a ``<cellStyle>`` points at; :meth:`add_cell_xf` writes
        the per-cell ``<cellXfs>`` pool instead.
        """
        return self._add_xf_record(record, LIB.fm_styles_add_cell_style_xf, "fm_styles_add_cell_style_xf")

    def _add_xf_record(self, record: CellXf, fn, op: str) -> int:
        h = self._require()
        ptr = S.alloc_struct(LIB, S.CELL_XF)
        out = _alloc_out_ptr()
        try:
            S.CELL_XF.pack(LIB, ptr, _cell_xf_fields(record))
            _check(fn(h, ptr, out), op)
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)

    def add_batch(
        self,
        *,
        fonts: Optional[Sequence[FontRecord]] = None,
        fills: Optional[Sequence[FillRecord]] = None,
        borders: Optional[Sequence[Dict[str, object]]] = None,
        num_fmts: Optional[Sequence[str]] = None,
        cell_xfs: Optional[Sequence[CellXf]] = None,
    ) -> StyleBatchIndices:
        """Add (dedup) whole style tables in one call.

        The per-record :meth:`add_font` / :meth:`add_fill` / :meth:`add_border`
        / :meth:`add_num_fmt` / :meth:`add_cell_xf` methods each rescan their
        table, so building thousands of formats one at a time is quadratic.
        This routes the same work through a single ABI call that dedups every
        table in linear time.

        Fonts, fills, borders and number formats are installed before the
        xfs, so a ``cell_xfs`` entry may reference an index this same call
        assigns. ``borders`` entries take :meth:`add_border`'s dict shape.

        The call is one transaction: on failure the workbook is unchanged
        and :class:`FormulonError` is raised.
        """
        h = self._require()
        font_list = list(fonts or ())
        fill_list = list(fills or ())
        border_list = list(borders or ())
        num_fmt_list = list(num_fmts or ())
        xf_list = list(cell_xfs or ())

        owned: List[int] = []
        batch = S.alloc_struct(LIB, S.STYLES_BATCH)
        try:
            fields: Dict[str, object] = {}

            def alloc_array(element_size: int, count: int) -> int:
                """Allocate a zeroed array of ``count`` elements, tracked for release."""
                ptr = LIB.alloc(element_size * count)
                owned.append(ptr)
                LIB.write_bytes(ptr, b"\x00" * (element_size * count))
                return ptr

            font_indices = 0
            if font_list:
                records = alloc_array(S.FONT_RECORD.size, len(font_list))
                font_indices = alloc_array(4, len(font_list))
                for i, record in enumerate(font_list):
                    _pack_font(records + i * S.FONT_RECORD.size, record, owned)
                fields.update(fonts=records, font_count=len(font_list), font_indices=font_indices)

            fill_indices = 0
            if fill_list:
                records = alloc_array(S.FILL_RECORD.size, len(fill_list))
                fill_indices = alloc_array(4, len(fill_list))
                for i, record in enumerate(fill_list):
                    _pack_fill(records + i * S.FILL_RECORD.size, record)
                fields.update(fills=records, fill_count=len(fill_list), fill_indices=fill_indices)

            border_indices = 0
            if border_list:
                records = alloc_array(S.BORDER_RECORD.size, len(border_list))
                border_indices = alloc_array(4, len(border_list))
                for i, sides in enumerate(border_list):
                    _pack_border_record(records + i * S.BORDER_RECORD.size, sides)
                fields.update(borders=records, border_count=len(border_list), border_indices=border_indices)

            xf_indices = 0
            if xf_list:
                records = alloc_array(S.CELL_XF.size, len(xf_list))
                xf_indices = alloc_array(4, len(xf_list))
                for i, record in enumerate(xf_list):
                    S.CELL_XF.pack(LIB, records + i * S.CELL_XF.size, _cell_xf_fields(record))
                fields.update(cell_xfs=records, cell_xf_count=len(xf_list), cell_xf_indices=xf_indices)

            num_fmt_ids = 0
            if num_fmt_list:
                # An array of `const char*`: one 4-byte pointer slot per code.
                codes = alloc_array(4, len(num_fmt_list))
                num_fmt_ids = alloc_array(2, len(num_fmt_list))
                for i, code in enumerate(num_fmt_list):
                    code_ptr, _ = LIB.alloc_utf8(code)
                    owned.append(code_ptr)
                    LIB.write_bytes(codes + i * 4, struct.pack("<I", code_ptr))
                fields.update(num_fmt_codes=codes, num_fmt_count=len(num_fmt_list), num_fmt_ids=num_fmt_ids)

            S.STYLES_BATCH.pack(LIB, batch, fields)
            _check(LIB.fm_styles_add_batch(h, batch), "fm_styles_add_batch")

            def read_u32_array(ptr: int, count: int) -> List[int]:
                if count == 0:
                    return []
                raw = LIB.read_bytes(ptr, 4 * count)
                return [struct.unpack_from("<I", raw, 4 * i)[0] for i in range(count)]

            def read_u16_array(ptr: int, count: int) -> List[int]:
                if count == 0:
                    return []
                raw = LIB.read_bytes(ptr, 2 * count)
                return [struct.unpack_from("<H", raw, 2 * i)[0] for i in range(count)]

            return StyleBatchIndices(
                fonts=read_u32_array(font_indices, len(font_list)),
                fills=read_u32_array(fill_indices, len(fill_list)),
                borders=read_u32_array(border_indices, len(border_list)),
                cell_xfs=read_u32_array(xf_indices, len(xf_list)),
                num_fmts=read_u16_array(num_fmt_ids, len(num_fmt_list)),
            )
        finally:
            LIB.free(batch)
            for ptr in owned:
                LIB.free(ptr)

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
                    "num_fmt_id": _uint(record.num_fmt_id or 0, "num_fmt_id", 16),
                },
            )
            offsets = S.DXF_RECORD.offsets
            if record.font is not None:
                _pack_font(ptr + offsets["font"][1], record.font, owned)
            if record.fill is not None:
                _pack_fill(ptr + offsets["fill"][1], record.fill)
            if record.border is not None:
                _pack_border_record(ptr + offsets["border"][1], record.border)
            S.write_str_field(LIB, ptr, S.DXF_RECORD, "num_fmt_code", record.num_fmt_code, owned)
            S.write_str_field(LIB, ptr, S.DXF_RECORD, "alignment_xml", record.alignment_xml, owned)
            S.write_str_field(LIB, ptr, S.DXF_RECORD, "protection_xml", record.protection_xml, owned)
            _check(LIB.fm_styles_add_dxf(h, ptr, out), "fm_styles_add_dxf")
            return LIB.read_u32(out)
        finally:
            LIB.free(ptr)
            LIB.free(out)
            for owned_ptr in owned:
                LIB.free(owned_ptr)

    def set_cell_style(self, name: str, xf_id: int, builtin_id: int = CELL_STYLE_BUILTIN_ID_NONE) -> None:
        """Add or replace a named ``<cellStyle>`` record.

        ``xf_id`` must reference an existing named-style xf -- register one
        with :meth:`add_cell_style_xf` first. ``builtin_id`` is an Excel
        built-in style ordinal (0..47); leave it at
        ``CELL_STYLE_BUILTIN_ID_NONE`` for a custom style.
        """
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(name)
        try:
            _check(
                LIB.fm_styles_set_cell_style(h, name_ptr, _uint(xf_id, "xf_id"), _uint(builtin_id, "builtin_id")),
                "fm_styles_set_cell_style",
            )
        finally:
            LIB.free(name_ptr)

    def get_cell_style(self, index: int) -> CellStyle:
        """Return the named cell style at ``index``."""
        h = self._require()
        ptr = S.alloc_struct(LIB, S.CELL_STYLE_RECORD)
        try:
            _check(
                LIB.fm_styles_get_cell_style(h, _uint(index, "index"), ptr),
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
        ptr = S.alloc_struct(LIB, S.CELL_XF)
        try:
            _check(
                LIB.fm_styles_get_cell_style_xf(h, _uint(index, "index"), ptr),
                "fm_styles_get_cell_style_xf",
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
                LIB.fm_workbook_pivot_count(h, _uint(sheet, "sheet_index"), out),
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
                LIB.fm_workbook_pivot_layout(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), out),
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
            # One scratch block for the whole walk; see iter_defined_names.
            cptr = S.alloc_struct(LIB, S.PIVOT_CELL)
            try:
                for i in range(n):
                    S.zero_struct(LIB, S.PIVOT_CELL, cptr)
                    _check(
                        LIB.fm_pivot_cells_at(handle, _uint(i, "idx"), cptr),
                        "fm_pivot_cells_at",
                    )
                    d = S.PIVOT_CELL.unpack(LIB, cptr)
                    value = Value._from_wasm(cptr + S.PIVOT_CELL_VALUE_OFFSET)
                    cells.append(
                        PivotCell(
                            row=d["row"],
                            col=d["col"],
                            value=value,
                            kind=PivotCellKind(d["kind"]),
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
                LIB.fm_workbook_pivot_cache_id_at(h, _uint(index, "idx"), out),
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
                LIB.fm_workbook_pivot_cache_create(h, _uint(requested_id, "requested_id"), out),
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
                    h, _uint(cache_id, "cache_id"), present, out_ref, out_sheet, out_name
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
                    h, _uint(cache_id, "cache_id"), 0 if source is None else 1, *pointers
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
            LIB.fm_workbook_pivot_cache_remove(h, _uint(cache_id, "cache_id")),
            "fm_workbook_pivot_cache_remove",
        )

    def pivot_cache_field_count(self, cache_id: int) -> int:
        """Return the number of fields on the cache."""
        h = self._require()
        return _read_count(LIB.fm_workbook_pivot_cache_field_count, h, _uint(cache_id, "cache_id"))

    def pivot_cache_field_name(self, cache_id: int, field_idx: int) -> str:
        """Read the name of cache field ``field_idx``."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_cache_field_name(
                    h, _uint(cache_id, "cache_id"), _uint(field_idx, "field_idx"), out
                ),
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
                LIB.fm_workbook_pivot_cache_field_add(h, _uint(cache_id, "cache_id"), name_ptr, out),
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
            LIB.fm_workbook_pivot_cache_field_clear(h, _uint(cache_id, "cache_id")),
            "fm_workbook_pivot_cache_field_clear",
        )

    def pivot_cache_field_shared_item_count(self, cache_id: int, field_idx: int) -> int:
        """Return the shared-item count on cache field ``field_idx``."""
        h = self._require()
        return _read_count(
            LIB.fm_workbook_pivot_cache_field_shared_item_count,
            h,
            _uint(cache_id, "cache_id"),
            _uint(field_idx, "field_idx"),
        )

    def pivot_cache_field_add_shared_item_number(self, cache_id: int, field_idx: int, value: float) -> None:
        """Append a numeric shared item to a cache field."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_field_add_shared_item_number(
                h, _uint(cache_id, "cache_id"), _uint(field_idx, "field_idx"), float(value)
            ),
            "fm_workbook_pivot_cache_field_add_shared_item_number",
        )

    def pivot_cache_field_add_shared_item_text(self, cache_id: int, field_idx: int, value: str) -> None:
        """Append a text shared item to a cache field."""
        h = self._require()
        vp, _ = LIB.alloc_utf8(value)
        try:
            _check(
                LIB.fm_workbook_pivot_cache_field_add_shared_item_text(
                    h, _uint(cache_id, "cache_id"), _uint(field_idx, "field_idx"), vp
                ),
                "fm_workbook_pivot_cache_field_add_shared_item_text",
            )
        finally:
            LIB.free(vp)

    def pivot_cache_field_add_shared_item_bool(self, cache_id: int, field_idx: int, value: bool) -> None:
        """Append a boolean shared item to a cache field."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_field_add_shared_item_bool(
                h, _uint(cache_id, "cache_id"), _uint(field_idx, "field_idx"), 1 if value else 0
            ),
            "fm_workbook_pivot_cache_field_add_shared_item_bool",
        )

    def pivot_cache_field_add_shared_item_blank(self, cache_id: int, field_idx: int) -> None:
        """Append a blank shared item to a cache field."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_field_add_shared_item_blank(
                h, _uint(cache_id, "cache_id"), _uint(field_idx, "field_idx")
            ),
            "fm_workbook_pivot_cache_field_add_shared_item_blank",
        )

    def pivot_cache_field_add_shared_item_error(self, cache_id: int, field_idx: int, error_code: int) -> None:
        """Append an Excel error shared item to a cache field."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_field_add_shared_item_error(
                h, _uint(cache_id, "cache_id"), _uint(field_idx, "field_idx"), _sint(error_code, "error")
            ),
            "fm_workbook_pivot_cache_field_add_shared_item_error",
        )

    def pivot_cache_field_clear_shared_items(self, cache_id: int, field_idx: int) -> None:
        """Drop every shared item from a cache field."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_field_clear_shared_items(
                h, _uint(cache_id, "cache_id"), _uint(field_idx, "field_idx")
            ),
            "fm_workbook_pivot_cache_field_clear_shared_items",
        )

    def pivot_cache_record_count(self, cache_id: int) -> int:
        """Return the number of records on the cache."""
        h = self._require()
        return _read_count(LIB.fm_workbook_pivot_cache_record_count, h, _uint(cache_id, "cache_id"))

    def pivot_cache_record_add(self, cache_id: int) -> int:
        """Append an empty record to the cache; return its index."""
        h = self._require()
        out = _alloc_out_ptr()
        try:
            _check(
                LIB.fm_workbook_pivot_cache_record_add(h, _uint(cache_id, "cache_id"), out),
                "fm_workbook_pivot_cache_record_add",
            )
            return LIB.read_u32(out)
        finally:
            LIB.free(out)

    def pivot_cache_record_clear(self, cache_id: int) -> None:
        """Drop every record from the cache."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_record_clear(h, _uint(cache_id, "cache_id")),
            "fm_workbook_pivot_cache_record_clear",
        )

    def pivot_cache_record_set_number(self, cache_id: int, record_idx: int, field_idx: int, value: float) -> None:
        """Set cache cell ``(record_idx, field_idx)`` to a number."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_record_set_number(
                h,
                _uint(cache_id, "cache_id"),
                _uint(record_idx, "record_idx"),
                _uint(field_idx, "field_idx"),
                float(value),
            ),
            "fm_workbook_pivot_cache_record_set_number",
        )

    def pivot_cache_record_set_text(self, cache_id: int, record_idx: int, field_idx: int, value: str) -> None:
        """Set cache cell ``(record_idx, field_idx)`` to text."""
        h = self._require()
        vp, _ = LIB.alloc_utf8(value)
        try:
            _check(
                LIB.fm_workbook_pivot_cache_record_set_text(
                    h, _uint(cache_id, "cache_id"), _uint(record_idx, "record_idx"), _uint(field_idx, "field_idx"), vp
                ),
                "fm_workbook_pivot_cache_record_set_text",
            )
        finally:
            LIB.free(vp)

    def pivot_cache_record_set_bool(self, cache_id: int, record_idx: int, field_idx: int, value: bool) -> None:
        """Set cache cell ``(record_idx, field_idx)`` to a boolean."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_record_set_bool(
                h,
                _uint(cache_id, "cache_id"),
                _uint(record_idx, "record_idx"),
                _uint(field_idx, "field_idx"),
                1 if value else 0,
            ),
            "fm_workbook_pivot_cache_record_set_bool",
        )

    def pivot_cache_record_set_blank(self, cache_id: int, record_idx: int, field_idx: int) -> None:
        """Set cache cell ``(record_idx, field_idx)`` to blank."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_record_set_blank(
                h, _uint(cache_id, "cache_id"), _uint(record_idx, "record_idx"), _uint(field_idx, "field_idx")
            ),
            "fm_workbook_pivot_cache_record_set_blank",
        )

    def pivot_cache_record_set_error(self, cache_id: int, record_idx: int, field_idx: int, error_code: int) -> None:
        """Set cache cell ``(record_idx, field_idx)`` to an Excel error."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_cache_record_set_error(
                h,
                _uint(cache_id, "cache_id"),
                _uint(record_idx, "record_idx"),
                _uint(field_idx, "field_idx"),
                _sint(error_code, "error"),
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
                    _uint(sheet, "sheet_index"),
                    name_ptr,
                    _uint(cache_id, "cache_id"),
                    _uint(anchor_row, "anchor_row"),
                    _uint(anchor_col, "anchor_col"),
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
                LIB.fm_workbook_pivot_get_layout(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), out
                ),
                "fm_workbook_pivot_get_layout",
            )
            return PivotReportLayout(LIB.read_i32(out))
        finally:
            LIB.free(out)

    def set_pivot_report_layout(self, sheet: int, pivot_index: int, layout: "PivotReportLayout | int") -> None:
        """Set the pivot's compact, tabular, or outline report layout."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_set_layout(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _sint(layout, "layout")
            ),
            "fm_workbook_pivot_set_layout",
        )

    def pivot_remove(self, sheet: int, pivot_index: int) -> None:
        """Remove the pivot table at ``pivot_index``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_remove(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")),
            "fm_workbook_pivot_remove",
        )

    def pivot_set_name(self, sheet: int, pivot_index: int, name: str) -> None:
        """Rename the pivot table."""
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(name)
        try:
            _check(
                LIB.fm_workbook_pivot_set_name(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), name_ptr
                ),
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
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(anchor_row, "anchor_row"),
                _uint(anchor_col, "anchor_col"),
                _uint(span_rows, "span_rows"),
                _uint(span_cols, "span_cols"),
            ),
            "fm_workbook_pivot_set_anchor",
        )

    def pivot_set_grand_totals(self, sheet: int, pivot_index: int, rows_enabled: bool, cols_enabled: bool) -> None:
        """Toggle the row / column grand-total bands."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_set_grand_totals(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                1 if rows_enabled else 0,
                1 if cols_enabled else 0,
            ),
            "fm_workbook_pivot_set_grand_totals",
        )

    def pivot_field_count(self, sheet: int, pivot_index: int) -> int:
        """Return the number of fields configured on the pivot."""
        h = self._require()
        return _read_count(
            LIB.fm_workbook_pivot_field_count, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )

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
                    "axis": _sint(spec.axis, "axis"),
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
                LIB.fm_workbook_pivot_field_add(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr, out
                ),
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
            LIB.fm_workbook_pivot_field_clear(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")),
            "fm_workbook_pivot_field_clear",
        )

    def pivot_field_set_axis(self, sheet: int, pivot_index: int, field_idx: int, axis: "PivotAxis | int") -> None:
        """Set the axis of pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_set_axis(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_idx, "field_idx"),
                _sint(axis, "axis"),
            ),
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
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_idx, "field_idx"),
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
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_idx, "field_idx"),
                1 if top else 0,
            ),
            "fm_workbook_pivot_field_set_subtotal_top",
        )

    def pivot_field_add_aggregation(
        self, sheet: int, pivot_index: int, field_idx: int, agg: "PivotAggregation | int"
    ) -> None:
        """Append an aggregation to pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_add_aggregation(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_idx, "field_idx"),
                _sint(agg, "agg"),
            ),
            "fm_workbook_pivot_field_add_aggregation",
        )

    def pivot_field_clear_aggregations(self, sheet: int, pivot_index: int, field_idx: int) -> None:
        """Drop every aggregation from pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_clear_aggregations(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(field_idx, "field_idx")
            ),
            "fm_workbook_pivot_field_clear_aggregations",
        )

    def pivot_field_add_item(self, sheet: int, pivot_index: int, field_idx: int, name: str, visible: bool) -> None:
        """Append a manual-filter item addressed by its label.

        The item carries no cache binding, so records are matched by
        comparing their rendered label against ``name``. The blank item has
        no label of its own; build it with
        :meth:`pivot_field_add_item_at` instead.
        """
        h = self._require()
        name_ptr, _ = LIB.alloc_utf8(name)
        try:
            _check(
                LIB.fm_workbook_pivot_field_add_item(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_idx, "field_idx"),
                    name_ptr,
                    1 if visible else 0,
                ),
                "fm_workbook_pivot_field_add_item",
            )
        finally:
            LIB.free(name_ptr)

    def pivot_field_add_item_at(
        self, sheet: int, pivot_index: int, field_idx: int, cache_index: int, visible: bool
    ) -> None:
        """Append a manual-filter item addressed by its cache shared-item index.

        ``cache_index`` selects ``shared_items[cache_index]`` of the bound
        cache field, the same index space as OOXML ``<item x="N">``. The
        label is resolved from that shared item, so a shared item that
        renders to nothing produces the blank item.
        """
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_add_item_at(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_idx, "field_idx"),
                _uint(cache_index, "cache_index"),
                1 if visible else 0,
            ),
            "fm_workbook_pivot_field_add_item_at",
        )

    def pivot_field_clear_items(self, sheet: int, pivot_index: int, field_idx: int) -> None:
        """Drop every manual-filter item from pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_clear_items(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(field_idx, "field_idx")
            ),
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
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_idx, "field_idx"),
                _uint(item_idx, "item_idx"),
                1 if visible else 0,
            ),
            "fm_workbook_pivot_field_set_item_visible",
        )

    def pivot_field_add_subtotal_fn(
        self, sheet: int, pivot_index: int, field_idx: int, agg: "PivotAggregation | int"
    ) -> None:
        """Append a subtotal-fn entry to pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_add_subtotal_fn(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_idx, "field_idx"),
                _sint(agg, "agg"),
            ),
            "fm_workbook_pivot_field_add_subtotal_fn",
        )

    def pivot_field_clear_subtotal_fns(self, sheet: int, pivot_index: int, field_idx: int) -> None:
        """Drop every subtotal-fn entry from pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_clear_subtotal_fns(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(field_idx, "field_idx")
            ),
            "fm_workbook_pivot_field_clear_subtotal_fns",
        )

    def pivot_field_set_date_group(
        self,
        sheet: int,
        pivot_index: int,
        field_idx: int,
        granularity: "PivotDateGrouping | int",
        calendar: "PivotCalendar | int",
        start_year: int = -1,
        end_year: int = -1,
    ) -> None:
        """Configure date-grouping on pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_set_date_group(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_idx, "field_idx"),
                _sint(granularity, "granularity"),
                _sint(calendar, "calendar"),
                _sint(start_year, "start_year_or_neg1"),
                _sint(end_year, "end_year_or_neg1"),
            ),
            "fm_workbook_pivot_field_set_date_group",
        )

    def pivot_field_clear_date_group(self, sheet: int, pivot_index: int, field_idx: int) -> None:
        """Remove the date-grouping config from pivot field ``field_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_field_clear_date_group(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(field_idx, "field_idx")
            ),
            "fm_workbook_pivot_field_clear_date_group",
        )

    def pivot_field_set_number_format(self, sheet: int, pivot_index: int, field_idx: int, fmt: str) -> None:
        """Set the OOXML number-format string on pivot field ``field_idx``."""
        h = self._require()
        fmt_ptr, _ = LIB.alloc_utf8(fmt)
        try:
            _check(
                LIB.fm_workbook_pivot_field_set_number_format(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_idx, "field_idx"),
                    fmt_ptr,
                ),
                "fm_workbook_pivot_field_set_number_format",
            )
        finally:
            LIB.free(fmt_ptr)

    def pivot_set_row_field_order(self, sheet: int, pivot_index: int, indices: Sequence[int]) -> None:
        """Replace the row-axis field order with ``indices``."""
        self._pivot_set_field_order(
            LIB.fm_workbook_pivot_set_row_field_order,
            _uint(sheet, "sheet_index"),
            _uint(pivot_index, "pivot_index"),
            indices,
        )

    def pivot_set_col_field_order(self, sheet: int, pivot_index: int, indices: Sequence[int]) -> None:
        """Replace the column-axis field order with ``indices``."""
        self._pivot_set_field_order(
            LIB.fm_workbook_pivot_set_col_field_order,
            _uint(sheet, "sheet_index"),
            _uint(pivot_index, "pivot_index"),
            indices,
        )

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
        return _read_count(
            LIB.fm_workbook_pivot_data_field_count, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )

    def pivot_data_field_add(self, sheet: int, pivot_index: int, spec: PivotDataFieldSpec) -> int:
        """Append a data-field entry; return its index."""
        h = self._require()
        out = _alloc_out_ptr()
        owned: List[int] = []
        ptr = self._pack_data_field_spec(spec, owned)
        try:
            _check(
                LIB.fm_workbook_pivot_data_field_add(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr, out
                ),
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
                LIB.fm_workbook_pivot_data_field_set(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(data_field_idx, "data_field_idx"),
                    ptr,
                ),
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
                "field_index": _uint(spec.field_index, "field_index"),
                "aggregation": _sint(spec.aggregation, "aggregation"),
                "show_as": _sint(spec.show_as, "show_as"),
                "show_as_base_field": _sint(spec.show_as_base_field, "show_as_base_field"),
                "show_as_base_item": _sint(spec.show_as_base_item, "show_as_base_item"),
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
            LIB.fm_workbook_pivot_data_field_clear(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")),
            "fm_workbook_pivot_data_field_clear",
        )

    def pivot_filter_count(self, sheet: int, pivot_index: int) -> int:
        """Return the number of active filters on the pivot.

        Counts only the entries this session added; see
        :meth:`pivot_filter_at` for why.
        """
        h = self._require()
        return _read_count(
            LIB.fm_workbook_pivot_filter_count, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )

    def pivot_filter_at(self, sheet: int, pivot_index: int, filter_idx: int) -> PivotFilterSpec:
        """Read the active filter at ``filter_idx`` without mutating it.

        The active-filter list is **session state**. An entry added through
        :meth:`pivot_filter_add` affects evaluation only while this
        workbook handle lives and is not written by :meth:`save`.
        Conversely a ``<filters>`` block Excel wrote is preserved verbatim
        on save but does not appear here, so :meth:`pivot_filter_count`
        reports only what this session added. The filter surface that does
        persist is pivot field item visibility
        (:meth:`pivot_field_set_item_visible`).
        """
        h = self._require()
        ptr = S.alloc_struct(LIB, S.PIVOT_FILTER_SPEC)
        try:
            _check(
                LIB.fm_workbook_pivot_filter_at(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(filter_idx, "filter_idx"),
                    ptr,
                ),
                "fm_workbook_pivot_filter_at",
            )
            d = S.PIVOT_FILTER_SPEC.unpack(LIB, ptr)
            return PivotFilterSpec(
                axis=PivotAxis(d["axis"]),
                field_name=LIB.read_cstr(d["field_name"]),
                type=PivotFilterType(d["type"]),
                value_kind=PivotFilterValueKind(d["value_kind"]),
                value_int=d["value_int"],
                value_double=d["value_double"],
                value_text=LIB.read_cstr(d["value_text"]),
                value_high_kind=PivotFilterValueKind(d["value_high_kind"]),
                value_high_int=d["value_high_int"],
                value_high_double=d["value_high_double"],
                data_field_index=d["data_field_index"],
            )
        finally:
            LIB.free(ptr)

    def pivot_filter_add(self, sheet: int, pivot_index: int, spec: PivotFilterSpec) -> None:
        """Append an active filter to the pivot.

        The entry is session state: it affects evaluation only while this
        workbook handle lives and is not written by :meth:`save`. See
        :meth:`pivot_filter_at`.
        """
        h = self._require()
        owned: List[int] = []
        ptr = S.alloc_struct(LIB, S.PIVOT_FILTER_SPEC)
        try:
            S.PIVOT_FILTER_SPEC.pack(
                LIB,
                ptr,
                {
                    "axis": _sint(spec.axis, "axis"),
                    "type": _sint(spec.type, "type"),
                    "data_field_index": _uint(spec.data_field_index, "data_field_index"),
                    "value_kind": _sint(spec.value_kind, "value_kind"),
                    "value_int": _sint(spec.value_int, "value_int"),
                    "value_double": float(spec.value_double),
                    "value_high_kind": _sint(spec.value_high_kind, "value_high_kind"),
                    "value_high_int": _sint(spec.value_high_int, "value_high_int"),
                    "value_high_double": float(spec.value_high_double),
                },
            )
            S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC, "field_name", spec.field_name, owned)
            S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC, "value_text", spec.value_text, owned)
            _check(
                LIB.fm_workbook_pivot_filter_add(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr
                ),
                "fm_workbook_pivot_filter_add",
            )
        finally:
            LIB.free(ptr)
            for p in owned:
                LIB.free(p)

    def pivot_filter_clear(self, sheet: int, pivot_index: int) -> None:
        """Drop every active filter from the pivot."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_filter_clear(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")),
            "fm_workbook_pivot_filter_clear",
        )

    def pivot_filter_remove_at(self, sheet: int, pivot_index: int, filter_idx: int) -> None:
        """Remove the active filter at ``filter_idx``."""
        h = self._require()
        _check(
            LIB.fm_workbook_pivot_filter_remove_at(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(filter_idx, "filter_idx")
            ),
            "fm_workbook_pivot_filter_remove_at",
        )

    # -- Dependency-graph trace --------------------------------------------
    def precedents(self, sheet: int, row: int, col: int, depth: int = 1) -> List[CellNode]:
        """Return the cells ``(sheet, row, col)`` reads (up to ``depth``)."""
        return self._trace(
            LIB.fm_workbook_precedents,
            _uint(sheet, "sheet"),
            _uint(row, "row"),
            _uint(col, "col"),
            _uint(depth, "depth"),
        )

    def dependents(self, sheet: int, row: int, col: int, depth: int = 1) -> List[CellNode]:
        """Return the cells that read ``(sheet, row, col)`` (up to ``depth``)."""
        return self._trace(
            LIB.fm_workbook_dependents,
            _uint(sheet, "sheet"),
            _uint(row, "row"),
            _uint(col, "col"),
            _uint(depth, "depth"),
        )

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
            # One scratch block for the whole walk; see iter_defined_names.
            nptr = S.alloc_struct(LIB, S.CELL_NODE)
            try:
                for i in range(n):
                    S.zero_struct(LIB, S.CELL_NODE, nptr)
                    _check(LIB.fm_cell_nodes_at(handle, _uint(i, "idx"), nptr), "fm_cell_nodes_at")
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
                LIB.fm_workbook_spill_info(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), ptr),
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
            _check(LIB.fm_function_name_at(_uint(index, "idx"), out), "fm_function_name_at")
            return LIB.read_cstr(LIB.read_u32(out))
        finally:
            LIB.free(out)

    @staticmethod
    def function_metadata(name: str, locale: int = 0) -> Optional[FunctionMetadata]:
        """Return metadata for ``name`` or ``None`` when unknown.

        ``locale`` is ``0`` for ``en-US`` and ``1`` for ``ja-JP``.

        Args:
          name: canonical function name, matched case-insensitively.
          locale: catalog locale ordinal.

        Returns:
          The metadata, or ``None`` when ``name`` matches no registered
          function.

        Raises:
          FormulonError: when ``locale`` is outside the supported range.
            The C ABI reports an invalid locale and an unknown function
            with the same status, so the locale is range-checked here to
            keep ``None`` meaning "unknown function" only -- matching
            ``localize_function_name`` / ``canonicalize_function_name``,
            which raise for the same bad ``locale``.
        """
        if not _LOCALE_MIN <= int(locale) <= _LOCALE_MAX:
            raise FormulonError(
                _STATUS_INVALID_ARGUMENT,
                op="fm_function_metadata",
                _diagnostic_override=("invalid locale", f"locale={int(locale)}"),
            )
        name_ptr, _ = LIB.alloc_utf8(name)
        ptr = S.alloc_struct(LIB, S.FUNCTION_METADATA)
        try:
            status = LIB.fm_function_metadata(name_ptr, _sint(locale, "locale"), ptr)
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
                LIB.fm_function_localize(name_ptr, _sint(locale, "locale"), out),
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
                LIB.fm_function_canonicalize(name_ptr, _sint(locale, "locale"), out),
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
                LIB.fm_workbook_external_link_at(h, _uint(index, "index"), ptr),
                "fm_workbook_external_link_at",
            )
            d = S.EXTERNAL_LINK_RECORD.unpack(LIB, ptr)
            return ExternalLink(
                index=d["index"],
                rel_id=LIB.read_cstr(d["rel_id"]),
                part_path=LIB.read_cstr(d["part_path"]),
                target=LIB.read_cstr(d["target"]),
                target_external=bool(d["target_external"]),
                kind=ExternalLinkKind(d["kind"]),
            )
        finally:
            LIB.free(ptr)

    def get_external_links(self) -> List[ExternalLink]:
        """Return every external-link record in document order."""
        n = self.external_link_count()
        return [self.get_external_link_at(i) for i in range(n)]
