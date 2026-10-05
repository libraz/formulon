# Hand-rolled type stubs for the Formulon public surface.
# IDEs and type-checkers consume this rather than the runtime module.

from enum import IntEnum
from typing import (
    Dict,
    Iterator,
    List,
    Mapping,
    NamedTuple,
    Optional,
    Sequence,
    Tuple,
    TypedDict,
    Union,
)

from ._c import ValueKind as ValueKind

__version__: str

# ---------------------------------------------------------------------------
# Value / iteration
# ---------------------------------------------------------------------------

class Value:
    kind: ValueKind
    number: Optional[float]
    boolean: Optional[bool]
    text: Optional[str]
    error_code: Optional[int]
    def to_python(self) -> Union[None, float, bool, str, "Value"]: ...

class Cell(NamedTuple):
    row: int
    col: int
    formula: Optional[str]
    value: Value

class DefinedName(NamedTuple):
    name: str
    formula: str
    local_sheet_id: int

class Table(NamedTuple):
    name: str
    display_name: str
    ref: str
    sheet_index: int

class PassthroughPart(NamedTuple):
    path: str

class CommentEntry(NamedTuple):
    row: int
    col: int
    author: str
    text: str

class IterativeSettings(NamedTuple):
    enabled: bool
    max_iterations: int
    max_change: float

class PaginationResult:
    page_count: int
    # The declared `_xlnm.Print_Area`; empty when the sheet declares none
    # (not backfilled with the used range, so `page_count` can be non-zero
    # while this is empty).
    print_area: List[tuple[int, int, int, int]]
    horizontal_breaks: List[int]
    vertical_breaks: List[int]
    paper: PaperInfo
    margins: MarginsPt
    printable: RectPt
    scale: float
    page_order: int
    print_titles: PrintTitles
    pages: List[PageLayout]
    horizontal_break_manual: List[bool]
    vertical_break_manual: List[bool]

class RectPt:
    x: float
    y: float
    width: float
    height: float

class PaperInfo:
    width_pt: float
    height_pt: float
    landscape: bool
    known: bool

class MarginsPt:
    left: float
    right: float
    top: float
    bottom: float
    header: float
    footer: float

class PrintTitles:
    has_rows: bool
    first_row: int
    last_row: int
    has_cols: bool
    first_col: int
    last_col: int

class PageLayout:
    area_index: int
    first_row: int
    last_row: int
    first_col: int
    last_col: int
    origin_x_pt: float
    origin_y_pt: float
    width_pt: float
    height_pt: float

class SheetFormatDefaults:
    default_col_width: float
    default_row_height: float
    base_col_width: float
    has_default_col_width: bool
    has_default_row_height: bool
    def __init__(
        self,
        default_col_width: float = ...,
        default_row_height: float = ...,
        base_col_width: float = ...,
        has_default_col_width: bool = ...,
        has_default_row_height: bool = ...,
    ) -> None: ...

class WidthModel:
    points_per_char: float
    padding_pt: float
    normal_font_size: float
    calibrated: bool
    normal_font_name: str
    platform: str

class PageBreak:
    id: int
    min: int
    max: int
    manual: bool

class PageSetup:
    orientation: int
    paper_size: int
    scale: int
    fit_to_width: int
    fit_to_height: int
    fit_to_page: bool
    # Whether the XML states the attribute at all, as opposed to the value
    # field carrying the effective default.
    orientation_stated: bool
    paper_size_stated: bool
    scale_stated: bool
    fit_to_width_stated: bool
    fit_to_height_stated: bool
    fit_to_page_stated: bool

class PageMargins:
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

class FormulonError(Exception):
    status: int
    status_name: str
    message: str
    context: str
    def __init__(self, status: int, *, op: str = ...) -> None: ...

class SaveDiagnostics:
    bytes: bytes
    downgraded_formula_count: int
    deferred_feature_count: int
    dropped_part_count: int
    dropped_relationship_count: int
    renumbered_part_count: int
    def __init__(
        self,
        bytes: bytes,
        downgraded_formula_count: int,
        deferred_feature_count: int,
        dropped_part_count: int,
        dropped_relationship_count: int,
        renumbered_part_count: int,
    ) -> None: ...

class ReadDiagnostics:
    undecoded_formula_count: int
    undecoded_defined_name_count: int
    undecoded_part_count: int
    skipped_feature_count: int
    unknown_content_type_count: int
    def __init__(
        self,
        undecoded_formula_count: int,
        undecoded_defined_name_count: int,
        undecoded_part_count: int,
        skipped_feature_count: int,
        unknown_content_type_count: int,
    ) -> None: ...

# ---------------------------------------------------------------------------
# Enumerations
# ---------------------------------------------------------------------------

class GeometryMode(IntEnum):
    DISPLAY = 0
    PRINT = 1

class DisplayStatus(IntEnum):
    OK = 0
    OVERFLOW = 1
    INVALID_FORMAT = 2

class FilterKind(IntEnum):
    NONE = 0
    VALUES = 1
    CUSTOM = 2
    TOP10 = 3
    DYNAMIC = 4
    COLOR = 5
    ICON = 6

class FilterOperator(IntEnum):
    EQUAL = 0
    LESS_THAN = 1
    LESS_THAN_OR_EQUAL = 2
    NOT_EQUAL = 3
    GREATER_THAN_OR_EQUAL = 4
    GREATER_THAN = 5

class DynamicFilterType(IntEnum):
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
    VALUE = 0
    CELL_COLOR = 1
    FONT_COLOR = 2
    ICON = 3

class SortMethod(IntEnum):
    NONE = 0
    PIN_YIN = 1
    STROKE = 2

class DateTimeGrouping(IntEnum):
    YEAR = 0
    MONTH = 1
    DAY = 2
    HOUR = 3
    MINUTE = 4
    SECOND = 5

class ValidationErrorStyle(IntEnum):
    STOP = 0
    WARNING = 1
    INFORMATION = 2

class ColorContext(IntEnum):
    FONT = 0
    FILL_FOREGROUND = 1
    FILL_BACKGROUND = 2
    BORDER = 3

class ImageFormat(IntEnum):
    UNKNOWN = 0
    PNG = 1
    JPEG = 2
    GIF = 3
    BMP = 4

class DrawingObjectKind(IntEnum):
    PICTURE = 0
    SHAPE = 1
    CHART = 2
    GROUP = 3
    CONNECTOR = 4
    GRAPHIC_FRAME = 5
    OTHER = 6

class AnchorKind(IntEnum):
    ONE_CELL = 0
    TWO_CELL = 1
    ABSOLUTE = 2

class AnchorEditAs(IntEnum):
    TWO_CELL = 0
    ONE_CELL = 1
    ABSOLUTE = 2

class ColorResolution(IntEnum):
    EXACT = 0
    DEFAULT_THEME = 1
    INDEX_OUT_OF_RANGE = 2
    THEME_UNPARSEABLE = 3
    AUTO_CONTEXT = 4

class ThemeSource(IntEnum):
    PART = 0
    DEFAULT = 1
    UNPARSEABLE = 2

class EffectiveStyleSource(IntEnum):
    CELL = 0
    ROW = 1
    COLUMN = 2
    DEFAULT = 3

class CalcMode(IntEnum):
    AUTO = 0
    MANUAL = 1
    AUTO_NO_TABLE = 2

class SheetVisibility(IntEnum):
    VISIBLE = 0
    HIDDEN = 1
    VERY_HIDDEN = 2

class LogLevel(IntEnum):
    DEBUG = 0
    INFO = 1
    WARN = 2
    ERROR = 3
    OFF = 4

class ErrorCode(IntEnum):
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
    UNKNOWN = 0
    EXTERNAL_BOOK = 1
    OLE = 2
    DDE = 3

class PivotAxis(IntEnum):
    ROW = 0
    COL = 1
    VALUE = 2
    PAGE = 3

class WorkbookFormat(IntEnum):
    UNKNOWN = 0
    XLSX = 1
    XLSB = 2

class PivotAggregation(IntEnum):
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
    VALUE_TOP_10 = 0
    VALUE_GREATER_THAN = 1
    VALUE_BETWEEN = 2
    LABEL_CONTAINS = 3
    LABEL_BEGINS_WITH = 4
    LABEL_DATE = 5

class PivotFilterValueKind(IntEnum):
    NONE = -1
    INT = 0
    DOUBLE = 1
    TEXT = 2

class PivotDateGrouping(IntEnum):
    DAY = 0
    MONTH = 1
    QUARTER = 2
    YEAR = 3
    DAYS = 4
    HOUR = 5
    MINUTE = 6
    SECOND = 7

class PivotCalendar(IntEnum):
    GREGORIAN = 0
    JAPANESE = 1

class PivotCellKind(IntEnum):
    HEADER = 0
    ROW_LABEL = 1
    COL_LABEL = 2
    DATA = 3
    ROW_SUBTOTAL = 4
    COL_SUBTOTAL = 5
    GRAND_TOTAL = 6
    BLANK = 7

# ---------------------------------------------------------------------------
# Structured value / input dataclasses
# ---------------------------------------------------------------------------

class MergeRange:
    first_row: int
    first_col: int
    last_row: int
    last_col: int
    def __init__(self, first_row: int, first_col: int, last_row: int, last_col: int) -> None: ...

class Hyperlink:
    row: int
    col: int
    last_row: int
    last_col: int
    target: str
    location: str
    display: str
    tooltip: str

class Comment:
    author: str
    text: str

class DataValidation:
    ranges: List[MergeRange]
    type: int
    op: int
    error_style: Union[ValidationErrorStyle, int]
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

class AutoFilterDateGroup:
    year: int
    month: int
    day: int
    hour: int
    minute: int
    second: int
    grouping: Union[DateTimeGrouping, int]
    def __init__(
        self,
        year: int = ...,
        month: int = ...,
        day: int = ...,
        hour: int = ...,
        minute: int = ...,
        second: int = ...,
        grouping: Union[DateTimeGrouping, int] = ...,
    ) -> None: ...

class AutoFilterColumn:
    col_id: int
    hidden_button: bool
    show_button: bool
    kind: Union[FilterKind, int]
    filter_blank: bool
    values: List[str]
    date_groups: List[AutoFilterDateGroup]
    custom_and: bool
    custom_count: int
    op1: Union[FilterOperator, int]
    val1: str
    op2: Union[FilterOperator, int]
    val2: str
    top: bool
    percent: bool
    has_filter_val: bool
    top_val: float
    filter_val: float
    dynamic_type: Union[DynamicFilterType, int]
    has_dyn_val: bool
    has_dyn_max_val: bool
    dyn_val: float
    dyn_max_val: float
    val_iso: str
    max_val_iso: str
    dxf_id: int
    cell_color: bool
    icon_set: int
    icon_id: int
    has_icon_id: bool
    def __init__(
        self,
        col_id: int,
        hidden_button: bool = ...,
        show_button: bool = ...,
        kind: Union[FilterKind, int] = ...,
        filter_blank: bool = ...,
        values: List[str] = ...,
        date_groups: List[AutoFilterDateGroup] = ...,
        custom_and: bool = ...,
        custom_count: int = ...,
        op1: Union[FilterOperator, int] = ...,
        val1: str = ...,
        op2: Union[FilterOperator, int] = ...,
        val2: str = ...,
        top: bool = ...,
        percent: bool = ...,
        has_filter_val: bool = ...,
        top_val: float = ...,
        filter_val: float = ...,
        dynamic_type: Union[DynamicFilterType, int] = ...,
        has_dyn_val: bool = ...,
        has_dyn_max_val: bool = ...,
        dyn_val: float = ...,
        dyn_max_val: float = ...,
        val_iso: str = ...,
        max_val_iso: str = ...,
        dxf_id: int = ...,
        cell_color: bool = ...,
        icon_set: int = ...,
        icon_id: int = ...,
        has_icon_id: bool = ...,
    ) -> None: ...

class AutoFilterSortCondition:
    ref: MergeRange
    descending: bool
    sort_by: Union[SortBy, int]
    custom_list: str
    dxf_id: int
    has_dxf_id: bool
    icon_set: int
    icon_id: int
    has_icon_id: bool
    def __init__(
        self,
        ref: MergeRange,
        descending: bool = ...,
        sort_by: Union[SortBy, int] = ...,
        custom_list: str = ...,
        dxf_id: int = ...,
        has_dxf_id: bool = ...,
        icon_set: int = ...,
        icon_id: int = ...,
        has_icon_id: bool = ...,
    ) -> None: ...

class AutoFilterSortState:
    ref: MergeRange
    column_sort: bool
    case_sensitive: bool
    sort_method: Union[SortMethod, int]
    conditions: List[AutoFilterSortCondition]
    def __init__(
        self,
        ref: MergeRange,
        column_sort: bool = ...,
        case_sensitive: bool = ...,
        sort_method: Union[SortMethod, int] = ...,
        conditions: List[AutoFilterSortCondition] = ...,
    ) -> None: ...

class AutoFilter:
    range: MergeRange
    columns: List[AutoFilterColumn]
    sort: Optional[AutoFilterSortState]
    def __init__(
        self,
        range: MergeRange,
        columns: List[AutoFilterColumn] = ...,
        sort: Optional[AutoFilterSortState] = ...,
    ) -> None: ...

class ValidationOutcome:
    has_rule: bool
    valid: bool
    rule_index: int
    error_style: Union[ValidationErrorStyle, int]
    def __init__(
        self,
        has_rule: bool,
        valid: bool,
        rule_index: int,
        error_style: Union[ValidationErrorStyle, int],
    ) -> None: ...

class Mention:
    person_id: str
    mention_id: str
    start: int
    length: int
    def __init__(
        self,
        person_id: str,
        mention_id: str,
        start: int,
        length: int,
    ) -> None: ...

class ThreadedComment:
    id: str
    row: int
    col: int
    person_id: str
    created: str
    text: str
    parent_id: str
    done: bool
    mentions: List[Mention]
    def __init__(
        self,
        id: str,
        row: int = ...,
        col: int = ...,
        person_id: str = ...,
        created: str = ...,
        text: str = ...,
        parent_id: str = ...,
        done: bool = ...,
        mentions: List[Mention] = ...,
    ) -> None: ...

class ImageInfo:
    format: Union[ImageFormat, int]
    px_width: int
    px_height: int
    def __init__(self, format: Union[ImageFormat, int], px_width: int, px_height: int) -> None: ...

class DrawingObject:
    object_id: int
    kind: Union[DrawingObjectKind, int]
    anchor_kind: Union[AnchorKind, int]
    edit_as: Union[AnchorEditAs, int]
    from_row: int
    from_col: int
    from_row_off: int
    from_col_off: int
    to_row: int
    to_col: int
    to_row_off: int
    to_col_off: int
    cx: int
    cy: int
    image_format: Union[ImageFormat, int]
    name: str
    descr: str
    media_path: str
    def __init__(
        self,
        object_id: int,
        kind: Union[DrawingObjectKind, int],
        anchor_kind: Union[AnchorKind, int],
        edit_as: Union[AnchorEditAs, int],
        from_row: int,
        from_col: int,
        from_row_off: int,
        from_col_off: int,
        to_row: int,
        to_col: int,
        to_row_off: int,
        to_col_off: int,
        cx: int,
        cy: int,
        image_format: Union[ImageFormat, int],
        name: str,
        descr: str,
        media_path: str,
    ) -> None: ...

class Person:
    id: str
    display_name: str
    user_id: str
    provider_id: str
    def __init__(
        self,
        id: str,
        display_name: str,
        user_id: str = ...,
        provider_id: str = ...,
    ) -> None: ...

class DataValidationInput:
    type: int
    ranges: List[MergeRange]
    op: int
    error_style: Union[ValidationErrorStyle, int]
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
    def __init__(
        self,
        type: int,
        ranges: List[MergeRange] = ...,
        op: int = ...,
        error_style: Union[ValidationErrorStyle, int] = ...,
        allow_blank: bool = ...,
        show_input_message: bool = ...,
        show_error_message: bool = ...,
        show_dropdown: bool = ...,
        formula1: str = ...,
        formula2: str = ...,
        error_title: str = ...,
        error_message: str = ...,
        prompt_title: str = ...,
        prompt_message: str = ...,
    ) -> None: ...

class SheetProtection:
    enabled: bool
    algorithm_name: str
    hash_value: str
    salt_value: str
    spin_count: int
    legacy_password: str
    sheet: bool
    objects: bool
    scenarios: bool
    format_cells: bool
    format_columns: bool
    format_rows: bool
    insert_columns: bool
    insert_rows: bool
    insert_hyperlinks: bool
    delete_columns: bool
    delete_rows: bool
    select_locked_cells: bool
    select_unlocked_cells: bool
    sort: bool
    auto_filter: bool
    pivot_tables: bool
    def __init__(
        self,
        enabled: bool = ...,
        algorithm_name: str = ...,
        hash_value: str = ...,
        salt_value: str = ...,
        spin_count: int = ...,
        legacy_password: str = ...,
        sheet: bool = ...,
        objects: bool = ...,
        scenarios: bool = ...,
        format_cells: bool = ...,
        format_columns: bool = ...,
        format_rows: bool = ...,
        insert_columns: bool = ...,
        insert_rows: bool = ...,
        insert_hyperlinks: bool = ...,
        delete_columns: bool = ...,
        delete_rows: bool = ...,
        select_locked_cells: bool = ...,
        select_unlocked_cells: bool = ...,
        sort: bool = ...,
        auto_filter: bool = ...,
        pivot_tables: bool = ...,
    ) -> None: ...

class CivilTime:
    year: int
    month: int
    day: int
    hour: int
    minute: int
    second: int

class SheetView:
    zoom_scale: int
    freeze_rows: int
    freeze_cols: int
    tab_hidden: bool
    visibility: SheetVisibility
    show_grid_lines: bool
    show_row_col_headers: bool
    show_zeros: bool
    right_to_left: bool
    tab_selected: bool
    view_mode: str

class ColumnLayout:
    first: int
    last: int
    width: float
    hidden: bool
    outline_level: int
    # True when width is logically explicit, including legacy non-zero widths.
    has_width: bool
    has_style: bool
    style_xf: int

class RowLayout:
    row: int
    height: float
    hidden: bool
    outline_level: int
    has_style: bool
    style_xf: int
    has_height: bool
    custom_height: bool

class CfMatch:
    kind: int
    priority: int
    dxf_id_engaged: bool
    dxf_id: int
    color: Tuple[int, int, int, int]
    bar_length_pct: float
    bar_axis_position_pct: float
    bar_is_negative: bool
    bar_fill: Tuple[int, int, int, int]
    bar_border_engaged: bool
    bar_border: Tuple[int, int, int, int]
    bar_gradient: bool
    icon_set_name: int
    icon_index: int
    # The rule's direction: 0 = context, 1 = left to right, 2 = right to left.
    bar_direction: int

class CfCellResult:
    row: int
    col: int
    matches: List[CfMatch]

class CfColor:
    r: int
    g: int
    b: int
    a: int
    def __init__(self, r: int, g: int, b: int, a: int = ...) -> None: ...

class CfValueObject:
    type: int
    value: Optional[str]
    gte: bool
    def __init__(self, type: int, value: Optional[str] = ..., gte: bool = ...) -> None: ...

class ColorScale:
    thresholds: List[CfValueObject]
    colors: List[CfColor]
    def __init__(self, thresholds: List[CfValueObject], colors: List[CfColor]) -> None: ...

class DataBar:
    minimum: CfValueObject
    maximum: CfValueObject
    fill: CfColor
    show_value: bool
    min_length_pct: int
    max_length_pct: int
    # x14 extension. `None` keeps the model default: gradient fill on,
    # automatic axis, negative fill equal to `fill`, no border, black axis.
    # `axis_position` is 0 = automatic, 1 = middle, 2 = none.
    gradient: Optional[bool]
    axis_position: Optional[int]
    negative_fill: Optional[CfColor]
    border: Optional[CfColor]
    negative_border: Optional[CfColor]
    axis_color: Optional[CfColor]
    # Edge the bar grows from: 0 = context (default), 1 = left to right,
    # 2 = right to left.
    direction: int
    def __init__(
        self,
        minimum: CfValueObject,
        maximum: CfValueObject,
        fill: CfColor,
        show_value: bool = ...,
        min_length_pct: int = ...,
        max_length_pct: int = ...,
        gradient: Optional[bool] = ...,
        axis_position: Optional[int] = ...,
        negative_fill: Optional[CfColor] = ...,
        border: Optional[CfColor] = ...,
        negative_border: Optional[CfColor] = ...,
        axis_color: Optional[CfColor] = ...,
        direction: int = ...,
    ) -> None: ...

class IconSet:
    name: int
    thresholds: List[CfValueObject]
    reverse: bool
    show_value: bool
    # Round-trip only: preserved across load and save but never consulted
    # during evaluation; each threshold's own `type` is authoritative.
    percent: bool
    # Lower bound of the lowest icon's bucket; a cell below it gets no icon.
    # `None` means Excel's default, `percent 0`.
    floor: Optional[CfValueObject]
    def __init__(
        self,
        name: int,
        thresholds: List[CfValueObject],
        reverse: bool = ...,
        show_value: bool = ...,
        percent: bool = ...,
        floor: Optional[CfValueObject] = ...,
    ) -> None: ...

class ConditionalFormat:
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

class ConditionalFormatInput:
    sqref: List[MergeRange]
    type: int
    priority: int
    stop_if_true: bool
    id: str
    dxf_id_engaged: bool
    dxf_id: int
    formula1: str
    formula2: str
    op_engaged: bool
    op: int
    rank_engaged: bool
    rank: int
    percent: bool
    bottom: bool
    above_average: bool
    equal_average: bool
    std_dev_engaged: bool
    std_dev: float
    text: str
    time_period_engaged: bool
    time_period: int
    color_scale: Optional[ColorScale]
    data_bar: Optional[DataBar]
    icon_set: Optional[IconSet]
    def __init__(
        self,
        sqref: List[MergeRange],
        type: int,
        priority: int = ...,
        stop_if_true: bool = ...,
        id: str = ...,
        dxf_id_engaged: bool = ...,
        dxf_id: int = ...,
        formula1: str = ...,
        formula2: str = ...,
        op_engaged: bool = ...,
        op: int = ...,
        rank_engaged: bool = ...,
        rank: int = ...,
        percent: bool = ...,
        bottom: bool = ...,
        above_average: bool = ...,
        equal_average: bool = ...,
        std_dev_engaged: bool = ...,
        std_dev: float = ...,
        text: str = ...,
        time_period_engaged: bool = ...,
        time_period: int = ...,
        color_scale: Optional[ColorScale] = ...,
        data_bar: Optional[DataBar] = ...,
        icon_set: Optional[IconSet] = ...,
    ) -> None: ...

class CellNode:
    sheet: int
    row: int
    col: int

class SpillInfo:
    engaged: bool
    anchor_row: int
    anchor_col: int
    rows: int
    cols: int

class FunctionMetadata:
    name: str
    min_arity: int
    max_arity: Optional[int]
    availability: int
    signature_template: Optional[str]
    description: Optional[str]

class FunctionMetadataLocalized(TypedDict, total=False):
    signature: str
    description: str

class FunctionMetadataEntry(TypedDict, total=False):
    signature: str
    description: str
    aliases: Dict[str, str]
    localized: Dict[str, FunctionMetadataLocalized]

FunctionMetadataProvider = Dict[str, FunctionMetadataEntry]

class MergedFunctionMetadata:
    name: str
    min_arity: int
    max_arity: Optional[int]
    availability: int
    signature_template: Optional[str]
    description: Optional[str]
    localized_name: str

def merge_function_metadata(
    base: FunctionMetadata,
    entry: Optional[Mapping[str, object]],
    locale: str,
) -> Union[FunctionMetadata, MergedFunctionMetadata]: ...

class CellXf:
    font_index: int
    fill_index: int
    border_index: int
    num_fmt_id: int
    horizontal_align: int
    vertical_align: int
    wrap_text: bool
    has_alignment: Optional[bool]
    justify_last_line: bool
    xf_id: int
    text_rotation: Optional[int]
    indent: Optional[int]
    relative_indent: Optional[int]
    shrink_to_fit: Optional[bool]
    reading_order: Optional[int]
    has_horizontal_align: Optional[bool]
    has_vertical_align: Optional[bool]
    has_wrap_text: Optional[bool]
    has_justify_last_line: Optional[bool]
    apply_number_format: bool
    apply_font: bool
    apply_fill: bool
    apply_border: bool
    apply_alignment: bool
    apply_protection: bool
    quote_prefix: bool
    has_protection: bool
    locked: bool
    hidden: bool
    def __init__(
        self,
        font_index: int,
        fill_index: int,
        border_index: int,
        num_fmt_id: int,
        horizontal_align: int,
        vertical_align: int,
        wrap_text: bool,
        has_alignment: Optional[bool] = ...,
        justify_last_line: bool = ...,
        xf_id: int = ...,
        text_rotation: Optional[int] = ...,
        indent: Optional[int] = ...,
        relative_indent: Optional[int] = ...,
        shrink_to_fit: Optional[bool] = ...,
        reading_order: Optional[int] = ...,
        has_horizontal_align: Optional[bool] = ...,
        has_vertical_align: Optional[bool] = ...,
        has_wrap_text: Optional[bool] = ...,
        has_justify_last_line: Optional[bool] = ...,
        apply_number_format: bool = ...,
        apply_font: bool = ...,
        apply_fill: bool = ...,
        apply_border: bool = ...,
        apply_alignment: bool = ...,
        apply_protection: bool = ...,
        quote_prefix: bool = ...,
        has_protection: bool = ...,
        locked: bool = ...,
        hidden: bool = ...,
    ) -> None: ...

class ColorSpec:
    kind: int
    rgb: int
    theme: int
    tint: float
    indexed: int
    def __init__(
        self,
        kind: int = ...,
        rgb: int = ...,
        theme: int = ...,
        tint: float = ...,
        indexed: int = ...,
    ) -> None: ...

class PhoneticRun:
    sb: int
    eb: int
    text: str
    def __init__(self, sb: int = ..., eb: int = ..., text: str = ...) -> None: ...

class PhoneticProperties:
    font_id: int
    type: int
    alignment: int
    def __init__(self, font_id: int = ..., type: int = ..., alignment: int = ...) -> None: ...

class FontRecord:
    name: str
    size: float
    color_argb: int
    bold: bool
    italic: bool
    strike: bool
    has_bold: bool
    has_italic: bool
    has_strike: bool
    underline: int
    vert_align: int
    has_family: bool
    family: int
    has_charset: bool
    charset: int
    scheme: int
    color: ColorSpec
    def __init__(
        self,
        name: str = ...,
        size: float = ...,
        color_argb: int = ...,
        bold: bool = ...,
        italic: bool = ...,
        strike: bool = ...,
        has_bold: bool = ...,
        has_italic: bool = ...,
        has_strike: bool = ...,
        underline: int = ...,
        vert_align: int = ...,
        has_family: bool = ...,
        family: int = ...,
        has_charset: bool = ...,
        charset: int = ...,
        scheme: int = ...,
        color: ColorSpec = ...,
    ) -> None: ...

class FillRecord:
    pattern: int
    fg_argb: int
    bg_argb: int
    fg: ColorSpec
    bg: ColorSpec
    def __init__(
        self,
        pattern: int = ...,
        fg_argb: int = ...,
        bg_argb: int = ...,
        fg: ColorSpec = ...,
        bg: ColorSpec = ...,
    ) -> None: ...

class DifferentialFormat:
    font: Optional[FontRecord]
    fill: Optional[FillRecord]
    border: Optional[Dict[str, object]]
    num_fmt_id: Optional[int]
    num_fmt_code: str
    alignment_xml: str
    protection_xml: str
    def __init__(
        self,
        font: Optional[FontRecord] = ...,
        fill: Optional[FillRecord] = ...,
        border: Optional[Dict[str, object]] = ...,
        num_fmt_id: Optional[int] = ...,
        num_fmt_code: str = ...,
        alignment_xml: str = ...,
        protection_xml: str = ...,
    ) -> None: ...

class StyleBatchIndices(NamedTuple):
    fonts: List[int]
    fills: List[int]
    borders: List[int]
    cell_xfs: List[int]
    num_fmts: List[int]

class CellStyle:
    name: str
    xf_id: int
    builtin_id: int
    i_level: int
    hidden: bool
    custom_builtin: bool
    def __init__(
        self,
        name: str,
        xf_id: int,
        builtin_id: int = ...,
        i_level: int = ...,
        hidden: bool = ...,
        custom_builtin: bool = ...,
    ) -> None: ...

class ThemeFonts:
    major_latin: str
    major_east_asian: str
    minor_latin: str
    minor_east_asian: str
    def __init__(
        self,
        major_latin: str = ...,
        major_east_asian: str = ...,
        minor_latin: str = ...,
        minor_east_asian: str = ...,
    ) -> None: ...

class Theme:
    source: ThemeSource
    colors: List[int]
    fonts: ThemeFonts
    def __init__(self, source: ThemeSource, colors: List[int], fonts: ThemeFonts) -> None: ...

class ResolvedColor:
    argb: int
    resolution: ColorResolution
    def __init__(self, argb: int, resolution: ColorResolution) -> None: ...

class EffectiveStyle:
    xf_index: int
    source: EffectiveStyleSource
    font_index: int
    fill_index: int
    border_index: int
    font: ResolvedColor
    fill_foreground: ResolvedColor
    fill_background: ResolvedColor
    borders: List[ResolvedColor]
    locked: bool
    hidden: bool
    num_fmt_code: str
    def __init__(
        self,
        xf_index: int,
        source: EffectiveStyleSource,
        font_index: int,
        fill_index: int,
        border_index: int,
        font: ResolvedColor,
        fill_foreground: ResolvedColor,
        fill_background: ResolvedColor,
        borders: List[ResolvedColor],
        locked: bool,
        hidden: bool,
        num_fmt_code: str,
    ) -> None: ...

class ExternalLink:
    index: int
    rel_id: str
    part_path: str
    target: str
    target_external: bool
    kind: ExternalLinkKind

class PivotCell:
    row: int
    col: int
    value: Value
    kind: PivotCellKind
    depth: int
    field_name: str
    number_format: str

class PivotLayout:
    top: int
    left: int
    rows: int
    cols: int
    cells: List[PivotCell]

class PivotReportLayout(IntEnum):
    COMPACT: int
    TABULAR: int
    OUTLINE: int

class PivotWorksheetSource:
    ref: Optional[str]
    sheet: Optional[str]
    name: Optional[str]
    def __init__(self, ref: Optional[str] = ..., sheet: Optional[str] = ..., name: Optional[str] = ...) -> None: ...

class PivotFieldSpec:
    source_name: str
    custom_name: str
    axis: Union[PivotAxis, int]
    subtotal_top: bool
    number_format: str
    def __init__(
        self,
        source_name: str,
        custom_name: str = ...,
        axis: Union[PivotAxis, int] = ...,
        subtotal_top: bool = ...,
        number_format: str = ...,
    ) -> None: ...

class PivotDataFieldSpec:
    name: str
    field_index: int
    aggregation: Union[PivotAggregation, int]
    number_format: str
    show_as: Union[PivotShowValuesAs, int]
    show_as_base_field: int
    show_as_base_item: int
    def __init__(
        self,
        name: str,
        field_index: int,
        aggregation: Union[PivotAggregation, int] = ...,
        number_format: str = ...,
        show_as: Union[PivotShowValuesAs, int] = ...,
        show_as_base_field: int = ...,
        show_as_base_item: int = ...,
    ) -> None: ...

class PivotFilterSpec:
    axis: Union[PivotAxis, int]
    field_name: str
    type: Union[PivotFilterType, int]
    value_kind: Union[PivotFilterValueKind, int]
    value_int: int
    value_double: float
    value_text: str
    value_high_kind: Union[PivotFilterValueKind, int]
    value_high_int: int
    value_high_double: float
    data_field_index: int
    def __init__(
        self,
        axis: Union[PivotAxis, int],
        field_name: str,
        type: Union[PivotFilterType, int],
        value_kind: Union[PivotFilterValueKind, int] = ...,
        value_int: int = ...,
        value_double: float = ...,
        value_text: str = ...,
        value_high_kind: Union[PivotFilterValueKind, int] = ...,
        value_high_int: int = ...,
        value_high_double: float = ...,
        data_field_index: int = ...,
    ) -> None: ...

# ---------------------------------------------------------------------------
# Workbook
# ---------------------------------------------------------------------------

class Workbook:
    def __init__(self) -> None: ...
    @classmethod
    def create_default(cls) -> "Workbook": ...
    @classmethod
    def create_empty(cls) -> "Workbook": ...
    @classmethod
    def load(cls, data: Union[bytes, bytearray, memoryview]) -> "Workbook": ...
    def __enter__(self) -> "Workbook": ...
    def __exit__(self, exc_type: object, exc: object, tb: object) -> None: ...
    def close(self) -> None: ...
    @property
    def is_valid(self) -> bool: ...
    def memory_usage(self) -> int: ...

    # Sheets.
    def sheet_count(self) -> int: ...
    def sheet_name(self, index: int) -> str: ...
    def add_sheet(self, name: str) -> None: ...
    def move_sheet(self, from_index: int, to_index: int) -> None: ...
    def remove_sheet(self, index: int) -> None: ...
    def rename_sheet(self, index: int, new_name: str) -> None: ...

    # Cell mutation / read.
    def set_number(self, sheet: int, row: int, col: int, value: float) -> None: ...
    def set_bool(self, sheet: int, row: int, col: int, value: bool) -> None: ...
    def set_error(self, sheet: int, row: int, col: int, error_code: int) -> None: ...
    def set_text(self, sheet: int, row: int, col: int, value: str) -> None: ...
    def set_blank(self, sheet: int, row: int, col: int) -> None: ...
    def set_formula(self, sheet: int, row: int, col: int, formula: str) -> None: ...
    def set_phonetic(self, sheet: int, row: int, col: int, text: str) -> None: ...
    def set_phonetic_runs(self, sheet: int, row: int, col: int, runs: Sequence[PhoneticRun]) -> None: ...
    def get_phonetic(self, sheet: int, row: int, col: int) -> str: ...
    def get_phonetic_runs(self, sheet: int, row: int, col: int) -> List[PhoneticRun]: ...
    def set_phonetic_properties(self, sheet: int, row: int, col: int, properties: PhoneticProperties) -> None: ...
    def get_phonetic_properties(self, sheet: int, row: int, col: int) -> PhoneticProperties: ...
    def get_value(self, sheet: int, row: int, col: int) -> Value: ...
    def evaluate_formula_array(self, sheet: int, row: int, col: int, formula: str) -> List[List[Value]]: ...
    def lambda_text_at(self, sheet: int, row: int, col: int) -> str: ...

    # Defined names.
    def set_defined_name(self, name: str, formula: str) -> None: ...
    def set_defined_name_scoped(self, name: str, formula: str, local_sheet_id: int) -> None: ...

    # Row / column structural edits.
    def insert_rows(self, sheet: int, row: int, count: int) -> None: ...
    def delete_rows(self, sheet: int, row: int, count: int) -> None: ...
    def insert_cols(self, sheet: int, col: int, count: int) -> None: ...
    def delete_cols(self, sheet: int, col: int, count: int) -> None: ...

    # Recalc + calc policy / profile.
    def recalc(self) -> None: ...
    def set_iterative(self, enabled: bool, max_iterations: int, max_change: float) -> None: ...
    def set_iterative_enabled(self, enabled: bool) -> None: ...
    def get_iterative(self) -> IterativeSettings: ...
    def partial_recalc(
        self,
        sheet: int,
        first_row: int,
        last_row: int,
        first_col: int,
        last_col: int,
    ) -> int: ...
    def calc_mode(self) -> CalcMode: ...
    def set_calc_mode(self, mode: Union[CalcMode, int]) -> None: ...
    def pinned_now(self) -> Optional[CivilTime]: ...
    def set_pinned_now(
        self,
        year: int,
        month: int,
        day: int,
        hour: int = ...,
        minute: int = ...,
        second: int = ...,
    ) -> None: ...
    def clear_pinned_now(self) -> None: ...
    def excel_profile_id(self) -> str: ...
    def set_excel_profile_id(self, profile_id: str) -> None: ...

    # Save.
    def save(self) -> bytes: ...
    def save_as(self, fmt: Union[WorkbookFormat, int]) -> bytes: ...
    def save_with_diagnostics(self, fmt: Union[WorkbookFormat, int]) -> SaveDiagnostics: ...
    def read_diagnostics(self) -> ReadDiagnostics: ...

    # Iteration.
    def iter_cells(self, sheet: int) -> Iterator[Cell]: ...
    def iter_defined_names(self) -> Iterator[DefinedName]: ...
    def iter_tables(self) -> Iterator[Table]: ...
    def iter_passthrough(self) -> Iterator[PassthroughPart]: ...
    def table_create(
        self,
        sheet: int,
        ref: str,
        name: str,
        display_name: str,
        column_names: Sequence[str],
        style_name: str = ...,
        header_row: bool = ...,
        totals_row: bool = ...,
    ) -> int: ...
    def table_update(
        self,
        index: int,
        ref: str,
        style_name: Optional[str] = ...,
        header_row: Optional[bool] = ...,
        totals_row: Optional[bool] = ...,
    ) -> None: ...
    def table_remove(self, index: int) -> None: ...

    # Merges.
    def add_merge(self, sheet: int, merge: MergeRange) -> None: ...
    def remove_merge(self, sheet: int, merge: MergeRange) -> None: ...
    def remove_merge_at(self, sheet: int, index: int) -> None: ...
    def clear_merges(self, sheet: int) -> None: ...
    def merge_count(self, sheet: int) -> int: ...
    def get_merges(self, sheet: int) -> List[MergeRange]: ...

    # Hyperlinks.
    def add_hyperlink(
        self,
        sheet: int,
        row: int,
        col: int,
        target: str,
        display: str = ...,
        tooltip: str = ...,
        location: str = ...,
    ) -> None: ...
    def add_hyperlink_range(
        self,
        sheet: int,
        row: int,
        col: int,
        last_row: int,
        last_col: int,
        target: str,
        display: str = ...,
        tooltip: str = ...,
        location: str = ...,
    ) -> None: ...
    def remove_hyperlink(self, sheet: int, row: int, col: int) -> None: ...
    def remove_hyperlink_at(self, sheet: int, index: int) -> None: ...
    def clear_hyperlinks(self, sheet: int) -> None: ...
    def hyperlink_count(self, sheet: int) -> int: ...
    def get_hyperlinks(self, sheet: int) -> List[Hyperlink]: ...

    # Comments.
    def get_comment(self, sheet: int, row: int, col: int) -> Optional[Comment]: ...
    def set_comment(self, sheet: int, row: int, col: int, author: str, text: str) -> None: ...
    def comment_count(self, sheet: int) -> int: ...
    def get_comments(self, sheet: int) -> List[CommentEntry]: ...

    # Ad-hoc conditional-format evaluation.
    def evaluate_cf_formula(
        self, sheet: int, row: int, col: int, anchor_row: int, anchor_col: int, formula: str
    ) -> bool: ...

    # Data validations.
    def validation_count(self, sheet: int) -> int: ...
    def get_validation_at(self, sheet: int, index: int) -> DataValidation: ...
    def get_validations(self, sheet: int) -> List[DataValidation]: ...
    def add_validation(self, sheet: int, validation: DataValidationInput) -> None: ...
    def remove_validation_at(self, sheet: int, index: int) -> None: ...
    def clear_validations(self, sheet: int) -> None: ...

    # Sheet protection.
    def get_sheet_protection(self, sheet: int) -> SheetProtection: ...
    def set_sheet_protection(self, sheet: int, protection: SheetProtection) -> None: ...

    # Sheet view / layout.
    def paginate(self, sheet: int) -> PaginationResult: ...
    def get_sheet_view(self, sheet: int) -> SheetView: ...
    def set_sheet_zoom(self, sheet: int, zoom_scale: int) -> None: ...
    def set_sheet_freeze(self, sheet: int, freeze_rows: int, freeze_cols: int) -> None: ...
    def set_sheet_tab_hidden(self, sheet: int, hidden: bool) -> None: ...
    def set_sheet_visibility(self, sheet: int, visibility: Union[SheetVisibility, int]) -> None: ...
    def set_sheet_show_grid_lines(self, sheet: int, show: bool) -> None: ...
    def set_sheet_show_row_col_headers(self, sheet: int, show: bool) -> None: ...
    def set_sheet_show_zeros(self, sheet: int, show: bool) -> None: ...
    def set_sheet_right_to_left(self, sheet: int, right_to_left: bool) -> None: ...
    def set_sheet_tab_selected(self, sheet: int, selected: bool) -> None: ...
    def set_sheet_view_mode(self, sheet: int, mode: str) -> None: ...
    def get_auto_filter_xml(self, sheet: int) -> str: ...
    def set_auto_filter_xml(self, sheet: int, xml: str) -> None: ...

    # Print settings.
    def get_page_setup_xml(self, sheet: int) -> str: ...
    def set_page_setup_xml(self, sheet: int, xml: str) -> None: ...
    def get_page_margins_xml(self, sheet: int) -> str: ...
    def set_page_margins_xml(self, sheet: int, xml: str) -> None: ...
    def get_print_options_xml(self, sheet: int) -> str: ...
    def set_print_options_xml(self, sheet: int, xml: str) -> None: ...
    def get_header_footer_xml(self, sheet: int) -> str: ...
    def set_header_footer_xml(self, sheet: int, xml: str) -> None: ...
    def get_sheet_pr_xml(self, sheet: int) -> str: ...
    def set_sheet_pr_xml(self, sheet: int, xml: str) -> None: ...
    def set_fit_to_page(self, sheet: int, enabled: bool) -> None: ...
    def get_print_area(self, sheet: int) -> str: ...
    def set_print_area(self, sheet: int, ranges_a1: str) -> None: ...
    def get_print_titles(self, sheet: int) -> tuple[str, str]: ...
    def set_print_titles(self, sheet: int, repeat_rows: str = ..., repeat_cols: str = ...) -> None: ...
    def add_row_break(self, sheet: int, row: int, manual: bool = ...) -> None: ...
    def add_col_break(self, sheet: int, col: int, manual: bool = ...) -> None: ...
    def remove_row_break(self, sheet: int, row: int) -> None: ...
    def remove_col_break(self, sheet: int, col: int) -> None: ...
    def clear_breaks(self, sheet: int) -> None: ...
    def get_row_breaks(self, sheet: int) -> List[PageBreak]: ...
    def get_col_breaks(self, sheet: int) -> List[PageBreak]: ...
    def set_page_setup(
        self,
        sheet: int,
        *,
        orientation: Optional[int] = ...,
        paper_size: Optional[int] = ...,
        scale: Optional[int] = ...,
        fit_to_width: Optional[int] = ...,
        fit_to_height: Optional[int] = ...,
        fit_to_page: Optional[bool] = ...,
    ) -> None: ...
    def set_page_margins(
        self,
        sheet: int,
        *,
        left: Optional[float] = ...,
        right: Optional[float] = ...,
        top: Optional[float] = ...,
        bottom: Optional[float] = ...,
        header: Optional[float] = ...,
        footer: Optional[float] = ...,
    ) -> None: ...
    def set_print_options(
        self,
        sheet: int,
        *,
        grid_lines: Optional[bool] = ...,
        headings: Optional[bool] = ...,
        horizontal_centered: Optional[bool] = ...,
        vertical_centered: Optional[bool] = ...,
    ) -> None: ...
    def set_header_footer(
        self,
        sheet: int,
        *,
        odd_header: Optional[str] = ...,
        odd_footer: Optional[str] = ...,
        even_header: Optional[str] = ...,
        even_footer: Optional[str] = ...,
        first_header: Optional[str] = ...,
        first_footer: Optional[str] = ...,
        different_odd_even: Optional[bool] = ...,
        different_first: Optional[bool] = ...,
        scale_with_doc: Optional[bool] = ...,
        align_with_margins: Optional[bool] = ...,
    ) -> None: ...
    def get_page_setup(self, sheet: int) -> PageSetup: ...
    def get_page_margins(self, sheet: int) -> PageMargins: ...
    def get_sheet_columns(self, sheet: int) -> List[ColumnLayout]: ...
    def set_column_width(self, sheet: int, first: int, last: int, width: float) -> None: ...
    def set_column_hidden(self, sheet: int, first: int, last: int, hidden: bool) -> None: ...
    def set_column_outline(self, sheet: int, first: int, last: int, level: int) -> None: ...
    def get_sheet_row_overrides(self, sheet: int) -> List[RowLayout]: ...
    def set_row_height(self, sheet: int, row: int, height: float) -> None: ...
    def clear_row_height(self, sheet: int, row: int) -> None: ...
    def get_sheet_format_defaults(self, sheet: int) -> SheetFormatDefaults: ...
    def set_sheet_format_defaults(self, sheet: int, defaults: SheetFormatDefaults) -> None: ...
    def get_cell_rect_pt(self, sheet: int, cell_range: MergeRange, mode: Union[GeometryMode, int] = ...) -> RectPt: ...
    def get_column_width_pt(self, sheet: int, col: int, mode: Union[GeometryMode, int] = ...) -> float: ...
    def get_row_height_pt(self, sheet: int, row: int) -> float: ...
    def get_width_model(self, sheet: int, mode: Union[GeometryMode, int] = ...) -> WidthModel: ...
    def column_chars_to_pt(self, sheet: int, chars: float, mode: Union[GeometryMode, int] = ...) -> float: ...
    def column_pt_to_chars(self, sheet: int, pt: float, mode: Union[GeometryMode, int] = ...) -> float: ...
    def get_formula(self, sheet: int, row: int, col: int) -> Optional[str]: ...
    def get_formula_r1c1(self, sheet: int, row: int, col: int) -> Optional[str]: ...
    def get_cells_in_range(
        self,
        sheet: int,
        cell_range: MergeRange,
        cursor: Optional[int] = ...,
        limit: Optional[int] = ...,
    ) -> Tuple[List[Cell], Optional[int]]: ...
    def get_auto_filter(self, sheet: int) -> Optional[AutoFilter]: ...
    def set_auto_filter(self, sheet: int, auto_filter: AutoFilter) -> None: ...
    def remove_auto_filter(self, sheet: int) -> None: ...
    def apply_auto_filter(self, sheet: int) -> None: ...
    def clear_auto_filter(self, sheet: int) -> None: ...
    def evaluate_auto_filter(self, sheet: int) -> Tuple[int, List[bool]]: ...
    def get_table_auto_filter(self, table_index: int) -> Optional[AutoFilter]: ...
    def set_table_auto_filter(self, table_index: int, auto_filter: AutoFilter) -> None: ...
    def remove_table_auto_filter(self, table_index: int) -> None: ...
    def apply_table_auto_filter(self, table_index: int) -> None: ...
    def clear_table_auto_filter(self, table_index: int) -> None: ...
    def evaluate_table_auto_filter(self, table_index: int) -> Tuple[int, List[bool]]: ...
    def validate_value(self, sheet: int, row: int, col: int, value: Value) -> ValidationOutcome: ...
    def list_invalid_cells(
        self,
        sheet: int,
        cursor: Optional[int] = ...,
        limit: Optional[int] = ...,
    ) -> Tuple[List[Cell], Optional[int]]: ...
    def get_threaded_comments(self, sheet: int) -> List[ThreadedComment]: ...
    def add_threaded_comment(self, sheet: int, comment: ThreadedComment) -> None: ...
    def edit_threaded_comment(
        self, sheet: int, comment_id: str, text: str, mentions: Sequence[Mention] = ...
    ) -> None: ...
    def set_thread_resolved(self, sheet: int, thread_id: str, done: bool) -> None: ...
    def remove_threaded_comment(self, sheet: int, comment_id: str) -> None: ...
    def probe_image(self, data: Union[bytes, bytearray, memoryview]) -> ImageInfo: ...
    def list_drawing_objects(self, sheet: int) -> List[DrawingObject]: ...
    def get_image(self, sheet: int, object_id: int) -> bytes: ...
    def insert_image(
        self,
        sheet: int,
        data: Union[bytes, bytearray, memoryview],
        *,
        name: str = ...,
        descr: str = ...,
        anchor_kind: Union[AnchorKind, int] = ...,
        edit_as: Union[AnchorEditAs, int] = ...,
        row: int = ...,
        col: int = ...,
        row_off_emu: int = ...,
        col_off_emu: int = ...,
        width_emu: int = ...,
        height_emu: int = ...,
    ) -> int: ...
    def remove_image(self, sheet: int, object_id: int) -> None: ...
    def get_persons(self) -> List[Person]: ...
    def add_person(self, person: Person) -> None: ...
    def remove_person(self, person_id: str) -> None: ...
    def get_merges_in_range(self, sheet: int, cell_range: MergeRange) -> List[MergeRange]: ...
    def get_display_text(self, sheet: int, row: int, col: int) -> Tuple[str, DisplayStatus]: ...
    def format_value(self, value: Value, format_code: str = ...) -> Tuple[str, DisplayStatus]: ...
    def set_row_hidden(self, sheet: int, row: int, hidden: bool) -> None: ...
    def set_row_outline(self, sheet: int, row: int, level: int) -> None: ...

    # Conditional formatting.
    def evaluate_cf_range(
        self,
        sheet: int,
        first_row: int,
        first_col: int,
        last_row: int,
        last_col: int,
        today_serial: float = ...,
    ) -> List[CfCellResult]: ...
    def cf_count(self, sheet: int) -> int: ...
    def get_conditional_format_at(self, sheet: int, index: int) -> ConditionalFormat: ...
    def get_conditional_formats(self, sheet: int) -> List[ConditionalFormat]: ...
    def add_conditional_format(self, sheet: int, rule: ConditionalFormatInput) -> int: ...
    def remove_conditional_format_at(self, sheet: int, index: int) -> None: ...
    def clear_conditional_formats(self, sheet: int) -> None: ...

    # Styles.
    def get_cell_xf_index(self, sheet: int, row: int, col: int) -> int: ...
    def set_cell_xf_index(self, sheet: int, row: int, col: int, xf_index: int) -> None: ...
    def set_range_xf_index(
        self, sheet: int, first_row: int, first_col: int, last_row: int, last_col: int, xf_index: int
    ) -> None: ...
    def get_cell_xf(self, xf_index: int) -> CellXf: ...
    def get_font(self, font_index: int) -> FontRecord: ...
    def get_fill(self, fill_index: int) -> FillRecord: ...
    def get_border(self, border_index: int) -> Dict[str, object]: ...
    def get_dxf(self, dxf_index: int) -> DifferentialFormat: ...
    def get_num_fmt(self, num_fmt_id: int) -> str: ...
    def font_count(self) -> int: ...
    def fill_count(self) -> int: ...
    def border_count(self) -> int: ...
    def cell_xf_count(self) -> int: ...
    def cell_style_count(self) -> int: ...
    def cell_style_xf_count(self) -> int: ...
    def dxf_count(self) -> int: ...
    def add_font(self, record: FontRecord) -> int: ...
    def set_font(self, font_index: int, record: FontRecord) -> None: ...
    def set_default_font(self, record: FontRecord) -> None: ...
    def add_fill(self, record: FillRecord) -> int: ...
    def add_border(self, sides: Dict[str, object]) -> int: ...
    def add_num_fmt(self, format_code: str) -> int: ...
    def add_cell_xf(self, record: CellXf) -> int: ...
    def add_cell_style_xf(self, record: CellXf) -> int: ...
    def add_batch(
        self,
        *,
        fonts: Optional[Sequence[FontRecord]] = ...,
        fills: Optional[Sequence[FillRecord]] = ...,
        borders: Optional[Sequence[Dict[str, object]]] = ...,
        num_fmts: Optional[Sequence[str]] = ...,
        cell_xfs: Optional[Sequence[CellXf]] = ...,
    ) -> StyleBatchIndices: ...
    def add_dxf(self, record: DifferentialFormat) -> int: ...
    def set_cell_style(self, record: CellStyle) -> None: ...
    def remove_cell_style(self, name: str) -> None: ...
    def get_theme(self) -> Theme: ...
    def set_theme_colors(self, colors: Sequence[int]) -> None: ...
    def set_theme_fonts(self, fonts: ThemeFonts) -> None: ...
    def resolve_color(self, spec: ColorSpec, context: Union[ColorContext, int] = ...) -> ResolvedColor: ...
    def get_effective_style(self, sheet: int, row: int, col: int) -> EffectiveStyle: ...
    def get_cell_style(self, index: int) -> CellStyle: ...
    def get_cell_style_xf(self, index: int) -> CellXf: ...

    # Pivot layout projection.
    def pivot_count(self, sheet: int) -> int: ...
    def pivot_layout(self, sheet: int, pivot_index: int) -> PivotLayout: ...
    def get_pivot_report_layout(self, sheet: int, pivot_index: int) -> PivotReportLayout: ...
    def set_pivot_report_layout(self, sheet: int, pivot_index: int, layout: Union[PivotReportLayout, int]) -> None: ...

    # Pivot caches.
    def pivot_cache_count(self) -> int: ...
    def pivot_cache_id_at(self, index: int) -> int: ...
    def pivot_cache_create(self, requested_id: int = ...) -> int: ...
    def get_pivot_cache_worksheet_source(self, cache_id: int) -> Optional[PivotWorksheetSource]: ...
    def set_pivot_cache_worksheet_source(self, cache_id: int, source: Optional[PivotWorksheetSource]) -> None: ...
    def pivot_cache_remove(self, cache_id: int) -> None: ...
    def pivot_cache_field_count(self, cache_id: int) -> int: ...
    def pivot_cache_field_name(self, cache_id: int, field_idx: int) -> str: ...
    def pivot_cache_field_add(self, cache_id: int, name: str) -> int: ...
    def pivot_cache_field_clear(self, cache_id: int) -> None: ...
    def pivot_cache_field_shared_item_count(self, cache_id: int, field_idx: int) -> int: ...
    def pivot_cache_field_add_shared_item_number(self, cache_id: int, field_idx: int, value: float) -> None: ...
    def pivot_cache_field_add_shared_item_text(self, cache_id: int, field_idx: int, value: str) -> None: ...
    def pivot_cache_field_add_shared_item_bool(self, cache_id: int, field_idx: int, value: bool) -> None: ...
    def pivot_cache_field_add_shared_item_blank(self, cache_id: int, field_idx: int) -> None: ...
    def pivot_cache_field_add_shared_item_error(self, cache_id: int, field_idx: int, error_code: int) -> None: ...
    def pivot_cache_field_clear_shared_items(self, cache_id: int, field_idx: int) -> None: ...
    def pivot_cache_record_count(self, cache_id: int) -> int: ...
    def pivot_cache_record_add(self, cache_id: int) -> int: ...
    def pivot_cache_record_clear(self, cache_id: int) -> None: ...
    def pivot_cache_record_set_number(self, cache_id: int, record_idx: int, field_idx: int, value: float) -> None: ...
    def pivot_cache_record_set_text(self, cache_id: int, record_idx: int, field_idx: int, value: str) -> None: ...
    def pivot_cache_record_set_bool(self, cache_id: int, record_idx: int, field_idx: int, value: bool) -> None: ...
    def pivot_cache_record_set_blank(self, cache_id: int, record_idx: int, field_idx: int) -> None: ...
    def pivot_cache_record_set_error(self, cache_id: int, record_idx: int, field_idx: int, error_code: int) -> None: ...

    # Pivot tables.
    def pivot_create(self, sheet: int, name: str, cache_id: int, anchor_row: int, anchor_col: int) -> int: ...
    def pivot_remove(self, sheet: int, pivot_index: int) -> None: ...
    def pivot_set_name(self, sheet: int, pivot_index: int, name: str) -> None: ...
    def pivot_set_anchor(
        self,
        sheet: int,
        pivot_index: int,
        anchor_row: int,
        anchor_col: int,
        span_rows: int,
        span_cols: int,
    ) -> None: ...
    def pivot_set_grand_totals(self, sheet: int, pivot_index: int, rows_enabled: bool, cols_enabled: bool) -> None: ...
    def pivot_field_count(self, sheet: int, pivot_index: int) -> int: ...
    def pivot_field_add(self, sheet: int, pivot_index: int, spec: PivotFieldSpec) -> int: ...
    def pivot_field_clear(self, sheet: int, pivot_index: int) -> None: ...
    def pivot_field_set_axis(
        self, sheet: int, pivot_index: int, field_idx: int, axis: Union[PivotAxis, int]
    ) -> None: ...
    def pivot_field_set_sort(
        self,
        sheet: int,
        pivot_index: int,
        field_idx: int,
        ascending: bool,
        by_field: str = ...,
    ) -> None: ...
    def pivot_field_set_subtotal_top(self, sheet: int, pivot_index: int, field_idx: int, top: bool) -> None: ...
    def pivot_field_add_item(
        self,
        sheet: int,
        pivot_index: int,
        field_idx: int,
        name: str,
        visible: bool,
    ) -> None: ...
    def pivot_field_add_item_at(
        self,
        sheet: int,
        pivot_index: int,
        field_idx: int,
        cache_index: int,
        visible: bool,
    ) -> None: ...
    def pivot_field_clear_items(self, sheet: int, pivot_index: int, field_idx: int) -> None: ...
    def pivot_field_set_item_visible(
        self,
        sheet: int,
        pivot_index: int,
        field_idx: int,
        item_idx: int,
        visible: bool,
    ) -> None: ...
    def pivot_field_add_subtotal_fn(
        self, sheet: int, pivot_index: int, field_idx: int, agg: Union[PivotAggregation, int]
    ) -> None: ...
    def pivot_field_clear_subtotal_fns(self, sheet: int, pivot_index: int, field_idx: int) -> None: ...
    def pivot_field_set_date_group(
        self,
        sheet: int,
        pivot_index: int,
        field_idx: int,
        granularity: Union[PivotDateGrouping, int],
        calendar: Union[PivotCalendar, int],
        start_year: int = ...,
        end_year: int = ...,
        interval_days: int = ...,
        start_serial: float = ...,
        end_serial: float = ...,
    ) -> None: ...
    def pivot_field_clear_date_group(self, sheet: int, pivot_index: int, field_idx: int) -> None: ...
    def pivot_field_set_number_format(self, sheet: int, pivot_index: int, field_idx: int, fmt: str) -> None: ...
    def pivot_set_row_field_order(self, sheet: int, pivot_index: int, indices: Sequence[int]) -> None: ...
    def pivot_set_col_field_order(self, sheet: int, pivot_index: int, indices: Sequence[int]) -> None: ...
    def pivot_data_field_count(self, sheet: int, pivot_index: int) -> int: ...
    def pivot_data_field_add(self, sheet: int, pivot_index: int, spec: PivotDataFieldSpec) -> int: ...
    def pivot_data_field_set(
        self,
        sheet: int,
        pivot_index: int,
        data_field_idx: int,
        spec: PivotDataFieldSpec,
    ) -> None: ...
    def pivot_data_field_clear(self, sheet: int, pivot_index: int) -> None: ...
    def pivot_filter_count(self, sheet: int, pivot_index: int) -> int: ...
    def pivot_filter_add(self, sheet: int, pivot_index: int, spec: PivotFilterSpec) -> None: ...
    def pivot_filter_at(self, sheet: int, pivot_index: int, filter_idx: int) -> PivotFilterSpec: ...
    def pivot_filter_clear(self, sheet: int, pivot_index: int) -> None: ...
    def pivot_filter_remove_at(self, sheet: int, pivot_index: int, filter_idx: int) -> None: ...

    # Dependency-graph trace + spill.
    def precedents(self, sheet: int, row: int, col: int, depth: int = ...) -> List[CellNode]: ...
    def dependents(self, sheet: int, row: int, col: int, depth: int = ...) -> List[CellNode]: ...
    def spill_info(self, sheet: int, row: int, col: int) -> SpillInfo: ...

    # Function catalog (workbook-independent).
    @staticmethod
    def function_count() -> int: ...
    @staticmethod
    def function_name_at(index: int) -> str: ...
    @staticmethod
    def function_metadata(name: str, locale: int = ...) -> Optional[FunctionMetadata]: ...
    @staticmethod
    def localize_function_name(canonical_name: str, locale: int = ...) -> str: ...
    @staticmethod
    def canonicalize_function_name(localized_name: str, locale: int = ...) -> str: ...

    # External links.
    def external_link_count(self) -> int: ...
    def get_external_link_at(self, index: int) -> ExternalLink: ...
    def get_external_links(self) -> List[ExternalLink]: ...

def library_version() -> str: ...
def version_string() -> str: ...
def error_display_name(error_code: int) -> str: ...
def set_log_min_level(level: Union[LogLevel, int]) -> None: ...
def eval_formula(formula: str) -> Value: ...
