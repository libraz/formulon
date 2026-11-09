"""The :class:`Sheet` handle: every operation that targets one sheet of a workbook.

A :class:`~formulon.Workbook` hands out one :class:`Sheet` per sheet and keeps it
pointing at the same sheet across insertions, moves, renames and removals;
see :class:`Sheet`.
"""

from __future__ import annotations

import builtins
import dataclasses
import datetime
import decimal
import inspect
import math
import numbers
import struct
from typing import (
    Any,
    Callable,
    Dict,
    Iterator,
    List,
    Mapping,
    Optional,
    Sequence,
    Tuple,
    TypeVar,
    Union,
    cast,
    overload,
)

from . import (
    _structs as S,
)
from ._c import (
    LIB,
    FormulonError,
    ValueKind,
    _argument_message,
    _call_label,
    _NoCopy,
    _require_bool,
    _sint,
    _uint,
    fm_value_t_size,
)
from ._handle_tables import READ_ONLY_KEYS
from ._inputs import cell_rect, civil_datetime, color_input, date_serial, normalize_record_ranges, serial_datetime
from ._refs import _fold, resolve_ref
from ._status import (
    StatusCode,
)
from ._workbook import (
    _CELL_RANGE_NO_CURSOR,
    _PATCH_CLEAR,
    _PATCH_SET,
    _SPARKLINE_COLORS,
    UNCHANGED,
    Alignment,
    AlignmentPatch,
    ArrayFormula,
    AutoFilter,
    AutoFilterEval,
    AutoFilterInput,
    BlanksAs,
    Border,
    BorderPatch,
    BorderStyle,
    CellAddress,
    CellFormat,
    CellNode,
    CellProtection,
    CellProtectionPatch,
    CellRect,
    CellShift,
    CellsInRange,
    CellXf,
    CellXfInput,
    CfCell,
    CfColorScale,
    CfColorScaleInput,
    CfDataBar,
    CfDataBarAxisPosition,
    CfDataBarDirection,
    CfDataBarInput,
    CfIcon,
    CfIconSet,
    CfIconSetInput,
    CfMatch,
    CfRuleId,
    CfRuleRef,
    CfValueObject,
    CfValueObjectType,
    ChartInput,
    Color,
    ColorInput,
    ColorKind,
    ColorResolution,
    ColumnBreak,
    ColumnLayout,
    Comment,
    ComparisonOperator,
    ConditionalFormat,
    ConditionalFormatInput,
    DataTable,
    DataTableInput,
    DataTableInputCell,
    DataValidation,
    DataValidationInput,
    DatePeriod,
    DisplayStatus,
    DisplayText,
    DrawingPlacementInput,
    EffectiveBorders,
    EffectiveStyle,
    EffectiveStyleSource,
    ErrorCode,
    ExcelError,
    Fill,
    FillPatch,
    FillPattern,
    Font,
    FontFamily,
    FontPatch,
    FontScheme,
    FontVerticalAlign,
    GeometryMode,
    GradientFill,
    HeaderFooter,
    HorizontalAlign,
    Hyperlink,
    HyperlinkInput,
    Mention,
    MentionInput,
    NativeValue,
    PageBreakInput,
    PageMargins,
    PageOrientation,
    PageSetup,
    Pagination,
    PaginationMargins,
    PaginationPage,
    PaginationPaper,
    PhoneticProperties,
    PhoneticRun,
    PivotInput,
    PivotRef,
    PrintOptions,
    PrintTitles,
    PrintTitlesInput,
    ProtectedRange,
    ProtectedRangeInput,
    ProtectedRangeRef,
    RangeCell,
    ReadingOrder,
    RectPt,
    ResolvedColor,
    RowBreak,
    RowLayout,
    ShapeInput,
    SheetComment,
    SheetFormatDefaults,
    SheetProtection,
    SheetProtectionInput,
    SheetRect,
    SheetRef,
    SheetSelection,
    SheetSelectionInput,
    SheetView,
    SheetViewMode,
    SheetVisibility,
    SlicerInput,
    Span,
    Sparkline,
    SparklineAxisType,
    SparklineGroup,
    SparklineGroupId,
    SparklineGroupInput,
    SparklineGroupRef,
    SparklineInput,
    SparklineType,
    Spill,
    TableInput,
    TableRef,
    ThreadedComment,
    ThreadedCommentInput,
    Unchanged,
    UnderlineStyle,
    ValidationErrorStyle,
    ValidationId,
    ValidationOutcome,
    ValidationRef,
    ValidationType,
    Value,
    VerticalAlign,
    WidthModel,
    Workbook,
    _ArgScope,
    _cell_rect_fields,
    _cell_rect_from,
    _cf_match_from,
    _check,
    _check_value,
    _decode_auto_filter,
    _decode_cell_rect_array,
    _decode_cell_rect_at,
    _decode_color,
    _decode_optional_color,
    _decode_sparkline,
    _decode_sparkline_date_axis,
    _encode_chart_insert,
    _enum,
    _finite,
    _format_created,
    _hyperlink_from,
    _install_workbook_locks,
    _pack_cell_rect_array,
    _pack_cell_rect_at,
    _pack_color,
    _pack_mentions,
    _pack_optional_color,
    _pack_phonetic_run_array,
    _pack_pivot_spec,
    _pack_shape_spec,
    _pack_slicer_spec,
    _pack_sparkline_group,
    _pack_sparklines,
    _placement_fields,
    _print_titles_from,
    _qualified_range_text,
    _read_count,
    _read_cstr_out,
    _read_optional_cstr_out,
    _threaded_comment_from,
    _tristate_default,
    _write_mentions,
    _write_value_input,
    to_native,
)

__all__ = ["Sheet", "SheetOoxml"]

_SPARKLINE_FLAGS = (
    "markers",
    "high",
    "low",
    "first",
    "last",
    "negative",
    "display_x_axis",
    "display_hidden",
    "date_axis",
)

_VALIDATION_FLAGS = ("allow_blank", "show_input_message", "show_error_message", "show_in_cell_dropdown")
_VALIDATION_TEXTS = ("formula1", "formula2", "error_title", "error_message", "prompt_title", "prompt_message")

# <facade-tables>
WORKBOOK_METHODS = (
    "active_sheet",
    "add_person",
    "add_pivot_cache",
    "add_sheet",
    "close",
    "closed",
    "copy_sheet",
    "format_reference",
    "format_value",
    "get_calc_mode",
    "get_calc_on_save",
    "get_cell_style",
    "get_date1904",
    "get_default_font",
    "get_defined_name",
    "get_diagnostics",
    "get_file_name",
    "get_full_precision",
    "get_iterative",
    "get_memory_usage",
    "get_person",
    "get_pinned_now",
    "get_profile_id",
    "get_ref_mode",
    "get_table_style",
    "get_table_style_defaults",
    "get_theme_colors",
    "get_theme_fonts",
    "get_theme_script_font",
    "get_workbook_protection",
    "has_dirty_cells",
    "list_cell_styles",
    "list_defined_names",
    "list_external_links",
    "list_people",
    "list_pivot_caches",
    "list_sheets",
    "list_table_styles",
    "list_theme_script_fonts",
    "ooxml",
    "parse_reference",
    "pivot_cache",
    "recalc",
    "recalc_data_tables",
    "remove_cell_style",
    "remove_defined_name",
    "remove_person",
    "remove_pivot_cache",
    "remove_sheet",
    "remove_table_style",
    "reset_theme",
    "resolve_color",
    "save",
    "save_with_diagnostics",
    "set_active_sheet",
    "set_calc_mode",
    "set_calc_on_save",
    "set_cell_style",
    "set_date1904",
    "set_default_font",
    "set_defined_name",
    "set_file_name",
    "set_full_precision",
    "set_pinned_now",
    "set_profile_id",
    "set_ref_mode",
    "set_table_style",
    "set_theme_colors",
    "set_theme_fonts",
    "set_theme_script_font",
    "set_workbook_protection",
    "sheet",
    "table",
    "update_iterative",
    "update_table_style_defaults",
)
SHEET_METHODS = (
    "add_cf_rule",
    "add_chart",
    "add_data_table",
    "add_merge",
    "add_picture",
    "add_pivot",
    "add_protected_range",
    "add_shape",
    "add_slicer",
    "add_sparkline_group",
    "add_table",
    "add_threaded_comment",
    "add_validation",
    "apply_auto_filter",
    "chart",
    "clear_auto_filter_criteria",
    "clear_breaks",
    "clear_cf_rules",
    "clear_hyperlinks",
    "clear_merges",
    "clear_validations",
    "convert_column_chars_to_pt",
    "convert_column_pt_to_chars",
    "delete_cells",
    "delete_columns",
    "delete_rows",
    "drawing",
    "evaluate_auto_filter",
    "evaluate_cf_formula",
    "evaluate_cf_range",
    "evaluate_formula",
    "evaluate_formula_array",
    "evaluate_typed_formula",
    "evaluate_typed_formula_array",
    "get_array_formula",
    "get_auto_filter",
    "get_cell_format",
    "get_cell_phonetic",
    "get_cell_phonetic_properties",
    "get_cell_phonetic_runs",
    "get_cell_rect_pt",
    "get_cell_show_phonetic",
    "get_cells_in_range",
    "get_cf_rule",
    "get_column_width_pt",
    "get_comment",
    "get_comment_visible",
    "get_data_table",
    "get_date",
    "get_dependents",
    "get_display_text",
    "get_format_defaults",
    "get_formula",
    "get_formula_r1c1",
    "get_header_footer",
    "get_hyperlink",
    "get_invalid_cells",
    "get_lambda_text",
    "get_merges_in_range",
    "get_page_margins",
    "get_page_setup",
    "get_pagination",
    "get_precedents",
    "get_print_area",
    "get_print_options",
    "get_print_titles",
    "get_protected_range",
    "get_row_height_pt",
    "get_selection",
    "get_sheet_protection",
    "get_sparkline_group",
    "get_spill",
    "get_tab_color",
    "get_threaded_comment",
    "get_typed_value",
    "get_typed_values",
    "get_used_range",
    "get_validation",
    "get_value",
    "get_values",
    "get_view",
    "get_visibility",
    "get_width_model",
    "insert_cells",
    "insert_columns",
    "insert_rows",
    "iter_rows",
    "list_cells",
    "list_cf_rules",
    "list_charts",
    "list_column_breaks",
    "list_column_overrides",
    "list_comments",
    "list_drawings",
    "list_hyperlinks",
    "list_merges",
    "list_pictures",
    "list_pivots",
    "list_protected_ranges",
    "list_row_breaks",
    "list_row_overrides",
    "list_shapes",
    "list_slicers",
    "list_sparkline_groups",
    "list_tables",
    "list_threaded_comments",
    "list_validations",
    "picture",
    "pivot",
    "recalc_viewport",
    "remove_cf_rule",
    "remove_data_table",
    "remove_drawing",
    "remove_merge",
    "remove_pivot",
    "remove_protected_range",
    "remove_sparkline_group",
    "remove_table",
    "remove_threaded_comment",
    "remove_validation",
    "reset_column_width",
    "reset_row_height",
    "restore_drawing",
    "set_array_formula",
    "set_auto_filter",
    "set_blank",
    "set_bool",
    "set_cell_phonetic",
    "set_cell_phonetic_properties",
    "set_cell_phonetic_runs",
    "set_cell_show_phonetic",
    "set_column_break",
    "set_column_collapsed",
    "set_column_hidden",
    "set_column_outline",
    "set_column_width",
    "set_comment",
    "set_comment_visible",
    "set_error",
    "set_formula",
    "set_hyperlink",
    "set_number",
    "set_print_area",
    "set_print_titles",
    "set_row_auto_height",
    "set_row_break",
    "set_row_collapsed",
    "set_row_height",
    "set_row_hidden",
    "set_row_outline",
    "set_selection",
    "set_sheet_protection",
    "set_tab_color",
    "set_text",
    "set_value",
    "set_values",
    "set_visibility",
    "shape",
    "slicer",
    "table",
    "update_cell_format",
    "update_cf_rule",
    "update_format_defaults",
    "update_header_footer",
    "update_page_margins",
    "update_page_setup",
    "update_print_options",
    "update_protected_range",
    "update_sheet_protection",
    "update_sparkline_group",
    "update_threaded_comment",
    "update_validation",
    "update_view",
    "validate_value",
)
OOXML_WORKBOOK_METHODS = (
    "add_border",
    "add_cell_style_xf",
    "add_cell_xf",
    "add_dxf",
    "add_fill",
    "add_font",
    "add_number_format",
    "get_border",
    "get_cell_style_xf",
    "get_cell_xf",
    "get_dxf",
    "get_fill",
    "get_font",
    "get_number_format",
    "has_passthrough",
    "list_borders",
    "list_cell_style_xfs",
    "list_cell_xfs",
    "list_dxfs",
    "list_fills",
    "list_fonts",
    "list_number_formats",
    "list_passthroughs",
    "set_border",
    "set_cell_style_xf",
    "set_cell_xf",
    "set_dxf",
    "set_fill",
    "set_font",
)
OOXML_SHEET_METHODS = (
    "get_auto_filter_xml",
    "get_cell_xf_index",
    "get_column_xf_index",
    "get_effective_style",
    "get_header_footer_xml",
    "get_page_margins_xml",
    "get_page_setup_xml",
    "get_print_options_xml",
    "get_range_xf_bases",
    "get_range_xf_bases_excluding",
    "get_row_xf_index",
    "get_sheet_pr_xml",
    "remap_range_xfs",
    "reset_cell_xf_index",
    "reset_column_xf_index",
    "reset_row_xf_index",
    "restore_range_xfs",
    "set_auto_filter_xml",
    "set_cell_xf_index",
    "set_column_xf_index",
    "set_header_footer_xml",
    "set_page_margins_xml",
    "set_page_setup_xml",
    "set_print_options_xml",
    "set_range_xf_index",
    "set_row_xf_index",
    "set_sheet_pr_xml",
    "snapshot_range_xfs",
)
SHEET_STRUCTURE_OPS = (
    "Pivot.show_details",
    "Sheet.move",
    "Sheet.rename",
    "Workbook.add_sheet",
    "Workbook.copy_sheet",
    "Workbook.remove_sheet",
)
LIST_FAMILIES = {
    "Sheet.list_cf_rules": ("get_cf_rule_count", "get_cf_rule"),
    "Sheet.list_protected_ranges": ("get_protected_range_count", "get_protected_range"),
    "Sheet.list_sparkline_groups": ("get_sparkline_group_count", "get_sparkline_group"),
    "Sheet.list_validations": ("get_validation_count", "get_validation"),
    "Workbook.list_borders": ("get_border_count", "get_border"),
    "Workbook.list_cell_style_xfs": ("get_cell_style_xf_count", "get_cell_style_xf"),
    "Workbook.list_cell_xfs": ("get_cell_xf_count", "get_cell_xf"),
    "Workbook.list_dxfs": ("get_dxf_count", "get_dxf"),
    "Workbook.list_external_links": ("get_external_link_count", "get_external_link"),
    "Workbook.list_fills": ("get_fill_count", "get_fill"),
    "Workbook.list_fonts": ("get_font_count", "get_font"),
}
HANDWRITTEN = (
    "active_sheet",
    "add_cf_rule",
    "add_data_table",
    "add_picture",
    "add_sheet",
    "chart",
    "copy_sheet",
    "drawing",
    "evaluate_formula",
    "evaluate_typed_formula",
    "evaluate_typed_formula_array",
    "format_reference",
    "get_cell_format",
    "get_cell_phonetic_runs",
    "get_cells_in_range",
    "get_column_xf_index",
    "get_date",
    "get_default_font",
    "get_formula",
    "get_invalid_cells",
    "get_range_xf_bases_excluding",
    "get_row_xf_index",
    "get_theme_colors",
    "get_theme_fonts",
    "get_typed_value",
    "get_typed_values",
    "get_used_range",
    "get_values",
    "iter_rows",
    "list_cells",
    "list_charts",
    "list_drawings",
    "list_pictures",
    "list_pivot_caches",
    "list_pivots",
    "list_shapes",
    "list_sheets",
    "list_slicers",
    "list_tables",
    "list_threaded_comments",
    "picture",
    "pivot",
    "pivot_cache",
    "remap_range_xfs",
    "remove_drawing",
    "remove_sheet",
    "set_active_sheet",
    "set_array_formula",
    "set_comment",
    "set_formula",
    "set_hyperlink",
    "set_pinned_now",
    "set_theme_colors",
    "set_theme_script_font",
    "set_value",
    "set_values",
    "shape",
    "slicer",
    "table",
    "update_cell_format",
    "update_cf_rule",
    "update_protected_range",
    "update_sparkline_group",
    "update_threaded_comment",
    "update_validation",
    "update_view",
)
CELL_ARGS = {
    "Sheet.evaluate_cf_formula": ((0, 1), (2, 3)),
    "Sheet.evaluate_formula": ((0, 1),),
    "Sheet.evaluate_formula_array": ((0, 1),),
    "Sheet.evaluate_typed_formula": ((0, 1),),
    "Sheet.evaluate_typed_formula_array": ((0, 1),),
    "Sheet.get_array_formula": ((0, 1),),
    "Sheet.get_cell_format": ((0, 1),),
    "Sheet.get_cell_phonetic": ((0, 1),),
    "Sheet.get_cell_phonetic_properties": ((0, 1),),
    "Sheet.get_cell_phonetic_runs": ((0, 1),),
    "Sheet.get_cell_show_phonetic": ((0, 1),),
    "Sheet.get_cell_xf_index": ((0, 1),),
    "Sheet.get_comment": ((0, 1),),
    "Sheet.get_comment_visible": ((0, 1),),
    "Sheet.get_data_table": ((0, 1),),
    "Sheet.get_date": ((0, 1),),
    "Sheet.get_dependents": ((0, 1),),
    "Sheet.get_display_text": ((0, 1),),
    "Sheet.get_effective_style": ((0, 1),),
    "Sheet.get_formula": ((0, 1),),
    "Sheet.get_formula_r1c1": ((0, 1),),
    "Sheet.get_hyperlink": ((0, 1),),
    "Sheet.get_lambda_text": ((0, 1),),
    "Sheet.get_precedents": ((0, 1),),
    "Sheet.get_spill": ((0, 1),),
    "Sheet.get_typed_value": ((0, 1),),
    "Sheet.get_value": ((0, 1),),
    "Sheet.remove_data_table": ((0, 1),),
    "Sheet.reset_cell_xf_index": ((0, 1),),
    "Sheet.set_blank": ((0, 1),),
    "Sheet.set_bool": ((0, 1),),
    "Sheet.set_cell_phonetic": ((0, 1),),
    "Sheet.set_cell_phonetic_properties": ((0, 1),),
    "Sheet.set_cell_phonetic_runs": ((0, 1),),
    "Sheet.set_cell_show_phonetic": ((0, 1),),
    "Sheet.set_cell_xf_index": ((0, 1),),
    "Sheet.set_comment": ((0, 1),),
    "Sheet.set_comment_visible": ((0, 1),),
    "Sheet.set_error": ((0, 1),),
    "Sheet.set_formula": ((0, 1),),
    "Sheet.set_hyperlink": ((0, 1),),
    "Sheet.set_number": ((0, 1),),
    "Sheet.set_text": ((0, 1),),
    "Sheet.set_value": ((0, 1),),
    "Sheet.set_values": ((0, 1),),
    "Sheet.validate_value": ((0, 1),),
}
RECT_ARGS = {
    "Sheet.add_merge": (0,),
    "Sheet.delete_cells": (0,),
    "Sheet.evaluate_cf_range": (0,),
    "Sheet.get_cell_rect_pt": (0,),
    "Sheet.get_cells_in_range": (0,),
    "Sheet.get_merges_in_range": (0,),
    "Sheet.get_range_xf_bases": (0,),
    "Sheet.get_range_xf_bases_excluding": (0,),
    "Sheet.get_typed_values": (0,),
    "Sheet.get_values": (0,),
    "Sheet.insert_cells": (0,),
    "Sheet.iter_rows": (0,),
    "Sheet.list_cells": (0,),
    "Sheet.recalc_viewport": (0,),
    "Sheet.remap_range_xfs": (0,),
    "Sheet.remove_merge": (0,),
    "Sheet.set_print_area": (0,),
    "Sheet.set_range_xf_index": (0,),
    "Sheet.snapshot_range_xfs": (0,),
    "Sheet.update_cell_format": (0,),
    "Workbook.format_reference": (0,),
}
SPAN_ARGS = {
    "Sheet.delete_columns": ((0, "col", "span"),),
    "Sheet.delete_rows": ((0, "row", "span"),),
    "Sheet.get_column_width_pt": ((0, "col", "single"),),
    "Sheet.get_column_xf_index": ((0, "col", "single"),),
    "Sheet.get_row_height_pt": ((0, "row", "single"),),
    "Sheet.get_row_xf_index": ((0, "row", "single"),),
    "Sheet.insert_columns": ((0, "col", "span"),),
    "Sheet.insert_rows": ((0, "row", "span"),),
    "Sheet.reset_column_width": ((0, "col", "span"),),
    "Sheet.reset_column_xf_index": ((0, "col", "span"),),
    "Sheet.reset_row_height": ((0, "row", "span"),),
    "Sheet.reset_row_xf_index": ((0, "row", "span"),),
    "Sheet.set_column_collapsed": ((0, "col", "span"),),
    "Sheet.set_column_hidden": ((0, "col", "span"),),
    "Sheet.set_column_outline": ((0, "col", "span"),),
    "Sheet.set_column_width": ((0, "col", "span"),),
    "Sheet.set_column_xf_index": ((0, "col", "span"),),
    "Sheet.set_row_auto_height": ((0, "row", "span"),),
    "Sheet.set_row_collapsed": ((0, "row", "span"),),
    "Sheet.set_row_height": ((0, "row", "span"),),
    "Sheet.set_row_hidden": ((0, "row", "span"),),
    "Sheet.set_row_outline": ((0, "row", "span"),),
    "Sheet.set_row_xf_index": ((0, "row", "span"),),
}
# </facade-tables>

#: A value :meth:`Sheet.set_value` stores.
CellValueInput = Union[
    None, bool, int, float, decimal.Decimal, str, ExcelError, Value, datetime.date, datetime.datetime, datetime.time
]

_LAST_ROW = 1048575
_LAST_COL = 16383
# The built-in number format each en-US date code a date write gives a General cell names.
_BUILTIN_DATE_NUM_FMT_IDS = {"m/d/yyyy": 14, "m/d/yyyy h:mm": 22, "h:mm:ss": 21}
# Largest rectangle :meth:`Sheet.get_values` reads.
_GET_VALUES_CELL_CAP = 1 << 22
_BLANK = Value(ValueKind.BLANK)
_WHOLE_SHEET = CellRect(first_row=0, first_col=0, last_row=_LAST_ROW, last_col=_LAST_COL)

# An argument slot the caller left empty; ``None`` is a real argument value.
_MISSING: Any = object()

_F = TypeVar("_F", bound=Callable[..., Any])
_COORDINATE_FORMS: Dict[str, inspect.Signature] = {}


def _coordinate_form(stub: _F) -> _F:
    """Record ``stub``, a method's coordinate-form overload, as that method's ``inspect`` signature."""
    _COORDINATE_FORMS[stub.__qualname__] = inspect.signature(stub)
    return stub


def _span_form(stub: _F) -> _F:
    """Record ``stub``, a method's number-pair overload of a row or column span, as its ``inspect`` signature."""
    _COORDINATE_FORMS[stub.__qualname__] = inspect.signature(stub)
    return stub


def _bind(args: Tuple[Any, ...], kwargs: Dict[str, Any], names: Sequence[str]) -> List[Any]:
    """``args`` and ``kwargs`` bound to the parameters ``names``, every one required."""
    if len(args) > len(names):
        raise TypeError(f"{_call_label()}() takes at most {len(names)} positional arguments")
    bound = dict(zip(names, args))
    for key, value in kwargs.items():
        if key not in names:
            raise TypeError(f"{_call_label()}() got an unexpected keyword argument {key!r}")
        if key in bound:
            raise TypeError(f"{_call_label()}() got multiple values for argument {key!r}")
        bound[key] = value
    missing = [name for name in names if name not in bound]
    if missing:
        raise TypeError(f"{_call_label()}() missing required argument {missing[0]!r}")
    return [bound[name] for name in names]


def _split_cell_args(
    args: Tuple[Any, ...],
    kwargs: Dict[str, Any],
    pairs: Tuple[Tuple[int, int], ...],
    rest_names: Sequence[Union[str, Tuple[str, Any]]],
) -> Tuple[List[Union[Tuple[Any, Any], str]], List[Any]]:
    """Split a cell method's arguments into its cells and its other arguments.

    ``pairs`` holds the ``(row, col)`` positions of the coordinate form, as
    ``CELL_ARGS`` lists them. ``rest_names`` names the other parameters in
    order; ``(name, default)`` marks an optional one. Cells are positional:
    either every pair is two indexes or every pair is one A1 string, chosen
    by the type at the first pair's position. The other parameters are
    positional or keyword.

    Returns:
      Each cell, as a ``(row, col)`` tuple or an unresolved A1 string, and
      the other arguments in ``rest_names`` order.

    Raises:
      TypeError: for a missing, surplus, repeated or unknown argument, or
        for cells given in both forms.
    """
    names = [spec if isinstance(spec, str) else spec[0] for spec in rest_names]
    a1 = len(args) > pairs[0][0] and isinstance(args[pairs[0][0]], str)
    row_at = {i: n for n, (i, _) in enumerate(pairs)}
    col_at = {j: n for n, (_, j) in enumerate(pairs)}
    # Each positional slot of the chosen form: ("row" | "col" | "rest", index).
    layout: List[Tuple[str, int]] = []
    rest_count = 0
    for position in range(2 * len(pairs) + len(names)):
        if position in row_at:
            layout.append(("row", row_at[position]))
        elif position in col_at:
            if not a1:
                layout.append(("col", col_at[position]))
        else:
            layout.append(("rest", rest_count))
            rest_count += 1
    if len(args) > len(layout):
        form = "A1 strings" if a1 else "(row, col) pairs"
        raise TypeError(f"{_call_label()}() takes at most {len(layout)} positional arguments with cells as {form}")
    rows: List[Any] = [_MISSING] * len(pairs)
    cols: List[Any] = [_MISSING] * len(pairs)
    rest: List[Any] = [_MISSING] * len(names)
    for (kind, index), arg in zip(layout, args):
        target = rows if kind == "row" else cols if kind == "col" else rest
        target[index] = arg
    if any(v is _MISSING for v in rows) or (not a1 and any(v is _MISSING for v in cols)):
        raise TypeError(_argument_message("cell", "must be (row, col) or one A1 string"))
    if any(isinstance(v, str) != a1 for v in rows) or any(isinstance(v, str) for v in cols):
        raise TypeError(_argument_message("cell", "must be (row, col) or one A1 string"))
    for key, value in kwargs.items():
        if key not in names:
            raise TypeError(f"{_call_label()}() got an unexpected keyword argument {key!r}")
        index = names.index(key)
        if rest[index] is not _MISSING:
            raise TypeError(f"{_call_label()}() got multiple values for argument {key!r}")
        rest[index] = value
    for index, spec in enumerate(rest_names):
        if rest[index] is _MISSING:
            if isinstance(spec, str):
                raise TypeError(f"{_call_label()}() missing required argument {spec!r}")
            rest[index] = spec[1]
    cells: List[Union[Tuple[Any, Any], str]] = list(rows) if a1 else list(zip(rows, cols))
    return cells, rest


def _geometry_mode(mode: Any) -> int:
    """The C ordinal of ``mode``, checked to be a :class:`GeometryMode` member.

    Raises:
      TypeError: when ``mode`` is not an integer.
      InvalidArgumentError: when ``mode`` is not a :class:`GeometryMode` member.
    """
    if isinstance(mode, bool) or not isinstance(mode, numbers.Integral):
        raise TypeError(_argument_message("mode", "must be a GeometryMode"))
    try:
        return int(GeometryMode(mode))
    except ValueError:
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT, _argument_message("mode", f"{mode!r} is not a GeometryMode"), f"mode={mode!r}"
        ) from None


def _cell_input(value: Any) -> Tuple[str, Tuple[Any, ...]]:
    """The setter :meth:`Sheet.set_value` calls for ``value`` and the arguments after the cell.

    Raises:
      TypeError: for a type no cell holds (a ``timedelta`` included), or a
        :class:`Value` of kind ``ARRAY``, ``REF`` or ``LAMBDA``.
      ValueError: for an integer too large to convert to a float, or a
        ``datetime`` / ``time`` carrying a time zone.
      InvalidArgumentError: for a NaN or infinite number.
    """
    if isinstance(value, Value):
        kind = value.kind
        if kind is ValueKind.BLANK:
            return "set_blank", ()
        if kind is ValueKind.NUMBER:
            value = value.number
        elif kind is ValueKind.BOOL:
            return "set_bool", (value.boolean,)
        elif kind is ValueKind.TEXT:
            return "set_text", (value.text,)
        elif kind is ValueKind.ERROR:
            return "set_error", (value.error_code,)
        else:
            raise TypeError(_argument_message("value", f"of kind {kind.name} cannot be stored in a cell"))
    if value is None:
        return "set_blank", ()
    if isinstance(value, bool):
        return "set_bool", (value,)
    if isinstance(value, decimal.Decimal):
        value = float(value)
    if isinstance(value, numbers.Real):
        try:
            number = float(value)
        except OverflowError as exc:
            raise ValueError(_argument_message("value", "is too large")) from exc
        if not math.isfinite(number):
            raise FormulonError(StatusCode.INVALID_ARGUMENT, _argument_message("value", "must be finite"), "")
        return "set_number", (number,)
    if isinstance(value, str):
        return ("set_formula" if value.startswith("=") else "set_text"), (value,)
    if isinstance(value, ExcelError):
        return "set_error", (value.code,)
    if isinstance(value, (datetime.date, datetime.time)):
        if getattr(value, "tzinfo", None) is not None:
            raise ValueError(_argument_message("value", f"must be a naive {type(value).__name__}"))
        return "_set_date", (value,)
    raise TypeError(_argument_message("value", "must be a cell value"))


def _value_of_input(value: Any, *, date1904: bool) -> Value:
    """A :class:`Value` of any :data:`CellValueInput`; a ``str`` is text, a date its serial.

    Raises:
      TypeError: for a type no cell holds.
    """
    if isinstance(value, Value):
        return value
    setter, payload = _cell_input(value)
    if setter == "set_blank":
        return Value(kind=ValueKind.BLANK)
    if setter == "set_bool":
        return Value(kind=ValueKind.BOOL, boolean=payload[0])
    if setter == "set_error":
        return Value(kind=ValueKind.ERROR, error_code=payload[0])
    if setter == "_set_date":
        return Value(kind=ValueKind.NUMBER, number=date_serial(payload[0], date1904=date1904)[0])
    if setter == "set_number":
        return Value(kind=ValueKind.NUMBER, number=payload[0])
    return Value(kind=ValueKind.TEXT, text=value)


def _normalize_auto_filter(auto_filter: Union[AutoFilterInput, AutoFilter], workbook: Workbook) -> AutoFilter:
    """The :class:`AutoFilter` ``auto_filter`` describes, with the A1 text of its ranges resolved."""
    if not isinstance(auto_filter, (AutoFilterInput, AutoFilter)):
        raise TypeError(_argument_message("auto_filter", "must be an AutoFilterInput"))
    sort = auto_filter.sort
    if sort is not None:
        sort = dataclasses.replace(
            sort,
            range=cell_rect(workbook, sort.range, "auto_filter.sort.range"),
            conditions=[
                dataclasses.replace(c, range=cell_rect(workbook, c.range, f"auto_filter.sort.conditions[{i}].range"))
                for i, c in enumerate(sort.conditions)
            ],
        )
    return AutoFilter(
        range=cell_rect(workbook, auto_filter.range, "auto_filter.range"),
        columns=list(auto_filter.columns),
        sort=sort,
    )


def _cell_rects(workbook: Workbook, value: Any, label: str) -> List[CellRect]:
    """A list of :class:`CellRect` or unqualified A1 ranges, or one A1 string, as :class:`CellRect` objects."""
    items = [value] if isinstance(value, str) else value
    if isinstance(items, (bytes, dict)) or not hasattr(items, "__iter__"):
        raise TypeError(_argument_message(label, "must be an array or one A1 string"))
    return [cell_rect(workbook, item, f"{label}[{i}]") for i, item in enumerate(items)]


def _range_cell(row: int, col: int, formula: Optional[str], value: Value) -> RangeCell:
    return RangeCell(row=row, col=col, formula=formula, value=to_native(value), kind=value.kind)


def _with_cell_context(exc: FormulonError, row: int, col: int) -> FormulonError:
    """``exc`` with ``row=`` / ``col=`` naming the failing cell in its context, unless it names one already."""
    if "row=" not in exc.context:
        exc.context = f"{exc.context} row={row} col={col}".strip()
        exc.args = (f"{exc.message} [{exc.context}]",)
    return exc


# ---------------------------------------------------------------------------
# Cell formats: ``Sheet.get_cell_format`` / ``Sheet.update_cell_format`` over the
# style records. tools/api/cell_format_cases.json pins the behaviour for the
# Python and JS facades alike.
# ---------------------------------------------------------------------------

# How each key of each group of a format patch is checked: ``(kind, nullable,
# argument)``; a nullable key takes ``None`` to clear the attribute. The keys
# are the patch type's keys and the read record's fields.
_FORMAT_FLAG = ("boolean", True, None)
_FORMAT_BOOLEAN = ("boolean", False, None)
_FORMAT_COLOR = ("color", False, None)
_FORMAT_SIDE = ("side", False, None)
_FORMAT_FIELDS: Dict[str, Dict[str, Tuple[str, bool, Any]]] = {
    "font": {
        "name": ("string", False, None),
        "size_pt": ("size", True, None),
        "bold": _FORMAT_FLAG,
        "italic": _FORMAT_FLAG,
        "strike": _FORMAT_FLAG,
        "outline": _FORMAT_FLAG,
        "shadow": _FORMAT_FLAG,
        "condense": _FORMAT_FLAG,
        "extend": _FORMAT_FLAG,
        "underline": ("enum", False, UnderlineStyle),
        "vert_align": ("enum", False, FontVerticalAlign),
        "family": ("enum", True, FontFamily),
        "charset": ("byte", True, None),
        "scheme": ("enum", False, FontScheme),
        "color": _FORMAT_COLOR,
    },
    "fill": {
        "pattern": ("enum", False, FillPattern),
        "foreground": _FORMAT_COLOR,
        "background": _FORMAT_COLOR,
        "gradient": ("gradient", True, None),
    },
    "border": {
        "left": _FORMAT_SIDE,
        "right": _FORMAT_SIDE,
        "top": _FORMAT_SIDE,
        "bottom": _FORMAT_SIDE,
        "diagonal": _FORMAT_SIDE,
        "diagonal_up": _FORMAT_BOOLEAN,
        "diagonal_down": _FORMAT_BOOLEAN,
    },
    "side": {"style": ("enum", False, BorderStyle), "color": _FORMAT_COLOR},
    "alignment": {
        "horizontal": ("enum", False, HorizontalAlign),
        "vertical": ("enum", False, VerticalAlign),
        "wrap_text": _FORMAT_BOOLEAN,
        "shrink_to_fit": _FORMAT_BOOLEAN,
        "indent": ("byte", False, None),
        "text_rotation": ("rotation", False, None),
        "reading_order": ("enum", False, ReadingOrder),
        "justify_last_line": _FORMAT_BOOLEAN,
        "relative_indent": ("signed", False, None),
    },
    "protection": {"locked": _FORMAT_BOOLEAN, "hidden": _FORMAT_BOOLEAN},
}

_APPLY_FLAGS = {
    "number_format": "apply_number_format",
    "font": "apply_font",
    "fill": "apply_fill",
    "border": "apply_border",
    "alignment": "apply_alignment",
    "protection": "apply_protection",
}
_INT32_MAX = 0x7FFFFFFF


def _out_of_range(label: str, value: Any) -> FormulonError:
    return FormulonError(
        StatusCode.INVALID_ARGUMENT, _argument_message(label, f"{value} is out of range"), f"value={value}"
    )


def _format_count(value: Any, label: str) -> int:
    """``value`` as a non-negative integer: ``TypeError`` for a non-number, ``ValueError`` for a fraction or negative."""
    if isinstance(value, bool) or not isinstance(value, numbers.Real):
        raise TypeError(_argument_message(label, "must be an integer"))
    if not isinstance(value, numbers.Integral) and not (math.isfinite(value) and float(value).is_integer()):
        raise TypeError(_argument_message(label, "must be an integer"))
    if value < 0:
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT, _argument_message(label, "must not be negative"), f"value={value}"
        )
    return int(value)


def _format_color(value: Any, label: str) -> Color:
    """A ``ColorInput`` as a :class:`Color`, its ``kind`` checked against :class:`ColorKind`."""
    color = color_input(value, label)
    if _format_count(color.kind, f"{label}.kind") > max(ColorKind):
        raise _out_of_range(f"{label}.kind", color.kind)
    return color


def _format_value(spec: Tuple[str, bool, Any], value: Any, label: str) -> Any:
    """One checked key of a format patch."""
    kind, nullable, arg = spec
    if value is None:
        if nullable:
            return None
        raise TypeError(_argument_message(label, "must not be null"))
    if kind == "string":
        if not isinstance(value, str):
            raise TypeError(_argument_message(label, "must be a string"))
        return value
    if kind == "boolean":
        return _require_bool(value, label)
    if kind == "gradient":
        if not isinstance(value, GradientFill):
            raise TypeError(_argument_message(label, "must be a GradientFill"))
        return value
    if kind == "color":
        return _format_color(value, label)
    if kind == "side":
        return _format_group("side", value, label)
    if kind == "size":
        if isinstance(value, bool) or not isinstance(value, numbers.Real):
            raise TypeError(_argument_message(label, "must be a number"))
        if not math.isfinite(value):
            raise ValueError(_argument_message(label, "must be finite"))
        return float(value)
    if kind == "enum":
        number = _format_count(value, label)
        if number > max(arg):
            raise _out_of_range(label, value)
        return arg(number)
    if kind == "byte":
        number = _format_count(value, label)
        if number > 255:
            raise _out_of_range(label, value)
        return number
    if kind == "rotation":
        number = _format_count(value, label)
        if number > 180 and number != 255:
            raise _out_of_range(label, value)
        return number
    # signed
    if isinstance(value, bool) or not isinstance(value, numbers.Real):
        raise TypeError(_argument_message(label, "must be an integer"))
    if not isinstance(value, numbers.Integral) and not (math.isfinite(value) and float(value).is_integer()):
        raise TypeError(_argument_message(label, "must be an integer"))
    if value > _INT32_MAX or value < -_INT32_MAX - 1:
        raise _out_of_range(label, value)
    return int(value)


def _format_group(group: str, value: Any, label: str) -> Dict[str, Any]:
    """One group of a format patch, checked: a patch mapping or the read record of that group."""
    if dataclasses.is_dataclass(value) and not isinstance(value, type):
        items = {f.name: getattr(value, f.name) for f in dataclasses.fields(value)}
    elif isinstance(value, Mapping):
        items = dict(value)
    else:
        raise TypeError(_argument_message(label, "must be an object"))
    fields = _FORMAT_FIELDS[group]
    out: Dict[str, Any] = {}
    for key, item in items.items():
        if key not in fields:
            raise TypeError(_argument_message(label, f"has no field {key}"))
        out[key] = _format_value(fields[key], item, f"{label}.{key}")
    return out


def _check_cell_format_patch(given: Dict[str, Any]) -> Dict[str, Any]:
    """Every given argument of :meth:`Sheet.update_cell_format`, checked before anything is written."""
    out: Dict[str, Any] = {}
    for key, value in given.items():
        if value is UNCHANGED:
            continue
        if key == "number_format":
            if not isinstance(value, str):
                raise TypeError(_argument_message("number_format", "must be a string"))
            if value == "":
                raise FormulonError(
                    StatusCode.INVALID_ARGUMENT, _argument_message("number_format", "must not be empty"), ""
                )
            out[key] = value
        else:
            out[key] = _format_group(key, value, key)
    return out


def _merge_font(current: Font, patch: Dict[str, Any]) -> Font:
    return dataclasses.replace(current, **patch)


def _merge_fill(current: Fill, patch: Dict[str, Any]) -> Fill:
    """Setting a colour without a pattern makes the fill solid (Excel's fill colour); a fill patch drops a
    gradient unless it names one."""
    changes: Dict[str, Any] = {"gradient": patch.get("gradient")}
    if "pattern" in patch:
        changes["pattern"] = patch["pattern"]
    for key in ("foreground", "background"):
        if key in patch:
            changes[key] = patch[key]
    colored = "foreground" in patch or "background" in patch
    if colored and "pattern" not in patch and current.pattern == FillPattern.NONE:
        changes["pattern"] = FillPattern.SOLID
    return dataclasses.replace(current, **changes)


def _merge_border(current: Border, patch: Dict[str, Any]) -> Border:
    changes: Dict[str, Any] = {}
    for key, value in patch.items():
        if key in ("diagonal_up", "diagonal_down"):
            changes[key] = value
            continue
        changes[key] = dataclasses.replace(getattr(current, key), **value)
    return dataclasses.replace(current, **changes)


def _cell_xf_input(xf: CellXf) -> Dict[str, Any]:
    """The :class:`CellXfInput` fields that rebuild ``xf`` as ``add_cell_xf`` reads them."""
    return {f.name: getattr(xf, f.name) for f in dataclasses.fields(CellXfInput)}


def _patched_xf(workbook: Workbook, xf: CellXf, patch: Dict[str, Any], num_fmt_id: Optional[int]) -> CellXfInput:
    """``xf`` with ``patch`` merged in, as the :class:`CellXfInput` of the new record."""
    fields = _cell_xf_input(xf)
    for group in patch:
        fields[_APPLY_FLAGS[group]] = True
    if num_fmt_id is not None:
        fields["num_fmt_id"] = num_fmt_id
    if "font" in patch:
        fields["font_index"] = workbook.ooxml.add_font(
            _merge_font(workbook.ooxml.get_font(xf.font_index), patch["font"])
        )
    if "fill" in patch:
        fields["fill_index"] = workbook.ooxml.add_fill(
            _merge_fill(workbook.ooxml.get_fill(xf.fill_index), patch["fill"])
        )
    if "border" in patch:
        fields["border_index"] = workbook.ooxml.add_border(
            _merge_border(workbook.ooxml.get_border(xf.border_index), patch["border"])
        )
    if "alignment" in patch:
        fields.update(patch["alignment"])
    if "protection" in patch:
        fields["locked"] = patch["protection"].get("locked", True if fields["locked"] is None else fields["locked"])
        fields["hidden"] = patch["protection"].get("hidden", False if fields["hidden"] is None else fields["hidden"])
    return CellXfInput(**fields)


def _cell_format_of(workbook: Workbook, xf: CellXf, source: EffectiveStyleSource) -> CellFormat:
    """The :class:`CellFormat` of the cell XF record ``xf``, found at ``source``."""
    code = None if xf.num_fmt_id == 0 else workbook.ooxml.get_number_format(xf.num_fmt_id)
    return CellFormat(
        source=source,
        number_format="General" if code is None else code,
        font=workbook.ooxml.get_font(xf.font_index),
        fill=workbook.ooxml.get_fill(xf.fill_index),
        border=workbook.ooxml.get_border(xf.border_index),
        alignment=Alignment(
            horizontal=HorizontalAlign.GENERAL if xf.horizontal is None else HorizontalAlign(xf.horizontal),
            vertical=VerticalAlign.BOTTOM if xf.vertical is None else VerticalAlign(xf.vertical),
            wrap_text=bool(xf.wrap_text),
            shrink_to_fit=bool(xf.shrink_to_fit),
            indent=xf.indent or 0,
            text_rotation=xf.text_rotation or 0,
            reading_order=ReadingOrder.CONTEXT if xf.reading_order is None else ReadingOrder(xf.reading_order),
            justify_last_line=bool(xf.justify_last_line),
            relative_indent=xf.relative_indent or 0,
        ),
        protection=CellProtection(
            locked=True if xf.locked is None else xf.locked,
            hidden=False if xf.hidden is None else xf.hidden,
        ),
    )


def _read_auto_filter(workbook: "Workbook", fn, *indices: int) -> Optional[AutoFilter]:
    """Call a ``(..., out_present*, out_filter*)`` AutoFilter getter; ``None`` when the target has no filter."""
    h = workbook._require()
    with _ArgScope() as scope:
        present = scope.out_ptr()
        ptr = scope.struct(S.AUTO_FILTER)
        _check(fn(h, *(_uint(i, "index", role="position") for i in indices), present, ptr))
        if LIB.read_u32(present) == 0:
            return None
        return _decode_auto_filter(ptr)


class Sheet(_NoCopy):
    """One sheet of a :class:`Workbook`; every sheet-scope operation lives here.

    Obtain a handle from :meth:`Workbook.sheet`, ``wb[key]``, iteration or
    :meth:`Workbook.list_sheets`. A workbook hands out one handle per sheet,
    so ``wb[0] is wb["Data"]`` when the first sheet is named ``Data``. The
    handle follows its sheet through :meth:`Workbook.add_sheet`,
    :meth:`Workbook.copy_sheet`, :meth:`move`, :meth:`rename` and
    :meth:`Pivot.show_details`. Once
    :meth:`Workbook.remove_sheet` removes the sheet, every operation raises
    :class:`NotFoundError`; once the workbook is closed, every operation
    except reading :attr:`workbook` raises :class:`WorkbookClosedError`.

    Rows and columns are 0-based. Every method taking a ``(row, col)`` pair
    also takes the cell as one unqualified A1 string in its place
    (``set_value("B3", 1)`` is ``set_value(2, 1, 1)``); a method taking two
    cells takes both pairs or both strings. The pair is positional only;
    the arguments after it may be passed by keyword. A range, a
    sheet-qualified reference, R1C1 text, surrounding whitespace or anything
    else :meth:`Workbook.parse_reference` does not read as one cell raises
    :class:`InvalidArgumentError`. A method taking a :class:`CellRect` also
    takes an unqualified A1 range (``"A1:B2"``; one cell is a 1x1 range).
    A method taking a column or row span also takes it as one A1 span string
    in its place: ``"B"`` or ``"B:D"`` for columns, ``"3"`` or ``"3:5"`` for
    rows (``set_column_width("B:D", 12)``); the whole sheet suits either axis.
    Every row or column span, insert and delete included, is the closed
    interval ``(first, last)`` or an A1 span string (``insert_rows(2, 4)``
    opens three rows).

    To read cells, pick by shape: :meth:`get_values` for a dense grid of
    values, :meth:`list_cells` for the populated cells of a range,
    :meth:`get_cells_in_range` to page through a large range, and
    :meth:`iter_rows` for the rows of the used range. A read returns native
    values (:data:`NativeValue`); its ``*_typed_*`` variant returns tagged
    :class:`Value` objects. A read computes dirty cells first under
    ``CalcMode.AUTO`` and returns the values as of the last recalc under
    ``CalcMode.MANUAL``; :meth:`evaluate_formula` computes every precedent it
    reads, even in manual mode, without storing the results.

    ``sheet["A1"]`` and ``sheet[row, col]`` read the cell's native value and
    ``sheet["A1:B2"]`` the grid :meth:`get_values` reads; ``sheet["A1"] = v``
    stores ``v`` as :meth:`set_value` does, and ``sheet["A1:B2"] = grid``
    stores a grid of that shape as :meth:`set_values` does. An ``int`` or a
    slice key raises ``TypeError``.

    The cell, row and column XF indices and the raw worksheet XML are
    reached through :attr:`ooxml`.

    Tables, pivot tables and drawing objects are handles too: :meth:`table`,
    :meth:`pivot`, :meth:`drawing` and the per-kind accessors return the same
    :class:`Table`, :class:`Pivot` or :class:`Drawing` object for the same
    table, pivot or object every time.
    """

    __slots__ = ("_workbook", "_position", "_name", "_pivots", "_drawings", "_ooxml")

    def __init__(self, *args: object, **kwargs: object) -> None:
        raise TypeError("obtain a Sheet from Workbook.sheet()")

    @classmethod
    def _new(cls, workbook: Workbook, position: int, name: str) -> "Sheet":
        handle = object.__new__(cls)
        handle._workbook = workbook
        handle._ooxml = SheetOoxml._new(handle)
        # The current 0-based position, or None once the sheet is removed.
        handle._position = position
        handle._name = name
        # One handle per pivot runtime id and per drawing object id, never dropped.
        handle._pivots = {}
        handle._drawings = {}
        return handle

    @property
    def workbook(self) -> Workbook:
        """The workbook this sheet belongs to."""
        return self._workbook

    @property
    def ooxml(self) -> SheetOoxml:
        """The OOXML-level layer of the sheet: XF indices and raw worksheet XML. Always the same object."""
        return self._ooxml

    @property
    def index(self) -> int:
        """The sheet's current 0-based position in the workbook.

        Raises:
          NotFoundError: when the sheet was removed.
          WorkbookClosedError: when the workbook is closed.
        """
        return self._index()

    @property
    def name(self) -> str:
        """The sheet's display name.

        Raises:
          NotFoundError: when the sheet was removed.
          WorkbookClosedError: when the workbook is closed.
        """
        self._index()
        return self._name

    def _index(self) -> int:
        self._workbook._require()
        if self._position is None:
            raise FormulonError(StatusCode.NOT_FOUND, "the sheet was removed", "")
        return self._position

    def __repr__(self) -> str:
        if self._position is None:
            return "Sheet(removed)"
        return f"Sheet(index={self._position}, name={self._name!r})"

    def rename(self, name: str) -> None:
        """Rename the sheet to ``name``; this handle keeps addressing it under the new name.

        Raises:
          FormulonError: with ``StatusCode.INVALID_SHEET_NAME`` under the
            :meth:`Workbook.add_sheet` naming rules.
        """
        index = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            _check(LIB.fm_workbook_rename_sheet(h, _uint(index, "sheet_index", role="sheet"), name_ptr))
        self._name = self._workbook._sheet_name(index)

    def move(self, insert_before: Optional[SheetRef]) -> None:
        """Move the sheet before ``insert_before`` (a sheet, a name or a position in the current order).

        ``None`` or the sheet count moves it to the end; moving a sheet before itself or before the sheet
        after it changes nothing. Every :class:`Sheet` handle keeps pointing at its sheet.

        Raises:
          IndexOutOfRangeError: when ``insert_before`` is a position past the sheet count.
        """
        from_index = self._index()
        h = self._workbook._require()
        handles = self._workbook._sheet_handles()
        before = self._workbook._insert_before_position(insert_before, len(handles))
        _check(
            LIB.fm_workbook_move_sheet(
                h, _uint(from_index, "from_index", role="sheet"), _uint(before, "insert_before", role="sheet")
            )
        )
        to_index = before if before < from_index else max(from_index, before - 1)
        handles.insert(to_index, handles.pop(from_index))
        self._workbook._renumber_sheets(handles)

    # -- A1 access ---------------------------------------------------------
    def _cell_args(
        self,
        name: str,
        args: Tuple[Any, ...],
        kwargs: Dict[str, Any],
        rest_names: Sequence[Union[str, Tuple[str, Any]]],
    ) -> List[Any]:
        """The arguments of method ``name`` in coordinate-form order, each A1 cell resolved to its row and column."""
        pairs = CELL_ARGS[f"Sheet.{name}"]
        cells, rest = _split_cell_args(args, kwargs, pairs, rest_names)
        resolved = [self._cell(cell) if isinstance(cell, str) else cell for cell in cells]
        row_at = {i: n for n, (i, _) in enumerate(pairs)}
        col_at = {j: n for n, (_, j) in enumerate(pairs)}
        others = iter(rest)
        return [
            resolved[row_at[p]][0] if p in row_at else resolved[col_at[p]][1] if p in col_at else next(others)
            for p in range(2 * len(pairs) + len(rest))
        ]

    def _rect(self, rect: Union[CellRect, str]) -> CellRect:
        """``rect`` itself, or the rectangle an unqualified A1 range or cell names."""
        if not isinstance(rect, str):
            return rect
        ref = self._workbook.parse_reference(rect)
        if ref is None or ref.sheet_name is not None:
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _argument_message("range", f"{rect!r} is not an unqualified A1 range"),
                "",
            )
        return ref.range

    def _span(self, text: str, axis: str) -> Tuple[int, int]:
        """The first and last column (``axis`` ``"col"``) or row (``"row"``) the A1 span ``text`` covers.

        A column span is ``"B"`` or ``"B:D"``, a row span ``"3"`` or
        ``"3:5"``; the whole sheet (``"A:XFD"``, ``"1:1048576"``) is both.
        """
        ref = self._workbook.parse_reference(text if ":" in text else f"{text}:{text}")
        rect = None if ref is None or ref.sheet_name is not None else ref.range
        if rect is not None and axis == "col" and rect.first_row == 0 and rect.last_row == _LAST_ROW:
            return rect.first_col, rect.last_col
        if rect is not None and axis == "row" and rect.first_col == 0 and rect.last_col == _LAST_COL:
            return rect.first_row, rect.last_row
        noun = "whole-column" if axis == "col" else "whole-row"
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT,
            _argument_message("span", f"{text!r} is not a {noun} A1 span without a sheet name"),
            "",
        )

    def _span_args(self, name: str, args: Tuple[Any, ...], kwargs: Dict[str, Any], names: Sequence[str]) -> List[Any]:
        """The arguments of method ``name`` in number-pair order, each span string resolved.

        ``SPAN_ARGS`` gives the span positions: a ``span`` form fills
        ``(first, last)``, a ``start_count`` form ``(first, count)``. The span
        string is positional or the keyword ``span``.
        """
        spans = SPAN_ARGS[f"Sheet.{name}"]
        if "span" in kwargs:
            at = spans[0][0]
            if len(args) > at:
                raise TypeError(f"{_call_label()}() got multiple values for argument 'span'")
            kwargs = dict(kwargs)
            args = (*args, kwargs.pop("span"))
        for at, axis, form in spans:
            if len(args) > at and isinstance(args[at], str):
                first, last = self._span(args[at], axis)
                pair = (first, last) if form == "span" else (first, last - first + 1)
                args = (*args[:at], *pair, *args[at + 1 :])
        return _bind(args, kwargs, names)

    def _cell(self, a1: str) -> Tuple[int, int]:
        ref = self._workbook.parse_reference(a1)
        if ref is not None and ref.sheet_name is None:
            rect = ref.range
            if rect.first_row == rect.last_row and rect.first_col == rect.last_col:
                return rect.first_row, rect.first_col
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT,
            _argument_message("cell", f"{a1!r} is not a single-cell A1 reference"),
            "",
        )

    def _key(self, key: Union[str, Tuple[int, int]]) -> Tuple[int, int]:
        """``(row, col)`` from a ``sheet[...]`` key: one A1 cell or a ``(row, col)`` pair."""
        if isinstance(key, str):
            return self._cell(key)
        if isinstance(key, tuple) and len(key) == 2 and all(isinstance(k, numbers.Integral) for k in key):
            return key[0], key[1]
        raise TypeError(_argument_message("key", "must be an A1 string or a (row, col) pair"))

    def __iter__(self) -> Iterator[Any]:
        """Not iterable: a sheet is not a sequence of cells; use :meth:`iter_rows`.

        Raises:
          TypeError: always.
        """
        raise TypeError("Sheet is not iterable; use iter_rows() to walk its cells")

    def __contains__(self, item: object) -> bool:
        """Not a container: membership is undefined for a sheet; use :meth:`iter_rows`.

        Raises:
          TypeError: always.
        """
        raise TypeError("Sheet has no membership test; use iter_rows() to walk its cells")

    def __getitem__(self, key: Union[str, Tuple[int, int]]) -> NativeValue:
        if isinstance(key, str) and ":" in key:
            return self.get_values(key)
        row, col = self._key(key)
        return self.get_value(row, col)

    def __setitem__(self, key: Union[str, Tuple[int, int]], value: Any) -> None:
        if isinstance(key, str) and ":" in key:
            rect = self._rect(key)
            rows = rect.last_row - rect.first_row + 1
            cols = rect.last_col - rect.first_col + 1
            if isinstance(value, (str, bytes)) or not isinstance(value, Sequence):
                raise TypeError(_argument_message("value", "must be an array of rows"))
            if len(value) != rows or any(not isinstance(line, Sequence) or len(line) != cols for line in value):
                raise FormulonError(
                    StatusCode.INVALID_ARGUMENT,
                    _argument_message("value", f"must be a {rows}x{cols} grid"),
                    "",
                )
            self.set_values(rect.first_row, rect.first_col, value)
            return
        row, col = self._key(key)
        self.set_value(row, col, value)

    # -- Cell mutation -----------------------------------------------------
    @overload
    @_coordinate_form
    def set_number(self, row: int, col: int, value: float) -> None: ...

    @overload
    def set_number(self, a1: str, value: float) -> None: ...

    def set_number(self, *args: Any, **kwargs: Any) -> None:
        """Store the number ``value`` at ``(row, col)``.

        ``row`` and ``col`` are 0-based.
        """
        row, col, value = self._cell_args("set_number", args, kwargs, ("value",))
        sheet = self._index()
        h = self._workbook._require()
        status = LIB.fm_sheet_set_number(
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(row, "row", role="coordinate"),
            _uint(col, "col", role="coordinate"),
            float(value),
        )
        _check(status)

    @overload
    @_coordinate_form
    def set_bool(self, row: int, col: int, value: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    @overload
    def set_bool(self, a1: str, value: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    def set_bool(self, *args: Any, **kwargs: Any) -> None:
        """Store the boolean ``value`` at ``(row, col)``.

        ``row`` and ``col`` are 0-based.
        """
        row, col, value = self._cell_args("set_bool", args, kwargs, ("value",))
        sheet = self._index()
        h = self._workbook._require()
        status = LIB.fm_sheet_set_bool(
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(row, "row", role="coordinate"),
            _uint(col, "col", role="coordinate"),
            int(_require_bool(value, "value")),
        )
        _check(status)

    @overload
    @_coordinate_form
    def set_error(self, row: int, col: int, error: Union[ErrorCode, ExcelError]) -> None: ...

    @overload
    def set_error(self, a1: str, error: Union[ErrorCode, ExcelError]) -> None: ...

    def set_error(self, *args: Any, **kwargs: Any) -> None:
        """Store an Excel cell error at ``(row, col)``.

        ``row`` and ``col`` are 0-based. ``error`` is an :class:`ErrorCode`
        member, not a display literal, or an :class:`ExcelError` as the cell
        reads return it; the same member comes back out through
        :attr:`Value.error_code`.
        """
        row, col, error = self._cell_args("set_error", args, kwargs, ("error",))
        if isinstance(error, ExcelError):
            error = error.code
        sheet = self._index()
        h = self._workbook._require()
        status = LIB.fm_sheet_set_error(
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(row, "row", role="coordinate"),
            _uint(col, "col", role="coordinate"),
            _sint(error, "error"),
        )
        _check(status)

    @overload
    @_coordinate_form
    def set_text(self, row: int, col: int, value: str) -> None: ...

    @overload
    def set_text(self, a1: str, value: str) -> None: ...

    def set_text(self, *args: Any, **kwargs: Any) -> None:
        """Store the string ``value`` at ``(row, col)``.

        ``row`` and ``col`` are 0-based. The text is stored
        verbatim; a leading ``=`` does not make it a formula (use
        :meth:`set_formula`). Text longer than 32,767 UTF-16 code units is
        rejected with ``InvalidArgumentError``.
        """
        row, col, value = self._cell_args("set_text", args, kwargs, ("value",))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            text_ptr = scope.cstr(value)
            status = LIB.fm_sheet_set_text(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
                text_ptr,
            )
            _check(status)

    @overload
    @_coordinate_form
    def set_blank(self, row: int, col: int) -> None: ...

    @overload
    def set_blank(self, a1: str) -> None: ...

    def set_blank(self, *args: Any, **kwargs: Any) -> None:
        """Clear the cell at ``(row, col)``.

        ``row`` and ``col`` are 0-based. Clearing drops the
        stored value or formula; cell formatting is unaffected.
        """
        row, col = self._cell_args("set_blank", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        status = LIB.fm_sheet_set_blank(
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(row, "row", role="coordinate"),
            _uint(col, "col", role="coordinate"),
        )
        _check(status)

    @overload
    @_coordinate_form
    def set_formula(self, row: int, col: int, formula: str) -> None: ...

    @overload
    def set_formula(self, a1: str, formula: str) -> None: ...

    def set_formula(self, *args: Any, **kwargs: Any) -> None:
        """Store a formula at ``(row, col)``.

        ``row`` and ``col`` are 0-based. The leading ``=`` is optional.

        Raises:
          FormulaSyntaxError: when the text does not parse (a syntax error, or
            a built-in called through a sheet or workbook qualifier); its
            ``start`` / ``end`` locate the failure and nothing changes.

        Example::

            ws.set_formula("C1", "=A1+B1")
            ws.get_value("C1")
        """
        row, col, formula = self._cell_args("set_formula", args, kwargs, ("formula",))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            formula_ptr = scope.cstr(formula)
            status = LIB.fm_sheet_set_formula(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
                formula_ptr,
            )
            _check(status)

    @overload
    @_coordinate_form
    def set_value(
        self,
        row: int,
        col: int,
        value: CellValueInput,  # noqa: FBT001 -- the value is any cell type; a bool is one of them
    ) -> None: ...

    @overload
    def set_value(
        self,
        a1: str,
        value: CellValueInput,  # noqa: FBT001 -- the value is any cell type; a bool is one of them
    ) -> None: ...

    def set_value(self, *args: Any, **kwargs: Any) -> None:
        """Store ``value`` at ``(row, col)``, choosing the setter by its type.

        ``None`` clears the cell (:meth:`set_blank`); a ``bool`` is a boolean
        (:meth:`set_bool`); any other real number or a ``Decimal`` is a
        number (:meth:`set_number`); a ``str`` starting with ``=`` is a
        formula (:meth:`set_formula`; ``"="`` alone is a syntax error, and
        ``+``, ``-``, ``@`` or the full-width ``＝`` do not start one); any
        other ``str`` is text (:meth:`set_text`); an :class:`ExcelError` is a
        cell error (:meth:`set_error`). A :class:`Value` is stored by its
        kind, a ``TEXT`` one always as text. Store untrusted text with
        :meth:`set_text`, which never reads it as a formula.

        A ``date``, naive ``datetime`` or naive ``time`` stores its serial
        under the workbook's date system (a ``time`` is the fraction of a
        day) and, when the cell's effective number format is General, gives
        the cell the format ``m/d/yyyy``, ``m/d/yyyy h:mm`` or ``h:mm:ss``.

        Raises:
          TypeError: for any other type (a ``timedelta`` included), or a
            :class:`Value` of kind ``ARRAY``, ``REF`` or ``LAMBDA``.
          ValueError: for an integer too large to convert to a float, or an
            aware ``datetime`` / ``time``.
          InvalidArgumentError: for a NaN or infinite number, or a date
            outside the date system.
          FormulaSyntaxError: for a formula that does not parse.

        Example::

            ws.set_value("A1", 42)
            ws.set_value("B1", "Total")
            ws.set_value("C1", "=A1*2")
            ws.set_value("D1", datetime.date(2024, 1, 31))
        """
        row, col, value = self._cell_args("set_value", args, kwargs, ("value",))
        setter, payload = _cell_input(value)
        getattr(self, setter)(row, col, *payload)

    def _set_date(self, row: int, col: int, value: Union[datetime.date, datetime.time]) -> None:
        """Store the serial of ``value`` and give a General cell the built-in date format of its kind."""
        workbook = self._workbook
        serial, code = date_serial(value, date1904=workbook.get_date1904())
        self.set_number(row, col, serial)
        ooxml = workbook.ooxml
        if ooxml.get_cell_xf(self.ooxml.get_effective_style(row, col).xf_index).num_fmt_id != 0:
            return
        rect = CellRect(first_row=row, first_col=col, last_row=row, last_col=col)
        patch: Dict[str, Any] = {"number_format": code}
        bases = self.ooxml.get_range_xf_bases(rect)
        num_fmt_id = _BUILTIN_DATE_NUM_FMT_IDS[code]
        xfs = [ooxml.add_cell_xf(_patched_xf(workbook, ooxml.get_cell_xf(base), patch, num_fmt_id)) for base in bases]
        self.ooxml.remap_range_xfs(rect, dict(zip(bases, xfs)))

    @overload
    @_coordinate_form
    def set_cell_phonetic(self, row: int, col: int, phonetic: Optional[str]) -> None: ...

    @overload
    def set_cell_phonetic(self, a1: str, phonetic: Optional[str]) -> None: ...

    def set_cell_phonetic(self, *args: Any, **kwargs: Any) -> None:
        """Set the cell's phonetic guide (OOXML ``<rPh>``, furigana).

        ``row`` and ``col`` are 0-based. ``None`` (or ``""``) removes the
        guide. The guide is display metadata attached to the
        cell's string; it does not affect evaluation. Text longer than 32,767
        UTF-16 code units is rejected with ``InvalidArgumentError``.
        """
        row, col, phonetic = self._cell_args("set_cell_phonetic", args, kwargs, ("phonetic",))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            text_ptr = scope.cstr(phonetic or "")
            status = LIB.fm_sheet_set_cell_phonetic(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
                text_ptr,
            )
            _check(status)

    @overload
    @_coordinate_form
    def set_cell_phonetic_runs(self, row: int, col: int, runs: Optional[Sequence[PhoneticRun]]) -> None: ...

    @overload
    def set_cell_phonetic_runs(self, a1: str, runs: Optional[Sequence[PhoneticRun]]) -> None: ...

    def set_cell_phonetic_runs(self, *args: Any, **kwargs: Any) -> None:
        """Set the cell's phonetic guide as one ``<rPh>`` block per run.

        ``row`` and ``col`` are 0-based. ``None`` or an empty ``runs`` removes
        the guide. Unlike :meth:`set_cell_phonetic`, which annotates the whole
        cell, this keeps the spans each reading covers -- the difference
        ``PHONETIC`` and a partially annotated cell depend on.

        The runs must be an ordered partition: each needs ``start <= end`` and
        must start at or after the previous run's ``end``.
        """
        row, col, runs = self._cell_args("set_cell_phonetic_runs", args, kwargs, ("runs",))
        sheet = self._index()
        h = self._workbook._require()
        runs = [] if runs is None else runs
        with _ArgScope() as scope:
            ptr = _pack_phonetic_run_array(runs, scope)
            status = LIB.fm_sheet_set_cell_phonetic_runs(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
                ptr,
                _uint(len(runs), "count"),
            )
            _check(status)

    @overload
    @_coordinate_form
    def get_cell_phonetic_runs(self, row: int, col: int) -> List[PhoneticRun]: ...

    @overload
    def get_cell_phonetic_runs(self, a1: str) -> List[PhoneticRun]: ...

    def get_cell_phonetic_runs(self, *args: Any, **kwargs: Any) -> List[PhoneticRun]:
        """Return the cell's ``<rPh>`` blocks with their spans; ``[]`` for an unannotated cell.

        ``row`` and ``col`` are 0-based. :meth:`get_cell_phonetic` returns the
        readings concatenated, without the spans.
        """
        row, col = self._cell_args("get_cell_phonetic_runs", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        count = _read_count(
            LIB.fm_sheet_get_cell_phonetic_run_count,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(row, "row", role="coordinate"),
            _uint(col, "col", role="coordinate"),
        )
        runs: List[PhoneticRun] = []
        with _ArgScope() as scope:
            run_ptr = scope.struct(S.PHONETIC_RUN)
            for run_index in range(count):
                S.zero_struct(LIB, S.PHONETIC_RUN, run_ptr)
                _check(
                    LIB.fm_sheet_get_cell_phonetic_run_at(
                        h,
                        _uint(sheet, "sheet_index", role="sheet"),
                        _uint(row, "row", role="coordinate"),
                        _uint(col, "col", role="coordinate"),
                        _uint(run_index, "run_index", role="position"),
                        run_ptr,
                    )
                )
                fields = S.PHONETIC_RUN.unpack(LIB, run_ptr)
                runs.append(
                    PhoneticRun(
                        start=fields["start"],
                        end=fields["end"],
                        text=LIB.read_cstr(fields["text"]) if fields["text"] else "",
                    )
                )
        return runs

    @overload
    @_coordinate_form
    def set_cell_phonetic_properties(self, row: int, col: int, properties: Optional[PhoneticProperties]) -> None: ...

    @overload
    def set_cell_phonetic_properties(self, a1: str, properties: Optional[PhoneticProperties]) -> None: ...

    def set_cell_phonetic_properties(self, *args: Any, **kwargs: Any) -> None:
        """Set how the cell's phonetic guide renders; ``None`` resets it to the unstated default.

        ``row`` and ``col`` are 0-based. ``properties.stated`` is ignored, so
        the :class:`PhoneticProperties` :meth:`get_cell_phonetic_properties`
        returns can be passed back. Independent of
        :meth:`set_cell_phonetic_runs` in both directions: editing the readings
        does not reset the rendering, and this does not touch the readings.
        Only observable on a cell that has runs, and every value-setting
        method clears both, so call it after the cell's text.

        Raises:
          TypeError: when ``properties`` is neither a :class:`PhoneticProperties` nor ``None``.
        """
        row, col, properties = self._cell_args("set_cell_phonetic_properties", args, kwargs, ("properties",))
        if properties is not None and not isinstance(properties, PhoneticProperties):
            raise TypeError(_argument_message("properties", "must be a PhoneticProperties or null"))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = 0
            if properties is not None:
                ptr = scope.struct(S.PHONETIC_PROPERTIES)
                S.PHONETIC_PROPERTIES.pack(
                    LIB,
                    ptr,
                    {
                        "font_index": _uint(properties.font_index, "font_index", role="position"),
                        "type": _uint(properties.type, "type"),
                        "alignment": _uint(properties.alignment, "alignment"),
                    },
                )
            _check(
                LIB.fm_sheet_set_cell_phonetic_properties(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    ptr,
                )
            )

    @overload
    @_coordinate_form
    def get_cell_phonetic_properties(self, row: int, col: int) -> PhoneticProperties: ...

    @overload
    def get_cell_phonetic_properties(self, a1: str) -> PhoneticProperties: ...

    def get_cell_phonetic_properties(self, *args: Any, **kwargs: Any) -> PhoneticProperties:
        """Return how the cell's phonetic guide renders.

        ``row`` and ``col`` are 0-based. A cell with no
        annotation reports the all-zero default. ``stated`` is true when the
        properties are explicit: a ``<phoneticPr>`` element or
        :meth:`set_cell_phonetic_properties`, an all-zero triple included.
        """
        row, col = self._cell_args("get_cell_phonetic_properties", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PHONETIC_PROPERTIES)
            S.zero_struct(LIB, S.PHONETIC_PROPERTIES, ptr)
            _check(
                LIB.fm_sheet_get_cell_phonetic_properties(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    ptr,
                )
            )
            d = S.PHONETIC_PROPERTIES.unpack(LIB, ptr)
            return PhoneticProperties(
                font_index=d["font_index"], type=d["type"], alignment=d["alignment"], stated=bool(d["stated"])
            )

    @overload
    @_coordinate_form
    def get_cell_show_phonetic(self, row: int, col: int) -> bool: ...

    @overload
    def get_cell_show_phonetic(self, a1: str) -> bool: ...

    def get_cell_show_phonetic(self, *args: Any, **kwargs: Any) -> bool:
        """Whether the cell shows its phonetic guide (``<c ph="1">``)."""
        row, col = self._cell_args("get_cell_show_phonetic", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        return bool(
            _read_count(
                LIB.fm_sheet_get_cell_show_phonetic,
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
            )
        )

    @overload
    @_coordinate_form
    def set_cell_show_phonetic(self, row: int, col: int, show: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    @overload
    def set_cell_show_phonetic(self, a1: str, show: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    def set_cell_show_phonetic(self, *args: Any, **kwargs: Any) -> None:
        """Set the cell's "Show Phonetic" display flag.

        Independent of the readings: value writes clear the readings but keep
        this flag. A formula cell's flag is neither saved nor read from a file.
        """
        row, col, show = self._cell_args("set_cell_show_phonetic", args, kwargs, ("show",))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_cell_show_phonetic(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
                int(_require_bool(show, "show")),
            )
        )

    @overload
    @_coordinate_form
    def get_cell_phonetic(self, row: int, col: int) -> Optional[str]: ...

    @overload
    def get_cell_phonetic(self, a1: str) -> Optional[str]: ...

    def get_cell_phonetic(self, *args: Any, **kwargs: Any) -> Optional[str]:
        """Return the cell's phonetic guide, or ``None`` when it has none.

        ``row`` and ``col`` are 0-based.
        """
        row, col = self._cell_args("get_cell_phonetic", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        return _read_optional_cstr_out(
            LIB.fm_sheet_get_cell_phonetic,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(row, "row", role="coordinate"),
            _uint(col, "col", role="coordinate"),
        )

    # -- Cell read ---------------------------------------------------------
    @overload
    @_coordinate_form
    def get_value(self, row: int, col: int) -> NativeValue: ...

    @overload
    def get_value(self, a1: str) -> NativeValue: ...

    def get_value(self, *args: Any, **kwargs: Any) -> NativeValue:
        """Read the value at ``(row, col)`` as native data (:data:`NativeValue`).

        ``None`` for an empty cell, ``float`` for a number (Excel numbers are
        doubles, so ``42`` reads ``42.0``), ``bool``, ``str``, or an
        :class:`ExcelError`. A read computes dirty cells first under
        ``CalcMode.AUTO`` and returns the values as of the last recalc under
        ``CalcMode.MANUAL``. A sheet-qualified address is refused; read
        ``wb.sheet("Sheet1").get_value("A1")`` instead.

        Example::

            ws.set_value("A1", 6)
            ws.set_value("B1", "=A1*7")
            ws.get_value("B1")  # 42.0
        """
        row, col = self._cell_args("get_value", args, kwargs, ())
        return to_native(self._value_at(row, col))

    @overload
    @_coordinate_form
    def get_typed_value(self, row: int, col: int) -> Value: ...

    @overload
    def get_typed_value(self, a1: str) -> Value: ...

    def get_typed_value(self, *args: Any, **kwargs: Any) -> Value:
        """Read the value at ``(row, col)`` tagged with its kind.

        Example::

            v = ws.get_typed_value("A1")
            if v.kind == ValueKind.NUMBER:
                print(v.number)
        """
        row, col = self._cell_args("get_typed_value", args, kwargs, ())
        return self._value_at(row, col)

    def _value_at(self, row: int, col: int) -> Value:
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            value_ptr = scope.alloc(fm_value_t_size)
            status = LIB.fm_sheet_get_value(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
                value_ptr,
            )
            _check(status)
            return Value._from_wasm(value_ptr)

    def _range_or_used(self, range: Union[CellRect, str, None]) -> Optional[CellRect]:
        """``range`` as a rectangle, or the used range when it is ``None`` (``None`` for a sheet with no cells)."""
        return self.get_used_range() if range is None else self._rect(range)

    def get_values(self, range: Union[CellRect, str, None] = None) -> List[List[NativeValue]]:
        """Read the values of ``range`` as a dense row-major grid of native data.

        ``range`` defaults to :meth:`get_used_range`; a sheet with no cells
        reads ``[]``. ``result[r][c]`` is the cell ``r`` rows and ``c`` columns
        from the top-left corner; an empty cell is ``None`` and a spilled cell
        holds its spilled value. ``sheet["A1:B2"]`` reads the same grid.

        Raises:
          InvalidArgumentError: when the range holds more than 4,194,304
            (2**22) cells.

        Example::

            grid = ws.get_values("A1:B2")  # grid[row][col]
        """
        return [[to_native(value) for value in line] for line in self.get_typed_values(range)]

    def get_typed_values(self, range: Union[CellRect, str, None] = None) -> List[List[Value]]:
        """:meth:`get_values` with each value tagged with its kind; :meth:`set_values` stores the grid back unchanged.

        Raises:
          InvalidArgumentError: when the range holds more than 4,194,304
            (2**22) cells.
        """
        rect = self._range_or_used(range)
        if rect is None:
            return []
        rows = rect.last_row - rect.first_row + 1
        cols = rect.last_col - rect.first_col + 1
        if rows <= 0 or cols <= 0 or rows * cols > _GET_VALUES_CELL_CAP:
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _argument_message("range", f"must hold at most {_GET_VALUES_CELL_CAP} cells, got {rows}x{cols}"),
                "",
            )
        grid = [[_BLANK] * cols for _ in builtins.range(rows)]  # `range` is the parameter here
        for row, col, _, value in self._typed_cells(rect):
            grid[row - rect.first_row][col - rect.first_col] = value
        return grid

    @overload
    @_coordinate_form
    def get_date(self, row: int, col: int) -> Optional[datetime.datetime]: ...

    @overload
    def get_date(self, a1: str) -> Optional[datetime.datetime]: ...

    def get_date(self, *args: Any, **kwargs: Any) -> Optional[datetime.datetime]:
        """Read the cell as a date: the naive ``datetime`` its serial names under the workbook's date system.

        The fraction of a day is rounded to the microsecond; ``None`` for an
        empty cell.

        Raises:
          InvalidArgumentError: for a cell holding no number, or a serial
            outside the date system (below 1 in the 1900 system, below 0 in
            the 1904 one, or past 9999-12-31).
        """
        row, col = self._cell_args("get_date", args, kwargs, ())
        value = self._value_at(row, col)
        if value.kind is ValueKind.BLANK:
            return None
        if value.kind is not ValueKind.NUMBER:
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                f"the cell holds {value.kind.name.lower()}, not a date serial",
                f"row={row} col={col}",
            )
        return serial_datetime(value.number, date1904=self._workbook.get_date1904())

    def get_used_range(self) -> Optional[CellRect]:
        """The bounding rectangle of the cells holding a value or a formula; ``None`` for a sheet with none.

        Formatting alone does not count.
        """
        first_row = first_col = last_row = last_col = -1
        for row, col, _, _ in self._typed_cells(_WHOLE_SHEET):
            if first_row < 0:
                first_row, first_col, last_row, last_col = row, col, row, col
                continue
            first_col = min(first_col, col)
            last_col = max(last_col, col)
            last_row = row
        if first_row < 0:
            return None
        return CellRect(first_row=first_row, first_col=first_col, last_row=last_row, last_col=last_col)

    def iter_rows(self, range: Union[CellRect, str, None] = None) -> Iterator[List[NativeValue]]:
        """Iterate the rows of ``range`` (default: :meth:`get_used_range`) as lists of native values, top to bottom.

        Each row is read when the iteration reaches it.

        Example::

            for row in ws.iter_rows():
                print(row)
        """
        rect = self._range_or_used(range)
        if rect is None:
            return iter(())
        return self._iter_rows(rect)

    def _iter_rows(self, rect: CellRect) -> Iterator[List[NativeValue]]:
        width = rect.last_col - rect.first_col + 1
        for row in builtins.range(rect.first_row, rect.last_row + 1):
            line: List[NativeValue] = [None] * width
            band = CellRect(first_row=row, first_col=rect.first_col, last_row=row, last_col=rect.last_col)
            for _, col, _, value in self._typed_cells(band):
                line[col - rect.first_col] = to_native(value)
            yield line

    @overload
    @_coordinate_form
    def set_values(self, row: int, col: int, values: Sequence[Sequence[CellValueInput]]) -> None: ...

    @overload
    def set_values(self, a1: str, values: Sequence[Sequence[CellValueInput]]) -> None: ...

    def set_values(self, *args: Any, **kwargs: Any) -> None:
        """Store a row-major grid of values with ``(row, col)`` as its top-left cell.

        Each element is stored as :meth:`set_value` stores it, so a string
        starting with ``=`` is a formula. Element types, finiteness, equal
        row lengths and the grid's fit on the sheet are all checked before
        any cell is written; an empty grid writes nothing. The write is not
        atomic: when the engine refuses a cell (one inside an array formula
        or a data table, say) or a formula does not parse, the cells before it
        stay written and the error's ``context`` names the cell (``row=`` /
        ``col=``). :meth:`get_typed_values` reads a block back so that it
        stores unchanged.

        Raises:
          TypeError: for a non-grid argument or an element :meth:`set_value`
            rejects by type.
          ValueError: for an integer too large to convert to a float.
          InvalidArgumentError: for rows of unequal length, a NaN or infinite
            number, or a grid reaching past the last row or column.
          FormulaSyntaxError: for a formula that does not parse.

        Example::

            ws.set_values("A1", [["Item", "Qty"], ["Apple", 3]])
        """
        row, col, values = self._cell_args("set_values", args, kwargs, ("values",))
        if isinstance(values, (str, bytes)):
            raise TypeError(_argument_message("values", "must be an array of rows"))
        grid = []
        for line in values:
            if isinstance(line, (str, bytes)):
                raise TypeError(_argument_message("values", "must be an array of rows of cell values"))
            grid.append([_cell_input(value) for value in line])
        width = len(grid[0]) if grid else 0
        if any(len(line) != width for line in grid):
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT, _argument_message("values", "must have rows of one length"), ""
            )
        if width == 0:
            return
        if row + len(grid) - 1 > _LAST_ROW or col + width - 1 > _LAST_COL:
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _argument_message("values", f"{len(grid)}x{width} at ({row}, {col}) is outside the grid"),
                "",
            )
        for r, line in enumerate(grid):
            for c, (setter, payload) in enumerate(line):
                try:
                    getattr(self, setter)(row + r, col + c, *payload)
                except FormulonError as exc:
                    raise _with_cell_context(exc, row + r, col + c) from None

    # -- Ad-hoc array evaluation ------------------------------------------
    @overload
    @_coordinate_form
    def evaluate_cf_formula(self, row: int, col: int, anchor_row: int, anchor_col: int, formula: str) -> bool: ...

    @overload
    def evaluate_cf_formula(self, a1: str, anchor: str, formula: str) -> bool: ...

    def evaluate_cf_formula(self, *args: Any, **kwargs: Any) -> bool:
        """Evaluate a conditional-format formula at one cell without mutation.

        The result is coerced with Excel's conditional-format rules: an error,
        blank, text or numeric-zero result is ``False`` (the rule does not
        fire), any non-zero number is ``True``. An evaluation error is not
        raised.
        """
        row, col, anchor_row, anchor_col, formula = self._cell_args("evaluate_cf_formula", args, kwargs, ("formula",))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            formula_ptr = scope.cstr(formula)
            value_ptr = scope.alloc(fm_value_t_size)
            _check(
                LIB.fm_sheet_evaluate_cf_formula(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    _uint(anchor_row, "anchor_row", role="coordinate"),
                    _uint(anchor_col, "anchor_col", role="coordinate"),
                    formula_ptr,
                    value_ptr,
                )
            )
            return Value._from_wasm(value_ptr).boolean is True

    @overload
    @_coordinate_form
    def evaluate_formula(self, row: int, col: int, formula: str) -> NativeValue: ...

    @overload
    def evaluate_formula(self, a1: str, formula: str) -> NativeValue: ...

    def evaluate_formula(self, *args: Any, **kwargs: Any) -> NativeValue:
        """Evaluate ``formula`` as if entered at ``(row, col)``, without mutating the workbook.

        A 1x1 result is returned as its native value, any other as a list of
        rows. A cell-level Excel error is an :class:`ExcelError`. A reference
        to the anchor cell itself reads its value. Every precedent the
        formula reads is computed, even under ``CalcMode.MANUAL``, without
        storing the results; :meth:`get_value` follows the calc mode. A blank
        cell the formula reads evaluates as 0, as it does in a cell (``=B9``
        on an empty B9 is 0); :meth:`get_value` reads the blank as ``None``.

        Raises:
          TypeError: when ``formula`` is not a ``str``.
          FormulaSyntaxError: when ``formula`` does not parse.

        Example::

            ws.evaluate_formula("A1", "=SUM(B1:B3)")
        """
        row, col, formula = self._cell_args("evaluate_formula", args, kwargs, ("formula",))
        return to_native(self._workbook._evaluate_typed_formula(self._index(), row, col, formula))

    @overload
    @_coordinate_form
    def evaluate_typed_formula(self, row: int, col: int, formula: str) -> Value: ...

    @overload
    def evaluate_typed_formula(self, a1: str, formula: str) -> Value: ...

    def evaluate_typed_formula(self, *args: Any, **kwargs: Any) -> Value:
        """:meth:`evaluate_formula` with the result tagged; a result larger than 1x1 is a ``ValueKind.ARRAY`` value."""
        row, col, formula = self._cell_args("evaluate_typed_formula", args, kwargs, ("formula",))
        return self._workbook._evaluate_typed_formula(self._index(), row, col, formula)

    @overload
    @_coordinate_form
    def evaluate_formula_array(self, row: int, col: int, formula: str) -> List[List[NativeValue]]: ...

    @overload
    def evaluate_formula_array(self, a1: str, formula: str) -> List[List[NativeValue]]: ...

    def evaluate_formula_array(self, *args: Any, **kwargs: Any) -> List[List[NativeValue]]:
        """Evaluate ``formula`` as if entered at ``(row, col)`` and return the whole result as a grid of native values.

        A scalar result such as ``=1+2`` is a 1x1 grid (``[[3.0]]``); a
        dynamic-array formula such as ``=SEQUENCE(2,3)`` yields its
        ``rows x cols`` grid in row-major order (``result[r][c]``). Nothing is
        written and no spill is committed. A blank cell the formula reads
        evaluates as 0, as it does in a cell; :meth:`get_value` reads the
        blank as ``None``.

        Raises:
          TypeError: when ``formula`` is not a ``str``.
          FormulaSyntaxError: when ``formula`` does not parse.
        """
        row, col, formula = self._cell_args("evaluate_formula_array", args, kwargs, ("formula",))
        grid = self._evaluate_grid(row, col, formula)
        return [[to_native(value) for value in line] for line in grid]

    @overload
    @_coordinate_form
    def evaluate_typed_formula_array(self, row: int, col: int, formula: str) -> List[List[Value]]: ...

    @overload
    def evaluate_typed_formula_array(self, a1: str, formula: str) -> List[List[Value]]: ...

    def evaluate_typed_formula_array(self, *args: Any, **kwargs: Any) -> List[List[Value]]:
        """:meth:`evaluate_formula_array` with each value tagged with its kind."""
        row, col, formula = self._cell_args("evaluate_typed_formula_array", args, kwargs, ("formula",))
        return self._evaluate_grid(row, col, formula)

    def _evaluate_grid(self, row: int, col: int, formula: Any) -> List[List[Value]]:
        if not isinstance(formula, str):
            raise TypeError(_argument_message("formula", "must be a string"))
        return self._workbook._evaluate_formula_cells(self._index(), row, col, formula, top_left_only=False)

    # -- Iteration ---------------------------------------------------------
    def list_cells(self, range: Union[CellRect, str, None] = None) -> List[RangeCell]:
        """Return every populated cell in ``range``, or in the used range when it is ``None``.

        Cells come in row-major order; each carries its stored formula, its
        value as native data and the value's kind, as
        :meth:`get_cells_in_range` pages them.
        """
        rect = self._range_or_used(range)
        if rect is None:
            return []
        return [_range_cell(*cell) for cell in self._typed_cells(rect)]

    def _typed_cells(self, rect: CellRect) -> Iterator[Tuple[int, int, Optional[str], Value]]:
        """``(row, col, formula, value)`` of every populated cell of ``rect``, row-major, a page at a time."""
        cursor: Optional[int] = None
        while True:
            cells, cursor = self._cell_page(rect, cursor, None)
            yield from cells
            if cursor is None:
                return

    @overload
    @_span_form
    def insert_rows(self, first: int, last: int) -> None: ...

    @overload
    def insert_rows(self, span: str) -> None: ...

    def insert_rows(self, *args: Any, **kwargs: Any) -> None:
        """Insert the rows ``first`` to ``last`` (inclusive) on the sheet, or the A1 span ``"3:5"``.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          InvalidArgumentError: when ``first`` is past ``last`` or ``last`` is past the sheet bound.
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them, a table, or a value or formula cell past the sheet bound
            (a format-only cell, note or hyperlink pushed off is dropped); the
            workbook is unchanged.
        """
        first, last = self._span_args("insert_rows", args, kwargs, ("first", "last"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_insert_rows(
                h,
                _uint(sheet, "sheet", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
            )
        )

    @overload
    @_span_form
    def delete_rows(self, first: int, last: int) -> None: ...

    @overload
    def delete_rows(self, span: str) -> None: ...

    def delete_rows(self, *args: Any, **kwargs: Any) -> None:
        """Delete the rows ``first`` to ``last`` (inclusive) on the sheet, or the A1 span ``"3:5"``.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          InvalidArgumentError: when ``first`` is past ``last`` or ``last`` is past the sheet bound.
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them or a table past the sheet bound; the workbook is unchanged.
        """
        first, last = self._span_args("delete_rows", args, kwargs, ("first", "last"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_delete_rows(
                h,
                _uint(sheet, "sheet", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
            )
        )

    @overload
    @_span_form
    def insert_columns(self, first: int, last: int) -> None: ...

    @overload
    def insert_columns(self, span: str) -> None: ...

    def insert_columns(self, *args: Any, **kwargs: Any) -> None:
        """Insert the columns ``first`` to ``last`` (inclusive) on the sheet, or the A1 span ``"C:D"``.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          InvalidArgumentError: when ``first`` is past ``last`` or ``last`` is past the sheet bound.
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them, a table, or a value or formula cell past the sheet bound
            (a format-only cell, note or hyperlink pushed off is dropped); the
            workbook is unchanged.
        """
        first, last = self._span_args("insert_columns", args, kwargs, ("first", "last"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_insert_columns(
                h,
                _uint(sheet, "sheet", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
            )
        )

    @overload
    @_span_form
    def delete_columns(self, first: int, last: int) -> None: ...

    @overload
    def delete_columns(self, span: str) -> None: ...

    def delete_columns(self, *args: Any, **kwargs: Any) -> None:
        """Delete the columns ``first`` to ``last`` (inclusive) on the sheet, or the A1 span ``"C:D"``.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          InvalidArgumentError: when ``first`` is past ``last`` or ``last`` is past the sheet bound.
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them or a table past the sheet bound; the workbook is unchanged.
        """
        first, last = self._span_args("delete_columns", args, kwargs, ("first", "last"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_delete_columns(
                h,
                _uint(sheet, "sheet", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
            )
        )

    def insert_cells(self, range: Union[CellRect, str], shift: CellShift) -> None:
        """Insert blank cells over ``range``, moving the cells of the band ``DOWN`` or ``RIGHT``.

        The band is the columns of ``range`` for ``DOWN`` and its rows for ``RIGHT``; only the band moves,
        from the range's leading edge to the grid's end, by the height or width of ``range``. A range
        covering the whole width or height of the sheet acts as :meth:`insert_rows` /
        :meth:`insert_columns`.

        A cell, range, defined-name target or hyperlink range is rewritten only when its perpendicular
        extent lies inside the band: it moves, or grows when it straddles the leading edge. One straddling
        the band, a whole column or a whole row keeps its address. A conditional-format or validation range
        straddling the moving area splits into its moving and fixed parts; a merge straddling it is
        unmerged. Notes and threaded comments move with their cell. Row and column layout, manual breaks,
        drawing anchors and the sheet view do not move. An inserted cell takes the format of the cell
        directly above (``DOWN``) or to its left (``RIGHT``), if that cell has one.

        Raises:
          InvalidArgumentError: for an inverted or off-grid ``range``, a ``shift`` of ``UP`` / ``LEFT``, or a
            ``shift`` along the full extent of the range (a whole-column range with ``DOWN``, a
            whole-row range with ``RIGHT``).
          FormulonError: with ``StatusCode.PRECONDITION_FAILED``, leaving the workbook unchanged, when
            the moving area meets a table, multi-cell array formula or data table whose perpendicular
            extent the band does not cover; meets a PivotTable, unless the band covers its whole
            perpendicular extent and ``range`` does not lie inside it; the insert would push a value or
            formula cell past the grid (a format-only cell, note or hyperlink pushed off is dropped); or
            Excel would accept the edit without changing anything (an insert inside a PivotTable).
        """
        self._shift_cells(LIB.fm_sheet_insert_cells, range, shift)

    def delete_cells(self, range: Union[CellRect, str], shift: CellShift) -> None:
        """Delete the cells of ``range``, moving the cells beyond the band ``UP`` or ``LEFT``.

        The band and the moving area are as in :meth:`insert_cells`, the deleted rectangle included.
        A reference to a wholly deleted cell becomes ``#REF!`` (a defined name ``=Sheet!#REF!``); a partly
        deleted one shrinks. A merge partly intersecting the deleted rectangle is unmerged and its deleted
        cells lose their values. Notes and threaded comments of a deleted cell are removed; a table, data
        table or PivotTable the deleted rectangle contains is removed.

        Raises as :meth:`insert_cells`, with the direction rule reversed (``DOWN`` and ``RIGHT`` are
        refused with ``InvalidArgumentError``) and no push-off refusal; a delete over part of a table or
        inside a PivotTable is refused with ``StatusCode.PRECONDITION_FAILED``.
        """
        self._shift_cells(LIB.fm_sheet_delete_cells, range, shift)

    def _shift_cells(self, fn, range: Union[CellRect, str], shift: CellShift) -> None:
        rect = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(ptr, rect)
            _check(fn(h, _uint(sheet, "sheet", role="sheet"), ptr, _sint(shift, "shift")))

    # -- Partial recalc ----------------------------------------------------
    #
    # Note: ``fm_workbook_set_iterative_progress`` (the per-sweep callback)
    # is intentionally NOT bound -- it takes a C function pointer that the
    # wasmtime-py host cannot synthesise into the module's function table.
    def recalc_viewport(self, viewport: Union[CellRect, str]) -> int:
        """Recalc the dependency closure for the viewport ``viewport``; return cell count."""
        rect = self._rect(viewport)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            vp = scope.struct(S.CELL_RECT)
            out = scope.out_ptr()
            _pack_cell_rect_at(vp, rect)
            _check(LIB.fm_sheet_recalc_viewport(h, _uint(sheet, "sheet", role="sheet"), vp, out))
            return LIB.read_u32(out)

    # -- Lambda text -------------------------------------------------------
    @overload
    @_coordinate_form
    def get_lambda_text(self, row: int, col: int) -> Optional[str]: ...

    @overload
    def get_lambda_text(self, a1: str) -> Optional[str]: ...

    def get_lambda_text(self, *args: Any, **kwargs: Any) -> Optional[str]:
        """Render the lambda cached at ``(row, col)`` as text; ``None`` when the cell holds no lambda."""
        row, col = self._cell_args("get_lambda_text", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        return _read_optional_cstr_out(
            LIB.fm_sheet_get_lambda_text,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(row, "row", role="coordinate"),
            _uint(col, "col", role="coordinate"),
        )

    # -- Merges ------------------------------------------------------------
    def add_merge(self, range: Union[CellRect, str]) -> CellRect:
        """Append a merge range to the sheet and return it as stored.

        Covered cells keep their values and stay writable; clear them first to
        match Excel's Merge.

        Raises:
          AlreadyExistsError: when ``range`` intersects a merge already on the sheet.
          PreconditionFailedError: when ``range`` intersects a table, a pivot table, an array formula or a
            data table.
        """
        merge = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_cell_rect_array([merge], scope)
            out = scope.struct(S.CELL_RECT)
            _check(LIB.fm_sheet_add_merge(h, _uint(sheet, "sheet", role="sheet"), ptr, out))
            return _decode_cell_rect_at(out)

    def remove_merge(self, range: Union[CellRect, str]) -> None:
        """Remove every merge on the sheet overlapping ``range``."""
        merge = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_cell_rect_array([merge], scope)
            _check(LIB.fm_sheet_remove_merge(h, _uint(sheet, "sheet", role="sheet"), ptr))

    def clear_merges(self) -> None:
        """Drop every merge range on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_merges(h, _uint(sheet, "sheet", role="sheet")))

    def list_merges(self) -> List[CellRect]:
        """Return every merge range on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(LIB.fm_sheet_get_merge_count, h, _uint(sheet, "sheet", role="sheet"))
        out: List[CellRect] = []
        # One scratch block for the whole walk; see list_defined_names.
        with _ArgScope() as scope:
            ptr = scope.struct(S.CELL_RECT)
            for i in range(n):
                S.zero_struct(LIB, S.CELL_RECT, ptr)
                _check(
                    LIB.fm_sheet_get_merge_at(
                        h, _uint(sheet, "sheet", role="sheet"), _uint(i, "index", role="position"), ptr
                    )
                )
                out.append(_decode_cell_rect_at(ptr))
        return out

    # -- Hyperlinks --------------------------------------------------------
    @overload
    @_coordinate_form
    def set_hyperlink(self, row: int, col: int, link: Union[HyperlinkInput, Hyperlink, None]) -> None: ...

    @overload
    def set_hyperlink(self, a1: str, link: Union[HyperlinkInput, Hyperlink, None]) -> None: ...

    def set_hyperlink(self, *args: Any, **kwargs: Any) -> None:
        """Set the hyperlink anchored at the cell ``(row, col)``, replacing the one there; ``None`` removes it.

        ``link.target`` is the external address (URL, file path); ``None``
        makes a link inside the workbook, which then needs ``link.location``
        (a reference such as ``"Sheet2!A1"`` or a defined name). A
        ``location`` beside a ``target`` is the fragment within it.
        ``display`` is the hyperlink's stored display attribute; the cell
        value is not written, so call :meth:`set_text` with the same text to
        show it. ``link.range`` is the extent the link covers and must start
        at ``(row, col)``; ``None`` covers the key cell alone, and a
        :class:`Hyperlink` read back keeps its extent. Removing a cell that
        has no hyperlink does nothing.

        Raises:
          TypeError: when ``link`` is neither a :class:`HyperlinkInput`, a :class:`Hyperlink` nor ``None``,
            or its ``target`` and ``location`` are both ``None``.
          InvalidArgumentError: when ``link.range`` does not start at ``(row, col)``.
        """
        row, col, link = self._cell_args("set_hyperlink", args, kwargs, ("link",))
        self._set_hyperlink(row, col, link)

    def _set_hyperlink(self, row: int, col: int, link: Union[HyperlinkInput, Hyperlink, None]) -> None:
        rect = CellRect(row, col, row, col)
        if link is not None:
            if not isinstance(link, (HyperlinkInput, Hyperlink)):
                raise TypeError(_argument_message("link", "must be a HyperlinkInput, a Hyperlink or null"))
            if link.target is None and link.location is None:
                raise TypeError(_argument_message("link", "needs a target or a location"))
            if link.range is not None:
                rect = self._rect(link.range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = 0
            if link is not None:
                ptr = scope.struct(S.HYPERLINK)
                S.HYPERLINK.pack(LIB, ptr, {"rect": _cell_rect_fields(rect)})
                S.write_str_field(LIB, ptr, S.HYPERLINK, "target", link.target, scope)
                S.write_str_field(LIB, ptr, S.HYPERLINK, "location", link.location, scope)
                S.write_str_field(LIB, ptr, S.HYPERLINK, "display", link.display, scope)
                S.write_str_field(LIB, ptr, S.HYPERLINK, "tooltip", link.tooltip, scope)
            _check(
                LIB.fm_sheet_set_hyperlink(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    ptr,
                )
            )

    def clear_hyperlinks(self) -> None:
        """Drop every hyperlink on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_hyperlinks(h, _uint(sheet, "sheet", role="sheet")))

    def list_hyperlinks(self) -> List[Hyperlink]:
        """Return every hyperlink on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(LIB.fm_sheet_get_hyperlink_count, h, _uint(sheet, "sheet", role="sheet"))
        out: List[Hyperlink] = []
        # One scratch block for the whole walk; see list_defined_names.
        with _ArgScope() as scope:
            ptr = scope.struct(S.HYPERLINK)
            for i in range(n):
                S.zero_struct(LIB, S.HYPERLINK, ptr)
                _check(
                    LIB.fm_sheet_get_hyperlink_at(
                        h, _uint(sheet, "sheet", role="sheet"), _uint(i, "index", role="position"), ptr
                    )
                )
                d = S.HYPERLINK.unpack(LIB, ptr)
                out.append(_hyperlink_from(d))
        return out

    @overload
    @_coordinate_form
    def get_hyperlink(self, row: int, col: int) -> Optional[Hyperlink]: ...

    @overload
    def get_hyperlink(self, a1: str) -> Optional[Hyperlink]: ...

    def get_hyperlink(self, *args: Any, **kwargs: Any) -> Optional[Hyperlink]:
        """Return the hyperlink whose range starts at ``(row, col)``, or ``None``."""
        row, col = self._cell_args("get_hyperlink", args, kwargs, ())
        sheet = self._index()
        return self._workbook._lookup(
            LIB.fm_sheet_get_hyperlink,
            (
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
            ),
            S.HYPERLINK,
            _hyperlink_from,
        )

    # -- Comments ----------------------------------------------------------
    @overload
    @_coordinate_form
    def get_comment(self, row: int, col: int) -> Optional[Comment]: ...

    @overload
    def get_comment(self, a1: str) -> Optional[Comment]: ...

    def get_comment(self, *args: Any, **kwargs: Any) -> Optional[Comment]:
        """Return the comment at ``(row, col)`` or ``None`` if absent."""
        row, col = self._cell_args("get_comment", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.COMMENT)
            _check(
                LIB.fm_sheet_get_comment(
                    h,
                    _uint(sheet, "sheet", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    present,
                    ptr,
                )
            )
            if not LIB.read_u32(present):
                return None
            d = S.COMMENT.unpack(LIB, ptr)
            return Comment(
                author=LIB.read_cstr(d["author"]),
                text=LIB.read_cstr(d["text"]),
            )

    @overload
    @_coordinate_form
    def set_comment(self, row: int, col: int, comment: Optional[Comment]) -> None: ...

    @overload
    def set_comment(self, a1: str, comment: Optional[Comment]) -> None: ...

    def set_comment(self, *args: Any, **kwargs: Any) -> None:
        """Insert or replace the comment (note) at ``(row, col)``; ``None`` removes it.

        An empty ``comment.text`` writes an empty note, as Excel keeps one.
        Removing a cell that has no note does nothing. The ``*_comment``
        family is Excel 365's Notes (legacy comments); Excel's comments are
        the ``*_threaded_comment`` family.

        Raises:
          AlreadyExistsError: when the cell holds a threaded comment: a cell
            carries a note or a thread, never both.
        """
        row, col, comment = self._cell_args("set_comment", args, kwargs, ("comment",))
        if comment is not None and not isinstance(comment, Comment):
            raise TypeError(_argument_message("comment", "must be a Comment or null"))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = 0
            if comment is not None:
                ptr = scope.struct(S.COMMENT)
                S.COMMENT.pack(LIB, ptr, {})
                S.write_str_field(LIB, ptr, S.COMMENT, "author", comment.author, scope)
                S.write_str_field(LIB, ptr, S.COMMENT, "text", comment.text, scope)
            _check(
                LIB.fm_sheet_set_comment(
                    h,
                    _uint(sheet, "sheet", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    ptr,
                )
            )

    def list_comments(self) -> List[SheetComment]:
        """Return every comment on the sheet in storage order.

        Unlike :meth:`get_comment`, this discovers comments on otherwise
        empty cells without requiring callers to scan worksheet coordinates.
        """
        sheet = self._index()
        h = self._workbook._require()
        out: List[SheetComment] = []
        # One scratch block for the whole walk; see list_defined_names.
        with _ArgScope() as scope:
            ptr = scope.struct(S.COMMENT)
            for index in range(_read_count(LIB.fm_sheet_get_comment_count, h, _uint(sheet, "sheet", role="sheet"))):
                S.zero_struct(LIB, S.COMMENT, ptr)
                _check(
                    LIB.fm_sheet_get_comment_at(
                        h, _uint(sheet, "sheet", role="sheet"), _uint(index, "index", role="position"), ptr
                    )
                )
                d = S.COMMENT.unpack(LIB, ptr)
                out.append(
                    SheetComment(
                        cell=CellAddress(row=d["row"], col=d["col"]),
                        author=LIB.read_cstr(d["author"]),
                        text=LIB.read_cstr(d["text"]),
                    )
                )
        return out

    @overload
    @_coordinate_form
    def get_comment_visible(self, row: int, col: int) -> Optional[bool]: ...

    @overload
    def get_comment_visible(self, a1: str) -> Optional[bool]: ...

    def get_comment_visible(self, *args: Any, **kwargs: Any) -> Optional[bool]:
        """Return whether the note at the cell is shown without hovering, or ``None`` if absent."""
        row, col = self._cell_args("get_comment_visible", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_get_comment_visible(
                    h,
                    _uint(sheet, "sheet", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    present,
                    out,
                )
            )
            if LIB.read_u32(present) == 0:
                return None
            return LIB.read_u32(out) != 0

    @overload
    @_coordinate_form
    def set_comment_visible(self, row: int, col: int, visible: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    @overload
    def set_comment_visible(self, a1: str, visible: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    def set_comment_visible(self, *args: Any, **kwargs: Any) -> None:
        """Show or hide the note at the cell; its box geometry is kept."""
        row, col, visible = self._cell_args("set_comment_visible", args, kwargs, ("visible",))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_comment_visible(
                h,
                _uint(sheet, "sheet", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
                int(_require_bool(visible, "visible")),
            )
        )

    _PROTECTED_RANGE_TEXT = (
        "name",
        "legacy_password_hash",
        "algorithm_name",
        "hash_value",
        "salt_value",
        "security_descriptor",
    )

    def list_protected_ranges(self) -> List[ProtectedRange]:
        """Every protected range of the sheet, in order."""
        return [self._protected_range_at(i) for i in range(self._protected_range_count())]

    def _protected_range_count(self) -> int:
        """Return the number of protected ranges on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(LIB.fm_sheet_get_protected_range_count, h, _uint(sheet, "sheet", role="sheet"))

    def _protected_range_at(self, protected_range_index: int) -> ProtectedRange:
        """Return the ``protected_range_index``-th protected range of the sheet in file order."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PROTECTED_RANGE)
            _check(
                LIB.fm_sheet_get_protected_range_at(
                    h,
                    _uint(sheet, "sheet", role="sheet"),
                    _uint(protected_range_index, "protected_range_index", role="position"),
                    ptr,
                )
            )
            return self._decode_protected_range(ptr)

    def get_protected_range(self, protected_range: ProtectedRangeRef) -> Optional[ProtectedRange]:
        """Return ``protected_range`` (a name or a position in :meth:`list_protected_ranges`), or ``None`` when absent.

        Raises:
          IndexOutOfRangeError: when ``protected_range`` is a position past the last range.
        """
        name = self._protected_range_name(protected_range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.PROTECTED_RANGE)
            _check(
                LIB.fm_sheet_get_protected_range(h, _uint(sheet, "sheet", role="sheet"), scope.cstr(name), present, ptr)
            )
            if not LIB.read_u32(present):
                return None
            return self._decode_protected_range(ptr)

    def _decode_protected_range(self, ptr: int) -> ProtectedRange:
        d = S.PROTECTED_RANGE.unpack(LIB, ptr)
        return ProtectedRange(
            spin_count=d["spin_count"],
            ranges=_decode_cell_rect_array(d["ranges"], d["range_count"]),
            legacy_password_hash=LIB.read_optional_cstr(d["legacy_password_hash"]),
            **{
                fld: LIB.read_optional_cstr(d[fld])
                for fld in self._PROTECTED_RANGE_TEXT
                if fld != "legacy_password_hash"
            },
        )

    def _protected_range_call(
        self,
        fn,
        leading: tuple,
        protected: Union[ProtectedRange, ProtectedRangeInput],
        *,
        keep_ranges: bool = False,
    ) -> str:
        h = self._workbook._require()
        rects = [] if keep_ranges else _cell_rects(self._workbook, protected.ranges, "ranges")
        with _ArgScope() as scope:
            ptr = scope.struct(S.PROTECTED_RANGE)
            # NULL with a zero count keeps the stored rectangles; an empty list needs a non-NULL pointer to stay empty.
            ranges_ptr = 0
            if not keep_ranges:
                ranges_ptr = _pack_cell_rect_array(rects, scope) or scope.struct_array(S.CELL_RECT, 1)
            S.PROTECTED_RANGE.pack(
                LIB,
                ptr,
                {
                    "spin_count": _uint(protected.spin_count, "spin_count"),
                    "ranges": ranges_ptr,
                    "range_count": 0 if keep_ranges else _uint(len(rects), "range_count"),
                },
            )
            for fld in self._PROTECTED_RANGE_TEXT:
                S.write_str_field(LIB, ptr, S.PROTECTED_RANGE, fld, getattr(protected, fld), scope)
            name_out = scope.out_ptr()
            _check(fn(h, *leading, ptr, name_out))
            return LIB.read_cstr(LIB.read_u32(name_out))

    def add_protected_range(self, input: Union[ProtectedRangeInput, ProtectedRange]) -> str:
        """Append a protected range and return its name; ``name`` must be non-empty and unique on the sheet.

        ``input.ranges`` is a list of :class:`CellRect` or unqualified A1 ranges, or one A1 string. A
        :class:`ProtectedRange` read back is accepted in place of a :class:`ProtectedRangeInput`.

        Raises:
          TypeError: when ``input`` is not one of the two record types.
        """
        if not isinstance(input, (ProtectedRangeInput, ProtectedRange)):
            raise TypeError(_argument_message("input", "must be a ProtectedRangeInput"))
        sheet = self._index()
        return self._protected_range_call(
            LIB.fm_sheet_add_protected_range,
            (_uint(sheet, "sheet", role="sheet"),),
            input,
        )

    def _protected_range_name(self, protected_range: ProtectedRangeRef) -> str:
        """The name ``protected_range`` addresses: a name as it is, a position through its range."""
        if isinstance(protected_range, str):
            return protected_range
        if isinstance(protected_range, bool):
            raise TypeError(_argument_message("protected_range", "must be an integer or a string"))
        return self._protected_range_at(_uint(protected_range, "protected_range", role="position")).name

    def update_protected_range(
        self,
        protected_range: ProtectedRangeRef,
        *,
        name: Union[str, Unchanged] = UNCHANGED,
        ranges: Union[Sequence[Union[CellRect, str]], str, Unchanged] = UNCHANGED,
        legacy_password_hash: Union[str, Unchanged] = UNCHANGED,
        algorithm_name: Union[str, Unchanged] = UNCHANGED,
        hash_value: Union[str, Unchanged] = UNCHANGED,
        salt_value: Union[str, Unchanged] = UNCHANGED,
        spin_count: Union[int, Unchanged] = UNCHANGED,
        security_descriptor: Union[str, Unchanged] = UNCHANGED,
    ) -> None:
        """Update ``protected_range``, an index or a protected range's name; an omitted argument keeps its value.

        The arguments are the :class:`ProtectedRange` fields (``ranges`` also
        takes A1 text); an empty string clears a text field. Omitting ``ranges`` keeps the stored rectangles,
        including any the reader could not parse.

        Raises:
          TypeError: for ``None`` on any argument but ``legacy_password_hash``, which ``None`` removes.
          NotFoundError: when no protected range has the name ``protected_range``.
          IndexOutOfRangeError: when ``protected_range`` is an index past the last one.
          InvalidArgumentError: when ``ranges`` is empty.
        """
        given = {
            key: value
            for key, value in (
                ("name", name),
                ("ranges", ranges),
                ("legacy_password_hash", legacy_password_hash),
                ("algorithm_name", algorithm_name),
                ("hash_value", hash_value),
                ("salt_value", salt_value),
                ("spin_count", spin_count),
                ("security_descriptor", security_descriptor),
            )
            if value is not UNCHANGED
        }
        for key, value in given.items():
            if value is None and key != "legacy_password_hash":
                raise TypeError(_argument_message(key, "must not be null"))
        range_name = self._protected_range_name(protected_range)
        sheet = self._index()
        h = self._workbook._require()
        values: Dict[str, object] = {}
        with _ArgScope() as scope:
            ptr = scope.struct(S.PROTECTED_RANGE_PATCH)
            if "spin_count" in given:
                values["spin_count_op"] = _PATCH_SET
                values["spin_count"] = _uint(given["spin_count"], "spin_count")
            if "ranges" in given:
                rects = _cell_rects(self._workbook, given["ranges"], "ranges")
                # An empty list still needs a non-NULL pointer; the engine refuses it by count.
                values["ranges_op"] = _PATCH_SET
                values["ranges"] = _pack_cell_rect_array(rects, scope) or scope.struct_array(S.CELL_RECT, 1)
                values["range_count"] = _uint(len(rects), "range_count")
            for key in self._PROTECTED_RANGE_TEXT:
                if key not in given:
                    continue
                if given[key] is None:
                    values[key + "_op"] = _PATCH_CLEAR
                else:
                    values[key + "_op"] = _PATCH_SET
                    values[key] = scope.cstr(given[key])
            S.PROTECTED_RANGE_PATCH.pack(LIB, ptr, values)
            _check(
                LIB.fm_sheet_update_protected_range(h, _uint(sheet, "sheet", role="sheet"), scope.cstr(range_name), ptr)
            )

    def remove_protected_range(self, protected_range: ProtectedRangeRef) -> None:
        """Remove ``protected_range``, an index or a protected range's name.

        Raises:
          NotFoundError: when no protected range has the name ``protected_range``.
          IndexOutOfRangeError: when ``protected_range`` is an index past the last one.
        """
        range_name = self._protected_range_name(protected_range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            _check(LIB.fm_sheet_remove_protected_range(h, _uint(sheet, "sheet", role="sheet"), scope.cstr(range_name)))

    # -- Data validations --------------------------------------------------
    def list_validations(self) -> List[DataValidation]:
        """Every data validation of the sheet, in order."""
        return [self._validation_at(i) for i in range(self._validation_count())]

    def _validation_count(self) -> int:
        """Return the number of data-validation rules on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(LIB.fm_sheet_get_validation_count, h, _uint(sheet, "sheet", role="sheet"))

    def _validation_at(self, validation_index: int) -> DataValidation:
        """Read the ``validation_index``-th data-validation rule on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.DATA_VALIDATION)
            _check(
                LIB.fm_sheet_get_validation_at(
                    h,
                    _uint(sheet, "sheet", role="sheet"),
                    _uint(validation_index, "validation_index", role="position"),
                    ptr,
                )
            )
            return self._decode_validation(ptr)

    @staticmethod
    def _decode_validation(ptr: int) -> DataValidation:
        d = S.DATA_VALIDATION.unpack(LIB, ptr)
        ranges: List[CellRect] = []
        base = d["ranges"]
        rsize = S.CELL_RECT.size
        for i in range(d["range_count"]):
            ranges.append(_decode_cell_rect_at(base + i * rsize))
        return DataValidation(
            id=LIB.read_cstr(d["id"]),
            ranges=ranges,
            type=d["type"],
            operator=d["op"],
            error_style=d["error_style"],
            allow_blank=d["allow_blank"] == S.TRISTATE_TRUE,
            show_input_message=d["show_input_message"] == S.TRISTATE_TRUE,
            show_error_message=d["show_error_message"] == S.TRISTATE_TRUE,
            show_in_cell_dropdown=d["show_in_cell_dropdown"] == S.TRISTATE_TRUE,
            formula1=LIB.read_cstr(d["formula1"]) or None,
            formula2=LIB.read_cstr(d["formula2"]) or None,
            error_title=LIB.read_cstr(d["error_title"]) or None,
            error_message=LIB.read_cstr(d["error_message"]) or None,
            prompt_title=LIB.read_cstr(d["prompt_title"]) or None,
            prompt_message=LIB.read_cstr(d["prompt_message"]) or None,
        )

    def add_validation(self, validation: Union[DataValidationInput, DataValidation]) -> ValidationId:
        """Append a data-validation rule to the sheet; return its :data:`ValidationId`.

        ``validation.ranges`` takes :class:`CellRect` values, A1 ranges or one A1 range string. A
        :class:`DataValidation` read back is accepted in place of a :class:`DataValidationInput`; its ``id``
        is ignored and the new rule gets a fresh one.
        """
        sheet = self._index()
        validation = normalize_record_ranges(validation, self._workbook, sheet=sheet, label="validation")
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.DATA_VALIDATION)
            S.DATA_VALIDATION.pack(
                LIB,
                ptr,
                {
                    "type": _uint(validation.type, "type", 8),
                    "op": _uint(validation.operator, "operator", 8),
                    "error_style": _uint(validation.error_style, "error_style", 8),
                    "allow_blank": _tristate_default(validation.allow_blank),
                    "show_input_message": _tristate_default(validation.show_input_message),
                    "show_error_message": _tristate_default(validation.show_error_message),
                    "show_in_cell_dropdown": _tristate_default(validation.show_in_cell_dropdown),
                },
            )
            ranges_ptr = _pack_cell_rect_array(validation.ranges, scope)
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
                S.write_str_field(LIB, ptr, S.DATA_VALIDATION, fld, getattr(validation, fld), scope)
            out = scope.out_ptr()
            _check(LIB.fm_sheet_add_validation(h, _uint(sheet, "sheet", role="sheet"), ptr, out))
            return LIB.read_cstr(LIB.read_u32(out))

    def _validation_id(self, validation: ValidationRef) -> str:
        """The :data:`ValidationId` ``validation`` names: an id as it is, a position through its rule."""
        if isinstance(validation, str):
            return validation
        if isinstance(validation, bool):
            raise TypeError(_argument_message("validation", "must be an integer or a string"))
        return self._validation_at(_uint(validation, "validation", role="position")).id

    def get_validation(self, validation: ValidationRef) -> Optional[DataValidation]:
        """Return ``validation`` (an id or a position in :meth:`list_validations`), or ``None`` when absent.

        Raises:
          IndexOutOfRangeError: when ``validation`` is a position past the last rule.
        """
        validation_id = self._validation_id(validation)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.DATA_VALIDATION)
            _check(
                LIB.fm_sheet_get_validation(
                    h, _uint(sheet, "sheet", role="sheet"), scope.cstr(validation_id), present, ptr
                )
            )
            if not LIB.read_u32(present):
                return None
            return self._decode_validation(ptr)

    def remove_validation(self, validation: ValidationRef) -> None:
        """Remove ``validation``, a rule's :data:`ValidationId` or its position in :meth:`list_validations`.

        Raises:
          NotFoundError: when no rule has the id ``validation``.
          IndexOutOfRangeError: when ``validation`` is a position past the last rule.
        """
        validation_id = self._validation_id(validation)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            _check(LIB.fm_sheet_remove_validation(h, _uint(sheet, "sheet", role="sheet"), scope.cstr(validation_id)))

    def update_validation(
        self,
        validation: ValidationRef,
        *,
        ranges: Union[Sequence[Union[CellRect, str]], str, Unchanged] = UNCHANGED,
        type: Union[ValidationType, Unchanged] = UNCHANGED,
        operator: Union[ComparisonOperator, Unchanged] = UNCHANGED,
        error_style: Union[ValidationErrorStyle, Unchanged] = UNCHANGED,
        allow_blank: Union[bool, Unchanged] = UNCHANGED,
        show_input_message: Union[bool, Unchanged] = UNCHANGED,
        show_error_message: Union[bool, Unchanged] = UNCHANGED,
        show_in_cell_dropdown: Union[bool, Unchanged] = UNCHANGED,
        formula1: Union[str, None, Unchanged] = UNCHANGED,
        formula2: Union[str, None, Unchanged] = UNCHANGED,
        error_title: Union[str, None, Unchanged] = UNCHANGED,
        error_message: Union[str, None, Unchanged] = UNCHANGED,
        prompt_title: Union[str, None, Unchanged] = UNCHANGED,
        prompt_message: Union[str, None, Unchanged] = UNCHANGED,
    ) -> None:
        """Update ``validation``, a rule's :data:`ValidationId` or position; an omitted argument keeps its value.

        The arguments are the :class:`DataValidationInput` fields. ``None`` clears ``formula1``, ``formula2``
        and the four texts, and is a ``TypeError`` for any other argument. The whole update is checked before
        anything changes.

        Raises:
          TypeError: for ``None`` on an argument that is not clearable.
          NotFoundError: when no rule has the id ``validation``.
          IndexOutOfRangeError: when ``validation`` is a position past the last rule.
          InvalidArgumentError: when ``ranges`` is empty or a value is outside its domain.
        """
        given = {
            key: value
            for key, value in (
                ("ranges", ranges),
                ("type", type),
                ("op", operator),
                ("error_style", error_style),
                ("allow_blank", allow_blank),
                ("show_input_message", show_input_message),
                ("show_error_message", show_error_message),
                ("show_in_cell_dropdown", show_in_cell_dropdown),
                ("formula1", formula1),
                ("formula2", formula2),
                ("error_title", error_title),
                ("error_message", error_message),
                ("prompt_title", prompt_title),
                ("prompt_message", prompt_message),
            )
            if value is not UNCHANGED
        }
        for key, value in given.items():
            if value is None and key not in _VALIDATION_TEXTS:
                raise TypeError(_argument_message("operator" if key == "op" else key, "must not be null"))
        validation_id = self._validation_id(validation)
        sheet = self._index()
        h = self._workbook._require()
        values: Dict[str, object] = {}
        with _ArgScope() as scope:
            ptr = scope.struct(S.DATA_VALIDATION_PATCH)
            if "ranges" in given:
                rects = _cell_rects(self._workbook, given["ranges"], "ranges")
                # An empty list still needs a non-NULL pointer; the engine refuses it by count.
                values["ranges_op"] = _PATCH_SET
                values["ranges"] = _pack_cell_rect_array(rects, scope) or scope.struct_array(S.CELL_RECT, 1)
                values["range_count"] = _uint(len(rects), "range_count")
            for key in ("type", "op", "error_style"):
                if key in given:
                    values[key + "_op"] = _PATCH_SET
                    values[key] = _sint(given[key], key)
            for key in _VALIDATION_FLAGS:
                if key in given:
                    _require_bool(given[key], key)
                    values[key + "_op"] = _PATCH_SET
                    values[key] = _tristate_default(given[key])
            for key in _VALIDATION_TEXTS:
                if key not in given:
                    continue
                if given[key] is None:
                    values[key + "_op"] = _PATCH_CLEAR
                else:
                    values[key + "_op"] = _PATCH_SET
                    values[key] = scope.cstr(given[key])
            S.DATA_VALIDATION_PATCH.pack(LIB, ptr, values)
            _check(
                LIB.fm_sheet_update_validation(h, _uint(sheet, "sheet", role="sheet"), scope.cstr(validation_id), ptr)
            )

    def clear_validations(self) -> None:
        """Drop every validation rule on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_validations(h, _uint(sheet, "sheet", role="sheet")))

    # -- Sheet protection --------------------------------------------------
    # The two select flags are ``fm_tristate_t`` in C (FALSE 1, TRUE 2); the others are 0 / 1.
    _PROTECT_TRISTATE = frozenset(("allow_select_locked_cells", "allow_select_unlocked_cells"))

    def get_sheet_protection(self) -> Optional[SheetProtection]:
        """Read the sheet's ``<sheetProtection>`` record, or ``None`` when the sheet is not protected."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.SHEET_PROTECTION)
            _check(LIB.fm_sheet_get_sheet_protection(h, _uint(sheet, "sheet_index", role="sheet"), present, ptr))
            if not LIB.read_i32(present):
                return None
            d = S.SHEET_PROTECTION.unpack(LIB, ptr)
            return SheetProtection(
                algorithm_name=LIB.read_cstr(d["algorithm_name"]) or None,
                hash_value=LIB.read_cstr(d["hash_value"]) or None,
                salt_value=LIB.read_cstr(d["salt_value"]) or None,
                spin_count=d["spin_count"],
                legacy_password_hash=LIB.read_cstr(d["legacy_password_hash"]) or None,
                **{
                    flag: d[flag] == S.TRISTATE_TRUE if flag in self._PROTECT_TRISTATE else d[flag] != 0
                    for flag in S.SHEET_PROTECTION_ALLOW_FLAGS
                },
            )

    def set_sheet_protection(self, protection: Union[SheetProtectionInput, SheetProtection, None]) -> None:
        """Protect the sheet with ``protection``, replacing any protection it has; ``None`` unprotects it.

        The defaults of :class:`SheetProtectionInput` are those of Excel's
        Protect Sheet dialog. A :class:`SheetProtection` read back is
        accepted as it is. Unprotecting an unprotected sheet does nothing.

        Raises:
          TypeError: when ``protection`` is neither a :class:`SheetProtectionInput`,
            a :class:`SheetProtection` nor ``None``.
        """
        if protection is not None and not isinstance(protection, (SheetProtectionInput, SheetProtection)):
            raise TypeError(
                _argument_message("protection", "must be a SheetProtectionInput, a SheetProtection or null")
            )
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = 0
            if protection is not None:
                ptr = scope.struct(S.SHEET_PROTECTION)
                values: Dict[str, int] = {"spin_count": _uint(protection.spin_count, "spin_count")}
                for flag in S.SHEET_PROTECTION_ALLOW_FLAGS:
                    value = getattr(protection, flag)
                    values[flag] = (
                        _tristate_default(value) if flag in self._PROTECT_TRISTATE else int(_require_bool(value, flag))
                    )
                S.SHEET_PROTECTION.pack(LIB, ptr, values)
                for fld in S.SHEET_PROTECTION_TEXT_FIELDS:
                    S.write_str_field(LIB, ptr, S.SHEET_PROTECTION, fld, getattr(protection, fld), scope)
            _check(LIB.fm_sheet_set_sheet_protection(h, _uint(sheet, "sheet_index", role="sheet"), ptr))

    def update_sheet_protection(
        self,
        *,
        algorithm_name: Union[str, None, Unchanged] = UNCHANGED,
        hash_value: Union[str, None, Unchanged] = UNCHANGED,
        salt_value: Union[str, None, Unchanged] = UNCHANGED,
        spin_count: Union[int, None, Unchanged] = UNCHANGED,
        legacy_password_hash: Union[str, None, Unchanged] = UNCHANGED,
        allow_edit_objects: Union[bool, Unchanged] = UNCHANGED,
        allow_edit_scenarios: Union[bool, Unchanged] = UNCHANGED,
        allow_format_cells: Union[bool, Unchanged] = UNCHANGED,
        allow_format_columns: Union[bool, Unchanged] = UNCHANGED,
        allow_format_rows: Union[bool, Unchanged] = UNCHANGED,
        allow_insert_columns: Union[bool, Unchanged] = UNCHANGED,
        allow_insert_rows: Union[bool, Unchanged] = UNCHANGED,
        allow_insert_hyperlinks: Union[bool, Unchanged] = UNCHANGED,
        allow_delete_columns: Union[bool, Unchanged] = UNCHANGED,
        allow_delete_rows: Union[bool, Unchanged] = UNCHANGED,
        allow_select_locked_cells: Union[bool, Unchanged] = UNCHANGED,
        allow_select_unlocked_cells: Union[bool, Unchanged] = UNCHANGED,
        allow_sort: Union[bool, Unchanged] = UNCHANGED,
        allow_auto_filter: Union[bool, Unchanged] = UNCHANGED,
        allow_pivot_tables: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the sheet's protection with the arguments given; an omitted argument keeps its value.

        ``None`` clears a text field or ``spin_count``; an ``allow_*`` flag has
        no cleared form. On an unprotected sheet the patch applies to the
        :class:`SheetProtectionInput` defaults and protects the sheet, unless
        every argument is omitted. The update is all-or-nothing.

        Raises:
          TypeError: when an ``allow_*`` flag is not a ``bool`` (``None``
            included), a text field is neither a ``str`` nor ``None``, or
            ``spin_count`` is neither an integer nor ``None``.
        """
        arguments = dict(locals())
        given = {
            name: arguments[name]
            for name in (*S.SHEET_PROTECTION_TEXT_FIELDS, "spin_count", *S.SHEET_PROTECTION_ALLOW_FLAGS)
            if arguments[name] is not UNCHANGED
        }
        for name, value in given.items():
            if name in S.SHEET_PROTECTION_ALLOW_FLAGS:
                _require_bool(value, name)
            elif name in S.SHEET_PROTECTION_TEXT_FIELDS:
                if value is not None and not isinstance(value, str):
                    raise TypeError(_argument_message(name, "must be a string or null"))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHEET_PROTECTION_PATCH)
            values: Dict[str, int] = {}
            for name, value in given.items():
                if value is None:
                    values[name + "_op"] = _PATCH_CLEAR
                    continue
                values[name + "_op"] = _PATCH_SET
                if name == "spin_count":
                    values[name] = _uint(value, name)
                elif name in S.SHEET_PROTECTION_ALLOW_FLAGS:
                    values[name] = int(_require_bool(value, name))
                else:
                    values[name] = scope.cstr(value)
            S.SHEET_PROTECTION_PATCH.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_update_sheet_protection(h, _uint(sheet, "sheet_index", role="sheet"), ptr))

    # -- Geometry / display ------------------------------------------------
    @overload
    @_span_form
    def reset_row_height(self, first: int, last: int) -> None: ...

    @overload
    def reset_row_height(self, span: str) -> None: ...

    def reset_row_height(self, *args: Any, **kwargs: Any) -> None:
        """Reset the height on ``[first, last]`` to the sheet default by dropping its override."""
        first, last = self._span_args("reset_row_height", args, kwargs, ("first", "last"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_reset_row_height(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
            )
        )

    @overload
    @_span_form
    def reset_column_width(self, first: int, last: int) -> None: ...

    @overload
    def reset_column_width(self, span: str) -> None: ...

    def reset_column_width(self, *args: Any, **kwargs: Any) -> None:
        """Reset the width on ``[first, last]`` to the sheet default by dropping its override."""
        first, last = self._span_args("reset_column_width", args, kwargs, ("first", "last"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_reset_column_width(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
            )
        )

    def get_format_defaults(self) -> SheetFormatDefaults:
        """Read the sheet's default column / row metrics (``<sheetFormatPr>``)."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHEET_FORMAT_DEFAULTS)
            _check(LIB.fm_sheet_get_format_defaults(h, _uint(sheet, "sheet_index", role="sheet"), ptr))
            d = S.SHEET_FORMAT_DEFAULTS.unpack(LIB, ptr)
            return SheetFormatDefaults(
                default_column_width_chars=(
                    d["default_column_width_chars"] if d["has_default_column_width_chars"] else None
                ),
                default_row_height_pt=d["default_row_height_pt"] if d["has_default_row_height_pt"] else None,
                base_column_width_chars=d["base_column_width_chars"],
                custom_height=bool(d["custom_height"]),
                zero_height=bool(d["zero_height"]),
                effective_default_column_width_chars=d["effective_default_column_width_chars"],
                effective_default_row_height_pt=d["effective_default_row_height_pt"],
            )

    def update_format_defaults(
        self,
        *,
        default_column_width_chars: Union[float, None, Unchanged] = UNCHANGED,
        default_row_height_pt: Union[float, None, Unchanged] = UNCHANGED,
        base_column_width_chars: Union[float, Unchanged] = UNCHANGED,
        custom_height: Union[bool, Unchanged] = UNCHANGED,
        zero_height: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the sheet's default column / row metrics; an omitted argument keeps its value.

        Arguments are the :class:`SheetFormatDefaults` fields.
        ``default_column_width_chars`` / ``default_row_height_pt`` given as
        ``None`` clear the stored value; the other arguments cannot be cleared.

        Raises:
          TypeError: when ``base_column_width_chars`` is ``None`` or a flag is not a ``bool``.
        """
        if base_column_width_chars is None:
            raise TypeError(_argument_message("base_column_width_chars", "must not be null"))
        values: Dict[str, object] = {}
        for name, value in (
            ("default_column_width_chars", default_column_width_chars),
            ("default_row_height_pt", default_row_height_pt),
        ):
            if value is UNCHANGED:
                continue
            values[name + "_op"] = _PATCH_CLEAR if value is None else _PATCH_SET
            values[name] = 0.0 if value is None else _finite(value, name)
        if base_column_width_chars is not UNCHANGED:
            values["base_column_width_chars_op"] = _PATCH_SET
            values["base_column_width_chars"] = _finite(base_column_width_chars, "base_column_width_chars")
        for name, value in (("custom_height", custom_height), ("zero_height", zero_height)):
            if value is not UNCHANGED:
                values[name + "_op"] = _PATCH_SET
                values[name] = _require_bool(value, name)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHEET_FORMAT_DEFAULTS_PATCH)
            S.SHEET_FORMAT_DEFAULTS_PATCH.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_update_format_defaults(h, _uint(sheet, "sheet_index", role="sheet"), ptr))

    def _read_f64_out(self, fn, *args) -> float:
        """Call a ``(..., double* out)`` ABI function and return the double."""
        with _ArgScope() as scope:
            out = scope.alloc(8)
            _check(fn(*args, out))
            return LIB.read_f64(out)

    def get_cell_rect_pt(self, range: Union[CellRect, str], mode: GeometryMode = GeometryMode.DISPLAY) -> RectPt:
        """Return the rectangle covering ``range`` in points from the top-left of A1."""
        cell_range = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, cell_range)
            ptr = scope.struct(S.RECT_PT)
            _check(
                LIB.fm_sheet_get_cell_rect_pt(
                    h, _uint(sheet, "sheet_index", role="sheet"), rp, _sint(mode, "mode"), ptr
                )
            )
            return RectPt(**S.RECT_PT.unpack(LIB, ptr))

    def _single_axis(self, value: Union[int, str], axis: str) -> int:
        """The column (``axis`` ``"col"``) or row (``"row"``) an index or a one-column / one-row A1 span names."""
        if not isinstance(value, str):
            return value
        first, last = self._span(value, axis)
        if first != last:
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _argument_message(axis, f"must name one {'column' if axis == 'col' else 'row'}"),
                f"{axis}={value}",
            )
        return first

    def get_column_width_pt(self, col: Union[int, str], mode: GeometryMode = GeometryMode.DISPLAY) -> float:
        """Return the effective width of ``col`` in points (0 when hidden).

        ``col`` is a zero-based index or one A1 column such as ``"B"``.

        Raises:
          InvalidArgumentError: when ``mode`` is not a :class:`GeometryMode` member or ``col`` spans several columns.
        """
        col = self._single_axis(col, "col")
        mode_c = _geometry_mode(mode)
        sheet = self._index()
        h = self._workbook._require()
        return self._read_f64_out(
            LIB.fm_sheet_get_column_width_pt,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(col, "col", role="coordinate"),
            mode_c,
        )

    def get_row_height_pt(self, row: Union[int, str]) -> float:
        """Return the effective height of ``row`` in points (0 when hidden).

        ``row`` is a zero-based index or one A1 row such as ``"3"``.
        """
        row = self._single_axis(row, "row")
        sheet = self._index()
        h = self._workbook._require()
        return self._read_f64_out(
            LIB.fm_sheet_get_row_height_pt,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(row, "row", role="coordinate"),
        )

    def get_width_model(self, mode: GeometryMode = GeometryMode.DISPLAY) -> WidthModel:
        """Return the character-unit to point conversion model in ``mode``."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.WIDTH_MODEL)
            _check(LIB.fm_sheet_get_width_model(h, _uint(sheet, "sheet_index", role="sheet"), _sint(mode, "mode"), ptr))
            d = S.WIDTH_MODEL.unpack(LIB, ptr)
            return WidthModel(
                points_per_char=d["points_per_char"],
                padding_pt=d["padding_pt"],
                normal_font_size_pt=d["normal_font_size_pt"],
                calibrated=bool(d["calibrated"]),
                normal_font_name=LIB.read_cstr(d["normal_font_name"]),
                platform=d["platform"],
            )

    def convert_column_chars_to_pt(self, chars: float, mode: GeometryMode = GeometryMode.DISPLAY) -> float:
        """Convert a stored column width (padding included) to points, snapped to whole screen pixels.

        Raises:
          InvalidArgumentError: when ``mode`` is not a :class:`GeometryMode` member.
        """
        mode_c = _geometry_mode(mode)
        sheet = self._index()
        h = self._workbook._require()
        return self._read_f64_out(
            LIB.fm_sheet_convert_column_chars_to_pt, h, _uint(sheet, "sheet_index", role="sheet"), float(chars), mode_c
        )

    def convert_column_pt_to_chars(self, pt: float, mode: GeometryMode = GeometryMode.DISPLAY) -> float:
        """Convert points to the stored column width Excel writes for that column.

        Raises:
          InvalidArgumentError: when ``mode`` is not a :class:`GeometryMode` member.
        """
        mode_c = _geometry_mode(mode)
        sheet = self._index()
        h = self._workbook._require()
        return self._read_f64_out(
            LIB.fm_sheet_convert_column_pt_to_chars, h, _uint(sheet, "sheet_index", role="sheet"), float(pt), mode_c
        )

    def _read_formula_text(self, fn, sheet: int, row: int, col: int) -> Optional[str]:
        h = self._workbook._require()
        return _read_optional_cstr_out(
            fn,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(row, "row", role="coordinate"),
            _uint(col, "col", role="coordinate"),
        )

    @overload
    @_coordinate_form
    def get_formula(self, row: int, col: int) -> Optional[str]: ...

    @overload
    def get_formula(self, a1: str) -> Optional[str]: ...

    def get_formula(self, *args: Any, **kwargs: Any) -> Optional[str]:
        """Return the stored A1 formula (leading ``=`` included), or ``None`` for a non-formula cell.

        A data table's result cell reads the table's formula,
        ``=TABLE(row_input,col_input)``, with an absent input left empty
        (``=TABLE(,B3)``) and a deleted one as ``#REF!``.
        """
        row, col = self._cell_args("get_formula", args, kwargs, ())
        sheet = self._index()
        return self._read_formula_text(LIB.fm_sheet_get_formula, sheet, row, col)

    @overload
    @_coordinate_form
    def get_formula_r1c1(self, row: int, col: int) -> Optional[str]: ...

    @overload
    def get_formula_r1c1(self, a1: str) -> Optional[str]: ...

    def get_formula_r1c1(self, *args: Any, **kwargs: Any) -> Optional[str]:
        """Return the formula in R1C1 notation relative to the cell (leading ``=`` included), or ``None`` for a non-formula cell."""
        row, col = self._cell_args("get_formula_r1c1", args, kwargs, ())
        sheet = self._index()
        return self._read_formula_text(LIB.fm_sheet_get_formula_r1c1, sheet, row, col)

    def get_cells_in_range(
        self,
        range: Union[CellRect, str, None] = None,
        *,
        cursor: Optional[int] = None,
        limit: Optional[int] = None,
    ) -> CellsInRange:
        """Return one page of the populated cells in ``range`` (row-major), each a :class:`RangeCell`.

        ``range`` defaults to :meth:`get_used_range`; a sheet with no cells reads an empty page.

        Pass ``next_cursor`` back to fetch the next page; it is ``None``
        when the rectangle holds no further populated cell. ``limit`` of
        ``None`` or ``0`` means the engine maximum.
        """
        rect = self._range_or_used(range)
        if rect is None:
            return CellsInRange([], None)
        cells, next_cursor = self._cell_page(rect, cursor, limit)
        return CellsInRange([_range_cell(*cell) for cell in cells], next_cursor)

    def _cell_page(
        self, rect: CellRect, cursor: Optional[int], limit: Optional[int]
    ) -> Tuple[List[Tuple[int, int, Optional[str], Value]], Optional[int]]:
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            return self._read_cell_range_page(
                LIB.fm_sheet_get_cells_in_range,
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                rp,
                _uint(cursor or 0, "cursor", 64),
                _uint(limit or 0, "limit"),
            )

    @staticmethod
    def _read_cell_range_page(fn, *args) -> Tuple[List[Tuple[int, int, Optional[str], Value]], Optional[int]]:
        """Call a ``(..., cell_range** out)`` engine function and drain its page as ``(row, col, formula, value)``."""
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(fn(*args, out))
            handle = LIB.read_u32(out)
        if handle == 0:
            raise FormulonError(
                StatusCode.BINDING_NULL_POINTER,
                "returned OK with a null cell-range handle",
                "",
            )
        try:
            count = _read_count(LIB.fm_cell_range_get_cell_count, handle)
            with _ArgScope() as scope:
                row_ptr = scope.out_ptr()
                col_ptr = scope.out_ptr()
                formula_ptr = scope.out_ptr()
                cursor_ptr = scope.out_u64()
                value_ptr = scope.alloc(fm_value_t_size)
                cells: List[Tuple[int, int, Optional[str], Value]] = []
                for i in range(count):
                    LIB.write_bytes(formula_ptr, b"\x00\x00\x00\x00")
                    _check(
                        LIB.fm_cell_range_get_cell_at(
                            handle, _uint(i, "idx", role="position"), row_ptr, col_ptr, formula_ptr, value_ptr
                        )
                    )
                    formula_addr = LIB.read_u32(formula_ptr)
                    cells.append(
                        (
                            LIB.read_u32(row_ptr),
                            LIB.read_u32(col_ptr),
                            LIB.read_cstr(formula_addr) if formula_addr else None,
                            Value._from_wasm(value_ptr),
                        )
                    )
                _check(LIB.fm_cell_range_next_cursor(handle, cursor_ptr))
                next_cursor = struct.unpack("<Q", LIB.read_bytes(cursor_ptr, 8))[0]
        finally:
            LIB.fm_cell_range_destroy(handle)
        return cells, (None if next_cursor == _CELL_RANGE_NO_CURSOR else next_cursor)

    def get_auto_filter(self) -> Optional[AutoFilter]:
        """Return the sheet AutoFilter, or ``None`` when the sheet has none."""
        sheet = self._index()
        return _read_auto_filter(self._workbook, LIB.fm_sheet_get_auto_filter, sheet)

    def set_auto_filter(self, filter: Optional[Union[AutoFilterInput, AutoFilter]]) -> None:
        """Replace the sheet AutoFilter, or remove it with ``None`` (no-op without one).

        Row visibility is unchanged; see :meth:`apply_auto_filter`. ``range``
        of the filter, of its sort state and of its sort conditions takes an
        unqualified A1 range too; an :class:`AutoFilter` read back is accepted
        in place of an :class:`AutoFilterInput`.
        """
        sheet = self._index()
        if filter is not None:
            filter = _normalize_auto_filter(filter, self._workbook)
        self._workbook._write_auto_filter(LIB.fm_sheet_set_auto_filter, filter, sheet)

    def apply_auto_filter(self) -> None:
        """Hide the body rows the sheet AutoFilter excludes and show the rest.

        Fails with NotFoundError when the sheet has no AutoFilter.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_apply_auto_filter(h, _uint(sheet, "sheet_index", role="sheet")))

    def clear_auto_filter_criteria(self) -> None:
        """Drop every column criterion of the sheet AutoFilter and show all its body rows.

        The AutoFilter itself stays (remove it with ``set_auto_filter(None)``);
        a filter without criteria is left as it is.

        Fails with NotFoundError when the sheet has no AutoFilter.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_auto_filter_criteria(h, _uint(sheet, "sheet_index", role="sheet")))

    def evaluate_auto_filter(self) -> AutoFilterEval:
        """Evaluate the sheet AutoFilter without changing the sheet.

        Body row ``first_row + i`` matches when ``match[i]`` is true.

        Fails with NotFoundError when the sheet has no AutoFilter.
        """
        sheet = self._index()
        return AutoFilterEval(*self._workbook._evaluate_auto_filter(LIB.fm_sheet_evaluate_auto_filter, sheet))

    # -- Validation checks -----------------------------------------------------
    @overload
    @_coordinate_form
    def validate_value(self, row: int, col: int, value: CellValueInput) -> ValidationOutcome: ...

    @overload
    def validate_value(self, a1: str, value: CellValueInput) -> ValidationOutcome: ...

    def validate_value(self, *args: Any, **kwargs: Any) -> ValidationOutcome:
        """Check ``value`` against the data-validation rule covering ``(row, col)`` without storing it.

        ``value`` is anything :meth:`set_value` accepts, except that a ``str``
        is always text (a leading ``=`` does not make it a formula).
        """
        row, col, value = self._cell_args("validate_value", args, kwargs, ("value",))
        sheet = self._index()
        h = self._workbook._require()
        value = _check_value(_value_of_input(value, date1904=self._workbook.get_date1904()))
        with _ArgScope() as scope:
            value_ptr = scope.alloc(fm_value_t_size)
            ptr = scope.struct(S.VALIDATION_OUTCOME)
            _write_value_input(value, value_ptr, scope)
            _check(
                LIB.fm_sheet_validate_value(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    value_ptr,
                    ptr,
                )
            )
            d = S.VALIDATION_OUTCOME.unpack(LIB, ptr)
        has_rule = bool(d["has_rule"])
        return ValidationOutcome(
            valid=bool(d["valid"]),
            rule_id=LIB.read_cstr(d["rule_id"]) if has_rule else None,
            error_style=ValidationErrorStyle(d["error_style"]) if has_rule else None,
        )

    def get_invalid_cells(self, *, cursor: Optional[int] = None, limit: Optional[int] = None) -> CellsInRange:
        """Return one page of the populated cells failing their data-validation rule (row-major).

        Paging follows :meth:`get_cells_in_range`: pass ``next_cursor`` back
        for the next page; it is ``None`` once no further invalid cell exists.
        """
        sheet = self._index()
        h = self._workbook._require()
        cells, next_cursor = self._read_cell_range_page(
            LIB.fm_sheet_get_invalid_cells,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(cursor or 0, "cursor", 64),
            _uint(limit or 0, "limit"),
        )
        return CellsInRange([_range_cell(*cell) for cell in cells], next_cursor)

    # -- Threaded comments and persons -------------------------------------------
    def _cell_address(self, cell: Union[CellAddress, Tuple[int, int], str], label: str) -> Tuple[int, int]:
        """The ``(row, col)`` a :class:`CellAddress`, a ``(row, col)`` pair or an unqualified A1 cell names."""
        if isinstance(cell, str):
            return self._cell(cell)
        if isinstance(cell, CellAddress):
            return cell.row, cell.col
        if isinstance(cell, tuple) and len(cell) == 2:
            return cell
        raise TypeError(_argument_message(label, "must be a CellAddress, a (row, col) pair or an A1 string"))

    @overload
    def list_threaded_comments(self) -> List[ThreadedComment]: ...

    @overload
    def list_threaded_comments(self, row: int, col: int) -> List[ThreadedComment]: ...

    @overload
    def list_threaded_comments(self, a1: str) -> List[ThreadedComment]: ...

    def list_threaded_comments(self, *args: Any) -> List[ThreadedComment]:
        """Return the sheet's threaded comments in file order (openers followed by their replies).

        A cell, given as ``row, col`` or one A1 string, keeps the comments anchored at that cell: its
        thread, or nothing when it has none.

        Raises:
          TypeError: when the cell is neither ``row, col`` nor one A1 string.
        """
        sheet = self._index()
        anchor: Optional[Tuple[int, int]] = None
        if len(args) == 1 and isinstance(args[0], str):
            anchor = self._cell(args[0])
        elif len(args) == 2 and not any(
            isinstance(a, (bool, str)) or not isinstance(a, numbers.Integral) for a in args
        ):
            anchor = (args[0], args[1])
        elif args:
            raise TypeError(_argument_message("cell", "must be (row, col) or one A1 string"))
        h = self._workbook._require()
        n = _read_count(LIB.fm_sheet_get_threaded_comment_count, h, _uint(sheet, "sheet_index", role="sheet"))
        with _ArgScope() as scope:
            ptr = scope.struct(S.THREADED_COMMENT)
            out: List[ThreadedComment] = []
            for i in range(n):
                S.zero_struct(LIB, S.THREADED_COMMENT, ptr)
                _check(
                    LIB.fm_sheet_get_threaded_comment_at(
                        h, _uint(sheet, "sheet_index", role="sheet"), _uint(i, "idx", role="position"), ptr
                    )
                )
                d = S.THREADED_COMMENT.unpack(LIB, ptr)
                out.append(_threaded_comment_from(d))
        if anchor is not None:
            out = [c for c in out if (c.cell.row, c.cell.col) == anchor]
        return out

    def add_threaded_comment(self, comment: Union[ThreadedCommentInput, ThreadedComment]) -> str:
        """Add a threaded comment, a new thread when ``parent_id`` is ``None`` else a reply; return its id.

        ``comment.cell`` is a :class:`CellAddress` or an unqualified A1 cell, ``A1`` when omitted. An omitted
        ``id`` is generated and an omitted ``created`` is the current UTC time, or the pinned time as written when
        the clock is pinned (:meth:`Workbook.set_pinned_now`). ``created`` is read as UTC: a naive ``datetime`` is UTC,
        an aware one is converted, a ``date`` is midnight. A
        :class:`ThreadedComment` read back is accepted in place of a :class:`ThreadedCommentInput`.

        Raises:
          AlreadyExistsError: when the id is already used on the sheet, or a
            new thread's cell already holds a note or a thread.
          NotFoundError: when ``person_id`` names no person of the workbook or ``parent_id`` no comment
            of the sheet.
          InvalidArgumentError: when ``parent_id`` is an empty string.
        """
        row, col = (0, 0) if comment.cell is None else self._cell_address(comment.cell, "cell")
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.THREADED_COMMENT, 1)
            S.THREADED_COMMENT.pack(
                LIB,
                ptr,
                {
                    "row": _uint(row, "row", role="coordinate"),
                    "col": _uint(col, "col", role="coordinate"),
                    "done": int(_require_bool(comment.done, "comment.done")),
                },
            )
            for fld, text in (
                ("id", comment.id),
                ("person_id", comment.person_id),
                ("created", _format_created(comment.created)),
                ("text", comment.text),
            ):
                S.write_str_field(LIB, ptr, S.THREADED_COMMENT, fld, text, scope)
            # An empty parent id reaches the engine, which refuses it.
            S.write_str_field(LIB, ptr, S.THREADED_COMMENT, "parent_id", comment.parent_id, scope, allow_empty=True)
            _write_mentions(ptr, S.THREADED_COMMENT, comment.mentions, scope)
            out = scope.out_ptr()
            _check(LIB.fm_sheet_add_threaded_comment(h, _uint(sheet, "sheet_index", role="sheet"), ptr, out))
            return LIB.read_cstr(LIB.read_u32(out))

    def update_threaded_comment(
        self,
        id: str,
        *,
        text: Union[str, Unchanged] = UNCHANGED,
        mentions: Union[Sequence[Union[MentionInput, Mention]], Unchanged] = UNCHANGED,
        done: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the text, mentions and resolved state of the comment ``id``; an omitted argument is kept.

        ``mentions`` takes :class:`MentionInput` or :class:`Mention` records. ``done`` applies to a thread's
        opening comment only.

        Raises:
          TypeError: when an argument is ``None`` or ``done`` is not a ``bool``.
          NotFoundError: when the sheet has no comment ``id``.
          InvalidArgumentError: when ``done`` is given for a reply.
        """
        for name, value in (("text", text), ("mentions", mentions), ("done", done)):
            if value is None:
                raise TypeError(_argument_message(name, "must not be null"))
        values: Dict[str, object] = {}
        if done is not UNCHANGED:
            values["done_op"] = _PATCH_SET
            values["done"] = int(_require_bool(done, "done"))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            id_ptr = scope.cstr(id)
            if text is not UNCHANGED:
                values["text_op"] = _PATCH_SET
                values["text"] = scope.cstr(text)
            if mentions is not UNCHANGED:
                values["mentions_op"] = _PATCH_SET
                values["mentions"] = _pack_mentions(mentions, scope)
                values["mentions_count"] = _uint(len(mentions), "mentions_count")
            ptr = scope.struct(S.THREADED_COMMENT_PATCH)
            S.THREADED_COMMENT_PATCH.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_update_threaded_comment(h, _uint(sheet, "sheet_index", role="sheet"), id_ptr, ptr))

    def get_threaded_comment(self, id: str) -> Optional[ThreadedComment]:
        """Return the threaded comment ``id``, or ``None`` when the sheet has none."""
        sheet = self._index()
        return self._workbook._lookup(
            LIB.fm_sheet_get_threaded_comment,
            (_uint(sheet, "sheet_index", role="sheet"), id),
            S.THREADED_COMMENT,
            _threaded_comment_from,
        )

    def remove_threaded_comment(self, id: str) -> None:
        """Remove the comment ``id``; removing a thread opener removes its whole thread."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            id_ptr = scope.cstr(id)
            _check(LIB.fm_sheet_remove_threaded_comment(h, _uint(sheet, "sheet_index", role="sheet"), id_ptr))

    # -- Drawing objects ---------------------------------------------------
    #
    # Each object is a handle of the class its kind selects (see
    # ``drawing._drawing_for``); the sheet keeps one handle per object id in
    # ``_drawings`` and never drops it, so a restore that puts an id back makes
    # its old handle address the object again.
    def _drawing_handle(self, object_id: int) -> Drawing:
        from ._drawing import _drawing_for

        return _drawing_for(self, object_id)

    def _drawing_id(self, ref: Union[int, Drawing], where: str) -> int:
        """The object id ``ref`` names: an id, or a :class:`Drawing` of this sheet."""
        from ._drawing import Drawing

        if isinstance(ref, Drawing):
            if ref.sheet is not self:
                raise FormulonError(
                    StatusCode.INVALID_ARGUMENT, _argument_message(where, "is a Drawing of another sheet"), ""
                )
            return ref.object_id
        if isinstance(ref, bool) or not isinstance(ref, numbers.Integral):
            raise TypeError(_argument_message(where, "must be an object id or a Drawing"))
        return int(ref)

    def _typed_drawing(self, ref: Union[int, Drawing], cls: type, where: str) -> Any:
        """The handle of ``ref`` when it is a ``cls``; ``InvalidArgumentError`` for another kind."""
        drawing = self._drawing_handle(self._drawing_id(ref, where))
        if not isinstance(drawing, cls):
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _argument_message(where, f"{drawing.object_id} is not a {cls.__name__}"),
                "",
            )
        return drawing

    def _object_ids(self, count_fn: Any, at_fn: Any, layout: Any) -> List[int]:
        """The ``object_id`` of each record a count / at pair of the sheet lists, in order."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(count_fn, h, _uint(sheet, "sheet_index", role="sheet"))
        out: List[int] = []
        with _ArgScope() as scope:
            ptr = scope.struct(layout)
            for i in range(n):
                S.zero_struct(LIB, layout, ptr)
                _check(at_fn(h, _uint(sheet, "sheet_index", role="sheet"), _uint(i, "idx", role="position"), ptr))
                out.append(layout.unpack(LIB, ptr)["object_id"])
        return out

    def drawing(self, object_id: Union[int, Drawing]) -> Drawing:
        """Return the handle of drawing object ``object_id``, an instance of the class its kind selects.

        Raises:
          NotFoundError: when the sheet's drawing has no such object.
          InvalidArgumentError: when ``object_id`` is a :class:`Drawing` of another sheet.
        """
        return self._drawing_handle(self._drawing_id(object_id, "object_id"))

    def picture(self, object_id: Union[int, Drawing]) -> Picture:
        """Return the :class:`Picture` handle of ``object_id``.

        Raises:
          NotFoundError: when the sheet's drawing has no such object.
          InvalidArgumentError: when the object is not a picture.
        """
        from ._drawing import Picture

        return cast("Picture", self._typed_drawing(object_id, Picture, "object_id"))

    def chart(self, object_id: Union[int, Drawing]) -> Chart:
        """Return the :class:`Chart` handle of ``object_id``; fails as :meth:`picture` does when it is not a chart."""
        from ._drawing import Chart

        return cast("Chart", self._typed_drawing(object_id, Chart, "object_id"))

    def shape(self, object_id: Union[int, Drawing]) -> Shape:
        """Return the :class:`Shape` handle of ``object_id`` (a shape, connector or group); fails as :meth:`picture` does otherwise."""
        from ._drawing import Shape

        return cast("Shape", self._typed_drawing(object_id, Shape, "object_id"))

    def slicer(self, object_id: Union[int, Drawing]) -> Slicer:
        """Return the :class:`Slicer` handle of ``object_id`` (a table slicer's frame); fails as :meth:`picture` does otherwise."""
        from ._drawing import Slicer

        return cast("Slicer", self._typed_drawing(object_id, Slicer, "object_id"))

    def list_drawings(self) -> List[Drawing]:
        """Return the handles of the anchored objects of the sheet's drawing in document order.

        That order is the stacking order; ``0`` is the back.
        """
        ids = self._object_ids(
            LIB.fm_sheet_get_drawing_object_count, LIB.fm_sheet_get_drawing_object_at, S.DRAWING_OBJECT
        )
        return [self._drawing_handle(object_id) for object_id in ids]

    def list_pictures(self) -> List[Picture]:
        """Return the handles of the sheet's pictures, in document order."""
        from ._drawing import Picture

        return [drawing for drawing in self.list_drawings() if isinstance(drawing, Picture)]

    def list_charts(self) -> List[Chart]:
        """Return the handles of the sheet's charts, in document order."""
        from ._drawing import Chart

        return [drawing for drawing in self.list_drawings() if isinstance(drawing, Chart)]

    def list_shapes(self) -> List[Shape]:
        """Return the handles of the shapes, connectors and groups of the sheet's drawing in z-order (back first).

        Children of a group are not listed.
        """
        from ._drawing import Shape

        return [drawing for drawing in self.list_drawings() if isinstance(drawing, Shape)]

    def list_slicers(self) -> List[Slicer]:
        """Return the handles of the table slicers the sheet hosts; slicers over a pivot cache are not listed."""
        from ._drawing import Slicer

        return [drawing for drawing in self.list_drawings() if isinstance(drawing, Slicer)]

    def add_picture(
        self, bytes: Union[bytes, bytearray, memoryview], input: Optional[DrawingPlacementInput] = None
    ) -> Picture:
        """Insert ``bytes`` as a picture placed as ``input`` says and return its handle.

        An empty ``name`` becomes ``Picture <id>``. An ``anchor`` of ``None``
        (and no ``input`` at all) is a one-cell anchor at A1. A ``width_emu``
        or ``height_emu`` of ``0`` means the image's pixel size times 9525.
        ``edit_as`` is honoured for a two-cell anchor only. Existing drawing
        content is kept.
        """
        sheet = self._index()
        h = self._workbook._require()
        raw = builtins.bytes(bytes)
        with _ArgScope() as scope:
            out = scope.out_ptr()
            buf = scope.buffer(raw)
            ptr = scope.struct(S.DRAWING_PLACEMENT_SPEC)
            S.DRAWING_PLACEMENT_SPEC.pack(LIB, ptr, _placement_fields(scope, input))
            _check(
                LIB.fm_sheet_add_picture(
                    h, _uint(sheet, "sheet_index", role="sheet"), buf, _uint(len(raw), "len"), ptr, out
                )
            )
            object_id = LIB.read_u32(out)
        return cast("Picture", self._drawing_handle(object_id))

    def add_shape(self, spec: ShapeInput) -> Shape:
        """Append the shape ``spec`` describes on top of the sheet's drawing and return its handle.

        A ``None`` field is not applied; see :class:`ShapeInput` for the
        defaults, the anchor rules and the units. A colour needs its kind
        (``TypeError`` otherwise), and a colour with ``NONE`` fails with a
        status error.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = _pack_shape_spec(scope, spec)
            _check(LIB.fm_sheet_add_shape(h, _uint(sheet, "sheet_index", role="sheet"), ptr, out))
            object_id = LIB.read_u32(out)
        return cast("Shape", self._drawing_handle(object_id))

    def add_slicer(self, table: Union[str, Table], column: Union[int, str], spec: SlicerInput) -> Slicer:
        """Add a slicer over column ``column`` of ``table`` (a :class:`Table` or its name); return its handle.

        ``table`` may live on another sheet; a name is compared
        case-insensitively. ``column`` is a 0-based position or the column's
        name, compared case-insensitively. The frame is two-cell: ``spec.anchor``
        gives its box or its size, as for :class:`ShapeInput`. The name and
        caption default to the column name (the name is made unique),
        ``row_height_pt`` to 19 and ``column_count`` to 1.

        Raises:
          NotFoundError: when no table has the name ``table`` or the table has no column named ``column``.
          TypeError: when ``table`` is neither a ``str`` nor a :class:`Table`, or ``column`` is neither an
            ``int`` nor a ``str``.
        """
        from ._table import Table

        if isinstance(table, Table):
            found = table
        elif isinstance(table, str):
            found = self._workbook.table(table)
        else:
            raise TypeError(_argument_message("table", "must be a Table or a string"))
        table_sheet, table_index = found.sheet.index, found.index
        if isinstance(column, str):
            wanted = _fold(column)
            names = [c.name for c in found.get_properties().columns]
            matches = [i for i, name in enumerate(names) if _fold(name) == wanted]
            if not matches:
                raise FormulonError(StatusCode.NOT_FOUND, f'the table has no column "{column}"', f"column={column}")
            column = matches[0]
        elif isinstance(column, bool) or not isinstance(column, int):
            raise TypeError(_argument_message("column", "must be an integer or a string"))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = _pack_slicer_spec(scope, spec)
            _check(
                LIB.fm_sheet_add_slicer(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(table_sheet, "table_sheet_index", role="sheet"),
                    _uint(table_index, "table_index", role="position"),
                    _uint(column, "column_index", role="position"),
                    ptr,
                    out,
                )
            )
            object_id = LIB.read_u32(out)
        return cast("Slicer", self._drawing_handle(object_id))

    def restore_drawing(
        self, bytes: Union[bytes, bytearray, memoryview], *, assign_new_id: bool = False
    ) -> Union[Picture, Chart, Shape]:
        """Put a ``snapshot()`` capture back and return the object's handle.

        An object with the same id is replaced and the capture returns to its
        recorded list position; a handle of that id addresses it again. With
        ``assign_new_id`` it is added on top as a copy with a fresh id.
        """
        sheet = self._index()
        h = self._workbook._require()
        raw = builtins.bytes(bytes)
        with _ArgScope() as scope:
            buf = scope.buffer(raw)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_restore_drawing(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    buf,
                    _uint(len(raw), "len"),
                    int(_require_bool(assign_new_id, "assign_new_id")),
                    out,
                )
            )
            object_id = LIB.read_u32(out)
        return cast("Union[Picture, Chart, Shape]", self._drawing_handle(object_id))

    def remove_drawing(self, object_id: Union[int, Drawing]) -> None:
        """Remove drawing object ``object_id`` and the parts only it reaches; other drawing content is kept.

        A shape, connector or group goes with its children; a slicer with its
        frame, and its cache when no other slicer uses it.

        Raises:
          NotFoundError: when the sheet's drawing has no such object.
          InvalidArgumentError: when the sheet is a chartsheet, or ``object_id``
            is a :class:`Drawing` of another sheet.
        """
        drawing = self.drawing(object_id)
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_remove_drawing(
                h, _uint(sheet, "sheet_index", role="sheet"), _uint(drawing.object_id, "object_id")
            )
        )

    def add_data_table(self, spec: Union[DataTableInput, DataTable]) -> CellAddress:
        """Create the data table ``spec`` describes and return the top-left cell of its result range.

        The returned cell is the one :meth:`get_data_table` and :meth:`remove_data_table` take.

        The formulas sit in the row above the range and the substitution values in the column left of it (a
        one-variable table), or the other way round; ``spec.row_input`` / ``spec.column_input`` are the input cells
        that receive the values of the row above / the column left. Give one for a one-variable table, both
        for a two-variable one. ``spec.range`` is a :class:`CellRect` or an unqualified A1 range; an input cell is a
        :class:`CellAddress`, a :class:`DataTableInputCell` or one A1 cell string. The result cells read 0 until the
        next recalc and refuse writes until the table is removed.
        """
        if isinstance(spec, DataTable):
            spec = DataTableInput(range=spec.range, row_input=spec.row_input, column_input=spec.column_input)
        if not isinstance(spec, DataTableInput):
            raise TypeError(_argument_message("spec", "must be a DataTableInput or DataTable"))
        rect = self._rect(spec.range)
        inputs = [
            self._cell(cell) if isinstance(cell, str) else None if cell is None else (cell.row, cell.col)
            for cell in (spec.row_input, spec.column_input)
        ]
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.DATA_TABLE, 1)
            values: Dict[str, object] = {"rect": _cell_rect_fields(rect)}
            for prefix, cell in (("row_input", inputs[0]), ("column_input", inputs[1])):
                if cell is not None:
                    values["has_" + prefix] = 1
                    values[prefix + "_row"] = _uint(cell[0], prefix + "_row")
                    values[prefix + "_col"] = _uint(cell[1], prefix + "_col")
            S.DATA_TABLE.pack(LIB, ptr, values)
            out_cell = scope.struct(S.CELL_ADDRESS)
            _check(LIB.fm_sheet_add_data_table(h, _uint(sheet, "sheet_index", role="sheet"), ptr, out_cell))
            d = S.CELL_ADDRESS.unpack(LIB, out_cell)
        return CellAddress(row=d["row"], col=d["col"])

    @overload
    @_coordinate_form
    def get_data_table(self, row: int, col: int) -> Optional[DataTable]: ...

    @overload
    def get_data_table(self, a1: str) -> Optional[DataTable]: ...

    def get_data_table(self, *args: Any, **kwargs: Any) -> Optional[DataTable]:
        """Return the data table whose result range contains ``(row, col)``, or ``None``."""
        row, col = self._cell_args("get_data_table", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.DATA_TABLE)
            _check(
                LIB.fm_sheet_get_data_table(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    present,
                    ptr,
                )
            )
            if not LIB.read_u32(present):
                return None
            d = S.DATA_TABLE.unpack(LIB, ptr)

        def _input(prefix: str) -> Optional[DataTableInputCell]:
            if not d["has_" + prefix]:
                return None
            return DataTableInputCell(d[prefix + "_row"], d[prefix + "_col"], bool(d[prefix + "_deleted"]))

        return DataTable(
            range=_cell_rect_from(d["rect"]),
            two_dimensional=bool(d["two_dimensional"]),
            row_oriented=bool(d["row_oriented"]),
            row_input=_input("row_input"),
            column_input=_input("column_input"),
        )

    @overload
    @_coordinate_form
    def remove_data_table(self, row: int, col: int) -> None: ...

    @overload
    def remove_data_table(self, a1: str) -> None: ...

    def remove_data_table(self, *args: Any, **kwargs: Any) -> None:
        """Remove the data table containing ``(row, col)`` and blank its result cells."""
        row, col = self._cell_args("remove_data_table", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_remove_data_table(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
            )
        )

    def list_sparkline_groups(self) -> List[SparklineGroup]:
        """Every sparkline group of the sheet, in order."""
        return [self._sparkline_group_at(i) for i in range(self._sparkline_group_count())]

    def _sparkline_group_count(self) -> int:
        """Return the number of sparkline groups on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(LIB.fm_sheet_get_sparkline_group_count, h, _uint(sheet, "sheet_index", role="sheet"))

    def _sparkline_group_at(self, group_index: int) -> SparklineGroup:
        """Return the sheet's ``group_index``-th sparkline group in document order with its sparklines."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            group_ptr = scope.struct(S.SPARKLINE_GROUP)
            item_ptr = scope.struct(S.SPARKLINE)
            _check(
                LIB.fm_sheet_get_sparkline_group_at(
                    h, _uint(sheet, "sheet_index", role="sheet"), _uint(group_index, "idx", role="position"), group_ptr
                )
            )
            d = S.SPARKLINE_GROUP.unpack(LIB, group_ptr)
            # Read the scratch-backed strings before the date-axis decode makes another scratch-backed read.
            group_id = LIB.read_cstr(d["id"])
            date_axis_range, stored_date_axis = _decode_sparkline_date_axis(
                self._workbook, sheet, LIB.read_cstr(d["date_axis_ref"])
            )
            group = SparklineGroup(
                id=group_id,
                type=d["type"],
                color_series=_decode_optional_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_series"][1]),
                color_negative=_decode_optional_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_negative"][1]),
                color_axis=_decode_optional_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_axis"][1]),
                color_markers=_decode_optional_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_markers"][1]),
                color_first=_decode_optional_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_first"][1]),
                color_last=_decode_optional_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_last"][1]),
                color_high=_decode_optional_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_high"][1]),
                color_low=_decode_optional_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_low"][1]),
                markers=bool(d["markers"]),
                high=bool(d["high"]),
                low=bool(d["low"]),
                first=bool(d["first"]),
                last=bool(d["last"]),
                negative=bool(d["negative"]),
                display_x_axis=bool(d["display_x_axis"]),
                display_hidden=bool(d["display_hidden"]),
                min_axis_type=d["min_axis_type"],
                max_axis_type=d["max_axis_type"],
                manual_min=d["manual_min"] if d["has_manual_min"] else None,
                manual_max=d["manual_max"] if d["has_manual_max"] else None,
                date_axis=bool(d["date_axis"]),
                date_axis_range=date_axis_range,
                line_width_pt=d["line_width_pt"],
                empty_cells_as=d["empty_cells_as"],
            )
            object.__setattr__(group, "_stored_date_axis", stored_date_axis)
            for k in range(d["sparkline_count"]):
                S.zero_struct(LIB, S.SPARKLINE, item_ptr)
                _check(
                    LIB.fm_sheet_get_sparkline_at(
                        h,
                        _uint(sheet, "sheet_index", role="sheet"),
                        _uint(group_index, "idx", role="position"),
                        _uint(k, "idx", role="position"),
                        item_ptr,
                    )
                )
                item = S.SPARKLINE.unpack(LIB, item_ptr)
                group.sparklines.append(_decode_sparkline(self._workbook, sheet, item))
            return group

    def get_sparkline_group(self, sparkline_group: SparklineGroupRef) -> Optional[SparklineGroup]:
        """Return ``sparkline_group`` (an id or a position in :meth:`list_sparkline_groups`), or ``None`` when absent.

        Raises:
          IndexOutOfRangeError: when ``sparkline_group`` is a position past the last group.
        """
        if not isinstance(sparkline_group, str):
            return self._sparkline_group_at(self._sparkline_group_position(sparkline_group))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.SPARKLINE_GROUP)
            _check(
                LIB.fm_sheet_get_sparkline_group(
                    h, _uint(sheet, "sheet_index", role="sheet"), scope.cstr(sparkline_group), present, ptr
                )
            )
            if not LIB.read_u32(present):
                return None
            # The sparklines of a group are read by its position.
            for position in range(self._sparkline_group_count()):
                S.zero_struct(LIB, S.SPARKLINE_GROUP, ptr)
                _check(
                    LIB.fm_sheet_get_sparkline_group_at(
                        h, _uint(sheet, "sheet_index", role="sheet"), _uint(position, "idx", role="position"), ptr
                    )
                )
                if LIB.read_cstr(S.SPARKLINE_GROUP.unpack(LIB, ptr)["id"]) == sparkline_group:
                    return self._sparkline_group_at(position)
        return None

    def _sparkline_group_position(self, group: SparklineGroupRef) -> int:
        if isinstance(group, bool):
            raise TypeError(_argument_message("group", "must be an integer or a string"))
        return _uint(group, "group", role="position")

    def add_sparkline_group(self, spec: Union[SparklineGroupInput, SparklineGroup]) -> SparklineGroupId:
        """Append the group ``spec`` with its ``sparklines`` (at least one) to the sheet's worksheet extension list.

        A sparkline's ``source`` and the group's ``date_axis_range`` are a
        :class:`SheetRect` or a sheet-qualified A1 range (an unqualified one is
        on this sheet); ``location`` is a :class:`CellAddress` or an A1 cell. A
        :class:`SparklineGroup` read back is accepted in place of a
        :class:`SparklineGroupInput`; its ``id`` is ignored. Returns the new group's :data:`SparklineGroupId`.

        Raises:
          TypeError: when ``spec`` is not one of the two record types.
          InvalidArgumentError: when ``spec.sparklines`` is empty.
        """
        if not isinstance(spec, (SparklineGroupInput, SparklineGroup)):
            raise TypeError(_argument_message("spec", "must be a SparklineGroupInput"))
        if not spec.sparklines:
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _argument_message("spec.sparklines", "needs at least one sparkline"),
                "sparklines_count=0",
            )
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_sparkline_group(scope, spec, self._workbook, sheet)
            items = _pack_sparklines(scope, spec.sparklines, self._workbook, sheet)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_add_sparkline_group(
                    h, _uint(sheet, "sheet_index", role="sheet"), ptr, items, _uint(len(spec.sparklines), "count"), out
                )
            )
            return LIB.read_cstr(LIB.read_u32(out))

    def _sparkline_group_id(self, group: SparklineGroupRef) -> str:
        """The :data:`SparklineGroupId` ``group`` names: an id as it is, a position through its group."""
        if isinstance(group, str):
            return group
        if isinstance(group, bool):
            raise TypeError(_argument_message("group", "must be an integer or a string"))
        return self._sparkline_group_at(_uint(group, "group", role="position")).id

    def update_sparkline_group(
        self,
        sparkline_group: SparklineGroupRef,
        *,
        type: Union[SparklineType, Unchanged] = UNCHANGED,
        color_series: Union[ColorInput, None, Unchanged] = UNCHANGED,
        color_negative: Union[ColorInput, None, Unchanged] = UNCHANGED,
        color_axis: Union[ColorInput, None, Unchanged] = UNCHANGED,
        color_markers: Union[ColorInput, None, Unchanged] = UNCHANGED,
        color_first: Union[ColorInput, None, Unchanged] = UNCHANGED,
        color_last: Union[ColorInput, None, Unchanged] = UNCHANGED,
        color_high: Union[ColorInput, None, Unchanged] = UNCHANGED,
        color_low: Union[ColorInput, None, Unchanged] = UNCHANGED,
        markers: Union[bool, Unchanged] = UNCHANGED,
        high: Union[bool, Unchanged] = UNCHANGED,
        low: Union[bool, Unchanged] = UNCHANGED,
        first: Union[bool, Unchanged] = UNCHANGED,
        last: Union[bool, Unchanged] = UNCHANGED,
        negative: Union[bool, Unchanged] = UNCHANGED,
        display_x_axis: Union[bool, Unchanged] = UNCHANGED,
        display_hidden: Union[bool, Unchanged] = UNCHANGED,
        min_axis_type: Union[SparklineAxisType, Unchanged] = UNCHANGED,
        max_axis_type: Union[SparklineAxisType, Unchanged] = UNCHANGED,
        manual_min: Union[float, None, Unchanged] = UNCHANGED,
        manual_max: Union[float, None, Unchanged] = UNCHANGED,
        date_axis: Union[bool, Unchanged] = UNCHANGED,
        date_axis_range: Union[SheetRect, str, None, Unchanged] = UNCHANGED,
        line_width_pt: Union[float, Unchanged] = UNCHANGED,
        empty_cells_as: Union[BlanksAs, Unchanged] = UNCHANGED,
        sparklines: Union[Sequence[Union[SparklineInput, Sparkline]], Unchanged] = UNCHANGED,
    ) -> None:
        """Update ``sparkline_group``, a :data:`SparklineGroupId` or a position in :meth:`list_sparkline_groups`; an
        omitted argument keeps the group's value.

        The arguments are the :class:`SparklineGroup` fields. ``None`` clears
        ``manual_min``, ``manual_max`` and ``date_axis_range`` and is a
        ``TypeError`` for any other argument. Every attribute or child outside
        the model is kept.

        Raises:
          TypeError: for ``None`` on an argument that is not clearable.
          NotFoundError: when no group has the id ``sparkline_group``.
          IndexOutOfRangeError: when ``sparkline_group`` is a position past the last group.
          InvalidArgumentError: when ``sparklines`` is empty (a group holds at
            least one sparkline; remove the group instead).
        """
        given = {
            key: value
            for key, value in (
                ("type", type),
                ("color_series", color_series),
                ("color_negative", color_negative),
                ("color_axis", color_axis),
                ("color_markers", color_markers),
                ("color_first", color_first),
                ("color_last", color_last),
                ("color_high", color_high),
                ("color_low", color_low),
                ("markers", markers),
                ("high", high),
                ("low", low),
                ("first", first),
                ("last", last),
                ("negative", negative),
                ("display_x_axis", display_x_axis),
                ("display_hidden", display_hidden),
                ("min_axis_type", min_axis_type),
                ("max_axis_type", max_axis_type),
                ("manual_min", manual_min),
                ("manual_max", manual_max),
                ("date_axis", date_axis),
                ("date_axis_range", date_axis_range),
                ("line_width_pt", line_width_pt),
                ("empty_cells_as", empty_cells_as),
                ("sparklines", sparklines),
            )
            if value is not UNCHANGED
        }
        for key, value in given.items():
            if value is None and key not in ("manual_min", "manual_max", "date_axis_range", *_SPARKLINE_COLORS):
                raise TypeError(_argument_message(key, "must not be null"))
        group_id = self._sparkline_group_id(sparkline_group)
        sheet = self._index()
        h = self._workbook._require()
        values: Dict[str, object] = {}
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.SPARKLINE_GROUP_PATCH, 1)
            for key in _SPARKLINE_COLORS:
                if key in given:
                    _pack_optional_color(
                        ptr + S.SPARKLINE_GROUP_PATCH.offsets[key][1], given[key], f"SparklineGroup.{key}"
                    )
                    values[key + "_op"] = _PATCH_SET
            for key in ("manual_min", "manual_max"):
                if key in given:
                    values[key + "_op"] = _PATCH_CLEAR if given[key] is None else _PATCH_SET
                    values[key] = 0.0 if given[key] is None else _finite(given[key], key)
            if "line_width_pt" in given:
                values["line_width_pt_op"] = _PATCH_SET
                values["line_width_pt"] = _finite(given["line_width_pt"], "line_width_pt")
            for key in ("type", "min_axis_type", "max_axis_type", "empty_cells_as"):
                if key in given:
                    values[key + "_op"] = _PATCH_SET
                    values[key] = _sint(given[key], key)
            for key in _SPARKLINE_FLAGS:
                if key in given:
                    values[key + "_op"] = _PATCH_SET
                    values[key] = int(_require_bool(given[key], key))
            if "date_axis_range" in given:
                text = _qualified_range_text(
                    self._workbook, given["date_axis_range"], sheet, "SparklineGroup.date_axis_range"
                )
                if text is None:
                    values["date_axis_ref_op"] = _PATCH_CLEAR
                else:
                    values["date_axis_ref_op"] = _PATCH_SET
                    values["date_axis_ref"] = scope.cstr(text)
            if "sparklines" in given:
                values["items_op"] = _PATCH_SET
                values["items"] = _pack_sparklines(scope, given["sparklines"], self._workbook, sheet)
                values["items_count"] = _uint(len(given["sparklines"]), "count")
            raw = bytearray(LIB.read_bytes(ptr, S.SPARKLINE_GROUP_PATCH.size))
            S.SPARKLINE_GROUP_PATCH.pack_into(raw, 0, values)
            LIB.write_bytes(ptr, bytes(raw))
            _check(
                LIB.fm_sheet_update_sparkline_group(
                    h, _uint(sheet, "sheet_index", role="sheet"), scope.cstr(group_id), ptr
                )
            )

    def remove_sparkline_group(self, sparkline_group: SparklineGroupRef) -> None:
        """Remove ``sparkline_group``, a :data:`SparklineGroupId` or a position in :meth:`list_sparkline_groups`.

        Raises:
          NotFoundError: when no group has the id ``sparkline_group``.
          IndexOutOfRangeError: when ``sparkline_group`` is a position past the last group.
        """
        group_id = self._sparkline_group_id(sparkline_group)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            _check(
                LIB.fm_sheet_remove_sparkline_group(h, _uint(sheet, "sheet_index", role="sheet"), scope.cstr(group_id))
            )

    def add_chart(self, input: ChartInput) -> Chart:
        """Create the chart ``input`` describes and return its handle.

        ``input.range`` is a :class:`SheetRect` or an A1 range, on this sheet
        when unqualified. ``input.grouping`` applies to bar, line and area
        charts. An omitted ``title`` (``ChartTitle(ChartTitleKind.AUTO)``) is
        Excel's automatic title, ``None`` is no title and an empty string is refused; ``legend`` is ``None`` for
        no legend. The placement fields (``name``, ``anchor``, ``width_emu`` ...)
        read as for :meth:`add_picture` (an empty ``name`` becomes
        ``Chart <id>``), except that a zero ``width_emu`` or ``height_emu`` means
        Excel's default 360 x 216 pt frame and an omitted anchor is a two-cell
        anchor at A1.

        Raises:
          TypeError: when ``input`` is not a :class:`ChartInput`, or for
            ``bar_direction`` on a type other than ``BAR``.
          InvalidArgumentError: when ``input.title`` is an empty string.

        Example::

            chart = ws.add_chart(ChartInput(type=ChartPlotType.BAR, range="A1:B5", title="Sales"))
        """
        if not isinstance(input, ChartInput):
            raise TypeError(_argument_message("input", "must be a ChartInput"))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = _encode_chart_insert(scope, self._workbook, sheet, input)
            _check(LIB.fm_sheet_add_chart(h, _uint(sheet, "sheet_index", role="sheet"), ptr, out))
            object_id = LIB.read_u32(out)
        return cast("Chart", self._drawing_handle(object_id))

    def get_merges_in_range(self, range: Union[CellRect, str]) -> List[CellRect]:
        """Return every merge range on the sheet that intersects ``range``."""
        cell_range = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            range_ptr = _pack_cell_rect_array([cell_range], scope)
            count_ptr = scope.out_ptr()
            # Size the buffer first, then fetch.
            _check(
                LIB.fm_sheet_get_merges_in_range(
                    h, _uint(sheet, "sheet_index", role="sheet"), range_ptr, 0, 0, count_ptr
                )
            )
            total = LIB.read_u32(count_ptr)
            if total == 0:
                return []
            size = S.CELL_RECT.size
            buf = scope.struct_array(S.CELL_RECT, total)
            _check(
                LIB.fm_sheet_get_merges_in_range(
                    h, _uint(sheet, "sheet_index", role="sheet"), range_ptr, buf, _uint(total, "capacity"), count_ptr
                )
            )
            out: List[CellRect] = []
            for i in builtins.range(min(total, LIB.read_u32(count_ptr))):  # `range` is the parameter here
                out.append(_decode_cell_rect_at(buf + i * size))
            return out

    @overload
    @_coordinate_form
    def get_display_text(self, row: int, col: int) -> DisplayText: ...

    @overload
    def get_display_text(self, a1: str) -> DisplayText: ...

    def get_display_text(self, *args: Any, **kwargs: Any) -> DisplayText:
        """Return the text Excel displays for the cell and the rendering outcome.

        Example::

            ws.get_display_text("A1").text  # "1,234.50" under a #,##0.00 format
        """
        row, col = self._cell_args("get_display_text", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            text_ptr = scope.out_ptr()
            status_ptr = scope.out_ptr()
            _check(
                LIB.fm_sheet_get_display_text(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    text_ptr,
                    status_ptr,
                )
            )
            return DisplayText(LIB.read_cstr(LIB.read_u32(text_ptr)), DisplayStatus(LIB.read_i32(status_ptr)))

    # -- Sheet view / layout -----------------------------------------------
    def get_pagination(self) -> Pagination:
        """Resolve the worksheet's print area, page breaks, and page count.

        The result is a snapshot: later workbook or layout changes do not
        mutate the returned lists.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_sheet_paginate(h, _uint(sheet, "sheet_index", role="sheet"), out))
            pagination = LIB.read_u32(out)
            if pagination == 0:
                raise FormulonError(
                    StatusCode.BINDING_NULL_POINTER,
                    "returned OK with a null pagination handle",
                    "",
                )
            try:
                page_count = _read_count(LIB.fm_pagination_get_page_count, pagination)
                range_count = _read_count(LIB.fm_pagination_get_print_area_count, pagination)
                break_count = _read_count(LIB.fm_pagination_get_row_break_count, pagination)
                column_break_count = _read_count(LIB.fm_pagination_get_column_break_count, pagination)
                range_ptr = scope.struct(S.CELL_RECT)
                value_ptr = scope.out_ptr()
                scale_ptr = scope.alloc(8)
                # Sized for the largest pagination detail struct (fm_page_layout_info_t).
                struct_ptr = scope.struct(S.PAGE_LAYOUT)
                print_area = []
                for i in range(range_count):
                    _check(
                        LIB.fm_pagination_get_print_area_at(pagination, _uint(i, "index", role="position"), range_ptr)
                    )
                    print_area.append(_decode_cell_rect_at(range_ptr))
                row_breaks = []
                row_break_ptr = scope.struct(S.ROW_BREAK)
                for i in range(break_count):
                    _check(
                        LIB.fm_pagination_get_row_break_at(
                            pagination, _uint(i, "index", role="position"), row_break_ptr
                        )
                    )
                    brk = S.ROW_BREAK.unpack(LIB, row_break_ptr)
                    row_breaks.append(
                        RowBreak(row=brk["row"], first=brk["first"], last=brk["last"], manual=bool(brk["manual"]))
                    )
                column_breaks = []
                column_break_ptr = scope.struct(S.COLUMN_BREAK)
                for i in range(column_break_count):
                    _check(
                        LIB.fm_pagination_get_column_break_at(
                            pagination, _uint(i, "index", role="position"), column_break_ptr
                        )
                    )
                    brk = S.COLUMN_BREAK.unpack(LIB, column_break_ptr)
                    column_breaks.append(
                        ColumnBreak(col=brk["col"], first=brk["first"], last=brk["last"], manual=bool(brk["manual"]))
                    )
                _check(LIB.fm_pagination_get_scale(pagination, scale_ptr))
                scale_pct = LIB.read_f64(scale_ptr)
                _check(LIB.fm_pagination_get_page_order(pagination, value_ptr))
                page_order = LIB.read_i32(value_ptr)

                def read_struct(layout: S.Struct, fn) -> Dict[str, int]:
                    S.zero_struct(LIB, layout, struct_ptr)
                    _check(fn(pagination, struct_ptr))
                    return layout.unpack(LIB, struct_ptr)

                d = read_struct(S.PAPER_INFO, LIB.fm_pagination_get_paper)
                paper = PaginationPaper(d["width_pt"], d["height_pt"], bool(d["landscape"]), bool(d["known"]))
                margins = PaginationMargins(**read_struct(S.MARGINS_PT, LIB.fm_pagination_get_margins))
                printable = RectPt(**read_struct(S.RECT_PT, LIB.fm_pagination_get_printable))
                d = read_struct(S.PRINT_TITLES, LIB.fm_pagination_get_print_titles)
                print_titles = _print_titles_from(d)
                if print_titles.rows is None and print_titles.columns is None:
                    print_titles = None
                pages = []
                for i in range(page_count):
                    S.zero_struct(LIB, S.PAGE_LAYOUT, struct_ptr)
                    _check(LIB.fm_pagination_get_page_at(pagination, _uint(i, "index", role="position"), struct_ptr))
                    page = S.PAGE_LAYOUT.unpack(LIB, struct_ptr)
                    page["range"] = _cell_rect_from(page.pop("rect"))
                    pages.append(PaginationPage(**page))
                return Pagination(
                    page_count,
                    print_area,
                    row_breaks,
                    column_breaks,
                    paper,
                    margins,
                    printable,
                    scale_pct,
                    page_order,
                    print_titles,
                    pages,
                )
            finally:
                LIB.fm_pagination_destroy(pagination)

    def get_view(self) -> SheetView:
        """Read the full per-sheet view (zoom, freeze, and the display / orientation flags).

        The sheet's visibility is :meth:`get_visibility`."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHEET_VIEW)
            _check(LIB.fm_sheet_get_view(h, _uint(sheet, "sheet_index", role="sheet"), ptr))
            d = S.SHEET_VIEW.unpack(LIB, ptr)
            return SheetView(
                zoom_pct=d["zoom_pct"],
                freeze_rows=d["freeze_rows"],
                freeze_cols=d["freeze_cols"],
                show_grid_lines=bool(d["show_grid_lines"]),
                show_row_column_headers=bool(d["show_row_column_headers"]),
                show_zeros=bool(d["show_zeros"]),
                right_to_left=bool(d["right_to_left"]),
                tab_selected=bool(d["tab_selected"]),
                view_mode=SheetViewMode(d["view_mode"]),
                top_left=CellAddress(row=d["top_left_row"], col=d["top_left_col"]),
                pane_top_left=CellAddress(row=d["pane_top_left_row"], col=d["pane_top_left_col"]),
                split_x_twips=d["split_x_twips"],
                split_y_twips=d["split_y_twips"],
                show_formulas=bool(d["show_formulas"]),
                show_outline_symbols=bool(d["show_outline_symbols"]),
            )

    def update_view(
        self,
        *,
        zoom_pct: Union[int, Unchanged] = UNCHANGED,
        freeze_rows: Union[int, Unchanged] = UNCHANGED,
        freeze_cols: Union[int, Unchanged] = UNCHANGED,
        split_x_twips: Union[float, Unchanged] = UNCHANGED,
        split_y_twips: Union[float, Unchanged] = UNCHANGED,
        top_left: Union[CellAddress, str, Unchanged] = UNCHANGED,
        pane_top_left: Union[CellAddress, str, Unchanged] = UNCHANGED,
        show_grid_lines: Union[bool, Unchanged] = UNCHANGED,
        show_row_column_headers: Union[bool, Unchanged] = UNCHANGED,
        show_zeros: Union[bool, Unchanged] = UNCHANGED,
        show_formulas: Union[bool, Unchanged] = UNCHANGED,
        show_outline_symbols: Union[bool, Unchanged] = UNCHANGED,
        right_to_left: Union[bool, Unchanged] = UNCHANGED,
        tab_selected: Union[bool, Unchanged] = UNCHANGED,
        view_mode: Union[SheetViewMode, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the sheet view in one call; an omitted argument keeps its value.

        Arguments are the :class:`SheetView` fields. ``zoom_pct`` outside
        ``[10, 400]`` is rejected. ``split_x_twips`` / ``split_y_twips`` are in
        twips; a split replaces any freeze and a freeze any split, and
        changing either resets the pane origin unless ``pane_top_left`` is
        given. ``top_left`` and ``pane_top_left`` are a :class:`CellAddress` or
        an A1 cell. One half of a pair (``freeze_rows`` without
        ``freeze_cols``...) keeps the other half.
        The sheet's visibility is :meth:`set_visibility`.

        The whole patch is checked before anything is written.

        Raises:
          TypeError: when an argument is ``None`` or a flag is not a ``bool``.
          InvalidArgumentError: when a non-zero freeze and a non-zero split
            are both given, for a pane origin on a view without panes, or for
            a value outside its domain.
        """
        flag_keys = (
            "show_grid_lines",
            "show_row_column_headers",
            "show_zeros",
            "show_formulas",
            "show_outline_symbols",
            "right_to_left",
            "tab_selected",
        )
        given = {
            key: value
            for key, value in (
                ("zoom_pct", zoom_pct),
                ("freeze_rows", freeze_rows),
                ("freeze_cols", freeze_cols),
                ("split_x_twips", split_x_twips),
                ("split_y_twips", split_y_twips),
                ("top_left", top_left),
                ("pane_top_left", pane_top_left),
                ("view_mode", view_mode),
                ("show_grid_lines", show_grid_lines),
                ("show_row_column_headers", show_row_column_headers),
                ("show_zeros", show_zeros),
                ("show_formulas", show_formulas),
                ("show_outline_symbols", show_outline_symbols),
                ("right_to_left", right_to_left),
                ("tab_selected", tab_selected),
            )
            if value is not UNCHANGED
        }
        for key, value in given.items():
            if value is None:
                raise TypeError(_argument_message(key, "must not be null"))
        for key in flag_keys:
            if key in given:
                _require_bool(given[key], key)
        if not given:
            return
        values: Dict[str, object] = {}
        for key in ("zoom_pct", "freeze_rows", "freeze_cols"):
            if key in given:
                values[key] = _uint(given[key], key)
        for key in ("split_x_twips", "split_y_twips"):
            if key in given:
                values[key] = _finite(given[key], key)
        for key in ("top_left", "pane_top_left"):
            if key in given:
                cell = CellAddress(*self._cell_address(given[key], key))
                values[key + "_row"] = _uint(cell.row, key + "_row")
                values[key + "_col"] = _uint(cell.col, key + "_col")
        if "view_mode" in given:
            values["view_mode"] = _sint(given["view_mode"], "view_mode")
        for key in flag_keys:
            if key in given:
                values[key] = int(_require_bool(given[key], key))
        for key in [k for k in values]:
            values[key + "_op"] = _PATCH_SET
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHEET_VIEW_PATCH)
            S.SHEET_VIEW_PATCH.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_update_view(h, _uint(sheet, "sheet_index", role="sheet"), ptr))

    def get_visibility(self) -> SheetVisibility:
        """The sheet's tab visibility: shown, hidden, or very hidden (absent from Excel's Unhide dialog)."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_sheet_get_visibility(h, _uint(sheet, "sheet_index", role="sheet"), out))
            return SheetVisibility(LIB.read_i32(out))

    def set_visibility(self, visibility: SheetVisibility) -> None:
        """Set the sheet's tab visibility.

        Raises:
          FormulonError: with ``code == StatusCode.PRECONDITION_FAILED`` when hiding the only visible sheet.
          InvalidArgumentError: for a value that is not a :class:`SheetVisibility` member.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_visibility(h, _uint(sheet, "sheet_index", role="sheet"), _sint(visibility, "visibility"))
        )

    def get_selection(self) -> SheetSelection:
        """Return the active cell and ranges of the active pane's selection.

        The selection reads as it is saved: a pane that stores none reports A1
        with the single range A1 (the pane's top-left cell for a frozen or split
        view's active pane), and an active cell outside every range moves to the
        first range's top-left.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            row_ptr = scope.out_ptr()
            col_ptr = scope.out_ptr()
            total_ptr = scope.out_ptr()
            array_ptr = 0
            # Size the buffer first, then fetch.
            _check(
                LIB.fm_sheet_get_selection(
                    h, _uint(sheet, "sheet_index", role="sheet"), row_ptr, col_ptr, 0, 0, total_ptr
                )
            )
            total = LIB.read_u32(total_ptr)
            array_ptr = scope.struct_array(S.CELL_RECT, total)
            _check(
                LIB.fm_sheet_get_selection(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    row_ptr,
                    col_ptr,
                    array_ptr,
                    _uint(total, "capacity"),
                    total_ptr,
                )
            )
            size = S.CELL_RECT.size
            ranges = [_decode_cell_rect_at(array_ptr + i * size) for i in range(total)]
            return SheetSelection(
                active=CellAddress(row=LIB.read_u32(row_ptr), col=LIB.read_u32(col_ptr)), ranges=ranges
            )

    def set_selection(self, selection: Union[SheetSelectionInput, SheetSelection]) -> None:
        """Replace the active pane's selection; empty ``ranges`` selects the active cell alone.

        ``selection.active`` is a :class:`CellAddress` or an A1 cell and
        ``selection.ranges`` takes :class:`CellRect` values, A1 ranges or one A1
        range string. A :class:`SheetSelection` read back is accepted in place
        of a :class:`SheetSelectionInput`.

        Raises:
          TypeError: when ``selection`` is not one of the two record types.
        """
        if not isinstance(selection, (SheetSelectionInput, SheetSelection)):
            raise TypeError(_argument_message("selection", "must be a SheetSelectionInput"))
        active_row, active_col = self._cell_address(selection.active, "selection.active")
        rects = _cell_rects(self._workbook, selection.ranges, "selection.ranges")
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_cell_rect_array(rects, scope)
            _check(
                LIB.fm_sheet_set_selection(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(active_row, "active_row", role="coordinate"),
                    _uint(active_col, "active_col", role="coordinate"),
                    ptr,
                    _uint(len(rects), "count"),
                )
            )

    def get_tab_color(self) -> Optional[Color]:
        """The sheet's tab colour (``<sheetPr><tabColor>``); ``None`` when the tab has none."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.COLOR_SPEC)
            _check(LIB.fm_sheet_get_tab_color(h, _uint(sheet, "sheet_index", role="sheet"), ptr))
            color = _decode_color(ptr)
        return None if color.kind == ColorKind.NONE else color

    def set_tab_color(self, color: Optional[ColorInput]) -> None:
        """Set the tab colour, a ``ColorInput``; ``None`` removes it."""
        spec = None if color is None else _format_color(color, "color")
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = 0
            if spec is not None:
                ptr = scope.struct(S.COLOR_SPEC)
                _pack_color(ptr, spec, "color")
            _check(LIB.fm_sheet_set_tab_color(h, _uint(sheet, "sheet_index", role="sheet"), ptr))

    def _get_print_xml(self, export: str) -> Optional[str]:
        sheet = self._index()
        h = self._workbook._require()
        return _read_cstr_out(getattr(LIB, export), h, _uint(sheet, "sheet_index", role="sheet")) or None

    def _set_print_xml(self, export: str, xml: Optional[str]) -> None:
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            xml_ptr = scope.cstr(xml or "")
            _check(getattr(LIB, export)(h, _uint(sheet, "sheet_index", role="sheet"), xml_ptr))

    def get_print_area(self) -> Optional[List[CellRect]]:
        """Return the print area's rectangles, or ``None`` when none is declared.

        A whole-axis area reads back expanded to the grid edge (``A:D``
        becomes ``A1:D1048576``), because this reports the rectangles the
        paginator resolves rather than the authored shape.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            count_out = scope.out_ptr()
            # Size the buffer first, then fetch.
            _check(LIB.fm_sheet_get_print_area(h, _uint(sheet, "sheet_index", role="sheet"), 0, 0, count_out))
            total = LIB.read_u32(count_out)
            if total == 0:
                return None
            ranges_ptr = scope.struct_array(S.CELL_RECT, total)
            _check(
                LIB.fm_sheet_get_print_area(
                    h, _uint(sheet, "sheet_index", role="sheet"), ranges_ptr, _uint(total, "capacity"), count_out
                )
            )
            return _decode_cell_rect_array(ranges_ptr, total)

    def set_print_area(self, ranges: Union[Sequence[Union[CellRect, str]], str, None]) -> None:
        """Write ``_xlnm.Print_Area`` from one or more rectangles, printed in order.

        A rectangle spanning every row or column is stored as whole columns
        or rows; one A1 range string is one rectangle. ``None`` or an empty
        list removes the name.
        """
        rects = [] if ranges is None else [self._rect(r) for r in ([ranges] if isinstance(ranges, str) else ranges)]
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ranges_ptr = _pack_cell_rect_array(rects, scope)
            _check(
                LIB.fm_sheet_set_print_area(
                    h, _uint(sheet, "sheet_index", role="sheet"), ranges_ptr, _uint(len(rects), "range_count")
                )
            )

    def get_print_titles(self) -> Optional[PrintTitles]:
        """Return the rows and columns repeated on every printed page, or ``None`` when neither repeats."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PRINT_TITLES)
            _check(LIB.fm_sheet_get_print_titles(h, _uint(sheet, "sheet_index", role="sheet"), ptr))
            titles = _print_titles_from(S.PRINT_TITLES.unpack(LIB, ptr))
            return None if titles.rows is None and titles.columns is None else titles

    def set_print_titles(self, titles: Optional[Union[PrintTitlesInput, PrintTitles]]) -> None:
        """Write the sheet's ``_xlnm.Print_Titles`` defined name.

        ``titles.rows`` and ``titles.columns`` are each a :class:`Span`, a span
        string (``"1:2"`` or ``"3"`` for rows, ``"A:B"`` or ``"B"`` for columns)
        or ``None``. The stored formula lists rows before columns and anchors
        every endpoint (``Sheet1!$1:$2,Sheet1!$A:$A``). ``None``, or titles with
        both axes ``None``, removes the name. A :class:`PrintTitles` read back is
        accepted in place of a :class:`PrintTitlesInput`.

        Raises:
          TypeError: when ``titles`` is not one of the two record types.
          InvalidArgumentError: for a span string that is not a whole-row or
            whole-column span of its axis.
        """
        if titles is not None and not isinstance(titles, (PrintTitlesInput, PrintTitles)):
            raise TypeError(_argument_message("titles", "must be a PrintTitlesInput"))
        sheet = self._index()
        h = self._workbook._require()
        if titles is None:
            _check(LIB.fm_sheet_set_print_titles(h, _uint(sheet, "sheet_index", role="sheet"), 0))
            return
        values: Dict[str, int] = {}
        for axis, kind, (flag, first, last) in (
            (titles.rows, "row", ("has_rows", "first_row", "last_row")),
            (titles.columns, "col", ("has_columns", "first_col", "last_col")),
        ):
            if isinstance(axis, str):
                axis = Span(*self._span(axis, kind))
            if axis is not None:
                values[flag] = 1
                values[first] = _uint(axis.first, first, role="coordinate")
                values[last] = _uint(axis.last, last, role="coordinate")
        with _ArgScope() as scope:
            ptr = scope.struct(S.PRINT_TITLES)
            S.PRINT_TITLES.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_set_print_titles(h, _uint(sheet, "sheet_index", role="sheet"), ptr))

    def set_row_break(self, row: int, input: Optional[Union[PageBreakInput, RowBreak]]) -> None:
        """Set the row break before ``row``, replacing the span and ``manual`` flag of one already there.

        ``input`` is a :class:`PageBreakInput`, or a :class:`RowBreak` /
        :class:`ColumnBreak` (only its ``first``, ``last`` and ``manual`` are
        read, ``row`` / ``col`` is ignored, so a break read back can be passed in). ``first == last == 0`` covers the
        whole sheet. ``None`` removes the break (a no-op when there is none).

        Raises:
          TypeError: when ``input`` is neither a break record nor ``None``, or ``manual`` is not a ``bool``.
          InvalidArgumentError: for a row past the grid, an inverted span or a
            span past the last column.
          FormulonError: with ``code == StatusCode.PRECONDITION_FAILED`` when the axis already holds 1026 breaks.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = self._pack_break(scope, S.ROW_BREAK, input)
            _check(
                LIB.fm_sheet_set_row_break(
                    h, _uint(sheet, "sheet_index", role="sheet"), _uint(row, "row", role="coordinate"), ptr
                )
            )

    def set_column_break(self, col: int, input: Optional[Union[PageBreakInput, ColumnBreak]]) -> None:
        """Set or remove the column break before ``col``, as :meth:`set_row_break` does for rows.

        The whole-sheet span (``first == last == 0``) covers every row.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = self._pack_break(scope, S.COLUMN_BREAK, input)
            _check(
                LIB.fm_sheet_set_column_break(
                    h, _uint(sheet, "sheet_index", role="sheet"), _uint(col, "col", role="coordinate"), ptr
                )
            )

    @staticmethod
    def _pack_break(scope: _ArgScope, layout: S.Struct, brk: Union[PageBreakInput, RowBreak, ColumnBreak, None]) -> int:
        """The break struct ``brk`` describes in ``layout``; the null pointer for ``None``."""
        if brk is None:
            return 0
        if not isinstance(brk, (PageBreakInput, RowBreak, ColumnBreak)):
            raise TypeError(_argument_message("input", "must be a PageBreakInput, a RowBreak, a ColumnBreak or null"))
        ptr = scope.struct(layout)
        layout.pack(
            LIB,
            ptr,
            {
                "first": _uint(brk.first, "first", role="coordinate"),
                "last": _uint(brk.last, "last", role="coordinate"),
                "manual": _tristate_default(_require_bool(brk.manual, "input.manual")),
            },
        )
        return ptr

    def clear_breaks(self) -> None:
        """Remove every manual break on both axes."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_breaks(h, _uint(sheet, "sheet_index", role="sheet")))

    def _read_breaks(self, count_fn, at_fn, layout: S.Struct, record, key: str) -> list:
        sheet = self._index()
        h = self._workbook._require()
        count = _read_count(count_fn, h, _uint(sheet, "sheet_index", role="sheet"))
        out = []
        with _ArgScope() as scope:
            ptr = scope.struct(layout)
            for i in range(count):
                _check(at_fn(h, _uint(sheet, "sheet_index", role="sheet"), _uint(i, "index", role="position"), ptr))
                d = layout.unpack(LIB, ptr)
                manual = d["manual"] == S.TRISTATE_TRUE
                out.append(record(**{key: d[key]}, first=d["first"], last=d["last"], manual=manual))
        return out

    def list_row_breaks(self) -> List[RowBreak]:
        """Return the row breaks, ascending by row."""
        return self._read_breaks(
            LIB.fm_sheet_get_row_break_count, LIB.fm_sheet_get_row_break_at, S.ROW_BREAK, RowBreak, "row"
        )

    def list_column_breaks(self) -> List[ColumnBreak]:
        """Return the column breaks, ascending by column."""
        return self._read_breaks(
            LIB.fm_sheet_get_column_break_count, LIB.fm_sheet_get_column_break_at, S.COLUMN_BREAK, ColumnBreak, "col"
        )

    def update_page_setup(
        self,
        *,
        orientation: Union[PageOrientation, None, Unchanged] = UNCHANGED,
        paper_size: Union[int, Unchanged] = UNCHANGED,
        scale_pct: Union[int, Unchanged] = UNCHANGED,
        fit_to_width: Union[int, Unchanged] = UNCHANGED,
        fit_to_height: Union[int, Unchanged] = UNCHANGED,
        fit_to_page: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update ``<pageSetup>``; an omitted argument keeps its attribute.

        ``paper_size`` is the OOXML code (9 = A4). ``scale_pct`` (a percentage) outside
        ``[10, 400]`` is rejected rather than clamped -- a mis-stated print
        scale lands on paper, so rounding it silently would hide the error.
        ``fit_to_page`` sets or clears ``<sheetPr><pageSetUpPr fitToPage>``
        without disturbing anything else ``<sheetPr>`` carries; ``False``
        removes the attribute.
        ``orientation=None`` removes the stated orientation; the sheet then
        prints portrait.

        Raises:
          TypeError: when a non-nullable argument is ``None``.
        """
        values: Dict[str, object] = {}
        for name, value in (
            ("orientation", orientation),
            ("paper_size", paper_size),
            ("scale_pct", scale_pct),
            ("fit_to_width", fit_to_width),
            ("fit_to_height", fit_to_height),
            ("fit_to_page", fit_to_page),
        ):
            if value is UNCHANGED:
                continue
            if value is None:
                if name != "orientation":
                    raise TypeError(_argument_message(name, "must not be null"))
                values[name + "_op"] = _PATCH_CLEAR
                continue
            values[name + "_op"] = _PATCH_SET
            if name == "orientation":
                values[name] = _sint(value, name)
            elif name == "fit_to_page":
                values[name] = int(_require_bool(value, name))
            else:
                values[name] = _uint(value, name)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PAGE_SETUP_PATCH)
            S.PAGE_SETUP_PATCH.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_update_page_setup(h, _uint(sheet, "sheet_index", role="sheet"), ptr))

    def update_page_margins(
        self,
        *,
        left_in: Union[float, Unchanged] = UNCHANGED,
        right_in: Union[float, Unchanged] = UNCHANGED,
        top_in: Union[float, Unchanged] = UNCHANGED,
        bottom_in: Union[float, Unchanged] = UNCHANGED,
        header_in: Union[float, Unchanged] = UNCHANGED,
        footer_in: Union[float, Unchanged] = UNCHANGED,
    ) -> None:
        """Update ``<pageMargins>``, in inches; an omitted argument keeps its margin.

        A negative, infinite or NaN margin is rejected: the paginator
        subtracts these from the paper, so such a value does not describe a
        printable body.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PAGE_MARGINS_PATCH)
            values: Dict[str, object] = {}
            for name, value in (
                ("left_in", left_in),
                ("right_in", right_in),
                ("top_in", top_in),
                ("bottom_in", bottom_in),
                ("header_in", header_in),
                ("footer_in", footer_in),
            ):
                if value is UNCHANGED:
                    continue
                if value is None:
                    raise TypeError(_argument_message(name, "must not be null"))
                values[f"{name}_op"] = _PATCH_SET
                values[name] = float(value)
            S.PAGE_MARGINS_PATCH.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_update_page_margins(h, _uint(sheet, "sheet_index", role="sheet"), ptr))

    def update_print_options(
        self,
        *,
        grid_lines: Union[bool, Unchanged] = UNCHANGED,
        headings: Union[bool, Unchanged] = UNCHANGED,
        horizontal_centered: Union[bool, Unchanged] = UNCHANGED,
        vertical_centered: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update ``<printOptions>``; an omitted argument keeps its attribute.

        Raises:
          TypeError: when an argument is not a ``bool`` (``None`` included).
        """
        values: Dict[str, object] = {}
        for name, value in (
            ("grid_lines", grid_lines),
            ("headings", headings),
            ("horizontal_centered", horizontal_centered),
            ("vertical_centered", vertical_centered),
        ):
            if value is UNCHANGED:
                continue
            values[f"{name}_op"] = _PATCH_SET
            values[name] = int(_require_bool(value, name))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PRINT_OPTIONS_PATCH)
            S.PRINT_OPTIONS_PATCH.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_update_print_options(h, _uint(sheet, "sheet_index", role="sheet"), ptr))

    def update_header_footer(
        self,
        *,
        odd_header: Union[str, None, Unchanged] = UNCHANGED,
        odd_footer: Union[str, None, Unchanged] = UNCHANGED,
        even_header: Union[str, None, Unchanged] = UNCHANGED,
        even_footer: Union[str, None, Unchanged] = UNCHANGED,
        first_header: Union[str, None, Unchanged] = UNCHANGED,
        first_footer: Union[str, None, Unchanged] = UNCHANGED,
        different_odd_even: Union[bool, Unchanged] = UNCHANGED,
        different_first: Union[bool, Unchanged] = UNCHANGED,
        scale_with_doc: Union[bool, Unchanged] = UNCHANGED,
        align_with_margins: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update ``<headerFooter>``; an omitted argument keeps its value.

        A section given as ``None`` or ``""`` is cleared. Pass decoded text --
        Excel's formatting codes go in plainly
        (``"&C&\\"MS Gothic\\"Report &P/&N"``) and the engine applies the XML
        escaping.

        Raises:
          TypeError: when a flag is not a ``bool`` (``None`` included) or a
            section is neither a ``str`` nor ``None``.
        """
        values: Dict[str, object] = {}
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            for name, value in (
                ("different_odd_even", different_odd_even),
                ("different_first", different_first),
                ("scale_with_doc", scale_with_doc),
                ("align_with_margins", align_with_margins),
            ):
                if value is UNCHANGED:
                    continue
                values[f"{name}_op"] = _PATCH_SET
                values[name] = int(_require_bool(value, name))
            for name, text in (
                ("odd_header", odd_header),
                ("odd_footer", odd_footer),
                ("even_header", even_header),
                ("even_footer", even_footer),
                ("first_header", first_header),
                ("first_footer", first_footer),
            ):
                if text is UNCHANGED:
                    continue
                if text is None:
                    values[f"{name}_op"] = _PATCH_CLEAR
                    continue
                if not isinstance(text, str):
                    raise TypeError(_argument_message(name, "must be a string or null"))
                values[f"{name}_op"] = _PATCH_SET
                values[name] = scope.cstr(text)
            ptr = scope.struct(S.HEADER_FOOTER_PATCH)
            S.HEADER_FOOTER_PATCH.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_update_header_footer(h, _uint(sheet, "sheet_index", role="sheet"), ptr))

    def get_page_setup(self) -> PageSetup:
        """Read the effective page setup and which attributes state it."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PAGE_SETUP)
            _check(LIB.fm_sheet_get_page_setup(h, _uint(sheet, "sheet_index", role="sheet"), ptr))
            d = S.PAGE_SETUP.unpack(LIB, ptr)
            return PageSetup(
                orientation=PageOrientation(d["orientation"]) if d["has_orientation"] else None,
                paper_size=d["paper_size"],
                scale_pct=d["scale_pct"],
                fit_to_width=d["fit_to_width"],
                fit_to_height=d["fit_to_height"],
                fit_to_page=bool(d["fit_to_page"]),
                paper_size_stated=bool(d["has_paper_size"]),
                scale_pct_stated=bool(d["has_scale_pct"]),
                fit_to_width_stated=bool(d["has_fit_to_width"]),
                fit_to_height_stated=bool(d["has_fit_to_height"]),
                fit_to_page_stated=bool(d["has_fit_to_page"]),
            )

    def get_page_margins(self) -> PageMargins:
        """Read the effective page margins in inches, with presence flags."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PAGE_MARGINS)
            _check(LIB.fm_sheet_get_page_margins(h, _uint(sheet, "sheet_index", role="sheet"), ptr))
            d = S.PAGE_MARGINS.unpack(LIB, ptr)
            return PageMargins(
                left_in=d["left_in"],
                right_in=d["right_in"],
                top_in=d["top_in"],
                bottom_in=d["bottom_in"],
                header_in=d["header_in"],
                footer_in=d["footer_in"],
                left_in_stated=bool(d["has_left_in"]),
                right_in_stated=bool(d["has_right_in"]),
                top_in_stated=bool(d["has_top_in"]),
                bottom_in_stated=bool(d["has_bottom_in"]),
                header_in_stated=bool(d["has_header_in"]),
                footer_in_stated=bool(d["has_footer_in"]),
            )

    def get_print_options(self) -> PrintOptions:
        """Read the ``<printOptions>`` flags and which of them the XML states."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PRINT_OPTIONS)
            _check(LIB.fm_sheet_get_print_options(h, _uint(sheet, "sheet_index", role="sheet"), ptr))
            d = S.PRINT_OPTIONS.unpack(LIB, ptr)
        names = ("grid_lines", "headings", "horizontal_centered", "vertical_centered")
        fields: Dict[str, bool] = {name: bool(d[name]) for name in names}
        fields.update({f"{name}_stated": bool(d[f"has_{name}"]) for name in names})
        return PrintOptions(**fields)

    def get_header_footer(self) -> HeaderFooter:
        """Read the ``<headerFooter>`` sections and flags, and which flags the XML states.

        A section the sheet does not carry reads ``None``; a present one is
        its decoded text, as :meth:`update_header_footer` takes it.
        """
        sheet = self._index()
        h = self._workbook._require()
        sections = ("odd_header", "odd_footer", "even_header", "even_footer", "first_header", "first_footer")
        flags = ("different_odd_even", "different_first", "scale_with_doc", "align_with_margins")
        with _ArgScope() as scope:
            ptr = scope.struct(S.HEADER_FOOTER)
            _check(LIB.fm_sheet_get_header_footer(h, _uint(sheet, "sheet_index", role="sheet"), ptr))
            d = S.HEADER_FOOTER.unpack(LIB, ptr)
            fields: Dict[str, object] = {name: LIB.read_cstr(d[name]) if d[name] else None for name in sections}
        fields.update({name: bool(d[name]) for name in flags})
        fields.update({f"{name}_stated": bool(d[f"has_{name}"]) for name in flags})
        return HeaderFooter(**fields)

    def list_column_overrides(self) -> List[ColumnLayout]:
        """Return the column-layout overrides on the sheet."""
        return [
            ColumnLayout(
                first=d["first"],
                last=d["last"],
                width_chars=d["width_chars"] if d["has_width_chars"] else None,
                hidden=bool(d["hidden"]),
                outline_level=d["outline_level"],
                collapsed=bool(d["collapsed"]),
            )
            for d in self._layout_overrides(
                LIB.fm_sheet_get_column_layout_count, LIB.fm_sheet_get_column_layout_at, S.COLUMN_LAYOUT
            )
        ]

    def _layout_overrides(self, count_fn: Any, at_fn: Any, layout: S.Struct) -> List[Dict[str, Any]]:
        """The decoded column or row layout structs of the sheet, in order."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(count_fn, h, _uint(sheet, "sheet_index", role="sheet"))
        out: List[Dict[str, Any]] = []
        # One scratch block for the whole walk; see list_defined_names.
        with _ArgScope() as scope:
            ptr = scope.struct(layout)
            for i in range(n):
                S.zero_struct(LIB, layout, ptr)
                _check(at_fn(h, _uint(sheet, "sheet_index", role="sheet"), _uint(i, "idx", role="position"), ptr))
                out.append(layout.unpack(LIB, ptr))
        return out

    @overload
    @_span_form
    def set_column_width(self, first: int, last: int, width_chars: float) -> None: ...

    @overload
    def set_column_width(self, span: str, width_chars: float) -> None: ...

    def set_column_width(self, *args: Any, **kwargs: Any) -> None:
        """Set the column width override on ``[first, last]``.

        ``width_chars`` is a stored width: characters with the column padding
        included, as ``<col width>`` holds it. A width above 255 is rejected
        with ``InvalidArgumentError``.
        """
        first, last, width_chars = self._span_args("set_column_width", args, kwargs, ("first", "last", "width_chars"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_column_width(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
                float(width_chars),
            )
        )

    @overload
    @_span_form
    def set_column_hidden(self, first: int, last: int, hidden: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    @overload
    def set_column_hidden(self, span: str, hidden: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    def set_column_hidden(self, *args: Any, **kwargs: Any) -> None:
        """Set the column hidden flag on ``[first, last]``."""
        first, last, hidden = self._span_args("set_column_hidden", args, kwargs, ("first", "last", "hidden"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_column_hidden(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
                int(_require_bool(hidden, "hidden")),
            )
        )

    @overload
    @_span_form
    def set_column_outline(self, first: int, last: int, level: int) -> None: ...

    @overload
    def set_column_outline(self, span: str, level: int) -> None: ...

    def set_column_outline(self, *args: Any, **kwargs: Any) -> None:
        """Set the column outline level (0-7) on ``[first, last]``.

        Raises:
          FormulonError: with ``StatusCode.INVALID_ARGUMENT`` when ``level`` exceeds 7.
        """
        first, last, level = self._span_args("set_column_outline", args, kwargs, ("first", "last", "level"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_column_outline(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
                _uint(level, "level", 32),
            )
        )

    @overload
    @_span_form
    def set_column_collapsed(self, first: int, last: int, collapsed: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    @overload
    def set_column_collapsed(self, span: str, collapsed: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    def set_column_collapsed(self, *args: Any, **kwargs: Any) -> None:
        """Set or clear the ``collapsed`` flag on the columns ``[first, last]``."""
        first, last, collapsed = self._span_args("set_column_collapsed", args, kwargs, ("first", "last", "collapsed"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_column_collapsed(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
                int(_require_bool(collapsed, "collapsed")),
            )
        )

    def list_row_overrides(self) -> List[RowLayout]:
        """Return the row-layout overrides on the sheet."""
        return [
            RowLayout(
                row=d["row"],
                height_pt=d["height_pt"] if d["has_height_pt"] else None,
                hidden=bool(d["hidden"]),
                outline_level=d["outline_level"],
                collapsed=bool(d["collapsed"]),
                custom_height=bool(d["custom_height"]),
            )
            for d in self._layout_overrides(
                LIB.fm_sheet_get_row_layout_count, LIB.fm_sheet_get_row_layout_at, S.ROW_LAYOUT
            )
        ]

    @overload
    @_span_form
    def set_row_height(self, first: int, last: int, height_pt: float) -> None: ...

    @overload
    def set_row_height(self, span: str, height_pt: float) -> None: ...

    def set_row_height(self, *args: Any, **kwargs: Any) -> None:
        """Set the row height, in points, on ``[first, last]`` as a manual height (``customHeight``). A height above 409 is rejected with ``InvalidArgumentError``."""
        first, last, height_pt = self._span_args("set_row_height", args, kwargs, ("first", "last", "height_pt"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_height(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
                float(height_pt),
            )
        )

    @overload
    @_span_form
    def set_row_auto_height(self, first: int, last: int, height_pt: float) -> None: ...

    @overload
    def set_row_auto_height(self, span: str, height_pt: float) -> None: ...

    def set_row_auto_height(self, *args: Any, **kwargs: Any) -> None:
        """Store ``height_pt`` (points) as the cached automatic height of the rows on ``[first, last]``.

        Saved as ``ht`` without ``customHeight``, so each row keeps following its
        content and font; the rows' other attributes are kept.
        """
        first, last, height_pt = self._span_args("set_row_auto_height", args, kwargs, ("first", "last", "height_pt"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_auto_height(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
                float(height_pt),
            )
        )

    @overload
    @_span_form
    def set_row_hidden(self, first: int, last: int, hidden: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    @overload
    def set_row_hidden(self, span: str, hidden: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    def set_row_hidden(self, *args: Any, **kwargs: Any) -> None:
        """Set the row hidden flag on ``[first, last]``."""
        first, last, hidden = self._span_args("set_row_hidden", args, kwargs, ("first", "last", "hidden"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_hidden(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
                int(_require_bool(hidden, "hidden")),
            )
        )

    @overload
    @_span_form
    def set_row_outline(self, first: int, last: int, level: int) -> None: ...

    @overload
    def set_row_outline(self, span: str, level: int) -> None: ...

    def set_row_outline(self, *args: Any, **kwargs: Any) -> None:
        """Set the row outline level (0-7) on ``[first, last]``.

        Raises:
          FormulonError: with ``StatusCode.INVALID_ARGUMENT`` when ``level`` exceeds 7.
        """
        first, last, level = self._span_args("set_row_outline", args, kwargs, ("first", "last", "level"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_outline(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
                _uint(level, "level", 32),
            )
        )

    @overload
    @_span_form
    def set_row_collapsed(self, first: int, last: int, collapsed: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    @overload
    def set_row_collapsed(self, span: str, collapsed: bool) -> None: ...  # noqa: FBT001 -- a single-boolean setter takes its flag positionally

    def set_row_collapsed(self, *args: Any, **kwargs: Any) -> None:
        """Set or clear the ``collapsed`` flag on ``[first, last]``."""
        first, last, collapsed = self._span_args("set_row_collapsed", args, kwargs, ("first", "last", "collapsed"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_collapsed(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
                int(_require_bool(collapsed, "collapsed")),
            )
        )

    # -- Conditional formatting --------------------------------------------
    def evaluate_cf_range(
        self,
        range: Union[CellRect, str],
        *,
        today: Optional[Union[datetime.datetime, datetime.date]] = None,
        time_period_rules: bool = True,
    ) -> List[CfCell]:
        """Evaluate every CF block on the sheet against the cells of ``range``.

        Args:
          range: the cells to evaluate.
          today: date basis for ``timePeriod`` rules, a ``date`` or
            ``datetime`` (a naive local wall clock; an aware one is a
            ValueError; only the day counts). Omitted, it is today: the pinned
            date when the workbook has one, else the host's local date.
          time_period_rules: ``False`` evaluates without a date, so no
            ``timePeriod`` rule matches; ``today`` is then ignored.

        Raises:
          TypeError: when ``today`` or ``time_period_rules`` is passed positionally.
        """
        today_serial: Optional[float] = None
        if time_period_rules and today is not None:
            day = civil_datetime(today, "today").date()
            today_serial = date_serial(day, date1904=self._workbook.get_date1904())[0]
        return self._evaluate_cf_rect(self._rect(range), today_serial, time_period_rules=time_period_rules)

    def _evaluate_cf_rect(
        self, rect: CellRect, today_serial: Optional[float], *, time_period_rules: bool = True
    ) -> List[CfCell]:
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rect_ptr = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rect_ptr, rect)
            options = scope.struct(S.CF_RANGE_EVALUATION_SPEC)
            S.CF_RANGE_EVALUATION_SPEC.pack(
                LIB,
                options,
                {
                    "time_period_rules": 0 if time_period_rules else 1,
                    "has_today_serial": 0 if today_serial is None else 1,
                    "today_serial": 0.0 if today_serial is None else float(today_serial),
                },
            )
            out = scope.out_ptr()
            results = 0
            try:
                _check(
                    LIB.fm_sheet_evaluate_cf_range(
                        h,
                        _uint(sheet, "sheet_index", role="sheet"),
                        rect_ptr,
                        options,
                        out,
                    )
                )
                results = LIB.read_u32(out)
                cells: List[CfCell] = []
                n = _read_count(LIB.fm_cf_results_get_cell_count, results)
                # One fixed set of out slots for the whole walk: allocating
                # per cell would put a locked WASM malloc/free pair on every
                # item, which dominates the cost of a viewport-sized range.
                rp = scope.out_ptr()
                cp = scope.out_ptr()
                mcp = scope.out_ptr()
                mptr = scope.struct(S.CF_MATCH)
                for ci in range(n):
                    _check(
                        LIB.fm_cf_results_get_cell_at(results, _uint(ci, "cell_index", role="position"), rp, cp, mcp)
                    )
                    row = LIB.read_u32(rp)
                    col = LIB.read_u32(cp)
                    mcount = LIB.read_u32(mcp)
                    matches: List[CfMatch] = []
                    for mi in range(mcount):
                        S.zero_struct(LIB, S.CF_MATCH, mptr)
                        _check(
                            LIB.fm_cf_results_get_match_at(
                                results,
                                _uint(ci, "cell_index", role="position"),
                                _uint(mi, "match_index", role="position"),
                                mptr,
                            )
                        )
                        d = S.CF_MATCH.unpack(LIB, mptr)
                        matches.append(_cf_match_from(d))
                    cells.append(CfCell(row=row, col=col, matches=matches))
                return cells
            finally:
                if results:
                    LIB.fm_cf_results_destroy(results)

    def list_cf_rules(self) -> List[ConditionalFormat]:
        """Every conditional-format rule of the sheet, in priority order."""
        return [self._cf_rule_at(i) for i in range(self._cf_rule_count())]

    def _cf_rule_count(self) -> int:
        """Return the number of CF rules on the sheet (flattened)."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(
            LIB.fm_sheet_get_cf_rule_count,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
        )

    @staticmethod
    def _decode_cfvo(ptr: int) -> CfValueObject:
        d = S.CFVO.unpack(LIB, ptr)
        return CfValueObject(
            type=d["type"], value=LIB.read_cstr(d["value"]) if d["value"] else None, gte=d["gte"] == S.TRISTATE_TRUE
        )

    @staticmethod
    def _write_cfvo(ptr: int, value: CfValueObject, scope: _ArgScope) -> None:
        S.CFVO.pack(LIB, ptr, {"type": _uint(value.type, "type", 8), "gte": _tristate_default(value.gte)})
        if value.value is not None:
            S.write_str_field(LIB, ptr, S.CFVO, "value", value.value, scope)

    def _cf_rule_at(self, cf_rule_index: int) -> ConditionalFormat:
        """Read the ``cf_rule_index``-th CF rule on the sheet (flattened order)."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CF_RULE)
            _check(
                LIB.fm_sheet_get_cf_rule_at(
                    h, _uint(sheet, "sheet_index", role="sheet"), _uint(cf_rule_index, "idx", role="position"), ptr
                )
            )
            return self._decode_cf_rule(ptr)

    def get_cf_rule(self, cf_rule: CfRuleRef) -> Optional[ConditionalFormat]:
        """Return ``cf_rule`` (an id or a position in :meth:`list_cf_rules`), or ``None`` when absent.

        Raises:
          IndexOutOfRangeError: when ``cf_rule`` is a position past the last rule.
        """
        cf_rule_id = self._cf_rule_id(cf_rule)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.CF_RULE)
            _check(
                LIB.fm_sheet_get_cf_rule(
                    h, _uint(sheet, "sheet_index", role="sheet"), scope.cstr(cf_rule_id), present, ptr
                )
            )
            if not LIB.read_u32(present):
                return None
            return self._decode_cf_rule(ptr)

    def _decode_cf_rule(self, ptr: int) -> ConditionalFormat:
        """Decode the rule record at ``ptr``."""
        d = S.CF_RULE.unpack(LIB, ptr)
        ranges = _decode_cell_rect_array(d["ranges"], d["range_count"])
        color_scale = None
        if d["color_scale_count"]:
            thresholds = [
                self._decode_cfvo(d["color_scale_thresholds"] + i * S.CFVO.size) for i in range(d["color_scale_count"])
            ]
            colors = [
                _decode_color(d["color_scale_colors"] + i * S.COLOR_SPEC.size) for i in range(d["color_scale_count"])
            ]
            color_scale = CfColorScale(thresholds, colors)
        offsets = S.CF_RULE.offsets
        data_bar = None
        if d["has_data_bar"]:
            data_bar = CfDataBar(
                self._decode_cfvo(ptr + offsets["data_bar_min"][1]),
                self._decode_cfvo(ptr + offsets["data_bar_max"][1]),
                _decode_color(ptr + offsets["data_bar_fill"][1]),
                bool(d["data_bar_show_value"]),
                d["data_bar_min_length_pct"],
                d["data_bar_max_length_pct"],
                # x14 extension. Each field is decoded only when the C
                # ABI reports its `has_*` flag, so `None` keeps its
                # "no override" meaning on the way back in.
                bool(d["data_bar_gradient"]) if d["has_data_bar_gradient"] else None,
                _enum(CfDataBarAxisPosition, d["data_bar_axis_position"]) if d["has_data_bar_axis_position"] else None,
                _decode_optional_color(ptr + offsets["data_bar_negative_fill"][1]),
                _decode_optional_color(ptr + offsets["data_bar_border"][1]),
                _decode_optional_color(ptr + offsets["data_bar_negative_border"][1]),
                _decode_optional_color(ptr + offsets["data_bar_axis_color"][1]),
                _enum(CfDataBarDirection, d["data_bar_direction"]),
            )
        icon_set = None
        if d["has_icon_set"]:
            icon_set = CfIconSet(
                d["icon_set"],
                [
                    self._decode_cfvo(d["icon_set_thresholds"] + i * S.CFVO.size)
                    for i in range(d["icon_set_threshold_count"])
                ],
                bool(d["icon_set_reverse"]),
                bool(d["icon_set_show_value"]),
                bool(d["icon_set_percent"]),
                self._decode_cfvo(ptr + offsets["icon_set_floor"][1]) if d["has_icon_set_floor"] else None,
                [
                    CfIcon(
                        LIB.read_i32(d["icon_set_icons"] + i * S.CF_ICON.size),
                        LIB.read_i32(d["icon_set_icons"] + i * S.CF_ICON.size + S.CF_ICON.offsets["icon_index"][1]),
                    )
                    for i in range(d["icon_set_icon_count"])
                ],
            )
        return ConditionalFormat(
            id=LIB.read_cstr(d["id"]),
            type=d["type"],
            priority=d["priority"],
            stop_if_true=bool(d["stop_if_true"]),
            ranges=ranges,
            dxf_index=d["dxf_index"] if d["has_dxf_index"] else None,
            formula1=LIB.read_optional_cstr(d["formula1"]),
            formula2=LIB.read_optional_cstr(d["formula2"]),
            operator=d["op"] if d["has_op"] else None,
            rank=d["rank"] if d["has_rank"] else None,
            percent=bool(d["percent"]),
            bottom=bool(d["bottom"]),
            above_average=bool(d["above_average"]),
            equal_average=bool(d["equal_average"]),
            std_dev=d["std_dev"] if d["has_std_dev"] else None,
            text=LIB.read_optional_cstr(d["text"]),
            time_period=d["time_period"] if d["has_time_period"] else None,
            color_scale=color_scale,
            data_bar=data_bar,
            icon_set=icon_set,
        )

    def add_cf_rule(self, rule: Union[ConditionalFormatInput, ConditionalFormat]) -> CfRuleId:
        """Append a CF rule to the sheet; return its :data:`CfRuleId`.

        ``rule.id`` is a brace-wrapped GUID (``{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}``, hex digits in either
        case); omitted, empty or already used on the sheet, the engine generates an upper-case one; anything
        else is ``InvalidArgumentError``.

        ``rule`` is the input dataclass of its type (:class:`CfCellIsRuleInput`, :class:`CfColorScaleRuleInput`
        and the others of :data:`ConditionalFormatInput`); a :class:`ConditionalFormat` read back is accepted
        too. A color scale, data bar or icon set rule carries its payload in ``rule.color_scale``,
        ``rule.data_bar`` or ``rule.icon_set``. ``rule.ranges`` takes :class:`CellRect` values, A1 ranges or
        one A1 range string.
        """
        sheet = self._index()
        ranges = _cell_rects(self._workbook, rule.ranges, "rule.ranges")
        operator = getattr(rule, "operator", None)
        time_period = getattr(rule, "time_period", None)
        dxf_index = getattr(rule, "dxf_index", None)
        rank = getattr(rule, "rank", None)
        std_dev = getattr(rule, "std_dev", None)
        color_scale = getattr(rule, "color_scale", None)
        data_bar = getattr(rule, "data_bar", None)
        icon_set = getattr(rule, "icon_set", None)
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CF_RULE)
            out = scope.out_ptr()
            S.CF_RULE.pack(
                LIB,
                ptr,
                {
                    "type": _uint(rule.type, "type", 8),
                    "op": _uint(operator or 0, "operator", 8),
                    "time_period": _uint(time_period or 0, "time_period", 8),
                    "priority": _sint(rule.priority, "priority"),
                    "stop_if_true": int(_require_bool(rule.stop_if_true, "rule.stop_if_true")),
                    "has_dxf_index": 1 if dxf_index is not None else 0,
                    "dxf_index": _uint(dxf_index or 0, "dxf_index", role="position"),
                    "has_op": 1 if operator is not None else 0,
                    "has_rank": 1 if rank is not None else 0,
                    "rank": _sint(rank or 0, "rank"),
                    "percent": int(_require_bool(getattr(rule, "percent", False), "rule.percent")),
                    "bottom": int(_require_bool(getattr(rule, "bottom", False), "rule.bottom")),
                    "above_average": int(_require_bool(getattr(rule, "above_average", False), "rule.above_average")),
                    "equal_average": int(_require_bool(getattr(rule, "equal_average", False), "rule.equal_average")),
                    "has_std_dev": 1 if std_dev is not None else 0,
                    "std_dev": float(std_dev or 0.0),
                    "has_time_period": 1 if time_period is not None else 0,
                },
            )
            ranges_ptr = _pack_cell_rect_array(ranges, scope)
            ro = S.CF_RULE.offsets
            LIB.write_bytes(ptr + ro["ranges"][1], struct.pack("<I", ranges_ptr))
            LIB.write_bytes(ptr + ro["range_count"][1], struct.pack("<I", len(ranges)))
            S.write_str_field(LIB, ptr, S.CF_RULE, "id", rule.id, scope)
            S.write_str_field(LIB, ptr, S.CF_RULE, "formula1", getattr(rule, "formula1", None), scope)
            S.write_str_field(LIB, ptr, S.CF_RULE, "formula2", getattr(rule, "formula2", None), scope)
            S.write_str_field(LIB, ptr, S.CF_RULE, "text", getattr(rule, "text", None), scope)
            if color_scale is not None:
                count = len(color_scale.thresholds)
                if count != len(color_scale.colors):
                    raise ValueError("color_scale thresholds and colors must have equal length")
                thresholds = scope.alloc(count * S.CFVO.size)
                colors = scope.alloc(count * S.COLOR_SPEC.size)
                for i, value in enumerate(color_scale.thresholds):
                    self._write_cfvo(thresholds + i * S.CFVO.size, value, scope)
                for i, color in enumerate(color_scale.colors):
                    _pack_color(colors + i * S.COLOR_SPEC.size, color, f"color_scale.colors[{i}]")
                LIB.write_bytes(ptr + ro["color_scale_thresholds"][1], struct.pack("<I", thresholds))
                LIB.write_bytes(ptr + ro["color_scale_colors"][1], struct.pack("<I", colors))
                LIB.write_bytes(ptr + ro["color_scale_count"][1], struct.pack("<I", count))
            if data_bar is not None:
                self._write_cfvo(ptr + ro["data_bar_min"][1], data_bar.min, scope)
                self._write_cfvo(ptr + ro["data_bar_max"][1], data_bar.max, scope)
                _pack_color(ptr + ro["data_bar_fill"][1], data_bar.fill, "data_bar.fill")
                LIB.write_bytes(ptr + ro["has_data_bar"][1], struct.pack("<i", 1))
                LIB.write_bytes(
                    ptr + ro["data_bar_show_value"][1],
                    struct.pack("<i", int(_require_bool(data_bar.show_value, "data_bar.show_value"))),
                )
                LIB.write_bytes(
                    ptr + ro["data_bar_min_length_pct"][1],
                    struct.pack("<B", _uint(data_bar.min_length_pct, "min_length_pct", 8)),
                )
                LIB.write_bytes(
                    ptr + ro["data_bar_max_length_pct"][1],
                    struct.pack("<B", _uint(data_bar.max_length_pct, "max_length_pct", 8)),
                )
                # x14 extension. A `None` leaves the field zero, which the C ABI
                # reads as "keep the model default".
                if data_bar.gradient is not None:
                    LIB.write_bytes(ptr + ro["has_data_bar_gradient"][1], struct.pack("<i", 1))
                    LIB.write_bytes(
                        ptr + ro["data_bar_gradient"][1],
                        struct.pack("<i", int(_require_bool(data_bar.gradient, "data_bar.gradient"))),
                    )
                if data_bar.axis_position is not None:
                    LIB.write_bytes(ptr + ro["has_data_bar_axis_position"][1], struct.pack("<i", 1))
                    LIB.write_bytes(
                        ptr + ro["data_bar_axis_position"][1],
                        struct.pack("<i", _sint(data_bar.axis_position, "axis_position")),
                    )
                for field, color in (
                    ("negative_fill", data_bar.negative_fill),
                    ("border", data_bar.border),
                    ("negative_border", data_bar.negative_border),
                    ("axis_color", data_bar.axis_color),
                ):
                    if color is not None:
                        _pack_color(ptr + ro[f"data_bar_{field}"][1], color, f"data_bar.{field}")
                LIB.write_bytes(
                    ptr + ro["data_bar_direction"][1], struct.pack("<i", _sint(data_bar.direction, "direction"))
                )
            if icon_set is not None:
                count = len(icon_set.thresholds)
                thresholds = scope.alloc(count * S.CFVO.size)
                for i, value in enumerate(icon_set.thresholds):
                    self._write_cfvo(thresholds + i * S.CFVO.size, value, scope)
                LIB.write_bytes(ptr + ro["icon_set_thresholds"][1], struct.pack("<I", thresholds))
                LIB.write_bytes(ptr + ro["has_icon_set"][1], struct.pack("<i", 1))
                LIB.write_bytes(ptr + ro["icon_set"][1], struct.pack("<i", _sint(icon_set.name, "icon_set")))
                LIB.write_bytes(ptr + ro["icon_set_threshold_count"][1], struct.pack("<I", count))
                LIB.write_bytes(
                    ptr + ro["icon_set_reverse"][1],
                    struct.pack("<i", int(_require_bool(icon_set.reverse, "icon_set.reverse"))),
                )
                LIB.write_bytes(
                    ptr + ro["icon_set_show_value"][1],
                    struct.pack("<i", int(_require_bool(icon_set.show_value, "icon_set.show_value"))),
                )
                LIB.write_bytes(
                    ptr + ro["icon_set_percent"][1],
                    struct.pack("<i", int(_require_bool(icon_set.percent, "icon_set.percent"))),
                )
                if icon_set.floor is not None:
                    LIB.write_bytes(ptr + ro["has_icon_set_floor"][1], struct.pack("<i", 1))
                    self._write_cfvo(ptr + ro["icon_set_floor"][1], icon_set.floor, scope)
                if icon_set.icons:
                    icons_ptr = scope.alloc(len(icon_set.icons) * S.CF_ICON.size)
                    for i, icon in enumerate(icon_set.icons):
                        S.CF_ICON.pack(
                            LIB,
                            icons_ptr + i * S.CF_ICON.size,
                            {
                                "icon_set": _sint(icon.icon_set, "icon_set"),
                                "icon_index": _sint(icon.icon_index, "icon_index"),
                            },
                        )
                    LIB.write_bytes(ptr + ro["icon_set_icons"][1], struct.pack("<I", icons_ptr))
                    LIB.write_bytes(ptr + ro["icon_set_icon_count"][1], struct.pack("<I", len(icon_set.icons)))
            _check(LIB.fm_sheet_add_cf_rule(h, _uint(sheet, "sheet_index", role="sheet"), ptr, out))
            return LIB.read_cstr(LIB.read_u32(out))

    def _cf_rule_id(self, cf_rule: CfRuleRef) -> CfRuleId:
        """The :data:`CfRuleId` ``cf_rule`` names: an id as it is, a position through its rule."""
        if isinstance(cf_rule, str):
            return cf_rule
        if isinstance(cf_rule, bool):
            raise TypeError(_argument_message("cf_rule", "must be an integer or a string"))
        return self._cf_rule_at(_uint(cf_rule, "cf_rule", role="position")).id

    def update_cf_rule(
        self,
        cf_rule: CfRuleRef,
        *,
        priority: Union[int, Unchanged] = UNCHANGED,
        stop_if_true: Union[bool, Unchanged] = UNCHANGED,
        dxf_index: Union[int, None, Unchanged] = UNCHANGED,
        ranges: Union[Sequence[Union[CellRect, str]], str, Unchanged] = UNCHANGED,
        formula1: Union[str, None, Unchanged] = UNCHANGED,
        formula2: Union[str, None, Unchanged] = UNCHANGED,
        operator: Union[ComparisonOperator, Unchanged] = UNCHANGED,
        rank: Union[int, Unchanged] = UNCHANGED,
        percent: Union[bool, Unchanged] = UNCHANGED,
        bottom: Union[bool, Unchanged] = UNCHANGED,
        above_average: Union[bool, Unchanged] = UNCHANGED,
        equal_average: Union[bool, Unchanged] = UNCHANGED,
        std_dev: Union[float, Unchanged] = UNCHANGED,
        text: Union[str, Unchanged] = UNCHANGED,
        time_period: Union[DatePeriod, Unchanged] = UNCHANGED,
        color_scale: Union[CfColorScaleInput, CfColorScale, Unchanged] = UNCHANGED,
        data_bar: Union[CfDataBarInput, CfDataBar, Unchanged] = UNCHANGED,
        icon_set: Union[CfIconSetInput, CfIconSet, Unchanged] = UNCHANGED,
    ) -> None:
        """Update ``cf_rule``, an index in :meth:`list_cf_rules` order or a rule's ``id``; an omitted argument
        keeps its value.

        The arguments are the fields of the rule's own input dataclass, except ``type`` and ``id``: a rule keeps
        its type, id and position. ``None`` clears ``dxf_index``, ``formula1`` and ``formula2`` and is a
        ``TypeError`` for any other argument. A payload (``color_scale``, ``data_bar``, ``icon_set``) replaces
        the whole payload; a ``None`` field of a data bar or icon set restores its default. The whole update is
        checked before anything changes.

        Raises:
          TypeError: for ``None`` on an argument that is not clearable.
          NotFoundError: when no rule has the id ``cf_rule``.
          IndexOutOfRangeError: when ``cf_rule`` is an index past the last rule.
          InvalidArgumentError: when an argument is not used by the rule's type, or a value is outside its domain.
        """
        given = {
            key: value
            for key, value in (
                ("priority", priority),
                ("stop_if_true", stop_if_true),
                ("dxf_index", dxf_index),
                ("ranges", ranges),
                ("formula1", formula1),
                ("formula2", formula2),
                ("op", operator),
                ("rank", rank),
                ("percent", percent),
                ("bottom", bottom),
                ("above_average", above_average),
                ("equal_average", equal_average),
                ("std_dev", std_dev),
                ("text", text),
                ("time_period", time_period),
                ("color_scale", color_scale),
                ("data_bar", data_bar),
                ("icon_set", icon_set),
            )
            if value is not UNCHANGED
        }
        for key, value in given.items():
            if value is None and key not in ("dxf_index", "formula1", "formula2"):
                raise TypeError(_argument_message("operator" if key == "op" else key, "must not be null"))
        for key in ("stop_if_true", "percent", "bottom", "above_average", "equal_average"):
            if key in given:
                _require_bool(given[key], key)
        rule_id = self._cf_rule_id(cf_rule)
        sheet = self._index()
        h = self._workbook._require()
        values: Dict[str, object] = {}
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CF_RULE_PATCH, 1)
            po = S.CF_RULE_PATCH.offsets
            if "priority" in given:
                values["priority_op"] = _PATCH_SET
                values["priority"] = _sint(given["priority"], "priority")
            for key in ("stop_if_true", "percent", "bottom", "above_average", "equal_average"):
                if key in given:
                    values[key + "_op"] = _PATCH_SET
                    values[key] = int(given[key])
            if "dxf_index" in given:
                if given["dxf_index"] is None:
                    values["dxf_index_op"] = _PATCH_CLEAR
                else:
                    values["dxf_index_op"] = _PATCH_SET
                    values["dxf_index"] = _uint(given["dxf_index"], "dxf_index", role="position")
            if "ranges" in given:
                rects = _cell_rects(self._workbook, given["ranges"], "ranges")
                values["ranges_op"] = _PATCH_SET
                values["ranges"] = _pack_cell_rect_array(rects, scope)
                values["range_count"] = _uint(len(rects), "range_count")
            for key in ("formula1", "formula2", "text"):
                if key not in given:
                    continue
                if given[key] is None:
                    values[key + "_op"] = _PATCH_CLEAR
                else:
                    values[key + "_op"] = _PATCH_SET
                    values[key] = scope.cstr(given[key])
            for key in ("op", "rank", "time_period"):
                if key in given:
                    values[key + "_op"] = _PATCH_SET
                    values[key] = _sint(given[key], key)
            if "std_dev" in given:
                values["std_dev_op"] = _PATCH_SET
                values["std_dev"] = _finite(given["std_dev"], "std_dev")
            if "color_scale" in given:
                payload = given["color_scale"]
                count = len(payload.thresholds)
                if count != len(payload.colors):
                    raise ValueError("color_scale thresholds and colors must have equal length")
                thresholds = scope.alloc(count * S.CFVO.size)
                colors = scope.alloc(count * S.COLOR_SPEC.size)
                for i, value in enumerate(payload.thresholds):
                    self._write_cfvo(thresholds + i * S.CFVO.size, value, scope)
                for i, color in enumerate(payload.colors):
                    _pack_color(colors + i * S.COLOR_SPEC.size, color, f"color_scale.colors[{i}]")
                values.update(
                    color_scale_thresholds_op=_PATCH_SET,
                    color_scale_thresholds=thresholds,
                    color_scale_thresholds_count=count,
                    color_scale_colors_op=_PATCH_SET,
                    color_scale_colors=colors,
                    color_scale_colors_count=count,
                )
            if "data_bar" in given:
                payload = given["data_bar"]
                self._write_cfvo(ptr + po["data_bar_min"][1], payload.min, scope)
                self._write_cfvo(ptr + po["data_bar_max"][1], payload.max, scope)
                _pack_color(ptr + po["data_bar_fill"][1], payload.fill, "data_bar.fill")
                for key in ("negative_fill", "border", "negative_border", "axis_color"):
                    _pack_optional_color(ptr + po[f"data_bar_{key}"][1], getattr(payload, key), f"data_bar.{key}")
                gradient = True if payload.gradient is None else payload.gradient
                axis_position = (
                    CfDataBarAxisPosition.AUTOMATIC if payload.axis_position is None else payload.axis_position
                )
                for key in (
                    "min",
                    "max",
                    "fill",
                    "show_value",
                    "min_length_pct",
                    "max_length_pct",
                    "gradient",
                    "axis_position",
                    "negative_fill",
                    "border",
                    "negative_border",
                    "axis_color",
                    "direction",
                ):
                    values[f"data_bar_{key}_op"] = _PATCH_SET
                values.update(
                    data_bar_show_value=int(_require_bool(payload.show_value, "data_bar.show_value")),
                    data_bar_min_length_pct=_uint(payload.min_length_pct, "min_length_pct", 8),
                    data_bar_max_length_pct=_uint(payload.max_length_pct, "max_length_pct", 8),
                    data_bar_gradient=int(_require_bool(gradient, "data_bar.gradient")),
                    data_bar_axis_position=_sint(axis_position, "axis_position"),
                    data_bar_direction=_sint(payload.direction, "direction"),
                )
            if "icon_set" in given:
                payload = given["icon_set"]
                count = len(payload.thresholds)
                thresholds = scope.alloc(count * S.CFVO.size)
                for i, value in enumerate(payload.thresholds):
                    self._write_cfvo(thresholds + i * S.CFVO.size, value, scope)
                floor = CfValueObject(CfValueObjectType.PERCENT, "0") if payload.floor is None else payload.floor
                self._write_cfvo(ptr + po["icon_set_floor"][1], floor, scope)
                icons = payload.icons or ()
                icons_ptr = scope.alloc(len(icons) * S.CF_ICON.size) if icons else 0
                for i, icon in enumerate(icons):
                    S.CF_ICON.pack(
                        LIB,
                        icons_ptr + i * S.CF_ICON.size,
                        {
                            "icon_set": _sint(icon.icon_set, "icon_set"),
                            "icon_index": _sint(icon.icon_index, "icon_index"),
                        },
                    )
                values["icon_set_op"] = _PATCH_SET
                for key in ("thresholds", "reverse", "show_value", "percent", "floor", "icons"):
                    values[f"icon_set_{key}_op"] = _PATCH_SET
                values.update(
                    icon_set=_sint(payload.name, "icon_set"),
                    icon_set_thresholds=thresholds,
                    icon_set_thresholds_count=count,
                    icon_set_reverse=int(_require_bool(payload.reverse, "icon_set.reverse")),
                    icon_set_show_value=int(_require_bool(payload.show_value, "icon_set.show_value")),
                    icon_set_percent=int(_require_bool(payload.percent, "icon_set.percent")),
                    icon_set_icons=icons_ptr,
                    icon_set_icons_count=len(icons),
                )
            raw = bytearray(LIB.read_bytes(ptr, S.CF_RULE_PATCH.size))
            S.CF_RULE_PATCH.pack_into(raw, 0, values)
            LIB.write_bytes(ptr, bytes(raw))
            _check(LIB.fm_sheet_update_cf_rule(h, _uint(sheet, "sheet_index", role="sheet"), scope.cstr(rule_id), ptr))

    def remove_cf_rule(self, cf_rule: CfRuleRef) -> None:
        """Remove ``cf_rule``, an index in :meth:`list_cf_rules` order or a rule's ``id``.

        Raises:
          NotFoundError: when no rule has the id ``cf_rule``.
          IndexOutOfRangeError: when ``cf_rule`` is an index past the last rule.
        """
        rule_id = self._cf_rule_id(cf_rule)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            _check(LIB.fm_sheet_remove_cf_rule(h, _uint(sheet, "sheet_index", role="sheet"), scope.cstr(rule_id)))

    def clear_cf_rules(self) -> None:
        """Drop every CF block on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_cf_rules(h, _uint(sheet, "sheet_index", role="sheet")))

    # -- Styles ------------------------------------------------------------
    @overload
    @_coordinate_form
    def get_cell_format(self, row: int, col: int) -> CellFormat: ...

    @overload
    def get_cell_format(self, a1: str) -> CellFormat: ...

    def get_cell_format(self, *args: Any, **kwargs: Any) -> CellFormat:
        """Return the format the cell at ``(row, col)`` shows.

        The number format, font, fill, border, alignment and protection of
        its effective XF (cell, then row, then column, then the default).
        Conditional formatting is not applied.
        """
        row, col = self._cell_args("get_cell_format", args, kwargs, ())
        workbook = self._workbook
        style = self.ooxml.get_effective_style(row, col)
        return _cell_format_of(workbook, workbook.ooxml.get_cell_xf(style.xf_index), style.source)

    def update_cell_format(
        self,
        range: Union[CellRect, str],
        *,
        number_format: Union[str, Unchanged] = UNCHANGED,
        font: Union[FontPatch, Font, Unchanged] = UNCHANGED,
        fill: Union[FillPatch, Fill, Unchanged] = UNCHANGED,
        border: Union[BorderPatch, Border, Unchanged] = UNCHANGED,
        alignment: Union[AlignmentPatch, Alignment, Unchanged] = UNCHANGED,
        protection: Union[CellProtectionPatch, CellProtection, Unchanged] = UNCHANGED,
        **read_only: Any,
    ) -> None:
        """Merge the given groups into the format of every cell of ``range``; an omitted group or key is kept.

        A group is a patch mapping (:class:`FontPatch`, ...) or the record
        :meth:`get_cell_format` reads, so passing every field of a
        :class:`CellFormat` copies that format (its read-only ``source`` is accepted and
        ignored): ``ws.update_cell_format("B1", **vars(ws.get_cell_format("A1")))``. ``None``
        clears a key the read record reports as nullable (``bold``, ``size_pt``,
        ``family``, ...). A colour is ``"#RRGGBB"``, ``"#AARRGGBB"`` or a
        :class:`Color`; a fill colour without a pattern makes the fill solid.
        A whole-column, whole-row or whole-sheet ``range`` (``"B:B"``,
        ``"3:3"``, ``"A:XFD"``) formats the column, row or sheet, so cells
        entered there later take the format too. Equal formats share one
        style record.

        Every argument is checked before anything is written and the sheet
        changes in one step at the end, but an engine failure on the way can
        leave unused style records behind (:meth:`WorkbookOoxml.list_cell_xfs`).

        Raises:
          TypeError: for an unknown key, a value of the wrong type, or
            ``None`` on a key that is not nullable.
          ValueError: for a negative or fractional count or enum value, or a
            non-finite size.
          InvalidArgumentError: for an empty ``number_format``, an enum value
            outside its table, a font size outside 1..409, ``indent`` above
            255, or ``text_rotation`` outside 0..180 and 255.

        Example::

            ws.update_cell_format("A1:B2", number_format="0.00", font={"bold": True})
        """
        for key in read_only:
            if key not in READ_ONLY_KEYS["Sheet.update_cell_format"]:
                raise TypeError(f"{_call_label()}() got an unexpected keyword argument {key!r}")
        patch = _check_cell_format_patch(
            {
                "number_format": number_format,
                "font": font,
                "fill": fill,
                "border": border,
                "alignment": alignment,
                "protection": protection,
            }
        )
        rect = self._rect(range)
        bases = self.ooxml.get_range_xf_bases(rect)
        if not patch:
            return
        workbook = self._workbook
        num_fmt_id: Optional[int] = None
        if "number_format" in patch:
            code = patch["number_format"]
            num_fmt_id = 0 if code == "General" else workbook.ooxml.add_number_format(code)
        ooxml = workbook.ooxml
        xfs = [ooxml.add_cell_xf(_patched_xf(workbook, ooxml.get_cell_xf(base), patch, num_fmt_id)) for base in bases]
        self.ooxml.remap_range_xfs(rect, dict(zip(bases, xfs)))

    # -- Pivot tables ------------------------------------------------------
    def _pivot_count(self) -> int:
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(LIB.fm_sheet_get_pivot_count, h, _uint(sheet, "sheet_index", role="sheet"))

    def _pivot_name(self, pivot_index: int) -> str:
        sheet = self._index()
        h = self._workbook._require()
        return _read_cstr_out(
            LIB.fm_sheet_get_pivot_name,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(pivot_index, "pivot_index", role="position"),
        )

    def _pivot_index(self, pivot: PivotRef) -> int:
        """The index ``pivot`` addresses: an index, a pivot table's name or a :class:`Pivot` of this sheet."""
        return resolve_ref(
            pivot, lambda: [self._pivot_name(i) for i in range(self._pivot_count())], "pivot", owner=self
        )

    def _pivot_at(self, pivot_index: int) -> Pivot:
        """The handle of the pivot at ``pivot_index``; the sheet keeps one per pivot."""
        from ._pivot import Pivot

        sheet = self._index()
        h = self._workbook._require()
        runtime_id = _read_count(
            LIB.fm_sheet_get_pivot_runtime_id,
            h,
            _uint(sheet, "sheet_index", role="sheet"),
            _uint(pivot_index, "pivot_index", role="position"),
        )
        pivot = self._pivots.get(runtime_id)
        if pivot is None:
            pivot = Pivot._new(self, runtime_id)
            self._pivots[runtime_id] = pivot
        return pivot

    def list_pivots(self) -> List[Pivot]:
        """Return the handles of the pivot tables anchored on the sheet, in index order."""
        return [self._pivot_at(i) for i in range(self._pivot_count())]

    def pivot(self, pivot: PivotRef) -> Pivot:
        """Return the handle of the pivot table ``pivot``: an index, a name or a :class:`Pivot` of this sheet.

        Raises:
          NotFoundError: when no pivot has the name ``pivot``.
          IndexOutOfRangeError: when the index is past the last pivot.
          InvalidArgumentError: when ``pivot`` is a :class:`Pivot` of another sheet.
        """
        pivot_index = _uint(self._pivot_index(pivot), "pivot_index", role="position")
        if pivot_index >= self._pivot_count():
            raise FormulonError(
                StatusCode.INDEX_OUT_OF_RANGE,
                _argument_message("pivot", f"{pivot_index} is out of range ({self._pivot_count()} items)"),
                "",
            )
        return self._pivot_at(pivot_index)

    def add_pivot(self, input: PivotInput) -> Pivot:
        """Create the empty pivot table ``input`` describes and return its handle.

        ``input.anchor`` is the top-left cell, a :class:`CellAddress` or one A1 cell; ``input.cache_id`` is a
        cache id or a :class:`PivotCache` of this workbook.

        Raises:
          InvalidArgumentError: when ``input.cache_id`` is a :class:`PivotCache` of another workbook.

        Example::

            cache = wb.add_pivot_cache()
            pivot = ws.add_pivot(PivotInput(name="Sales", cache_id=cache, anchor="E1"))
        """
        if not isinstance(input, PivotInput):
            raise TypeError(_argument_message("input", "must be a PivotInput"))
        anchor_row, anchor_col = self._cell_address(input.anchor, "input.anchor")
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            spec = _pack_pivot_spec(
                scope, input.name, self._workbook._pivot_cache_id(input.cache_id), anchor_row, anchor_col
            )
            out = scope.out_ptr()
            _check(LIB.fm_sheet_add_pivot(h, _uint(sheet, "sheet_index", role="sheet"), spec, out))
            pivot_index = LIB.read_u32(out)
        return self._pivot_at(pivot_index)

    def remove_pivot(self, pivot: PivotRef) -> None:
        """Remove the pivot table ``pivot``; its handle raises :class:`NotFoundError` from then on."""
        pivot_index = self._pivot_index(pivot)
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_remove_pivot(
                h, _uint(sheet, "sheet_index", role="sheet"), _uint(pivot_index, "pivot_index", role="position")
            )
        )

    # -- Tables ------------------------------------------------------------
    def _table_count(self) -> int:
        return self._workbook._table_count(self._index())

    def _table_index(self, table: TableRef) -> int:
        """The sheet-local index ``table`` addresses: an index, a table name or a :class:`Table` of this sheet."""
        from ._table import Table

        if isinstance(table, Table):
            if table.sheet is not self:
                raise FormulonError(
                    StatusCode.INVALID_ARGUMENT, _argument_message("table", "is a Table of another sheet"), ""
                )
            return table.index
        sheet = self._index()
        h = self._workbook._require()
        return resolve_ref(
            table,
            lambda: [Table._name(h, sheet, i) for i in range(self._table_count())],
            "table",
        )

    def list_tables(self) -> List[Table]:
        """Return the handles of the sheet's worksheet tables; a table's position is its sheet-local index."""
        sheet = self._index()
        return [self._workbook._table_handle(sheet, i) for i in range(self._table_count())]

    def table(self, table: TableRef) -> Table:
        """Return the handle of the table ``table`` on this sheet: an index, a name or a :class:`Table` of this sheet.

        A name is compared case-insensitively and a string is always a name,
        digits included. :meth:`Workbook.table` finds a table on any sheet.

        Raises:
          NotFoundError: when no table of the sheet has the name ``table``.
          IndexOutOfRangeError: when the index is past the last table.
          InvalidArgumentError: when ``table`` is a :class:`Table` of another sheet.
        """
        return self._workbook._table_handle(
            self._index(), _uint(self._table_index(table), "table_index", role="position")
        )

    def add_table(self, input: TableInput) -> Table:
        """Create the worksheet table ``input`` describes on this sheet and return its handle.

        A ``name`` another table already has (compared case-insensitively)
        raises ``AlreadyExistsError``. See :class:`TableInput` for the defaults.

        Example::

            ws.set_values("A1", [["Item", "Qty"], ["Apple", 3]])
            table = ws.add_table(TableInput(range="A1:B2", name="Sales"))
        """
        return self._workbook._add_table(self._index(), input)

    def remove_table(self, table: TableRef) -> None:
        """Remove the table ``table``; its handle raises :class:`NotFoundError` from then on.

        Formulas that name the table keep their structured references and
        evaluate to ``#NAME?``. Fails as :meth:`table` does when there is no such table.
        """
        table_index = _uint(self._table_index(table), "table_index", role="position")
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_table(h, _uint(sheet, "sheet_index", role="sheet"), table_index))

    # -- Dependency-graph trace --------------------------------------------
    @overload
    @_coordinate_form
    def get_precedents(self, row: int, col: int, depth: int = 1) -> List[CellNode]: ...

    @overload
    def get_precedents(self, a1: str, depth: int = 1) -> List[CellNode]: ...

    def get_precedents(self, *args: Any, **kwargs: Any) -> List[CellNode]:
        """Return the cells ``(row, col)`` reads (up to ``depth``)."""
        row, col, depth = self._cell_args("get_precedents", args, kwargs, (("depth", 1),))
        sheet = self._index()
        return self._trace(
            LIB.fm_sheet_get_precedents,
            _uint(sheet, "sheet", role="sheet"),
            _uint(row, "row", role="coordinate"),
            _uint(col, "col", role="coordinate"),
            _uint(depth, "depth"),
        )

    @overload
    @_coordinate_form
    def get_dependents(self, row: int, col: int, depth: int = 1) -> List[CellNode]: ...

    @overload
    def get_dependents(self, a1: str, depth: int = 1) -> List[CellNode]: ...

    def get_dependents(self, *args: Any, **kwargs: Any) -> List[CellNode]:
        """Return the cells that read ``(row, col)`` (up to ``depth``)."""
        row, col, depth = self._cell_args("get_dependents", args, kwargs, (("depth", 1),))
        sheet = self._index()
        return self._trace(
            LIB.fm_sheet_get_dependents,
            _uint(sheet, "sheet", role="sheet"),
            _uint(row, "row", role="coordinate"),
            _uint(col, "col", role="coordinate"),
            _uint(depth, "depth"),
        )

    def _trace(self, fn, sheet: int, row: int, col: int, depth: int) -> List[CellNode]:
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            handle = 0
            try:
                _check(fn(h, sheet, row, col, depth, out))
                handle = LIB.read_u32(out)
                nodes: List[CellNode] = []
                n = _read_count(LIB.fm_cell_nodes_get_node_count, handle)
                # One scratch block for the whole walk; see list_defined_names.
                nptr = scope.struct(S.CELL_NODE)
                for i in range(n):
                    S.zero_struct(LIB, S.CELL_NODE, nptr)
                    _check(LIB.fm_cell_nodes_get_node_at(handle, _uint(i, "idx", role="position"), nptr))
                    d = S.CELL_NODE.unpack(LIB, nptr)
                    nodes.append(CellNode(sheet=d["sheet_index"], row=d["row"], col=d["col"]))
                return nodes
            finally:
                if handle:
                    LIB.fm_cell_nodes_destroy(handle)

    # -- Dynamic-array spill -----------------------------------------------
    @overload
    @_coordinate_form
    def get_spill(self, row: int, col: int) -> Optional[Spill]: ...

    @overload
    def get_spill(self, a1: str) -> Optional[Spill]: ...

    def get_spill(self, *args: Any, **kwargs: Any) -> Optional[Spill]:
        """Return the spill region ``(row, col)`` belongs to, or ``None`` when it is in no spill.

        A blocked anchor reads as a :class:`Spill` with ``blocked`` set and the footprint it attempted.
        """
        row, col = self._cell_args("get_spill", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.SPILL_INFO)
            _check(
                LIB.fm_sheet_get_spill(
                    h,
                    _uint(sheet, "sheet", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    present,
                    ptr,
                )
            )
            if LIB.read_i32(present) == 0:
                return None
            d = S.SPILL_INFO.unpack(LIB, ptr)
            return Spill(blocked=bool(d["blocked"]), range=_cell_rect_from(d["rect"]))

    # -- Legacy array (CSE) formulas ---------------------------------------
    @overload
    @_coordinate_form
    def get_array_formula(self, row: int, col: int) -> Optional[ArrayFormula]: ...

    @overload
    def get_array_formula(self, a1: str) -> Optional[ArrayFormula]: ...

    def get_array_formula(self, *args: Any, **kwargs: Any) -> Optional[ArrayFormula]:
        """Return the legacy array formula covering ``(row, col)``, or ``None``.

        Every cell of the block reports the same :class:`ArrayFormula`.
        """
        row, col = self._cell_args("get_array_formula", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            ptr = scope.struct(S.ARRAY_FORMULA)
            _check(
                LIB.fm_sheet_get_array_formula(
                    h,
                    _uint(sheet, "sheet", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    present,
                    ptr,
                )
            )
            if LIB.read_i32(present) == 0:
                return None
            d = S.ARRAY_FORMULA.unpack(LIB, ptr)
            return ArrayFormula(
                range=_cell_rect_from(d["rect"]),
                formula=LIB.read_cstr(d["formula"]),
            )

    def set_array_formula(self, array_formula: ArrayFormula) -> None:
        """Enter ``array_formula.formula`` as one legacy array formula over ``array_formula.range``.

        ``array_formula.range`` is a :class:`CellRect` or an unqualified A1 range.
        The top-left cell of the rectangle anchors the block. Every cell of the rectangle loses its value and formula, and
        the next :meth:`Workbook.recalc` distributes the result. Refused when the
        rectangle partially overlaps another array, or intersects a merged
        area, a data table, a PivotTable or (for more than one cell) a table.
        Writing to the top-left cell of a block replaces the whole block.
        """
        rect = self._rect(array_formula.range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            formula_ptr = scope.cstr(array_formula.formula)
            status = LIB.fm_sheet_set_array_formula(h, _uint(sheet, "sheet", role="sheet"), rp, formula_ptr)
            _check(status)


class SheetOoxml(_NoCopy):
    """The OOXML-level layer of a sheet, :attr:`Sheet.ooxml`.

    The cell, row and column XF indices, the range-style remap that works on
    them, and the raw XML of the worksheet elements the typed setters cover.
    Most code uses :meth:`Sheet.update_cell_format`, which takes whole columns
    (``"B:B"``) and rows (``"3:3"``) too.
    """

    __slots__ = ("_sheet",)

    def __init__(self, *args: object, **kwargs: object) -> None:
        raise TypeError("obtain the OOXML layer from Sheet.ooxml")

    @classmethod
    def _new(cls, sheet: Sheet) -> "SheetOoxml":
        layer = object.__new__(cls)
        layer._sheet = sheet
        return layer

    def __repr__(self) -> str:
        return f"{self._sheet!r}.ooxml"

    def get_column_xf_index(self, col: Union[int, str]) -> Optional[int]:
        """The cell-style XF index column ``col`` carries (:meth:`set_column_xf_index`); ``None`` when it has none.

        ``col`` is a zero-based index or one A1 column such as ``"B"``.
        """
        col = _uint(self._sheet._single_axis(col, "col"), "col", role="coordinate")
        spans = self._sheet._layout_overrides(
            LIB.fm_sheet_get_column_layout_count, LIB.fm_sheet_get_column_layout_at, S.COLUMN_LAYOUT
        )
        for d in spans:
            if d["first"] <= col <= d["last"]:
                return d["xf_index"] if d["has_xf_index"] else None
        return None

    def get_row_xf_index(self, row: Union[int, str]) -> Optional[int]:
        """The cell-style XF index row ``row`` makes effective (:meth:`set_row_xf_index`); ``None`` when it has none.

        ``row`` is a zero-based index or one A1 row such as ``"3"``.
        """
        row = _uint(self._sheet._single_axis(row, "row"), "row", role="coordinate")
        rows = self._sheet._layout_overrides(
            LIB.fm_sheet_get_row_layout_count, LIB.fm_sheet_get_row_layout_at, S.ROW_LAYOUT
        )
        for d in rows:
            if d["row"] == row:
                return d["xf_index"] if d["has_xf_index"] else None
        return None

    def get_auto_filter_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`Sheet.get_auto_filter`.

        Return the sheet's ``<autoFilter>`` element, or ``None`` without one.
        The fragment carries the filter-column criteria, sort state and
        extension payload verbatim, so a host can preserve definitions it
        does not interpret.
        """
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        return (
            _read_cstr_out(
                LIB.fm_sheet_get_auto_filter_xml,
                h,
                _uint(sheet, "sheet_index", role="sheet"),
            )
            or None
        )

    def set_auto_filter_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`Sheet.set_auto_filter`.

        Replace the sheet's ``<autoFilter>`` element; ``None`` (or ``""``) removes the AutoFilter. Other input must be a
        complete ``<autoFilter ...>`` fragment; the engine preserves filter
        extensions verbatim rather than validating them in detail.
        """
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        with _ArgScope() as scope:
            xml_ptr = scope.cstr(xml or "")
            _check(LIB.fm_sheet_set_auto_filter_xml(h, _uint(sheet, "sheet_index", role="sheet"), xml_ptr))

    def get_page_setup_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`Sheet.get_page_setup`.

        Return the sheet's ``<pageSetup>`` element, or ``None`` without one.
        """
        return self._sheet._get_print_xml("fm_sheet_get_page_setup_xml")

    def set_page_setup_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`Sheet.update_page_setup`.

        Replace the sheet's ``<pageSetup>`` element; ``None`` (or ``""``)
        removes it. ``fitToPage`` is not a ``<pageSetup>`` attribute but lives
        in ``<sheetPr>``. A fragment carrying ``r:id`` is rejected
        unless the sheet already has a printerSettings part, because the
        reference would otherwise dangle and Excel would repair the file.
        """
        self._sheet._set_print_xml("fm_sheet_set_page_setup_xml", xml)

    def get_page_margins_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`Sheet.get_page_margins`.

        Return the sheet's ``<pageMargins>`` element, or ``None`` without one.
        """
        return self._sheet._get_print_xml("fm_sheet_get_page_margins_xml")

    def set_page_margins_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`Sheet.update_page_margins`.

        Replace the sheet's ``<pageMargins>`` element; ``None`` (or ``""``) removes it.
        """
        self._sheet._set_print_xml("fm_sheet_set_page_margins_xml", xml)

    def get_print_options_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`Sheet.get_print_options`.

        Return the sheet's ``<printOptions>`` element, or ``None`` without one.
        """
        return self._sheet._get_print_xml("fm_sheet_get_print_options_xml")

    def set_print_options_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`Sheet.update_print_options`.

        Replace the sheet's ``<printOptions>`` element; ``None`` (or ``""``) removes it.
        """
        self._sheet._set_print_xml("fm_sheet_set_print_options_xml", xml)

    def get_header_footer_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`Sheet.get_header_footer`.

        Return the sheet's ``<headerFooter>`` element, or ``None`` without one.
        """
        return self._sheet._get_print_xml("fm_sheet_get_header_footer_xml")

    def set_header_footer_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`Sheet.update_header_footer`.

        Replace the sheet's ``<headerFooter>`` element; ``None`` (or ``""``)
        removes it. Header and footer codes are introduced by ``&``, which lives in XML
        element text and so must arrive escaped: Excel writes
        ``<oddHeader>&amp;C...</oddHeader>``. Pass decoded text to
        :meth:`Sheet.update_header_footer` instead to let the engine escape it.
        """
        self._sheet._set_print_xml("fm_sheet_set_header_footer_xml", xml)

    def get_sheet_pr_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`Sheet.get_page_setup`, which reads its ``fitToPage``.

        Return the sheet's ``<sheetPr>`` element, or ``None`` without one. It
        also carries the tab colour and the VBA code name.
        """
        return self._sheet._get_print_xml("fm_sheet_get_sheet_pr_xml")

    def set_sheet_pr_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`Sheet.update_page_setup`, which writes its ``fitToPage``.

        Replace the sheet's ``<sheetPr>`` element; ``None`` (or ``""``) removes it.
        """
        self._sheet._set_print_xml("fm_sheet_set_sheet_pr_xml", xml)

    @overload
    @_span_form
    def set_column_xf_index(self, first: int, last: int, xf_index: int) -> None: ...

    @overload
    def set_column_xf_index(self, span: str, xf_index: int) -> None: ...

    def set_column_xf_index(self, *args: Any, **kwargs: Any) -> None:
        """Set the cell-style XF on ``[first, last]``, keeping the other column attributes.

        Raises:
          IndexOutOfRangeError: when ``xf_index`` is not below the workbook's cell-XF count.
        """
        first, last, xf_index = self._sheet._span_args(
            "set_column_xf_index", args, kwargs, ("first", "last", "xf_index")
        )
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        _check(
            LIB.fm_sheet_set_column_xf_index(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
                _uint(xf_index, "xf_index", role="position"),
            )
        )

    @overload
    @_span_form
    def reset_column_xf_index(self, first: int, last: int) -> None: ...

    @overload
    def reset_column_xf_index(self, span: str) -> None: ...

    def reset_column_xf_index(self, *args: Any, **kwargs: Any) -> None:
        """Remove the cell-style XF on ``[first, last]``, keeping the other column attributes."""
        first, last = self._sheet._span_args("reset_column_xf_index", args, kwargs, ("first", "last"))
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        _check(
            LIB.fm_sheet_reset_column_xf_index(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
            )
        )

    @overload
    @_span_form
    def set_row_xf_index(self, first: int, last: int, xf_index: int) -> None: ...

    @overload
    def set_row_xf_index(self, span: str, xf_index: int) -> None: ...

    def set_row_xf_index(self, *args: Any, **kwargs: Any) -> None:
        """Set the cell-style XF on the inclusive row span ``first..last``, keeping the other row attributes.

        Every row of the span gets an override. The span and ``xf_index`` are
        validated before anything changes.

        Raises:
          IndexOutOfRangeError: when ``xf_index`` is not below the workbook's cell-XF count.
        """
        first, last, xf_index = self._sheet._span_args("set_row_xf_index", args, kwargs, ("first", "last", "xf_index"))
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        _check(
            LIB.fm_sheet_set_row_xf_index(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
                _uint(xf_index, "xf_index", role="position"),
            )
        )

    @overload
    @_span_form
    def reset_row_xf_index(self, first: int, last: int) -> None: ...

    @overload
    def reset_row_xf_index(self, span: str) -> None: ...

    def reset_row_xf_index(self, *args: Any, **kwargs: Any) -> None:
        """Remove the cell-style XF on the inclusive row span ``first..last``, keeping the other row attributes."""
        first, last = self._sheet._span_args("reset_row_xf_index", args, kwargs, ("first", "last"))
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        _check(
            LIB.fm_sheet_reset_row_xf_index(
                h,
                _uint(sheet, "sheet_index", role="sheet"),
                _uint(first, "first", role="coordinate"),
                _uint(last, "last", role="coordinate"),
            )
        )

    @overload
    @_coordinate_form
    def get_effective_style(self, row: int, col: int) -> EffectiveStyle: ...

    @overload
    def get_effective_style(self, a1: str) -> EffectiveStyle: ...

    def get_effective_style(self, *args: Any, **kwargs: Any) -> EffectiveStyle:
        """Read the effective style of the cell at ``(row, col)``; the cell need not exist."""
        row, col = self._sheet._cell_args("get_effective_style", args, kwargs, ())
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.EFFECTIVE_STYLE)
            _check(
                LIB.fm_sheet_get_effective_style(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    ptr,
                )
            )
            d = S.EFFECTIVE_STYLE.unpack(LIB, ptr)
            offsets = S.EFFECTIVE_STYLE.offsets
            border_argb = struct.unpack("<5I", LIB.read_bytes(ptr + offsets["border_argb"][1], 20))
            border_res = struct.unpack("<5i", LIB.read_bytes(ptr + offsets["border_resolution"][1], 20))
            return EffectiveStyle(
                xf_index=d["xf_index"],
                source=EffectiveStyleSource(d["source"]),
                font_index=d["font_index"],
                fill_index=d["fill_index"],
                border_index=d["border_index"],
                font_color=ResolvedColor(d["font_argb"], ColorResolution(d["font_resolution"])),
                fill_foreground=ResolvedColor(
                    d["fill_foreground_argb"], ColorResolution(d["fill_foreground_resolution"])
                ),
                fill_background=ResolvedColor(
                    d["fill_background_argb"], ColorResolution(d["fill_background_resolution"])
                ),
                borders=EffectiveBorders(
                    *(ResolvedColor(a, ColorResolution(r)) for a, r in zip(border_argb, border_res))
                ),
                protection=CellProtection(locked=bool(d["locked"]), hidden=bool(d["hidden"])),
                number_format=LIB.read_cstr(d["number_format"]),
            )

    def snapshot_range_xfs(self, range: Union[CellRect, str]) -> bytes:
        """Capture the style state of a selection as opaque bytes for :meth:`restore_range_xfs`.

        The capture holds the covered column styles (column selection), row
        styles (row selection) and every stored cell's XF and explicit mark.
        Take it before :meth:`remap_range_xfs` to be able to undo the edit.
        """
        rect = self._sheet._rect(range)
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            bytes_ptr = scope.out_ptr()
            len_ptr = scope.out_ptr()
            data_ptr = 0
            try:
                _check(
                    LIB.fm_sheet_snapshot_range_xfs(
                        h,
                        _uint(sheet, "sheet_index", role="sheet"),
                        rp,
                        bytes_ptr,
                        len_ptr,
                    )
                )
                data_ptr = LIB.read_u32(bytes_ptr)
                return LIB.read_bytes(data_ptr, LIB.read_u32(len_ptr))
            finally:
                if data_ptr:
                    LIB.fm_buffer_destroy(data_ptr)

    def restore_range_xfs(self, bytes: Union[bytes, bytearray, memoryview]) -> None:
        """Put a :meth:`snapshot_range_xfs` capture back on the sheet.

        Captured axis styles and cell XFs are restored exactly; a stored cell of
        the captured rectangle that was not captured loses its style. Values and
        formulas are untouched. An invalid capture raises :class:`FormulonError`
        and changes nothing.
        """
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        raw = builtins.bytes(bytes)
        with _ArgScope() as scope:
            buf = scope.buffer(raw) if raw else 0
            _check(
                LIB.fm_sheet_restore_range_xfs(
                    h, _uint(sheet, "sheet_index", role="sheet"), buf, _uint(len(raw), "len")
                )
            )

    def get_range_xf_bases(self, range: Union[CellRect, str]) -> List[int]:
        """List, ascending, every cell XF some cell of the selection resolves to.

        ``range`` is a selection: full height is a column selection, full
        width a row selection, both the whole sheet, neither a cell range. It
        is the first step of an attribute patch (for example a font size over a selection): read the
        bases, build a patched XF for each with :meth:`WorkbookOoxml.add_font` /
        :meth:`WorkbookOoxml.add_cell_xf`, then apply them with :meth:`remap_range_xfs`.
        """
        rect = self._sheet._rect(range)
        return self.get_range_xf_bases_excluding(rect, ())

    def get_range_xf_bases_excluding(
        self,
        range: Union[CellRect, str],
        excluded: Sequence[Union[CellAddress, Tuple[int, int], str]],
    ) -> List[int]:
        """:meth:`get_range_xf_bases` over the selection less the cells in ``excluded``.

        Each excluded cell is a :class:`CellAddress`, a ``(row, col)`` pair or an A1 cell.

        An excluded cell contributes no XF, neither its own nor the row, column
        or default style beneath it. Cells outside the selection and repeats are
        ignored; excluding every cell yields an empty list.
        """
        rect = self._sheet._rect(range)
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        cells = [self._sheet._cell_address(cell, "excluded cell") for cell in excluded]
        count = len(cells)
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            cells_buf = 0
            if count:
                cells_buf = scope.buffer(
                    b"".join(
                        struct.pack(
                            "<II",
                            _uint(r, "excluded_row", role="coordinate"),
                            _uint(c, "excluded_col", role="coordinate"),
                        )
                        for r, c in cells
                    )
                )
            len_ptr = scope.out_ptr()
            # Size the buffer first, then fetch.
            _check(
                LIB.fm_sheet_get_range_xf_bases_excluding(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    rp,
                    cells_buf,
                    _uint(count, "excluded_count"),
                    0,
                    0,
                    len_ptr,
                )
            )
            total = LIB.read_u32(len_ptr)
            if total == 0:
                return []
            buf = scope.alloc(4 * total)
            _check(
                LIB.fm_sheet_get_range_xf_bases_excluding(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    rp,
                    cells_buf,
                    _uint(count, "excluded_count"),
                    buf,
                    _uint(total, "cap"),
                    len_ptr,
                )
            )
            n = min(total, LIB.read_u32(len_ptr))
            return list(struct.unpack("<%dI" % n, LIB.read_bytes(buf, 4 * n)))

    def remap_range_xfs(self, range: Union[CellRect, str], mapping: Mapping[int, int]) -> None:
        """Remap the XFs of a selection: a cell resolving to key ``k`` ends at ``mapping[k]``.

        An XF absent from ``mapping`` is kept. Each target must be a valid cell
        XF, else :class:`FormulonError` and nothing changes. Patch an attribute
        by reading :meth:`get_range_xf_bases`, adding a patched XF for each base
        and passing ``{base: patched}``. Undo with :meth:`snapshot_range_xfs`
        taken beforehand and :meth:`restore_range_xfs`.
        """
        rect = self._sheet._rect(range)
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        pairs = list(mapping.items())
        count = len(pairs)
        pairs_buf = 0
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            if count:
                pairs_buf = scope.buffer(
                    b"".join(
                        struct.pack("<II", _uint(k, "from_xf", role="position"), _uint(v, "to_xf", role="position"))
                        for k, v in pairs
                    )
                )
            _check(
                LIB.fm_sheet_remap_range_xfs(
                    h,
                    _uint(sheet, "sheet_index", role="sheet"),
                    rp,
                    pairs_buf,
                    _uint(count, "count"),
                )
            )

    @overload
    @_coordinate_form
    def get_cell_xf_index(self, row: int, col: int) -> Optional[int]: ...

    @overload
    def get_cell_xf_index(self, a1: str) -> Optional[int]: ...

    def get_cell_xf_index(self, *args: Any, **kwargs: Any) -> Optional[int]:
        """Return the xf (style record) index attached to a cell, or ``None`` when the cell has none of its own."""
        row, col = self._sheet._cell_args("get_cell_xf_index", args, kwargs, ())
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_get_cell_xf_index(
                    h,
                    _uint(sheet, "sheet", role="sheet"),
                    _uint(row, "row", role="coordinate"),
                    _uint(col, "col", role="coordinate"),
                    present,
                    out,
                )
            )
            return LIB.read_u32(out) if LIB.read_i32(present) else None

    @overload
    @_coordinate_form
    def set_cell_xf_index(self, row: int, col: int, xf_index: int) -> None: ...

    @overload
    def set_cell_xf_index(self, a1: str, xf_index: int) -> None: ...

    def set_cell_xf_index(self, *args: Any, **kwargs: Any) -> None:
        """Store ``xf_index`` on the cell at ``(row, col)``.

        Raises:
          IndexOutOfRangeError: when ``xf_index`` is not below the workbook's cell-XF count.
        """
        row, col, xf_index = self._sheet._cell_args("set_cell_xf_index", args, kwargs, ("xf_index",))
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        _check(
            LIB.fm_sheet_set_cell_xf_index(
                h,
                _uint(sheet, "sheet", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
                _uint(xf_index, "xf_index", role="position"),
            )
        )

    @overload
    @_coordinate_form
    def reset_cell_xf_index(self, row: int, col: int) -> None: ...

    @overload
    def reset_cell_xf_index(self, a1: str) -> None: ...

    def reset_cell_xf_index(self, *args: Any, **kwargs: Any) -> None:
        """Reset the cell at ``(row, col)`` to its inherited XF by removing its direct style assignment.

        A blank cell inherits the row / column style again; value, formula and
        comments are kept. An absent cell is a no-op. A non-blank cell keeps
        owning the default xf. Use the ``source`` of :meth:`get_effective_style`
        to tell an unassigned cell from one directly assigned the Normal xf.
        """
        row, col = self._sheet._cell_args("reset_cell_xf_index", args, kwargs, ())
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        _check(
            LIB.fm_sheet_reset_cell_xf_index(
                h,
                _uint(sheet, "sheet", role="sheet"),
                _uint(row, "row", role="coordinate"),
                _uint(col, "col", role="coordinate"),
            )
        )

    def set_range_xf_index(self, range: Union[CellRect, str], xf_index: int) -> None:
        """Store ``xf_index`` on every cell of ``range``.

        Cells that do not exist yet are materialised as blanks carrying only
        the style, so an empty ruled box renders. Ruling a report this way
        costs one call instead of one per cell.

        Raises:
          IndexOutOfRangeError: when ``xf_index`` is not below the workbook's cell-XF count.
        """
        rect = self._sheet._rect(range)
        sheet = self._sheet._index()
        h = self._sheet._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            _check(
                LIB.fm_sheet_set_range_xf_index(
                    h, _uint(sheet, "sheet", role="sheet"), rp, _uint(xf_index, "xf_index", role="position")
                )
            )


for _qualname, _signature in _COORDINATE_FORMS.items():
    _owner, _, _name = _qualname.partition(".")
    getattr({"Sheet": Sheet, "SheetOoxml": SheetOoxml}[_owner], _name).__signature__ = _signature
del _qualname, _owner, _name, _signature

_install_workbook_locks(Sheet)
_install_workbook_locks(SheetOoxml)

# Bound after `Sheet` exists so `typing.get_type_hints` resolves the handle types its annotations name; the
# package imports `workbook` first, which imports this module before any of these.
from ._drawing import Chart, Drawing, Picture, Shape, Slicer  # noqa: E402
from ._pivot import Pivot  # noqa: E402
from ._table import Table  # noqa: E402
