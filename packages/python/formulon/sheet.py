"""The :class:`Sheet` handle: every operation that targets one sheet of a workbook.

A :class:`~formulon.Workbook` hands out one :class:`Sheet` per sheet and keeps it
pointing at the same sheet across insertions, moves, renames and removals;
see :class:`Sheet`.
"""

from __future__ import annotations

import builtins
import dataclasses
import inspect
import math
import numbers
import re
import struct
from typing import (
    TYPE_CHECKING,
    Any,
    Callable,
    Dict,
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
    _sint,
    _uint,
    fm_value_t_size,
)
from ._refs import resolve_ref
from ._status import (
    StatusCode,
)
from .workbook import (
    _CELL_RANGE_NO_CURSOR,
    _STATUS_NOT_FOUND,
    UNCHANGED,
    Alignment,
    AlignmentPatch,
    AnchorEditAs,
    AnchorKind,
    ArrayFormula,
    AutoFilter,
    AutoFilterEval,
    Border,
    BorderPatch,
    BorderStyle,
    Cell,
    CellFormat,
    CellNode,
    CellProtection,
    CellProtectionPatch,
    CellRect,
    CellsInRange,
    CellXf,
    CellXfInput,
    CfCell,
    CfColor,
    CfDataBarAxisPosition,
    CfDataBarDirection,
    CfIcon,
    CfMatch,
    CfMatchKind,
    Cfvo,
    ChartGrouping,
    ChartInsertKind,
    ChartLegendPosition,
    ChartSeriesDirection,
    Color,
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
    DisplayStatus,
    DisplayText,
    DrawingAnchor,
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
    IconSet,
    MarginsPt,
    Mention,
    PageBreak,
    PageLayout,
    PageMargins,
    PageOrientation,
    PageSetup,
    Pagination,
    Paper,
    PhoneticProperties,
    PhoneticRun,
    PrintOptions,
    PrintTitles,
    ProtectedRange,
    ReadingOrder,
    RectPt,
    ResolvedColor,
    RowLayout,
    ShapePaint,
    SheetComment,
    SheetFormatDefaults,
    SheetProtection,
    SheetSelection,
    SheetView,
    SheetViewMode,
    SheetVisibility,
    SlicerCrossFilter,
    SlicerSortOrder,
    Sparkline,
    SparklineAxisType,
    SparklineEmptyCells,
    SparklineGroup,
    SparklineType,
    Spill,
    ThreadedComment,
    ThreadedCommentInput,
    Unchanged,
    UnderlineStyle,
    ValidationErrorStyle,
    ValidationOutcome,
    Value,
    VerticalAlign,
    WidthModel,
    Workbook,
    _ArgScope,
    _cell_rect_fields,
    _cell_rect_from,
    _check,
    _check_value,
    _decode_cell_rect_array,
    _decode_cell_rect_at,
    _decode_color,
    _enum,
    _header_mode,
    _hyperlink_from,
    _image_anchor_fields,
    _install_workbook_locks,
    _nullable_enum,
    _pack_cell_rect_array,
    _pack_cell_rect_at,
    _pack_mentions,
    _pack_phonetic_run_array,
    _pack_shape,
    _pack_slicer,
    _pack_sparkline_group,
    _pack_sparklines,
    _print_titles_from,
    _read_count,
    _read_cstr_out,
    _read_optional_cstr_out,
    _threaded_comment_from,
    _write_mentions,
    _write_value_input,
)

if TYPE_CHECKING:
    from .drawing import Chart, Drawing, Picture, Shape, Slicer
    from .pivot import Pivot
    from .table import Table

__all__ = ["Sheet"]

# <facade-tables>
WORKBOOK_METHODS = (
    "add_border",
    "add_cell_style_xf",
    "add_cell_xf",
    "add_dxf",
    "add_fill",
    "add_font",
    "add_num_fmt",
    "add_person",
    "add_pivot_cache",
    "add_sheet",
    "add_table",
    "clear_theme",
    "close",
    "closed",
    "copy_sheet",
    "format_reference",
    "format_value",
    "get_active_sheet",
    "get_border",
    "get_calc_mode",
    "get_calc_on_save",
    "get_cell_style",
    "get_cell_style_xf",
    "get_cell_xf",
    "get_date1904",
    "get_defined_name",
    "get_diagnostics",
    "get_dxf",
    "get_excel_profile_id",
    "get_file_name",
    "get_fill",
    "get_font",
    "get_full_precision",
    "get_iterative",
    "get_memory_usage",
    "get_num_fmt_string",
    "get_passthrough",
    "get_person",
    "get_pinned_now",
    "get_ref_mode",
    "get_table_style",
    "get_table_style_defaults",
    "get_theme_colors",
    "get_theme_fonts",
    "get_theme_script_font",
    "get_workbook_protection",
    "has_dirty_cells",
    "list_borders",
    "list_cell_style_xfs",
    "list_cell_styles",
    "list_cell_xfs",
    "list_defined_names",
    "list_dxfs",
    "list_external_links",
    "list_fills",
    "list_fonts",
    "list_passthroughs",
    "list_people",
    "list_pivot_caches",
    "list_sheets",
    "list_table_styles",
    "list_tables",
    "parse_reference",
    "pivot_cache",
    "recalc",
    "recalc_data_tables",
    "remove_cell_style",
    "remove_defined_name",
    "remove_person",
    "remove_pivot_cache",
    "remove_sheet",
    "remove_table",
    "remove_table_style",
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
    "set_excel_profile_id",
    "set_file_name",
    "set_font",
    "set_full_precision",
    "set_iterative",
    "set_pinned_now",
    "set_ref_mode",
    "set_table_style",
    "set_table_style_defaults",
    "set_theme_colors",
    "set_theme_fonts",
    "set_theme_script_font",
    "set_workbook_protection",
    "sheet",
    "table",
)
SHEET_METHODS = (
    "add_cf_rule",
    "add_chart",
    "add_column_break",
    "add_data_table",
    "add_hyperlink",
    "add_hyperlink_range",
    "add_merge",
    "add_picture",
    "add_pivot",
    "add_protected_range",
    "add_row_break",
    "add_shape",
    "add_slicer",
    "add_sparkline_group",
    "add_threaded_comment",
    "add_validation",
    "apply_auto_filter",
    "chart",
    "clear_auto_filter",
    "clear_breaks",
    "clear_cell_xf_index",
    "clear_cf_rules",
    "clear_column_style",
    "clear_column_width",
    "clear_hyperlinks",
    "clear_merges",
    "clear_row_height",
    "clear_row_style",
    "clear_validations",
    "convert_column_chars_to_pt",
    "convert_column_pt_to_chars",
    "delete_columns",
    "delete_rows",
    "drawing",
    "evaluate_auto_filter",
    "evaluate_cf_formula",
    "evaluate_cf_range",
    "evaluate_formula",
    "evaluate_formula_array",
    "get_array_formula",
    "get_auto_filter",
    "get_auto_filter_xml",
    "get_cell_format",
    "get_cell_phonetic",
    "get_cell_phonetic_properties",
    "get_cell_phonetic_runs",
    "get_cell_rect_pt",
    "get_cell_show_phonetic",
    "get_cell_xf_index",
    "get_cells_in_range",
    "get_column_width_pt",
    "get_comment",
    "get_comment_visible",
    "get_data_table",
    "get_dependents",
    "get_display_text",
    "get_effective_style",
    "get_format_defaults",
    "get_formula",
    "get_formula_r1c1",
    "get_header_footer",
    "get_header_footer_xml",
    "get_hyperlink",
    "get_invalid_cells",
    "get_lambda_text",
    "get_merges_in_range",
    "get_page_margins",
    "get_page_margins_xml",
    "get_page_setup",
    "get_page_setup_xml",
    "get_pagination",
    "get_precedents",
    "get_print_area",
    "get_print_options",
    "get_print_options_xml",
    "get_print_titles",
    "get_range_style_bases",
    "get_range_style_bases_excluding",
    "get_row_height_pt",
    "get_selection",
    "get_sheet_pr_xml",
    "get_sheet_protection",
    "get_spill_info",
    "get_threaded_comment",
    "get_value",
    "get_values",
    "get_view",
    "get_width_model",
    "insert_columns",
    "insert_rows",
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
    "list_threaded_comments",
    "list_validations",
    "picture",
    "pivot",
    "recalc_viewport",
    "remap_range_style",
    "remove_auto_filter",
    "remove_cf_rule",
    "remove_column_break",
    "remove_comment",
    "remove_data_table",
    "remove_drawing",
    "remove_hyperlink",
    "remove_merge",
    "remove_pivot",
    "remove_protected_range",
    "remove_row_break",
    "remove_sparkline_group",
    "remove_threaded_comment",
    "remove_validation",
    "restore_drawing",
    "restore_range_style",
    "set_array_formula",
    "set_auto_filter",
    "set_auto_filter_xml",
    "set_blank",
    "set_bool",
    "set_cell_format",
    "set_cell_phonetic",
    "set_cell_phonetic_properties",
    "set_cell_phonetic_runs",
    "set_cell_show_phonetic",
    "set_cell_xf_index",
    "set_column_collapsed",
    "set_column_hidden",
    "set_column_outline",
    "set_column_style",
    "set_column_width",
    "set_comment",
    "set_comment_visible",
    "set_error",
    "set_format_defaults",
    "set_formula",
    "set_header_footer",
    "set_header_footer_xml",
    "set_number",
    "set_page_margins",
    "set_page_margins_xml",
    "set_page_setup",
    "set_page_setup_xml",
    "set_print_area",
    "set_print_options",
    "set_print_options_xml",
    "set_print_titles",
    "set_range_xf_index",
    "set_row_auto_height",
    "set_row_collapsed",
    "set_row_height",
    "set_row_hidden",
    "set_row_outline",
    "set_row_style",
    "set_selection",
    "set_sheet_pr_xml",
    "set_sheet_protection",
    "set_text",
    "set_value",
    "set_values",
    "set_view",
    "shape",
    "slicer",
    "snapshot_range_style",
    "update_protected_range",
    "update_sparkline_group",
    "update_threaded_comment",
    "validate_value",
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
    "add_data_table",
    "add_hyperlink",
    "add_hyperlink_range",
    "add_picture",
    "add_sheet",
    "chart",
    "copy_sheet",
    "drawing",
    "get_active_sheet",
    "get_cell_format",
    "get_cell_phonetic_runs",
    "get_formula",
    "get_theme_colors",
    "get_theme_fonts",
    "get_values",
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
    "picture",
    "pivot",
    "pivot_cache",
    "remove_drawing",
    "remove_sheet",
    "set_active_sheet",
    "set_array_formula",
    "set_cell_format",
    "set_comment",
    "set_formula",
    "set_pinned_now",
    "set_theme_colors",
    "set_theme_script_font",
    "set_value",
    "set_values",
    "set_view",
    "shape",
    "slicer",
    "table",
    "update_protected_range",
    "update_sparkline_group",
    "update_threaded_comment",
)
CELL_ARGS = {
    "Sheet.add_hyperlink": ((0, 1),),
    "Sheet.clear_cell_xf_index": ((0, 1),),
    "Sheet.evaluate_cf_formula": ((0, 1), (2, 3)),
    "Sheet.evaluate_formula": ((0, 1),),
    "Sheet.evaluate_formula_array": ((0, 1),),
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
    "Sheet.get_dependents": ((0, 1),),
    "Sheet.get_display_text": ((0, 1),),
    "Sheet.get_effective_style": ((0, 1),),
    "Sheet.get_formula": ((0, 1),),
    "Sheet.get_formula_r1c1": ((0, 1),),
    "Sheet.get_hyperlink": ((0, 1),),
    "Sheet.get_lambda_text": ((0, 1),),
    "Sheet.get_precedents": ((0, 1),),
    "Sheet.get_spill_info": ((0, 1),),
    "Sheet.get_value": ((0, 1),),
    "Sheet.remove_comment": ((0, 1),),
    "Sheet.remove_data_table": ((0, 1),),
    "Sheet.remove_hyperlink": ((0, 1),),
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
    "Sheet.set_number": ((0, 1),),
    "Sheet.set_text": ((0, 1),),
    "Sheet.set_value": ((0, 1),),
    "Sheet.set_values": ((0, 1),),
    "Sheet.validate_value": ((0, 1),),
}
RECT_ARGS = {
    "Sheet.add_hyperlink_range": (0,),
    "Sheet.add_merge": (0,),
    "Sheet.evaluate_cf_range": (0,),
    "Sheet.get_cell_rect_pt": (0,),
    "Sheet.get_cells_in_range": (0,),
    "Sheet.get_merges_in_range": (0,),
    "Sheet.get_range_style_bases": (0,),
    "Sheet.get_range_style_bases_excluding": (0,),
    "Sheet.get_values": (0,),
    "Sheet.list_cells": (0,),
    "Sheet.recalc_viewport": (0,),
    "Sheet.remap_range_style": (0,),
    "Sheet.remove_merge": (0,),
    "Sheet.set_cell_format": (0,),
    "Sheet.set_print_area": (0,),
    "Sheet.set_range_xf_index": (0,),
    "Sheet.snapshot_range_style": (0,),
}
SPAN_ARGS = {
    "Sheet.clear_column_style": ((0, "col", "span"),),
    "Sheet.clear_column_width": ((0, "col", "span"),),
    "Sheet.clear_row_height": ((0, "row", "span"),),
    "Sheet.clear_row_style": ((0, "row", "span"),),
    "Sheet.delete_columns": ((0, "col", "start_count"),),
    "Sheet.delete_rows": ((0, "row", "start_count"),),
    "Sheet.insert_columns": ((0, "col", "start_count"),),
    "Sheet.insert_rows": ((0, "row", "start_count"),),
    "Sheet.set_column_collapsed": ((0, "col", "span"),),
    "Sheet.set_column_hidden": ((0, "col", "span"),),
    "Sheet.set_column_outline": ((0, "col", "span"),),
    "Sheet.set_column_style": ((0, "col", "span"),),
    "Sheet.set_column_width": ((0, "col", "span"),),
    "Sheet.set_row_auto_height": ((0, "row", "span"),),
    "Sheet.set_row_collapsed": ((0, "row", "span"),),
    "Sheet.set_row_height": ((0, "row", "span"),),
    "Sheet.set_row_hidden": ((0, "row", "span"),),
    "Sheet.set_row_outline": ((0, "row", "span"),),
    "Sheet.set_row_style": ((0, "row", "span"),),
}
# </facade-tables>

_LAST_ROW = 1048575
_LAST_COL = 16383
# Largest rectangle :meth:`Sheet.get_values` reads.
_GET_VALUES_CELL_CAP = 1 << 22
_BLANK = Value(ValueKind.BLANK)

# An argument slot the caller left empty; ``None`` is a real argument value.
_MISSING: Any = object()

_F = TypeVar("_F", bound=Callable[..., Any])
_COORDINATE_FORMS: Dict[str, inspect.Signature] = {}


def _coordinate_form(stub: _F) -> _F:
    """Record ``stub``, a method's coordinate-form overload, as that method's ``inspect`` signature."""
    _COORDINATE_FORMS[stub.__name__] = inspect.signature(stub)
    return stub


def _span_form(stub: _F) -> _F:
    """Record ``stub``, a method's number-pair overload of a row or column span, as its ``inspect`` signature."""
    _COORDINATE_FORMS[stub.__name__] = inspect.signature(stub)
    return stub


def _bind(args: Tuple[Any, ...], kwargs: Dict[str, Any], names: Sequence[str]) -> List[Any]:
    """``args`` and ``kwargs`` bound to the parameters ``names``, every one required."""
    if len(args) > len(names):
        raise TypeError(f"formulon: takes at most {len(names)} positional arguments")
    bound = dict(zip(names, args))
    for key, value in kwargs.items():
        if key not in names:
            raise TypeError(f"formulon: unexpected keyword argument {key!r}")
        if key in bound:
            raise TypeError(f"formulon: got multiple values for argument {key!r}")
        bound[key] = value
    missing = [name for name in names if name not in bound]
    if missing:
        raise TypeError(f"formulon: missing argument {missing[0]!r}")
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
        raise TypeError(f"formulon: takes at most {len(layout)} positional arguments with cells as {form}")
    rows: List[Any] = [_MISSING] * len(pairs)
    cols: List[Any] = [_MISSING] * len(pairs)
    rest: List[Any] = [_MISSING] * len(names)
    for (kind, index), arg in zip(layout, args):
        target = rows if kind == "row" else cols if kind == "col" else rest
        target[index] = arg
    if any(v is _MISSING for v in rows) or (not a1 and any(v is _MISSING for v in cols)):
        raise TypeError("formulon: each cell is positional, a (row, col) pair or one A1 string")
    if any(isinstance(v, str) != a1 for v in rows) or any(isinstance(v, str) for v in cols):
        raise TypeError("formulon: give every cell as a (row, col) pair or every cell as one A1 string")
    for key, value in kwargs.items():
        if key not in names:
            raise TypeError(f"formulon: unexpected keyword argument {key!r}")
        index = names.index(key)
        if rest[index] is not _MISSING:
            raise TypeError(f"formulon: got multiple values for argument {key!r}")
        rest[index] = value
    for index, spec in enumerate(rest_names):
        if rest[index] is _MISSING:
            if isinstance(spec, str):
                raise TypeError(f"formulon: missing argument {spec!r}")
            rest[index] = spec[1]
    cells: List[Union[Tuple[Any, Any], str]] = list(rows) if a1 else list(zip(rows, cols))
    return cells, rest


def _geometry_mode(mode: Any) -> int:
    """The C ordinal of ``mode``, checked to be a :class:`GeometryMode` member.

    Raises:
      TypeError: when ``mode`` is not an integer.
      ValueError: when ``mode`` is not a :class:`GeometryMode` member.
    """
    if isinstance(mode, bool) or not isinstance(mode, numbers.Integral):
        raise TypeError(f"formulon: mode must be a GeometryMode, got {type(mode).__name__}")
    try:
        return int(GeometryMode(mode))
    except ValueError:
        raise ValueError(f"formulon: mode must be a GeometryMode, got {mode!r}") from None


def _drawing_anchor_fields(anchor: DrawingAnchor) -> Dict[str, int]:
    """The ABI fields of ``anchor``."""
    return _image_anchor_fields(
        anchor.anchor_kind,
        anchor.edit_as,
        anchor.row,
        anchor.col,
        anchor.row_offset_emu,
        anchor.col_offset_emu,
        anchor.width_emu,
        anchor.height_emu,
    )


def _flag(value: Any, name: str) -> int:
    """``value`` as a C 0/1 flag; ``TypeError`` unless it is a ``bool``."""
    if not isinstance(value, bool):
        raise TypeError(f"formulon: {name} must be a bool, got {type(value).__name__}")
    return 1 if value else 0


def _cell_input(value: Any) -> Tuple[str, Tuple[Any, ...]]:
    """The setter :meth:`Sheet.set_value` calls for ``value`` and the arguments after the cell.

    Raises:
      TypeError: for a type no cell holds, or a :class:`Value` of kind
        ``ARRAY``, ``REF`` or ``LAMBDA``.
      ValueError: for an integer too large to convert to a float.
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
            raise TypeError(f"formulon: a Value of kind {kind.name} cannot be stored in a cell")
    if value is None:
        return "set_blank", ()
    if isinstance(value, bool):
        return "set_bool", (value,)
    if isinstance(value, numbers.Real):
        try:
            number = float(value)
        except OverflowError as exc:
            raise ValueError("formulon: the integer is too large for a float") from exc
        if not math.isfinite(number):
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT, _diagnostic_override=(f"{number!r} cannot be stored in a cell", "")
            )
        return "set_number", (number,)
    if isinstance(value, str):
        return "set_text", (value,)
    if isinstance(value, ExcelError):
        return "set_error", (value.code,)
    raise TypeError(f"formulon: cannot store a {type(value).__name__} in a cell")


# ---------------------------------------------------------------------------
# Cell formats: ``Sheet.get_cell_format`` / ``Sheet.set_cell_format`` over the
# style records. tools/api/cell_format_cases.json pins the behaviour for the
# Python and JS facades alike.
# ---------------------------------------------------------------------------

# How each key of each group of a format patch is checked: ``(kind, nullable,
# argument)``; a nullable key takes ``None`` to clear the attribute. The keys
# are the patch type's keys and the read record's fields.
_FORMAT_FLAG = ("boolean", True, None)
_FORMAT_BOOLEAN = ("boolean", False, None)
_FORMAT_COLOR = ("color", False, None)
_FORMAT_ARGB = ("argb", False, None)
_FORMAT_SIDE = ("side", False, None)
_FORMAT_FIELDS: Dict[str, Dict[str, Tuple[str, bool, Any]]] = {
    "font": {
        "name": ("string", False, None),
        "size": ("size", True, None),
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
        "color_argb": _FORMAT_ARGB,
    },
    "fill": {
        "pattern": ("enum", False, FillPattern),
        "fg": _FORMAT_COLOR,
        "bg": _FORMAT_COLOR,
        "fg_argb": _FORMAT_ARGB,
        "bg_argb": _FORMAT_ARGB,
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
    "side": {"style": ("enum", False, BorderStyle), "color": _FORMAT_COLOR, "color_argb": _FORMAT_ARGB},
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

# The ``CellXfInput`` field each alignment key writes, and the presence flag it sets when it has one.
_ALIGNMENT_XF: Dict[str, Tuple[str, Optional[str]]] = {
    "horizontal": ("horizontal_align", "has_horizontal_align"),
    "vertical": ("vertical_align", "has_vertical_align"),
    "wrap_text": ("wrap_text", "has_wrap_text"),
    "justify_last_line": ("justify_last_line", "has_justify_last_line"),
    "shrink_to_fit": ("shrink_to_fit", None),
    "indent": ("indent", None),
    "text_rotation": ("text_rotation", None),
    "reading_order": ("reading_order", None),
    "relative_indent": ("relative_indent", None),
}
_OPTIONAL_XF_FIELDS = ("text_rotation", "indent", "shrink_to_fit", "reading_order", "relative_indent")
_APPLY_FLAGS = {
    "number_format": "apply_number_format",
    "font": "apply_font",
    "fill": "apply_fill",
    "border": "apply_border",
    "alignment": "apply_alignment",
    "protection": "apply_protection",
}
_COLOR_TEXT = re.compile(r"#(?:[0-9a-fA-F]{6}|[0-9a-fA-F]{8})")
_INT32_MAX = 0x7FFFFFFF


def _out_of_range(label: str, value: Any) -> FormulonError:
    return FormulonError(
        StatusCode.INVALID_ARGUMENT, _diagnostic_override=(f"{label} is out of range", f"value={value}")
    )


def _format_count(value: Any, label: str) -> int:
    """``value`` as a non-negative integer: ``TypeError`` for a non-number, ``ValueError`` for a fraction or negative."""
    if isinstance(value, bool) or not isinstance(value, numbers.Real):
        raise TypeError(f"formulon: {label} must be an int, got {type(value).__name__}")
    if not isinstance(value, numbers.Integral) and not (math.isfinite(value) and float(value).is_integer()):
        raise ValueError(f"formulon: {label} must be a non-negative integer, got {value!r}")
    if value < 0:
        raise ValueError(f"formulon: {label} must be a non-negative integer, got {value!r}")
    return int(value)


def _format_color(value: Any, label: str) -> Color:
    """A ``ColorInput`` as a :class:`Color`: ``"#RRGGBB"`` (opaque) or ``"#AARRGGBB"``, or a :class:`Color`."""
    if isinstance(value, str):
        if not _COLOR_TEXT.fullmatch(value):
            raise TypeError(f'formulon: {label} must be "#RRGGBB" or "#AARRGGBB", got {value!r}')
        digits = value[1:] if len(value) == 9 else "FF" + value[1:]
        return Color(kind=ColorKind.RGB, rgb=int(digits, 16))
    if not isinstance(value, Color):
        raise TypeError(f"formulon: {label} must be a colour string or a Color, got {type(value).__name__}")
    if _format_count(value.kind, f"{label}.kind") > max(ColorKind):
        raise _out_of_range(f"{label}.kind", value.kind)
    return value


def _format_value(spec: Tuple[str, bool, Any], value: Any, label: str) -> Any:
    """One checked key of a format patch."""
    kind, nullable, arg = spec
    if value is None:
        if nullable:
            return None
        raise TypeError(f"formulon: {label} cannot be None")
    if kind == "string":
        if not isinstance(value, str):
            raise TypeError(f"formulon: {label} must be a str, got {type(value).__name__}")
        return value
    if kind == "boolean":
        _flag(value, label)
        return value
    if kind == "gradient":
        if not isinstance(value, GradientFill):
            raise TypeError(f"formulon: {label} must be a GradientFill, got {type(value).__name__}")
        return value
    if kind == "color":
        return _format_color(value, label)
    if kind == "side":
        return _format_group("side", value, label)
    if kind == "size":
        if isinstance(value, bool) or not isinstance(value, numbers.Real):
            raise TypeError(f"formulon: {label} must be a number, got {type(value).__name__}")
        if not math.isfinite(value) or value < 0:
            raise ValueError(f"formulon: {label} must be finite and not negative, got {value!r}")
        return float(value)
    if kind == "argb":
        if _format_count(value, label) > 0xFFFFFFFF:
            raise ValueError(f"formulon: {label} is outside uint32 range")
        return int(value)
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
        raise TypeError(f"formulon: {label} must be an int, got {type(value).__name__}")
    if not isinstance(value, numbers.Integral) and not (math.isfinite(value) and float(value).is_integer()):
        raise ValueError(f"formulon: {label} must be an integer, got {value!r}")
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
        raise TypeError(f"formulon: {label} must be a mapping or a record, got {type(value).__name__}")
    fields = _FORMAT_FIELDS[group]
    out: Dict[str, Any] = {}
    for key, item in items.items():
        if key not in fields:
            raise TypeError(f"formulon: {label} has no field {key!r}")
        out[key] = _format_value(fields[key], item, f"{label}.{key}")
    return out


def _check_cell_format_patch(given: Dict[str, Any]) -> Dict[str, Any]:
    """Every given argument of :meth:`Sheet.set_cell_format`, checked before anything is written."""
    out: Dict[str, Any] = {}
    for key, value in given.items():
        if value is UNCHANGED:
            continue
        if key == "number_format":
            if not isinstance(value, str):
                raise TypeError(f"formulon: number_format must be a str, got {type(value).__name__}")
            if value == "":
                raise FormulonError(StatusCode.INVALID_ARGUMENT, _diagnostic_override=("number_format is empty", ""))
            out[key] = value
        else:
            out[key] = _format_group(key, value, key)
    return out


def _merge_color(target: Dict[str, Any], patch: Dict[str, Any], color_key: str, argb_key: str) -> None:
    """Write the colour pair of ``patch`` onto ``target``; an RGB colour also sets the literal ARGB, so it wins."""
    if argb_key in patch:
        target[argb_key] = patch[argb_key]
    if color_key in patch:
        color = patch[color_key]
        target[color_key] = color
        if color.kind == ColorKind.RGB:
            target[argb_key] = color.rgb


def _merge_font(current: Font, patch: Dict[str, Any]) -> Font:
    changes = {key: value for key, value in patch.items() if key not in ("color", "color_argb")}
    _merge_color(changes, patch, "color", "color_argb")
    return dataclasses.replace(current, **changes)


def _merge_fill(current: Fill, patch: Dict[str, Any]) -> Fill:
    """Setting a colour without a pattern makes the fill solid (Excel's fill colour); a fill patch drops a
    gradient unless it names one."""
    changes: Dict[str, Any] = {"gradient": patch.get("gradient")}
    if "pattern" in patch:
        changes["pattern"] = patch["pattern"]
    _merge_color(changes, patch, "fg", "fg_argb")
    _merge_color(changes, patch, "bg", "bg_argb")
    colored = "fg" in patch or "bg" in patch
    if colored and "pattern" not in patch and current.pattern == FillPattern.NONE:
        changes["pattern"] = FillPattern.SOLID
    return dataclasses.replace(current, **changes)


def _merge_border(current: Border, patch: Dict[str, Any]) -> Border:
    changes: Dict[str, Any] = {}
    for key, value in patch.items():
        if key in ("diagonal_up", "diagonal_down"):
            changes[key] = value
            continue
        side: Dict[str, Any] = {}
        if "style" in value:
            side["style"] = value["style"]
        _merge_color(side, value, "color", "color_argb")
        changes[key] = dataclasses.replace(getattr(current, key), **side)
    return dataclasses.replace(current, **changes)


def _cell_xf_input(xf: CellXf) -> Dict[str, Any]:
    """The :class:`CellXfInput` fields that rebuild ``xf`` as ``add_cell_xf`` reads them.

    An absent attribute (``None``) keeps its input default with its presence
    flag false; an optional alignment attribute is copied only when present.
    """
    fields: Dict[str, Any] = {
        "font_index": xf.font_index,
        "fill_index": xf.fill_index,
        "border_index": xf.border_index,
        "num_fmt_id": xf.num_fmt_id,
        "horizontal_align": HorizontalAlign.GENERAL if xf.horizontal_align is None else xf.horizontal_align,
        "has_horizontal_align": xf.horizontal_align is not None,
        "vertical_align": VerticalAlign.BOTTOM if xf.vertical_align is None else xf.vertical_align,
        "has_vertical_align": xf.vertical_align is not None,
        "wrap_text": bool(xf.wrap_text),
        "has_wrap_text": xf.wrap_text is not None,
        "justify_last_line": bool(xf.justify_last_line),
        "has_justify_last_line": xf.justify_last_line is not None,
        "apply_number_format": xf.apply_number_format,
        "apply_font": xf.apply_font,
        "apply_fill": xf.apply_fill,
        "apply_border": xf.apply_border,
        "apply_alignment": xf.apply_alignment,
        "apply_protection": xf.apply_protection,
        "quote_prefix": xf.quote_prefix,
        "has_protection": xf.locked is not None or xf.hidden is not None,
        "locked": True if xf.locked is None else xf.locked,
        "hidden": False if xf.hidden is None else xf.hidden,
        "xf_id": xf.xf_id,
    }
    for key in _OPTIONAL_XF_FIELDS:
        if getattr(xf, key) is not None:
            fields[key] = getattr(xf, key)
    fields["has_alignment"] = any(
        fields[flag]
        for flag in ("has_horizontal_align", "has_vertical_align", "has_wrap_text", "has_justify_last_line")
    ) or any(key in fields for key in _OPTIONAL_XF_FIELDS)
    return fields


def _patched_xf(workbook: Workbook, xf: CellXf, patch: Dict[str, Any], num_fmt_id: Optional[int]) -> CellXfInput:
    """``xf`` with ``patch`` merged in, as the :class:`CellXfInput` of the new record."""
    fields = _cell_xf_input(xf)
    for group in patch:
        fields[_APPLY_FLAGS[group]] = True
    if num_fmt_id is not None:
        fields["num_fmt_id"] = num_fmt_id
    if "font" in patch:
        fields["font_index"] = workbook.add_font(_merge_font(workbook.get_font(xf.font_index), patch["font"]))
    if "fill" in patch:
        fields["fill_index"] = workbook.add_fill(_merge_fill(workbook.get_fill(xf.fill_index), patch["fill"]))
    if "border" in patch:
        fields["border_index"] = workbook.add_border(
            _merge_border(workbook.get_border(xf.border_index), patch["border"])
        )
    if "alignment" in patch:
        for key, value in patch["alignment"].items():
            field, flag = _ALIGNMENT_XF[key]
            fields[field] = value
            if flag is not None:
                fields[flag] = True
        fields["has_alignment"] = True
    if "protection" in patch:
        fields["locked"] = patch["protection"].get("locked", fields["locked"])
        fields["hidden"] = patch["protection"].get("hidden", fields["hidden"])
        fields["has_protection"] = True
    return CellXfInput(**fields)


def _cell_format_of(workbook: Workbook, xf: CellXf) -> CellFormat:
    """The :class:`CellFormat` of the cell XF record ``xf``."""
    code = None if xf.num_fmt_id == 0 else workbook.get_num_fmt_string(xf.num_fmt_id)
    return CellFormat(
        number_format="General" if code is None else code,
        font=workbook.get_font(xf.font_index),
        fill=workbook.get_fill(xf.fill_index),
        border=workbook.get_border(xf.border_index),
        alignment=Alignment(
            horizontal=HorizontalAlign.GENERAL if xf.horizontal_align is None else HorizontalAlign(xf.horizontal_align),
            vertical=VerticalAlign.BOTTOM if xf.vertical_align is None else VerticalAlign(xf.vertical_align),
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


class Sheet:
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
    A method taking a column or row span (``first, last`` or ``col, count``
    / ``row, count``) also takes it as one A1 span string in its place:
    ``"B"`` or ``"B:D"`` for columns, ``"3"`` or ``"3:5"`` for rows
    (``set_column_width("B:D", 12)``); the whole sheet suits either axis.
    Inserting or deleting rows or columns takes ``(start, count)``
    (``insert_rows(row, count)``); every other row or column span is the
    closed interval ``(first, last)``.

    To read cells, use :meth:`get_values` for a dense grid of values,
    :meth:`list_cells` to enumerate the cells that hold something, and
    :meth:`get_cells_in_range` to page through a large range.

    ``sheet["A1"]`` and ``sheet[row, col]`` return the cell's value as
    :meth:`Value.to_python` converts it; ``sheet["A1"] = v`` and
    ``sheet[row, col] = v`` store ``v`` as :meth:`set_value` does. A range
    key raises :class:`InvalidArgumentError`; an ``int`` or a slice raises
    ``TypeError``.

    Pivot tables and drawing objects are handles too: :meth:`pivot`,
    :meth:`drawing` and the per-kind accessors return the same
    :class:`Pivot` or :class:`Drawing` object for the same pivot or object
    every time.

    Reads return the values of the last full :meth:`Workbook.recalc`.
    """

    __slots__ = ("_workbook", "_position", "_name", "_pivots", "_drawings")

    def __init__(self, workbook: Workbook, position: int, name: str) -> None:
        self._workbook = workbook
        # The current 0-based position, or None once the sheet is removed.
        self._position: Optional[int] = position
        self._name = name
        # One handle per pivot runtime id and per drawing object id, never dropped.
        self._pivots: Dict[int, Pivot] = {}
        self._drawings: Dict[int, Drawing] = {}

    @property
    def workbook(self) -> Workbook:
        """The workbook this sheet belongs to."""
        return self._workbook

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
            raise FormulonError(StatusCode.NOT_FOUND, _diagnostic_override=("the sheet was removed", ""))
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
            _check(LIB.fm_workbook_rename_sheet(h, _uint(index, "sheet_index"), name_ptr))
        self._name = self._workbook._sheet_name(index)

    def move(self, to_index: int) -> None:
        """Move the sheet to ``to_index`` (the position after it is taken out).

        Every :class:`Sheet` handle keeps pointing at its sheet.
        """
        from_index = self._index()
        h = self._workbook._require()
        handles = self._workbook._sheet_handles()
        _check(LIB.fm_workbook_move_sheet(h, _uint(from_index, "from_index"), _uint(to_index, "to_index")))
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
                _diagnostic_override=(f"{rect!r} is not an unqualified A1 range", ""),
            )
        return ref.rect

    def _span(self, text: str, axis: str) -> Tuple[int, int]:
        """The first and last column (``axis`` ``"col"``) or row (``"row"``) the A1 span ``text`` covers.

        A column span is ``"B"`` or ``"B:D"``, a row span ``"3"`` or
        ``"3:5"``; the whole sheet (``"A:XFD"``, ``"1:1048576"``) is both.
        """
        ref = self._workbook.parse_reference(text if ":" in text else f"{text}:{text}")
        rect = None if ref is None or ref.sheet_name is not None else ref.rect
        if rect is not None and axis == "col" and rect.first_row == 0 and rect.last_row == _LAST_ROW:
            return rect.first_col, rect.last_col
        if rect is not None and axis == "row" and rect.first_col == 0 and rect.last_col == _LAST_COL:
            return rect.first_row, rect.last_row
        noun = "whole-column" if axis == "col" else "whole-row"
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT,
            _diagnostic_override=(f"{text!r} is not a {noun} A1 span without a sheet name", ""),
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
                raise TypeError("formulon: got multiple values for argument 'span'")
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
            rect = ref.rect
            if rect.first_row == rect.last_row and rect.first_col == rect.last_col:
                return rect.first_row, rect.first_col
        raise FormulonError(
            StatusCode.INVALID_ARGUMENT,
            _diagnostic_override=(f"{a1!r} is not a single-cell A1 reference", ""),
        )

    def _key(self, key: Union[str, Tuple[int, int]]) -> Tuple[int, int]:
        """``(row, col)`` from a ``sheet[...]`` key: one A1 cell or a ``(row, col)`` pair."""
        if isinstance(key, str):
            return self._cell(key)
        if isinstance(key, tuple) and len(key) == 2 and all(isinstance(k, numbers.Integral) for k in key):
            return key[0], key[1]
        raise TypeError(f"formulon: a sheet key is an A1 string or a (row, col) pair, got {key!r}")

    def __getitem__(self, key: Union[str, Tuple[int, int]]) -> Union[None, float, bool, str, ExcelError, Value]:
        row, col = self._key(key)
        return self.get_value(row, col).to_python()

    def __setitem__(
        self, key: Union[str, Tuple[int, int]], value: Union[None, bool, int, float, str, ExcelError, Value]
    ) -> None:
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
            h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), float(value)
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
            h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), 1 if value else 0
        )
        _check(status)

    @overload
    @_coordinate_form
    def set_error(self, row: int, col: int, error_code: ErrorCode) -> None: ...

    @overload
    def set_error(self, a1: str, error_code: ErrorCode) -> None: ...

    def set_error(self, *args: Any, **kwargs: Any) -> None:
        """Store an Excel cell error at ``(row, col)``.

        ``row`` and ``col`` are 0-based. ``error_code`` is an
        :class:`ErrorCode` member, not a display literal; the same member
        comes back out through :attr:`Value.error_code`.
        """
        row, col, error_code = self._cell_args("set_error", args, kwargs, ("error_code",))
        sheet = self._index()
        h = self._workbook._require()
        status = LIB.fm_sheet_set_error(
            h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), _sint(error_code, "error")
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
        :meth:`set_formula`).
        """
        row, col, value = self._cell_args("set_text", args, kwargs, ("value",))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            text_ptr = scope.cstr(value)
            status = LIB.fm_sheet_set_text(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), text_ptr
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
        status = LIB.fm_sheet_set_blank(h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"))
        _check(status)

    @overload
    @_coordinate_form
    def set_formula(self, row: int, col: int, formula: str) -> None: ...

    @overload
    def set_formula(self, a1: str, formula: str) -> None: ...

    def set_formula(self, *args: Any, **kwargs: Any) -> None:
        """Store a formula at ``(row, col)``.

        ``row`` and ``col`` are 0-based. The leading ``=`` is optional. The cell keeps its previous cached
        value until the next :meth:`Workbook.recalc`.

        Example::

            ws.set_formula("C1", "=A1+B1")
            wb.recalc()
        """
        row, col, formula = self._cell_args("set_formula", args, kwargs, ("formula",))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            formula_ptr = scope.cstr(formula)
            status = LIB.fm_sheet_set_formula(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), formula_ptr
            )
            _check(status)

    @overload
    @_coordinate_form
    def set_value(
        self,
        row: int,
        col: int,
        value: Union[None, bool, int, float, str, ExcelError, Value],  # noqa: FBT001 -- the value is any cell type; a bool is one of them
    ) -> None: ...

    @overload
    def set_value(
        self,
        a1: str,
        value: Union[None, bool, int, float, str, ExcelError, Value],  # noqa: FBT001 -- the value is any cell type; a bool is one of them
    ) -> None: ...

    def set_value(self, *args: Any, **kwargs: Any) -> None:
        """Store ``value`` at ``(row, col)``, choosing the setter by its type.

        ``None`` clears the cell (:meth:`set_blank`); a ``bool`` is a boolean
        (:meth:`set_bool`); any other real number is a number
        (:meth:`set_number`); a ``str`` is text, never a formula, even with a
        leading ``=`` (:meth:`set_text`); an :class:`ExcelError` is a cell
        error (:meth:`set_error`). A :class:`Value` is stored by its kind.

        Raises:
          TypeError: for any other type, or a :class:`Value` of kind
            ``ARRAY``, ``REF`` or ``LAMBDA``.
          ValueError: for an integer too large to convert to a float.
          InvalidArgumentError: for a NaN or infinite number.

        Example::

            ws.set_value("A1", 42)
            ws.set_value("B1", "Total")
            ws.set_value("C1", None)
        """
        row, col, value = self._cell_args("set_value", args, kwargs, ("value",))
        setter, payload = _cell_input(value)
        getattr(self, setter)(row, col, *payload)

    @overload
    @_coordinate_form
    def set_cell_phonetic(self, row: int, col: int, text: Optional[str]) -> None: ...

    @overload
    def set_cell_phonetic(self, a1: str, text: Optional[str]) -> None: ...

    def set_cell_phonetic(self, *args: Any, **kwargs: Any) -> None:
        """Set the cell's phonetic guide (OOXML ``<rPh>``, furigana).

        ``row`` and ``col`` are 0-based. ``None`` (or ``""``) removes the
        guide. The guide is display metadata attached to the
        cell's string; it does not affect evaluation.
        """
        row, col, text = self._cell_args("set_cell_phonetic", args, kwargs, ("text",))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            text_ptr = scope.cstr(text or "")
            status = LIB.fm_sheet_set_cell_phonetic(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), text_ptr
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

        The runs must be an ordered partition: each needs ``sb <= eb`` and
        must start at or after the previous run's ``eb``.
        """
        row, col, runs = self._cell_args("set_cell_phonetic_runs", args, kwargs, ("runs",))
        sheet = self._index()
        h = self._workbook._require()
        runs = [] if runs is None else runs
        with _ArgScope() as scope:
            ptr = _pack_phonetic_run_array(runs, scope)
            status = LIB.fm_sheet_set_cell_phonetic_runs(
                h,
                _uint(sheet, "sheet_index"),
                _uint(row, "row"),
                _uint(col, "col"),
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
            _uint(sheet, "sheet_index"),
            _uint(row, "row"),
            _uint(col, "col"),
        )
        runs: List[PhoneticRun] = []
        with _ArgScope() as scope:
            run_ptr = scope.struct(S.PHONETIC_RUN)
            for run_index in range(count):
                S.zero_struct(LIB, S.PHONETIC_RUN, run_ptr)
                _check(
                    LIB.fm_sheet_get_cell_phonetic_run(
                        h,
                        _uint(sheet, "sheet_index"),
                        _uint(row, "row"),
                        _uint(col, "col"),
                        _uint(run_index, "run_index"),
                        run_ptr,
                    )
                )
                fields = S.PHONETIC_RUN.unpack(LIB, run_ptr)
                runs.append(
                    PhoneticRun(
                        sb=fields["sb"],
                        eb=fields["eb"],
                        text=LIB.read_cstr(fields["text"]) if fields["text"] else "",
                    )
                )
        return runs

    @overload
    @_coordinate_form
    def set_cell_phonetic_properties(self, row: int, col: int, *, font_id: int, type: int, alignment: int) -> None: ...

    @overload
    def set_cell_phonetic_properties(self, a1: str, *, font_id: int, type: int, alignment: int) -> None: ...

    def set_cell_phonetic_properties(self, *args: Any, **kwargs: Any) -> None:
        """Set how the cell's phonetic guide renders.

        ``row`` and ``col`` are 0-based. Independent of
        :meth:`set_cell_phonetic_runs` in both directions: editing the readings
        does not reset the rendering, and this does not touch the readings.
        Only observable on a cell that has runs, and every value-setting
        method clears both, so call it after the cell's text.
        """
        row, col, font_id, type, alignment = self._cell_args(
            "set_cell_phonetic_properties", args, kwargs, ("font_id", "type", "alignment")
        )
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_cell_phonetic_properties(
                h,
                _uint(sheet, "sheet_index"),
                _uint(row, "row"),
                _uint(col, "col"),
                _uint(font_id, "font_id"),
                _uint(type, "type"),
                _uint(alignment, "alignment"),
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
        annotation reports the all-zero default. ``present`` is true when the
        properties are explicit: a ``<phoneticPr>`` element or
        :meth:`set_cell_phonetic_properties`, an all-zero triple included.
        """
        row, col = self._cell_args("get_cell_phonetic_properties", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            font_ptr = scope.out_ptr()
            type_ptr = scope.out_ptr()
            alignment_ptr = scope.out_ptr()
            _check(
                LIB.fm_sheet_get_cell_phonetic_properties(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(row, "row"),
                    _uint(col, "col"),
                    font_ptr,
                    type_ptr,
                    alignment_ptr,
                )
            )
            return PhoneticProperties(
                font_id=LIB.read_u32(font_ptr),
                type=LIB.read_u32(type_ptr),
                alignment=LIB.read_u32(alignment_ptr),
                present=bool(
                    _read_count(
                        LIB.fm_sheet_has_cell_phonetic_properties,
                        h,
                        _uint(sheet, "sheet_index"),
                        _uint(row, "row"),
                        _uint(col, "col"),
                    )
                ),
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
                _uint(sheet, "sheet_index"),
                _uint(row, "row"),
                _uint(col, "col"),
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
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), 1 if show else 0
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
        return (
            _read_cstr_out(
                LIB.fm_sheet_get_cell_phonetic,
                h,
                _uint(sheet, "sheet_index"),
                _uint(row, "row"),
                _uint(col, "col"),
            )
            or None
        )

    # -- Cell read ---------------------------------------------------------
    @overload
    @_coordinate_form
    def get_value(self, row: int, col: int) -> Value: ...

    @overload
    def get_value(self, a1: str) -> Value: ...

    def get_value(self, *args: Any, **kwargs: Any) -> Value:
        """Read the cached value at ``(row, col)``.

        Reflects the most recent recalc; mutations made after the last
        :meth:`Workbook.recalc` are not visible to dependents until you call
        :meth:`Workbook.recalc` again.

        Example::

            ws.set_value("A1", 42)
            ws.get_value("A1").to_python()  # 42.0
        """
        row, col = self._cell_args("get_value", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            value_ptr = scope.alloc(fm_value_t_size)
            status = LIB.fm_sheet_get_value(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), value_ptr
            )
            _check(status)
            return Value._from_wasm(value_ptr)

    def get_values(self, range: Union[CellRect, str]) -> List[List[Value]]:
        """Read the cached values of ``range`` as a dense row-major grid.

        ``result[r][c]`` is the cell ``r`` rows and ``c`` columns from the
        top-left corner; an empty cell is a ``BLANK`` :class:`Value`, and a
        spilled cell holds its spill value. Values are those of the last full
        :meth:`Workbook.recalc`.

        Raises:
          InvalidArgumentError: when the range holds more than 4,194,304
            (2**22) cells.

        Example::

            grid = ws.get_values("A1:B2")
            grid[0][1].to_python()
        """
        rect = self._rect(range)
        rows = rect.last_row - rect.first_row + 1
        cols = rect.last_col - rect.first_col + 1
        if rows <= 0 or cols <= 0 or rows * cols > _GET_VALUES_CELL_CAP:
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _diagnostic_override=(f"get_values reads at most {_GET_VALUES_CELL_CAP} cells, got {rows}x{cols}", ""),
            )
        grid = [[_BLANK] * cols for _ in builtins.range(rows)]  # `range` is the parameter here
        for cell in self.list_cells(rect):
            grid[cell.row - rect.first_row][cell.col - rect.first_col] = cell.value
        return grid

    @overload
    @_coordinate_form
    def set_values(
        self, row: int, col: int, values: Sequence[Sequence[Union[None, bool, int, float, str, ExcelError, Value]]]
    ) -> None: ...

    @overload
    def set_values(
        self, a1: str, values: Sequence[Sequence[Union[None, bool, int, float, str, ExcelError, Value]]]
    ) -> None: ...

    def set_values(self, *args: Any, **kwargs: Any) -> None:
        """Store a row-major grid of values with ``(row, col)`` as its top-left cell.

        Each element is stored as :meth:`set_value` stores it, so a string
        starting with ``=`` is text. Element types, finiteness, equal row
        lengths and the grid's fit on the sheet are all checked before any
        cell is written; an empty grid writes nothing. The write is not
        atomic: when the engine refuses a cell (one inside an array formula
        or a data table, say), the cells before it stay written.

        Raises:
          TypeError: for a non-grid argument or an element :meth:`set_value`
            rejects by type.
          ValueError: for an integer too large to convert to a float.
          InvalidArgumentError: for a NaN or infinite number, or a grid
            reaching past the last row or column.

        Example::

            ws.set_values("A1", [[1, 2], [3, 4]])
        """
        row, col, values = self._cell_args("set_values", args, kwargs, ("values",))
        if isinstance(values, (str, bytes)):
            raise TypeError("formulon: values is a sequence of rows")
        grid = []
        for line in values:
            if isinstance(line, (str, bytes)):
                raise TypeError("formulon: each row of values is a sequence of cell values")
            grid.append([_cell_input(value) for value in line])
        width = len(grid[0]) if grid else 0
        if any(len(line) != width for line in grid):
            raise TypeError("formulon: every row of values must have the same length")
        if width == 0:
            return
        if row + len(grid) - 1 > _LAST_ROW or col + width - 1 > _LAST_COL:
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _diagnostic_override=(f"a {len(grid)}x{width} grid at ({row}, {col}) runs past the sheet", ""),
            )
        for r, line in enumerate(grid):
            for c, (setter, payload) in enumerate(line):
                getattr(self, setter)(row + r, col + c, *payload)

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
                    _uint(sheet, "sheet_index"),
                    _uint(row, "row"),
                    _uint(col, "col"),
                    _uint(anchor_row, "anchor_row"),
                    _uint(anchor_col, "anchor_col"),
                    formula_ptr,
                    value_ptr,
                )
            )
            return Value._from_wasm(value_ptr).boolean is True

    @overload
    @_coordinate_form
    def evaluate_formula(self, row: int, col: int, formula: str) -> Value: ...

    @overload
    def evaluate_formula(self, a1: str, formula: str) -> Value: ...

    def evaluate_formula(self, *args: Any, **kwargs: Any) -> Value:
        """Evaluate ``formula`` as if entered at ``(row, col)``, without mutating the workbook.

        An array result is reduced to its top-left element; use
        :meth:`evaluate_formula_array` for the whole grid. A cell-level Excel
        error is returned as a :class:`Value` of kind ``ERROR``. A reference
        to the anchor cell itself reads its cached value.

        Example::

            ws.evaluate_formula("A1", "=SUM(B1:B3)").to_python()
        """
        row, col, formula = self._cell_args("evaluate_formula", args, kwargs, ("formula",))
        sheet = self._index()
        return self._workbook._evaluate_formula_top_left(sheet, row, col, formula)

    @overload
    @_coordinate_form
    def evaluate_formula_array(self, row: int, col: int, formula: str) -> List[List[Value]]: ...

    @overload
    def evaluate_formula_array(self, a1: str, formula: str) -> List[List[Value]]: ...

    def evaluate_formula_array(self, *args: Any, **kwargs: Any) -> List[List[Value]]:
        """Evaluate ``formula`` as if entered at ``(row, col)`` and
        return the whole multi-cell result, without mutating the workbook.

        Unlike a scalar evaluation (which reduces an array to its top-left
        element), a dynamic-array formula such as ``=SEQUENCE(2,3)`` yields
        the full ``rows x cols`` grid as a nested list in row-major order
        (``result[r][c]``); a scalar result such as ``=1+2`` is returned as
        a ``1 x 1`` grid (``[[Value(...)]]``).

        The evaluation is read-only: no cell value is written and no spill
        is committed. Cell-level Excel errors surface as
        :class:`Value` entries with ``kind == ValueKind.ERROR``; only
        host-side failures (a closed workbook, an out-of-range sheet) raise
        :class:`FormulonError`.

        Args:
          row: 0-based anchor row (drives ``ROW()`` and relative refs).
          col: 0-based anchor column (drives ``COLUMN()``).
          formula: the formula text, with or without a leading ``=``.

        Returns:
          A ``rows x cols`` nested list of :class:`Value`.
        """
        row, col, formula = self._cell_args("evaluate_formula_array", args, kwargs, ("formula",))
        sheet = self._index()
        return self._workbook._evaluate_formula_cells(sheet, row, col, formula, top_left_only=False)

    # -- Iteration ---------------------------------------------------------
    def list_cells(self, range: Union[CellRect, str, None] = None) -> List[Cell]:
        """Return every populated cell in ``range``, or on the whole sheet when it is ``None``.

        Cells come in row-major order; each carries its stored formula and
        cached value, as :meth:`get_cells_in_range` pages them.
        """
        rect = (
            CellRect(first_row=0, first_col=0, last_row=_LAST_ROW, last_col=_LAST_COL)
            if range is None
            else self._rect(range)
        )
        cells: List[Cell] = []
        cursor: Optional[int] = None
        while True:
            page = self.get_cells_in_range(rect, cursor)
            cells.extend(page.cells)
            if page.next_cursor is None:
                return cells
            cursor = page.next_cursor

    @overload
    @_span_form
    def insert_rows(self, row: int, count: int) -> None: ...

    @overload
    def insert_rows(self, span: str) -> None: ...

    def insert_rows(self, *args: Any, **kwargs: Any) -> None:
        """Insert ``count`` rows at ``row`` on the sheet.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them or a table past the sheet bound; the workbook is unchanged.
        """
        row, count = self._span_args("insert_rows", args, kwargs, ("row", "count"))
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_insert_rows(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(count, "count")))

    @overload
    @_span_form
    def delete_rows(self, row: int, count: int) -> None: ...

    @overload
    def delete_rows(self, span: str) -> None: ...

    def delete_rows(self, *args: Any, **kwargs: Any) -> None:
        """Delete ``count`` rows starting at ``row`` on the sheet.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them or a table past the sheet bound; the workbook is unchanged.
        """
        row, count = self._span_args("delete_rows", args, kwargs, ("row", "count"))
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_delete_rows(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(count, "count")))

    @overload
    @_span_form
    def insert_columns(self, col: int, count: int) -> None: ...

    @overload
    def insert_columns(self, span: str) -> None: ...

    def insert_columns(self, *args: Any, **kwargs: Any) -> None:
        """Insert ``count`` columns at ``col`` on the sheet.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them or a table past the sheet bound; the workbook is unchanged.
        """
        col, count = self._span_args("insert_columns", args, kwargs, ("col", "count"))
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_insert_columns(h, _uint(sheet, "sheet"), _uint(col, "col"), _uint(count, "count")))

    @overload
    @_span_form
    def delete_columns(self, col: int, count: int) -> None: ...

    @overload
    def delete_columns(self, span: str) -> None: ...

    def delete_columns(self, *args: Any, **kwargs: Any) -> None:
        """Delete ``count`` columns starting at ``col`` on the sheet.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them or a table past the sheet bound; the workbook is unchanged.
        """
        col, count = self._span_args("delete_columns", args, kwargs, ("col", "count"))
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_delete_columns(h, _uint(sheet, "sheet"), _uint(col, "col"), _uint(count, "count")))

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
            _check(LIB.fm_sheet_recalc_viewport(h, _uint(sheet, "sheet"), vp, out))
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
            _uint(sheet, "sheet_index"),
            _uint(row, "row"),
            _uint(col, "col"),
        )

    # -- Merges ------------------------------------------------------------
    def add_merge(self, range: Union[CellRect, str]) -> None:
        """Append a merge range to the sheet."""
        merge = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_cell_rect_array([merge], scope)
            _check(LIB.fm_sheet_add_merge(h, _uint(sheet, "sheet"), ptr))

    def remove_merge(self, range: Union[CellRect, str]) -> None:
        """Remove every merge on the sheet overlapping ``range``."""
        merge = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_cell_rect_array([merge], scope)
            _check(LIB.fm_sheet_remove_merge(h, _uint(sheet, "sheet"), ptr))

    def clear_merges(self) -> None:
        """Drop every merge range on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_merges(h, _uint(sheet, "sheet")))

    def list_merges(self) -> List[CellRect]:
        """Return every merge range on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(LIB.fm_sheet_get_merge_count, h, _uint(sheet, "sheet"))
        out: List[CellRect] = []
        # One scratch block for the whole walk; see list_defined_names.
        with _ArgScope() as scope:
            ptr = scope.struct(S.CELL_RECT)
            for i in range(n):
                S.zero_struct(LIB, S.CELL_RECT, ptr)
                _check(LIB.fm_sheet_get_merge_at(h, _uint(sheet, "sheet"), _uint(i, "index"), ptr))
                out.append(_decode_cell_rect_at(ptr))
        return out

    # -- Hyperlinks --------------------------------------------------------
    @overload
    @_coordinate_form
    def add_hyperlink(
        self,
        row: int,
        col: int,
        *,
        target: Optional[str],
        location: Optional[str] = None,
        display: Optional[str] = None,
        tooltip: Optional[str] = None,
    ) -> None: ...

    @overload
    def add_hyperlink(
        self,
        a1: str,
        *,
        target: Optional[str],
        location: Optional[str] = None,
        display: Optional[str] = None,
        tooltip: Optional[str] = None,
    ) -> None: ...

    def add_hyperlink(
        self,
        *args: Any,
        target: Optional[str],
        location: Optional[str] = None,
        display: Optional[str] = None,
        tooltip: Optional[str] = None,
        **kwargs: Any,
    ) -> None:
        """Append a hyperlink on the cell ``(row, col)``.

        ``target`` is the external address (URL, file path); ``None`` makes a
        link inside the workbook, which then needs ``location`` (a reference
        such as ``"Sheet2!A1"`` or a defined name). ``location`` beside a
        ``target`` is the fragment within it.

        Raises:
          TypeError: when ``target`` and ``location`` are both ``None``.
        """
        row, col = self._cell_args("add_hyperlink", args, kwargs, ())
        self._add_hyperlink(CellRect(row, col, row, col), target, location, display, tooltip)

    def add_hyperlink_range(
        self,
        range: Union[CellRect, str],
        *,
        target: Optional[str],
        location: Optional[str] = None,
        display: Optional[str] = None,
        tooltip: Optional[str] = None,
    ) -> None:
        """Append a hyperlink covering the cells of ``range``; the arguments read as in :meth:`add_hyperlink`."""
        rect = self._rect(range)
        self._add_hyperlink(rect, target, location, display, tooltip)

    def _add_hyperlink(
        self,
        rect: CellRect,
        target: Optional[str],
        location: Optional[str],
        display: Optional[str],
        tooltip: Optional[str],
    ) -> None:
        if target is None and location is None:
            raise TypeError("formulon: a hyperlink needs a target or a location")
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.HYPERLINK)
            S.HYPERLINK.pack(LIB, ptr, {"rect": _cell_rect_fields(rect)})
            S.write_str_field(LIB, ptr, S.HYPERLINK, "target", target, scope)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "location", location, scope)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "display", display, scope)
            S.write_str_field(LIB, ptr, S.HYPERLINK, "tooltip", tooltip, scope)
            _check(LIB.fm_sheet_add_hyperlink(h, _uint(sheet, "sheet"), ptr))

    @overload
    @_coordinate_form
    def remove_hyperlink(self, row: int, col: int) -> None: ...

    @overload
    def remove_hyperlink(self, a1: str) -> None: ...

    def remove_hyperlink(self, *args: Any, **kwargs: Any) -> None:
        """Remove every hyperlink anchored at ``(row, col)`` on the sheet."""
        row, col = self._cell_args("remove_hyperlink", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_hyperlink(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col")))

    def clear_hyperlinks(self) -> None:
        """Drop every hyperlink on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_hyperlinks(h, _uint(sheet, "sheet")))

    def list_hyperlinks(self) -> List[Hyperlink]:
        """Return every hyperlink on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(LIB.fm_sheet_get_hyperlink_count, h, _uint(sheet, "sheet"))
        out: List[Hyperlink] = []
        # One scratch block for the whole walk; see list_defined_names.
        with _ArgScope() as scope:
            ptr = scope.struct(S.HYPERLINK)
            for i in range(n):
                S.zero_struct(LIB, S.HYPERLINK, ptr)
                _check(LIB.fm_sheet_get_hyperlink_at(h, _uint(sheet, "sheet"), _uint(i, "index"), ptr))
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
            (_uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col")),
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
                LIB.fm_sheet_get_comment(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), present, ptr)
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
    def set_comment(self, row: int, col: int, comment: Comment) -> None: ...

    @overload
    def set_comment(self, a1: str, comment: Comment) -> None: ...

    def set_comment(self, *args: Any, **kwargs: Any) -> None:
        """Insert or replace the comment (note) at ``(row, col)``.

        An empty ``comment.text`` writes an empty note, as Excel keeps one;
        :meth:`remove_comment` removes a note.
        """
        row, col, comment = self._cell_args("set_comment", args, kwargs, ("comment",))
        if not isinstance(comment, Comment):
            raise TypeError(f"formulon: comment must be a Comment, got {type(comment).__name__}")
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            author_ptr = scope.nonempty_cstr(comment.author)
            text_ptr = scope.cstr(comment.text)
            _check(
                LIB.fm_sheet_set_comment(
                    h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), author_ptr, text_ptr
                )
            )

    @overload
    @_coordinate_form
    def remove_comment(self, row: int, col: int) -> None: ...

    @overload
    def remove_comment(self, a1: str) -> None: ...

    def remove_comment(self, *args: Any, **kwargs: Any) -> None:
        """Remove the comment (note) at ``(row, col)``.

        Raises:
          NotFoundError: when the cell has no comment.
        """
        row, col = self._cell_args("remove_comment", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_comment(h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col")))

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
            for index in range(_read_count(LIB.fm_sheet_get_comment_count, h, _uint(sheet, "sheet"))):
                S.zero_struct(LIB, S.COMMENT, ptr)
                _check(LIB.fm_sheet_get_comment_at(h, _uint(sheet, "sheet"), _uint(index, "index"), ptr))
                d = S.COMMENT.unpack(LIB, ptr)
                out.append(
                    SheetComment(
                        row=d["row"], col=d["col"], author=LIB.read_cstr(d["author"]), text=LIB.read_cstr(d["text"])
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
            out = scope.out_ptr()
            status = LIB.fm_sheet_get_comment_visible(
                h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), out
            )
            if status == _STATUS_NOT_FOUND:
                return None
            _check(status)
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
                h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), 1 if visible else 0
            )
        )

    _PROTECTED_RANGE_TEXT = (
        "name",
        "password",
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
        return _read_count(LIB.fm_sheet_get_protected_range_count, h, _uint(sheet, "sheet"))

    def _protected_range_at(self, protected_range_index: int) -> ProtectedRange:
        """Return the ``protected_range_index``-th protected range of the sheet in file order."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PROTECTED_RANGE)
            _check(
                LIB.fm_sheet_get_protected_range(
                    h, _uint(sheet, "sheet"), _uint(protected_range_index, "protected_range_index"), ptr
                )
            )
            d = S.PROTECTED_RANGE.unpack(LIB, ptr)
            return ProtectedRange(
                spin_count=d["spin_count"],
                ranges=_decode_cell_rect_array(d["ranges"], d["range_count"]),
                **{fld: LIB.read_cstr(d[fld]) for fld in self._PROTECTED_RANGE_TEXT},
            )

    def _protected_range_call(
        self, fn, leading: tuple, protected: ProtectedRange, *, returns_index: bool = False, keep_ranges: bool = False
    ) -> Optional[int]:
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PROTECTED_RANGE)
            # NULL with a zero count keeps the stored rectangles; an empty list needs a non-NULL pointer to stay empty.
            ranges_ptr = 0
            if not keep_ranges:
                ranges_ptr = _pack_cell_rect_array(protected.ranges, scope) or scope.struct_array(S.CELL_RECT, 1)
            S.PROTECTED_RANGE.pack(
                LIB,
                ptr,
                {
                    "spin_count": _uint(protected.spin_count, "spin_count"),
                    "ranges": ranges_ptr,
                    "range_count": 0 if keep_ranges else _uint(len(protected.ranges), "range_count"),
                },
            )
            for fld in self._PROTECTED_RANGE_TEXT:
                S.write_str_field(LIB, ptr, S.PROTECTED_RANGE, fld, getattr(protected, fld), scope)
            if not returns_index:
                _check(fn(h, *leading, ptr))
                return None
            out = scope.out_ptr()
            _check(fn(h, *leading, ptr, out))
            return LIB.read_u32(out)

    def add_protected_range(self, protected: ProtectedRange) -> int:
        """Append a protected range and return its index; ``name`` must be non-empty and unique on the sheet."""
        sheet = self._index()
        index = self._protected_range_call(
            LIB.fm_sheet_add_protected_range,
            (_uint(sheet, "sheet"),),
            protected,
            returns_index=True,
        )
        assert index is not None
        return index

    def _protected_range_index(self, protected_range: Union[int, str]) -> int:
        """The index ``protected_range`` addresses: an index, or a protected range's name."""
        return resolve_ref(
            protected_range,
            lambda: [self._protected_range_at(i).name for i in range(self._protected_range_count())],
            "protected range",
        )

    def update_protected_range(
        self,
        protected_range: Union[int, str],
        *,
        name: Union[str, Unchanged] = UNCHANGED,
        ranges: Union[Sequence[CellRect], Unchanged] = UNCHANGED,
        password: Union[str, Unchanged] = UNCHANGED,
        algorithm_name: Union[str, Unchanged] = UNCHANGED,
        hash_value: Union[str, Unchanged] = UNCHANGED,
        salt_value: Union[str, Unchanged] = UNCHANGED,
        spin_count: Union[int, Unchanged] = UNCHANGED,
        security_descriptor: Union[str, Unchanged] = UNCHANGED,
    ) -> None:
        """Update ``protected_range``, an index or a protected range's name; an omitted argument keeps its value.

        The arguments are the :class:`ProtectedRange` fields; an empty string
        clears a text field. Omitting ``ranges`` keeps the stored rectangles,
        including any the reader could not parse.

        Raises:
          TypeError: for ``None`` on any argument.
          NotFoundError: when no protected range has the name ``protected_range``.
          IndexOutOfRangeError: when ``protected_range`` is an index past the last one.
          InvalidArgumentError: when ``ranges`` is empty.
        """
        given = {
            key: value
            for key, value in (
                ("name", name),
                ("ranges", ranges),
                ("password", password),
                ("algorithm_name", algorithm_name),
                ("hash_value", hash_value),
                ("salt_value", salt_value),
                ("spin_count", spin_count),
                ("security_descriptor", security_descriptor),
            )
            if value is not UNCHANGED
        }
        for key, value in given.items():
            if value is None:
                raise TypeError(f"formulon: {key} cannot be None")
        index = self._protected_range_index(protected_range)
        sheet = self._index()
        protected = dataclasses.replace(self._protected_range_at(index), **given)
        self._protected_range_call(
            LIB.fm_sheet_update_protected_range,
            (_uint(sheet, "sheet"), _uint(index, "protected_range_index")),
            protected,
            keep_ranges="ranges" not in given,
        )

    def remove_protected_range(self, protected_range: Union[int, str]) -> None:
        """Remove ``protected_range``, an index or a protected range's name.

        Raises:
          NotFoundError: when no protected range has the name ``protected_range``.
          IndexOutOfRangeError: when ``protected_range`` is an index past the last one.
        """
        index = self._protected_range_index(protected_range)
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_protected_range(h, _uint(sheet, "sheet"), _uint(index, "protected_range_index")))

    # -- Data validations --------------------------------------------------
    def list_validations(self) -> List[DataValidation]:
        """Every data validation of the sheet, in order."""
        return [self._validation_at(i) for i in range(self._validation_count())]

    def _validation_count(self) -> int:
        """Return the number of data-validation rules on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(LIB.fm_sheet_get_validation_count, h, _uint(sheet, "sheet"))

    def _validation_at(self, validation_index: int) -> DataValidation:
        """Read the ``validation_index``-th data-validation rule on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.DATA_VALIDATION)
            _check(
                LIB.fm_sheet_get_validation(h, _uint(sheet, "sheet"), _uint(validation_index, "validation_index"), ptr)
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
            ranges=ranges,
            type=d["type"],
            op=d["op"],
            error_style=d["error_style"],
            allow_blank=bool(d["allow_blank"]),
            show_input_message=bool(d["show_input_message"]),
            show_error_message=bool(d["show_error_message"]),
            show_dropdown=bool(d["show_dropdown"]),
            formula1=LIB.read_cstr(d["formula1"]) or None,
            formula2=LIB.read_cstr(d["formula2"]) or None,
            error_title=LIB.read_cstr(d["error_title"]) or None,
            error_message=LIB.read_cstr(d["error_message"]) or None,
            prompt_title=LIB.read_cstr(d["prompt_title"]) or None,
            prompt_message=LIB.read_cstr(d["prompt_message"]) or None,
        )

    def add_validation(self, validation: DataValidationInput) -> int:
        """Append a data-validation rule to the sheet; return its index."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.DATA_VALIDATION)
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
            _check(LIB.fm_sheet_add_validation(h, _uint(sheet, "sheet"), ptr, out))
            return LIB.read_u32(out)

    def remove_validation(self, validation_index: int) -> None:
        """Remove the validation rule at ``validation_index`` on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_validation(h, _uint(sheet, "sheet"), _uint(validation_index, "validation_index")))

    def clear_validations(self) -> None:
        """Drop every validation rule on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_validations(h, _uint(sheet, "sheet")))

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

    def get_sheet_protection(self) -> SheetProtection:
        """Read the sheet's ``<sheetProtection>`` flags."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHEET_PROTECTION)
            _check(LIB.fm_sheet_get_sheet_protection(h, _uint(sheet, "sheet_index"), ptr))
            d = S.SHEET_PROTECTION.unpack(LIB, ptr)
            return SheetProtection(
                enabled=bool(d["enabled"]),
                algorithm_name=LIB.read_cstr(d["algorithm_name"]) or None,
                hash_value=LIB.read_cstr(d["hash_value"]) or None,
                salt_value=LIB.read_cstr(d["salt_value"]) or None,
                spin_count=d["spin_count"],
                legacy_password=LIB.read_cstr(d["legacy_password"]) or None,
                **{flag: bool(d[flag]) for flag in self._PROTECT_FLAGS},
            )

    def set_sheet_protection(self, protection: Optional[SheetProtection]) -> None:
        """Replace the sheet's ``<sheetProtection>`` flags wholesale; ``None`` removes the element."""
        if protection is None:
            protection = SheetProtection()
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHEET_PROTECTION)
            values: Dict[str, int] = {
                "enabled": 1 if protection.enabled else 0,
                "spin_count": _uint(protection.spin_count, "spin_count"),
            }
            for flag in self._PROTECT_FLAGS:
                values[flag] = 1 if getattr(protection, flag) else 0
            S.SHEET_PROTECTION.pack(LIB, ptr, values)
            for fld in ("algorithm_name", "hash_value", "salt_value", "legacy_password"):
                S.write_str_field(LIB, ptr, S.SHEET_PROTECTION, fld, getattr(protection, fld), scope)
            _check(LIB.fm_sheet_set_sheet_protection(h, _uint(sheet, "sheet_index"), ptr))

    # -- Geometry / display ------------------------------------------------
    @overload
    @_span_form
    def clear_row_height(self, first: int, last: int) -> None: ...

    @overload
    def clear_row_height(self, span: str) -> None: ...

    def clear_row_height(self, *args: Any, **kwargs: Any) -> None:
        """Remove the height override on ``[first, last]`` (back to the sheet default)."""
        first, last = self._span_args("clear_row_height", args, kwargs, ("first", "last"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_clear_row_height(h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"))
        )

    @overload
    @_span_form
    def clear_column_width(self, first: int, last: int) -> None: ...

    @overload
    def clear_column_width(self, span: str) -> None: ...

    def clear_column_width(self, *args: Any, **kwargs: Any) -> None:
        """Remove the width override on ``[first, last]`` (back to the sheet default)."""
        first, last = self._span_args("clear_column_width", args, kwargs, ("first", "last"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_clear_column_width(h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"))
        )

    def get_format_defaults(self) -> SheetFormatDefaults:
        """Read the sheet's default column / row metrics (``<sheetFormatPr>``)."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHEET_FORMAT_DEFAULTS)
            _check(LIB.fm_sheet_get_format_defaults(h, _uint(sheet, "sheet_index"), ptr))
            d = S.SHEET_FORMAT_DEFAULTS.unpack(LIB, ptr)
            return SheetFormatDefaults(
                default_col_width=d["default_col_width"] if d["has_default_col_width"] else None,
                default_row_height=d["default_row_height"] if d["has_default_row_height"] else None,
                base_col_width=d["base_col_width"],
                custom_height=bool(d["custom_height"]),
                zero_height=bool(d["zero_height"]),
                effective_default_col_width=d["effective_default_col_width"],
                effective_default_row_height=d["effective_default_row_height"],
            )

    def set_format_defaults(
        self,
        *,
        default_col_width: Union[float, None, Unchanged] = UNCHANGED,
        default_row_height: Union[float, None, Unchanged] = UNCHANGED,
        base_col_width: Union[float, Unchanged] = UNCHANGED,
        custom_height: Union[bool, Unchanged] = UNCHANGED,
        zero_height: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the sheet's default column / row metrics; an omitted argument keeps its value.

        Arguments are the :class:`SheetFormatDefaults` fields.
        ``default_col_width`` / ``default_row_height`` given as ``None`` clear
        the stored value; the other arguments cannot be cleared.

        Raises:
          TypeError: when ``base_col_width`` is ``None`` or a flag is not a ``bool``.
        """
        if base_col_width is None:
            raise TypeError("formulon: base_col_width cannot be None")
        for name, value in (("custom_height", custom_height), ("zero_height", zero_height)):
            if value is not UNCHANGED:
                _flag(value, name)
        given = (default_col_width, default_row_height, base_col_width, custom_height, zero_height)
        if any(value is UNCHANGED for value in given):
            stored = self.get_format_defaults()
            if default_col_width is UNCHANGED:
                default_col_width = stored.default_col_width
            if default_row_height is UNCHANGED:
                default_row_height = stored.default_row_height
            if base_col_width is UNCHANGED:
                base_col_width = stored.base_col_width
            if custom_height is UNCHANGED:
                custom_height = stored.custom_height
            if zero_height is UNCHANGED:
                zero_height = stored.zero_height
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHEET_FORMAT_DEFAULTS)
            S.SHEET_FORMAT_DEFAULTS.pack(
                LIB,
                ptr,
                {
                    "default_col_width": 0.0 if default_col_width is None else float(default_col_width),
                    "default_row_height": 0.0 if default_row_height is None else float(default_row_height),
                    "base_col_width": float(base_col_width),
                    "has_default_col_width": 0 if default_col_width is None else 1,
                    "has_default_row_height": 0 if default_row_height is None else 1,
                    "custom_height": 1 if custom_height else 0,
                    "zero_height": 1 if zero_height else 0,
                },
            )
            _check(LIB.fm_sheet_set_format_defaults(h, _uint(sheet, "sheet_index"), ptr))

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
            _check(LIB.fm_sheet_get_cell_rect_pt(h, _uint(sheet, "sheet_index"), rp, _sint(mode, "mode"), ptr))
            return RectPt(**S.RECT_PT.unpack(LIB, ptr))

    def get_column_width_pt(self, col: int, mode: GeometryMode = GeometryMode.DISPLAY) -> float:
        """Return the effective width of ``col`` in points (0 when hidden).

        Raises:
          ValueError: when ``mode`` is not a :class:`GeometryMode` member.
        """
        mode_c = _geometry_mode(mode)
        sheet = self._index()
        h = self._workbook._require()
        return self._read_f64_out(
            LIB.fm_sheet_get_column_width_pt, h, _uint(sheet, "sheet_index"), _uint(col, "col"), mode_c
        )

    def get_row_height_pt(self, row: int) -> float:
        """Return the effective height of ``row`` in points (0 when hidden)."""
        sheet = self._index()
        h = self._workbook._require()
        return self._read_f64_out(LIB.fm_sheet_get_row_height_pt, h, _uint(sheet, "sheet_index"), _uint(row, "row"))

    def get_width_model(self, mode: GeometryMode = GeometryMode.DISPLAY) -> WidthModel:
        """Return the character-unit to point conversion model in ``mode``."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.WIDTH_MODEL)
            _check(LIB.fm_sheet_get_width_model(h, _uint(sheet, "sheet_index"), _sint(mode, "mode"), ptr))
            d = S.WIDTH_MODEL.unpack(LIB, ptr)
            return WidthModel(
                points_per_char=d["points_per_char"],
                padding_pt=d["padding_pt"],
                normal_font_size=d["normal_font_size"],
                calibrated=bool(d["calibrated"]),
                normal_font_name=LIB.read_cstr(d["normal_font_name"]),
                platform=LIB.read_cstr(d["platform"]),
            )

    def convert_column_chars_to_pt(self, chars: float, mode: GeometryMode = GeometryMode.DISPLAY) -> float:
        """Convert a stored column width (padding included) to points, snapped to whole screen pixels.

        Raises:
          ValueError: when ``mode`` is not a :class:`GeometryMode` member.
        """
        mode_c = _geometry_mode(mode)
        sheet = self._index()
        h = self._workbook._require()
        return self._read_f64_out(
            LIB.fm_sheet_convert_column_chars_to_pt, h, _uint(sheet, "sheet_index"), mode_c, float(chars)
        )

    def convert_column_pt_to_chars(self, pt: float, mode: GeometryMode = GeometryMode.DISPLAY) -> float:
        """Convert points to the stored column width Excel writes for that column.

        Raises:
          ValueError: when ``mode`` is not a :class:`GeometryMode` member.
        """
        mode_c = _geometry_mode(mode)
        sheet = self._index()
        h = self._workbook._require()
        return self._read_f64_out(
            LIB.fm_sheet_convert_column_pt_to_chars, h, _uint(sheet, "sheet_index"), mode_c, float(pt)
        )

    def _read_formula_text(self, fn, sheet: int, row: int, col: int) -> Optional[str]:
        h = self._workbook._require()
        text = _read_cstr_out(
            fn,
            h,
            _uint(sheet, "sheet_index"),
            _uint(row, "row"),
            _uint(col, "col"),
        )
        return text or None

    @overload
    @_coordinate_form
    def get_formula(self, row: int, col: int) -> Optional[str]: ...

    @overload
    def get_formula(self, a1: str) -> Optional[str]: ...

    def get_formula(self, *args: Any, **kwargs: Any) -> Optional[str]:
        """Return the stored A1 formula (leading ``=`` included), or ``None`` for a non-formula cell."""
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
        range: Union[CellRect, str],
        cursor: Optional[int] = None,
        limit: Optional[int] = None,
    ) -> CellsInRange:
        """Return one page of the populated cells in ``range`` (row-major).

        Pass ``next_cursor`` back to fetch the next page; it is ``None``
        when the rectangle holds no further populated cell. ``limit`` of
        ``None`` or ``0`` means the engine maximum.
        """
        cell_range = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, cell_range)
            return CellsInRange(
                *self._read_cell_range_page(
                    LIB.fm_sheet_get_cells_in_range,
                    h,
                    _uint(sheet, "sheet_index"),
                    rp,
                    _uint(cursor or 0, "cursor", 64),
                    _uint(limit or 0, "limit"),
                )
            )

    @staticmethod
    def _read_cell_range_page(fn, *args) -> "tuple[List[Cell], Optional[int]]":
        """Call a ``(..., cell_range** out)`` engine function and drain its page."""
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(fn(*args, out))
            handle = LIB.read_u32(out)
        if handle == 0:
            raise FormulonError(
                StatusCode.BINDING_NULL_POINTER,
                _diagnostic_override=("returned OK with a null cell-range handle", ""),
            )
        try:
            count = _read_count(LIB.fm_cell_range_get_cell_count, handle)
            with _ArgScope() as scope:
                row_ptr = scope.out_ptr()
                col_ptr = scope.out_ptr()
                formula_ptr = scope.out_ptr()
                cursor_ptr = scope.out_u64()
                value_ptr = scope.alloc(fm_value_t_size)
                cells: List[Cell] = []
                for i in range(count):
                    LIB.write_bytes(formula_ptr, b"\x00\x00\x00\x00")
                    _check(
                        LIB.fm_cell_range_get_cell_at(handle, _uint(i, "idx"), row_ptr, col_ptr, formula_ptr, value_ptr)
                    )
                    formula_addr = LIB.read_u32(formula_ptr)
                    cells.append(
                        Cell(
                            row=LIB.read_u32(row_ptr),
                            col=LIB.read_u32(col_ptr),
                            formula=LIB.read_cstr(formula_addr) if formula_addr else None,
                            value=Value._from_wasm(value_ptr),
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
        return self._workbook._read_auto_filter(LIB.fm_sheet_get_auto_filter, sheet, "sheet_index")

    def set_auto_filter(self, auto_filter: AutoFilter) -> None:
        """Replace the sheet AutoFilter. Row visibility is unchanged; see :meth:`apply_auto_filter`."""
        sheet = self._index()
        self._workbook._write_auto_filter(LIB.fm_sheet_set_auto_filter, sheet, auto_filter, "sheet_index")

    def remove_auto_filter(self) -> None:
        """Remove the sheet AutoFilter (no-op when absent). Row visibility is unchanged."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_auto_filter(h, _uint(sheet, "sheet_index")))

    def apply_auto_filter(self) -> None:
        """Hide the body rows the sheet AutoFilter excludes and show the rest."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_apply_auto_filter(h, _uint(sheet, "sheet_index")))

    def clear_auto_filter(self) -> None:
        """Drop every column criterion of the sheet AutoFilter and show all its body rows."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_auto_filter(h, _uint(sheet, "sheet_index")))

    def evaluate_auto_filter(self) -> AutoFilterEval:
        """Evaluate the sheet AutoFilter without changing the sheet.

        Body row ``first_row + i`` matches when ``match[i]`` is true.
        """
        sheet = self._index()
        return AutoFilterEval(
            *self._workbook._evaluate_auto_filter(LIB.fm_sheet_evaluate_auto_filter, sheet, "sheet_index")
        )

    # -- Validation checks -----------------------------------------------------
    @overload
    @_coordinate_form
    def validate_value(self, row: int, col: int, value: Value) -> ValidationOutcome: ...

    @overload
    def validate_value(self, a1: str, value: Value) -> ValidationOutcome: ...

    def validate_value(self, *args: Any, **kwargs: Any) -> ValidationOutcome:
        """Check ``value`` against the data-validation rule covering ``(row, col)`` without storing it."""
        row, col, value = self._cell_args("validate_value", args, kwargs, ("value",))
        sheet = self._index()
        h = self._workbook._require()
        _check_value(value)
        with _ArgScope() as scope:
            value_ptr = scope.alloc(fm_value_t_size)
            ptr = scope.struct(S.VALIDATION_OUTCOME)
            _write_value_input(value, value_ptr, scope)
            _check(
                LIB.fm_sheet_validate_value(
                    h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), value_ptr, ptr
                )
            )
            d = S.VALIDATION_OUTCOME.unpack(LIB, ptr)
        has_rule = bool(d["has_rule"])
        return ValidationOutcome(
            valid=bool(d["valid"]),
            rule_index=d["rule_index"] if has_rule else None,
            error_style=ValidationErrorStyle(d["error_style"]) if has_rule else None,
        )

    def get_invalid_cells(self, cursor: Optional[int] = None, limit: Optional[int] = None) -> CellsInRange:
        """Return one page of the populated cells failing their data-validation rule (row-major).

        Paging follows :meth:`get_cells_in_range`: pass ``next_cursor`` back
        for the next page; it is ``None`` once no further invalid cell exists.
        """
        sheet = self._index()
        h = self._workbook._require()
        return CellsInRange(
            *self._read_cell_range_page(
                LIB.fm_sheet_get_invalid_cells,
                h,
                _uint(sheet, "sheet_index"),
                _uint(cursor or 0, "cursor", 64),
                _uint(limit or 0, "limit"),
            )
        )

    # -- Threaded comments and persons -------------------------------------------
    def list_threaded_comments(self) -> List[ThreadedComment]:
        """Return the sheet's threaded comments in file order (openers followed by their replies)."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(LIB.fm_sheet_get_threaded_comment_count, h, _uint(sheet, "sheet_index"))
        with _ArgScope() as scope:
            ptr = scope.struct(S.THREADED_COMMENT)
            out: List[ThreadedComment] = []
            for i in range(n):
                S.zero_struct(LIB, S.THREADED_COMMENT, ptr)
                _check(LIB.fm_sheet_get_threaded_comment_at(h, _uint(sheet, "sheet_index"), _uint(i, "idx"), ptr))
                d = S.THREADED_COMMENT.unpack(LIB, ptr)
                out.append(_threaded_comment_from(d))
        return out

    def add_threaded_comment(self, comment: ThreadedCommentInput) -> None:
        """Add a threaded comment: a new thread when ``parent_id`` is ``None``, else a reply."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.THREADED_COMMENT, 1)
            S.THREADED_COMMENT.pack(
                LIB,
                ptr,
                {
                    "row": _uint(comment.row, "row"),
                    "col": _uint(comment.col, "col"),
                    "done": 1 if comment.done else 0,
                },
            )
            for fld in ("id", "person_id", "created", "text", "parent_id"):
                S.write_str_field(LIB, ptr, S.THREADED_COMMENT, fld, getattr(comment, fld), scope)
            _write_mentions(ptr, S.THREADED_COMMENT, comment.mentions, scope)
            _check(LIB.fm_sheet_add_threaded_comment(h, _uint(sheet, "sheet_index"), ptr))

    def update_threaded_comment(
        self,
        comment_id: str,
        *,
        text: Union[str, Unchanged] = UNCHANGED,
        mentions: Union[Sequence[Mention], Unchanged] = UNCHANGED,
        done: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the text, mentions and resolved state of the comment ``comment_id``; an omitted argument is kept.

        ``done`` applies to a thread's opening comment only.

        Raises:
          TypeError: when an argument is ``None`` or ``done`` is not a ``bool``.
          NotFoundError: when the sheet has no comment ``comment_id``.
          FormulonError: with ``StatusCode.THREADED_COMMENT_INVALID`` when
            ``done`` is given for a reply.
        """
        for name, value in (("text", text), ("mentions", mentions), ("done", done)):
            if value is None:
                raise TypeError(f"formulon: {name} cannot be None")
        values: Dict[str, object] = {"done": -1 if done is UNCHANGED else _flag(done, "done")}
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            id_ptr = scope.cstr(comment_id)
            if text is not UNCHANGED:
                values["text"] = scope.cstr(text)
            if mentions is not UNCHANGED:
                values["mentions"] = _pack_mentions(mentions, scope)
                values["mention_count"] = _uint(len(mentions), "mention_count")
                values["mentions_engaged"] = 1
            ptr = scope.struct(S.THREADED_COMMENT_PATCH)
            S.THREADED_COMMENT_PATCH.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_update_threaded_comment(h, _uint(sheet, "sheet_index"), id_ptr, ptr))

    def get_threaded_comment(self, comment_id: str) -> Optional[ThreadedComment]:
        """Return the threaded comment ``comment_id``, or ``None`` when the sheet has none."""
        sheet = self._index()
        return self._workbook._lookup(
            LIB.fm_sheet_get_threaded_comment,
            (_uint(sheet, "sheet_index"), comment_id),
            S.THREADED_COMMENT,
            _threaded_comment_from,
        )

    def remove_threaded_comment(self, comment_id: str) -> None:
        """Remove the comment ``comment_id``; removing a thread opener removes its whole thread."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            id_ptr = scope.cstr(comment_id)
            _check(LIB.fm_sheet_remove_threaded_comment(h, _uint(sheet, "sheet_index"), id_ptr))

    # -- Drawing objects ---------------------------------------------------
    #
    # Each object is a handle of the class its kind selects (see
    # ``drawing._drawing_for``); the sheet keeps one handle per object id in
    # ``_drawings`` and never drops it, so a restore that puts an id back makes
    # its old handle address the object again.
    def _drawing_handle(self, object_id: int) -> Drawing:
        from .drawing import _drawing_for

        return _drawing_for(self, object_id)

    def _drawing_id(self, ref: Union[int, Drawing], where: str) -> int:
        """The object id ``ref`` names: an id, or a :class:`Drawing` of this sheet."""
        from .drawing import Drawing

        if isinstance(ref, Drawing):
            if ref.sheet is not self:
                raise FormulonError(
                    StatusCode.INVALID_ARGUMENT, _diagnostic_override=(f"{where} is a Drawing of another sheet", "")
                )
            return ref.object_id
        if isinstance(ref, bool) or not isinstance(ref, numbers.Integral):
            raise TypeError(f"formulon: {where} must be an object id or a Drawing, got {type(ref).__name__}")
        return int(ref)

    def _typed_drawing(self, ref: Union[int, Drawing], cls: type, where: str) -> Any:
        """The handle of ``ref`` when it is a ``cls``; ``InvalidArgumentError`` for another kind."""
        drawing = self._drawing_handle(self._drawing_id(ref, where))
        if not isinstance(drawing, cls):
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _diagnostic_override=(f"drawing object {drawing.object_id} is not a {cls.__name__}", ""),
            )
        return drawing

    def _object_ids(self, count_fn: Any, at_fn: Any, layout: Any) -> List[int]:
        """The ``object_id`` of each record a count / at pair of the sheet lists, in order."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(count_fn, h, _uint(sheet, "sheet_index"))
        out: List[int] = []
        with _ArgScope() as scope:
            ptr = scope.struct(layout)
            for i in range(n):
                S.zero_struct(LIB, layout, ptr)
                _check(at_fn(h, _uint(sheet, "sheet_index"), _uint(i, "idx"), ptr))
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
        from .drawing import Picture

        return cast("Picture", self._typed_drawing(object_id, Picture, "object_id"))

    def chart(self, object_id: Union[int, Drawing]) -> Chart:
        """Return the :class:`Chart` handle of ``object_id``; fails as :meth:`picture` does when it is not a chart."""
        from .drawing import Chart

        return cast("Chart", self._typed_drawing(object_id, Chart, "object_id"))

    def shape(self, object_id: Union[int, Drawing]) -> Shape:
        """Return the :class:`Shape` handle of ``object_id`` (a shape, connector or group); fails as :meth:`picture` does otherwise."""
        from .drawing import Shape

        return cast("Shape", self._typed_drawing(object_id, Shape, "object_id"))

    def slicer(self, object_id: Union[int, Drawing]) -> Slicer:
        """Return the :class:`Slicer` handle of ``object_id`` (a table slicer's frame); fails as :meth:`picture` does otherwise."""
        from .drawing import Slicer

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
        from .drawing import Picture

        return [drawing for drawing in self.list_drawings() if isinstance(drawing, Picture)]

    def list_charts(self) -> List[Chart]:
        """Return the handles of the sheet's charts, in document order."""
        from .drawing import Chart

        return [drawing for drawing in self.list_drawings() if isinstance(drawing, Chart)]

    def list_shapes(self) -> List[Shape]:
        """Return the handles of the shapes, connectors and groups of the sheet's drawing in z-order (back first).

        Children of a group are not listed.
        """
        ids = self._object_ids(LIB.fm_sheet_get_shape_count, LIB.fm_sheet_get_shape_at, S.SHAPE)
        return [cast("Shape", self._drawing_handle(object_id)) for object_id in ids]

    def list_slicers(self) -> List[Slicer]:
        """Return the handles of the table slicers the sheet hosts; slicers over a pivot cache are not listed."""
        ids = self._object_ids(LIB.fm_sheet_get_slicer_count, LIB.fm_sheet_get_slicer_at, S.SLICER)
        return [cast("Slicer", self._drawing_handle(object_id)) for object_id in ids]

    def add_picture(
        self,
        data: Union[bytes, bytearray, memoryview],
        *,
        name: str = "",
        description: str = "",
        anchor: Optional[DrawingAnchor] = None,
    ) -> Picture:
        """Insert ``data`` as a picture placed by ``anchor`` and return its handle.

        An empty ``name`` becomes ``Picture <id>``. ``anchor`` ``None`` is a
        one-cell anchor at A1. A ``width_emu`` or ``height_emu`` of ``0``
        means the image's pixel size times 9525. ``edit_as`` is honoured for a
        two-cell anchor only. Existing drawing content is kept.
        """
        if anchor is None:
            anchor = DrawingAnchor()
        sheet = self._index()
        h = self._workbook._require()
        raw = bytes(data)
        with _ArgScope() as scope:
            out = scope.out_ptr()
            buf = scope.buffer(raw)
            ptr = scope.struct_array(S.IMAGE_INSERT, 1)
            S.IMAGE_INSERT.pack(LIB, ptr, {"anchor": _drawing_anchor_fields(anchor)})
            S.write_str_field(LIB, ptr, S.IMAGE_INSERT, "name", name, scope)
            S.write_str_field(LIB, ptr, S.IMAGE_INSERT, "description", description, scope)
            _check(LIB.fm_sheet_add_image(h, _uint(sheet, "sheet_index"), buf, _uint(len(raw), "len"), ptr, out))
            object_id = LIB.read_u32(out)
        return cast("Picture", self._drawing_handle(object_id))

    def add_shape(
        self,
        *,
        anchor_kind: AnchorKind,
        edit_as: AnchorEditAs = AnchorEditAs.TWO_CELL,
        from_row: int = 0,
        from_col: int = 0,
        from_row_offset_emu: int = 0,
        from_col_offset_emu: int = 0,
        to_row: int = 0,
        to_col: int = 0,
        to_row_offset_emu: int = 0,
        to_col_offset_emu: int = 0,
        width_emu: int = 0,
        height_emu: int = 0,
        name: Optional[str] = None,
        preset_geometry: Optional[str] = None,
        text: Optional[str] = None,
        text_box: Optional[bool] = None,
        fill_kind: Optional[ShapePaint] = None,
        fill_rgb: int = 0,
        line_kind: Optional[ShapePaint] = None,
        line_rgb: int = 0,
        line_width_emu: Optional[int] = None,
        rotation: Optional[int] = None,
        hidden: Optional[bool] = None,
    ) -> Shape:
        """Append a shape on top of the sheet's drawing and return its handle.

        A ``None`` argument is not applied. ``preset_geometry`` defaults to ``rect`` and
        ``name`` to ``Shape <id>``; a ``preset_geometry`` naming a line or connector
        geometry makes a connector, which holds no text. A two-cell anchor's
        size is its box (``width_emu`` / ``height_emu`` are ignored); an absolute anchor's
        position is ``from_col_offset_emu`` (x) and ``from_row_offset_emu`` (y). ``fill_kind``
        and ``line_kind`` must be ``SOLID`` (with the rgb) or ``NONE``.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = _pack_shape(
                scope, name, preset_geometry, text, text_box, anchor_kind, edit_as, from_row, from_col, from_row_offset_emu,
                from_col_offset_emu, to_row, to_col, to_row_offset_emu, to_col_offset_emu, width_emu, height_emu, fill_kind, fill_rgb, line_kind,
                line_rgb, line_width_emu, rotation, hidden,
            )  # fmt: skip
            _check(LIB.fm_sheet_add_shape(h, _uint(sheet, "sheet_index"), ptr, out))
            object_id = LIB.read_u32(out)
        return cast("Shape", self._drawing_handle(object_id))

    def add_slicer(
        self,
        table: Union[int, str, Table],
        column: int,
        *,
        from_row: int,
        from_col: int,
        to_row: int,
        to_col: int,
        from_row_offset_emu: int = 0,
        from_col_offset_emu: int = 0,
        to_row_offset_emu: int = 0,
        to_col_offset_emu: int = 0,
        name: Optional[str] = None,
        caption: Optional[str] = None,
        style: Optional[str] = None,
        column_count: Optional[int] = None,
        row_height_emu: Optional[int] = None,
        sort_order: Optional[SlicerSortOrder] = None,
        cross_filter: Optional[SlicerCrossFilter] = None,
    ) -> Slicer:
        """Add a slicer over column ``column`` of ``table`` (an index, a display name or a :class:`Table`); return its handle.

        The frame is the two-cell anchor ``from_*`` .. ``to_*`` (0-based cells
        plus EMU offsets), named as in :meth:`add_shape`. The name and caption
        default to the column name (the name is made unique),
        ``row_height_emu`` to 241300 and ``column_count`` to 1. The table may
        live on another sheet.
        """
        table_index = self._workbook._table_index(table)
        sheet = self._index()
        h = self._workbook._require()
        anchor = (
            from_row, from_col, to_row, to_col,
            from_row_offset_emu, from_col_offset_emu, to_row_offset_emu, to_col_offset_emu,
        )  # fmt: skip
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = _pack_slicer(
                scope, name, caption, style, column_count, row_height_emu, anchor, sort_order, cross_filter
            )
            _check(
                LIB.fm_sheet_add_slicer(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(table_index, "table_index"),
                    _uint(column, "column_index"),
                    ptr,
                    out,
                )
            )
            object_id = LIB.read_u32(out)
        return cast("Slicer", self._drawing_handle(object_id))

    def restore_drawing(self, data: Union[bytes, bytearray, memoryview], *, new_id: bool = False) -> Drawing:
        """Put a ``snapshot()`` capture back and return the object's handle.

        An object with the same id is replaced and the capture returns to its
        recorded list position; a handle of that id addresses it again. With
        ``new_id`` it is added on top as a copy with a fresh id.
        """
        sheet = self._index()
        h = self._workbook._require()
        raw = bytes(data)
        with _ArgScope() as scope:
            buf = scope.buffer(raw)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_restore_drawing(
                    h, _uint(sheet, "sheet_index"), buf, _uint(len(raw), "len"), 1 if new_id else 0, out
                )
            )
            object_id = LIB.read_u32(out)
        return self._drawing_handle(object_id)

    def remove_drawing(self, object_id: Union[int, Drawing]) -> None:
        """Remove drawing object ``object_id`` and the parts only it reaches; other drawing content is kept.

        A shape, connector or group goes with its children; a slicer with its
        frame, and its cache when no other slicer uses it.

        Raises:
          NotFoundError: when the sheet's drawing has no such object.
          InvalidArgumentError: when the object is not a picture, chart, shape
            or table slicer, or ``object_id`` is a :class:`Drawing` of another sheet.
        """
        from .drawing import Chart, Picture, Shape, Slicer

        drawing = self.drawing(object_id)
        sheet = self._index()
        h = self._workbook._require()
        args = (h, _uint(sheet, "sheet_index"), _uint(drawing.object_id, "object_id"))
        if isinstance(drawing, (Picture, Chart)):
            _check(LIB.fm_sheet_remove_drawing(*args))
        elif isinstance(drawing, Shape):
            _check(LIB.fm_sheet_remove_shape(*args))
        elif isinstance(drawing, Slicer):
            _check(LIB.fm_sheet_remove_slicer(*args))
        else:
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _diagnostic_override=(
                    f"drawing object {drawing.object_id} is not a picture, chart, shape or table slicer",
                    "",
                ),
            )

    def add_data_table(
        self,
        rect: Union[CellRect, str],
        *,
        row_input: Union[Tuple[int, int], DataTableInputCell, str, None] = None,
        col_input: Union[Tuple[int, int], DataTableInputCell, str, None] = None,
    ) -> None:
        """Create a data table whose result range is ``rect``.

        The formulas sit in the row above the range and the substitution values in the column left of it (a
        one-variable table), or the other way round; ``row_input`` / ``col_input`` are the ``(row, col)`` input
        cells that receive the values of the row above / the column left. Give one for a one-variable table, both
        for a two-variable one. ``rect`` is a :class:`CellRect` or an unqualified A1 range; an input cell is a
        ``(row, col)`` pair, a :class:`DataTableInputCell` or one A1 cell string. The result cells read 0 until the
        next recalc and refuse writes until the table is removed.
        """
        rect = self._rect(rect)
        inputs = [
            self._cell(cell)
            if isinstance(cell, str)
            else (cell.row, cell.col)
            if isinstance(cell, DataTableInputCell)
            else cell
            for cell in (row_input, col_input)
        ]
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.DATA_TABLE, 1)
            values: Dict[str, object] = {"rect": _cell_rect_fields(rect)}
            for prefix, cell in (("row_input", inputs[0]), ("col_input", inputs[1])):
                if cell is not None:
                    values["has_" + prefix] = 1
                    values[prefix + "_row"] = _uint(cell[0], prefix + "_row")
                    values[prefix + "_col"] = _uint(cell[1], prefix + "_col")
            S.DATA_TABLE.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_add_data_table(h, _uint(sheet, "sheet_index"), ptr))

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
                    h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), present, ptr
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
            rect=_cell_rect_from(d["rect"]),
            two_dimensional=bool(d["two_dimensional"]),
            row_oriented=bool(d["row_oriented"]),
            row_input=_input("row_input"),
            col_input=_input("col_input"),
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
        _check(LIB.fm_sheet_remove_data_table(h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col")))

    def list_sparkline_groups(self) -> List[SparklineGroup]:
        """Every sparkline group of the sheet, in order."""
        return [self._sparkline_group_at(i) for i in range(self._sparkline_group_count())]

    def _sparkline_group_count(self) -> int:
        """Return the number of sparkline groups on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(LIB.fm_sheet_get_sparkline_group_count, h, _uint(sheet, "sheet_index"))

    def _sparkline_group_at(self, group_index: int) -> SparklineGroup:
        """Return the sheet's ``group_index``-th sparkline group in document order with its sparklines."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            group_ptr = scope.struct(S.SPARKLINE_GROUP)
            item_ptr = scope.struct(S.SPARKLINE)
            _check(
                LIB.fm_sheet_get_sparkline_group(h, _uint(sheet, "sheet_index"), _uint(group_index, "idx"), group_ptr)
            )
            d = S.SPARKLINE_GROUP.unpack(LIB, group_ptr)
            group = SparklineGroup(
                type=d["type"],
                color_series=_decode_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_series"][1]),
                color_negative=_decode_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_negative"][1]),
                color_axis=_decode_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_axis"][1]),
                color_markers=_decode_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_markers"][1]),
                color_first=_decode_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_first"][1]),
                color_last=_decode_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_last"][1]),
                color_high=_decode_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_high"][1]),
                color_low=_decode_color(group_ptr + S.SPARKLINE_GROUP.offsets["color_low"][1]),
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
                date_axis_ref=LIB.read_cstr(d["date_axis_ref"]) or None,
                line_weight=d["line_weight"],
                empty_cells_as=d["empty_cells_as"],
            )
            for k in range(d["sparkline_count"]):
                S.zero_struct(LIB, S.SPARKLINE, item_ptr)
                _check(
                    LIB.fm_sheet_get_sparkline(
                        h, _uint(sheet, "sheet_index"), _uint(group_index, "idx"), _uint(k, "idx"), item_ptr
                    )
                )
                item = S.SPARKLINE.unpack(LIB, item_ptr)
                group.sparklines.append(
                    Sparkline(formula=LIB.read_cstr(item["formula"]), location=LIB.read_cstr(item["location"]))
                )
            return group

    def add_sparkline_group(self, group: SparklineGroup) -> int:
        """Append ``group`` with its ``sparklines`` (at least one) to the sheet's worksheet extension list.

        Returns the new group's index.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_sparkline_group(scope, group)
            items = _pack_sparklines(scope, group.sparklines)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_add_sparkline_group(
                    h, _uint(sheet, "sheet_index"), ptr, items, _uint(len(group.sparklines), "count"), out
                )
            )
            return LIB.read_u32(out)

    def update_sparkline_group(
        self,
        group_index: int,
        *,
        type: Union[SparklineType, Unchanged] = UNCHANGED,
        color_series: Union[Color, Unchanged] = UNCHANGED,
        color_negative: Union[Color, Unchanged] = UNCHANGED,
        color_axis: Union[Color, Unchanged] = UNCHANGED,
        color_markers: Union[Color, Unchanged] = UNCHANGED,
        color_first: Union[Color, Unchanged] = UNCHANGED,
        color_last: Union[Color, Unchanged] = UNCHANGED,
        color_high: Union[Color, Unchanged] = UNCHANGED,
        color_low: Union[Color, Unchanged] = UNCHANGED,
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
        date_axis_ref: Union[str, None, Unchanged] = UNCHANGED,
        line_weight: Union[float, Unchanged] = UNCHANGED,
        empty_cells_as: Union[SparklineEmptyCells, Unchanged] = UNCHANGED,
        sparklines: Union[Sequence[Sparkline], Unchanged] = UNCHANGED,
    ) -> None:
        """Update group ``group_index``; an omitted argument keeps the group's value.

        The arguments are the :class:`SparklineGroup` fields. ``None`` clears
        ``manual_min``, ``manual_max`` and ``date_axis_ref`` and is a
        ``TypeError`` for any other argument. Every attribute or child outside
        the model is kept.

        Raises:
          TypeError: for ``None`` on an argument that is not clearable.
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
                ("date_axis_ref", date_axis_ref),
                ("line_weight", line_weight),
                ("empty_cells_as", empty_cells_as),
                ("sparklines", sparklines),
            )
            if value is not UNCHANGED
        }
        for key, value in given.items():
            if value is None and key not in ("manual_min", "manual_max", "date_axis_ref"):
                raise TypeError(f"formulon: {key} cannot be None")
        group = dataclasses.replace(self._sparkline_group_at(group_index), **given)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_sparkline_group(scope, group)
            # A NULL list keeps the stored sparklines.
            items = _pack_sparklines(scope, group.sparklines) if sparklines is not UNCHANGED else 0
            count = len(group.sparklines) if sparklines is not UNCHANGED else 0
            _check(
                LIB.fm_sheet_update_sparkline_group(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(group_index, "idx"),
                    ptr,
                    items,
                    _uint(count, "count"),
                )
            )

    def remove_sparkline_group(self, group_index: int) -> None:
        """Remove sparkline group ``group_index``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_sparkline_group(h, _uint(sheet, "sheet_index"), _uint(group_index, "idx")))

    def add_chart(
        self,
        kind: ChartInsertKind,
        source: str,
        *,
        grouping: ChartGrouping = ChartGrouping.CLUSTERED,
        direction: ChartSeriesDirection = ChartSeriesDirection.AUTO,
        header_row: Optional[bool] = None,
        header_column: Optional[bool] = None,
        title: Optional[str] = None,
        legend: Optional[ChartLegendPosition] = ChartLegendPosition.BOTTOM,
        name: str = "",
        description: str = "",
        anchor: Optional[DrawingAnchor] = None,
    ) -> Chart:
        """Create a chart over the A1 rectangle ``source`` and return its handle.

        ``source`` is sheet-qualified or on the sheet. ``grouping`` applies to
        column, bar, line and area charts. ``header_row`` / ``header_column``
        ``None`` follow Excel's detection. An omitted ``title`` is Excel's
        automatic title and an empty one is no title; ``legend`` is ``None``
        for no legend. The placement reads as
        for :meth:`add_picture` (an empty ``name`` becomes ``Chart <id>``),
        except that a zero ``width_emu`` or ``height_emu`` means Excel's default
        360 x 216 pt frame and ``anchor`` ``None`` is a two-cell anchor at A1.

        Example::

            chart = ws.add_chart(ChartInsertKind.COLUMN, "A1:B5", title="Sales")
        """
        if anchor is None:
            anchor = DrawingAnchor(anchor_kind=AnchorKind.TWO_CELL)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = scope.struct_array(S.CHART_INSERT, 1)
            placement = {
                "name": scope.nonempty_cstr(name),
                "description": scope.nonempty_cstr(description),
                "anchor": _drawing_anchor_fields(anchor),
            }
            S.CHART_INSERT.pack(
                LIB,
                ptr,
                {
                    "kind": _sint(kind, "kind"),
                    "grouping": _sint(grouping, "grouping"),
                    "source": scope.cstr(source),
                    "direction": _sint(direction, "direction"),
                    "header_row": _sint(_header_mode(header_row), "header_row"),
                    "header_column": _sint(_header_mode(header_column), "header_column"),
                    "title": scope.opt_cstr(title),
                    "legend": _sint(_nullable_enum(legend, "legend"), "legend"),
                    "placement": placement,
                },
            )
            _check(LIB.fm_sheet_add_chart(h, _uint(sheet, "sheet_index"), ptr, out))
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
            _check(LIB.fm_sheet_get_merges_in_range(h, _uint(sheet, "sheet_index"), range_ptr, 0, 0, count_ptr))
            total = LIB.read_u32(count_ptr)
            if total == 0:
                return []
            size = S.CELL_RECT.size
            buf = scope.struct_array(S.CELL_RECT, total)
            _check(
                LIB.fm_sheet_get_merges_in_range(
                    h, _uint(sheet, "sheet_index"), range_ptr, buf, _uint(total, "capacity"), count_ptr
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
                    h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), text_ptr, status_ptr
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
            _check(LIB.fm_sheet_get_pagination(h, _uint(sheet, "sheet_index"), out))
            pagination = LIB.read_u32(out)
            if pagination == 0:
                raise FormulonError(
                    StatusCode.BINDING_NULL_POINTER,
                    _diagnostic_override=("returned OK with a null pagination handle", ""),
                )
            try:
                page_count = _read_count(LIB.fm_pagination_get_page_count, pagination)
                range_count = _read_count(LIB.fm_pagination_get_print_area_count, pagination)
                break_count = _read_count(LIB.fm_pagination_get_row_break_count, pagination)
                column_break_count = _read_count(LIB.fm_pagination_get_column_break_count, pagination)
                range_ptr = scope.struct(S.CELL_RECT)
                value_ptr = scope.out_ptr()
                scale_ptr = scope.alloc(8)
                # Sized for the largest pagination detail struct (fm_page_layout_t).
                struct_ptr = scope.struct(S.PAGE_LAYOUT)
                print_area = []
                for i in range(range_count):
                    _check(LIB.fm_pagination_get_print_area_at(pagination, _uint(i, "index"), range_ptr))
                    print_area.append(_decode_cell_rect_at(range_ptr))
                row_breaks = []
                for i in range(break_count):
                    _check(LIB.fm_pagination_get_row_break_at(pagination, _uint(i, "index"), value_ptr))
                    row_breaks.append(LIB.read_u32(value_ptr))
                column_breaks = []
                for i in range(column_break_count):
                    _check(LIB.fm_pagination_get_column_break_at(pagination, _uint(i, "index"), value_ptr))
                    column_breaks.append(LIB.read_u32(value_ptr))
                row_break_manual = []
                for i in range(break_count):
                    _check(LIB.fm_pagination_is_row_break_manual(pagination, _uint(i, "index"), value_ptr))
                    row_break_manual.append(bool(LIB.read_i32(value_ptr)))
                column_break_manual = []
                for i in range(column_break_count):
                    _check(LIB.fm_pagination_is_column_break_manual(pagination, _uint(i, "index"), value_ptr))
                    column_break_manual.append(bool(LIB.read_i32(value_ptr)))
                _check(LIB.fm_pagination_get_scale(pagination, scale_ptr))
                scale = LIB.read_f64(scale_ptr)
                _check(LIB.fm_pagination_get_page_order(pagination, value_ptr))
                page_order = LIB.read_i32(value_ptr)

                def read_struct(layout: S.Struct, fn) -> Dict[str, int]:
                    S.zero_struct(LIB, layout, struct_ptr)
                    _check(fn(pagination, struct_ptr))
                    return layout.unpack(LIB, struct_ptr)

                d = read_struct(S.PAPER_INFO, LIB.fm_pagination_get_paper)
                paper = Paper(d["width_pt"], d["height_pt"], bool(d["landscape"]), bool(d["known"]))
                margins = MarginsPt(**read_struct(S.MARGINS_PT, LIB.fm_pagination_get_margins))
                printable = RectPt(**read_struct(S.RECT_PT, LIB.fm_pagination_get_printable))
                d = read_struct(S.PRINT_TITLES, LIB.fm_pagination_get_print_titles)
                print_titles = _print_titles_from(d)
                pages = []
                for i in range(page_count):
                    S.zero_struct(LIB, S.PAGE_LAYOUT, struct_ptr)
                    _check(LIB.fm_pagination_get_page_at(pagination, _uint(i, "index"), struct_ptr))
                    page = S.PAGE_LAYOUT.unpack(LIB, struct_ptr)
                    page["rect"] = _cell_rect_from(page["rect"])
                    pages.append(PageLayout(**page))
                return Pagination(
                    page_count,
                    print_area,
                    row_breaks,
                    column_breaks,
                    paper,
                    margins,
                    printable,
                    scale,
                    page_order,
                    print_titles,
                    pages,
                    row_break_manual,
                    column_break_manual,
                )
            finally:
                LIB.fm_pagination_destroy(pagination)

    def get_view(self) -> SheetView:
        """Read the full per-sheet view (zoom, freeze, visibility, and the
        display / orientation flags)."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHEET_VIEW)
            _check(LIB.fm_sheet_get_view(h, _uint(sheet, "sheet_index"), ptr))
            d = S.SHEET_VIEW.unpack(LIB, ptr)
            return SheetView(
                zoom_scale=d["zoom_scale"],
                freeze_rows=d["freeze_rows"],
                freeze_cols=d["freeze_cols"],
                visibility=SheetVisibility(d["visibility"]),
                show_grid_lines=bool(d["show_grid_lines"]),
                show_row_col_headers=bool(d["show_row_col_headers"]),
                show_zeros=bool(d["show_zeros"]),
                right_to_left=bool(d["right_to_left"]),
                tab_selected=bool(d["tab_selected"]),
                view_mode=SheetViewMode(d["view_mode"]),
                top_left_row=d["top_left_row"],
                top_left_col=d["top_left_col"],
                pane_top_left_row=d["pane_top_left_row"],
                pane_top_left_col=d["pane_top_left_col"],
                split_x=d["split_x"],
                split_y=d["split_y"],
                show_formulas=bool(d["show_formulas"]),
                show_outline_symbols=bool(d["show_outline_symbols"]),
            )

    def set_view(
        self,
        *,
        zoom_scale: Union[int, Unchanged] = UNCHANGED,
        freeze_rows: Union[int, Unchanged] = UNCHANGED,
        freeze_cols: Union[int, Unchanged] = UNCHANGED,
        split_x: Union[float, Unchanged] = UNCHANGED,
        split_y: Union[float, Unchanged] = UNCHANGED,
        top_left_row: Union[int, Unchanged] = UNCHANGED,
        top_left_col: Union[int, Unchanged] = UNCHANGED,
        pane_top_left_row: Union[int, Unchanged] = UNCHANGED,
        pane_top_left_col: Union[int, Unchanged] = UNCHANGED,
        visibility: Union[SheetVisibility, Unchanged] = UNCHANGED,
        show_grid_lines: Union[bool, Unchanged] = UNCHANGED,
        show_row_col_headers: Union[bool, Unchanged] = UNCHANGED,
        show_zeros: Union[bool, Unchanged] = UNCHANGED,
        show_formulas: Union[bool, Unchanged] = UNCHANGED,
        show_outline_symbols: Union[bool, Unchanged] = UNCHANGED,
        right_to_left: Union[bool, Unchanged] = UNCHANGED,
        tab_selected: Union[bool, Unchanged] = UNCHANGED,
        view_mode: Union[SheetViewMode, Unchanged] = UNCHANGED,
    ) -> None:
        """Update the sheet view; an omitted argument keeps its value.

        Arguments are the :class:`SheetView` fields. ``zoom_scale`` is clamped
        to ``[10, 400]``. ``split_x`` / ``split_y`` are in twips; a split
        replaces any freeze and a freeze any split, and changing either resets
        the pane origin unless ``pane_top_left_*`` is given. One half of a
        pair (``freeze_rows`` without ``freeze_cols``, ``top_left_row``
        without ``top_left_col``...) keeps the other half.
        :attr:`SheetVisibility.VERY_HIDDEN` hides the tab from Excel's Unhide
        dialog as well.

        Fields apply in this order: freeze or split, pane origin, scroll
        origin, display flags, zoom, view mode, direction, tab selection,
        visibility. Hiding the only visible sheet and a pane origin on a view
        left without panes are refused before anything changes; any other
        failure leaves the fields before it applied.

        Raises:
          TypeError: when an argument is ``None`` or a flag is not a ``bool``,
            or when a non-zero freeze and a non-zero split are both given.
          InvalidArgumentError: when hiding the only visible sheet, for a pane
            origin on a view without panes, or for a value the sheet rejects.
        """
        flags = (
            ("show_grid_lines", show_grid_lines, LIB.fm_sheet_set_show_grid_lines),
            ("show_row_col_headers", show_row_col_headers, LIB.fm_sheet_set_show_row_col_headers),
            ("show_zeros", show_zeros, LIB.fm_sheet_set_show_zeros),
            ("show_formulas", show_formulas, LIB.fm_sheet_set_show_formulas),
            ("show_outline_symbols", show_outline_symbols, LIB.fm_sheet_set_show_outline_symbols),
        )
        given = {
            key: value
            for key, value in (
                ("zoom_scale", zoom_scale),
                ("freeze_rows", freeze_rows),
                ("freeze_cols", freeze_cols),
                ("split_x", split_x),
                ("split_y", split_y),
                ("top_left_row", top_left_row),
                ("top_left_col", top_left_col),
                ("pane_top_left_row", pane_top_left_row),
                ("pane_top_left_col", pane_top_left_col),
                ("visibility", visibility),
                ("right_to_left", right_to_left),
                ("tab_selected", tab_selected),
                ("view_mode", view_mode),
                *((key, value) for key, value, _ in flags),
            )
            if value is not UNCHANGED
        }
        for key, value in given.items():
            if value is None:
                raise TypeError(f"formulon: {key} cannot be None")
        for key in ("right_to_left", "tab_selected", *(key for key, _, _ in flags)):
            if key in given and not isinstance(given[key], bool):
                raise TypeError(f"formulon: {key} must be a bool, got {type(given[key]).__name__}")

        def non_zero(*keys: str) -> bool:
            return any(key in given and given[key] != 0 for key in keys)

        if non_zero("freeze_rows", "freeze_cols") and non_zero("split_x", "split_y"):
            raise TypeError("formulon: cannot freeze and split the view at once")
        if not given:
            return
        view = self.get_view()
        merged = {**{f.name: getattr(view, f.name) for f in dataclasses.fields(view)}, **given}
        freeze = merged["freeze_rows"] != view.freeze_rows or merged["freeze_cols"] != view.freeze_cols
        split = merged["split_x"] != view.split_x or merged["split_y"] != view.split_y
        # The panes once applied: a freeze replaces a split and a split a freeze.
        panes = [view.freeze_rows, view.freeze_cols, view.split_x, view.split_y]
        if freeze:
            panes = [merged["freeze_rows"], merged["freeze_cols"], 0, 0]
        if split:
            cleared = merged["split_x"] == 0 and merged["split_y"] == 0
            panes = [panes[0], panes[1], 0, 0] if cleared else [0, 0, merged["split_x"], merged["split_y"]]
        sheet = self._index()
        pane_origin = "pane_top_left_row" in given or "pane_top_left_col" in given
        if pane_origin and all(n == 0 for n in panes):
            raise FormulonError(
                StatusCode.INVALID_ARGUMENT,
                _diagnostic_override=("the view has no panes to set an origin on", f"sheet_index={sheet}"),
            )
        if merged["visibility"] != SheetVisibility.VISIBLE and view.visibility == SheetVisibility.VISIBLE:
            if not any(
                other is not self and other.get_view().visibility == SheetVisibility.VISIBLE for other in self._workbook
            ):
                raise FormulonError(
                    StatusCode.INVALID_ARGUMENT,
                    _diagnostic_override=("cannot hide the only visible sheet", f"sheet_index={sheet}"),
                )
        h = self._workbook._require()
        if freeze:
            _check(
                LIB.fm_sheet_set_freeze(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(merged["freeze_rows"], "freeze_rows"),
                    _uint(merged["freeze_cols"], "freeze_cols"),
                )
            )
        if split:
            _check(
                LIB.fm_sheet_set_split(
                    h, _uint(sheet, "sheet_index"), float(merged["split_x"]), float(merged["split_y"])
                )
            )
        # Changing the panes resets their origin, so a missing partner is read afresh.
        current = self.get_view() if freeze or split else view
        if pane_origin:
            row = given.get("pane_top_left_row", current.pane_top_left_row)
            col = given.get("pane_top_left_col", current.pane_top_left_col)
            _check(
                LIB.fm_sheet_set_pane_top_left(
                    h, _uint(sheet, "sheet_index"), _uint(row, "pane_top_left_row"), _uint(col, "pane_top_left_col")
                )
            )
        if "top_left_row" in given or "top_left_col" in given:
            row = given.get("top_left_row", current.top_left_row)
            col = given.get("top_left_col", current.top_left_col)
            _check(
                LIB.fm_sheet_set_top_left(
                    h, _uint(sheet, "sheet_index"), _uint(row, "top_left_row"), _uint(col, "top_left_col")
                )
            )
        for key, value, setter in flags:
            if key in given:
                _check(setter(h, _uint(sheet, "sheet_index"), 1 if value else 0))
        if "zoom_scale" in given:
            _check(LIB.fm_sheet_set_zoom(h, _uint(sheet, "sheet_index"), _uint(zoom_scale, "zoom_scale")))
        if "view_mode" in given:
            _check(LIB.fm_sheet_set_view_mode(h, _uint(sheet, "sheet_index"), _sint(view_mode, "view_mode")))
        if "right_to_left" in given:
            _check(LIB.fm_sheet_set_right_to_left(h, _uint(sheet, "sheet_index"), 1 if right_to_left else 0))
        if "tab_selected" in given:
            _check(LIB.fm_sheet_set_tab_selected(h, _uint(sheet, "sheet_index"), 1 if tab_selected else 0))
        if "visibility" in given:
            _check(LIB.fm_sheet_set_visibility(h, _uint(sheet, "sheet_index"), _sint(visibility, "visibility")))

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
            _check(LIB.fm_sheet_get_selection(h, _uint(sheet, "sheet_index"), row_ptr, col_ptr, 0, 0, total_ptr))
            total = LIB.read_u32(total_ptr)
            array_ptr = scope.struct_array(S.CELL_RECT, total)
            _check(
                LIB.fm_sheet_get_selection(
                    h, _uint(sheet, "sheet_index"), row_ptr, col_ptr, array_ptr, _uint(total, "capacity"), total_ptr
                )
            )
            size = S.CELL_RECT.size
            ranges = [_decode_cell_rect_at(array_ptr + i * size) for i in range(total)]
            return SheetSelection(active_row=LIB.read_u32(row_ptr), active_col=LIB.read_u32(col_ptr), ranges=ranges)

    def set_selection(self, *, active_row: int, active_col: int, ranges: Sequence[CellRect] = ()) -> None:
        """Replace the active pane's selection; empty ``ranges`` selects the active cell alone."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_cell_rect_array(list(ranges), scope)
            _check(
                LIB.fm_sheet_set_selection(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(active_row, "active_row"),
                    _uint(active_col, "active_col"),
                    ptr,
                    _uint(len(ranges), "count"),
                )
            )

    def get_auto_filter_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`get_auto_filter`.

        Return the sheet's ``<autoFilter>`` element, or ``None`` without one.
        The fragment carries the filter-column criteria, sort state and
        extension payload verbatim, so a host can preserve definitions it
        does not interpret.
        """
        sheet = self._index()
        h = self._workbook._require()
        return (
            _read_cstr_out(
                LIB.fm_sheet_get_auto_filter_xml,
                h,
                _uint(sheet, "sheet_index"),
            )
            or None
        )

    def set_auto_filter_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`set_auto_filter`.

        Replace the sheet's ``<autoFilter>`` element; ``None`` (or ``""``) removes the AutoFilter. Other input must be a
        complete ``<autoFilter ...>`` fragment; the engine preserves filter
        extensions verbatim rather than validating them in detail.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            xml_ptr = scope.cstr(xml or "")
            _check(LIB.fm_sheet_set_auto_filter_xml(h, _uint(sheet, "sheet_index"), xml_ptr))

    def _get_print_xml(self, export: str) -> Optional[str]:
        sheet = self._index()
        h = self._workbook._require()
        return _read_cstr_out(getattr(LIB, export), h, _uint(sheet, "sheet_index")) or None

    def _set_print_xml(self, export: str, xml: Optional[str]) -> None:
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            xml_ptr = scope.cstr(xml or "")
            _check(getattr(LIB, export)(h, _uint(sheet, "sheet_index"), xml_ptr))

    def get_page_setup_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`get_page_setup`.

        Return the sheet's ``<pageSetup>`` element, or ``None`` without one.
        """
        return self._get_print_xml("fm_sheet_get_page_setup_xml")

    def set_page_setup_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`set_page_setup`.

        Replace the sheet's ``<pageSetup>`` element; ``None`` (or ``""``)
        removes it. ``fitToPage`` is not a ``<pageSetup>`` attribute but lives
        in ``<sheetPr>``. A fragment carrying ``r:id`` is rejected
        unless the sheet already has a printerSettings part, because the
        reference would otherwise dangle and Excel would repair the file.
        """
        self._set_print_xml("fm_sheet_set_page_setup_xml", xml)

    def get_page_margins_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`get_page_margins`.

        Return the sheet's ``<pageMargins>`` element, or ``None`` without one.
        """
        return self._get_print_xml("fm_sheet_get_page_margins_xml")

    def set_page_margins_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`set_page_margins`.

        Replace the sheet's ``<pageMargins>`` element; ``None`` (or ``""``) removes it.
        """
        self._set_print_xml("fm_sheet_set_page_margins_xml", xml)

    def get_print_options_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`get_print_options`.

        Return the sheet's ``<printOptions>`` element, or ``None`` without one.
        """
        return self._get_print_xml("fm_sheet_get_print_options_xml")

    def set_print_options_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`set_print_options`.

        Replace the sheet's ``<printOptions>`` element; ``None`` (or ``""``) removes it.
        """
        self._set_print_xml("fm_sheet_set_print_options_xml", xml)

    def get_header_footer_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`get_header_footer`.

        Return the sheet's ``<headerFooter>`` element, or ``None`` without one.
        """
        return self._get_print_xml("fm_sheet_get_header_footer_xml")

    def set_header_footer_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`set_header_footer`.

        Replace the sheet's ``<headerFooter>`` element; ``None`` (or ``""``)
        removes it. Header and footer codes are introduced by ``&``, which lives in XML
        element text and so must arrive escaped: Excel writes
        ``<oddHeader>&amp;C...</oddHeader>``. Pass decoded text to
        :meth:`set_header_footer` instead to let the engine escape it.
        """
        self._set_print_xml("fm_sheet_set_header_footer_xml", xml)

    def get_sheet_pr_xml(self) -> Optional[str]:
        """OOXML fragment for round trips; normally use :meth:`get_page_setup`, which reads its ``fitToPage``.

        Return the sheet's ``<sheetPr>`` element, or ``None`` without one. It
        also carries the tab colour and the VBA code name.
        """
        return self._get_print_xml("fm_sheet_get_sheet_pr_xml")

    def set_sheet_pr_xml(self, xml: Optional[str]) -> None:
        """OOXML fragment for round trips; normally use :meth:`set_page_setup`, which writes its ``fitToPage``.

        Replace the sheet's ``<sheetPr>`` element; ``None`` (or ``""``) removes it.
        """
        self._set_print_xml("fm_sheet_set_sheet_pr_xml", xml)

    def get_print_area(self) -> Optional[List[CellRect]]:
        """Return the print area's rectangles, or ``None`` when none is declared.

        A whole-axis area reads back expanded to the grid edge (``A:D``
        becomes ``A1:D1048576``), because this reports the rectangles the
        paginator resolves rather than the authored shape.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ranges_out = scope.out_ptr()
            count_out = scope.out_ptr()
            _check(LIB.fm_sheet_get_print_area(h, _uint(sheet, "sheet_index"), ranges_out, count_out))
            return _decode_cell_rect_array(LIB.read_u32(ranges_out), LIB.read_u32(count_out)) or None

    def set_print_area(self, ranges: Optional[Sequence[Union[CellRect, str]]]) -> None:
        """Write ``_xlnm.Print_Area`` from one or more rectangles, printed in order.

        A rectangle spanning every row or column is stored as whole columns
        or rows. ``None`` or an empty list removes the name.

        Raises:
          TypeError: when ``ranges`` is a single ``str`` rather than a sequence.
        """
        if isinstance(ranges, str):
            raise TypeError("formulon: ranges must be a sequence of CellRect or A1 ranges, got str")
        rects = [] if ranges is None else [self._rect(r) for r in ranges]
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ranges_ptr = _pack_cell_rect_array(rects, scope)
            _check(
                LIB.fm_sheet_set_print_area(
                    h, _uint(sheet, "sheet_index"), ranges_ptr, _uint(len(rects), "range_count")
                )
            )

    def get_print_titles(self) -> Optional[PrintTitles]:
        """Return the rows and columns repeated on every printed page, or ``None`` when neither repeats."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PRINT_TITLES)
            _check(LIB.fm_sheet_get_print_titles(h, _uint(sheet, "sheet_index"), ptr))
            titles = _print_titles_from(S.PRINT_TITLES.unpack(LIB, ptr))
            return None if titles.rows is None and titles.columns is None else titles

    def set_print_titles(self, titles: Optional[PrintTitles]) -> None:
        """Write the sheet's ``_xlnm.Print_Titles`` defined name.

        The stored formula lists rows before columns and anchors every
        endpoint (``Sheet1!$1:$2,Sheet1!$A:$A``). ``None``, or a
        :class:`PrintTitles` with both axes ``None``, removes the name.
        """
        sheet = self._index()
        h = self._workbook._require()
        if titles is None:
            _check(LIB.fm_sheet_set_print_titles(h, _uint(sheet, "sheet_index"), 0))
            return
        values: Dict[str, int] = {}
        for axis, (flag, first, last) in (
            (titles.rows, ("has_rows", "first_row", "last_row")),
            (titles.columns, ("has_cols", "first_col", "last_col")),
        ):
            if axis is not None:
                values[flag] = 1
                values[first] = _uint(axis.first, first)
                values[last] = _uint(axis.last, last)
        with _ArgScope() as scope:
            ptr = scope.struct(S.PRINT_TITLES)
            S.PRINT_TITLES.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_set_print_titles(h, _uint(sheet, "sheet_index"), ptr))

    def add_row_break(self, row: int, *, manual: bool = True) -> None:
        """Upsert a row break before ``row``, spanning the whole sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_add_row_break(h, _uint(sheet, "sheet_index"), _uint(row, "row"), 1 if manual else 0))

    def add_column_break(self, col: int, *, manual: bool = True) -> None:
        """Upsert a column break before ``col``, spanning the whole sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_add_column_break(h, _uint(sheet, "sheet_index"), _uint(col, "col"), 1 if manual else 0))

    def remove_row_break(self, row: int) -> None:
        """Remove the row break at ``row``. An absent break is a no-op."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_row_break(h, _uint(sheet, "sheet_index"), _uint(row, "row")))

    def remove_column_break(self, col: int) -> None:
        """Remove the column break at ``col``. An absent break is a no-op."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_column_break(h, _uint(sheet, "sheet_index"), _uint(col, "col")))

    def clear_breaks(self) -> None:
        """Remove every manual break on both axes."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_breaks(h, _uint(sheet, "sheet_index")))

    def _read_breaks(self, count_export: str, at_export: str) -> List[PageBreak]:
        sheet = self._index()
        h = self._workbook._require()
        count = _read_count(getattr(LIB, count_export), h, _uint(sheet, "sheet_index"))
        out: List[PageBreak] = []
        with _ArgScope() as scope:
            ptr = scope.struct(S.PAGE_BREAK)
            for i in range(count):
                _check(getattr(LIB, at_export)(h, _uint(sheet, "sheet_index"), _uint(i, "index"), ptr))
                d = S.PAGE_BREAK.unpack(LIB, ptr)
                out.append(PageBreak(id=d["id"], min=d["min"], max=d["max"], manual=bool(d["manual"])))
        return out

    def list_row_breaks(self) -> List[PageBreak]:
        """Return the row breaks, ascending by row."""
        return self._read_breaks("fm_sheet_get_row_break_count", "fm_sheet_get_row_break_at")

    def list_column_breaks(self) -> List[PageBreak]:
        """Return the column breaks, ascending by column."""
        return self._read_breaks("fm_sheet_get_column_break_count", "fm_sheet_get_column_break_at")

    def set_page_setup(
        self,
        *,
        orientation: Union[PageOrientation, None, Unchanged] = UNCHANGED,
        paper_size: Union[int, Unchanged] = UNCHANGED,
        scale: Union[int, Unchanged] = UNCHANGED,
        fit_to_width: Union[int, Unchanged] = UNCHANGED,
        fit_to_height: Union[int, Unchanged] = UNCHANGED,
        fit_to_page: Union[bool, Unchanged] = UNCHANGED,
    ) -> None:
        """Update ``<pageSetup>``; an omitted argument keeps its attribute.

        ``paper_size`` is the OOXML code (9 = A4). ``scale`` outside
        ``[10, 400]`` is rejected rather than clamped -- a mis-stated print
        scale lands on paper, so rounding it silently would hide the error.
        ``fit_to_page`` sets or clears ``<sheetPr><pageSetUpPr fitToPage>``
        without disturbing anything else ``<sheetPr>`` carries; ``False``
        removes the attribute.
        ``orientation=None`` leaves a sheet that states no orientation as it
        is; the engine cannot remove a stated orientation.

        Raises:
          InvalidArgumentError: when ``orientation`` is ``None`` and the sheet
            states one.
          TypeError: when a non-nullable argument is ``None``.
        """
        sheet = self._index()
        if orientation is None:
            if self.get_page_setup().orientation is not None:
                raise FormulonError(
                    StatusCode.INVALID_ARGUMENT,
                    _diagnostic_override=("a stated page orientation cannot be removed", ""),
                )
            orientation = UNCHANGED
        values: Dict[str, object] = {}
        for name, value in (
            ("orientation", orientation),
            ("paper_size", paper_size),
            ("scale", scale),
            ("fit_to_width", fit_to_width),
            ("fit_to_height", fit_to_height),
            ("fit_to_page", fit_to_page),
        ):
            if value is UNCHANGED:
                continue
            if value is None:
                raise TypeError(f"formulon: {name} cannot be None")
            values[f"{name}_engaged"] = 1
            if name == "orientation":
                values[name] = _sint(value, name)
            elif name == "fit_to_page":
                values[name] = 1 if value else 0
            else:
                values[name] = _uint(value, name)
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PAGE_SETUP)
            S.PAGE_SETUP.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_set_page_setup(h, _uint(sheet, "sheet_index"), ptr))

    def set_page_margins(
        self,
        *,
        left: Union[float, Unchanged] = UNCHANGED,
        right: Union[float, Unchanged] = UNCHANGED,
        top: Union[float, Unchanged] = UNCHANGED,
        bottom: Union[float, Unchanged] = UNCHANGED,
        header: Union[float, Unchanged] = UNCHANGED,
        footer: Union[float, Unchanged] = UNCHANGED,
    ) -> None:
        """Update ``<pageMargins>``, in inches; an omitted argument keeps its margin.

        A negative, infinite or NaN margin is rejected: the paginator
        subtracts these from the paper, so such a value does not describe a
        printable body.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PAGE_MARGINS)
            values: Dict[str, object] = {}
            for name, value in (
                ("left", left),
                ("right", right),
                ("top", top),
                ("bottom", bottom),
                ("header", header),
                ("footer", footer),
            ):
                if value is UNCHANGED:
                    continue
                if value is None:
                    raise TypeError(f"formulon: {name} cannot be None")
                values[f"{name}_engaged"] = 1
                values[name] = float(value)
            S.PAGE_MARGINS.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_set_page_margins(h, _uint(sheet, "sheet_index"), ptr))

    def set_print_options(
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
            _flag(value, name)
            values[f"{name}_engaged"] = 1
            values[name] = 1 if value else 0
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PRINT_OPTIONS)
            S.PRINT_OPTIONS.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_set_print_options(h, _uint(sheet, "sheet_index"), ptr))

    def set_header_footer(
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
        for name, value in (
            ("different_odd_even", different_odd_even),
            ("different_first", different_first),
            ("scale_with_doc", scale_with_doc),
            ("align_with_margins", align_with_margins),
        ):
            if value is UNCHANGED:
                continue
            _flag(value, name)
            values[f"{name}_engaged"] = 1
            values[name] = 1 if value else 0
        sections: List[Tuple[str, str]] = []
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
            if text is not None and not isinstance(text, str):
                raise TypeError(f"formulon: {name} must be a str or None, got {type(text).__name__}")
            sections.append((name, text or ""))
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.HEADER_FOOTER)
            S.HEADER_FOOTER.pack(LIB, ptr, values)
            # Written after `pack`, which zeroes the block: a NULL pointer keeps
            # the section, an empty string clears it.
            for name, text in sections:
                _, off = S.HEADER_FOOTER.offsets[name]
                LIB.write_bytes(ptr + off, struct.pack("<I", scope.cstr(text)))
            _check(LIB.fm_sheet_set_header_footer(h, _uint(sheet, "sheet_index"), ptr))

    def get_page_setup(self) -> PageSetup:
        """Read the effective page setup and which attributes state it."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PAGE_SETUP)
            _check(LIB.fm_sheet_get_page_setup(h, _uint(sheet, "sheet_index"), ptr))
            d = S.PAGE_SETUP.unpack(LIB, ptr)
            return PageSetup(
                orientation=PageOrientation(d["orientation"]) if d["orientation_engaged"] else None,
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

    def get_page_margins(self) -> PageMargins:
        """Read the effective page margins in inches, with presence flags."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PAGE_MARGINS)
            _check(LIB.fm_sheet_get_page_margins(h, _uint(sheet, "sheet_index"), ptr))
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

    def get_print_options(self) -> PrintOptions:
        """Read the ``<printOptions>`` flags and which of them the XML states."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PRINT_OPTIONS)
            _check(LIB.fm_sheet_get_print_options(h, _uint(sheet, "sheet_index"), ptr))
            d = S.PRINT_OPTIONS.unpack(LIB, ptr)
        names = ("grid_lines", "headings", "horizontal_centered", "vertical_centered")
        fields: Dict[str, bool] = {name: bool(d[name]) for name in names}
        fields.update({f"{name}_stated": bool(d[f"{name}_engaged"]) for name in names})
        return PrintOptions(**fields)

    def get_header_footer(self) -> HeaderFooter:
        """Read the ``<headerFooter>`` sections and flags, and which flags the XML states.

        A section the sheet does not carry reads ``None``; a present one is
        its decoded text, as :meth:`set_header_footer` takes it.
        """
        sheet = self._index()
        h = self._workbook._require()
        sections = ("odd_header", "odd_footer", "even_header", "even_footer", "first_header", "first_footer")
        flags = ("different_odd_even", "different_first", "scale_with_doc", "align_with_margins")
        with _ArgScope() as scope:
            ptr = scope.struct(S.HEADER_FOOTER)
            _check(LIB.fm_sheet_get_header_footer(h, _uint(sheet, "sheet_index"), ptr))
            d = S.HEADER_FOOTER.unpack(LIB, ptr)
            fields: Dict[str, object] = {name: LIB.read_cstr(d[name]) if d[name] else None for name in sections}
        fields.update({name: bool(d[name]) for name in flags})
        fields.update({f"{name}_stated": bool(d[f"{name}_engaged"]) for name in flags})
        return HeaderFooter(**fields)

    def list_column_overrides(self) -> List[ColumnLayout]:
        """Return the column-layout overrides on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(LIB.fm_sheet_get_column_override_count, h, _uint(sheet, "sheet_index"))
        out: List[ColumnLayout] = []
        # One scratch block for the whole walk; see list_defined_names.
        with _ArgScope() as scope:
            ptr = scope.struct(S.COLUMN_LAYOUT)
            for i in range(n):
                S.zero_struct(LIB, S.COLUMN_LAYOUT, ptr)
                _check(LIB.fm_sheet_get_column_override_at(h, _uint(sheet, "sheet_index"), _uint(i, "idx"), ptr))
                d = S.COLUMN_LAYOUT.unpack(LIB, ptr)
                out.append(
                    ColumnLayout(
                        first=d["first"],
                        last=d["last"],
                        width=d["width"] if d["has_width"] else None,
                        hidden=bool(d["hidden"]),
                        outline_level=d["outline_level"],
                        collapsed=bool(d["collapsed"]),
                        style_xf=d["style_xf"] if d["has_style"] else None,
                    )
                )
        return out

    @overload
    @_span_form
    def set_column_width(self, first: int, last: int, width: float) -> None: ...

    @overload
    def set_column_width(self, span: str, width: float) -> None: ...

    def set_column_width(self, *args: Any, **kwargs: Any) -> None:
        """Set the column width override on ``[first, last]``.

        ``width`` is a stored width: characters with the column padding
        included, as ``<col width>`` holds it.
        """
        first, last, width = self._span_args("set_column_width", args, kwargs, ("first", "last", "width"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_column_width(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), float(width)
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
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), 1 if hidden else 0
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
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), _uint(level, "level", 8)
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
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), 1 if collapsed else 0
            )
        )

    @overload
    @_span_form
    def set_column_style(self, first: int, last: int, xf_index: int) -> None: ...

    @overload
    def set_column_style(self, span: str, xf_index: int) -> None: ...

    def set_column_style(self, *args: Any, **kwargs: Any) -> None:
        """Set the cell-style XF on ``[first, last]``, keeping the other column attributes."""
        first, last, xf_index = self._span_args("set_column_style", args, kwargs, ("first", "last", "xf_index"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_column_style(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), _uint(xf_index, "xf_index")
            )
        )

    @overload
    @_span_form
    def clear_column_style(self, first: int, last: int) -> None: ...

    @overload
    def clear_column_style(self, span: str) -> None: ...

    def clear_column_style(self, *args: Any, **kwargs: Any) -> None:
        """Remove the cell-style XF on ``[first, last]``, keeping the other column attributes."""
        first, last = self._span_args("clear_column_style", args, kwargs, ("first", "last"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_clear_column_style(h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"))
        )

    def list_row_overrides(self) -> List[RowLayout]:
        """Return the row-layout overrides on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(LIB.fm_sheet_get_row_override_count, h, _uint(sheet, "sheet_index"))
        out: List[RowLayout] = []
        # One scratch block for the whole walk; see list_defined_names.
        with _ArgScope() as scope:
            ptr = scope.struct(S.ROW_LAYOUT)
            for i in range(n):
                S.zero_struct(LIB, S.ROW_LAYOUT, ptr)
                _check(LIB.fm_sheet_get_row_override_at(h, _uint(sheet, "sheet_index"), _uint(i, "idx"), ptr))
                d = S.ROW_LAYOUT.unpack(LIB, ptr)
                out.append(
                    RowLayout(
                        row=d["row"],
                        height=d["height"] if d["has_height"] else None,
                        hidden=bool(d["hidden"]),
                        outline_level=d["outline_level"],
                        collapsed=bool(d["collapsed"]),
                        style_xf=d["style_xf"] if d["has_style"] else None,
                        custom_height=bool(d["custom_height"]),
                    )
                )
        return out

    @overload
    @_span_form
    def set_row_height(self, first: int, last: int, height: float) -> None: ...

    @overload
    def set_row_height(self, span: str, height: float) -> None: ...

    def set_row_height(self, *args: Any, **kwargs: Any) -> None:
        """Set the row height on ``[first, last]`` as a manual height (``customHeight``)."""
        first, last, height = self._span_args("set_row_height", args, kwargs, ("first", "last", "height"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_height(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), float(height)
            )
        )

    @overload
    @_span_form
    def set_row_auto_height(self, first: int, last: int, height: float) -> None: ...

    @overload
    def set_row_auto_height(self, span: str, height: float) -> None: ...

    def set_row_auto_height(self, *args: Any, **kwargs: Any) -> None:
        """Store ``height`` as the cached automatic height of the rows on ``[first, last]``.

        Saved as ``ht`` without ``customHeight``, so each row keeps following its
        content and font; the rows' other attributes are kept.
        """
        first, last, height = self._span_args("set_row_auto_height", args, kwargs, ("first", "last", "height"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_auto_height(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), float(height)
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
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), 1 if hidden else 0
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
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), _uint(level, "level", 8)
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
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), 1 if collapsed else 0
            )
        )

    @overload
    @_span_form
    def set_row_style(self, first: int, last: int, xf_index: int) -> None: ...

    @overload
    def set_row_style(self, span: str, xf_index: int) -> None: ...

    def set_row_style(self, *args: Any, **kwargs: Any) -> None:
        """Set the cell-style XF on the inclusive row span ``first..last``, keeping the other row attributes.

        Every row of the span gets an override. The span and ``xf_index`` are
        validated before anything changes.
        """
        first, last, xf_index = self._span_args("set_row_style", args, kwargs, ("first", "last", "xf_index"))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_style(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), _uint(xf_index, "xf_index")
            )
        )

    @overload
    @_span_form
    def clear_row_style(self, first: int, last: int) -> None: ...

    @overload
    def clear_row_style(self, span: str) -> None: ...

    def clear_row_style(self, *args: Any, **kwargs: Any) -> None:
        """Remove the cell-style XF on the inclusive row span ``first..last``, keeping the other row attributes."""
        first, last = self._span_args("clear_row_style", args, kwargs, ("first", "last"))
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_row_style(h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last")))

    def get_range_style_bases(self, range: Union[CellRect, str]) -> List[int]:
        """List, ascending, every cell XF some cell of the selection resolves to.

        ``range`` is a selection: full height is a column selection, full
        width a row selection, both the whole sheet, neither a cell range. It
        is the first step of an attribute patch (for example a font size over a selection): read the
        bases, build a patched XF for each with :meth:`Workbook.add_font` /
        :meth:`Workbook.add_cell_xf`, then apply them with :meth:`remap_range_style`.
        """
        rect = self._rect(range)
        return self.get_range_style_bases_excluding(rect, ())

    def get_range_style_bases_excluding(
        self,
        range: Union[CellRect, str],
        excluded: Sequence[Tuple[int, int]],
    ) -> List[int]:
        """:meth:`get_range_style_bases` over the selection less the ``(row, col)`` cells in ``excluded``.

        An excluded cell contributes no XF, neither its own nor the row, column
        or default style beneath it. Cells outside the selection and repeats are
        ignored; excluding every cell yields an empty list.
        """
        rect = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        cells = list(excluded)
        count = len(cells)
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            rows_buf = cols_buf = 0
            if count:
                rows_buf = scope.buffer(b"".join(struct.pack("<I", _uint(r, "excluded_row")) for r, _ in cells))
                cols_buf = scope.buffer(b"".join(struct.pack("<I", _uint(c, "excluded_col")) for _, c in cells))
            len_ptr = scope.out_ptr()
            # Size the buffer first, then fetch.
            _check(
                LIB.fm_sheet_get_range_style_bases_excluding(
                    h,
                    _uint(sheet, "sheet_index"),
                    rp,
                    rows_buf,
                    cols_buf,
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
                LIB.fm_sheet_get_range_style_bases_excluding(
                    h,
                    _uint(sheet, "sheet_index"),
                    rp,
                    rows_buf,
                    cols_buf,
                    _uint(count, "excluded_count"),
                    buf,
                    _uint(total, "cap"),
                    len_ptr,
                )
            )
            n = min(total, LIB.read_u32(len_ptr))
            return list(struct.unpack("<%dI" % n, LIB.read_bytes(buf, 4 * n)))

    def remap_range_style(self, range: Union[CellRect, str], mapping: Dict[int, int]) -> None:
        """Remap the XFs of a selection: a cell resolving to key ``k`` ends at ``mapping[k]``.

        An XF absent from ``mapping`` is kept. Each target must be a valid cell
        XF, else :class:`FormulonError` and nothing changes. Patch an attribute
        by reading :meth:`get_range_style_bases`, adding a patched XF for each base
        and passing ``{base: patched}``. Undo with :meth:`snapshot_range_style`
        taken beforehand and :meth:`restore_range_style`.
        """
        rect = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        pairs = list(mapping.items())
        count = len(pairs)
        from_buf = to_buf = 0
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            if count:
                from_buf = scope.buffer(b"".join(struct.pack("<I", _uint(k, "from_xf")) for k, _ in pairs))
                to_buf = scope.buffer(b"".join(struct.pack("<I", _uint(v, "to_xf")) for _, v in pairs))
            _check(
                LIB.fm_sheet_remap_range_style(
                    h,
                    _uint(sheet, "sheet_index"),
                    rp,
                    from_buf,
                    to_buf,
                    _uint(count, "count"),
                )
            )

    def snapshot_range_style(self, range: Union[CellRect, str]) -> bytes:
        """Capture the style state of a selection as opaque bytes for :meth:`restore_range_style`.

        The capture holds the covered column styles (column selection), row
        styles (row selection) and every stored cell's XF and explicit mark.
        Take it before :meth:`remap_range_style` to be able to undo the edit.
        """
        rect = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            bytes_ptr = scope.out_ptr()
            len_ptr = scope.out_ptr()
            data_ptr = 0
            try:
                _check(
                    LIB.fm_sheet_snapshot_range_style(
                        h,
                        _uint(sheet, "sheet_index"),
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

    def restore_range_style(self, data: Union[bytes, bytearray, memoryview]) -> None:
        """Put a :meth:`snapshot_range_style` capture back on the sheet.

        Captured axis styles and cell XFs are restored exactly; a stored cell of
        the captured rectangle that was not captured loses its style. Values and
        formulas are untouched. An invalid capture raises :class:`FormulonError`
        and changes nothing.
        """
        sheet = self._index()
        h = self._workbook._require()
        raw = bytes(data)
        with _ArgScope() as scope:
            buf = scope.buffer(raw) if raw else 0
            _check(LIB.fm_sheet_restore_range_style(h, _uint(sheet, "sheet_index"), buf, _uint(len(raw), "len")))

    # -- Conditional formatting --------------------------------------------
    @staticmethod
    def _decode_cf_color(d: Dict[str, Any], key: str) -> Tuple[int, int, int, int]:
        spec = d[key]
        if spec["kind"] != ColorKind.RGB:
            return (0, 0, 0, 0)
        argb = spec["rgb"]
        return ((argb >> 16) & 0xFF, (argb >> 8) & 0xFF, argb & 0xFF, argb >> 24)

    def evaluate_cf_range(
        self,
        range: Union[CellRect, str],
        today_serial: float = float("nan"),
    ) -> List[CfCell]:
        """Evaluate every CF block on the sheet against the cells of ``range``.

        Args:
          range: the cells to evaluate.
          today_serial: date basis pinned for ``timePeriod`` rules, as an
            Excel serial date. The default ``float("nan")`` disables them:
            no ``timePeriod`` rule can match, so pass a serial date to use
            them. ``0.0`` is *not* a disabling value -- it is the valid
            serial for 1899-12-30 and would evaluate those rules against
            that date.
        """
        return self._evaluate_cf_rect(self._rect(range), today_serial)

    def _evaluate_cf_rect(self, rect: CellRect, today_serial: float) -> List[CfCell]:
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rect_ptr = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rect_ptr, rect)
            out = scope.out_ptr()
            results = 0
            try:
                _check(
                    LIB.fm_sheet_evaluate_cf_range(
                        h,
                        _uint(sheet, "sheet_index"),
                        rect_ptr,
                        float(today_serial),
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
                    _check(LIB.fm_cf_results_get_cell_at(results, _uint(ci, "cell_index"), rp, cp, mcp))
                    row = LIB.read_u32(rp)
                    col = LIB.read_u32(cp)
                    mcount = LIB.read_u32(mcp)
                    matches: List[CfMatch] = []
                    for mi in range(mcount):
                        S.zero_struct(LIB, S.CF_MATCH, mptr)
                        _check(
                            LIB.fm_cf_results_get_match_at(
                                results, _uint(ci, "cell_index"), _uint(mi, "match_index"), mptr
                            )
                        )
                        d = S.CF_MATCH.unpack(LIB, mptr)
                        matches.append(
                            CfMatch(
                                kind=_enum(CfMatchKind, d["kind"]),
                                priority=d["priority"],
                                dxf_id_engaged=bool(d["dxf_id_engaged"]),
                                dxf_id=d["dxf_id"],
                                color=self._decode_cf_color(d, "color"),
                                bar_length_pct=d["bar_length_pct"],
                                bar_axis_position_pct=d["bar_axis_position_pct"],
                                bar_is_negative=bool(d["bar_is_negative"]),
                                bar_fill=self._decode_cf_color(d, "bar_fill"),
                                bar_border_engaged=bool(d["bar_border_engaged"]),
                                bar_border=self._decode_cf_color(d, "bar_border"),
                                bar_gradient=bool(d["bar_gradient"]),
                                icon_set_name=d["icon_set_name"],
                                icon_index=d["icon_index"],
                                bar_direction=_enum(CfDataBarDirection, d["bar_direction"]),
                            )
                        )
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
            _uint(sheet, "sheet_index"),
        )

    @staticmethod
    def _decode_cfvo(ptr: int) -> Cfvo:
        d = S.CFVO.unpack(LIB, ptr)
        return Cfvo(type=d["type"], value=LIB.read_cstr(d["value"]) if d["value"] else None, gte=bool(d["gte"]))

    @staticmethod
    def _decode_cf_color_at(ptr: int) -> CfColor:
        d = S.COLOR_SPEC.unpack(LIB, ptr)
        if d["kind"] != ColorKind.RGB:
            return CfColor(0, 0, 0, 0)
        argb = d["rgb"]
        return CfColor((argb >> 16) & 0xFF, (argb >> 8) & 0xFF, argb & 0xFF, argb >> 24)

    @staticmethod
    def _write_cfvo(ptr: int, value: Cfvo, scope: _ArgScope) -> None:
        S.CFVO.pack(LIB, ptr, {"type": _uint(value.type, "type", 8), "gte": 1 if value.gte else 0})
        if value.value is not None:
            S.write_str_field(LIB, ptr, S.CFVO, "value", value.value, scope)

    @staticmethod
    def _write_cf_color(ptr: int, color: CfColor) -> None:
        r, g, b, a = (_uint(color.r, "r", 8), _uint(color.g, "g", 8), _uint(color.b, "b", 8), _uint(color.a, "a", 8))
        if (r, g, b, a) == (0, 0, 0, 0):
            S.COLOR_SPEC.pack(LIB, ptr, {})
        else:
            S.COLOR_SPEC.pack(LIB, ptr, {"kind": ColorKind.RGB, "rgb": (a << 24) | (r << 16) | (g << 8) | b})

    def _cf_rule_at(self, cf_rule_index: int) -> ConditionalFormat:
        """Read the ``cf_rule_index``-th CF rule on the sheet (flattened order)."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CF_RULE)
            _check(LIB.fm_sheet_get_cf_rule(h, _uint(sheet, "sheet_index"), _uint(cf_rule_index, "idx"), ptr))
            d = S.CF_RULE.unpack(LIB, ptr)
            ranges = _decode_cell_rect_array(d["ranges"], d["range_count"])
            color_scale = None
            if d["color_scale_count"]:
                thresholds = [
                    self._decode_cfvo(d["color_scale_thresholds"] + i * S.CFVO.size)
                    for i in range(d["color_scale_count"])
                ]
                colors = [
                    self._decode_cf_color_at(d["color_scale_colors"] + i * S.COLOR_SPEC.size)
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
                    _enum(CfDataBarAxisPosition, d["data_bar_axis_position"])
                    if d["data_bar_axis_position_engaged"]
                    else None,
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
                    _enum(CfDataBarDirection, d["data_bar_direction"]),
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
                    self._decode_cfvo(ptr + offsets["icon_set_floor"][1]) if d["icon_set_floor_engaged"] else None,
                    [
                        CfIcon(
                            LIB.read_cstr(LIB.read_u32(d["icon_set_icons"] + i * S.CF_ICON.size)),
                            LIB.read_i32(d["icon_set_icons"] + i * S.CF_ICON.size + S.CF_ICON.offsets["icon_id"][1]),
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
                dxf_id=d["dxf_id"] if d["dxf_id_engaged"] else None,
                formula1=LIB.read_cstr(d["formula1"]),
                formula2=LIB.read_cstr(d["formula2"]),
                op=d["op"] if d["op_engaged"] else None,
                rank=d["rank"] if d["rank_engaged"] else None,
                percent=bool(d["percent"]),
                bottom=bool(d["bottom"]),
                above_average=bool(d["above_average"]),
                equal_average=bool(d["equal_average"]),
                std_dev=d["std_dev"] if d["std_dev_engaged"] else None,
                text=LIB.read_cstr(d["text"]),
                time_period=d["time_period"] if d["time_period_engaged"] else None,
                color_scale=color_scale,
                data_bar=data_bar,
                icon_set=icon_set,
            )

    def add_cf_rule(self, rule: ConditionalFormatInput) -> int:
        """Append a CF rule to the sheet; return its index.

        Every rule type is accepted, the visual ones included: a color scale,
        data bar or icon set rule carries its payload in ``rule.color_scale``,
        ``rule.data_bar`` or ``rule.icon_set``.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.CF_RULE)
            out = scope.out_ptr()
            S.CF_RULE.pack(
                LIB,
                ptr,
                {
                    "type": _uint(rule.type, "type", 8),
                    "op": _uint(rule.op or 0, "op", 8),
                    "time_period": _uint(rule.time_period or 0, "time_period", 8),
                    "priority": _sint(rule.priority, "priority"),
                    "stop_if_true": 1 if rule.stop_if_true else 0,
                    "dxf_id_engaged": 1 if rule.dxf_id is not None else 0,
                    "dxf_id": _uint(rule.dxf_id or 0, "dxf_id"),
                    "op_engaged": 1 if rule.op is not None else 0,
                    "rank_engaged": 1 if rule.rank is not None else 0,
                    "rank": _sint(rule.rank or 0, "rank"),
                    "percent": 1 if rule.percent else 0,
                    "bottom": 1 if rule.bottom else 0,
                    "above_average": 1 if rule.above_average else 0,
                    "equal_average": 1 if rule.equal_average else 0,
                    "std_dev_engaged": 1 if rule.std_dev is not None else 0,
                    "std_dev": float(rule.std_dev or 0.0),
                    "time_period_engaged": 1 if rule.time_period is not None else 0,
                },
            )
            ranges_ptr = _pack_cell_rect_array(rule.ranges, scope)
            ro = S.CF_RULE.offsets
            LIB.write_bytes(ptr + ro["ranges"][1], struct.pack("<I", ranges_ptr))
            LIB.write_bytes(ptr + ro["range_count"][1], struct.pack("<I", len(rule.ranges)))
            S.write_str_field(LIB, ptr, S.CF_RULE, "id", rule.id, scope)
            S.write_str_field(LIB, ptr, S.CF_RULE, "formula1", rule.formula1, scope)
            S.write_str_field(LIB, ptr, S.CF_RULE, "formula2", rule.formula2, scope)
            S.write_str_field(LIB, ptr, S.CF_RULE, "text", rule.text, scope)
            if rule.color_scale is not None:
                count = len(rule.color_scale.thresholds)
                if count != len(rule.color_scale.colors):
                    raise ValueError("color_scale thresholds and colors must have equal length")
                thresholds = scope.alloc(count * S.CFVO.size)
                colors = scope.alloc(count * S.COLOR_SPEC.size)
                for i, value in enumerate(rule.color_scale.thresholds):
                    self._write_cfvo(thresholds + i * S.CFVO.size, value, scope)
                for i, color in enumerate(rule.color_scale.colors):
                    self._write_cf_color(colors + i * S.COLOR_SPEC.size, color)
                LIB.write_bytes(ptr + ro["color_scale_thresholds"][1], struct.pack("<I", thresholds))
                LIB.write_bytes(ptr + ro["color_scale_colors"][1], struct.pack("<I", colors))
                LIB.write_bytes(ptr + ro["color_scale_count"][1], struct.pack("<I", count))
            if rule.data_bar is not None:
                data_bar = rule.data_bar
                self._write_cfvo(ptr + ro["data_bar_min"][1], data_bar.minimum, scope)
                self._write_cfvo(ptr + ro["data_bar_max"][1], data_bar.maximum, scope)
                self._write_cf_color(ptr + ro["data_bar_fill"][1], data_bar.fill)
                LIB.write_bytes(ptr + ro["data_bar_engaged"][1], struct.pack("<i", 1))
                LIB.write_bytes(ptr + ro["data_bar_show_value"][1], struct.pack("<i", 1 if data_bar.show_value else 0))
                LIB.write_bytes(
                    ptr + ro["data_bar_min_length_pct"][1],
                    struct.pack("<B", _uint(data_bar.min_length_pct, "min_length_pct", 8)),
                )
                LIB.write_bytes(
                    ptr + ro["data_bar_max_length_pct"][1],
                    struct.pack("<B", _uint(data_bar.max_length_pct, "max_length_pct", 8)),
                )
                # x14 extension. A `None` leaves the `*_engaged` flag clear,
                # which the C ABI reads as "keep the model default".
                if data_bar.gradient is not None:
                    LIB.write_bytes(ptr + ro["data_bar_gradient_engaged"][1], struct.pack("<i", 1))
                    LIB.write_bytes(ptr + ro["data_bar_gradient"][1], struct.pack("<i", 1 if data_bar.gradient else 0))
                if data_bar.axis_position is not None:
                    LIB.write_bytes(ptr + ro["data_bar_axis_position_engaged"][1], struct.pack("<i", 1))
                    LIB.write_bytes(
                        ptr + ro["data_bar_axis_position"][1],
                        struct.pack("<B", _uint(data_bar.axis_position, "axis_position", 8)),
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
                LIB.write_bytes(
                    ptr + ro["data_bar_direction"][1], struct.pack("<B", _uint(data_bar.direction, "direction", 8))
                )
            if rule.icon_set is not None:
                icon_set = rule.icon_set
                count = len(icon_set.thresholds)
                thresholds = scope.alloc(count * S.CFVO.size)
                for i, value in enumerate(icon_set.thresholds):
                    self._write_cfvo(thresholds + i * S.CFVO.size, value, scope)
                LIB.write_bytes(ptr + ro["icon_set_thresholds"][1], struct.pack("<I", thresholds))
                LIB.write_bytes(ptr + ro["icon_set_engaged"][1], struct.pack("<i", 1))
                LIB.write_bytes(ptr + ro["icon_set_name"][1], struct.pack("<B", _uint(icon_set.name, "icon_set", 8)))
                LIB.write_bytes(ptr + ro["icon_set_threshold_count"][1], struct.pack("<I", count))
                LIB.write_bytes(ptr + ro["icon_set_reverse"][1], struct.pack("<i", 1 if icon_set.reverse else 0))
                LIB.write_bytes(ptr + ro["icon_set_show_value"][1], struct.pack("<i", 1 if icon_set.show_value else 0))
                LIB.write_bytes(ptr + ro["icon_set_percent"][1], struct.pack("<i", 1 if icon_set.percent else 0))
                if icon_set.floor is not None:
                    LIB.write_bytes(ptr + ro["icon_set_floor_engaged"][1], struct.pack("<i", 1))
                    self._write_cfvo(ptr + ro["icon_set_floor"][1], icon_set.floor, scope)
                if icon_set.icons:
                    icons_ptr = scope.alloc(len(icon_set.icons) * S.CF_ICON.size)
                    for i, icon in enumerate(icon_set.icons):
                        name_ptr = scope.cstr(icon.icon_set)
                        S.CF_ICON.pack(
                            LIB,
                            icons_ptr + i * S.CF_ICON.size,
                            {"icon_set": name_ptr, "icon_id": _sint(icon.icon_id, "icon_id")},
                        )
                    LIB.write_bytes(ptr + ro["icon_set_icons"][1], struct.pack("<I", icons_ptr))
                    LIB.write_bytes(ptr + ro["icon_set_icon_count"][1], struct.pack("<I", len(icon_set.icons)))
            _check(LIB.fm_sheet_add_cf_rule(h, _uint(sheet, "sheet_index"), ptr, out))
            return LIB.read_u32(out)

    def remove_cf_rule(self, cf_rule: Union[int, str]) -> None:
        """Remove ``cf_rule``, an index in :meth:`list_cf_rules` order or a rule's ``id``.

        Raises:
          NotFoundError: when no rule has the id ``cf_rule``.
          IndexOutOfRangeError: when ``cf_rule`` is an index past the last rule.
        """
        index = resolve_ref(cf_rule, lambda: [rule.id for rule in self.list_cf_rules()], "CF rule")
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_cf_rule(h, _uint(sheet, "sheet_index"), _uint(index, "idx")))

    def clear_cf_rules(self) -> None:
        """Drop every CF block on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_cf_rules(h, _uint(sheet, "sheet_index")))

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
        xf_index = self.get_effective_style(row, col).xf_index
        return _cell_format_of(workbook, workbook.get_cell_xf(xf_index))

    def set_cell_format(
        self,
        range: Union[CellRect, str],
        *,
        number_format: Union[str, Unchanged] = UNCHANGED,
        font: Union[FontPatch, Font, Unchanged] = UNCHANGED,
        fill: Union[FillPatch, Fill, Unchanged] = UNCHANGED,
        border: Union[BorderPatch, Border, Unchanged] = UNCHANGED,
        alignment: Union[AlignmentPatch, Alignment, Unchanged] = UNCHANGED,
        protection: Union[CellProtectionPatch, CellProtection, Unchanged] = UNCHANGED,
    ) -> None:
        """Merge the given groups into the format of every cell of ``range``; an omitted group or key is kept.

        A group is a patch mapping (:class:`FontPatch`, ...) or the record
        :meth:`get_cell_format` reads, so passing every field of a
        :class:`CellFormat` copies that format. ``None``
        clears a key the read record reports as nullable (``bold``, ``size``,
        ``family``, ...). A colour is ``"#RRGGBB"``, ``"#AARRGGBB"`` or a
        :class:`Color`; a fill colour without a pattern makes the fill solid.
        A whole-column, whole-row or whole-sheet ``range`` (``"B:B"``,
        ``"3:3"``, ``"A:XFD"``) formats the column, row or sheet, so cells
        entered there later take the format too. Equal formats share one
        style record.

        Every argument is checked before anything is written and the sheet
        changes in one step at the end, but an engine failure on the way can
        leave unused style records behind (:meth:`Workbook.list_cell_xfs`).

        Raises:
          TypeError: for an unknown key, a value of the wrong type, or
            ``None`` on a key that is not nullable.
          ValueError: for a negative or fractional count or enum value, or a
            negative or non-finite size.
          InvalidArgumentError: for an empty ``number_format``, an enum value
            outside its table, ``indent`` above 255, or ``text_rotation``
            outside 0..180 and 255.

        Example::

            ws.set_cell_format("A1:B2", number_format="0.00", font={"bold": True})
        """
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
        bases = self.get_range_style_bases(rect)
        if not patch:
            return
        workbook = self._workbook
        num_fmt_id: Optional[int] = None
        if "number_format" in patch:
            code = patch["number_format"]
            num_fmt_id = 0 if code == "General" else workbook.add_num_fmt(code)
        xfs = [
            workbook.add_cell_xf(_patched_xf(workbook, workbook.get_cell_xf(base), patch, num_fmt_id)) for base in bases
        ]
        self.remap_range_style(rect, dict(zip(bases, xfs)))

    @overload
    @_coordinate_form
    def get_cell_xf_index(self, row: int, col: int) -> int: ...

    @overload
    def get_cell_xf_index(self, a1: str) -> int: ...

    def get_cell_xf_index(self, *args: Any, **kwargs: Any) -> int:
        """Return the xf (style record) index attached to a cell."""
        row, col = self._cell_args("get_cell_xf_index", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_sheet_get_cell_xf_index(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), out))
            return LIB.read_u32(out)

    @overload
    @_coordinate_form
    def set_cell_xf_index(self, row: int, col: int, xf_index: int) -> None: ...

    @overload
    def set_cell_xf_index(self, a1: str, xf_index: int) -> None: ...

    def set_cell_xf_index(self, *args: Any, **kwargs: Any) -> None:
        """Store ``xf_index`` on the cell at ``(row, col)``.

        Raises:
          FormulonError: with ``StatusCode.INVALID_ARGUMENT`` when ``xf_index`` is not
            below the workbook's cell-XF count.
        """
        row, col, xf_index = self._cell_args("set_cell_xf_index", args, kwargs, ("xf_index",))
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_cell_xf_index(
                h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), _uint(xf_index, "xf_index")
            )
        )

    @overload
    @_coordinate_form
    def clear_cell_xf_index(self, row: int, col: int) -> None: ...

    @overload
    def clear_cell_xf_index(self, a1: str) -> None: ...

    def clear_cell_xf_index(self, *args: Any, **kwargs: Any) -> None:
        """Remove the direct style assignment from the cell at ``(row, col)``.

        A blank cell inherits the row / column style again; value, formula and
        comments are kept. An absent cell is a no-op. A non-blank cell keeps
        owning the default xf. Use the ``source`` of :meth:`get_effective_style`
        to tell an unassigned cell from one directly assigned the Normal xf.
        """
        row, col = self._cell_args("clear_cell_xf_index", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_cell_xf_index(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col")))

    def set_range_xf_index(self, range: Union[CellRect, str], xf_index: int) -> None:
        """Store ``xf_index`` on every cell of ``range``.

        Cells that do not exist yet are materialised as blanks carrying only
        the style, so an empty ruled box renders. Ruling a report this way
        costs one call instead of one per cell.
        """
        rect = self._rect(range)
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            _check(LIB.fm_sheet_set_range_xf_index(h, _uint(sheet, "sheet"), rp, _uint(xf_index, "xf_index")))

    @overload
    @_coordinate_form
    def get_effective_style(self, row: int, col: int) -> EffectiveStyle: ...

    @overload
    def get_effective_style(self, a1: str) -> EffectiveStyle: ...

    def get_effective_style(self, *args: Any, **kwargs: Any) -> EffectiveStyle:
        """Read the effective style of the cell at ``(row, col)``; the cell need not exist."""
        row, col = self._cell_args("get_effective_style", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.EFFECTIVE_STYLE)
            _check(
                LIB.fm_sheet_get_effective_style(
                    h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), ptr
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
                font=ResolvedColor(d["font_argb"], ColorResolution(d["font_resolution"])),
                fill_foreground=ResolvedColor(d["fill_fg_argb"], ColorResolution(d["fill_fg_resolution"])),
                fill_background=ResolvedColor(d["fill_bg_argb"], ColorResolution(d["fill_bg_resolution"])),
                borders=[ResolvedColor(a, ColorResolution(r)) for a, r in zip(border_argb, border_res)],
                locked=bool(d["locked"]),
                hidden=bool(d["hidden"]),
                num_fmt_code=LIB.read_cstr(d["num_fmt_code"]),
            )

    # -- Pivot tables ------------------------------------------------------
    def _pivot_count(self) -> int:
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(LIB.fm_sheet_get_pivot_count, h, _uint(sheet, "sheet_index"))

    def _pivot_name(self, pivot_index: int) -> str:
        sheet = self._index()
        h = self._workbook._require()
        return _read_cstr_out(
            LIB.fm_sheet_get_pivot_name, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )

    def _pivot_index(self, pivot: Union[int, str, Pivot]) -> int:
        """The index ``pivot`` addresses: an index, a pivot table's name or a :class:`Pivot` of this sheet."""
        return resolve_ref(
            pivot, lambda: [self._pivot_name(i) for i in range(self._pivot_count())], "pivot", owner=self
        )

    def _pivot_at(self, pivot_index: int) -> Pivot:
        """The handle of the pivot at ``pivot_index``; the sheet keeps one per pivot."""
        from .pivot import Pivot

        sheet = self._index()
        h = self._workbook._require()
        runtime_id = _read_count(
            LIB.fm_sheet_get_pivot_runtime_id, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )
        pivot = self._pivots.get(runtime_id)
        if pivot is None:
            pivot = Pivot(self, runtime_id)
            self._pivots[runtime_id] = pivot
        return pivot

    def list_pivots(self) -> List[Pivot]:
        """Return the handles of the pivot tables anchored on the sheet, in index order."""
        return [self._pivot_at(i) for i in range(self._pivot_count())]

    def pivot(self, pivot: Union[int, str, Pivot]) -> Pivot:
        """Return the handle of the pivot table ``pivot``: an index, a name or a :class:`Pivot` of this sheet.

        Raises:
          NotFoundError: when no pivot has the name ``pivot``.
          IndexOutOfRangeError: when the index is past the last pivot.
          InvalidArgumentError: when ``pivot`` is a :class:`Pivot` of another sheet.
        """
        pivot_index = _uint(self._pivot_index(pivot), "pivot_index")
        if pivot_index >= self._pivot_count():
            raise FormulonError(
                StatusCode.INDEX_OUT_OF_RANGE, _diagnostic_override=(f"pivot index {pivot_index} is out of range", "")
            )
        return self._pivot_at(pivot_index)

    def add_pivot(self, *, name: str, cache_id: int, anchor_row: int, anchor_col: int) -> Pivot:
        """Create an empty pivot table over cache ``cache_id`` with its top-left cell at
        ``(anchor_row, anchor_col)``; return its handle.

        Example::

            cache = wb.add_pivot_cache()
            pivot = ws.add_pivot(name="Sales", cache_id=cache.cache_id, anchor_row=0, anchor_col=4)
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_add_pivot(
                    h,
                    _uint(sheet, "sheet_index"),
                    name_ptr,
                    _uint(cache_id, "cache_id"),
                    _uint(anchor_row, "anchor_row"),
                    _uint(anchor_col, "anchor_col"),
                    out,
                )
            )
            pivot_index = LIB.read_u32(out)
        return self._pivot_at(pivot_index)

    def remove_pivot(self, pivot: Union[int, str, Pivot]) -> None:
        """Remove the pivot table ``pivot``; its handle raises :class:`NotFoundError` from then on."""
        pivot_index = self._pivot_index(pivot)
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_pivot(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")))

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
            _uint(sheet, "sheet"),
            _uint(row, "row"),
            _uint(col, "col"),
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
            _uint(sheet, "sheet"),
            _uint(row, "row"),
            _uint(col, "col"),
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
                    _check(LIB.fm_cell_nodes_get_node_at(handle, _uint(i, "idx"), nptr))
                    d = S.CELL_NODE.unpack(LIB, nptr)
                    nodes.append(CellNode(sheet=d["sheet"], row=d["row"], col=d["col"]))
                return nodes
            finally:
                if handle:
                    LIB.fm_cell_nodes_destroy(handle)

    # -- Dynamic-array spill -----------------------------------------------
    @overload
    @_coordinate_form
    def get_spill_info(self, row: int, col: int) -> Spill: ...

    @overload
    def get_spill_info(self, a1: str) -> Spill: ...

    def get_spill_info(self, *args: Any, **kwargs: Any) -> Spill:
        """Return dynamic-array spill info for ``(row, col)``."""
        row, col = self._cell_args("get_spill_info", args, kwargs, ())
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.SPILL_INFO)
            _check(LIB.fm_sheet_get_spill_info(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), ptr))
            d = S.SPILL_INFO.unpack(LIB, ptr)
            engaged = bool(d["engaged"])
            blocked = bool(d["blocked"])
            return Spill(
                engaged=engaged,
                blocked=blocked,
                rect=_cell_rect_from(d["rect"]) if engaged or blocked else None,
            )

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
            ptr = scope.struct(S.ARRAY_FORMULA)
            _check(LIB.fm_sheet_get_array_formula(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), ptr))
            d = S.ARRAY_FORMULA.unpack(LIB, ptr)
            if not d["engaged"]:
                return None
            return ArrayFormula(
                rect=_cell_rect_from(d["rect"]),
                formula=LIB.read_cstr(d["formula"]),
            )

    def set_array_formula(self, array_formula: ArrayFormula) -> None:
        """Enter ``array_formula.formula`` as one legacy array formula over ``array_formula.rect``.

        The top-left cell of the rectangle anchors the block. Every cell of the rectangle loses its value and formula, and
        the next :meth:`Workbook.recalc` distributes the result. Refused when the
        rectangle partially overlaps another array, or intersects a merged
        area, a data table, a PivotTable or (for more than one cell) a table.
        Writing to the top-left cell of a block replaces the whole block.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, array_formula.rect)
            formula_ptr = scope.cstr(array_formula.formula)
            status = LIB.fm_sheet_set_array_formula(h, _uint(sheet, "sheet"), rp, formula_ptr)
            _check(status)


for _name, _signature in _COORDINATE_FORMS.items():
    getattr(Sheet, _name).__signature__ = _signature
del _name, _signature

_install_workbook_locks(Sheet)
