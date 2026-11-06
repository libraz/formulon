"""The :class:`Sheet` handle: every operation that targets one sheet of a workbook.

A :class:`~formulon.Workbook` hands out one :class:`Sheet` per sheet and keeps it
pointing at the same sheet across insertions, moves, renames and removals;
see :class:`Sheet`.
"""

from __future__ import annotations

import numbers
import struct
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union, cast, overload

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
from ._status import (
    StatusCode,
)
from .workbook import (
    _CELL_RANGE_NO_CURSOR,
    _CHART_TITLE_NONE,
    _DISPLAY_PATCH_SPEC,
    _PIVOT_NO_DATA_FIELD,
    _PIVOT_VALUES_FIELD_C,
    _PLOT_PATCH_SPEC,
    _PREVIEW_OVERRIDE_SPEC,
    _STATUS_NOT_FOUND,
    PIVOT_VALUES_FIELD,
    UNCHANGED,
    AnchorEditAs,
    AnchorKind,
    ArrayFormula,
    AutoFilter,
    AutoFilterEval,
    Cell,
    CellNode,
    CellRect,
    CellsInRange,
    CfCell,
    CfColor,
    CfDataBarAxisPosition,
    CfDataBarDirection,
    CfIcon,
    CfMatch,
    CfMatchKind,
    Cfvo,
    Chart,
    ChartAxisPatch,
    ChartData,
    ChartDisplayPatch,
    ChartGrouping,
    ChartInsertKind,
    ChartLegendPosition,
    ChartPlotPatch,
    ChartPointPatch,
    ChartPreviewOverridesInput,
    ChartSeriesDirection,
    ChartSeriesPatch,
    ChartTitle,
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
    DrawingObject,
    EffectiveStyle,
    EffectiveStyleSource,
    ErrorCode,
    ExcelError,
    GeometryMode,
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
    PhoneticPropertiesInput,
    PhoneticRun,
    PivotAggregation,
    PivotAxis,
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
    PivotItem,
    PivotLayout,
    PivotOptions,
    PivotReportLayout,
    PivotShowValuesAs,
    PivotSortOrder,
    PivotStyle,
    PivotStyleInput,
    PrintTitles,
    ProtectedRange,
    RectPt,
    ResolvedColor,
    RowLayout,
    Shape,
    ShapePaint,
    SheetComment,
    SheetFormatDefaults,
    SheetProtection,
    SheetSelection,
    SheetView,
    SheetViewMode,
    SheetVisibility,
    Slicer,
    SlicerCrossFilter,
    SlicerSortOrder,
    Sparkline,
    SparklineGroup,
    Spill,
    ThreadedComment,
    ThreadedCommentInput,
    ValidationErrorStyle,
    ValidationOutcome,
    Value,
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
    _drawing_object_from,
    _encode_axis_patch,
    _encode_patch,
    _encode_point_patch,
    _encode_series_patch,
    _enum,
    _enum_or_none,
    _header_mode,
    _hyperlink_from,
    _id_array,
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
    _read_chart,
    _read_chart_data,
    _read_count,
    _read_cstr_out,
    _read_optional_cstr_out,
    _shape_from,
    _slicer_from,
    _threaded_comment_from,
    _title_fields,
    _write_mentions,
    _write_value_input,
)

__all__ = ["Sheet"]

# <facade-tables>
WORKBOOK_METHODS = (
    "add_batch",
    "add_border",
    "add_cell_style_xf",
    "add_cell_xf",
    "add_dxf",
    "add_fill",
    "add_font",
    "add_num_fmt",
    "add_person",
    "add_pivot_cache",
    "add_pivot_cache_calculated_field",
    "add_pivot_cache_calculated_item",
    "add_pivot_cache_field",
    "add_pivot_cache_field_shared_item_blank",
    "add_pivot_cache_field_shared_item_bool",
    "add_pivot_cache_field_shared_item_date",
    "add_pivot_cache_field_shared_item_error",
    "add_pivot_cache_field_shared_item_number",
    "add_pivot_cache_field_shared_item_text",
    "add_pivot_cache_record",
    "add_sheet",
    "add_table",
    "apply_table_auto_filter",
    "clear_pivot_cache_field_shared_items",
    "clear_pivot_cache_fields",
    "clear_pivot_cache_records",
    "clear_table_auto_filter",
    "clear_theme",
    "close",
    "closed",
    "copy_sheet",
    "evaluate_table_auto_filter",
    "format_reference",
    "format_value",
    "get_active_sheet",
    "get_border",
    "get_border_count",
    "get_calc_mode",
    "get_calc_on_save",
    "get_cell_style",
    "get_cell_style_xf",
    "get_cell_style_xf_count",
    "get_cell_xf",
    "get_cell_xf_count",
    "get_date1904",
    "get_defined_name",
    "get_diagnostics",
    "get_dxf",
    "get_dxf_count",
    "get_excel_profile_id",
    "get_file_name",
    "get_fill",
    "get_fill_count",
    "get_font",
    "get_font_count",
    "get_full_precision",
    "get_image_info",
    "get_iterative",
    "get_memory_usage",
    "get_num_fmt_string",
    "get_passthrough",
    "get_person",
    "get_pinned_now",
    "get_pivot_cache_field_count",
    "get_pivot_cache_field_formula",
    "get_pivot_cache_field_name",
    "get_pivot_cache_field_shared_item_count",
    "get_pivot_cache_field_shared_item_value",
    "get_pivot_cache_options",
    "get_pivot_cache_record_count",
    "get_pivot_cache_worksheet_source",
    "get_ref_mode",
    "get_table_auto_filter",
    "get_table_style",
    "get_table_style_defaults",
    "get_theme_colors",
    "get_theme_fonts",
    "get_theme_script_font",
    "get_workbook_protection",
    "is_pivot_cache_field_shared_item_date",
    "list_cell_styles",
    "list_defined_names",
    "list_external_links",
    "list_passthroughs",
    "list_people",
    "list_pivot_cache_calculated_items",
    "list_pivot_caches",
    "list_sheets",
    "list_table_styles",
    "list_tables",
    "move_pivot_cache_calculated_item",
    "move_sheet",
    "parse_reference",
    "recalc",
    "recalc_data_tables",
    "remove_cell_style",
    "remove_defined_name",
    "remove_person",
    "remove_pivot_cache",
    "remove_pivot_cache_calculated_field",
    "remove_pivot_cache_calculated_item",
    "remove_sheet",
    "remove_table",
    "remove_table_auto_filter",
    "remove_table_style",
    "rename_sheet",
    "rename_table",
    "rename_table_column",
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
    "set_pivot_cache_calculated_field_formula",
    "set_pivot_cache_calculated_item_formula",
    "set_pivot_cache_options",
    "set_pivot_cache_record_blank",
    "set_pivot_cache_record_bool",
    "set_pivot_cache_record_date",
    "set_pivot_cache_record_error",
    "set_pivot_cache_record_number",
    "set_pivot_cache_record_text",
    "set_pivot_cache_worksheet_source",
    "set_ref_mode",
    "set_table_auto_filter",
    "set_table_column_totals",
    "set_table_style",
    "set_table_style_defaults",
    "set_table_style_info",
    "set_theme_colors",
    "set_theme_fonts",
    "set_theme_script_font",
    "set_workbook_protection",
    "sheet",
    "update_table",
)
SHEET_METHODS = (
    "add_cf_rule",
    "add_chart",
    "add_chart_series",
    "add_column_break",
    "add_data_table",
    "add_hyperlink",
    "add_hyperlink_range",
    "add_image",
    "add_merge",
    "add_pivot",
    "add_pivot_data_field",
    "add_pivot_field",
    "add_pivot_field_item",
    "add_pivot_field_shared_item",
    "add_pivot_field_subtotal_fn",
    "add_pivot_filter",
    "add_protected_range",
    "add_row_break",
    "add_shape",
    "add_slicer",
    "add_sparkline_group",
    "add_threaded_comment",
    "add_validation",
    "apply_auto_filter",
    "clear_auto_filter",
    "clear_breaks",
    "clear_cell_xf_index",
    "clear_cf_rules",
    "clear_column_style",
    "clear_column_width",
    "clear_hyperlinks",
    "clear_merges",
    "clear_pivot_data_fields",
    "clear_pivot_field_date_group",
    "clear_pivot_field_items",
    "clear_pivot_field_subtotal_fns",
    "clear_pivot_fields",
    "clear_pivot_filters",
    "clear_row_height",
    "clear_row_style",
    "clear_validations",
    "convert_column_chars_to_pt",
    "convert_column_pt_to_chars",
    "delete_columns",
    "delete_rows",
    "evaluate_auto_filter",
    "evaluate_cf_formula",
    "evaluate_cf_range",
    "evaluate_formula_array",
    "get_array_formula",
    "get_auto_filter",
    "get_auto_filter_xml",
    "get_cell_at",
    "get_cell_count",
    "get_cell_phonetic",
    "get_cell_phonetic_properties",
    "get_cell_phonetic_runs",
    "get_cell_rect_pt",
    "get_cell_show_phonetic",
    "get_cell_xf_index",
    "get_cells_in_range",
    "get_chart",
    "get_chart_data",
    "get_chart_preview",
    "get_column_width_pt",
    "get_comment",
    "get_comment_visible",
    "get_data_table",
    "get_dependents",
    "get_display_text",
    "get_drawing_object",
    "get_effective_style",
    "get_format_defaults",
    "get_formula",
    "get_formula_r1c1",
    "get_header_footer_xml",
    "get_hyperlink",
    "get_image",
    "get_invalid_cells",
    "get_lambda_text",
    "get_merges_in_range",
    "get_page_margins",
    "get_page_margins_xml",
    "get_page_setup",
    "get_page_setup_xml",
    "get_pagination",
    "get_pivot_cache_id",
    "get_pivot_column_field_order",
    "get_pivot_count",
    "get_pivot_data_field_default_name",
    "get_pivot_field_item",
    "get_pivot_field_item_count",
    "get_pivot_grand_totals",
    "get_pivot_layout",
    "get_pivot_name",
    "get_pivot_options",
    "get_pivot_report_layout",
    "get_pivot_row_field_order",
    "get_pivot_style_info",
    "get_precedents",
    "get_print_area",
    "get_print_options_xml",
    "get_print_titles",
    "get_range_style_bases",
    "get_range_style_bases_excluding",
    "get_row_height_pt",
    "get_selection",
    "get_shape",
    "get_sheet_pr_xml",
    "get_sheet_protection",
    "get_slicer",
    "get_spill_info",
    "get_threaded_comment",
    "get_value",
    "get_view",
    "get_width_model",
    "has_cell_phonetic_properties",
    "insert_columns",
    "insert_rows",
    "is_pivot_projection_pending",
    "list_cf_rules",
    "list_column_breaks",
    "list_column_overrides",
    "list_comments",
    "list_drawing_objects",
    "list_hyperlinks",
    "list_merges",
    "list_pivot_data_fields",
    "list_pivot_fields",
    "list_pivot_filters",
    "list_protected_ranges",
    "list_row_breaks",
    "list_row_overrides",
    "list_shapes",
    "list_slicers",
    "list_sparkline_groups",
    "list_threaded_comments",
    "list_validations",
    "project_pivot",
    "recalc_viewport",
    "remap_range_style",
    "remove_auto_filter",
    "remove_cf_rule",
    "remove_chart_series",
    "remove_column_break",
    "remove_comment",
    "remove_data_table",
    "remove_drawing_object",
    "remove_hyperlink",
    "remove_hyperlink_at",
    "remove_merge",
    "remove_merge_at",
    "remove_pivot",
    "remove_pivot_filter",
    "remove_protected_range",
    "remove_row_break",
    "remove_shape",
    "remove_slicer",
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
    "set_cell_phonetic",
    "set_cell_phonetic_properties",
    "set_cell_phonetic_runs",
    "set_cell_show_phonetic",
    "set_cell_xf_index",
    "set_chart_axis",
    "set_chart_display",
    "set_chart_filter",
    "set_chart_legend",
    "set_chart_plot",
    "set_chart_point",
    "set_chart_series",
    "set_chart_source",
    "set_chart_title",
    "set_chart_type",
    "set_column_collapsed",
    "set_column_hidden",
    "set_column_outline",
    "set_column_style",
    "set_column_width",
    "set_comment",
    "set_comment_visible",
    "set_drawing_anchor",
    "set_drawing_name",
    "set_drawing_z_order",
    "set_error",
    "set_fit_to_page",
    "set_format_defaults",
    "set_formula",
    "set_freeze",
    "set_header_footer",
    "set_header_footer_xml",
    "set_number",
    "set_page_margins",
    "set_page_margins_xml",
    "set_page_setup",
    "set_page_setup_xml",
    "set_pane_top_left",
    "set_pivot_anchor",
    "set_pivot_column_field_order",
    "set_pivot_data_field",
    "set_pivot_field_axis",
    "set_pivot_field_custom_name",
    "set_pivot_field_date_group",
    "set_pivot_field_default_subtotal",
    "set_pivot_field_fill_down_labels",
    "set_pivot_field_insert_blank_row",
    "set_pivot_field_item_caption",
    "set_pivot_field_item_show_detail",
    "set_pivot_field_item_visible",
    "set_pivot_field_number_format",
    "set_pivot_field_show_all",
    "set_pivot_field_show_detail",
    "set_pivot_field_sort",
    "set_pivot_field_subtotal_top",
    "set_pivot_grand_totals",
    "set_pivot_name",
    "set_pivot_options",
    "set_pivot_report_layout",
    "set_pivot_row_field_order",
    "set_pivot_style_info",
    "set_print_area",
    "set_print_options",
    "set_print_options_xml",
    "set_print_titles",
    "set_range_xf_index",
    "set_right_to_left",
    "set_row_auto_height",
    "set_row_collapsed",
    "set_row_height",
    "set_row_hidden",
    "set_row_outline",
    "set_row_style",
    "set_selection",
    "set_sheet_pr_xml",
    "set_sheet_protection",
    "set_show_formulas",
    "set_show_grid_lines",
    "set_show_outline_symbols",
    "set_show_row_col_headers",
    "set_show_zeros",
    "set_sparkline_group_sparklines",
    "set_split",
    "set_tab_selected",
    "set_text",
    "set_thread_resolved",
    "set_top_left",
    "set_value",
    "set_view_mode",
    "set_visibility",
    "set_zoom",
    "show_pivot_details",
    "snapshot_drawing",
    "snapshot_range_style",
    "update_protected_range",
    "update_shape",
    "update_slicer",
    "update_sparkline_group",
    "update_threaded_comment",
    "validate_value",
)
SHEET_STRUCTURE_OPS = (
    "add_sheet",
    "copy_sheet",
    "move_sheet",
    "remove_sheet",
    "rename_sheet",
    "show_pivot_details",
)
LIST_FAMILIES = {
    "list_cf_rules": ("get_cf_rule_count", "get_cf_rule"),
    "list_external_links": ("get_external_link_count", "get_external_link"),
    "list_pivot_cache_calculated_items": ("get_pivot_cache_calculated_item_count", "get_pivot_cache_calculated_item"),
    "list_pivot_data_fields": ("get_pivot_data_field_count", "get_pivot_data_field"),
    "list_pivot_fields": ("get_pivot_field_count", "get_pivot_field"),
    "list_pivot_filters": ("get_pivot_filter_count", "get_pivot_filter"),
    "list_protected_ranges": ("get_protected_range_count", "get_protected_range"),
    "list_sparkline_groups": ("get_sparkline_group_count", "get_sparkline_group"),
    "list_tables": ("get_table_count", "get_table"),
    "list_validations": ("get_validation_count", "get_validation"),
}
HANDWRITTEN = (
    "add_hyperlink",
    "add_hyperlink_range",
    "add_sheet",
    "copy_sheet",
    "get_active_sheet",
    "get_cell_phonetic_runs",
    "get_formula",
    "get_theme_colors",
    "get_theme_fonts",
    "list_sheets",
    "move_sheet",
    "remove_sheet",
    "rename_sheet",
    "set_active_sheet",
    "set_array_formula",
    "set_chart_source",
    "set_comment",
    "set_formula",
    "set_pinned_now",
    "set_pivot_field_date_group",
    "set_pivot_grand_totals",
    "set_theme_colors",
    "set_theme_script_font",
    "set_value",
    "show_pivot_details",
    "update_threaded_comment",
)
CELL_ARGS = {
    "add_hyperlink": ((0, 1),),
    "add_pivot": ((2, 3),),
    "clear_cell_xf_index": ((0, 1),),
    "evaluate_cf_formula": ((0, 1), (2, 3)),
    "evaluate_formula_array": ((0, 1),),
    "get_array_formula": ((0, 1),),
    "get_cell_phonetic": ((0, 1),),
    "get_cell_phonetic_properties": ((0, 1),),
    "get_cell_phonetic_runs": ((0, 1),),
    "get_cell_show_phonetic": ((0, 1),),
    "get_cell_xf_index": ((0, 1),),
    "get_comment": ((0, 1),),
    "get_comment_visible": ((0, 1),),
    "get_data_table": ((0, 1),),
    "get_dependents": ((0, 1),),
    "get_display_text": ((0, 1),),
    "get_effective_style": ((0, 1),),
    "get_formula": ((0, 1),),
    "get_formula_r1c1": ((0, 1),),
    "get_hyperlink": ((0, 1),),
    "get_lambda_text": ((0, 1),),
    "get_precedents": ((0, 1),),
    "get_spill_info": ((0, 1),),
    "get_value": ((0, 1),),
    "remove_comment": ((0, 1),),
    "remove_data_table": ((0, 1),),
    "remove_hyperlink": ((0, 1),),
    "set_blank": ((0, 1),),
    "set_bool": ((0, 1),),
    "set_cell_phonetic": ((0, 1),),
    "set_cell_phonetic_properties": ((0, 1),),
    "set_cell_phonetic_runs": ((0, 1),),
    "set_cell_show_phonetic": ((0, 1),),
    "set_cell_xf_index": ((0, 1),),
    "set_comment": ((0, 1),),
    "set_comment_visible": ((0, 1),),
    "set_error": ((0, 1),),
    "set_formula": ((0, 1),),
    "set_number": ((0, 1),),
    "set_pane_top_left": ((0, 1),),
    "set_text": ((0, 1),),
    "set_top_left": ((0, 1),),
    "set_value": ((0, 1),),
    "show_pivot_details": ((1, 2),),
    "validate_value": ((0, 1),),
}
RECT_ARGS = {
    "add_hyperlink_range": (0,),
    "add_merge": (0,),
    "evaluate_cf_range": (0,),
    "get_cell_rect_pt": (0,),
    "get_cells_in_range": (0,),
    "get_merges_in_range": (0,),
    "get_range_style_bases": (0,),
    "get_range_style_bases_excluding": (0,),
    "recalc_viewport": (0,),
    "remap_range_style": (0,),
    "remove_merge": (0,),
    "set_pivot_anchor": (1,),
    "set_print_area": (0,),
    "set_range_xf_index": (0,),
    "snapshot_range_style": (0,),
}
# </facade-tables>


class Sheet:
    """One sheet of a :class:`Workbook`; every sheet-scope operation lives here.

    Obtain a handle from :meth:`Workbook.sheet`, ``wb[key]``, iteration or
    :meth:`Workbook.list_sheets`. A workbook hands out one handle per sheet,
    so ``wb[0] is wb["Data"]`` when the first sheet is named ``Data``. The
    handle follows its sheet through :meth:`Workbook.add_sheet`,
    :meth:`Workbook.copy_sheet`, :meth:`Workbook.move_sheet`,
    :meth:`Workbook.rename_sheet` and :meth:`show_pivot_details`. Once
    :meth:`Workbook.remove_sheet` removes the sheet, every operation raises
    :class:`NotFoundError`; once the workbook is closed, every operation
    except reading :attr:`workbook` raises :class:`WorkbookClosedError`.

    Rows and columns are 0-based. :meth:`get`, :meth:`set`,
    :meth:`get_formula` and :meth:`set_formula` also take one unqualified A1
    cell (``"B3"``, ``"$B$3"``, ``"b3"``); a range, a sheet-qualified
    reference, R1C1 text, surrounding whitespace or anything else
    :meth:`Workbook.parse_reference` does not read as one cell raises
    :class:`InvalidArgumentError`. ``sheet["A1"]`` returns
    ``sheet.get("A1").to_python()``, and ``sheet["A1"] = v`` is
    ``sheet.set("A1", v)``.
    """

    __slots__ = ("_workbook", "_position", "_name")

    def __init__(self, workbook: Workbook, position: int, name: str) -> None:
        self._workbook = workbook
        # The current 0-based position, or None once the sheet is removed.
        self._position: Optional[int] = position
        self._name = name

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

    # -- A1 access ---------------------------------------------------------
    def _address(self, row_or_a1: Union[int, str], col: Optional[int]) -> Tuple[int, int]:
        """``(row, col)`` from an A1 address or a 0-based pair."""
        if isinstance(row_or_a1, str):
            if col is not None:
                raise TypeError("formulon: an A1 address takes no column argument")
            return self._cell(row_or_a1)
        if col is None:
            raise TypeError("formulon: a row index needs a column index")
        return row_or_a1, col

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

    def get(self, a1: str) -> Value:
        """The value at ``a1``, as :meth:`get_value` returns it."""
        self._index()
        row, col = self._cell(a1)
        return self.get_value(row, col)

    def set(self, a1: str, value: Union[None, bool, int, float, str, ExcelError, Value]) -> None:  # noqa: FBT001 -- the value is any cell type; a bool is one of them
        """Store ``value`` at ``a1``, as :meth:`set_value` does."""
        self._index()
        row, col = self._cell(a1)
        self.set_value(row, col, value)

    def __getitem__(self, a1: str) -> Union[None, float, bool, str, ExcelError, Value]:
        return self.get(a1).to_python()

    def __setitem__(self, a1: str, value: Union[None, bool, int, float, str, ExcelError, Value]) -> None:
        self.set(a1, value)

    # -- Cell mutation -----------------------------------------------------
    def set_number(self, row: int, col: int, value: float) -> None:
        """Store the number ``value`` at ``(row, col)``.

        ``row`` and ``col`` are 0-based.
        """
        sheet = self._index()
        h = self._workbook._require()
        status = LIB.fm_sheet_set_number(
            h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), float(value)
        )
        _check(status)

    def set_bool(self, row: int, col: int, value: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Store the boolean ``value`` at ``(row, col)``.

        ``row`` and ``col`` are 0-based.
        """
        sheet = self._index()
        h = self._workbook._require()
        status = LIB.fm_sheet_set_bool(
            h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), 1 if value else 0
        )
        _check(status)

    def set_error(self, row: int, col: int, error_code: ErrorCode) -> None:
        """Store an Excel cell error at ``(row, col)``.

        ``row`` and ``col`` are 0-based. ``error_code`` is an
        :class:`ErrorCode` member, not a display literal; the same member
        comes back out through :attr:`Value.error_code`.
        """
        sheet = self._index()
        h = self._workbook._require()
        status = LIB.fm_sheet_set_error(
            h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), _sint(error_code, "error")
        )
        _check(status)

    def set_text(self, row: int, col: int, value: str) -> None:
        """Store the string ``value`` at ``(row, col)``.

        ``row`` and ``col`` are 0-based. The text is stored
        verbatim; a leading ``=`` does not make it a formula (use
        :meth:`set_formula`).
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            text_ptr = scope.cstr(value)
            status = LIB.fm_sheet_set_text(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), text_ptr
            )
            _check(status)

    def set_blank(self, row: int, col: int) -> None:
        """Clear the cell at ``(row, col)``.

        ``row`` and ``col`` are 0-based. Clearing drops the
        stored value or formula; cell formatting is unaffected.
        """
        sheet = self._index()
        h = self._workbook._require()
        status = LIB.fm_sheet_set_blank(h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"))
        _check(status)

    @overload
    def set_formula(self, row_or_a1: str, col: str) -> None: ...

    @overload
    def set_formula(self, row_or_a1: int, col: int, formula: str) -> None: ...

    def set_formula(self, row_or_a1: Union[int, str], col: Union[int, str], formula: Optional[str] = None) -> None:
        """Store a formula at one cell.

        The cell is either one A1 address (``set_formula("B3", "=1+2")``) or a
        0-based ``(row, col)`` pair (``set_formula(2, 1, "=1+2")``). The
        leading ``=`` is optional. The cell keeps its previous cached value
        until the next :meth:`Workbook.recalc`.
        """
        sheet = self._index()
        if isinstance(row_or_a1, str):
            if formula is not None:
                raise TypeError("formulon: set_formula takes (a1, formula) or (row, col, formula)")
            formula = cast(str, col)
            row, col = self._address(row_or_a1, None)
        else:
            if formula is None:
                raise TypeError("formulon: set_formula(row, col, formula) is missing formula")
            row, col = self._address(row_or_a1, cast(int, col))
        h = self._workbook._require()
        with _ArgScope() as scope:
            formula_ptr = scope.cstr(formula)
            status = LIB.fm_sheet_set_formula(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), formula_ptr
            )
            _check(status)

    def set_value(
        self,
        row: int,
        col: int,
        value: Union[None, bool, int, float, str, "ExcelError", Value],  # noqa: FBT001 -- the value is any cell type; a bool is one of them
    ) -> None:
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
        """
        if isinstance(value, Value):
            kind = value.kind
            if kind is ValueKind.BLANK:
                self.set_blank(row, col)
            elif kind is ValueKind.NUMBER:
                self.set_number(row, col, cast(float, value.number))
            elif kind is ValueKind.BOOL:
                self.set_bool(row, col, cast(bool, value.boolean))
            elif kind is ValueKind.TEXT:
                self.set_text(row, col, cast(str, value.text))
            elif kind is ValueKind.ERROR:
                self.set_error(row, col, cast(ErrorCode, value.error_code))
            else:
                raise TypeError(f"formulon: a Value of kind {kind.name} cannot be stored in a cell")
        elif value is None:
            self.set_blank(row, col)
        elif isinstance(value, bool):
            self.set_bool(row, col, value)
        elif isinstance(value, numbers.Real):
            try:
                number = float(value)
            except OverflowError as exc:
                raise ValueError("formulon: the integer is too large for a float") from exc
            self.set_number(row, col, number)
        elif isinstance(value, str):
            self.set_text(row, col, value)
        elif isinstance(value, ExcelError):
            self.set_error(row, col, value.code)
        else:
            raise TypeError(f"formulon: cannot store a {type(value).__name__} in a cell")

    def set_cell_phonetic(self, row: int, col: int, text: Optional[str]) -> None:
        """Set the cell's phonetic guide (OOXML ``<rPh>``, furigana).

        ``row`` and ``col`` are 0-based. ``None`` (or ``""``) removes the
        guide. The guide is display metadata attached to the
        cell's string; it does not affect evaluation.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            text_ptr = scope.cstr(text or "")
            status = LIB.fm_sheet_set_cell_phonetic(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), text_ptr
            )
            _check(status)

    def set_cell_phonetic_runs(self, row: int, col: int, runs: Optional[Sequence[PhoneticRun]]) -> None:
        """Set the cell's phonetic guide as one ``<rPh>`` block per run.

        ``row`` and ``col`` are 0-based. ``None`` or an empty ``runs`` removes
        the guide. Unlike :meth:`set_cell_phonetic`, which annotates the whole
        cell, this keeps the spans each reading covers -- the difference
        ``PHONETIC`` and a partially annotated cell depend on.

        The runs must be an ordered partition: each needs ``sb <= eb`` and
        must start at or after the previous run's ``eb``.
        """
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

    def get_cell_phonetic_runs(self, row: int, col: int) -> List[PhoneticRun]:
        """Return the cell's ``<rPh>`` blocks with their spans; ``[]`` for an unannotated cell.

        ``row`` and ``col`` are 0-based. :meth:`get_cell_phonetic` returns the
        readings concatenated, without the spans.
        """
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

    def set_cell_phonetic_properties(self, row: int, col: int, properties: PhoneticPropertiesInput) -> None:
        """Set how the cell's phonetic guide renders.

        ``row`` and ``col`` are 0-based. Independent of
        :meth:`set_cell_phonetic_runs` in both directions: editing the readings
        does not reset the rendering, and this does not touch the readings.
        Only observable on a cell that has runs, and every value-setting
        method clears both, so call it after the cell's text.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_cell_phonetic_properties(
                h,
                _uint(sheet, "sheet_index"),
                _uint(row, "row"),
                _uint(col, "col"),
                _uint(properties.font_id, "font_id"),
                _uint(properties.type, "type"),
                _uint(properties.alignment, "alignment"),
            )
        )

    def get_cell_phonetic_properties(self, row: int, col: int) -> PhoneticProperties:
        """Return how the cell's phonetic guide renders.

        ``row`` and ``col`` are 0-based. A cell with no
        annotation reports the all-zero default.
        """
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
            )

    def has_cell_phonetic_properties(self, row: int, col: int) -> bool:
        """Whether the cell's phonetic properties are explicit.

        True when they came from a ``<phoneticPr>`` element or
        :meth:`set_cell_phonetic_properties` (an all-zero triple included), false
        for the state an absent element resolves to. The JS bindings report
        it as ``present`` on the phonetic-properties result.
        """
        sheet = self._index()
        h = self._workbook._require()
        return bool(
            _read_count(
                LIB.fm_sheet_has_cell_phonetic_properties,
                h,
                _uint(sheet, "sheet_index"),
                _uint(row, "row"),
                _uint(col, "col"),
            )
        )

    def get_cell_show_phonetic(self, row: int, col: int) -> bool:
        """Whether the cell shows its phonetic guide (``<c ph="1">``)."""
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

    def set_cell_show_phonetic(self, row: int, col: int, show: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the cell's "Show Phonetic" display flag.

        Independent of the readings: value writes clear the readings but keep
        this flag. A formula cell's flag is neither saved nor read from a file.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_cell_show_phonetic(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), 1 if show else 0
            )
        )

    def get_cell_phonetic(self, row: int, col: int) -> Optional[str]:
        """Return the cell's phonetic guide, or ``None`` when it has none.

        ``row`` and ``col`` are 0-based.
        """
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
    def get_value(self, row: int, col: int) -> Value:
        """Read the cached value at ``(row, col)``.

        Reflects the most recent recalc; mutations made after the last
        :meth:`Workbook.recalc` are not visible to dependents until you call
        :meth:`Workbook.recalc` again.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            value_ptr = scope.alloc(fm_value_t_size)
            status = LIB.fm_sheet_get_value(
                h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col"), value_ptr
            )
            _check(status)
            return Value._from_wasm(value_ptr)

    # -- Ad-hoc array evaluation ------------------------------------------
    def evaluate_cf_formula(self, row: int, col: int, anchor_row: int, anchor_col: int, formula: str) -> bool:
        """Evaluate a conditional-format formula at one cell without mutation.

        The result is coerced with Excel's conditional-format rules: an error,
        blank, text or numeric-zero result is ``False`` (the rule does not
        fire), any non-zero number is ``True``. An evaluation error is not
        raised.
        """
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

    def evaluate_formula_array(self, row: int, col: int, formula: str) -> List[List[Value]]:
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
        sheet = self._index()
        return self._workbook._evaluate_formula_cells(sheet, row, col, formula, top_left_only=False)

    # -- Iteration ---------------------------------------------------------
    def get_cell_count(self) -> int:
        """Return the number of written coordinates on the sheet.

        Counts each cell set by a value, formula or format setter (clearing
        keeps the slot) and each spill phantom. A workbook mutation
        invalidates the enumeration :meth:`get_cell_at` walks.
        """
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(LIB.fm_sheet_get_cell_count, h, _uint(sheet, "sheet_index"))

    def get_cell_at(self, index: int) -> Cell:
        """Return the ``index``-th written coordinate on the sheet, sorted by ``(row, col)`` ascending."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            row_ptr = scope.out_ptr()
            col_ptr = scope.out_ptr()
            formula_ptr = scope.out_ptr()
            value_ptr = scope.alloc(fm_value_t_size)
            status = LIB.fm_sheet_get_cell_at(
                h, _uint(sheet, "sheet_index"), _uint(index, "idx"), row_ptr, col_ptr, formula_ptr, value_ptr
            )
            _check(status)
            formula_addr = LIB.read_u32(formula_ptr)
            return Cell(
                row=LIB.read_u32(row_ptr),
                col=LIB.read_u32(col_ptr),
                formula=LIB.read_cstr(formula_addr) if formula_addr else None,
                value=Value._from_wasm(value_ptr),
            )

    def insert_rows(self, row: int, count: int) -> None:
        """Insert ``count`` rows at ``row`` on the sheet.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them or a table past the sheet bound; the workbook is unchanged.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_insert_rows(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(count, "count")))

    def delete_rows(self, row: int, count: int) -> None:
        """Delete ``count`` rows starting at ``row`` on the sheet.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them or a table past the sheet bound; the workbook is unchanged.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_delete_rows(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(count, "count")))

    def insert_columns(self, col: int, count: int) -> None:
        """Insert ``count`` columns at ``col`` on the sheet.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them or a table past the sheet bound; the workbook is unchanged.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_insert_columns(h, _uint(sheet, "sheet"), _uint(col, "col"), _uint(count, "count")))

    def delete_columns(self, col: int, count: int) -> None:
        """Delete ``count`` columns starting at ``col`` on the sheet.

        The anchor of an embedded OLE object or form control is not
        remapped; see the note above this method group.

        Raises:
          FormulonError: with ``StatusCode.PRECONDITION_FAILED`` when the edit cuts into
            a data table, legacy array block or PivotTable, or would push one
            of them or a table past the sheet bound; the workbook is unchanged.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_delete_columns(h, _uint(sheet, "sheet"), _uint(col, "col"), _uint(count, "count")))

    # -- Partial recalc ----------------------------------------------------
    #
    # Note: ``fm_workbook_set_iterative_progress`` (the per-sweep callback)
    # is intentionally NOT bound -- it takes a C function pointer that the
    # wasmtime-py host cannot synthesise into the module's function table.
    def recalc_viewport(self, rect: CellRect) -> int:
        """Recalc the dependency closure for the viewport ``rect``; return cell count."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            vp = scope.struct(S.CELL_RECT)
            out = scope.out_ptr()
            _pack_cell_rect_at(vp, rect)
            _check(LIB.fm_sheet_recalc_viewport(h, _uint(sheet, "sheet"), vp, out))
            return LIB.read_u32(out)

    # -- Lambda text -------------------------------------------------------
    def get_lambda_text(self, row: int, col: int) -> Optional[str]:
        """Render the lambda cached at ``(row, col)`` as text; ``None`` when the cell holds no lambda."""
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
    def add_merge(self, merge: CellRect) -> None:
        """Append a merge range to the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_cell_rect_array([merge], scope)
            _check(LIB.fm_sheet_add_merge(h, _uint(sheet, "sheet"), ptr))

    def remove_merge(self, merge: CellRect) -> None:
        """Remove every merge on the sheet overlapping ``merge``."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_cell_rect_array([merge], scope)
            _check(LIB.fm_sheet_remove_merge(h, _uint(sheet, "sheet"), ptr))

    def remove_merge_at(self, index: int) -> None:
        """Remove the merge at ``index`` on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_merge_at(h, _uint(sheet, "sheet"), _uint(index, "index")))

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
    def add_hyperlink(
        self,
        row: int,
        col: int,
        *,
        target: Optional[str],
        location: Optional[str] = None,
        display: Optional[str] = None,
        tooltip: Optional[str] = None,
    ) -> None:
        """Append a hyperlink on the cell ``(row, col)``.

        ``target`` is the external address (URL, file path); ``None`` makes a
        link inside the workbook, which then needs ``location`` (a reference
        such as ``"Sheet2!A1"`` or a defined name). ``location`` beside a
        ``target`` is the fragment within it.

        Raises:
          TypeError: when ``target`` and ``location`` are both ``None``.
        """
        self._add_hyperlink(CellRect(row, col, row, col), target, location, display, tooltip)

    def add_hyperlink_range(
        self,
        rect: CellRect,
        *,
        target: Optional[str],
        location: Optional[str] = None,
        display: Optional[str] = None,
        tooltip: Optional[str] = None,
    ) -> None:
        """Append a hyperlink covering the cells of ``rect``; the arguments read as in :meth:`add_hyperlink`."""
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

    def remove_hyperlink(self, row: int, col: int) -> None:
        """Remove every hyperlink anchored at ``(row, col)`` on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_hyperlink(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col")))

    def remove_hyperlink_at(self, index: int) -> None:
        """Remove the hyperlink at ``index`` on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_hyperlink_at(h, _uint(sheet, "sheet"), _uint(index, "index")))

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

    def get_hyperlink(self, row: int, col: int) -> Optional[Hyperlink]:
        """Return the hyperlink whose range starts at ``(row, col)``, or ``None``."""
        sheet = self._index()
        return self._workbook._lookup(
            LIB.fm_sheet_get_hyperlink,
            (_uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col")),
            S.HYPERLINK,
            _hyperlink_from,
        )

    # -- Comments ----------------------------------------------------------
    def get_comment(self, row: int, col: int) -> Optional[Comment]:
        """Return the comment at ``(row, col)`` or ``None`` if absent."""
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

    def set_comment(self, row: int, col: int, comment: Comment) -> None:
        """Insert or replace the comment (note) at ``(row, col)``.

        An empty ``comment.text`` writes an empty note, as Excel keeps one;
        :meth:`remove_comment` removes a note.
        """
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

    def remove_comment(self, row: int, col: int) -> None:
        """Remove the comment (note) at ``(row, col)``.

        Raises:
          NotFoundError: when the cell has no comment.
        """
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

    def get_comment_visible(self, row: int, col: int) -> Optional[bool]:
        """Return whether the note at the cell is shown without hovering, or ``None`` if absent."""
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

    def set_comment_visible(self, row: int, col: int, visible: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Show or hide the note at the cell; its box geometry is kept."""
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
        self, fn, leading: tuple, protected: ProtectedRange, *, returns_index: bool = False
    ) -> Optional[int]:
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PROTECTED_RANGE)
            S.PROTECTED_RANGE.pack(
                LIB,
                ptr,
                {
                    "spin_count": _uint(protected.spin_count, "spin_count"),
                    "ranges": _pack_cell_rect_array(protected.ranges, scope),
                    "range_count": _uint(len(protected.ranges), "range_count"),
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

    def update_protected_range(self, protected_range_index: int, protected: ProtectedRange) -> None:
        """Replace the ``protected_range_index``-th protected range of the sheet."""
        sheet = self._index()
        self._protected_range_call(
            LIB.fm_sheet_update_protected_range,
            (_uint(sheet, "sheet"), _uint(protected_range_index, "protected_range_index")),
            protected,
        )

    def remove_protected_range(self, protected_range_index: int) -> None:
        """Remove the ``protected_range_index``-th protected range of the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_remove_protected_range(
                h, _uint(sheet, "sheet"), _uint(protected_range_index, "protected_range_index")
            )
        )

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
    def clear_row_height(self, first: int, last: int) -> None:
        """Remove the height override on ``[first, last]`` (back to the sheet default)."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_clear_row_height(h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"))
        )

    def clear_column_width(self, first: int, last: int) -> None:
        """Remove the width override on ``[first, last]`` (back to the sheet default)."""
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

    def set_format_defaults(self, defaults: SheetFormatDefaults) -> None:
        """Replace the sheet's default column / row metrics.

        ``None`` for ``default_col_width`` / ``default_row_height`` clears the stored value;
        :data:`UNCHANGED` keeps it. A ``None`` ``custom_height`` / ``zero_height`` keeps the stored flag.
        """
        sheet = self._index()
        h = self._workbook._require()
        col_width, row_height = defaults.default_col_width, defaults.default_row_height
        if col_width is UNCHANGED or row_height is UNCHANGED:
            stored = self.get_format_defaults()
            col_width = stored.default_col_width if col_width is UNCHANGED else col_width
            row_height = stored.default_row_height if row_height is UNCHANGED else row_height
        if defaults.custom_height is None or defaults.zero_height is None:
            stored = self.get_format_defaults()
            custom_height = stored.custom_height if defaults.custom_height is None else defaults.custom_height
            zero_height = stored.zero_height if defaults.zero_height is None else defaults.zero_height
        else:
            custom_height, zero_height = defaults.custom_height, defaults.zero_height
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHEET_FORMAT_DEFAULTS)
            S.SHEET_FORMAT_DEFAULTS.pack(
                LIB,
                ptr,
                {
                    "default_col_width": 0.0 if col_width is None else float(col_width),
                    "default_row_height": 0.0 if row_height is None else float(row_height),
                    "base_col_width": float(defaults.base_col_width),
                    "has_default_col_width": 0 if col_width is None else 1,
                    "has_default_row_height": 0 if row_height is None else 1,
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

    def get_cell_rect_pt(self, cell_range: CellRect, mode: GeometryMode = GeometryMode.DISPLAY) -> RectPt:
        """Return the rectangle covering ``cell_range`` in points from the top-left of A1."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, cell_range)
            ptr = scope.struct(S.RECT_PT)
            _check(LIB.fm_sheet_get_cell_rect_pt(h, _uint(sheet, "sheet_index"), rp, _sint(mode, "mode"), ptr))
            return RectPt(**S.RECT_PT.unpack(LIB, ptr))

    def get_column_width_pt(self, col: int, mode: GeometryMode = GeometryMode.DISPLAY) -> float:
        """Return the effective width of ``col`` in points (0 when hidden)."""
        sheet = self._index()
        h = self._workbook._require()
        return self._read_f64_out(
            LIB.fm_sheet_get_column_width_pt, h, _uint(sheet, "sheet_index"), _uint(col, "col"), _sint(mode, "mode")
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
        """Convert a stored column width (padding included) to points, snapped to whole screen pixels."""
        sheet = self._index()
        h = self._workbook._require()
        return self._read_f64_out(
            LIB.fm_sheet_convert_column_chars_to_pt, h, _uint(sheet, "sheet_index"), _sint(mode, "mode"), float(chars)
        )

    def convert_column_pt_to_chars(self, pt: float, mode: GeometryMode = GeometryMode.DISPLAY) -> float:
        """Convert points to the stored column width Excel writes for that column."""
        sheet = self._index()
        h = self._workbook._require()
        return self._read_f64_out(
            LIB.fm_sheet_convert_column_pt_to_chars, h, _uint(sheet, "sheet_index"), _sint(mode, "mode"), float(pt)
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
    def get_formula(self, row_or_a1: str) -> Optional[str]: ...

    @overload
    def get_formula(self, row_or_a1: int, col: int) -> Optional[str]: ...

    def get_formula(self, row_or_a1: Union[int, str], col: Optional[int] = None) -> Optional[str]:
        """Return the stored A1 formula (leading ``=`` included), or ``None`` for a non-formula cell.

        The cell is either one A1 address (``get_formula("B3")``) or a 0-based
        ``(row, col)`` pair.
        """
        sheet = self._index()
        row, col = self._address(row_or_a1, col)
        return self._read_formula_text(LIB.fm_sheet_get_formula, sheet, row, col)

    def get_formula_r1c1(self, row: int, col: int) -> Optional[str]:
        """Return the formula in R1C1 notation relative to the cell, or ``None`` for a non-formula cell."""
        sheet = self._index()
        return self._read_formula_text(LIB.fm_sheet_get_formula_r1c1, sheet, row, col)

    def get_cells_in_range(
        self,
        cell_range: CellRect,
        cursor: Optional[int] = None,
        limit: Optional[int] = None,
    ) -> CellsInRange:
        """Return one page of the populated cells in ``cell_range`` (row-major).

        Pass ``next_cursor`` back to fetch the next page; it is ``None``
        when the rectangle holds no further populated cell. ``limit`` of
        ``None`` or ``0`` means the engine maximum.
        """
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
    def validate_value(self, row: int, col: int, value: Value) -> ValidationOutcome:
        """Check ``value`` against the data-validation rule covering ``(row, col)`` without storing it."""
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
        text: str = UNCHANGED,
        mentions: Sequence[Mention] = UNCHANGED,
    ) -> None:
        """Replace the text and / or mentions of the comment ``comment_id``; an omitted argument is kept.

        Raises:
          NotFoundError: when the sheet has no comment ``comment_id``.
        """
        sheet = self._index()
        if text is UNCHANGED or mentions is UNCHANGED:
            current = self.get_threaded_comment(comment_id)
            if current is None:
                raise FormulonError(
                    StatusCode.NOT_FOUND, _diagnostic_override=(f"no threaded comment has the id {comment_id!r}", "")
                )
            text = current.text if text is UNCHANGED else text
            mentions = current.mentions if mentions is UNCHANGED else mentions
        h = self._workbook._require()
        with _ArgScope() as scope:
            id_ptr = scope.cstr(comment_id)
            text_ptr = scope.cstr(text)
            mentions_ptr = _pack_mentions(mentions, scope)
            _check(
                LIB.fm_sheet_update_threaded_comment(
                    h, _uint(sheet, "sheet_index"), id_ptr, text_ptr, mentions_ptr, _uint(len(mentions), "count")
                )
            )

    def set_thread_resolved(self, thread_id: str, done: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the resolved state of the thread opened by ``thread_id``."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            id_ptr = scope.cstr(thread_id)
            _check(LIB.fm_sheet_set_thread_resolved(h, _uint(sheet, "sheet_index"), id_ptr, 1 if done else 0))

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

    def list_drawing_objects(self) -> List[DrawingObject]:
        """Return the anchored objects of the sheet's drawing in document order."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(LIB.fm_sheet_get_drawing_object_count, h, _uint(sheet, "sheet_index"))
        with _ArgScope() as scope:
            ptr = scope.struct(S.DRAWING_OBJECT)
            out: List[DrawingObject] = []
            for i in range(n):
                S.zero_struct(LIB, S.DRAWING_OBJECT, ptr)
                _check(LIB.fm_sheet_get_drawing_object_at(h, _uint(sheet, "sheet_index"), _uint(i, "idx"), ptr))
                d = S.DRAWING_OBJECT.unpack(LIB, ptr)
                out.append(_drawing_object_from(d))
        return out

    def get_drawing_object(self, object_id: int) -> Optional[DrawingObject]:
        """Return the drawing object ``object_id``, or ``None`` when the sheet's drawing has none."""
        sheet = self._index()
        return self._workbook._lookup(
            LIB.fm_sheet_get_drawing_object,
            (_uint(sheet, "sheet_index"), _uint(object_id, "object_id")),
            S.DRAWING_OBJECT,
            _drawing_object_from,
        )

    def get_image(self, object_id: int) -> Optional[bytes]:
        """Return the bytes of the picture ``object_id`` (copied out of engine memory).

        ``None`` when the sheet has no such picture with a present media part.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            bytes_ptr = scope.out_ptr()
            len_ptr = scope.out_ptr()
            info = scope.struct(S.IMAGE_INFO)
            _check(
                LIB.fm_sheet_get_image(
                    h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), present, bytes_ptr, len_ptr, info
                )
            )
            if not LIB.read_u32(present):
                return None
            return LIB.read_bytes(LIB.read_u32(bytes_ptr), LIB.read_u32(len_ptr))

    def add_image(
        self,
        data: Union[bytes, bytearray, memoryview],
        *,
        name: str = "",
        description: str = "",
        anchor_kind: AnchorKind = AnchorKind.ONE_CELL,
        edit_as: AnchorEditAs = AnchorEditAs.TWO_CELL,
        row: int = 0,
        col: int = 0,
        row_offset_emu: int = 0,
        col_offset_emu: int = 0,
        width_emu: int = 0,
        height_emu: int = 0,
    ) -> int:
        """Insert an image anchored at ``(row, col)`` and return its object id.

        A ``width_emu`` or ``height_emu`` of ``0`` means the image's pixel size
        times 9525. ``edit_as`` is honoured for a two-cell anchor only.
        """
        sheet = self._index()
        h = self._workbook._require()
        raw = bytes(data)
        with _ArgScope() as scope:
            out = scope.out_ptr()
            buf = scope.buffer(raw)
            ptr = scope.struct_array(S.IMAGE_INSERT, 1)
            S.IMAGE_INSERT.pack(
                LIB,
                ptr,
                {
                    "anchor": _image_anchor_fields(
                        anchor_kind,
                        edit_as,
                        row,
                        col,
                        row_offset_emu,
                        col_offset_emu,
                        width_emu,
                        height_emu,
                    )
                },
            )
            S.write_str_field(LIB, ptr, S.IMAGE_INSERT, "name", name, scope)
            S.write_str_field(LIB, ptr, S.IMAGE_INSERT, "description", description, scope)
            _check(LIB.fm_sheet_add_image(h, _uint(sheet, "sheet_index"), buf, _uint(len(raw), "len"), ptr, out))
            return LIB.read_u32(out)

    def remove_drawing_object(self, object_id: int) -> None:
        """Remove the picture or chart ``object_id`` and the parts only it reaches; other drawing content is kept."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_drawing(h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id")))

    def set_drawing_anchor(
        self,
        object_id: int,
        *,
        anchor_kind: AnchorKind = AnchorKind.ONE_CELL,
        edit_as: AnchorEditAs = AnchorEditAs.TWO_CELL,
        row: int = 0,
        col: int = 0,
        row_offset_emu: int = 0,
        col_offset_emu: int = 0,
        width_emu: int = 0,
        height_emu: int = 0,
    ) -> None:
        """Move and resize the picture or chart ``object_id``, keeping its id and every other element.

        The placement is in the quantities :meth:`list_drawing_objects`
        reports: ``(row, col)`` and the offsets are the top-left marker of the
        anchor box, ``width_emu`` and ``height_emu`` the object's own unrotated
        size; ``0`` keeps the current size. ``edit_as`` is
        honoured for a two-cell anchor only.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.DRAWING_ANCHOR, 1)
            anchor = _image_anchor_fields(
                anchor_kind,
                edit_as,
                row,
                col,
                row_offset_emu,
                col_offset_emu,
                width_emu,
                height_emu,
            )
            S.DRAWING_ANCHOR.pack(LIB, ptr, anchor)
            _check(LIB.fm_sheet_set_drawing_anchor(h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), ptr))

    def set_drawing_z_order(self, object_id: int, z_order_index: int) -> None:
        """Make the picture or chart ``object_id`` item ``z_order_index`` of :meth:`list_drawing_objects` (``0`` is the back)."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_drawing_z_order(
                h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), _uint(z_order_index, "z_order_index")
            )
        )

    def set_drawing_name(self, object_id: int, name: str) -> None:
        """Rename the picture or chart ``object_id``. Duplicate names are allowed; ``name`` must be non-empty."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            _check(
                LIB.fm_sheet_set_drawing_name(h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), name_ptr)
            )

    def snapshot_drawing(self, object_id: int) -> bytes:
        """Capture the picture or chart ``object_id`` as opaque bytes for :meth:`restore_drawing`.

        The capture holds the object's list position, element and every part it reaches.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            bytes_ptr = scope.out_ptr()
            len_ptr = scope.out_ptr()
            data_ptr = 0
            try:
                _check(
                    LIB.fm_sheet_snapshot_drawing(
                        h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), bytes_ptr, len_ptr
                    )
                )
                data_ptr = LIB.read_u32(bytes_ptr)
                return LIB.read_bytes(data_ptr, LIB.read_u32(len_ptr))
            finally:
                if data_ptr:
                    LIB.fm_buffer_destroy(data_ptr)

    def restore_drawing(self, data: Union[bytes, bytearray, memoryview], *, new_id: bool = False) -> int:
        """Put a :meth:`snapshot_drawing` capture back and return the object's id.

        An object with the same id is replaced and the capture returns to its
        recorded list position. With ``new_id`` it is added on top as a copy
        with a fresh id.
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
            return LIB.read_u32(out)

    def list_shapes(self) -> List[Shape]:
        """Return the shapes, connectors and groups of the sheet's drawing in z-order (back first).

        Children of a group are not listed.
        """
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(LIB.fm_sheet_get_shape_count, h, _uint(sheet, "sheet_index"))
        with _ArgScope() as scope:
            ptr = scope.struct(S.SHAPE)
            out: List[Shape] = []
            for i in range(n):
                S.zero_struct(LIB, S.SHAPE, ptr)
                _check(LIB.fm_sheet_get_shape_at(h, _uint(sheet, "sheet_index"), _uint(i, "idx"), ptr))
                d = S.SHAPE.unpack(LIB, ptr)
                out.append(_shape_from(d))
        return out

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
    ) -> int:
        """Append a shape on top of the sheet's drawing and return its object id.

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
            return LIB.read_u32(out)

    def update_shape(
        self,
        object_id: int,
        *,
        anchor_kind: Optional[AnchorKind] = None,
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
    ) -> None:
        """Apply the arguments that are not ``None`` to the shape ``object_id``; the rest of its XML is kept.

        The anchor is applied as a whole when ``anchor_kind`` is given, and its
        kind cannot change. Text applies to a shape only and keeps the first
        paragraph's properties and first run's formatting.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_shape(
                scope, name, preset_geometry, text, text_box, anchor_kind, edit_as, from_row, from_col, from_row_offset_emu,
                from_col_offset_emu, to_row, to_col, to_row_offset_emu, to_col_offset_emu, width_emu, height_emu, fill_kind, fill_rgb, line_kind,
                line_rgb, line_width_emu, rotation, hidden,
            )  # fmt: skip
            _check(LIB.fm_sheet_update_shape(h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), ptr))

    def get_shape(self, object_id: int) -> Optional[Shape]:
        """Return the shape ``object_id``, or ``None`` when the sheet's drawing has none."""
        sheet = self._index()
        return self._workbook._lookup(
            LIB.fm_sheet_get_shape,
            (_uint(sheet, "sheet_index"), _uint(object_id, "object_id")),
            S.SHAPE,
            _shape_from,
        )

    def remove_shape(self, object_id: int) -> None:
        """Remove the shape, connector or group ``object_id`` with its children."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_shape(h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id")))

    def list_slicers(self) -> List[Slicer]:
        """Return the table slicers hosted by the sheet; slicers over a pivot cache are not listed."""
        sheet = self._index()
        h = self._workbook._require()
        n = _read_count(LIB.fm_sheet_get_slicer_count, h, _uint(sheet, "sheet_index"))
        with _ArgScope() as scope:
            ptr = scope.struct(S.SLICER)
            out: List[Slicer] = []
            for i in range(n):
                S.zero_struct(LIB, S.SLICER, ptr)
                _check(LIB.fm_sheet_get_slicer_at(h, _uint(sheet, "sheet_index"), _uint(i, "idx"), ptr))
                d = S.SLICER.unpack(LIB, ptr)
                out.append(_slicer_from(d))
        return out

    def add_slicer(
        self,
        table: int,
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
    ) -> int:
        """Add a slicer over column ``column`` of table ``table`` and return its frame's object id.

        The frame is the two-cell anchor ``from_*`` .. ``to_*`` (0-based cells
        plus EMU offsets), named as in :meth:`add_shape`. The name and caption
        default to the column name (the name is made unique),
        ``row_height_emu`` to 241300 and ``column_count`` to 1. The table may
        live on another sheet.
        """
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
                    h, _uint(sheet, "sheet_index"), _uint(table, "table_index"), _uint(column, "column_index"), ptr, out
                )
            )
            return LIB.read_u32(out)

    def get_slicer(self, object_id: int) -> Optional[Slicer]:
        """Return the slicer ``object_id`` (the id :meth:`add_slicer` returns), or ``None``.

        A slicer without a drawing frame has no id and is not found.
        """
        sheet = self._index()
        return self._workbook._lookup(
            LIB.fm_sheet_get_slicer,
            (_uint(sheet, "sheet_index"), _uint(object_id, "object_id")),
            S.SLICER,
            _slicer_from,
        )

    def update_slicer(
        self,
        object_id: int,
        *,
        from_row: Optional[int] = None,
        from_col: Optional[int] = None,
        to_row: Optional[int] = None,
        to_col: Optional[int] = None,
        from_row_offset_emu: int = 0,
        from_col_offset_emu: int = 0,
        to_row_offset_emu: int = 0,
        to_col_offset_emu: int = 0,
        caption: Optional[str] = None,
        style: Optional[str] = None,
        column_count: Optional[int] = None,
        row_height_emu: Optional[int] = None,
        sort_order: Optional[SlicerSortOrder] = None,
        cross_filter: Optional[SlicerCrossFilter] = None,
    ) -> None:
        """Apply the arguments that are not ``None`` to slicer ``object_id``; the rest of its XML is kept.

        An empty ``style`` removes the style. The four ``from_*`` / ``to_*``
        cells, given together, move the slicer's two-cell frame (see
        :meth:`add_slicer`); the EMU offsets apply only with them.
        """
        sheet = self._index()
        h = self._workbook._require()
        corners = (from_row, from_col, to_row, to_col)
        anchor: Optional[Tuple[int, ...]] = None
        if any(c is not None for c in corners):
            if any(c is None for c in corners):
                raise TypeError("formulon: from_row, from_col, to_row and to_col move the frame together")
            anchor = (
                *cast(Tuple[int, int, int, int], corners),
                from_row_offset_emu, from_col_offset_emu, to_row_offset_emu, to_col_offset_emu,
            )  # fmt: skip
        with _ArgScope() as scope:
            ptr = _pack_slicer(
                scope, None, caption, style, column_count, row_height_emu, anchor, sort_order, cross_filter
            )
            _check(LIB.fm_sheet_update_slicer(h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), ptr))

    def remove_slicer(self, object_id: int) -> None:
        """Remove slicer ``object_id`` with its drawing frame, and its cache when no other slicer uses it."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_slicer(h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id")))

    def add_data_table(
        self,
        rect: CellRect,
        *,
        row_input: Optional[Tuple[int, int]] = None,
        col_input: Optional[Tuple[int, int]] = None,
    ) -> None:
        """Create a data table whose result range is ``rect``.

        The formulas sit in the row above the range and the substitution values in the column left of it (a
        one-variable table), or the other way round; ``row_input`` / ``col_input`` are the ``(row, col)`` input
        cells that receive the values of the row above / the column left. Give one for a one-variable table, both
        for a two-variable one. The result cells read 0 until the next recalc and refuse writes until the table
        is removed.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.DATA_TABLE, 1)
            values: Dict[str, object] = {"rect": _cell_rect_fields(rect)}
            for prefix, cell in (("row_input", row_input), ("col_input", col_input)):
                if cell is not None:
                    values["has_" + prefix] = 1
                    values[prefix + "_row"] = _uint(cell[0], prefix + "_row")
                    values[prefix + "_col"] = _uint(cell[1], prefix + "_col")
            S.DATA_TABLE.pack(LIB, ptr, values)
            _check(LIB.fm_sheet_add_data_table(h, _uint(sheet, "sheet_index"), ptr))

    def get_data_table(self, row: int, col: int) -> Optional[DataTable]:
        """Return the data table whose result range contains ``(row, col)``, or ``None``."""
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

    def remove_data_table(self, row: int, col: int) -> None:
        """Remove the data table containing ``(row, col)`` and blank its result cells."""
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

    def update_sparkline_group(self, group_index: int, group: SparklineGroup) -> None:
        """Rewrite the modelled attributes, colours and date range of group ``group_index`` from ``group``.

        Its sparklines (``group.sparklines`` is ignored) and every attribute or
        child outside the model are kept; read, modify and write back.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = _pack_sparkline_group(scope, group)
            _check(LIB.fm_sheet_update_sparkline_group(h, _uint(sheet, "sheet_index"), _uint(group_index, "idx"), ptr))

    def set_sparkline_group_sparklines(self, group_index: int, sparklines: Sequence[Sparkline]) -> None:
        """Replace the sparklines of group ``group_index`` with ``sparklines`` (at least one)."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            items = _pack_sparklines(scope, sparklines)
            _check(
                LIB.fm_sheet_set_sparkline_group_sparklines(
                    h, _uint(sheet, "sheet_index"), _uint(group_index, "idx"), items, _uint(len(sparklines), "count")
                )
            )

    def remove_sparkline_group(self, group_index: int) -> None:
        """Remove sparkline group ``group_index``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_sparkline_group(h, _uint(sheet, "sheet_index"), _uint(group_index, "idx")))

    def get_chart(self, object_id: int) -> Optional[Chart]:
        """Return the typed view of the chart ``object_id``: plots, series, axes, legend and source.

        Returns ``None`` when the sheet has no chart with that id; raises
        :class:`FormulonError` when the chart part does not parse.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            present = scope.out_ptr()
            out = scope.out_ptr()
            _check(LIB.fm_sheet_get_chart(h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), present, out))
            if not LIB.read_u32(present):
                return None
            chart = LIB.read_u32(out)
        try:
            return _read_chart(chart)
        finally:
            LIB.fm_chart_destroy(chart)

    def get_chart_data(self, object_id: int) -> ChartData:
        """Resolve the chart's series against the current cell values by Excel's rules.

        Raises:
          NotFoundError: when the sheet has no chart ``object_id``.
        """
        sheet = self._index()
        h = self._workbook._require()
        data = _read_count(LIB.fm_sheet_get_chart_data, h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"))
        try:
            return _read_chart_data(data)
        finally:
            LIB.fm_chart_data_destroy(data)

    def get_chart_preview(self, object_id: int, overrides: Optional[ChartPreviewOverridesInput] = None) -> ChartData:
        """Resolve the data of a copy of the chart with ``overrides`` applied; the workbook is unchanged.

        Series are rebuilt as :meth:`set_chart_source` does: over the
        overridden ``source`` (an unset direction or header is auto), else over
        the chart's derived range.

        Raises:
          NotFoundError: when the sheet has no chart ``object_id``.
        """
        sheet = self._index()
        h = self._workbook._require()
        given = {k: v for k, v in vars(overrides).items() if v is not None} if overrides is not None else {}
        with _ArgScope() as scope:
            fields = _encode_patch(_PREVIEW_OVERRIDE_SPEC, given, ("source",), "preview override", scope)
            if "source" in given:
                fields["source"] = scope.cstr(given["source"])
            ptr = scope.struct(S.CHART_PREVIEW_OVERRIDES)
            S.CHART_PREVIEW_OVERRIDES.pack(LIB, ptr, fields)
            data = _read_count(
                LIB.fm_sheet_get_chart_preview,
                h,
                _uint(sheet, "sheet_index"),
                _uint(object_id, "object_id"),
                ptr,
            )
        try:
            return _read_chart_data(data)
        finally:
            if data:
                LIB.fm_chart_data_destroy(data)

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
        anchor_kind: AnchorKind = AnchorKind.TWO_CELL,
        edit_as: AnchorEditAs = AnchorEditAs.TWO_CELL,
        row: int = 0,
        col: int = 0,
        row_offset_emu: int = 0,
        col_offset_emu: int = 0,
        width_emu: int = 0,
        height_emu: int = 0,
    ) -> int:
        """Create a chart over the A1 rectangle ``source`` and return its object id.

        ``source`` is sheet-qualified or on the sheet. ``grouping`` applies to
        column, bar, line and area charts. ``header_row`` / ``header_column``
        ``None`` follow Excel's detection. An omitted ``title`` is Excel's
        automatic title and an empty one is no title; ``legend`` is ``None``
        for no legend. The placement reads as
        for :meth:`add_image` (an empty ``name`` becomes ``Chart <id>``),
        except that a zero ``width_emu`` or ``height_emu`` means Excel's default
        360 x 216 pt frame and the anchor defaults to two-cell.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = scope.struct_array(S.CHART_INSERT, 1)
            placement = {
                "name": scope.nonempty_cstr(name),
                "description": scope.nonempty_cstr(description),
                "anchor": _image_anchor_fields(
                    anchor_kind, edit_as, row, col, row_offset_emu, col_offset_emu, width_emu, height_emu
                ),
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
            return LIB.read_u32(out)

    def set_chart_type(
        self,
        object_id: int,
        plot: int,
        kind: ChartInsertKind,
        grouping: ChartGrouping = ChartGrouping.CLUSTERED,
    ) -> None:
        """Replace plot group ``plot`` with ``kind``, keeping each series' references, idx, order and explicit colors."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_chart_type(
                h,
                _uint(sheet, "sheet_index"),
                _uint(object_id, "object_id"),
                _uint(plot, "plot"),
                _sint(kind, "kind"),
                _sint(grouping, "grouping"),
            )
        )

    def set_chart_title(self, object_id: int, title: Union[str, None, ChartTitle]) -> None:
        """Set the chart title to ``title``; ``None`` removes it.

        A :class:`ChartTitle` sets that kind, so ``ChartTitle(ChartTitleKind.AUTO, "", "")``
        returns the title to Excel's automatic one.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_TITLE, 1)
            fields = {"kind": _sint(_CHART_TITLE_NONE, "kind")} if title is None else _title_fields(title, scope)
            S.CHART_TITLE.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_set_chart_title(h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), ptr))

    def set_chart_legend(
        self,
        object_id: int,
        position: Optional[ChartLegendPosition],
        *,
        overlay: bool = False,
        hidden_series: Optional[Sequence[int]] = None,
    ) -> None:
        """Set the legend position (``None`` removes the legend) and overlay.

        ``hidden_series`` replaces the hidden legend entries (series ``c:idx``
        values); ``None`` leaves them alone.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_LEGEND_PATCH, 1)
            fields: Dict[str, object] = {
                "position": _sint(_nullable_enum(position, "position"), "position"),
                "overlay": 1 if overlay else 0,
            }
            if hidden_series is not None:
                fields["hidden_series_op"] = 2
                fields["hidden_series"] = _id_array(hidden_series, "hidden_series", scope)
                fields["hidden_series_count"] = _uint(len(hidden_series), "hidden_series_count")
            S.CHART_LEGEND_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_set_chart_legend(h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), ptr))

    def set_chart_axis(self, object_id: int, axis_id: int, patch: ChartAxisPatch) -> None:
        """Patch axis ``axis_id``: an omitted key is kept, ``None`` removes the element."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_AXIS_PATCH, 1)
            fields = _encode_axis_patch(patch, scope)
            S.CHART_AXIS_PATCH.pack(LIB, ptr, fields)
            _check(
                LIB.fm_sheet_set_chart_axis(
                    h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), _uint(axis_id, "axis_id"), ptr
                )
            )

    def set_chart_display(self, object_id: int, patch: ChartDisplayPatch) -> None:
        """Patch how blanks plot, whether hidden cells are skipped and whether ``#N/A`` plots as blank.

        An omitted key is kept, ``None`` removes the element.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            fields = _encode_patch(_DISPLAY_PATCH_SPEC, patch, (), "display patch", scope)
            ptr = scope.struct(S.CHART_DISPLAY_PATCH)
            S.CHART_DISPLAY_PATCH.pack(LIB, ptr, fields)
            _check(LIB.fm_sheet_set_chart_display(h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), ptr))

    def set_chart_series(self, object_id: int, plot: int, series_index: int, patch: ChartSeriesPatch) -> None:
        """Patch the series with ``c:idx`` ``series_index`` of plot group ``plot``; a changed reference refreshes its cache."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_SERIES_PATCH, 1)
            fields = _encode_series_patch(patch, scope)
            S.CHART_SERIES_PATCH.pack(LIB, ptr, fields)
            _check(
                LIB.fm_sheet_set_chart_series(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(object_id, "object_id"),
                    _uint(plot, "plot"),
                    _uint(series_index, "series_index"),
                    ptr,
                )
            )

    def set_chart_point(
        self,
        object_id: int,
        plot: int,
        series_index: int,
        point_index: int,
        patch: Optional[ChartPointPatch],
    ) -> None:
        """Patch point ``point_index`` of a series; ``None`` (or an empty patch) removes the point's formatting."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_POINT_PATCH, 1)
            fields = _encode_point_patch(patch, scope)
            S.CHART_POINT_PATCH.pack(LIB, ptr, fields)
            _check(
                LIB.fm_sheet_set_chart_point(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(object_id, "object_id"),
                    _uint(plot, "plot"),
                    _uint(series_index, "series_index"),
                    _uint(point_index, "point_index"),
                    ptr,
                )
            )

    def set_chart_plot(self, object_id: int, plot: int, patch: ChartPlotPatch) -> None:
        """Patch plot group ``plot``; a field the plot type does not have is rejected."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct_array(S.CHART_PLOT_PATCH, 1)
            fields = _encode_patch(_PLOT_PATCH_SPEC, patch, (), "plot patch", scope)
            S.CHART_PLOT_PATCH.pack(LIB, ptr, fields)
            _check(
                LIB.fm_sheet_set_chart_plot(
                    h, _uint(sheet, "sheet_index"), _uint(object_id, "object_id"), _uint(plot, "plot"), ptr
                )
            )

    def set_chart_filter(
        self,
        object_id: int,
        hidden_series: Sequence[int],
        hidden_categories: Sequence[int],
    ) -> None:
        """Replace the chart filter: the listed series (by ``c:idx``) and category positions are hidden."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            _check(
                LIB.fm_sheet_set_chart_filter(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(object_id, "object_id"),
                    _id_array(hidden_series, "hidden_series", scope),
                    _uint(len(hidden_series), "hidden_series_count"),
                    _id_array(hidden_categories, "hidden_categories", scope),
                    _uint(len(hidden_categories), "hidden_category_count"),
                )
            )

    def add_chart_series(self, object_id: int, plot: int, patch: ChartSeriesPatch) -> int:
        """Append a series to plot group ``plot`` and return its ``c:idx``.

        ``patch`` must set ``values`` (``y`` for a scatter plot).
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = scope.struct_array(S.CHART_SERIES_PATCH, 1)
            fields = _encode_series_patch(patch, scope)
            S.CHART_SERIES_PATCH.pack(LIB, ptr, fields)
            _check(
                LIB.fm_sheet_add_chart_series(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(object_id, "object_id"),
                    _uint(plot, "plot"),
                    ptr,
                    out,
                )
            )
            return LIB.read_u32(out)

    def remove_chart_series(self, object_id: int, plot: int, series_index: int) -> None:
        """Remove the series with ``c:idx`` ``series_index`` from plot group ``plot``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_remove_chart_series(
                h,
                _uint(sheet, "sheet_index"),
                _uint(object_id, "object_id"),
                _uint(plot, "plot"),
                _uint(series_index, "series_index"),
            )
        )

    def set_chart_source(
        self,
        object_id: int,
        source: str,
        *,
        direction: ChartSeriesDirection = ChartSeriesDirection.AUTO,
        header_row: Optional[bool] = None,
        header_column: Optional[bool] = None,
    ) -> None:
        """Rebuild every series from ``source`` by the insert splitting rules.

        The chart kind and each series' formatting (by idx) are kept and the
        filter is cleared. ``header_row`` / ``header_column`` ``None`` follow
        Excel's detection.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            _check(
                LIB.fm_sheet_set_chart_source(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(object_id, "object_id"),
                    scope.cstr(source),
                    _sint(direction, "direction"),
                    _sint(_header_mode(header_row), "header_row"),
                    _sint(_header_mode(header_column), "header_column"),
                )
            )

    def get_merges_in_range(self, cell_range: CellRect) -> List[CellRect]:
        """Return every merge range on the sheet that intersects ``cell_range``."""
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
            for i in range(min(total, LIB.read_u32(count_ptr))):
                out.append(_decode_cell_rect_at(buf + i * size))
            return out

    def get_display_text(self, row: int, col: int) -> DisplayText:
        """Return the text Excel displays for the cell and the rendering outcome."""
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

    def set_zoom(self, zoom_scale: int) -> None:
        """Set the sheet zoom percentage (clamped to ``[10, 400]``)."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_zoom(h, _uint(sheet, "sheet_index"), _uint(zoom_scale, "zoom_scale")))

    def set_freeze(self, freeze_rows: int, freeze_cols: int) -> None:
        """Set the frozen pane in ``(rows, cols)``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_freeze(
                h, _uint(sheet, "sheet_index"), _uint(freeze_rows, "freeze_rows"), _uint(freeze_cols, "freeze_cols")
            )
        )

    def set_visibility(self, visibility: SheetVisibility) -> None:
        """Set the sheet tab's visibility to one of the three states.

        :attr:`SheetVisibility.VERY_HIDDEN` hides the tab from Excel's
        Unhide dialog as well.

        Raises:
          FormulonError: with ``StatusCode.INVALID_ARGUMENT`` for an unknown state or
            when hiding the only visible sheet.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_visibility(h, _uint(sheet, "sheet_index"), _sint(visibility, "visibility")))

    def set_show_grid_lines(self, show: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the sheet's ``showGridLines`` flag."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_show_grid_lines(h, _uint(sheet, "sheet_index"), 1 if show else 0))

    def set_show_row_col_headers(self, show: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the sheet's ``showRowColHeaders`` flag."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_show_row_col_headers(h, _uint(sheet, "sheet_index"), 1 if show else 0))

    def set_show_zeros(self, show: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the sheet's ``showZeros`` flag."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_show_zeros(h, _uint(sheet, "sheet_index"), 1 if show else 0))

    def set_show_formulas(self, show: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the sheet's ``showFormulas`` flag."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_show_formulas(h, _uint(sheet, "sheet_index"), 1 if show else 0))

    def set_show_outline_symbols(self, show: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the sheet's ``showOutlineSymbols`` flag."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_show_outline_symbols(h, _uint(sheet, "sheet_index"), 1 if show else 0))

    def set_right_to_left(self, right_to_left: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the sheet's ``rightToLeft`` flag."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_right_to_left(h, _uint(sheet, "sheet_index"), 1 if right_to_left else 0))

    def set_top_left(self, row: int, col: int) -> None:
        """Set the sheet view's scroll origin (``topLeftCell``)."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_top_left(h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col")))

    def set_pane_top_left(self, row: int, col: int) -> None:
        """Set the first visible cell of a frozen or split view's lower-right pane.

        The cell must lie past any frozen rows and columns, and the sheet must
        be frozen or split. Changing the freeze resets it to the first unfrozen
        cell.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_pane_top_left(h, _uint(sheet, "sheet_index"), _uint(row, "row"), _uint(col, "col")))

    def set_split(self, x_twips: float, y_twips: float) -> None:
        """Split the view at ``x_twips`` / ``y_twips`` (0 for none on that axis).

        ``(0, 0)`` removes the split. A split replaces any freeze, as
        :meth:`set_freeze` replaces a split. Negative or non-finite
        values raise and leave the view unchanged.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_split(h, _uint(sheet, "sheet_index"), float(x_twips), float(y_twips)))

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

    def set_selection(self, selection: SheetSelection) -> None:
        """Replace the active pane's selection; empty ``ranges`` selects the active cell alone."""
        sheet = self._index()
        h = self._workbook._require()
        active_row, active_col, ranges = selection.active_row, selection.active_col, selection.ranges
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

    def set_tab_selected(self, selected: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the sheet's ``tabSelected`` flag."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_tab_selected(h, _uint(sheet, "sheet_index"), 1 if selected else 0))

    def set_view_mode(self, mode: SheetViewMode) -> None:
        """Set the sheet's ``<sheetView view="...">`` mode."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_view_mode(h, _uint(sheet, "sheet_index"), _sint(mode, "mode")))

    def get_auto_filter_xml(self) -> Optional[str]:
        """Return the sheet's ``<autoFilter>`` XML fragment, or ``None`` without one.

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
        """Replace the sheet's ``<autoFilter>`` XML fragment.

        ``None`` (or ``""``) removes the AutoFilter. Other input must be a
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
        """Return the sheet's ``<pageSetup>`` fragment, or ``None`` without one."""
        return self._get_print_xml("fm_sheet_get_page_setup_xml")

    def set_page_setup_xml(self, xml: Optional[str]) -> None:
        """Replace the sheet's ``<pageSetup>`` fragment; ``None`` (or ``""``) removes it.

        Note that ``fitToPage`` is not a ``<pageSetup>`` attribute -- use
        :meth:`set_fit_to_page`. A fragment carrying ``r:id`` is rejected
        unless the sheet already has a printerSettings part, because the
        reference would otherwise dangle and Excel would repair the file.
        """
        self._set_print_xml("fm_sheet_set_page_setup_xml", xml)

    def get_page_margins_xml(self) -> Optional[str]:
        """Return the sheet's ``<pageMargins>`` fragment, or ``None`` without one."""
        return self._get_print_xml("fm_sheet_get_page_margins_xml")

    def set_page_margins_xml(self, xml: Optional[str]) -> None:
        """Replace the sheet's ``<pageMargins>`` fragment; ``None`` (or ``""``) removes it."""
        self._set_print_xml("fm_sheet_set_page_margins_xml", xml)

    def get_print_options_xml(self) -> Optional[str]:
        """Return the sheet's ``<printOptions>`` fragment, or ``None`` without one."""
        return self._get_print_xml("fm_sheet_get_print_options_xml")

    def set_print_options_xml(self, xml: Optional[str]) -> None:
        """Replace the sheet's ``<printOptions>`` fragment; ``None`` (or ``""``) removes it."""
        self._set_print_xml("fm_sheet_set_print_options_xml", xml)

    def get_header_footer_xml(self) -> Optional[str]:
        """Return the sheet's ``<headerFooter>`` fragment, or ``None`` without one."""
        return self._get_print_xml("fm_sheet_get_header_footer_xml")

    def set_header_footer_xml(self, xml: Optional[str]) -> None:
        """Replace the sheet's ``<headerFooter>`` fragment; ``None`` (or ``""``) removes it.

        Header and footer codes are introduced by ``&``, which lives in XML
        element text and so must arrive escaped: Excel writes
        ``<oddHeader>&amp;C...</oddHeader>``. Pass decoded text to
        :meth:`set_header_footer` instead to let the engine escape it.
        """
        self._set_print_xml("fm_sheet_set_header_footer_xml", xml)

    def get_sheet_pr_xml(self) -> Optional[str]:
        """Return the sheet's ``<sheetPr>`` fragment, or ``None`` without one.

        ``<sheetPr>`` is not print-only: it also carries the tab colour and
        the VBA code name, which is why toggling fit-to-page has its own
        entry point rather than going through this setter.
        """
        return self._get_print_xml("fm_sheet_get_sheet_pr_xml")

    def set_sheet_pr_xml(self, xml: Optional[str]) -> None:
        """Replace the sheet's ``<sheetPr>`` fragment; ``None`` (or ``""``) removes it."""
        self._set_print_xml("fm_sheet_set_sheet_pr_xml", xml)

    def set_fit_to_page(self, enabled: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set or clear ``<sheetPr><pageSetUpPr fitToPage>``.

        Leaves every other ``<sheetPr>`` child and attribute alone.
        ``fitToPage`` only selects the mode; state the target with
        ``fit_to_width`` / ``fit_to_height`` on :meth:`set_page_setup`.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_set_fit_to_page(h, _uint(sheet, "sheet_index"), 1 if enabled else 0))

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

    def set_print_area(self, ranges: Optional[Sequence[CellRect]]) -> None:
        """Write ``_xlnm.Print_Area`` from one or more rectangles, printed in order.

        A rectangle spanning every row or column is stored as whole columns
        or rows. ``None`` or an empty list removes the name.
        """
        sheet = self._index()
        h = self._workbook._require()
        ranges = [] if ranges is None else ranges
        with _ArgScope() as scope:
            ranges_ptr = _pack_cell_rect_array(ranges, scope)
            _check(
                LIB.fm_sheet_set_print_area(
                    h, _uint(sheet, "sheet_index"), ranges_ptr, _uint(len(ranges), "range_count")
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
        orientation: Optional[PageOrientation] = UNCHANGED,
        paper_size: int = UNCHANGED,
        scale: int = UNCHANGED,
        fit_to_width: int = UNCHANGED,
        fit_to_height: int = UNCHANGED,
        fit_to_page: bool = UNCHANGED,
    ) -> None:
        """Update ``<pageSetup>``; an omitted argument keeps its attribute.

        ``paper_size`` is the OOXML code (9 = A4). ``scale`` outside
        ``[10, 400]`` is rejected rather than clamped -- a mis-stated print
        scale lands on paper, so rounding it silently would hide the error.
        ``fit_to_page`` is routed to ``<sheetPr><pageSetUpPr>``.
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
        left: float = UNCHANGED,
        right: float = UNCHANGED,
        top: float = UNCHANGED,
        bottom: float = UNCHANGED,
        header: float = UNCHANGED,
        footer: float = UNCHANGED,
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
        grid_lines: Optional[bool] = None,
        headings: Optional[bool] = None,
        horizontal_centered: Optional[bool] = None,
        vertical_centered: Optional[bool] = None,
    ) -> None:
        """Apply a partial ``<printOptions>`` update."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PRINT_OPTIONS)
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
            _check(LIB.fm_sheet_set_print_options(h, _uint(sheet, "sheet_index"), ptr))

    def set_header_footer(
        self,
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
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.HEADER_FOOTER)
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
                text_ptr = scope.opt_cstr(value)
                if text_ptr:
                    _, off = S.HEADER_FOOTER.offsets[field]
                    LIB.write_bytes(ptr + off, struct.pack("<I", text_ptr))
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

    def set_column_width(self, first: int, last: int, width: float) -> None:
        """Set the column width override on ``[first, last]``.

        ``width`` is a stored width: characters with the column padding
        included, as ``<col width>`` holds it.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_column_width(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), float(width)
            )
        )

    def set_column_hidden(self, first: int, last: int, hidden: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the column hidden flag on ``[first, last]``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_column_hidden(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), 1 if hidden else 0
            )
        )

    def set_column_outline(self, first: int, last: int, level: int) -> None:
        """Set the column outline level (0-7) on ``[first, last]``.

        Raises:
          FormulonError: with ``StatusCode.INVALID_ARGUMENT`` when ``level`` exceeds 7.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_column_outline(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), _uint(level, "level", 8)
            )
        )

    def set_column_collapsed(self, first: int, last: int, collapsed: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set or clear the ``collapsed`` flag on the columns ``[first, last]``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_column_collapsed(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), 1 if collapsed else 0
            )
        )

    def set_column_style(self, first: int, last: int, xf_index: int) -> None:
        """Set the cell-style XF on ``[first, last]``, keeping the other column attributes."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_column_style(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), _uint(xf_index, "xf_index")
            )
        )

    def clear_column_style(self, first: int, last: int) -> None:
        """Remove the cell-style XF on ``[first, last]``, keeping the other column attributes."""
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

    def set_row_height(self, first: int, last: int, height: float) -> None:
        """Set the row height on ``[first, last]`` as a manual height (``customHeight``)."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_height(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), float(height)
            )
        )

    def set_row_auto_height(self, first: int, last: int, height: float) -> None:
        """Store ``height`` as the cached automatic height of the rows on ``[first, last]``.

        Saved as ``ht`` without ``customHeight``, so each row keeps following its
        content and font; the rows' other attributes are kept.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_auto_height(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), float(height)
            )
        )

    def set_row_hidden(self, first: int, last: int, hidden: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the row hidden flag on ``[first, last]``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_hidden(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), 1 if hidden else 0
            )
        )

    def set_row_outline(self, first: int, last: int, level: int) -> None:
        """Set the row outline level (0-7) on ``[first, last]``.

        Raises:
          FormulonError: with ``StatusCode.INVALID_ARGUMENT`` when ``level`` exceeds 7.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_outline(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), _uint(level, "level", 8)
            )
        )

    def set_row_collapsed(self, first: int, last: int, collapsed: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set or clear the ``collapsed`` flag on ``[first, last]``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_collapsed(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), 1 if collapsed else 0
            )
        )

    def set_row_style(self, first: int, last: int, xf_index: int) -> None:
        """Set the cell-style XF on the inclusive row span ``first..last``, keeping the other row attributes.

        Every row of the span gets an override. The span and ``xf_index`` are
        validated before anything changes.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_row_style(
                h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last"), _uint(xf_index, "xf_index")
            )
        )

    def clear_row_style(self, first: int, last: int) -> None:
        """Remove the cell-style XF on the inclusive row span ``first..last``, keeping the other row attributes."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_row_style(h, _uint(sheet, "sheet_index"), _uint(first, "first"), _uint(last, "last")))

    def get_range_style_bases(self, rect: CellRect) -> List[int]:
        """List, ascending, every cell XF some cell of the selection resolves to.

        ``rect`` is a selection: full height is a column selection, full
        width a row selection, both the whole sheet, neither a cell range. It
        is the first step of an attribute patch (for example a font size over a selection): read the
        bases, build a patched XF for each with :meth:`Workbook.add_font` /
        :meth:`Workbook.add_cell_xf`, then apply them with :meth:`remap_range_style`.
        """
        return self.get_range_style_bases_excluding(rect, ())

    def get_range_style_bases_excluding(
        self,
        rect: CellRect,
        excluded: Sequence[Tuple[int, int]],
    ) -> List[int]:
        """:meth:`get_range_style_bases` over the selection less the ``(row, col)`` cells in ``excluded``.

        An excluded cell contributes no XF, neither its own nor the row, column
        or default style beneath it. Cells outside the selection and repeats are
        ignored; excluding every cell yields an empty list.
        """
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

    def remap_range_style(self, rect: CellRect, mapping: Dict[int, int]) -> None:
        """Remap the XFs of a selection: a cell resolving to key ``k`` ends at ``mapping[k]``.

        An XF absent from ``mapping`` is kept. Each target must be a valid cell
        XF, else :class:`FormulonError` and nothing changes. Patch an attribute
        by reading :meth:`get_range_style_bases`, adding a patched XF for each base
        and passing ``{base: patched}``. Undo with :meth:`snapshot_range_style`
        taken beforehand and :meth:`restore_range_style`.
        """
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

    def snapshot_range_style(self, rect: CellRect) -> bytes:
        """Capture the style state of a selection as opaque bytes for :meth:`restore_range_style`.

        The capture holds the covered column styles (column selection), row
        styles (row selection) and every stored cell's XF and explicit mark.
        Take it before :meth:`remap_range_style` to be able to undo the edit.
        """
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
    def _decode_cf_color(d: Dict[str, int], prefix: str) -> Tuple[int, int, int, int]:
        return (
            d[prefix + "r"],
            d[prefix + "g"],
            d[prefix + "b"],
            d[prefix + "a"],
        )

    def evaluate_cf_range(
        self,
        rect: CellRect,
        today_serial: float = float("nan"),
    ) -> List[CfCell]:
        """Evaluate every CF block on the sheet against the cells of ``rect``.

        Args:
          sheet: 0-based sheet index.
          rect: the cells to evaluate.
          today_serial: date basis pinned for ``timePeriod`` rules, as an
            Excel serial date. The default ``float("nan")`` disables them:
            no ``timePeriod`` rule can match, so pass a serial date to use
            them. ``0.0`` is *not* a disabling value -- it is the valid
            serial for 1899-12-30 and would evaluate those rules against
            that date.
        """
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
        d = S.CF_COLOR.unpack(LIB, ptr)
        return CfColor(d["r"], d["g"], d["b"], d["a"])

    @staticmethod
    def _write_cfvo(ptr: int, value: Cfvo, scope: _ArgScope) -> None:
        S.CFVO.pack(LIB, ptr, {"type": _uint(value.type, "type", 8), "gte": 1 if value.gte else 0})
        if value.value is not None:
            S.write_str_field(LIB, ptr, S.CFVO, "value", value.value, scope)

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
                colors = scope.alloc(count * S.CF_COLOR.size)
                for i, value in enumerate(rule.color_scale.thresholds):
                    self._write_cfvo(thresholds + i * S.CFVO.size, value, scope)
                for i, color in enumerate(rule.color_scale.colors):
                    self._write_cf_color(colors + i * S.CF_COLOR.size, color)
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

    def remove_cf_rule(self, cf_rule_index: int) -> None:
        """Remove the CF rule at ``cf_rule_index`` on the sheet (flattened order)."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_cf_rule(h, _uint(sheet, "sheet_index"), _uint(cf_rule_index, "idx")))

    def clear_cf_rules(self) -> None:
        """Drop every CF block on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_cf_rules(h, _uint(sheet, "sheet_index")))

    # -- Styles ------------------------------------------------------------
    def get_cell_xf_index(self, row: int, col: int) -> int:
        """Return the xf (style record) index attached to a cell."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(LIB.fm_sheet_get_cell_xf_index(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), out))
            return LIB.read_u32(out)

    def set_cell_xf_index(self, row: int, col: int, xf_index: int) -> None:
        """Store ``xf_index`` on the cell at ``(row, col)``.

        Raises:
          FormulonError: with ``StatusCode.INVALID_ARGUMENT`` when ``xf_index`` is not
            below the workbook's cell-XF count.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_cell_xf_index(
                h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col"), _uint(xf_index, "xf_index")
            )
        )

    def clear_cell_xf_index(self, row: int, col: int) -> None:
        """Remove the direct style assignment from the cell at ``(row, col)``.

        A blank cell inherits the row / column style again; value, formula and
        comments are kept. An absent cell is a no-op. A non-blank cell keeps
        owning the default xf. Use the ``source`` of :meth:`get_effective_style`
        to tell an unassigned cell from one directly assigned the Normal xf.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_cell_xf_index(h, _uint(sheet, "sheet"), _uint(row, "row"), _uint(col, "col")))

    def set_range_xf_index(self, rect: CellRect, xf_index: int) -> None:
        """Store ``xf_index`` on every cell of ``rect``.

        Cells that do not exist yet are materialised as blanks carrying only
        the style, so an empty ruled box renders. Ruling a report this way
        costs one call instead of one per cell.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            _check(LIB.fm_sheet_set_range_xf_index(h, _uint(sheet, "sheet"), rp, _uint(xf_index, "xf_index")))

    def get_effective_style(self, row: int, col: int) -> EffectiveStyle:
        """Read the effective style of the cell at ``(row, col)``; the cell need not exist."""
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

    # -- Pivot layout projection -------------------------------------------
    def get_pivot_count(self) -> int:
        """Return the number of PivotTables anchored on the sheet."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(
            LIB.fm_sheet_get_pivot_count,
            h,
            _uint(sheet, "sheet_index"),
        )

    def get_pivot_layout(self, pivot_index: int) -> PivotLayout:
        """Evaluate and project a PivotTable into concrete grid cells."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            handle = 0
            try:
                _check(
                    LIB.fm_sheet_get_pivot_layout(
                        h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), out
                    )
                )
                handle = LIB.read_u32(out)
                present_p = scope.out_ptr()
                rect_p = scope.struct(S.CELL_RECT)
                _check(LIB.fm_pivot_cells_get_bounds(handle, present_p, rect_p))
                rect = _decode_cell_rect_at(rect_p) if LIB.read_u32(present_p) else None
                cells: List[PivotCell] = []
                n = _read_count(LIB.fm_pivot_cells_get_cell_count, handle)
                # One scratch block for the whole walk; see list_defined_names.
                cptr = scope.struct(S.PIVOT_CELL)
                for i in range(n):
                    S.zero_struct(LIB, S.PIVOT_CELL, cptr)
                    _check(LIB.fm_pivot_cells_get_cell_at(handle, _uint(i, "idx"), cptr))
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
                return PivotLayout(rect=rect, cells=cells)
            finally:
                if handle:
                    LIB.fm_pivot_cells_destroy(handle)

    # -- Pivot tables ------------------------------------------------------
    def add_pivot(self, name: str, cache_id: int, anchor_row: int, anchor_col: int) -> int:
        """Create a new empty pivot table; return its flat index."""
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
            return LIB.read_u32(out)

    def get_pivot_report_layout(self, pivot_index: int) -> PivotReportLayout:
        """Return the pivot's compact, tabular, or outline report layout."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_get_pivot_report_layout(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), out
                )
            )
            return PivotReportLayout(LIB.read_i32(out))

    def set_pivot_report_layout(self, pivot_index: int, layout: PivotReportLayout) -> None:
        """Set the pivot's compact, tabular, or outline report layout."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_pivot_report_layout(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _sint(layout, "layout")
            )
        )

    def remove_pivot(self, pivot_index: int) -> None:
        """Remove the pivot table at ``pivot_index``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_remove_pivot(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")))

    def set_pivot_name(self, pivot_index: int, name: str) -> None:
        """Rename the pivot table."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            _check(
                LIB.fm_sheet_set_pivot_name(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), name_ptr)
            )

    def set_pivot_anchor(
        self,
        pivot_index: int,
        rect: CellRect,
    ) -> None:
        """Set the rectangle the pivot reserves (``<location ref>``); its top-left cell is the anchor."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rp = scope.struct(S.CELL_RECT)
            _pack_cell_rect_at(rp, rect)
            _check(LIB.fm_sheet_set_pivot_anchor(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), rp))

    def set_pivot_grand_totals(
        self,
        pivot_index: int,
        *,
        rows: bool = UNCHANGED,
        columns: bool = UNCHANGED,
    ) -> None:
        """Toggle the row / column grand-total bands; an omitted argument keeps its band."""
        sheet = self._index()
        if rows is UNCHANGED or columns is UNCHANGED:
            current = self.get_pivot_grand_totals(pivot_index)
            rows = current.rows if rows is UNCHANGED else rows
            columns = current.columns if columns is UNCHANGED else columns
        for name, value in (("rows", rows), ("columns", columns)):
            if not isinstance(value, bool):
                raise TypeError(f"formulon: {name} must be a bool, got {type(value).__name__}")
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_pivot_grand_totals(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                1 if rows else 0,
                1 if columns else 0,
            )
        )

    def set_pivot_style_info(self, pivot_index: int, style: PivotStyleInput) -> None:
        """Set the pivot's table style and toggles, making the style element present.

        A ``None`` or empty ``style.name`` writes the element without a style name.
        """
        sheet = self._index()
        h = self._workbook._require()
        show_row_headers = style.show_row_headers
        show_col_headers = style.show_col_headers
        show_row_stripes = style.show_row_stripes
        show_col_stripes = style.show_col_stripes
        show_last_column = style.show_last_column
        with _ArgScope() as scope:
            name_ptr = scope.nonempty_cstr(style.name)
            _check(
                LIB.fm_sheet_set_pivot_style_info(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    name_ptr,
                    1 if show_row_headers else 0,
                    1 if show_col_headers else 0,
                    1 if show_row_stripes else 0,
                    1 if show_col_stripes else 0,
                    1 if show_last_column else 0,
                )
            )

    def get_pivot_cache_id(self, pivot_index: int) -> int:
        """Return the ``cache_id`` of the pivot cache the pivot draws from."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(
            LIB.fm_sheet_get_pivot_cache_id, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )

    def get_pivot_name(self, pivot_index: int) -> str:
        """Return the pivot's name."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_cstr_out(
            LIB.fm_sheet_get_pivot_name, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )

    def get_pivot_grand_totals(self, pivot_index: int) -> PivotGrandTotals:
        """Return the row and column flags :meth:`set_pivot_grand_totals` writes."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            rows = scope.out_ptr()
            cols = scope.out_ptr()
            _check(
                LIB.fm_sheet_get_pivot_grand_totals(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), rows, cols
                )
            )
            return PivotGrandTotals(rows=LIB.read_i32(rows) != 0, columns=LIB.read_i32(cols) != 0)

    def get_pivot_style_info(self, pivot_index: int) -> Optional[PivotStyle]:
        """Return the table style :meth:`set_pivot_style_info` writes, or ``None`` without a style element."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_STYLE_INFO)
            _check(
                LIB.fm_sheet_get_pivot_style_info(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr
                )
            )
            d = S.PIVOT_STYLE_INFO.unpack(LIB, ptr)
            if not d["present"]:
                return None
            return PivotStyle(
                name=LIB.read_cstr(d["name"]) or None,
                show_row_headers=d["show_row_headers"] != 0,
                show_col_headers=d["show_col_headers"] != 0,
                show_row_stripes=d["show_row_stripes"] != 0,
                show_col_stripes=d["show_col_stripes"] != 0,
                show_last_column=d["show_last_column"] != 0,
            )

    def _pivot_get_field_order(self, fn, sheet: int, pivot_index: int) -> List[int]:
        h = self._workbook._require()
        with _ArgScope() as scope:
            out_ptr = scope.out_ptr()
            out_count = scope.out_ptr()
            _check(fn(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), out_ptr, out_count))
            count = LIB.read_u32(out_count)
            if count == 0:
                return []
            raw = LIB.read_bytes(LIB.read_u32(out_ptr), 4 * count)
            return [PIVOT_VALUES_FIELD if v == _PIVOT_VALUES_FIELD_C else v for v in struct.unpack("<%dI" % count, raw)]

    def get_pivot_row_field_order(self, pivot_index: int) -> List[int]:
        """Return the row-axis field order (the list :meth:`set_pivot_row_field_order` takes)."""
        sheet = self._index()
        return self._pivot_get_field_order(LIB.fm_sheet_get_pivot_row_field_order, sheet, pivot_index)

    def get_pivot_column_field_order(self, pivot_index: int) -> List[int]:
        """Return the column-axis field order. Same contract as rows."""
        sheet = self._index()
        return self._pivot_get_field_order(LIB.fm_sheet_get_pivot_column_field_order, sheet, pivot_index)

    def _pivot_field_at(self, pivot_index: int, field_index: int) -> PivotField:
        """Read the definition of pivot field ``field_index``."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_FIELD_INFO)
            _check(
                LIB.fm_sheet_get_pivot_field(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    ptr,
                )
            )
            d = S.PIVOT_FIELD_INFO.unpack(LIB, ptr)
            return PivotField(
                source_name=LIB.read_cstr(d["source_name"]),
                custom_name=LIB.read_cstr(d["custom_name"]) or None,
                axis=PivotAxis(d["axis"]) if d["axis"] >= 0 else None,
                subtotal_top=d["subtotal_top"] != 0,
                number_format=LIB.read_cstr(d["number_format"]) or None,
                cache_field_index=None if d["cache_field_index"] < 0 else d["cache_field_index"],
                sort_order=PivotSortOrder(d["sort_order"]),
                sort_by_field=LIB.read_cstr(d["sort_by_field"]) or None,
                date_periods=PivotDatePeriod(d["date_periods"]),
                date_calendar=PivotCalendar(d["date_calendar"]),
                group_base_index=None if d["group_base_index"] < 0 else d["group_base_index"],
                default_subtotal=d["default_subtotal"] != 0,
                subtotal_fns=[agg for agg in PivotAggregation if d["subtotal_fns"] & (1 << int(agg))],
                insert_blank_row=d["insert_blank_row"] != 0,
                fill_down_labels=d["fill_down_labels"] != 0,
                show_all=d["show_all"] != 0,
            )

    def get_pivot_field_item_count(self, pivot_index: int, field_index: int) -> int:
        """Return the number of items in pivot field ``field_index``'s item list."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(
            LIB.fm_sheet_get_pivot_field_item_count,
            h,
            _uint(sheet, "sheet_index"),
            _uint(pivot_index, "pivot_index"),
            _uint(field_index, "field_index"),
        )

    def get_pivot_field_item(self, pivot_index: int, field_index: int, item_index: int) -> PivotItem:
        """Read item ``item_index`` of pivot field ``field_index``."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_ITEM_INFO)
            _check(
                LIB.fm_sheet_get_pivot_field_item(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    _uint(item_index, "item_index"),
                    ptr,
                )
            )
            d = S.PIVOT_ITEM_INFO.unpack(LIB, ptr)
            return PivotItem(
                name=LIB.read_cstr(d["name"]) or None,
                cache_index=None if d["cache_index"] < 0 else d["cache_index"],
                visible=d["visible"] != 0,
                caption=LIB.read_cstr(d["caption"]) if d["caption"] else None,
                show_detail=d["show_detail"] != 0,
            )

    def _pivot_data_field_at(self, pivot_index: int, data_field_index: int) -> PivotDataField:
        """Read data field ``data_field_index`` in the shape ``pivot_data_field_add`` accepts."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_DATA_FIELD_SPEC)
            _check(
                LIB.fm_sheet_get_pivot_data_field(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(data_field_index, "data_field_index"),
                    ptr,
                )
            )
            d = S.PIVOT_DATA_FIELD_SPEC.unpack(LIB, ptr)
            return PivotDataField(
                name=LIB.read_cstr(d["name"]),
                field_index=d["field_index"],
                aggregation=PivotAggregation(d["aggregation"]),
                number_format=LIB.read_cstr(d["number_format"]) or None,
                show_as=PivotShowValuesAs(d["show_as"]),
                show_as_base_field=None if d["show_as_base_field"] < 0 else d["show_as_base_field"],
                show_as_base_item=None if d["show_as_base_item"] < 0 else d["show_as_base_item"],
            )

    def list_pivot_fields(self, pivot_index: int) -> List[PivotField]:
        """Every field of the pivot, in cache-field order."""
        return [self._pivot_field_at(pivot_index, i) for i in range(self._pivot_field_count(pivot_index))]

    def _pivot_field_count(self, pivot_index: int) -> int:
        """Return the number of fields configured on the pivot."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(
            LIB.fm_sheet_get_pivot_field_count, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )

    def add_pivot_field(self, pivot_index: int, spec: PivotFieldInput) -> int:
        """Place a field on the pivot; return its index.

        A pivot's fields are positional over its cache's fields (field ``i``
        is cache field ``i``). When ``source_name`` names a cache field the
        pivot first gets one bare, axis-less field per cache field and the
        named one is configured in place, so the index is that cache field's
        position and adding it again updates it. A name the cache lacks
        appends a field. A field added on the row or column axis is appended
        to that axis's field order.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_FIELD_SPEC)
            out = scope.out_ptr()
            S.PIVOT_FIELD_SPEC.pack(
                LIB,
                ptr,
                {
                    "axis": _sint(spec.axis, "axis"),
                    "subtotal_top": 1 if spec.subtotal_top else 0,
                },
            )
            S.write_str_field(LIB, ptr, S.PIVOT_FIELD_SPEC, "source_name", spec.source_name, scope)
            S.write_str_field(LIB, ptr, S.PIVOT_FIELD_SPEC, "custom_name", spec.custom_name, scope)
            S.write_str_field(
                LIB,
                ptr,
                S.PIVOT_FIELD_SPEC,
                "number_format",
                spec.number_format,
                scope,
            )
            _check(
                LIB.fm_sheet_add_pivot_field(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr, out
                )
            )
            return LIB.read_u32(out)

    def clear_pivot_fields(self, pivot_index: int) -> None:
        """Drop every field from the pivot."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_pivot_fields(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")))

    def set_pivot_field_axis(self, pivot_index: int, field_index: int, axis: Optional[PivotAxis]) -> None:
        """Set the axis of pivot field ``field_index``.

        Moving onto the row or column axis appends the field to that axis's
        order; moving off removes it from the order it was in. ``None`` takes
        the field off every axis: its data fields are removed (later
        data-field indices shift down), as are filters naming it and
        value filters scoring a removed data field.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_pivot_field_axis(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_index, "field_index"),
                _sint(_nullable_enum(axis, "axis"), "axis"),
            )
        )

    def set_pivot_field_sort(
        self,
        pivot_index: int,
        field_index: int,
        order: PivotSortOrder,
        by_field: str = "",
    ) -> None:
        """Set the sibling order of pivot field ``field_index``.

        ``PivotSortOrder.MANUAL`` rejects a non-empty ``by_field``.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            by_ptr = scope.nonempty_cstr(by_field)
            _check(
                LIB.fm_sheet_set_pivot_field_sort(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    _sint(order, "order"),
                    by_ptr,
                )
            )

    def set_pivot_field_subtotal_top(self, pivot_index: int, field_index: int, top: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set the ``subtotal_top`` flag on pivot field ``field_index``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_pivot_field_subtotal_top(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_index, "field_index"),
                1 if top else 0,
            )
        )

    def _pivot_field_set_flag(self, fn, sheet: int, pivot_index: int, field_index: int, on: bool) -> None:  # noqa: FBT001 -- shared body of the flag setters
        h = self._workbook._require()
        _check(
            fn(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_index, "field_index"),
                1 if on else 0,
            )
        )

    def set_pivot_field_default_subtotal(self, pivot_index: int, field_index: int, on: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Turn the automatic subtotal of pivot field ``field_index`` on or off.

        Custom subtotal functions replace the automatic subtotal whatever this flag says.
        """
        sheet = self._index()
        self._pivot_field_set_flag(LIB.fm_sheet_set_pivot_field_default_subtotal, sheet, pivot_index, field_index, on)

    def set_pivot_field_insert_blank_row(self, pivot_index: int, field_index: int, on: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set whether a blank row follows each item block of pivot field ``field_index``."""
        sheet = self._index()
        self._pivot_field_set_flag(LIB.fm_sheet_set_pivot_field_insert_blank_row, sheet, pivot_index, field_index, on)

    def set_pivot_field_fill_down_labels(self, pivot_index: int, field_index: int, on: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set whether pivot field ``field_index`` repeats its item labels on every row of a block."""
        sheet = self._index()
        self._pivot_field_set_flag(LIB.fm_sheet_set_pivot_field_fill_down_labels, sheet, pivot_index, field_index, on)

    def set_pivot_field_show_detail(self, pivot_index: int, field_index: int, on: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Expand (``True``) or collapse (``False``) every item of pivot field ``field_index`` and its subtotal row."""
        sheet = self._index()
        self._pivot_field_set_flag(LIB.fm_sheet_set_pivot_field_show_detail, sheet, pivot_index, field_index, on)

    def set_pivot_field_item_show_detail(
        self,
        pivot_index: int,
        field_index: int,
        item_index: int,
        on: bool,  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
    ) -> None:
        """Expand (``True``) or collapse (``False``) item ``item_index`` of pivot field ``field_index``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_pivot_field_item_show_detail(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_index, "field_index"),
                _uint(item_index, "item_index"),
                1 if on else 0,
            )
        )

    def set_pivot_field_item_caption(
        self, pivot_index: int, field_index: int, item_index: int, caption: Optional[str]
    ) -> None:
        """Show ``caption`` instead of the value of item ``item_index``; ``None`` or ``""`` clears it.

        A caption longer than 255 UTF-16 code units, or one matching the label another item of
        the field shows, is rejected.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            cap_ptr = scope.nonempty_cstr(caption)
            _check(
                LIB.fm_sheet_set_pivot_field_item_caption(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    _uint(item_index, "item_index"),
                    cap_ptr,
                )
            )

    def get_pivot_options(self, pivot_index: int) -> PivotOptions:
        """Read the PivotTable Options of the pivot; see :class:`PivotOptions`."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_OPTIONS)
            _check(
                LIB.fm_sheet_get_pivot_options(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr)
            )
            d = S.PIVOT_OPTIONS.unpack(LIB, ptr)

            return PivotOptions(
                show_error=d["show_error"] != 0,
                error_caption=LIB.read_cstr(d["error_caption"]),
                show_missing=d["show_missing"] != 0,
                missing_caption=LIB.read_cstr(d["missing_caption"]),
                show_headers=d["show_headers"] != 0,
                show_drill=d["show_drill"] != 0,
                show_data_tips=d["show_data_tips"] != 0,
                grid_drop_zones=d["grid_drop_zones"] != 0,
                show_empty_row=d["show_empty_row"] != 0,
                show_empty_col=d["show_empty_col"] != 0,
                show_data_drop_down=d["show_data_drop_down"] != 0,
                print_drill=d["print_drill"] != 0,
                field_print_titles=d["field_print_titles"] != 0,
                item_print_titles=d["item_print_titles"] != 0,
                merge_item=d["merge_item"] != 0,
                indent=d["indent"],
                page_over_then_down=d["page_over_then_down"] != 0,
                page_wrap=d["page_wrap"],
                multiple_field_filters=d["multiple_field_filters"] != 0,
                preserve_formatting=d["preserve_formatting"] != 0,
                custom_list_sort=d["custom_list_sort"] != 0,
                alt_text=LIB.read_cstr(d["alt_text"]),
                alt_text_summary=LIB.read_cstr(d["alt_text_summary"]),
                grand_total_caption=LIB.read_cstr(d["grand_total_caption"]),
                auto_formatting=d["auto_formatting"] != 0,
                enable_drill=d["enable_drill"] != 0,
            )

    def set_pivot_options(self, pivot_index: int, options: PivotOptions) -> None:
        """Change the PivotTable Options named by ``options``; members left ``None`` keep their value.

        See :class:`PivotOptions`.
        """
        sheet = self._index()
        h = self._workbook._require()
        if not isinstance(options, PivotOptions):
            raise TypeError(f"expected PivotOptions, got {type(options).__name__}")
        current = self.get_pivot_options(pivot_index)

        def pick(name: str) -> Any:
            value = getattr(options, name)
            return getattr(current, name) if value is None else value

        flags = (
            "show_error",
            "show_missing",
            "show_headers",
            "show_drill",
            "show_data_tips",
            "grid_drop_zones",
            "show_empty_row",
            "show_empty_col",
            "show_data_drop_down",
            "print_drill",
            "field_print_titles",
            "item_print_titles",
            "merge_item",
            "page_over_then_down",
            "multiple_field_filters",
            "preserve_formatting",
            "custom_list_sort",
            "auto_formatting",
            "enable_drill",
        )
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_OPTIONS)
            values: Dict[str, object] = {name: 1 if pick(name) else 0 for name in flags}
            values["indent"] = _uint(pick("indent"), "indent")
            values["page_wrap"] = _uint(pick("page_wrap"), "page_wrap")
            S.PIVOT_OPTIONS.pack(LIB, ptr, values)
            for name in ("error_caption", "missing_caption", "alt_text", "alt_text_summary", "grand_total_caption"):
                S.write_str_field(LIB, ptr, S.PIVOT_OPTIONS, name, pick(name), scope)
            _check(
                LIB.fm_sheet_set_pivot_options(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr)
            )

    def project_pivot(self, pivot_index: int) -> None:
        """Write the pivot's report into its sheet cells now and clear its pending flag.

        Raises when the report would overlap another structure; the pivot stays pending.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_project_pivot(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")))

    def is_pivot_projection_pending(self, pivot_index: int) -> bool:
        """Return whether a cache change left the pivot's sheet cells stale."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_is_pivot_projection_pending(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), out
                )
            )
            return LIB.read_i32(out) != 0

    def set_pivot_field_custom_name(self, pivot_index: int, field_index: int, name: Optional[str]) -> None:
        """Rename pivot field ``field_index`` in the report; ``None`` or ``""`` restores its source name."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            name_ptr = scope.nonempty_cstr(name)
            _check(
                LIB.fm_sheet_set_pivot_field_custom_name(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    name_ptr,
                )
            )

    def set_pivot_field_show_all(self, pivot_index: int, field_index: int, on: bool) -> None:  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
        """Set whether pivot field ``field_index`` lists items without data."""
        sheet = self._index()
        self._pivot_field_set_flag(LIB.fm_sheet_set_pivot_field_show_all, sheet, pivot_index, field_index, on)

    def get_pivot_data_field_default_name(
        self,
        pivot_index: int,
        field_index: int,
        aggregation: PivotAggregation,
        exclude_data_field: Optional[int] = None,
    ) -> str:
        """Return the name a data field over ``field_index`` would get from an empty name.

        Nothing changes. ``exclude_data_field`` is the entry :meth:`set_pivot_data_field`
        would replace, whose name is not counted as taken; ``None`` for the add case.
        """
        sheet = self._index()
        h = self._workbook._require()
        exclude = (
            _PIVOT_NO_DATA_FIELD if exclude_data_field is None else _uint(exclude_data_field, "exclude_data_field")
        )
        return _read_cstr_out(
            LIB.fm_sheet_get_pivot_data_field_default_name,
            h,
            _uint(sheet, "sheet_index"),
            _uint(pivot_index, "pivot_index"),
            _uint(field_index, "field_index"),
            _sint(aggregation, "aggregation"),
            exclude,
        )

    def show_pivot_details(self, pivot_index: int, row: int, col: int) -> Sheet:
        """Write the source records behind pivot cell (``row``, ``col``) onto a new sheet and return its handle.

        The sheet is inserted immediately before the pivot's sheet and becomes
        the active sheet; every existing :class:`Sheet` handle keeps pointing
        at its sheet.
        """
        sheet = self._index()
        h = self._workbook._require()
        handles = self._workbook._sheet_handles()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_show_pivot_details(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(row, "row"),
                    _uint(col, "col"),
                    out,
                )
            )
            index = LIB.read_u32(out)
        return self._workbook._insert_sheet_handle(handles, index)

    def add_pivot_field_item(self, pivot_index: int, field_index: int, name: str, *, visible: bool) -> int:
        """Append a manual-filter item addressed by its label.

        The item carries no cache binding, so records are matched by
        comparing their rendered label against ``name``. The blank item has
        no label of its own; build it with
        :meth:`add_pivot_field_shared_item` instead.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            name_ptr = scope.cstr(name)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_add_pivot_field_item(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    name_ptr,
                    1 if visible else 0,
                    out,
                )
            )
            return LIB.read_u32(out)

    def add_pivot_field_shared_item(
        self, pivot_index: int, field_index: int, cache_index: int, *, visible: bool
    ) -> int:
        """Append a manual-filter item addressed by its cache shared-item index.

        ``cache_index`` selects ``shared_items[cache_index]`` of the bound
        cache field, the same index space as OOXML ``<item x="N">``. The
        label is resolved from that shared item, so a shared item that
        renders to nothing produces the blank item.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_add_pivot_field_shared_item(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    _uint(cache_index, "cache_index"),
                    1 if visible else 0,
                    out,
                )
            )
            return LIB.read_u32(out)

    def clear_pivot_field_items(self, pivot_index: int, field_index: int) -> None:
        """Drop every manual-filter item from pivot field ``field_index``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_clear_pivot_field_items(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(field_index, "field_index")
            )
        )

    def set_pivot_field_item_visible(
        self,
        pivot_index: int,
        field_index: int,
        item_index: int,
        visible: bool,  # noqa: FBT001 -- a single-boolean setter takes its flag positionally
    ) -> None:
        """Toggle the visibility of item ``item_index`` on field ``field_index``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_pivot_field_item_visible(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_index, "field_index"),
                _uint(item_index, "item_index"),
                1 if visible else 0,
            )
        )

    def add_pivot_field_subtotal_fn(self, pivot_index: int, field_index: int, agg: PivotAggregation) -> None:
        """Append a subtotal-fn entry to pivot field ``field_index``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_add_pivot_field_subtotal_fn(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_index, "field_index"),
                _sint(agg, "agg"),
            )
        )

    def clear_pivot_field_subtotal_fns(self, pivot_index: int, field_index: int) -> None:
        """Drop every subtotal-fn entry from pivot field ``field_index``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_clear_pivot_field_subtotal_fns(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(field_index, "field_index")
            )
        )

    def set_pivot_field_date_group(
        self,
        pivot_index: int,
        field_index: int,
        periods: PivotDatePeriod,
        calendar: PivotCalendar = PivotCalendar.GREGORIAN,
        interval_days: int = 1,
        start: float = float("nan"),
        end: float = float("nan"),
    ) -> None:
        """Group the date field ``field_index`` by the periods of ``periods``.

        ``field_index`` is a source field or one of its group fields; any grouping
        it has is replaced. Each period becomes a group field appended to the
        pivot cache, finest first, named and labelled in the workbook's locale;
        this pivot shows them in the field's place, coarsest outermost, and
        another pivot showing the field shows them before it.
        ``PivotCalendar.JAPANESE`` groups the years by era year.
        ``interval_days`` above 1 groups by runs of days and needs ``periods``
        to be ``PivotDatePeriod.DAYS`` alone. ``start`` / ``end`` are date
        serials; NaN takes the field's first value and its last (the midnight
        after it for a date without a time). An empty or unknown mask, a zero
        interval, an invalid bound, a start after the end, or a field holding
        anything but dates raises ``FormulonError``.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_set_pivot_field_date_group(
                h,
                _uint(sheet, "sheet_index"),
                _uint(pivot_index, "pivot_index"),
                _uint(field_index, "field_index"),
                _uint(periods, "periods_mask"),
                _sint(calendar, "calendar"),
                _uint(interval_days, "interval_days"),
                float(start),
                float(end),
            )
        )

    def clear_pivot_field_date_group(self, pivot_index: int, field_index: int) -> None:
        """Remove the date grouping of pivot field ``field_index``.

        ``field_index`` is a source field or one of its group fields. The group
        fields leave every pivot on the cache and the field returns to the axis
        position they held; a field that is not grouped is left as it is.
        """
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_clear_pivot_field_date_group(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(field_index, "field_index")
            )
        )

    def set_pivot_field_number_format(self, pivot_index: int, field_index: int, fmt: str) -> None:
        """Set the ``numFmtId`` of pivot field ``field_index``.

        ``fmt`` is a decimal id string naming a built-in format or one returned
        by :meth:`Workbook.add_num_fmt`; the empty string clears it.
        """
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            fmt_ptr = scope.cstr(fmt)
            _check(
                LIB.fm_sheet_set_pivot_field_number_format(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(field_index, "field_index"),
                    fmt_ptr,
                )
            )

    def set_pivot_row_field_order(self, pivot_index: int, indices: Sequence[int]) -> None:
        """Replace the row-axis field order with ``indices``.

        The list is the axis's whole membership: listed fields move onto the
        row axis, and fields left out come off it. Each index may appear once.
        ``PIVOT_VALUES_FIELD`` (``-2``) places the Σ Values level.
        """
        sheet = self._index()
        self._pivot_set_field_order(
            LIB.fm_sheet_set_pivot_row_field_order,
            _uint(sheet, "sheet_index"),
            _uint(pivot_index, "pivot_index"),
            indices,
        )

    def set_pivot_column_field_order(self, pivot_index: int, indices: Sequence[int]) -> None:
        """Replace the column-axis field order with ``indices``.

        Same contract as :meth:`set_pivot_row_field_order`.
        """
        sheet = self._index()
        self._pivot_set_field_order(
            LIB.fm_sheet_set_pivot_column_field_order,
            _uint(sheet, "sheet_index"),
            _uint(pivot_index, "pivot_index"),
            indices,
        )

    def _pivot_set_field_order(self, fn, sheet: int, pivot_index: int, indices: Sequence[int]) -> None:
        h = self._workbook._require()
        arr_ptr = 0
        with _ArgScope() as scope:
            if indices:
                arr_ptr = scope.alloc(4 * len(indices))
                buf = b"".join(
                    struct.pack("<I", _PIVOT_VALUES_FIELD_C if i == PIVOT_VALUES_FIELD else _uint(i, "index"))
                    for i in indices
                )
                LIB.write_bytes(arr_ptr, buf)
            _check(fn(h, sheet, pivot_index, arr_ptr, _uint(len(indices), "count")))

    def list_pivot_data_fields(self, pivot_index: int) -> List[PivotDataField]:
        """Every data field of the pivot, in order."""
        return [self._pivot_data_field_at(pivot_index, i) for i in range(self._pivot_data_field_count(pivot_index))]

    def _pivot_data_field_count(self, pivot_index: int) -> int:
        """Return the number of ``<dataField>`` entries on the pivot."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(
            LIB.fm_sheet_get_pivot_data_field_count, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )

    def add_pivot_data_field(self, pivot_index: int, spec: PivotDataFieldInput) -> int:
        """Append a data-field entry; return its index."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            out = scope.out_ptr()
            ptr = self._pack_data_field_spec(spec, scope)
            _check(
                LIB.fm_sheet_add_pivot_data_field(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr, out
                )
            )
            return LIB.read_u32(out)

    def set_pivot_data_field(
        self,
        pivot_index: int,
        data_field_index: int,
        spec: PivotDataFieldInput,
    ) -> None:
        """Replace the data-field entry at ``data_field_index`` in place."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = self._pack_data_field_spec(spec, scope)
            _check(
                LIB.fm_sheet_set_pivot_data_field(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(data_field_index, "data_field_index"),
                    ptr,
                )
            )

    @staticmethod
    def _pack_data_field_spec(spec: PivotDataFieldInput, scope: _ArgScope) -> int:
        ptr = scope.struct(S.PIVOT_DATA_FIELD_SPEC)
        S.PIVOT_DATA_FIELD_SPEC.pack(
            LIB,
            ptr,
            {
                "field_index": _uint(spec.field_index, "field_index"),
                "aggregation": _sint(spec.aggregation, "aggregation"),
                "show_as": _sint(spec.show_as, "show_as"),
                "show_as_base_field": _sint(
                    -1 if spec.show_as_base_field is None else spec.show_as_base_field, "show_as_base_field"
                ),
                "show_as_base_item": _sint(
                    -1 if spec.show_as_base_item is None else spec.show_as_base_item, "show_as_base_item"
                ),
            },
        )
        S.write_str_field(LIB, ptr, S.PIVOT_DATA_FIELD_SPEC, "name", spec.name, scope)
        S.write_str_field(
            LIB,
            ptr,
            S.PIVOT_DATA_FIELD_SPEC,
            "number_format",
            spec.number_format,
            scope,
        )
        return ptr

    def clear_pivot_data_fields(self, pivot_index: int) -> None:
        """Drop every data-field entry from the pivot."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_pivot_data_fields(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")))

    def list_pivot_filters(self, pivot_index: int) -> List[PivotFilter]:
        """Every listed filter of the pivot, in order.

        The filter list is saved as the definition's ``<filters>`` block. A
        loaded file's caption contains / begins-with, top-N count, value
        greater-than / between and date-between entries appear here, naming
        their field by source name. Its other entries are not listed, still
        filter, survive :meth:`clear_pivot_filters` /
        :meth:`remove_pivot_filter`, and are saved verbatim.
        """
        return [self._pivot_filter_at(pivot_index, i) for i in range(self._pivot_filter_count(pivot_index))]

    def _pivot_filter_count(self, pivot_index: int) -> int:
        """Return the number of filters on the pivot; see :meth:`list_pivot_filters`."""
        sheet = self._index()
        h = self._workbook._require()
        return _read_count(
            LIB.fm_sheet_get_pivot_filter_count, h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")
        )

    def _pivot_filter_at(self, pivot_index: int, filter_index: int) -> PivotFilter:
        """Read the filter at ``filter_index`` without mutating it."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_FILTER_SPEC)
            _check(
                LIB.fm_sheet_get_pivot_filter(
                    h,
                    _uint(sheet, "sheet_index"),
                    _uint(pivot_index, "pivot_index"),
                    _uint(filter_index, "filter_index"),
                    ptr,
                )
            )
            d = S.PIVOT_FILTER_SPEC.unpack(LIB, ptr)
            return PivotFilter(
                axis=PivotAxis(d["axis"]),
                field_name=LIB.read_cstr(d["field_name"]),
                type=PivotFilterType(d["type"]),
                value_kind=_enum_or_none(PivotFilterValueKind, d["value_kind"]),
                value_int=d["value_int"],
                value_double=d["value_double"],
                value_text=LIB.read_cstr(d["value_text"]),
                value_high_kind=_enum_or_none(PivotFilterValueKind, d["value_high_kind"]),
                value_high_int=d["value_high_int"],
                value_high_double=d["value_high_double"],
                data_field_index=d["data_field_index"],
                value_high_text=LIB.read_cstr(d["value_high_text"]),
            )

    def add_pivot_filter(self, pivot_index: int, spec: PivotFilterInput) -> int:
        """Append a filter to the pivot and return its index; it is saved with it. See :meth:`list_pivot_filters`."""
        sheet = self._index()
        h = self._workbook._require()
        with _ArgScope() as scope:
            ptr = scope.struct(S.PIVOT_FILTER_SPEC)
            S.PIVOT_FILTER_SPEC.pack(
                LIB,
                ptr,
                {
                    "axis": _sint(spec.axis, "axis"),
                    "type": _sint(spec.type, "type"),
                    "data_field_index": _uint(spec.data_field_index, "data_field_index"),
                    "value_kind": _sint(_nullable_enum(spec.value_kind, "value_kind"), "value_kind"),
                    "value_int": _sint(spec.value_int, "value_int"),
                    "value_double": float(spec.value_double),
                    "value_high_kind": _sint(
                        _nullable_enum(spec.value_high_kind, "value_high_kind"), "value_high_kind"
                    ),
                    "value_high_int": _sint(spec.value_high_int, "value_high_int"),
                    "value_high_double": float(spec.value_high_double),
                },
            )
            S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC, "field_name", spec.field_name, scope)
            S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC, "value_text", spec.value_text, scope)
            S.write_str_field(LIB, ptr, S.PIVOT_FILTER_SPEC, "value_high_text", spec.value_high_text, scope)
            out = scope.out_ptr()
            _check(
                LIB.fm_sheet_add_pivot_filter(
                    h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), ptr, out
                )
            )
            return LIB.read_u32(out)

    def clear_pivot_filters(self, pivot_index: int) -> None:
        """Drop every listed filter from the pivot."""
        sheet = self._index()
        h = self._workbook._require()
        _check(LIB.fm_sheet_clear_pivot_filters(h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index")))

    def remove_pivot_filter(self, pivot_index: int, filter_index: int) -> None:
        """Remove the filter at ``filter_index``."""
        sheet = self._index()
        h = self._workbook._require()
        _check(
            LIB.fm_sheet_remove_pivot_filter(
                h, _uint(sheet, "sheet_index"), _uint(pivot_index, "pivot_index"), _uint(filter_index, "filter_index")
            )
        )

    # -- Dependency-graph trace --------------------------------------------
    def get_precedents(self, row: int, col: int, depth: int = 1) -> List[CellNode]:
        """Return the cells ``(row, col)`` reads (up to ``depth``)."""
        sheet = self._index()
        return self._trace(
            LIB.fm_sheet_get_precedents,
            _uint(sheet, "sheet"),
            _uint(row, "row"),
            _uint(col, "col"),
            _uint(depth, "depth"),
        )

    def get_dependents(self, row: int, col: int, depth: int = 1) -> List[CellNode]:
        """Return the cells that read ``(row, col)`` (up to ``depth``)."""
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
    def get_spill_info(self, row: int, col: int) -> Spill:
        """Return dynamic-array spill info for ``(row, col)``."""
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
    def get_array_formula(self, row: int, col: int) -> Optional[ArrayFormula]:
        """Return the legacy array formula covering ``(row, col)``, or ``None``.

        Every cell of the block reports the same :class:`ArrayFormula`.
        """
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


_install_workbook_locks(Sheet)
