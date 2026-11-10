//
// Compile-only C11 contract probe for the public C ABI header.

#include "c_api/formulon_c.h"

// The header declares no `bool` / `_Bool` anywhere and does not include
// <stdbool.h>, so a C11 consumer that never includes it must still compile.
// Booleans cross the ABI as `int32_t` (0 = false), including this callback's
// return type.
static int32_t ContinueIteration(uint32_t iteration, double max_residual, uint32_t max_iterations, void* user_data) {
  (void)iteration;
  (void)max_residual;
  (void)max_iterations;
  (void)user_data;
  return 1;
}

// A status code the API never returns is not declared.
#ifdef FM_STATUS_CANNOT_REMOVE_LAST_SHEET
#error "FM_STATUS_CANNOT_REMOVE_LAST_SHEET is no longer part of the C ABI"
#endif

// A custom cell style says so through `has_builtin_id`, not a sentinel id.
#ifdef FM_CELL_STYLE_BUILTIN_ID_NONE
#error "FM_CELL_STYLE_BUILTIN_ID_NONE is no longer part of the C ABI"
#endif

// An optional value inside a record is announced by an `int32_t has_<field>`
// flag; an optional index sits right after its flag as a `uint32_t`.
#define EXPECT_HAS_FLAG(T, field) \
  _Static_assert(sizeof(((T*)0)->field) == sizeof(int32_t), #T "." #field " is an int32_t flag")
#define EXPECT_FLAGGED_INDEX(T, field)                                                                                \
  _Static_assert(                                                                                                     \
      offsetof(T, field) == offsetof(T, has_##field) + sizeof(int32_t) && sizeof(((T*)0)->field) == sizeof(uint32_t), \
      #T "." #field " follows has_" #field)

EXPECT_HAS_FLAG(fm_cf_match_info_t, has_dxf_index);
EXPECT_HAS_FLAG(fm_page_setup_t, has_orientation);
EXPECT_HAS_FLAG(fm_page_margins_t, has_left_in);
EXPECT_HAS_FLAG(fm_print_options_t, has_grid_lines);
EXPECT_HAS_FLAG(fm_header_footer_t, has_different_first);
EXPECT_HAS_FLAG(fm_cf_rule_t, has_op);
EXPECT_HAS_FLAG(fm_cf_rule_t, has_icon_set_floor);
EXPECT_HAS_FLAG(fm_dxf_t, has_font);
EXPECT_HAS_FLAG(fm_chart_spec_t, has_title);
EXPECT_FLAGGED_INDEX(fm_function_metadata_info_t, max_arity);
EXPECT_FLAGGED_INDEX(fm_filter_column_t, dxf_index);
EXPECT_FLAGGED_INDEX(fm_pivot_cache_options_t, missing_items_limit);
EXPECT_FLAGGED_INDEX(fm_pivot_field_info_t, cache_field_index);
EXPECT_FLAGGED_INDEX(fm_pivot_field_info_t, group_base_index);
EXPECT_FLAGGED_INDEX(fm_pivot_item_info_t, cache_index);
EXPECT_FLAGGED_INDEX(fm_pivot_data_field_t, show_as_base_field);
EXPECT_FLAGGED_INDEX(fm_pivot_data_field_t, show_as_base_item);

// Absent enumerations read their `_ABSENT` member; the other sentinels are
// named constants.
_Static_assert(FM_PIVOT_AXIS_ABSENT == -1, "FM_PIVOT_AXIS_ABSENT changed");
_Static_assert(FM_CHART_EXTENDED_LAYOUT_ABSENT == -1, "FM_CHART_EXTENDED_LAYOUT_ABSENT changed");
_Static_assert(FM_DEFINED_NAME_SCOPE_WORKBOOK == -1, "FM_DEFINED_NAME_SCOPE_WORKBOOK changed");
_Static_assert(FM_PIVOT_DATA_FIELD_NONE == SIZE_MAX, "FM_PIVOT_DATA_FIELD_NONE changed");
_Static_assert(FM_PIVOT_SHOW_AS_BASE_PREVIOUS == 1048828u, "FM_PIVOT_SHOW_AS_BASE_PREVIOUS changed");
_Static_assert(FM_PIVOT_SHOW_AS_BASE_NEXT == 1048829u, "FM_PIVOT_SHOW_AS_BASE_NEXT changed");

int main(void) {
  fm_iterative_progress_cb callback = ContinueIteration;
  fm_parallel_recalc_stats_info_t stats = FM_STRUCT_INIT(fm_parallel_recalc_stats_info_t);
  // A record initialised from C: `struct_size` leads, every other field is
  // zero, and the version macro links against the run-time query.
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  _Static_assert(offsetof(fm_cf_rule_t, struct_size) == 0, "struct_size must lead a record");
  int version_matches = FM_CHECK_VERSION();
  fm_status_t (*parallel_recalc)(fm_workbook_t*, uint32_t, fm_parallel_recalc_stats_info_t*) =
      fm_workbook_recalc_parallel;
  fm_status_t (*save_diagnostics)(const fm_workbook_t*, int32_t, uint8_t**, size_t*, fm_save_diagnostics_info_t*) =
      fm_workbook_save_with_diagnostics;
  fm_status_t (*read_diagnostics)(const fm_workbook_t*, fm_read_diagnostics_info_t*) = fm_workbook_get_diagnostics;
  fm_status_t (*save)(const fm_workbook_t*, int32_t, uint8_t**, size_t*) = fm_workbook_save;
  _Static_assert(FM_WORKBOOK_FORMAT_XLSX == 0, "a zeroed save format is xlsx");
  // Insert / Delete Cells take the rectangle by const pointer and the shift as
  // a raw ordinal; zero is Down, valid for an insert only.
  fm_status_t (*insert_cells)(fm_workbook_t*, size_t, const fm_cell_rect_t*, int32_t) = fm_sheet_insert_cells;
  fm_status_t (*delete_cells)(fm_workbook_t*, size_t, const fm_cell_rect_t*, int32_t) = fm_sheet_delete_cells;
  _Static_assert(FM_CELL_SHIFT_DOWN == 0 && FM_CELL_SHIFT_LEFT == 3, "fm_cell_shift_t ordinals changed");
  _Static_assert(offsetof(fm_phonetic_run_t, start) == 0 && offsetof(fm_phonetic_run_t, end) == 4 &&
                     offsetof(fm_phonetic_run_t, start_cp) == 8 && offsetof(fm_phonetic_run_t, end_cp) == 12,
                 "fm_phonetic_run_t field layout changed");
  _Static_assert(FM_TEXT_OFFSET_UNIT_UTF16 == 0 && FM_TEXT_OFFSET_UNIT_CODE_POINT == 1,
                 "fm_text_offset_unit_t ordinals changed");
  // Written out in full because these are the entry points a third-party C
  // consumer binds by signature: the XF record is read through a const
  // pointer, and the sheet-view record is written through a caller-supplied
  // pointer.
  fm_status_t (*get_cell_xf)(const fm_workbook_t*, uint32_t, fm_cell_xf_t*) = fm_workbook_get_cell_xf_at;
  fm_status_t (*add_cell_xf)(fm_workbook_t*, const fm_cell_xf_t*, uint32_t*) = fm_workbook_add_cell_xf;
  fm_status_t (*get_cell_style_xf)(const fm_workbook_t*, uint32_t, fm_cell_xf_t*) = fm_workbook_get_cell_style_xf_at;
  fm_status_t (*add_cell_style_xf)(fm_workbook_t*, const fm_cell_xf_t*, uint32_t*) = fm_workbook_add_cell_style_xf;
  fm_status_t (*get_view)(const fm_workbook_t*, size_t, fm_sheet_view_info_t*) = fm_sheet_get_view;
  fm_status_t (*update_view)(fm_workbook_t*, size_t, const fm_sheet_view_patch_t*) = fm_sheet_update_view;
  fm_sheet_view_patch_t view_patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  view_patch.zoom_pct_op = FM_PATCH_OP_SET;
  view_patch.zoom_pct = 150U;
  fm_status_t (*defined_name_at)(const fm_workbook_t*, size_t, fm_defined_name_t*) = fm_workbook_get_defined_name_at;
  fm_status_t (*pivot_add_item_at)(fm_workbook_t*, size_t, size_t, size_t, size_t, int32_t, size_t*) =
      fm_sheet_add_pivot_field_item_by_cache_index;
  // The two entry points whose records no binding marshals, written out in
  // full so a parameter added to either is a compile error here rather than a
  // silently-wrong C caller.
  fm_status_t (*add_styles_batch)(fm_workbook_t*, const fm_styles_batch_spec_t*) = fm_workbook_add_batch;
  fm_status_t (*print_area_at)(const fm_pagination_t*, size_t, fm_cell_rect_t*) = fm_pagination_get_print_area_at;
  // Both counter structs must lay out identically on native and wasm32, so a
  // binding's hand-written offsets cannot drift from the compiler's.
  // `struct_size` and five 4-byte counters, no padding, in either target.
  _Static_assert(sizeof(fm_read_diagnostics_info_t) == 6 * sizeof(uint32_t),
                 "fm_read_diagnostics_info_t must be packed");
  _Static_assert(sizeof(fm_save_diagnostics_info_t) == 6 * sizeof(uint32_t),
                 "fm_save_diagnostics_info_t must be packed");
  // Pinned from C as well as C++: a C consumer built against a narrower
  // definition of either record is miscompiled rather than diagnosed.
  _Static_assert(sizeof(fm_cell_xf_t) == 136, "fm_cell_xf_t ABI layout changed");
  _Static_assert(sizeof(fm_row_layout_info_t) == 48, "fm_row_layout_info_t ABI layout changed");
  _Static_assert(sizeof(fm_sheet_view_info_t) == 80, "fm_sheet_view_info_t ABI layout changed");
  _Static_assert(sizeof(fm_sheet_view_patch_t) == 152, "fm_sheet_view_patch_t ABI layout changed");
  // The three records a C consumer marshals without help from any binding:
  // `fm_value_t` comes back from every cell read, `fm_cell_rect_t` from
  // pagination, and `fm_styles_batch_spec_t` is passed in by pointer with
  // `struct_size` and fifteen pointers, sixteen pointer-width slots. The first two are target-independent; the batch is
  // pinned in terms of `sizeof(void*)` because its pointers are not.
  _Static_assert(sizeof(fm_value_t) == 16, "fm_value_t ABI layout changed");
  _Static_assert(offsetof(fm_value_t, u) == 8, "fm_value_t.u offset changed");
  _Static_assert(sizeof(fm_cell_rect_t) == 16, "fm_cell_rect_t ABI layout changed");
  _Static_assert(offsetof(fm_cell_rect_t, last_col) == 12, "fm_cell_rect_t.last_col offset changed");
  _Static_assert(sizeof(fm_styles_batch_spec_t) == 16 * sizeof(void*), "fm_styles_batch_spec_t ABI layout changed");
  _Static_assert(offsetof(fm_styles_batch_spec_t, num_fmt_ids) == 15 * sizeof(void*),
                 "fm_styles_batch_spec_t.num_fmt_ids offset changed");
  // Records the geometry, display and pagination-detail entry points write
  // through caller pointers. All but the width model are target-independent.
  _Static_assert(sizeof(fm_sheet_format_defaults_t) == 64, "fm_sheet_format_defaults_t ABI layout changed");
  _Static_assert(sizeof(fm_sheet_format_defaults_patch_t) == 64, "fm_sheet_format_defaults_patch_t ABI layout changed");
  fm_status_t (*update_format_defaults)(fm_workbook_t*, size_t, const fm_sheet_format_defaults_patch_t*) =
      fm_sheet_update_format_defaults;
  fm_status_t (*chars_to_pt)(const fm_workbook_t*, size_t, double, int32_t, double*) =
      fm_sheet_convert_column_chars_to_pt;
  fm_status_t (*column_outline)(fm_workbook_t*, size_t, uint32_t, uint32_t, uint32_t) = fm_sheet_set_column_outline;
  fm_status_t (*get_cell_xf_index)(const fm_workbook_t*, size_t, uint32_t, uint32_t, int32_t*, uint32_t*) =
      fm_sheet_get_cell_xf_index;
  fm_status_t (*add_dxf)(fm_workbook_t*, const fm_dxf_t*, uint32_t*) = fm_workbook_add_dxf;
  fm_status_t (*last_parse_error)(int32_t*, fm_parse_error_info_t*) = fm_get_last_parse_error;
  fm_status_t (*update_table_style_defaults)(fm_workbook_t*, const fm_table_style_defaults_patch_t*) =
      fm_workbook_update_table_style_defaults;
  fm_sheet_format_defaults_patch_t defaults_patch = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  defaults_patch.default_row_height_pt_op = FM_PATCH_OP_CLEAR;
  fm_cell_style_t custom_style = {0};
  custom_style.has_builtin_id = 0;
  _Static_assert(sizeof(fm_rect_pt_info_t) == 40, "fm_rect_pt_info_t ABI layout changed");
  _Static_assert(sizeof(fm_width_model_info_t) == 48, "fm_width_model_info_t ABI layout changed");
  _Static_assert(sizeof(fm_paper_info_t) == 32, "fm_paper_info_t ABI layout changed");
  _Static_assert(sizeof(fm_margins_pt_info_t) == 56, "fm_margins_pt_info_t ABI layout changed");
  _Static_assert(sizeof(fm_print_titles_t) == 28, "fm_print_titles_t ABI layout changed");
  _Static_assert(sizeof(fm_page_layout_info_t) == 56, "fm_page_layout_info_t ABI layout changed");
  fm_status_t (*cells_in_range)(const fm_workbook_t*, size_t, const fm_cell_rect_t*, uint64_t, uint32_t,
                                fm_cell_range_t**) = fm_sheet_get_cells_in_range;
  fm_status_t (*format_value)(const fm_workbook_t*, const fm_value_t*, const char*, const char**, int32_t*) =
      fm_workbook_format_value;
  // Records the AutoFilter, validation and threaded-comment entry points read
  // and write. The pointer-bearing ones are pinned per target.
  _Static_assert(sizeof(fm_date_group_item_t) == 12, "fm_date_group_item_t ABI layout changed");
  _Static_assert(sizeof(fm_filter_condition_t) == (sizeof(void*) == 4 ? 8 : 16),
                 "fm_filter_condition_t ABI layout changed");
  _Static_assert(sizeof(fm_filter_column_t) == (sizeof(void*) == 4 ? 144 : 168),
                 "fm_filter_column_t ABI layout changed");
  _Static_assert(sizeof(fm_sort_condition_t) == (sizeof(void*) == 4 ? 48 : 56),
                 "fm_sort_condition_t ABI layout changed");
  _Static_assert(sizeof(fm_auto_filter_t) == (sizeof(void*) == 4 ? 68 : 88), "fm_auto_filter_t ABI layout changed");
  _Static_assert(sizeof(fm_validation_outcome_info_t) == (sizeof(void*) == 4 ? 20 : 24),
                 "fm_validation_outcome_info_t ABI layout changed");
  _Static_assert(sizeof(fm_mention_t) == (sizeof(void*) == 4 ? 24 : 32), "fm_mention_t ABI layout changed");
  _Static_assert(sizeof(fm_threaded_comment_t) == (sizeof(void*) == 4 ? 44 : 80),
                 "fm_threaded_comment_t ABI layout changed");
  _Static_assert(sizeof(fm_person_t) == 5 * sizeof(void*), "fm_person_t ABI layout changed");
  _Static_assert(sizeof(fm_image_info_t) == 16, "fm_image_info_t ABI layout changed");
  _Static_assert(sizeof(fm_drawing_object_info_t) == (sizeof(void*) == 4 ? 104 : 112),
                 "fm_drawing_object_info_t ABI layout changed");
  _Static_assert(sizeof(fm_drawing_placement_spec_t) == (sizeof(void*) == 4 ? 88 : 104),
                 "fm_drawing_placement_spec_t ABI layout changed");
  _Static_assert(offsetof(fm_drawing_placement_spec_t, struct_size) == 0, "fm_drawing_placement_spec_t is a record");
  _Static_assert(sizeof(fm_drawing_anchor_t) == 72, "fm_drawing_anchor_t ABI layout changed");
  _Static_assert(sizeof(fm_drawing_anchor_patch_t) == 144, "fm_drawing_anchor_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_drawing_patch_t) == (sizeof(void*) == 4 ? 176 : 184),
                 "fm_drawing_patch_t ABI layout changed");
  _Static_assert(FM_ANCHOR_KIND_DEFAULT == 0, "a zeroed anchor kind asks for the operation's default");
  _Static_assert(FM_ANCHOR_EDIT_AS_DEFAULT == 0, "a zeroed editAs asks for the operation's default");
  // A shape's paint is the chart paint record pair.
  _Static_assert(sizeof(((fm_shape_info_t*)0)->fill) == sizeof(fm_chart_fill_t),
                 "fm_shape_info_t.fill is a chart fill");
  _Static_assert(sizeof(((fm_shape_patch_t*)0)->line) == sizeof(fm_chart_line_patch_t),
                 "fm_shape_patch_t.line is a chart line patch");
  _Static_assert(FM_DRAWING_RESTORE_NEW_ID == 1u, "FM_DRAWING_RESTORE_NEW_ID changed");
  _Static_assert(sizeof(fm_shape_info_t) == (sizeof(void*) == 4 ? 160 : 168), "fm_shape_info_t ABI layout changed");
  _Static_assert(sizeof(fm_shape_spec_t) == (sizeof(void*) == 4 ? 240 : 264), "fm_shape_spec_t ABI layout changed");
  _Static_assert(sizeof(fm_slicer_info_t) == (sizeof(void*) == 4 ? 56 : 72), "fm_slicer_info_t ABI layout changed");
  _Static_assert(sizeof(fm_slicer_spec_t) == (sizeof(void*) == 4 ? 128 : 152), "fm_slicer_spec_t ABI layout changed");
  _Static_assert(sizeof(fm_data_table_t) == 60, "fm_data_table_t ABI layout changed");
  _Static_assert(sizeof(fm_sparkline_t) == 2 * sizeof(void*), "fm_sparkline_t ABI layout changed");
  _Static_assert(sizeof(fm_sparkline_group_t) == (sizeof(void*) == 4 ? 296 : 312),
                 "fm_sparkline_group_t ABI layout changed");
  // The common drawing fields share one patch; each kind's update takes its own fields only.
  _Static_assert(sizeof(fm_shape_patch_t) == (sizeof(void*) == 4 ? 184 : 200), "fm_shape_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_slicer_patch_t) == (sizeof(void*) == 4 ? 56 : 72), "fm_slicer_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_sparkline_group_patch_t) == (sizeof(void*) == 4 ? 432 : 448),
                 "fm_sparkline_group_patch_t ABI layout changed");
  fm_status_t (*add_shape)(fm_workbook_t*, size_t, const fm_shape_spec_t*, uint32_t*) = fm_sheet_add_shape;
  fm_status_t (*add_slicer)(fm_workbook_t*, size_t, size_t, size_t, size_t, const fm_slicer_spec_t*, uint32_t*) =
      fm_sheet_add_slicer;
  fm_status_t (*update_drawing)(fm_workbook_t*, size_t, uint32_t, const fm_drawing_patch_t*) = fm_sheet_update_drawing;
  fm_status_t (*update_shape)(fm_workbook_t*, size_t, uint32_t, const fm_drawing_patch_t*, const fm_shape_patch_t*) =
      fm_sheet_update_shape;
  fm_status_t (*update_slicer)(fm_workbook_t*, size_t, uint32_t, const fm_drawing_patch_t*, const fm_slicer_patch_t*) =
      fm_sheet_update_slicer;
  fm_status_t (*update_sparkline_group)(fm_workbook_t*, size_t, const char*, const fm_sparkline_group_patch_t*) =
      fm_sheet_update_sparkline_group;
  fm_status_t (*add_sparkline_group)(fm_workbook_t*, size_t, const fm_sparkline_group_t*, const fm_sparkline_t*, size_t,
                                     const char**) = fm_sheet_add_sparkline_group;
  fm_status_t (*get_sparkline_group)(const fm_workbook_t*, size_t, const char*, int32_t*, fm_sparkline_group_t*) =
      fm_sheet_get_sparkline_group;
  fm_status_t (*chart_data)(const fm_workbook_t*, size_t, uint32_t, int32_t*, fm_chart_data_t**) =
      fm_sheet_get_chart_data;
  fm_status_t (*update_chart_series)(fm_workbook_t*, size_t, uint32_t, size_t, uint32_t,
                                     const fm_chart_series_patch_t*) = fm_sheet_update_chart_series;
  fm_shape_patch_t shape_patch = FM_STRUCT_INIT(fm_shape_patch_t);
  shape_patch.text_box_op = FM_PATCH_OP_SET;
  shape_patch.text_box = 1;
  fm_drawing_patch_t drawing_patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  drawing_patch.hidden_op = FM_PATCH_OP_SET;
  drawing_patch.hidden = 1;
  drawing_patch.anchor.from_row_op = FM_PATCH_OP_SET;
  drawing_patch.anchor.from_row = 3;
  // Table and custom table-style records, pinned per target.
  _Static_assert(sizeof(fm_table_info_t) == (sizeof(void*) == 4 ? 56 : 72), "fm_table_info_t ABI layout changed");
  _Static_assert(sizeof(fm_table_patch_t) == (sizeof(void*) == 4 ? 52 : 64), "fm_table_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_table_column_patch_t) == (sizeof(void*) == 4 ? 32 : 56),
                 "fm_table_column_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_table_column_info_t) == 6 * sizeof(void*), "fm_table_column_info_t ABI layout changed");
  _Static_assert(sizeof(fm_table_style_element_t) == 16, "fm_table_style_element_t ABI layout changed");
  _Static_assert(sizeof(fm_table_style_t) == (sizeof(void*) == 4 ? 24 : 40), "fm_table_style_t ABI layout changed");
  // Protection records, pinned per target.
  _Static_assert(sizeof(fm_workbook_protection_t) == (sizeof(void*) == 4 ? 56 : 88),
                 "fm_workbook_protection_t ABI layout changed");
  _Static_assert(sizeof(fm_protected_range_t) == (sizeof(void*) == 4 ? 40 : 72),
                 "fm_protected_range_t ABI layout changed");
  _Static_assert(sizeof(fm_protected_range_patch_t) == (sizeof(void*) == 4 ? 72 : 120),
                 "fm_protected_range_patch_t ABI layout changed");
  // Threaded-comment and iterative-calculation patches: one `fm_patch_op_t`
  // before each field.
  _Static_assert(sizeof(fm_threaded_comment_patch_t) == (sizeof(void*) == 4 ? 32 : 48),
                 "fm_threaded_comment_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_iterative_settings_patch_t) == 32, "fm_iterative_settings_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_iterative_settings_info_t) == 24, "fm_iterative_settings_info_t ABI layout changed");
  // Keyed reads take `out_present` before their result.
  fm_status_t (*get_comment_visible)(const fm_workbook_t*, size_t, uint32_t, uint32_t, int32_t*, int32_t*) =
      fm_sheet_get_comment_visible;
  fm_status_t (*get_pinned_now)(const fm_workbook_t*, int32_t*, fm_civil_time_t*) = fm_workbook_get_pinned_now;
  fm_status_t (*get_file_name)(const fm_workbook_t*, int32_t*, const char**) = fm_workbook_get_file_name;
  fm_status_t (*get_iterative)(const fm_workbook_t*, fm_iterative_settings_info_t*) = fm_workbook_get_iterative;
  fm_status_t (*get_spill)(const fm_workbook_t*, size_t, uint32_t, uint32_t, int32_t*, fm_spill_info_t*) =
      fm_sheet_get_spill;
  fm_status_t (*get_array_formula)(const fm_workbook_t*, size_t, uint32_t, uint32_t, int32_t*,
                                   fm_array_formula_info_t*) = fm_sheet_get_array_formula;
  fm_status_t (*get_table_style_options)(const fm_workbook_t*, size_t, size_t, int32_t*, fm_table_style_options_t*) =
      fm_sheet_get_table_style_options;
  fm_status_t (*get_pivot_style_options)(const fm_workbook_t*, size_t, size_t, int32_t*, fm_pivot_style_options_t*) =
      fm_sheet_get_pivot_style_options;
  fm_status_t (*get_defined_name)(const fm_workbook_t*, const char*, int32_t, int32_t*, fm_defined_name_t*) =
      fm_workbook_get_defined_name;
  // A date grouping is removed by passing a `NULL` group to its setter.
  fm_status_t (*set_pivot_field_date_group)(fm_workbook_t*, size_t, size_t, size_t, const fm_pivot_date_group_spec_t*) =
      fm_sheet_set_pivot_field_date_group;
  fm_status_t (*pivot_data_field_default_name)(const fm_workbook_t*, size_t, size_t, size_t, int32_t, size_t,
                                               const char**) = fm_sheet_get_pivot_data_field_default_name;
  fm_status_t (*merges_in_range)(const fm_workbook_t*, size_t, const fm_cell_rect_t*, fm_cell_rect_t*, size_t,
                                 size_t*) = fm_sheet_get_merges_in_range;
  // Protected ranges are keyed by name; the add returns nothing because the
  // caller's name is the key.
  fm_status_t (*get_protected_range)(const fm_workbook_t*, size_t, const char*, int32_t*, fm_protected_range_t*) =
      fm_sheet_get_protected_range;
  fm_status_t (*add_protected_range)(fm_workbook_t*, size_t, const fm_protected_range_t*, const char**) =
      fm_sheet_add_protected_range;
  fm_status_t (*update_protected_range)(fm_workbook_t*, size_t, const char*, const fm_protected_range_patch_t*) =
      fm_sheet_update_protected_range;
  fm_status_t (*remove_protected_range)(fm_workbook_t*, size_t, const char*) = fm_sheet_remove_protected_range;
  fm_threaded_comment_patch_t comment_patch = FM_STRUCT_INIT(fm_threaded_comment_patch_t);
  comment_patch.done_op = FM_PATCH_OP_SET;
  comment_patch.done = 1;
  fm_status_t (*get_auto_filter)(const fm_workbook_t*, size_t, int32_t*, fm_auto_filter_t*) = fm_sheet_get_auto_filter;
  fm_status_t (*set_table_auto_filter)(fm_workbook_t*, size_t, size_t, const fm_auto_filter_t*) =
      fm_sheet_set_table_auto_filter;
  fm_status_t (*evaluate_auto_filter)(const fm_workbook_t*, size_t, uint8_t*, size_t, size_t*, uint32_t*) =
      fm_sheet_evaluate_auto_filter;
  fm_status_t (*validate_value)(const fm_workbook_t*, size_t, uint32_t, uint32_t, const fm_value_t*,
                                fm_validation_outcome_info_t*) = fm_sheet_validate_value;
  fm_status_t (*list_invalid_cells)(const fm_workbook_t*, size_t, uint64_t, uint32_t, fm_cell_range_t**) =
      fm_sheet_get_invalid_cells;
  fm_status_t (*add_threaded_comment)(fm_workbook_t*, size_t, int32_t, const fm_threaded_comment_t*, const char**) =
      fm_sheet_add_threaded_comment;
  fm_status_t (*edit_threaded_comment)(fm_workbook_t*, size_t, const char*, int32_t,
                                       const fm_threaded_comment_patch_t*) = fm_sheet_update_threaded_comment;
  fm_status_t (*person_at)(const fm_workbook_t*, size_t, fm_person_t*) = fm_workbook_get_person_at;
  fm_status_t (*insert_image)(fm_workbook_t*, size_t, const uint8_t*, size_t, const fm_drawing_placement_spec_t*,
                              uint32_t*) = fm_sheet_add_picture;
  fm_status_t (*get_image)(const fm_workbook_t*, size_t, uint32_t, int32_t*, const uint8_t**, size_t*,
                           fm_image_info_t*) = fm_sheet_get_picture;
  fm_status_t (*set_drawing_z_order)(fm_workbook_t*, size_t, uint32_t, size_t) = fm_sheet_set_drawing_z_order;
  fm_status_t (*snapshot_drawing)(const fm_workbook_t*, size_t, uint32_t, uint8_t**, size_t*) =
      fm_sheet_snapshot_drawing;
  fm_status_t (*restore_drawing)(fm_workbook_t*, size_t, const uint8_t*, size_t, uint32_t, uint32_t*) =
      fm_sheet_restore_drawing;
  fm_status_t (*reset_theme)(fm_workbook_t*) = fm_workbook_reset_theme;
  fm_status_t (*set_table_style)(fm_workbook_t*, const fm_table_style_t*) = fm_workbook_set_table_style;
  fm_status_t (*table_at)(const fm_workbook_t*, size_t, size_t, fm_table_info_t*) = fm_sheet_get_table_at;
  fm_status_t (*update_table)(fm_workbook_t*, size_t, size_t, const fm_table_patch_t*) = fm_sheet_update_table;
  fm_table_patch_t table_patch = FM_STRUCT_INIT(fm_table_patch_t);
  table_patch.header_row_op = FM_PATCH_OP_SET;
  table_patch.header_row = 1;
  fm_table_column_patch_t column_patch = {0};
  column_patch.totals_row_function_op = FM_PATCH_OP_SET;
  column_patch.totals_row_function = FM_TABLE_TOTALS_ROW_FUNCTION_SUM;
  column_patch.calculated_formula_op = FM_PATCH_OP_CLEAR;
  table_patch.columns_op = FM_PATCH_OP_SET;
  table_patch.columns = &column_patch;
  table_patch.columns_count = 1;
  // Print settings: each typed setter takes its own patch, and the print area
  // is copied into a caller buffer.
  _Static_assert(sizeof(fm_page_setup_patch_t) == 52, "fm_page_setup_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_page_margins_patch_t) == 96, "fm_page_margins_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_print_options_patch_t) == 36, "fm_print_options_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_header_footer_patch_t) == (sizeof(void*) == 4 ? 84 : 128),
                 "fm_header_footer_patch_t ABI layout changed");
  fm_status_t (*update_page_setup)(fm_workbook_t*, size_t, const fm_page_setup_patch_t*) = fm_sheet_update_page_setup;
  fm_status_t (*update_page_margins)(fm_workbook_t*, size_t, const fm_page_margins_patch_t*) =
      fm_sheet_update_page_margins;
  fm_status_t (*update_print_options)(fm_workbook_t*, size_t, const fm_print_options_patch_t*) =
      fm_sheet_update_print_options;
  fm_status_t (*update_header_footer)(fm_workbook_t*, size_t, const fm_header_footer_patch_t*) =
      fm_sheet_update_header_footer;
  fm_status_t (*get_print_area)(const fm_workbook_t*, size_t, fm_cell_rect_t*, size_t, size_t*) =
      fm_sheet_get_print_area;
  fm_status_t (*evaluate_cf_range)(const fm_workbook_t*, size_t, const fm_cell_rect_t*,
                                   const fm_cf_range_evaluation_spec_t*, fm_cf_results_t**) =
      fm_sheet_evaluate_cf_range;
  fm_page_setup_patch_t setup_patch = FM_STRUCT_INIT(fm_page_setup_patch_t);
  setup_patch.orientation_op = FM_PATCH_OP_CLEAR;
  fm_header_footer_patch_t hf_patch = FM_STRUCT_INIT(fm_header_footer_patch_t);
  hf_patch.odd_header_op = FM_PATCH_OP_SET;
  hf_patch.odd_header = "&CPage &P";
  // Pivots: options, cache options and grand totals merge a patch, the field
  // orders fill a caller buffer, and a non-calculated field reads no formula.
  _Static_assert(sizeof(fm_pivot_cache_options_patch_t) == 28, "fm_pivot_cache_options_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_pivot_grand_totals_patch_t) == 20, "fm_pivot_grand_totals_patch_t ABI layout changed");
  _Static_assert(sizeof(fm_pivot_options_patch_t) == (sizeof(void*) == 4 ? 212 : 248),
                 "fm_pivot_options_patch_t ABI layout changed");
  fm_status_t (*update_pivot_options)(fm_workbook_t*, size_t, size_t, const fm_pivot_options_patch_t*) =
      fm_sheet_update_pivot_options;
  fm_status_t (*update_pivot_cache_options)(fm_workbook_t*, uint32_t, const fm_pivot_cache_options_patch_t*) =
      fm_workbook_update_pivot_cache_options;
  fm_status_t (*update_pivot_grand_totals)(fm_workbook_t*, size_t, size_t, const fm_pivot_grand_totals_patch_t*) =
      fm_sheet_update_pivot_grand_totals;
  fm_status_t (*pivot_row_field_order)(const fm_workbook_t*, size_t, size_t, uint32_t*, size_t, size_t*) =
      fm_sheet_get_pivot_row_field_order;
  fm_status_t (*pivot_page_field_order)(const fm_workbook_t*, size_t, size_t, uint32_t*, size_t, size_t*) =
      fm_sheet_get_pivot_page_field_order;
  fm_status_t (*set_pivot_page_field_order)(fm_workbook_t*, size_t, size_t, const uint32_t*, size_t) =
      fm_sheet_set_pivot_page_field_order;
  // Every add returns its key or index.
  fm_status_t (*add_shared_item_text)(fm_workbook_t*, uint32_t, size_t, const char*, size_t*) =
      fm_workbook_add_pivot_cache_field_shared_item_text;
  fm_status_t (*add_shared_item_blank)(fm_workbook_t*, uint32_t, size_t, size_t*) =
      fm_workbook_add_pivot_cache_field_shared_item_blank;
  fm_status_t (*add_person)(fm_workbook_t*, const fm_person_t*, const char**) = fm_workbook_add_person;
  fm_status_t (*pivot_cache_field_formula)(const fm_workbook_t*, uint32_t, size_t, int32_t*, const char**) =
      fm_workbook_get_pivot_cache_field_formula;
  fm_pivot_field_patch_t pivot_field_patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  pivot_field_patch.custom_name_op = FM_PATCH_OP_CLEAR;
  // Sheets are placed by `insert_before` in the current list; a `NULL` name
  // takes the profile's default; row and column edits take an inclusive span.
  fm_status_t (*add_sheet)(fm_workbook_t*, const char*, size_t, size_t*) = fm_workbook_add_sheet;
  fm_status_t (*move_sheet)(fm_workbook_t*, size_t, size_t) = fm_workbook_move_sheet;
  fm_status_t (*insert_rows)(fm_workbook_t*, size_t, uint32_t, uint32_t) = fm_sheet_insert_rows;
  fm_status_t (*delete_columns)(fm_workbook_t*, size_t, uint32_t, uint32_t) = fm_sheet_delete_columns;
  fm_status_t (*set_value)(fm_workbook_t*, size_t, uint32_t, uint32_t, const fm_value_t*) = fm_sheet_set_value;
  fm_status_t (*get_visibility)(const fm_workbook_t*, size_t, int32_t*) = fm_sheet_get_visibility;
  fm_status_t (*set_visibility)(fm_workbook_t*, size_t, int32_t) = fm_sheet_set_visibility;
  fm_status_t (*profile_id_count)(size_t*) = fm_get_profile_id_count;
  fm_status_t (*profile_id_at)(size_t, const char**) = fm_get_profile_id_at;
  // Data validations are addressed by their engine id; add returns it.
  _Static_assert(sizeof(fm_data_validation_t) == (sizeof(void*) == 4 ? 68 : 104),
                 "fm_data_validation_t ABI layout changed");
  _Static_assert(offsetof(fm_data_validation_t, id) == sizeof(void*), "fm_data_validation_t leads with its id");
  _Static_assert(sizeof(fm_data_validation_patch_t) == (sizeof(void*) == 4 ? 120 : 168),
                 "fm_data_validation_patch_t ABI layout changed");
  fm_status_t (*get_validation)(const fm_workbook_t*, size_t, const char*, int32_t*, fm_data_validation_t*) =
      fm_sheet_get_validation;
  fm_status_t (*add_validation)(fm_workbook_t*, size_t, const fm_data_validation_t*, const char**) =
      fm_sheet_add_validation;
  fm_status_t (*update_validation)(fm_workbook_t*, size_t, const char*, const fm_data_validation_patch_t*) =
      fm_sheet_update_validation;
  fm_status_t (*remove_validation)(fm_workbook_t*, size_t, const char*) = fm_sheet_remove_validation;
  fm_data_validation_patch_t validation_patch = FM_STRUCT_INIT(fm_data_validation_patch_t);
  validation_patch.formula1_op = FM_PATCH_OP_CLEAR;
  _Static_assert(sizeof(fm_cf_rule_patch_t) == (sizeof(void*) == 4 ? 472 : 544),
                 "fm_cf_rule_patch_t ABI layout changed");
  // CF rules are keyed by id.
  fm_status_t (*get_cf_rule)(const fm_workbook_t*, size_t, const char*, int32_t*, fm_cf_rule_t*) = fm_sheet_get_cf_rule;
  fm_status_t (*add_cf_rule)(fm_workbook_t*, size_t, const fm_cf_rule_t*, const char**) = fm_sheet_add_cf_rule;
  fm_status_t (*update_cf_rule)(fm_workbook_t*, size_t, const char*, const fm_cf_rule_patch_t*) =
      fm_sheet_update_cf_rule;
  fm_status_t (*remove_cf_rule)(fm_workbook_t*, size_t, const char*) = fm_sheet_remove_cf_rule;
  fm_cf_rule_patch_t cf_patch = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  // A chart's insert title is the setter's record behind `has_title`.
  fm_chart_spec_t chart_spec = FM_STRUCT_INIT(fm_chart_spec_t);
  chart_spec.has_title = 1;
  chart_spec.title.kind = FM_CHART_TITLE_KIND_NONE;
  cf_patch.dxf_index_op = FM_PATCH_OP_CLEAR;
  // A pivot filter names its field by index; subtotal functions are set as a
  // whole; cache records read back.
  _Static_assert(sizeof(fm_pivot_filter_t) == (sizeof(void*) == 4 ? 64 : 72), "fm_pivot_filter_t ABI layout changed");
  _Static_assert(offsetof(fm_pivot_filter_t, field_index) == 4, "fm_pivot_filter_t names its field by index");
  fm_status_t (*set_subtotal_functions)(fm_workbook_t*, size_t, size_t, size_t, const int32_t*, size_t) =
      fm_sheet_set_pivot_field_subtotal_functions;
  fm_status_t (*cache_record_value)(const fm_workbook_t*, uint32_t, size_t, size_t, fm_value_t*) =
      fm_workbook_get_pivot_cache_record_value;
  _Static_assert(sizeof(fm_external_link_info_t) == (sizeof(void*) == 4 ? 28 : 40),
                 "fm_external_link_info_t ABI layout changed");
  _Static_assert(offsetof(fm_external_link_info_t, book_number) == 4, "fm_external_link_info_t book_number moved");
  // Each style table with an `add` also overwrites a slot in place.
  fm_status_t (*set_fill)(fm_workbook_t*, uint32_t, const fm_fill_t*) = fm_workbook_set_fill;
  fm_status_t (*set_border)(fm_workbook_t*, uint32_t, const fm_border_t*) = fm_workbook_set_border;
  fm_status_t (*set_cell_xf)(fm_workbook_t*, uint32_t, const fm_cell_xf_t*) = fm_workbook_set_cell_xf;
  fm_status_t (*set_dxf)(fm_workbook_t*, uint32_t, const fm_dxf_t*) = fm_workbook_set_dxf;
  fm_status_t (*set_cell_style_xf)(fm_workbook_t*, uint32_t, const fm_cell_xf_t*) = fm_workbook_set_cell_style_xf;
  // Font, fill and border indices are uint32_t like every other style index,
  // and a numFmt id is 32 bits wide in the xf record.
  fm_status_t (*get_font_at)(const fm_workbook_t*, uint32_t, fm_font_t*) = fm_workbook_get_font_at;
  fm_status_t (*add_font)(fm_workbook_t*, const fm_font_t*, uint32_t*) = fm_workbook_add_font;
  fm_status_t (*set_font)(fm_workbook_t*, uint32_t, const fm_font_t*) = fm_workbook_set_font;
  fm_status_t (*get_fill_at)(const fm_workbook_t*, uint32_t, fm_fill_t*) = fm_workbook_get_fill_at;
  fm_status_t (*add_fill)(fm_workbook_t*, const fm_fill_t*, uint32_t*) = fm_workbook_add_fill;
  fm_status_t (*get_border_at)(const fm_workbook_t*, uint32_t, fm_border_t*) = fm_workbook_get_border_at;
  fm_status_t (*add_border)(fm_workbook_t*, const fm_border_t*, uint32_t*) = fm_workbook_add_border;
  _Static_assert(sizeof(((fm_cell_xf_t*)0)->num_fmt_id) == sizeof(uint32_t), "fm_cell_xf_t.num_fmt_id is 32 bits");
  // The declared number formats list as records; grand totals and the style
  // defaults read into records; a cache cell says whether it holds a date.
  fm_status_t (*number_format_count)(const fm_workbook_t*, size_t*) = fm_workbook_get_number_format_count;
  fm_status_t (*number_format_at)(const fm_workbook_t*, size_t, fm_number_format_info_t*) =
      fm_workbook_get_number_format_at;
  fm_status_t (*get_pivot_grand_totals)(const fm_workbook_t*, size_t, size_t, fm_pivot_grand_totals_t*) =
      fm_sheet_get_pivot_grand_totals;
  fm_status_t (*get_table_style_defaults)(const fm_workbook_t*, fm_table_style_defaults_t*) =
      fm_workbook_get_table_style_defaults;
  fm_status_t (*cache_record_is_date)(const fm_workbook_t*, uint32_t, size_t, size_t, int32_t*) =
      fm_workbook_is_pivot_cache_record_value_date;
  fm_table_style_defaults_patch_t style_defaults_patch = FM_STRUCT_INIT(fm_table_style_defaults_patch_t);
  style_defaults_patch.pivot_style_name_op = FM_PATCH_OP_CLEAR;
  fm_pivot_filter_t pivot_filter = FM_STRUCT_INIT(fm_pivot_filter_t);
  pivot_filter.kind = FM_PIVOT_FILTER_KIND_VALUE_TOP_COUNT;
  _Static_assert(FM_PIVOT_FILTER_KIND_DATE_BETWEEN == 5, "fm_pivot_filter_kind_t ordinals changed");
  _Static_assert(FM_ICON_SET_NAME_THREE_ARROWS == 0 && FM_ICON_SET_NAME_FIVE_BOXES == 19,
                 "fm_icon_set_name_t ordinals changed");
  _Static_assert(FM_COLOR_CONTEXT_FILL_FOREGROUND == 1 && FM_COLOR_CONTEXT_FILL_BACKGROUND == 2,
                 "fm_color_context_t ordinals changed");
  (void)add_sheet;
  (void)move_sheet;
  (void)insert_rows;
  (void)delete_columns;
  (void)set_value;
  (void)get_visibility;
  (void)set_visibility;
  (void)profile_id_count;
  (void)profile_id_at;
  (void)get_validation;
  (void)add_validation;
  (void)update_validation;
  (void)remove_validation;
  (void)validation_patch;
  (void)get_cf_rule;
  (void)add_cf_rule;
  (void)update_cf_rule;
  (void)remove_cf_rule;
  (void)cf_patch;
  (void)chart_spec;
  (void)set_subtotal_functions;
  (void)cache_record_value;
  (void)set_fill;
  (void)set_border;
  (void)set_cell_xf;
  (void)set_dxf;
  (void)set_cell_style_xf;
  (void)get_font_at;
  (void)add_font;
  (void)set_font;
  (void)get_fill_at;
  (void)add_fill;
  (void)get_border_at;
  (void)add_border;
  (void)number_format_count;
  (void)number_format_at;
  (void)get_pivot_grand_totals;
  (void)get_table_style_defaults;
  (void)cache_record_is_date;
  (void)style_defaults_patch;
  (void)pivot_filter;
  (void)cells_in_range;
  (void)set_table_style;
  (void)table_at;
  (void)update_table;
  (void)table_patch;
  (void)add_shape;
  (void)add_slicer;
  (void)update_drawing;
  (void)add_sparkline_group;
  (void)get_sparkline_group;
  (void)drawing_patch;
  (void)update_shape;
  (void)update_slicer;
  (void)update_sparkline_group;
  (void)chart_data;
  (void)update_chart_series;
  (void)shape_patch;
  (void)update_page_setup;
  (void)update_page_margins;
  (void)update_print_options;
  (void)update_header_footer;
  (void)get_print_area;
  (void)evaluate_cf_range;
  (void)setup_patch;
  (void)hf_patch;
  (void)update_pivot_options;
  (void)update_pivot_cache_options;
  (void)update_pivot_grand_totals;
  (void)pivot_row_field_order;
  (void)pivot_page_field_order;
  (void)set_pivot_page_field_order;
  (void)add_shared_item_text;
  (void)add_shared_item_blank;
  (void)add_person;
  (void)pivot_cache_field_formula;
  (void)pivot_field_patch;
  (void)format_value;
  (void)get_auto_filter;
  (void)set_table_auto_filter;
  (void)evaluate_auto_filter;
  (void)validate_value;
  (void)list_invalid_cells;
  (void)add_threaded_comment;
  (void)edit_threaded_comment;
  (void)get_comment_visible;
  (void)get_pinned_now;
  (void)get_file_name;
  (void)get_iterative;
  (void)get_spill;
  (void)get_array_formula;
  (void)get_table_style_options;
  (void)get_pivot_style_options;
  (void)get_defined_name;
  (void)set_pivot_field_date_group;
  (void)pivot_data_field_default_name;
  (void)merges_in_range;
  (void)get_protected_range;
  (void)add_protected_range;
  (void)update_protected_range;
  (void)remove_protected_range;
  (void)comment_patch;
  (void)person_at;
  (void)insert_image;
  (void)get_image;
  (void)set_drawing_z_order;
  (void)snapshot_drawing;
  (void)restore_drawing;
  (void)reset_theme;
  (void)save_diagnostics;
  (void)read_diagnostics;
  (void)save;
  (void)insert_cells;
  (void)delete_cells;
  (void)get_cell_xf;
  (void)add_cell_xf;
  (void)get_cell_style_xf;
  (void)add_cell_style_xf;
  (void)get_view;
  (void)update_view;
  (void)view_patch;
  (void)update_format_defaults;
  (void)chars_to_pt;
  (void)column_outline;
  (void)get_cell_xf_index;
  (void)add_dxf;
  (void)last_parse_error;
  (void)update_table_style_defaults;
  (void)defaults_patch;
  (void)custom_style;
  (void)defined_name_at;
  (void)pivot_add_item_at;
  (void)add_styles_batch;
  (void)print_area_at;
  (void)parallel_recalc;
  (void)stats;
  (void)rule;
  (void)version_matches;
  return callback(1U, 0.0, 1U, 0) ? 0 : 1;
}
