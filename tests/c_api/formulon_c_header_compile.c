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

int main(void) {
  fm_iterative_progress_cb callback = ContinueIteration;
  fm_parallel_recalc_stats_t stats = {0};
  fm_status_t (*parallel_recalc)(fm_workbook_t*, uint32_t, fm_parallel_recalc_stats_t*) = fm_workbook_recalc_parallel;
  fm_status_t (*save_diagnostics)(const fm_workbook_t*, int32_t, uint8_t**, size_t*, fm_save_diagnostics_t*) =
      fm_workbook_save_with_diagnostics;
  fm_status_t (*read_diagnostics)(const fm_workbook_t*, fm_read_diagnostics_t*) = fm_workbook_get_diagnostics;
  fm_status_t (*save_as)(const fm_workbook_t*, int32_t, uint8_t**, size_t*) = fm_workbook_save_as;
  // Written out in full because these are the entry points a third-party C
  // consumer binds by signature: the XF record is read through a const
  // pointer, and the sheet-view record is written through a caller-supplied
  // pointer.
  fm_status_t (*get_cell_xf)(const fm_workbook_t*, size_t, fm_cell_xf_t*) = fm_workbook_get_cell_xf;
  fm_status_t (*add_cell_xf)(fm_workbook_t*, const fm_cell_xf_t*, size_t*) = fm_workbook_add_cell_xf;
  fm_status_t (*get_cell_style_xf)(const fm_workbook_t*, size_t, fm_cell_xf_t*) = fm_workbook_get_cell_style_xf;
  fm_status_t (*add_cell_style_xf)(fm_workbook_t*, const fm_cell_xf_t*, size_t*) = fm_workbook_add_cell_style_xf;
  fm_status_t (*get_view)(const fm_workbook_t*, size_t, fm_sheet_view_t*) = fm_sheet_get_view;
  fm_status_t (*defined_name_at)(const fm_workbook_t*, size_t, fm_defined_name_t*) = fm_workbook_get_defined_name_at;
  fm_status_t (*pivot_add_item_at)(fm_workbook_t*, size_t, size_t, size_t, size_t, int32_t, size_t*) =
      fm_sheet_add_pivot_field_shared_item;
  // The two entry points whose records no binding marshals, written out in
  // full so a parameter added to either is a compile error here rather than a
  // silently-wrong C caller.
  fm_status_t (*add_styles_batch)(fm_workbook_t*, const fm_styles_batch_t*) = fm_workbook_add_batch;
  fm_status_t (*print_area_at)(const fm_pagination_t*, size_t, fm_cell_rect_t*) = fm_pagination_get_print_area_at;
  // Both counter structs must lay out identically on native and wasm32, so a
  // binding's hand-written offsets cannot drift from the compiler's. Five
  // 4-byte counters, no padding, in either target.
  _Static_assert(sizeof(fm_read_diagnostics_t) == 5 * sizeof(uint32_t), "fm_read_diagnostics_t must be packed");
  _Static_assert(sizeof(fm_save_diagnostics_t) == 5 * sizeof(uint32_t), "fm_save_diagnostics_t must be packed");
  // Pinned from C as well as C++: a C consumer built against a narrower
  // definition of either record is miscompiled rather than diagnosed.
  _Static_assert(sizeof(fm_cell_xf_t) == 128, "fm_cell_xf_t ABI layout changed");
  _Static_assert(sizeof(fm_row_layout_t) == 40, "fm_row_layout_t ABI layout changed");
  _Static_assert(sizeof(fm_sheet_view_t) == 80, "fm_sheet_view_t ABI layout changed");
  // The three records a C consumer marshals without help from any binding:
  // `fm_value_t` comes back from every cell read, `fm_cell_rect_t` from
  // pagination, and `fm_styles_batch_t` is passed in by pointer with fifteen
  // pointer-width slots. The first two are target-independent; the batch is
  // pinned in terms of `sizeof(void*)` because its pointers are not.
  _Static_assert(sizeof(fm_value_t) == 16, "fm_value_t ABI layout changed");
  _Static_assert(offsetof(fm_value_t, u) == 8, "fm_value_t.u offset changed");
  _Static_assert(sizeof(fm_cell_rect_t) == 16, "fm_cell_rect_t ABI layout changed");
  _Static_assert(offsetof(fm_cell_rect_t, last_col) == 12, "fm_cell_rect_t.last_col offset changed");
  _Static_assert(sizeof(fm_styles_batch_t) == 15 * sizeof(void*), "fm_styles_batch_t ABI layout changed");
  _Static_assert(offsetof(fm_styles_batch_t, num_fmt_ids) == 14 * sizeof(void*),
                 "fm_styles_batch_t.num_fmt_ids offset changed");
  // Records the geometry, display and pagination-detail entry points write
  // through caller pointers. All but the width model are target-independent.
  _Static_assert(sizeof(fm_sheet_format_defaults_t) == 56, "fm_sheet_format_defaults_t ABI layout changed");
  _Static_assert(sizeof(fm_rect_pt_t) == 32, "fm_rect_pt_t ABI layout changed");
  _Static_assert(sizeof(fm_width_model_t) == (sizeof(void*) == 4 ? 40 : 48), "fm_width_model_t ABI layout changed");
  _Static_assert(sizeof(fm_paper_info_t) == 24, "fm_paper_info_t ABI layout changed");
  _Static_assert(sizeof(fm_margins_pt_t) == 48, "fm_margins_pt_t ABI layout changed");
  _Static_assert(sizeof(fm_print_titles_t) == 24, "fm_print_titles_t ABI layout changed");
  _Static_assert(sizeof(fm_page_layout_t) == 56, "fm_page_layout_t ABI layout changed");
  fm_status_t (*cells_in_range)(const fm_workbook_t*, size_t, const fm_cell_rect_t*, uint64_t, uint32_t,
                                fm_cell_range_t**) = fm_sheet_get_cells_in_range;
  fm_status_t (*format_value)(const fm_workbook_t*, const fm_value_t*, const char*, const char**, int32_t*) =
      fm_workbook_format_value;
  // Records the AutoFilter, validation and threaded-comment entry points read
  // and write. The pointer-bearing ones are pinned per target.
  _Static_assert(sizeof(fm_date_group_item_t) == 8, "fm_date_group_item_t ABI layout changed");
  _Static_assert(sizeof(fm_filter_column_t) == (sizeof(void*) == 4 ? 152 : 192),
                 "fm_filter_column_t ABI layout changed");
  _Static_assert(sizeof(fm_sort_condition_t) == (sizeof(void*) == 4 ? 48 : 56),
                 "fm_sort_condition_t ABI layout changed");
  _Static_assert(sizeof(fm_auto_filter_t) == (sizeof(void*) == 4 ? 64 : 80), "fm_auto_filter_t ABI layout changed");
  _Static_assert(sizeof(fm_validation_outcome_t) == 16, "fm_validation_outcome_t ABI layout changed");
  _Static_assert(sizeof(fm_mention_t) == (sizeof(void*) == 4 ? 16 : 24), "fm_mention_t ABI layout changed");
  _Static_assert(sizeof(fm_threaded_comment_t) == (sizeof(void*) == 4 ? 40 : 72),
                 "fm_threaded_comment_t ABI layout changed");
  _Static_assert(sizeof(fm_person_t) == 4 * sizeof(void*), "fm_person_t ABI layout changed");
  _Static_assert(sizeof(fm_image_info_t) == 12, "fm_image_info_t ABI layout changed");
  _Static_assert(sizeof(fm_drawing_object_t) == (sizeof(void*) == 4 ? 96 : 112),
                 "fm_drawing_object_t ABI layout changed");
  _Static_assert(sizeof(fm_image_insert_t) == (sizeof(void*) == 4 ? 56 : 64), "fm_image_insert_t ABI layout changed");
  _Static_assert(sizeof(fm_drawing_anchor_t) == 48, "fm_drawing_anchor_t ABI layout changed");
  _Static_assert(FM_DRAWING_RESTORE_NEW_ID == 1u, "FM_DRAWING_RESTORE_NEW_ID changed");
  _Static_assert(sizeof(fm_shape_t) == (sizeof(void*) == 4 ? 136 : 144), "fm_shape_t ABI layout changed");
  _Static_assert(sizeof(fm_slicer_t) == (sizeof(void*) == 4 ? 104 : 128), "fm_slicer_t ABI layout changed");
  _Static_assert(sizeof(fm_data_table_t) == 56, "fm_data_table_t ABI layout changed");
  _Static_assert(sizeof(fm_sparkline_t) == 2 * sizeof(void*), "fm_sparkline_t ABI layout changed");
  _Static_assert(sizeof(fm_sparkline_group_t) == (sizeof(void*) == 4 ? 288 : 296),
                 "fm_sparkline_group_t ABI layout changed");
  // Table and custom table-style records, pinned per target.
  _Static_assert(sizeof(fm_table_info_t) == (sizeof(void*) == 4 ? 36 : 40), "fm_table_info_t ABI layout changed");
  _Static_assert(sizeof(fm_table_column_info_t) == 4 * sizeof(void*), "fm_table_column_info_t ABI layout changed");
  _Static_assert(sizeof(fm_table_style_element_t) == (sizeof(void*) == 4 ? 16 : 24),
                 "fm_table_style_element_t ABI layout changed");
  _Static_assert(sizeof(fm_table_style_t) == (sizeof(void*) == 4 ? 20 : 32), "fm_table_style_t ABI layout changed");
  // Protection records, pinned per target.
  _Static_assert(sizeof(fm_workbook_protection_t) == (sizeof(void*) == 4 ? 52 : 88),
                 "fm_workbook_protection_t ABI layout changed");
  _Static_assert(sizeof(fm_protected_range_t) == (sizeof(void*) == 4 ? 36 : 72),
                 "fm_protected_range_t ABI layout changed");
  fm_status_t (*get_auto_filter)(const fm_workbook_t*, size_t, fm_auto_filter_t*, int32_t*) = fm_sheet_get_auto_filter;
  fm_status_t (*set_table_auto_filter)(fm_workbook_t*, size_t, const fm_auto_filter_t*) =
      fm_workbook_set_table_auto_filter;
  fm_status_t (*evaluate_auto_filter)(const fm_workbook_t*, size_t, uint8_t*, size_t, size_t*, uint32_t*) =
      fm_sheet_evaluate_auto_filter;
  fm_status_t (*validate_value)(const fm_workbook_t*, size_t, uint32_t, uint32_t, const fm_value_t*,
                                fm_validation_outcome_t*) = fm_sheet_validate_value;
  fm_status_t (*list_invalid_cells)(const fm_workbook_t*, size_t, uint64_t, uint32_t, fm_cell_range_t**) =
      fm_sheet_get_invalid_cells;
  fm_status_t (*add_threaded_comment)(fm_workbook_t*, size_t, const fm_threaded_comment_t*) =
      fm_sheet_add_threaded_comment;
  fm_status_t (*edit_threaded_comment)(fm_workbook_t*, size_t, const char*, const char*, const fm_mention_t*,
                                       uint32_t) = fm_sheet_update_threaded_comment;
  fm_status_t (*person_at)(const fm_workbook_t*, size_t, fm_person_t*) = fm_workbook_get_person_at;
  fm_status_t (*insert_image)(fm_workbook_t*, size_t, const uint8_t*, size_t, const fm_image_insert_t*, uint32_t*) =
      fm_sheet_add_image;
  fm_status_t (*get_image)(const fm_workbook_t*, size_t, uint32_t, int32_t*, const uint8_t**, size_t*,
                           fm_image_info_t*) = fm_sheet_get_image;
  fm_status_t (*set_drawing_anchor)(fm_workbook_t*, size_t, uint32_t, const fm_drawing_anchor_t*) =
      fm_sheet_set_drawing_anchor;
  fm_status_t (*set_drawing_z_order)(fm_workbook_t*, size_t, uint32_t, size_t) = fm_sheet_set_drawing_z_order;
  fm_status_t (*snapshot_drawing)(const fm_workbook_t*, size_t, uint32_t, uint8_t**, size_t*) =
      fm_sheet_snapshot_drawing;
  fm_status_t (*restore_drawing)(fm_workbook_t*, size_t, const uint8_t*, size_t, uint32_t, uint32_t*) =
      fm_sheet_restore_drawing;
  fm_status_t (*reset_theme)(fm_workbook_t*) = fm_workbook_clear_theme;
  fm_status_t (*set_table_style)(fm_workbook_t*, const fm_table_style_t*) = fm_workbook_set_table_style;
  fm_status_t (*set_column_totals)(fm_workbook_t*, size_t, size_t, const char*, const char*, const char*) =
      fm_workbook_set_table_column_totals;
  (void)cells_in_range;
  (void)set_table_style;
  (void)set_column_totals;
  (void)format_value;
  (void)get_auto_filter;
  (void)set_table_auto_filter;
  (void)evaluate_auto_filter;
  (void)validate_value;
  (void)list_invalid_cells;
  (void)add_threaded_comment;
  (void)edit_threaded_comment;
  (void)person_at;
  (void)insert_image;
  (void)get_image;
  (void)set_drawing_anchor;
  (void)set_drawing_z_order;
  (void)snapshot_drawing;
  (void)restore_drawing;
  (void)reset_theme;
  (void)save_diagnostics;
  (void)read_diagnostics;
  (void)save_as;
  (void)get_cell_xf;
  (void)add_cell_xf;
  (void)get_cell_style_xf;
  (void)add_cell_style_xf;
  (void)get_view;
  (void)defined_name_at;
  (void)pivot_add_item_at;
  (void)add_styles_batch;
  (void)print_area_at;
  (void)parallel_recalc;
  (void)stats;
  return callback(1U, 0.0, 1U, 0) ? 0 : 1;
}
