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
  fm_parallel_recalc_stats stats = {0};
  fm_status_t (*parallel_recalc)(fm_workbook_t*, uint32_t, fm_parallel_recalc_stats*) = fm_workbook_recalc_parallel;
  fm_status_t (*save_diagnostics)(const fm_workbook_t*, int32_t, uint8_t**, size_t*, fm_save_diagnostics_t*) =
      fm_workbook_save_with_diagnostics;
  fm_status_t (*read_diagnostics)(const fm_workbook_t*, fm_read_diagnostics_t*) = fm_workbook_read_diagnostics;
  fm_status_t (*save_as)(const fm_workbook_t*, int32_t, uint8_t**, size_t*) = fm_workbook_save_as;
  // Written out in full because these are the entry points a third-party C
  // consumer binds by signature: the XF record crosses by value, so its width
  // is part of the calling convention, and the sheet-view record is written
  // through a caller-supplied pointer.
  fm_status_t (*get_cell_xf)(fm_workbook_t*, uint32_t, fm_cell_xf*) = fm_styles_get_cell_xf;
  fm_status_t (*add_cell_xf)(fm_workbook_t*, fm_cell_xf, uint32_t*) = fm_styles_add_cell_xf;
  fm_status_t (*get_cell_style_xf)(fm_workbook_t*, uint32_t, fm_cell_xf*) = fm_styles_get_cell_style_xf;
  fm_status_t (*add_cell_style_xf)(fm_workbook_t*, fm_cell_xf, uint32_t*) = fm_styles_add_cell_style_xf;
  fm_status_t (*get_view)(const fm_workbook_t*, size_t, fm_sheet_view_t*) = fm_sheet_get_view;
  fm_status_t (*defined_name_at)(const fm_workbook_t*, size_t, const char**, const char**, int32_t*) =
      fm_workbook_defined_name_at;
  fm_status_t (*pivot_add_item_at)(fm_workbook_t*, size_t, size_t, size_t, uint32_t, int32_t) =
      fm_workbook_pivot_field_add_item_at;
  // The two entry points whose records no binding marshals, written out in
  // full so a parameter added to either is a compile error here rather than a
  // silently-wrong C caller.
  fm_status_t (*add_styles_batch)(fm_workbook_t*, const fm_styles_batch*) = fm_styles_add_batch;
  fm_status_t (*print_area_at)(const fm_pagination_t*, size_t, fm_print_range_t*) = fm_pagination_print_area_at;
  // Both counter structs must lay out identically on native and wasm32, so a
  // binding's hand-written offsets cannot drift from the compiler's. Five
  // 4-byte counters, no padding, in either target.
  _Static_assert(sizeof(fm_read_diagnostics_t) == 5 * sizeof(uint32_t), "fm_read_diagnostics_t must be packed");
  _Static_assert(sizeof(fm_save_diagnostics_t) == 5 * sizeof(uint32_t), "fm_save_diagnostics_t must be packed");
  // Pinned from C as well as C++: a C consumer built against a narrower
  // definition of either record is miscompiled rather than diagnosed.
  _Static_assert(sizeof(fm_cell_xf) == 128, "fm_cell_xf ABI layout changed");
  _Static_assert(sizeof(fm_row_layout_t) == 40, "fm_row_layout_t ABI layout changed");
  _Static_assert(sizeof(fm_sheet_view_t) == (sizeof(void*) == 4 ? 44 : 48), "fm_sheet_view_t ABI layout changed");
  // The three records a C consumer marshals without help from any binding:
  // `fm_value_t` comes back from every cell read, `fm_print_range_t` from
  // pagination, and `fm_styles_batch` is passed in by pointer with fifteen
  // pointer-width slots. The first two are target-independent; the batch is
  // pinned in terms of `sizeof(void*)` because its `size_t` counts are not.
  _Static_assert(sizeof(fm_value_t) == 16, "fm_value_t ABI layout changed");
  _Static_assert(offsetof(fm_value_t, u) == 8, "fm_value_t.u offset changed");
  _Static_assert(sizeof(fm_print_range_t) == 16, "fm_print_range_t ABI layout changed");
  _Static_assert(offsetof(fm_print_range_t, last_col) == 12, "fm_print_range_t.last_col offset changed");
  _Static_assert(sizeof(fm_styles_batch) == 15 * sizeof(void*), "fm_styles_batch ABI layout changed");
  _Static_assert(offsetof(fm_styles_batch, num_fmt_ids) == 14 * sizeof(void*),
                 "fm_styles_batch.num_fmt_ids offset changed");
  // Records the geometry, display and pagination-detail entry points write
  // through caller pointers. All but the width model are target-independent.
  _Static_assert(sizeof(fm_sheet_format_defaults) == 32, "fm_sheet_format_defaults ABI layout changed");
  _Static_assert(sizeof(fm_rect_pt) == 32, "fm_rect_pt ABI layout changed");
  _Static_assert(sizeof(fm_width_model) == (sizeof(void*) == 4 ? 40 : 48), "fm_width_model ABI layout changed");
  _Static_assert(sizeof(fm_paper_info) == 24, "fm_paper_info ABI layout changed");
  _Static_assert(sizeof(fm_margins_pt) == 48, "fm_margins_pt ABI layout changed");
  _Static_assert(sizeof(fm_print_titles) == 24, "fm_print_titles ABI layout changed");
  _Static_assert(sizeof(fm_page_layout) == 56, "fm_page_layout ABI layout changed");
  fm_status_t (*cells_in_range)(const fm_workbook_t*, size_t, uint32_t, uint32_t, uint32_t, uint32_t, uint64_t,
                                uint32_t, fm_cell_range_t**) = fm_sheet_cells_in_range;
  fm_status_t (*format_value)(const fm_workbook_t*, const fm_value_t*, const char*, const char**, int32_t*) =
      fm_workbook_format_value;
  // Records the AutoFilter, validation and threaded-comment entry points read
  // and write. The pointer-bearing ones are pinned per target.
  _Static_assert(sizeof(fm_date_group_item) == 8, "fm_date_group_item ABI layout changed");
  _Static_assert(sizeof(fm_filter_column) == (sizeof(void*) == 4 ? 152 : 192), "fm_filter_column ABI layout changed");
  _Static_assert(sizeof(fm_sort_condition) == (sizeof(void*) == 4 ? 48 : 56), "fm_sort_condition ABI layout changed");
  _Static_assert(sizeof(fm_auto_filter) == (sizeof(void*) == 4 ? 64 : 80), "fm_auto_filter ABI layout changed");
  _Static_assert(sizeof(fm_validation_outcome) == 16, "fm_validation_outcome ABI layout changed");
  _Static_assert(sizeof(fm_mention) == (sizeof(void*) == 4 ? 16 : 24), "fm_mention ABI layout changed");
  _Static_assert(sizeof(fm_threaded_comment) == (sizeof(void*) == 4 ? 40 : 72),
                 "fm_threaded_comment ABI layout changed");
  _Static_assert(sizeof(fm_person) == 4 * sizeof(void*), "fm_person ABI layout changed");
  fm_status_t (*get_auto_filter)(const fm_workbook_t*, size_t, fm_auto_filter*, int32_t*) = fm_sheet_get_auto_filter;
  fm_status_t (*set_table_auto_filter)(fm_workbook_t*, size_t, const fm_auto_filter*) = fm_table_set_auto_filter;
  fm_status_t (*evaluate_auto_filter)(const fm_workbook_t*, size_t, uint8_t*, size_t, size_t*, uint32_t*) =
      fm_sheet_evaluate_auto_filter;
  fm_status_t (*validate_value)(const fm_workbook_t*, size_t, uint32_t, uint32_t, const fm_value_t*,
                                fm_validation_outcome*) = fm_sheet_validate_value;
  fm_status_t (*list_invalid_cells)(const fm_workbook_t*, size_t, uint64_t, uint32_t, fm_cell_range_t**) =
      fm_sheet_list_invalid_cells;
  fm_status_t (*add_threaded_comment)(fm_workbook_t*, size_t, const fm_threaded_comment*) =
      fm_sheet_add_threaded_comment;
  fm_status_t (*edit_threaded_comment)(fm_workbook_t*, size_t, const char*, const char*, const fm_mention*, uint32_t) =
      fm_sheet_edit_threaded_comment;
  fm_status_t (*person_at)(const fm_workbook_t*, size_t, fm_person*) = fm_workbook_person_at;
  (void)cells_in_range;
  (void)format_value;
  (void)get_auto_filter;
  (void)set_table_auto_filter;
  (void)evaluate_auto_filter;
  (void)validate_value;
  (void)list_invalid_cells;
  (void)add_threaded_comment;
  (void)edit_threaded_comment;
  (void)person_at;
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
