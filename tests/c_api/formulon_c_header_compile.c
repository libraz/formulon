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
  _Static_assert(sizeof(fm_cell_xf) == 88, "fm_cell_xf ABI layout changed");
  _Static_assert(sizeof(fm_sheet_view_t) == (sizeof(void*) == 4 ? 40 : 48), "fm_sheet_view_t ABI layout changed");
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
