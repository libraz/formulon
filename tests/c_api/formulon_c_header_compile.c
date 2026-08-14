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
  fm_status_t (*save_diagnostics)(const fm_workbook_t*, int32_t, uint8_t**, size_t*, size_t*, size_t*) =
      fm_workbook_save_ex_with_diagnostics;
  fm_status_t (*read_diagnostics)(const fm_workbook_t*, size_t*, size_t*, size_t*) =
      fm_workbook_xlsb_read_diagnostics_ex;
  (void)save_diagnostics;
  (void)read_diagnostics;
  (void)parallel_recalc;
  (void)stats;
  return callback(1U, 0.0, 1U, 0) ? 0 : 1;
}
