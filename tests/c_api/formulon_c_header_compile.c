//
// Compile-only C11 contract probe for the public C ABI header.

#include "c_api/formulon_c.h"

static bool ContinueIteration(uint32_t iteration, double max_residual, uint32_t max_iterations, void* user_data) {
  (void)iteration;
  (void)max_residual;
  (void)max_iterations;
  (void)user_data;
  return true;
}

int main(void) {
  fm_iterative_progress_cb callback = ContinueIteration;
  fm_status_t (*save_diagnostics)(const fm_workbook_t*, fm_workbook_format_t, uint8_t**, size_t*, size_t*, size_t*) =
      fm_workbook_save_ex_with_diagnostics;
  fm_status_t (*read_diagnostics)(const fm_workbook_t*, size_t*, size_t*, size_t*) =
      fm_workbook_xlsb_read_diagnostics_ex;
  (void)save_diagnostics;
  (void)read_diagnostics;
  return callback(1U, 0.0, 1U, 0) ? 0 : 1;
}
