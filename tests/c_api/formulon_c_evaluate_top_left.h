#pragma once

#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"

// Evaluates `formula` at `(row, col)` through `fm_sheet_evaluate_formula_array`
// and writes the result's top-left cell to `*out`, the scalar a binding's
// `evaluateFormula` folds a 1 x 1 result to. A text view stays valid until the
// next call on the same thread.
inline fm_status_t evaluate_top_left(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                     std::uint32_t col, const char* formula, fm_value_t* out) {
  struct Held {
    fm_value_array_t* array = nullptr;
    ~Held() { fm_value_array_destroy(array); }
  };
  thread_local Held held;
  fm_value_array_t* array = nullptr;
  const fm_status_t rc = fm_sheet_evaluate_formula_array(wb, sheet_index, row, col, formula, &array);
  if (rc != FM_STATUS_OK) {
    return rc;
  }
  fm_value_array_destroy(held.array);
  held.array = array;
  return fm_value_array_get_value_at(array, 0, out);
}
