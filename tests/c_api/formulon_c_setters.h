#pragma once

#include <cstdint>

#include "c_api/formulon_c.h"

// Positional wrappers over the C setters that take a record, for tests that
// set every field.

// Sets the defined name `name` to `formula` in scope `local_sheet_id` (`-1`
// for workbook scope), with no comment and not hidden.
inline fm_status_t SetDefinedName(fm_workbook_t* wb, const char* name, const char* formula,
                                  std::int32_t local_sheet_id = -1) {
  const fm_defined_name_input_t input{name, formula, local_sheet_id, nullptr, 0};
  return fm_workbook_set_defined_name(wb, &input);
}

// Sets all three iterative-calculation settings.
inline fm_status_t SetIterative(fm_workbook_t* wb, std::int32_t enabled, std::int32_t max_iterations,
                                double max_change) {
  const fm_iterative_settings_patch_t patch{enabled, max_iterations, max_change};
  return fm_workbook_set_iterative(wb, &patch);
}
