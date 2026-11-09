#pragma once

#include <cstddef>

#include "c_api/formulon_c.h"

// An out-index slot for C ABI adds whose allocated index the test does not read.
inline std::size_t* ignored_index() {
  static std::size_t slot = 0;
  return &slot;
}

// An out-id slot for C ABI adds that return an id the test does not read.
inline const char** ignored_id() {
  static const char* slot = nullptr;
  return &slot;
}

// The `insert_before` that appends: the workbook's current sheet count.
inline std::size_t append_position(const fm_workbook_t* wb) {
  std::size_t count = 0;
  fm_workbook_get_sheet_count(wb, &count);
  return count;
}
