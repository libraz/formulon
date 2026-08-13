// Shared validation and size accounting for dynamic-array spill footprints.
// The OOXML and XLSB readers carry the same inclusive rectangle shape but
// keep format-specific error codes and record state, so only the arithmetic
// belongs in this seam.

#ifndef FORMULON_IO_ARRAY_ANCHOR_BUDGET_H_
#define FORMULON_IO_ARRAY_ANCHOR_BUDGET_H_

#include <cstdint>
#include <limits>
#include <string>
#include <utility>

#include "sheet.h"
#include "utils/budget_charge.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"

namespace formulon {
namespace io {

/// Checks an inclusive dynamic-array anchor rectangle and returns its cell
/// count without narrowing to `size_t`. The caller charges the returned count
/// to its sheet-scoped `ResourceBudget` before reserving or walking cells.
inline Expected<std::uint64_t, Error> checked_array_anchor_cells(std::uint32_t row, std::uint32_t col,
                                                                 std::uint32_t last_row, std::uint32_t last_col,
                                                                 FormulonErrorCode error_code, const char* context) {
  if (row >= Sheet::kMaxRows || last_row >= Sheet::kMaxRows || col >= Sheet::kMaxCols || last_col >= Sheet::kMaxCols ||
      last_row < row || last_col < col) {
    return make_error(error_code, "array anchor rectangle out of grid", std::string(context));
  }

  const std::uint64_t row_count = static_cast<std::uint64_t>(last_row) - row + 1U;
  const std::uint64_t col_count = static_cast<std::uint64_t>(last_col) - col + 1U;
  if (row_count != 0U && col_count > std::numeric_limits<std::uint64_t>::max() / row_count) {
    return make_error(error_code, "array anchor cell count overflow", std::string(context));
  }
  return row_count * col_count;
}

/// Charges one validated anchor to the sheet-scoped budget. If the charge
/// fails, retain the caller's format/anchor context alongside the budget's
/// `used/requested/ceiling` diagnostics.
inline Expected<void, Error> consume_array_anchor_budget(ResourceBudget& budget, std::uint64_t cell_count,
                                                         std::string context) {
  return charge(budget, cell_count, std::move(context));
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_ARRAY_ANCHOR_BUDGET_H_
