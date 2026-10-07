// Dynamic-array spill footprints shared by the OOXML and XLSB readers:
// the anchor rectangle, its validation and size accounting, and the
// registration pass that turns each footprint into a spill region. Error
// codes and diagnostic context stay format-specific and are passed in.

#ifndef FORMULON_IO_ARRAY_ANCHOR_BUDGET_H_
#define FORMULON_IO_ARRAY_ANCHOR_BUDGET_H_

#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "sheet.h"
#include "utils/budget_charge.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "value.h"

namespace formulon {
namespace io {

/// One dynamic-array formula anchor and its spill footprint, as recorded by
/// `<f t="array" ref="F6:F8">` (OOXML) or a `BrtArrFmla` record's `RfX`
/// (XLSB). The 0-based inclusive rectangle `[row..last_row] x
/// [col..last_col]` is the footprint; `(row, col)` is the anchor. One-cell
/// anchors are kept too so dynamic-array metadata survives a rewrite.
struct ArrayAnchor {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;
};

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

/// Registers each anchor's footprint as a spill region on `sheet`: captures
/// the footprint's current cached values, blanks the non-anchor cells so
/// they block neither the commit's collision scan nor the anchor's re-spill
/// on recalc, then commits the region. Must run after the whole sheet has
/// been decoded. `reader` and `format` name the caller in error context.
inline Expected<void, Error> register_array_spills(Sheet& sheet, const std::vector<ArrayAnchor>& anchors,
                                                   FormulonErrorCode error_code, std::string_view reader,
                                                   std::string_view format) {
  // Validate and charge every footprint before the first reserve or cell
  // walk. This keeps a later malformed/over-budget anchor from arriving
  // after an earlier one has already started an attacker-sized operation.
  // The budget is local to this sheet so independent sheets cannot consume
  // one another's dynamic-array allowance.
  ResourceBudget budget(kMaxDynamicArrayCells, error_code);
  std::string cells_context("context=");
  cells_context.append(reader);
  cells_context.append(" array_anchor");
  for (const ArrayAnchor& a : anchors) {
    auto cells_or = checked_array_anchor_cells(a.row, a.col, a.last_row, a.last_col, error_code, cells_context.c_str());
    if (!cells_or) {
      return std::move(cells_or.error());
    }
    std::string context("context=");
    context.append(reader);
    context.append(" format=");
    context.append(format);
    context.append(" anchor_row=");
    context.append(std::to_string(a.row));
    context.append(" anchor_col=");
    context.append(std::to_string(a.col));
    context.append(" last_row=");
    context.append(std::to_string(a.last_row));
    context.append(" last_col=");
    context.append(std::to_string(a.last_col));
    auto charged = consume_array_anchor_budget(budget, cells_or.value(), std::move(context));
    if (!charged) {
      return std::move(charged.error());
    }
  }

  for (const ArrayAnchor& a : anchors) {
    const std::uint32_t rows = a.last_row - a.row + 1U;
    const std::uint32_t cols = a.last_col - a.col + 1U;
    const std::uint64_t cell_count = static_cast<std::uint64_t>(rows) * cols;
    std::vector<Value> values;
    values.reserve(static_cast<std::size_t>(cell_count));
    for (std::uint32_t r = a.row; r <= a.last_row; ++r) {
      for (std::uint32_t c = a.col; c <= a.last_col; ++c) {
        const Cell* cell = sheet.cell_at(r, c);
        values.push_back(cell != nullptr ? cell->cached_value : Value::blank());
      }
    }
    // Blank the non-anchor cells: their values now live in the region as
    // phantoms.
    for (std::uint32_t r = a.row; r <= a.last_row; ++r) {
      for (std::uint32_t c = a.col; c <= a.last_col; ++c) {
        if (r == a.row && c == a.col) {
          continue;
        }
        sheet.set_cell_cached_value(r, c, Value::blank());
      }
    }
    sheet.commit_spill(a.row, a.col, rows, cols, std::move(values));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_ARRAY_ANCHOR_BUDGET_H_
