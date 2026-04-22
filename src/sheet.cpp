// Copyright 2026 libraz. Licensed under the MIT License.
//
// Out-of-line implementation of the row-sparse, column-dense cell store
// owned by `Sheet`. See `sheet.h` for the storage-layer contract.

#include "sheet.h"

#include <cassert>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "value.h"

namespace formulon {

namespace {

// Grows `row_cells` so that index `col` is addressable, padding with
// default-constructed cells, and returns a reference to the slot at `col`.
Cell& EnsureSlot(std::vector<Cell>& row_cells, std::uint32_t col) {
  const std::size_t needed = static_cast<std::size_t>(col) + 1U;
  if (row_cells.size() < needed) {
    row_cells.resize(needed);
  }
  return row_cells[col];
}

}  // namespace

void Sheet::set_cell_value(std::uint32_t row, std::uint32_t col, Value v) {
  // Bounds checks are advisory: callers above this layer (parser, OOXML
  // reader) own coordinate validation. A debug assert catches programming
  // errors without imposing a release-mode branch.
  assert(row < kMaxRows && col < kMaxCols);

  std::vector<Cell>& row_cells = rows_[row];
  Cell& slot = EnsureSlot(row_cells, col);
  slot.formula_text.clear();
  slot.cached_value = v;
}

void Sheet::set_cell_formula(std::uint32_t row, std::uint32_t col, std::string formula) {
  assert(row < kMaxRows && col < kMaxCols);

  std::vector<Cell>& row_cells = rows_[row];
  Cell& slot = EnsureSlot(row_cells, col);
  slot.formula_text = std::move(formula);
  slot.cached_value = Value::blank();
}

const Cell* Sheet::cell_at(std::uint32_t row, std::uint32_t col) const noexcept {
  const auto it = rows_.find(row);
  if (it == rows_.end()) {
    return nullptr;
  }
  const std::vector<Cell>& row_cells = it->second;
  if (col >= row_cells.size()) {
    return nullptr;
  }
  return &row_cells[col];
}

bool Sheet::has_cell(std::uint32_t row, std::uint32_t col) const noexcept { return cell_at(row, col) != nullptr; }

std::size_t Sheet::cell_count() const noexcept {
  std::size_t total = 0;
  for (const auto& kv : rows_) {
    total += kv.second.size();
  }
  return total;
}

}  // namespace formulon
