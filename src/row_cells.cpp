//
// `RowCells` column insertion and removal: the sorted column-key array and its
// parallel cell array move together.

#include "row_cells.h"

#include <algorithm>
#include <cassert>
#include <cstddef>
#include <cstdint>
#include <utility>

#include "cell.h"
#include "utils/a1_ref.h"

namespace formulon {

Cell& RowCells::ensure(std::uint32_t col) {
  assert(col < a1::kMaxCols);
  // Right-append fast path: rows are mostly written left to right.
  if (cols_.empty() || col > cols_.back()) {
    cols_.push_back(static_cast<std::uint16_t>(col));
    cells_.emplace_back();
    return cells_.back();
  }
  if (Cell* existing = find(col); existing != nullptr) {
    return *existing;
  }
  const std::size_t index = lower_index(col);
  cols_.insert(cols_.begin() + static_cast<std::ptrdiff_t>(index), static_cast<std::uint16_t>(col));
  return *cells_.emplace(cells_.begin() + static_cast<std::ptrdiff_t>(index));
}

void RowCells::shift_cols_from(std::uint32_t col, std::uint32_t count) {
  const std::size_t first = lower_index(col);
  std::size_t kept = cols_.size();
  for (std::size_t i = first; i < cols_.size(); ++i) {
    const std::uint64_t shifted = static_cast<std::uint64_t>(cols_[i]) + count;
    if (shifted >= a1::kMaxCols) {
      kept = i;
      break;
    }
    cols_[i] = static_cast<std::uint16_t>(shifted);
  }
  cols_.resize(kept);
  cells_.erase(cells_.begin() + static_cast<std::ptrdiff_t>(kept), cells_.end());
}

void RowCells::erase(std::uint32_t col) {
  const std::size_t index = index_of(col);
  if (index == kNpos) {
    return;
  }
  cols_.erase(cols_.begin() + static_cast<std::ptrdiff_t>(index));
  cells_.erase(cells_.begin() + static_cast<std::ptrdiff_t>(index));
}

void RowCells::erase_cols(std::uint32_t col, std::uint32_t count) {
  const std::uint64_t band_end = static_cast<std::uint64_t>(col) + count;
  const std::size_t first = lower_index(col);
  const std::size_t last = std::max(first, lower_index(band_end));
  cols_.erase(cols_.begin() + static_cast<std::ptrdiff_t>(first), cols_.begin() + static_cast<std::ptrdiff_t>(last));
  cells_.erase(cells_.begin() + static_cast<std::ptrdiff_t>(first), cells_.begin() + static_cast<std::ptrdiff_t>(last));
  for (std::size_t i = first; i < cols_.size(); ++i) {
    cols_[i] = static_cast<std::uint16_t>(cols_[i] - count);
  }
}

void RowCells::erase_in(std::uint32_t first, std::uint64_t end) {
  const std::size_t lo = lower_index(first);
  const std::size_t hi = std::max(lo, lower_index(end));
  cols_.erase(cols_.begin() + static_cast<std::ptrdiff_t>(lo), cols_.begin() + static_cast<std::ptrdiff_t>(hi));
  cells_.erase(cells_.begin() + static_cast<std::ptrdiff_t>(lo), cells_.begin() + static_cast<std::ptrdiff_t>(hi));
}

void RowCells::move_in_to(std::uint32_t first, std::uint64_t end, RowCells& dest) {
  dest.erase_in(first, end);
  const std::size_t lo = lower_index(first);
  const std::size_t hi = std::max(lo, lower_index(end));
  for (std::size_t i = lo; i < hi; ++i) {
    dest.ensure(cols_[i]) = std::move(cells_[i]);
  }
  erase_in(first, end);
}

}  // namespace formulon
