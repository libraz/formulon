//
// Rectangle scans over the sparse cell store: populated extent, ranged
// cell enumeration and bulk range reads. All of them read `rows_` and the
// spill table under the sheet's one lock; see `sheet.h` for the contracts.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <mutex>
#include <optional>
#include <vector>

#include "cell.h"
#include "sheet.h"
#include "sheet_spill_table.h"
#include "utils/arena.h"
#include "utils/index_sort.h"
#include "value.h"

namespace formulon {

namespace {

/// A stored cell is content when it has a formula or a non-blank value.
bool has_cell_content(const Cell& cell) noexcept {
  return !cell.formula_text.empty() || !cell.cached_value.is_blank();
}

// Calls `fn(row, cells)` for each stored row in [row_begin, row_end), walking
// whichever of the stored rows and the rectangle's rows is the smaller set.
template <typename Rows, typename Fn>
void for_each_stored_row_in(const Rows& rows, std::uint64_t row_begin, std::uint64_t row_end, Fn&& fn) {
  if (row_end <= row_begin) {
    return;
  }
  if (static_cast<std::uint64_t>(rows.size()) <= row_end - row_begin) {
    for (const auto& [row, cells] : rows) {
      if (static_cast<std::uint64_t>(row) >= row_begin && static_cast<std::uint64_t>(row) < row_end) {
        fn(row, cells);
      }
    }
    return;
  }
  for (std::uint64_t row = row_begin; row < row_end; ++row) {
    const auto it = rows.find(static_cast<std::uint32_t>(row));
    if (it != rows.end()) {
      fn(static_cast<std::uint32_t>(row), it->second);
    }
  }
}

}  // namespace

std::size_t Sheet::materialised_cells_in_rect_locked(std::uint32_t row_begin, std::uint64_t row_end,
                                                     std::uint32_t col_begin, std::uint64_t col_end) const noexcept {
  std::size_t total = 0;
  for_each_stored_row_in(rows_, row_begin, row_end,
                         [&](std::uint32_t, const RowCells& cells) { total += cells.count_in(col_begin, col_end); });
  return total;
}

std::optional<Sheet::PopulatedExtent> Sheet::populated_extent(std::uint32_t first_row, std::uint32_t first_col,
                                                              std::uint32_t last_row,
                                                              std::uint32_t last_col) const noexcept {
  if (!rect_in_grid(first_row, first_col, last_row, last_col)) {
    return std::nullopt;
  }

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  PopulatedExtent extent;
  bool any = false;
  const auto include = [&](std::uint32_t row, std::uint32_t col) {
    if (!any) {
      extent.first_row = extent.last_row = row;
      extent.first_col = extent.last_col = col;
      any = true;
      return;
    }
    extent.first_row = std::min(extent.first_row, row);
    extent.first_col = std::min(extent.first_col, col);
    extent.last_row = std::max(extent.last_row, row);
    extent.last_col = std::max(extent.last_col, col);
  };

  // Only stored cells in the column span are visited; a stored cell with no
  // formula and a blank value is not content.
  const auto scan_row = [&](std::uint32_t row, const RowCells& cells) {
    for (const auto& [col, cell] : cells.entries_in(first_col, static_cast<std::uint64_t>(last_col) + 1U)) {
      if (has_cell_content(cell)) {
        include(row, col);
      }
    }
  };
  for_each_stored_row_in(rows_, first_row, static_cast<std::uint64_t>(last_row) + 1U, scan_row);

  // A committed spill is represented by one rectangle, so fold its clipped
  // intersection into the result without walking its phantom payload.
  if (spill_table_ != nullptr) {
    for (const auto& [unused, region] : spill_table_->by_anchor) {
      (void)unused;
      const std::uint64_t region_last_row = static_cast<std::uint64_t>(region.anchor_row) + region.rows - 1U;
      const std::uint64_t region_last_col = static_cast<std::uint64_t>(region.anchor_col) + region.cols - 1U;
      if (region.anchor_row > last_row || region.anchor_col > last_col || region_last_row < first_row ||
          region_last_col < first_col) {
        continue;
      }
      const std::uint32_t clipped_first_row = std::max(first_row, region.anchor_row);
      const std::uint32_t clipped_first_col = std::max(first_col, region.anchor_col);
      const std::uint32_t clipped_last_row = std::min(last_row, static_cast<std::uint32_t>(region_last_row));
      const std::uint32_t clipped_last_col = std::min(last_col, static_cast<std::uint32_t>(region_last_col));
      include(clipped_first_row, clipped_first_col);
      include(clipped_last_row, clipped_last_col);
    }
  }

  if (!any) {
    return std::nullopt;
  }
  return extent;
}

std::uint64_t Sheet::cells_in_range(std::uint32_t first_row, std::uint32_t first_col, std::uint32_t last_row,
                                    std::uint32_t last_col, std::uint64_t cursor, std::uint32_t limit,
                                    void (*visit)(const RangeCell& cell, void* ctx), void* ctx) const {
  constexpr std::uint64_t kStride = kMaxCols;
  if (!rect_in_grid(first_row, first_col, last_row, last_col) || cursor >= kStride * kMaxRows) {
    return kCellCursorEnd;
  }
  auto start_row = static_cast<std::uint32_t>(cursor / kStride);
  auto start_col = static_cast<std::uint32_t>(cursor % kStride);
  if (start_row < first_row) {
    start_row = first_row;
    start_col = first_col;
  }
  start_col = std::max(start_col, first_col);
  if (start_col > last_col) {
    ++start_row;
    start_col = first_col;
  }
  if (start_row > last_row) {
    return kCellCursorEnd;
  }

  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  std::vector<const SpillRegion*> regions;
  if (spill_table_ != nullptr) {
    for (const auto& [unused, region] : spill_table_->by_anchor) {
      (void)unused;
      if (RectIntersectsSpan(region, start_row, static_cast<std::uint64_t>(last_row) + 1U, first_col,
                             static_cast<std::uint64_t>(last_col) + 1U)) {
        regions.push_back(&region);
      }
    }
  }

  // Phantom-free rectangles taller than the stored row count walk the sorted
  // stored rows instead of probing every row, as `populated_extent` does.
  std::vector<std::uint32_t> stored_rows;
  const bool walk_stored =
      regions.empty() && static_cast<std::uint64_t>(rows_.size()) <= static_cast<std::uint64_t>(last_row - start_row);
  if (walk_stored) {
    for (const auto& entry : rows_) {
      if (entry.first >= start_row && entry.first <= last_row) {
        stored_rows.push_back(entry.first);
      }
    }
    sort_ascending(stored_rows);
  }

  std::uint32_t emitted = 0;
  std::vector<std::uint32_t> cols;
  const auto scan_row = [&](std::uint32_t row) -> std::uint64_t {
    const std::uint32_t col_begin = row == start_row ? start_col : first_col;
    cols.clear();
    const auto row_it = rows_.find(row);
    const RowCells* cells = row_it == rows_.end() ? nullptr : &row_it->second;
    if (cells != nullptr) {
      for (const auto& [col, cell] : cells->entries_in(col_begin, static_cast<std::uint64_t>(last_col) + 1U)) {
        if (has_cell_content(cell)) {
          cols.push_back(col);
        }
      }
    }
    bool phantom_cols = false;
    for (const SpillRegion* region : regions) {
      if (row < region->anchor_row || row - region->anchor_row >= region->rows) {
        continue;
      }
      const std::uint32_t span_first = std::max(col_begin, region->anchor_col);
      const std::uint32_t span_last = std::min(last_col, region->anchor_col + region->cols - 1U);
      for (std::uint32_t col = span_first; col <= span_last; ++col) {
        if (row != region->anchor_row || col != region->anchor_col) {
          cols.push_back(col);
          phantom_cols = true;
        }
      }
    }
    if (phantom_cols) {
      sort_ascending(cols);
      cols.erase(std::unique(cols.begin(), cols.end()), cols.end());
    }
    for (const std::uint32_t col : cols) {
      if (emitted == limit) {
        return static_cast<std::uint64_t>(row) * kStride + col;
      }
      RangeCell out;
      out.row = row;
      out.col = col;
      // Same precedence as `read_formula_cell`: formula cell, then phantom,
      // then the stored literal.
      const Cell* cell = cells == nullptr ? nullptr : cells->find(col);
      if (cell != nullptr && !cell->formula_text.empty()) {
        out.formula_text = cell->formula_text;
        out.value = cell->cached_value;
      } else if (const SpillRegion* covering = spill_region_covering_locked(row, col); covering != nullptr) {
        out.value = covering->cells[static_cast<std::size_t>(row - covering->anchor_row) * covering->cols +
                                    (col - covering->anchor_col)];
      } else if (cell != nullptr) {
        out.value = cell->cached_value;
      }
      visit(out, ctx);
      ++emitted;
    }
    return kCellCursorEnd;
  };

  if (walk_stored) {
    for (const std::uint32_t row : stored_rows) {
      if (const std::uint64_t next = scan_row(row); next != kCellCursorEnd) {
        return next;
      }
    }
  } else {
    for (std::uint32_t row = start_row; row <= last_row; ++row) {
      if (const std::uint64_t next = scan_row(row); next != kCellCursorEnd) {
        return next;
      }
    }
  }
  return kCellCursorEnd;
}

void Sheet::read_range(std::uint32_t first_row, std::uint32_t last_row, std::uint32_t first_col, std::uint32_t last_col,
                       Arena& text_arena, std::vector<Value>& out, std::vector<std::size_t>& formula_indices) const {
  if (first_row > last_row || first_col > last_col || last_row >= kMaxRows || last_col >= kMaxCols) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);

  // Narrow the spill table to the regions this rectangle can actually reach,
  // once, so the per-coordinate test scans only those instead of paying a
  // banded-index lookup per cell.
  std::vector<const SpillRegion*> regions;
  if (spill_table_ != nullptr) {
    const std::uint64_t row_end = static_cast<std::uint64_t>(last_row) + 1U;
    const std::uint64_t col_end = static_cast<std::uint64_t>(last_col) + 1U;
    for (const auto& entry : spill_table_->by_anchor) {
      if (RectIntersectsSpan(entry.second, first_row, row_end, first_col, col_end)) {
        regions.push_back(&entry.second);
      }
    }
  }
  const auto covering = [&regions](std::uint32_t row, std::uint32_t col) noexcept -> const SpillRegion* {
    for (const SpillRegion* region : regions) {
      if (!RectIntersectsSpan(*region, row, static_cast<std::uint64_t>(row) + 1U, col,
                              static_cast<std::uint64_t>(col) + 1U)) {
        continue;
      }
      // The anchor reads as its own stored cell, matching the scalar path.
      return (row == region->anchor_row && col == region->anchor_col) ? nullptr : region;
    }
    return nullptr;
  };

  for (std::uint32_t row = first_row; row <= last_row; ++row) {
    const auto row_it = rows_.find(row);
    RowCells::const_iterator next;
    RowCells::const_iterator stop;
    if (row_it != rows_.end()) {
      const RowCells::EntryRange entries = row_it->second.entries_from(first_col);
      next = entries.begin();
      stop = entries.end();
    }
    for (std::uint32_t col = first_col; col <= last_col; ++col) {
      // The stored entries are ascending, so one cursor per row follows `col`.
      // It starts at `first_col` and never needs the range's end.
      const Cell* cell = nullptr;
      if (next != stop && (*next).col == col) {
        cell = &(*next).cell;
        ++next;
      }
      // Order matters and mirrors the scalar path: a stored formula wins over
      // a covering spill region, which in turn wins over a stored literal.
      if (cell != nullptr && !cell->formula_text.empty()) {
        formula_indices.push_back(out.size());
        out.push_back(adopt_text_into(text_arena, cell->cached_value));
        continue;
      }
      if (const SpillRegion* region = covering(row, col); region != nullptr) {
        const std::size_t index = static_cast<std::size_t>(row - region->anchor_row) * region->cols +
                                  static_cast<std::size_t>(col - region->anchor_col);
        out.push_back(adopt_text_into(text_arena, region->cells[index]));
        continue;
      }
      out.push_back(cell != nullptr ? adopt_text_into(text_arena, cell->cached_value) : Value::blank());
    }
  }
}

}  // namespace formulon
