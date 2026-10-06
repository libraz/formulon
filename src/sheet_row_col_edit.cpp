//
// Row and column insert/delete for `Sheet`: the cell-store shift and the
// coordinate remap of every sheet-attached structure that follows it.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "utils/index_sort.h"

namespace formulon {

namespace {

// Shifts every entry in `items` whose anchor lies on or past `index` by the
// insert / delete rule encoded in `count` and `is_delete`. Entries whose
// anchor falls inside the deleted interval, or whose insert would push it
// past `bound`, are removed. `field` names the axis coordinate, so one
// instantiation per anchor type serves both axes.
template <typename Anchor>
void ShiftAnchored(std::vector<Anchor>& items, std::uint32_t Anchor::*field, std::uint32_t index, std::uint32_t count,
                   bool is_delete, std::uint32_t bound) {
  std::vector<Anchor> retained;
  retained.reserve(items.size());
  for (Anchor& item : items) {
    if (item.*field < index) {
      retained.push_back(std::move(item));
      continue;
    }
    if (is_delete) {
      if (item.*field < index + count) {
        continue;  // Anchor inside deleted interval; drop the entry.
      }
      item.*field -= count;
    } else {
      // Insert. Anchors at or past `index` shift forward; entries pushed
      // past the sheet bound are dropped.
      const std::uint64_t shifted = static_cast<std::uint64_t>(item.*field) + count;
      if (shifted >= bound) {
        continue;
      }
      item.*field = static_cast<std::uint32_t>(shifted);
    }
    retained.push_back(std::move(item));
  }
  items = std::move(retained);
}

template <typename Anchor>
void ShiftRowAnchored(std::vector<Anchor>& items, std::uint32_t row, std::uint32_t count, bool is_delete) {
  ShiftAnchored(items, &Anchor::row, row, count, is_delete, Sheet::kMaxRows);
}

template <typename Anchor>
void ShiftColAnchored(std::vector<Anchor>& items, std::uint32_t col, std::uint32_t count, bool is_delete) {
  ShiftAnchored(items, &Anchor::col, col, count, is_delete, Sheet::kMaxCols);
}

// Rectangular merge / validation range shifter along one axis.
// Spans that fall entirely inside the deleted interval are dropped;
// spans that straddle the deletion are clamped so the survivors stay
// contiguous (Excel's "shrink the merge" behaviour). Inserts that would
// push `last` past `bound - 1` clamp to the sheet bound.
void ShiftSpan(std::uint32_t& first, std::uint32_t& last, std::uint32_t index, std::uint32_t count, bool is_delete,
               std::uint32_t bound, bool* out_drop) {
  *out_drop = false;
  if (is_delete) {
    const std::uint32_t del_end = index + count;  // exclusive
    // Both endpoints below the deletion: unchanged.
    if (last < index) {
      return;
    }
    // Both endpoints inside the deletion: drop the range entirely.
    if (first >= index && last < del_end) {
      *out_drop = true;
      return;
    }
    // Split shifts depending on which endpoints fall inside.
    if (first < index && last >= index && last < del_end) {
      // Trailing endpoint inside the deletion; clamp to index-1.
      last = index - 1U;
      return;
    }
    if (first >= index && first < del_end && last >= del_end) {
      // Leading endpoint inside the deletion; clamp to the line after the
      // deletion (which after the shift becomes `index`).
      first = index;
      last -= count;
      return;
    }
    if (first < index && last >= del_end) {
      // Range straddles the entire deletion: shrink by `count`. The
      // leading endpoint stays put; the trailing one shifts back.
      last -= count;
      return;
    }
    // Both endpoints past the deletion: shift back.
    first -= count;
    last -= count;
    return;
  }
  // Insert. Endpoints at or past `index` shift forward; clamp to bound.
  if (last < index) {
    return;  // Both endpoints below the insert; unchanged.
  }
  // A range pushed wholly beyond the sheet has no surviving cells; drop it rather than clamp to a singleton.
  if (first >= index && static_cast<std::uint64_t>(first) + count >= bound) {
    *out_drop = true;
    return;
  }
  auto shift_one = [count, bound](std::uint32_t value) -> std::uint32_t {
    const std::uint64_t shifted = static_cast<std::uint64_t>(value) + count;
    if (shifted >= bound) {
      return bound - 1U;
    }
    return static_cast<std::uint32_t>(shifted);
  };
  if (first >= index) {
    first = shift_one(first);
  }
  last = shift_one(last);
}

void ShiftRowRange(MergeRange& range, std::uint32_t row, std::uint32_t count, bool is_delete, bool* out_drop) {
  ShiftSpan(range.first_row, range.last_row, row, count, is_delete, Sheet::kMaxRows, out_drop);
}

void ShiftColRange(MergeRange& range, std::uint32_t col, std::uint32_t count, bool is_delete, bool* out_drop) {
  ShiftSpan(range.first_col, range.last_col, col, count, is_delete, Sheet::kMaxCols, out_drop);
}

void ShiftRangeList(std::vector<MergeRange>& ranges, std::uint32_t index, std::uint32_t count, bool is_delete,
                    bool row_axis) {
  std::vector<MergeRange> retained;
  retained.reserve(ranges.size());
  for (MergeRange& range : ranges) {
    bool drop = false;
    if (row_axis) {
      ShiftRowRange(range, index, count, is_delete, &drop);
    } else {
      ShiftColRange(range, index, count, is_delete, &drop);
    }
    if (drop) {
      continue;
    }
    retained.push_back(range);
  }
  ranges = std::move(retained);
}

// Hyperlinks use the same inclusive rectangle semantics as merge ranges.
// Keeping this conversion next to ShiftRangeList is intentional: row/column
// insertions and deletions must apply the exact same shift, shrink, clamp and
// drop rules to both metadata shapes.
void ShiftHyperlinkList(std::vector<Hyperlink>& hyperlinks, std::uint32_t index, std::uint32_t count, bool is_delete,
                        bool row_axis) {
  std::vector<Hyperlink> retained;
  retained.reserve(hyperlinks.size());
  for (Hyperlink& hyperlink : hyperlinks) {
    MergeRange range{hyperlink.row, hyperlink.col, hyperlink.last_row, hyperlink.last_col};
    bool drop = false;
    if (row_axis) {
      ShiftRowRange(range, index, count, is_delete, &drop);
    } else {
      ShiftColRange(range, index, count, is_delete, &drop);
    }
    if (drop) {
      continue;
    }
    hyperlink.row = range.first_row;
    hyperlink.col = range.first_col;
    hyperlink.last_row = range.last_row;
    hyperlink.last_col = range.last_col;
    retained.push_back(std::move(hyperlink));
  }
  hyperlinks = std::move(retained);
}

void ShiftConditionalFormats(std::vector<cf::ConditionalFormat>& formats, std::uint32_t index, std::uint32_t count,
                             bool is_delete, bool row_axis) {
  std::vector<cf::ConditionalFormat> retained_formats;
  retained_formats.reserve(formats.size());
  for (cf::ConditionalFormat& format : formats) {
    std::vector<MergeRange> ranges;
    ranges.reserve(format.sqref.size());
    for (const cf::CFCellRange& range : format.sqref) {
      ranges.push_back(MergeRange{range.first.row, range.first.col, range.last.row, range.last.col});
    }
    shift_sqref_ranges(ranges, index, count, is_delete, row_axis);
    format.sqref.clear();
    for (const MergeRange& range : ranges) {
      format.sqref.push_back(
          cf::CFCellRange{CellAddress{range.first_row, range.first_col}, CellAddress{range.last_row, range.last_col}});
    }
    // A conditionalFormatting element without an sqref applies nowhere and
    // is invalid OOXML. Drop it when a row/column deletion consumed every
    // one of its ranges.
    if (!format.sqref.empty()) {
      retained_formats.push_back(std::move(format));
    }
  }
  formats = std::move(retained_formats);
}

void ShiftRowLayouts(std::vector<RowLayout>& rows, std::uint32_t index, std::uint32_t count, bool is_delete) {
  ShiftRowAnchored(rows, index, count, is_delete);
}

void ShiftColumnLayouts(std::vector<ColumnLayout>& columns, std::uint32_t index, std::uint32_t count, bool is_delete) {
  std::vector<ColumnLayout> retained;
  retained.reserve(columns.size());
  for (ColumnLayout& column : columns) {
    MergeRange shifted{0, column.first, 0, column.last};
    bool drop = false;
    ShiftColRange(shifted, index, count, is_delete, &drop);
    if (!drop) {
      column.first = shifted.first_col;
      column.last = shifted.last_col;
      retained.push_back(std::move(column));
    }
  }
  columns = std::move(retained);
}

void ShiftPivotAnchors(std::vector<std::unique_ptr<pivot::PivotTable>>& pivots, std::uint32_t index,
                       std::uint32_t count, bool is_delete, bool row_axis) {
  for (std::unique_ptr<pivot::PivotTable>& pivot : pivots) {
    if (pivot == nullptr) {
      continue;
    }
    std::uint32_t anchor = row_axis ? pivot->anchor_row() : pivot->anchor_col();
    if (anchor >= index) {
      if (is_delete) {
        anchor = anchor < index + count ? index : anchor - count;
      } else {
        const std::uint64_t shifted = static_cast<std::uint64_t>(anchor) + count;
        anchor = static_cast<std::uint32_t>(
            std::min<std::uint64_t>(shifted, row_axis ? Sheet::kMaxRows - 1U : Sheet::kMaxCols - 1U));
      }
    }
    std::uint32_t span_rows = pivot->span_rows();
    std::uint32_t span_cols = pivot->span_cols();
    // Clip the span on the edited axis to the grid; a zero span stays zero.
    if (row_axis && span_rows != 0U) {
      span_rows = std::min(span_rows, Sheet::kMaxRows - anchor);
    } else if (!row_axis && span_cols != 0U) {
      span_cols = std::min(span_cols, Sheet::kMaxCols - anchor);
    }
    pivot->set_anchor(row_axis ? anchor : pivot->anchor_row(), row_axis ? pivot->anchor_col() : anchor, span_rows,
                      span_cols);
  }
}

}  // namespace

// Beyond the per-range rules, a delete joins the two ranges it brings edge to
// edge across the removed band when they share the other axis's span, as
// Excel does: deleting column B turns `A1:A6 C1:C6` into `A1:B6`.
void shift_sqref_ranges(std::vector<MergeRange>& ranges, std::uint32_t index, std::uint32_t count, bool is_delete,
                        bool row_axis) {
  ShiftRangeList(ranges, index, count, is_delete, row_axis);
  if (!is_delete || index == 0U) {
    return;
  }
  const auto seam = [index, row_axis](const MergeRange& before, const MergeRange& after) {
    return row_axis ? before.last_row + 1U == index && after.first_row == index &&
                          before.first_col == after.first_col && before.last_col == after.last_col
                    : before.last_col + 1U == index && after.first_col == index &&
                          before.first_row == after.first_row && before.last_row == after.last_row;
  };
  for (std::size_t i = 0; i < ranges.size(); ++i) {
    for (std::size_t j = 0; j < ranges.size(); ++j) {
      if (i == j || !seam(ranges[i], ranges[j])) {
        continue;
      }
      (row_axis ? ranges[i].last_row : ranges[i].last_col) = row_axis ? ranges[j].last_row : ranges[j].last_col;
      ranges.erase(ranges.begin() + static_cast<std::ptrdiff_t>(j));
      i = static_cast<std::size_t>(-1);
      break;
    }
  }
}

bool shift_auto_filter(AutoFilter& filter, std::uint32_t index, std::uint32_t count, bool is_delete, bool row_axis,
                       bool header_delete_removes) {
  if (filter.is_opaque()) {
    return true;
  }
  const MergeRange before = filter.range;
  if (row_axis && is_delete && header_delete_removes && before.first_row >= index && before.first_row - index < count) {
    return false;
  }
  bool drop = false;
  if (row_axis) {
    ShiftRowRange(filter.range, index, count, is_delete, &drop);
  } else {
    ShiftColRange(filter.range, index, count, is_delete, &drop);
  }
  if (drop) {
    return false;
  }
  if (filter.sort) {
    bool sort_drop = false;
    if (row_axis) {
      ShiftRowRange(filter.sort->ref, index, count, is_delete, &sort_drop);
    } else {
      ShiftColRange(filter.sort->ref, index, count, is_delete, &sort_drop);
    }
    if (sort_drop) {
      filter.sort.reset();
    } else {
      std::vector<SortCondition> kept;
      for (SortCondition& cond : filter.sort->conditions) {
        bool cond_drop = false;
        if (row_axis) {
          ShiftRowRange(cond.ref, index, count, is_delete, &cond_drop);
        } else {
          ShiftColRange(cond.ref, index, count, is_delete, &cond_drop);
        }
        if (!cond_drop) {
          kept.push_back(std::move(cond));
        }
      }
      filter.sort->conditions = std::move(kept);
    }
  }
  if (row_axis) {
    return true;
  }
  // `colId` is an offset from the range's first column. A column delete
  // drops the criteria of deleted columns and pulls later ones back by the
  // deleted columns before them; an insert inside the range (not at its left
  // edge, which moves the whole range) pushes the columns at or after it.
  std::vector<FilterColumn> kept;
  kept.reserve(filter.columns.size());
  for (FilterColumn& col : filter.columns) {
    const std::uint64_t absolute = static_cast<std::uint64_t>(before.first_col) + col.col_id;
    if (is_delete) {
      const std::uint64_t del_end = static_cast<std::uint64_t>(index) + count;
      if (absolute >= index && absolute < del_end) {
        continue;
      }
      const std::uint64_t lo = std::max<std::uint64_t>(index, before.first_col);
      const std::uint64_t hi = std::min<std::uint64_t>(del_end, absolute);
      col.col_id -= static_cast<std::uint32_t>(hi > lo ? hi - lo : 0U);
    } else if (index > before.first_col && absolute >= index) {
      col.col_id += count;
    }
    if (col.col_id > filter.range.last_col - filter.range.first_col) {
      continue;  // Pushed past the sheet edge with the clamped range.
    }
    kept.push_back(std::move(col));
  }
  filter.columns = std::move(kept);
  return true;
}

// Every sheet-attached structure derives its new coordinates from the one
// `StructuralEdit` description. This list is the enumeration: a structure
// added to `Sheet` and not added here does not follow a row/column edit.
void Sheet::shift_sheet_metadata(const StructuralEdit& edit) {
  const std::uint32_t index = edit.index;
  const std::uint32_t count = edit.count;
  const bool is_delete = edit.is_delete;
  const bool row_axis = edit.row_axis;
  if (row_axis) {
    ShiftHyperlinkList(hyperlinks_, index, count, is_delete, /*row_axis=*/true);
    ShiftRowAnchored(comments_, index, count, is_delete);
    ShiftRowAnchored(threaded_comments_, index, count, is_delete);
  } else {
    ShiftHyperlinkList(hyperlinks_, index, count, is_delete, /*row_axis=*/false);
    ShiftColAnchored(comments_, index, count, is_delete);
    ShiftColAnchored(threaded_comments_, index, count, is_delete);
  }
  ShiftRangeList(merges_, index, count, is_delete, row_axis);
  for (DataValidation& dv : validations_) {
    ShiftRangeList(dv.ranges, index, count, is_delete, row_axis);
  }
  validations_.erase(std::remove_if(validations_.begin(), validations_.end(),
                                    [](const DataValidation& dv) { return dv.ranges.empty(); }),
                     validations_.end());
  ShiftConditionalFormats(conditional_formats_, index, count, is_delete, row_axis);
  if (row_axis) {
    ShiftRowLayouts(layout_.row_overrides, index, count, is_delete);
    ShiftAnchored(print_settings_.manual_row_breaks, &ManualBreak::id, index, count, is_delete, Sheet::kMaxRows);
  } else {
    ShiftColumnLayouts(layout_.columns, index, count, is_delete);
    ShiftAnchored(print_settings_.manual_col_breaks, &ManualBreak::id, index, count, is_delete, Sheet::kMaxCols);
  }
  ShiftPivotAnchors(pivot_tables_, index, count, is_delete, row_axis);
  if (AutoFilter* filter = auto_filter_.get();
      filter != nullptr && !shift_auto_filter(*filter, index, count, is_delete, row_axis,
                                              /*header_delete_removes=*/true)) {
    auto_filter_.reset();
  }
}

void Sheet::insert_rows(std::uint32_t row, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  // Walk populated rows in descending key order so a moved row never
  // collides with an existing row that still needs to move.
  std::vector<std::uint32_t> keys;
  keys.reserve(rows_.size());
  for (const auto& kv : rows_) {
    keys.push_back(kv.first);
  }
  // Ascending sort, walked backwards, so the shared ascending sort serves this
  // direction too.
  sort_ascending(keys);
  for (auto it = keys.rbegin(); it != keys.rend(); ++it) {
    const std::uint32_t key = *it;
    if (key < row) {
      continue;
    }
    const std::uint64_t shifted = static_cast<std::uint64_t>(key) + count;
    auto node = rows_.extract(key);
    if (shifted >= kMaxRows) {
      continue;  // Row pushed past sheet bound; drop the cells.
    }
    node.key() = static_cast<std::uint32_t>(shifted);
    rows_.insert(std::move(node));
  }
  clear_committed_spills_locked();
  rebuild_formula_index_locked();
  const StructuralEdit edit{row, count, /*is_delete=*/false, /*row_axis=*/true};
  shift_blocked_spills_locked(edit);
  shift_sheet_metadata(edit);
  cell_enumeration_revision_.bump();
}

void Sheet::delete_rows(std::uint32_t row, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  // Drop cells inside the deletion interval, then shift trailing rows
  // up. Walk in ascending key order — every shifted destination key is
  // strictly less than its source so no collision can occur.
  std::vector<std::uint32_t> keys;
  keys.reserve(rows_.size());
  for (const auto& kv : rows_) {
    keys.push_back(kv.first);
  }
  sort_ascending(keys);
  for (std::uint32_t key : keys) {
    if (key < row) {
      continue;
    }
    auto node = rows_.extract(key);
    if (key < row + count) {
      continue;  // Deleted row; drop the node.
    }
    node.key() = key - count;
    rows_.insert(std::move(node));
  }
  clear_committed_spills_locked();
  rebuild_formula_index_locked();
  const StructuralEdit edit{row, count, /*is_delete=*/true, /*row_axis=*/true};
  shift_blocked_spills_locked(edit);
  shift_sheet_metadata(edit);
  cell_enumeration_revision_.bump();
}

void Sheet::insert_cols(std::uint32_t col, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  for (auto& kv : rows_) {
    RowCells& row = kv.second;
    if (row.empty() || col >= row.size()) {
      continue;
    }
    // Insertion entirely left of the run: the run keeps its contents and only
    // its origin moves. A run pushed wholly past the last column is dropped.
    if (col <= row.first_col()) {
      const std::uint64_t shifted = static_cast<std::uint64_t>(row.first_col()) + count;
      if (shifted >= kMaxCols) {
        row.mutable_run().clear();
        row.set_first_col(0U);
        continue;
      }
      row.set_first_col(static_cast<std::uint32_t>(shifted));
      // The tail may now overhang the sheet; drop what no longer fits.
      const std::size_t capacity = static_cast<std::size_t>(kMaxCols) - row.first_col();
      if (row.mutable_run().size() > capacity) {
        row.mutable_run().resize(capacity);
      }
      continue;
    }
    // Insertion point sits inside this row's run. Pad with default Cells at
    // the local index. Cells past `kMaxCols` are dropped wholesale.
    std::vector<Cell>& cells = row.mutable_run();
    const std::size_t local_col = static_cast<std::size_t>(col) - row.first_col();
    const std::size_t old_size = cells.size();
    const std::uint64_t new_size_unbounded = static_cast<std::uint64_t>(old_size) + count;
    const std::size_t new_size =
        static_cast<std::size_t>(std::min<std::uint64_t>(new_size_unbounded, kMaxCols - row.first_col()));
    cells.resize(new_size);
    // Move the tail to its new position. Walk from the back so source
    // and destination never overlap during the swap.
    const std::size_t tail_count = old_size - local_col;
    for (std::size_t i = 0; i < tail_count; ++i) {
      const std::size_t src = old_size - 1U - i;
      const std::size_t dst = src + count;
      if (dst >= new_size) {
        continue;  // Cell pushed past sheet bound; drop it.
      }
      cells[dst] = std::move(cells[src]);
    }
    // Clear the inserted slots so they read as default-blank.
    const std::size_t clear_end = std::min<std::size_t>(local_col + count, new_size);
    for (std::size_t i = local_col; i < clear_end; ++i) {
      cells[i] = Cell{};
    }
  }
  clear_committed_spills_locked();
  rebuild_formula_index_locked();
  const StructuralEdit edit{col, count, /*is_delete=*/false, /*row_axis=*/false};
  shift_blocked_spills_locked(edit);
  shift_sheet_metadata(edit);
  cell_enumeration_revision_.bump();
}

void Sheet::delete_cols(std::uint32_t col, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  for (auto& kv : rows_) {
    RowCells& row = kv.second;
    if (row.empty() || col >= row.size()) {
      continue;
    }
    std::vector<Cell>& cells = row.mutable_run();
    const std::uint64_t band_end = static_cast<std::uint64_t>(col) + count;
    if (band_end <= row.first_col()) {
      // The deleted band lies entirely left of the run: only the origin moves.
      row.set_first_col(row.first_col() - count);
      continue;
    }
    // The band reaches into the run. Everything from the run's start up to the
    // band's end goes; what remains starts at `col`.
    const std::size_t local_first = col > row.first_col() ? static_cast<std::size_t>(col) - row.first_col() : 0U;
    const std::size_t local_last =
        static_cast<std::size_t>(std::min<std::uint64_t>(band_end - row.first_col(), cells.size()));
    cells.erase(cells.begin() + static_cast<std::ptrdiff_t>(local_first),
                cells.begin() + static_cast<std::ptrdiff_t>(local_last));
    if (col < row.first_col()) {
      row.set_first_col(col);
    }
  }
  clear_committed_spills_locked();
  rebuild_formula_index_locked();
  const StructuralEdit edit{col, count, /*is_delete=*/true, /*row_axis=*/false};
  shift_blocked_spills_locked(edit);
  shift_sheet_metadata(edit);
  cell_enumeration_revision_.bump();
}

}  // namespace formulon
