//
// Row and column insert/delete for `Sheet`: the cell-store shift and the
// coordinate remap of every sheet-attached structure that follows it.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "utils/a1_column.h"
#include "utils/a1_ref.h"
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
    std::vector<cf::CFCellRange> retained;
    retained.reserve(format.sqref.size());
    for (cf::CFCellRange& range : format.sqref) {
      MergeRange shifted{range.first.row, range.first.col, range.last.row, range.last.col};
      bool drop = false;
      if (row_axis) {
        ShiftRowRange(shifted, index, count, is_delete, &drop);
      } else {
        ShiftColRange(shifted, index, count, is_delete, &drop);
      }
      if (!drop) {
        range.first = CellAddress{shifted.first_row, shifted.first_col};
        range.last = CellAddress{shifted.last_row, shifted.last_col};
        retained.push_back(std::move(range));
      }
    }
    format.sqref = std::move(retained);
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

/// Renders `[first_row..last_row] x [first_col..last_col]` as an OOXML
/// `ref` rectangle, collapsing a single-cell rectangle to one address the
/// way Excel writes it. Returns false when a coordinate is outside the grid.
bool AppendRefRectangle(std::string& out, const MergeRange& rect) {
  const auto append_cell = [&out](std::uint32_t row, std::uint32_t col) {
    if (!a1::append_column_letters(out, col)) {
      return false;
    }
    out += std::to_string(static_cast<std::uint64_t>(row) + 1U);
    return true;
  };
  if (!append_cell(rect.first_row, rect.first_col)) {
    return false;
  }
  if (rect.first_row == rect.last_row && rect.first_col == rect.last_col) {
    return true;
  }
  out += ':';
  return append_cell(rect.last_row, rect.last_col);
}

/// Rewrites the `ref` rectangle of a raw `<autoFilter>` element through the
/// same span rules the modelled rectangles follow, and clears the element
/// when the edit consumed its whole rectangle.
///
/// The element is retained verbatim, so this is the one coordinate inside it
/// that moves. It is also the one that decides which cells the filter is
/// attached to: leaving it behind points the filter at whatever occupies the
/// old rectangle after the edit. The `colId` offsets on any `<filterColumn>`
/// children are relative to this rectangle's first column and are not
/// remapped — see the field's declaration for what that costs.
void ShiftAutoFilterRef(std::string& xml, std::uint32_t index, std::uint32_t count, bool is_delete, bool row_axis) {
  if (xml.empty()) {
    return;
  }
  // Confine the search to the start tag: only the `<autoFilter>` element
  // itself carries the rectangle.
  const std::size_t tag_end = xml.find('>');
  const std::size_t attr = xml.find("ref=\"");
  if (tag_end == std::string::npos || attr == std::string::npos || attr > tag_end) {
    return;
  }
  const std::size_t value_begin = attr + 5U;
  const std::size_t value_end = xml.find('"', value_begin);
  if (value_end == std::string::npos || value_end > tag_end) {
    return;
  }

  const std::string_view value(xml.data() + value_begin, value_end - value_begin);
  const std::size_t colon = value.find(':');
  MergeRange rect;
  if (!a1::parse_a1_ref(value.substr(0, colon), &rect.first_row, &rect.first_col)) {
    return;  // Not a plain A1 rectangle; leave the element untouched.
  }
  if (colon == std::string_view::npos) {
    rect.last_row = rect.first_row;
    rect.last_col = rect.first_col;
  } else if (!a1::parse_a1_ref(value.substr(colon + 1U), &rect.last_row, &rect.last_col)) {
    return;
  }
  if (rect.first_row > rect.last_row || rect.first_col > rect.last_col) {
    return;
  }

  bool drop = false;
  if (row_axis) {
    ShiftRowRange(rect, index, count, is_delete, &drop);
  } else {
    ShiftColRange(rect, index, count, is_delete, &drop);
  }
  if (drop) {
    // Every filtered cell was deleted, which is what Excel resolves by
    // removing the filter rather than by keeping an empty one.
    xml.clear();
    return;
  }
  std::string replacement;
  if (!AppendRefRectangle(replacement, rect)) {
    return;
  }
  xml.replace(value_begin, value_end - value_begin, replacement);
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

void ShiftBreaks(std::vector<ManualBreak>& breaks, std::uint32_t index, std::uint32_t count, bool is_delete,
                 std::uint32_t bound) {
  std::vector<ManualBreak> retained;
  retained.reserve(breaks.size());
  for (ManualBreak& page_break : breaks) {
    if (page_break.id < index) {
      retained.push_back(std::move(page_break));
      continue;
    }
    if (is_delete) {
      if (page_break.id < index + count) {
        continue;
      }
      page_break.id -= count;
    } else {
      const std::uint64_t shifted = static_cast<std::uint64_t>(page_break.id) + count;
      if (shifted >= bound) {
        continue;
      }
      page_break.id = static_cast<std::uint32_t>(shifted);
    }
    retained.push_back(std::move(page_break));
  }
  breaks = std::move(retained);
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
    pivot->set_anchor(row_axis ? anchor : pivot->anchor_row(), row_axis ? pivot->anchor_col() : anchor,
                      pivot->span_rows(), pivot->span_cols());
  }
}

}  // namespace

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
  } else {
    ShiftHyperlinkList(hyperlinks_, index, count, is_delete, /*row_axis=*/false);
    ShiftColAnchored(comments_, index, count, is_delete);
  }
  ShiftRangeList(merges_, index, count, is_delete, row_axis);
  for (DataValidation& dv : validations_) {
    ShiftRangeList(dv.ranges, index, count, is_delete, row_axis);
  }
  ShiftConditionalFormats(conditional_formats_, index, count, is_delete, row_axis);
  if (row_axis) {
    ShiftRowLayouts(layout_.row_overrides, index, count, is_delete);
    ShiftBreaks(print_settings_.manual_row_breaks, index, count, is_delete, Sheet::kMaxRows);
  } else {
    ShiftColumnLayouts(layout_.columns, index, count, is_delete);
    ShiftBreaks(print_settings_.manual_col_breaks, index, count, is_delete, Sheet::kMaxCols);
  }
  ShiftPivotAnchors(pivot_tables_, index, count, is_delete, row_axis);
  ShiftAutoFilterRef(auto_filter_xml_, index, count, is_delete, row_axis);
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
