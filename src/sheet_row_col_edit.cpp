//
// Row and column insert/delete for `Sheet`: the cell-store shift and the
// coordinate remap of every sheet-attached structure that follows it.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "utils/index_sort.h"

namespace formulon {

namespace {

// Shifts every entry in `items` whose anchor lies on or past `index` by the
// insert / delete rule encoded in `count` and `is_delete`. Entries whose
// anchor falls inside the deleted interval, or whose insert would push it
// past `bound`, are removed. `field` names the axis coordinate, so one
// instantiation per anchor type serves both axes. Under a `band`, an entry
// whose `perp` coordinate lies outside it stays.
template <typename Anchor>
void ShiftAnchored(std::vector<Anchor>& items, std::uint32_t Anchor::*field, std::uint32_t index, std::uint32_t count,
                   bool is_delete, std::uint32_t bound, const std::optional<Band>& band = std::nullopt,
                   std::uint32_t Anchor::*perp = nullptr) {
  std::vector<Anchor> retained;
  retained.reserve(items.size());
  for (Anchor& item : items) {
    if (item.*field < index || (band.has_value() && (item.*perp < band->first || item.*perp > band->last))) {
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

// A note or threaded comment moves with its cell.
template <typename Anchor>
void ShiftCellAnchored(std::vector<Anchor>& items, const StructuralEdit& edit) {
  if (edit.row_axis) {
    ShiftAnchored(items, &Anchor::row, edit.index, edit.count, edit.is_delete, Sheet::kMaxRows, edit.band,
                  &Anchor::col);
  } else {
    ShiftAnchored(items, &Anchor::col, edit.index, edit.count, edit.is_delete, Sheet::kMaxCols, edit.band,
                  &Anchor::row);
  }
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

void ShiftAxisRange(MergeRange& range, std::uint32_t index, std::uint32_t count, bool is_delete, bool row_axis,
                    bool* out_drop) {
  if (row_axis) {
    ShiftRowRange(range, index, count, is_delete, out_drop);
  } else {
    ShiftColRange(range, index, count, is_delete, out_drop);
  }
}

void ShiftRangeList(std::vector<MergeRange>& ranges, std::uint32_t index, std::uint32_t count, bool is_delete,
                    bool row_axis) {
  std::size_t kept = 0;
  for (MergeRange& range : ranges) {
    bool drop = false;
    ShiftAxisRange(range, index, count, is_delete, row_axis, &drop);
    if (!drop) {
      ranges[kept++] = range;
    }
  }
  ranges.erase(ranges.begin() + static_cast<std::ptrdiff_t>(kept), ranges.end());
}

// Under a band, a merge inside it moves; one straddling the moving region, or
// one a delete cuts into, is unmerged.
void ShiftMerges(std::vector<MergeRange>& merges, const StructuralEdit& edit) {
  if (!edit.band.has_value()) {
    ShiftRangeList(merges, edit.index, edit.count, edit.is_delete, edit.row_axis);
    return;
  }
  const std::uint64_t delete_end = static_cast<std::uint64_t>(edit.index) + edit.count;
  std::size_t kept = 0;
  for (MergeRange& merge : merges) {
    const BandFit fit = band_fit(merge, edit);
    bool drop = fit == BandFit::kStraddles;
    if (fit == BandFit::kInside) {
      const std::uint32_t first = edit.row_axis ? merge.first_row : merge.first_col;
      const std::uint32_t last = edit.row_axis ? merge.last_row : merge.last_col;
      drop = edit.is_delete && first < delete_end && last >= edit.index;
      if (!drop) {
        ShiftAxisRange(merge, edit.index, edit.count, edit.is_delete, edit.row_axis, &drop);
      }
    }
    if (!drop) {
      merges[kept++] = merge;
    }
  }
  merges.erase(merges.begin() + static_cast<std::ptrdiff_t>(kept), merges.end());
}

// Moves a data table's input cell along one axis. An input cell that is
// deleted keeps its (clamped) position and is flagged `deleted`, the way the
// `del1` / `del2` attributes record it. Under a band, a cell outside it stays.
void ShiftDataTableInput(DataTableInput& input, const StructuralEdit& edit) {
  const std::uint32_t index = edit.index;
  const std::uint32_t count = edit.count;
  const bool is_delete = edit.is_delete;
  const bool row_axis = edit.row_axis;
  const std::uint32_t perp = row_axis ? input.col : input.row;
  if (!input.present || input.deleted ||
      (edit.band.has_value() && (perp < edit.band->first || perp > edit.band->last))) {
    return;
  }
  std::uint32_t& pos = row_axis ? input.row : input.col;
  const std::uint32_t bound = a1::axis_extent(row_axis);
  if (pos < index) {
    return;
  }
  if (is_delete) {
    if (pos < index + count) {
      pos = index < bound ? index : bound - 1U;
      input.deleted = true;
    } else {
      pos -= count;
    }
    return;
  }
  const std::uint64_t shifted = static_cast<std::uint64_t>(pos) + count;
  if (shifted >= bound) {
    pos = bound - 1U;
    input.deleted = true;
  } else {
    pos = static_cast<std::uint32_t>(shifted);
  }
}

// A table the edit misses shifts whole, its input cells with it; one the edit
// cuts into goes (the workbook refuses a cut that would leave part of it). A
// band edit moves only a table inside its band.
void ShiftDataTables(std::vector<DataTable>& tables, const StructuralEdit& edit) {
  const std::uint32_t index = edit.index;
  const std::uint32_t count = edit.count;
  const bool is_delete = edit.is_delete;
  const bool row_axis = edit.row_axis;
  std::size_t kept = 0;
  for (DataTable& table : tables) {
    const bool stays = edit.band.has_value() && band_fit(data_table_block(table), edit) != BandFit::kInside;
    if (!stays && data_table_structure_cut(table, edit) != StructureEditCut::kMisses) {
      continue;
    }
    if (stays) {
      if (&tables[kept] != &table) {
        tables[kept] = std::move(table);
      }
      ++kept;
      continue;
    }
    bool drop = false;
    ShiftAxisRange(table.ref, index, count, is_delete, row_axis, &drop);
    if (drop) {
      continue;
    }
    const bool was_deleted1 = table.r1.deleted;
    const bool was_deleted2 = table.r2.deleted;
    ShiftDataTableInput(table.r1, edit);
    ShiftDataTableInput(table.r2, edit);
    if (table.r1.deleted && !was_deleted1) {
      table.explicit_attrs |= kDataTableExplicitDel1;
    }
    if (table.r2.deleted && !was_deleted2) {
      table.explicit_attrs |= kDataTableExplicitDel2;
    }
    if (&tables[kept] != &table) {
      tables[kept] = std::move(table);
    }
    ++kept;
  }
  tables.erase(tables.begin() + static_cast<std::ptrdiff_t>(kept), tables.end());
}

// Hyperlinks use the same inclusive rectangle semantics as merge ranges.
// Keeping this conversion next to ShiftRangeList is intentional: row/column
// insertions and deletions must apply the exact same shift, shrink, clamp and
// drop rules to both metadata shapes.
// Under a band, a hyperlink moves only when its range lies inside it; one
// straddling the band stays with its key cell.
void ShiftHyperlinkList(std::vector<Hyperlink>& hyperlinks, const StructuralEdit& edit) {
  std::size_t kept = 0;
  for (Hyperlink& hyperlink : hyperlinks) {
    MergeRange range{hyperlink.row, hyperlink.col, hyperlink.last_row, hyperlink.last_col};
    bool drop = false;
    if (!edit.band.has_value() || band_fit(range, edit) == BandFit::kInside) {
      ShiftAxisRange(range, edit.index, edit.count, edit.is_delete, edit.row_axis, &drop);
    }
    if (drop) {
      continue;
    }
    hyperlink.row = range.first_row;
    hyperlink.col = range.first_col;
    hyperlink.last_row = range.last_row;
    hyperlink.last_col = range.last_col;
    if (&hyperlinks[kept] != &hyperlink) {
      hyperlinks[kept] = std::move(hyperlink);
    }
    ++kept;
  }
  hyperlinks.erase(hyperlinks.begin() + static_cast<std::ptrdiff_t>(kept), hyperlinks.end());
}

// Whether `ranges` (two or more) tile exactly one rectangle.
bool TilesOneRectangle(const std::vector<MergeRange>& ranges) {
  if (ranges.size() < 2U) {
    return false;
  }
  MergeRange bounds = ranges.front();
  std::uint64_t area = 0;
  for (std::size_t i = 0; i < ranges.size(); ++i) {
    const MergeRange& r = ranges[i];
    bounds = MergeRange{std::min(bounds.first_row, r.first_row), std::min(bounds.first_col, r.first_col),
                        std::max(bounds.last_row, r.last_row), std::max(bounds.last_col, r.last_col)};
    area += (static_cast<std::uint64_t>(r.last_row) - r.first_row + 1U) * (r.last_col - r.first_col + 1U);
    for (std::size_t j = 0; j < i; ++j) {
      const MergeRange& o = ranges[j];
      if (r.first_row <= o.last_row && o.first_row <= r.last_row && r.first_col <= o.last_col &&
          o.first_col <= r.last_col) {
        return false;
      }
    }
  }
  return area == (static_cast<std::uint64_t>(bounds.last_row) - bounds.first_row + 1U) *
                     (bounds.last_col - bounds.first_col + 1U);
}

void ShiftConditionalFormats(std::vector<cf::ConditionalFormat>& formats, const StructuralEdit& edit) {
  std::vector<cf::ConditionalFormat> retained_formats;
  retained_formats.reserve(formats.size());
  for (cf::ConditionalFormat& format : formats) {
    std::vector<MergeRange> ranges;
    ranges.reserve(format.sqref.size());
    for (const cf::CFCellRange& range : format.sqref) {
      ranges.push_back(MergeRange{range.first.row, range.first.col, range.last.row, range.last.col});
    }
    shift_cf_sqref_ranges(ranges, edit);
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

// A pivot the edit misses shifts whole; one the edit deletes, or cuts into
// (which the workbook refuses), goes. A band edit moves only a pivot inside
// its band.
void ShiftPivotAnchors(std::vector<std::unique_ptr<pivot::PivotTable>>& pivots, const StructuralEdit& edit) {
  const std::uint32_t index = edit.index;
  const std::uint32_t count = edit.count;
  const bool is_delete = edit.is_delete;
  const bool row_axis = edit.row_axis;
  for (std::unique_ptr<pivot::PivotTable>& pivot : pivots) {
    if (pivot == nullptr || (edit.band.has_value() && band_fit(pivot_location(*pivot), edit) != BandFit::kInside)) {
      continue;
    }
    if (pivot_structure_cut(*pivot, edit) != StructureEditCut::kMisses) {
      pivot.reset();
      continue;
    }
    std::uint32_t anchor = row_axis ? pivot->anchor_row() : pivot->anchor_col();
    if (anchor >= index) {
      anchor = is_delete ? anchor - count : anchor + count;
    }
    pivot->set_anchor(row_axis ? anchor : pivot->anchor_row(), row_axis ? pivot->anchor_col() : anchor,
                      pivot->span_rows(), pivot->span_cols());
    // The cells and merges the last projection wrote move with the report.
    std::vector<MergeRange> projected;
    if (pivot->projected_extent().has_value()) {
      const pivot::PivotRect& r = *pivot->projected_extent();
      projected.push_back(MergeRange{r.row0, r.col0, r.row1, r.col1});
    }
    const std::size_t extent_count = projected.size();
    for (const pivot::PivotRect& r : pivot->projected_merges()) {
      projected.push_back(MergeRange{r.row0, r.col0, r.row1, r.col1});
    }
    std::vector<bool> dropped(projected.size(), false);
    for (std::size_t i = 0; i < projected.size(); ++i) {
      bool drop = false;
      ShiftAxisRange(projected[i], index, count, is_delete, row_axis, &drop);
      dropped[i] = drop;
    }
    const auto to_rect = [](const MergeRange& m) {
      return pivot::PivotRect{m.first_row, m.first_col, m.last_row, m.last_col};
    };
    if (extent_count != 0U) {
      pivot->set_projected_extent(dropped[0] ? std::nullopt : std::optional<pivot::PivotRect>(to_rect(projected[0])));
    }
    std::vector<pivot::PivotRect>& merges = pivot->mutable_projected_merges();
    merges.clear();
    for (std::size_t i = extent_count; i < projected.size(); ++i) {
      if (!dropped[i]) {
        merges.push_back(to_rect(projected[i]));
      }
    }
  }
  pivots.erase(std::remove(pivots.begin(), pivots.end(), nullptr), pivots.end());
}

// Populated row keys of `rows` that are at least `first`, ascending.
std::vector<std::uint32_t> SortedRowKeys(const std::unordered_map<std::uint32_t, RowCells>& rows, std::uint32_t first) {
  std::vector<std::uint32_t> keys;
  for (const auto& kv : rows) {
    if (kv.first >= first) {
      keys.push_back(kv.first);
    }
  }
  sort_ascending(keys);
  return keys;
}

}  // namespace

StructureEditCut structure_edit_cut(std::uint32_t first, std::uint32_t last, const StructuralEdit& edit) noexcept {
  if (!edit.is_delete) {
    if (edit.index > last) {
      return StructureEditCut::kMisses;
    }
    const std::uint32_t bound = a1::axis_extent(edit.row_axis);
    if (static_cast<std::uint64_t>(last) + edit.count >= bound) {
      return StructureEditCut::kPushesOff;
    }
    return edit.index > first ? StructureEditCut::kSplits : StructureEditCut::kMisses;
  }
  const std::uint64_t end = static_cast<std::uint64_t>(edit.index) + edit.count;  // exclusive
  if (end <= first || edit.index > last) {
    return StructureEditCut::kMisses;
  }
  return edit.index <= first && end > last ? StructureEditCut::kRemoves : StructureEditCut::kSplits;
}

StructureEditCut data_table_structure_cut(const DataTable& table, const StructuralEdit& edit) noexcept {
  const std::uint32_t first = edit.row_axis ? table.ref.first_row : table.ref.first_col;
  const std::uint32_t last = edit.row_axis ? table.ref.last_row : table.ref.last_col;
  return structure_edit_cut(first == 0U ? 0U : first - 1U, last, edit);
}

MergeRange data_table_block(const DataTable& table) noexcept {
  MergeRange block = table.ref;
  block.first_row -= block.first_row == 0U ? 0U : 1U;
  block.first_col -= block.first_col == 0U ? 0U : 1U;
  return block;
}

MergeRange pivot_location(const pivot::PivotTable& pivot) noexcept {
  const std::uint32_t rows = pivot.span_rows() == 0U ? 1U : pivot.span_rows();
  const std::uint32_t cols = pivot.span_cols() == 0U ? 1U : pivot.span_cols();
  return MergeRange{pivot.anchor_row(), pivot.anchor_col(), pivot.anchor_row() + rows - 1U,
                    pivot.anchor_col() + cols - 1U};
}

BandFit band_fit(const MergeRange& rect, const StructuralEdit& edit) noexcept {
  const Band& band = *edit.band;
  const std::uint32_t perp_first = edit.row_axis ? rect.first_col : rect.first_row;
  const std::uint32_t perp_last = edit.row_axis ? rect.last_col : rect.last_row;
  const std::uint32_t axial_last = edit.row_axis ? rect.last_row : rect.last_col;
  if (perp_last < band.first || perp_first > band.last || axial_last < edit.index) {
    return BandFit::kOutside;
  }
  return band.first <= perp_first && perp_last <= band.last ? BandFit::kInside : BandFit::kStraddles;
}

StructureEditCut pivot_structure_cut(const pivot::PivotTable& pivot, const StructuralEdit& edit) noexcept {
  const std::uint32_t first = edit.row_axis ? pivot.anchor_row() : pivot.anchor_col();
  const std::uint32_t span = edit.row_axis ? pivot.span_rows() : pivot.span_cols();
  return structure_edit_cut(first, span == 0U ? first : first + span - 1U, edit);
}

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

namespace {

// Appends `range` as an edit moves a cut sqref: the part before the edit stays
// put, the rest moves as a range of its own.
void CutOneRange(const MergeRange& range, std::uint32_t index, std::uint32_t count, bool is_delete, bool row_axis,
                 std::vector<MergeRange>& out) {
  MergeRange rest = range;
  if ((row_axis ? range.first_row : range.first_col) < index && index <= (row_axis ? range.last_row : range.last_col)) {
    MergeRange head = range;
    (row_axis ? head.last_row : head.last_col) = index - 1U;
    out.push_back(head);
    (row_axis ? rest.first_row : rest.first_col) = index;
  }
  bool drop = false;
  ShiftAxisRange(rest, index, count, is_delete, row_axis, &drop);
  if (!drop) {
    out.push_back(rest);
  }
}

// How an sqref follows a `StructuralEdit`; the modes differ only under a band.
enum class SqrefMode : std::uint8_t {
  kShift,      // A straddling range splits into its moved and unmoved parts.
  kCf,         // `kShift`, then a delete joins what it brings into one tiled rectangle.
  kReference,  // Only a range inside the band moves.
  kCut,        // Only a range inside the band is cut.
};

void EditSqrefRanges(std::vector<MergeRange>& ranges, const StructuralEdit& edit, SqrefMode mode) {
  if (!edit.band.has_value()) {
    if (mode == SqrefMode::kCut) {
      cut_sqref_ranges(ranges, edit.index, edit.count, edit.is_delete, edit.row_axis);
    } else {
      shift_sqref_ranges(ranges, edit.index, edit.count, edit.is_delete, edit.row_axis);
    }
    return;
  }
  const bool join_tiles = mode == SqrefMode::kCf && edit.is_delete && !TilesOneRectangle(ranges);
  const bool splits = mode == SqrefMode::kShift || mode == SqrefMode::kCf;
  const Band band = *edit.band;
  std::vector<MergeRange> out;
  out.reserve(ranges.size() * 3U);
  // `range` narrowed to the perpendicular interval [first, last].
  const auto across = [&edit](MergeRange range, std::uint32_t first, std::uint32_t last) {
    (edit.row_axis ? range.first_col : range.first_row) = first;
    (edit.row_axis ? range.last_col : range.last_row) = last;
    return range;
  };
  for (const MergeRange& range : ranges) {
    const BandFit fit = band_fit(range, edit);
    if (fit == BandFit::kOutside || (fit == BandFit::kStraddles && !splits)) {
      out.push_back(range);
      continue;
    }
    if (mode == SqrefMode::kCut) {
      CutOneRange(range, edit.index, edit.count, edit.is_delete, edit.row_axis, out);
      continue;
    }
    MergeRange moving = range;
    if (fit == BandFit::kStraddles) {
      const std::uint32_t first = edit.row_axis ? range.first_col : range.first_row;
      const std::uint32_t last = edit.row_axis ? range.last_col : range.last_row;
      if (first < band.first) {
        out.push_back(across(range, first, band.first - 1U));
      }
      if (last > band.last) {
        out.push_back(across(range, band.last + 1U, last));
      }
      moving = across(range, std::max(first, band.first), std::min(last, band.last));
    }
    bool drop = false;
    ShiftAxisRange(moving, edit.index, edit.count, edit.is_delete, edit.row_axis, &drop);
    if (!drop) {
      out.push_back(moving);
    }
  }
  ranges = std::move(out);
  if (!join_tiles || !TilesOneRectangle(ranges)) {
    return;
  }
  MergeRange joined = ranges.front();
  for (const MergeRange& r : ranges) {
    joined = MergeRange{std::min(joined.first_row, r.first_row), std::min(joined.first_col, r.first_col),
                        std::max(joined.last_row, r.last_row), std::max(joined.last_col, r.last_col)};
  }
  ranges.assign(1U, joined);
}

}  // namespace

void cut_sqref_ranges(std::vector<MergeRange>& ranges, std::uint32_t index, std::uint32_t count, bool is_delete,
                      bool row_axis) {
  std::vector<MergeRange> out;
  out.reserve(ranges.size() * 2U);
  for (const MergeRange& range : ranges) {
    CutOneRange(range, index, count, is_delete, row_axis, out);
  }
  ranges = std::move(out);
}

void shift_sqref_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit) {
  EditSqrefRanges(ranges, edit, SqrefMode::kShift);
}

void shift_cf_sqref_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit) {
  EditSqrefRanges(ranges, edit, SqrefMode::kCf);
}

void shift_reference_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit) {
  EditSqrefRanges(ranges, edit, SqrefMode::kReference);
}

void cut_sqref_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit) {
  EditSqrefRanges(ranges, edit, SqrefMode::kCut);
}

bool shift_auto_filter(AutoFilter& filter, const StructuralEdit& edit, bool header_delete_removes) {
  if (edit.band.has_value() && (filter.is_opaque() || band_fit(filter.range, edit) != BandFit::kInside)) {
    return true;
  }
  return shift_auto_filter(filter, edit.index, edit.count, edit.is_delete, edit.row_axis, header_delete_removes);
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
  ShiftAxisRange(filter.range, index, count, is_delete, row_axis, &drop);
  if (drop) {
    return false;
  }
  if (filter.sort) {
    bool sort_drop = false;
    ShiftAxisRange(filter.sort->ref, index, count, is_delete, row_axis, &sort_drop);
    if (sort_drop) {
      filter.sort.reset();
    } else {
      std::vector<SortCondition> kept;
      for (SortCondition& cond : filter.sort->conditions) {
        bool cond_drop = false;
        ShiftAxisRange(cond.ref, index, count, is_delete, row_axis, &cond_drop);
        if (!cond_drop) {
          kept.push_back(std::move(cond));
        }
      }
      // A sort state goes with the last of its conditions.
      if (kept.empty() && !filter.sort->conditions.empty()) {
        filter.sort.reset();
      } else {
        filter.sort->conditions = std::move(kept);
      }
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
    std::uint64_t col_id = col.col_id;
    const std::uint64_t absolute = static_cast<std::uint64_t>(before.first_col) + col_id;
    if (is_delete) {
      const std::uint64_t del_end = static_cast<std::uint64_t>(index) + count;
      if (absolute >= index && absolute < del_end) {
        continue;
      }
      const std::uint64_t lo = std::max<std::uint64_t>(index, before.first_col);
      const std::uint64_t hi = std::min<std::uint64_t>(del_end, absolute);
      col_id -= hi > lo ? hi - lo : 0U;
    } else if (index > before.first_col && absolute >= index) {
      col_id += count;
    }
    if (col_id > static_cast<std::uint64_t>(filter.range.last_col - filter.range.first_col)) {
      continue;  // Pushed past the sheet edge with the clamped range.
    }
    col.col_id = static_cast<std::uint32_t>(col_id);
    kept.push_back(std::move(col));
  }
  filter.columns = std::move(kept);
  return true;
}

// Every sheet-attached structure derives its new coordinates from the one
// `StructuralEdit` description. This list is the enumeration: a structure
// added to `Sheet` and not added here does not follow a row/column edit.
// A band edit leaves the row and column geometry (layout, page breaks) alone.
void Sheet::shift_sheet_metadata(const StructuralEdit& edit) {
  const std::uint32_t index = edit.index;
  const std::uint32_t count = edit.count;
  const bool is_delete = edit.is_delete;
  const bool row_axis = edit.row_axis;
  ShiftHyperlinkList(hyperlinks_, edit);
  ShiftCellAnchored(comments_, edit);
  ShiftCellAnchored(threaded_comments_, edit);
  ShiftMerges(merges_, edit);
  ShiftDataTables(data_tables_, edit);
  for (DataValidation& dv : validations_) {
    if (edit.band.has_value()) {
      shift_sqref_ranges(dv.ranges, edit);
    } else {
      ShiftRangeList(dv.ranges, index, count, is_delete, row_axis);
    }
  }
  validations_.erase(std::remove_if(validations_.begin(), validations_.end(),
                                    [](const DataValidation& dv) { return dv.ranges.empty(); }),
                     validations_.end());
  ShiftConditionalFormats(conditional_formats_, edit);
  if (!edit.band.has_value() && row_axis) {
    ShiftRowAnchored(layout_.row_overrides, index, count, is_delete);
    ShiftAnchored(print_settings_.manual_row_breaks, &ManualBreak::id, index, count, is_delete, Sheet::kMaxRows);
  } else if (!edit.band.has_value()) {
    ShiftColumnLayouts(layout_.columns, index, count, is_delete);
    ShiftAnchored(print_settings_.manual_col_breaks, &ManualBreak::id, index, count, is_delete, Sheet::kMaxCols);
  }
  ShiftPivotAnchors(pivot_tables_, edit);
  if (AutoFilter* filter = auto_filter_.get();
      filter != nullptr && !shift_auto_filter(*filter, edit, /*header_delete_removes=*/true)) {
    auto_filter_.reset();
  }
}

void Sheet::finalize_structural_edit_locked(const StructuralEdit& edit) {
  clear_committed_spills_locked();
  rebuild_formula_index_locked();
  shift_blocked_spills_locked(edit);
  shift_sheet_metadata(edit);
  if (!edit.is_delete) {
    inherit_inserted_formatting_locked(edit);
  }
  cell_enumeration_revision_.bump();
}

void Sheet::inherit_inserted_formatting_locked(const StructuralEdit& edit) {
  const std::uint32_t bound = a1::axis_extent(edit.row_axis);
  if (edit.index == 0U || edit.index >= bound) {
    return;
  }
  const std::uint32_t source = edit.index - 1U;
  const std::uint32_t end =
      static_cast<std::uint32_t>(std::min<std::uint64_t>(static_cast<std::uint64_t>(edit.index) + edit.count, bound));
  const auto owns_xf = [](const Cell& cell) { return cell.has_explicit_xf || cell.xf_index != 0U; };
  if (edit.band.has_value()) {
    // Insert Cells: each inserted cell takes the XF of the cell above (row
    // axis) or to the left; the row and column styles stay where they are.
    const auto give = [&](Cell& to, std::uint32_t xf_index) {
      to.xf_index = xf_index;
      to.has_explicit_xf = true;
    };
    if (edit.row_axis) {
      const auto it = rows_.find(source);
      if (it == rows_.end()) {
        return;
      }
      std::vector<std::pair<std::uint32_t, std::uint32_t>> styled;  // (col, xf)
      for (const RowCells::Entry entry :
           it->second.entries_in(edit.band->first, static_cast<std::uint64_t>(edit.band->last) + 1U)) {
        if (owns_xf(entry.cell)) {
          styled.emplace_back(entry.col, entry.cell.xf_index);
        }
      }
      for (const auto& [col, xf_index] : styled) {
        for (std::uint32_t row = edit.index; row < end; ++row) {
          give(rows_[row].ensure(col), xf_index);
        }
      }
      return;
    }
    for (auto& kv : rows_) {
      if (kv.first < edit.band->first || kv.first > edit.band->last) {
        continue;
      }
      const Cell* from = kv.second.find(source);
      if (from == nullptr || !owns_xf(*from)) {
        continue;
      }
      const std::uint32_t xf_index = from->xf_index;
      for (std::uint32_t col = edit.index; col < end; ++col) {
        give(kv.second.ensure(col), xf_index);
      }
    }
    return;
  }
  // Gives the cell `cell_at(i)` the style of the source for every i in the inserted band.
  const auto copy_xf = [&](std::uint32_t xf_index, auto&& cell_at) {
    for (std::uint32_t i = edit.index; i < end; ++i) {
      Cell& to = cell_at(i);
      to.xf_index = xf_index;
      to.has_explicit_xf = true;
    }
  };
  if (edit.row_axis) {
    // Style and an explicit height carry over; hidden state and outline level do not.
    // Copied by value: inserting overrides below invalidates pointers into the list.
    const RowLayout* found = find_row_override(layout_, source);
    const RowLayout styled = found != nullptr ? *found : RowLayout{};
    const bool inherits_height = styled.has_height && styled.custom_height;
    if (styled.has_style || inherits_height) {
      for (std::uint32_t row = edit.index; row < end; ++row) {
        RowLayout entry;
        entry.row = row;
        entry.has_style = styled.has_style;
        entry.style_xf = styled.has_style ? styled.style_xf : 0U;
        if (inherits_height) {
          entry.has_height = true;
          entry.height = styled.height;
          entry.custom_height = true;
        }
        insert_row_override(layout_, entry);
      }
    }
    const auto it = rows_.find(source);
    if (it == rows_.end()) {
      return;
    }
    for (const auto& [col, cell] : it->second) {
      if (!owns_xf(cell)) {
        continue;
      }
      const std::uint32_t source_col = col;
      copy_xf(cell.xf_index, [&](std::uint32_t row) -> Cell& { return rows_[row].ensure(source_col); });
    }
    return;
  }
  // A style span that reaches past the source column already covers the
  // inserted ones; only a span ending there needs a new entry.
  const auto span = std::find_if(layout_.columns.begin(), layout_.columns.end(),
                                 [&](const ColumnLayout& c) { return c.first <= source && source <= c.last; });
  if (span != layout_.columns.end() && (span->has_style || HasExplicitColumnWidth(*span)) && span->last == source) {
    ColumnLayout entry;
    entry.first = edit.index;
    entry.last = end - 1U;
    entry.has_style = span->has_style;
    entry.style_xf = span->has_style ? span->style_xf : 0U;
    entry.has_width = HasExplicitColumnWidth(*span);
    entry.width = span->width;
    layout_.columns.insert(span + 1, entry);
  }
  for (auto& kv : rows_) {
    const Cell* from = kv.second.find(source);
    if (from == nullptr || !owns_xf(*from)) {
      continue;
    }
    copy_xf(from->xf_index, [&](std::uint32_t col) -> Cell& { return kv.second.ensure(col); });
  }
}

void Sheet::insert_rows(std::uint32_t row, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  // Walk populated rows in descending key order so a moved row never
  // collides with an existing row that still needs to move.
  // Ascending keys, walked backwards, so the shared ascending sort serves this
  // direction too.
  const std::vector<std::uint32_t> keys = SortedRowKeys(rows_, row);
  for (auto it = keys.rbegin(); it != keys.rend(); ++it) {
    const std::uint32_t key = *it;
    const std::uint64_t shifted = static_cast<std::uint64_t>(key) + count;
    auto node = rows_.extract(key);
    if (shifted >= kMaxRows) {
      continue;  // Row pushed past sheet bound; drop the cells.
    }
    node.key() = static_cast<std::uint32_t>(shifted);
    rows_.insert(std::move(node));
  }
  const StructuralEdit edit{row, count, /*is_delete=*/false, /*row_axis=*/true};
  finalize_structural_edit_locked(edit);
}

void Sheet::delete_rows(std::uint32_t row, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  // Drop cells inside the deletion interval, then shift trailing rows
  // up. Walk in ascending key order — every shifted destination key is
  // strictly less than its source so no collision can occur.
  const std::vector<std::uint32_t> keys = SortedRowKeys(rows_, row);
  for (std::uint32_t key : keys) {
    auto node = rows_.extract(key);
    if (key < row + count) {
      continue;  // Deleted row; drop the node.
    }
    node.key() = key - count;
    rows_.insert(std::move(node));
  }
  const StructuralEdit edit{row, count, /*is_delete=*/true, /*row_axis=*/true};
  finalize_structural_edit_locked(edit);
}

void Sheet::insert_cols(std::uint32_t col, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  for (auto& kv : rows_) {
    kv.second.shift_cols_from(col, count);
  }
  const StructuralEdit edit{col, count, /*is_delete=*/false, /*row_axis=*/false};
  finalize_structural_edit_locked(edit);
}

void Sheet::edit_cells_in_band(const StructuralEdit& edit) {
  if (edit.count == 0U || !edit.band.has_value()) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  const Band band = *edit.band;
  if (!edit.row_axis) {
    // A horizontal shift moves each band row's tail, as a column edit moves every row's.
    for (auto& kv : rows_) {
      if (kv.first < band.first || kv.first > band.last) {
        continue;
      }
      if (edit.is_delete) {
        kv.second.erase_cols(edit.index, edit.count);
      } else {
        kv.second.shift_cols_from(edit.index, edit.count);
      }
    }
  } else {
    // A vertical shift moves the band's columns between rows: an insert from
    // the bottom up and a delete from the top down, so a destination is
    // always vacated before it is filled.
    const std::uint64_t band_end = static_cast<std::uint64_t>(band.last) + 1U;
    const std::uint64_t delete_end = static_cast<std::uint64_t>(edit.index) + edit.count;
    const auto move = [&](std::uint32_t from, std::uint64_t to, bool drop) {
      const auto it = rows_.find(from);
      if (it == rows_.end() || it->second.count_in(band.first, band_end) == 0U) {
        return;
      }
      RowCells& source = it->second;
      if (drop || to >= kMaxRows) {
        source.erase_in(band.first, band_end);
        return;
      }
      source.move_in_to(band.first, band_end, rows_[static_cast<std::uint32_t>(to)]);
    };
    const std::vector<std::uint32_t> keys = SortedRowKeys(rows_, edit.index);
    if (edit.is_delete) {
      for (const std::uint32_t key : keys) {
        move(key, key < delete_end ? key : key - edit.count, key < delete_end);
      }
    } else {
      for (auto it = keys.rbegin(); it != keys.rend(); ++it) {
        move(*it, static_cast<std::uint64_t>(*it) + edit.count, false);
      }
    }
  }
  finalize_structural_edit_locked(edit);
}

void Sheet::delete_cols(std::uint32_t col, std::uint32_t count) {
  if (count == 0U) {
    return;
  }
  const std::lock_guard<std::mutex> guard(*spill_mutex_);
  for (auto& kv : rows_) {
    kv.second.erase_cols(col, count);
  }
  const StructuralEdit edit{col, count, /*is_delete=*/true, /*row_axis=*/false};
  finalize_structural_edit_locked(edit);
}

}  // namespace formulon
