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
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "io/cf_writer.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "style_resolve.h"
#include "utils/index_sort.h"

namespace formulon {

namespace {

// Shifts every entry in `items` whose anchor lies on or past `index` by the
// insert / delete rule encoded in `count` and `is_delete`. Entries whose
// anchor falls inside the deleted interval, or whose insert would push it
// past `bound`, are removed. `field` names the axis coordinate, so one
// instantiation per anchor type serves both axes. Under a `band`, an entry
// whose `perp` coordinate lies outside it stays.
// Applies the insert / delete rule to one anchor coordinate at or past the
// edit index; false when the anchor is dropped.
bool ShiftCoord(std::uint32_t& value, std::uint32_t count, std::uint32_t index, bool is_delete, std::uint32_t bound) {
  if (is_delete) {
    if (value < index + count) {
      return false;  // Anchor inside deleted interval.
    }
    value -= count;
    return true;
  }
  const std::uint64_t shifted = static_cast<std::uint64_t>(value) + count;
  if (shifted >= bound) {
    return false;  // Pushed past the sheet bound.
  }
  value = static_cast<std::uint32_t>(shifted);
  return true;
}

template <typename Anchor>
void ShiftAnchored(std::vector<Anchor>& items, std::uint32_t Anchor::*field, std::uint32_t index, std::uint32_t count,
                   bool is_delete, std::uint32_t bound, const std::optional<Band>& band = std::nullopt,
                   std::uint32_t Anchor::*perp = nullptr) {
  std::size_t kept = 0;
  for (Anchor& item : items) {
    const bool outside =
        item.*field < index || (band.has_value() && (item.*perp < band->first || item.*perp > band->last));
    if (outside || ShiftCoord(item.*field, count, index, is_delete, bound)) {
      if (kept != static_cast<std::size_t>(&item - items.data())) {
        items[kept] = std::move(item);
      }
      ++kept;
    }
  }
  items.erase(items.begin() + static_cast<std::ptrdiff_t>(kept), items.end());
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

void ShiftAxisRange(MergeRange& range, std::uint32_t index, std::uint32_t count, bool is_delete, bool row_axis,
                    bool* out_drop) {
  if (row_axis) {
    ShiftSpan(range.first_row, range.last_row, index, count, is_delete, Sheet::kMaxRows, out_drop);
  } else {
    ShiftSpan(range.first_col, range.last_col, index, count, is_delete, Sheet::kMaxCols, out_drop);
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

// The smallest rectangle covering both `a` and `b`.
MergeRange BoundingBox(const MergeRange& a, const MergeRange& b) {
  return MergeRange{std::min(a.first_row, b.first_row), std::min(a.first_col, b.first_col),
                    std::max(a.last_row, b.last_row), std::max(a.last_col, b.last_col)};
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
    bounds = BoundingBox(bounds, r);
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

// True for a rule whose conditional format takes the rank order: an icon set
// with custom icons whatever its id, and any other rank rule without an id.
// An id-bearing rank rule and an icon set with a formula threshold keep their
// pieces in place.
bool TakesRankOrder(const cf::CFRule& rule) noexcept {
  switch (rule.type) {
    case cf::RuleType::IconSet:
      if (rule.icon_set.has_value() &&
          (rule.icon_set->floor.type == cf::CfvoType::Formula ||
           std::any_of(rule.icon_set->thresholds.begin(), rule.icon_set->thresholds.end(),
                       [](const cf::CfValueObject& v) { return v.type == cf::CfvoType::Formula; }))) {
        return false;
      }
      if (rule.icon_set.has_value() && !rule.icon_set->custom_icons.empty()) {
        return true;
      }
      return rule.id.empty();
    case cf::RuleType::Top10:
    case cf::RuleType::AboveAverage:
    case cf::RuleType::DuplicateValues:
    case cf::RuleType::UniqueValues:
      return rule.id.empty();
    default:
      return false;
  }
}

// The order a band edit leaves a conditional format's sqref pieces in: the
// value, text, blank, error and date rules share one, the rank rules another;
// any other rule keeps each range's pieces in place. An icon set whose
// threshold names a sheet other than `own_sheet` is stored in the x14
// records, which order it with the value rules.
SqrefOrder CfSqrefOrder(const std::vector<cf::CFRule>& rules, std::string_view own_sheet,
                        const std::vector<std::string>* sheet_list) noexcept {
  if (!rules.empty() && std::all_of(rules.begin(), rules.end(), TakesRankOrder)) {
    return SqrefOrder::kRankRule;
  }
  bool value_rules = !rules.empty();
  for (const cf::CFRule& rule : rules) {
    switch (rule.type) {
      case cf::RuleType::Expression:
      case cf::RuleType::CellIs:
      case cf::RuleType::ContainsText:
      case cf::RuleType::NotContainsText:
      case cf::RuleType::BeginsWith:
      case cf::RuleType::EndsWith:
      case cf::RuleType::ContainsBlanks:
      case cf::RuleType::NotContainsBlanks:
      case cf::RuleType::ContainsErrors:
      case cf::RuleType::NotContainsErrors:
      case cf::RuleType::TimePeriod:
        break;
      case cf::RuleType::IconSet:
        value_rules = value_rules && io::cf_rule_is_x14_only(rule, own_sheet, sheet_list);
        break;
      default:
        value_rules = false;
        break;
    }
  }
  return value_rules ? SqrefOrder::kValueRule : SqrefOrder::kInPlace;
}

void ShiftConditionalFormats(std::vector<cf::ConditionalFormat>& formats, const StructuralEdit& edit,
                             std::string_view own_sheet) {
  std::vector<cf::ConditionalFormat> retained_formats;
  retained_formats.reserve(formats.size());
  for (cf::ConditionalFormat& format : formats) {
    std::vector<MergeRange> ranges;
    ranges.reserve(format.sqref.size());
    for (const cf::CFCellRange& range : format.sqref) {
      ranges.push_back(MergeRange{range.first.row, range.first.col, range.last.row, range.last.col});
    }
    shift_cf_sqref_ranges(ranges, edit, CfSqrefOrder(format.rules, own_sheet, nullptr));
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
    ShiftSpan(shifted.first_col, shifted.last_col, index, count, is_delete, Sheet::kMaxCols, &drop);
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
  kShift,      // A range reaching the moving region splits into its unmoved and moved parts.
  kCf,         // `kShift`, then a delete joins what it brings into one tiled rectangle.
  kReference,  // Only a range inside the band moves.
  kCut,        // Only a range inside the band is cut.
};

// Appends `rect` minus `hole` (which lies inside it) row by row: the rows
// above, the parts left and right of it in its rows, the rows below.
void AppendDifference(const MergeRange& rect, const MergeRange& hole, std::vector<MergeRange>& out) {
  if (rect.first_row < hole.first_row) {
    out.push_back(MergeRange{rect.first_row, rect.first_col, hole.first_row - 1U, rect.last_col});
  }
  if (rect.first_col < hole.first_col) {
    out.push_back(MergeRange{hole.first_row, rect.first_col, hole.last_row, hole.first_col - 1U});
  }
  if (hole.last_col < rect.last_col) {
    out.push_back(MergeRange{hole.first_row, hole.last_col + 1U, hole.last_row, rect.last_col});
  }
  if (hole.last_row < rect.last_row) {
    out.push_back(MergeRange{hole.last_row + 1U, rect.first_col, rect.last_row, rect.last_col});
  }
}

// Appends `rect` minus `hole` (which lies inside it), cut around `around`: the
// rows above `around`, the parts left and right of it in its rows, the rows
// below. Each part is split into column runs row by row, and consecutive rows
// with the same runs make one piece.
// Whether `ranges` cover every cell of line `line` (a row when `row_axis`,
// else a column) from cross index `first` through `last`.
bool CoversLine(const std::vector<MergeRange>& ranges, bool row_axis, std::uint32_t line, std::uint32_t first,
                std::uint32_t last) noexcept {
  std::uint32_t at = first;
  for (bool advanced = true; advanced && at <= last;) {
    advanced = false;
    for (const MergeRange& r : ranges) {
      const bool on_line =
          row_axis ? r.first_row <= line && line <= r.last_row : r.first_col <= line && line <= r.last_col;
      const std::uint32_t lo = row_axis ? r.first_col : r.first_row;
      const std::uint32_t hi = row_axis ? r.last_col : r.last_row;
      if (on_line && lo <= at && at <= hi) {
        if (hi >= last) {
          return true;
        }
        at = hi + 1U;
        advanced = true;
      }
    }
  }
  return at > last;
}

void AppendDifferenceAround(const MergeRange& rect, const MergeRange& hole, const MergeRange& around,
                            std::vector<MergeRange>& out) {
  struct Run {
    std::uint32_t first;
    std::uint32_t last;
  };
  std::vector<std::uint32_t> cuts = {rect.first_row,     rect.last_row + 1U, hole.first_row,
                                     hole.last_row + 1U, around.first_row,   around.last_row + 1U};
  for (std::uint32_t& row : cuts) {
    row = std::clamp(row, rect.first_row, rect.last_row + 1U);
  }
  sort_ascending(cuts);
  cuts.erase(std::unique(cuts.begin(), cuts.end()), cuts.end());
  // The runs of row segment [first, last] between columns `lo` and `hi`.
  const auto runs = [&](std::uint32_t first, std::uint32_t last, std::uint32_t lo, std::uint32_t hi) {
    std::vector<Run> found;
    const bool holed = hole.first_row <= first && last <= hole.last_row;
    const auto add = [&](std::uint32_t a, std::uint32_t b) {
      a = std::max(a, lo);
      b = std::min(b, hi);
      if (a <= b) {
        found.push_back(Run{a, b});
      }
    };
    if (!holed) {
      add(rect.first_col, rect.last_col);
      return found;
    }
    if (rect.first_col < hole.first_col) {
      add(rect.first_col, hole.first_col - 1U);
    }
    if (hole.last_col < rect.last_col) {
      add(hole.last_col + 1U, rect.last_col);
    }
    return found;
  };
  const auto emit = [&](std::uint32_t row_first, std::uint32_t row_last, std::uint32_t lo, std::uint32_t hi) {
    std::vector<Run> open;
    std::uint32_t open_first = 0;
    std::uint32_t open_last = 0;
    const auto flush = [&]() {
      for (const Run& run : open) {
        out.push_back(MergeRange{open_first, run.first, open_last, run.last});
      }
      open.clear();
    };
    for (std::size_t i = 0; i + 1U < cuts.size(); ++i) {
      const std::uint32_t first = cuts[i];
      const std::uint32_t last = cuts[i + 1U] - 1U;
      if (first < row_first || last > row_last) {
        continue;
      }
      std::vector<Run> here = runs(first, last, lo, hi);
      const bool same = !open.empty() && here.size() == open.size() && open_last + 1U == first &&
                        std::equal(here.begin(), here.end(), open.begin(),
                                   [](const Run& a, const Run& b) { return a.first == b.first && a.last == b.last; });
      if (same) {
        open_last = last;
        continue;
      }
      flush();
      open = std::move(here);
      open_first = first;
      open_last = last;
    }
    flush();
  };
  if (rect.first_row < around.first_row) {
    emit(rect.first_row, around.first_row - 1U, rect.first_col, rect.last_col);
  }
  if (rect.first_col < around.first_col) {
    emit(around.first_row, around.last_row, rect.first_col, around.first_col - 1U);
  }
  emit(around.first_row, around.last_row, around.last_col + 1U, rect.last_col);
  if (around.last_row < rect.last_row) {
    emit(around.last_row + 1U, rect.last_row, rect.first_col, rect.last_col);
  }
}

// Whether `a` and `b` together are exactly one rectangle, edge to edge.
bool JoinsAsRectangle(const MergeRange& a, const MergeRange& b) noexcept {
  const bool same_cols = a.first_col == b.first_col && a.last_col == b.last_col;
  const bool same_rows = a.first_row == b.first_row && a.last_row == b.last_row;
  return (same_cols && (a.last_row + 1U == b.first_row || b.last_row + 1U == a.first_row)) ||
         (same_rows && (a.last_col + 1U == b.first_col || b.last_col + 1U == a.first_col));
}

// What a piece is to the edit that produced it.
enum class Piece : std::uint8_t {
  kWhole,      // A range the edit leaves where it is.
  kRemainder,  // The unmoved part of a split range.
  kMoved,      // The moved part of a split range, or a piece a join produced.
  kShifted,    // A range that only moved.
};

bool IsMoved(Piece kind) noexcept {
  return kind == Piece::kMoved || kind == Piece::kShifted;
}

// How a band edit arranges the pieces of the ranges it splits or cuts.
enum class Arrangement : std::uint8_t {
  kInPlace,  // Each range's pieces where the range was; a join takes the earlier piece's place.
  kFront,    // As `kInPlace`, but a join takes the moved piece's place, and the first range that only
             // moved, unjoined, then goes to the front.
  kToEnd,    // After the other ranges, in order; a join takes the moved piece's place.
};

// Joins, pairwise until none is left, a moved piece with another moved piece
// or a remainder when their union is a rectangle; the joined piece is a moved
// piece; a range the edit leaves alone joins nothing. Under `kInPlace` the
// joined piece takes the earlier one's place, else the moved piece's (the
// earlier one's when both moved).
void JoinSplitPieces(std::vector<MergeRange>& ranges, std::vector<Piece>& kinds, Arrangement arrangement) {
  bool joined = true;
  while (joined) {
    joined = false;
    for (std::size_t i = 0; !joined && i < ranges.size(); ++i) {
      for (std::size_t j = i + 1U; !joined && kinds[i] != Piece::kWhole && j < ranges.size(); ++j) {
        const bool moved = IsMoved(kinds[i]) || IsMoved(kinds[j]);
        if (!moved || kinds[j] == Piece::kWhole || !JoinsAsRectangle(ranges[i], ranges[j])) {
          continue;
        }
        const bool later = arrangement != Arrangement::kInPlace && !IsMoved(kinds[i]);
        const std::size_t kept = later ? j : i;
        const std::size_t gone = later ? i : j;
        ranges[kept] = BoundingBox(ranges[i], ranges[j]);
        kinds[kept] = Piece::kMoved;
        ranges.erase(ranges.begin() + static_cast<std::ptrdiff_t>(gone));
        kinds.erase(kinds.begin() + static_cast<std::ptrdiff_t>(gone));
        joined = true;
      }
    }
  }
}

// Joins, pairwise until none is left, two remainders whose union is a
// rectangle; the joined piece takes the earlier one's place.
void FuseRemainders(std::vector<MergeRange>& ranges, std::vector<Piece>& kinds) {
  bool fused = true;
  while (fused) {
    fused = false;
    for (std::size_t i = 0; !fused && i < ranges.size(); ++i) {
      for (std::size_t j = i + 1U; !fused && kinds[i] == Piece::kRemainder && j < ranges.size(); ++j) {
        if (kinds[j] != Piece::kRemainder || !JoinsAsRectangle(ranges[i], ranges[j])) {
          continue;
        }
        ranges[i] = BoundingBox(ranges[i], ranges[j]);
        ranges.erase(ranges.begin() + static_cast<std::ptrdiff_t>(j));
        kinds.erase(kinds.begin() + static_cast<std::ptrdiff_t>(j));
        fused = true;
      }
    }
  }
}

// The inserted cells of an insert `edit` across `band` that join `ranges` beside the pieces already made (`made`),
// one list per range: a cell joins when the cell on the line before the insert, in its column (row), was in the range.
std::vector<std::vector<MergeRange>> InheritedCells(const std::vector<MergeRange>& ranges, const StructuralEdit& edit,
                                                    const Band& band,
                                                    const std::vector<const std::vector<MergeRange>*>& made) {
  const std::uint32_t line = edit.index - 1U;
  std::vector<MergeRange> taken;
  for (const std::vector<MergeRange>* list : made) {
    taken.insert(taken.end(), list->begin(), list->end());
  }
  std::vector<std::vector<MergeRange>> inherited(ranges.size());
  for (std::size_t i = 0; i < ranges.size(); ++i) {
    const MergeRange& r = ranges[i];
    const bool on_line =
        edit.row_axis ? r.first_row <= line && line <= r.last_row : r.first_col <= line && line <= r.last_col;
    const std::uint32_t lo = std::max(edit.row_axis ? r.first_col : r.first_row, band.first);
    const std::uint32_t hi = std::min(edit.row_axis ? r.last_col : r.last_row, band.last);
    if (!on_line || lo > hi) {
      continue;
    }
    const std::uint32_t last_line = edit.index + edit.count - 1U;
    std::vector<MergeRange> cells{edit.row_axis ? MergeRange{edit.index, lo, last_line, hi}
                                                : MergeRange{lo, edit.index, hi, last_line}};
    for (const MergeRange& hole : taken) {
      std::vector<MergeRange> rest;
      for (const MergeRange& cell : cells) {
        const MergeRange cut{std::max(cell.first_row, hole.first_row), std::max(cell.first_col, hole.first_col),
                             std::min(cell.last_row, hole.last_row), std::min(cell.last_col, hole.last_col)};
        if (cut.first_row > cut.last_row || cut.first_col > cut.last_col) {
          rest.push_back(cell);
        } else {
          AppendDifference(cell, cut, rest);
        }
      }
      cells = std::move(rest);
    }
    taken.insert(taken.end(), cells.begin(), cells.end());
    inherited[i] = std::move(cells);
  }
  return inherited;
}

void EditSqrefRanges(std::vector<MergeRange>& ranges, const StructuralEdit& edit, SqrefMode mode, SqrefOrder order,
                     bool grow) {
  if (!edit.band.has_value()) {
    if (mode == SqrefMode::kCut) {
      cut_sqref_ranges(ranges, edit.index, edit.count, edit.is_delete, edit.row_axis);
    } else if (grow && !edit.is_delete && edit.index > 0U) {
      // An insert moves each range whole; the inserted cells a range inherits follow it.
      std::vector<MergeRange> out;
      std::vector<std::size_t> ends;
      for (MergeRange range : ranges) {
        bool drop = false;
        ShiftAxisRange(range, edit.index, edit.count, false, edit.row_axis, &drop);
        if (!drop) {
          out.push_back(range);
        }
        ends.push_back(out.size());
      }
      const Band whole{0U, (edit.row_axis ? Sheet::kMaxCols : Sheet::kMaxRows) - 1U};
      const std::vector<std::vector<MergeRange>> inherited = InheritedCells(ranges, edit, whole, {&out});
      for (std::size_t i = ranges.size(); i-- > 0U;) {
        out.insert(out.begin() + static_cast<std::ptrdiff_t>(ends[i]), inherited[i].begin(), inherited[i].end());
      }
      ranges = std::move(out);
    } else {
      shift_sqref_ranges(ranges, edit.index, edit.count, edit.is_delete, edit.row_axis);
    }
    return;
  }
  const bool join_tiles = mode == SqrefMode::kCf && edit.is_delete && !TilesOneRectangle(ranges);
  const bool splits = mode == SqrefMode::kShift || mode == SqrefMode::kCf;
  const Band band = *edit.band;
  const bool several_ranges = ranges.size() > 1U;
  const Arrangement arrangement = order == SqrefOrder::kInPlace    ? Arrangement::kInPlace
                                  : order == SqrefOrder::kRankRule ? Arrangement::kFront
                                                                   : Arrangement::kToEnd;
  std::vector<MergeRange> out;
  out.reserve(ranges.size() * 3U);
  std::vector<Piece> kinds;
  // Under `kToEnd`, the pieces of the ranges the edit splits or cuts.
  std::vector<MergeRange> edited;
  std::vector<Piece> edited_kinds;
  // Where each range's pieces end: in `edited` or `out`, and at which size.
  std::vector<std::pair<bool, std::size_t>> ends;
  ends.reserve(ranges.size());
  for (const MergeRange& range : ranges) {
    const BandFit fit = band_fit(range, edit);
    if (fit == BandFit::kOutside || (fit == BandFit::kStraddles && !splits)) {
      out.push_back(range);
      kinds.resize(out.size(), Piece::kWhole);
      ends.emplace_back(false, out.size());
      continue;
    }
    if (mode == SqrefMode::kCut) {
      CutOneRange(range, edit.index, edit.count, edit.is_delete, edit.row_axis, out);
      kinds.resize(out.size(), Piece::kWhole);
      ends.emplace_back(false, out.size());
      continue;
    }
    // The part in the moving region (the band from the edit's edge on) moves;
    // the rest of the range stays where it is.
    MergeRange moving = range;
    std::uint32_t& axial_first = edit.row_axis ? moving.first_row : moving.first_col;
    axial_first = std::max(axial_first, edit.index);
    (edit.row_axis ? moving.first_col : moving.first_row) =
        std::max(edit.row_axis ? range.first_col : range.first_row, band.first);
    (edit.row_axis ? moving.last_col : moving.last_row) =
        std::min(edit.row_axis ? range.last_col : range.last_row, band.last);
    const MergeRange source = moving;
    const bool cuts = edit.is_delete && axial_first < edit.index + edit.count;
    bool drop = false;
    ShiftAxisRange(moving, edit.index, edit.count, edit.is_delete, edit.row_axis, &drop);
    // An insert inside the element leaves the inserted cells covered: inside
    // the range, or on its edge when the line before is covered too.
    if (!edit.is_delete) {
      const std::uint32_t range_first = edit.row_axis ? range.first_row : range.first_col;
      const bool inside = range_first < edit.index || (range_first == edit.index && edit.index > 0U &&
                                                       CoversLine(ranges, edit.row_axis, edit.index - 1U,
                                                                  edit.row_axis ? moving.first_col : moving.first_row,
                                                                  edit.row_axis ? moving.last_col : moving.last_row));
      if (inside) {
        axial_first = edit.index;
      }
    }
    // A moved part one cell deep along the shift re-cuts the unmoved part
    // around where it lands, except in a validation of several ranges;
    // anything else is cut around the moving region.
    const bool one_deep = edit.row_axis ? moving.first_row == moving.last_row : moving.first_col == moving.last_col;
    const bool recut = !drop && arrangement != Arrangement::kInPlace && one_deep &&
                       !(order == SqrefOrder::kValidation && several_ranges);
    std::vector<MergeRange> remainder;
    if (recut) {
      AppendDifferenceAround(range, source, moving, remainder);
    } else {
      AppendDifference(range, source, remainder);
    }
    const bool splits_range = !remainder.empty();
    const bool to_end = arrangement == Arrangement::kToEnd && (splits_range || cuts);
    std::vector<MergeRange>& target = to_end ? edited : out;
    std::vector<Piece>& target_kinds = to_end ? edited_kinds : kinds;
    // A validation of several ranges puts each range's moved part before its unmoved parts.
    const bool moved_first = !drop && to_end && order == SqrefOrder::kValidation && several_ranges;
    if (moved_first) {
      target.push_back(moving);
      target_kinds.push_back(Piece::kMoved);
    }
    target.insert(target.end(), remainder.begin(), remainder.end());
    target_kinds.resize(target.size(), Piece::kRemainder);
    if (!drop && !moved_first) {
      target.push_back(moving);
      target_kinds.push_back(splits_range || cuts ? Piece::kMoved : Piece::kShifted);
    }
    ends.emplace_back(to_end, target.size());
  }
  // Inserted cells a range inherits follow that range's pieces; later ranges first, so the recorded ends hold.
  if (grow && !edit.is_delete && edit.index > 0U) {
    const std::vector<std::vector<MergeRange>> inherited = InheritedCells(ranges, edit, band, {&out, &edited});
    for (std::size_t i = ranges.size(); i-- > 0U;) {
      std::vector<MergeRange>& target = ends[i].first ? edited : out;
      std::vector<Piece>& target_kinds = ends[i].first ? edited_kinds : kinds;
      const auto at = static_cast<std::ptrdiff_t>(ends[i].second);
      target.insert(target.begin() + at, inherited[i].begin(), inherited[i].end());
      target_kinds.insert(target_kinds.begin() + at, inherited[i].size(), Piece::kMoved);
    }
  }
  out.insert(out.end(), edited.begin(), edited.end());
  kinds.insert(kinds.end(), edited_kinds.begin(), edited_kinds.end());
  JoinSplitPieces(out, kinds, arrangement);
  if (arrangement != Arrangement::kInPlace) {
    FuseRemainders(out, kinds);
  }
  if (arrangement == Arrangement::kFront) {
    const auto first = std::find(kinds.begin(), kinds.end(), Piece::kShifted);
    if (first != kinds.end()) {
      const auto at = out.begin() + (first - kinds.begin());
      std::rotate(out.begin(), at, at + 1);
    }
  }
  ranges = std::move(out);
  if (!join_tiles || !TilesOneRectangle(ranges)) {
    return;
  }
  MergeRange joined = ranges.front();
  for (const MergeRange& r : ranges) {
    joined = BoundingBox(joined, r);
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

void shift_sqref_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit, SqrefOrder order) {
  EditSqrefRanges(ranges, edit, SqrefMode::kShift, order, /*grow=*/false);
}

void shift_cf_sqref_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit, SqrefOrder order) {
  EditSqrefRanges(ranges, edit, SqrefMode::kCf, order, /*grow=*/true);
}

void shift_validation_sqref_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit) {
  EditSqrefRanges(ranges, edit, SqrefMode::kShift, SqrefOrder::kValidation, /*grow=*/true);
}

SqrefOrder cf_sqref_order(const cf::ConditionalFormat& format, std::string_view own_sheet,
                          const std::vector<std::string>* sheet_list) noexcept {
  return CfSqrefOrder(format.rules, own_sheet, sheet_list);
}

void shift_reference_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit) {
  EditSqrefRanges(ranges, edit, SqrefMode::kReference, SqrefOrder::kInPlace, /*grow=*/false);
}

void cut_sqref_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit) {
  EditSqrefRanges(ranges, edit, SqrefMode::kCut, SqrefOrder::kInPlace, /*grow=*/false);
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
  // A band edit moves only part of a row, so the address order is rebuilt.
  ShiftCellAnchored(rich_values_, edit);
  sort_by_index(rich_values_, [](const CellRichValue& a, const CellRichValue& b) {
    return a.row != b.row ? a.row < b.row : a.col < b.col;
  });
  rich_value_count_.set(rich_values_.size());
  ShiftMerges(merges_, edit);
  ShiftDataTables(data_tables_, edit);
  for (DataValidation& dv : validations_) {
    if (edit.band.has_value() || !is_delete) {
      shift_validation_sqref_ranges(dv.ranges, edit);
    } else {
      ShiftRangeList(dv.ranges, index, count, is_delete, row_axis);
    }
  }
  validations_.erase(std::remove_if(validations_.begin(), validations_.end(),
                                    [](const DataValidation& dv) { return dv.ranges.empty(); }),
                     validations_.end());
  ShiftConditionalFormats(conditional_formats_, edit, name_);
  if (!edit.band.has_value() && row_axis) {
    ShiftAnchored(layout_.row_overrides, &RowLayout::row, index, count, is_delete, Sheet::kMaxRows);
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
  // The inserted coordinates are vacant, so each takes its neighbour's own
  // format as a style run.
  const auto owns_xf = [](const StyleSlot& slot) { return slot.explicit_xf || slot.xf != 0U; };
  struct Piece {
    std::uint32_t first;
    std::uint64_t end;
    std::uint32_t xf;
  };
  // The own-format pieces of row `source` inside `[first, stop)`.
  const auto source_row_pieces = [&](std::uint32_t first, std::uint64_t stop) {
    std::vector<Piece> pieces;
    const auto it = rows_.find(source);
    if (it == rows_.end()) {
      return pieces;
    }
    for (const RowCells::Entry entry : it->second.entries_in(first, stop)) {
      if (owns_xf(StyleSlot{entry.cell.xf_index, entry.cell.has_explicit_xf})) {
        pieces.push_back({entry.col, static_cast<std::uint64_t>(entry.col) + 1U, entry.cell.xf_index});
      }
    }
    const auto [lo, hi] = it->second.runs_in(first, stop);
    for (const StyleRun* run = lo; run != hi; ++run) {
      pieces.push_back({std::max<std::uint32_t>(run->first, first),
                        std::min<std::uint64_t>(static_cast<std::uint64_t>(run->last) + 1U, stop), run->xf});
    }
    return pieces;
  };
  // (row, xf) of every row in `[first_row, last_row]` whose column `source` owns a format.
  const auto source_col_styles = [&](std::uint32_t first_row, std::uint32_t last_row) {
    std::vector<std::pair<std::uint32_t, std::uint32_t>> styled;
    for (const auto& kv : rows_) {
      if (kv.first < first_row || kv.first > last_row) {
        continue;
      }
      if (const std::optional<StyleSlot> slot = kv.second.find_style(source); slot && owns_xf(*slot)) {
        styled.emplace_back(kv.first, slot->xf);
      }
    }
    return styled;
  };
  if (edit.band.has_value()) {
    // Insert Cells: each inserted cell takes the XF of the cell above (row
    // axis) or to the left; the row and column styles stay where they are.
    if (edit.row_axis) {
      for (const Piece& piece : source_row_pieces(edit.band->first, static_cast<std::uint64_t>(edit.band->last) + 1U)) {
        for (std::uint32_t row = edit.index; row < end; ++row) {
          apply_style_runs_locked(row, piece.first, piece.end, constant_xf_map(piece.xf));
        }
      }
      return;
    }
    for (const auto& [row, xf_index] : source_col_styles(edit.band->first, edit.band->last)) {
      apply_style_runs_locked(row, edit.index, end, constant_xf_map(xf_index));
    }
    return;
  }
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
    for (const Piece& piece : source_row_pieces(0U, kMaxCols)) {
      for (std::uint32_t row = edit.index; row < end; ++row) {
        apply_style_runs_locked(row, piece.first, piece.end, constant_xf_map(piece.xf));
      }
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
    entry.custom_width = span->custom_width;
    entry.width = span->width;
    layout_.columns.insert(span + 1, entry);
  }
  for (const auto& [row, xf_index] : source_col_styles(0U, kMaxRows - 1U)) {
    apply_style_runs_locked(row, edit.index, end, constant_xf_map(xf_index));
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
  // The blank an insert shifts into the edge cell keeps the style it showed;
  // where the edge's row or column style differs, it becomes a cell record.
  std::vector<std::pair<std::uint32_t, std::uint32_t>> edge_xfs;  // (band line, xf)
  const std::uint32_t edge = a1::axis_extent(edit.row_axis) - 1U;
  if (!edit.is_delete && static_cast<std::uint64_t>(edit.index) + edit.count <= edge) {
    const std::uint32_t from = edge - edit.count;
    for (std::uint32_t line = band.first; line <= band.last; ++line) {
      const std::uint32_t row = edit.row_axis ? from : line;
      const std::uint32_t col = edit.row_axis ? line : from;
      if (cell_style_at_locked(row, col).has_value()) {
        continue;
      }
      const std::uint32_t shown = select_axis_xf(layout_, row, col).xf_index;
      if (shown != select_axis_xf(layout_, edit.row_axis ? edge : line, edit.row_axis ? line : edge).xf_index) {
        edge_xfs.emplace_back(line, shown);
      }
    }
  }
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
      if (it == rows_.end() || !it->second.any_in(band.first, band_end)) {
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
  for (const auto& [line, xf_index] : edge_xfs) {
    const std::uint32_t row = edit.row_axis ? edge : line;
    const std::uint32_t col = edit.row_axis ? line : edge;
    apply_style_runs_locked(row, col, static_cast<std::uint64_t>(col) + 1U, constant_xf_map(xf_index));
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
