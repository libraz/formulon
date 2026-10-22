// Shared editors for `SheetLayout`: the column-span overlay used by every
// column setter and the emptiness tests that decide when an axis entry is
// dropped. The C ABI setters and the region style operations go through the
// same code so a span edit splits and coalesces one way everywhere.

#ifndef FORMULON_SHEET_LAYOUT_EDIT_H_
#define FORMULON_SHEET_LAYOUT_EDIT_H_

#include <algorithm>
#include <cstdint>
#include <limits>
#include <utility>
#include <vector>

#include "sheet_layout.h"
#include "utils/index_sort.h"

namespace formulon {

inline bool SameColumnLayoutState(const ColumnLayout& lhs, const ColumnLayout& rhs) {
  return lhs.width == rhs.width && lhs.hidden == rhs.hidden && lhs.outline_level == rhs.outline_level &&
         lhs.collapsed == rhs.collapsed && lhs.has_width == rhs.has_width && lhs.has_style == rhs.has_style &&
         lhs.style_xf == rhs.style_xf && lhs.show_phonetic == rhs.show_phonetic;
}

// Orders column spans by start, then by end. This lives at namespace scope
// rather than inside `overlay_column_span` on purpose: a comparator declared
// inside a template has a distinct closure type per instantiation, so every
// setter would otherwise reach the shared sort through its own trampoline.
struct ColumnSpanOrder {
  bool operator()(const ColumnLayout& lhs, const ColumnLayout& rhs) const {
    if (lhs.first != rhs.first) {
      return lhs.first < rhs.first;
    }
    return lhs.last < rhs.last;
  }
};

// Applies one column setter to every intersection with `[first, last]`.
// Existing entries are split into left residual / updated intersection / right
// residual, so fields not owned by this setter remain local to the source
// span. Any uncovered part of the target gets a default layout carrying only
// the requested field. The final sort + coalesce keeps the vector canonical
// without merging spans whose presence/style/visibility state differs.
template <typename Apply>
void overlay_column_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last, Apply&& apply) {
  std::vector<ColumnLayout> next;
  next.reserve(layout.columns.size() + 1U);
  std::vector<std::pair<std::uint32_t, std::uint32_t>> covered;
  covered.reserve(layout.columns.size());

  for (const ColumnLayout& entry : layout.columns) {
    if (entry.last < first || entry.first > last) {
      next.push_back(entry);
      continue;
    }

    if (entry.first < first) {
      ColumnLayout left = entry;
      left.last = first - 1U;
      next.push_back(left);
    }

    const std::uint32_t intersection_first = std::max(entry.first, first);
    const std::uint32_t intersection_last = std::min(entry.last, last);
    ColumnLayout intersection = entry;
    intersection.first = intersection_first;
    intersection.last = intersection_last;
    apply(intersection);
    next.push_back(intersection);
    covered.emplace_back(intersection_first, intersection_last);

    if (entry.last > last) {
      ColumnLayout right = entry;
      right.first = last + 1U;
      next.push_back(right);
    }
  }

  // Merge covered intervals before synthesising gaps. Normal SheetLayout
  // entries are disjoint, but this also avoids duplicating a default segment
  // when a caller supplied overlapping aggregate entries directly.
  sort_by_index(covered, [](const std::pair<std::uint32_t, std::uint32_t>& lhs,
                            const std::pair<std::uint32_t, std::uint32_t>& rhs) { return lhs < rhs; });
  std::vector<std::pair<std::uint32_t, std::uint32_t>> covered_union;
  for (const auto& interval : covered) {
    if (covered_union.empty() ||
        static_cast<std::uint64_t>(interval.first) > static_cast<std::uint64_t>(covered_union.back().second) + 1U) {
      covered_union.push_back(interval);
    } else if (interval.second > covered_union.back().second) {
      covered_union.back().second = interval.second;
    }
  }

  std::uint64_t cursor = first;
  for (const auto& interval : covered_union) {
    if (cursor < interval.first) {
      ColumnLayout gap;
      gap.first = static_cast<std::uint32_t>(cursor);
      gap.last = interval.first - 1U;
      apply(gap);
      next.push_back(gap);
    }
    cursor = std::max(cursor, static_cast<std::uint64_t>(interval.second) + 1U);
  }
  if (cursor <= last) {
    ColumnLayout gap;
    gap.first = static_cast<std::uint32_t>(cursor);
    gap.last = last;
    apply(gap);
    next.push_back(gap);
  }

  sort_by_index(next, ColumnSpanOrder{});
  std::vector<ColumnLayout> merged;
  merged.reserve(next.size());
  for (const ColumnLayout& entry : next) {
    if (!merged.empty() && merged.back().last != std::numeric_limits<std::uint32_t>::max() &&
        merged.back().last + 1U == entry.first && SameColumnLayoutState(merged.back(), entry)) {
      merged.back().last = entry.last;
    } else {
      merged.push_back(entry);
    }
  }
  layout.columns = std::move(merged);
}

// Drops column spans left with no state (including empty ones synthesised for
// columns that had no entry); they carry nothing to save.
inline void drop_stateless_columns(SheetLayout& layout) {
  auto& columns = layout.columns;
  columns.erase(std::remove_if(columns.begin(), columns.end(),
                               [](const ColumnLayout& entry) {
                                 return !HasExplicitColumnWidth(entry) && !entry.hidden && entry.outline_level == 0U &&
                                        !entry.collapsed && !entry.has_style && !entry.show_phonetic;
                               }),
                columns.end());
}

inline bool row_override_is_empty(const RowLayout& entry) {
  return !entry.has_height && !entry.custom_height && !entry.hidden && entry.outline_level == 0U && !entry.collapsed &&
         !entry.has_style && !entry.show_phonetic;
}

// Intrinsic emptiness does not account for zeroHeight: an existing empty row
// with hidden=false is an explicit visible exception to the sheet default.
// Keep that exception when an operation removes its last unrelated field.
inline bool should_erase_row_override(const RowLayout& entry, bool default_hidden, bool existed_before) {
  return row_override_is_empty(entry) && (!default_hidden || !existed_before);
}

inline std::vector<RowLayout>::iterator row_override_lower_bound(SheetLayout& layout, std::uint32_t row) {
  return std::lower_bound(layout.row_overrides.begin(), layout.row_overrides.end(), row,
                          [](const RowLayout& entry, std::uint32_t key) { return entry.row < key; });
}

inline std::vector<RowLayout>::const_iterator row_override_lower_bound(const SheetLayout& layout, std::uint32_t row) {
  return std::lower_bound(layout.row_overrides.begin(), layout.row_overrides.end(), row,
                          [](const RowLayout& entry, std::uint32_t key) { return entry.row < key; });
}

/// The entry for `row`, or nullptr.
inline const RowLayout* find_row_override(const SheetLayout& layout, std::uint32_t row) {
  const auto it = row_override_lower_bound(layout, row);
  return (it != layout.row_overrides.end() && it->row == row) ? &*it : nullptr;
}

/// Stores `entry` in row order unless the row already has one (a reader keeps
/// the first of duplicated rows). Appending in ascending order stays O(1).
inline void insert_row_override(SheetLayout& layout, const RowLayout& entry) {
  auto& rows = layout.row_overrides;
  if (rows.empty() || rows.back().row < entry.row) {
    rows.push_back(entry);
    return;
  }
  const auto it = row_override_lower_bound(layout, entry.row);
  if (it == rows.end() || it->row != entry.row) {
    rows.insert(rows.begin() + (it - rows.cbegin()), entry);
  }
}

/// The entry for `row`, inserted in order when absent with no state but the
/// sheet's default visibility `initial_hidden`. The reference is valid until
/// the next change to `row_overrides`.
inline RowLayout& upsert_row_override(SheetLayout& layout, std::uint32_t row, bool initial_hidden) {
  RowLayout fresh;
  fresh.row = row;
  fresh.hidden = initial_hidden;
  insert_row_override(layout, fresh);
  auto& rows = layout.row_overrides;
  return rows[static_cast<std::size_t>(row_override_lower_bound(layout, row) - rows.begin())];
}

}  // namespace formulon

#endif  // FORMULON_SHEET_LAYOUT_EDIT_H_
