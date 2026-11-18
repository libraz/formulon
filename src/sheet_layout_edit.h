// Shared editors for `SheetLayout`: the column-span overlay used by every
// column setter and the emptiness tests that decide when an axis entry is
// dropped. The C ABI setters and the region style operations go through the
// same code so a span edit splits and coalesces one way everywhere.

#ifndef FORMULON_SHEET_LAYOUT_EDIT_H_
#define FORMULON_SHEET_LAYOUT_EDIT_H_

#include <algorithm>
#include <cstdint>
#include <vector>

#include "sheet_layout.h"
#include "utils/index_sort.h"

namespace formulon {

inline bool SameColumnLayoutState(const ColumnLayout& lhs, const ColumnLayout& rhs) {
  return lhs.width == rhs.width && lhs.hidden == rhs.hidden && lhs.outline_level == rhs.outline_level &&
         lhs.collapsed == rhs.collapsed && lhs.has_width == rhs.has_width && lhs.custom_width == rhs.custom_width &&
         lhs.has_style == rhs.has_style && lhs.style_xf == rhs.style_xf && lhs.show_phonetic == rhs.show_phonetic;
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

// Puts loaded column spans in the canonical form every reader of
// `SheetLayout::columns` assumes: sorted by (first, last) and disjoint. On an
// overlap the span earlier in that order keeps its columns and the later one
// is clipped to start after it; a span left empty is dropped. Equal
// neighbours are not merged.
inline void normalize_column_spans(SheetLayout& layout) {
  std::vector<ColumnLayout>& columns = layout.columns;
  const auto disjoint_ascending = [](const ColumnLayout& lhs, const ColumnLayout& rhs) { return lhs.last < rhs.first; };
  bool canonical = true;
  for (std::size_t i = 0; i < columns.size() && canonical; ++i) {
    canonical = columns[i].first <= columns[i].last && (i == 0U || disjoint_ascending(columns[i - 1U], columns[i]));
  }
  if (canonical) {
    return;
  }
  sort_by_index(columns, ColumnSpanOrder{});
  std::size_t kept = 0;
  std::uint64_t next_free = 0;
  for (std::size_t i = 0; i < columns.size(); ++i) {
    ColumnLayout span = columns[i];
    if (span.first > span.last || span.last < next_free) {
      continue;
    }
    span.first = static_cast<std::uint32_t>(std::max<std::uint64_t>(span.first, next_free));
    next_free = static_cast<std::uint64_t>(span.last) + 1U;
    columns[kept++] = span;
  }
  columns.resize(kept);
}

// Applies one column setter to every intersection with `[first, last]`.
// Existing entries are split into left residual / updated intersection / right
// residual, so fields not owned by this setter remain local to the source
// span. Any uncovered part of the target gets a default layout carrying only
// the requested field. The final sort + coalesce keeps the vector canonical
// without merging spans whose presence/style/visibility state differs.
// `apply` receives each touched span and `ctx` (the setter's value).
using ColumnSpanApply = void (*)(ColumnLayout& entry, const void* ctx);
void overlay_column_span(SheetLayout& layout, std::uint32_t first, std::uint32_t last, ColumnSpanApply apply,
                         const void* ctx);

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
