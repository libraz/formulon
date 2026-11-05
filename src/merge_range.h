//
// The 0-based inclusive cell rectangle shared by merges, AutoFilter and sort
// ranges.

#ifndef FORMULON_MERGE_RANGE_H_
#define FORMULON_MERGE_RANGE_H_

#include <algorithm>
#include <cstdint>

namespace formulon {

/// One `<mergeCell ref="A1:B2"/>` block: a rectangular merged range
/// stored as 0-based, inclusive `(row, col)` corners. The two corners
/// are normalised so that `first_row <= last_row` and
/// `first_col <= last_col`; degenerate `first == last` rectangles are
/// permitted (single-cell merges, which Excel emits). Also the rectangle
/// type of AutoFilter and sort ranges.
struct MergeRange {
  std::uint32_t first_row = 0;
  std::uint32_t first_col = 0;
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;

  friend bool operator==(const MergeRange& a, const MergeRange& b) noexcept {
    return a.first_row == b.first_row && a.first_col == b.first_col && a.last_row == b.last_row &&
           a.last_col == b.last_col;
  }
};

/// Orders both corners without changing the rectangle they cover.
inline constexpr MergeRange normalized(MergeRange range) noexcept {
  return MergeRange{std::min(range.first_row, range.last_row), std::min(range.first_col, range.last_col),
                    std::max(range.first_row, range.last_row), std::max(range.first_col, range.last_col)};
}

/// True when the two rectangles share at least one cell.
inline constexpr bool intersects(const MergeRange& a, const MergeRange& b) noexcept {
  return a.first_row <= b.last_row && b.first_row <= a.last_row && a.first_col <= b.last_col &&
         b.first_col <= a.last_col;
}

/// True when `inner` lies wholly inside `outer`.
inline constexpr bool contains(const MergeRange& outer, const MergeRange& inner) noexcept {
  return outer.first_row <= inner.first_row && inner.last_row <= outer.last_row && outer.first_col <= inner.first_col &&
         inner.last_col <= outer.last_col;
}

}  // namespace formulon

#endif  // FORMULON_MERGE_RANGE_H_
