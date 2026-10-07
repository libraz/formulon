//
// The 0-based inclusive cell rectangle shared by merges, AutoFilter and sort
// ranges.

#ifndef FORMULON_MERGE_RANGE_H_
#define FORMULON_MERGE_RANGE_H_

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

}  // namespace formulon

#endif  // FORMULON_MERGE_RANGE_H_
