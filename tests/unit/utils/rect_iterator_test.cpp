//
// Unit tests for `formulon::utils::RectRange`. The iterator is the shared
// replacement for the hand-rolled `for row × for col` double loops scattered
// across cf/, eval/, and io/, so empty-range, single-cell, and row-major
// ordering invariants matter for every consumer.

#include "utils/rect_iterator.h"

#include <cstdint>
#include <utility>
#include <vector>

#include "gtest/gtest.h"

namespace formulon::utils {
namespace {

// Helper: walk the range and snapshot every yielded `(row, col)` pair into a
// vector. Tests then assert against the exact expected sequence.
std::vector<std::pair<std::uint32_t, std::uint32_t>> Collect(const RectRange& range) {
  std::vector<std::pair<std::uint32_t, std::uint32_t>> out;
  for (auto [row, col] : range) {
    out.emplace_back(row, col);
  }
  return out;
}

TEST(RectRange, EmptyWhenRowsInverted) {
  // r0 > r1 must collapse to the empty range so call sites need no extra
  // empty checks before iterating.
  RectRange range(5U, 0U, 3U, 4U);
  EXPECT_TRUE(range.empty());
  EXPECT_EQ(range.size(), 0U);
  EXPECT_EQ(range.begin(), range.end());
  EXPECT_TRUE(Collect(range).empty());
}

TEST(RectRange, EmptyWhenColsInverted) {
  // Mirror of the rows-inverted case; both axes must short-circuit
  // independently so neither order silently iterates.
  RectRange range(0U, 7U, 4U, 2U);
  EXPECT_TRUE(range.empty());
  EXPECT_EQ(range.size(), 0U);
  EXPECT_EQ(range.begin(), range.end());
  EXPECT_TRUE(Collect(range).empty());
}

TEST(RectRange, SingleCell) {
  // The smallest non-empty range. `<=` bounds mean (r, r, c, c) must yield
  // exactly one cell; this is the regression case for accidentally writing
  // `<` instead of `<=`.
  RectRange range(3U, 7U, 3U, 7U);
  EXPECT_FALSE(range.empty());
  EXPECT_EQ(range.size(), 1U);
  const auto cells = Collect(range);
  ASSERT_EQ(cells.size(), 1U);
  EXPECT_EQ(cells[0].first, 3U);
  EXPECT_EQ(cells[0].second, 7U);
}

TEST(RectRange, RowMajorOrder2x3) {
  // 2 rows × 3 cols must produce the canonical row-major sequence. The
  // assertion fixes the contract: every consumer (cf rule evaluation,
  // recalc seeding, dep extraction) relies on this exact order.
  RectRange range(10U, 4U, 11U, 6U);
  EXPECT_EQ(range.size(), 6U);
  const auto cells = Collect(range);
  const std::vector<std::pair<std::uint32_t, std::uint32_t>> expected = {
      {10U, 4U}, {10U, 5U}, {10U, 6U},  //
      {11U, 4U}, {11U, 5U}, {11U, 6U},
  };
  EXPECT_EQ(cells, expected);
}

TEST(RectRange, SingleRowSpan) {
  // 1×N walks must visit only the outer (row) value once and step the
  // column. Guards against an `operator++` that wraps to the next row at
  // the wrong column boundary.
  RectRange range(2U, 0U, 2U, 3U);
  EXPECT_EQ(range.size(), 4U);
  const auto cells = Collect(range);
  const std::vector<std::pair<std::uint32_t, std::uint32_t>> expected = {
      {2U, 0U},
      {2U, 1U},
      {2U, 2U},
      {2U, 3U},
  };
  EXPECT_EQ(cells, expected);
}

TEST(RectRange, SingleColumnSpan) {
  // N×1 walks: column stays pinned, row advances. Symmetric guard to
  // `SingleRowSpan` for the row-bump path in `operator++`.
  RectRange range(4U, 9U, 7U, 9U);
  EXPECT_EQ(range.size(), 4U);
  const auto cells = Collect(range);
  const std::vector<std::pair<std::uint32_t, std::uint32_t>> expected = {
      {4U, 9U},
      {5U, 9U},
      {6U, 9U},
      {7U, 9U},
  };
  EXPECT_EQ(cells, expected);
}

TEST(RectRange, SizeDoesNotOverflowOnFullSheet) {
  // Excel's sheet max is 1,048,576 × 16,384 = ~17e9 cells, which overflows
  // uint32_t. `size()` must return uint64_t so callers reserving capacity
  // do not silently truncate.
  constexpr std::uint32_t kMaxRow = 1048576U - 1U;
  constexpr std::uint32_t kMaxCol = 16384U - 1U;
  RectRange range(0U, 0U, kMaxRow, kMaxCol);
  const std::uint64_t expected = static_cast<std::uint64_t>(1048576U) * 16384U;
  EXPECT_EQ(range.size(), expected);
}

TEST(RectRange, ConstexprUsable) {
  // The whole adapter must be evaluable at compile time so callers can
  // `static_assert` on dimensions if they want; this test makes the
  // contract visible.
  constexpr RectRange range(0U, 0U, 1U, 1U);
  static_assert(!range.empty(), "non-empty 2x2 range");
  static_assert(range.size() == 4U, "2x2 range yields 4 cells");
  EXPECT_EQ(range.size(), 4U);
}

TEST(RectRange, IteratorPostIncrement) {
  // Post-increment must return the pre-step value but still advance the
  // iterator. Range-`for` only uses pre-increment, but the iterator is
  // public API so post-increment must behave correctly too.
  RectRange range(0U, 0U, 0U, 2U);
  auto it = range.begin();
  const auto first = *it;
  const auto returned = *(it++);
  EXPECT_EQ(first.row, returned.row);
  EXPECT_EQ(first.col, returned.col);
  EXPECT_NE(it, range.begin());
  const auto second = *it;
  EXPECT_EQ(second.row, 0U);
  EXPECT_EQ(second.col, 1U);
}

}  // namespace
}  // namespace formulon::utils
