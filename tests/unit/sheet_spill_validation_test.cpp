// Sheet spill shape validation and death tests.

#include <algorithm>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace {
// Bounds and shape rejection are assertion-based in debug builds and
// return-false checks in release builds; retain both paths exactly.
#if defined(NDEBUG)
TEST(SheetSpillTest, ZeroRowsOrZeroColsIsRejected) {
  Sheet s("Sheet1");
  EXPECT_FALSE(s.commit_spill(0U, 0U, 0U, 1U, std::vector<Value>{}));
  EXPECT_FALSE(s.commit_spill(0U, 0U, 1U, 0U, std::vector<Value>{}));
  EXPECT_EQ(s.spill_region_at_anchor(0U, 0U), nullptr);
}
TEST(SheetSpillTest, MismatchedCellsLengthIsRejected) {
  Sheet s("Sheet1");
  std::vector<Value> wrong_size = {Value::number(1.0), Value::number(2.0)};  // expected 3
  EXPECT_FALSE(s.commit_spill(0U, 0U, 3U, 1U, std::move(wrong_size)));
  EXPECT_EQ(s.spill_region_at_anchor(0U, 0U), nullptr);
}
TEST(SheetSpillTest, FootprintOverTheDynamicArrayCellCeilingIsRejected) {
  Sheet s("Sheet1");
  // A whole-column rectangle is 1,048,576 cells and commits. A hundred of
  // them side by side is past the ceiling the evaluator's array allocator
  // applies, and the sheet refuses it on its own rather than trusting the
  // caller to have checked.
  //
  // This asserts the refusal, not which check produced it: a payload
  // matching a shape that large cannot be built in a test, so the refusal
  // here is indistinguishable from the payload-length one. The assertion
  // that separates them is the death test below, which matches the message,
  // and it is the debug build that runs it.
  EXPECT_TRUE(s.commit_spill(0U, 0U, Sheet::kMaxRows, 1U, std::vector<Value>(Sheet::kMaxRows, Value::number(1.0))));
  EXPECT_FALSE(s.commit_spill(0U, 1U, Sheet::kMaxRows, 100U, std::vector<Value>{}));
  EXPECT_EQ(s.spill_region_at_anchor(0U, 1U), nullptr);
  // Nor is the refusal recorded as a blocked footprint to retry: the shape
  // can never become admissible, so there is nothing for the release path
  // to wake up.
  EXPECT_TRUE(s.blocked_spill_footprints().empty());
}
TEST(SheetSpillTest, FootprintOverflowingSheetBoundsIsRejected) {
  Sheet s("Sheet1");
  // Anchor at the last row, height 2 — footprint extends past kMaxRows.
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0)};
  EXPECT_FALSE(s.commit_spill(Sheet::kMaxRows - 1U, 0U, 2U, 1U, std::move(cells)));
  EXPECT_EQ(s.spill_region_at_anchor(Sheet::kMaxRows - 1U, 0U), nullptr);

  // Anchor at the last column, width 2 — footprint extends past kMaxCols.
  std::vector<Value> cells2 = {Value::number(1.0), Value::number(2.0)};
  EXPECT_FALSE(s.commit_spill(0U, Sheet::kMaxCols - 1U, 1U, 2U, std::move(cells2)));
  EXPECT_EQ(s.spill_region_at_anchor(0U, Sheet::kMaxCols - 1U), nullptr);
}
#elif GTEST_HAS_DEATH_TEST
TEST(SheetSpillDeathTest, ZeroRowsAborts) {
  Sheet s("Sheet1");
  EXPECT_DEATH(s.commit_spill(0U, 0U, 0U, 1U, std::vector<Value>{}), "zero-sized spill region");
}
TEST(SheetSpillDeathTest, ZeroColsAborts) {
  Sheet s("Sheet1");
  EXPECT_DEATH(s.commit_spill(0U, 0U, 1U, 0U, std::vector<Value>{}), "zero-sized spill region");
}
TEST(SheetSpillDeathTest, MismatchedCellsLengthAborts) {
  Sheet s("Sheet1");
  std::vector<Value> wrong_size = {Value::number(1.0), Value::number(2.0)};
  EXPECT_DEATH(s.commit_spill(0U, 0U, 3U, 1U, std::move(wrong_size)), "does not match");
}
TEST(SheetSpillDeathTest, FootprintOverTheDynamicArrayCellCeilingAborts) {
  Sheet s("Sheet1");
  EXPECT_DEATH(s.commit_spill(0U, 0U, Sheet::kMaxRows, 100U, std::vector<Value>{}),
               "exceeds the dynamic-array cell ceiling");
}
TEST(SheetSpillDeathTest, FootprintOverflowsRowsAborts) {
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0)};
  EXPECT_DEATH(s.commit_spill(Sheet::kMaxRows - 1U, 0U, 2U, 1U, std::move(cells)), "footprint exceeds sheet bounds");
}
TEST(SheetSpillDeathTest, FootprintOverflowsColsAborts) {
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0)};
  EXPECT_DEATH(s.commit_spill(0U, Sheet::kMaxCols - 1U, 1U, 2U, std::move(cells)), "footprint exceeds sheet bounds");
}
#endif  // NDEBUG / GTEST_HAS_DEATH_TEST

}  // namespace
}  // namespace formulon
