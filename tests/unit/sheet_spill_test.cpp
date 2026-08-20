//
// Unit tests for the cell-level dynamic-array spill API on `Sheet`. The
// tests cover registration, collision detection, deep-copy semantics for
// Text payloads, eager invalidation when phantom cells are written, and
// the contract that `cell_at` remains spill-blind while
// `resolve_cell_value` is the spill-aware reader.

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

// ---------------------------------------------------------------------------
// Basic registration and read-back
// ---------------------------------------------------------------------------

TEST(SheetSpillTest, CommitThreeByOneSucceedsAndExposesValues) {
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(10.0), Value::number(20.0), Value::number(30.0)};

  EXPECT_TRUE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  const SpillRegion* region = s.spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->anchor_row, 0U);
  EXPECT_EQ(region->anchor_col, 0U);
  EXPECT_EQ(region->rows, 3U);
  EXPECT_EQ(region->cols, 1U);

  // resolve_cell_value reads the anchor through the cell store and the
  // phantoms through the spill table; both must yield the right value.
  EXPECT_EQ(s.resolve_cell_value(0U, 0U), Value::number(10.0));
  EXPECT_EQ(s.resolve_cell_value(1U, 0U), Value::number(20.0));
  EXPECT_EQ(s.resolve_cell_value(2U, 0U), Value::number(30.0));

  // Anchor's stored cached_value must equal cells[0] per the contract.
  const Cell* anchor = s.cell_at(0U, 0U);
  ASSERT_NE(anchor, nullptr);
  EXPECT_EQ(anchor->cached_value, Value::number(10.0));

  // covering() returns the region for phantoms only.
  EXPECT_EQ(s.spill_region_covering(0U, 0U), nullptr);
  EXPECT_NE(s.spill_region_covering(1U, 0U), nullptr);
  EXPECT_NE(s.spill_region_covering(2U, 0U), nullptr);
}

TEST(SheetSpillTest, CommittedSpillFootprintSnapshotExcludesBlockedRecords) {
  Sheet s("Sheet1");
  ASSERT_TRUE(s.commit_spill(4U, 5U, 2U, 3U,
                             {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0),
                              Value::number(5.0), Value::number(6.0)}));
  s.set_cell_value(10U, 10U, Value::number(9.0));
  ASSERT_FALSE(s.commit_spill(9U, 10U, 2U, 2U,
                              {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0)}));

  const std::vector<SpillFootprint> footprints = s.committed_spill_footprints();
  ASSERT_EQ(footprints.size(), 1U);
  EXPECT_EQ(footprints[0].anchor_row, 4U);
  EXPECT_EQ(footprints[0].anchor_col, 5U);
  EXPECT_EQ(footprints[0].rows, 2U);
  EXPECT_EQ(footprints[0].cols, 3U);
  ASSERT_EQ(s.blocked_spill_footprints().size(), 1U);
}

TEST(SheetSpillTest, CommitTwoByThreeRowMajorOrderingMatches) {
  Sheet s("Sheet1");
  // 2x3 matrix:
  //   1 2 3
  //   4 5 6
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0),
                              Value::number(4.0), Value::number(5.0), Value::number(6.0)};
  EXPECT_TRUE(s.commit_spill(5U, 7U, 2U, 3U, std::move(cells)));

  EXPECT_EQ(s.resolve_cell_value(5U, 7U), Value::number(1.0));
  EXPECT_EQ(s.resolve_cell_value(5U, 8U), Value::number(2.0));
  EXPECT_EQ(s.resolve_cell_value(5U, 9U), Value::number(3.0));
  EXPECT_EQ(s.resolve_cell_value(6U, 7U), Value::number(4.0));
  EXPECT_EQ(s.resolve_cell_value(6U, 8U), Value::number(5.0));
  EXPECT_EQ(s.resolve_cell_value(6U, 9U), Value::number(6.0));
}

TEST(SheetSpillTest, LargeSpillFindsFarPhantomWithoutPerCellReverseIndex) {
  Sheet s("Sheet1");
  constexpr std::uint32_t kRows = 1000U;
  constexpr std::uint32_t kCols = 100U;
  std::vector<Value> cells(static_cast<std::size_t>(kRows) * kCols, Value::number(7.0));

  ASSERT_TRUE(s.commit_spill(10U, 20U, kRows, kCols, std::move(cells)));
  const SpillRegion* region = s.spill_region_covering(10U + kRows - 1U, 20U + kCols - 1U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->anchor_row, 10U);
  EXPECT_EQ(region->anchor_col, 20U);
  EXPECT_EQ(s.resolve_cell_value(10U + kRows - 1U, 20U + kCols - 1U), Value::number(7.0));

  s.clear_spill(10U, 20U);
  EXPECT_EQ(s.spill_region_covering(10U + kRows - 1U, 20U + kCols - 1U), nullptr);
}

// ---------------------------------------------------------------------------
// clear_spill
// ---------------------------------------------------------------------------

TEST(SheetSpillTest, ClearSpillRemovesRegionAndRevertsResolveToBlank) {
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));
  ASSERT_NE(s.spill_region_covering(2U, 0U), nullptr);

  s.clear_spill(0U, 0U);

  EXPECT_EQ(s.spill_region_at_anchor(0U, 0U), nullptr);
  EXPECT_EQ(s.spill_region_covering(1U, 0U), nullptr);
  EXPECT_EQ(s.spill_region_covering(2U, 0U), nullptr);

  // Anchor cell still carries the previously-set cached_value (clear_spill
  // does not touch the anchor cell's own slot per the documented contract).
  const Cell* anchor = s.cell_at(0U, 0U);
  ASSERT_NE(anchor, nullptr);
  EXPECT_EQ(anchor->cached_value, Value::number(1.0));

  // Phantom rows now revert to blank: no cell stored, no spill region.
  EXPECT_EQ(s.resolve_cell_value(1U, 0U), Value::blank());
  EXPECT_EQ(s.resolve_cell_value(2U, 0U), Value::blank());
}

TEST(SheetSpillTest, ClearSpillIsNoOpWhenNothingAnchored) {
  Sheet s("Sheet1");
  // Should not crash and should not allocate the lazy table.
  s.clear_spill(0U, 0U);
  s.clear_spill(99U, 99U);
  EXPECT_EQ(s.spill_region_at_anchor(0U, 0U), nullptr);
}

// ---------------------------------------------------------------------------
// Collision: literal cell in the footprint
// ---------------------------------------------------------------------------

TEST(SheetSpillTest, CollidingWithExistingLiteralReturnsFalseAndSurfacesSpill) {
  Sheet s("Sheet1");
  // Plant a literal at A2 (row 1, col 0). Then try to spill A1:A3 over it.
  s.set_cell_value(1U, 0U, Value::number(7.0));

  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  EXPECT_FALSE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  // No region was registered.
  EXPECT_EQ(s.spill_region_at_anchor(0U, 0U), nullptr);
  EXPECT_EQ(s.spill_region_covering(1U, 0U), nullptr);

  // Anchor's cached_value is #SPILL!.
  const Cell* anchor = s.cell_at(0U, 0U);
  ASSERT_NE(anchor, nullptr);
  ASSERT_TRUE(anchor->cached_value.is_error());
  EXPECT_EQ(anchor->cached_value.as_error(), ErrorCode::Spill);

  // Existing literal is preserved verbatim.
  const Cell* literal = s.cell_at(1U, 0U);
  ASSERT_NE(literal, nullptr);
  ASSERT_TRUE(literal->cached_value.is_number());
  EXPECT_EQ(literal->cached_value.as_number(), 7.0);
}

TEST(SheetSpillTest, CollidingWithExistingFormulaReturnsFalseAndSurfacesSpill) {
  Sheet s("Sheet1");
  s.set_cell_formula(1U, 0U, "=A1+1");

  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  EXPECT_FALSE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  const Cell* anchor = s.cell_at(0U, 0U);
  ASSERT_NE(anchor, nullptr);
  ASSERT_TRUE(anchor->cached_value.is_error());
  EXPECT_EQ(anchor->cached_value.as_error(), ErrorCode::Spill);

  const Cell* formula = s.cell_at(1U, 0U);
  ASSERT_NE(formula, nullptr);
  EXPECT_EQ(formula->formula_text, "=A1+1");
}

// ---------------------------------------------------------------------------
// Collision: overlap with another spill's phantom
// ---------------------------------------------------------------------------

TEST(SheetSpillTest, CollidingWithAnotherSpillsPhantomReturnsFalse) {
  Sheet s("Sheet1");
  // First spill: anchor B1 (row 0, col 1), 3x1 — phantoms at B2, B3.
  std::vector<Value> first = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 1U, 3U, 1U, std::move(first)));
  ASSERT_NE(s.spill_region_covering(1U, 1U), nullptr);  // B2 is a phantom

  // Second spill: anchor A2 (row 1, col 0), 1x3 — would cover B2 (phantom).
  std::vector<Value> second = {Value::number(10.0), Value::number(20.0), Value::number(30.0)};
  EXPECT_FALSE(s.commit_spill(1U, 0U, 1U, 3U, std::move(second)));

  // Anchor A2 carries #SPILL!.
  const Cell* anchor = s.cell_at(1U, 0U);
  ASSERT_NE(anchor, nullptr);
  ASSERT_TRUE(anchor->cached_value.is_error());
  EXPECT_EQ(anchor->cached_value.as_error(), ErrorCode::Spill);

  // First spill is intact.
  EXPECT_NE(s.spill_region_at_anchor(0U, 1U), nullptr);
  EXPECT_EQ(s.resolve_cell_value(1U, 1U), Value::number(2.0));
}

TEST(SheetSpillTest, SpillWouldCollideTreatsForeignPhantomAsBlocker) {
  Sheet s("Sheet1");
  ASSERT_TRUE(s.commit_spill(0U, 0U, 2U, 1U, {Value::number(1.0), Value::number(2.0)}));

  // A2 is a phantom, not a stored Cell. The foreign region still blocks a
  // read-only spill anchored there; the existing region remains untouched.
  EXPECT_TRUE(s.spill_would_collide(1U, 0U, 2U, 1U));
  EXPECT_NE(s.spill_region_at_anchor(0U, 0U), nullptr);
  EXPECT_EQ(s.resolve_cell_value(1U, 0U), Value::number(2.0));
}

TEST(SheetSpillTest, SpillWouldCollideIgnoresCurrentAnchorSpill) {
  Sheet s("Sheet1");
  ASSERT_TRUE(s.commit_spill(0U, 0U, 2U, 1U, {Value::number(1.0), Value::number(2.0)}));

  // Ad-hoc re-evaluation is allowed to inspect the producer's own current
  // spill without reporting that spill as a foreign blocker.
  EXPECT_FALSE(s.spill_would_collide(0U, 0U, 3U, 1U));
}

TEST(SheetSpillTest, MergedRangeBlocksSpillAndPreservesMetadata) {
  Sheet s("Sheet1");
  s.set_cell_formula(0U, 0U, "=SEQUENCE(3,3)");
  s.set_cell_xf_index(0U, 0U, 17U);
  s.mutable_merges().push_back(MergeRange{1U, 1U, 2U, 2U});

  EXPECT_TRUE(s.spill_would_collide(0U, 0U, 3U, 3U));
  EXPECT_FALSE(s.commit_spill(
      0U, 0U, 3U, 3U,
      {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0), Value::number(5.0),
       Value::number(6.0), Value::number(7.0), Value::number(8.0), Value::number(9.0)}));

  const Cell* anchor = s.cell_at(0U, 0U);
  ASSERT_NE(anchor, nullptr);
  EXPECT_EQ(anchor->formula_text, "=SEQUENCE(3,3)");
  EXPECT_EQ(anchor->xf_index, 17U);
  ASSERT_TRUE(anchor->cached_value.is_error());
  EXPECT_EQ(anchor->cached_value.as_error(), ErrorCode::Spill);
  ASSERT_EQ(s.merges().size(), 1U);
  EXPECT_EQ(s.merges()[0].first_row, 1U);
  EXPECT_EQ(s.merges()[0].last_col, 2U);
}

TEST(SheetSpillTest, MergeAtAnchorBlocksVerticalSpill) {
  Sheet s("Sheet1");
  s.set_cell_formula(0U, 0U, "=SEQUENCE(2,1)");
  s.set_cell_xf_index(0U, 0U, 23U);
  // Excel's A1:B1 merge occupies the anchor itself and its adjacent cell.
  s.mutable_merges().push_back(MergeRange{0U, 0U, 0U, 1U});

  EXPECT_TRUE(s.spill_would_collide(0U, 0U, 2U, 1U));
  EXPECT_FALSE(s.commit_spill(0U, 0U, 2U, 1U, {Value::number(1.0), Value::number(2.0)}));

  const Cell* anchor = s.cell_at(0U, 0U);
  ASSERT_NE(anchor, nullptr);
  EXPECT_EQ(anchor->formula_text, "=SEQUENCE(2,1)");
  EXPECT_EQ(anchor->xf_index, 23U);
  ASSERT_TRUE(anchor->cached_value.is_error());
  EXPECT_EQ(anchor->cached_value.as_error(), ErrorCode::Spill);
  EXPECT_EQ(s.spill_region_at_anchor(0U, 0U), nullptr);
  ASSERT_EQ(s.merges().size(), 1U);
  EXPECT_EQ(s.merges()[0].first_row, 0U);
  EXPECT_EQ(s.merges()[0].first_col, 0U);
  EXPECT_EQ(s.merges()[0].last_row, 0U);
  EXPECT_EQ(s.merges()[0].last_col, 1U);
}

TEST(SheetSpillTest, SpillWouldCollideRejectsMalformedFootprints) {
  Sheet s("Sheet1");
  EXPECT_TRUE(s.spill_would_collide(0U, 0U, 0U, 1U));
  EXPECT_TRUE(s.spill_would_collide(0U, 0U, 1U, 0U));
  EXPECT_TRUE(s.spill_would_collide(Sheet::kMaxRows, 0U, 1U, 1U));
  EXPECT_TRUE(s.spill_would_collide(0U, Sheet::kMaxCols, 1U, 1U));
  EXPECT_TRUE(s.spill_would_collide(Sheet::kMaxRows - 1U, 0U, 2U, 1U));
  EXPECT_TRUE(s.spill_would_collide(0U, Sheet::kMaxCols - 1U, 1U, 2U));
}

// ---------------------------------------------------------------------------
// Text deep-copy
// ---------------------------------------------------------------------------

TEST(SheetSpillTest, TextValuesAreDeepCopiedAndOutliveSourceVector) {
  Sheet s("Sheet1");
  // Build the source string in scoped storage so its address is recycled
  // after the vector is consumed by commit_spill — guarantees the spill
  // payload is genuinely independent of the input.
  std::string scratch_a = "alpha";
  std::string scratch_b = "beta";
  std::vector<Value> cells = {Value::text(scratch_a), Value::text(scratch_b)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 2U, 1U, std::move(cells)));

  // Mutate / drop the original strings.
  scratch_a.assign(64U, 'x');
  scratch_b.assign(64U, 'y');

  // Spill payload still observes the original values.
  const Value v_anchor = s.resolve_cell_value(0U, 0U);
  ASSERT_TRUE(v_anchor.is_text());
  EXPECT_EQ(v_anchor.as_text(), "alpha");

  const Value v_phantom = s.resolve_cell_value(1U, 0U);
  ASSERT_TRUE(v_phantom.is_text());
  EXPECT_EQ(v_phantom.as_text(), "beta");
}

// ---------------------------------------------------------------------------
// Degenerate 1x1 region
// ---------------------------------------------------------------------------

TEST(SheetSpillTest, OneByOneRegionHasNoPhantoms) {
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(42.0)};
  EXPECT_TRUE(s.commit_spill(3U, 4U, 1U, 1U, std::move(cells)));

  const SpillRegion* region = s.spill_region_at_anchor(3U, 4U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 1U);
  EXPECT_EQ(region->cols, 1U);

  // Anchor's cached_value matches cells[0].
  const Cell* anchor = s.cell_at(3U, 4U);
  ASSERT_NE(anchor, nullptr);
  EXPECT_EQ(anchor->cached_value, Value::number(42.0));

  // Adjacent cells are not covered (no phantoms exist).
  EXPECT_EQ(s.spill_region_covering(3U, 5U), nullptr);
  EXPECT_EQ(s.spill_region_covering(4U, 4U), nullptr);
}

// ---------------------------------------------------------------------------
// Eager invalidation
// ---------------------------------------------------------------------------

TEST(SheetSpillTest, WriteToPhantomEagerlyClearsSpill) {
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));
  ASSERT_NE(s.spill_region_covering(1U, 0U), nullptr);

  // Writing to A2 (a phantom) must eagerly drop the spill.
  s.set_cell_value(1U, 0U, Value::number(99.0));

  EXPECT_EQ(s.spill_region_at_anchor(0U, 0U), nullptr);
  EXPECT_EQ(s.spill_region_covering(1U, 0U), nullptr);
  EXPECT_EQ(s.spill_region_covering(2U, 0U), nullptr);

  // The literal write happened.
  EXPECT_EQ(s.resolve_cell_value(1U, 0U), Value::number(99.0));
  // A3 (previously phantom, now uncovered and unset) is back to blank.
  EXPECT_EQ(s.resolve_cell_value(2U, 0U), Value::blank());
}

TEST(SheetSpillTest, FormulaWriteToPhantomEagerlyClearsSpill) {
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  s.set_cell_formula(1U, 0U, "=B1");
  EXPECT_EQ(s.spill_region_at_anchor(0U, 0U), nullptr);

  const Cell* phantom = s.cell_at(1U, 0U);
  ASSERT_NE(phantom, nullptr);
  EXPECT_EQ(phantom->formula_text, "=B1");
}

TEST(SheetSpillTest, WriteToAnchorEagerlyClearsSpill) {
  // Replacing a dynamic-array formula with a literal invalidates the entire
  // former spill footprint, just like a direct write to one of its phantoms.
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  s.set_cell_value(0U, 0U, Value::number(500.0));

  EXPECT_EQ(s.spill_region_at_anchor(0U, 0U), nullptr);
  EXPECT_EQ(s.spill_region_covering(1U, 0U), nullptr);
  // The anchor cell's literal value reflects the recent write.
  const Cell* anchor = s.cell_at(0U, 0U);
  ASSERT_NE(anchor, nullptr);
  EXPECT_EQ(anchor->cached_value, Value::number(500.0));
  EXPECT_TRUE(s.resolve_cell_value(1U, 0U).is_blank());
}

// ---------------------------------------------------------------------------
// Idempotent re-registration
// ---------------------------------------------------------------------------

TEST(SheetSpillTest, CommitOverExistingAnchorClearsOldRegionFirst) {
  Sheet s("Sheet1");
  std::vector<Value> first = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 3U, 1U, std::move(first)));
  ASSERT_NE(s.spill_region_covering(2U, 0U), nullptr);

  // Re-register at the same anchor with a smaller region.
  std::vector<Value> second = {Value::number(100.0), Value::number(200.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 2U, 1U, std::move(second)));

  // New region replaced the old one entirely.
  const SpillRegion* region = s.spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 2U);
  EXPECT_EQ(region->cols, 1U);

  EXPECT_EQ(s.resolve_cell_value(0U, 0U), Value::number(100.0));
  EXPECT_EQ(s.resolve_cell_value(1U, 0U), Value::number(200.0));
  // A3 is no longer a phantom: the old region was cleared.
  EXPECT_EQ(s.spill_region_covering(2U, 0U), nullptr);
  EXPECT_EQ(s.resolve_cell_value(2U, 0U), Value::blank());
}

// ---------------------------------------------------------------------------
// Bounds and shape rejection
// ---------------------------------------------------------------------------

// `commit_spill` validates shape and footprint via `assert(false)` followed
// by a `return false` safety net. In debug builds the assert fires first; in
// release builds (NDEBUG) the safety net is the only barrier. Cover both.
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

// ---------------------------------------------------------------------------
// cell_at remains spill-blind
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Phantom enumeration (cell_count / spill_phantom_addresses)
// ---------------------------------------------------------------------------

TEST(SheetSpillTest, SpillPhantomAddressesListsPhantomsExcludingAnchor) {
  Sheet s("Sheet1");
  // Anchor D1 (row 0, col 3), 3x1 spill: phantoms at D2 (1,3) and D3 (2,3).
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 3U, 3U, 1U, std::move(cells)));

  std::vector<CellAddress> phantoms = s.spill_phantom_addresses();
  std::sort(phantoms.begin(), phantoms.end(),
            [](CellAddress a, CellAddress b) { return a.row != b.row ? a.row < b.row : a.col < b.col; });
  ASSERT_EQ(phantoms.size(), 2U);
  EXPECT_EQ(phantoms[0].row, 1U);
  EXPECT_EQ(phantoms[0].col, 3U);
  EXPECT_EQ(phantoms[1].row, 2U);
  EXPECT_EQ(phantoms[1].col, 3U);
}

TEST(SheetSpillTest, SpillPhantomAddressesIsEmptyWithoutRegions) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(1.0));
  EXPECT_TRUE(s.spill_phantom_addresses().empty());
}

TEST(SheetSpillTest, CellCountIncludesPhantomsWithoutStoredSlots) {
  Sheet s("Sheet1");
  // Anchor D1 (row 0, col 3), 3x1 spill. The row's run starts at column D, so
  // committing materialises exactly one slot; the two phantoms D2, D3 live in
  // rows 1 and 2, which hold no stored slots at all, so each adds one.
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 3U, 3U, 1U, std::move(cells)));

  EXPECT_EQ(s.cell_count(), 1U + 2U);
}

TEST(SheetSpillTest, CellCountDoesNotDoubleCountPhantomOverImplicitSlot) {
  Sheet s("Sheet1");
  // Spill D1:D3 (phantoms at D2, D3), then populate D2 and F2 in row 1. The
  // literal at D2 is a phantom coordinate that now also holds a stored slot,
  // and the write to F2 extends that row's run to cols D..F. Phantom D2 must
  // be counted once; phantom D3 (row 2) still has no slot and adds one.
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 3U, 3U, 1U, std::move(cells)));
  s.set_cell_value(1U, 5U, Value::number(9.0));
  s.set_cell_xf_index(1U, 3U, 7U);

  // Neither write targets the anchor, so the spill stays intact.
  ASSERT_NE(s.spill_region_at_anchor(0U, 3U), nullptr);
  // row 0: 1 slot (D1), row 1: 3 slots (D2..F2), phantom D3 (row 2): +1.
  EXPECT_EQ(s.cell_count(), 1U + 3U + 1U);
}

TEST(SheetSpillTest, CellAtIsUnchangedForPhantomCoordinates) {
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  // Anchor: cell exists in storage (commit_spill writes its cached_value).
  EXPECT_NE(s.cell_at(0U, 0U), nullptr);

  // Phantoms: no Cell stored. cell_at returns nullptr; resolve_cell_value
  // is the spill-aware reader for these coordinates.
  EXPECT_EQ(s.cell_at(1U, 0U), nullptr);
  EXPECT_EQ(s.cell_at(2U, 0U), nullptr);
  EXPECT_EQ(s.resolve_cell_value(1U, 0U), Value::number(2.0));
  EXPECT_EQ(s.resolve_cell_value(2U, 0U), Value::number(3.0));
}

TEST(SheetSpillTest, PopulatedExtentUnifiesStoredCellsAndCommittedSpills) {
  Sheet s("Sheet1");
  s.set_cell_value(10U, 2U, Value::number(7.0));
  s.set_cell_formula(20U, 8U, "=1");
  ASSERT_TRUE(s.commit_spill(30U, 4U, 3U, 2U,
                             {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0),
                              Value::number(5.0), Value::number(6.0)}));

  const auto extent = s.populated_extent(0U, 0U, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U);
  ASSERT_TRUE(extent.has_value());
  EXPECT_EQ(extent->first_row, 10U);
  EXPECT_EQ(extent->first_col, 2U);
  EXPECT_EQ(extent->last_row, 32U);
  EXPECT_EQ(extent->last_col, 8U);

  const auto phantom_only_column = s.populated_extent(0U, 5U, Sheet::kMaxRows - 1U, 5U);
  ASSERT_TRUE(phantom_only_column.has_value());
  EXPECT_EQ(phantom_only_column->first_row, 30U);
  EXPECT_EQ(phantom_only_column->last_row, 32U);
  EXPECT_EQ(phantom_only_column->first_col, 5U);
  EXPECT_EQ(phantom_only_column->last_col, 5U);

  const auto empty = s.populated_extent(0U, 0U, 9U, 1U);
  EXPECT_FALSE(empty.has_value());
}

TEST(SheetSpillTest, PopulatedExtentExcludesBlockedSpillFootprints) {
  Sheet s("Sheet1");
  s.set_cell_formula(0U, 0U, "=SEQUENCE(1,3)");
  s.set_cell_value(0U, 2U, Value::number(9.0));
  EXPECT_FALSE(s.commit_spill(0U, 0U, 1U, 3U, {Value::number(1.0), Value::number(2.0), Value::number(3.0)}));

  // The failed footprint is retained for dependency invalidation, but its
  // blank middle coordinate is not populated content.
  EXPECT_FALSE(s.populated_extent(0U, 1U, 0U, 1U).has_value());
}

// ---------------------------------------------------------------------------
// Bulk range reads
// ---------------------------------------------------------------------------
//
// `read_range` exists to take the sheet lock once for a whole rectangle
// instead of twice per cell. These tests pin it against the per-coordinate
// readers it replaces, because the two must not drift.

TEST(SheetReadRange, MatchesResolveCellValueOverLiteralsAndGaps) {
  Sheet s("Sheet1");
  s.set_cell_value(1U, 1U, Value::number(1.0));
  s.set_cell_value(1U, 3U, Value::number(2.0));
  s.set_cell_value(3U, 2U, Value::text("x"));

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 4U, 0U, 4U, arena, bulk, formula_indices);

  ASSERT_EQ(bulk.size(), 25U);
  EXPECT_TRUE(formula_indices.empty());
  std::size_t index = 0;
  for (std::uint32_t row = 0; row <= 4U; ++row) {
    for (std::uint32_t col = 0; col <= 4U; ++col, ++index) {
      EXPECT_EQ(bulk[index], s.resolve_cell_value(row, col)) << "at (" << row << "," << col << ")";
    }
  }
}

TEST(SheetReadRange, SurfacesSpillPhantomValues) {
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 2U, 0U, 0U, arena, bulk, formula_indices);

  ASSERT_EQ(bulk.size(), 3U);
  EXPECT_TRUE(formula_indices.empty());
  EXPECT_EQ(bulk[0], Value::number(1.0));
  EXPECT_EQ(bulk[1], Value::number(2.0));
  EXPECT_EQ(bulk[2], Value::number(3.0));
}

TEST(SheetReadRange, ReportsFormulaCoordinatesInsteadOfEvaluatingThem) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(5.0));
  s.set_cell_formula(0U, 1U, "=A1*2");

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 0U, 0U, 1U, arena, bulk, formula_indices);

  ASSERT_EQ(bulk.size(), 2U);
  ASSERT_EQ(formula_indices.size(), 1U);
  EXPECT_EQ(formula_indices[0], 1U);
  // The formula slot carries only its cached value; evaluation is the
  // caller's job because it re-enters the sheet.
  EXPECT_TRUE(bulk[1].is_blank());
}

TEST(SheetReadRange, AppendsToTheCallersBufferAndIndexesAbsolutely) {
  Sheet s("Sheet1");
  s.set_cell_formula(0U, 0U, "=1");

  Arena arena;
  std::vector<Value> bulk = {Value::number(99.0)};
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 0U, 0U, 0U, arena, bulk, formula_indices);

  ASSERT_EQ(bulk.size(), 2U);
  EXPECT_EQ(bulk[0], Value::number(99.0));
  ASSERT_EQ(formula_indices.size(), 1U);
  EXPECT_EQ(formula_indices[0], 1U);
}

TEST(SheetReadRange, ReversedOrOutOfRangeRectangleAppendsNothing) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(1.0));

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(2U, 1U, 0U, 0U, arena, bulk, formula_indices);
  s.read_range(0U, 0U, 2U, 1U, arena, bulk, formula_indices);
  s.read_range(0U, Sheet::kMaxRows, 0U, 0U, arena, bulk, formula_indices);
  s.read_range(0U, 0U, 0U, Sheet::kMaxCols, arena, bulk, formula_indices);

  EXPECT_TRUE(bulk.empty());
  EXPECT_TRUE(formula_indices.empty());
}

// ---------------------------------------------------------------------------
// Text lifetime across the reads that copy values out of the sheet
// ---------------------------------------------------------------------------
//
// A Text `Value` is a view, and inside the sheet it views bytes the sheet
// owns and frees. A copy handed to a caller therefore has to be re-pointed at
// storage the caller controls before the lock is released; otherwise the copy
// looks self-contained while depending on a sheet mutation not happening.
// These tests assert both halves: the read does not alias sheet storage, and
// the value it produced still reads correctly after the mutation that frees
// what it came from.

// Overwrites enough unrelated cells to make an allocator hand the freed bytes
// out again, so a stale view is unlikely to still find its old contents.
void ChurnTextAllocations(Sheet& sheet) {
  for (std::uint32_t row = 100U; row < 200U; ++row) {
    sheet.set_cell_text(row, 0U, "filler payload of about the same length as the original");
  }
}

TEST(SheetReadRange, TextOutlivesTheWriteThatFreesTheCellsOwnBytes) {
  Sheet s("Sheet1");
  const std::string original = "a text payload long enough to live on the heap";
  s.set_cell_text(0U, 0U, original);

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 0U, 0U, 0U, arena, bulk, formula_indices);
  ASSERT_EQ(bulk.size(), 1U);
  ASSERT_TRUE(bulk[0].is_text());

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_NE(cell->cached_text_owned, nullptr);
  EXPECT_NE(bulk[0].as_text().data(), cell->cached_text_owned->data())
      << "the read handed back a view into the cell's own allocation";

  // The next cached write replaces that allocation.
  s.set_cell_cached_value(0U, 0U, Value::text("a replacement payload of a comparable length"));
  ChurnTextAllocations(s);
  EXPECT_EQ(bulk[0].as_text(), original);
}

TEST(SheetReadRange, TextOutlivesTheClearThatFreesTheSpillRegionsBytes) {
  Sheet s("Sheet1");
  const std::string original = "a spilled text payload long enough to live on the heap";
  std::vector<Value> cells = {Value::number(1.0), Value::text(original)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 2U, 1U, std::move(cells)));

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 1U, 0U, 0U, arena, bulk, formula_indices);
  ASSERT_EQ(bulk.size(), 2U);
  ASSERT_TRUE(bulk[1].is_text());

  const SpillRegion* region = s.spill_region_at_anchor(0U, 0U);
  ASSERT_NE(region, nullptr);
  ASSERT_EQ(region->owned_strings.size(), 1U);
  EXPECT_NE(bulk[1].as_text().data(), region->owned_strings.front().data())
      << "the read handed back a view into the region's own storage";

  s.clear_spill(0U, 0U);
  ChurnTextAllocations(s);
  EXPECT_EQ(bulk[1].as_text(), original);
}

TEST(SheetSpillTest, ReadSpillRegionAtAnchorOutlivesTheRegionItCopied) {
  Sheet s("Sheet1");
  const std::string original = "a spilled text payload long enough to live on the heap";
  std::vector<Value> cells = {Value::number(1.0), Value::text(original)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 2U, 1U, std::move(cells)));

  Arena arena;
  std::vector<Value> copied;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_TRUE(s.read_spill_region_at_anchor(0U, 0U, arena, copied, &rows, &cols));
  EXPECT_EQ(rows, 2U);
  EXPECT_EQ(cols, 1U);
  ASSERT_EQ(copied.size(), 2U);
  EXPECT_EQ(copied[0], Value::number(1.0));
  ASSERT_TRUE(copied[1].is_text());

  s.clear_spill(0U, 0U);
  ChurnTextAllocations(s);
  EXPECT_EQ(copied[1].as_text(), original);

  // A coordinate that anchors nothing appends nothing and says so.
  EXPECT_FALSE(s.read_spill_region_at_anchor(9U, 9U, arena, copied, &rows, &cols));
  EXPECT_EQ(copied.size(), 2U);
}

// ---------------------------------------------------------------------------
// Cost characteristics
// ---------------------------------------------------------------------------
//
// Both of these are about the shape of the work, not its duration, so they
// compare two shapes against each other rather than against a clock. An
// absolute bound would only say how loaded the machine is; a ratio says
// whether the cost follows the thing it must not follow. Each shape is timed
// several times and the fastest run is kept, which is the measurement least
// disturbed by whatever else is running.

/// Fastest of `kTrials` runs of `work`, in nanoseconds.
template <typename Fn>
double FastestRunNanos(Fn&& work) {
  constexpr int kTrials = 5;
  double best = 0.0;
  for (int trial = 0; trial < kTrials; ++trial) {
    const auto start = std::chrono::steady_clock::now();
    work();
    const double elapsed = std::chrono::duration<double, std::nano>(std::chrono::steady_clock::now() - start).count();
    if (trial == 0 || elapsed < best) {
      best = elapsed;
    }
  }
  return best;
}

/// How much slower the loaded shape may be before the cost is following the
/// wrong quantity. The shapes below differ by two orders of magnitude in the
/// quantity under test, so anything near 1 passes and a linear dependence
/// fails by a wide margin — the gap is what keeps this from flaking.
constexpr double kMaxCostRatio = 10.0;

TEST(SheetReadRange, CostDoesNotFollowTheNumberOfSpillRegionsOnTheSheet) {
  // A rectangle of empty cells, read on two sheets that differ only in how
  // many spill regions sit far away from it. Consulting the spill table per
  // coordinate makes the second sheet hundreds of times slower for a result
  // that is identical.
  constexpr std::uint32_t kLastRow = 199U;
  constexpr std::uint32_t kLastCol = 99U;
  constexpr std::uint32_t kManyRegions = 400U;

  const auto build = [](std::uint32_t regions) {
    Sheet sheet("Sheet1");
    for (std::uint32_t i = 0; i < regions; ++i) {
      // Parked well below and to the right of the read rectangle.
      EXPECT_TRUE(sheet.commit_spill(1000U + i, 500U, 1U, 1U, {Value::number(i)}));
    }
    return sheet;
  };
  Sheet few = build(1U);
  Sheet many = build(kManyRegions);

  const auto read = [&](Sheet& sheet) {
    return [&sheet]() {
      Arena arena;
      std::vector<Value> out;
      std::vector<std::size_t> formula_indices;
      sheet.read_range(0U, kLastRow, 0U, kLastCol, arena, out, formula_indices);
      EXPECT_EQ(out.size(), static_cast<std::size_t>(kLastRow + 1U) * (kLastCol + 1U));
    };
  };

  const double one_region = FastestRunNanos(read(few));
  const double many_regions = FastestRunNanos(read(many));
  EXPECT_LT(many_regions, one_region * kMaxCostRatio)
      << "read_range cost follows the spill-table size: " << one_region << "ns with one region, " << many_regions
      << "ns with " << kManyRegions;
}

TEST(SheetSpillTest, CellCountDoesNotFollowTheSpillArea) {
  // A spill is one rectangle with one payload, so counting the coordinates it
  // covers is arithmetic. Walking them instead puts a lookup per spilled cell
  // under the sheet lock, which a whole-column spill turns into a million.
  constexpr std::uint32_t kSmallSide = 4U;
  constexpr std::uint32_t kLargeSide = 400U;

  const auto build = [](std::uint32_t side) {
    Sheet sheet("Sheet1");
    std::vector<Value> cells;
    cells.reserve(static_cast<std::size_t>(side) * side);
    for (std::size_t i = 0; i < static_cast<std::size_t>(side) * side; ++i) {
      cells.push_back(Value::number(static_cast<double>(i)));
    }
    EXPECT_TRUE(sheet.commit_spill(0U, 0U, side, side, std::move(cells)));
    return sheet;
  };
  Sheet small = build(kSmallSide);
  Sheet large = build(kLargeSide);

  EXPECT_EQ(small.cell_count(), static_cast<std::size_t>(kSmallSide) * kSmallSide);
  EXPECT_EQ(large.cell_count(), static_cast<std::size_t>(kLargeSide) * kLargeSide);

  const auto count = [](Sheet& sheet) {
    return [&sheet]() {
      for (int i = 0; i < 20; ++i) {
        EXPECT_GT(sheet.cell_count(), 0U);
      }
    };
  };

  const double small_area = FastestRunNanos(count(small));
  const double large_area = FastestRunNanos(count(large));
  EXPECT_LT(large_area, small_area * kMaxCostRatio)
      << "cell_count cost follows the spill area: " << small_area << "ns over " << kSmallSide * kSmallSide << " cells, "
      << large_area << "ns over " << kLargeSide * kLargeSide;
}

TEST(SheetSpillTest, CellCountKeepsCountingPhantomsThatShareAMaterialisedSlot) {
  Sheet s("Sheet1");
  // Writing A1 and then E1 materialises the whole run A1..E1, three of them
  // implicitly blank. A spill over A1:C1 then covers slots that the stored
  // count already includes.
  s.set_cell_value(0U, 0U, Value::blank());
  s.set_cell_value(0U, 4U, Value::number(9.0));
  ASSERT_EQ(s.cell_count(), 5U);
  ASSERT_TRUE(s.commit_spill(0U, 0U, 1U, 3U, {Value::number(1.0), Value::number(2.0), Value::number(3.0)}));
  EXPECT_EQ(s.cell_count(), 5U) << "phantoms that coincide with materialised slots must not be counted twice";

  // A spill running past the end of the row's materialised run adds only the
  // coordinates that are not already slots.
  Sheet t("Sheet1");
  ASSERT_TRUE(t.commit_spill(
      0U, 0U, 1U, 5U,
      {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0), Value::number(5.0)}));
  EXPECT_EQ(t.cell_count(), 5U);
}

// ---------------------------------------------------------------------------
// Admission probing without materialising the footprint
// ---------------------------------------------------------------------------
//
// A producer whose result would be refused should learn that before building
// it. For a footprint the size of a grid axis the difference is 1,048,576
// cells against a handful, so these tests pin both the verdict and the work
// it took to reach it.

TEST(SheetSpillTest, ProbeAcceptsAClearFootprint) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(1.0));
  EXPECT_EQ(s.probe_spill_footprint(0U, 25U, 4U, 2U), Sheet::SpillAdmission::kAdmissible);
}

TEST(SheetSpillTest, ProbeReportsAFootprintThatLeavesTheGrid) {
  Sheet s("Sheet1");
  // A whole-column rectangle only fits when it starts at row 1; one row down
  // its far edge is past the last row of the grid.
  EXPECT_EQ(s.probe_spill_footprint(0U, 25U, Sheet::kMaxRows, 1U), Sheet::SpillAdmission::kAdmissible);
  EXPECT_EQ(s.probe_spill_footprint(1U, 25U, Sheet::kMaxRows, 1U), Sheet::SpillAdmission::kOutsideGrid);
  // Same rule on the other axis for a whole-row rectangle.
  EXPECT_EQ(s.probe_spill_footprint(4U, 0U, 2U, Sheet::kMaxCols), Sheet::SpillAdmission::kAdmissible);
  EXPECT_EQ(s.probe_spill_footprint(4U, 1U, 2U, Sheet::kMaxCols), Sheet::SpillAdmission::kOutsideGrid);
  // Degenerate shapes never commit either.
  EXPECT_EQ(s.probe_spill_footprint(0U, 0U, 0U, 1U), Sheet::SpillAdmission::kOutsideGrid);
  EXPECT_EQ(s.probe_spill_footprint(0U, 0U, 1U, 0U), Sheet::SpillAdmission::kOutsideGrid);
}

TEST(SheetSpillTest, ProbeFindsABlockerFarOutsideThePopulatedData) {
  Sheet s("Sheet1");
  // Three cells of data at the top of column A, and one unrelated cell 500
  // rows down in the column the rectangle would occupy. The blocker sits far
  // below everything else, which is what distinguishes the declared
  // rectangle from the sheet's used range.
  s.set_cell_value(0U, 0U, Value::number(1.0));
  s.set_cell_value(1U, 0U, Value::number(2.0));
  s.set_cell_value(2U, 0U, Value::number(3.0));
  s.set_cell_value(499U, 25U, Value::number(7.0));

  std::uint64_t steps = 0;
  EXPECT_EQ(s.probe_spill_footprint(0U, 25U, Sheet::kMaxRows, 1U, &steps), Sheet::SpillAdmission::kBlocked);

  // The verdict alone cannot tell a sparse scan from an area-proportional
  // one — both refuse, one merely takes a million probes to do it. Bound the
  // work by what the sheet stores, well under the 1,048,576-cell rectangle.
  EXPECT_LE(steps, 8U) << "admission scan must follow the stored cells, not the rectangle's area";
}

TEST(SheetSpillTest, ProbeRefusesAFullWidthRectangleOnABlockedRow) {
  Sheet s("Sheet1");
  s.set_cell_value(2U, 0U, Value::number(3.0));
  s.set_cell_value(2U, 1U, Value::number(30.0));
  s.set_cell_value(4U, 25U, Value::number(9.0));  // blocker inside the full-width span

  EXPECT_EQ(s.probe_spill_footprint(4U, 0U, 1U, Sheet::kMaxCols), Sheet::SpillAdmission::kBlocked);
  EXPECT_EQ(s.probe_spill_footprint(5U, 0U, 1U, Sheet::kMaxCols), Sheet::SpillAdmission::kAdmissible);

  // No step bound here on purpose. The proportionality claim lives on the row
  // axis: a row's cells are one dense run, so a blocker inside the span is
  // reached in as many steps as the run is wide either way, and an occupied
  // cell short-circuits the sweep at the first hit. Asserting a bound here
  // would hold whatever the scan does, which is worth less than no assertion
  // because it reads like coverage. The bound that does discriminate is in
  // `ProbeFindsABlockerFarOutsideThePopulatedData`.
}

TEST(SheetSpillTest, WouldCollideIsTheProbeVerdictInBooleanForm) {
  // The two refusal paths must agree on every input, which they do by both
  // being the same body. Pin that over a sheet carrying each kind of blocker
  // and over shapes that hit every branch: clear, blocked by a literal, by a
  // merge, by another region, degenerate, and off the grid.
  Sheet s("Sheet1");
  s.set_cell_value(4U, 4U, Value::number(1.0));
  s.mutable_merges().push_back(MergeRange{8U, 0U, 9U, 1U});
  ASSERT_TRUE(s.commit_spill(12U, 0U, 2U, 2U,
                             {Value::number(1.0), Value::number(2.0), Value::number(3.0), Value::number(4.0)}));

  const std::vector<SpillFootprint> shapes = {
      {0U, 0U, 1U, 1U},                    // clear
      {4U, 0U, 1U, 6U},                    // crosses the literal
      {7U, 0U, 3U, 1U},                    // crosses the merge
      {11U, 0U, 3U, 1U},                   // crosses the committed region
      {12U, 0U, 2U, 2U},                   // the committed region's own anchor
      {0U, 0U, 0U, 4U},                    // degenerate
      {0U, 0U, 4U, 0U},                    // degenerate
      {1U, 0U, Sheet::kMaxRows, 1U},       // off the bottom edge
      {0U, 1U, 1U, Sheet::kMaxCols},       // off the right edge
      {Sheet::kMaxRows - 1U, 0U, 1U, 1U},  // last row, in grid
  };
  for (const SpillFootprint& shape : shapes) {
    const bool collides = s.spill_would_collide(shape.anchor_row, shape.anchor_col, shape.rows, shape.cols);
    const Sheet::SpillAdmission admission =
        s.probe_spill_footprint(shape.anchor_row, shape.anchor_col, shape.rows, shape.cols);
    EXPECT_EQ(collides, admission != Sheet::SpillAdmission::kAdmissible)
        << "at (" << shape.anchor_row << "," << shape.anchor_col << ") " << shape.rows << "x" << shape.cols;
  }
}

TEST(SheetSpillTest, ProbeLeavesTheSheetUnchanged) {
  Sheet s("Sheet1");
  s.set_cell_value(499U, 25U, Value::number(7.0));
  const std::uint64_t revision_before = s.cell_enumeration_revision();
  const std::size_t cells_before = s.cell_count();

  EXPECT_EQ(s.probe_spill_footprint(0U, 25U, Sheet::kMaxRows, 1U), Sheet::SpillAdmission::kBlocked);
  EXPECT_EQ(s.probe_spill_footprint(0U, 25U, 2U, 1U), Sheet::SpillAdmission::kAdmissible);

  // No blocked record, no #SPILL! written, no row materialised by probing.
  EXPECT_TRUE(s.blocked_spill_footprints().empty());
  EXPECT_EQ(s.cell_count(), cells_before);
  EXPECT_EQ(s.cell_enumeration_revision(), revision_before);
  EXPECT_EQ(s.spill_region_at_anchor(0U, 25U), nullptr);
}

TEST(SheetSpillTest, RejectRecordsTheFootprintTheReleasePathRetries) {
  Sheet s("Sheet1");
  s.set_cell_value(499U, 25U, Value::number(7.0));
  ASSERT_EQ(s.probe_spill_footprint(0U, 25U, Sheet::kMaxRows, 1U), Sheet::SpillAdmission::kBlocked);

  s.reject_spill_footprint(0U, 25U, Sheet::kMaxRows, 1U);

  // The anchor reads #SPILL! and the rectangle is remembered, which is what
  // the blocked-spill release machinery retries once the blocker is gone.
  const Value anchor = s.resolve_cell_value(0U, 25U);
  ASSERT_TRUE(anchor.is_error());
  EXPECT_EQ(anchor.as_error(), ErrorCode::Spill);
  const std::vector<BlockedSpillFootprint> blocked = s.blocked_spill_footprints();
  ASSERT_EQ(blocked.size(), 1U);
  EXPECT_EQ(blocked[0].anchor_row, 0U);
  EXPECT_EQ(blocked[0].anchor_col, 25U);
  EXPECT_EQ(blocked[0].rows, Sheet::kMaxRows);
  EXPECT_EQ(blocked[0].cols, 1U);
}

TEST(SheetSpillTest, CommitRefusalRecordsTheSameFootprintAsAnExplicitReject) {
  // `commit_spill` and a caller that probes first and refuses on its own must
  // leave the sheet in the same state, or the two paths drift.
  Sheet s("Sheet1");
  s.set_cell_value(1U, 0U, Value::number(9.0));
  ASSERT_FALSE(s.commit_spill(0U, 0U, 3U, 1U, {Value::number(1.0), Value::number(2.0), Value::number(3.0)}));
  const std::vector<BlockedSpillFootprint> from_commit = s.blocked_spill_footprints();

  Sheet other("Sheet1");
  other.set_cell_value(1U, 0U, Value::number(9.0));
  ASSERT_EQ(other.probe_spill_footprint(0U, 0U, 3U, 1U), Sheet::SpillAdmission::kBlocked);
  other.reject_spill_footprint(0U, 0U, 3U, 1U);
  const std::vector<BlockedSpillFootprint> from_reject = other.blocked_spill_footprints();

  ASSERT_EQ(from_commit.size(), 1U);
  ASSERT_EQ(from_reject.size(), 1U);
  EXPECT_EQ(from_commit[0].anchor_row, from_reject[0].anchor_row);
  EXPECT_EQ(from_commit[0].anchor_col, from_reject[0].anchor_col);
  EXPECT_EQ(from_commit[0].rows, from_reject[0].rows);
  EXPECT_EQ(from_commit[0].cols, from_reject[0].cols);
  EXPECT_EQ(s.resolve_cell_value(0U, 0U), other.resolve_cell_value(0U, 0U));
}

}  // namespace
}  // namespace formulon
