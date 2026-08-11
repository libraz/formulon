//
// Unit tests for the cell-level dynamic-array spill API on `Sheet`. The
// tests cover registration, collision detection, deep-copy semantics for
// Text payloads, eager invalidation when phantom cells are written, and
// the contract that `cell_at` remains spill-blind while
// `resolve_cell_value` is the spill-aware reader.

#include <algorithm>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "sheet.h"
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

  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 4U, 0U, 4U, bulk, formula_indices);

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

  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 2U, 0U, 0U, bulk, formula_indices);

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

  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 0U, 0U, 1U, bulk, formula_indices);

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

  std::vector<Value> bulk = {Value::number(99.0)};
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 0U, 0U, 0U, bulk, formula_indices);

  ASSERT_EQ(bulk.size(), 2U);
  EXPECT_EQ(bulk[0], Value::number(99.0));
  ASSERT_EQ(formula_indices.size(), 1U);
  EXPECT_EQ(formula_indices[0], 1U);
}

TEST(SheetReadRange, ReversedOrOutOfRangeRectangleAppendsNothing) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(1.0));

  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(2U, 1U, 0U, 0U, bulk, formula_indices);
  s.read_range(0U, 0U, 2U, 1U, bulk, formula_indices);
  s.read_range(0U, Sheet::kMaxRows, 0U, 0U, bulk, formula_indices);
  s.read_range(0U, 0U, 0U, Sheet::kMaxCols, bulk, formula_indices);

  EXPECT_TRUE(bulk.empty());
  EXPECT_TRUE(formula_indices.empty());
}

}  // namespace
}  // namespace formulon
