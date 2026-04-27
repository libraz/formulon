// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the cell-level dynamic-array spill API on `Sheet`. The
// tests cover registration, collision detection, deep-copy semantics for
// Text payloads, eager invalidation when phantom cells are written, and
// the contract that `cell_at` remains spill-blind while
// `resolve_cell_value` is the spill-aware reader.

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

TEST(SheetSpillTest, WriteToAnchorDoesNotEagerlyClearSpill) {
  // The eager-invalidation rule is "writing to a phantom drops the spill".
  // Writing to the anchor itself is a separate, intentional caller action
  // (e.g. overwriting the formula). The anchor write must succeed and the
  // spill table is not consulted for the anchor coordinate.
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  s.set_cell_value(0U, 0U, Value::number(500.0));

  // Per the contract, writing to the anchor does not auto-clear; the spill
  // table still references the original anchor.
  EXPECT_NE(s.spill_region_at_anchor(0U, 0U), nullptr);
  EXPECT_NE(s.spill_region_covering(1U, 0U), nullptr);
  // The anchor cell's literal value reflects the recent write.
  const Cell* anchor = s.cell_at(0U, 0U);
  ASSERT_NE(anchor, nullptr);
  EXPECT_EQ(anchor->cached_value, Value::number(500.0));
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

// ---------------------------------------------------------------------------
// cell_at remains spill-blind
// ---------------------------------------------------------------------------

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

}  // namespace
}  // namespace formulon
