// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the `Sheet` row-sparse, column-dense cell store.
// Verifies CellAddress equality, literal/formula storage, growth and
// sparseness semantics, overwrite behaviour, value-kind round-trips,
// row iteration, and the maximum-coordinate boundary.

#include "sheet.h"

#include <cstddef>
#include <cstdint>
#include <set>
#include <string>

#include "cell.h"
#include "gtest/gtest.h"
#include "value.h"

namespace formulon {
namespace {

// ---------------------------------------------------------------------------
// CellAddress equality
// ---------------------------------------------------------------------------

TEST(CellAddressTest, EqualWhenRowAndColMatch) {
  EXPECT_TRUE(CellAddress({1U, 2U}) == CellAddress({1U, 2U}));
  EXPECT_FALSE(CellAddress({1U, 2U}) != CellAddress({1U, 2U}));
}

TEST(CellAddressTest, NotEqualWhenColumnDiffers) {
  EXPECT_FALSE(CellAddress({1U, 2U}) == CellAddress({1U, 3U}));
  EXPECT_TRUE(CellAddress({1U, 2U}) != CellAddress({1U, 3U}));
}

TEST(CellAddressTest, NotEqualWhenRowDiffers) {
  EXPECT_FALSE(CellAddress({1U, 2U}) == CellAddress({2U, 2U}));
  EXPECT_TRUE(CellAddress({1U, 2U}) != CellAddress({2U, 2U}));
}

TEST(CellAddressTest, OriginEqualsOrigin) {
  EXPECT_TRUE(CellAddress({0U, 0U}) == CellAddress({0U, 0U}));
}

// ---------------------------------------------------------------------------
// Default Sheet state
// ---------------------------------------------------------------------------

TEST(SheetTest, DefaultHasZeroCellsAndEmptyRowMap) {
  Sheet s("Sheet1");
  EXPECT_EQ(s.cell_count(), 0U);
  EXPECT_TRUE(s.rows().empty());
}

TEST(SheetTest, CellAtOnEmptySheetReturnsNullptr) {
  Sheet s("Sheet1");
  EXPECT_EQ(s.cell_at(0U, 0U), nullptr);
  EXPECT_EQ(s.cell_at(123U, 456U), nullptr);
}

TEST(SheetTest, HasCellOnEmptySheetReturnsFalse) {
  Sheet s("Sheet1");
  EXPECT_FALSE(s.has_cell(0U, 0U));
  EXPECT_FALSE(s.has_cell(7U, 9U));
}

// ---------------------------------------------------------------------------
// Literal value round-trip
// ---------------------------------------------------------------------------

TEST(SheetTest, SetCellValueStoresNumberAndIsReadable) {
  Sheet s("Sheet1");
  s.set_cell_value(5U, 3U, Value::number(42.0));

  const Cell* cell = s.cell_at(5U, 3U);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->cached_value.is_number());
  EXPECT_EQ(cell->cached_value.as_number(), 42.0);
  EXPECT_TRUE(cell->formula_text.empty());
  EXPECT_TRUE(s.has_cell(5U, 3U));
}

TEST(SheetTest, SetCellValueStoresBoolean) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::boolean(true));

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_boolean());
  EXPECT_TRUE(cell->cached_value.as_boolean());
  EXPECT_TRUE(cell->formula_text.empty());
}

TEST(SheetTest, SetCellValueStoresText) {
  Sheet s("Sheet1");
  s.set_cell_value(2U, 4U, Value::text("hello"));

  const Cell* cell = s.cell_at(2U, 4U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_text());
  EXPECT_EQ(cell->cached_value.as_text(), "hello");
}

TEST(SheetTest, SetCellValueStoresError) {
  Sheet s("Sheet1");
  s.set_cell_value(1U, 1U, Value::error(ErrorCode::Div0));

  const Cell* cell = s.cell_at(1U, 1U);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_error());
  EXPECT_EQ(cell->cached_value.as_error(), ErrorCode::Div0);
}

TEST(SheetTest, SetCellValueStoresExplicitBlank) {
  Sheet s("Sheet1");
  s.set_cell_value(3U, 7U, Value::blank());

  const Cell* cell = s.cell_at(3U, 7U);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->cached_value.is_blank());
  EXPECT_TRUE(cell->formula_text.empty());
}

// ---------------------------------------------------------------------------
// Formula storage
// ---------------------------------------------------------------------------

TEST(SheetTest, SetCellFormulaStoresStringAndBlanksCached) {
  Sheet s("Sheet1");
  s.set_cell_formula(0U, 0U, "=A1+1");

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=A1+1");
  EXPECT_TRUE(cell->cached_value.is_blank());
}

TEST(SheetTest, SetCellFormulaStoresStringVerbatimWithoutValidation) {
  // Storage layer does not validate the leading '='; the parser owns that.
  Sheet s("Sheet1");
  s.set_cell_formula(2U, 2U, "no leading equals");

  const Cell* cell = s.cell_at(2U, 2U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "no leading equals");
}

// ---------------------------------------------------------------------------
// Overwrite semantics
// ---------------------------------------------------------------------------

TEST(SheetTest, OverwriteLiteralWithLiteralKeepsCountAndUpdatesValue) {
  Sheet s("Sheet1");
  s.set_cell_value(1U, 1U, Value::number(1.0));
  ASSERT_EQ(s.cell_count(), 2U);  // columns 0 and 1 grown

  s.set_cell_value(1U, 1U, Value::number(99.0));
  EXPECT_EQ(s.cell_count(), 2U);

  const Cell* cell = s.cell_at(1U, 1U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->cached_value.as_number(), 99.0);
  EXPECT_TRUE(cell->formula_text.empty());
}

TEST(SheetTest, OverwriteLiteralWithFormulaResetsCachedValue) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(7.0));
  ASSERT_EQ(s.cell_count(), 1U);

  s.set_cell_formula(0U, 0U, "=B1*2");
  EXPECT_EQ(s.cell_count(), 1U);

  const Cell* cell = s.cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=B1*2");
  EXPECT_TRUE(cell->cached_value.is_blank());
}

TEST(SheetTest, OverwriteFormulaWithLiteralClearsFormulaText) {
  Sheet s("Sheet1");
  s.set_cell_formula(4U, 4U, "=SUM(A1:A3)");
  s.set_cell_value(4U, 4U, Value::number(6.0));

  const Cell* cell = s.cell_at(4U, 4U);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->formula_text.empty());
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_EQ(cell->cached_value.as_number(), 6.0);
}

// ---------------------------------------------------------------------------
// Sparseness: columns and rows
// ---------------------------------------------------------------------------

TEST(SheetTest, SparseColumnsGrowVectorImplicitly) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(1.0));
  s.set_cell_value(0U, 100U, Value::number(2.0));

  // Both explicitly-set columns are readable.
  const Cell* a = s.cell_at(0U, 0U);
  const Cell* z = s.cell_at(0U, 100U);
  ASSERT_NE(a, nullptr);
  ASSERT_NE(z, nullptr);
  EXPECT_EQ(a->cached_value.as_number(), 1.0);
  EXPECT_EQ(z->cached_value.as_number(), 2.0);

  // A column in the gap exists but is default-constructed (empty formula,
  // blank cached value); the storage contract documents this behaviour.
  const Cell* gap = s.cell_at(0U, 50U);
  ASSERT_NE(gap, nullptr);
  EXPECT_TRUE(gap->formula_text.empty());
  EXPECT_TRUE(gap->cached_value.is_blank());

  // A column past the row vector's end is not in storage.
  EXPECT_EQ(s.cell_at(0U, 200U), nullptr);
  EXPECT_FALSE(s.has_cell(0U, 200U));

  // cell_count counts every slot in the populated row, including implicit.
  EXPECT_EQ(s.cell_count(), 101U);
}

TEST(SheetTest, SparseRowsLeaveUnvisitedRowsAbsent) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(1.0));
  s.set_cell_value(100U, 0U, Value::number(2.0));

  EXPECT_EQ(s.cell_count(), 2U);

  // Untouched intermediate row is absent from the map entirely.
  EXPECT_EQ(s.cell_at(50U, 0U), nullptr);
  EXPECT_FALSE(s.has_cell(50U, 0U));

  EXPECT_TRUE(s.has_cell(0U, 0U));
  EXPECT_TRUE(s.has_cell(100U, 0U));
}

// ---------------------------------------------------------------------------
// rows() iteration
// ---------------------------------------------------------------------------

TEST(SheetTest, RowsIterationExposesAllPopulatedRows) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(10.0));
  s.set_cell_value(5U, 0U, Value::number(20.0));
  s.set_cell_value(42U, 1U, Value::number(30.0));

  std::set<std::uint32_t> seen;
  for (const auto& kv : s.rows()) {
    seen.insert(kv.first);
  }
  EXPECT_EQ(seen.size(), 3U);
  EXPECT_NE(seen.find(0U), seen.end());
  EXPECT_NE(seen.find(5U), seen.end());
  EXPECT_NE(seen.find(42U), seen.end());
}

TEST(SheetTest, RowsIterationReportsExpectedColumnCounts) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(1.0));   // row 0, vector size 1
  s.set_cell_value(1U, 4U, Value::number(2.0));   // row 1, vector size 5
  s.set_cell_formula(2U, 2U, "=A1");              // row 2, vector size 3

  ASSERT_EQ(s.rows().size(), 3U);
  EXPECT_EQ(s.rows().at(0U).size(), 1U);
  EXPECT_EQ(s.rows().at(1U).size(), 5U);
  EXPECT_EQ(s.rows().at(2U).size(), 3U);
}

// ---------------------------------------------------------------------------
// Boundary
// ---------------------------------------------------------------------------

TEST(SheetTest, MaxValidCoordinateIsAccepted) {
  Sheet s("Sheet1");
  const std::uint32_t max_row = Sheet::kMaxRows - 1U;  // 1,048,575
  const std::uint32_t max_col = Sheet::kMaxCols - 1U;  // 16,383

  s.set_cell_value(max_row, max_col, Value::number(-1.0));

  const Cell* cell = s.cell_at(max_row, max_col);
  ASSERT_NE(cell, nullptr);
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_EQ(cell->cached_value.as_number(), -1.0);
  EXPECT_TRUE(s.has_cell(max_row, max_col));
}

// ---------------------------------------------------------------------------
// Sheet name continues to work alongside the cell store
// ---------------------------------------------------------------------------

TEST(SheetTest, NameAccessorsUnaffectedByCellStore) {
  Sheet s("Daten");
  EXPECT_EQ(s.name(), "Daten");

  s.set_cell_value(0U, 0U, Value::number(1.0));
  EXPECT_EQ(s.name(), "Daten");

  s.set_name("Renamed");
  EXPECT_EQ(s.name(), "Renamed");
  EXPECT_EQ(s.cell_count(), 1U);
}

}  // namespace
}  // namespace formulon
