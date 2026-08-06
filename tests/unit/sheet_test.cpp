//
// Unit tests for the `Sheet` row-sparse, column-dense cell store.
// Verifies CellAddress equality, literal/formula storage, growth and
// sparseness semantics, overwrite behaviour, value-kind round-trips,
// row iteration, and the maximum-coordinate boundary.

#include "sheet.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <set>
#include <string>

#include "cell.h"
#include "gtest/gtest.h"
#include "pivot/pivot_table.h"
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

TEST(SheetTest, StructuralEditsShiftPivotAnchorOnBothAxes) {
  Sheet s("Sheet1");
  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_anchor(/*row=*/5U, /*col=*/7U, /*rows=*/3U, /*cols=*/4U);
  s.add_pivot_table(std::move(pivot));

  s.insert_rows(/*row=*/3U, /*count=*/2U);
  s.insert_cols(/*col=*/4U, /*count=*/3U);
  ASSERT_EQ(s.pivot_tables().size(), 1U);
  EXPECT_EQ(s.pivot_tables()[0]->anchor_row(), 7U);
  EXPECT_EQ(s.pivot_tables()[0]->anchor_col(), 10U);

  // Deleting across the anchor clamps it to the start of the deleted band.
  s.delete_rows(/*row=*/7U, /*count=*/1U);
  s.delete_cols(/*col=*/9U, /*count=*/2U);
  EXPECT_EQ(s.pivot_tables()[0]->anchor_row(), 7U);
  EXPECT_EQ(s.pivot_tables()[0]->anchor_col(), 9U);
}

// ---------------------------------------------------------------------------
// Overwrite semantics
// ---------------------------------------------------------------------------

TEST(SheetTest, OverwriteLiteralWithLiteralKeepsCountAndUpdatesValue) {
  Sheet s("Sheet1");
  s.set_cell_value(1U, 1U, Value::number(1.0));
  // The row's run starts at its first populated column, so writing B2 does
  // not materialise column A.
  ASSERT_EQ(s.cell_count(), 1U);

  s.set_cell_value(1U, 1U, Value::number(99.0));
  EXPECT_EQ(s.cell_count(), 1U);

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
// Leading-gap storage
// ---------------------------------------------------------------------------

TEST(SheetTest, ColumnsBeforeTheFirstPopulatedOneAreNotMaterialised) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 5000U, Value::number(1.0));

  const RowCells& row = s.rows().at(0U);
  EXPECT_EQ(row.first_col(), 5000U);
  // One slot held, but the row still reports its full addressable width so
  // index-based scans behave as they would against a dense vector.
  EXPECT_EQ(row.stored_count(), 1U);
  EXPECT_EQ(row.size(), 5001U);
  EXPECT_EQ(s.cell_count(), 1U);

  // A column in the leading gap reads as a blank cell...
  EXPECT_TRUE(row[0U].formula_text.empty());
  EXPECT_TRUE(row[0U].cached_value.is_blank());
  // ...but is reported as absent, because it was never written.
  EXPECT_EQ(s.cell_at(0U, 0U), nullptr);
  EXPECT_FALSE(s.has_cell(0U, 0U));
}

TEST(SheetTest, WritingLeftOfTheRunExtendsItWithoutLosingCells) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 10U, Value::number(1.0));
  s.set_cell_text(0U, 12U, "kept");
  s.set_cell_value(0U, 4U, Value::number(2.0));

  const RowCells& row = s.rows().at(0U);
  EXPECT_EQ(row.first_col(), 4U);
  EXPECT_EQ(row.stored_count(), 9U);  // columns 4..12
  EXPECT_EQ(s.cell_at(0U, 4U)->cached_value.as_number(), 2.0);
  EXPECT_EQ(s.cell_at(0U, 10U)->cached_value.as_number(), 1.0);
  // The heap-owned text payload survives the run being re-seated.
  EXPECT_EQ(s.cell_at(0U, 12U)->cached_value.as_text(), "kept");
}

TEST(SheetTest, InsertColsLeftOfTheRunOnlyMovesItsOrigin) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 10U, Value::number(1.0));
  s.insert_cols(/*col=*/2U, /*count=*/3U);

  const RowCells& row = s.rows().at(0U);
  EXPECT_EQ(row.first_col(), 13U);
  EXPECT_EQ(row.stored_count(), 1U);
  EXPECT_EQ(s.cell_at(0U, 13U)->cached_value.as_number(), 1.0);
  EXPECT_EQ(s.cell_at(0U, 10U), nullptr);
}

TEST(SheetTest, DeleteColsLeftOfTheRunOnlyMovesItsOrigin) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 10U, Value::number(1.0));
  s.delete_cols(/*col=*/2U, /*count=*/3U);

  const RowCells& row = s.rows().at(0U);
  EXPECT_EQ(row.first_col(), 7U);
  EXPECT_EQ(row.stored_count(), 1U);
  EXPECT_EQ(s.cell_at(0U, 7U)->cached_value.as_number(), 1.0);
}

TEST(SheetTest, DeleteColsStraddlingTheRunStartDropsTheCoveredCells) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 10U, Value::number(1.0));
  s.set_cell_value(0U, 12U, Value::number(2.0));
  // The band covers columns 8..10, i.e. the run's first stored column.
  s.delete_cols(/*col=*/8U, /*count=*/3U);

  const RowCells& row = s.rows().at(0U);
  EXPECT_EQ(row.first_col(), 8U);
  EXPECT_EQ(row.stored_count(), 2U);  // former columns 11 and 12
  EXPECT_TRUE(s.cell_at(0U, 8U)->cached_value.is_blank());
  EXPECT_EQ(s.cell_at(0U, 9U)->cached_value.as_number(), 2.0);
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
  s.set_cell_value(0U, 0U, Value::number(1.0));  // row 0, vector size 1
  s.set_cell_value(1U, 4U, Value::number(2.0));  // row 1, vector size 5
  s.set_cell_formula(2U, 2U, "=A1");             // row 2, vector size 3

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
