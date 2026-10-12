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
#include <vector>

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

TEST(SheetTest, StructuralEditsKeepPivotSpansInsideGridAtEdges) {
  for (const bool rows : {false, true}) {
    SCOPED_TRACE(rows ? "rows" : "columns");
    Sheet s("Sheet1");
    const std::uint32_t bound = rows ? Sheet::kMaxRows : Sheet::kMaxCols;
    const std::uint32_t insert_at = bound - 3U;

    auto near_edge = std::make_unique<pivot::PivotTable>();
    auto fully_shifted = std::make_unique<pivot::PivotTable>();
    auto zero_span = std::make_unique<pivot::PivotTable>();
    if (rows) {
      near_edge->set_anchor(insert_at, 4U, 3U, 2U);
      fully_shifted->set_anchor(bound - 1U, 8U, 1U, 1U);
      zero_span->set_anchor(bound - 1U, 12U, 0U, 1U);
    } else {
      near_edge->set_anchor(4U, insert_at, 2U, 3U);
      fully_shifted->set_anchor(8U, bound - 1U, 1U, 1U);
      zero_span->set_anchor(12U, bound - 1U, 1U, 0U);
    }
    s.add_pivot_table(std::move(near_edge));
    s.add_pivot_table(std::move(fully_shifted));
    s.add_pivot_table(std::move(zero_span));

    if (rows) {
      s.insert_rows(insert_at, 1U);
    } else {
      s.insert_cols(insert_at, 1U);
    }

    // A pivot pushed wholly off the grid is dropped, like merges and validations.
    ASSERT_EQ(s.pivot_tables().size(), 1U);
    EXPECT_EQ(s.pivot_tables()[0]->anchor_row(), rows ? bound - 2U : 4U);
    EXPECT_EQ(s.pivot_tables()[0]->anchor_col(), rows ? 4U : bound - 2U);
    EXPECT_EQ(s.pivot_tables()[0]->span_rows(), 2U);
    EXPECT_EQ(s.pivot_tables()[0]->span_cols(), 2U);
  }
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
// Formula-cell index
// ---------------------------------------------------------------------------

TEST(SheetTest, FormulaCellsInReportsOnlyFormulaCellsInsideTheRectangle) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(1.0));
  s.set_cell_formula(1U, 0U, "=A1");
  s.set_cell_formula(5U, 0U, "=A2");
  s.set_cell_formula(2U, 2U, "=A1");
  s.set_cell_formula(9U, 2U, "=A1");
  s.set_cell_formula(3U, 4U, "=A1");

  // Columns A..C, rows 2..6: skips C10 below the band and E4 right of it.
  const std::vector<CellAddress> hits = s.formula_cells_in(1U, 0U, 5U, 2U);
  ASSERT_EQ(hits.size(), 3U);
  EXPECT_EQ(hits[0], (CellAddress{1U, 0U}));
  EXPECT_EQ(hits[1], (CellAddress{5U, 0U}));
  EXPECT_EQ(hits[2], (CellAddress{2U, 2U}));

  EXPECT_TRUE(s.formula_cells_in(0U, 1U, Sheet::kMaxRows - 1U, 1U).empty());
  EXPECT_TRUE(s.formula_cells_in(5U, 0U, 1U, 0U).empty());
}

TEST(SheetTest, FormulaCellsInTracksOverwritesAndStructuralEdits) {
  Sheet s("Sheet1");
  s.set_cell_formula(2U, 1U, "=A1");
  s.set_cell_formula(4U, 1U, "=A2");
  s.set_cell_value(4U, 1U, Value::number(3.0));
  s.set_cell_formula(6U, 1U, "=A3");
  s.set_cell_text(6U, 1U, "text");
  ASSERT_EQ(s.formula_cells_in(0U, 0U, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U).size(), 1U);

  s.insert_rows(0U, 2U);
  s.insert_cols(0U, 1U);
  std::vector<CellAddress> hits = s.formula_cells_in(0U, 0U, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U);
  ASSERT_EQ(hits.size(), 1U);
  EXPECT_EQ(hits[0], (CellAddress{4U, 2U}));

  s.delete_cols(2U, 1U);
  EXPECT_TRUE(s.formula_cells_in(0U, 0U, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U).empty());
}

// ---------------------------------------------------------------------------
// Sparseness: columns and rows
// ---------------------------------------------------------------------------

TEST(SheetTest, SparseColumnsStoreOnlyWrittenCells) {
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

  // A column between the written ones was never written, so it is absent.
  EXPECT_EQ(s.cell_at(0U, 50U), nullptr);
  EXPECT_FALSE(s.has_cell(0U, 50U));

  // A column past the last written one is not in storage either.
  EXPECT_EQ(s.cell_at(0U, 200U), nullptr);
  EXPECT_FALSE(s.has_cell(0U, 200U));

  // cell_count counts written coordinates only.
  EXPECT_EQ(s.cell_count(), 2U);
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

// Lowest and highest stored column of a non-empty row.
std::uint32_t FirstStoredCol(const RowCells& row) {
  return (*row.begin()).col;
}

std::uint32_t LastStoredCol(const RowCells& row) {
  std::uint32_t last = 0U;
  for (const auto& [col, cell] : row) {
    (void)cell;
    last = col;
  }
  return last;
}

TEST(SheetTest, ColumnsBeforeTheFirstPopulatedOneAreNotMaterialised) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 5000U, Value::number(1.0));

  const RowCells& row = s.rows().at(0U);
  ASSERT_EQ(row.stored_count(), 1U);
  EXPECT_EQ(FirstStoredCol(row), 5000U);
  EXPECT_EQ(LastStoredCol(row), 5000U);
  EXPECT_EQ(s.cell_count(), 1U);

  // A column in the leading gap is absent, because it was never written.
  EXPECT_EQ(row.find(0U), nullptr);
  EXPECT_EQ(s.cell_at(0U, 0U), nullptr);
  EXPECT_FALSE(s.has_cell(0U, 0U));
}

TEST(SheetTest, WritingLeftOfTheRunExtendsItWithoutLosingCells) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 10U, Value::number(1.0));
  s.set_cell_text(0U, 12U, "kept");
  s.set_cell_value(0U, 4U, Value::number(2.0));

  const RowCells& row = s.rows().at(0U);
  EXPECT_EQ(FirstStoredCol(row), 4U);
  EXPECT_EQ(row.stored_count(), 3U);  // columns 4, 10 and 12
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
  EXPECT_EQ(FirstStoredCol(row), 13U);
  EXPECT_EQ(row.stored_count(), 1U);
  EXPECT_EQ(s.cell_at(0U, 13U)->cached_value.as_number(), 1.0);
  EXPECT_EQ(s.cell_at(0U, 10U), nullptr);
}

TEST(SheetTest, DeleteColsLeftOfTheRunOnlyMovesItsOrigin) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 10U, Value::number(1.0));
  s.delete_cols(/*col=*/2U, /*count=*/3U);

  const RowCells& row = s.rows().at(0U);
  EXPECT_EQ(FirstStoredCol(row), 7U);
  EXPECT_EQ(row.stored_count(), 1U);
  EXPECT_EQ(s.cell_at(0U, 7U)->cached_value.as_number(), 1.0);
}

TEST(SheetTest, DeleteColsStraddlingTheRunStartDropsTheCoveredCells) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 10U, Value::number(1.0));
  s.set_cell_value(0U, 12U, Value::number(2.0));
  // The band covers columns 8..10, i.e. the row's first stored column.
  s.delete_cols(/*col=*/8U, /*count=*/3U);

  const RowCells& row = s.rows().at(0U);
  EXPECT_EQ(FirstStoredCol(row), 9U);
  EXPECT_EQ(row.stored_count(), 1U);  // former column 12
  EXPECT_EQ(s.cell_at(0U, 8U), nullptr);
  ASSERT_NE(s.cell_at(0U, 9U), nullptr);
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
  s.set_cell_value(0U, 0U, Value::number(1.0));  // row 0, col 0
  s.set_cell_value(1U, 4U, Value::number(2.0));  // row 1, col 4
  s.set_cell_formula(2U, 2U, "=A1");             // row 2, col 2

  ASSERT_EQ(s.rows().size(), 3U);
  EXPECT_EQ(s.rows().at(0U).stored_count(), 1U);
  EXPECT_EQ(LastStoredCol(s.rows().at(0U)), 0U);
  EXPECT_EQ(s.rows().at(1U).stored_count(), 1U);
  EXPECT_EQ(LastStoredCol(s.rows().at(1U)), 4U);
  EXPECT_EQ(s.rows().at(2U).stored_count(), 1U);
  EXPECT_EQ(LastStoredCol(s.rows().at(2U)), 2U);
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

// ---------------------------------------------------------------------------
// RowCells invariants
// ---------------------------------------------------------------------------

// Writes `col` as its own numeric value so a cell identifies its column.
void Put(RowCells& row, std::uint32_t col) {
  row.ensure(col).cached_value = Value::number(static_cast<double>(col));
}

std::vector<std::uint32_t> Keys(const RowCells& row) {
  std::vector<std::uint32_t> keys;
  for (const auto& [col, cell] : row) {
    EXPECT_EQ(cell.cached_value.as_number(), static_cast<double>(col)) << "cell paired with the wrong key";
    keys.push_back(col);
  }
  return keys;
}

TEST(RowCellsTest, RightToLeftWritesKeepKeysAscending) {
  RowCells row;
  for (std::uint32_t col = 10U; col-- > 0U;) {
    Put(row, col * 3U);
  }
  EXPECT_EQ(Keys(row), (std::vector<std::uint32_t>{0U, 3U, 6U, 9U, 12U, 15U, 18U, 21U, 24U, 27U}));
  EXPECT_EQ(row.stored_count(), 10U);
}

TEST(RowCellsTest, MiddleInsertKeepsNeighboursAndReturnsExisting) {
  RowCells row;
  Put(row, 1U);
  Put(row, 9U);
  Put(row, 5U);
  EXPECT_EQ(Keys(row), (std::vector<std::uint32_t>{1U, 5U, 9U}));
  // A second ensure of a held column returns the same cell without inserting.
  EXPECT_EQ(&row.ensure(5U), row.find(5U));
  EXPECT_EQ(row.stored_count(), 3U);
}

TEST(RowCellsTest, FindAroundAHole) {
  RowCells row;
  for (const std::uint32_t col : {4U, 5U, 6U, 20U, 21U}) {
    Put(row, col);
  }
  EXPECT_EQ(row.find(3U), nullptr);  // before the first key
  ASSERT_NE(row.find(4U), nullptr);
  ASSERT_NE(row.find(6U), nullptr);  // left of the hole
  EXPECT_EQ(row.find(6U)->cached_value.as_number(), 6.0);
  EXPECT_EQ(row.find(7U), nullptr);  // in the hole
  EXPECT_EQ(row.find(19U), nullptr);
  ASSERT_NE(row.find(20U), nullptr);  // right of the hole
  EXPECT_EQ(row.find(20U)->cached_value.as_number(), 20.0);
  ASSERT_NE(row.find(21U), nullptr);
  EXPECT_EQ(row.find(21U)->cached_value.as_number(), 21.0);
  EXPECT_EQ(row.find(22U), nullptr);  // after the last key
  EXPECT_EQ(row.find(16383U), nullptr);
}

TEST(RowCellsTest, GapIsAbsent) {
  RowCells row;
  Put(row, 0U);
  Put(row, 16383U);
  EXPECT_EQ(row.find(8000U), nullptr);
  EXPECT_EQ(row.stored_count(), 2U);
  EXPECT_EQ(LastStoredCol(row), 16383U);
  ASSERT_NE(row.find(16383U), nullptr);
  EXPECT_EQ(row.find(16383U)->cached_value.as_number(), 16383.0);
}

TEST(RowCellsTest, ShiftColsFromDropsKeysPushedPastTheLastColumn) {
  RowCells row;
  for (const std::uint32_t col : {2U, 10U, 16380U, 16382U, 16383U}) {
    Put(row, col);
  }
  row.shift_cols_from(10U, 2U);
  // 2 stays; 10 -> 12; 16380 -> 16382; 16382 and 16383 fall off the sheet.
  std::vector<std::uint32_t> keys;
  for (const auto& [col, cell] : row) {
    (void)cell;
    keys.push_back(col);
  }
  EXPECT_EQ(keys, (std::vector<std::uint32_t>{2U, 12U, 16382U}));
  ASSERT_NE(row.find(16382U), nullptr);
  EXPECT_EQ(row.find(16382U)->cached_value.as_number(), 16380.0);
}

TEST(RowCellsTest, ShiftColsFromPastEveryKeyEmptiesTheRow) {
  RowCells row;
  Put(row, 0U);
  Put(row, 100U);
  row.shift_cols_from(0U, 16384U);
  EXPECT_TRUE(row.empty());
  EXPECT_EQ(row.stored_count(), 0U);
}

TEST(RowCellsTest, EraseColsDropsTheBandAndDecrementsLaterKeys) {
  RowCells row;
  for (const std::uint32_t col : {1U, 4U, 5U, 7U, 8U, 20U}) {
    Put(row, col);
  }
  row.erase_cols(4U, 4U);  // band [4, 8)
  std::vector<std::uint32_t> keys;
  std::vector<double> values;
  for (const auto& [col, cell] : row) {
    keys.push_back(col);
    values.push_back(cell.cached_value.as_number());
  }
  EXPECT_EQ(keys, (std::vector<std::uint32_t>{1U, 4U, 16U}));
  EXPECT_EQ(values, (std::vector<double>{1.0, 8.0, 20.0}));

  // A band reaching past the last column removes the tail.
  row.erase_cols(2U, 16383U);
  EXPECT_EQ(row.stored_count(), 1U);
  EXPECT_NE(row.find(1U), nullptr);
}

TEST(RowCellsTest, EntriesInAndCountInAreHalfOpen) {
  RowCells row;
  for (const std::uint32_t col : {3U, 5U, 9U, 16383U}) {
    Put(row, col);
  }
  const auto cols_in = [&](std::uint32_t first, std::uint64_t end) {
    std::vector<std::uint32_t> out;
    for (const auto& [col, cell] : row.entries_in(first, end)) {
      (void)cell;
      out.push_back(col);
    }
    EXPECT_EQ(out.size(), row.count_in(first, end));
    return out;
  };
  EXPECT_EQ(cols_in(3U, 9U), (std::vector<std::uint32_t>{3U, 5U}));
  EXPECT_EQ(cols_in(4U, 10U), (std::vector<std::uint32_t>{5U, 9U}));
  EXPECT_EQ(cols_in(5U, 5U), (std::vector<std::uint32_t>{}));
  EXPECT_EQ(cols_in(6U, 9U), (std::vector<std::uint32_t>{}));
  EXPECT_EQ(cols_in(9U, 2U), (std::vector<std::uint32_t>{}));
  EXPECT_EQ(cols_in(0U, 16384U), (std::vector<std::uint32_t>{3U, 5U, 9U, 16383U}));
  EXPECT_EQ(cols_in(16383U, 16384U), (std::vector<std::uint32_t>{16383U}));
  EXPECT_EQ(cols_in(0U, std::uint64_t{1} << 40U), (std::vector<std::uint32_t>{3U, 5U, 9U, 16383U}));
  EXPECT_EQ(row.count_in(0U, 3U), 0U);
}

TEST(RowCellsTest, EmptyRowHoldsNothing) {
  RowCells row;
  EXPECT_TRUE(row.empty());
  EXPECT_EQ(row.stored_count(), 0U);
  EXPECT_EQ(row.find(0U), nullptr);
  EXPECT_EQ(row.count_in(0U, 16384U), 0U);
  EXPECT_TRUE(row.entries_in(0U, 16384U).empty());
  EXPECT_TRUE(row.begin() == row.end());
}

}  // namespace
}  // namespace formulon
