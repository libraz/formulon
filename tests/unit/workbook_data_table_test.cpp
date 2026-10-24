// What-If data table calculation and creation preconditions on `Workbook`.

#include <cstdint>

#include "data_table.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

DataTableInput InputAt(std::uint32_t row, std::uint32_t col) {
  DataTableInput input;
  input.present = true;
  input.row = row;
  input.col = col;
  return input;
}

Workbook OneSheet() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  return wb;
}

TEST(WorkbookDataTable, SecondTableHeadersReadTheFirstTablesInputAfterItIsRestored) {
  Workbook wb = OneSheet();
  ASSERT_TRUE(wb.set_cell_value(0, 0, 0, Value::number(1.0)));
  ASSERT_TRUE(wb.set_cell_value(0, 1, 0, Value::number(5.0)));
  // Table A: D2:D4 over input A1.
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 3, "=A1*2"));
  for (std::uint32_t k = 0; k < 3; ++k) {
    ASSERT_TRUE(wb.set_cell_value(0, 1 + k, 2, Value::number(10.0 * (k + 1))));
  }
  ASSERT_TRUE(wb.add_data_table(0, MergeRange{1, 3, 3, 3}, DataTableInput{}, InputAt(0, 0)));
  // Table B: H2:H4 over input A2, with left headers that read A1.
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 7, "=A2*2"));
  for (std::uint32_t k = 0; k < 3; ++k) {
    ASSERT_TRUE(wb.set_cell_formula(0, 1 + k, 6, "=A1"));
  }
  ASSERT_TRUE(wb.add_data_table(0, MergeRange{1, 7, 3, 7}, DataTableInput{}, InputAt(1, 0)));
  ASSERT_TRUE(wb.recalc(eval::default_registry()));
  for (std::uint32_t k = 0; k < 3; ++k) {
    const Cell* cell = wb.sheet(0).cell_at(1 + k, 7);
    ASSERT_NE(cell, nullptr);
    ASSERT_TRUE(cell->cached_value.is_number());
    EXPECT_EQ(cell->cached_value.as_number(), 2.0) << "row " << k;
  }
}

TEST(WorkbookDataTable, RefusesARefLargerThanTheCellCeiling) {
  Workbook wb = OneSheet();
  auto added = wb.add_data_table(0, MergeRange{1, 1, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U}, DataTableInput{},
                                 InputAt(0, 0));
  ASSERT_FALSE(static_cast<bool>(added));
  EXPECT_EQ(added.error().code, FormulonErrorCode::kSecResourceLimit);
  EXPECT_EQ(wb.sheet(0).cell_at(1, 1), nullptr);
  EXPECT_TRUE(wb.sheet(0).data_tables().empty());
}

TEST(WorkbookDataTable, ATableOverTheRecalcPassCeilingKeepsItsCachedResults) {
  Workbook wb = OneSheet();
  // A 300 x 300 two-variable table needs 90,001 passes, over the ceiling.
  ASSERT_TRUE(wb.set_cell_formula(0, 1, 1, "=A1+A2+1"));
  ASSERT_TRUE(wb.add_data_table(0, MergeRange{2, 2, 301, 301}, InputAt(0, 0), InputAt(1, 0)));
  // A small table after it still computes from the remaining budget.
  ASSERT_TRUE(wb.set_cell_formula(0, 1, 400, "=A1+A2+1"));
  ASSERT_TRUE(wb.add_data_table(0, MergeRange{2, 401, 3, 402}, InputAt(0, 0), InputAt(1, 0)));
  ASSERT_TRUE(wb.recalc(eval::default_registry()));
  const Cell* skipped = wb.sheet(0).cell_at(2, 2);
  ASSERT_NE(skipped, nullptr);
  ASSERT_TRUE(skipped->cached_value.is_number());
  EXPECT_EQ(skipped->cached_value.as_number(), 0.0);
  const Cell* computed = wb.sheet(0).cell_at(2, 401);
  ASSERT_NE(computed, nullptr);
  ASSERT_TRUE(computed->cached_value.is_number());
  EXPECT_EQ(computed->cached_value.as_number(), 1.0);
}

TEST(WorkbookDataTable, RefusesARefOverALegacyArrayBlock) {
  Workbook wb = OneSheet();
  ASSERT_TRUE(wb.set_cell_formula(0, 5, 3, "=1"));
  wb.sheet(0).set_cell_array_block(5, 3, 2, 1);
  auto added = wb.add_data_table(0, MergeRange{4, 3, 6, 3}, DataTableInput{}, InputAt(0, 0));
  ASSERT_FALSE(static_cast<bool>(added));
  EXPECT_EQ(added.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.sheet(0).cell_at(4, 3), nullptr);
  EXPECT_TRUE(wb.sheet(0).data_tables().empty());
}

TEST(WorkbookDataTable, RefusesARefOverASpill) {
  Workbook wb = OneSheet();
  ASSERT_TRUE(wb.set_cell_formula(0, 5, 3, "=SEQUENCE(3)"));
  ASSERT_TRUE(wb.recalc(eval::default_registry()));
  ASSERT_NE(wb.sheet(0).spill_region_covering(6, 3), nullptr);
  auto added = wb.add_data_table(0, MergeRange{6, 3, 7, 3}, DataTableInput{}, InputAt(0, 0));
  ASSERT_FALSE(static_cast<bool>(added));
  EXPECT_EQ(added.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_TRUE(wb.sheet(0).data_tables().empty());
}

TEST(WorkbookDataTable, AnInputInsideASpillYieldsRefErrors) {
  Workbook wb = OneSheet();
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 3, "=SEQUENCE(3)"));
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 5, "=D2*2"));
  ASSERT_TRUE(wb.set_cell_value(0, 1, 4, Value::number(7.0)));
  ASSERT_TRUE(wb.set_cell_value(0, 2, 4, Value::number(8.0)));
  ASSERT_TRUE(wb.recalc(eval::default_registry()));
  ASSERT_NE(wb.sheet(0).spill_region_covering(1, 3), nullptr);
  ASSERT_TRUE(wb.add_data_table(0, MergeRange{1, 5, 2, 5}, DataTableInput{}, InputAt(1, 3)));
  ASSERT_TRUE(wb.recalc(eval::default_registry()));
  for (std::uint32_t row = 1; row <= 2; ++row) {
    const Cell* cell = wb.sheet(0).cell_at(row, 5);
    ASSERT_NE(cell, nullptr);
    ASSERT_TRUE(cell->cached_value.is_error()) << "row " << row;
    EXPECT_EQ(cell->cached_value.as_error(), ErrorCode::Ref);
  }
}

TEST(WorkbookDataTable, AnInputThatWasAbsentIsAbsentAgainAfterTheTableRuns) {
  Workbook wb = OneSheet();
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 5, "=J10*2"));
  ASSERT_TRUE(wb.set_cell_value(0, 1, 4, Value::number(7.0)));
  ASSERT_TRUE(wb.add_data_table(0, MergeRange{1, 5, 1, 5}, DataTableInput{}, InputAt(9, 9)));
  ASSERT_TRUE(wb.recalc(eval::default_registry()));
  const Cell* result = wb.sheet(0).cell_at(1, 5);
  ASSERT_NE(result, nullptr);
  EXPECT_EQ(result->cached_value.as_number(), 14.0);
  EXPECT_EQ(wb.sheet(0).cell_at(9, 9), nullptr);
}

}  // namespace
}  // namespace formulon
