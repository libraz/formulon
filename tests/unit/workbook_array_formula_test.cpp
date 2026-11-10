// Authoring and reading legacy array (CSE) formulas through `Workbook`.

#include <cstdint>
#include <memory>
#include <string>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "table.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

Workbook OneSheet() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  return wb;
}

void Recalc(Workbook& wb) {
  ASSERT_TRUE(wb.recalc(eval::default_registry()));
}

Value At(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  return wb.sheet(0).resolve_cell_value(row, col);
}

void FillA1A3(Workbook& wb) {
  for (std::uint32_t r = 0; r < 3; ++r) {
    ASSERT_TRUE(wb.set_cell_value(0, r, 0, Value::number(static_cast<double>(r + 1))));
  }
}

TEST(WorkbookArrayFormula, SetDistributesValuesAndEveryCellReadsTheSameBlock) {
  Workbook wb = OneSheet();
  FillA1A3(wb);
  ASSERT_TRUE(wb.set_array_formula(0, 0, 2, 3, 1, "=A1:A3*2"));
  Recalc(wb);
  EXPECT_EQ(At(wb, 0, 2), Value::number(2.0));
  EXPECT_EQ(At(wb, 1, 2), Value::number(4.0));
  EXPECT_EQ(At(wb, 2, 2), Value::number(6.0));
  for (std::uint32_t r = 0; r < 3; ++r) {
    const auto info = wb.array_formula_at(0, r, 2);
    ASSERT_TRUE(info.has_value()) << r;
    EXPECT_EQ(info->anchor_row, 0U);
    EXPECT_EQ(info->anchor_col, 2U);
    EXPECT_EQ(info->rows, 3U);
    EXPECT_EQ(info->cols, 1U);
    EXPECT_EQ(info->formula, wb.sheet(0).cell_at(0, 2)->formula_text);
  }
  EXPECT_FALSE(wb.array_formula_at(0, 3, 2).has_value());
  EXPECT_FALSE(wb.array_formula_at(0, 0, 3).has_value());
  EXPECT_FALSE(wb.array_formula_at(5, 0, 2).has_value());
  const Cell* anchor = wb.sheet(0).cell_at(0, 2);
  ASSERT_NE(anchor, nullptr);
  EXPECT_FALSE(anchor->dynamic_array);
}

TEST(WorkbookArrayFormula, EveryCellReadsTheBlockBeforeARecalc) {
  Workbook wb = OneSheet();
  ASSERT_TRUE(wb.set_array_formula(0, 1, 1, 2, 2, "=1"));
  for (std::uint32_t r = 1; r < 3; ++r) {
    for (std::uint32_t c = 1; c < 3; ++c) {
      const auto info = wb.array_formula_at(0, r, c);
      ASSERT_TRUE(info.has_value()) << r << "," << c;
      EXPECT_EQ(info->anchor_row, 1U);
      EXPECT_EQ(info->anchor_col, 1U);
    }
  }
  EXPECT_FALSE(wb.array_formula_at(0, 3, 1).has_value());
  EXPECT_FALSE(wb.array_formula_at(0, 1, 3).has_value());
  EXPECT_FALSE(wb.array_formula_at(0, 0, 0).has_value());
}

TEST(WorkbookArrayFormula, SingleCellBlockIsAllowed) {
  Workbook wb = OneSheet();
  FillA1A3(wb);
  ASSERT_TRUE(wb.set_array_formula(0, 0, 2, 1, 1, "=SUM(A1:A3*2)"));
  Recalc(wb);
  EXPECT_EQ(At(wb, 0, 2), Value::number(12.0));
  const auto info = wb.array_formula_at(0, 0, 2);
  ASSERT_TRUE(info.has_value());
  EXPECT_EQ(info->rows, 1U);
  EXPECT_EQ(info->cols, 1U);
}

TEST(WorkbookArrayFormula, ShortResultFillsNotAvailable) {
  Workbook wb = OneSheet();
  FillA1A3(wb);
  ASSERT_TRUE(wb.set_array_formula(0, 0, 2, 5, 1, "=A1:A3*2"));
  Recalc(wb);
  EXPECT_EQ(At(wb, 2, 2), Value::number(6.0));
  EXPECT_EQ(At(wb, 3, 2), Value::error(ErrorCode::NA));
  EXPECT_EQ(At(wb, 4, 2), Value::error(ErrorCode::NA));
}

TEST(WorkbookArrayFormula, ScalarResultRepeatsAcrossTheBlock) {
  Workbook wb = OneSheet();
  ASSERT_TRUE(wb.set_array_formula(0, 0, 0, 2, 2, "=7"));
  Recalc(wb);
  for (std::uint32_t r = 0; r < 2; ++r) {
    for (std::uint32_t c = 0; c < 2; ++c) {
      EXPECT_EQ(At(wb, r, c), Value::number(7.0));
    }
  }
}

TEST(WorkbookArrayFormula, SettingClearsExistingValuesAndKeepsStyles) {
  Workbook wb = OneSheet();
  FillA1A3(wb);
  ASSERT_TRUE(wb.set_cell_value(0, 1, 2, Value::number(99.0)));
  ASSERT_TRUE(wb.set_cell_formula(0, 2, 2, "=1+1"));
  ASSERT_TRUE(wb.set_cell_xf_index(0, 1, 2, 0U));
  ASSERT_TRUE(wb.set_array_formula(0, 0, 2, 3, 1, "=A1:A3*2"));
  Recalc(wb);
  EXPECT_EQ(At(wb, 1, 2), Value::number(4.0));
  EXPECT_EQ(At(wb, 2, 2), Value::number(6.0));
  EXPECT_TRUE(wb.sheet(0).cell_at(2, 2) == nullptr || wb.sheet(0).cell_at(2, 2)->formula_text.empty());
}

TEST(WorkbookArrayFormula, WriteIntoNonAnchorIsRefusedAndAnchorBlankClearsTheBlock) {
  Workbook wb = OneSheet();
  FillA1A3(wb);
  ASSERT_TRUE(wb.set_array_formula(0, 0, 2, 3, 1, "=A1:A3*2"));
  Recalc(wb);
  const auto refused = wb.set_cell_value(0, 1, 2, Value::number(1.0));
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_NE(refused.error().message.find("Cannot change part of an array"), std::string::npos);

  ASSERT_TRUE(wb.set_cell_value(0, 0, 2, Value::blank()));
  Recalc(wb);
  for (std::uint32_t r = 0; r < 3; ++r) {
    EXPECT_FALSE(wb.array_formula_at(0, r, 2).has_value()) << r;
    EXPECT_TRUE(At(wb, r, 2).is_blank()) << r;
  }
  // The cleared cells are ordinary again.
  EXPECT_TRUE(wb.set_cell_value(0, 1, 2, Value::number(1.0)));
}

TEST(WorkbookArrayFormula, OrdinaryFormulaOnTheAnchorReplacesTheWholeBlock) {
  Workbook wb = OneSheet();
  FillA1A3(wb);
  ASSERT_TRUE(wb.set_array_formula(0, 0, 2, 3, 1, "=A1:A3*2"));
  Recalc(wb);
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 2, "=5"));
  Recalc(wb);
  EXPECT_FALSE(wb.array_formula_at(0, 0, 2).has_value());
  EXPECT_EQ(At(wb, 0, 2), Value::number(5.0));
  EXPECT_TRUE(At(wb, 1, 2).is_blank());
}

TEST(WorkbookArrayFormula, PartialOverlapIsRefusedAndLeavesTheSheetUnchanged) {
  Workbook wb = OneSheet();
  FillA1A3(wb);
  ASSERT_TRUE(wb.set_array_formula(0, 0, 2, 3, 1, "=A1:A3*2"));
  Recalc(wb);
  // Rows 2..3 of the existing block C1:C3, and a larger block over its tail.
  struct Rect {
    std::uint32_t row, col, rows, cols;
  };
  for (const Rect rect : {Rect{1, 2, 1, 1}, Rect{2, 1, 2, 2}}) {
    const auto refused = wb.set_array_formula(0, rect.row, rect.col, rect.rows, rect.cols, "=1");
    ASSERT_FALSE(static_cast<bool>(refused)) << rect.row << "," << rect.col;
    EXPECT_EQ(refused.error().code, FormulonErrorCode::kPreconditionFailed);
    EXPECT_NE(refused.error().message.find("overlaps part of an array formula"), std::string::npos);
  }
  Recalc(wb);
  const auto info = wb.array_formula_at(0, 1, 2);
  ASSERT_TRUE(info.has_value());
  EXPECT_EQ(info->rows, 3U);
  EXPECT_EQ(At(wb, 2, 2), Value::number(6.0));
  EXPECT_EQ(wb.sheet(0).cell_at(2, 1), nullptr);
}

TEST(WorkbookArrayFormula, BlockWhollyInsideTheRectangleIsReplaced) {
  Workbook wb = OneSheet();
  FillA1A3(wb);
  ASSERT_TRUE(wb.set_array_formula(0, 1, 3, 2, 1, "=A1:A2"));
  Recalc(wb);
  ASSERT_TRUE(wb.set_array_formula(0, 0, 2, 4, 3, "=A1:A3+1"));
  Recalc(wb);
  const auto inner = wb.array_formula_at(0, 2, 3);
  ASSERT_TRUE(inner.has_value());
  EXPECT_EQ(inner->anchor_row, 0U);
  EXPECT_EQ(inner->anchor_col, 2U);
  EXPECT_EQ(inner->rows, 4U);
  EXPECT_EQ(inner->cols, 3U);
  EXPECT_EQ(At(wb, 0, 3), Value::number(2.0));
}

TEST(WorkbookArrayFormula, MergedAreaIsRefused) {
  Workbook wb = OneSheet();
  ASSERT_TRUE(wb.add_merge(0, MergeRange{1, 2, 1, 3}));
  const auto refused = wb.set_array_formula(0, 0, 2, 3, 1, "=1");
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_NE(refused.error().message.find("array formulas are not valid in merged cells"), std::string::npos);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 2), nullptr);
}

TEST(WorkbookArrayFormula, MultiCellOverATableIsRefusedButOneCellIsAllowed) {
  Workbook wb = OneSheet();
  TableMetadata table;
  table.id = 1;
  table.name = "T1";
  table.display_name = "T1";
  table.ref = "A1:B4";
  table.sheet_index = 0;
  wb.mutable_tables().push_back(table);
  const auto refused = wb.set_array_formula(0, 1, 0, 2, 1, "=1");
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_TRUE(wb.set_array_formula(0, 1, 0, 1, 1, "=1"));
  EXPECT_TRUE(wb.set_array_formula(0, 5, 0, 2, 1, "=1"));
}

TEST(WorkbookArrayFormula, DataTableResultRangeIsRefused) {
  Workbook wb = OneSheet();
  ASSERT_TRUE(wb.set_cell_formula(0, 0, 3, "=A1*2"));
  DataTableInput input;
  input.present = true;
  input.row = 0;
  input.col = 0;
  for (std::uint32_t k = 0; k < 3; ++k) {
    ASSERT_TRUE(wb.set_cell_value(0, 1 + k, 2, Value::number(1.0 + k)));
  }
  ASSERT_TRUE(wb.add_data_table(0, MergeRange{1, 3, 3, 3}, DataTableInput{}, input));
  const auto refused = wb.set_array_formula(0, 3, 3, 2, 1, "=1");
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kPreconditionFailed);
}

TEST(WorkbookArrayFormula, PivotTableAreaIsRefused) {
  Workbook wb = OneSheet();
  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_anchor(5, 5, 4, 3);
  wb.sheet(0).add_pivot_table(std::move(pivot));
  const auto refused = wb.set_array_formula(0, 7, 6, 1, 1, "=1");
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_TRUE(wb.set_array_formula(0, 9, 5, 1, 1, "=1"));
}

TEST(WorkbookArrayFormula, InvalidArgumentsAreRefused) {
  Workbook wb = OneSheet();
  EXPECT_EQ(wb.set_array_formula(3, 0, 0, 1, 1, "=1").error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.set_array_formula(0, Sheet::kMaxRows, 0, 1, 1, "=1").error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.set_array_formula(0, 0, 0, 0, 1, "=1").error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.set_array_formula(0, 0, 0, 1, 0, "=1").error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.set_array_formula(0, Sheet::kMaxRows - 1U, 0, 2, 1, "=1").error().code,
            FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.set_array_formula(0, 0, Sheet::kMaxCols - 1U, 1, 2, "=1").error().code,
            FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0), nullptr);
}

TEST(WorkbookArrayFormula, UnparseableFormulaIsStoredLikeTheOrdinarySetter) {
  Workbook wb = OneSheet();
  ASSERT_TRUE(wb.set_array_formula(0, 0, 0, 2, 1, "=1+"));
  Recalc(wb);
  EXPECT_EQ(At(wb, 0, 0), Value::error(ErrorCode::Name));
  EXPECT_EQ(At(wb, 1, 0), Value::error(ErrorCode::Name));
}

TEST(WorkbookArrayFormula, BlockSurvivesAnXlsxRoundTrip) {
  Workbook wb = OneSheet();
  FillA1A3(wb);
  ASSERT_TRUE(wb.set_array_formula(0, 0, 2, 3, 1, "=A1:A3*2"));
  Recalc(wb);
  auto saved = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved));
  const std::vector<std::uint8_t> bytes = saved.value();
  auto loaded = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& dst = loaded.value().workbook;
  ASSERT_TRUE(dst.recalc(eval::default_registry()));
  for (std::uint32_t r = 0; r < 3; ++r) {
    const auto info = dst.array_formula_at(0, r, 2);
    ASSERT_TRUE(info.has_value()) << r;
    EXPECT_EQ(info->anchor_row, 0U);
    EXPECT_EQ(info->anchor_col, 2U);
    EXPECT_EQ(info->rows, 3U);
    EXPECT_EQ(info->cols, 1U);
    EXPECT_EQ(info->formula, wb.array_formula_at(0, 0, 2)->formula);
  }
  EXPECT_EQ(dst.sheet(0).resolve_cell_value(2, 2), Value::number(6.0));
}

}  // namespace
}  // namespace formulon
