//
// A `:` chain names the bounding box of its endpoints however it is
// spelled, and FORMULATEXT reports the spelling as written, including
// after a column insert rewrites the formula.

#include <cstdint>
#include <string>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

Value Evaluate(Workbook& wb, const char* formula) {
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 9U, 20U, formula))) << formula;
  EXPECT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  const Cell* cell = wb.sheet(0).cell_at(9U, 20U);
  return cell == nullptr ? Value::blank() : cell->cached_value;
}

// A1:G5 hold row * 10 + column (1-based).
Workbook SeededBook() {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 5U; ++r) {
    for (std::uint32_t c = 0; c < 7U; ++c) {
      EXPECT_TRUE(
          static_cast<bool>(wb.set_cell_value(0U, r, c, Value::number(static_cast<double>((r + 1U) * 10U + c + 1U)))));
    }
  }
  return wb;
}

TEST(RangeChainEval, ShapesNameTheBoundingBox) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=COLUMNS(A:C:E:G)", 7.0},     {"=COLUMNS(G:E:C:A)", 7.0},     {"=COLUMNS(H:H:E:E)", 4.0},
      {"=COLUMNS(H:E)", 4.0},         {"=ROWS(A1:C3:E5)", 5.0},       {"=COLUMNS(A1:C3:E5)", 5.0},
      {"=COLUMNS(A:C:1:3)", 16384.0}, {"=ROWS(A:C:1:3)", 1048576.0},  {"=COLUMNS(A:C:E1)", 5.0},
      {"=ROWS(A:C:E1)", 1048576.0},   {"=COLUMNS(A1:B2:C3:D4)", 4.0},
  };
  Workbook wb = SeededBook();
  for (const Case& c : cases) {
    const Value v = Evaluate(wb, c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(RangeChainEval, SumsReadTheBoundingBox) {
  Workbook wb = SeededBook();
  // A1:E5 sums to 5 * (11 + ... + 15) + 10 * 5 * (0 + ... + 4).
  const double a1_e5 = 5.0 * (11.0 + 12.0 + 13.0 + 14.0 + 15.0) + 10.0 * 5.0 * 10.0;
  for (const char* formula : {"=SUM(A1:C3:E5)", "=SUM(E5:A1)", "=SUM(A1:B2:C3:D4:E5)", "=SUM(A1:E5)"}) {
    const Value v = Evaluate(wb, formula);
    ASSERT_TRUE(v.is_number()) << formula;
    EXPECT_EQ(v.as_number(), a1_e5) << formula;
  }
  const Value cols = Evaluate(wb, "=SUM(A:A:C:C)");
  const Value same = Evaluate(wb, "=SUM(A:C)");
  ASSERT_TRUE(cols.is_number());
  ASSERT_TRUE(same.is_number());
  EXPECT_EQ(cols.as_number(), same.as_number());
}

TEST(RangeChainEval, FormulaTextReportsTheTypedSpelling) {
  Workbook wb = SeededBook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 9U, 9U, "=SUM(H:H:E:E)+SUM(A:C:E:G)")));
  const Value v = Evaluate(wb, "=FORMULATEXT(J10)");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "=SUM(H:H:E:E)+SUM(A:C:E:G)");
}

TEST(RangeChainEval, ColumnInsertRewritesEachToken) {
  Workbook wb = SeededBook();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 9U, 9U, "=SUM(H:H:E:E)+SUM(A:C:E:G)")));
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0U, 3U, 1U)));
  const Cell* cell = wb.sheet(0).cell_at(9U, 10U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=SUM(I:I:F:F)+SUM(A:C:F:H)");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
