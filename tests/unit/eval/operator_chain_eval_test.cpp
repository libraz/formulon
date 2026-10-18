//
// Evaluation, dependency tracking and structural edits over flat operator
// chains (`=A1+A2+...+An`), whose length is bounded only by the
// 8,192-character formula limit.

#include <algorithm>
#include <cstdint>
#include <string>

#include "cell.h"
#include "dep_extractor_test_helpers.h"
#include "eval/dep_extractor.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using dep_extractor_test_helpers::ParseFormula;

// `A1+A2+...+An`, without the leading `=`.
std::string RefChain(std::uint32_t terms) {
  std::string out = "A1";
  for (std::uint32_t i = 2; i <= terms; ++i) {
    out.append("+A").append(std::to_string(i));
  }
  return out;
}

// Seeds A1:An with 1..n and puts `formula` in B1.
Workbook ChainBook(std::uint32_t terms, const std::string& formula) {
  Workbook wb = Workbook::create();
  for (std::uint32_t i = 0; i < terms; ++i) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0U, i, 0U, Value::number(static_cast<double>(i + 1U)))));
  }
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 1U, formula)));
  EXPECT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  return wb;
}

Value ValueAt(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Cell* cell = wb.sheet(0).cell_at(row, col);
  return cell == nullptr ? Value::blank() : cell->cached_value;
}

double SumTo(std::uint32_t n) {
  return static_cast<double>(n) * static_cast<double>(n + 1U) / 2.0;
}

TEST(OperatorChainEval, OneHundredFiftyTermChainEvaluates) {
  const Workbook wb = ChainBook(150U, "=" + RefChain(150U));
  const Value v = ValueAt(wb, 0U, 1U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 11325.0);
}

// `A1+A2+...+A9+A1+...` with `terms` operands: 2,700 of them fit in
// 8,192 characters.
std::string ShortRefChain(std::uint32_t terms) {
  std::string out = "A1";
  for (std::uint32_t i = 1; i < terms; ++i) {
    out.append("+A").append(std::to_string(i % 9U + 1U));
  }
  return out;
}

TEST(OperatorChainEval, ThousandTermChainEvaluates) {
  const Workbook wb = ChainBook(1000U, "=" + RefChain(1000U));
  const Value v = ValueAt(wb, 0U, 1U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), SumTo(1000U));
}

TEST(OperatorChainEval, ChainNearTheFormulaLengthLimitEvaluates) {
  const std::string formula = "=" + ShortRefChain(2700U);
  ASSERT_LT(formula.size(), 8192U);
  const Workbook wb = ChainBook(9U, formula);
  const Value v = ValueAt(wb, 0U, 1U);
  ASSERT_TRUE(v.is_number());
  // 300 passes over 1..9.
  EXPECT_EQ(v.as_number(), 300.0 * SumTo(9U));
}

TEST(OperatorChainEval, DensestChainEvaluates) {
  std::string formula = "=1";
  for (std::uint32_t i = 2; i <= 4095U; ++i) {
    formula.append("+1");
  }
  ASSERT_LT(formula.size(), 8192U);
  const Workbook wb = ChainBook(0U, formula);
  const Value v = ValueAt(wb, 0U, 1U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4095.0);
}

TEST(OperatorChainEval, MixedPrecedenceChainEvaluates) {
  std::string formula = "=1";
  double expected = 1.0;
  for (std::uint32_t i = 0; i < 700U; ++i) {
    const std::uint32_t a = 2U * i + 2U;
    const std::uint32_t b = 2U * i + 3U;
    formula.append("+").append(std::to_string(a)).append("*").append(std::to_string(b));
    expected += static_cast<double>(a) * static_cast<double>(b);
  }
  ASSERT_LT(formula.size(), 8192U);
  const Workbook wb = ChainBook(0U, formula);
  const Value v = ValueAt(wb, 0U, 1U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), expected);
}

TEST(OperatorChainEval, MixedOperatorsApplyLeftToRight) {
  const Workbook wb = ChainBook(0U, "=10-2+3-4&\"x\"&5");
  const Value v = ValueAt(wb, 0U, 1U);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "7x5");
}

TEST(OperatorChainEval, LeftmostErrorWins) {
  const Workbook wb = ChainBook(0U, "=1+NA()+1/0+#REF!");
  const Value v = ValueAt(wb, 0U, 1U);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(OperatorChainEval, RootCancellationSnapAppliesToTheLastOperator) {
  const Workbook wb = ChainBook(0U, "=0.1+0.2-0.3");
  const Value v = ValueAt(wb, 0U, 1U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(OperatorChainEval, ArrayOperandsBroadcastAcrossTheChain) {
  // (1+2+1) + (2+4+1) + (3+6+1).
  const Workbook wb = ChainBook(3U, "=SUM(A1:A3+A1:A3*2+1)");
  const Value v = ValueAt(wb, 0U, 1U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 21.0);
}

TEST(OperatorChainEval, DependenciesCoverEveryOperand) {
  Workbook wb = Workbook::create();
  Arena arena;
  const std::string text = RefChain(1000U);
  const parser::AstNode* root = ParseFormula(text, arena);
  ASSERT_NE(root, nullptr);
  const ExtractedDeps deps = extract_deps(*root, /*current_sheet_id=*/0U, wb);
  EXPECT_EQ(deps.cell_deps.size(), 1000U);
  EXPECT_NE(std::find(deps.cell_deps.begin(), deps.cell_deps.end(), (CellNodeId{0U, 999U, 0U})), deps.cell_deps.end());
  EXPECT_FALSE(deps.is_volatile);
}

TEST(OperatorChainEval, EditingTheLastOperandRecalculatesTheChain) {
  Workbook wb = ChainBook(1000U, "=" + RefChain(1000U));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 999U, 0U, Value::number(0.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  const Value v = ValueAt(wb, 0U, 1U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), SumTo(1000U) - 1000.0);
}

TEST(OperatorChainEval, RowInsertRewritesTheChainAndKeepsItsValue) {
  Workbook wb = ChainBook(150U, "=" + RefChain(150U));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0U, 100U, 2U)));
  const Cell* cell = wb.sheet(0).cell_at(0U, 1U);
  ASSERT_NE(cell, nullptr);
  const std::string& text = cell->formula_text;
  EXPECT_NE(text.find("+A100+A103+"), std::string::npos);
  EXPECT_EQ(text.substr(text.size() - 5U), "+A152");
  ASSERT_TRUE(static_cast<bool>(wb.recalc(default_registry())));
  const Value v = ValueAt(wb, 0U, 1U);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 11325.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
