// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the SERIESSUM power-series builtin.
//
// Formula: SERIESSUM(x, n, m, coefficients) evaluates
//   Σᵢ aᵢ · x^(n + (i-1)·m)   for i = 1..k
// so the first coefficient is paired with x^n, the second with
// x^(n+m), and so on.

#include <cmath>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

Value EvalSource(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena);
}

Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  EvalState state;
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

TEST(BuiltinsSeriesSum, BinaryPolynomial) {
  // 1·2^0 + 1·2^1 + 1·2^2 + 1·2^3 = 1 + 2 + 4 + 8 = 15.
  const Value v = EvalSource("=SERIESSUM(2, 0, 1, {1;1;1;1})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(BuiltinsSeriesSum, AllPowersOfOneAreOne) {
  // x = 1 makes every x^power term equal 1, so the result is just the
  // sum of coefficients: 5 + 5 + 5 = 15.
  const Value v = EvalSource("=SERIESSUM(1, 0, 1, {5;5;5})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(BuiltinsSeriesSum, StepTwoSkipsOddPowers) {
  // 1·2^2 + 1·2^4 = 4 + 16 = 20.
  const Value v = EvalSource("=SERIESSUM(2, 2, 2, {1;1})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);
}

TEST(BuiltinsSeriesSum, CoefficientsFromRange) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  // SERIESSUM(2, 0, 1, A1:A3) = 1·2^0 + 2·2^1 + 3·2^2 = 1 + 4 + 12 = 17.
  const Value v = EvalSourceIn("=SERIESSUM(2, 0, 1, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 17.0);
}

TEST(BuiltinsSeriesSum, NonNumericXIsValue) {
  const Value v = EvalSource("=SERIESSUM(\"abc\", 0, 1, {1;1})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSeriesSum, NonNumericNIsValue) {
  const Value v = EvalSource("=SERIESSUM(2, \"abc\", 1, {1;1})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSeriesSum, NonNumericMIsValue) {
  const Value v = EvalSource("=SERIESSUM(2, 0, \"abc\", {1;1})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSeriesSum, ErrorInCoefficientPropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  const Value v = EvalSourceIn("=SERIESSUM(2, 0, 1, A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsSeriesSum, ArityUnder) {
  const Value v = EvalSource("=SERIESSUM(2, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSeriesSum, ArityOver) {
  const Value v = EvalSource("=SERIESSUM(2, 0, 1, {1;1}, {1;1})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSeriesSum, ErrorInXShortCircuits) {
  // #DIV/0! in x beats a later #VALUE! in n: left-to-right.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::error(ErrorCode::Div0));
  const Value v = EvalSourceIn("=SERIESSUM(A1, \"abc\", 1, {1;1})", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
