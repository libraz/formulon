// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the SUMSQ aggregate. SUMSQ rides the same
// eager / range-aware path as SUM: direct scalar arguments coerce
// through `coerce_to_number` (so TRUE -> 1, "5" -> 5, "abc" -> #VALUE!),
// while Bool / Text / Blank cells sourced from a range are silently
// dropped by the dispatcher's `range_filter_numeric_only` flag before
// the impl runs.

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

TEST(BuiltinsSumSq, TwoScalarArguments) {
  // 3^2 + 4^2 = 25.
  const Value v = EvalSource("=SUMSQ(3, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 25.0);
}

TEST(BuiltinsSumSq, Range) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(4.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=SUMSQ(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 50.0);
}

TEST(BuiltinsSumSq, MixedRangeAndScalar) {
  // A1:A2 = {3, 4} contributes 9+16; trailing scalar 5 contributes 25;
  // total 50.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(4.0));
  const Value v = EvalSourceIn("=SUMSQ(A1:A2, 5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 50.0);
}

TEST(BuiltinsSumSq, SkipsTextInRange) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=SUMSQ(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 9.0 + 25.0);
}

TEST(BuiltinsSumSq, SkipsBoolAndBlankInRange) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::boolean(true));  // dropped
  // row 2 left blank on purpose
  wb.sheet(0).set_cell_value(3, 0, Value::number(6.0));
  const Value v = EvalSourceIn("=SUMSQ(A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0 + 36.0);
}

TEST(BuiltinsSumSq, DirectTextCoerces) {
  // Direct text argument: numeric-looking text coerces to the number;
  // 3^2 + 5^2 = 34.
  const Value v = EvalSource("=SUMSQ(3, \"5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 34.0);
}

TEST(BuiltinsSumSq, DirectNonNumericTextIsValue) {
  const Value v = EvalSource("=SUMSQ(3, \"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSumSq, DirectBooleanCoerces) {
  // TRUE -> 1, so =SUMSQ(3, TRUE) = 9 + 1 = 10.
  const Value v = EvalSource("=SUMSQ(3, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsSumSq, ErrorInRangePropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=SUMSQ(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsSumSq, ArrayLiteral) {
  // {1;2;3;4} -> 1+4+9+16 = 30.
  const Value v = EvalSource("=SUMSQ({1;2;3;4})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
