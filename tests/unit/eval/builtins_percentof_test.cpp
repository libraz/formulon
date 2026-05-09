// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// End-to-end tests for the PERCENTOF aggregate. PERCENTOF is implemented
// as a lazy impl (see `eval/builtins/aggregate.cpp` and
// `tree_walker.cpp`'s `kLazyDispatch` table) because it must compute
// per-argument totals — the eager dispatch path concatenates every
// flattened argument into a single values vector before invoking the
// impl, which would erase the boundary between `data_subset` and
// `data_all`. The semantics still mirror SUM: range-sourced cells skip
// Bool / Text / Blank, while direct scalar arguments coerce strictly.

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
#include "test_eval_helpers.h"
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
  return evaluate(*root, eval_arena, default_registry(), test::mac_context());
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
  const EvalContext ctx = test::mac_context(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

TEST(BuiltinsPercentof, ScalarBasic) {
  // 2 / 10 = 0.2.
  const Value v = EvalSource("=PERCENTOF(2, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.2);
}

TEST(BuiltinsPercentof, ArrayLiteralSubsetAndAll) {
  // SUM({2,3}) / SUM({2,3,5,10}) = 5 / 20 = 0.25.
  const Value v = EvalSource("=PERCENTOF({2,3}, {2,3,5,10})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.25);
}

TEST(BuiltinsPercentof, AllRatioOne) {
  // Subset == All -> 1.0.
  const Value v = EvalSource("=PERCENTOF({1,2,3}, {1,2,3})");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsPercentof, ZeroNumeratorIsZero) {
  // 0 / 100 = 0; not #DIV/0!.
  const Value v = EvalSource("=PERCENTOF(0, 100)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsPercentof, ZeroDenominatorIsDiv0) {
  // SUM(data_all) == 0 -> #DIV/0!.
  const Value v = EvalSource("=PERCENTOF(50, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsPercentof, AllTextInDenominatorIsDiv0) {
  // Text cells inside an array literal are skipped (range-style filter),
  // so SUM({"a","b"}) == 0 -> #DIV/0!.
  const Value v = EvalSource("=PERCENTOF({1,2}, {\"a\",\"b\"})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsPercentof, AllBoolInDenominatorIsDiv0) {
  // Boolean cells inside an array literal are skipped (range-style filter),
  // so SUM({TRUE,TRUE}) == 0 -> #DIV/0!.
  const Value v = EvalSource("=PERCENTOF({1,2}, {TRUE,TRUE})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsPercentof, DirectScalarBoolCoerces) {
  // Direct scalar TRUE coerces to 1: 2 / 1 = 2.
  const Value v = EvalSource("=PERCENTOF(2, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsPercentof, DirectScalarTextNumericCoerces) {
  // Direct scalar numeric-looking text coerces: 2 / 10 = 0.2.
  const Value v = EvalSource("=PERCENTOF(2, \"10\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.2);
}

TEST(BuiltinsPercentof, DirectScalarTextNonNumericIsValue) {
  const Value v = EvalSource("=PERCENTOF(2, \"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsPercentof, SubsetErrorPropagates) {
  // Leftmost error wins: subset is #N/A.
  const Value v = EvalSource("=PERCENTOF(#N/A, 10)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsPercentof, AllErrorPropagates) {
  // Subset OK; data_all is #DIV/0!.
  const Value v = EvalSource("=PERCENTOF(10, #DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsPercentof, BothErrorSubsetWins) {
  // Both args error: subset's code wins (#N/A leftmost).
  const Value v = EvalSource("=PERCENTOF(#N/A, #DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsPercentof, ArityOneIsValue) {
  const Value v = EvalSource("=PERCENTOF(2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsPercentof, ArityThreeIsValue) {
  const Value v = EvalSource("=PERCENTOF(2, 3, 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsPercentof, RangesFromCells) {
  // A1=10, A2=20, A3=30 (sum 60); B1=2, B2=3 (sum 5).
  // PERCENTOF(B1:B2, A1:A3) = 5/60 ~= 0.08333...
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(3.0));
  const Value v = EvalSourceIn("=PERCENTOF(B1:B2, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0 / 60.0);
}

TEST(BuiltinsPercentof, RangeSkipsTextAndBool) {
  // A1=2, A2="x" (skipped), A3=TRUE (skipped), A4=8 -> SUM(A1:A4)=10.
  // B1=2, B2=3 -> SUM(B1:B2)=5. Result = 5/10 = 0.5.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(3, 0, Value::number(8.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(3.0));
  const Value v = EvalSourceIn("=PERCENTOF(B1:B2, A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(BuiltinsPercentof, ErrorInRangePropagates) {
  // A1=3, A2=#DIV/0!, A3=5 -> error in data_all propagates.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=PERCENTOF(2, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
