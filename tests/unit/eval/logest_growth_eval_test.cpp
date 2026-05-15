// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for the LOGEST and GROWTH exponential-regression lazy builtins
// (`linest_lazy.{h,cpp}`). Both fit `y = b * m_1^x_1 * ...` by running
// LINEST on `ln(y)`. LOGEST exponentiates the row-1 coefficients and
// keeps the rest of the stats matrix on the linear scale; GROWTH
// exponentiates the predictions.
//
// Reference values are computed analytically in the test comments
// using the inverse: `ln(y) = ln(b) + x*ln(m)`. With a perfect-fit
// data set the returned coefficients should agree with the analytic
// `b` and `m` to within FP tolerance.

#include <cmath>
#include <cstdint>
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

Value EvalUnder(std::string_view src, Arena* parse_arena, Arena* eval_arena, const EvalContext& ctx) {
  parser::Parser p(src, *parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr);
  EXPECT_TRUE(p.errors().empty());
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, *eval_arena, default_registry(), ctx);
}

constexpr double kLooseEps = 1e-7;

// ---------------------------------------------------------------------------
// LOGEST
// ---------------------------------------------------------------------------

TEST(BuiltinsLogest, PerfectFitSingleVariable) {
  // y = 3 * 2^x: at x=1..4 -> y = 6, 12, 24, 48. LOGEST should recover
  // [m, b] = [2, 3]. (Output is 1 x 2: [m, b] in right-to-left order.)
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LOGEST({6,12,24,48}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kLooseEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 3.0, kLooseEps);
}

TEST(BuiltinsLogest, DefaultKnownX) {
  // No known_x -> defaults to {1, 2, 3, 4}. y = 3 * 2^x same as above.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LOGEST({6,12,24,48})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kLooseEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 3.0, kLooseEps);
}

TEST(BuiltinsLogest, ConstFalseInterceptSlotIsOne) {
  // y = m^x with no leading b: const=FALSE forces b=1. y = 2^x at
  // x=1..4 -> y = 2, 4, 8, 16. Output [m, b] = [2, 1].
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LOGEST({2,4,8,16}, {1,2,3,4}, FALSE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kLooseEps);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 1.0);
}

TEST(BuiltinsLogest, MultiVariableRecoversCoefficients) {
  // y = 2 * 3^x1 * 5^x2.
  //   x1 = {1, 1, 2, 2}, x2 = {1, 2, 1, 2}
  //   y = 2 * 3^x1 * 5^x2:
  //     (1,1) -> 2*3*5  = 30
  //     (1,2) -> 2*3*25 = 150
  //     (2,1) -> 2*9*5  = 90
  //     (2,2) -> 2*9*25 = 450
  // LOGEST output (1x3): [m_2, m_1, b] = [5, 3, 2].
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // y at A1:A4
  sheet.set_cell_value(0, 0, Value::number(30));
  sheet.set_cell_value(1, 0, Value::number(150));
  sheet.set_cell_value(2, 0, Value::number(90));
  sheet.set_cell_value(3, 0, Value::number(450));
  // x1 at B1:B4, x2 at C1:C4
  sheet.set_cell_value(0, 1, Value::number(1));
  sheet.set_cell_value(1, 1, Value::number(1));
  sheet.set_cell_value(2, 1, Value::number(2));
  sheet.set_cell_value(3, 1, Value::number(2));
  sheet.set_cell_value(0, 2, Value::number(1));
  sheet.set_cell_value(1, 2, Value::number(2));
  sheet.set_cell_value(2, 2, Value::number(1));
  sheet.set_cell_value(3, 2, Value::number(2));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LOGEST(A1:A4, B1:C4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* c = v.as_array_cells();
  EXPECT_NEAR(c[0].as_number(), 5.0, kLooseEps);  // m_2
  EXPECT_NEAR(c[1].as_number(), 3.0, kLooseEps);  // m_1
  EXPECT_NEAR(c[2].as_number(), 2.0, kLooseEps);  // b
}

TEST(BuiltinsLogest, StatsTrueRowOneIsExpRowsTwoToFiveAreLinear) {
  // y = 3 * 2^x at x=1..4. Perfect fit: r^2 = 1, ss_resid = 0,
  // SE row = 0 (perfect fit on log scale).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LOGEST({6,12,24,48}, {1,2,3,4}, TRUE, TRUE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 5U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* c = v.as_array_cells();
  // Row 1: [m, b] = [2, 3] (exponentiated).
  EXPECT_NEAR(c[0].as_number(), 2.0, kLooseEps);
  EXPECT_NEAR(c[1].as_number(), 3.0, kLooseEps);
  // Row 3 col 0: r^2 = 1 (on linear/log scale, perfect fit).
  EXPECT_NEAR(c[4].as_number(), 1.0, kLooseEps);
  // Row 5 col 1: ss_resid = 0.
  EXPECT_NEAR(c[9].as_number(), 0.0, kLooseEps);
}

TEST(BuiltinsLogest, NonPositiveYReturnsNum) {
  // y contains 0 -> ln(0) is undefined -> #NUM!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LOGEST({2,0,8,16}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsLogest, NegativeYReturnsNum) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LOGEST({2,4,-8,16}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsLogest, ErrorInYPropagates) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LOGEST({2,#N/A,8,16}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsLogest, ZeroArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LOGEST()", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsLogest, FiveArgsRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LOGEST({2,4,8}, {1,2,3}, TRUE, TRUE, 99)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// GROWTH
// ---------------------------------------------------------------------------

TEST(BuiltinsGrowth, DefaultNewXReturnsFittedYHat) {
  // y = 3 * 2^x at x=1..4. Fitted values match y exactly (perfect fit).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=GROWTH({6,12,24,48}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_cols(), 4U);
  const Value* c = v.as_array_cells();
  EXPECT_NEAR(c[0].as_number(), 6.0, 1e-6);
  EXPECT_NEAR(c[1].as_number(), 12.0, 1e-6);
  EXPECT_NEAR(c[2].as_number(), 24.0, 1e-6);
  EXPECT_NEAR(c[3].as_number(), 48.0, 1e-6);
}

TEST(BuiltinsGrowth, ExplicitNewXSingleVariable) {
  // y = 3 * 2^x. Predict at x=5 -> 3*32=96; x=6 -> 3*64=192.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=GROWTH({6,12,24,48}, {1,2,3,4}, {5,6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 96.0, 1e-5);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 192.0, 1e-4);
}

TEST(BuiltinsGrowth, ColumnYReturnsColumn) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=GROWTH({6;12;24;48}, {1;2;3;4}, {5;6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 96.0, 1e-5);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 192.0, 1e-4);
}

TEST(BuiltinsGrowth, ConstFalseForcesUnitScale) {
  // y = 2^x: through-origin (logged: through-zero, b=1). Predict at
  // x=5 -> 32, x=6 -> 64.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=GROWTH({2,4,8,16}, {1,2,3,4}, {5,6}, FALSE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 32.0, kLooseEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 64.0, kLooseEps);
}

TEST(BuiltinsGrowth, OmittedNewXWithConstFalseUsesKnownX) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=GROWTH({2;4;8;16}, {1;2;3;4}, , FALSE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 4U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kLooseEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 4.0, kLooseEps);
  EXPECT_NEAR(v.as_array_cells()[2].as_number(), 8.0, kLooseEps);
  EXPECT_NEAR(v.as_array_cells()[3].as_number(), 16.0, kLooseEps);
}

TEST(BuiltinsGrowth, NonPositiveYReturnsNum) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=GROWTH({6,0,24,48}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsGrowth, ErrorInNewXPropagates) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=GROWTH({6,12,24,48}, {1,2,3,4}, {5,#N/A})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsGrowth, NewXFeatureCountMismatchReturnsRef) {
  // k=1 known_x, new_x has 2 features (2x2) -> #REF!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=GROWTH({6;12;24;48}, {1;2;3;4}, {5,6;7,8})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsGrowth, ZeroArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=GROWTH()", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
