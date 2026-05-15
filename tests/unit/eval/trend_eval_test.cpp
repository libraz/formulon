// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for the TREND lazy builtin (`linest_lazy.{h,cpp}`). TREND fits
// the same multivariate least-squares model as LINEST and evaluates the
// fit at `new_x`. Coverage:
//   - default new_x (returns fitted values at the training observations);
//   - explicit new_x for single- and multi-variable fits;
//   - column / row orientation of known_y / new_x;
//   - const=FALSE forcing the regression through the origin;
//   - error and shape diagnostics shared with LINEST.
//
// Reference values are computed analytically in the test comments.

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

constexpr double kEps = 1e-9;

// ---------------------------------------------------------------------------
// Default new_x — fitted values at training observations
// ---------------------------------------------------------------------------

TEST(BuiltinsTrend, DefaultNewXReturnsFittedValuesPerfectLine) {
  // Perfect fit y = 2x: TREND should reproduce y exactly.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({2,4,6,8}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  const Value* c = v.as_array_cells();
  EXPECT_NEAR(c[0].as_number(), 2.0, kEps);
  EXPECT_NEAR(c[1].as_number(), 4.0, kEps);
  EXPECT_NEAR(c[2].as_number(), 6.0, kEps);
  EXPECT_NEAR(c[3].as_number(), 8.0, kEps);
}

TEST(BuiltinsTrend, DefaultNewXNoisyDataReturnsFittedYHat) {
  // y = {1, 2, 4, 7}, x = {1, 2, 3, 4}.
  // Slope = 2, intercept = -1.5, so y_hat = {0.5, 2.5, 4.5, 6.5}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({1,2,4,7}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_cols(), 4U);
  const Value* c = v.as_array_cells();
  EXPECT_NEAR(c[0].as_number(), 0.5, kEps);
  EXPECT_NEAR(c[1].as_number(), 2.5, kEps);
  EXPECT_NEAR(c[2].as_number(), 4.5, kEps);
  EXPECT_NEAR(c[3].as_number(), 6.5, kEps);
}

TEST(BuiltinsTrend, DefaultKnownXAndDefaultNewX) {
  // Both known_x and new_x defaulted: known_x = {1,2,3,4}, fitted at the
  // same. With y = {2,4,6,8}: slope=2, int=0, y_hat = y exactly.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({2,4,6,8})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_cols(), 4U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[3].as_number(), 8.0, kEps);
}

// ---------------------------------------------------------------------------
// Explicit new_x
// ---------------------------------------------------------------------------

TEST(BuiltinsTrend, ExplicitNewXSingleVariable) {
  // Same {1,2,4,7} / {1,2,3,4} fit as above; predict at {5, 6}.
  // y_hat(5) = 2*5 - 1.5 = 8.5; y_hat(6) = 2*6 - 1.5 = 10.5.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({1,2,4,7}, {1,2,3,4}, {5,6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  // Output orientation matches known_y (row of length 4 -> row of length 2).
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 8.5, kEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 10.5, kEps);
}

TEST(BuiltinsTrend, ExplicitNewXMultiVariable) {
  // y = 1 + 2*x1 + 3*x2 with x1={1..5}, x2={2,3,5,7,11}. Predict at
  // [(6, 13), (7, 17)]: y_hat = 1 + 12 + 39 = 52; 1 + 14 + 51 = 66.
  // y is column 5x1, known_x is 5x2, new_x is 2x2 (each row = obs).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // y at A1:A5
  sheet.set_cell_value(0, 0, Value::number(9));
  sheet.set_cell_value(1, 0, Value::number(14));
  sheet.set_cell_value(2, 0, Value::number(22));
  sheet.set_cell_value(3, 0, Value::number(30));
  sheet.set_cell_value(4, 0, Value::number(44));
  // x at B1:C5
  sheet.set_cell_value(0, 1, Value::number(1));
  sheet.set_cell_value(1, 1, Value::number(2));
  sheet.set_cell_value(2, 1, Value::number(3));
  sheet.set_cell_value(3, 1, Value::number(4));
  sheet.set_cell_value(4, 1, Value::number(5));
  sheet.set_cell_value(0, 2, Value::number(2));
  sheet.set_cell_value(1, 2, Value::number(3));
  sheet.set_cell_value(2, 2, Value::number(5));
  sheet.set_cell_value(3, 2, Value::number(7));
  sheet.set_cell_value(4, 2, Value::number(11));
  // new_x at E1:F2
  sheet.set_cell_value(0, 4, Value::number(6));
  sheet.set_cell_value(0, 5, Value::number(13));
  sheet.set_cell_value(1, 4, Value::number(7));
  sheet.set_cell_value(1, 5, Value::number(17));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND(A1:A5, B1:C5, E1:F2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  const Value* c = v.as_array_cells();
  EXPECT_NEAR(c[0].as_number(), 52.0, 1e-7);
  EXPECT_NEAR(c[1].as_number(), 66.0, 1e-7);
}

TEST(BuiltinsTrend, ConstFalseForcesThroughOrigin) {
  // const=FALSE, y = {2,4,6,8}, x = {1,2,3,4}, predict at {0, 5}.
  // Through-origin slope = sum(xy)/sum(x^2) = 60/30 = 2. y_hat(0)=0, y_hat(5)=10.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({2,4,6,8}, {1,2,3,4}, {0,5}, FALSE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 0.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 10.0, kEps);
}

TEST(BuiltinsTrend, OmittedNewXWithConstFalseUsesKnownX) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({2;4;6}, {1;2;3}, , FALSE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  ASSERT_EQ(v.as_array_rows(), 3U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 4.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[2].as_number(), 6.0, kEps);
}

// ---------------------------------------------------------------------------
// Orientation — column / row inputs
// ---------------------------------------------------------------------------

TEST(BuiltinsTrend, ColumnYColumnXColumnNewXReturnsColumn) {
  // y is 4x1, known_x is 4x1, new_x is 2x1 -> output 2x1.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({2;4;6;8}, {1;2;3;4}, {5;6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 10.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 12.0, kEps);
}

TEST(BuiltinsTrend, RowYRowXRowNewXReturnsRow) {
  // y is 1x4, known_x is 1x4, new_x is 1x2 -> output 1x2.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({2,4,6,8}, {1,2,3,4}, {5,6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 10.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 12.0, kEps);
}

// ---------------------------------------------------------------------------
// Error / shape semantics
// ---------------------------------------------------------------------------

TEST(BuiltinsTrend, ZeroArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND()", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsTrend, FiveArgsRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({1,2,3}, {1,2,3}, {4,5}, TRUE, 99)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsTrend, NewXFeatureCountMismatchReturnsRef) {
  // known_x has k=2 columns; new_x has 3 columns -> #REF!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // y at A1:A5
  sheet.set_cell_value(0, 0, Value::number(9));
  sheet.set_cell_value(1, 0, Value::number(14));
  sheet.set_cell_value(2, 0, Value::number(22));
  sheet.set_cell_value(3, 0, Value::number(30));
  sheet.set_cell_value(4, 0, Value::number(44));
  // x at B1:C5 (2 features)
  sheet.set_cell_value(0, 1, Value::number(1));
  sheet.set_cell_value(1, 1, Value::number(2));
  sheet.set_cell_value(2, 1, Value::number(3));
  sheet.set_cell_value(3, 1, Value::number(4));
  sheet.set_cell_value(4, 1, Value::number(5));
  sheet.set_cell_value(0, 2, Value::number(2));
  sheet.set_cell_value(1, 2, Value::number(3));
  sheet.set_cell_value(2, 2, Value::number(5));
  sheet.set_cell_value(3, 2, Value::number(7));
  sheet.set_cell_value(4, 2, Value::number(11));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND(A1:A5, B1:C5, {1,2,3})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsTrend, ErrorInYPropagates) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({1,#N/A,3}, {1,2,3})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsTrend, ErrorInNewXPropagates) {
  // Error in new_x cells -> propagate that error.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({1,2,3}, {1,2,3}, {#N/A,5})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsTrend, NonNumericYReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(1, 0, Value::text("x"));
  sheet.set_cell_value(2, 0, Value::number(3));
  sheet.set_cell_value(0, 1, Value::number(1));
  sheet.set_cell_value(1, 1, Value::number(2));
  sheet.set_cell_value(2, 1, Value::number(3));
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND(A1:A3, B1:B3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsTrend, CollinearXReturnsMeanYFit) {
  // x is constant (collinear with intercept) -> X^T X singular. The
  // rank-aware Gauss-Jordan kernel drops the predictor and lets the
  // intercept absorb mean(y) = 2.5. TREND's fitted values therefore
  // collapse to mean(y) at every known_x, matching Mac Excel 365.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TREND({1,2,3,4}, {1,1,1,1})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  const Value* c = v.as_array_cells();
  for (std::uint32_t i = 0; i < 4U; ++i) {
    EXPECT_NEAR(c[i].as_number(), 2.5, 1e-9) << "i=" << i;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
