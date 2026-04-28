// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tests for the LINEST lazy builtin (`linest_lazy.{h,cpp}`). Coverage:
// single-variable / multivariable regression with and without an
// intercept, the 5x(k+1) stats matrix, error propagation, shape rules
// (column-y vs row-y inputs), and degenerate-system rejection.
//
// Reference values are computed analytically in the test comments so the
// expected numbers can be re-derived without running Excel.

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

constexpr double kEps = 1e-9;

// ---------------------------------------------------------------------------
// stats=FALSE — coefficients only
// ---------------------------------------------------------------------------

TEST(BuiltinsLinest, PerfectFitSingleVariableViaArrayLiterals) {
  // y = 2x: known_y = {2,4,6,8}, known_x = {1,2,3,4}
  // Expected output: 1x2 row [slope, intercept] = [2, 0].
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({2,4,6,8}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* c = v.as_array_cells();
  EXPECT_NEAR(c[0].as_number(), 2.0, kEps);
  EXPECT_NEAR(c[1].as_number(), 0.0, kEps);
}

TEST(BuiltinsLinest, NonZeroInterceptSingleVariable) {
  // y = 2x + 1: {3,5,7,9} / {1,2,3,4}. Expect [2, 1].
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({3,5,7,9}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 1.0, kEps);
}

TEST(BuiltinsLinest, DefaultKnownXUsesOneTwoThreeSequence) {
  // No second arg -> known_x defaults to {1, 2, ..., m}. With y = {2,4,6,8}
  // this is the same as the perfect-fit case above: [2, 0].
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({2,4,6,8})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_cols(), 2U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 0.0, kEps);
}

TEST(BuiltinsLinest, NoisyDataSingleVariableMatchesAnalyticSlope) {
  // y = {1, 2, 4, 7}, x = {1, 2, 3, 4}.
  //   sum_xy / sum_xx = 10 / 5 = 2 -> slope = 2.
  //   intercept = mean_y - 2 * mean_x = 3.5 - 5 = -1.5.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({1,2,4,7}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), -1.5, kEps);
}

TEST(BuiltinsLinest, ConstFalseForcesInterceptToZero) {
  // const=FALSE and y = {2,4,6,8} / x = {1,2,3,4}: regression through
  // origin gives slope = sum(xy)/sum(x^2) = 60 / 30 = 2. The output is
  // still 1x2 with the trailing intercept slot reported as 0.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({2,4,6,8}, {1,2,3,4}, FALSE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_cols(), 2U);
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kEps);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 0.0);
}

TEST(BuiltinsLinest, MultiVariableRecoversCoefficientsInReverseOrder) {
  // y = 1 + 2*x1 + 3*x2 with x1 = {1..5}, x2 = {2,3,5,7,11}.
  // Expected output 1x3 in right-to-left order: [b2, b1, b0] = [3, 2, 1].
  // known_y is column-shaped (5x1) and known_x is 5x2 (each row = one obs).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // y at A1:A5
  sheet.set_cell_value(0, 0, Value::number(9));
  sheet.set_cell_value(1, 0, Value::number(14));
  sheet.set_cell_value(2, 0, Value::number(22));
  sheet.set_cell_value(3, 0, Value::number(30));
  sheet.set_cell_value(4, 0, Value::number(44));
  // x1 at B1:B5
  sheet.set_cell_value(0, 1, Value::number(1));
  sheet.set_cell_value(1, 1, Value::number(2));
  sheet.set_cell_value(2, 1, Value::number(3));
  sheet.set_cell_value(3, 1, Value::number(4));
  sheet.set_cell_value(4, 1, Value::number(5));
  // x2 at C1:C5
  sheet.set_cell_value(0, 2, Value::number(2));
  sheet.set_cell_value(1, 2, Value::number(3));
  sheet.set_cell_value(2, 2, Value::number(5));
  sheet.set_cell_value(3, 2, Value::number(7));
  sheet.set_cell_value(4, 2, Value::number(11));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST(A1:A5, B1:C5)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* c = v.as_array_cells();
  EXPECT_NEAR(c[0].as_number(), 3.0, 1e-7);  // b2
  EXPECT_NEAR(c[1].as_number(), 2.0, 1e-7);  // b1
  EXPECT_NEAR(c[2].as_number(), 1.0, 1e-7);  // b0
}

TEST(BuiltinsLinest, RowOrientedKnownYAccepted) {
  // y is 1x4 (row); x is also 1x4. Expect the same answer as the
  // column-y perfect-fit case.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(2));
  sheet.set_cell_value(0, 1, Value::number(4));
  sheet.set_cell_value(0, 2, Value::number(6));
  sheet.set_cell_value(0, 3, Value::number(8));
  sheet.set_cell_value(1, 0, Value::number(1));
  sheet.set_cell_value(1, 1, Value::number(2));
  sheet.set_cell_value(1, 2, Value::number(3));
  sheet.set_cell_value(1, 3, Value::number(4));
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST(A1:D1, A2:D2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 0.0, kEps);
}

TEST(BuiltinsLinest, ColumnYRowXMixedOrientationAccepted) {
  // y is 4x1 (column), x is 1x4 (row). Mac Excel accepts this:
  // x is treated as a single variable transposed to match y.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({2;4;6;8}, {1,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 0.0, kEps);
}

// ---------------------------------------------------------------------------
// stats=TRUE — 5x(k+1) statistics matrix
// ---------------------------------------------------------------------------

TEST(BuiltinsLinest, StatsBlockMatchesAnalyticReference) {
  // y = {1, 2, 4, 7}, x = {1, 2, 3, 4}.
  // slope = 2, intercept = -1.5, ss_resid = 1.0, ss_total = 21.0
  //   r2 = 1 - 1/21 = 20/21
  //   df_resid = 2; sigma2 = 0.5; se_y = sqrt(0.5)
  //   X^T X = [[30,10],[10,4]]; (X^T X)^-1 = [[0.2, -0.5],[-0.5, 1.5]]
  //   se(slope)     = sqrt(0.5 * 0.2)  = sqrt(0.1)
  //   se(intercept) = sqrt(0.5 * 1.5) = sqrt(0.75)
  //   F = (ss_reg / k) / (ss_resid / df_resid) = (20/1)/(1/2) = 40
  //   ss_reg = 20, ss_resid = 1
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({1,2,4,7}, {1,2,3,4}, TRUE, TRUE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 5U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* c = v.as_array_cells();
  // Row 1: [slope, intercept]
  EXPECT_NEAR(c[0].as_number(), 2.0, 1e-9);
  EXPECT_NEAR(c[1].as_number(), -1.5, 1e-9);
  // Row 2: [se(slope), se(intercept)]
  EXPECT_NEAR(c[2].as_number(), std::sqrt(0.1), 1e-9);
  EXPECT_NEAR(c[3].as_number(), std::sqrt(0.75), 1e-9);
  // Row 3: [r^2, se_y]
  EXPECT_NEAR(c[4].as_number(), 20.0 / 21.0, 1e-9);
  EXPECT_NEAR(c[5].as_number(), std::sqrt(0.5), 1e-9);
  // Row 4: [F, df]
  EXPECT_NEAR(c[6].as_number(), 40.0, 1e-9);
  EXPECT_NEAR(c[7].as_number(), 2.0, 1e-9);
  // Row 5: [ss_reg, ss_resid]
  EXPECT_NEAR(c[8].as_number(), 20.0, 1e-9);
  EXPECT_NEAR(c[9].as_number(), 1.0, 1e-9);
}

TEST(BuiltinsLinest, StatsBlockTrailingSlotsAreNAForMultiVariable) {
  // 5x3 output for k=2: rows 3-5 should have #N/A in column index 2.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // y = 9,14,22,30,44 at A1:A5; x at B1:C5 (perfect fit, see earlier test).
  sheet.set_cell_value(0, 0, Value::number(9));
  sheet.set_cell_value(1, 0, Value::number(14));
  sheet.set_cell_value(2, 0, Value::number(22));
  sheet.set_cell_value(3, 0, Value::number(30));
  sheet.set_cell_value(4, 0, Value::number(44));
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
  const Value v = EvalUnder("=LINEST(A1:A5, B1:C5, TRUE, TRUE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 5U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* c = v.as_array_cells();
  // Row 3 column 2, row 4 column 2, row 5 column 2 -> #N/A.
  EXPECT_TRUE(c[2 * 3 + 2].is_error());
  EXPECT_EQ(c[2 * 3 + 2].as_error(), ErrorCode::NA);
  EXPECT_TRUE(c[3 * 3 + 2].is_error());
  EXPECT_EQ(c[3 * 3 + 2].as_error(), ErrorCode::NA);
  EXPECT_TRUE(c[4 * 3 + 2].is_error());
  EXPECT_EQ(c[4 * 3 + 2].as_error(), ErrorCode::NA);
}

TEST(BuiltinsLinest, PerfectFitYieldsRSquaredOne) {
  // y = 2x exactly. r^2 = 1, ss_resid = 0, F is undefined (0/0) so
  // surfaces as #N/A. ss_resid is exactly zero.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({2,4,6,8}, {1,2,3,4}, TRUE, TRUE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  const Value* c = v.as_array_cells();
  // Row 3 col 0: r^2 = 1.
  EXPECT_NEAR(c[4].as_number(), 1.0, 1e-9);
  // Row 5 col 1: ss_resid = 0 (might be tiny FP noise; tolerate).
  EXPECT_NEAR(c[9].as_number(), 0.0, 1e-9);
}

TEST(BuiltinsLinest, StatsConstFalseInterceptSeIsNA) {
  // const=false -> intercept slot in row 1 is exactly 0; intercept SE
  // in row 2 is #N/A (no intercept estimated).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({2,4,6,8}, {1,2,3,4}, FALSE, TRUE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 5U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[1].as_number(), 0.0);
  EXPECT_TRUE(c[3].is_error());
  EXPECT_EQ(c[3].as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Error propagation and shape checking
// ---------------------------------------------------------------------------

TEST(BuiltinsLinest, ZeroArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST()", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsLinest, FiveArgsRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({1,2,3}, {1,2,3}, TRUE, TRUE, 99)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsLinest, NonNumericYCellReturnsValue) {
  // Text cell in y -> #VALUE! (matrix-strict coercion).
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
  const Value v = EvalUnder("=LINEST(A1:A3, B1:B3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsLinest, ErrorCellInYPropagatesVerbatim) {
  // #N/A in y -> the impl surfaces #N/A unchanged.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({1,#N/A,3}, {1,2,3})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsLinest, ShapeMismatchReturnsRef) {
  // y is 3x1, x is 4x1 -> #REF!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({1;2;3}, {1;2;3;4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsLinest, CollinearXReturnsPartialFit) {
  // All x equal to the intercept column (`{1,1,1,1}`) -> X^T X is exactly
  // singular. The rank-aware Gauss-Jordan kernel drops the redundant
  // predictor column and lets the intercept absorb mean(y) = 2.5,
  // matching Mac Excel 365: output is the 1x2 spill `[0, 2.5]` (slope
  // dropped, intercept = mean of y).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({1,2,3,4}, {1,1,1,1})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* c = v.as_array_cells();
  EXPECT_NEAR(c[0].as_number(), 0.0, kEps);
  EXPECT_NEAR(c[1].as_number(), 2.5, kEps);
}

TEST(BuiltinsLinest, UnderdeterminedSystemReturnsNum) {
  // m=2, k=2 + intercept => p=3 > m. Under-determined -> #NUM!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({1;2}, {1,2;3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsLinest, BooleanCellsCoerceToNumeric) {
  // Booleans coerce to 1/0. With y={2,4,6,8} and x={TRUE,2,3,4}={1,2,3,4}
  // we get the same perfect-fit answer as before.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=LINEST({2,4,6,8}, {TRUE,2,3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_NEAR(v.as_array_cells()[0].as_number(), 2.0, kEps);
  EXPECT_NEAR(v.as_array_cells()[1].as_number(), 0.0, kEps);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
