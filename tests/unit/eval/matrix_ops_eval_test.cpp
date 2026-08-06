//
// Tests for the linear-algebra lazy builtins (`MMULT`, `MDETERM`,
// `MINVERSE`). Each function lives in `matrix_ops_lazy.{h,cpp}`; the
// dispatch wiring is in `tree_walker.cpp`.

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

// Populates a 2x3 matrix A at A1:C2:
//   1 2 3
//   4 5 6
void PopulateA_2x3(Sheet& sheet) {
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(2));
  sheet.set_cell_value(0, 2, Value::number(3));
  sheet.set_cell_value(1, 0, Value::number(4));
  sheet.set_cell_value(1, 1, Value::number(5));
  sheet.set_cell_value(1, 2, Value::number(6));
}

// Populates a 3x2 matrix B at E1:F3:
//   7  8
//   9 10
//  11 12
void PopulateB_3x2(Sheet& sheet) {
  sheet.set_cell_value(0, 4, Value::number(7));
  sheet.set_cell_value(0, 5, Value::number(8));
  sheet.set_cell_value(1, 4, Value::number(9));
  sheet.set_cell_value(1, 5, Value::number(10));
  sheet.set_cell_value(2, 4, Value::number(11));
  sheet.set_cell_value(2, 5, Value::number(12));
}

// ---------------------------------------------------------------------------
// MMULT
// ---------------------------------------------------------------------------

TEST(BuiltinsMmult, BasicProductFromSheetRanges) {
  // (2x3) * (3x2) = (2x2). [[58,64],[139,154]] is the canonical worked
  // example for these matrices and is what Excel reports.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  PopulateA_2x3(sheet);
  PopulateB_3x2(sheet);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MMULT(A1:C2, E1:F3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 58.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 64.0);
  EXPECT_DOUBLE_EQ(c[2].as_number(), 139.0);
  EXPECT_DOUBLE_EQ(c[3].as_number(), 154.0);
}

TEST(BuiltinsMmult, ArrayLiteralArguments) {
  // Same product computed from inline `{...}` literals.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MMULT({1,2,3;4,5,6}, {7,8;9,10;11,12})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 58.0);
  EXPECT_DOUBLE_EQ(c[3].as_number(), 154.0);
}

TEST(BuiltinsMmult, IdentityIsLeftAndRightNeutral) {
  // I * M = M and M * I = M for any 2x2 M.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MMULT({1,0;0,1}, {2,3;4,5})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(c[2].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(c[3].as_number(), 5.0);
}

TEST(BuiltinsMmult, InnerDimensionMismatchReturnsValue) {
  // (2x3) * (2x2) — cols(A) != rows(B), so #VALUE!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MMULT({1,2,3;4,5,6}, {1,2;3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMmult, NonNumericCellReturnsValue) {
  // Text cell anywhere in either input -> #VALUE!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MMULT({1,\"x\";3,4}, {1,2;3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMmult, ErrorCellPropagates) {
  // #N/A inside an input matrix propagates verbatim.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MMULT({1,#N/A;3,4}, {1,2;3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMmult, ScalarArgumentTreatedAsOneByOne) {
  // 1x1 dotted with 1x1 yields a 1x1 output (= product). Excel accepts
  // this.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MMULT(3, 4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 12.0);
}

TEST(BuiltinsMmult, OneArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MMULT({1,2;3,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMmult, BooleanCellsCoerceToNumeric) {
  // TRUE/FALSE coerce to 1/0 just like the rest of the matrix family.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MMULT({TRUE,FALSE;FALSE,TRUE}, {2;3})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// MDETERM
// ---------------------------------------------------------------------------

TEST(BuiltinsMdeterm, TwoByTwoDeterminant) {
  // det({{a,b},{c,d}}) = a*d - b*c. {{3,8},{4,6}} -> 18 - 32 = -14.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MDETERM({3,8;4,6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -14.0);
}

TEST(BuiltinsMdeterm, ThreeByThreeDeterminant) {
  // det({{6,1,1},{4,-2,5},{2,8,7}}) = -306 (Khan-Academy worked example).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MDETERM({6,1,1;4,-2,5;2,8,7})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -306.0, 1e-9);
}

TEST(BuiltinsMdeterm, IdentityHasDeterminantOne) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MDETERM({1,0,0;0,1,0;0,0,1})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMdeterm, SingularMatrixYieldsZero) {
  // Two identical rows -> singular matrix.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MDETERM({1,2,3;1,2,3;4,5,6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMdeterm, NonSquareReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MDETERM({1,2,3;4,5,6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMdeterm, NonNumericCellReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MDETERM({1,2;\"x\",4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMdeterm, ScalarArgumentReturnsItself) {
  // 1x1 matrix's determinant is the cell value.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MDETERM(7)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsMdeterm, ZeroArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MDETERM()", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// MINVERSE
// ---------------------------------------------------------------------------

TEST(BuiltinsMinverse, TwoByTwoInverse) {
  // inv({{4,7},{2,6}}) = {{0.6,-0.7},{-0.2,0.4}} (1/det = 0.1; det=10).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MINVERSE({4,7;2,6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* c = v.as_array_cells();
  EXPECT_NEAR(c[0].as_number(), 0.6, 1e-12);
  EXPECT_NEAR(c[1].as_number(), -0.7, 1e-12);
  EXPECT_NEAR(c[2].as_number(), -0.2, 1e-12);
  EXPECT_NEAR(c[3].as_number(), 0.4, 1e-12);
}

TEST(BuiltinsMinverse, IdentityIsItsOwnInverse) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MINVERSE({1,0;0,1})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 0.0);
  EXPECT_DOUBLE_EQ(c[2].as_number(), 0.0);
  EXPECT_DOUBLE_EQ(c[3].as_number(), 1.0);
}

TEST(BuiltinsMinverse, ProductWithOriginalIsIdentity) {
  // M * inv(M) = I (within floating-point tolerance). Embeds MINVERSE in
  // MMULT so the result lifecycle (arena allocation + dispatch) is
  // exercised end-to-end.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MMULT({3,8;4,6}, MINVERSE({3,8;4,6}))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* c = v.as_array_cells();
  EXPECT_NEAR(c[0].as_number(), 1.0, 1e-10);
  EXPECT_NEAR(c[1].as_number(), 0.0, 1e-10);
  EXPECT_NEAR(c[2].as_number(), 0.0, 1e-10);
  EXPECT_NEAR(c[3].as_number(), 1.0, 1e-10);
}

TEST(BuiltinsMinverse, ThreeByThreeInverse) {
  // inv({{1,2,3},{0,1,4},{5,6,0}}) computed analytically:
  // Det = 1*(1*0 - 4*6) - 2*(0*0 - 4*5) + 3*(0*6 - 1*5)
  //     = 1*(-24) - 2*(-20) + 3*(-5)
  //     = -24 + 40 - 15 = 1.
  // Inverse = adjugate / det. Hand-derived adjugate (transposed cofactors):
  //   {{-24,18,5},{20,-15,-4},{-5,4,1}}.
  // So inv = {{-24,18,5},{20,-15,-4},{-5,4,1}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MINVERSE({1,2,3;0,1,4;5,6,0})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* c = v.as_array_cells();
  EXPECT_NEAR(c[0].as_number(), -24.0, 1e-10);
  EXPECT_NEAR(c[1].as_number(), 18.0, 1e-10);
  EXPECT_NEAR(c[2].as_number(), 5.0, 1e-10);
  EXPECT_NEAR(c[3].as_number(), 20.0, 1e-10);
  EXPECT_NEAR(c[4].as_number(), -15.0, 1e-10);
  EXPECT_NEAR(c[5].as_number(), -4.0, 1e-10);
  EXPECT_NEAR(c[6].as_number(), -5.0, 1e-10);
  EXPECT_NEAR(c[7].as_number(), 4.0, 1e-10);
  EXPECT_NEAR(c[8].as_number(), 1.0, 1e-10);
}

TEST(BuiltinsMinverse, SingularMatrixReturnsNum) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MINVERSE({1,2;2,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMinverse, NonSquareReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MINVERSE({1,2,3;4,5,6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMinverse, NonNumericCellReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MINVERSE({1,2;\"x\",4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMinverse, ErrorCellPropagates) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MINVERSE({1,2;#N/A,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsMinverse, ScalarArgumentReturnsReciprocalAsOneByOne) {
  // 1x1 matrix's inverse is the reciprocal cell.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=MINVERSE(4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 0.25);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
