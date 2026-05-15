// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for the lazy `HSTACK(array1, array2, ...)` and
// `VSTACK(array1, array2, ...)` dynamic-array spilling builtins. These
// concatenate inputs along columns / rows respectively; cells absent from
// a shorter / narrower input are filled with `#N/A` per Mac Excel.

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
#include "test_eval_helpers.h"
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

// ---------------------------------------------------------------------------
// HSTACK
// ---------------------------------------------------------------------------

TEST(BuiltinsHstack, TwoColumnVectorsSameLength) {
  // A1:A2 = {1, 2}; B1:B2 = {10, 20}. HSTACK -> 2x2 = {{1,10},{2,20}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(1, 0, Value::number(2));
  sheet.set_cell_value(0, 1, Value::number(10));
  sheet.set_cell_value(1, 1, Value::number(20));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=HSTACK(A1:A2, B1:B2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 20.0);
}

TEST(BuiltinsHstack, ThreeArraysVariadic) {
  // Three 2-row column vectors stacked horizontally -> 2x3.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(1, 0, Value::number(2));
  sheet.set_cell_value(0, 1, Value::number(3));
  sheet.set_cell_value(1, 1, Value::number(4));
  sheet.set_cell_value(0, 2, Value::number(5));
  sheet.set_cell_value(1, 2, Value::number(6));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=HSTACK(A1:A2, B1:B2, C1:C2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 6.0);
}

TEST(BuiltinsHstack, RowCountMismatchPadsWithNa) {
  // A1:A3 (3 rows) stacked with B1:B2 (2 rows). Output is 3 rows; the
  // missing cell at (row 2, col 1) becomes #N/A.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(1, 0, Value::number(2));
  sheet.set_cell_value(2, 0, Value::number(3));
  sheet.set_cell_value(0, 1, Value::number(10));
  sheet.set_cell_value(1, 1, Value::number(20));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=HSTACK(A1:A3, B1:B2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 3.0);
  ASSERT_TRUE(cells[5].is_error());
  EXPECT_EQ(cells[5].as_error(), ErrorCode::NA);
}

TEST(BuiltinsHstack, ScalarTreatedAsOneByOne) {
  // Scalar args become 1x1; stacking with a 2-row vector pads the scalar
  // column with #N/A in row 2.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(1, 0, Value::number(2));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=HSTACK(A1:A2, 99)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 99.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 2.0);
  ASSERT_TRUE(cells[3].is_error());
  EXPECT_EQ(cells[3].as_error(), ErrorCode::NA);
}

TEST(BuiltinsHstack, MixedShapesSumsColumns) {
  // 2x2 array A1:B2 stacked with 2x1 column C1:C2 -> 2x3.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(2));
  sheet.set_cell_value(1, 0, Value::number(3));
  sheet.set_cell_value(1, 1, Value::number(4));
  sheet.set_cell_value(0, 2, Value::number(5));
  sheet.set_cell_value(1, 2, Value::number(6));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=HSTACK(A1:B2, C1:C2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 6.0);
}

TEST(BuiltinsHstack, ZeroArgsRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=HSTACK()", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsHstack, ScalarErrorArgPropagates) {
  // NA() is a scalar function call yielding scalar Value::error. Per the
  // FILTER / SORT / SORTBY precedent, scalar errors at argument level
  // propagate as call-level errors (they do NOT become a 1x1 error cell
  // stacked into the output).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=HSTACK({1;2}, NA())", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsHstack, ErrorCellInArrayArgPreserved) {
  // An error embedded INSIDE an array argument is preserved verbatim in
  // the output (HSTACK does not short-circuit on it).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(1, 0, Value::error(ErrorCode::Ref));
  sheet.set_cell_value(0, 1, Value::number(10));
  sheet.set_cell_value(1, 1, Value::number(20));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=HSTACK(A1:A2, B1:B2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 10.0);
  ASSERT_TRUE(cells[2].is_error());
  EXPECT_EQ(cells[2].as_error(), ErrorCode::Ref);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 20.0);
}

TEST(BuiltinsHstack, ErrorCellInArrayLiteralArgPreserved) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=HSTACK({1;#N/A;3}, {4;5;6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 4.0);
  ASSERT_TRUE(cells[2].is_error());
  EXPECT_EQ(cells[2].as_error(), ErrorCode::NA);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 6.0);
}

// ---------------------------------------------------------------------------
// VSTACK
// ---------------------------------------------------------------------------

TEST(BuiltinsVstack, TwoRowVectorsSameWidth) {
  // A1:B1 = {1,2}; A2:B2 = {3,4}. VSTACK -> 2x2 = {{1,2},{3,4}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(2));
  sheet.set_cell_value(1, 0, Value::number(3));
  sheet.set_cell_value(1, 1, Value::number(4));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=VSTACK(A1:B1, A2:B2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 4.0);
}

TEST(BuiltinsVstack, ColumnCountMismatchPadsWithNa) {
  // 1x3 row stacked with 1x2 row. Output is 2x3 with #N/A at (1, 2).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(2));
  sheet.set_cell_value(0, 2, Value::number(3));
  sheet.set_cell_value(1, 0, Value::number(10));
  sheet.set_cell_value(1, 1, Value::number(20));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=VSTACK(A1:C1, A2:B2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 20.0);
  ASSERT_TRUE(cells[5].is_error());
  EXPECT_EQ(cells[5].as_error(), ErrorCode::NA);
}

TEST(BuiltinsVstack, ThreeArraysVariadic) {
  // Three 1x2 row vectors -> 3x2.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(2));
  sheet.set_cell_value(1, 0, Value::number(3));
  sheet.set_cell_value(1, 1, Value::number(4));
  sheet.set_cell_value(2, 0, Value::number(5));
  sheet.set_cell_value(2, 1, Value::number(6));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=VSTACK(A1:B1, A2:B2, A3:B3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  for (std::uint32_t i = 0; i < 6U; ++i) {
    EXPECT_DOUBLE_EQ(cells[i].as_number(), static_cast<double>(i + 1));
  }
}

TEST(BuiltinsVstack, ScalarTreatedAsOneByOne) {
  // Scalar 99 stacked under a 1x2 row -> 2x2 with #N/A in (1,1).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(2));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=VSTACK(A1:B1, 99)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 99.0);
  ASSERT_TRUE(cells[3].is_error());
  EXPECT_EQ(cells[3].as_error(), ErrorCode::NA);
}

TEST(BuiltinsVstack, ZeroArgsRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=VSTACK()", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsVstack, PreservesTextCells) {
  // Verify text payloads survive the stack copy.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("alpha"));
  sheet.set_cell_value(0, 1, Value::text("beta"));
  sheet.set_cell_value(1, 0, Value::text("gamma"));
  sheet.set_cell_value(1, 1, Value::text("delta"));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=VSTACK(A1:B1, A2:B2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_EQ(cells[0].as_text(), "alpha");
  EXPECT_EQ(cells[1].as_text(), "beta");
  EXPECT_EQ(cells[2].as_text(), "gamma");
  EXPECT_EQ(cells[3].as_text(), "delta");
}

TEST(BuiltinsVstack, ErrorCellInArrayLiteralArgPreserved) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=VSTACK({1,2,3}, {#VALUE!,5,6})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 3.0);
  ASSERT_TRUE(cells[3].is_error());
  EXPECT_EQ(cells[3].as_error(), ErrorCode::Value);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 6.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
