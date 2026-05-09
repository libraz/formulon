// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for the lazy `TAKE(array, rows, [columns])` and
// `DROP(array, rows, [columns])` dynamic-array spilling builtins. Both
// pick a corner sub-array from `array`; sign of `rows` / `columns`
// selects edge (positive = leading, negative = trailing). Excessive
// take/drop clamps for TAKE and surfaces #CALC! for DROP. TAKE with
// explicit zero on either axis surfaces #CALC! (the result has zero
// cells along that axis).

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

// Populates a 3x4 array at A1:D3:
//   row 0: 1 2 3 4
//   row 1: 5 6 7 8
//   row 2: 9 10 11 12
void Populate3x4(Sheet& sheet) {
  for (std::uint32_t r = 0; r < 3U; ++r) {
    for (std::uint32_t c = 0; c < 4U; ++c) {
      sheet.set_cell_value(r, c, Value::number(static_cast<double>(r * 4U + c + 1U)));
    }
  }
}

// ---------------------------------------------------------------------------
// TAKE
// ---------------------------------------------------------------------------

TEST(BuiltinsTake, PositiveRowsLeadingEdge) {
  // Take first 2 rows -> 2x4.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TAKE(A1:D3, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[7].as_number(), 8.0);
}

TEST(BuiltinsTake, NegativeRowsTrailingEdge) {
  // Take last 2 rows -> 2x4 (rows 1..2).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TAKE(A1:D3, -2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[7].as_number(), 12.0);
}

TEST(BuiltinsTake, PositiveRowsAndColumns) {
  // Take 2 rows from top and 3 columns from left -> 2x3.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TAKE(A1:D3, 2, 3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 6.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 7.0);
}

TEST(BuiltinsTake, NegativeRowsAndColumnsTrailingCorner) {
  // Last 2 rows and last 2 columns -> 2x2 = {{7,8},{11,12}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TAKE(A1:D3, -2, -2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 7.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 8.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 12.0);
}

TEST(BuiltinsTake, RowsExceedingClampsToFullAxis) {
  // 100 rows requested from a 3-row source -> all 3 rows returned.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TAKE(A1:D3, 100)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 4U);
}

TEST(BuiltinsTake, FractionalCountTruncates) {
  // 2.7 truncates to 2.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TAKE(A1:D3, 2.7)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
}

TEST(BuiltinsTake, ZeroRowsReturnsCalc) {
  // TAKE with zero rows leaves nothing -> #CALC!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TAKE(A1:D3, 0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

TEST(BuiltinsTake, ZeroColumnsReturnsCalc) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TAKE(A1:D3, 1, 0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

TEST(BuiltinsTake, OneArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TAKE(A1:D3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// DROP
// ---------------------------------------------------------------------------

TEST(BuiltinsDrop, PositiveRowsDropsFromTop) {
  // Drop first row -> 2x4 (rows 1..2).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=DROP(A1:D3, 1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[7].as_number(), 12.0);
}

TEST(BuiltinsDrop, NegativeRowsDropsFromBottom) {
  // Drop last row -> 2x4 (rows 0..1).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=DROP(A1:D3, -1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[7].as_number(), 8.0);
}

TEST(BuiltinsDrop, ZeroDoesNothing) {
  // Drop 0 rows -> full 3x4 array unchanged.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=DROP(A1:D3, 0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 4U);
}

TEST(BuiltinsDrop, BothAxesPositive) {
  // Drop first row and first column -> 2x3.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=DROP(A1:D3, 1, 1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 6.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 7.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 8.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 11.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 12.0);
}

TEST(BuiltinsDrop, OverDropRowsReturnsCalc) {
  // Drop all rows -> #CALC!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=DROP(A1:D3, 3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

TEST(BuiltinsDrop, OverDropColumnsReturnsCalc) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=DROP(A1:D3, 0, 4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

TEST(BuiltinsDrop, NegativeColumnsDropsFromRight) {
  // Drop the last 2 columns -> 3x2.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=DROP(A1:D3, 0, -2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 6.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 9.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 10.0);
}

TEST(BuiltinsDrop, OneArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate3x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=DROP(A1:D3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
