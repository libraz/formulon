// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for the lazy `CHOOSECOLS(array, col_num1, ...)` and
// `CHOOSEROWS(array, row_num1, ...)` dynamic-array spilling builtins.
// Indices are 1-based with negative-from-end support; out-of-range or `0`
// surfaces #VALUE!.

#include <cstdint>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "test_eval_helpers.h"
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

// Populates a 2x4 array at A1:D2:
//   row 0: 1 2 3 4
//   row 1: 5 6 7 8
void Populate2x4(Sheet& sheet) {
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(2));
  sheet.set_cell_value(0, 2, Value::number(3));
  sheet.set_cell_value(0, 3, Value::number(4));
  sheet.set_cell_value(1, 0, Value::number(5));
  sheet.set_cell_value(1, 1, Value::number(6));
  sheet.set_cell_value(1, 2, Value::number(7));
  sheet.set_cell_value(1, 3, Value::number(8));
}

// ---------------------------------------------------------------------------
// CHOOSECOLS
// ---------------------------------------------------------------------------

TEST(BuiltinsChoosecols, SinglePositiveIndex) {
  // Pick column 2 -> {{2}, {6}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSECOLS(A1:D2, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 6.0);
}

TEST(BuiltinsChoosecols, MultiplePositiveIndicesPreserveOrder) {
  // Pick columns 3, 1 in that order -> 2x2 = {{3,1},{7,5}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSECOLS(A1:D2, 3, 1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 7.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 5.0);
}

TEST(BuiltinsChoosecols, DuplicatedIndicesYieldDuplicatedColumns) {
  // CHOOSECOLS allows repeats -> column 2 twice.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSECOLS(A1:D2, 2, 2, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 6.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 6.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 6.0);
}

TEST(BuiltinsChoosecols, NegativeIndexCountsFromEnd) {
  // -1 is the last column (col 4); -4 is the first.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSECOLS(A1:D2, -1, -4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 8.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 5.0);
}

TEST(BuiltinsChoosecols, FractionalIndexTruncates) {
  // 2.9 truncates to 2; -1.7 truncates to -1.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSECOLS(A1:D2, 2.9, -1.7)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 4.0);
}

TEST(BuiltinsChoosecols, ZeroIndexReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSECOLS(A1:D2, 0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsChoosecols, OutOfRangePositiveReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSECOLS(A1:D2, 5)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsChoosecols, OutOfRangeNegativeReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSECOLS(A1:D2, -5)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsChoosecols, MissingIndexArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSECOLS(A1:D2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// CHOOSEROWS
// ---------------------------------------------------------------------------

TEST(BuiltinsChooserows, SinglePositiveIndex) {
  // Pick row 2 -> {{5,6,7,8}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSEROWS(A1:D2, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 6.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 7.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 8.0);
}

TEST(BuiltinsChooserows, MultipleIndicesIncludingDuplicateAndNegative) {
  // Pick rows 2, 1, -1 (-1 == row 2). Output is 3x4.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSEROWS(A1:D2, 2, 1, -1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  const Value* cells = v.as_array_cells();
  // Row 0 of output = source row 2.
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 8.0);
  // Row 1 of output = source row 1.
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[7].as_number(), 4.0);
  // Row 2 of output = source row 2 (via -1).
  EXPECT_DOUBLE_EQ(cells[8].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(cells[11].as_number(), 8.0);
}

TEST(BuiltinsChooserows, OutOfRangeReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSEROWS(A1:D2, 3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsChooserows, ZeroIndexReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x4(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSEROWS(A1:D2, 0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsChooserows, ScalarArrayTreatedAsOneByOne) {
  // A scalar treated as a 1x1 array; choosing row 1 returns it as 1x1.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=CHOOSEROWS(42, 1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 42.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
