// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for the lazy `SORT(array, [sort_index], [sort_order], [by_col])`
// builtin. Shares its TU with FILTER / UNIQUE; uses the same array-context
// seam (`eval_node_as_array`) to keep range-shaped args 2D, then performs
// a stable permutation sort by the chosen key column / row using an
// Excel-canonical cell ordering: Number < Text < Bool < Error < Blank,
// with text compared ASCII case-insensitively. Blanks always sink to the
// end regardless of `sort_order`.

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
// Single-column ascending / descending
// ---------------------------------------------------------------------------

TEST(BuiltinsSort, AscendingNumbers) {
  // A1:A4 = {3, 1, 4, 2}. Default ascending -> {1, 2, 3, 4}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(3));
  sheet.set_cell_value(1, 0, Value::number(1));
  sheet.set_cell_value(2, 0, Value::number(4));
  sheet.set_cell_value(3, 0, Value::number(2));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:A4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 4.0);
}

TEST(BuiltinsSort, DescendingNumbers) {
  // sort_order = -1 reverses ascending. {3,1,4,2} -> {4,3,2,1}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(3));
  sheet.set_cell_value(1, 0, Value::number(1));
  sheet.set_cell_value(2, 0, Value::number(4));
  sheet.set_cell_value(3, 0, Value::number(2));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:A4, 1, -1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 1.0);
}

TEST(BuiltinsSort, TextCaseInsensitiveLex) {
  // {"banana","Apple","cherry","apple"} ascending -> Apple/apple stable
  // (input order preserved for equal keys), then banana, then cherry.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("banana"));
  sheet.set_cell_value(1, 0, Value::text("Apple"));
  sheet.set_cell_value(2, 0, Value::text("cherry"));
  sheet.set_cell_value(3, 0, Value::text("apple"));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:A4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  // Apple was at row 1, apple at row 3. Stable sort preserves this order
  // since both compare equal under case-insensitive compare.
  EXPECT_EQ(v.as_array_cells()[0].as_text(), "Apple");
  EXPECT_EQ(v.as_array_cells()[1].as_text(), "apple");
  EXPECT_EQ(v.as_array_cells()[2].as_text(), "banana");
  EXPECT_EQ(v.as_array_cells()[3].as_text(), "cherry");
}

// ---------------------------------------------------------------------------
// Multi-column with sort_index
// ---------------------------------------------------------------------------

TEST(BuiltinsSort, MultiColumnSortByIndex2) {
  // A1:B4 = {{"a",3},{"b",1},{"c",4},{"d",2}}. SORT by column 2 ascending:
  // expected row order = {"b",1},{"d",2},{"a",3},{"c",4}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("a"));
  sheet.set_cell_value(0, 1, Value::number(3));
  sheet.set_cell_value(1, 0, Value::text("b"));
  sheet.set_cell_value(1, 1, Value::number(1));
  sheet.set_cell_value(2, 0, Value::text("c"));
  sheet.set_cell_value(2, 1, Value::number(4));
  sheet.set_cell_value(3, 0, Value::text("d"));
  sheet.set_cell_value(3, 1, Value::number(2));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:B4, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_EQ(cells[0].as_text(), "b");
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 1.0);
  EXPECT_EQ(cells[2].as_text(), "d");
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 2.0);
  EXPECT_EQ(cells[4].as_text(), "a");
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 3.0);
  EXPECT_EQ(cells[6].as_text(), "c");
  EXPECT_DOUBLE_EQ(cells[7].as_number(), 4.0);
}

// ---------------------------------------------------------------------------
// by_col = TRUE -- sort columns by a row key
// ---------------------------------------------------------------------------

TEST(BuiltinsSort, ByColTrueSortsColumnsByRowKey) {
  // A1:D2 = {{4,1,3,2},{40,10,30,20}}. Sort columns by row 1 ascending:
  // permutation = (B,D,C,A) -> {{1,2,3,4},{10,20,30,40}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(4));
  sheet.set_cell_value(0, 1, Value::number(1));
  sheet.set_cell_value(0, 2, Value::number(3));
  sheet.set_cell_value(0, 3, Value::number(2));
  sheet.set_cell_value(1, 0, Value::number(40));
  sheet.set_cell_value(1, 1, Value::number(10));
  sheet.set_cell_value(1, 2, Value::number(30));
  sheet.set_cell_value(1, 3, Value::number(20));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:D2, 1, 1, TRUE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 20.0);
  EXPECT_DOUBLE_EQ(cells[6].as_number(), 30.0);
  EXPECT_DOUBLE_EQ(cells[7].as_number(), 40.0);
}

// ---------------------------------------------------------------------------
// Cross-kind ordering: Number < Text < Bool < Error < Blank
// ---------------------------------------------------------------------------

TEST(BuiltinsSort, CrossKindOrderingFollowsExcelRanks) {
  // Mixed-kind column: blank, error, TRUE, FALSE, "z", "a", 2, 1.
  // Ascending expected order: 1, 2, "a", "z", FALSE, TRUE, #N/A, blank.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::blank());
  sheet.set_cell_value(1, 0, Value::error(ErrorCode::NA));
  sheet.set_cell_value(2, 0, Value::boolean(true));
  sheet.set_cell_value(3, 0, Value::boolean(false));
  sheet.set_cell_value(4, 0, Value::text("z"));
  sheet.set_cell_value(5, 0, Value::text("a"));
  sheet.set_cell_value(6, 0, Value::number(2));
  sheet.set_cell_value(7, 0, Value::number(1));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:A8)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 8U);
  const Value* cells = v.as_array_cells();
  EXPECT_TRUE(cells[0].is_number());
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_TRUE(cells[1].is_number());
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_TRUE(cells[2].is_text());
  EXPECT_EQ(cells[2].as_text(), "a");
  EXPECT_TRUE(cells[3].is_text());
  EXPECT_EQ(cells[3].as_text(), "z");
  EXPECT_TRUE(cells[4].is_boolean());
  EXPECT_FALSE(cells[4].as_boolean());
  EXPECT_TRUE(cells[5].is_boolean());
  EXPECT_TRUE(cells[5].as_boolean());
  EXPECT_TRUE(cells[6].is_error());
  EXPECT_EQ(cells[6].as_error(), ErrorCode::NA);
  EXPECT_TRUE(cells[7].is_blank());
}

TEST(BuiltinsSort, BlanksStayLastInDescending) {
  // Descending {3, blank, 1, 2} -> {3, 2, 1, blank}. Blanks must sink to
  // the end regardless of `sort_order`.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(3));
  sheet.set_cell_value(1, 0, Value::blank());
  sheet.set_cell_value(2, 0, Value::number(1));
  sheet.set_cell_value(3, 0, Value::number(2));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:A4, 1, -1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 1.0);
  EXPECT_TRUE(v.as_array_cells()[3].is_blank());
}

// ---------------------------------------------------------------------------
// Errors in `array` are preserved in the output (not call-level)
// ---------------------------------------------------------------------------

TEST(BuiltinsSort, ErrorCellsParticipateInOrdering) {
  // {1, #N/A, 2} ascending -> {1, 2, #N/A}. Errors sit between Bool and
  // Blank in the ranking and are kept in the output verbatim.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(1, 0, Value::error(ErrorCode::NA));
  sheet.set_cell_value(2, 0, Value::number(2));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:A3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  EXPECT_TRUE(v.as_array_cells()[0].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_TRUE(v.as_array_cells()[1].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 2.0);
  EXPECT_TRUE(v.as_array_cells()[2].is_error());
  EXPECT_EQ(v.as_array_cells()[2].as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Argument validation
// ---------------------------------------------------------------------------

TEST(BuiltinsSort, SortIndexOutOfRangeReturnsValue) {
  // 1-column array: only sort_index = 1 is valid. Index 2 -> #VALUE!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(1, 0, Value::number(2));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:A2, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSort, SortIndexZeroReturnsValue) {
  // 1-based index, so 0 is invalid.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:A1, 0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSort, InvalidSortOrderReturnsValue) {
  // sort_order must be 1 or -1; 2 is rejected.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:A1, 1, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSort, ZeroArgsRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT()", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSort, FiveArgsRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:A1, 1, 1, FALSE, FALSE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSort, StableSortPreservesEqualKeyOrder) {
  // Two-column input where column 1 is the key and column 2 is a marker.
  // Two rows have key=5; their markers must appear in input order after
  // sort. {5,"a"},{3,"b"},{5,"c"},{1,"d"} sorted asc -> {1,"d"},{3,"b"},
  // {5,"a"},{5,"c"}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(5));
  sheet.set_cell_value(0, 1, Value::text("a"));
  sheet.set_cell_value(1, 0, Value::number(3));
  sheet.set_cell_value(1, 1, Value::text("b"));
  sheet.set_cell_value(2, 0, Value::number(5));
  sheet.set_cell_value(2, 1, Value::text("c"));
  sheet.set_cell_value(3, 0, Value::number(1));
  sheet.set_cell_value(3, 1, Value::text("d"));

  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=SORT(A1:B4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_EQ(cells[1].as_text(), "d");
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 3.0);
  EXPECT_EQ(cells[3].as_text(), "b");
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 5.0);
  EXPECT_EQ(cells[5].as_text(), "a");
  EXPECT_DOUBLE_EQ(cells[6].as_number(), 5.0);
  EXPECT_EQ(cells[7].as_text(), "c");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
