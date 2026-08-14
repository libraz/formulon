//
// Tests for the lazy array-reshape dynamic-array spilling builtins:
// `EXPAND`, `TOCOL`, `TOROW`, `WRAPCOLS`, `WRAPROWS`. `EXPAND` lives in
// `dynamic_array/reshape.cpp` and the four flatten / wrap functions live
// in `dynamic_array/layout.cpp`; they share helpers
// (`collect_tocol_torow_cells`, `resolve_wrap_args`,
// `dynamic_array::materialise_vector`, `dynamic_array::allocate_array_value`)
// across the `dynamic_array/common.{h,cpp}` seam.

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

// Populates a 2x3 array at A1:C2:
//   row 0: 1 2 3
//   row 1: 4 5 6
void Populate2x3(Sheet& sheet) {
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::number(2));
  sheet.set_cell_value(0, 2, Value::number(3));
  sheet.set_cell_value(1, 0, Value::number(4));
  sheet.set_cell_value(1, 1, Value::number(5));
  sheet.set_cell_value(1, 2, Value::number(6));
}

// ---------------------------------------------------------------------------
// EXPAND
// ---------------------------------------------------------------------------

TEST(BuiltinsExpand, GrowsRowsAndColumnsWithDefaultPadNa) {
  // Expand 2x3 source to 4x5; new cells are #N/A.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=EXPAND(A1:C2, 4, 5)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  ASSERT_EQ(v.as_array_cols(), 5U);
  const Value* cells = v.as_array_cells();
  // Original cells stay anchored at top-left.
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[7].as_number(), 6.0);
  // New cells (col 3, 4 of row 0; rows 2, 3 entirely) are #N/A.
  ASSERT_TRUE(cells[3].is_error());
  EXPECT_EQ(cells[3].as_error(), ErrorCode::NA);
  ASSERT_TRUE(cells[10].is_error());
  EXPECT_EQ(cells[10].as_error(), ErrorCode::NA);
}

TEST(BuiltinsExpand, ColumnsArgOmittedKeepsExistingColumnCount) {
  // Expand to 5 rows but keep the existing 3 columns.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=EXPAND(A1:C2, 5)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 5U);
  ASSERT_EQ(v.as_array_cols(), 3U);
}

TEST(BuiltinsExpand, RejectsOverflowingDenseResultBeforeAllocation) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1));
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, wb.sheet(0), state);
  Arena parse_arena;
  Arena eval_arena;
  // Both dimensions fit Excel's grid. Their product would otherwise reserve
  // tens of gigabytes on 64-bit or wrap a wasm32 size_t.
  const Value v = EvalUnder("=EXPAND(A1,65536,65537)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsExpand, ExplicitPadValueIsUsed) {
  // pad_with = 0; new cells become 0.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=EXPAND(A1:C2, 3, 4, 0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  const Value* cells = v.as_array_cells();
  // Original cells preserved; new cells == 0.
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 0.0);
  EXPECT_DOUBLE_EQ(cells[8].as_number(), 0.0);
  EXPECT_DOUBLE_EQ(cells[11].as_number(), 0.0);
}

TEST(BuiltinsExpand, OmittedDimensionsAndBlankPadsProjectAtGridBoundary) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  // D1 is intentionally left blank. A reference to it is an explicit pad
  // value and, like a present empty slot, remains an internal Blank for
  // nested consumers while direct grid output projects it to zero.
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;

  const Value omitted_rows = EvalUnder("=EXPAND(A1:C2,,5)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(omitted_rows.is_array());
  ASSERT_EQ(omitted_rows.as_array_rows(), 2U);
  ASSERT_EQ(omitted_rows.as_array_cols(), 5U);
  ASSERT_TRUE(omitted_rows.as_array_cells()[3].is_error());
  EXPECT_EQ(omitted_rows.as_array_cells()[3].as_error(), ErrorCode::NA);

  parse_arena.reset();
  eval_arena.reset();
  const Value omitted_columns = EvalUnder("=EXPAND(A1:C2,4,)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(omitted_columns.is_array());
  ASSERT_EQ(omitted_columns.as_array_rows(), 4U);
  ASSERT_EQ(omitted_columns.as_array_cols(), 3U);
  ASSERT_TRUE(omitted_columns.as_array_cells()[6].is_error());
  EXPECT_EQ(omitted_columns.as_array_cells()[6].as_error(), ErrorCode::NA);

  parse_arena.reset();
  eval_arena.reset();
  const Value empty_slot_pad = EvalUnder("=EXPAND(A1:C2,4,5,)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(empty_slot_pad.is_array());
  ASSERT_EQ(empty_slot_pad.as_array_rows(), 4U);
  ASSERT_EQ(empty_slot_pad.as_array_cols(), 5U);
  ASSERT_TRUE(empty_slot_pad.as_array_cells()[3].is_number());
  EXPECT_DOUBLE_EQ(empty_slot_pad.as_array_cells()[3].as_number(), 0.0);
  ASSERT_TRUE(empty_slot_pad.as_array_cells()[10].is_number());
  EXPECT_DOUBLE_EQ(empty_slot_pad.as_array_cells()[10].as_number(), 0.0);

  parse_arena.reset();
  eval_arena.reset();
  const Value blank_pad = EvalUnder("=EXPAND(A1:C2,4,5,D1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(blank_pad.is_array());
  ASSERT_EQ(blank_pad.as_array_rows(), 4U);
  ASSERT_EQ(blank_pad.as_array_cols(), 5U);
  ASSERT_TRUE(blank_pad.as_array_cells()[3].is_number());
  EXPECT_DOUBLE_EQ(blank_pad.as_array_cells()[3].as_number(), 0.0);
  ASSERT_TRUE(blank_pad.as_array_cells()[10].is_number());
  EXPECT_DOUBLE_EQ(blank_pad.as_array_cells()[10].as_number(), 0.0);

  parse_arena.reset();
  eval_arena.reset();
  const Value blank_range_pad = EvalUnder("=EXPAND(A1:C2,4,5,D1:D2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(blank_range_pad.is_array());
  ASSERT_EQ(blank_range_pad.as_array_rows(), 4U);
  ASSERT_EQ(blank_range_pad.as_array_cols(), 5U);
  ASSERT_TRUE(blank_range_pad.as_array_cells()[3].is_number());
  EXPECT_DOUBLE_EQ(blank_range_pad.as_array_cells()[3].as_number(), 0.0);
  ASSERT_TRUE(blank_range_pad.as_array_cells()[10].is_number());
  EXPECT_DOUBLE_EQ(blank_range_pad.as_array_cells()[10].as_number(), 0.0);

  parse_arena.reset();
  eval_arena.reset();
  const Value empty_text_pad = EvalUnder("=EXPAND(A1:C2,4,5,\"\")", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(empty_text_pad.is_array());
  ASSERT_EQ(empty_text_pad.as_array_rows(), 4U);
  ASSERT_EQ(empty_text_pad.as_array_cols(), 5U);
  ASSERT_TRUE(empty_text_pad.as_array_cells()[3].is_text());
  EXPECT_TRUE(empty_text_pad.as_array_cells()[3].as_text().empty());
  ASSERT_TRUE(empty_text_pad.as_array_cells()[10].is_text());
  EXPECT_TRUE(empty_text_pad.as_array_cells()[10].as_text().empty());

  parse_arena.reset();
  eval_arena.reset();
  const Value nested_count = EvalUnder("=COUNT(EXPAND({1,2;3,4},3,4,))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(nested_count.is_number());
  EXPECT_DOUBLE_EQ(nested_count.as_number(), 4.0);

  parse_arena.reset();
  eval_arena.reset();
  const Value nested_counta = EvalUnder("=COUNTA(EXPAND({2;1},3,1,))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(nested_counta.is_number());
  EXPECT_DOUBLE_EQ(nested_counta.as_number(), 3.0);

  parse_arena.reset();
  eval_arena.reset();
  const Value nested_range_count = EvalUnder("=COUNT(EXPAND({1,2;3,4},3,4,D1:D2))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(nested_range_count.is_number());
  EXPECT_DOUBLE_EQ(nested_range_count.as_number(), 4.0);

  parse_arena.reset();
  eval_arena.reset();
  const Value nested_blank = EvalUnder("=ISBLANK(INDEX(EXPAND({1,2;3,4},3,4,),3,4))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(nested_blank.is_boolean());
  EXPECT_TRUE(nested_blank.as_boolean());
}

TEST(BuiltinsHstack, ExpandBlankPadProjectsAfterStackAndCountsAsValueArrayNested) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;

  const Value direct = EvalUnder("=HSTACK(EXPAND({1},2,1,),{2;3})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(direct.is_array());
  ASSERT_EQ(direct.as_array_rows(), 2U);
  ASSERT_EQ(direct.as_array_cols(), 2U);
  const Value* cells = direct.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  ASSERT_TRUE(cells[2].is_number());
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 0.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 3.0);

  parse_arena.reset();
  eval_arena.reset();
  const Value count = EvalUnder("=COUNT(HSTACK(EXPAND({1},2,1,),{2;3}))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(count.is_number());
  EXPECT_DOUBLE_EQ(count.as_number(), 3.0);

  parse_arena.reset();
  eval_arena.reset();
  const Value counta = EvalUnder("=COUNTA(HSTACK(EXPAND({1},2,1,),{2;3}))", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(counta.is_number());
  EXPECT_DOUBLE_EQ(counta.as_number(), 4.0);
}

TEST(BuiltinsExpand, ShrinkingRowsReturnsValue) {
  // Cannot shrink: requesting fewer rows than the source has -> #VALUE!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=EXPAND(A1:C2, 1, 3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsExpand, ShrinkingColumnsReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=EXPAND(A1:C2, 2, 1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsExpand, OneArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=EXPAND(A1:C2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// TOCOL
// ---------------------------------------------------------------------------

TEST(BuiltinsTocol, FlattensRowMajorByDefault) {
  // Default scan order is row-major: 1,2,3,4,5,6 -> 6x1 column.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOCOL(A1:C2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 6U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  for (std::uint32_t i = 0; i < 6U; ++i) {
    EXPECT_DOUBLE_EQ(v.as_array_cells()[i].as_number(), static_cast<double>(i + 1));
  }
}

TEST(BuiltinsTocol, ScanByColumnReordersToColumnMajor) {
  // scan_by_column=TRUE -> 1,4,2,5,3,6.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOCOL(A1:C2, 0, TRUE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 6U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 6.0);
}

TEST(BuiltinsTocol, IgnoreOneSkipsBlanks) {
  // Set A1 = 1, B1 = blank (cleared), C1 = 3 -> with ignore=1, result is
  // 1, 3 (2 cells).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  // B1 left blank.
  sheet.set_cell_value(0, 2, Value::number(3));
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOCOL(A1:C1, 1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 3.0);
}

TEST(BuiltinsTocol, IgnoreTwoSkipsErrors) {
  // A1=1, B1=#N/A, C1=3 -> with ignore=2, result is 1, 3.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::error(ErrorCode::NA));
  sheet.set_cell_value(0, 2, Value::number(3));
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOCOL(A1:C1, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 3.0);
}

TEST(BuiltinsTocol, IgnoreThreeSkipsBothBlanksAndErrors) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  sheet.set_cell_value(0, 1, Value::error(ErrorCode::NA));
  sheet.set_cell_value(0, 2, Value::number(3));
  // B2 left blank for testing blank skip.
  sheet.set_cell_value(1, 0, Value::number(4));
  sheet.set_cell_value(1, 2, Value::number(6));
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOCOL(A1:C2, 3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 4U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 6.0);
}

TEST(BuiltinsTocol, EmptyTextCellsBecomeBlanksButAreNotIgnored) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOCOL({1,\"\",2;\"\",3,\"\"}, 1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 6U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_TRUE(cells[1].is_blank());
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 2.0);
  EXPECT_TRUE(cells[3].is_blank());
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 3.0);
  EXPECT_TRUE(cells[5].is_blank());
}

TEST(BuiltinsTocol, ArrayLiteralErrorsCanBeIgnored) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOCOL({1,#N/A;2,3}, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 3.0);
}

TEST(BuiltinsTocol, AllSkippedReturnsCalc) {
  // A 3-cell range of blanks with ignore=1 -> nothing kept -> #CALC!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // A1:A3 left blank intentionally.
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOCOL(A1:A3, 1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

TEST(BuiltinsTocol, OutOfRangeIgnoreMaskReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOCOL(A1:C2, 4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// TOROW
// ---------------------------------------------------------------------------

TEST(BuiltinsTorow, FlattensRowMajorByDefault) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOROW(A1:C2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 6U);
  for (std::uint32_t i = 0; i < 6U; ++i) {
    EXPECT_DOUBLE_EQ(v.as_array_cells()[i].as_number(), static_cast<double>(i + 1));
  }
}

TEST(BuiltinsTorow, ScanByColumnFlag) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOROW(A1:C2, 0, TRUE)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 6U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 6.0);
}

TEST(BuiltinsTorow, EmptyTextCellsBecomeBlanksButAreNotIgnored) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOROW({\"\",1;2,\"\"}, 1)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 4U);
  const Value* cells = v.as_array_cells();
  EXPECT_TRUE(cells[0].is_blank());
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 2.0);
  EXPECT_TRUE(cells[3].is_blank());
}

TEST(BuiltinsTorow, ArrayLiteralErrorsCanBeIgnored) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOROW({#VALUE!,1,#N/A,2}, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 2.0);
}

TEST(BuiltinsTorow, AllArrayLiteralErrorsIgnoredReturnsCalc) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=TOROW({#N/A,#VALUE!}, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

// ---------------------------------------------------------------------------
// WRAPROWS
// ---------------------------------------------------------------------------

TEST(BuiltinsWraprows, BasicWrapWithDefaultPad) {
  // A1:F1 = {1,2,3,4,5,6,7}; n=7, wrap_count=3 -> 3x3 with one #N/A pad.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t c = 0; c < 7U; ++c) {
    sheet.set_cell_value(0, c, Value::number(static_cast<double>(c + 1)));
  }
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPROWS(A1:G1, 3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 6.0);
  EXPECT_DOUBLE_EQ(cells[6].as_number(), 7.0);
  ASSERT_TRUE(cells[7].is_error());
  EXPECT_EQ(cells[7].as_error(), ErrorCode::NA);
  ASSERT_TRUE(cells[8].is_error());
  EXPECT_EQ(cells[8].as_error(), ErrorCode::NA);
}

TEST(BuiltinsWraprows, ExactDivisionNoPadding) {
  // 6 cells with wrap_count=3 -> 2x3, no pad cells.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t c = 0; c < 6U; ++c) {
    sheet.set_cell_value(0, c, Value::number(static_cast<double>(c + 1)));
  }
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPROWS(A1:F1, 3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 3U);
}

TEST(BuiltinsWraprows, WrapCountExceedingLengthClipsToVectorLength) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t c = 0; c < 3U; ++c) {
    sheet.set_cell_value(0, c, Value::number(static_cast<double>(c + 1)));
  }
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPROWS(A1:C1, 10)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 3.0);
}

TEST(BuiltinsWraprows, ExplicitPadValueIsUsed) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t c = 0; c < 5U; ++c) {
    sheet.set_cell_value(0, c, Value::number(static_cast<double>(c + 1)));
  }
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPROWS(A1:E1, 3, 99)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 99.0);
}

TEST(BuiltinsWraprows, ColumnVectorIsAllowed) {
  // A column vector (Nx1) is also a 1D vector.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t r = 0; r < 4U; ++r) {
    sheet.set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPROWS(A1:A4, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 4.0);
}

TEST(BuiltinsWraprows, TwoDimensionalSourceReturnsValue) {
  // A 2D rectangle is rejected by WRAPROWS -> #VALUE!.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPROWS(A1:C2, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsWraprows, ZeroWrapCountReturnsNum) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1));
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPROWS(A1:A1, 0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// WRAPCOLS
// ---------------------------------------------------------------------------

TEST(BuiltinsWrapcols, WrapsIntoColumnMajor) {
  // n=7, wrap_count=3 -> 3x3 with last column padded:
  //   row 0: 1 4 7
  //   row 1: 2 5 #N/A
  //   row 2: 3 6 #N/A
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t c = 0; c < 7U; ++c) {
    sheet.set_cell_value(0, c, Value::number(static_cast<double>(c + 1)));
  }
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPCOLS(A1:G1, 3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 7.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 5.0);
  ASSERT_TRUE(cells[5].is_error());
  EXPECT_EQ(cells[5].as_error(), ErrorCode::NA);
  EXPECT_DOUBLE_EQ(cells[6].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[7].as_number(), 6.0);
  ASSERT_TRUE(cells[8].is_error());
  EXPECT_EQ(cells[8].as_error(), ErrorCode::NA);
}

TEST(BuiltinsWrapcols, ExactDivisionNoPadding) {
  // 6 cells with wrap_count=2 -> 2 rows, 3 cols.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t c = 0; c < 6U; ++c) {
    sheet.set_cell_value(0, c, Value::number(static_cast<double>(c + 1)));
  }
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPCOLS(A1:F1, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 3U);
  const Value* cells = v.as_array_cells();
  // Column-major fill: col 0 = {1, 2}, col 1 = {3, 4}, col 2 = {5, 6}.
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 5.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(cells[4].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 6.0);
}

TEST(BuiltinsWrapcols, WrapCountExceedingLengthClipsToVectorLength) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t c = 0; c < 3U; ++c) {
    sheet.set_cell_value(0, c, Value::number(static_cast<double>(c + 1)));
  }
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPCOLS(A1:C1, 10)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 3.0);
}

TEST(BuiltinsWrapcols, TwoDimensionalSourceReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Populate2x3(sheet);
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPCOLS(A1:C2, 2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsWrapcols, ExplicitPadValueIsUsed) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t c = 0; c < 5U; ++c) {
    sheet.set_cell_value(0, c, Value::number(static_cast<double>(c + 1)));
  }
  EvalState state;
  const EvalContext ctx = test::mac_context(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=WRAPCOLS(A1:E1, 3, 0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  // Cells:
  //   row 0: 1, 4
  //   row 1: 2, 5
  //   row 2: 3, 0  (pad)
  EXPECT_DOUBLE_EQ(cells[5].as_number(), 0.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
