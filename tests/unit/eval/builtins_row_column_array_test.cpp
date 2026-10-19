//
// Aggregator-context tests for ROW / COLUMN. The standalone scalar form
// (`=ROW(A1:A5)` -> 1) is pinned in `builtins_shape_test.cpp`; this file
// exercises the seam where ROW / COLUMN appear as range-shaped arguments
// to range-aware aggregators (SUM / SUMPRODUCT / PRODUCT). Mac Excel 365
// spills the call to `{1;2;3;4;5}` / `{1,2,3,4,5}` and the surrounding
// aggregator iterates the spill; Formulon mirrors that without a
// `Value::Array` runtime by routing through `expand_row_call` /
// `expand_column_call` in the aggregator dispatch path.

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

// Parse + evaluate against an empty workbook. References resolve through
// the bound sheet so RangeOps inside ROW / COLUMN parse correctly.
Value EvalIn(std::string_view src, const Workbook& wb, const Sheet& current) {
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

Value EvalUnbound(std::string_view src) {
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
  return evaluate(*root, eval_arena, default_registry());
}

// ---------------------------------------------------------------------------
// SUM / SUMPRODUCT / PRODUCT over ROW(range) / COLUMN(range)
// ---------------------------------------------------------------------------

TEST(BuiltinsRowColumnArray, SumOfRowOverVerticalRange) {
  // Mac Excel 365 spills ROW(A1:A5) to {1;2;3;4;5}; SUM consumes it as 15.
  // The two oracle divergences this fixes: shape_sum_of_row_over_vertical_range.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUM(ROW(A1:A5))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(BuiltinsRowColumnArray, SumOfColumnOverHorizontalRange) {
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUM(COLUMN(A1:E1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(BuiltinsRowColumnArray, SumproductOfColumnOverHorizontalRange) {
  // Matches oracle case shape_sumproduct_of_column_over_horizontal_range.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUMPRODUCT(COLUMN(A1:E1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 15.0);
}

TEST(BuiltinsRowColumnArray, SumproductOfRowOverVerticalRange) {
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUMPRODUCT(ROW(A1:A3))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsRowColumnArray, SumOfRowOverOffsetRange) {
  // Verify the offset-row case (top != 0): ROW(B5:B7) -> {5;6;7} -> 18.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUM(ROW(B5:B7))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 18.0);
}

TEST(BuiltinsRowColumnArray, ProductOfRowAcceptsRange) {
  // PRODUCT also has accepts_ranges = true. ROW(A2:A4) -> {2;3;4} -> 24.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=PRODUCT(ROW(A2:A4))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 24.0);
}

// ---------------------------------------------------------------------------
// Direct calls return the same dynamic arrays exposed to aggregators.
// ---------------------------------------------------------------------------

TEST(BuiltinsRowColumnArray, BinaryOpOverRowBroadcasts) {
  // `ROW(A1:A5)*2` -> {2;4;6;8;10}, whose sum is 30.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUM(ROW(A1:A5)*2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsRowColumnArray, StandaloneRowOverRangeSpills) {
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=ROW(A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 5U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[4].as_number(), 5.0);
}

TEST(BuiltinsRowColumnArray, StandaloneColumnOverRangeSpills) {
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=COLUMN(A1:E1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 5U);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[4].as_number(), 5.0);
}

TEST(BuiltinsRowColumnArray, RowOverRangeWithComputedEndpoint) {
  // `A1:INDEX(A1:A10,3)` is the reference A1:A3 once INDEX resolves.
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=ROW(A1:INDEX(A1:A10,3))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  for (std::uint32_t r = 0; r < 3U; ++r) {
    EXPECT_DOUBLE_EQ(v.as_array_cells()[r].as_number(), static_cast<double>(r + 1U));
  }
}

TEST(BuiltinsRowColumnArray, RowOfMissingSheetRangeIsRef) {
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=ROW(Missing!A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsRowColumnArray, SumOfRowOverMissingSheetRangeIsRef) {
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=SUM(ROW(Missing!A1:A3))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsRowColumnArray, ColumnOfMissingSheetCellIsRef) {
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=COLUMN(Missing!B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsRowColumnArray, RowOfMissingSheetChosenReferenceIsRef) {
  Workbook wb = Workbook::create();
  const Value v = EvalIn("=ROW(IF(TRUE,Missing!A1,A1))", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsRowColumnArray, KnownSheetReferenceKeepsDeclaredShapeWithoutReadingCells) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  wb.sheet(1).set_cell_value(0U, 1U, Value::number(42.0));
  EvalState state;
  state.observe_reads_of(&wb.sheet(1), 0U, 1U, 1U);

  Arena parse_arena;
  Arena eval_arena;
  parser::Parser p("=ROW(data!B1:B3)", parse_arena);
  parser::AstNode* root = p.parse();
  ASSERT_NE(root, nullptr);
  const EvalContext ctx = test::mac_context(wb, wb.sheet(0), state);
  const Value v = evaluate(*root, eval_arena, default_registry(), ctx);

  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[2].as_number(), 3.0);
  EXPECT_TRUE(state.observed_reads().empty());
}

TEST(BuiltinsRowColumnArray, UnboundContextRetainsLocalReferenceShapes) {
  const Value row = EvalUnbound("=ROW(A1:A3)");
  ASSERT_TRUE(row.is_array());
  ASSERT_EQ(row.as_array_rows(), 3U);
  ASSERT_EQ(row.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(row.as_array()->cells[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(row.as_array()->cells[2].as_number(), 3.0);

  const Value sum = EvalUnbound("=SUM(ROW(A1:A3))");
  ASSERT_TRUE(sum.is_number());
  EXPECT_DOUBLE_EQ(sum.as_number(), 6.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
