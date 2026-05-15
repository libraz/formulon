// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for `FREQUENCY(data_array, bins_array)`. The impl lives in
// `regression_lazy.{h,cpp}` and dispatches via the central
// `kLazyDispatch` table in `tree_walker.cpp`.

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

TEST(BuiltinsFrequency, BasicBucketingFromArrayLiterals) {
  // data = {1,2,3,4,5}, bins = {2,4}.
  // bin0: x <= 2 -> {1,2} = 2
  // bin1: 2 < x <= 4 -> {3,4} = 2
  // bin2 (extra): x > 4 -> {5} = 1
  // Result: column {2; 2; 1}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({1;2;3;4;5}, {2;4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(c[2].as_number(), 1.0);
}

TEST(BuiltinsFrequency, FromSheetRanges) {
  // A1:A5 = {1, 2, 3, 4, 5}; C1:C2 = {2, 4}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t r = 0; r < 5; ++r) {
    sheet.set_cell_value(r, 0, Value::number(r + 1));
  }
  sheet.set_cell_value(0, 2, Value::number(2));
  sheet.set_cell_value(1, 2, Value::number(4));
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY(A1:A5, C1:C2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(c[2].as_number(), 1.0);
}

TEST(BuiltinsFrequency, BoundaryValuesGoToTheBinTheyEqual) {
  // data = {2, 4}, bins = {2, 4}. Both values fall ON a bin edge.
  // x <= 2 -> bin0 (value 2 is here); 2 < x <= 4 -> bin1 (value 4); x > 4 -> bin2 (empty).
  // Result: {1; 1; 0}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({2;4}, {2;4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(c[2].as_number(), 0.0);
}

TEST(BuiltinsFrequency, AllValuesExceedHighestBinFillExtra) {
  // data = {10, 20, 30}, bins = {1, 2}.
  // None match any bin -> extra slot gets 3.
  // Result: {0; 0; 3}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({10;20;30}, {1;2})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 0.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 0.0);
  EXPECT_DOUBLE_EQ(c[2].as_number(), 3.0);
}

TEST(BuiltinsFrequency, NonNumericDataCellsAreSkipped) {
  // data = {1, "x", 3, TRUE}, bins = {2}. Only 1 and 3 are numeric.
  // bin0: x <= 2 -> {1} = 1
  // bin1 (extra): x > 2 -> {3} = 1
  // Result: {1; 1}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({1;\"x\";3;TRUE}, {2})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 1.0);
}

TEST(BuiltinsFrequency, NonNumericBinCellsAreSkipped) {
  // bins = {2, "x", 4}. Only 2 and 4 count as bins; output has 3 slots.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({1;2;3;4;5}, {2;\"x\";4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(c[2].as_number(), 1.0);
}

TEST(BuiltinsFrequency, EmptyBinsArrayReturnsTotalCount) {
  // bins entirely non-numeric -> a 1x1 array containing total numeric
  // data count.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({1;2;3}, {\"x\"})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 1U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 3.0);
}

TEST(BuiltinsFrequency, AllNonNumericDataReturnsAllZeros) {
  // No numeric data -> every output slot is 0.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({\"a\";\"b\"}, {1;2})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 0.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 0.0);
  EXPECT_DOUBLE_EQ(c[2].as_number(), 0.0);
}

TEST(BuiltinsFrequency, ErrorInDataArrayPropagates) {
  // #N/A in data_array propagates, with leftmost-wins (data_array before
  // bins_array).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({1;#N/A;3}, {2})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsFrequency, ErrorInBinsArraySilentlySkipped) {
  // Errors in bins_array are treated as non-numeric and silently
  // skipped (matches Mac Excel). With every bin cell being an error,
  // the bins reduce to the empty list and the result is the degenerate
  // 1x1 array containing count(numeric_data).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({1;2;3}, {#DIV/0!})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array()) << v.debug_to_string();
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array()->cells[0].as_number(), 3.0);
}

TEST(BuiltinsFrequency, DataErrorBeatsBinError) {
  // Both arrays contain errors; data_array's error wins (left-to-right).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({#N/A}, {#DIV/0!})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(BuiltinsFrequency, OneArgRejected) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({1;2;3})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsFrequency, BinsTakenAsGivenNotSorted) {
  // bins given out of order -> Excel sorts numeric bins ascending before
  // bucketing.
  // data = {1, 5, 10}, bins = {7, 3}.
  // sorted bins = {3, 7}; result: {1; 1; 1}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({1;5;10}, {7;3})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(c[2].as_number(), 1.0);
}

TEST(BuiltinsFrequency, RowVectorBinsAreAcceptedAsArray) {
  // bins given as a row vector instead of column vector. Both shapes are
  // valid since FREQUENCY only inspects cell values, not rectangle
  // shape.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({1;2;3;4;5}, {2,4})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  const Value* c = v.as_array_cells();
  EXPECT_DOUBLE_EQ(c[0].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(c[1].as_number(), 2.0);
  EXPECT_DOUBLE_EQ(c[2].as_number(), 1.0);
}

TEST(BuiltinsFrequency, OutputIsAlwaysColumnEvenForRowVectorBins) {
  // Verify the output rectangle is (n+1) x 1 regardless of bin layout.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FREQUENCY({1;2;3}, {1,2,3})", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cols(), 1U);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
