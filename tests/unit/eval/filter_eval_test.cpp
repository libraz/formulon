// Copyright 2026 libraz. Licensed under the MIT License.
//
// Tests for the lazy `FILTER(array, include, [if_empty])` builtin.
//
// FILTER is the first non-shape dynamic-array spilling builtin in Formulon.
// It evaluates both `array` and `include` via `eval_node_as_array` to
// preserve their 2D shapes, decides whether to filter rows or cols by
// matching `include`'s shape against `array`, then copies the kept rows /
// cols into a freshly arena-allocated output array. Empty results return
// `if_empty` if provided, else `#CALC!`.

#include <cstdint>
#include <string_view>
#include <vector>

#include "cell.h"
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

// Fills A1:Cn with the given column-major payload (each column's cells run
// top-to-bottom). Used to seed a 2D array for FILTER tests.
void SeedRect(Sheet& sheet, std::uint32_t rows, std::uint32_t cols, std::initializer_list<double> values) {
  EXPECT_EQ(values.size(), static_cast<std::size_t>(rows) * cols);
  auto it = values.begin();
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      sheet.set_cell_value(r, c, Value::number(*it++));
    }
  }
}

// ---------------------------------------------------------------------------
// Row-axis filtering (column-vector mask)
// ---------------------------------------------------------------------------

TEST(BuiltinsFilter, FilterRowsByColumnVectorMask) {
  // A1:B4 = {{10,1},{20,2},{30,3},{40,4}}; mask = {TRUE;FALSE;TRUE;FALSE}.
  // Expected: 2x2 array {{10,1},{30,3}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  SeedRect(sheet, 4U, 2U, {10, 1, 20, 2, 30, 3, 40, 4});
  // Mask in D1:D4
  sheet.set_cell_value(0, 3, Value::boolean(true));
  sheet.set_cell_value(1, 3, Value::boolean(false));
  sheet.set_cell_value(2, 3, Value::boolean(true));
  sheet.set_cell_value(3, 3, Value::boolean(false));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FILTER(A1:B4, D1:D4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* cells = v.as_array_cells();
  EXPECT_DOUBLE_EQ(cells[0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(cells[1].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(cells[2].as_number(), 30.0);
  EXPECT_DOUBLE_EQ(cells[3].as_number(), 3.0);
}

TEST(BuiltinsFilter, FilterRowsWithComputedMask) {
  // A1:A5 = {1,2,3,4,5}. Mask via comparison: A1:A5>2 -> {F,F,T,T,T}.
  // Expected: 3x1 column vector {3,4,5}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t r = 0; r < 5U; ++r) {
    sheet.set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FILTER(A1:A5, A1:A5>2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 3U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 3.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 4.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 5.0);
}

// ---------------------------------------------------------------------------
// Column-axis filtering (row-vector mask)
// ---------------------------------------------------------------------------

TEST(BuiltinsFilter, FilterColsByRowVectorMask) {
  // A1:D2 = {{10,20,30,40},{1,2,3,4}}; mask = {TRUE,FALSE,TRUE,FALSE} as a
  // row vector. Expected: 2x2 array {{10,30},{1,3}}.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(10));
  sheet.set_cell_value(0, 1, Value::number(20));
  sheet.set_cell_value(0, 2, Value::number(30));
  sheet.set_cell_value(0, 3, Value::number(40));
  sheet.set_cell_value(1, 0, Value::number(1));
  sheet.set_cell_value(1, 1, Value::number(2));
  sheet.set_cell_value(1, 2, Value::number(3));
  sheet.set_cell_value(1, 3, Value::number(4));
  // Mask A4:D4 (row 3, 0-based)
  sheet.set_cell_value(3, 0, Value::boolean(true));
  sheet.set_cell_value(3, 1, Value::boolean(false));
  sheet.set_cell_value(3, 2, Value::boolean(true));
  sheet.set_cell_value(3, 3, Value::boolean(false));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FILTER(A1:D2, A4:D4)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 10.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 30.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[2].as_number(), 1.0);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[3].as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// Empty-result handling
// ---------------------------------------------------------------------------

TEST(BuiltinsFilter, NoMatchWithIfEmptyReturnsScalar) {
  // FILTER over an all-FALSE mask with `if_empty` returns the scalar value
  // verbatim (Mac Excel does NOT spill it).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t r = 0; r < 3U; ++r) {
    sheet.set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FILTER(A1:A3, A1:A3<0, \"none\")", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "none");
}

TEST(BuiltinsFilter, NoMatchWithoutIfEmptyReturnsCalc) {
  // Without if_empty, an all-FALSE mask surfaces #CALC! per Mac Excel.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t r = 0; r < 3U; ++r) {
    sheet.set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FILTER(A1:A3, A1:A3<0)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Calc);
}

// ---------------------------------------------------------------------------
// Shape errors
// ---------------------------------------------------------------------------

TEST(BuiltinsFilter, ShapeMismatchYieldsValue) {
  // Mask shape (2x1) does not match array's row-axis (3x1) nor cols (1).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t r = 0; r < 3U; ++r) {
    sheet.set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  sheet.set_cell_value(0, 1, Value::boolean(true));
  sheet.set_cell_value(1, 1, Value::boolean(false));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FILTER(A1:A3, B1:B2)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Error propagation
// ---------------------------------------------------------------------------

TEST(BuiltinsFilter, ErrorInArrayPropagatesPerCell) {
  // An error cell inside `array` is preserved in the kept output position
  // -- FILTER does not coerce array cells, only routes them.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(10.0));
  sheet.set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  sheet.set_cell_value(2, 0, Value::number(30.0));
  sheet.set_cell_value(0, 1, Value::boolean(true));
  sheet.set_cell_value(1, 1, Value::boolean(true));
  sheet.set_cell_value(2, 1, Value::boolean(false));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FILTER(A1:A3, B1:B3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 1U);
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 10.0);
  ASSERT_TRUE(v.as_array_cells()[1].is_error());
  EXPECT_EQ(v.as_array_cells()[1].as_error(), ErrorCode::Div0);
}

TEST(BuiltinsFilter, ErrorInIncludeShortCircuits) {
  // An error cell in the `include` mask short-circuits the entire FILTER
  // call (conservative behaviour matching the dominant Mac Excel path).
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t r = 0; r < 3U; ++r) {
    sheet.set_cell_value(r, 0, Value::number(static_cast<double>(r + 1)));
  }
  sheet.set_cell_value(0, 1, Value::boolean(true));
  sheet.set_cell_value(1, 1, Value::error(ErrorCode::NA));
  sheet.set_cell_value(2, 1, Value::boolean(true));

  EvalState state;
  const EvalContext ctx(wb, sheet, state);

  Arena parse_arena;
  Arena eval_arena;
  const Value v = EvalUnder("=FILTER(A1:A3, B1:B3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Arity gates
// ---------------------------------------------------------------------------

TEST(BuiltinsFilter, ArityMismatchReturnsValue) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  EvalState state;
  const EvalContext ctx(wb, sheet, state);
  Arena parse_arena;
  Arena eval_arena;
  // Only 1 arg provided -> #VALUE!.
  const Value v = EvalUnder("=FILTER(A1:A3)", &parse_arena, &eval_arena, ctx);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
