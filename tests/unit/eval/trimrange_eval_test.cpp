// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Tests for the lazy `TRIMRANGE(range, [trim_rows]=3, [trim_cols]=3)`
// builtin. Trim modes: 0 = none, 1 = leading, 2 = trailing, 3 = both. Only
// the `Blank` value variant counts as trimmable; `""`, `0`, `FALSE`, and
// errors are preserved.

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

Value EvalSrc(std::string_view src) {
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
  return evaluate(*root, eval_arena);
}

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
  const EvalContext ctx(wb, current, state);
  return evaluate(*root, eval_arena, default_registry(), ctx);
}

// ---------------------------------------------------------------------------
// Default mode (both edges, both axes)
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, DefaultTrimsBothEdgesOnColumn) {
  // Column literal {""; 1; ""}: leading "" (text) is preserved (not Blank),
  // so emulate Blank using a Workbook with empty cells around a single value.
  Workbook wb = Workbook::create();
  // A1 blank (default), A2=1, A3 blank.
  wb.sheet(0).set_cell_value(1, 0, Value::number(1.0));
  const Value v = EvalIn("=TRIMRANGE(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  ASSERT_TRUE(v.as_array_cells()[0].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
}

TEST(BuiltinsTrimRange, DefaultTrimsBlankBorderOf3x3) {
  // 5x5 grid with a 3x3 inner block of values surrounded by blanks.
  // Default trim-mode 3 should collapse to the 3x3 inner block.
  Workbook wb = Workbook::create();
  // Inner block is rows 1..3, cols 1..3 (zero-indexed); fill with 1..9.
  double k = 1.0;
  for (std::uint32_t r = 1; r <= 3; ++r) {
    for (std::uint32_t c = 1; c <= 3; ++c) {
      wb.sheet(0).set_cell_value(r, c, Value::number(k));
      k += 1.0;
    }
  }
  const Value v = EvalIn("=TRIMRANGE(A1:E5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 3U);
  for (std::uint32_t i = 0; i < 9; ++i) {
    ASSERT_TRUE(v.as_array_cells()[i].is_number()) << "i=" << i;
    EXPECT_DOUBLE_EQ(v.as_array_cells()[i].as_number(), static_cast<double>(i + 1));
  }
}

TEST(BuiltinsTrimRange, DefaultTrimsTopLeftBlanks) {
  // Inner block at rows 2..3, cols 2..3.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(2, 2, Value::number(1));
  wb.sheet(0).set_cell_value(2, 3, Value::number(2));
  wb.sheet(0).set_cell_value(3, 2, Value::number(3));
  wb.sheet(0).set_cell_value(3, 3, Value::number(4));
  const Value v = EvalIn("=TRIMRANGE(A1:E5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
}

TEST(BuiltinsTrimRange, DefaultTrimsBottomRightBlanks) {
  // Inner block at rows 0..1, cols 0..1.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3));
  wb.sheet(0).set_cell_value(1, 1, Value::number(4));
  const Value v = EvalIn("=TRIMRANGE(A1:E5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 2U);
  EXPECT_EQ(v.as_array_cols(), 2U);
}

// ---------------------------------------------------------------------------
// Mode 0: no trimming
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, ModeZeroPreservesAllRowsAndCols) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(2, 2, Value::number(42));
  const Value v = EvalIn("=TRIMRANGE(A1:E5, 0, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 5U);
  EXPECT_EQ(v.as_array_cols(), 5U);
}

// ---------------------------------------------------------------------------
// Mode 1: leading-only
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, ModeOneLeadingOnlyKeepsTrailingBlanks) {
  // Value at (1, 1); leading-only trim keeps the trailing blanks at rows 2..4
  // and cols 2..4.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 1, Value::number(7));
  const Value v = EvalIn("=TRIMRANGE(A1:E5, 1, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  ASSERT_TRUE(v.as_array_cells()[0].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 7.0);
}

// ---------------------------------------------------------------------------
// Mode 2: trailing-only
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, ModeTwoTrailingOnlyKeepsLeadingBlanks) {
  // Value at (3, 3); trailing-only trim keeps the leading blanks at rows
  // 0..2 and cols 0..2.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(3, 3, Value::number(9));
  const Value v = EvalIn("=TRIMRANGE(A1:E5, 2, 2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 4U);
  EXPECT_EQ(v.as_array_cols(), 4U);
  // The non-blank cell sits at the bottom-right of the kept rectangle.
  const std::uint32_t rows = v.as_array_rows();
  const std::uint32_t cols = v.as_array_cols();
  const std::size_t last = static_cast<std::size_t>(rows - 1U) * cols + (cols - 1U);
  ASSERT_TRUE(v.as_array_cells()[last].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[last].as_number(), 9.0);
}

// ---------------------------------------------------------------------------
// Explicit mode 3 == default
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, ModeThreeMatchesDefault) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(2, 2, Value::number(42));
  const Value v_default = EvalIn("=TRIMRANGE(A1:E5)", wb, wb.sheet(0));
  const Value v_explicit = EvalIn("=TRIMRANGE(A1:E5, 3, 3)", wb, wb.sheet(0));
  ASSERT_TRUE(v_default.is_array());
  ASSERT_TRUE(v_explicit.is_array());
  EXPECT_EQ(v_default.as_array_rows(), v_explicit.as_array_rows());
  EXPECT_EQ(v_default.as_array_cols(), v_explicit.as_array_cols());
  EXPECT_EQ(v_default.as_array_rows(), 1U);
  EXPECT_EQ(v_default.as_array_cols(), 1U);
}

// ---------------------------------------------------------------------------
// Independent row/col modes
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, RowsTrimmedColsPreserved) {
  // Inner row stripe at row 2 only; trim_rows=1 (leading), trim_cols=0.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(2, 2, Value::number(5));
  const Value v = EvalIn("=TRIMRANGE(A1:E5, 1, 0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  // Leading rows 0,1 trimmed; trailing rows 3,4 kept (mode 1 = leading only).
  EXPECT_EQ(v.as_array_rows(), 3U);
  // Columns untouched.
  EXPECT_EQ(v.as_array_cols(), 5U);
}

// ---------------------------------------------------------------------------
// Non-blank tokens that look blank but are not
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, EmptyStringIsNotBlank) {
  // {"";1;""} — leading "" cell is empty TEXT (not Blank), so the row should
  // be preserved by default trim. Output is the full 3x1.
  const Value v = EvalSrc("=TRIMRANGE({\"\";1;\"\"})");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
}

TEST(BuiltinsTrimRange, ZeroIsNotBlank) {
  const Value v = EvalSrc("=TRIMRANGE({0;1;0})");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
}

TEST(BuiltinsTrimRange, FalseIsNotBlank) {
  const Value v = EvalSrc("=TRIMRANGE({FALSE;1;FALSE})");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
}

TEST(BuiltinsTrimRange, ErrorCellIsNotBlank) {
  // #N/A in an otherwise-blank trailing row keeps that row in the output.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1));
  wb.sheet(0).set_cell_value(2, 0, Value::error(ErrorCode::NA));
  const Value v = EvalIn("=TRIMRANGE(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  // Row 1 (blank) is between two non-blank rows, so default trim keeps the
  // full 3x1 column.
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  // The error cell survives at row 2.
  ASSERT_TRUE(v.as_array_cells()[2].is_error());
  EXPECT_EQ(v.as_array_cells()[2].as_error(), ErrorCode::NA);
}

TEST(BuiltinsTrimRange, MixedRowWithBlankAndNumberPreserved) {
  // 2x2 with row 0 = [1, blank], row 1 = [blank, blank]. Row 0 is mixed
  // (one blank cell, one number); not fully blank, so it survives. Row 1
  // is fully blank trailing -> trimmed. Column 0 has a value at row 0;
  // column 1 is fully blank -> trimmed. Result is 1x1 with the number 1.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1));
  const Value v = EvalIn("=TRIMRANGE(A1:B2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  ASSERT_TRUE(v.as_array_cells()[0].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 1.0);
}

// ---------------------------------------------------------------------------
// Single-cell pass-through
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, ScalarInputYields1x1) {
  const Value v = EvalSrc("=TRIMRANGE(42)");
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 1U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  ASSERT_TRUE(v.as_array_cells()[0].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 42.0);
}

// ---------------------------------------------------------------------------
// All-blank input -> #REF!
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, AllBlankInputYieldsRef) {
  Workbook wb = Workbook::create();
  // A1:B2 all blank.
  const Value v = EvalIn("=TRIMRANGE(A1:B2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// Single-column, leading-only trimming
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, SingleColumnLeadingTrim) {
  // 4x1 column with blanks at rows 0 and 3, value at rows 1 and 2. Mode 1
  // (leading) on rows trims row 0 only; trailing blank at row 3 kept.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 0, Value::number(10));
  wb.sheet(0).set_cell_value(2, 0, Value::number(20));
  const Value v = EvalIn("=TRIMRANGE(A1:A4, 1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 3U);
  EXPECT_EQ(v.as_array_cols(), 1U);
  ASSERT_TRUE(v.as_array_cells()[0].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[0].as_number(), 10.0);
  ASSERT_TRUE(v.as_array_cells()[1].is_number());
  EXPECT_DOUBLE_EQ(v.as_array_cells()[1].as_number(), 20.0);
  EXPECT_TRUE(v.as_array_cells()[2].is_blank());
}

// ---------------------------------------------------------------------------
// Out-of-range / non-numeric / error mode args
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, ModeFourRejected) {
  const Value v = EvalSrc("=TRIMRANGE({1}, 4)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsTrimRange, ModeNegativeRejected) {
  const Value v = EvalSrc("=TRIMRANGE({1}, -1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsTrimRange, NonNumericModeArgRejected) {
  const Value v = EvalSrc("=TRIMRANGE({1}, \"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsTrimRange, ModeArgErrorPropagates) {
  const Value v = EvalSrc("=TRIMRANGE({1}, 1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsTrimRange, RangeArgErrorPropagates) {
  const Value v = EvalSrc("=TRIMRANGE(1/0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Arity rejection
// ---------------------------------------------------------------------------

TEST(BuiltinsTrimRange, ZeroArgsRejected) {
  const Value v = EvalSrc("=TRIMRANGE()");
  ASSERT_TRUE(v.is_error());
}

TEST(BuiltinsTrimRange, FourArgsRejected) {
  const Value v = EvalSrc("=TRIMRANGE({1}, 0, 0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
