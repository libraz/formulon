// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for SUBTOTAL — the multi-mode aggregator that dispatches
// on a numeric function code (1..11 / 101..111). Tests pin:
//   * Each of the 11 modes returns the right scalar.
//   * The 100+ "ignore-hidden" codes fold to identical results (Formulon
//     does not yet model row visibility).
//   * Code 3 (COUNTA) sees Bool / Text / Error cells inside a range and
//     counts them, while every other code skips non-numeric range cells.
//   * Out-of-range or non-finite codes surface `#VALUE!`.
//   * Truncating dispatch: 9.7 -> SUM, 109.4 -> SUM (matches Excel).

#include <cstdint>
#include <string>
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

Value EvalSource(std::string_view src) {
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

Value EvalSourceIn(std::string_view src, const Workbook& wb, const Sheet& current) {
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

// Seeds A1:A5 with 10, 20, 30, 40, 50 and returns a workbook bound to that
// sheet so the parameterized tests below can assert against a known column.
Workbook MakeNumericRange() {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(40.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(50.0));
  return wb;
}

// ---------------------------------------------------------------------------
// Registry
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotalRegistry, Registered) {
  const FunctionDef* def = default_registry().lookup("SUBTOTAL");
  ASSERT_NE(def, nullptr);
  EXPECT_TRUE(def->accepts_ranges);
  EXPECT_FALSE(def->propagate_errors);  // COUNTA must see error cells.
  EXPECT_FALSE(def->range_filter_numeric_only);
  EXPECT_EQ(def->min_arity, 2u);
}

// ---------------------------------------------------------------------------
// Mode dispatch
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, Code1IsAverage) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(1,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsSubtotal, Code2IsCount) {
  Workbook wb = MakeNumericRange();
  // Add a non-numeric cell to A6 to confirm COUNT skips it.
  wb.sheet(0).set_cell_value(5, 0, Value::text("text"));
  const Value v = EvalSourceIn("=SUBTOTAL(2,A1:A6)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsSubtotal, Code3IsCountA) {
  Workbook wb = MakeNumericRange();
  // Mix of types: 5 numbers + 1 text + 1 bool + 1 blank should yield 7.
  wb.sheet(0).set_cell_value(5, 0, Value::text("text"));
  wb.sheet(0).set_cell_value(6, 0, Value::boolean(true));
  // Row 8 (index 7) intentionally left blank; COUNTA must skip it.
  const Value v = EvalSourceIn("=SUBTOTAL(3,A1:A8)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 7.0);
}

TEST(BuiltinsSubtotal, Code4IsMax) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(4,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 50.0);
}

TEST(BuiltinsSubtotal, Code5IsMin) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(5,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsSubtotal, Code6IsProduct) {
  // Use a smaller range to keep the product representable as an exact int.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(4.0));
  const Value v = EvalSourceIn("=SUBTOTAL(6,A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 24.0);
}

TEST(BuiltinsSubtotal, Code9IsSum) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(9,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 150.0);
}

// VAR / VARP / STDEV / STDEVP — small symmetric range so the expected values
// are easy to verify by hand. Sample variance of {1, 2, 3, 4, 5} = 2.5;
// population variance = 2.0.
TEST(BuiltinsSubtotal, Code10IsVarSample) {
  Workbook wb = Workbook::create();
  for (uint32_t i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(i, 0, Value::number(static_cast<double>(i + 1)));
  }
  const Value v = EvalSourceIn("=SUBTOTAL(10,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.5);
}

TEST(BuiltinsSubtotal, Code11IsVarPopulation) {
  Workbook wb = Workbook::create();
  for (uint32_t i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(i, 0, Value::number(static_cast<double>(i + 1)));
  }
  const Value v = EvalSourceIn("=SUBTOTAL(11,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsSubtotal, Code7IsStdevSample) {
  Workbook wb = Workbook::create();
  for (uint32_t i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(i, 0, Value::number(static_cast<double>(i + 1)));
  }
  const Value v = EvalSourceIn("=SUBTOTAL(7,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), std::sqrt(2.5));
}

TEST(BuiltinsSubtotal, Code8IsStdevPopulation) {
  Workbook wb = Workbook::create();
  for (uint32_t i = 0; i < 5; ++i) {
    wb.sheet(0).set_cell_value(i, 0, Value::number(static_cast<double>(i + 1)));
  }
  const Value v = EvalSourceIn("=SUBTOTAL(8,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), std::sqrt(2.0));
}

// ---------------------------------------------------------------------------
// 100+ ignore-hidden codes fold to identical results.
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, Code109EqualsCode9) {
  Workbook wb = MakeNumericRange();
  const Value v9 = EvalSourceIn("=SUBTOTAL(9,A1:A5)", wb, wb.sheet(0));
  const Value v109 = EvalSourceIn("=SUBTOTAL(109,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v9.is_number());
  ASSERT_TRUE(v109.is_number());
  EXPECT_DOUBLE_EQ(v9.as_number(), v109.as_number());
}

TEST(BuiltinsSubtotal, Code103EqualsCode3) {
  Workbook wb = MakeNumericRange();
  wb.sheet(0).set_cell_value(5, 0, Value::text("foo"));
  const Value v3 = EvalSourceIn("=SUBTOTAL(3,A1:A6)", wb, wb.sheet(0));
  const Value v103 = EvalSourceIn("=SUBTOTAL(103,A1:A6)", wb, wb.sheet(0));
  ASSERT_TRUE(v3.is_number());
  ASSERT_TRUE(v103.is_number());
  EXPECT_DOUBLE_EQ(v3.as_number(), v103.as_number());
}

// ---------------------------------------------------------------------------
// Truncating code (matches Excel's TRUNC-before-dispatch rule).
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, FractionalCodeTruncatesToSum) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(9.7,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 150.0);
}

// ---------------------------------------------------------------------------
// Bad codes
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, ZeroCodeRejected) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(0,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSubtotal, OutOfRangeCodeRejected) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(12,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSubtotal, BetweenBlockCodeRejected) {
  // Codes 12..100 fall in the gap between the two valid blocks.
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(50,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSubtotal, AboveHiddenBlockCodeRejected) {
  Workbook wb = MakeNumericRange();
  const Value v = EvalSourceIn("=SUBTOTAL(112,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Provenance: range cells of non-numeric kind are skipped (numeric modes)
// or counted (COUNTA mode). Errors propagate in numeric modes; COUNTA
// counts them.
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, SumSkipsNonNumericInRange) {
  Workbook wb = MakeNumericRange();
  wb.sheet(0).set_cell_value(5, 0, Value::text("nope"));
  wb.sheet(0).set_cell_value(6, 0, Value::boolean(true));
  // Blank at row 8 (index 7) silently skipped.
  const Value v = EvalSourceIn("=SUBTOTAL(9,A1:A8)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 150.0);
}

TEST(BuiltinsSubtotal, CountASkipsBlankButCountsErrors) {
  Workbook wb = MakeNumericRange();
  wb.sheet(0).set_cell_value(5, 0, Value::error(ErrorCode::NA));
  // Row 7 blank.
  const Value v = EvalSourceIn("=SUBTOTAL(3,A1:A7)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  // 5 numbers + 1 error = 6; the blank at A7 is the only skipped cell.
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsSubtotal, SumPropagatesErrorInRange) {
  Workbook wb = MakeNumericRange();
  wb.sheet(0).set_cell_value(2, 0, Value::error(ErrorCode::NA));
  const Value v = EvalSourceIn("=SUBTOTAL(9,A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Empty / degenerate ranges
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, AverageOverEmptyRangeIsDiv0) {
  Workbook wb = Workbook::create();
  // A1:A3 entirely blank -> SUBTOTAL(1, A1:A3) sees no numerics.
  const Value v = EvalSourceIn("=SUBTOTAL(1,A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsSubtotal, MinOverEmptyRangeIsZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=SUBTOTAL(5,A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsSubtotal, VarSingleSampleIsDiv0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value v = EvalSourceIn("=SUBTOTAL(10,A1:A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// Direct (literal) arguments
// ---------------------------------------------------------------------------

TEST(BuiltinsSubtotal, DirectArgsSum) {
  // Mac Excel 365 accepts direct numeric literals as inputs to SUBTOTAL.
  // Our impl drops non-Number direct values silently (matching the range
  // rule); for purely numeric direct args the behaviour matches Mac.
  const Value v = EvalSource("=SUBTOTAL(9,1,2,3,4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsSubtotal, MissingRangeRejected) {
  // Single-arg form is rejected by min_arity = 2 (no implicit empty data).
  const Value v = EvalSource("=SUBTOTAL(9)");
  ASSERT_TRUE(v.is_error());
}

}  // namespace
}  // namespace eval
}  // namespace formulon
