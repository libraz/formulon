// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the counting aggregators: COUNT, COUNTA,
// COUNTBLANK. Each test parses a formula source, evaluates the AST
// through the default registry, and asserts the resulting Value.
//
// COUNTA and COUNTBLANK are registered with `propagate_errors = false`
// and `accepts_ranges = true`, so errors inside expanded ranges reach the
// impl directly and are skipped or counted per Excel semantics.
//
// COUNT itself is NOT in the eager registry: it lives in the lazy
// dispatch table (`eval_count_lazy` in special_forms_lazy.cpp) because
// Excel's "direct-arg Bool counts, range-sourced Bool doesn't" rule
// requires per-arg AST inspection that the eager dispatcher's flattened
// values vector has already erased.

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

// Parses `src` and evaluates it via the default function registry. Arenas
// are reset on each call to avoid cross-test contamination while keeping
// text payloads readable for the assertions that follow.
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

// Parses `src` and evaluates it against a bound workbook + current sheet.
// Used by tests that exercise range expansion through `expand_range`.
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

// Invokes a registered function impl directly with `arity` arguments.
// Used for Blank-containing cases that cannot be expressed in formula
// syntax without a workbook.
Value CallDirect(std::string_view name, const Value* args, std::uint32_t arity) {
  static thread_local Arena arena;
  arena.reset();
  const FunctionDef* def = default_registry().lookup(name);
  EXPECT_NE(def, nullptr) << "function not registered: " << name;
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return def->impl(args, arity, arena);
}

// ---------------------------------------------------------------------------
// Registry pins
// ---------------------------------------------------------------------------

TEST(BuiltinsCountRegistry, CountaAndCountblankRegistered) {
  // COUNTA / COUNTBLANK ride the eager registry path. COUNT is lazy and
  // therefore absent from the registry; the dispatch table in
  // tree_walker.cpp routes it directly.
  EXPECT_EQ(default_registry().lookup("COUNT"), nullptr);
  EXPECT_NE(default_registry().lookup("COUNTA"), nullptr);
  EXPECT_NE(default_registry().lookup("COUNTBLANK"), nullptr);
}

TEST(BuiltinsCountRegistry, CountaAndCountblankRangeAwareAndNonPropagating) {
  const FunctionDef* counta = default_registry().lookup("COUNTA");
  const FunctionDef* countblank = default_registry().lookup("COUNTBLANK");
  ASSERT_NE(counta, nullptr);
  ASSERT_NE(countblank, nullptr);
  EXPECT_TRUE(counta->accepts_ranges);
  EXPECT_TRUE(countblank->accepts_ranges);
  EXPECT_FALSE(counta->propagate_errors);
  EXPECT_FALSE(countblank->propagate_errors);
}

// ---------------------------------------------------------------------------
// COUNT - only Number values contribute.
// ---------------------------------------------------------------------------

TEST(BuiltinsCount, SingleNumberIsOne) {
  const Value v = EvalSource("=COUNT(42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCount, MultipleNumbers) {
  const Value v = EvalSource("=COUNT(1, 2, 3, 4, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsCount, NegativeAndFractionalNumbers) {
  const Value v = EvalSource("=COUNT(-1, 0, 3.14, -2.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsCount, BooleanLiteralIsCountedAsDirectArg) {
  // Excel pin: COUNT counts booleans when they are direct arguments, but
  // NOT when they are sourced from a range. Verified against Mac Excel 365
  // (`=COUNT(TRUE)` -> 1).
  const Value v = EvalSource("=COUNT(TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCount, NumericTextLiteralIsCoerced) {
  // Excel pin: a direct text-literal argument that parses as a finite
  // number counts. Verified against Mac Excel 365 and the IronCalc oracle
  // fixture calc_tests/COUNT.xlsx F27 (`=COUNT("23")` -> 1).
  const Value v = EvalSource("=COUNT(\"5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCount, NonNumericTextIsSkipped) {
  const Value v = EvalSource("=COUNT(\"hello\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCount, DirectErrorArgIsSkipped) {
  // Accepted divergence: Excel propagates #DIV/0! here. Our range-vs-direct
  // parity rule skips the error instead. See builtins.cpp block comment.
  const Value v = EvalSource("=COUNT(1, #DIV/0!)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCount, MixedDirectArgsCountsNumbersBoolsAndCoercibleText) {
  // Direct-arg rule: numbers, booleans, and numeric-looking text all
  // count; non-coercible text and errors skip. Of the seven args below,
  // {1, 3, 4.5} are numbers, {TRUE, FALSE} are booleans, and "2" parses
  // as a finite number -> 6. "text" fails coercion and is skipped.
  const Value v = EvalSource("=COUNT(1, TRUE, \"2\", 3, FALSE, \"text\", 4.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsCount, SingleCellBoolRefNotCounted) {
  // Excel pin: a single-cell Ref to a Bool behaves like a Bool inside a
  // range -- it does NOT count, even though a direct Bool literal does.
  // IronCalc oracle fixture calc_tests/COUNT.xlsx B16 codifies this:
  // `=COUNT(B2,...,B10)` with B5=TRUE yields 3, not 4.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::boolean(true));
  const Value v = EvalSourceIn("=COUNT(A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCount, DirectLiteralTextCoerced) {
  // Excel pin: a text literal that parses as a finite number counts.
  // IronCalc oracle fixture calc_tests/COUNT.xlsx F27 (`=COUNT("23")` -> 1).
  const Value v = EvalSource("=COUNT(\"23\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCount, DirectLiteralTextNotCoerced) {
  // A text literal that cannot be parsed as a number is skipped.
  const Value v = EvalSource("=COUNT(\"Hello\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCount, DirectLiteralBoolCounted) {
  // Excel pin (duplicate of BooleanLiteralIsCountedAsDirectArg under the
  // requested naming): a direct Bool literal counts.
  const Value v = EvalSource("=COUNT(TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCount, BlankCellInRangeIsSkipped) {
  // Blank has no formula literal and COUNT is lazy (not in the registry),
  // so we can't call the impl directly. Instead route a Blank through a
  // range: A2 is unset, so the range flattens as [1, Blank, 2]. COUNT
  // skips the blank and returns 2.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  // row 1 left blank
  wb.sheet(0).set_cell_value(2, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=COUNT(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCount, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=COUNT()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCount, RangeWithNumbersOnly) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30.0));
  const Value v = EvalSourceIn("=COUNT(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsCount, RangeWithMixedTypes) {
  // Half populated: numbers, text, blank, bool. Only numbers counted.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("hello"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(3, 0, Value::boolean(true));
  // row 4 left blank
  wb.sheet(0).set_cell_value(5, 0, Value::number(3.0));
  const Value v = EvalSourceIn("=COUNT(A1:A6)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsCount, RangeWithErrorCellSkipsError) {
  // A2 is =1/0 → #DIV/0!. COUNT must silently skip it (propagate_errors=false)
  // and return the count of the valid numeric cells.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_formula(1, 0, "=1/0");
  wb.sheet(0).set_cell_value(2, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=COUNT(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCount, CrossSheetRange) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.sheet(1).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 0, Value::text("skip"));
  wb.sheet(1).set_cell_value(2, 0, Value::number(3.0));
  const Value v = EvalSourceIn("=COUNT(Sheet2!A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// COUNTA - everything except Blank contributes.
// ---------------------------------------------------------------------------

TEST(BuiltinsCountA, SingleNumberIsOne) {
  const Value v = EvalSource("=COUNTA(42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCountA, TextIsCounted) {
  const Value v = EvalSource("=COUNTA(\"hello\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCountA, EmptyStringLiteralIsCounted) {
  // A formula returning "" produces a Text value, NOT Blank, so COUNTA
  // counts it. Matches Excel exactly.
  const Value v = EvalSource("=COUNTA(\"\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCountA, BooleanIsCounted) {
  const Value v = EvalSource("=COUNTA(TRUE, FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountA, ErrorIsCounted) {
  // Errors are non-blank, so COUNTA counts them. Combined with
  // propagate_errors=false this lets COUNTA report "populated cell count"
  // even when some cells are in an error state.
  const Value v = EvalSource("=COUNTA(1, #DIV/0!, \"x\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsCountA, MixedTypesAllCountedExceptBlank) {
  const Value v = EvalSource("=COUNTA(1, \"a\", TRUE, FALSE, \"\", 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6.0);
}

TEST(BuiltinsCountA, BlankArgIsSkipped) {
  const Value args[] = {Value::number(1.0), Value::blank(), Value::text("x"),
                        Value::blank(), Value::boolean(true)};
  const Value v = CallDirect("COUNTA", args, 5u);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsCountA, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=COUNTA()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCountA, RangeMixedHalfPopulated) {
  // A1..A10 with half populated → COUNTA == 5.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(3, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(5, 0, Value::text(""));
  wb.sheet(0).set_cell_value(9, 0, Value::number(-2.0));
  const Value v = EvalSourceIn("=COUNTA(A1:A10)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsCountA, RangeWithErrorCellIsCounted) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_formula(1, 0, "=1/0");
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  const Value v = EvalSourceIn("=COUNTA(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsCountA, CrossSheetRange) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.sheet(1).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 0, Value::text("x"));
  // row 2 left blank
  wb.sheet(1).set_cell_value(3, 0, Value::boolean(false));
  const Value v = EvalSourceIn("=COUNTA(Sheet2!A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// COUNTBLANK - Blank scalars and Text "" only.
// ---------------------------------------------------------------------------

TEST(BuiltinsCountBlank, EmptyStringLiteralCounts) {
  // Excel pin: a formula returning "" is counted as blank by COUNTBLANK.
  const Value v = EvalSource("=COUNTBLANK(\"\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCountBlank, ZeroIsNotBlank) {
  const Value v = EvalSource("=COUNTBLANK(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCountBlank, FalseIsNotBlank) {
  const Value v = EvalSource("=COUNTBLANK(FALSE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCountBlank, NonEmptyTextIsNotBlank) {
  const Value v = EvalSource("=COUNTBLANK(\"hello\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCountBlank, ErrorIsNotBlank) {
  const Value v = EvalSource("=COUNTBLANK(#DIV/0!)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsCountBlank, MixedDirectArgsOnlyEmptyStringsCounted) {
  // Out of five args only the two "" count.
  const Value v = EvalSource("=COUNTBLANK(0, FALSE, \"\", \"x\", \"\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountBlank, BlankArgCounts) {
  const Value args[] = {Value::number(1.0), Value::blank(), Value::text(""),
                        Value::number(0.0), Value::blank()};
  const Value v = CallDirect("COUNTBLANK", args, 5u);
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsCountBlank, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=COUNTBLANK()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCountBlank, RangeHalfPopulated) {
  // A1..A10, five populated with non-blank content → five blanks.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("a"));
  wb.sheet(0).set_cell_value(3, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(5, 0, Value::number(0.0));
  wb.sheet(0).set_cell_value(9, 0, Value::number(-2.0));
  const Value v = EvalSourceIn("=COUNTBLANK(A1:A10)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsCountBlank, RangeIncludesEmptyStringCell) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text(""));
  // row 2 left blank
  const Value v = EvalSourceIn("=COUNTBLANK(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountBlank, RangeWithErrorCellSkipsError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_formula(0, 0, "=1/0");
  // row 1 blank
  wb.sheet(0).set_cell_value(2, 0, Value::text(""));
  const Value v = EvalSourceIn("=COUNTBLANK(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountBlank, CrossSheetRange) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  // Sheet2!A1 blank (not set), A2 = 5, A3 blank.
  wb.sheet(1).set_cell_value(1, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=COUNTBLANK(Sheet2!A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// Combined range-integration sanity check - one workbook, three calls.
// ---------------------------------------------------------------------------

TEST(BuiltinsCountCombined, OneRangeThreeCalls) {
  // A1..A10: 3 numbers, 2 text (one ""), 1 bool, 1 error, 3 blanks.
  //   COUNT      -> 3
  //   COUNTA     -> 7  (blanks only are skipped)
  //   COUNTBLANK -> 4  (3 true blanks + 1 empty-string text)
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("hello"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(2.5));
  wb.sheet(0).set_cell_value(3, 0, Value::text(""));
  wb.sheet(0).set_cell_value(4, 0, Value::boolean(true));
  // row 5 blank
  wb.sheet(0).set_cell_value(6, 0, Value::number(-3.0));
  wb.sheet(0).set_cell_formula(7, 0, "=1/0");
  // rows 8, 9 blank

  const Value c = EvalSourceIn("=COUNT(A1:A10)", wb, wb.sheet(0));
  ASSERT_TRUE(c.is_number());
  EXPECT_DOUBLE_EQ(c.as_number(), 3.0);

  const Value ca = EvalSourceIn("=COUNTA(A1:A10)", wb, wb.sheet(0));
  ASSERT_TRUE(ca.is_number());
  EXPECT_DOUBLE_EQ(ca.as_number(), 7.0);

  const Value cb = EvalSourceIn("=COUNTBLANK(A1:A10)", wb, wb.sheet(0));
  ASSERT_TRUE(cb.is_number());
  EXPECT_DOUBLE_EQ(cb.as_number(), 4.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
