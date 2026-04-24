// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the conditional aggregators COUNTIF, SUMIF, and
// AVERAGEIF. All three route through the lazy dispatch table in
// `tree_walker.cpp` because their first argument must be seen as AST so a
// single-cell `Ref` can be promoted to a 1-cell range, and because SUMIF /
// AVERAGEIF need a parallel second-range iteration shape that the eager
// `accepts_ranges` path does not provide.

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

// Parses `src` and evaluates it through the default function registry with
// no bound workbook. Each call resets the shared arenas to avoid
// cross-test contamination.
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
// Used by every range-integration test below.
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

// ---------------------------------------------------------------------------
// COUNTIF — numeric criteria
// ---------------------------------------------------------------------------

TEST(BuiltinsCountIf, RangeWithNumericEqCriterion) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(7.0));
  const Value v = EvalSourceIn("=COUNTIF(A1:A4, 5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountIf, RangeWithGtNumericCriterion) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(7.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(10.0));
  const Value v = EvalSourceIn("=COUNTIF(A1:A4, \">5\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountIf, RangeWithLtEqNumericCriterion) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(11.0));
  const Value v = EvalSourceIn("=COUNTIF(A1:A3, \"<=10\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountIf, RangeWithNotEqNumericCriterion) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(0.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(0.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(-1.0));
  const Value v = EvalSourceIn("=COUNTIF(A1:A4, \"<>0\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// COUNTIF — text criteria
// ---------------------------------------------------------------------------

TEST(BuiltinsCountIf, CaseInsensitiveTextEquality) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("Apple"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("APPLE"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("banana"));
  wb.sheet(0).set_cell_value(3, 0, Value::text("apple"));
  const Value v = EvalSourceIn("=COUNTIF(A1:A4, \"apple\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsCountIf, WildcardStarSuffix) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("Apple"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("Apricot"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("banana"));
  wb.sheet(0).set_cell_value(3, 0, Value::text("Ant"));
  const Value v = EvalSourceIn("=COUNTIF(A1:A4, \"A*\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsCountIf, WildcardQuestionInteriorMatch) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("bed"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("red"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("ed"));
  wb.sheet(0).set_cell_value(3, 0, Value::text("breed"));
  const Value v = EvalSourceIn("=COUNTIF(A1:A4, \"?ed\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountIf, EmptyStringCriterionCountsBlanks) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  // row 1: blank
  wb.sheet(0).set_cell_value(2, 0, Value::text(""));
  // row 3: blank
  wb.sheet(0).set_cell_value(4, 0, Value::text("y"));
  const Value v = EvalSourceIn("=COUNTIF(A1:A5, \"\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsCountIf, NotEqEmptyCountsNonBlank) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  // row 1: blank
  wb.sheet(0).set_cell_value(2, 0, Value::number(7.0));
  // row 3: blank
  wb.sheet(0).set_cell_value(4, 0, Value::text(""));
  // "<>" matches non-blank cells. The empty-string text cell ("") is NOT
  // blank — it's a concrete Text value — so it IS counted.
  const Value v = EvalSourceIn("=COUNTIF(A1:A5, \"<>\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

// ---------------------------------------------------------------------------
// COUNTIF — cross-sheet and single-cell ref
// ---------------------------------------------------------------------------

TEST(BuiltinsCountIf, CrossSheetRange) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Sheet2");
  wb.sheet(1).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 0, Value::number(10.0));
  wb.sheet(1).set_cell_value(2, 0, Value::number(100.0));
  const Value v = EvalSourceIn("=COUNTIF(Sheet2!A1:A3, \">=10\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountIf, SingleCellRefTreatedAsOneCellRange) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  const Value yes = EvalSourceIn("=COUNTIF(A1, \">0\")", wb, wb.sheet(0));
  ASSERT_TRUE(yes.is_number());
  EXPECT_DOUBLE_EQ(yes.as_number(), 1.0);
  const Value no = EvalSourceIn("=COUNTIF(A1, \"<0\")", wb, wb.sheet(0));
  ASSERT_TRUE(no.is_number());
  EXPECT_DOUBLE_EQ(no.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// COUNTIF — error / arity
// ---------------------------------------------------------------------------

TEST(BuiltinsCountIf, ScalarFirstArgIsValueError) {
  const Value v = EvalSource("=COUNTIF(1, \">0\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCountIf, CriterionIsErrorFiltersMatchingErrorCells) {
  // An error criterion is not propagated; it counts cells with the same
  // error code. Matches Excel 365 behaviour verified against IronCalc
  // golden COUNTIF_Columns G20 (#N/A) and H20 (#DIV/0!).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_formula(0, 0, "=1/0");
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_formula(2, 0, "=1/0");
  wb.sheet(0).set_cell_formula(3, 0, "=NA()");
  const Value v = EvalSourceIn("=COUNTIF(A1:A4, #DIV/0!)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountIf, ErrorCellInRangeIsSkipped) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_formula(1, 0, "=1/0");
  wb.sheet(0).set_cell_value(2, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=COUNTIF(A1:A3, \">0\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountIf, Arity1IsValueError) {
  const Value v = EvalSource("=COUNTIF(1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCountIf, Arity3IsValueError) {
  const Value v = EvalSource("=COUNTIF(1, 2, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// SUMIF
// ---------------------------------------------------------------------------

TEST(BuiltinsSumIf, SumOverCriteriaRangeWithoutSumRange) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(6.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(8.0));
  const Value v = EvalSourceIn("=SUMIF(A1:A4, \">5\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 14.0);
}

TEST(BuiltinsSumIf, ExplicitSumRangeSameShape) {
  Workbook wb = Workbook::create();
  // Criteria column: A1..A4 = {low, low, high, high}
  wb.sheet(0).set_cell_value(0, 0, Value::text("low"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("low"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("high"));
  wb.sheet(0).set_cell_value(3, 0, Value::text("high"));
  // Sum column: B1..B4 = {10, 20, 30, 40}
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 1, Value::number(40.0));
  const Value v =
      EvalSourceIn("=SUMIF(A1:A4, \"high\", B1:B4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 70.0);
}

TEST(BuiltinsSumIf, SumRangeSmallerClampsToMin) {
  // Accepted divergence: Excel reshapes the sum range to match the
  // criteria range's anchor. We clamp to min(size) instead.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(7.0));
  // B3 left blank — only two of three matches have a numeric sum cell.
  const Value v =
      EvalSourceIn("=SUMIF(A1:A3, \"x\", B1:B2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 12.0);
}

TEST(BuiltinsSumIf, NonNumericMatchingCellsAreExcluded) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("50"));   // skipped
  wb.sheet(0).set_cell_value(2, 1, Value::boolean(true));  // skipped
  const Value v =
      EvalSourceIn("=SUMIF(A1:A3, \"x\", B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsSumIf, CrossSheetQualifiedRanges) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  wb.sheet(1).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 0, Value::number(10.0));
  wb.sheet(1).set_cell_value(2, 0, Value::number(100.0));
  wb.sheet(1).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(1).set_cell_value(2, 1, Value::number(3.0));
  const Value v =
      EvalSourceIn("=SUMIF(Data!A1:A3, \">=10\", Data!B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsSumIf, NoMatchesReturnsZero) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  const Value v = EvalSourceIn("=SUMIF(A1:A2, \">999\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsSumIf, ArityViolations) {
  EXPECT_EQ(EvalSource("=SUMIF(1)").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=SUMIF(1, 2, 3, 4)").as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// AVERAGEIF
// ---------------------------------------------------------------------------

TEST(BuiltinsAverageIf, AverageOverMatchingNumerics) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(20.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(30.0));
  const Value v = EvalSourceIn("=AVERAGEIF(A1:A4, \">=10\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);  // (10+20+30)/3
}

TEST(BuiltinsAverageIf, NoMatchesReturnsDiv0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  const Value v =
      EvalSourceIn("=AVERAGEIF(A1:A2, \">999\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsAverageIf, TextOrBoolMatchesExcludedFromPool) {
  // Three cells match ("x") but only two are numeric on the average side;
  // the text-typed sum cell is excluded from both numerator and
  // denominator so the result is (10+30)/2 = 20, not (10+30)/3.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("20"));   // excluded
  wb.sheet(0).set_cell_value(2, 1, Value::number(30.0));
  const Value v =
      EvalSourceIn("=AVERAGEIF(A1:A3, \"x\", B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);
}

TEST(BuiltinsAverageIf, ExplicitAverageRange) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("cat"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("dog"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("cat"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(100.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(8.0));
  const Value v =
      EvalSourceIn("=AVERAGEIF(A1:A3, \"cat\", B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 5.0);
}

TEST(BuiltinsAverageIf, ArityViolations) {
  EXPECT_EQ(EvalSource("=AVERAGEIF(1)").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=AVERAGEIF(1, 2, 3, 4)").as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Dispatcher pins
// ---------------------------------------------------------------------------

TEST(BuiltinsCountIfDispatch, AllThreeGoThroughLazyPath) {
  // These names must NOT be looked up through the eager function
  // registry — if they were, the dispatcher would pre-evaluate arg 0 as a
  // scalar (bypassing the AST-shape check) and a Ref argument would lose
  // its range semantics. The registry should therefore NOT have entries
  // for them.
  EXPECT_EQ(default_registry().lookup("COUNTIF"), nullptr);
  EXPECT_EQ(default_registry().lookup("SUMIF"), nullptr);
  EXPECT_EQ(default_registry().lookup("AVERAGEIF"), nullptr);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
