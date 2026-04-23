// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the multi-criteria aggregators COUNTIFS, SUMIFS,
// AVERAGEIFS, MAXIFS, and MINIFS. All five route through the lazy
// dispatch table in `tree_walker.cpp` via the shared
// `resolve_criteria_pairs` helper: arg 0 must be seen as AST so a single-
// cell Ref is a 1-cell range, and the parallel criteria / result ranges
// are iterated in lockstep rather than flattened.

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
// no bound workbook. Each call resets the shared arenas.
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
// COUNTIFS
// ---------------------------------------------------------------------------

TEST(BuiltinsCountIfs, SinglePairMatchesCountIf) {
  // With a single (range, crit) pair COUNTIFS is numerically equivalent to
  // COUNTIF; pin that here so later refactors can't drift them apart.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(7.0));
  const Value ifs = EvalSourceIn("=COUNTIFS(A1:A4, 5)", wb, wb.sheet(0));
  const Value cif = EvalSourceIn("=COUNTIF(A1:A4, 5)", wb, wb.sheet(0));
  ASSERT_TRUE(ifs.is_number());
  ASSERT_TRUE(cif.is_number());
  EXPECT_DOUBLE_EQ(ifs.as_number(), cif.as_number());
  EXPECT_DOUBLE_EQ(ifs.as_number(), 2.0);
}

TEST(BuiltinsCountIfs, TwoPairsAndLogicAcrossParallelRanges) {
  // Column A: category, Column B: score. We want rows where category="x"
  // AND score > 5. Matches are (x, 7) and (x, 10), so count = 2.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(3.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(9.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(7.0));
  wb.sheet(0).set_cell_value(3, 1, Value::number(10.0));
  const Value v = EvalSourceIn("=COUNTIFS(A1:A4, \"x\", B1:B4, \">5\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountIfs, WildcardAndComparisonInSameCall) {
  // A="A*" (wildcard match) AND B>=10. Rows 1 and 3 qualify.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("Apple"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("Ant"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("Banana"));
  wb.sheet(0).set_cell_value(3, 0, Value::text("Apricot"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(3, 1, Value::number(30.0));
  const Value v = EvalSourceIn("=COUNTIFS(A1:A4, \"A*\", B1:B4, \">=10\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsCountIfs, ThreePairsAllMustMatch) {
  // Three criteria. Position 0: (x, 1, a) – matches. Position 1: (x, 2, b)
  // – fails last. Position 2: (y, 1, a) – fails first. So count = 1.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 2, Value::text("a"));
  wb.sheet(0).set_cell_value(1, 2, Value::text("b"));
  wb.sheet(0).set_cell_value(2, 2, Value::text("a"));
  const Value v = EvalSourceIn("=COUNTIFS(A1:A3, \"x\", B1:B3, 1, C1:C3, \"a\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCountIfs, RangeSizeMismatchIsValueError) {
  // A1:A3 has 3 cells, B1:B2 has 2 cells -> #VALUE! (Excel 365 strict).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  const Value v = EvalSourceIn("=COUNTIFS(A1:A3, \"x\", B1:B2, \">0\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCountIfs, OddArityIsValueError) {
  // Must be an even number of args (>= 2). Three args is invalid.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  const Value v = EvalSourceIn("=COUNTIFS(A1, \">0\", A1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCountIfs, ZeroArityIsValueError) {
  EXPECT_EQ(EvalSource("=COUNTIFS()").as_error(), ErrorCode::Value);
}

TEST(BuiltinsCountIfs, CriterionErrorPropagates) {
  // An error-valued criterion propagates, even on the second pair.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  const Value v = EvalSourceIn("=COUNTIFS(A1:A2, \">0\", B1:B2, #DIV/0!)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsCountIfs, CrossSheetRangesBothQualified) {
  Workbook wb = Workbook::create();
  wb.add_sheet("Data");
  wb.sheet(1).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(1).set_cell_value(1, 0, Value::number(10.0));
  wb.sheet(1).set_cell_value(2, 0, Value::number(100.0));
  wb.sheet(1).set_cell_value(0, 1, Value::text("a"));
  wb.sheet(1).set_cell_value(1, 1, Value::text("b"));
  wb.sheet(1).set_cell_value(2, 1, Value::text("a"));
  // a on rows 1, 3; >=10 on rows 2, 3 => row 3 qualifies, count = 1.
  const Value v = EvalSourceIn("=COUNTIFS(Data!A1:A3, \">=10\", Data!B1:B3, \"a\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsCountIfs, SingleCellRefActsAsOneCellRange) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(42.0));
  wb.sheet(0).set_cell_value(0, 1, Value::text("hit"));
  const Value v = EvalSourceIn("=COUNTIFS(A1, \">0\", B1, \"hit\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
  const Value no = EvalSourceIn("=COUNTIFS(A1, \">0\", B1, \"miss\")", wb, wb.sheet(0));
  ASSERT_TRUE(no.is_number());
  EXPECT_DOUBLE_EQ(no.as_number(), 0.0);
}

TEST(BuiltinsCountIfs, ScalarFirstArgIsValueError) {
  const Value v = EvalSource("=COUNTIFS(1, \">0\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsCountIfs, ErrorCellInCriteriaRangeIsSkipped) {
  // An error cell in a criteria range fails the match silently — it does
  // NOT propagate. Expect count of numeric rows matching >0.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_formula(1, 0, "=1/0");
  wb.sheet(0).set_cell_value(2, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(10.0));
  const Value v = EvalSourceIn("=COUNTIFS(A1:A3, \">0\", B1:B3, 10)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.0);
}

// ---------------------------------------------------------------------------
// SUMIFS
// ---------------------------------------------------------------------------

TEST(BuiltinsSumIfs, SinglePairAgreesWithSumIf) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(100.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(30.0));
  const Value ifs = EvalSourceIn("=SUMIFS(B1:B3, A1:A3, \"x\")", wb, wb.sheet(0));
  const Value sif = EvalSourceIn("=SUMIF(A1:A3, \"x\", B1:B3)", wb, wb.sheet(0));
  ASSERT_TRUE(ifs.is_number());
  ASSERT_TRUE(sif.is_number());
  EXPECT_DOUBLE_EQ(ifs.as_number(), sif.as_number());
  EXPECT_DOUBLE_EQ(ifs.as_number(), 40.0);
}

TEST(BuiltinsSumIfs, TwoPairsAndLogic) {
  // Sum sales where region="West" AND product="Apple".
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("West"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("East"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("West"));
  wb.sheet(0).set_cell_value(3, 0, Value::text("West"));
  wb.sheet(0).set_cell_value(0, 1, Value::text("Apple"));
  wb.sheet(0).set_cell_value(1, 1, Value::text("Apple"));
  wb.sheet(0).set_cell_value(2, 1, Value::text("Apple"));
  wb.sheet(0).set_cell_value(3, 1, Value::text("Banana"));
  wb.sheet(0).set_cell_value(0, 2, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 2, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 2, Value::number(40.0));
  const Value v = EvalSourceIn("=SUMIFS(C1:C4, A1:A4, \"West\", B1:B4, \"Apple\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 40.0);  // 10 + 30
}

TEST(BuiltinsSumIfs, SumRangeShapeMismatchIsValueError) {
  // sum_range (2 cells) doesn't match criteria range (3 cells).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  const Value v = EvalSourceIn("=SUMIFS(B1:B2, A1:A3, \"x\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsSumIfs, NonNumericSumCellsExcluded) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("50"));     // skipped
  wb.sheet(0).set_cell_value(2, 1, Value::boolean(true));  // skipped
  const Value v = EvalSourceIn("=SUMIFS(B1:B3, A1:A3, \"x\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsSumIfs, NoMatchesReturnsZero) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  const Value v = EvalSourceIn("=SUMIFS(B1:B2, A1:A2, \">999\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsSumIfs, ArityViolations) {
  // Arity < 3 is always rejected.
  EXPECT_EQ(EvalSource("=SUMIFS()").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=SUMIFS(1)").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=SUMIFS(1, 2)").as_error(), ErrorCode::Value);
  // Even arity is rejected (needs sum + N pairs = odd >= 3).
  EXPECT_EQ(EvalSource("=SUMIFS(1, 2, 3, 4)").as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// AVERAGEIFS
// ---------------------------------------------------------------------------

TEST(BuiltinsAverageIfs, MeanOfMatchingNumerics) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(100.0));  // excluded (y)
  wb.sheet(0).set_cell_value(3, 1, Value::number(30.0));
  const Value v = EvalSourceIn("=AVERAGEIFS(B1:B4, A1:A4, \"x\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);  // (10+20+30)/3
}

TEST(BuiltinsAverageIfs, NoMatchesReturnsDiv0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  const Value v = EvalSourceIn("=AVERAGEIFS(B1:B2, A1:A2, \">999\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsAverageIfs, NonNumericAvgCellsExcludedFromPool) {
  // All three rows match on criterion A="x", but only two have numeric
  // averaging cells. Expect mean over those two, NOT two numerics averaged
  // over three denominators.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("20"));  // excluded
  wb.sheet(0).set_cell_value(2, 1, Value::number(30.0));
  const Value v = EvalSourceIn("=AVERAGEIFS(B1:B3, A1:A3, \"x\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 20.0);  // (10+30)/2
}

TEST(BuiltinsAverageIfs, TwoPairsFilter) {
  // Mean of C where A>=2 AND B<=20. Rows 2 (A=2, B=20, C=50) and 3 (A=3,
  // B=10, C=70) qualify; mean = 60.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(4.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(3, 1, Value::number(30.0));
  wb.sheet(0).set_cell_value(0, 2, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(50.0));
  wb.sheet(0).set_cell_value(2, 2, Value::number(70.0));
  wb.sheet(0).set_cell_value(3, 2, Value::number(90.0));
  const Value v = EvalSourceIn("=AVERAGEIFS(C1:C4, A1:A4, \">=2\", B1:B4, \"<=20\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 60.0);
}

TEST(BuiltinsAverageIfs, ArityViolations) {
  EXPECT_EQ(EvalSource("=AVERAGEIFS()").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=AVERAGEIFS(1)").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=AVERAGEIFS(1, 2)").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=AVERAGEIFS(1, 2, 3, 4)").as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// MAXIFS
// ---------------------------------------------------------------------------

TEST(BuiltinsMaxIfs, ReturnsMaxOverMatchingNumerics) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(999.0));  // excluded (y)
  wb.sheet(0).set_cell_value(2, 1, Value::number(50.0));
  wb.sheet(0).set_cell_value(3, 1, Value::number(30.0));
  const Value v = EvalSourceIn("=MAXIFS(B1:B4, A1:A4, \"x\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 50.0);
}

TEST(BuiltinsMaxIfs, NonNumericMatchesIgnored) {
  // Three rows match "x"; one has a text cell in max_range that is
  // excluded. Remaining numerics are {10, 30}.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("999"));  // excluded
  wb.sheet(0).set_cell_value(2, 1, Value::number(30.0));
  const Value v = EvalSourceIn("=MAXIFS(B1:B3, A1:A3, \"x\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30.0);
}

TEST(BuiltinsMaxIfs, NoMatchesReturnsZero) {
  // Excel quirk: empty pool -> 0, not #NUM!.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(2.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(20.0));
  const Value v = EvalSourceIn("=MAXIFS(B1:B2, A1:A2, \">999\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMaxIfs, NegativeOnlyPoolStillPicksLargest) {
  // Regression guard: don't let a 0.0 sentinel beat real (negative) data.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(-10.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(-3.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(-7.0));
  const Value v = EvalSourceIn("=MAXIFS(B1:B3, A1:A3, \"x\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -3.0);
}

TEST(BuiltinsMaxIfs, ArityViolations) {
  EXPECT_EQ(EvalSource("=MAXIFS()").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=MAXIFS(1, 2)").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=MAXIFS(1, 2, 3, 4)").as_error(), ErrorCode::Value);
}

TEST(BuiltinsMaxIfs, ShapeMismatchIsValueError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(2.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(3.0));
  const Value v = EvalSourceIn("=MAXIFS(B1:B3, A1:A2, \"x\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// MINIFS
// ---------------------------------------------------------------------------

TEST(BuiltinsMinIfs, ReturnsMinOverMatchingNumerics) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(50.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(-999.0));  // excluded (y)
  wb.sheet(0).set_cell_value(2, 1, Value::number(10.0));
  wb.sheet(0).set_cell_value(3, 1, Value::number(30.0));
  const Value v = EvalSourceIn("=MINIFS(B1:B4, A1:A4, \"x\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 10.0);
}

TEST(BuiltinsMinIfs, NonNumericMatchesIgnored) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 1, Value::text("1"));  // excluded
  wb.sheet(0).set_cell_value(2, 1, Value::number(50.0));
  const Value v = EvalSourceIn("=MINIFS(B1:B3, A1:A3, \"x\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 50.0);
}

TEST(BuiltinsMinIfs, NoMatchesReturnsZero) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(0, 1, Value::number(10.0));
  const Value v = EvalSourceIn("=MINIFS(B1:B1, A1:A1, \">999\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMinIfs, PositiveOnlyPoolStillPicksSmallest) {
  // Regression guard: don't let a 0.0 sentinel sneak in below real data.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(0, 1, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 1, Value::number(3.0));
  wb.sheet(0).set_cell_value(2, 1, Value::number(7.0));
  const Value v = EvalSourceIn("=MINIFS(B1:B3, A1:A3, \"x\")", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsMinIfs, ArityViolations) {
  EXPECT_EQ(EvalSource("=MINIFS()").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=MINIFS(1, 2)").as_error(), ErrorCode::Value);
  EXPECT_EQ(EvalSource("=MINIFS(1, 2, 3, 4)").as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Shared smoke — five aggregators, one dataset, cross-validation
// ---------------------------------------------------------------------------

TEST(BuiltinsIfsSmoke, AllFiveAgreeOnSharedDataset) {
  // Shape the data so all five aggregators have something to bite on.
  // A (category): x, y, x, x, y
  // B (flag):    T,  F, T, F, T
  // C (value):   10, 20, 30, 40, 50
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(1, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(4, 0, Value::text("y"));
  wb.sheet(0).set_cell_value(0, 1, Value::boolean(true));
  wb.sheet(0).set_cell_value(1, 1, Value::boolean(false));
  wb.sheet(0).set_cell_value(2, 1, Value::boolean(true));
  wb.sheet(0).set_cell_value(3, 1, Value::boolean(false));
  wb.sheet(0).set_cell_value(4, 1, Value::boolean(true));
  wb.sheet(0).set_cell_value(0, 2, Value::number(10.0));
  wb.sheet(0).set_cell_value(1, 2, Value::number(20.0));
  wb.sheet(0).set_cell_value(2, 2, Value::number(30.0));
  wb.sheet(0).set_cell_value(3, 2, Value::number(40.0));
  wb.sheet(0).set_cell_value(4, 2, Value::number(50.0));

  // Two-pair filter: category="x" AND flag=TRUE. Rows 1, 3 qualify.
  // Matching C values: {10, 30}.
  const Value cnt = EvalSourceIn("=COUNTIFS(A1:A5, \"x\", B1:B5, TRUE)", wb, wb.sheet(0));
  const Value sum = EvalSourceIn("=SUMIFS(C1:C5, A1:A5, \"x\", B1:B5, TRUE)", wb, wb.sheet(0));
  const Value avg = EvalSourceIn("=AVERAGEIFS(C1:C5, A1:A5, \"x\", B1:B5, TRUE)", wb, wb.sheet(0));
  const Value mx = EvalSourceIn("=MAXIFS(C1:C5, A1:A5, \"x\", B1:B5, TRUE)", wb, wb.sheet(0));
  const Value mn = EvalSourceIn("=MINIFS(C1:C5, A1:A5, \"x\", B1:B5, TRUE)", wb, wb.sheet(0));

  ASSERT_TRUE(cnt.is_number());
  ASSERT_TRUE(sum.is_number());
  ASSERT_TRUE(avg.is_number());
  ASSERT_TRUE(mx.is_number());
  ASSERT_TRUE(mn.is_number());
  EXPECT_DOUBLE_EQ(cnt.as_number(), 2.0);
  EXPECT_DOUBLE_EQ(sum.as_number(), 40.0);
  EXPECT_DOUBLE_EQ(avg.as_number(), 20.0);
  EXPECT_DOUBLE_EQ(mx.as_number(), 30.0);
  EXPECT_DOUBLE_EQ(mn.as_number(), 10.0);
  // SUM / COUNT must equal AVERAGE.
  EXPECT_DOUBLE_EQ(sum.as_number() / cnt.as_number(), avg.as_number());
}

// ---------------------------------------------------------------------------
// Dispatcher pins
// ---------------------------------------------------------------------------

TEST(BuiltinsIfsDispatch, AllFiveGoThroughLazyPath) {
  // These names must not live in the eager registry; if they did, the
  // dispatcher would pre-evaluate the result range as a scalar and drop
  // range semantics.
  EXPECT_EQ(default_registry().lookup("COUNTIFS"), nullptr);
  EXPECT_EQ(default_registry().lookup("SUMIFS"), nullptr);
  EXPECT_EQ(default_registry().lookup("AVERAGEIFS"), nullptr);
  EXPECT_EQ(default_registry().lookup("MAXIFS"), nullptr);
  EXPECT_EQ(default_registry().lookup("MINIFS"), nullptr);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
