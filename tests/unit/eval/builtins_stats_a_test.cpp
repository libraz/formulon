// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the statistical "A"-family aggregators:
// AVERAGEA / MAXA / MINA / VARA / VARPA / STDEVA / STDEVPA.
//
// These differ from AVERAGE / MAX / MIN / VAR / STDEV in how non-number
// values arriving from a range are treated: Text cells evaluate to 0,
// Bool cells to 0 / 1, and Blank cells are still skipped. Direct scalar
// arguments follow the same rules inside `collect_a` for symmetry, so
// `=AVERAGEA(1, TRUE, 3)` counts TRUE as 1.
//
// The registered flag `range_filter_a_coerce` is what drives the
// dispatcher-level Bool/Text/Blank transform on range-sourced cells; the
// AllRangeAwareAndPropagating pin below guards against accidental drift.

#include <cmath>
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

// Parses `src` and evaluates via the default registry (no workbook).
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

// Parses `src` and evaluates against a bound workbook + current sheet.
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
// Registry pins
// ---------------------------------------------------------------------------

TEST(BuiltinsStatsARegistry, AllNamesRegistered) {
  for (const char* name : {"AVERAGEA", "MAXA", "MINA", "VARA", "VARPA", "STDEVA", "STDEVPA", "VARP", "STDEVP"}) {
    EXPECT_NE(default_registry().lookup(name), nullptr) << "missing: " << name;
  }
}

TEST(BuiltinsStatsARegistry, ACoerceFlagOnAVariantsOnly) {
  // A-variants must carry the range_filter_a_coerce flag so the dispatcher
  // does the Bool->0/1 / Text->0 / Blank-skip transform on range cells.
  for (const char* name : {"AVERAGEA", "MAXA", "MINA", "VARA", "VARPA", "STDEVA", "STDEVPA"}) {
    const FunctionDef* def = default_registry().lookup(name);
    ASSERT_NE(def, nullptr) << name;
    EXPECT_TRUE(def->accepts_ranges) << name;
    EXPECT_TRUE(def->propagate_errors) << name;
    EXPECT_TRUE(def->range_filter_a_coerce) << name;
    // Must not double-up with the numeric-only or bool-coercible filters.
    EXPECT_FALSE(def->range_filter_numeric_only) << name;
    EXPECT_FALSE(def->range_filter_bool_coercible) << name;
  }
  // Non-A legacy aliases keep the numeric-only filter (they are VAR.P /
  // STDEV.P under the old spelling, not A-variants).
  for (const char* name : {"VARP", "STDEVP"}) {
    const FunctionDef* def = default_registry().lookup(name);
    ASSERT_NE(def, nullptr) << name;
    EXPECT_FALSE(def->range_filter_a_coerce) << name;
  }
}

// ---------------------------------------------------------------------------
// AVERAGEA
// ---------------------------------------------------------------------------

TEST(BuiltinsAverageA, NumericOnlyMatchesAverage) {
  const Value v = EvalSource("=AVERAGEA(1,2,3,4,5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsAverageA, DirectTrueCountsAsOne) {
  const Value v = EvalSource("=AVERAGEA(1, TRUE, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), (1.0 + 1.0 + 3.0) / 3.0);
}

TEST(BuiltinsAverageA, DirectFalseCountsAsZero) {
  const Value v = EvalSource("=AVERAGEA(1, FALSE, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), (1.0 + 0.0 + 5.0) / 3.0);
}

TEST(BuiltinsAverageA, RangeTextCountsAsZero) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("text"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  const Value v = EvalSourceIn("=AVERAGEA(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), (1.0 + 0.0 + 3.0) / 3.0);
}

TEST(BuiltinsAverageA, RangeBoolCountsAsZeroOne) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(2, 0, Value::boolean(false));
  const Value v = EvalSourceIn("=AVERAGEA(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), (1.0 + 1.0 + 0.0) / 3.0);
}

TEST(BuiltinsAverageA, RangeBlankIsSkipped) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  // A2 left blank on purpose — must not count toward the denominator.
  wb.sheet(0).set_cell_value(2, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=AVERAGEA(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), (1.0 + 5.0) / 2.0);
}

// ---------------------------------------------------------------------------
// MAXA / MINA
// ---------------------------------------------------------------------------

TEST(BuiltinsMaxA, RangeTextCountsAsZero) {
  // Values -5, "big", -2 -> {-5, 0, -2} -> max = 0.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(-5.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("big"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(-2.0));
  const Value v = EvalSourceIn("=MAXA(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMaxA, RangeTrueBeatsHalf) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(0.5));
  wb.sheet(0).set_cell_value(1, 0, Value::boolean(true));
  const Value v = EvalSourceIn("=MAXA(A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMinA, RangeTextCountsAsZero) {
  // Values 5, "small", 3 -> {5, 0, 3} -> min = 0.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("small"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(3.0));
  const Value v = EvalSourceIn("=MINA(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMinA, EmptyArityIsZero) {
  // All blanks in the range -> MINA returns 0 like MIN.
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=MINA(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// VARA / VARPA / STDEVA / STDEVPA
// ---------------------------------------------------------------------------

TEST(BuiltinsVarA, SampleVarianceWithBoolAndText) {
  // {1, TRUE, "x", 4} -> values [1, 1, 0, 4], n=4, mean=1.5.
  // SS = 0.25 + 0.25 + 2.25 + 6.25 = 9, sample var = 9 / 3 = 3.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 0, Value::number(4.0));
  const Value v = EvalSourceIn("=VARA(A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsVarA, SampleRequiresAtLeastTwo) {
  // Only one non-blank value even under A-coercion -> #DIV/0!.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(5.0));
  const Value v = EvalSourceIn("=VARA(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsVarPA, PopulationVarianceWithBoolAndText) {
  // Same data as VARA: mean=1.5, SS=9, pop var = 9 / 4 = 2.25.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 0, Value::number(4.0));
  const Value v = EvalSourceIn("=VARPA(A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2.25);
}

TEST(BuiltinsStdevA, SqrtOfVara) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 0, Value::number(4.0));
  const Value v = EvalSourceIn("=STDEVA(A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), std::sqrt(3.0));
}

TEST(BuiltinsStdevPA, SqrtOfVarpa) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(2, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(3, 0, Value::number(4.0));
  const Value v = EvalSourceIn("=STDEVPA(A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), std::sqrt(2.25));
}

// ---------------------------------------------------------------------------
// Legacy VARP / STDEVP (aliases of VAR.P / STDEV.P, NOT A-variants).
// ---------------------------------------------------------------------------

TEST(BuiltinsLegacyVarP, MatchesVarP) {
  const Value a = EvalSource("=VARP(1,2,3,4,5)");
  const Value b = EvalSource("=VAR.P(1,2,3,4,5)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsLegacyStdevP, MatchesStdevP) {
  const Value a = EvalSource("=STDEVP(1,2,3,4,5)");
  const Value b = EvalSource("=STDEV.P(1,2,3,4,5)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsLegacyVarP, SkipsRangeText) {
  // VARP is NOT an A-variant: range-sourced text is skipped, not 0.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("x"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(5.0));
  const Value legacy = EvalSourceIn("=VARP(A1:A3)", wb, wb.sheet(0));
  const Value modern = EvalSourceIn("=VAR.P(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(legacy.is_number());
  ASSERT_TRUE(modern.is_number());
  EXPECT_DOUBLE_EQ(legacy.as_number(), modern.as_number());
}

}  // namespace
}  // namespace eval
}  // namespace formulon
