// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for XIRR and XNPV. Both ride the lazy dispatch seam in
// `src/eval/financial_lazy.cpp` because each needs the un-flattened AST
// of two parallel range arguments (`values` alongside `dates`) plus, for
// XIRR, a trailing scalar `guess`. The harness here mirrors the pattern
// used by the MIRR range tests in `builtins_financial2_test.cpp`: a
// workbook is populated with literal cells and the formula is evaluated
// through the default registry + bound EvalContext.

#include <cmath>
#include <iomanip>
#include <sstream>
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

// Evaluates `src` through the default registry with no bound workbook.
// Used for scalar-only cases (inline ArrayLiteral args, arity checks).
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

// Evaluates `src` against a workbook + current sheet.
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

// Builds the Microsoft XIRR / XNPV documentation fixture:
//   values = {-10000, 2750, 4250, 3250, 2750}
//   dates  = {1/1/2008, 3/1/2008, 10/30/2008, 2/15/2009, 4/1/2009}
// Excel 1900 serials for those dates (verified against Mac Excel):
//   39448, 39508, 39751, 39859, 39904
// Column A holds values, column B holds date serials.
Workbook MakeMicrosoftFixture() {
  Workbook wb = Workbook::create();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(-10000.0));
  s.set_cell_value(1, 0, Value::number(2750.0));
  s.set_cell_value(2, 0, Value::number(4250.0));
  s.set_cell_value(3, 0, Value::number(3250.0));
  s.set_cell_value(4, 0, Value::number(2750.0));
  s.set_cell_value(0, 1, Value::number(39448.0));
  s.set_cell_value(1, 1, Value::number(39508.0));
  s.set_cell_value(2, 1, Value::number(39751.0));
  s.set_cell_value(3, 1, Value::number(39859.0));
  s.set_cell_value(4, 1, Value::number(39904.0));
  return wb;
}

// ---------------------------------------------------------------------------
// XNPV
// ---------------------------------------------------------------------------

TEST(FinancialXnpv, MicrosoftReference) {
  // Microsoft docs: XNPV(0.09, values, dates) ~= 2086.647602.
  Workbook wb = MakeMicrosoftFixture();
  const Value v = EvalSourceIn("=XNPV(0.09, A1:A5, B1:B5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 2086.647602, 1e-6);
}

TEST(FinancialXnpv, ZeroRateIsNum) {
  // Mathematically rate = 0 collapses every term to v[i] and yields the
  // undiscounted schedule total, but Mac Excel 365 (16.108.1, ja-JP)
  // treats rate == 0 as an input error and surfaces #NUM! before any
  // math runs. 1-bit parity requires the same behaviour.
  Workbook wb = MakeMicrosoftFixture();
  const Value v = EvalSourceIn("=XNPV(0, A1:A5, B1:B5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialXnpv, RateMinusOneIsNum) {
  // rate = -1 drives (1+rate) to zero; division would blow up. Excel
  // rejects this with #NUM! before any math runs.
  Workbook wb = MakeMicrosoftFixture();
  const Value v = EvalSourceIn("=XNPV(-1, A1:A5, B1:B5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialXnpv, RateBelowMinusOneIsNum) {
  Workbook wb = MakeMicrosoftFixture();
  const Value v = EvalSourceIn("=XNPV(-1.5, A1:A5, B1:B5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialXnpv, MismatchedRangeSizesIsNum) {
  Workbook wb = MakeMicrosoftFixture();
  // 5 values vs 4 dates -> #NUM!.
  const Value v = EvalSourceIn("=XNPV(0.1, A1:A5, B1:B4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialXnpv, ArgumentErrorPropagates) {
  Workbook wb = MakeMicrosoftFixture();
  // The rate arg is `1/0` which is #DIV/0!; that error must win.
  const Value v = EvalSourceIn("=XNPV(1/0, A1:A5, B1:B5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// XIRR
// ---------------------------------------------------------------------------

TEST(FinancialXirr, MicrosoftReference) {
  // Microsoft docs: XIRR(values, dates, 0.1) ~= 0.373362535.
  Workbook wb = MakeMicrosoftFixture();
  const Value v = EvalSourceIn("=XIRR(A1:A5, B1:B5, 0.1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 0.373362535, 1e-6);
}

TEST(FinancialXirr, DefaultGuessMatchesExplicitTenPercent) {
  // Arity-2 XIRR must produce the same rate as arity-3 with guess=0.1.
  Workbook wb = MakeMicrosoftFixture();
  const Value implicit = EvalSourceIn("=XIRR(A1:A5, B1:B5)", wb, wb.sheet(0));
  const Value explicit_guess = EvalSourceIn("=XIRR(A1:A5, B1:B5, 0.1)", wb, wb.sheet(0));
  ASSERT_TRUE(implicit.is_number());
  ASSERT_TRUE(explicit_guess.is_number());
  EXPECT_NEAR(implicit.as_number(), explicit_guess.as_number(), 1e-9);
}

TEST(FinancialXirr, XirrIsRootOfXnpv) {
  // XIRR is by definition the rate r such that XNPV(r, values, dates)==0.
  // Plug the XIRR result back into XNPV and verify the residual collapses.
  // `std::to_string` truncates to ~6 significant digits; Newton converges
  // to |f(r)| < 1e-10 internally, so we need full double precision in the
  // re-emitted formula to observe the residual shrink.
  Workbook wb = MakeMicrosoftFixture();
  const Value r = EvalSourceIn("=XIRR(A1:A5, B1:B5)", wb, wb.sheet(0));
  ASSERT_TRUE(r.is_number());
  std::ostringstream oss;
  oss << std::setprecision(17) << r.as_number();
  const std::string formula = "=XNPV(" + oss.str() + ", A1:A5, B1:B5)";
  const Value residual = EvalSourceIn(formula, wb, wb.sheet(0));
  ASSERT_TRUE(residual.is_number());
  EXPECT_LT(std::fabs(residual.as_number()), 1e-5);
}

TEST(FinancialXirr, AllPositiveIsNum) {
  Workbook wb = Workbook::create();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(100.0));
  s.set_cell_value(1, 0, Value::number(200.0));
  s.set_cell_value(0, 1, Value::number(39448.0));
  s.set_cell_value(1, 1, Value::number(39508.0));
  const Value v = EvalSourceIn("=XIRR(A1:A2, B1:B2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialXirr, AllNegativeIsNum) {
  Workbook wb = Workbook::create();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(-100.0));
  s.set_cell_value(1, 0, Value::number(-200.0));
  s.set_cell_value(0, 1, Value::number(39448.0));
  s.set_cell_value(1, 1, Value::number(39508.0));
  const Value v = EvalSourceIn("=XIRR(A1:A2, B1:B2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialXirr, SingleCashFlowIsNum) {
  // Fewer than two surviving cash flows: Excel rejects with #NUM!.
  Workbook wb = Workbook::create();
  auto& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(-100.0));
  s.set_cell_value(0, 1, Value::number(39448.0));
  const Value v = EvalSourceIn("=XIRR(A1:A1, B1:B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialXirr, MismatchedRangeSizesIsNum) {
  Workbook wb = MakeMicrosoftFixture();
  const Value v = EvalSourceIn("=XIRR(A1:A5, B1:B4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialXirr, GuessAtMinusOneIsNum) {
  Workbook wb = MakeMicrosoftFixture();
  const Value v = EvalSourceIn("=XIRR(A1:A5, B1:B5, -1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialXirr, GuessBelowMinusOneIsNum) {
  Workbook wb = MakeMicrosoftFixture();
  const Value v = EvalSourceIn("=XIRR(A1:A5, B1:B5, -1.5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialXirr, NonRangeValuesIsValue) {
  // The lazy dispatcher requires the values argument to resolve to a
  // range, a bare Ref, or an inline ArrayLiteral. Passing a scalar
  // arithmetic expression surfaces #VALUE! (not a usable cash-flow
  // source).
  const Value v = EvalSource("=XIRR(1+2, {1,2})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialXirr, ArgumentErrorPropagates) {
  Workbook wb = MakeMicrosoftFixture();
  // guess = 1/0 -> #DIV/0! wins before the iteration starts.
  const Value v = EvalSourceIn("=XIRR(A1:A5, B1:B5, 1/0)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(FinancialXirr, ArrayLiteralArgsWork) {
  // The lazy dispatcher must also handle inline ArrayLiteral ranges,
  // matching the IRR / MIRR path.
  const Value v = EvalSource("=XIRR({-10000, 2750, 4250, 3250, 2750}, {39448, 39508, 39751, 39859, 39904}, 0.1)");
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 0.373362535, 1e-6);
}

TEST(FinancialXirr, ArityTooFew) {
  const Value v = EvalSource("=XIRR({-1,1})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialXnpv, ArityTooFew) {
  const Value v = EvalSource("=XNPV(0.1, {-1,1})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialXnpv, ArityTooMany) {
  const Value v = EvalSource("=XNPV(0.1, {-1,1}, {1,2}, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
