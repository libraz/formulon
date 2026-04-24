// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the non-day-count-convention financial built-ins:
// DOLLARDE, DOLLARFR, EFFECT, NOMINAL, FVSCHEDULE, PDURATION, RRI, MIRR,
// and ISPMT. These all live alongside the time-value-of-money family in
// `eval/builtins/financial.cpp`, except MIRR which rides the lazy
// dispatch seam in `eval/financial_lazy.cpp` (same reason as IRR — its
// `values` argument must reach the impl as un-flattened AST so a bare
// Ref, a RangeOp, or an inline ArrayLiteral can each be walked
// cell-by-cell).
//
// Tests that need a bound workbook populate the current sheet and use
// `EvalSourceIn`; pure-scalar cases run through `EvalSource`.

#include <cmath>
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

// Parses `src` and evaluates it through the default function registry
// with no bound workbook. Shared between scalar-only financial tests.
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
// Used by FVSCHEDULE-from-range and MIRR-from-range tests.
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
// DOLLARDE
// ---------------------------------------------------------------------------

TEST(FinancialDollarDe, BasicSixteenths) {
  // DOLLARDE(1.02, 16): "1 + 2/16" in scale-100 notation -> 1 + 2*100/16/... wait.
  // Excel rule: trunc(1.02)=1, frac=0.02, scale=100 (ceil(log10(16))=2),
  // so decimal = 1 + 0.02 * 100 / 16 = 1.125.
  const Value v = EvalSource("=DOLLARDE(1.02, 16)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.125, 1e-10);
}

TEST(FinancialDollarDe, HalvesSingleDigitDenom) {
  // denom=2: scale=10 (ceil(log10(2))=1), so DOLLARDE(1.1, 2) = 1 + 0.1*10/2 = 1.5.
  const Value v = EvalSource("=DOLLARDE(1.1, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.5, 1e-10);
}

TEST(FinancialDollarDe, Denom1IsIdentity) {
  // denom=1: scale = 10^ceil(log10(1)) = 10^0 = 1. Result = trunc + frac*1/1 = price.
  const Value v = EvalSource("=DOLLARDE(2.5, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.5, 1e-10);
}

TEST(FinancialDollarDe, NegativeDenomIsNum) {
  const Value v = EvalSource("=DOLLARDE(1.1, -2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDollarDe, FractionalDenomZeroIsDiv0) {
  // denom=0.5 truncates to 0 -> divide by zero domain.
  const Value v = EvalSource("=DOLLARDE(1.1, 0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(FinancialDollarDe, FractionalDenomTruncates) {
  // denom=16.9 truncates to 16 -> identical to DOLLARDE(1.02, 16).
  const Value v = EvalSource("=DOLLARDE(1.02, 16.9)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.125, 1e-10);
}

TEST(FinancialDollarDe, RejectsBoolPrice) {
  // Excel 365 / IronCalc oracle: Bool in the price slot yields #VALUE!
  // (not silent coerce to 0/1).
  const Value v = EvalSource("=DOLLARDE(TRUE, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// DOLLARFR
// ---------------------------------------------------------------------------

TEST(FinancialDollarFr, BasicSixteenths) {
  // Inverse of the DOLLARDE example: 1.125 -> 1 + 0.125*16/100 = 1.02.
  const Value v = EvalSource("=DOLLARFR(1.125, 16)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.02, 1e-10);
}

TEST(FinancialDollarFr, HalvesSingleDigitDenom) {
  const Value v = EvalSource("=DOLLARFR(1.5, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.1, 1e-10);
}

TEST(FinancialDollarFr, NegativeDenomIsNum) {
  const Value v = EvalSource("=DOLLARFR(1.5, -2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDollarFr, ZeroDenomIsDiv0) {
  const Value v = EvalSource("=DOLLARFR(1.5, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(FinancialDollarFr, RejectsBoolDenom) {
  // Excel 365 / IronCalc oracle: Bool in the denom slot yields #VALUE!
  // (not silent coerce to 0/1).
  const Value v = EvalSource("=DOLLARFR(1.5, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// EFFECT
// ---------------------------------------------------------------------------

TEST(FinancialEffect, MonthlyCompounding) {
  // 5.25% nominal compounded monthly -> (1 + 0.0525/12)^12 - 1 ~= 0.0537819.
  // Full double-precision reference: 0.05378188672746131.
  const Value v = EvalSource("=EFFECT(0.0525, 12)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.05378188672746131, 1e-12);
}

TEST(FinancialEffect, AnnualCompoundingIsIdentity) {
  // npery=1: effect == nominal.
  const Value v = EvalSource("=EFFECT(0.05, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.05, 1e-10);
}

TEST(FinancialEffect, NegativeNominalIsNum) {
  const Value v = EvalSource("=EFFECT(-0.05, 12)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialEffect, NperyLessThanOneIsNum) {
  // 0.5 truncates to 0 -> below the minimum 1.
  const Value v = EvalSource("=EFFECT(0.05, 0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// NOMINAL
// ---------------------------------------------------------------------------

TEST(FinancialNominal, InverseOfEffect) {
  // NOMINAL(EFFECT(x, n), n) == x. Round-trip through the effective rate
  // using the full-precision effective-rate constant above.
  const Value v = EvalSource("=NOMINAL(0.05378188672746131, 12)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0525, 1e-12);
}

TEST(FinancialNominal, AnnualCompoundingIsIdentity) {
  const Value v = EvalSource("=NOMINAL(0.05, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.05, 1e-10);
}

TEST(FinancialNominal, NegativeEffectIsNum) {
  const Value v = EvalSource("=NOMINAL(-0.05, 12)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// FVSCHEDULE
// ---------------------------------------------------------------------------

TEST(FinancialFvSchedule, BasicRangeOfRates) {
  // principal=1000, rates {0.05, 0.05, 0.10} -> 1000 * 1.05 * 1.05 * 1.10 = 1212.75.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(0.05));
  wb.sheet(0).set_cell_value(1, 0, Value::number(0.05));
  wb.sheet(0).set_cell_value(2, 0, Value::number(0.10));
  const Value v = EvalSourceIn("=FVSCHEDULE(1000, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1000.0 * 1.05 * 1.05 * 1.10, 1e-6);
}

TEST(FinancialFvSchedule, ScalarRatesWork) {
  // All scalar tail args: 100 * 1.1 * 1.2 = 132.
  const Value v = EvalSource("=FVSCHEDULE(100, 0.1, 0.2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 132.0, 1e-10);
}

TEST(FinancialFvSchedule, RangeSkipsTextCells) {
  // `range_filter_numeric_only` drops Text cells silently, matching NPV's
  // behaviour. 1000 * 1.05 * 1.10 = 1155.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(0.05));
  wb.sheet(0).set_cell_value(1, 0, Value::text("skip"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(0.10));
  const Value v = EvalSourceIn("=FVSCHEDULE(1000, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1000.0 * 1.05 * 1.10, 1e-6);
}

TEST(FinancialFvSchedule, ArityTooFew) {
  const Value v = EvalSource("=FVSCHEDULE(100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// PDURATION
// ---------------------------------------------------------------------------

TEST(FinancialPDuration, BasicGrowth) {
  // log(2000/1000) / log(1.1) ~= 7.2725 periods to double at 10%/period.
  const Value v = EvalSource("=PDURATION(0.1, 1000, 2000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 7.272540897, 1e-6);
}

TEST(FinancialPDuration, ZeroRateIsNum) {
  const Value v = EvalSource("=PDURATION(0, 1000, 2000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPDuration, NegativePvIsNum) {
  const Value v = EvalSource("=PDURATION(0.1, -1000, 2000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPDuration, NegativeFvIsNum) {
  const Value v = EvalSource("=PDURATION(0.1, 1000, -2000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// RRI
// ---------------------------------------------------------------------------

TEST(FinancialRri, BasicGrowthRate) {
  // (2000/1000)^(1/5) - 1 ~= 0.1487 (14.87% CAGR).
  const Value v = EvalSource("=RRI(5, 1000, 2000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.148698355, 1e-8);
}

TEST(FinancialRri, ZeroNperIsNum) {
  const Value v = EvalSource("=RRI(0, 1000, 2000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialRri, NegativePvIsNum) {
  const Value v = EvalSource("=RRI(5, -1000, 2000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialRri, RoundTripWithPDuration) {
  // RRI(PDURATION(r, pv, fv), pv, fv) == r. Numeric round-trip check.
  const Value v = EvalSource("=RRI(7.272540897, 1000, 2000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.1, 1e-6);
}

// ---------------------------------------------------------------------------
// MIRR
// ---------------------------------------------------------------------------

TEST(FinancialMirr, ClassicCase) {
  // Classic MIRR from Microsoft's example:
  //   values = {-120000, 39000, 30000, 21000, 37000, 46000}
  //   finance_rate = 0.10, reinvest_rate = 0.12
  //   MIRR ~= 0.126094 (12.61%).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(-120000.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(39000.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(30000.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(21000.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(37000.0));
  wb.sheet(0).set_cell_value(5, 0, Value::number(46000.0));
  const Value v = EvalSourceIn("=MIRR(A1:A6, 0.1, 0.12)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 0.126094, 1e-4);
}

TEST(FinancialMirr, AllPositiveIsDiv0) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(200.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(300.0));
  const Value v = EvalSourceIn("=MIRR(A1:A3, 0.1, 0.1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(FinancialMirr, AllNegativeIsMinusOne) {
  // Mac Excel 365: an all-negative cash-flow schedule yields exactly -1.
  // Intuition: npv_pos collapses to 0, so ratio = 0 / npv_neg = 0, and
  // `pow(0, 1/(n-1)) - 1 == -1.0` (IEEE-754 `pow(0, x) == 0` for any
  // positive finite x). This is NOT a #DIV/0! — only the all-positive
  // schedule divides by zero in the closed-form MIRR.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(-100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(-200.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(-300.0));
  const Value v = EvalSourceIn("=MIRR(A1:A3, 0.1, 0.1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), -1.0, 1e-12);
}

TEST(FinancialMirr, EqualFinanceAndReinvestAgreesWithIRR) {
  // When finance_rate == reinvest_rate, MIRR reduces to the IRR of the
  // identical cash flows (after the algebraic identity simplifies). Use a
  // small sequence where IRR is easy to verify independently.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(-1000.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(300.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(400.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(500.0));
  const Value mirr_v = EvalSourceIn("=MIRR(A1:A4, 0.0889633947, 0.0889633947)", wb, wb.sheet(0));
  ASSERT_TRUE(mirr_v.is_number());
  EXPECT_NEAR(mirr_v.as_number(), 0.0889633947, 1e-6);
}

TEST(FinancialMirr, ArityTooFew) {
  const Value v = EvalSource("=MIRR({-100,200}, 0.1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialMirr, ArrayLiteralValues) {
  // Inline ArrayLiteral also walks cell-by-cell: {-1000, 500, 700} with
  // finance=10%, reinvest=10% -> MIRR via closed form.
  const Value v = EvalSource("=MIRR({-1000, 500, 700}, 0.1, 0.1)");
  ASSERT_TRUE(v.is_number());
  // Hand-computed: npv_pos = 500/1.1 + 700/1.21 = 1033.05785;
  // npv_neg = -1000; (-(-1000)*1.21/... wait let me recompute:
  // Actually per canonical form: (-npv_pos * (1.1)^2 / npv_neg)^(1/2) - 1.
  // npv_pos = 500/1.1^1 + 700/1.1^2 = 454.5454... + 578.5124 = 1033.0579.
  // npv_neg = -1000/1.1^0 = -1000.
  // ratio = -1033.0579 * 1.21 / -1000 = 1249.9999... = 1.25 (actually 1.24999...).
  // result = sqrt(1.25) - 1 ~= 0.1180.
  const double npv_pos = 500.0 / 1.1 + 700.0 / (1.1 * 1.1);
  const double ratio = -npv_pos * (1.1 * 1.1) / -1000.0;
  const double expected = std::sqrt(ratio) - 1.0;
  EXPECT_NEAR(v.as_number(), expected, 1e-8);
}

// ---------------------------------------------------------------------------
// ISPMT
// ---------------------------------------------------------------------------

TEST(FinancialIsPmt, BasicMidloan) {
  // ISPMT(0.1/12, 1, 36, 8000000) = -8000000 * 0.1/12 * (1 - 1/36)
  //                                = -8000000 * 0.0083333 * 0.97222
  //                                ~= -64814.8.
  const Value v = EvalSource("=ISPMT(0.1/12, 1, 36, 8000000)");
  ASSERT_TRUE(v.is_number());
  const double expected = -8000000.0 * (0.1 / 12.0) * (1.0 - 1.0 / 36.0);
  EXPECT_NEAR(v.as_number(), expected, 1e-4);
}

TEST(FinancialIsPmt, FirstPeriodIsFullInterest) {
  // ISPMT uses simple-interest schedule, so at per=0 the interest is the
  // full -pv * rate. Excel-compat: ISPMT(r, 0, n, pv) = -pv * r * (1 - 0) = -pv * r.
  const Value v = EvalSource("=ISPMT(0.05, 0, 10, 1000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -1000.0 * 0.05, 1e-10);
}

TEST(FinancialIsPmt, LastPeriodIsZero) {
  // per == nper: (1 - 1) = 0 -> no interest (principal fully paid).
  const Value v = EvalSource("=ISPMT(0.05, 10, 10, 1000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0, 1e-10);
}

TEST(FinancialIsPmt, NperZeroIsDiv0) {
  const Value v = EvalSource("=ISPMT(0.05, 0, 0, 1000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(FinancialIsPmt, NegativePvFlipsSign) {
  // Negative pv (loan "asset" from the lender's view) flips the sign.
  const Value v = EvalSource("=ISPMT(0.05, 1, 10, -1000)");
  ASSERT_TRUE(v.is_number());
  const double expected = -(-1000.0) * 0.05 * (1.0 - 1.0 / 10.0);
  EXPECT_NEAR(v.as_number(), expected, 1e-10);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
