// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the financial built-ins: PV, FV, PMT, NPER, NPV,
// IRR, RATE, IPMT, PPMT, CUMIPMT, and CUMPRINC. All but IRR run through
// the eager registry dispatcher; IRR rides the lazy dispatch table
// because its first argument must reach the impl as un-flattened AST so
// range / Ref / ArrayLiteral shapes can all be walked cell-by-cell for
// Newton-Raphson.
//
// Each test parses a formula source, evaluates the AST through the
// default registry, and asserts the resulting Value. Tests that need a
// bound workbook populate the current sheet and use `EvalSourceIn`.

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
// Used by NPV-from-range and IRR-from-range tests.
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
// PV
// ---------------------------------------------------------------------------

TEST(FinancialPV, BasicEndOfPeriod) {
  // =PV(0.05/12, 60, -500) - 5-year car loan at 5% APR; negative pmt
  // (cash out) yields positive PV (loan amount in hand today).
  const Value v = EvalSource("=PV(0.05/12, 60, -500)");
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 26495.35316, 1e-4);
}

TEST(FinancialPV, ZeroRate) {
  // rate==0 branch: PV = -(pmt*nper + fv) = -(-500*60 + 0) = 30000.
  const Value v = EvalSource("=PV(0, 60, -500)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30000.0);
}

TEST(FinancialPV, ArityUnder) {
  const Value v = EvalSource("=PV(0.05, 60)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialPV, ArityOver) {
  const Value v = EvalSource("=PV(0.05, 60, -500, 0, 0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialPV, ErrorArgPropagates) {
  const Value v = EvalSource("=PV(#REF!, 60, -500)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(FinancialPV, NonNumericArgIsValue) {
  const Value v = EvalSource("=PV(\"abc\", 60, -500)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// FV
// ---------------------------------------------------------------------------

TEST(FinancialFV, BasicMonthlySavings) {
  // =FV(0.08/12, 120, -200) - 10-year deposit at 8% APR, $200/mo cash
  // out, zero PV. Should be ~36589.21 positive (cash back to investor).
  const Value v = EvalSource("=FV(0.08/12, 120, -200)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 36589.21, 1e-2);
}

TEST(FinancialFV, ZeroRate) {
  // rate==0: FV = -(pv + pmt*nper) = -(0 + -500*60) = 30000.
  const Value v = EvalSource("=FV(0, 60, -500)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 30000.0);
}

TEST(FinancialFV, WithPvAndType) {
  // Sanity: FV of same inputs at begin-of-period is greater than end-of-
  // period in magnitude (one more period of compounding).
  const Value end = EvalSource("=FV(0.05/12, 60, -500, -10000, 0)");
  const Value begin = EvalSource("=FV(0.05/12, 60, -500, -10000, 1)");
  ASSERT_TRUE(end.is_number());
  ASSERT_TRUE(begin.is_number());
  EXPECT_GT(begin.as_number(), end.as_number());
}

// ---------------------------------------------------------------------------
// PMT
// ---------------------------------------------------------------------------

TEST(FinancialPMT, CarLoanPayment) {
  // =PMT(0.05/12, 60, 25000) - pv positive (borrow), pmt negative (pay).
  const Value v = EvalSource("=PMT(0.05/12, 60, 25000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -471.78084, 1e-4);
}

TEST(FinancialPMT, ZeroRate) {
  // rate==0: PMT = -(pv+fv)/nper = -25000/60 ~= -416.6667.
  const Value v = EvalSource("=PMT(0, 60, 25000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -25000.0 / 60.0, 1e-10);
}

TEST(FinancialPMT, ZeroRateZeroNperIsNum) {
  // rate==0, nper==0 -> divide-by-zero in rate-0 branch -> #NUM!.
  const Value v = EvalSource("=PMT(0, 0, 25000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPMT, BeginOfPeriodIsSmallerMagnitude) {
  // Annuity-due (type=1) pays less per period than ordinary annuity
  // because each payment earns an extra period of interest.
  const Value end = EvalSource("=PMT(0.05/12, 60, 25000)");
  const Value begin = EvalSource("=PMT(0.05/12, 60, 25000, 0, 1)");
  ASSERT_TRUE(end.is_number());
  ASSERT_TRUE(begin.is_number());
  // Both negative; begin should be closer to zero.
  EXPECT_LT(std::fabs(begin.as_number()), std::fabs(end.as_number()));
}

TEST(FinancialPmt, RateAtMinusOneIsNum) {
  // Excel 365 / IronCalc oracle: rate == -1 is rejected outright as a
  // domain error, even though nper==1 would give a finite closed form.
  const Value v = EvalSource("=PMT(-1, 1, -100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPmt, RateBelowMinusOneIsNum) {
  // Excel 365 / IronCalc oracle: rate < -1 is a domain error. Without
  // the guard, (1 + rate)^nper = (-2)^5 = -32 would yield a finite
  // answer that disagrees with Excel.
  const Value v = EvalSource("=PMT(-3, 5, 100, 300, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPv, RateAtMinusOneIsDiv0) {
  // Excel 365 / IronCalc oracle (docs__PV A14): rate == -1 makes (1+r)^n
  // vanish so the closed form divides by zero; Excel surfaces this as
  // #DIV/0!, not the generic #NUM! used by PMT.
  const Value v = EvalSource("=PV(-1, 5, 100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(FinancialPv, RateBelowMinusOneIntegerNperIsFinite) {
  // calc_tests__PV F8/G8: PV accepts rate < -1 as long as nper is an
  // integer (the negative base to an integer power yields a real number).
  const Value v = EvalSource("=PV(-3, 5, 100)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 34.375);
}

TEST(FinancialFv, RateAtMinusOneIsDiv0) {
  // Excel 365 / IronCalc oracle (FV Sheet2 A1 and docs FV A14): rate == -1
  // collapses (1+r)^n to 0 or infinity depending on nper sign; Excel
  // surfaces both with #DIV/0!.
  const Value v = EvalSource("=FV(-1, 5, 100)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(FinancialFv, RateBelowMinusOneIntegerNperIsFinite) {
  // calc_tests__FV Sheet1 G8: FV accepts rate < -1 as long as nper is an
  // integer. FV(-3, 5, 100) = -pmt * ((-2)^5 - 1)/-3 = -100 * -33/-3.
  const Value v = EvalSource("=FV(-3, 5, 100)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -1100.0);
}

// ---------------------------------------------------------------------------
// NPER
// ---------------------------------------------------------------------------

TEST(FinancialNPER, BasicLoan) {
  // =NPER(0.05/12, -500, 25000) -> ~56.18 periods.
  const Value v = EvalSource("=NPER(0.05/12, -500, 25000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 56.18429, 1e-4);
}

TEST(FinancialNPER, ZeroRate) {
  // rate==0: NPER = -(pv+fv)/pmt = -25000/-500 = 50.
  const Value v = EvalSource("=NPER(0, -500, 25000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 50.0);
}

TEST(FinancialNPER, ZeroRateZeroPmtIsNum) {
  const Value v = EvalSource("=NPER(0, 0, 25000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialNPER, SignMismatchReturnsNegative) {
  // Excel 365 Mac: both pv and pmt positive produces the raw algebraic
  // answer (~-45.51 periods). Oracle confirmed. Not #NUM!.
  const Value v = EvalSource("=NPER(0.05/12, 500, 25000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -45.51263534, 1e-6);
}

TEST(FinancialNPER, RateBelowMinusOneIsNum) {
  const Value v = EvalSource("=NPER(-1.5, -500, 25000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// NPV
// ---------------------------------------------------------------------------

TEST(FinancialNPV, BasicScalars) {
  // 100/1.1 + 200/1.1^2 + 300/1.1^3 ~= 481.59.
  const Value v = EvalSource("=NPV(0.1, 100, 200, 300)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 481.5928, 1e-3);
}

TEST(FinancialNPV, ZeroRateIsSum) {
  // rate==0 reduces to sum since discount factor is 1.
  const Value v = EvalSource("=NPV(0, 100, 200, 300)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 600.0);
}

TEST(FinancialNPV, FromRange) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(200.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(300.0));
  const Value v = EvalSourceIn("=NPV(0.1, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 481.5928, 1e-3);
}

TEST(FinancialNPV, RangeSkipsTextCells) {
  // `range_filter_numeric_only` drops Text cells silently. Only the two
  // numeric cells participate: 100/1.1 + 300/1.1^2 = 338.84.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::text("skip"));
  wb.sheet(0).set_cell_value(2, 0, Value::number(300.0));
  const Value v = EvalSourceIn("=NPV(0.1, A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 100.0 / 1.1 + 300.0 / (1.1 * 1.1), 1e-10);
}

TEST(FinancialNPV, MixedScalarAndRange) {
  // Scalar args advance the positional index just like range cells.
  // =NPV(0.1, 50, A1:A2, 300) =
  //   50/1.1 + 100/1.1^2 + 200/1.1^3 + 300/1.1^4.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(200.0));
  const Value v = EvalSourceIn("=NPV(0.1, 50, A1:A2, 300)", wb, wb.sheet(0));
  const double expected = 50.0 / 1.1 + 100.0 / std::pow(1.1, 2) + 200.0 / std::pow(1.1, 3) + 300.0 / std::pow(1.1, 4);
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), expected, 1e-10);
}

TEST(FinancialNPV, ArityTooFew) {
  const Value v = EvalSource("=NPV(0.1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialNPV, ErrorInScalarPropagates) {
  const Value v = EvalSource("=NPV(0.1, 100, #DIV/0!, 300)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// IRR
// ---------------------------------------------------------------------------

TEST(FinancialIRR, BasicRangeInvestment) {
  // Classic textbook IRR: -1000 followed by +300, +400, +500 over three
  // periods. Excel gives ~0.0890 (8.9% IRR).
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(-1000.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(300.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(400.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(500.0));
  const Value v = EvalSourceIn("=IRR(A1:A4)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 0.0889633947, 1e-8);
}

TEST(FinancialIRR, WithExplicitGuess) {
  // Same sequence with an explicit starting guess should converge to
  // the same answer.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(-1000.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(300.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(400.0));
  wb.sheet(0).set_cell_value(3, 0, Value::number(500.0));
  const Value v = EvalSourceIn("=IRR(A1:A4, 0.2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0889633947, 1e-8);
}

TEST(FinancialIRR, AllPositiveIsNum) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(200.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(300.0));
  const Value v = EvalSourceIn("=IRR(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialIRR, AllNegativeIsNum) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(-100.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(-200.0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(-300.0));
  const Value v = EvalSourceIn("=IRR(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialIRR, ArrayLiteral) {
  // Inline array literal should walk exactly like a range.
  const Value v = EvalSource("=IRR({-1000,300,400,500})");
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 0.0889633947, 1e-8);
}

TEST(FinancialIRR, ArityNone) {
  const Value v = EvalSource("=IRR()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialIRR, ArityTooMany) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(-1000.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(500.0));
  const Value v = EvalSourceIn("=IRR(A1:A2, 0.1, 0.2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialIRR, ScalarFirstArgIsValue) {
  // A bare scalar cannot be walked as a cash flow sequence.
  const Value v = EvalSource("=IRR(5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialIRR, RangeErrorPropagates) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(-1000.0));
  wb.sheet(0).set_cell_value(1, 0, Value::error(ErrorCode::Div0));
  wb.sheet(0).set_cell_value(2, 0, Value::number(500.0));
  const Value v = EvalSourceIn("=IRR(A1:A3)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(FinancialIRR, RangeSkipsNonNumeric) {
  // Non-numeric cells inside the range are silently skipped, matching
  // Excel SUM/AVERAGE/IRR provenance behaviour. The effective sequence
  // is [-1000, 300, 400, 500] so the answer matches BasicRangeInvestment.
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(-1000.0));
  wb.sheet(0).set_cell_value(1, 0, Value::number(300.0));
  wb.sheet(0).set_cell_value(2, 0, Value::text("note"));
  wb.sheet(0).set_cell_value(3, 0, Value::number(400.0));
  wb.sheet(0).set_cell_value(4, 0, Value::number(500.0));
  const Value v = EvalSourceIn("=IRR(A1:A5)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0889633947, 1e-8);
}

// ---------------------------------------------------------------------------
// RATE
// ---------------------------------------------------------------------------

TEST(FinancialRATE, CarLoanConverges) {
  // =RATE(60, -500, 25000): 60 monthly payments of $500 amortise a
  // $25,000 principal. Total outflow 30000 implies 5000 interest over
  // five years, giving a monthly rate of ~0.006183 (APR ~7.42%). This
  // matches the analytic closed-form solution of the annuity equation.
  const Value v = EvalSource("=RATE(60, -500, 25000)");
  ASSERT_TRUE(v.is_number()) << "kind=" << static_cast<int>(v.kind());
  EXPECT_NEAR(v.as_number(), 0.006183413161254, 1e-8);
}

TEST(FinancialRATE, ZeroRateShortcut) {
  // pv + pmt*nper + fv == 0 -> exact zero-rate answer (no NR needed).
  const Value v = EvalSource("=RATE(10, -100, 1000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(FinancialRATE, ExplicitGuessMatches) {
  // Supplying an explicit guess should converge to the same answer as
  // the default guess.
  const Value def = EvalSource("=RATE(60, -500, 25000)");
  const Value guess = EvalSource("=RATE(60, -500, 25000, 0, 0, 0.05)");
  ASSERT_TRUE(def.is_number());
  ASSERT_TRUE(guess.is_number());
  EXPECT_NEAR(def.as_number(), guess.as_number(), 1e-9);
}

TEST(FinancialRATE, TypeOneDiffersFromTypeZero) {
  // Annuity-due solves to a slightly different rate from ordinary
  // annuity for the same (nper, pmt, pv).
  const Value type0 = EvalSource("=RATE(60, -500, 25000)");
  const Value type1 = EvalSource("=RATE(60, -500, 25000, 0, 1)");
  ASSERT_TRUE(type0.is_number());
  ASSERT_TRUE(type1.is_number());
  EXPECT_NE(type0.as_number(), type1.as_number());
  // Both should still be plausible positive rates.
  EXPECT_GT(type0.as_number(), 0.0);
  EXPECT_GT(type1.as_number(), 0.0);
}

TEST(FinancialRATE, NperBelowOneIsNum) {
  const Value v = EvalSource("=RATE(0, -500, 25000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialRATE, ArityUnder) {
  const Value v = EvalSource("=RATE(60, -500)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialRATE, ArityOver) {
  const Value v = EvalSource("=RATE(60, -500, 25000, 0, 0, 0.1, 0.2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// IPMT
// ---------------------------------------------------------------------------

TEST(FinancialIPMT, FirstPeriodInterestDominates) {
  // =IPMT(0.05/12, 1, 60, 25000): balance at start of period 1 is the
  // full pv (25000), so interest = -25000 * (0.05/12) ~= -104.1667.
  const Value v = EvalSource("=IPMT(0.05/12, 1, 60, 25000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -25000.0 * (0.05 / 12.0), 1e-8);
}

TEST(FinancialIPMT, LastPeriodInterestNearZero) {
  // Interest in the final period is small in magnitude (most of the
  // final payment is principal).
  const Value v = EvalSource("=IPMT(0.05/12, 60, 60, 25000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_LT(std::fabs(v.as_number()), 5.0);
  EXPECT_LT(v.as_number(), 0.0);  // still negative (interest payment)
}

TEST(FinancialIPMT, IntegerPerBeyondNperIsNum) {
  // Integer per > nper is an out-of-schedule period: Mac Excel 365 (and
  // the IronCalc oracle) return #NUM! in this case.
  const Value v = EvalSource("=IPMT(0.05/12, 61, 60, 25000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialIPMT, FractionalPerBeyondNperMatchesOracle) {
  // IronCalc / Mac Excel 365 oracle value for a fractional per > nper case
  // at type == 0.
  const Value v = EvalSource("=IPMT(0.1, 3.9, 3, 8000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -30.51485756246893, 1e-9);
}

TEST(FinancialIPMT, FractionalPerBeyondNperType1) {
  // Same fractional per > nper case with fv=10 and annuity-due (type=1).
  const Value v = EvalSource("=IPMT(0.1, 3.9, 3, 8000, 10, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -26.866364667656, 1e-9);
}

TEST(FinancialIPMT, PerZeroIsNum) {
  const Value v = EvalSource("=IPMT(0.05/12, 0, 60, 25000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialIPMT, TypeOneFirstPeriodIsZero) {
  // Annuity-due, first period: payment is at start so no interest has
  // accrued. IPMT must be exactly 0.
  const Value v = EvalSource("=IPMT(0.05/12, 1, 60, 25000, 0, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(FinancialIPMT, ZeroRateIsZero) {
  // No interest ever accrues at rate=0.
  const Value v = EvalSource("=IPMT(0, 5, 60, 25000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(FinancialIPMT, ArityUnder) {
  const Value v = EvalSource("=IPMT(0.05/12, 1, 60)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// PPMT
// ---------------------------------------------------------------------------

TEST(FinancialPPMT, FirstPeriodPrincipal) {
  // PPMT = PMT - IPMT for the same period.
  const Value pmt = EvalSource("=PMT(0.05/12, 60, 25000)");
  const Value ipmt = EvalSource("=IPMT(0.05/12, 1, 60, 25000)");
  const Value ppmt = EvalSource("=PPMT(0.05/12, 1, 60, 25000)");
  ASSERT_TRUE(pmt.is_number());
  ASSERT_TRUE(ipmt.is_number());
  ASSERT_TRUE(ppmt.is_number());
  EXPECT_NEAR(ppmt.as_number(), pmt.as_number() - ipmt.as_number(), 1e-9);
}

TEST(FinancialPPMT, IdentityAcrossAllPeriods) {
  // For every period i in [1, nper], IPMT(i) + PPMT(i) == PMT.
  const Value pmt = EvalSource("=PMT(0.05/12, 60, 25000)");
  ASSERT_TRUE(pmt.is_number());
  // Probe a handful of representative periods (full sweep is expensive
  // via EvalSource and adds no signal over ~5 strategic probes).
  for (int per : {1, 2, 15, 30, 59, 60}) {
    const std::string formula = "=IPMT(0.05/12, " + std::to_string(per) + ", 60, 25000) + PPMT(0.05/12, " +
                                std::to_string(per) + ", 60, 25000)";
    const Value sum = EvalSource(formula);
    ASSERT_TRUE(sum.is_number()) << "per=" << per;
    EXPECT_NEAR(sum.as_number(), pmt.as_number(), 1e-8) << "per=" << per;
  }
}

TEST(FinancialPPMT, IntegerPerBeyondNperIsNum) {
  // Mirrors FinancialIPMT.IntegerPerBeyondNperIsNum — integer per > nper
  // is still rejected as #NUM! to match the Mac Excel 365 oracle.
  const Value v = EvalSource("=PPMT(0.05/12, 61, 60, 25000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialPPMT, FractionalPerBeyondNperMatchesOracle) {
  // IronCalc / Mac Excel 365 oracle value for a fractional per > nper case.
  const Value v = EvalSource("=PPMT(0.1, 3.9, 3, 8000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -3186.4035714405495, 1e-9);
}

TEST(FinancialPPMT, ZeroRatePrincipalEqualsPmt) {
  // rate=0 -> IPMT=0, so PPMT == PMT for every period.
  const Value pmt = EvalSource("=PMT(0, 60, 25000)");
  const Value ppmt = EvalSource("=PPMT(0, 5, 60, 25000)");
  ASSERT_TRUE(pmt.is_number());
  ASSERT_TRUE(ppmt.is_number());
  EXPECT_DOUBLE_EQ(ppmt.as_number(), pmt.as_number());
}

// ---------------------------------------------------------------------------
// CUMIPMT
// ---------------------------------------------------------------------------

TEST(FinancialCUMIPMT, FullRangeMatchesSumOfIPMT) {
  // Sum of IPMT over [1, nper] must equal CUMIPMT over the same range.
  const Value cum = EvalSource("=CUMIPMT(0.05/12, 60, 25000, 1, 60, 0)");
  ASSERT_TRUE(cum.is_number());
  double expected_sum = 0.0;
  for (int per = 1; per <= 60; ++per) {
    const std::string f = "=IPMT(0.05/12, " + std::to_string(per) + ", 60, 25000)";
    const Value v = EvalSource(f);
    ASSERT_TRUE(v.is_number());
    expected_sum += v.as_number();
  }
  EXPECT_NEAR(cum.as_number(), expected_sum, 1e-6);
}

TEST(FinancialCUMIPMT, FirstYearSubset) {
  // Sum of IPMT over periods 1..12 equals CUMIPMT for the first year.
  const Value cum = EvalSource("=CUMIPMT(0.05/12, 60, 25000, 1, 12, 0)");
  ASSERT_TRUE(cum.is_number());
  double expected_sum = 0.0;
  for (int per = 1; per <= 12; ++per) {
    const std::string f = "=IPMT(0.05/12, " + std::to_string(per) + ", 60, 25000)";
    const Value v = EvalSource(f);
    ASSERT_TRUE(v.is_number());
    expected_sum += v.as_number();
  }
  EXPECT_NEAR(cum.as_number(), expected_sum, 1e-6);
}

TEST(FinancialCUMIPMT, StartZeroIsNum) {
  const Value v = EvalSource("=CUMIPMT(0.05/12, 60, 25000, 0, 12, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCUMIPMT, ReversedRangeIsNum) {
  // start > end is rejected.
  const Value v = EvalSource("=CUMIPMT(0.05/12, 60, 25000, 24, 12, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCUMIPMT, ZeroRateIsNum) {
  // Excel's documented rule: rate must be strictly positive.
  const Value v = EvalSource("=CUMIPMT(0, 60, 25000, 1, 12, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCUMIPMT, NegativePvIsNum) {
  const Value v = EvalSource("=CUMIPMT(0.05/12, 60, -25000, 1, 12, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCUMIPMT, ArityTooFew) {
  // `type` is required; CUMIPMT without the 6th arg is an arity error.
  const Value v = EvalSource("=CUMIPMT(0.05/12, 60, 25000, 1, 12)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// CUMPRINC
// ---------------------------------------------------------------------------

TEST(FinancialCUMPRINC, FullRangeEqualsNegativePv) {
  // Sum of all principal payments on a standard loan should equal -pv
  // (the loan gets fully paid off). Small tolerance for compounding
  // round-off over 60 periods.
  const Value v = EvalSource("=CUMPRINC(0.05/12, 60, 25000, 1, 60, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -25000.0, 1e-6);
}

TEST(FinancialCUMPRINC, CumPrincPlusCumIpmtEqualsTotalPayments) {
  // CUMPRINC + CUMIPMT over [1, nper] = PMT * nper (sum of payments).
  const Value cp = EvalSource("=CUMPRINC(0.05/12, 60, 25000, 1, 60, 0)");
  const Value ci = EvalSource("=CUMIPMT(0.05/12, 60, 25000, 1, 60, 0)");
  const Value pmt = EvalSource("=PMT(0.05/12, 60, 25000)");
  ASSERT_TRUE(cp.is_number());
  ASSERT_TRUE(ci.is_number());
  ASSERT_TRUE(pmt.is_number());
  EXPECT_NEAR(cp.as_number() + ci.as_number(), pmt.as_number() * 60.0, 1e-6);
}

TEST(FinancialCUMPRINC, StartOutOfRangeIsNum) {
  const Value v = EvalSource("=CUMPRINC(0.05/12, 60, 25000, 0, 12, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCUMPRINC, EndBeyondNperIsNum) {
  const Value v = EvalSource("=CUMPRINC(0.05/12, 60, 25000, 1, 61, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCUMPRINC, ArityTooFew) {
  // `type` is required.
  const Value v = EvalSource("=CUMPRINC(0.05/12, 60, 25000, 1, 12)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialCUMPRINC, FractionalStartRoundsUp) {
  // Mac Excel 365 ceils `start_period` and floors `end_period` before
  // iterating; start=1.2 therefore skips period 1 and begins at period 2.
  const Value ref = EvalSource("=CUMPRINC(0.01, 36, 8000, 2, 8, 1)");
  const Value frac = EvalSource("=CUMPRINC(0.01, 36, 8000, 1.2, 8.53, 1)");
  ASSERT_TRUE(ref.is_number());
  ASSERT_TRUE(frac.is_number());
  EXPECT_DOUBLE_EQ(frac.as_number(), ref.as_number());
}

TEST(FinancialCUMPRINC, FractionalTypeIsNum) {
  // `type` must be strictly 0 or 1; fractional values yield #NUM!.
  const Value v = EvalSource("=CUMPRINC(0.01, 36, 8000, 1, 10, 0.7)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCUMIPMT, FractionalStartRoundsUp) {
  const Value ref = EvalSource("=CUMIPMT(0.01, 36, 8000, 2, 8, 1)");
  const Value frac = EvalSource("=CUMIPMT(0.01, 36, 8000, 1.2, 8.53, 1)");
  ASSERT_TRUE(ref.is_number());
  ASSERT_TRUE(frac.is_number());
  EXPECT_DOUBLE_EQ(frac.as_number(), ref.as_number());
}

TEST(FinancialCUMIPMT, FractionalTypeIsNum) {
  const Value v = EvalSource("=CUMIPMT(0.01, 36, 8000, 1, 10, 1.2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialCumipmt, RejectsBoolStart) {
  // Excel 365 rejects Bool in any position with #VALUE! (no silent
  // coercion to 1.0). Mirrors the oracle's CUMPRINC_CUMIPMT G27/H27.
  const Value v = EvalSource("=CUMIPMT(0.05, 10, 8000, TRUE, 10, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialCumprinc, RejectsBoolType) {
  // Oracle CUMPRINC_CUMIPMT G28/H28: Bool in the `type` position is
  // rejected with #VALUE! rather than being folded to 1.
  const Value v = EvalSource("=CUMPRINC(0.05, 10, 8000, 3, 10, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// SLN — straight-line depreciation
// ---------------------------------------------------------------------------

TEST(FinancialSLN, Basic) {
  const Value v = EvalSource("=SLN(10000, 1000, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1800.0);
}

TEST(FinancialSLN, ZeroSalvage) {
  const Value v = EvalSource("=SLN(10000, 0, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1000.0);
}

TEST(FinancialSLN, IdentitySlnTimesLifeEqualsDepreciableBase) {
  // SLN * life == cost - salvage for any legal inputs.
  const Value v = EvalSource("=SLN(10000, 1000, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number() * 5.0, 10000.0 - 1000.0);
}

TEST(FinancialSLN, LifeZeroIsDiv0) {
  const Value v = EvalSource("=SLN(10000, 1000, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(FinancialSLN, NegativeLifeProducesNegativeDepreciation) {
  // Oracle-verified: Excel 365 Mac returns the raw algebraic answer
  // (-1800 for these inputs) rather than #NUM!.
  const Value v = EvalSource("=SLN(10000, 1000, -5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -1800.0);
}

TEST(FinancialSLN, ArityMismatchIsValue) {
  const Value v = EvalSource("=SLN(10000, 1000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// SYD — sum-of-years'-digits depreciation
// ---------------------------------------------------------------------------

TEST(FinancialSYD, FirstPeriod) {
  // (10000-1000)*5*2/(5*6) = 3000
  const Value v = EvalSource("=SYD(10000, 1000, 5, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3000.0);
}

TEST(FinancialSYD, LastPeriod) {
  // (10000-1000)*1*2/(5*6) = 600
  const Value v = EvalSource("=SYD(10000, 1000, 5, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 600.0);
}

TEST(FinancialSYD, SumAcrossAllPeriodsEqualsDepreciableBase) {
  // sum_{p=1..life} SYD(cost, salvage, life, p) == cost - salvage.
  double total = 0.0;
  for (int p = 1; p <= 5; ++p) {
    const std::string src = "=SYD(10000, 1000, 5, " + std::to_string(p) + ")";
    const Value v = EvalSource(src);
    ASSERT_TRUE(v.is_number()) << "p=" << p;
    total += v.as_number();
  }
  EXPECT_NEAR(total, 9000.0, 1e-9);
}

TEST(FinancialSYD, PeriodOutOfRangeIsNum) {
  const Value v = EvalSource("=SYD(10000, 1000, 5, 6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialSYD, PeriodBelowOneIsNum) {
  const Value v = EvalSource("=SYD(10000, 1000, 5, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialSYD, LifeZeroIsNum) {
  const Value v = EvalSource("=SYD(10000, 1000, 0, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// DDB — double-declining-balance depreciation
// ---------------------------------------------------------------------------

TEST(FinancialDDB, FirstPeriod) {
  // rate = 2/5 = 0.4; dep_1 = 10000 * 0.4 = 4000
  const Value v = EvalSource("=DDB(10000, 1000, 5, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 4000.0);
}

TEST(FinancialDDB, SecondPeriod) {
  // book = 6000 after period 1; dep_2 = 6000 * 0.4 = 2400
  const Value v = EvalSource("=DDB(10000, 1000, 5, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 2400.0);
}

TEST(FinancialDDB, CustomFactorTriples) {
  // factor=3 -> rate = 3/5 = 0.6; dep_1 = 10000 * 0.6 = 6000
  const Value v = EvalSource("=DDB(10000, 1000, 5, 1, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 6000.0);
}

TEST(FinancialDDB, SalvageFloorCaps) {
  // Running the full schedule: dep must never take book value below
  // salvage. Verify total == cost - salvage (9000) exactly — DDB
  // depreciates the full base over the asset's life, with the salvage
  // floor absorbing whatever remains in the last period. For this
  // schedule the final period is the capped one: 1296 * 0.4 = 518.4
  // would overshoot, so dep_5 = book - salvage = 296.
  double total = 0.0;
  for (int p = 1; p <= 5; ++p) {
    const std::string src = "=DDB(10000, 1000, 5, " + std::to_string(p) + ")";
    const Value v = EvalSource(src);
    ASSERT_TRUE(v.is_number()) << "p=" << p;
    total += v.as_number();
  }
  EXPECT_NEAR(total, 9000.0, 1e-9);
  const Value last = EvalSource("=DDB(10000, 1000, 5, 5)");
  ASSERT_TRUE(last.is_number());
  EXPECT_NEAR(last.as_number(), 296.0, 1e-9);
}

TEST(FinancialDDB, PeriodOutOfRangeIsNum) {
  const Value v = EvalSource("=DDB(10000, 1000, 5, 6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDDB, NegativeCostIsNum) {
  const Value v = EvalSource("=DDB(-100, 0, 5, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDDB, FactorZeroIsNum) {
  const Value v = EvalSource("=DDB(10000, 1000, 5, 1, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// DB — fixed-rate declining-balance depreciation
// ---------------------------------------------------------------------------

TEST(FinancialDB, FullYearFirstPeriod) {
  // rate = round(1 - (0.1)^0.2, 3) = 0.369; dep_1 = 10000 * 0.369 = 3690.
  const Value v = EvalSource("=DB(10000, 1000, 5, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 3690.0, 1e-6);
}

TEST(FinancialDB, PartialFirstYearProrated) {
  // month=6: dep_1 = 10000 * 0.369 * 6/12 = 1845.
  const Value v = EvalSource("=DB(10000, 1000, 5, 1, 6)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1845.0, 1e-6);
}

TEST(FinancialDB, PartialYearSumApproxEqualsDepreciableBase) {
  // With month=6 the depreciation runs life+1 = 6 periods; the sum
  // should approximately equal cost - salvage within the 3-decimal
  // rate-rounding error.
  double total = 0.0;
  for (int p = 1; p <= 6; ++p) {
    const std::string src = "=DB(10000, 1000, 5, " + std::to_string(p) + ", 6)";
    const Value v = EvalSource(src);
    ASSERT_TRUE(v.is_number()) << "p=" << p;
    total += v.as_number();
  }
  // Excel's 3-decimal rate rounding produces a small residual between
  // the sum of depreciations and the true depreciable base. Observed
  // residual for this fixture (cost=10000, salvage=1000, life=5,
  // month=6) is ~54 — tolerance of 100 leaves margin for rounding
  // accumulation without masking a real regression.
  EXPECT_NEAR(total, 9000.0, 100.0);
}

TEST(FinancialDB, PeriodLifePlusOneNeedsPartialFirstYear) {
  // month=12 (full year) with period == life+1 is invalid.
  const Value v = EvalSource("=DB(10000, 1000, 5, 6)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDB, PeriodLifePlusOneWithPartialYearIsValid) {
  // month=6, period=6 is the partial last year; must produce a
  // positive, finite charge.
  const Value v = EvalSource("=DB(10000, 1000, 5, 6, 6)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
}

TEST(FinancialDB, InvalidMonthIsNum) {
  const Value v = EvalSource("=DB(10000, 1000, 5, 1, 13)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDB, ZeroMonthIsNum) {
  const Value v = EvalSource("=DB(10000, 1000, 5, 1, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDB, ZeroCostIsNum) {
  const Value v = EvalSource("=DB(0, 0, 5, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
