// Financial builtin tests grouped by schedule family.

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
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {
using formulon::test::EvalSource;
using formulon::test::EvalSourceIn;

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

TEST(FinancialNPV, DirectBoolContributesAtPeriod) {
  const Value v = EvalSource("=NPV(0.1, TRUE)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0 / 1.1, 1e-12);
}

TEST(FinancialNPV, DirectBoolAdvancesSubsequentPeriods) {
  const Value v = EvalSource("=NPV(0.1, 100, TRUE, 300)");
  const double expected = 100.0 / 1.1 + 1.0 / std::pow(1.1, 2) + 300.0 / std::pow(1.1, 3);
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), expected, 1e-10);
}

TEST(FinancialNPV, DirectNumericTextContributes) {
  const Value v = EvalSource("=NPV(0.1, \"5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 5.0 / 1.1, 1e-12);
}

TEST(FinancialNPV, DirectNonNumericTextIsIgnored) {
  const Value v = EvalSource("=NPV(0.1, \"x\", 100)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 100.0 / 1.1, 1e-12);
}

TEST(FinancialNPV, RangeBoolStillDropped) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::boolean(true));
  wb.sheet(0).set_cell_value(1, 0, Value::number(100.0));
  const Value v = EvalSourceIn("=NPV(0.1, A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  // The TRUE cell is filtered out; only 100 remains at period 1.
  EXPECT_NEAR(v.as_number(), 100.0 / 1.1, 1e-12);
}

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

}  // namespace
}  // namespace eval
}  // namespace formulon
