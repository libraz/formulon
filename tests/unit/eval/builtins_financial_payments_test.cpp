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

}  // namespace
}  // namespace eval
}  // namespace formulon
