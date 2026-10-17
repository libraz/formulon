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

// Excel 365 doubles, bit-exact.
TEST(ExcelExactTimeValue, MatchesExcelBitForBit) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=PV(0.05/12, 60, -500)", 26495.35316196374},
      {"=PV(0.05/12, 60, -500, -10000)", 34287.407065133426},
      {"=PV(0.05/12, 60, -500, 0, 1)", 26605.750466805257},
      {"=PV(0.5, 10, -500)", 982.6584700841673},
      {"=FV(0.08/12, 120, -200)", 36589.20703634143},
      {"=FV(0.05/12, 60, -500, -10000)", 46836.62820545683},
      {"=FV(0.08/12, 120, -200, 0, 1)", 36833.13508325037},
      {"=FV(0.1, 20, -100)", 5727.499949325609},
      {"=FV(-0.99, 1000, 1)", -1.0101010101010102},
      {"=EFFECT(0.0525,12)", 0.0537818867274622},
      {"=EFFECT(0.05,4)", 0.050945336914062445},
      {"=EFFECT(0.1,4)", 0.10381289062499977},
      {"=EFFECT(0.05,365)", 0.0512674964674229},
      {"=EFFECT(0.0525,1)", 0.05249999999999999},
      {"=EFFECT(0.0525,2)", 0.05318906250000022},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(FinancialPaymentsClosedForm, MatchesExcelValues) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=IPMT(0.05/12, 60, 60, 25000)", -1.9575968510384782},
      {"=PPMT(0.05/12, 1, 60, 25000)", -367.6141744336067},
      {"=PPMT(0.05/12, 60, 60, 25000)", -469.8232442492349},
      {"=PPMT(0.05/12, 1, 60, 25000, 0, 1)", -469.8232442492349},
      {"=CUMPRINC(0.05/12, 60, 25000, 1, 60, 0)", -25000.0},
      {"=CUMPRINC(0.05/12, 60, 25000, 1, 12, 1)", -4598.88596634685},
      {"=CUMIPMT(0.05/12, 60, 25000, 1, 12, 0)", -1147.4887686633192},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(FinancialKernels, MatchesExcelValues) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=PDURATION(0.01,1,2)", 69.66071689357483},
      {"=DB(2400,300,10,0.5)", 451.2},
      {"=DB(2400,300,10,10.5)", 69.24259796758413},
      {"=MULTINOMIAL(1,2,3)", 59.99999999999993},
      {"=YIELDDISC(DATE(2005,12,31),DATE(2015,7,15),99.795,1000,0)", 0.9453843260780671},
      {"=YIELDDISC(DATE(2008,2,29),DATE(2021,2,28),90.5,1000)", 0.7730556736081599},
      {"=YIELDDISC(DATE(2014,2,28),DATE(2016,7,1),97.3,100,2)", 0.011697567339236688},
      {"=DURATION(DATE(2005,12,31),DATE(2021,2,28),0.1,0.05,1,0)", 9.17079402150038},
      {"=MDURATION(DATE(2008,2,29),DATE(2021,2,28),0.1,0.15,2,0)", 6.106342401129919},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(FinancialSchedules, MatchesExcelValues) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=VDB(1234.56,300,20,3,6)", 243.89843904000003},
      {"=VDB(0.75,0,7,1,3,3,FALSE)", 0.2886297376093295},
      {"=VDB(1000000,1,5,3,4,2,TRUE)", 86400.0},
      {"=PRICE(DATE(2014,2,28),DATE(2034,8,31),0.08,0.09,100,1,4)", 90.70079326566801},
      {"=PRICE(DATE(2016,1,1),DATE(2025,3,31),0.065,0.1,100,1,4)", 79.43965091302346},
      {"=ACCRINT(DATE(2008,3,1),DATE(2010,8,31),DATE(2008,12,31),0.05,100,1,1)", 4.171232876712328},
      {"=ACCRINT(DATE(2007,8,31),DATE(2012,8,31),DATE(2013,3,1),0.1,1000,1,4)", 550.2777777777778},
      {"=AMORLINC(5000,DATE(2008,8,19),DATE(2008,12,31),123.45,2,0.15,1)", 750.0},
      {"=AMORLINC(1234.56,DATE(2012,2,29),DATE(2012,12,31),0,5,0.15)", 185.184},
      {"=MIRR({-1000,500,700},0.1,0.1)", 0.1180339887498949},
      {"=MIRR({-1000,500,700},0.1,0.12)", 0.12249721603218244},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(FinancialSchedules, AmorlincBasisTwoIsNum) {
  const Value v = EvalSource("=AMORLINC(5000,DATE(2008,8,19),DATE(2008,12,31),123.45,2,0.15,2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ComplexReciprocalTrig, MatchesExcelValues) {
  const Value csc = EvalSource("=IMCSC(\"1+i\")");
  ASSERT_TRUE(csc.is_text());
  EXPECT_EQ(csc.as_text(), "0.621518017170428-0.303931001628427i");
  const Value csc_conj = EvalSource("=IMCSC(\"1-i\")");
  ASSERT_TRUE(csc_conj.is_text());
  EXPECT_EQ(csc_conj.as_text(), "0.621518017170428+0.303931001628427i");
}

TEST(FinancialSolvers, MatchesExcelValues) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=NPV(0.1,{-10000,3000,4200,6800})", 1188.4434123352207},
      {"=NPV(0.08,{-10000,3000,4200,6800})", 1645.055612946685},
      {"=ASIN(0.5)", 0.5235987755982989},
      {"=ACOS(0.5)", 1.0471975511965976},
      {"=ASIN(0.25)", 0.25268025514207865},
      {"=ACOS(0.25)", 1.318116071652818},
      {"=YIELD(DATE(2014,2,28),DATE(2034,8,31),0.0575,88.3,100,4,2)", 0.0681335077911804},
      {"=YIELD(DATE(2008,2,29),DATE(2026,1,1),0.0575,110.25,100,1,3)", 0.0487513713275858},
      {"=YIELD(DATE(2008,1,1),DATE(2021,2,28),0.0575,101.5,1000,1,3)", 0.21522473374983334},
      {"=YIELD(DATE(2008,1,1),DATE(2015,7,15),0.1,95.04287,100,2,2)", 0.1098339003443397},
      {"=YIELD(DATE(2010,7,15),DATE(2015,7,15),0,120,100,4,0)", -0.036298609507787176},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(ComplexText, MatchesExcelValues) {
  struct Case {
    const char* formula;
    const char* expected;
  };
  const Case cases[] = {
      {"=COMPLEX(0.0000001,0.0000002)", "0.0000001+0.0000002i"},
      {"=COMPLEX(0.00001,0.000002)", "0.00001+0.000002i"},
      {"=COMPLEX(0.001,123456789012345)", "0.001+123456789012345i"},
      {"=COMPLEX(1E-20,1E20)", "1E-20+100000000000000000000i"},
      {"=COMPLEX(1.23456789012345E-6,0)", "1.23456789012345E-06"},
      {"=COMPLEX(0,1.23456789012345E-6)", "1.23456789012345E-06i"},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_text()) << c.formula;
    EXPECT_EQ(v.as_text(), c.expected) << c.formula;
  }
}

TEST(FinancialOddFirst, MatchesExcelValues) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0785,0.065,100,2,0)",
       111.31988204177658},
      {"=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,6,15),DATE(2009,3,1),0.0785,0.065,100,2,0)",
       111.26923534442776},
      {"=ODDFPRICE(DATE(2009,1,10),DATE(2021,3,1),DATE(2008,6,15),DATE(2009,3,1),0.0785,0.065,100,2,0)",
       111.18886595243998},
      {"=ODDFYIELD(DATE(2012,3,1),DATE(2022,2,28),DATE(2012,2,10),DATE(2012,8,31),0.0785,84.5,100,2,0)",
       0.09649993958525949},
      {"=ODDFYIELD(DATE(2010,1,15),DATE(2020,6,30),DATE(2009,10,1),DATE(2010,6,30),0.0785,84.5,100,2,1)",
       0.09574388098247555},
      {"=ODDFYIELD(DATE(2009,1,10),DATE(2021,3,1),DATE(2008,6,15),DATE(2009,3,1),0.0785,84.5,100,2,2)",
       0.10082869870066025},
      {"=ODDFYIELD(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,99.74,100,2,0)",
       0.05780093761027392},
      {"=ODDFYIELD(DATE(2008,11,11),DATE(2009,9,1),DATE(2008,10,15),DATE(2009,3,1),0.0575,105.5,100,2,0)",
       -0.010321719871718392},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(FinancialOddFirst, FirstCouponOffTheMaturityScheduleIsNum) {
  const Value v =
      EvalSource("=ODDFPRICE(DATE(2008,11,11),DATE(2021,3,1),DATE(2008,10,15),DATE(2009,3,5),0.0785,0.065,100,2,0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ComplexTrigHyperbolic, MatchesExcelValues) {
  struct Case {
    const char* formula;
    const char* expected;
  };
  const Case cases[] = {
      {"=IMSIN(\"1+i\")", "1.29845758141598+0.634963914784736i"},
      {"=IMCOS(\"1+i\")", "0.833730025131149-0.988897705762865i"},
      {"=IMTAN(\"1+i\")", "0.271752585319512+1.08392332733869i"},
      {"=IMCOT(\"1+i\")", "0.217621561854403-0.868014142895925i"},
      {"=IMSINH(\"1+i\")", "0.634963914784736+1.29845758141598i"},
      {"=IMCOSH(\"1+i\")", "0.833730025131149+0.988897705762865i"},
      {"=IMCSCH(\"1+i\")", "0.303931001628426-0.621518017170429i"},
      {"=IMSECH(\"1+i\")", "0.498337030555187-0.591083841721045i"},
      {"=IMPRODUCT(IMCSCH(\"1+i\"),IMSINH(\"1+i\"))", "1-5.55111512312578E-17i"},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_text()) << c.formula;
    EXPECT_EQ(v.as_text(), c.expected) << c.formula;
  }
}

TEST(SpecialDistributions, MatchesExcelValues) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=GAMMA(0.5)", 1.772453850905516},
      {"=GAMMA(1.5)", 0.886226925452758},
      {"=GAMMA(2.5)", 1.329340388179137},
      {"=GAMMALN(0.5)", 0.5723649429247001},
      {"=GAMMALN(1.5)", -0.12078223763524523},
      {"=GAMMALN.PRECISE(0.5)", 0.5723649429247001},
      {"=WEIBULL.DIST(105,20,100,TRUE)", 0.929581390069277},
      {"=WEIBULL.DIST(105,20,100,FALSE)", 0.035588864024503564},
      {"=LOGNORM.DIST(1,0,1,FALSE)", 0.39894228040143265},
      {"=LOGNORM.DIST(2,0,1,FALSE)", 0.1568740192789811},
      {"=LOGNORM.DIST(0.5,0,1,FALSE)", 0.6274960771159244},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(SpecialDistributions, GammaDomain) {
  for (const char* formula : {"=GAMMA(0)", "=GAMMA(-3)", "=GAMMA(172)"}) {
    const Value v = EvalSource(formula);
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Num) << formula;
  }
  const Value tiny = EvalSource("=GAMMA(-171.5)");
  ASSERT_TRUE(tiny.is_number());
  EXPECT_EQ(tiny.as_number(), 0.0);
}

TEST(ExactRemainderAndAngles, MatchesExcelValues) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=MOD(-771618.708616,-22.69)", -22.56861599997459},
      {"=MOD(931127,63.4079)", 45.39640000003038},
      {"=MOD(-108663,8.9404)", 7.562000000004215},
      {"=MOD(953952.3,-55.3608)", -25.00559999991151},
      {"=MOD(751068.5,-85.85)", -33.14999999995027},
      {"=MOD(0.7*134217727,0.7)", 0.6999999970197677},
      {"=MOD(12345.678*134217727+12345.678/2,12345.678)", 6172.8391328125},
      {"=ATAN2(-379273,48.77)", 3.141464065475872},
      {"=ATAN2(-549252,-70.7077)", -3.141463919056993},
      {"=ATAN2(881469.044,93.3196)", 0.00010586826648827292},
      {"=ATAN2(697982.245,-66.96)", -9.593367205863957e-05},
      {"=ATAN2(0,2)", 1.5707963267948966},
      {"=ATAN2(0,-2)", -1.5707963267948966},
      {"=ATAN2(-3,0)", 3.141592653589793},
      {"=RRI(3,25232.91,976406.04)", 2.382369422704911},
      {"=RRI(38.9,54123.17,821882.7)", 0.07243471583009375},
      {"=XNPV(0.2036,{-6330.59;3480.01;1130.15;3587.31},{44205;45070;46123;46186})", -2348.615433109875},
      {"=XNPV(0.263,{-2570.64;102.19;2936.41;1394.33;-117.67},{38615;38960;39138;39696;40565})", 277.26245335237104},
      {"=WORKDAY(2958465.9,-30)", 2958423.0},
      {"=WORKDAY(2958465.9,0)", 2958465.0},
      {"=WORKDAY.INTL(2958465.9,-30,11)", 2958430.0},
      {"=WORKDAY.INTL(2958465.9,-1,\"0101000\")", 2958463.0},
      {"=DATE(1900,3,0)", 60.0},
      {"=DAYS360(59,60)", 1.0},
      {"=DAYS360(60,61)", 1.0},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(ExactRemainderAndAngles, EomonthBeforeTheCalendarStartIsNum) {
  const Value v = EvalSource("=EOMONTH(1,-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(ExactRemainderAndAngles, PercentKeepsEveryDigitAfterACarry) {
  const Value v = EvalSource("=TEXT(VALUE(\"1E+21\"),\"0%\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "100000000000000000000000%");
  const Value w = EvalSource("=TEXT(VALUE(\"1E+22\"),\"0.0%\")");
  ASSERT_TRUE(w.is_text());
  EXPECT_EQ(w.as_text(), "1000000000000000000000000.0%");
}

TEST(RoundingFamily, MatchesExcelValues) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=ROUNDUP(0.285,1)", 0.30000000000000004},
      {"=ROUNDUP(2.675,2)", 2.6799999999999997},
      {"=ROUNDUP(123456789012.345,2)", 123456789012.34999},
      {"=ROUNDUP(-0.285,1)", -0.30000000000000004},
      {"=ROUND(1E+15,2)", 1000000000000000.0},
      {"=ROUND(1234567.891,12)", 1234567.891},
      {"=ROUND(99.995,2)", 100.0},
      {"=ROUND(-2.5,0)", -3.0},
      {"=ROUND(1234.5,-2)", 1200.0},
      {"=TRUNC(123456789012.345,7)", 123456789012.345},
      {"=ROUNDDOWN(123456789012.345,7)", 123456789012.345},
      {"=ROUNDDOWN(0.0049,2)", 0.0},
      {"=CEILING(1E+15,0.3)", 999999999999999.0},
      {"=FLOOR(1E+15,0.07)", 1000000000000001.1},
      {"=CEILING(1.23456789012345E+15,0.07)", 1234567890123449.0},
      {"=CEILING(1E+12,1E-5)", 1000000000000.0001},
      {"=CEILING.MATH(1E+15,0.3)", 999999999999999.0},
      {"=FLOOR.MATH(1E+15,0.07)", 1000000000000001.1},
      {"=CEILING.PRECISE(1E+15,0.3)", 999999999999999.0},
      {"=FLOOR.PRECISE(1E+15,0.07)", 1000000000000001.1},
      {"=ISO.CEILING(1E+15,0.3)", 999999999999999.0},
      {"=MOD(1125899999999.5,1)", 0.5},
      {"=MOD(2251799999999,2)", 1.0},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(RoundingFamily, ModBeyondTheQuotientCapIsNum) {
  for (const char* formula :
       {"=MOD(1125900000000,1)", "=MOD(2251800000001,2)", "=MOD(-1125900000000,1)", "=MOD(1125900000000,-1)"}) {
    const Value v = EvalSource(formula);
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Num) << formula;
  }
}

TEST(DecliningBalance, MatchesExcelValues) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=DDB(10000,1000,5,0.5)", 4000.0},
      {"=DDB(10000,1000,5,1.5)", 3098.386676965934},
      {"=DDB(10000,1000,5,4.25)", 760.4165405894919},
      {"=DDB(973279,25039.99,4,4)", 60829.9375},
      {"=DDB(135229.88,42371.69,7,3)", 19712.81049562682},
      {"=DDB(2400,2000,10,3)", 0.0},
      {"=DDB(5000,500,6,2,3)", 1250.0},
      {"=DDB(1000,100,1.5,1,3)", 900.0},
      {"=DDB(1000,100,1.5,1.2,3)", 0.0},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(DecliningBalance, OutOfDomainIsNum) {
  for (const char* formula : {"=DDB(10000,1000,5,0)", "=DDB(10000,1000,5,5.5)", "=DDB(10000,1000,5,2,-1)",
                              "=DDB(-1,0,5,1)", "=DDB(10000,-1,5,1)", "=DDB(10000,1000,0,1)"}) {
    const Value v = EvalSource(formula);
    ASSERT_TRUE(v.is_error()) << formula;
    EXPECT_EQ(v.as_error(), ErrorCode::Num) << formula;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
