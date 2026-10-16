//
// End-to-end tests for the security-rate / T-Bill financial built-ins:
// DISC, INTRATE, RECEIVED, TBILLPRICE, TBILLYIELD, TBILLEQ. These all
// live in `eval/builtins/financial_rates.cpp` and share the day-count
// helpers exported from `utils/date_time.h`.

#include <cmath>
#include <iterator>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "test_eval_helpers.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;

// Evaluates `src` under the 1904 date system. See
// `builtins_financial_coupon_test.cpp` / `builtins_datetime_test.cpp` for
// the same helper.
Value EvalSource1904(std::string_view src) {
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
  return evaluate(*root, eval_arena, default_registry(), test::mac_context().with_date1904(true));
}

// ---------------------------------------------------------------------------
// DISC
// ---------------------------------------------------------------------------

TEST(FinancialDisc, MicrosoftDocExample) {
  // From the DISC docs: settlement 2018-01-25, maturity 2018-06-15,
  // pr 97.975, redemption 100, basis 1 (actual/actual). 2018 is not a
  // leap year, so the span is 141 / 365 = 0.38630... of a year, giving
  // DISC = 0.02025 / 0.38630... = 0.052420... (the Microsoft docs page
  // shows 0.052420 for the same shape).
  const Value v = EvalSource("=DISC(DATE(2018,1,25), DATE(2018,6,15), 97.975, 100, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.05242, 1e-4);
}

TEST(FinancialDisc, DefaultBasisZero) {
  // Basis 0 (US 30/360) is the default when the arg is omitted.
  const Value v = EvalSource("=DISC(DATE(2018,1,25), DATE(2018,6,15), 97.975, 100)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
  EXPECT_LT(v.as_number(), 0.1);
}

TEST(FinancialDisc1904, AgreesWithThe1900EquivalentCalendarDate) {
  // Exercises DISC's `yearfrac_for_basis` choke point: decomposing the
  // serial without threading date1904 would shift the computed year
  // fraction (and hence the rate) by the ~1462-day epoch gap.
  const Value v1900 = EvalSource("=DISC(DATE(2018,1,25), DATE(2018,6,15), 97.975, 100, 1)");
  const Value v1904 = EvalSource1904("=DISC(DATE(2018,1,25), DATE(2018,6,15), 97.975, 100, 1)");
  ASSERT_TRUE(v1900.is_number());
  ASSERT_TRUE(v1904.is_number());
  EXPECT_EQ(v1900.as_number(), v1904.as_number());
}

TEST(FinancialDisc, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=DISC(DATE(2018,6,15), DATE(2018,1,25), 97.975, 100, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDisc, SettlementEqualsMaturityIsNum) {
  const Value v = EvalSource("=DISC(DATE(2018,1,25), DATE(2018,1,25), 97.975, 100, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDisc, PriceNonPositiveIsNum) {
  const Value v = EvalSource("=DISC(DATE(2018,1,25), DATE(2018,6,15), 0, 100, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDisc, RedemptionNonPositiveIsNum) {
  const Value v = EvalSource("=DISC(DATE(2018,1,25), DATE(2018,6,15), 97.975, -1, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDisc, InvalidBasisIsNum) {
  const Value v = EvalSource("=DISC(DATE(2018,1,25), DATE(2018,6,15), 97.975, 100, 5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialDisc, NonNumericIsValue) {
  const Value v = EvalSource("=DISC(\"oops\", DATE(2018,6,15), 97.975, 100, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// INTRATE
// ---------------------------------------------------------------------------

TEST(FinancialIntrate, MicrosoftDocExample) {
  // From the INTRATE docs: settlement 2008-02-15, maturity 2008-05-15,
  // investment 1,000,000, redemption 1,014,420, basis 2 (actual/360)
  // -> approx 0.05768.
  const Value v = EvalSource("=INTRATE(DATE(2008,2,15), DATE(2008,5,15), 1000000, 1014420, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.05768, 1e-4);
}

TEST(FinancialIntrate, DefaultBasisZero) {
  const Value v = EvalSource("=INTRATE(DATE(2008,2,15), DATE(2008,5,15), 1000000, 1014420)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
}

TEST(FinancialIntrate, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=INTRATE(DATE(2008,5,15), DATE(2008,2,15), 1000000, 1014420, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialIntrate, InvestmentNonPositiveIsNum) {
  const Value v = EvalSource("=INTRATE(DATE(2008,2,15), DATE(2008,5,15), 0, 1014420, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialIntrate, RedemptionNonPositiveIsNum) {
  const Value v = EvalSource("=INTRATE(DATE(2008,2,15), DATE(2008,5,15), 1000000, -1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialIntrate, InvalidBasisIsNum) {
  const Value v = EvalSource("=INTRATE(DATE(2008,2,15), DATE(2008,5,15), 1000000, 1014420, 7)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// RECEIVED
// ---------------------------------------------------------------------------

TEST(FinancialReceived, MicrosoftDocExample) {
  // From the RECEIVED docs: settlement 2008-02-15, maturity 2008-05-15,
  // investment 1,000,000, discount 5.75%, basis 2 (actual/360)
  // -> approx 1,014,584.65.
  const Value v = EvalSource("=RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 1000000, 0.0575, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1014584.65, 1.0);
}

TEST(FinancialReceived, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=RECEIVED(DATE(2008,5,15), DATE(2008,2,15), 1000000, 0.0575, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialReceived, InvestmentNonPositiveIsNum) {
  const Value v = EvalSource("=RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 0, 0.0575, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialReceived, DiscountNonPositiveIsNum) {
  const Value v = EvalSource("=RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 1000000, 0, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialReceived, DiscountConsumesFaceIsNum) {
  // A discount rate of 5.0 applied over 1/4 year gives
  // 1 - 5.0*(0.25) = -0.25 -> #NUM!.
  const Value v = EvalSource("=RECEIVED(DATE(2008,2,15), DATE(2008,5,15), 1000000, 5.0, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// TBILLPRICE
// ---------------------------------------------------------------------------

TEST(FinancialTBillPrice, MicrosoftDocExample) {
  // From the TBILLPRICE docs: settlement 2008-03-31, maturity 2008-06-01,
  // discount 9% -> ~98.45.
  const Value v = EvalSource("=TBILLPRICE(DATE(2008,3,31), DATE(2008,6,1), 0.09)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 98.45, 0.01);
}

TEST(FinancialTBillPrice, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=TBILLPRICE(DATE(2008,6,1), DATE(2008,3,31), 0.09)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillPrice, DsmOver365IsNum) {
  // Maturity two years out -> T-Bill domain violation.
  const Value v = EvalSource("=TBILLPRICE(DATE(2008,3,31), DATE(2010,3,31), 0.09)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillPrice, RejectsMoreThanOneCalendarYear) {
  // Settlement 1902-09-26 (serial 1000) -> maturity 1903-09-27 (serial
  // 1366). The anniversary serial is 1365 (1903-09-26), so maturity is
  // one day past it; Excel returns #NUM! even though the raw DSM is 366.
  const Value v = EvalSource("=TBILLPRICE(1000, 1366, 0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillPrice, DiscountNonPositiveIsNum) {
  const Value v = EvalSource("=TBILLPRICE(DATE(2008,3,31), DATE(2008,6,1), 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillPrice, OverLargeDiscountYieldsNonPositivePrice) {
  // discount * DSM/360 > 1 -> result <= 0 -> #NUM!.
  const Value v = EvalSource("=TBILLPRICE(DATE(2008,1,1), DATE(2008,12,31), 1.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillPrice, RejectsBoolSettlement) {
  // Bool settlement must surface #VALUE! rather than coerce to 1/0.
  const Value v = EvalSource("=TBILLPRICE(TRUE, DATE(2008,6,1), 0.09)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// TBILLYIELD
// ---------------------------------------------------------------------------

TEST(FinancialTBillYield, MicrosoftDocExample) {
  // From the TBILLYIELD docs: settlement 2008-03-31, maturity 2008-06-01,
  // pr 98.45 -> ~0.09141.
  const Value v = EvalSource("=TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), 98.45)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.09141, 1e-4);
}

TEST(FinancialTBillYield1904, AgreesWithThe1900EquivalentCalendarDate) {
  // Exercises `t_bill_dsm`'s direct `ymd_from_serial` call (the
  // anniversary-rule check), a separate choke point from
  // `yearfrac_for_basis`.
  const Value v1900 = EvalSource("=TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), 98.45)");
  const Value v1904 = EvalSource1904("=TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), 98.45)");
  ASSERT_TRUE(v1900.is_number());
  ASSERT_TRUE(v1904.is_number());
  EXPECT_EQ(v1900.as_number(), v1904.as_number());
}

TEST(FinancialTBillYield, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=TBILLYIELD(DATE(2008,6,1), DATE(2008,3,31), 98.45)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillYield, DsmOver365IsNum) {
  const Value v = EvalSource("=TBILLYIELD(DATE(2008,3,31), DATE(2010,3,31), 98.45)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillYield, PriceNonPositiveIsNum) {
  const Value v = EvalSource("=TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillYield, RejectsBoolPrice) {
  // Bool `pr` must surface #VALUE! rather than coerce TRUE/FALSE to 1/0.
  const Value v = EvalSource("=TBILLYIELD(DATE(2008,3,31), DATE(2008,6,1), TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// TBILLEQ
// ---------------------------------------------------------------------------

TEST(FinancialTBillEq, MicrosoftDocExampleShortBranch) {
  // From the TBILLEQ docs: settlement 2008-03-31, maturity 2008-06-01,
  // discount 9.14% -> ~0.09414. DSM = 62 days, so the short branch
  // (DSM <= 182) fires: TBILLEQ = 365*0.0914 / (360 - 0.0914*62) ~ 0.09414.
  const Value v = EvalSource("=TBILLEQ(DATE(2008,3,31), DATE(2008,6,1), 0.0914)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.09414, 1e-4);
}

TEST(FinancialTBillEq, ShortBranchMatchesClosedForm) {
  // Synthetic case firmly inside the short branch (DSM = 31 days).
  // Closed form: 365*0.05 / (360 - 0.05*31) = 18.25 / 358.45 = 0.050914...
  const Value v = EvalSource("=TBILLEQ(DATE(2020,1,1), DATE(2020,2,1), 0.05)");
  ASSERT_TRUE(v.is_number());
  const double expected = (365.0 * 0.05) / (360.0 - 0.05 * 31.0);
  EXPECT_NEAR(v.as_number(), expected, 1e-10);
}

TEST(FinancialTBillEq, LongBranchFinitePositive) {
  // DSM = 300 days (> 182) exercises the semi-annual-compounding
  // quadratic branch. Expected result is a small positive yield; the
  // precise value is verified against Excel via the oracle fixture.
  const Value v = EvalSource("=TBILLEQ(DATE(2020,1,1), DATE(2020,10,27), 0.05)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
  EXPECT_LT(v.as_number(), 0.5);
}

TEST(FinancialTBillEq, SettlementAfterMaturityIsNum) {
  const Value v = EvalSource("=TBILLEQ(DATE(2008,6,1), DATE(2008,3,31), 0.09)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillEq, DsmOver365IsNum) {
  const Value v = EvalSource("=TBILLEQ(DATE(2008,3,31), DATE(2010,3,31), 0.09)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillEq, DiscountNonPositiveIsNum) {
  const Value v = EvalSource("=TBILLEQ(DATE(2008,3,31), DATE(2008,6,1), 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillEq, NonNumericIsValue) {
  const Value v = EvalSource("=TBILLEQ(\"oops\", DATE(2008,6,1), 0.09)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(FinancialTBillEq, LongBranchMatchesMicrosoftFormula) {
  // DSM = 360 days, rate = 0.8 -> exercises the long-branch quadratic.
  // Per the Microsoft TBILLEQ formula:
  //   a     = 360 / 365
  //   price = 1 - 0.8 * 360 / 360 = 0.2
  //   disc  = a^2 - (2a - 1) * (1 - 1/price)
  //   res   = (-a + sqrt(disc)) / (a - 0.5)
  // which evaluates to ~2.506604825374815 (verified against Excel/IronCalc).
  const Value v = EvalSource("=TBILLEQ(DATE(2020,1,1), DATE(2020,12,26), 0.8)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.506604825374815, 1e-6);
}

TEST(FinancialTBillEq, AcceptsLeapYearDsm366) {
  // A 366-day span (e.g. a one-year maturity across a leap-year boundary)
  // must not be rejected. 2020 is a leap year so DATE(2020,1,1) ->
  // DATE(2020,12,31) is exactly 365 days; use 2019-12-31 -> 2020-12-31 to
  // get a 366-day span. The calendar-year anniversary of 2019-12-31 is
  // 2020-12-31, exactly matching maturity, so the span is accepted.
  const Value v = EvalSource("=TBILLEQ(DATE(2019,12,31), DATE(2020,12,31), 0.05)");
  ASSERT_TRUE(v.is_number());
  EXPECT_GT(v.as_number(), 0.0);
}

TEST(FinancialTBillEq, RejectsMoreThanOneCalendarYear) {
  // Settlement 1902-09-26 (serial 1000) -> maturity 1903-09-27 (serial
  // 1366). Anniversary is 1903-09-26 (serial 1365); maturity is one day
  // past it. Raw DSM is 366 but Excel applies a calendar-year rule and
  // returns #NUM!.
  const Value v = EvalSource("=TBILLEQ(1000, 1366, 0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(FinancialTBillEq, AcceptsExactOneYearOnLeapAnniversary) {
  // Settlement 1940-02-16 (serial 14640, inside a leap year) ->
  // maturity 1941-02-16 (serial 15006). The calendar span is exactly one
  // year and the raw DSM is 366 (1940 is a leap year). Anniversary is
  // 1941-02-16 (serial 15006), so maturity == anniversary and the span
  // is accepted. The closed-form long-branch result is ~0.85229613 (from
  // the IronCalc oracle).
  const Value v = EvalSource("=TBILLEQ(14640, 15006, 0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.8522961312491701, 1e-6);
}

// RATE / YIELD / PRICE follow Excel's own iteration and summation order; the
// expected values are Mac Excel 365 results compared bit for bit, and kErr
// marks the cases Excel reports as #NUM!.
struct SolverCase {
  const char* formula;
  double expected;
};

constexpr double kErr = -1.0e300;

void ExpectSolverCases(const SolverCase* begin, const SolverCase* end) {
  for (const SolverCase* c = begin; c != end; ++c) {
    const Value v = EvalSource(c->formula);
    if (c->expected == kErr) {
      ASSERT_TRUE(v.is_error()) << c->formula;
      EXPECT_EQ(v.as_error(), ErrorCode::Num) << c->formula;
    } else {
      ASSERT_TRUE(v.is_number()) << c->formula;
      EXPECT_EQ(v.as_number(), c->expected) << c->formula;
    }
  }
}

TEST(FinancialSolverBits, Rate) {
  const SolverCase kRate[] = {
      {"=RATE(2,0,-1,4,0,0.1)", 1.0},
      {"=RATE(10,-100,1000,0,0,-0.05)", -2.041911992328346e-9},
      {"=RATE(10,-100,1000,0,0,0.0005)", -1.3775478725413001e-9},
      {"=RATE(12,-100,1200,0,0,0.0291)", -1.2895408274247784e-9},
      {"=RATE(10,-100,1000,0,0,0.0139)", -7.883954662430739e-10},
      {"=RATE(10,-100,1000,0,0,0.0655)", -1.2040782954525884e-9},
      {"=RATE(12,-100,1200,0,0,0.0212)", 2.556466805331766e-9},
      {"=RATE(10,-100,1000,0,0,0.0013)", 2.0990922206472353e-9},
      {"=RATE(12,-100,1200,0,0,0.0758)", -2.445262151554231e-9},
      {"=RATE(10,-100,1000,0,0,0.2761)", -1.7975479238456246e-9},
      {"=RATE(1,0,-100,0.00001,0,-0.9)", kErr},
      {"=RATE(1,0,-100,1E-14,0,-0.9)", kErr},
      {"=RATE(8.25,-500,1000,-5000,0,0.05)", kErr},
      {"=RATE(60.25,-1000,-200,0,1,0.01)", kErr},
      {"=RATE(8.3,-500,0,0,1,0.01)", kErr},
      {"=RATE(10.3,-250,-200,0,0,0.03)", kErr},
      {"=RATE(6,-250,100000,0,0,0.1)", -0.5990280207085503},
      {"=RATE(10,-50,-200,0,0,0.05)", kErr},
      {"=RATE(6,-100,-200,0,0,0.03)", kErr},
      {"=RATE(1,-1234.5,500,-5000,1,0.1)", kErr},
      {"=RATE(120,-250,500,-5000,1,0.03)", kErr},
      {"=RATE(6,-10,1000,-1000,0,0.1)", 0.010000000000000916},
      {"=RATE(5,-250,5000,100,1,0.1)", -0.5254109439807213},
      {"=RATE(5,-100,100000,2000,1,0.2)", kErr},
  };
  ExpectSolverCases(std::begin(kRate), std::end(kRate));
}

TEST(FinancialSolverBits, YieldNewton) {
  const SolverCase kYield[] = {
      {"=YIELD(DATE(2010,1,1),DATE(2012,1,1),0,50.0,100,1,0)", 0.4142135623115364},
      {"=YIELD(DATE(2010,1,1),DATE(2012,1,1),0,66.5,100,1,0)", 0.226278678969924},
      {"=YIELD(DATE(2010,1,1),DATE(2012,1,1),0,83.0,100,1,0)", 0.09764259989690346},
      {"=YIELD(DATE(2010,1,1),DATE(2012,1,1),0,99.5,100,1,0)", 0.0025094142304488295},
      {"=YIELD(DATE(2010,1,1),DATE(2012,1,1),0,116.0,100,1,0)", -0.0715233091235927},
      {"=YIELD(DATE(2010,1,1),DATE(2012,1,1),0,132.5,100,1,0)", -0.13125551447386127},
      {"=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,80.04287,100,2,0)", 0.09104999647018831},
      {"=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,104.04287,100,2,0)", 0.051684707256933166},
      {"=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,86.04287,100,2,1)", 0.07996342919553746},
      {"=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,112.04287,100,2,1)", 0.04096936923122717},
      {"=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,97.04287,100,2,2)", 0.06191379763438008},
      {"=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,81.04287,100,2,3)", 0.089136431041181},
      {"=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,112.04287,100,2,3)", 0.040970386688201015},
      {"=YIELD(DATE(2008,2,15),DATE(2016,11,15),0.0575,94.04287,100,2,4)", 0.06657419391469298},
      {"=YIELD(DATE(2010,7,1),DATE(2013,1,1),0,70,100,1,1)", 0.15307909296387448},
      {"=YIELD(DATE(2010,7,1),DATE(2013,1,1),0.05,99,100,1,1)", 0.054233433068821994},
      {"=YIELD(DATE(2010,7,1),DATE(2013,1,1),0.05,80,100,1,2)", 0.1516183573014838},
      {"=YIELD(DATE(2010,7,1),DATE(2013,1,1),0,110,100,1,3)", -0.037346269009127527},
  };
  ExpectSolverCases(std::begin(kYield), std::end(kYield));
}

TEST(FinancialSolverBits, PriceSeparateSums) {
  const SolverCase kPrice[] = {
      {"=PRICE(DATE(2010,1,1),DATE(2013,1,1),0.05,0.2391304347826087,100.0,1,0)", 62.47865740065779},
      {"=PRICE(DATE(2010,1,1),DATE(2013,1,1),0.05,0.11617885378139524,100.0,1,0)", 84.00000002769346},
      {"=PRICE(DATE(2010,1,1),DATE(2013,1,1),0.05,0.009257137767498935,100.0,1,0)", 112.00000487927517},
      {"=PRICE(DATE(2010,1,1),DATE(2013,1,1),0.1,0.16911764705882354,100.0,1,0)", 84.70600761901646},
      {"=PRICE(DATE(2010,1,1),DATE(2013,1,1),0.1,0.06954140510295907,100.0,1,0)", 108.00000000012301},
      {"=PRICE(DATE(2010,1,1),DATE(2013,1,1),0.05,0.19487439078245164,105.0,1,0)", 72.16674886033783},
      {"=PRICE(DATE(2010,1,1),DATE(2013,1,1),0.05,0.05,105.0,1,0)", 104.31918799265738},
      {"=PRICE(DATE(2010,1,1),DATE(2013,1,1),0.03,0.1342822950739185,100.0,1,0)", 75.555295838188},
      {"=PRICE(DATE(2010,1,1),DATE(2013,1,1),0.0575,0.10108407472960881,105.0,1,0)", 92.92742751301819},
      {"=PRICE(DATE(2008,2,15),DATE(2016,11,15),0.0575,0.065,100,2,1)", 95.04403378062287},
  };
  ExpectSolverCases(std::begin(kPrice), std::end(kPrice));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
