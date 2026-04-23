// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the security-rate / T-Bill financial built-ins:
// DISC, INTRATE, RECEIVED, TBILLPRICE, TBILLYIELD, TBILLEQ. These all
// live in `eval/builtins/financial_rates.cpp` and share the day-count
// helpers exported from `eval/date_time.h`.

#include <cmath>
#include <string>
#include <string_view>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Parses `src` and evaluates it through the default function registry
// with no bound workbook. Shared between the scalar-only rate tests.
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

}  // namespace
}  // namespace eval
}  // namespace formulon
