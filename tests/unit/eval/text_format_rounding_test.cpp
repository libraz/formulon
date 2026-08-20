//
// End-to-end tests pinning the display-rounding contract shared by TEXT,
// FIXED and DOLLAR: a decimal tie renders the same digits that ROUND()
// produces for the same value and place. The format-engine-level cases live
// in `text_format_number_test.cpp`; this file exercises the built-ins so a
// divergence between the render helper and the arithmetic one is caught at
// the surface a workbook actually sees.

#include <string>
#include <string_view>

#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;

// Evaluates `formula` and returns its text payload copied out of the shared
// arena, so the caller can hold it across a second evaluation.
std::string EvalText(std::string_view formula) {
  const Value v = EvalSource(formula);
  EXPECT_TRUE(v.is_text()) << formula;
  return v.is_text() ? std::string(v.as_text()) : std::string();
}

double EvalNumber(std::string_view formula) {
  const Value v = EvalSource(formula);
  EXPECT_TRUE(v.is_number()) << formula;
  return v.is_number() ? v.as_number() : 0.0;
}

// ---------------------------------------------------------------------------
// TEXT / FIXED / DOLLAR agree with ROUND on decimal ties
// ---------------------------------------------------------------------------

TEST(DisplayRoundingTies, TextMatchesRoundAtTwoPlaces) {
  EXPECT_EQ(EvalText("=TEXT(1.005,\"0.00\")"), "1.01");
  EXPECT_EQ(EvalNumber("=ROUND(1.005,2)"), 1.01);
}

TEST(DisplayRoundingTies, TextMatchesRoundOnTextifiedRound) {
  // The same value through both paths must print identically.
  const std::string via_text = EvalText("=TEXT(2.675,\"0.00\")");
  const std::string via_round = EvalText("=TEXT(ROUND(2.675,2),\"0.00\")");
  EXPECT_EQ(via_text, "2.68");
  EXPECT_EQ(via_text, via_round);
}

TEST(DisplayRoundingTies, FixedMatchesRound) {
  EXPECT_EQ(EvalText("=FIXED(1.005,2)"), "1.01");
  EXPECT_EQ(EvalText("=FIXED(2.675,2)"), "2.68");
  EXPECT_EQ(EvalText("=FIXED(8.835,2)"), "8.84");
}

TEST(DisplayRoundingTies, DollarMatchesRound) {
  // ja-JP renders the yen sign `¥` (UTF-8 0xC2 0xA5) as the prefix.
  EXPECT_EQ(EvalText("=DOLLAR(1.005,2)"),
            "\xC2\xA5"
            "1.01");
}

TEST(DisplayRoundingTies, NegativeTiesRoundAwayFromZero) {
  EXPECT_EQ(EvalText("=TEXT(-1.005,\"0.00\")"), "-1.01");
  EXPECT_EQ(EvalText("=FIXED(-1.005,2)"), "-1.01");
  EXPECT_EQ(EvalNumber("=ROUND(-1.005,2)"), -1.01);
}

TEST(DisplayRoundingTies, ValuesBelowTheTieStillRoundDown) {
  EXPECT_EQ(EvalText("=TEXT(1.0049999,\"0.00\")"), "1.00");
  EXPECT_EQ(EvalText("=FIXED(1.0049999,2)"), "1.00");
  EXPECT_EQ(EvalNumber("=ROUND(1.0049999,2)"), 1.0);
}

TEST(DisplayRoundingTies, FixedKeepsLargeMagnitudesIntact) {
  // The tie handling must not perturb a value whose ULP is wider than 0.5.
  EXPECT_EQ(EvalText("=FIXED(9E+15,0)"), "9,000,000,000,000,000");
}

TEST(DisplayRoundingTies, FixedNegativeDecimals) {
  EXPECT_EQ(EvalText("=FIXED(1500,-3)"), "2,000");
  EXPECT_EQ(EvalNumber("=ROUND(1500,-3)"), 2000.0);
}

// ---------------------------------------------------------------------------
// Trailing-comma scaling reaches formats that carry a fractional part
// ---------------------------------------------------------------------------

TEST(DisplayRoundingScale, TextThousandsScaleWithFraction) {
  EXPECT_EQ(EvalText("=TEXT(1234567,\"0.0,\")"), "1234.6");
  EXPECT_EQ(EvalText("=TEXT(1234567,\"#,##0.0,\")"), "1,234.6");
}

TEST(DisplayRoundingScale, TextMillionsScaleWithFraction) {
  EXPECT_EQ(EvalText("=TEXT(1234567,\"0.00,,\")"), "1.23");
}

}  // namespace
}  // namespace eval
}  // namespace formulon
