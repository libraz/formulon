//
// XIRR reproduces Mac Excel 365's bracket-and-bisect search bit for bit:
// the doubling bracket [0, 2s], the negative brackets [-1, -s] and [-s, 0],
// the midpoint taken from the upper end, and the stop rule. Every expected
// value below is the full double Excel returned for the same formula.

#include <string>

#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;

void ExpectBits(const std::string& formula, double expected) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_number()) << formula << " kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), expected) << formula;
}

TEST(XirrSearch, DocumentationScheduleWithDefaultGuess) {
  ExpectBits("=XIRR({-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})", 0x1.7e52bf9999998p-2);
}

TEST(XirrSearch, BracketDoublesUpToTheRoot) {
  ExpectBits("=XIRR({-1,2},{39448,39813})", 0x1.ffffff999999ap-1);
}

TEST(XirrSearch, RootOnAMidpointIsApproachedFromBelow) {
  ExpectBits("=XIRR({-1,2},{39448,39813},1)", 0x1.ffffff8000000p-1);
}

TEST(XirrSearch, RootAtZeroStopsOnAHalvedGuess) {
  ExpectBits("=XIRR({-1,1},{39448,39813},0.5)", 0x1.0000000000000p-28);
}

TEST(XirrSearch, ZeroGuessStartsFromOneHundredThousandth) {
  ExpectBits("=XIRR({-1,2},{39448,39848},0)", 0x1.c3bd691a75cd0p-1);
}

TEST(XirrSearch, GuessAboveTheRootStillBisectsFromZero) {
  ExpectBits("=XIRR({-100,110},{39448,39813},3)", 0x1.9999998000000p-4);
}

TEST(XirrSearch, NegativeRootBelowMinusGuessUsesTheLowerBracket) {
  ExpectBits("=XIRR({-2,1},{39448,39813},0.5)", -0x1.0000004000000p-1);
  ExpectBits("=XIRR({-1,0.2},{39448,39848},0.5)", -0x1.8a1d59c000000p-1);
}

TEST(XirrSearch, NegativeRootAboveMinusGuessUsesTheUpperBracket) {
  ExpectBits("=XIRR({-1,0.8},{39448,39848},0.3)", -0x1.794bb9cccccccp-3);
}

TEST(XirrSearch, NanAtMinusOneStartsTheLowerBracketOneStepAbove) {
  // Later flows of both signs make f(-1) inf - inf.
  ExpectBits("=XIRR({-100;-20;60},{40000;40100;40200},0.1)", -0x1.82041f799999ap-1);
}

TEST(XirrSearch, NegativeRootIsNumWhenGuessIsAtLeastOne) {
  const Value v = EvalSource("=XIRR({-2,1},{39448,39813},1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
