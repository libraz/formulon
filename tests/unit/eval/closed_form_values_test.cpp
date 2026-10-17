// Exact values of the power core, SERIESSUM, the normal density and the
// exponential distribution.

#include <string>

#include "gtest/gtest.h"
#include "util/test_eval_helpers.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;

void ExpectNum(const std::string& formula, double expected) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_number()) << formula << " kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_number(), expected) << formula;
}

TEST(PowerOfTen, MatchesExcelValues) {
  ExpectNum("=10^100", 1e100);
  ExpectNum("=10^200", 1e200);
  ExpectNum("=10^300", 1e300);
  ExpectNum("=10^308", 1e308);
  ExpectNum("=POWER(10,150)", 1e150);
  ExpectNum("=10^-5", 1e-5);
  const Value over = EvalSource("=10^309");
  ASSERT_TRUE(over.is_error());
  EXPECT_EQ(over.as_error(), ErrorCode::Num);
}

TEST(SeriesSumExponent, MatchesExcelValues) {
  ExpectNum("=SERIESSUM(1.5,0,0.1,{1;1;1;1;1;1;1;1;1;1;1;1})", 15.145276319531888);
  ExpectNum("=SERIESSUM(1.5,1,0.1,{1;1;1;1;1;1;1;1;1;1;1;1})", 22.71791447929784);
  ExpectNum("=SERIESSUM(1.5,0.2,0.1,{1;1;1;1;1;1;1;1;1;1;1;1})", 16.424624635521308);
  ExpectNum("=SERIESSUM(10,100,100,{1;1;1})", 1e300);
}

TEST(SeriesSumExponent, NegativeBaseDropsFractionalTerms) {
  ExpectNum("=SERIESSUM(-2.2,0,0.5,{1;2})", 1.0);
  ExpectNum("=SERIESSUM(-8,1/3,1,{1})", 0.0);
  ExpectNum("=SERIESSUM(-2.2,0,1,{1;2})", -3.4000000000000004);
}

TEST(NormalDensity, MatchesExcelValues) {
  ExpectNum("=PHI(4.2697)", 4.387755789814979e-05);
  ExpectNum("=PHI(5.7036)", 3.442561205188857e-08);
  ExpectNum("=PHI(-2.8921)", 0.006090289032507509);
  ExpectNum("=PHI(38.5)", 0.0);
  ExpectNum("=NORM.DIST(17.947,-3.482,1.962,FALSE)", 2.538700187017102e-27);
  ExpectNum("=NORM.DIST(19.353,-7.699,3.836,FALSE)", 1.6509151245796868e-12);
  ExpectNum("=NORM.DIST(3.474,2.227,1.206,FALSE)", 0.19382061706176956);
  ExpectNum("=NORM.DIST(39,1,1,FALSE)", 0.0);
  ExpectNum("=NORM.S.DIST(4.493,FALSE)", 1.6494838680992867e-05);
  ExpectNum("=NORM.S.DIST(-3.1222,FALSE)", 0.0030488068756613657);
}

TEST(ExponentialDistribution, MatchesExcelValues) {
  ExpectNum("=EXPON.DIST(6.999,4.305,TRUE)", 0.9999999999999178);
  ExpectNum("=EXPON.DIST(6.999,4.305,FALSE)", 3.5349117938495164e-13);
  ExpectNum("=EXPONDIST(19.898,1.523,TRUE)", 0.9999999999999309);
  ExpectNum("=EXPON.DIST(19.898,1.523,FALSE)", 1.0508866987863447e-13);
  ExpectNum("=EXPON.DIST(1E-300,1,TRUE)", 1e-300);
  ExpectNum("=EXPON.DIST(1E-300,1,FALSE)", 1.0);
  ExpectNum("=EXPON.DIST(745,1,TRUE)", 1.0);
  ExpectNum("=EXPON.DIST(745,1,FALSE)", 0.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
