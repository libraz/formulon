//
// Mac Excel 365 reports a nonzero result below DBL_MIN as 0 for multiply,
// divide, POWER, `%`, per-product and per-square accumulation and the final
// division of AVERAGE, while + - SUM and comparisons keep it. Subnormals are
// read back by scaling with powers of two, which is lossless.

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

void ExpectErr(const std::string& formula, ErrorCode code) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_error()) << formula << " kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_error(), code) << formula;
}

void ExpectBool(const std::string& formula, bool expected) {
  const Value v = EvalSource(formula);
  ASSERT_TRUE(v.is_boolean()) << formula << " kind=" << static_cast<int>(v.kind());
  EXPECT_EQ(v.as_boolean(), expected) << formula;
}

TEST(SubnormalFlushPower, SubnormalResultIsZero) {
  ExpectNum("=SQRT(ABS(2^-1023))*1E200", 0.0);
  ExpectNum("=2^-1060.5", 0.0);
  ExpectNum("=0.5^1022*2^600*2^474", 4503599627370496.0);
  ExpectNum("=SQRT(ABS(0.5^1023))*1E200", 0.0);
  ExpectNum("=(2^-354)^3*2^600*2^474", 0.0);
}

TEST(SubnormalFlushPower, NegativeExponentOfUnderflowIsDivByZero) {
  ExpectErr("=(1E-200)^-2", ErrorCode::Div0);
  ExpectErr("=(1E-160)^-2", ErrorCode::Div0);
}

TEST(SubnormalFlushPower, OperatorSquareIsNotFlushed) {
  ExpectNum("=(2^-530)^2*2^600*2^474", 16384.0);
  ExpectNum("=(2^-530)^(1+1)*2^600*2^474", 16384.0);
  ExpectNum("=POWER(2^-530,2)*2^600*2^474", 0.0);
}

TEST(SubnormalFlushPower, SquaringLoopLimit) {
  ExpectNum("=(1-2^-30)^(2^31)", 0.13533528122258104);
  ExpectNum("=(1-2^-30)^(2^32)", 0.018315638854618645);
  ExpectNum("=IF(EXP(-720/4294967294)^4294967294>0,1,0)", 0.0);
  ExpectNum("=IF(EXP(-720/4294967295)^4294967295>0,1,0)", 1.0);
}

TEST(SubnormalFlushProducts, SumProduct) {
  ExpectNum("=SUMPRODUCT({1.5E-154,1.5E-154},{1.1E-154,1.1E-154})", 0.0);
  ExpectNum("=SUMPRODUCT({1E-200},{1E-110},{1E180})", 0.0);
  ExpectNum("=SUMPRODUCT({1E180},{1E-200},{1E-110})", 1.0000000000000001e-130);
}

TEST(SubnormalFlushProducts, Mmult) {
  ExpectNum("=INDEX(MMULT({1.5E-154,1.5E-154},{1.1E-154;1.1E-154}),1,1)", 0.0);
}

TEST(SubnormalFlushProducts, Squares) {
  ExpectNum("=SUMSQ(1.3E-154,1.3E-154)", 0.0);
  ExpectNum("=SUMXMY2({1.3E-154,1.3E-154},{0,0})", 0.0);
  ExpectNum("=DEVSQ(1.3E-154,-1.3E-154)", 0.0);
}

TEST(SubnormalFlushProducts, SeriesSum) {
  ExpectNum("=SQRT(ABS(SERIESSUM(2^-530,2,0,{1})))*1E200", 0.0);
}

TEST(SubnormalFlushDivision, AverageAndPercent) {
  ExpectNum("=SQRT(ABS(AVERAGE(EXP(-720),EXP(-720))))*1E200", 0.0);
  ExpectNum("=SQRT(ABS((1E-307)%))*1E200", 0.0);
}

TEST(SubnormalFlushZeroTest, SignAndTruthiness) {
  ExpectNum("=SIGN(EXP(-720))", 0.0);
  ExpectNum("=SIGN(2^-1022*(1+2^-4)-2^-1022)", 1.0);
  ExpectNum("=IF(EXP(-720),1,0)", 0.0);
  ExpectNum("=IF(EXP(-720)<>0,1,0)", 1.0);
  ExpectBool("=NOT(EXP(-720))", true);
}

TEST(SubnormalFlushZeroTest, UnaryMinusDropsSignBelowThreshold) {
  ExpectNum("=-EXP(-720)*2^600", 8.4327733541944975e-133);
  ExpectNum("=(0-EXP(-720))*2^600", -8.4327733541944975e-133);
  ExpectNum("=-(2^-1022*(1+2^-4)-2^-1022)*2^600", -5.7706116361160849e-129);
}

TEST(SubnormalFlushMod, SubnormalRemainderAndDivisor) {
  ExpectErr("=MOD(5E-308,3E-308)", ErrorCode::Num);
  ExpectNum("=MOD(9E-308,3.3E-308)", 2.4000000000000006e-308);
  ExpectErr("=MOD(EXP(-720),1)", ErrorCode::Num);
  ExpectErr("=MOD(1,EXP(-720))", ErrorCode::Div0);
  ExpectErr("=MROUND(EXP(-720),EXP(-720))", ErrorCode::Num);
}

TEST(SubnormalFlushControls, BinarySubtractionKeepsSubnormal) {
  ExpectNum("=(1E-300*1.000000000001-1E-300)*2^600*2^474", 202433888256.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
