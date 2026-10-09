//
// End-to-end tests for the math built-in functions: ABS, SIGN, INT, TRUNC,
// SQRT, MOD, POWER, ROUND, ROUNDDOWN, ROUNDUP, MIN, MAX, AVERAGE, PRODUCT.
// Each test parses a formula source, evaluates the AST through the default
// registry, and asserts the resulting Value.

#include <string_view>

#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;
using formulon::test::EvalSourceIn;

// ---------------------------------------------------------------------------
// ABS
// ---------------------------------------------------------------------------

TEST(MathAbs, PositiveNumber) {
  const Value v = EvalSource("=ABS(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(MathAbs, NegativeNumber) {
  const Value v = EvalSource("=ABS(-5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(MathAbs, Zero) {
  const Value v = EvalSource("=ABS(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(MathAbs, NumericTextCoerces) {
  const Value v = EvalSource("=ABS(\"3\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(MathAbs, ErrorPropagates) {
  const Value v = EvalSource("=ABS(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(MathAbs, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=ABS()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(MathAbs, TwoArgsIsArityViolation) {
  const Value v = EvalSource("=ABS(1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// SIGN
// ---------------------------------------------------------------------------

TEST(MathSign, NegativeIsMinusOne) {
  const Value v = EvalSource("=SIGN(-3.7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -1.0);
}

TEST(MathSign, ZeroIsZero) {
  const Value v = EvalSource("=SIGN(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(MathSign, PositiveIsOne) {
  const Value v = EvalSource("=SIGN(3.7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(MathSign, ErrorPropagates) {
  const Value v = EvalSource("=SIGN(#NAME?)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// INT (floor toward negative infinity, NOT toward zero)
// ---------------------------------------------------------------------------

TEST(MathInt, PositiveFloors) {
  const Value v = EvalSource("=INT(2.7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(MathInt, NegativeFloors) {
  // Excel quirk pin: INT(-2.7) -> -3, NOT -2. INT uses std::floor, not
  // std::trunc. This is the canonical INT-vs-TRUNC distinction.
  const Value v = EvalSource("=INT(-2.7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -3.0);
}

TEST(MathInt, IntegerInputUnchanged) {
  const Value v = EvalSource("=INT(2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(MathInt, NegativeIntegerInputUnchanged) {
  const Value v = EvalSource("=INT(-2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -2.0);
}

TEST(MathInt, ErrorPropagates) {
  const Value v = EvalSource("=INT(#DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// TRUNC (truncate toward zero, opposite of INT for negatives)
// ---------------------------------------------------------------------------

TEST(MathTrunc, PositiveTruncates) {
  const Value v = EvalSource("=TRUNC(2.7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(MathTrunc, NegativeTruncatesTowardZero) {
  // Excel quirk pin: TRUNC(-2.7) -> -2 (toward zero), unlike INT(-2.7) -> -3.
  const Value v = EvalSource("=TRUNC(-2.7)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -2.0);
}

TEST(MathTrunc, PositiveDigits) {
  const Value v = EvalSource("=TRUNC(2.789, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.78, 1e-12);
}

TEST(MathTrunc, NegativeValuePositiveDigits) {
  const Value v = EvalSource("=TRUNC(-2.789, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -2.78, 1e-12);
}

TEST(MathTrunc, NegativeDigits) {
  const Value v = EvalSource("=TRUNC(1234.5, -1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1230.0);
}

TEST(MathTrunc, SnapsBinaryRepresentationNoiseLikeRound) {
  // 0.29 * 100 == 28.999999999999996 in IEEE-754 double; without the
  // snap-to-integer compensation ROUND/CEILING/FLOOR already apply,
  // TRUNC(0.29, 2) truncates the (wrong) 28.999... down to 28 and returns
  // 0.28 instead of 0.29.
  const Value v = EvalSource("=TRUNC(0.29, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.29);
}

TEST(MathTrunc, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=TRUNC()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// A digit count far outside `int`'s range must saturate rather than wrap.
// Converting such a double to `int` is undefined, and the architectures
// disagree on the result, so an unclamped cast made these truncate to a no-op
// on one host and collapse to zero on another.
TEST(MathTrunc, HugePositiveDigitsIsANoOp) {
  const Value v = EvalSource("=TRUNC(1E100, 1E50)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1e100);

  const Value w = EvalSource("=TRUNC(5, 9.99999E307)");
  ASSERT_TRUE(w.is_number());
  EXPECT_EQ(w.as_number(), 5.0);
}

TEST(MathTrunc, HugeNegativeDigitsIsZero) {
  const Value v = EvalSource("=TRUNC(1E100, -1E50)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(MathRoundFamily, HugeDigitsSaturateConsistently) {
  // ROUND, ROUNDDOWN and ROUNDUP read the digit count through the same helper.
  for (const char* src : {"=ROUND(5, 1E50)", "=ROUNDDOWN(5, 1E50)", "=ROUNDUP(5, 1E50)"}) {
    const Value v = EvalSource(src);
    ASSERT_TRUE(v.is_number()) << src;
    EXPECT_EQ(v.as_number(), 5.0) << src;
  }
  for (const char* src : {"=ROUND(5, -1E50)", "=ROUNDDOWN(5, -1E50)", "=ROUNDUP(5, -1E50)"}) {
    const Value v = EvalSource(src);
    ASSERT_TRUE(v.is_number()) << src;
    EXPECT_EQ(v.as_number(), 0.0) << src;
  }
}

// ---------------------------------------------------------------------------
// SQRT
// ---------------------------------------------------------------------------

TEST(MathSqrt, PerfectSquare) {
  const Value v = EvalSource("=SQRT(4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(MathSqrt, Zero) {
  const Value v = EvalSource("=SQRT(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(MathSqrt, IrrationalApproximation) {
  const Value v = EvalSource("=SQRT(2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.4142135, 1e-6);
}

TEST(MathSqrt, NegativeYieldsNum) {
  const Value v = EvalSource("=SQRT(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// MOD (sign of result follows sign of divisor)
// ---------------------------------------------------------------------------

TEST(MathMod, BothPositive) {
  const Value v = EvalSource("=MOD(7, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(MathMod, NegativeDividendPositiveDivisor) {
  // Excel quirk pin: result has SIGN OF DIVISOR. C `%` would give -1 here.
  const Value v = EvalSource("=MOD(-7, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(MathMod, PositiveDividendNegativeDivisor) {
  const Value v = EvalSource("=MOD(7, -3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -2.0);
}

TEST(MathMod, BothNegative) {
  const Value v = EvalSource("=MOD(-7, -3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -1.0);
}

TEST(MathMod, SignFollowsDivisor) {
  // Compact summary of the four sign combinations above.
  EXPECT_EQ(EvalSource("=MOD(7, 3)").as_number(), 1.0);
  EXPECT_EQ(EvalSource("=MOD(-7, 3)").as_number(), 2.0);
  EXPECT_EQ(EvalSource("=MOD(7, -3)").as_number(), -2.0);
  EXPECT_EQ(EvalSource("=MOD(-7, -3)").as_number(), -1.0);
}

TEST(MathMod, DivisorZeroYieldsDiv0) {
  const Value v = EvalSource("=MOD(7, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(MathMod, FractionalArguments) {
  const Value v = EvalSource("=MOD(7.5, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.5, 1e-12);
}

// ---------------------------------------------------------------------------
// POWER
// ---------------------------------------------------------------------------

TEST(MathPower, BasicPositiveExponent) {
  const Value v = EvalSource("=POWER(2, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1024.0);
}

// Excel reports #NUM! for 0^0, diverging from the IEEE-754 pow convention.
// Both POWER() and the `^` operator share this rule via apply_pow.
TEST(MathPower, ZeroPowZeroIsNum) {
  const Value v = EvalSource("=POWER(0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// A zero base with a negative exponent is a division by zero, and Excel
// reports it with the `/` operator's code rather than the `#NUM!` that
// std::pow's `+Inf` would collapse into. `IFERROR` / `ERROR.TYPE` routing
// depends on the distinction, so both spellings are pinned.
TEST(MathPower, ZeroPowNegativeIsDiv0) {
  for (const std::string_view source : {"=POWER(0, -1)", "=POWER(0, -0.5)", "=POWER(0, -2)"}) {
    const Value v = EvalSource(source);
    ASSERT_TRUE(v.is_error()) << source;
    EXPECT_EQ(v.as_error(), ErrorCode::Div0) << source;
  }
}

TEST(MathPower, BinaryOpZeroPowNegativeIsDiv0) {
  for (const std::string_view source : {"=0^-1", "=0^-2", "=0^-0.5"}) {
    const Value v = EvalSource(source);
    ASSERT_TRUE(v.is_error()) << source;
    EXPECT_EQ(v.as_error(), ErrorCode::Div0) << source;
  }
}

TEST(MathPower, NegativeBaseFractionalExpYieldsNum) {
  const Value v = EvalSource("=POWER(-1, 0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// A negative base with a non-integer exponent e is an odd root when 1/e lies
// within half a unit of its 15th significant digit of an odd integer. Values
// are bit-exact Mac Excel results.
TEST(MathPower, NegativeBaseOddRootValues) {
  struct Case {
    const char* source;
    double expected;
  };
  const Case cases[] = {
      {"=(-8)^(1/3)", -1.9999999999999998},
      {"=POWER(-8,1/3)", -1.9999999999999998},
      {"=(-8)^(-1/3)", -0.5000000000000001},
      {"=(-27)^(1/3)", -2.9999999999999996},
      {"=(-32)^(1/5)", -2.0},
      {"=(-8)^0.333333333333333", -1.9999999999999984},
      {"=(-2)^(1/1000001)", -1.0000006931467276},
      {"=(-2)^(1/1001)", -1.0006926945279553},
      {"=(-1)^(1/3)", -1.0},
      {"=(-0.5)^(1/3)", -0.7937005259840998},
      {"=(-1E300)^(1/3)", -9.999999999999825e+99},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.source);
    ASSERT_TRUE(v.is_number()) << c.source;
    EXPECT_EQ(v.as_number(), c.expected) << c.source;
  }
}

TEST(MathPower, NegativeBaseNonOddRootIsNum) {
  for (const std::string_view source :
       {"=(-8)^(2/3)", "=(-8)^(4/3)", "=(-8)^(5/3)", "=(-8)^(-2/3)", "=(-8)^(1/6)", "=(-8)^1.5", "=(-8)^(1/3+1E-15)",
        "=(-8)^(1/3+1E-12)", "=(-8)^0.3333333333", "=(-8)^0.333", "=(-2)^0.5", "=(-2)^(3/7)", "=(-32)^(2/5)"}) {
    const Value v = EvalSource(source);
    ASSERT_TRUE(v.is_error()) << source;
    EXPECT_EQ(v.as_error(), ErrorCode::Num) << source;
  }
}

TEST(MathPower, NegativeExponent) {
  const Value v = EvalSource("=POWER(2, -2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.25);
}

// Confirms the BinaryOp::Pow path still works after the apply_pow refactor.
TEST(MathPower, BinaryOpPowStillWorks) {
  const Value v = EvalSource("=2^-2");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.25);
}

// ---------------------------------------------------------------------------
// ROUND (half away from zero)
// ---------------------------------------------------------------------------

TEST(MathRound, HalfAwayFromZeroPositive) {
  const Value v = EvalSource("=ROUND(2.5, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(MathRound, HalfAwayFromZeroNegative) {
  // Excel quirk pin: ROUND(-2.5, 0) -> -3, not -2. This distinguishes ROUND
  // from banker's rounding (which would round to even).
  const Value v = EvalSource("=ROUND(-2.5, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -3.0);
}

TEST(MathRound, PositiveDigits) {
  const Value v = EvalSource("=ROUND(2.345, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.35, 1e-12);
}

TEST(MathRound, NegativeDigits) {
  const Value v = EvalSource("=ROUND(1234, -2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1200.0);
}

TEST(MathRound, OneAndAHalfRoundsUp) {
  const Value v = EvalSource("=ROUND(1.5, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(MathRound, ValueOutsideUlpToleranceDoesNotRoundUp) {
  // The decimal is about 90 ULPs below 0.5. ROUND must not mistake it for
  // arithmetic noise around a half boundary.
  const Value v = EvalSource("=ROUND(0.49999999999999, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(MathRound, ExtremePositiveDigitsIsNoOp) {
  // `10^400` overflows to +Inf; without a clamp this used to surface a
  // spurious #NUM! instead of the mathematically correct no-op (a double
  // has no meaningful digits past ~17 decimal places).
  const Value v = EvalSource("=ROUND(1.5, 400)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.5);
}

TEST(MathRound, ExtremeNegativeDigitsIsZero) {
  // Rounding to the nearest 10^400 always lands on 0 for any finite double.
  const Value v = EvalSource("=ROUND(1.5, -400)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// ROUNDDOWN (always toward zero)
// ---------------------------------------------------------------------------

TEST(MathRoundDown, PositiveTowardZero) {
  const Value v = EvalSource("=ROUNDDOWN(2.99, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(MathRoundDown, NegativeTowardZero) {
  // Excel quirk pin: ROUNDDOWN(-2.99, 0) -> -2 (toward zero, not down).
  const Value v = EvalSource("=ROUNDDOWN(-2.99, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -2.0);
}

TEST(MathRoundDown, PositiveDigits) {
  const Value v = EvalSource("=ROUNDDOWN(1.999, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.99, 1e-12);
}

TEST(MathRoundDown, NegativeDigits) {
  const Value v = EvalSource("=ROUNDDOWN(1234, -2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1200.0);
}

TEST(MathRoundDown, SnapsBinaryRepresentationNoiseLikeRound) {
  // 4.35 * 100 == 434.99999999999994 in IEEE-754 double; without the snap
  // compensation, ROUNDDOWN(4.35, 2) truncates to 4.34 instead of 4.35.
  const Value v = EvalSource("=ROUNDDOWN(4.35, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.35);
}

TEST(MathRoundDown, ExtremePositiveDigitsIsNoOp) {
  const Value v = EvalSource("=ROUNDDOWN(1.5, 400)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.5);
}

TEST(MathRoundDown, ExtremeNegativeDigitsIsZero) {
  const Value v = EvalSource("=ROUNDDOWN(1.5, -400)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// ROUNDUP (always away from zero)
// ---------------------------------------------------------------------------

TEST(MathRoundUp, PositiveAwayFromZero) {
  const Value v = EvalSource("=ROUNDUP(2.01, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(MathRoundUp, NegativeAwayFromZero) {
  // Excel quirk pin: ROUNDUP(-2.01, 0) -> -3 (away from zero, not up).
  const Value v = EvalSource("=ROUNDUP(-2.01, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -3.0);
}

TEST(MathRoundUp, PositiveDigits) {
  const Value v = EvalSource("=ROUNDUP(1.001, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.01, 1e-12);
}

TEST(MathRoundUp, NegativeDigits) {
  const Value v = EvalSource("=ROUNDUP(1201, -2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1300.0);
}

TEST(MathRoundUp, SnapsBinaryRepresentationNoiseLikeRound) {
  // 0.07 * 100 == 7.000000000000001 in IEEE-754 double; without the snap
  // compensation this is the sharpest of the three failures: `ceil` on
  // the unsnapped product rounds a value already exact at 2 digits UP an
  // extra step, so ROUNDUP(0.07, 2) returns 0.08 instead of 0.07.
  const Value v = EvalSource("=ROUNDUP(0.07, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.07);
}

TEST(MathRoundUp, ExtremePositiveDigitsIsNoOp) {
  const Value v = EvalSource("=ROUNDUP(1.5, 400)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.5);
}

TEST(MathRoundUp, ExtremeNegativeDigitsIsZero) {
  const Value v = EvalSource("=ROUNDUP(1.5, -400)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// MIN
// ---------------------------------------------------------------------------

TEST(MathMin, ThreePositives) {
  const Value v = EvalSource("=MIN(3, 1, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(MathMin, MixedNegatives) {
  const Value v = EvalSource("=MIN(-3, -1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -3.0);
}

TEST(MathMin, SingleArgument) {
  const Value v = EvalSource("=MIN(2.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.5);
}

TEST(MathMin, EmptyArgListIsArityViolation) {
  const Value v = EvalSource("=MIN()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(MathMin, NonNumericTextYieldsValue) {
  // Literal text args do NOT get the cell-reference skip rule; they coerce
  // through coerce_to_number and surface #VALUE! on failure.
  const Value v = EvalSource("=MIN(1, \"abc\", 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(MathMin, ErrorPropagates) {
  const Value v = EvalSource("=MIN(1, #REF!, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// MAX
// ---------------------------------------------------------------------------

TEST(MathMax, ThreePositives) {
  const Value v = EvalSource("=MAX(1, 3, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(MathMax, MixedNegatives) {
  const Value v = EvalSource("=MAX(-3, -1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -1.0);
}

TEST(MathMax, SingleArgument) {
  const Value v = EvalSource("=MAX(2.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.5);
}

TEST(MathMax, EmptyArgListIsArityViolation) {
  const Value v = EvalSource("=MAX()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(MathMax, NonNumericTextYieldsValue) {
  const Value v = EvalSource("=MAX(1, \"abc\", 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(MathMax, ErrorPropagates) {
  const Value v = EvalSource("=MAX(1, #REF!, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// AVERAGE
// ---------------------------------------------------------------------------

TEST(MathAverage, FourValues) {
  const Value v = EvalSource("=AVERAGE(1, 2, 3, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.5);
}

TEST(MathAverage, SingleValue) {
  const Value v = EvalSource("=AVERAGE(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(MathAverage, NumericTextCoerces) {
  const Value v = EvalSource("=AVERAGE(1, \"2\", 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(MathAverage, NonNumericTextYieldsValue) {
  const Value v = EvalSource("=AVERAGE(\"abc\", 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(MathAverage, ErrorPropagates) {
  const Value v = EvalSource("=AVERAGE(1, #DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(MathAverage, EmptyArgListIsArityViolation) {
  const Value v = EvalSource("=AVERAGE()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(MathAverage, SumOverflowIsNumError) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::number(1.0e308));
  wb.sheet(0).set_cell_value(1, 0, Value::number(1.0e308));
  const Value average = EvalSourceIn("=AVERAGE(A1:A2)", wb, wb.sheet(0));
  ASSERT_TRUE(average.is_error()) << average.debug_to_string();
  EXPECT_EQ(average.as_error(), ErrorCode::Num);

  const Value direct = EvalSource("=AVERAGE(9E307,9E307)");
  ASSERT_TRUE(direct.is_error()) << direct.debug_to_string();
  EXPECT_EQ(direct.as_error(), ErrorCode::Num);
}

TEST(MathAverage, PlainSumKeepsLastBitAndAbsorption) {
  const Value last_bit = EvalSource("=(AVERAGE(0.1,0.2,0.3)-0.2)*1");
  ASSERT_TRUE(last_bit.is_number()) << last_bit.debug_to_string();
  EXPECT_DOUBLE_EQ(last_bit.as_number(), 2.7755575615628914e-17);

  const Value absorbed = EvalSource("=AVERAGE(1E16,1,-1E16)");
  ASSERT_TRUE(absorbed.is_number()) << absorbed.debug_to_string();
  EXPECT_DOUBLE_EQ(absorbed.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// PRODUCT
// ---------------------------------------------------------------------------

TEST(MathProduct, ThreeValues) {
  const Value v = EvalSource("=PRODUCT(2, 3, 4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 24.0);
}

TEST(MathProduct, SingleValue) {
  const Value v = EvalSource("=PRODUCT(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(MathProduct, ZeroAnnihilates) {
  const Value v = EvalSource("=PRODUCT(0, 100, 200)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(MathProduct, OverflowYieldsNum) {
  const Value v = EvalSource("=PRODUCT(1e200, 1e200)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(MathProduct, EmptyArgListIsArityViolation) {
  const Value v = EvalSource("=PRODUCT()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Analysis-ToolPak argument rule (QUOTIENT / MROUND / GCD)
// ---------------------------------------------------------------------------

TEST(MathAnalysisToolPak, QuotientBoolLiteralIsValue) {
  const Value v = EvalSource("=QUOTIENT(TRUE,1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(MathAnalysisToolPak, QuotientBoolCellIsValue) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::boolean(true));
  const Value v = EvalSourceIn("=QUOTIENT(A1,1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(MathAnalysisToolPak, QuotientOmittedRequiredIsNA) {
  const Value v = EvalSource("=QUOTIENT(,1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(MathAnalysisToolPak, QuotientBlankRefIsZero) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=QUOTIENT(A1,1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(MathAnalysisToolPak, MroundBoolLiteralIsValue) {
  const Value v = EvalSource("=MROUND(TRUE,1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(MathAnalysisToolPak, GcdOmittedIsNA) {
  const Value v = EvalSource("=GCD(,2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::NA);
}

TEST(MathAnalysisToolPak, GcdBoolCellIsValue) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(0, 0, Value::boolean(true));
  const Value v = EvalSourceIn("=GCD(A1,2)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(MathAnalysisToolPak, GcdBlankRefsKeepValueError) {
  Workbook wb = Workbook::create();
  const Value v = EvalSourceIn("=GCD(A1,B1)", wb, wb.sheet(0));
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ROUND's digits snap up to a near integer; 1.999999 is outside 2^-22.
TEST(MathRound, DigitsSnapToNearInteger) {
  const Value snapped = EvalSource("=ROUND(1.23456,1.9999999)");
  ASSERT_TRUE(snapped.is_number());
  EXPECT_DOUBLE_EQ(snapped.as_number(), 1.23);
  const Value truncated = EvalSource("=ROUND(1.23456,1.999999)");
  ASSERT_TRUE(truncated.is_number());
  EXPECT_DOUBLE_EQ(truncated.as_number(), 1.2);
}

TEST(MathAnalysisToolPak, GcdBoolInArrayLiteralIsValue) {
  const Value v = EvalSource("=GCD({TRUE,2})");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
