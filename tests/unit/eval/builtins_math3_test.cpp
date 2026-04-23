// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the hyperbolic / reciprocal-trig / parity-rounding
// math built-ins: SINH, COSH, TANH, ASINH, ACOSH, ATANH, SEC, CSC, COT,
// ACOT, SECH, CSCH, COTH, ACOTH, EVEN, ODD, QUOTIENT. Each test parses a
// formula source, evaluates the AST through the default registry, and
// asserts the resulting Value. Reference values are computed from the
// mathematical definitions (SINH(1) = (e - 1/e) / 2, etc.) to keep these
// tests independent of any golden oracle fixture.

#include <cmath>
#include <string_view>

#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Local pi copy; see `builtins_math2_test.cpp` for the rationale.
constexpr double kPi = 3.14159265358979323846;

// Parses `src` and evaluates it via the default function registry. Arenas
// are thread-local and reset per call.
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
// SINH
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Sinh, ZeroIsZero) {
  const Value v = EvalSource("=SINH(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath3Sinh, OneIsHalfEMinusOneOverE) {
  const Value v = EvalSource("=SINH(1)");
  ASSERT_TRUE(v.is_number());
  const double expected = (std::exp(1.0) - std::exp(-1.0)) / 2.0;
  EXPECT_NEAR(v.as_number(), expected, 1e-12);
}

TEST(BuiltinsMath3Sinh, NegativeIsAntisymmetric) {
  const Value v = EvalSource("=SINH(-1)");
  ASSERT_TRUE(v.is_number());
  const double expected = -(std::exp(1.0) - std::exp(-1.0)) / 2.0;
  EXPECT_NEAR(v.as_number(), expected, 1e-12);
}

TEST(BuiltinsMath3Sinh, OverflowYieldsNum) {
  // std::sinh(1000) = +Inf in double precision; caught by the finite check.
  const Value v = EvalSource("=SINH(1000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath3Sinh, ErrorPropagates) {
  const Value v = EvalSource("=SINH(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// COSH
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Cosh, ZeroIsOne) {
  const Value v = EvalSource("=COSH(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath3Cosh, SymmetricAroundZero) {
  const Value a = EvalSource("=COSH(1)");
  const Value b = EvalSource("=COSH(-1)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

TEST(BuiltinsMath3Cosh, OneIsHalfEPlusOneOverE) {
  const Value v = EvalSource("=COSH(1)");
  ASSERT_TRUE(v.is_number());
  const double expected = (std::exp(1.0) + std::exp(-1.0)) / 2.0;
  EXPECT_NEAR(v.as_number(), expected, 1e-12);
}

TEST(BuiltinsMath3Cosh, OverflowYieldsNum) {
  const Value v = EvalSource("=COSH(1000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// TANH
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Tanh, ZeroIsZero) {
  const Value v = EvalSource("=TANH(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath3Tanh, LargeApproachesOne) {
  const Value v = EvalSource("=TANH(100)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsMath3Tanh, NegativeLargeApproachesMinusOne) {
  const Value v = EvalSource("=TANH(-100)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -1.0, 1e-12);
}

TEST(BuiltinsMath3Tanh, OneIsKnownValue) {
  const Value v = EvalSource("=TANH(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::tanh(1.0), 1e-12);
}

// ---------------------------------------------------------------------------
// ASINH
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Asinh, ZeroIsZero) {
  const Value v = EvalSource("=ASINH(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath3Asinh, SinhRoundTrip) {
  const Value v = EvalSource("=ASINH(SINH(0.75))");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.75, 1e-12);
}

TEST(BuiltinsMath3Asinh, NegativeIsAntisymmetric) {
  const Value a = EvalSource("=ASINH(1)");
  const Value b = EvalSource("=ASINH(-1)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), -b.as_number(), 1e-12);
}

// ---------------------------------------------------------------------------
// ACOSH
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Acosh, OneIsZero) {
  const Value v = EvalSource("=ACOSH(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath3Acosh, TwoMatchesStdAcosh) {
  const Value v = EvalSource("=ACOSH(2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::acosh(2.0), 1e-12);
}

TEST(BuiltinsMath3Acosh, BelowOneYieldsNum) {
  const Value v = EvalSource("=ACOSH(0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath3Acosh, NegativeYieldsNum) {
  const Value v = EvalSource("=ACOSH(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath3Acosh, ErrorPropagates) {
  const Value v = EvalSource("=ACOSH(#DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// ATANH
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Atanh, ZeroIsZero) {
  const Value v = EvalSource("=ATANH(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath3Atanh, HalfMatchesStdAtanh) {
  const Value v = EvalSource("=ATANH(0.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::atanh(0.5), 1e-12);
}

TEST(BuiltinsMath3Atanh, OneYieldsNum) {
  // Excel quirk pin: |x| == 1 is OUTSIDE the open interval and yields #NUM!.
  const Value v = EvalSource("=ATANH(1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath3Atanh, MinusOneYieldsNum) {
  const Value v = EvalSource("=ATANH(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath3Atanh, AboveOneYieldsNum) {
  const Value v = EvalSource("=ATANH(2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// SEC
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Sec, ZeroIsOne) {
  const Value v = EvalSource("=SEC(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath3Sec, QuarterPiIsSqrtTwo) {
  const Value v = EvalSource("=SEC(PI()/4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::sqrt(2.0), 1e-12);
}

TEST(BuiltinsMath3Sec, PiIsMinusOne) {
  const Value v = EvalSource("=SEC(PI())");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -1.0, 1e-12);
}

TEST(BuiltinsMath3Sec, HalfPiIsLargeButFinite) {
  // COS(PI/2) in double precision is a tiny non-zero value, so SEC is
  // finite; this mirrors the TAN(PI/2) behaviour.
  const Value v = EvalSource("=SEC(PI()/2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
}

// ---------------------------------------------------------------------------
// CSC
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Csc, ZeroYieldsDiv0) {
  const Value v = EvalSource("=CSC(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMath3Csc, HalfPiIsOne) {
  const Value v = EvalSource("=CSC(PI()/2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsMath3Csc, QuarterPiIsSqrtTwo) {
  const Value v = EvalSource("=CSC(PI()/4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::sqrt(2.0), 1e-12);
}

// ---------------------------------------------------------------------------
// COT
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Cot, ZeroYieldsDiv0) {
  const Value v = EvalSource("=COT(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMath3Cot, QuarterPiIsOne) {
  const Value v = EvalSource("=COT(PI()/4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsMath3Cot, NegativeArgIsAntisymmetric) {
  const Value a = EvalSource("=COT(PI()/4)");
  const Value b = EvalSource("=COT(-PI()/4)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_NEAR(a.as_number(), -b.as_number(), 1e-12);
}

// ---------------------------------------------------------------------------
// ACOT
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Acot, ZeroIsHalfPi) {
  const Value v = EvalSource("=ACOT(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi / 2.0, 1e-12);
}

TEST(BuiltinsMath3Acot, OneIsQuarterPi) {
  const Value v = EvalSource("=ACOT(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi / 4.0, 1e-12);
}

TEST(BuiltinsMath3Acot, NegativeArgInUpperHalfPlane) {
  // ACOT range is (0, PI), so ACOT(-1) > PI/2 (not -PI/4!).
  const Value v = EvalSource("=ACOT(-1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 3.0 * kPi / 4.0, 1e-12);
}

// ---------------------------------------------------------------------------
// SECH
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Sech, ZeroIsOne) {
  const Value v = EvalSource("=SECH(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath3Sech, LargeArgIsZero) {
  // cosh overflows beyond ~710; Excel (and this impl) yield 0 there.
  const Value v = EvalSource("=SECH(1000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath3Sech, SymmetricAroundZero) {
  const Value a = EvalSource("=SECH(2)");
  const Value b = EvalSource("=SECH(-2)");
  ASSERT_TRUE(a.is_number());
  ASSERT_TRUE(b.is_number());
  EXPECT_DOUBLE_EQ(a.as_number(), b.as_number());
}

// ---------------------------------------------------------------------------
// CSCH
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Csch, ZeroYieldsDiv0) {
  const Value v = EvalSource("=CSCH(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMath3Csch, OneMatchesOneOverSinhOne) {
  const Value v = EvalSource("=CSCH(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0 / std::sinh(1.0), 1e-12);
}

TEST(BuiltinsMath3Csch, LargeArgIsZero) {
  const Value v = EvalSource("=CSCH(1000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

// ---------------------------------------------------------------------------
// COTH
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Coth, ZeroYieldsDiv0) {
  const Value v = EvalSource("=COTH(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMath3Coth, OneMatchesCoshOneOverSinhOne) {
  const Value v = EvalSource("=COTH(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::cosh(1.0) / std::sinh(1.0), 1e-12);
}

TEST(BuiltinsMath3Coth, LargePositiveApproachesOne) {
  const Value v = EvalSource("=COTH(1000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath3Coth, LargeNegativeApproachesMinusOne) {
  const Value v = EvalSource("=COTH(-1000)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -1.0);
}

// ---------------------------------------------------------------------------
// ACOTH
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Acoth, TwoMatchesAtanhOfHalf) {
  const Value v = EvalSource("=ACOTH(2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), std::atanh(0.5), 1e-12);
}

TEST(BuiltinsMath3Acoth, OneYieldsNum) {
  // Excel quirk pin: endpoint is NOT in the domain.
  const Value v = EvalSource("=ACOTH(1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath3Acoth, MinusOneYieldsNum) {
  const Value v = EvalSource("=ACOTH(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath3Acoth, ZeroYieldsNum) {
  const Value v = EvalSource("=ACOTH(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath3Acoth, InsideInclusiveIntervalYieldsNum) {
  const Value v = EvalSource("=ACOTH(0.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// EVEN
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Even, PositiveFractional) {
  const Value v = EvalSource("=EVEN(1.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMath3Even, PositiveOddIntegerRoundsUp) {
  const Value v = EvalSource("=EVEN(3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsMath3Even, PositiveEvenIntegerUnchanged) {
  const Value v = EvalSource("=EVEN(4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 4.0);
}

TEST(BuiltinsMath3Even, NegativeFractional) {
  // Excel quirk pin: `EVEN(-1.5) = -2` (AWAY from zero).
  const Value v = EvalSource("=EVEN(-1.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -2.0);
}

TEST(BuiltinsMath3Even, NegativeFractionalWithOddCeil) {
  // `-2.1` rounds away from zero to `-3` (odd), then steps to `-4`.
  const Value v = EvalSource("=EVEN(-2.1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -4.0);
}

TEST(BuiltinsMath3Even, ZeroIsZero) {
  const Value v = EvalSource("=EVEN(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath3Even, ErrorPropagates) {
  const Value v = EvalSource("=EVEN(#NAME?)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// ODD
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Odd, PositiveFractional) {
  const Value v = EvalSource("=ODD(1.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsMath3Odd, PositiveEvenIntegerRoundsUp) {
  const Value v = EvalSource("=ODD(2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsMath3Odd, PositiveOddIntegerUnchanged) {
  const Value v = EvalSource("=ODD(3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsMath3Odd, NegativeFractional) {
  const Value v = EvalSource("=ODD(-1.5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -3.0);
}

TEST(BuiltinsMath3Odd, NegativeEvenIntegerRoundsDown) {
  const Value v = EvalSource("=ODD(-2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -3.0);
}

TEST(BuiltinsMath3Odd, ZeroIsOne) {
  // Excel quirk pin: `ODD(0) = 1`, not 0. No odd value has magnitude 0, so
  // the away-from-zero rule promotes to +1.
  const Value v = EvalSource("=ODD(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath3Odd, ErrorPropagates) {
  const Value v = EvalSource("=ODD(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// QUOTIENT
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Quotient, PositiveOverPositive) {
  const Value v = EvalSource("=QUOTIENT(10, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsMath3Quotient, NegativeOverPositiveTruncatesTowardZero) {
  // Excel quirk pin: `QUOTIENT(-10, 3) = -3`, NOT -4 (floor-division).
  const Value v = EvalSource("=QUOTIENT(-10, 3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -3.0);
}

TEST(BuiltinsMath3Quotient, PositiveOverNegative) {
  const Value v = EvalSource("=QUOTIENT(10, -3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), -3.0);
}

TEST(BuiltinsMath3Quotient, BothNegative) {
  const Value v = EvalSource("=QUOTIENT(-10, -3)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 3.0);
}

TEST(BuiltinsMath3Quotient, FractionalTruncatesNumerator) {
  // `5.9 / 2 = 2.95`; trunc -> 2.
  const Value v = EvalSource("=QUOTIENT(5.9, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 2.0);
}

TEST(BuiltinsMath3Quotient, ZeroDenominatorYieldsDiv0) {
  const Value v = EvalSource("=QUOTIENT(5, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMath3Quotient, ErrorInNumeratorPropagates) {
  const Value v = EvalSource("=QUOTIENT(#VALUE!, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMath3Quotient, ErrorInDenominatorPropagates) {
  const Value v = EvalSource("=QUOTIENT(10, #REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsMath3Quotient, OneArgIsArityViolation) {
  const Value v = EvalSource("=QUOTIENT(10)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Registration sanity: every new name must be reachable via the registry.
// ---------------------------------------------------------------------------

TEST(BuiltinsMath3Registration, AllNamesRegistered) {
  const FunctionRegistry& reg = default_registry();
  for (const char* name : {"SINH",  "COSH",  "TANH",  "ASINH", "ACOSH",    "ATANH", "SEC",
                           "CSC",   "COT",   "ACOT",  "SECH",  "CSCH",     "COTH",  "ACOTH",
                           "EVEN",  "ODD",   "QUOTIENT"}) {
    EXPECT_NE(reg.lookup(name), nullptr) << "not registered: " << name;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
