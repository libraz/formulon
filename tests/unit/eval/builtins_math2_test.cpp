// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the exponential / logarithmic / trigonometric math
// built-ins: EXP, LN, LOG, LOG10, PI, RADIANS, DEGREES, SIN, COS, TAN,
// ASIN, ACOS, ATAN, ATAN2. Each test parses a formula source, evaluates
// the AST through the default registry, and asserts the resulting Value.

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

// Local copy of the same constant used inside builtins.cpp. Tests should
// never reach across translation units for an internal-linkage constant;
// reproducing the value here keeps the test self-contained.
constexpr double kPi = 3.14159265358979323846;

// Parses `src` and evaluates it via the default function registry. The
// thread-local arenas keep text payloads readable for the immediately
// following EXPECT_*. Each call resets the arenas to avoid cross-test
// contamination.
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

// Invokes a registered function impl directly. The parser now rewrites
// CellRef-shaped tokens to Ident when immediately followed by `(`, so
// `LOG10(x)` reaches the function-call path through formula syntax. The
// direct-invocation helper is retained for cases where the formula-source
// path cannot synthesise the desired argument shape (for example, Blank
// values per `builtins_info_test.cpp`).
Value CallRegistered(std::string_view name, const Value* args, std::uint32_t arity) {
  static thread_local Arena arena;
  arena.reset();
  const FunctionDef* def = default_registry().lookup(name);
  EXPECT_NE(def, nullptr) << "function not registered: " << name;
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return def->impl(args, arity, arena);
}

// ---------------------------------------------------------------------------
// EXP
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Exp, ZeroIsOne) {
  const Value v = EvalSource("=EXP(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath2Exp, OneIsE) {
  const Value v = EvalSource("=EXP(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.71828182845904523536, 1e-12);
}

TEST(BuiltinsMath2Exp, NegativeArgument) {
  const Value v = EvalSource("=EXP(-1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0 / 2.71828182845904523536, 1e-12);
}

TEST(BuiltinsMath2Exp, OverflowYieldsNum) {
  // EXP(1000) overflows to +Inf in double precision; finite check trips.
  const Value v = EvalSource("=EXP(1000)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Exp, TextCoerces) {
  const Value v = EvalSource("=EXP(\"1\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.71828182845904523536, 1e-12);
}

TEST(BuiltinsMath2Exp, ErrorPropagates) {
  const Value v = EvalSource("=EXP(#DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMath2Exp, ZeroArgsIsArityViolation) {
  const Value v = EvalSource("=EXP()");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMath2Exp, TwoArgsIsArityViolation) {
  const Value v = EvalSource("=EXP(1, 2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// LN
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Ln, OneIsZero) {
  const Value v = EvalSource("=LN(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath2Ln, EIsOne) {
  const Value v = EvalSource("=LN(EXP(1))");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsMath2Ln, ZeroYieldsNum) {
  const Value v = EvalSource("=LN(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Ln, NegativeYieldsNum) {
  const Value v = EvalSource("=LN(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Ln, ErrorPropagates) {
  const Value v = EvalSource("=LN(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// LOG (default base 10; quirky DIV/0 when base == 1)
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Log, DefaultBaseIsTen) {
  const Value v = EvalSource("=LOG(100)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.0, 1e-12);
}

TEST(BuiltinsMath2Log, ExplicitBaseTen) {
  const Value v = EvalSource("=LOG(1000, 10)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 3.0, 1e-12);
}

TEST(BuiltinsMath2Log, BaseTwo) {
  const Value v = EvalSource("=LOG(8, 2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 3.0, 1e-12);
}

TEST(BuiltinsMath2Log, OneOfOneIsZero) {
  // Numerator zero with non-trivial base - LOG(1, 5) = 0.
  const Value v = EvalSource("=LOG(1, 5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath2Log, ZeroXYieldsNum) {
  const Value v = EvalSource("=LOG(0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Log, NegativeXYieldsNum) {
  const Value v = EvalSource("=LOG(-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Log, NegativeBaseYieldsNum) {
  const Value v = EvalSource("=LOG(10, -2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Log, ZeroBaseYieldsNum) {
  const Value v = EvalSource("=LOG(10, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Log, BaseOneYieldsDiv0) {
  // Excel quirk pin: ln(1) == 0 sits in the divisor, so LOG(x, 1) is
  // #DIV/0! rather than #NUM!.
  const Value v = EvalSource("=LOG(10, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMath2Log, ErrorPropagates) {
  const Value v = EvalSource("=LOG(#NAME?)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsMath2Log, ErrorInBasePropagates) {
  const Value v = EvalSource("=LOG(10, #REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// LOG10
//
// `LOG10` lexes as a CellRef (column "LOG", row 10) because the cellref
// scanner accepts any 1-3 letter run followed by digits. The parser now
// disambiguates this at the token-stream stage by rewriting CellRef to
// Ident when immediately followed by `(`, so `LOG10(x)` reaches the
// function-call path through formula syntax. The first two tests below
// exercise that path end-to-end; the remaining tests use the
// `CallRegistered` helper because the math itself is the focus.
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Log10, HundredIsTwo) {
  const Value v = EvalSource("=LOG10(100)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.0, 1e-12);
}

TEST(BuiltinsMath2Log10, OneIsZero) {
  const Value v = EvalSource("=LOG10(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath2Log10, ZeroYieldsNum) {
  const Value arg = Value::number(0.0);
  const Value v = CallRegistered("LOG10", &arg, 1u);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Log10, NegativeYieldsNum) {
  const Value arg = Value::number(-10.0);
  const Value v = CallRegistered("LOG10", &arg, 1u);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Log10, ThousandIsThree) {
  const Value arg = Value::number(1000.0);
  const Value v = CallRegistered("LOG10", &arg, 1u);
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 3.0, 1e-12);
}

TEST(BuiltinsMath2Log10, RegisteredInDefaultRegistry) {
  // Defensive: confirm LOG10 is reachable by name in the default registry
  // (mirrors what `CallRegistered` checks per call, but as an explicit pin
  // in case future tests stop exercising the helper).
  EXPECT_NE(default_registry().lookup("LOG10"), nullptr);
}

// ---------------------------------------------------------------------------
// PI
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Pi, ReturnsPi) {
  const Value v = EvalSource("=PI()");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 3.14159265358979, 1e-13);
}

TEST(BuiltinsMath2Pi, OneArgIsArityViolation) {
  const Value v = EvalSource("=PI(1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// RADIANS
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Radians, ZeroIsZero) {
  const Value v = EvalSource("=RADIANS(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath2Radians, OneEightyIsPi) {
  const Value v = EvalSource("=RADIANS(180)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi, 1e-12);
}

TEST(BuiltinsMath2Radians, NinetyIsHalfPi) {
  const Value v = EvalSource("=RADIANS(90)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi / 2.0, 1e-12);
}

TEST(BuiltinsMath2Radians, NegativeAngle) {
  const Value v = EvalSource("=RADIANS(-180)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -kPi, 1e-12);
}

TEST(BuiltinsMath2Radians, BoolTrueCoerces) {
  // TRUE -> 1 degree -> pi/180 radians.
  const Value v = EvalSource("=RADIANS(TRUE())");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi / 180.0, 1e-12);
}

TEST(BuiltinsMath2Radians, ErrorPropagates) {
  const Value v = EvalSource("=RADIANS(#NAME?)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// DEGREES
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Degrees, ZeroIsZero) {
  const Value v = EvalSource("=DEGREES(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath2Degrees, PiIsOneEighty) {
  const Value v = EvalSource("=DEGREES(PI())");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 180.0, 1e-12);
}

TEST(BuiltinsMath2Degrees, HalfPiIsNinety) {
  const Value v = EvalSource("=DEGREES(PI()/2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 90.0, 1e-12);
}

TEST(BuiltinsMath2Degrees, NegativeRadians) {
  const Value v = EvalSource("=DEGREES(-PI())");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -180.0, 1e-12);
}

TEST(BuiltinsMath2Degrees, ErrorPropagates) {
  const Value v = EvalSource("=DEGREES(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// SIN
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Sin, ZeroIsZero) {
  const Value v = EvalSource("=SIN(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath2Sin, HalfPiIsOne) {
  const Value v = EvalSource("=SIN(PI()/2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsMath2Sin, PiIsAlmostZero) {
  const Value v = EvalSource("=SIN(PI())");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0, 1e-12);
}

TEST(BuiltinsMath2Sin, BoolTrueIsSinOne) {
  // TRUE -> 1 radian.
  const Value v = EvalSource("=SIN(TRUE())");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.8414709848078965, 1e-12);
}

TEST(BuiltinsMath2Sin, ErrorPropagates) {
  const Value v = EvalSource("=SIN(#DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// COS
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Cos, ZeroIsOne) {
  const Value v = EvalSource("=COS(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsMath2Cos, HalfPiIsAlmostZero) {
  const Value v = EvalSource("=COS(PI()/2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0, 1e-12);
}

TEST(BuiltinsMath2Cos, PiIsMinusOne) {
  const Value v = EvalSource("=COS(PI())");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -1.0, 1e-12);
}

TEST(BuiltinsMath2Cos, ErrorPropagates) {
  const Value v = EvalSource("=COS(#NAME?)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// TAN
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Tan, ZeroIsZero) {
  const Value v = EvalSource("=TAN(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath2Tan, QuarterPiIsOne) {
  const Value v = EvalSource("=TAN(PI()/4)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsMath2Tan, HalfPiIsLargeButFinite) {
  // Excel quirk pin: TAN(PI/2) does NOT yield #NUM!. Because PI/2 in
  // double precision differs slightly from the true pole, std::tan returns
  // a huge but finite value.
  const Value v = EvalSource("=TAN(PI()/2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_TRUE(std::isfinite(v.as_number()));
}

TEST(BuiltinsMath2Tan, ErrorPropagates) {
  const Value v = EvalSource("=TAN(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// ASIN
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Asin, ZeroIsZero) {
  const Value v = EvalSource("=ASIN(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath2Asin, OneIsHalfPi) {
  const Value v = EvalSource("=ASIN(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi / 2.0, 1e-12);
}

TEST(BuiltinsMath2Asin, MinusOneIsMinusHalfPi) {
  const Value v = EvalSource("=ASIN(-1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -kPi / 2.0, 1e-12);
}

TEST(BuiltinsMath2Asin, AboveOneYieldsNum) {
  const Value v = EvalSource("=ASIN(1.0001)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Asin, BelowMinusOneYieldsNum) {
  const Value v = EvalSource("=ASIN(-1.0001)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Asin, ErrorPropagates) {
  const Value v = EvalSource("=ASIN(#NAME?)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

// ---------------------------------------------------------------------------
// ACOS
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Acos, OneIsZero) {
  const Value v = EvalSource("=ACOS(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath2Acos, MinusOneIsPi) {
  const Value v = EvalSource("=ACOS(-1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi, 1e-12);
}

TEST(BuiltinsMath2Acos, ZeroIsHalfPi) {
  const Value v = EvalSource("=ACOS(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi / 2.0, 1e-12);
}

TEST(BuiltinsMath2Acos, AboveOneYieldsNum) {
  const Value v = EvalSource("=ACOS(1.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Acos, BelowMinusOneYieldsNum) {
  const Value v = EvalSource("=ACOS(-1.5)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsMath2Acos, ErrorPropagates) {
  const Value v = EvalSource("=ACOS(#DIV/0!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

// ---------------------------------------------------------------------------
// ATAN
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Atan, ZeroIsZero) {
  const Value v = EvalSource("=ATAN(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath2Atan, OneIsQuarterPi) {
  const Value v = EvalSource("=ATAN(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi / 4.0, 1e-12);
}

TEST(BuiltinsMath2Atan, MinusOneIsMinusQuarterPi) {
  const Value v = EvalSource("=ATAN(-1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -kPi / 4.0, 1e-12);
}

TEST(BuiltinsMath2Atan, LargeArgApproachesHalfPi) {
  const Value v = EvalSource("=ATAN(1e15)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi / 2.0, 1e-10);
}

TEST(BuiltinsMath2Atan, ErrorPropagates) {
  const Value v = EvalSource("=ATAN(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// ATAN2 - argument order is (x, y), opposite of C's atan2(y, x)
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Atan2, OneOneIsQuarterPi) {
  // (1, 1) is 45 degrees in the first quadrant.
  const Value v = EvalSource("=ATAN2(1, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi / 4.0, 1e-12);
}

TEST(BuiltinsMath2Atan2, ZeroOneIsHalfPi) {
  // Excel quirk pin: ATAN2(0, 1) treats x=0, y=1 - the point on the
  // positive y-axis - which is 90 degrees, NOT 0. This is the canonical
  // proof that the argument order is (x, y).
  const Value v = EvalSource("=ATAN2(0, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi / 2.0, 1e-12);
}

TEST(BuiltinsMath2Atan2, OneZeroIsZero) {
  // (1, 0) is the positive x-axis.
  const Value v = EvalSource("=ATAN2(1, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsMath2Atan2, MinusOneZeroIsPi) {
  // (-1, 0) is the negative x-axis -> +pi.
  const Value v = EvalSource("=ATAN2(-1, 0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), kPi, 1e-12);
}

TEST(BuiltinsMath2Atan2, ZeroMinusOneIsMinusHalfPi) {
  // (0, -1) is the negative y-axis -> -pi/2.
  const Value v = EvalSource("=ATAN2(0, -1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -kPi / 2.0, 1e-12);
}

TEST(BuiltinsMath2Atan2, BothZeroYieldsDiv0) {
  // Excel quirk pin: ATAN2(0, 0) is #DIV/0!, even though C's atan2(0, 0)
  // returns 0.
  const Value v = EvalSource("=ATAN2(0, 0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Div0);
}

TEST(BuiltinsMath2Atan2, SecondQuadrant) {
  // x = -1, y = 1 -> 3pi/4.
  const Value v = EvalSource("=ATAN2(-1, 1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 3.0 * kPi / 4.0, 1e-12);
}

TEST(BuiltinsMath2Atan2, FourthQuadrant) {
  // x = 1, y = -1 -> -pi/4.
  const Value v = EvalSource("=ATAN2(1, -1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -kPi / 4.0, 1e-12);
}

TEST(BuiltinsMath2Atan2, ErrorInXPropagates) {
  const Value v = EvalSource("=ATAN2(#NAME?, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Name);
}

TEST(BuiltinsMath2Atan2, ErrorInYPropagates) {
  const Value v = EvalSource("=ATAN2(1, #REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(BuiltinsMath2Atan2, OneArgIsArityViolation) {
  const Value v = EvalSource("=ATAN2(1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsMath2Atan2, ThreeArgsIsArityViolation) {
  const Value v = EvalSource("=ATAN2(1, 2, 3)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Cross-function identities (light sanity checks across the family)
// ---------------------------------------------------------------------------

TEST(BuiltinsMath2Identities, RadiansDegreesRoundTrip) {
  const Value v = EvalSource("=DEGREES(RADIANS(45))");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 45.0, 1e-12);
}

TEST(BuiltinsMath2Identities, ExpLnRoundTrip) {
  const Value v = EvalSource("=LN(EXP(2))");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 2.0, 1e-12);
}

TEST(BuiltinsMath2Identities, AsinSinRoundTrip) {
  const Value v = EvalSource("=ASIN(SIN(0.5))");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

TEST(BuiltinsMath2Identities, AcosCosRoundTrip) {
  const Value v = EvalSource("=ACOS(COS(0.5))");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

TEST(BuiltinsMath2Identities, AtanTanRoundTrip) {
  const Value v = EvalSource("=ATAN(TAN(0.5))");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.5, 1e-12);
}

TEST(BuiltinsMath2Identities, SinSquaredPlusCosSquared) {
  const Value v = EvalSource("=SIN(0.7)^2+COS(0.7)^2");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
