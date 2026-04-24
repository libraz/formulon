// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the ERF and BESSEL engineering built-ins:
//   * ERF family (1/2-arg ERF, ERF.PRECISE, ERFC, ERFC.PRECISE):
//     canonical values, oddness, 2-arg integral form, arity enforcement,
//     and #VALUE! propagation for non-numeric input.
//   * BESSEL family (BESSELJ / BESSELY / BESSELI / BESSELK): canonical
//     handbook values, order truncation, negative-order rejection, and
//     singularity handling (x <= 0) for the second-kind members.

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

// Parses `src` and evaluates it via the default function registry. Arenas
// are reset between calls to avoid cross-test contamination.
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
// Registry pin -- catches accidental drops / renames during refactors.
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineeringSpecialRegistry, NamesRegistered) {
  const FunctionRegistry& reg = default_registry();
  EXPECT_NE(reg.lookup("ERF"), nullptr);
  EXPECT_NE(reg.lookup("ERF.PRECISE"), nullptr);
  EXPECT_NE(reg.lookup("ERFC"), nullptr);
  EXPECT_NE(reg.lookup("ERFC.PRECISE"), nullptr);
  EXPECT_NE(reg.lookup("BESSELJ"), nullptr);
  EXPECT_NE(reg.lookup("BESSELY"), nullptr);
  EXPECT_NE(reg.lookup("BESSELI"), nullptr);
  EXPECT_NE(reg.lookup("BESSELK"), nullptr);
}

// ---------------------------------------------------------------------------
// ERF
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineeringSpecial, ErfZero) {
  const Value v = EvalSource("=ERF(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 0.0);
}

TEST(BuiltinsEngineeringSpecial, ErfOne) {
  const Value v = EvalSource("=ERF(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.8427007929, 1e-9);
}

TEST(BuiltinsEngineeringSpecial, ErfOdd) {
  // erf(-x) = -erf(x).
  const Value v = EvalSource("=ERF(-1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), -0.8427007929, 1e-9);
}

TEST(BuiltinsEngineeringSpecial, ErfTwoArgFromZero) {
  // Integral from 0 to 1 is erf(1) - erf(0) = erf(1).
  const Value v = EvalSource("=ERF(0,1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.8427007929, 1e-9);
}

TEST(BuiltinsEngineeringSpecial, ErfTwoArgInterval) {
  // erf(2) - erf(1) ~= 0.995322265 - 0.842700793 ~= 0.152621472.
  const Value v = EvalSource("=ERF(1,2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.152621472, 1e-9);
}

TEST(EngineeringErf, RejectsBoolArg) {
  // Excel 365 rejects Bool for the ERF family: ERF(TRUE) -> #VALUE!.
  const Value v = EvalSource("=ERF(TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(BuiltinsEngineeringSpecial, ErfNonNumericReturnsValue) {
  const Value v = EvalSource("=ERF(\"not a number\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(EngineeringErfc, RejectsBoolArg) {
  // Excel 365 rejects Bool for ERFC: ERFC(TRUE) -> #VALUE!.
  const Value v = EvalSource("=ERFC(TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(EngineeringBesselI, RejectsBoolX) {
  // Excel 365 rejects Bool as the `x` argument of BESSELI.
  const Value v = EvalSource("=BESSELI(TRUE, 1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(EngineeringBesselJ, RejectsBoolOrder) {
  // Excel 365 rejects Bool as the `n` (order) argument of BESSELJ.
  const Value v = EvalSource("=BESSELJ(1, TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// ERF.PRECISE
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineeringSpecial, ErfPreciseOne) {
  const Value v = EvalSource("=ERF.PRECISE(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.8427007929, 1e-9);
}

TEST(BuiltinsEngineeringSpecial, ErfPreciseRejectsExtraArg) {
  // Arity is fixed at 1; the registry dispatcher surfaces #VALUE! on
  // arity mismatch (Formulon's engine-level convention; Excel surfaces
  // the same error surface visually as "too many arguments").
  const Value v = EvalSource("=ERF.PRECISE(1,2)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// ERFC / ERFC.PRECISE
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineeringSpecial, ErfcZero) {
  const Value v = EvalSource("=ERFC(0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.0);
}

TEST(BuiltinsEngineeringSpecial, ErfcOne) {
  // erfc(1) = 1 - erf(1) ~= 0.157299207.
  const Value v = EvalSource("=ERFC(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.157299207, 1e-9);
}

TEST(BuiltinsEngineeringSpecial, ErfcPreciseMatchesErfc) {
  const Value v = EvalSource("=ERFC.PRECISE(1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.157299207, 1e-9);
}

// ---------------------------------------------------------------------------
// BESSELJ (POSIX jn)
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineeringSpecial, BesselJZeroZero) {
  // J_0(0) = 1.
  const Value v = EvalSource("=BESSELJ(0,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsEngineeringSpecial, BesselJOneZero) {
  // J_0(1) = 0.765197686557966551...
  const Value v = EvalSource("=BESSELJ(1,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.7651976866, 1e-8);
}

TEST(BuiltinsEngineeringSpecial, BesselJZeroOne) {
  // J_1(0) = 0.
  const Value v = EvalSource("=BESSELJ(0,1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0, 1e-12);
}

TEST(BuiltinsEngineeringSpecial, BesselJOneOne) {
  // J_1(1) = 0.440050585744933516...
  const Value v = EvalSource("=BESSELJ(1,1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.4400505857, 1e-8);
}

TEST(BuiltinsEngineeringSpecial, BesselJNegativeOrderRejected) {
  const Value v = EvalSource("=BESSELJ(1,-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineeringSpecial, BesselJOrderTruncates) {
  // 2.7 truncates to 2; should match BESSELJ(1, 2).
  const Value v1 = EvalSource("=BESSELJ(1,2.7)");
  const Value v2 = EvalSource("=BESSELJ(1,2)");
  ASSERT_TRUE(v1.is_number());
  ASSERT_TRUE(v2.is_number());
  EXPECT_NEAR(v1.as_number(), v2.as_number(), 1e-12);
}

// ---------------------------------------------------------------------------
// BESSELY (POSIX yn)
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineeringSpecial, BesselYOneZero) {
  // Y_0(1) = 0.088256964215676957...
  const Value v = EvalSource("=BESSELY(1,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0882569642, 1e-8);
}

TEST(BuiltinsEngineeringSpecial, BesselYZeroIsNum) {
  // Y_n is singular at x=0.
  const Value v = EvalSource("=BESSELY(0,0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineeringSpecial, BesselYNegativeXIsNum) {
  // Excel rejects negative x for the second-kind Bessel functions.
  const Value v = EvalSource("=BESSELY(-1,0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// BESSELI (power series)
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineeringSpecial, BesselIZeroZero) {
  // I_0(0) = 1.
  const Value v = EvalSource("=BESSELI(0,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.0, 1e-12);
}

TEST(BuiltinsEngineeringSpecial, BesselIOneZero) {
  // I_0(1) = 1.266065877752008336...
  const Value v = EvalSource("=BESSELI(1,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.2660658733, 1e-8);
}

TEST(BuiltinsEngineeringSpecial, BesselITwoOne) {
  // I_1(2) = 1.590636854637328814...
  const Value v = EvalSource("=BESSELI(2,1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 1.5906368546, 1e-8);
}

TEST(BuiltinsEngineeringSpecial, BesselIZeroOne) {
  // I_1(0) = 0.
  const Value v = EvalSource("=BESSELI(0,1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.0, 1e-12);
}

TEST(BuiltinsEngineeringSpecial, BesselINegativeOrderRejected) {
  const Value v = EvalSource("=BESSELI(1,-1)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

// ---------------------------------------------------------------------------
// BESSELK (A&S approximations)
// ---------------------------------------------------------------------------

TEST(BuiltinsEngineeringSpecial, BesselKOneZero) {
  // K_0(1) = 0.421024438240708332...
  // A&S 9.8.5-9.8.6 approximations are rated |epsilon| < ~1.9e-7; real
  // accuracy is around 5e-8 on modern hardware, so 1e-6 tolerance is
  // conservative.
  const Value v = EvalSource("=BESSELK(1,0)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.4210244382, 1e-6);
}

TEST(BuiltinsEngineeringSpecial, BesselKTwoOne) {
  // K_1(2) = 0.139865881816522427...
  const Value v = EvalSource("=BESSELK(2,1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.1398658818, 1e-6);
}

TEST(BuiltinsEngineeringSpecial, BesselKOneOne) {
  // K_1(1) ~= 0.6019072302 (cross-check against handbook).
  const Value v = EvalSource("=BESSELK(1,1)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.6019072302, 1e-6);
}

TEST(BuiltinsEngineeringSpecial, BesselKRecurrenceTwoTwo) {
  // K_2(2) ~= 0.253759755 (Wolfram reference). Exercises the forward
  // recurrence path (n >= 2).
  const Value v = EvalSource("=BESSELK(2,2)");
  ASSERT_TRUE(v.is_number());
  EXPECT_NEAR(v.as_number(), 0.2537597545, 1e-6);
}

TEST(BuiltinsEngineeringSpecial, BesselKZeroIsNum) {
  // K_n is singular at x=0.
  const Value v = EvalSource("=BESSELK(0,0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

TEST(BuiltinsEngineeringSpecial, BesselKNegativeXIsNum) {
  // Excel rejects negative x for the second-kind modified Bessel.
  const Value v = EvalSource("=BESSELK(-1,0)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Num);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
