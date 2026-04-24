// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the CONVERT built-in: canonical factors for each
// category, SI / binary prefix resolution, affine temperature conversion,
// and the error-surface (unknown unit / mismatched category / unsupported
// prefix). The oracle fixture `ironcalc/calc_tests/CONVERT.xlsx` exercises
// the tables more exhaustively; these tests are the fast-feedback safety
// net that stays green even when the oracle target is unavailable.

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

TEST(BuiltinsConvertRegistry, NameRegistered) {
  const FunctionRegistry& reg = default_registry();
  EXPECT_NE(reg.lookup("CONVERT"), nullptr);
}

// ---------------------------------------------------------------------------
// Linear categories
// ---------------------------------------------------------------------------

TEST(BuiltinsConvert, DistanceSimpleIdentity) {
  const Value r = EvalSource("=CONVERT(1, \"m\", \"m\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_DOUBLE_EQ(r.as_number(), 1.0);
}

TEST(BuiltinsConvert, DistanceMiToKm) {
  // 1 mi = 1609.344 m; to km -> divide by 1000 = 1.609344.
  const Value r = EvalSource("=CONVERT(1, \"mi\", \"km\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_NEAR(r.as_number(), 1.609344, 1e-12);
}

TEST(BuiltinsConvert, WeightLbmToGram) {
  const Value r = EvalSource("=CONVERT(1, \"lbm\", \"g\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_DOUBLE_EQ(r.as_number(), 453.59237);
}

TEST(BuiltinsConvert, EnergyCalToJoule) {
  const Value r = EvalSource("=CONVERT(1, \"cal\", \"J\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_DOUBLE_EQ(r.as_number(), 4.1868);
}

TEST(BuiltinsConvert, PowerHorsepowerHIsHorsepower) {
  // `h` in the Power category is horsepower, not hecto-; the exact-name
  // lookup must win over any prefix parse attempt.
  const Value r = EvalSource("=CONVERT(1, \"h\", \"W\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_NEAR(r.as_number(), 745.6998715822702, 1e-9);
}

// ---------------------------------------------------------------------------
// SI prefixes (linear and with dimensional exponentiation)
// ---------------------------------------------------------------------------

TEST(BuiltinsConvert, SiPrefixKilogramToGram) {
  const Value r = EvalSource("=CONVERT(1, \"kg\", \"g\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_DOUBLE_EQ(r.as_number(), 1000.0);
}

TEST(BuiltinsConvert, SiPrefixDekaMeterToMeter) {
  // `da` is the 2-character deca prefix.
  const Value r = EvalSource("=CONVERT(1, \"dam\", \"m\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_DOUBLE_EQ(r.as_number(), 10.0);
}

TEST(BuiltinsConvert, SiPrefixOnSquaredUnitSquaresTheMultiplier) {
  // (mega-meter)^2 = 10^12 m^2; 0.05 * 10^12 = 5e10.
  const Value r = EvalSource("=CONVERT(0.05, \"Mm^2\", \"m^2\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_DOUBLE_EQ(r.as_number(), 5e10);
}

TEST(BuiltinsConvert, SiPrefixOnCubedUnitCubesTheMultiplier) {
  // (mega-meter)^3 = 10^18 m^3; 0.0002 * 10^18 = 2e14.
  const Value r = EvalSource("=CONVERT(0.0002, \"Mm^3\", \"m^3\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_DOUBLE_EQ(r.as_number(), 2e14);
}

TEST(BuiltinsConvert, SiPrefixRejectedOnImperialUnit) {
  // Imperial units do not accept SI prefixes. `Mft` -> #N/A.
  const Value r = EvalSource("=CONVERT(1, \"Mft\", \"m\")");
  ASSERT_TRUE(r.is_error());
  EXPECT_EQ(r.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Binary prefixes (bit / byte only)
// ---------------------------------------------------------------------------

TEST(BuiltinsConvert, BinaryPrefixKibibit) {
  const Value r = EvalSource("=CONVERT(1, \"kibit\", \"bit\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_DOUBLE_EQ(r.as_number(), 1024.0);
}

TEST(BuiltinsConvert, BinaryPrefixMebibyteToBit) {
  // 1 Mibyte = 2^20 byte = 2^20 * 8 bit = 8388608 bit.
  const Value r = EvalSource("=CONVERT(1, \"Mibyte\", \"bit\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_DOUBLE_EQ(r.as_number(), 8388608.0);
}

TEST(BuiltinsConvert, BinaryPrefixRejectedOutsideInformation) {
  const Value r = EvalSource("=CONVERT(1, \"Mim\", \"m\")");
  ASSERT_TRUE(r.is_error());
  EXPECT_EQ(r.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Temperature (affine)
// ---------------------------------------------------------------------------

TEST(BuiltinsConvert, CelsiusToKelvinAddsOffset) {
  const Value r = EvalSource("=CONVERT(1, \"C\", \"K\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_DOUBLE_EQ(r.as_number(), 274.15);
}

TEST(BuiltinsConvert, FahrenheitToKelvin) {
  // 32 F == 273.15 K exactly.
  const Value r = EvalSource("=CONVERT(32, \"F\", \"K\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_NEAR(r.as_number(), 273.15, 1e-12);
}

TEST(BuiltinsConvert, CelsiusToFahrenheit) {
  // 100 C == 212 F.
  const Value r = EvalSource("=CONVERT(100, \"C\", \"F\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_NEAR(r.as_number(), 212.0, 1e-9);
}

TEST(BuiltinsConvert, MegaKelvinIsLinearPrefix) {
  const Value r = EvalSource("=CONVERT(12, \"MK\", \"K\")");
  ASSERT_TRUE(r.is_number());
  EXPECT_DOUBLE_EQ(r.as_number(), 12000000.0);
}

TEST(BuiltinsConvert, PrefixOnCelsiusIsRejected) {
  // SI prefixes apply only to K / kel in the temperature category.
  const Value r = EvalSource("=CONVERT(1, \"MC\", \"K\")");
  ASSERT_TRUE(r.is_error());
  EXPECT_EQ(r.as_error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

TEST(BuiltinsConvert, UnknownUnitIsNA) {
  const Value r = EvalSource("=CONVERT(1, \"xyz\", \"m\")");
  ASSERT_TRUE(r.is_error());
  EXPECT_EQ(r.as_error(), ErrorCode::NA);
}

TEST(BuiltinsConvert, IncompatibleCategoriesIsNA) {
  const Value r = EvalSource("=CONVERT(1, \"m\", \"g\")");
  ASSERT_TRUE(r.is_error());
  EXPECT_EQ(r.as_error(), ErrorCode::NA);
}

TEST(BuiltinsConvert, NonNumericValueIsValue) {
  const Value r = EvalSource("=CONVERT(\"hello\", \"m\", \"m\")");
  ASSERT_TRUE(r.is_error());
  EXPECT_EQ(r.as_error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
