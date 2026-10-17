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
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using formulon::test::EvalSource;

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

// Excel 365 doubles, bit-exact.
TEST(ExcelExactConvert, MatchesExcelBitForBit) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=CONVERT(1,\"m\",\"ft\")", 3.2808398950131235},
      {"=CONVERT(1,\"m\",\"in\")", 39.37007874015748},
      {"=CONVERT(1,\"ft\",\"m\")", 0.3048},
      {"=CONVERT(1,\"mi\",\"km\")", 1.609344},
      {"=CONVERT(1,\"km\",\"mi\")", 0.621371192237334},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

// ---------------------------------------------------------------------------
// Exact values: unit factors, prefix resolution order, the "pc" reading and
// the temperature pair formulas.
// ---------------------------------------------------------------------------

TEST(ConvertTable, MatchesExcelValues) {
  struct Case {
    const char* formula;
    double expected;
  };
  const Case cases[] = {
      {"=CONVERT(1,\"hweight\",\"g\")", 50802.34544},
      {"=CONVERT(1,\"hweight\",\"cwt\")", 1.1199999999999999},
      {"=CONVERT(1,\"SI_at\",\"Pa\")", 98066.51},
      {"=CONVERT(1,\"at\",\"SI_at\")", 1.033227347440018},
      {"=CONVERT(1,\"dam\",\"m\")", 10.0},
      {"=CONVERT(1,\"em\",\"m\")", 10.0},
      {"=CONVERT(1,\"dam2\",\"m2\")", 100.0},
      {"=CONVERT(1,\"km2\",\"m2\")", 1000000.0},
      {"=CONVERT(1,\"mm^3\",\"m^3\")", 1e-09},
      {"=CONVERT(1,\"cm3\",\"ml\")", 1.0},
      {"=CONVERT(1,\"Ym3\",\"ym3\")", 1e+144},
      {"=CONVERT(1,\"kibit\",\"bit\")", 1024.0},
      {"=CONVERT(1,\"ly3\",\"ang3\")", 8.467866646237152e+77},
      {"=CONVERT(1,\"ang\",\"ly\")", 1.0570008340246154e-26},
      {"=CONVERT(1,\"pc\",\"J\")", 4.184e-12},
      {"=CONVERT(1,\"c\",\"pc\")", 1000000000000.0},
      {"=CONVERT(1,\"pc\",\"m\")", 3.0856775812815532e+16},
      {"=CONVERT(1,\"kpc\",\"m\")", 3.085677581281553e+19},
      {"=CONVERT(1,\"kK\",\"C\")", 726.85},
      {"=CONVERT(1,\"C\",\"kK\")", 0.27415},
      {"=CONVERT(1,\"Reau\",\"kK\")", 274.4},
      {"=CONVERT(1,\"kK\",\"Rank\")", 1.8},
      {"=CONVERT(1,\"kK\",\"F\")", 1340.3300000000002},
      {"=CONVERT(1,\"F\",\"kK\")", 0.25592777777777775},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_number()) << c.formula;
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
  }
}

TEST(ConvertTable, ErrorsMatchExcel) {
  struct Case {
    const char* formula;
    ErrorCode expected;
  };
  const Case cases[] = {
      {"=CONVERT(1,\"eam\",\"m\")", ErrorCode::NA},
      {"=CONVERT(1,\"da\",\"m\")", ErrorCode::NA},
      {"=CONVERT(1,\"Kibit\",\"bit\")", ErrorCode::NA},
      {"=CONVERT(1,\"kim\",\"m\")", ErrorCode::NA},
      {"=CONVERT(1,\"KM\",\"m\")", ErrorCode::NA},
      {"=CONVERT(1,\"km \",\"m\")", ErrorCode::NA},
      {"=CONVERT(1,\"\",\"m\")", ErrorCode::NA},
      {"=CONVERT(1,\"kpc\",\"c\")", ErrorCode::NA},
      {"=CONVERT(1,\"parsec\",\"c\")", ErrorCode::NA},
      {"=CONVERT(1,\"kC\",\"K\")", ErrorCode::NA},
      {"=CONVERT(9.99999999999999E307,\"m\",\"ft\")", ErrorCode::Num},
      {"=CONVERT(9.99999999999999E307,\"C\",\"F\")", ErrorCode::Num},
      {"=CONVERT(9.99999999999999E307,\"K\",\"Rank\")", ErrorCode::Num},
      {"=CONVERT(9.99999999999999E307,\"Yg\",\"yg\")", ErrorCode::Num},
  };
  for (const Case& c : cases) {
    const Value v = EvalSource(c.formula);
    ASSERT_TRUE(v.is_error()) << c.formula;
    EXPECT_EQ(v.as_error(), c.expected) << c.formula;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
