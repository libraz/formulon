//
// Unit tests for `formulon::format_double`. The shared formatter feeds the
// AST S-expression dumper and the tree-walk evaluator's text coercion, so
// any behaviour change here is observable in both layers.

#include "utils/double_format.h"

#include <cmath>
#include <limits>
#include <string>

#include "gtest/gtest.h"

namespace formulon {
namespace {

// Helper: format `v` into a fresh string and return it.
std::string Format(double v) {
  std::string out;
  format_double(out, v);
  return out;
}

TEST(FormatDouble, IntegerInRange) {
  EXPECT_EQ(Format(0.0), "0");
  EXPECT_EQ(Format(1.0), "1");
  EXPECT_EQ(Format(42.0), "42");
  EXPECT_EQ(Format(-7.0), "-7");
}

TEST(FormatDouble, NegativeZeroCollapses) {
  EXPECT_EQ(Format(-0.0), "0");
}

TEST(FormatDouble, FractionalTrimsTrailingZeros) {
  EXPECT_EQ(Format(3.14), "3.14");
  EXPECT_EQ(Format(1.5), "1.5");
}

TEST(FormatDouble, NaN) {
  EXPECT_EQ(Format(std::numeric_limits<double>::quiet_NaN()), "nan");
}

TEST(FormatDouble, PositiveInfinity) {
  EXPECT_EQ(Format(std::numeric_limits<double>::infinity()), "inf");
}

TEST(FormatDouble, NegativeInfinity) {
  EXPECT_EQ(Format(-std::numeric_limits<double>::infinity()), "-inf");
}

TEST(FormatDouble, AppendsDoesNotOverwrite) {
  std::string out("prefix:");
  format_double(out, 12.0);
  EXPECT_EQ(out, "prefix:12");
}

TEST(FormatDouble, LargeIntegerRoundtrip) {
  // Exactly representable integers below 1e16 take the integer fast path.
  EXPECT_EQ(Format(1234567890.0), "1234567890");
}

TEST(FormatDouble, JustAboveFastPathFallsBack) {
  // 1e16 itself is outside the fast path; we just want a non-empty,
  // dot-trimmed result.
  const std::string s = Format(1e16);
  EXPECT_FALSE(s.empty());
  EXPECT_NE(s.back(), '.');
}

// -- Excel General-format threshold ---------------------------------------
// Mac Excel 365 (ja-JP, build 16.108) keeps decimal notation as long as
// the unsigned magnitude string is at most 20 chars; beyond that it
// switches to scientific. The leading `-` sign is NOT counted toward
// the 20-char ceiling. Test cases pinned via 87 oracle probes — see
// tests/oracle/cases/valuetotext_general_threshold_probes*.yaml.

TEST(FormatDouble, NegativeExponentBoundary) {
  // 1e-9 -> 17 chars decimal -> stays decimal.
  EXPECT_EQ(Format(7.123456e-9), "0.000000007123456");
  // 1e-18 -> exactly 20 chars decimal -> stays decimal.
  EXPECT_EQ(Format(1e-18), "0.000000000000000001");
  // 1e-19 -> would need 21 chars decimal -> switches to scientific.
  EXPECT_EQ(Format(1e-19), "1E-19");
  EXPECT_EQ(Format(1e-100), "1E-100");
}

TEST(FormatDouble, PositiveExponentBoundary) {
  // 1e19 -> exactly 20 chars decimal -> stays decimal.
  EXPECT_EQ(Format(1e19), "10000000000000000000");
  // 1e20 -> would need 21 chars decimal -> switches to scientific.
  EXPECT_EQ(Format(1e20), "1E+20");
  EXPECT_EQ(Format(1e100), "1E+100");
}

TEST(FormatDouble, NegativeSignDoesNotCount) {
  // 20-char unsigned magnitude with a leading '-': stays decimal even
  // though the total width is 21. Mac does not count the sign.
  EXPECT_EQ(Format(-1e19), "-10000000000000000000");
  EXPECT_EQ(Format(-1e-18), "-0.000000000000000001");
}

TEST(FormatDouble, ExcelStyleScientific) {
  // Exponent zero-padded to two digits.
  EXPECT_EQ(Format(1.5e-20), "1.5E-20");
  EXPECT_EQ(Format(1.5e25), "1.5E+25");
  // |exp| >= 100 expands naturally to three digits.
  EXPECT_EQ(Format(1.5e-105), "1.5E-105");
}

TEST(FormatDouble, FifteenSigDigitPrecision) {
  // %.15g rounds to 15 significant digits (Excel's General-format limit).
  EXPECT_EQ(Format(1.0 / 3.0), "0.333333333333333");
  EXPECT_EQ(Format(2.0 / 3.0), "0.666666666666667");
}

std::string FormatLiteral(double v) {
  std::string out;
  format_formula_number(out, v);
  return out;
}

TEST(FormatFormulaNumber, PositionalUpToTwentyOneCharacters) {
  EXPECT_EQ(FormatLiteral(1e20), "100000000000000000000");
  EXPECT_EQ(FormatLiteral(9.99999999999999e20), "999999999999999000000");
  EXPECT_EQ(FormatLiteral(1.23456789012345e20), "123456789012345000000");
  EXPECT_EQ(FormatLiteral(1e-19), "0.0000000000000000001");
  EXPECT_EQ(FormatLiteral(2.5e-18), "0.0000000000000000025");
  EXPECT_EQ(FormatLiteral(1.23456789012345e-5), "0.0000123456789012345");
}

TEST(FormatFormulaNumber, ScientificBeyondTwentyOneCharactersWithoutPadding) {
  EXPECT_EQ(FormatLiteral(1e21), "1E+21");
  EXPECT_EQ(FormatLiteral(1.5e21), "1.5E+21");
  EXPECT_EQ(FormatLiteral(1.23456789012345e21), "1.23456789012345E+21");
  EXPECT_EQ(FormatLiteral(1e-20), "1E-20");
  EXPECT_EQ(FormatLiteral(2.5e-19), "2.5E-19");
  EXPECT_EQ(FormatLiteral(1.2345e-18), "1.2345E-18");
  EXPECT_EQ(FormatLiteral(1.23456789012345e-10), "1.23456789012345E-10");
  EXPECT_EQ(FormatLiteral(1.23456789012345e-9), "1.23456789012345E-9");
  EXPECT_EQ(FormatLiteral(1e307), "1E+307");
  EXPECT_EQ(FormatLiteral(1e-100), "1E-100");
}

TEST(FormatFormulaNumber, SignIsOutsideTheWidthRule) {
  EXPECT_EQ(FormatLiteral(-1e20), "-100000000000000000000");
  EXPECT_EQ(FormatLiteral(-1e21), "-1E+21");
  EXPECT_EQ(FormatLiteral(-0.0), "0");
  EXPECT_EQ(FormatLiteral(-0.5), "-0.5");
}

TEST(FormatFormulaNumber, FifteenSignificantDigitsTrailingZerosDropped) {
  EXPECT_EQ(FormatLiteral(1.5), "1.5");
  EXPECT_EQ(FormatLiteral(1234567890123450.0), "1234567890123450");
  EXPECT_EQ(FormatLiteral(0.30000000000000004), "0.3");
  EXPECT_EQ(FormatLiteral(9.99999999999999e307), "9.99999999999999E+307");
}

}  // namespace
}  // namespace formulon
