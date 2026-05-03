// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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

}  // namespace
}  // namespace formulon
