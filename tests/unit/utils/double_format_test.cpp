// Copyright 2026 libraz. Licensed under the MIT License.
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

TEST(FormatDouble, NegativeZeroCollapses) { EXPECT_EQ(Format(-0.0), "0"); }

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

}  // namespace
}  // namespace formulon
