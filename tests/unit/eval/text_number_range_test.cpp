// Tests for the text-to-number range rules: digit truncation, folded percent,
// exponent window and the length limit.

#include <cfloat>
#include <cmath>
#include <string>

#include "eval/coerce.h"
#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"

namespace formulon::eval {
namespace {

class TextNumberRange : public ::testing::Test {
 protected:
  ScopedEvalProfile scope_{mac_365_en_us_profile()};

  static bool accepts(const std::string& text, double* out) {
    const auto r = coerce_text_to_number(text);
    if (!r.has_value()) {
      return false;
    }
    *out = r.value();
    return true;
  }

  void expect_value(const std::string& text, double want) {
    double got = -1.0;
    ASSERT_TRUE(accepts(text, &got)) << text;
    EXPECT_EQ(got, want) << text;
  }

  void expect_rejected(const std::string& text) {
    double got = 0.0;
    EXPECT_FALSE(accepts(text, &got)) << text;
  }
};

TEST_F(TextNumberRange, UpperBound) {
  expect_value("9.99999999999999E+307", 9.99999999999999e307);
  expect_value("9.999999999999995e307", 9.99999999999999e307);
  expect_value("0.1e308", 1e307);
  expect_value("1,000e304", 1e307);
  expect_value("1234567890123456789e289", 1.23456789012345e307);
  expect_rejected("1e308");
  expect_rejected("1000e305");
  expect_rejected("1234567890123456789e290");
  expect_rejected("1,000e305");
}

TEST_F(TextNumberRange, LowerBoundFlushesToZero) {
  for (const char* t : {"1e-309", "10e-310", "2.2250738585072014e-308", "-1e-309", "1.00000000000000000001e-309"}) {
    expect_value(t, 0.0);
  }
  expect_value("2.22507385850721e-308", 2.2250738585072098e-308);
  for (const char* t : {"1e-310", "0.1e-309", "9.99e-310", "9.99999999999999999999e-310"}) {
    expect_rejected(t);
  }
}

TEST_F(TextNumberRange, ZeroMantissa) {
  expect_value("0.0e309", 0.0);
  expect_value("0e-308", 0.0);
  expect_value("0000000000000000.0e309", 0.0);
  for (const char* t : {"0e309", "0e-309", "0.e309", "0.0e-308", "0e999", "-0e309"}) {
    expect_rejected(t);
  }
}

TEST_F(TextNumberRange, PercentFoldsIntoTheExponent) {
  expect_value("1e-9%", 1e-11);
  expect_value("1e309%", 1e307);
  expect_value("%1e309", 1e307);
  expect_value(" 1e309 % ", 1e307);
  expect_value("-1e309%", -1e307);
  expect_rejected("1e-308%");
  expect_rejected("1e310%");
  expect_rejected("1e-306%%");
  expect_rejected("%1%");
}

TEST_F(TextNumberRange, DigitsPastTheFifteenthAreTruncated) {
  expect_value("9999999999999999", 9999999999999990.0);
  expect_value("0.99999999999999999", 0.999999999999999);
  expect_value("1.234567890123456789", 1.23456789012345);
  expect_value("12345678901234567890", 1.23456789012345e19);
}

TEST_F(TextNumberRange, HugeExponentDigitsDoNotWrap) {
  expect_value("1e00000000000000000000007", 1e7);
  for (const char* t : {"1e4294967597", "1e2147483948", "1e65837", "1e99999999999999999999", "1e-4294967597"}) {
    expect_rejected(t);
  }
}

TEST_F(TextNumberRange, LengthLimitCountsSpaces) {
  expect_value(std::string(254, ' ') + "1", 1.0);
  expect_rejected(std::string(255, ' ') + "1");
  expect_value("1" + std::string(254, ' '), 1.0);
  expect_rejected("1" + std::string(255, ' '));
}

TEST_F(TextNumberRange, LengthCountsCharactersNotBytes) {
  const ScopedEvalProfile ja_scope(mac_365_ja_jp_profile());
  const std::string fullwidth_one = "\xEF\xBC\x91";
  const std::string ideographic_space = "\xE3\x80\x80";
  const auto repeat = [](const std::string& unit, int n) {
    std::string out;
    for (int i = 0; i < n; ++i) {
      out += unit;
    }
    return out;
  };
  expect_value(std::string(254, ' ') + fullwidth_one, 1.0);
  expect_rejected(std::string(255, ' ') + fullwidth_one);
  expect_value(repeat(ideographic_space, 254) + "1", 1.0);
  expect_rejected(repeat(ideographic_space, 255) + "1");
}

TEST_F(TextNumberRange, SinglePercentAnywhereOutsideTheDigits) {
  for (const char* t : {"%5", "-%5", "%-5", "% 5", "5 %", "+%5", "%+5"}) {
    double got = 0.0;
    ASSERT_TRUE(accepts(t, &got)) << t;
    EXPECT_DOUBLE_EQ(std::fabs(got), 0.05) << t;
  }
  expect_value("-%5", -0.05);
  expect_value("%-5", -0.05);
  for (const char* t : {"(5)%", "%(5)", "(%5)", "(5%)"}) {
    expect_value(t, -0.05);
  }
  expect_rejected("%5%");
  expect_rejected("5%5");
}

}  // namespace
}  // namespace formulon::eval
