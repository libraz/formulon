//
// Lexical-space tests for `io::parse_xsd_double`, the single decoder for
// every OOXML numeric payload (`<v>` cell bodies and pivot-cache
// `<n v=...>` attributes).
//
// The interesting cases are the ones `std::strtod` accepts and xs:double
// does not: a value that reads as a different number (hexadecimal), and
// values the writer cannot re-emit so that they reload identically
// (±infinity, NaN, overflow, underflow to zero).

#include "io/xsd_double.h"

#include <cmath>
#include <limits>
#include <string>

#include "gtest/gtest.h"

namespace formulon {
namespace io {
namespace {

double ParsedOr(const char* text, double fallback) {
  double out = fallback;
  return parse_xsd_double(text, &out) ? out : fallback;
}

TEST(XsdDouble, AcceptsTheOrdinaryLexicalForms) {
  EXPECT_DOUBLE_EQ(ParsedOr("0", -1.0), 0.0);
  EXPECT_DOUBLE_EQ(ParsedOr("42", -1.0), 42.0);
  EXPECT_DOUBLE_EQ(ParsedOr("-42", 0.0), -42.0);
  EXPECT_DOUBLE_EQ(ParsedOr("+3.5", 0.0), 3.5);
  EXPECT_DOUBLE_EQ(ParsedOr("1.5e3", 0.0), 1500.0);
  EXPECT_DOUBLE_EQ(ParsedOr("1.5E-3", 0.0), 0.0015);
  EXPECT_DOUBLE_EQ(ParsedOr(" 2.5 ", 0.0), 2.5);
}

TEST(XsdDouble, RejectsEmptyAndTrailingGarbage) {
  double out = 0.0;
  EXPECT_FALSE(parse_xsd_double("", &out));
  EXPECT_FALSE(parse_xsd_double("abc", &out));
  EXPECT_FALSE(parse_xsd_double("1.5q", &out));
  EXPECT_DOUBLE_EQ(out, 0.0) << "a rejected value must leave *out untouched";
}

TEST(XsdDouble, RejectsHexLiterals) {
  // `strtod` reads `0x10` as 16; xs:double has no hexadecimal form, so
  // accepting it would silently load a different number than the file
  // states.
  double out = 0.0;
  EXPECT_FALSE(parse_xsd_double("0x10", &out));
  EXPECT_FALSE(parse_xsd_double("0X10", &out));
  EXPECT_FALSE(parse_xsd_double("-0x1p4", &out));
  EXPECT_FALSE(parse_xsd_double(" 0x10", &out)) << "the gate must run after the whitespace strtod skips";
  EXPECT_DOUBLE_EQ(out, 0.0);
}

TEST(XsdDouble, RejectsInfinityAndNanSpellings) {
  // The writer cannot express a non-finite number in a `<v>` body — it
  // downgrades one to `#NUM!` — so a reader that accepted these would
  // turn a load into a silent value change at the next save.
  double out = 0.0;
  EXPECT_FALSE(parse_xsd_double("inf", &out));
  EXPECT_FALSE(parse_xsd_double("INF", &out));
  EXPECT_FALSE(parse_xsd_double("-INF", &out));
  EXPECT_FALSE(parse_xsd_double("Infinity", &out));
  EXPECT_FALSE(parse_xsd_double("nan", &out));
  EXPECT_FALSE(parse_xsd_double("NaN", &out));
  EXPECT_FALSE(parse_xsd_double("-nan", &out));
  EXPECT_FALSE(parse_xsd_double("  inf", &out));
  EXPECT_DOUBLE_EQ(out, 0.0);
}

TEST(XsdDouble, RejectsOverflowToInfinity) {
  double out = 0.0;
  EXPECT_FALSE(parse_xsd_double("1e999", &out));
  EXPECT_FALSE(parse_xsd_double("-1e999", &out));
  EXPECT_DOUBLE_EQ(out, 0.0);
}

TEST(XsdDouble, RejectsUnderflowToZero) {
  double out = 7.0;
  EXPECT_FALSE(parse_xsd_double("1e-999", &out));
  EXPECT_DOUBLE_EQ(out, 7.0);
}

TEST(XsdDouble, AcceptsTheRepresentableExtremes) {
  // The boundary values are in range and must keep loading: a subnormal
  // may carry ERANGE from the C library but is representable, and the
  // largest finite double round-trips through the writer unchanged.
  const std::string max_text = "1.7976931348623157e308";
  EXPECT_DOUBLE_EQ(ParsedOr(max_text.c_str(), 0.0), std::numeric_limits<double>::max());
  const double denormal = ParsedOr("5e-324", 0.0);
  EXPECT_GT(denormal, 0.0);
  EXPECT_TRUE(std::isfinite(denormal));
}

}  // namespace
}  // namespace io
}  // namespace formulon
