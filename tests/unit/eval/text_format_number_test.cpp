// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the number-format engine driving TEXT() through the
// numeric side of Excel's format-string language. Date/time coverage lives
// in `text_format_date_test.cpp`; top-level TEXT semantics (value
// coercion, text section, error propagation) live in
// `value_numbervalue_test.cpp`.

#include <string>
#include <string_view>

#include "eval/text_format/number_format.h"
#include "gtest/gtest.h"

namespace formulon {
namespace eval {
namespace text_format {
namespace {

// Convenience wrapper: render `value` through `format` and return the
// resulting string. On any engine failure the returned optional is empty.
std::string Render(double value, std::string_view format) {
  std::string out;
  const FormatStatus s = apply_format(value, format, out);
  EXPECT_EQ(s, FormatStatus::kOk);
  return out;
}

// ---------------------------------------------------------------------------
// Integer-only formats (no decimal point)
// ---------------------------------------------------------------------------

TEST(NumberFormatIntegers, ZeroPad) {
  EXPECT_EQ(Render(5.0, "000"), "005");
}

TEST(NumberFormatIntegers, ZeroPadMany) {
  EXPECT_EQ(Render(42.0, "00000"), "00042");
}

TEST(NumberFormatIntegers, HashNoPad) {
  EXPECT_EQ(Render(5.0, "###"), "5");
}

TEST(NumberFormatIntegers, MixedHashZero) {
  EXPECT_EQ(Render(5.0, "##0"), "5");
}

TEST(NumberFormatIntegers, ThousandsSeparator) {
  EXPECT_EQ(Render(1234567.0, "#,##0"), "1,234,567");
}

TEST(NumberFormatIntegers, ThousandsSeparatorSmall) {
  EXPECT_EQ(Render(12.0, "#,##0"), "12");
}

TEST(NumberFormatIntegers, Negative) {
  EXPECT_EQ(Render(-5.0, "000"), "-005");
}

TEST(NumberFormatIntegers, Zero) {
  EXPECT_EQ(Render(0.0, "000"), "000");
}

// ---------------------------------------------------------------------------
// Decimal formats
// ---------------------------------------------------------------------------

TEST(NumberFormatDecimal, TwoDecimalPlaces) {
  EXPECT_EQ(Render(3.14159, "0.00"), "3.14");
}

TEST(NumberFormatDecimal, RoundUp) {
  EXPECT_EQ(Render(3.145, "0.00"), "3.15");
}

TEST(NumberFormatDecimal, RoundDown) {
  EXPECT_EQ(Render(3.144, "0.00"), "3.14");
}

TEST(NumberFormatDecimal, ZeroWithDecimals) {
  EXPECT_EQ(Render(0.0, "0.00"), "0.00");
}

TEST(NumberFormatDecimal, Negative) {
  EXPECT_EQ(Render(-1.5, "0.00"), "-1.50");
}

TEST(NumberFormatDecimal, HashFractional) {
  // "#.##" trims trailing zero fraction digits.
  EXPECT_EQ(Render(1.5, "#.##"), "1.5");
}

TEST(NumberFormatDecimal, HashFractionalIntegerResult) {
  EXPECT_EQ(Render(1.0, "#.##"), "1.");
}

TEST(NumberFormatDecimal, CombinedThousandsAndDecimal) {
  EXPECT_EQ(Render(1234.5, "#,##0.00"), "1,234.50");
}

// ---------------------------------------------------------------------------
// Percent
// ---------------------------------------------------------------------------

TEST(NumberFormatPercent, BasicHalf) {
  EXPECT_EQ(Render(0.5, "0%"), "50%");
}

TEST(NumberFormatPercent, TwoDecimal) {
  EXPECT_EQ(Render(0.1234, "0.00%"), "12.34%");
}

TEST(NumberFormatPercent, NegativePercent) {
  EXPECT_EQ(Render(-0.25, "0%"), "-25%");
}

// ---------------------------------------------------------------------------
// Trailing commas (scale by 1e3)
// ---------------------------------------------------------------------------

TEST(NumberFormatScale, TrailingCommaDividesByThousand) {
  EXPECT_EQ(Render(1200000.0, "#,##0,"), "1,200");
}

TEST(NumberFormatScale, DoubleTrailingCommaDividesByMillion) {
  // 2_700_000 / 1e6 = 2.7, rounds unambiguously up to 3.
  EXPECT_EQ(Render(2700000.0, "0,,"), "3");
}

// ---------------------------------------------------------------------------
// Scientific notation
// ---------------------------------------------------------------------------

TEST(NumberFormatScientific, ExpPlus) {
  EXPECT_EQ(Render(12345.0, "0.00E+00"), "1.23E+04");
}

TEST(NumberFormatScientific, ExpPlusLargeExponent) {
  EXPECT_EQ(Render(1.5e10, "0.0E+00"), "1.5E+10");
}

TEST(NumberFormatScientific, ExpMinus) {
  // `E-` only emits the sign for negative exponents.
  EXPECT_EQ(Render(12345.0, "0E-00"), "1E04");
}

TEST(NumberFormatScientific, ExpNegativeExponent) {
  EXPECT_EQ(Render(0.0001234, "0.00E+00"), "1.23E-04");
}

// ---------------------------------------------------------------------------
// Literal passthrough / escapes / quoted text
// ---------------------------------------------------------------------------

TEST(NumberFormatLiteral, SuffixJapaneseYen) {
  // `円` is 3 UTF-8 bytes; the engine copies them verbatim.
  EXPECT_EQ(Render(123.0, "0円"), "123円");
}

TEST(NumberFormatLiteral, QuotedText) {
  EXPECT_EQ(Render(7.0, "\"items: \"0"), "items: 7");
}

TEST(NumberFormatLiteral, BackslashEscape) {
  EXPECT_EQ(Render(10.0, "\\$0"), "$10");
}

TEST(NumberFormatLiteral, BangEscape) {
  EXPECT_EQ(Render(10.0, "!@0"), "@10");
}

// Note: per-digit-position literals (e.g. "000-0000" for phone number
// rendering) are not supported by the current engine, which emits the
// integer block as a single run. Revisit if the oracle trips this shape.

// ---------------------------------------------------------------------------
// Section separators
// ---------------------------------------------------------------------------

TEST(NumberFormatSections, TwoSectionsNegative) {
  EXPECT_EQ(Render(-5.0, "0.00;(0.00)"), "(5.00)");
}

TEST(NumberFormatSections, TwoSectionsPositive) {
  EXPECT_EQ(Render(5.0, "0.00;(0.00)"), "5.00");
}

TEST(NumberFormatSections, TwoSectionsZero) {
  EXPECT_EQ(Render(0.0, "0.00;(0.00)"), "0.00");
}

TEST(NumberFormatSections, ThreeSectionsZero) {
  EXPECT_EQ(Render(0.0, "0.00;(0.00);\"zero\""), "zero");
}

TEST(NumberFormatSections, FourSectionsNumericValueUsesFirst) {
  // With a numeric value and four sections, the text section is not used.
  EXPECT_EQ(Render(1.0, "0.0;(0.0);\"z\";@"), "1.0");
}

// ---------------------------------------------------------------------------
// Bracketed specifiers: colour names rejected, locale-currency discarded.
// ---------------------------------------------------------------------------

TEST(NumberFormatBracketed, ColorRedIsValueError) {
  // Mac Excel ja-JP rejects bracketed colour names in TEXT.
  std::string out;
  const FormatStatus s = apply_format(5.0, "[Red]0.00", out);
  EXPECT_EQ(s, FormatStatus::kValueError);
}

TEST(NumberFormatBracketed, LocaleCurrencyDiscarded) {
  // `[$...]` locale-currency markers are accepted and silently dropped.
  EXPECT_EQ(Render(5.0, "[$-409]0.00"), "5.00");
}

// ---------------------------------------------------------------------------
// Empty format
// ---------------------------------------------------------------------------

TEST(NumberFormatEmpty, EmptyFormatYieldsEmpty) {
  std::string out;
  const FormatStatus s = apply_format(42.0, "", out);
  EXPECT_EQ(s, FormatStatus::kOk);
  EXPECT_EQ(out, "");
}

}  // namespace
}  // namespace text_format
}  // namespace eval
}  // namespace formulon
