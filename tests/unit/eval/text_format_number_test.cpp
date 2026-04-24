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
// Bracketed specifiers: named colours tolerated, locale-currency discarded,
// unknown qualifiers rejected.
// ---------------------------------------------------------------------------

TEST(NumberFormatBracketed, NamedColorSilentlyDropped) {
  // Mac Excel 365 and IronCalc silently discard colour qualifiers in TEXT
  // and format the value with the rest of the section.
  EXPECT_EQ(Render(5.0, "[Red]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[Blue]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[green]0.00"), "5.00");
}

TEST(NumberFormatBracketed, IndexedColorSilentlyDropped) {
  // `[ColorN]` for N in 1..56 is also dropped silently.
  EXPECT_EQ(Render(5.0, "[Color12]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[Color1]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[Color56]0.00"), "5.00");
}

TEST(NumberFormatBracketed, IndexedColorOutOfRangeIsValueError) {
  // `[Color57]` (and higher) is not a recognised colour -> invalid.
  std::string out;
  EXPECT_EQ(apply_format(5.0, "[Color57]0.00", out), FormatStatus::kValueError);
}

TEST(NumberFormatBracketed, ConditionalBracketStillRejected) {
  // Conditional qualifiers such as `[>100]` are not implemented.
  std::string out;
  EXPECT_EQ(apply_format(5.0, "[>100]0.00", out), FormatStatus::kValueError);
}

TEST(NumberFormatBracketed, LocaleCurrencyDiscarded) {
  // `[$...]` locale-currency markers are accepted and silently dropped.
  EXPECT_EQ(Render(5.0, "[$-409]0.00"), "5.00");
}

// ---------------------------------------------------------------------------
// Underscore-skip `_X`: emits a single space placeholder.
// ---------------------------------------------------------------------------

TEST(NumberFormatUnderscoreSkip, AccountingParens) {
  // Classic accounting format; positive branch renders a trailing space so
  // the digits line up with the parenthesised negative branch.
  EXPECT_EQ(Render(1234.0, "#,##0_);(#,##0)"), "1,234 ");
  EXPECT_EQ(Render(-1234.0, "#,##0_);(#,##0)"), "(1,234)");
}

TEST(NumberFormatUnderscoreSkip, SpaceBeforeDigits) {
  // `_(` in front of digits reserves a leading space.
  EXPECT_EQ(Render(5.0, "_(0.00"), " 5.00");
}

TEST(NumberFormatUnderscoreSkip, TrailingUnderscoreIsLiteral) {
  // A dangling `_` at end-of-format has no following byte to reserve;
  // it falls back to the single-byte literal path.
  EXPECT_EQ(Render(5.0, "0_"), "5_");
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

// ---------------------------------------------------------------------------
// `General` keyword -- 11-character-wide Excel default display.
// ---------------------------------------------------------------------------

TEST(NumberFormatGeneral, IntegerPositiveAndNegative) {
  // Whole numbers round-trip via the integer fast path: no decimal, sign
  // propagated by the outer walker.
  EXPECT_EQ(Render(12.0, "General"), "12");
  EXPECT_EQ(Render(-12.0, "General"), "-12");
  // Large-but-still-integral values skip scientific notation when they fit
  // within the fixed-width budget.
  EXPECT_EQ(Render(1234567890.0, "General"), "1234567890");
}

TEST(NumberFormatGeneral, FractionTrimmedAndScientific) {
  // 1/3 prints 9 fractional digits (exactly what Mac Excel / IronCalc
  // goldens emit), with trailing zeros trimmed.
  EXPECT_EQ(Render(1.0 / 3.0, "General"), "0.333333333");
  // Large magnitudes switch to scientific with an exponent zero-padded to
  // two digits; trailing mantissa zeros still collapse.
  EXPECT_EQ(Render(250000000000.0, "General"), "2.5E+11");
  EXPECT_EQ(Render(123456789012.0, "General"), "1.23457E+11");
}

// ---------------------------------------------------------------------------
// Interleaved digit + literal positional rendering.
// ---------------------------------------------------------------------------

TEST(NumberFormatInterleavedDigits, DashSeparatedDigits) {
  // The eight `0` tokens consume the eight right-aligned digits of `12`,
  // leaving the interleaved `-` literals in their original positions.
  EXPECT_EQ(Render(12.0, "00-00-00-00"), "00-00-00-12");
  EXPECT_EQ(Render(12345678.0, "00-00-00-00"), "12-34-56-78");
}

// ---------------------------------------------------------------------------
// Signed-zero suppression.
// ---------------------------------------------------------------------------

TEST(NumberFormatSignedZero, TwoSectionAccountingZero) {
  // `0` is exactly zero; the positive section (including the `_)` trailing
  // space placeholder) must render, not the negative branch.
  EXPECT_EQ(Render(0.0, "#,##0_);(#,##0)"), "0 ");
  // `-0.0` is IEEE-754-signed zero: signbit is true, value is still zero.
  // The format must still pick the positive section.
  EXPECT_EQ(Render(-0.0, "#,##0_);(#,##0)"), "0 ");
  // A tiny negative value that rounds to zero under the format must also
  // strip the leading minus sign (Excel's "effective zero" rule).
  EXPECT_EQ(Render(-1.0 / 3.0, "0"), "0");
}

}  // namespace
}  // namespace text_format
}  // namespace eval
}  // namespace formulon
