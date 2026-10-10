//
// Unit tests for the number-format engine driving TEXT() through the
// numeric side of Excel's format-string language. Date/time coverage lives
// in `text_format_date_test.cpp`; top-level TEXT semantics (value
// coercion, text section, error propagation) live in
// `value_numbervalue_test.cpp`.

#include <cmath>
#include <limits>
#include <string>
#include <string_view>

#include "eval/eval_profile_scope.h"
#include "eval/text_format/display_text.h"
#include "eval/text_format/number_format.h"
#include "gtest/gtest.h"

namespace formulon {
namespace text_format {
namespace {

// Convenience wrapper: render `value` through `format` and return the
// resulting string. On any engine failure the returned optional is empty.
std::string Render(double value, std::string_view format) {
  const eval::ScopedEvalProfile profile_scope(mac_365_ja_jp_profile());
  std::string out;
  const FormatStatus s = apply_format(value, format, out);
  EXPECT_EQ(s, FormatStatus::kOk);
  return out;
}

// `Render` under en-US, where the English `General` keyword is accepted.
std::string RenderEnUs(double value, std::string_view format) {
  const eval::ScopedEvalProfile profile_scope(mac_365_en_us_profile());
  std::string out;
  const FormatStatus s = apply_format(value, format, out);
  EXPECT_EQ(s, FormatStatus::kOk);
  return out;
}

// ---------------------------------------------------------------------------
// Integer-only formats (no decimal point)
// ---------------------------------------------------------------------------

TEST(NumberFormatIntegers, LargeFiniteValuesProduceCompleteDigitsWithoutNulls) {
  const std::string digits = "1" + std::string(308, '0');
  EXPECT_EQ(Render(1E308, "0"), digits);
  EXPECT_EQ(Render(-1E308, "0.00"), "-" + digits + ".00");
  const auto display = format_value_for_display(Value::number(1E308), "0", false, mac_365_ja_jp_profile());
  EXPECT_EQ(display.status, DisplayStatus::kOk);
  EXPECT_EQ(display.text, digits);
}

TEST(NumberFormatIntegers, MaximumFiniteDoubleCanRoundBeyondItsBinaryRange) {
  const double value = std::numeric_limits<double>::max();
  const std::string expected = "179769313486232" + std::string(294, '0');
  EXPECT_EQ(Render(value, "0"), expected);
  EXPECT_EQ(Render(-value, "0"), "-" + expected);
  // General budgets 11 characters; the three-digit exponent leaves four fractional digits.
  EXPECT_EQ(RenderEnUs(value, "General"), "1.7977E+308");
}

TEST(NumberFormatPercent, ScalingOverflowReturnsOverflowWithoutAppending) {
  std::string out = "prefix";
  EXPECT_EQ(apply_format(1E307, "0%", out), FormatStatus::kOverflow);
  EXPECT_EQ(out, "prefix");
  const auto display = format_value_for_display(Value::number(1E307), "0%", false, mac_365_ja_jp_profile());
  EXPECT_EQ(display.status, DisplayStatus::kOverflow);
  EXPECT_EQ(display.text, "########");
}

TEST(NumberFormatScientific, WideRequiredMantissaPaddingRemainsRenderable) {
  const std::string format = std::string(310, '0') + "E+00";
  std::string out;
  EXPECT_EQ(apply_format(1E308, format, out), FormatStatus::kOk);
  EXPECT_EQ(out.size(), 314U);
  EXPECT_EQ(out.substr(310), "E+00");
  EXPECT_EQ(out[0], '0');
  EXPECT_EQ(out[1], '1');
}

TEST(NumberFormatScientific, ExcessiveEngineeringGroupOverflowLeavesOutputUntouched) {
  const std::string format = std::string(1000, '0') + "E+00";
  std::string out = "prefix";
  EXPECT_EQ(apply_format(1E-300, format, out), FormatStatus::kOverflow);
  EXPECT_EQ(out, "prefix");
}

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

TEST(NumberFormatIntegers, PadPlaceholderUsesSpace) {
  // `?` reserves a position without zero-padding it.
  EXPECT_EQ(Render(5.0, "?0"), " 5");
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

TEST(NumberFormatScale, TrailingCommaAfterFractionDigits) {
  // The scaling comma sits behind the fractional placeholder, so it follows
  // the last digit placeholder rather than the decimal point.
  EXPECT_EQ(Render(1234567.0, "0.0,"), "1234.6");
}

TEST(NumberFormatScale, DoubleTrailingCommaAfterFractionDigits) {
  EXPECT_EQ(Render(1234567.0, "0.00,,"), "1.23");
}

TEST(NumberFormatScale, TrailingCommaKeepsGroupSeparator) {
  // The leading `#,##` still groups; only the comma past the last digit
  // placeholder scales.
  EXPECT_EQ(Render(1234567.0, "#,##0.0,"), "1,234.6");
}

TEST(NumberFormatScale, TrailingCommaBeforeLiteralSuffix) {
  EXPECT_EQ(Render(1234567.0, "#,##0.0,\"K\""), "1,234.6K");
}

TEST(NumberFormatScale, GroupSeparatorWithFractionIsNotScaled) {
  // A group separator inside the integer part never scales the value.
  EXPECT_EQ(Render(1234567.891, "#,##0.00"), "1,234,567.89");
}

// ---------------------------------------------------------------------------
// Decimal ties (rounded on Excel's 15-significant-digit decimal view)
// ---------------------------------------------------------------------------

TEST(NumberFormatTies, HalfAwayFromZeroBelowBinaryValue) {
  // 1.005 is stored just under its decimal value; Excel still rounds up.
  EXPECT_EQ(Render(1.005, "0.00"), "1.01");
}

TEST(NumberFormatTies, HalfAwayFromZeroNegative) {
  EXPECT_EQ(Render(-1.005, "0.00"), "-1.01");
}

TEST(NumberFormatTies, HalfAwayFromZeroCarriesIntoIntegerPart) {
  EXPECT_EQ(Render(9.995, "0.00"), "10.00");
  EXPECT_EQ(Render(99.995, "0.00"), "100.00");
}

TEST(NumberFormatTies, HalfAwayFromZeroMoreTies) {
  EXPECT_EQ(Render(2.675, "0.00"), "2.68");
  EXPECT_EQ(Render(8.835, "0.00"), "8.84");
  EXPECT_EQ(Render(0.045, "0.00"), "0.05");
}

TEST(NumberFormatTies, BelowTieStillRoundsDown) {
  // Only genuine 15-digit ties move: 1.0049999 is short of the boundary.
  EXPECT_EQ(Render(1.0049999, "0.00"), "1.00");
  EXPECT_EQ(Render(0.0449, "0.00"), "0.04");
}

TEST(NumberFormatTies, BinaryResidueIsRemoved) {
  // 0.1 + 0.2 == 0.30000000000000004; the 15-digit view is a flat 0.3.
  EXPECT_EQ(Render(0.1 + 0.2, "0.00"), "0.30");
  EXPECT_EQ(Render(0.1 + 0.2, "0.00000000000000000"), "0.30000000000000000");
}

TEST(NumberFormatTies, LargeIntegerIsNotNudged) {
  // The tie correction must not perturb magnitudes whose ULP exceeds 0.5.
  EXPECT_EQ(Render(9.0e15, "0"), "9000000000000000");
}

TEST(NumberFormatTies, IntegerPast15SignificantDigitsIsZeroed) {
  // 2^60 == 1152921504606846976 (19 digits) is exactly representable, so
  // "%.0f" prints every digit verbatim; Excel shows only the first 15
  // significant digits and zeros the rest, rounding the 15th against the
  // 16th (here 115292150460684|6... rounds the last kept digit up).
  const double v = std::pow(2.0, 60.0);
  EXPECT_EQ(Render(v, "0"), "1152921504606850000");
}

TEST(NumberFormatTies, IntegerPast15SignificantDigitsRoundsDown) {
  // 10^16 exactly: the first 15 significant digits are "100000000000000"
  // followed by a 16th-digit "0", so no round-up carry fires.
  EXPECT_EQ(Render(1.0e16, "0"), "10000000000000000");
}

TEST(NumberFormatTies, ScaledTieUsesScaledValue) {
  // 1005 / 1000 = 1.005, which then ties away from zero.
  EXPECT_EQ(Render(1005.0, "0.00,"), "1.01");
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

TEST(NumberFormatScientific, ExcelMantissaNormalizationAndStoredDisplayParity) {
  struct Case {
    double value;
    const char* format;
    const char* expected;
  };
  const Case cases[] = {
      {9.999, "0.00E+00", "1.00E+01"},
      {-9.999, "0.00E+00", "-1.00E+01"},
      {999.999, "##0.00E+00", "1.00E+03"},
      {12345.0, "00.00E+00", "01.23E+04"},
      {12345.0, "##0.00E+00", "12.35E+03"},
      {0.012345, "##0.00E+00", "12.35E-03"},
      {12345.0, "###.00E+00", "12.35E+03"},
      {0.0012345, "#.00E+00", "1.23E-03"},
      {1234.0, "00.00E+00", "12.34E+02"},
      {1234.0, "000.00E+00", "001.23E+03"},
      {1234.0, "##00.00E+00", "1234.00E+00"},
      {12345.0, "?0.00E+00", " 1.23E+04"},
      {12345.0, "??0.00E+00", " 12.35E+03"},
      {1.0, "0.##E+00", "1.E+00"},
      {1.23, "\"USD \"0.00E+00\" end\"", "USD 1.23E+00 end"},
  };
  for (const Case& test_case : cases) {
    std::string out;
    EXPECT_EQ(apply_format(test_case.value, test_case.format, out), FormatStatus::kOk) << test_case.format;
    EXPECT_EQ(out, test_case.expected) << test_case.format;

    const auto display =
        format_value_for_display(Value::number(test_case.value), test_case.format, false, mac_365_ja_jp_profile());
    EXPECT_EQ(display.status, DisplayStatus::kOk) << test_case.format;
    EXPECT_EQ(display.text, test_case.expected) << test_case.format;
  }
}

TEST(NumberFormatScientific, OptionalFractionPlaceholdersTrimOnlyOptionalDigits) {
  EXPECT_EQ(Render(1.201, "0.#0"), "1.20");
  EXPECT_EQ(Render(1.201, "0.0#"), "1.2");
}

TEST(NumberFormatNumeric, RepeatedPointsAndInterleavedLiteralsRemainStable) {
  EXPECT_EQ(Render(1.2345, "0.00.00"), "1.23.45");
  EXPECT_EQ(Render(1.23995, "0.00.00"), "1.24.00");
  EXPECT_EQ(Render(1.2345, "0.00\"x\"00"), "1.23x45");
}

TEST(NumberFormatSections, NumericAndTextPlaceholderIsInvalid) {
  for (const char* format : {"0 @", "0.00@", "yyyy@", "General@", "0.00E+00@"}) {
    std::string out = "prefix";
    EXPECT_EQ(apply_format(5.0, format, out), FormatStatus::kValueError) << format;
    EXPECT_EQ(out, "prefix") << format;

    const auto display = format_value_for_display(Value::number(5.0), format, false, mac_365_ja_jp_profile());
    EXPECT_EQ(display.status, DisplayStatus::kInvalidFormat) << format;
    EXPECT_EQ(display.text, "5") << format;
  }
}

TEST(NumberFormatSections, RepeatedPercentAndScientificPercentAreInvalid) {
  for (const char* format : {"0%%", "0.00E+00%"}) {
    std::string out = "prefix";
    EXPECT_EQ(apply_format(5.0, format, out), FormatStatus::kValueError) << format;
    EXPECT_EQ(out, "prefix") << format;

    const auto display = format_value_for_display(Value::number(5.0), format, false, mac_365_ja_jp_profile());
    EXPECT_EQ(display.status, DisplayStatus::kInvalidFormat) << format;
    EXPECT_EQ(display.text, "5") << format;
  }
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

TEST(NumberFormatLiteral, IncompleteSyntaxReturnsErrorWithoutAppending) {
  const eval::ScopedEvalProfile profile_scope(mac_365_ja_jp_profile());
  for (const char* code : {"0!", "0\\", "0_", "0*", "0\"open", "[h", "[$-411", "@!", "@\\", "@_", "@*", "@\"open"}) {
    SCOPED_TRACE(code);
    std::string out = "prefix";
    EXPECT_EQ(apply_format(12.0, code, out), FormatStatus::kValueError);
    EXPECT_EQ(out, "prefix");
    EXPECT_EQ(apply_text_format("hello", code, out), FormatStatus::kValueError);
    EXPECT_EQ(out, "prefix");
    const auto number_display = format_value_for_display(Value::number(12), code, false, mac_365_ja_jp_profile());
    EXPECT_EQ(number_display.status, DisplayStatus::kInvalidFormat);
    EXPECT_EQ(number_display.text, "12");
    const auto text_display = format_value_for_display(Value::text("hello"), code, false, mac_365_ja_jp_profile());
    EXPECT_EQ(text_display.status, DisplayStatus::kInvalidFormat);
    EXPECT_EQ(text_display.text, "hello");
  }
}

TEST(NumberFormatLiteral, CompleteEscapesAndDelimitersRemainValid) {
  EXPECT_EQ(Render(12, "0\\!"), "12!");
  EXPECT_EQ(Render(12, "0!x"), "12x");
  EXPECT_EQ(Render(12, "0\"open\""), "12open");
  std::string out;
  EXPECT_EQ(apply_text_format("hello", "@\\!", out), FormatStatus::kOk);
  EXPECT_EQ(out, "hello!");
}

// Mac Excel 16.111.3 (ja-JP) accepts the syntax-bearing full-width forms
// below. The normalizer intentionally folds only those forms outside opaque
// quoted/escaped payloads; a full-width A remains a literal.
TEST(NumberFormatJaJpFullWidth, NumericPlaceholdersAndPunctuation) {
  EXPECT_EQ(Render(5.0, "０.００"), "5.00");
  EXPECT_EQ(Render(5.0, "１0"), "15");
  EXPECT_EQ(Render(5.0, "０１"), "51");
  EXPECT_EQ(Render(5.0, "＃0"), "5");
  EXPECT_EQ(Render(5.0, "？0"), " 5");
  EXPECT_EQ(Render(5.0, "０％"), "500%");
  EXPECT_EQ(Render(1234.0, "＃，＃＃０"), "1,234");
  EXPECT_EQ(Render(5.0, "Ａ0"), "Ａ5");
}

TEST(NumberFormatJaJpFullWidth, SectionsAndBracketOperators) {
  EXPECT_EQ(Render(5.0, "０．０；(０．０)"), "5.0");
  EXPECT_EQ(Render(-5.0, "０．０；(０．０)"), "(5.0)");
  EXPECT_EQ(Render(5.0, "［＞１］０；０"), "5");
  EXPECT_EQ(Render(0.5, "０．００；［＜１］０"), "0.50");
}

TEST(NumberFormatJaJpFullWidth, DBNumDirective) {
  // Full-width DBNum spelling is syntax inside a bracket directive. DBNum3
  // emits full-width Arabic digits in the existing ja-JP renderer.
  EXPECT_EQ(Render(123.0, "［ＤＢＮｕｍ３］０"), "１２３");
}

TEST(NumberFormatJaJpFullWidth, EscapesQuotesAndMultibytePayloads) {
  EXPECT_EQ(Render(5.0, "\"＃０\"0"), "＃０5");
  EXPECT_EQ(Render(5.0, "\\円0"), "円5");
  EXPECT_EQ(Render(5.0, "＼０"), "０");
  EXPECT_EQ(Render(5.0, "＿円0"), " 5");
  EXPECT_EQ(Render(5.0, "＊円0"), "5");
}

TEST(NumberFormatJaJpFullWidth, ScientificAndMalformedUtf8) {
  EXPECT_EQ(Render(12345.0, "０．００Ｅ＋００"), "1.23E+04");
  EXPECT_EQ(Render(0.0001234, "０．００Ｅ－００"), "1.23E-04");

  std::string malformed;
  malformed.push_back(static_cast<char>(0xC3));
  malformed.push_back('0');
  std::string out;
  EXPECT_EQ(apply_format(5.0, malformed, out), FormatStatus::kOk);
  ASSERT_EQ(out.size(), 2U);
  EXPECT_EQ(static_cast<unsigned char>(out[0]), 0xC3U);
  EXPECT_EQ(out[1], '5');
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

TEST(NumberFormatSections, TextSectionIsReachableThroughPublicFormatter) {
  std::string out;
  EXPECT_EQ(apply_text_format("abc", "0.0;(0.0);\"z\";\"text: \"@", out), FormatStatus::kOk);
  EXPECT_EQ(out, "text: abc");
}

TEST(NumberFormatSections, EmptyTextTakesTheTextSection) {
  std::string out;
  EXPECT_EQ(apply_text_format("", "0;0;0;\"<\"@\">\"", out), FormatStatus::kOk);
  EXPECT_EQ(out, "<>");
}

TEST(NumberFormatSections, TextWithoutATextSectionIsUnchanged) {
  std::string out;
  EXPECT_EQ(apply_text_format("abc", "0.00", out), FormatStatus::kOk);
  EXPECT_EQ(out, "abc");
}

// ---------------------------------------------------------------------------
// Bracketed specifiers: named colours tolerated, locale-currency discarded,
// unknown qualifiers rejected.
// ---------------------------------------------------------------------------

TEST(NumberFormatBracketed, NamedColorSilentlyDropped) {
  // Excel discards a colour qualifier in TEXT and formats the value with the
  // rest of the section. All eight ja-JP names are recognised.
  EXPECT_EQ(Render(5.0, "[黒]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[青]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[水]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[緑]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[紫]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[赤]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[白]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[黄]0.00"), "5.00");
}

TEST(NumberFormatBracketed, EnglishColorNameIsValueError) {
  // A format string is read in the UI locale, so the English spellings are
  // not colours under the ja-JP profile and fall through to the
  // invalid-bracket path. Excel answers #VALUE! for all of them.
  const eval::ScopedEvalProfile profile_scope(mac_365_ja_jp_profile());
  std::string out;
  EXPECT_EQ(apply_format(5.0, "[Red]0.00", out), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[Blue]0.00", out), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[green]0.00", out), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[Color12]0.00", out), FormatStatus::kValueError);
}

TEST(NumberFormatBracketed, StoredDialectTakesEnglishColorNames) {
  // A stored cell format spells colours in English, case-insensitively, and
  // the ja-JP names are not colours there.
  std::string out;
  for (const char* code : {"[Black]0.00", "[Blue]0.00", "[Cyan]0.00", "[Green]0.00", "[Magenta]0.00", "[Red]0.00",
                           "[White]0.00", "[Yellow]0.00", "[red]0.00", "[Color1]0.00", "[Color56]0.00"}) {
    out.clear();
    EXPECT_EQ(apply_format(5.0, code, out, false, FormatDialect::kStored), FormatStatus::kOk) << code;
    EXPECT_EQ(out, "5.00") << code;
  }
  EXPECT_EQ(apply_format(5.0, "[Color57]0.00", out, false, FormatDialect::kStored), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[\xE8\xB5\xA4]0.00", out, false, FormatDialect::kStored),  // [赤]
            FormatStatus::kValueError);
}

TEST(NumberFormatBracketed, NegativeSectionColorFollowsTheDialect) {
  const eval::ScopedEvalProfile profile_scope(mac_365_ja_jp_profile());
  EXPECT_TRUE(negative_section_has_color("0;[Red]0", FormatDialect::kStored));
  EXPECT_FALSE(negative_section_has_color("[Red]0;0", FormatDialect::kStored));
  EXPECT_TRUE(negative_section_has_color("[Color3]0", FormatDialect::kStored));
  EXPECT_FALSE(negative_section_has_color("0;[Red]0", FormatDialect::kLocalized));
  EXPECT_TRUE(negative_section_has_color("0;[\xE8\xB5\xA4]0", FormatDialect::kLocalized));
}

TEST(NumberFormatBracketed, ColorNameMatchesAsAPrefix) {
  // Anything trailing the name inside the same bracket is ignored, which is
  // how `[水色]` and `[黄色]` come to be accepted.
  EXPECT_EQ(Render(5.0, "[水色]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[黄色]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[赤abc]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[赤 ]0.00"), "5.00");
}

TEST(NumberFormatBracketed, LeadingBlankBeforeColorNameIsValueError) {
  // Trailing bytes are ignored but leading ones are not.
  std::string out;
  EXPECT_EQ(apply_format(5.0, "[ 赤]0.00", out), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[　赤]0.00", out), FormatStatus::kValueError);
}

TEST(NumberFormatBracketed, IndexedColorSilentlyDropped) {
  // `[色N]` for N in 1..56 is dropped the same way a name is.
  EXPECT_EQ(Render(5.0, "[色1]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[色12]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[色56]0.00"), "5.00");
  // Blanks between `色` and the index are allowed, leading zeros are fine,
  // and the digits may be full-width.
  EXPECT_EQ(Render(5.0, "[色 1]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[色001]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[色１]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[色５６]0.00"), "5.00");
  // The index too matches as a prefix.
  EXPECT_EQ(Render(5.0, "[色1x]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "[色1.0]0.00"), "5.00");
}

TEST(NumberFormatBracketed, IndexedColorOutOfRangeIsValueError) {
  std::string out;
  EXPECT_EQ(apply_format(5.0, "[色57]0.00", out), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[色０]0.00", out), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[色00]0.00", out), FormatStatus::kValueError);
  // The digit run is greedy, so this is index 12345 rather than index 1 with
  // `2345` trailing.
  EXPECT_EQ(apply_format(5.0, "[色12345]0.00", out), FormatStatus::kValueError);
}

TEST(NumberFormatBracketed, IndexedColorWithoutAnIndexIsValueError) {
  std::string out;
  EXPECT_EQ(apply_format(5.0, "[色]0.00", out), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[色x]0.00", out), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[色-1]0.00", out), FormatStatus::kValueError);
}

TEST(NumberFormatBracketed, SecondColorInOneSectionIsValueError) {
  // Either bracket alone is inert, but a section carries at most one colour.
  const eval::ScopedEvalProfile profile_scope(mac_365_ja_jp_profile());
  std::string out;
  EXPECT_EQ(apply_format(5.0, "[赤][青]0.00", out), FormatStatus::kValueError);
  // One per section is still fine when the sections differ.
  EXPECT_EQ(Render(-5.0, "0.00;[赤]0.00"), "5.00");
}

TEST(NumberFormatBracketed, ColorCombinesWithOtherDirectives) {
  // A colour does not consume the section's one conditional predicate, and
  // it may sit anywhere in the section.
  EXPECT_EQ(Render(5.0, "[赤][>1]0.00"), "5.00");
  EXPECT_EQ(Render(5.0, "0.00[赤]"), "5.00");
}

TEST(NumberFormatBracketed, ConditionalGtMatch) {
  // `[>100]` predicate holds for 1500: section 0 renders.
  EXPECT_EQ(Render(1500.0, "[>1000]#,##0\\K;0"), "1,500K");
}

TEST(NumberFormatBracketed, ConditionalGtNoMatchFallsThrough) {
  // `[>1000]` fails for 500: section 1 (the `0` arm) renders.
  EXPECT_EQ(Render(500.0, "[>1000]#,##0\\K;0"), "500");
}

TEST(NumberFormatBracketed, ConditionalLeMatchesNegative) {
  // `[<=0]` predicate holds for 0 and for negatives; with a single section
  // the value is rendered verbatim (sign included).
  EXPECT_EQ(Render(-5.0, "[<=0]0.00;0.00"), "-5.00");
  EXPECT_EQ(Render(0.0, "[<=0]0.00;0.00"), "0.00");
}

TEST(NumberFormatBracketed, ConditionalLiteralMinusPreservesSignedPredicate) {
  // Excel keeps the negative sign selected by the value and also copies the
  // explicit literal minus from a matching conditional section.
  EXPECT_EQ(Render(-1.5, "[<=0]-0.00;0.00"), "--1.50");
  EXPECT_EQ(Render(-1.5, "[<0]-0.00;0.00"), "-1.50");
}

TEST(NumberFormatBracketed, ConditionalEqOperator) {
  // `[=42]` matches exactly the predicate value.
  EXPECT_EQ(Render(42.0, "[=42]\"yes\";\"no\""), "yes");
  EXPECT_EQ(Render(7.0, "[=42]\"yes\";\"no\""), "no");
}

TEST(NumberFormatBracketed, ConditionalNeOperator) {
  // `[<>0]` triggers when the value is non-zero.
  EXPECT_EQ(Render(7.0, "[<>0]\"nz\";\"zero\""), "nz");
  EXPECT_EQ(Render(0.0, "[<>0]\"nz\";\"zero\""), "zero");
}

TEST(NumberFormatBracketed, ConditionalInvalidNumberStillRejected) {
  // Predicate-shaped body with a non-numeric tail still surfaces #VALUE!.
  std::string out;
  EXPECT_EQ(apply_format(5.0, "[>abc]0.00", out), FormatStatus::kValueError);
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

TEST(NumberFormatUnderscoreSkip, LiteralTrailingUnderscoreRequiresEscape) {
  // Mac Excel 16.113.3 rejects a dangling `_`; an escaped underscore is literal.
  std::string out = "prefix";
  EXPECT_EQ(apply_format(5.0, "0_", out), FormatStatus::kValueError);
  EXPECT_EQ(out, "prefix");
  EXPECT_EQ(Render(5.0, "0\\_"), "5_");
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
  EXPECT_EQ(RenderEnUs(12.0, "General"), "12");
  EXPECT_EQ(RenderEnUs(-12.0, "General"), "-12");
  // Large-but-still-integral values skip scientific notation when they fit
  // within the fixed-width budget.
  EXPECT_EQ(RenderEnUs(1234567890.0, "General"), "1234567890");
}

TEST(NumberFormatGeneral, FractionTrimmedAndScientific) {
  // 1/3 prints 9 fractional digits (exactly what Mac Excel / IronCalc
  // goldens emit), with trailing zeros trimmed.
  EXPECT_EQ(RenderEnUs(1.0 / 3.0, "General"), "0.333333333");
  EXPECT_EQ(RenderEnUs(-1.0 / 3.0, "General"), "-0.333333333");
  // Large magnitudes switch to scientific with an exponent zero-padded to
  // two digits; trailing mantissa zeros still collapse.
  EXPECT_EQ(RenderEnUs(250000000000.0, "General"), "2.5E+11");
  EXPECT_EQ(RenderEnUs(123456789012.0, "General"), "1.23457E+11");
  EXPECT_EQ(RenderEnUs(-2.7e-18, "General"), "-2.7E-18");
}

TEST(NumberFormatGeneral, JaJpKeywordMatchesEnglishGeneral) {
  // ja-JP spells the built-in General format code "G/標準" (this is the
  // literal numFmtId=0 keyword Excel 365 ja-JP stores, not a translation
  // applied at render time). Must render identically to "General", not be
  // misread as an era-code token (a lone leading 'G' would otherwise scan
  // as `EraG`, a date token).
  EXPECT_EQ(Render(1234.0, "G/\xE6\xA8\x99\xE6\xBA\x96"), "1234");
  EXPECT_EQ(Render(1.0 / 3.0, "G/\xE6\xA8\x99\xE6\xBA\x96"), "0.333333333");
  EXPECT_EQ(RenderEnUs(1234.0, "General"), Render(1234.0, "G/\xE6\xA8\x99\xE6\xBA\x96"));
  // ja-JP rejects the English keyword (locale_tokens.text_general_english).
  const eval::ScopedEvalProfile profile_scope(mac_365_ja_jp_profile());
  std::string out;
  EXPECT_EQ(apply_format(1234.0, "General", out), FormatStatus::kValueError);
}

TEST(NumberFormatGeneral, JaJpKeywordHonoursDbNumQualifier) {
  // `[DBNumN]` spells General's integer part with place units, unlike a
  // digit-token format such as `[DBNum1]0`, which substitutes per digit
  // (locale_tokens.dbnum_general_ja_jp, text_format.text_dbnum1).
  EXPECT_EQ(Render(1234.0, "[DBNum1]G/\xE6\xA8\x99\xE6\xBA\x96"),
            "\xE5\x8D\x83\xE4\xBA\x8C\xE7\x99\xBE\xE4\xB8\x89\xE5\x8D\x81\xE5\x9B\x9B");  // 千二百三十四
  EXPECT_EQ(Render(1234.0, "[DBNum3]G/\xE6\xA8\x99\xE6\xBA\x96"),
            "\xE5\x8D\x83\xEF\xBC\x92\xE7\x99\xBE\xEF\xBC\x93\xE5\x8D\x81\xEF\xBC\x94");  // 千２百３十４
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

// ---------------------------------------------------------------------------
// Fraction format (`# ?/?`, `# ??/??`, ...).
// ---------------------------------------------------------------------------

TEST(NumberFormatFraction, ProperFractionZeroIntegerSuppressed) {
  // `# ?/?` with 0.5: integer 0 is suppressed by `#`, then a literal space,
  // numerator '1', '/', denominator '2'.
  EXPECT_EQ(Render(0.5, "# ?/?"), " 1/2");
}

TEST(NumberFormatFraction, MixedFractionWithIntegerOne) {
  // 1.5 -> "1 1/2": integer 1 emits, literal space, then 1/2.
  EXPECT_EQ(Render(1.5, "# ?/?"), "1 1/2");
}

TEST(NumberFormatFraction, WholeValueBlanksTheFractionComponent) {
  // An exact integer shows no 0/1 fraction, but the separator, numerator,
  // slash and denominator keep their width as blanks, and a zero integer is
  // shown even under `#` (Excel Range.Text).
  EXPECT_EQ(Render(2.0, "# ?/?"), "2    ");
  EXPECT_EQ(Render(-2.0, "# ?/?"), "-2    ");
  EXPECT_EQ(Render(0.0, "# ?/?"), "0    ");
  EXPECT_EQ(Render(1e-05, "# ?/?"), "0    ");
  EXPECT_EQ(Render(123456789012345678.0, "# ?/?"), "123456789012346000    ");
}

TEST(NumberFormatFraction, NegativeMixedFraction) {
  // -1.5 -> "-1 1/2": leading minus, then absolute-value rendering.
  EXPECT_EQ(Render(-1.5, "# ?/?"), "-1 1/2");
}

TEST(NumberFormatFraction, TwoDigitNumeratorDenominator) {
  // 0.123 with `# ??/??`. The Stern-Brocot best 2/2-digit approximation
  // is 8/65 = 0.123076..., padded to 2-wide right-aligned with spaces.
  EXPECT_EQ(Render(0.123, "# ?\?/?\?"), "  8/65");
}

TEST(NumberFormatFraction, ZeroPadPlaceholderDigitZero) {
  // `0/0` (no `?`/`#`, only `0`): leading positions zero-pad rather than
  // space-pad. 0.5 with `# 0/0` -> the integer is 0 with `#` (suppressed),
  // then literal space, "1/2".
  EXPECT_EQ(Render(0.5, "# 0/0"), " 1/2");
}

TEST(NumberFormatFraction, ImproperFractionNoIntegerGroup) {
  // `?/?` (no leading integer group): the full magnitude becomes the
  // numerator/denominator search target. 0.5 -> "1/2" with no integer
  // group and no preceding space.
  EXPECT_EQ(Render(0.5, "?/?"), "1/2");
}

TEST(NumberFormatFraction, PercentScalingAppliesBeforeMixedFraction) {
  EXPECT_EQ(Render(0.0125, "# ?/?%"), "1 1/4%");
}

TEST(NumberFormatFraction, ScalingCommaIsInvalidForFractionFormats) {
  for (const char* format : {"# ?/?,", "# ?/?,\"K\""}) {
    std::string out = "prefix";
    EXPECT_EQ(apply_format(1234.5, format, out), FormatStatus::kValueError) << format;
    EXPECT_EQ(out, "prefix") << format;
  }
}

TEST(NumberFormatFraction, FixedDenominatorAndPlaceholderWidths) {
  EXPECT_EQ(Render(0.3, "# ?/8"), " 2/8");
  EXPECT_EQ(Render(1.3, "# ?/8"), "1 2/8");
  EXPECT_EQ(Render(1.3, "?/8"), "10/8");
  EXPECT_EQ(Render(0.3, "# ?/16"), " 5/16");
  EXPECT_EQ(Render(0.5, "# ?/10"), " 5/10");
  EXPECT_EQ(Render(0.5,
                   "?"
                   "?/??"),
            " 1/2 ");
  EXPECT_EQ(Render(0.123,
                   "# ?/"
                   "??"),
            " 8/65");
  EXPECT_EQ(Render(0.123,
                   "# ??"
                   "/?"),
            "  1/8");
}

TEST(NumberFormatFraction, ImproperFractionExpandsBeyondPlaceholderWidth) {
  EXPECT_EQ(Render(12.5, "?/?"), "25/2");
  EXPECT_EQ(Render(1234.5, "?/?"), "2469/2");
  EXPECT_EQ(Render(1000000.125, "?/?"), "8000001/8");
  EXPECT_EQ(Render(0.00005,
                   "?/"
                   "?????"),
            "1/20000");
  const std::string tiny_expected = "0/1" + std::string(17, ' ');
  EXPECT_EQ(Render(std::ldexp(1.0, -64),
                   "?/"
                   "??????????????????"),
            tiny_expected);
  EXPECT_EQ(Render(std::numeric_limits<double>::denorm_min(),
                   "?/"
                   "??????????????????"),
            tiny_expected);
  const std::string max_expected = "179769313486232" + std::string(294, '0') + "/1";
  EXPECT_EQ(Render(std::numeric_limits<double>::max(), "?/?"), max_expected);
  std::string out = "prefix";
  EXPECT_EQ(apply_format(std::numeric_limits<double>::max(), "# ?/?%", out), FormatStatus::kOverflow);
  EXPECT_EQ(out, "prefix");
}

TEST(NumberFormatFraction, NearEqualApproximationsPreferSmallerDenominator) {
  // Mac Excel 16.113.3 chooses 1/8 at this Farey-neighbour boundary,
  // including the adjacent binary64 values represented by these literals.
  for (double value : {17.0 / 144.0, 0.11805555555555554, 0.11805555555555557}) {
    EXPECT_EQ(Render(value, "?/?"), "1/8") << value;
  }
}

TEST(NumberFormatFraction, CandidatesUseBinary64Arithmetic) {
  // A wider long double (WASM, x86-64 Linux) keeps 0.015 * 100 below 1.5
  // and picks 1/100, 4/1000 and 3/5 here.
  EXPECT_EQ(Render(0.015, "?/100"), "2/100");
  EXPECT_EQ(Render(0.0045, "?/1000"), "5/1000");
  EXPECT_EQ(Render(0.6125, "?/?"), "5/8");
}

TEST(NumberFormatFraction, VariableDenominatorSearchCapsAtSevenDigits) {
  const std::string format = "?/" + std::string(18, '?');
  EXPECT_EQ(Render(1.0 / 9999999.0, format), "1/9999999" + std::string(11, ' '));
  EXPECT_EQ(Render(1.0 / 10000000.0, format), "0/1" + std::string(17, ' '));
  EXPECT_EQ(Render(0.5000000000000001, format), "1/2" + std::string(17, ' '));
  EXPECT_EQ(Render(1e-16, format), "0/1" + std::string(17, ' '));
}

TEST(NumberFormatFraction, FirstReciprocalOutsidePlaceholderRangeRendersZero) {
  const std::string format = "?/" + std::string(6, '?');
  EXPECT_EQ(Render(1.0 / 999999.0, format), "1/999999");
  EXPECT_EQ(Render(1e-6, format), "0/1" + std::string(5, ' '));
  EXPECT_EQ(Render(0.09, "?/?"), "0/1");
  EXPECT_EQ(Render(0.1, "?/?"), "0/1");
  EXPECT_EQ(Render(0.11, "?/?"), "1/9");
  // Later continued-fraction boundaries still compare intermediate ratios.
  EXPECT_EQ(Render(0.3, "?/?"), "2/7");
}

TEST(NumberFormatFraction, ZeroAndOptionalPlaceholderControls) {
  EXPECT_EQ(Render(0.5, "00/00"), "01/02");
  EXPECT_EQ(Render(0.5, "##/##"), "1/2");
  EXPECT_EQ(Render(0.5, "# ?/0"), " 1/2");
  EXPECT_EQ(Render(0.8, "# ?/2"), "1    ");
  EXPECT_EQ(Render(0.75, "# ?/2"), "1    ");
  EXPECT_EQ(Render(0.3, "# ?/8%"), "30    %");
}

TEST(NumberFormatFraction, ZeroPrefixedFixedDenominatorsKeepLiteralWidth) {
  EXPECT_EQ(Render(0.5, "# ?/08"), " 4/0 ");
  EXPECT_EQ(Render(0.5, "# ?/008"), " 4/0  ");
  EXPECT_EQ(Render(0.5, "# ?/0010"), " 5/00  ");
  EXPECT_EQ(Render(0.5, "?/008"), "4/0  ");
  EXPECT_EQ(Render(0.75, "# ?/008"), " 6/0  ");
  EXPECT_EQ(Render(1.0, "# ?/008"), "1      ");
  // An all-zero denominator remains a variable placeholder run.
  EXPECT_EQ(Render(0.5, "# ?/000"), " 1/002");
  EXPECT_EQ(Render(0.5, "# ?/0?"), " 1/02");
  EXPECT_EQ(Render(0.5, "# ?/?0"), " 1/20");
}

TEST(NumberFormatFraction, QuotedAndEscapedSlashIsLiteral) {
  EXPECT_EQ(Render(0.5, "# ?\"/\"?"), "  /1");
  EXPECT_EQ(Render(0.5, "# ?\\/?"), "  /1");
}

TEST(NumberFormatFraction, QuotedAndEscapedFixedDenominatorIsInvalid) {
  for (const char* format : {"# ?/\"8\"", "# ?/\\8"}) {
    std::string out = "prefix";
    EXPECT_EQ(apply_format(0.5, format, out), FormatStatus::kValueError) << format;
    EXPECT_EQ(out, "prefix") << format;
  }
}

TEST(NumberFormatScale, CommaBeforeDecimalScalesIntegerSection) {
  EXPECT_EQ(Render(1234.5, "0,.00"), "1.23");
  EXPECT_EQ(Render(1234.5, "#,##0,.00"), "1.23");
  EXPECT_EQ(Render(1234.5, "0,,.00"), "0.00");
  // A comma between fractional placeholders is a passthrough separator, not
  // a trailing scale marker.
  EXPECT_EQ(Render(1.2345, "0.0,0"), "1.23");
}

}  // namespace
}  // namespace text_format
}  // namespace formulon
