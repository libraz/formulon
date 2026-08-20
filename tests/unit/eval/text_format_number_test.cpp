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
  EXPECT_EQ(apply_format(0.0, "0.0;(0.0);\"z\";\"text: \"@", "abc", out), FormatStatus::kOk);
  EXPECT_EQ(out, "text: abc");
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
  std::string out;
  EXPECT_EQ(apply_format(5.0, "[Red]0.00", out), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[Blue]0.00", out), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[green]0.00", out), FormatStatus::kValueError);
  EXPECT_EQ(apply_format(5.0, "[Color12]0.00", out), FormatStatus::kValueError);
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

TEST(NumberFormatBracketed, ConditionalExplicitMinusIsNotDoubled) {
  // A conditional section is not inherently a negative section. Its literal
  // minus supplies the sign itself, so the numeric renderer must use the
  // magnitude rather than adding a second prefix.
  EXPECT_EQ(Render(-1.5, "[<=0]-0.00;0.00"), "-1.50");
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
  EXPECT_EQ(Render(-1.0 / 3.0, "General"), "-0.333333333");
  // Large magnitudes switch to scientific with an exponent zero-padded to
  // two digits; trailing mantissa zeros still collapse.
  EXPECT_EQ(Render(250000000000.0, "General"), "2.5E+11");
  EXPECT_EQ(Render(123456789012.0, "General"), "1.23457E+11");
  EXPECT_EQ(Render(-2.7e-18, "General"), "-2.7E-18");
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

TEST(NumberFormatFraction, WholeValueSuppressesFractionComponent) {
  // An exact integer must not render the synthetic 0/1 approximation.
  EXPECT_EQ(Render(2.0, "# ?/?"), "2");
  EXPECT_EQ(Render(-2.0, "# ?/?"), "-2");
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

}  // namespace
}  // namespace text_format
}  // namespace formulon
