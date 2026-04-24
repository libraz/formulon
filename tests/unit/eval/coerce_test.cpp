// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the scalar coercion helpers in `eval/coerce.{h,cpp}`. The
// focus is `coerce_to_number`'s text branch, which now falls back to the
// shared date / datetime parser whenever `std::strtod` rejects the input.
// This locks in Mac Excel 365 ja-JP coercion behaviour: date-shaped strings
// such as `"2024-01-10"`, `"2024/01/10"`, `"2024年1月10日"`, and
// `"2024-01-10 12:00"` resolve to their serial form wherever any function
// routes through `coerce_to_number`.

#include "eval/coerce.h"

#include <string>

#include "gtest/gtest.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

TEST(CoerceToNumberTextDate, IsoDashedDate) {
  // 2024-01-10 -> serial 45301 (1900-based system).
  auto r = coerce_to_number(Value::text("2024-01-10"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 45301.0);
}

TEST(CoerceToNumberTextDate, SlashSeparatedDate) {
  auto r = coerce_to_number(Value::text("2024/01/10"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 45301.0);
}

TEST(CoerceToNumberTextDate, KanjiDate) {
  auto r = coerce_to_number(Value::text("2024\xE5\xB9\xB4""1\xE6\x9C\x88""10\xE6\x97\xA5"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 45301.0);
}

TEST(CoerceToNumberTextDate, IsoDateWithTime) {
  // 2024-01-10 12:00 -> 45301 + 0.5.
  auto r = coerce_to_number(Value::text("2024-01-10 12:00"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 45301.5);
}

TEST(CoerceToNumberTextDate, TimeOnly) {
  // No date component: result is just the fractional day.
  auto r = coerce_to_number(Value::text("12:00"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 0.5);
}

TEST(CoerceToNumberTextNumeric, PlainDecimalStillFastPath) {
  // strtod must come first; the date-parse fallback only fires on rejection.
  auto r = coerce_to_number(Value::text("3.14"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 3.14);
}

TEST(CoerceToNumberTextNumeric, NegativeInteger) {
  auto r = coerce_to_number(Value::text("-42"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), -42.0);
}

TEST(CoerceToNumberTextNumeric, ScientificNotation) {
  auto r = coerce_to_number(Value::text("1.5e3"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 1500.0);
}

TEST(CoerceToNumberTextRejection, GarbageReturnsValue) {
  auto r = coerce_to_number(Value::text("hello"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextRejection, EmptyStringReturnsValue) {
  auto r = coerce_to_number(Value::text(""));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextRejection, WhitespaceOnlyReturnsValue) {
  auto r = coerce_to_number(Value::text("   "));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextRejection, WhitespacePaddedDateRejected) {
  // strtod path tolerates whitespace via the leading trim, but the date-parse
  // fallback runs against the raw, untrimmed text. Padded date strings must
  // therefore be rejected with #VALUE!, even though the unpadded form
  // ("2024-01-10") coerces to its serial. This matches IronCalc's calc_test
  // fixtures and likely Mac's implicit-coercion contract (DATEVALUE is more
  // permissive — see `parse_date_time_text` callers).
  auto r = coerce_to_number(Value::text(" 2024-01-10 "));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextNumeric, WhitespacePaddedNumericStillCoerces) {
  // Confirm the numeric path keeps its whitespace tolerance (existing
  // behaviour). The trim happens before strtod.
  auto r = coerce_to_number(Value::text(" 100 "));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
}

TEST(CoerceToNumberOtherKinds, BlankIsZero) {
  auto r = coerce_to_number(Value::blank());
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 0.0);
}

TEST(CoerceToNumberOtherKinds, BoolTrueIsOne) {
  auto r = coerce_to_number(Value::boolean(true));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 1.0);
}

TEST(CoerceToNumberOtherKinds, BoolFalseIsZero) {
  auto r = coerce_to_number(Value::boolean(false));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 0.0);
}

TEST(CoerceToNumberOtherKinds, NumberRoundTrips) {
  auto r = coerce_to_number(Value::number(2.5));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 2.5);
}

TEST(CoerceToNumberOtherKinds, ErrorPropagates) {
  auto r = coerce_to_number(Value::error(ErrorCode::NA));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::NA);
}

// Percent-suffixed text: Mac Excel 365 divides by 100 after stripping a
// trailing '%'. Leading '%' is not a percent literal and must stay #VALUE!.

TEST(CoerceToNumberTextPercent, SingleDigit) {
  auto r = coerce_to_number(Value::text("8%"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 0.08);
}

TEST(CoerceToNumberTextPercent, Hundred) {
  auto r = coerce_to_number(Value::text("100%"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 1.0);
}

TEST(CoerceToNumberTextPercent, NegativeHalf) {
  auto r = coerce_to_number(Value::text("-50%"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), -0.5);
}

TEST(CoerceToNumberTextPercent, ScientificBody) {
  auto r = coerce_to_number(Value::text("1.5e2%"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 1.5);
}

TEST(CoerceToNumberTextPercent, LeadingPercentRejected) {
  auto r = coerce_to_number(Value::text("%5"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

// Currency-prefixed/suffixed text: Mac Excel 365 accepts the common symbols
// `$ ¢ £ ¥ € ₩` on either side of a numeric body.

TEST(CoerceToNumberTextCurrency, LeadingDollar) {
  auto r = coerce_to_number(Value::text("$100"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
}

TEST(CoerceToNumberTextCurrency, LeadingEuro) {
  auto r = coerce_to_number(Value::text("\xE2\x82\xAC""100"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
}

TEST(CoerceToNumberTextCurrency, LeadingYen) {
  auto r = coerce_to_number(Value::text("\xC2\xA5""1000"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 1000.0);
}

TEST(CoerceToNumberTextCurrency, LeadingPoundWithFraction) {
  auto r = coerce_to_number(Value::text("\xC2\xA3""42.5"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 42.5);
}

TEST(CoerceToNumberTextCurrency, LeadingDollarNegativeBody) {
  auto r = coerce_to_number(Value::text("$-100"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), -100.0);
}

TEST(CoerceToNumberTextCurrency, TrailingDollar) {
  auto r = coerce_to_number(Value::text("100$"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
}

TEST(CoerceToNumberTextCurrency, TrailingEuro) {
  auto r = coerce_to_number(Value::text("100\xE2\x82\xAC"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
}

TEST(CoerceToNumberTextCurrency, TrailingYen) {
  auto r = coerce_to_number(Value::text("1000\xC2\xA5"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 1000.0);
}

TEST(CoerceToNumberTextCurrency, LeadingEuroPadded) {
  // Outer ASCII trim strips the spaces; the currency strip handles `€`.
  auto r = coerce_to_number(Value::text(" \xE2\x82\xAC""100 "));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
}

TEST(CoerceToNumberTextCurrency, TrailingEuroPadded) {
  auto r = coerce_to_number(Value::text(" 100\xE2\x82\xAC "));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
}

TEST(CoerceToNumberTextCurrency, LeadingDollarPadded) {
  auto r = coerce_to_number(Value::text("  $50  "));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 50.0);
}

TEST(CoerceToNumberTextCurrency, BothEndsRejected) {
  // Mismatched markers: strtod fails on either trim, so the whole input is
  // #VALUE!. IronCalc behaviour is undefined here and Mac is conservative.
  auto r = coerce_to_number(Value::text("$100\xE2\x82\xAC"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextCurrency, CurrencyOnlyRejected) {
  auto r = coerce_to_number(Value::text("\xE2\x82\xAC"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextCurrency, NonNumericBodyRejected) {
  auto r = coerce_to_number(Value::text("$abc"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
