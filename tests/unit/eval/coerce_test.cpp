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

}  // namespace
}  // namespace eval
}  // namespace formulon
