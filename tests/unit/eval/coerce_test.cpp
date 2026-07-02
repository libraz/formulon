// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for the scalar coercion helpers in `eval/coerce.{h,cpp}`. The
// focus is `coerce_to_number`'s text branch, which now falls back to the
// shared date / datetime parser whenever `std::strtod` rejects the input.
// This locks in Mac Excel 365 ja-JP coercion behaviour: date-shaped strings
// such as `"2024-01-10"`, `"2024/01/10"`, `"2024年1月10日"`, and
// `"2024-01-10 12:00"` resolve to their serial form wherever any function
// routes through `coerce_to_number`.

#include "eval/coerce.h"

#include <locale.h>

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <string>

#include "eval/number_parse.h"
#include "gtest/gtest.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Builds an arena-owned Array `Value` for the matrix-strict / collect_numerics
// tests. Mirrors `MakeArray` in `spill_committer_test.cpp` — keep them in
// sync if the `ArrayValue` shape changes.
Value MakeArray(Arena& arena, std::uint32_t rows, std::uint32_t cols, std::initializer_list<Value> entries) {
  Value* cells = arena.create_array<Value>(static_cast<std::size_t>(rows) * cols);
  std::size_t i = 0;
  for (const Value& v : entries) {
    cells[i++] = v;
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  arr->rows = rows;
  arr->cols = cols;
  arr->cells = cells;
  return Value::array(arr);
}

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
  auto r =
      coerce_to_number(Value::text("2024\xE5\xB9\xB4"
                                   "1\xE6\x9C\x88"
                                   "10\xE6\x97\xA5"));
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

// Currency-prefixed/suffixed text: Mac Excel 365 (ja-JP) accepts ONLY the
// symbols {$, ¥, ￥, €}, on the leading OR trailing side but not both. `£`,
// `¢`, `₩`, and the kanji `円` are rejected. The suffix position accepts only
// `€`; a trailing `$` / `¥` is #VALUE!. VALUE() and implicit coercion agree.

TEST(CoerceToNumberTextCurrency, LeadingDollar) {
  auto r = coerce_to_number(Value::text("$100"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
}

TEST(CoerceToNumberTextCurrency, LeadingEuro) {
  auto r =
      coerce_to_number(Value::text("\xE2\x82\xAC"
                                   "100"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
}

TEST(CoerceToNumberTextCurrency, LeadingYen) {
  auto r =
      coerce_to_number(Value::text("\xC2\xA5"
                                   "1000"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 1000.0);
}

TEST(CoerceToNumberTextCurrency, LeadingPoundRejected) {
  // `£` (U+00A3) is not in the ja-JP currency set, so "£42.5" is #VALUE!.
  auto r =
      coerce_to_number(Value::text("\xC2\xA3"
                                   "42.5"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextCurrency, LeadingDollarNegativeBody) {
  auto r = coerce_to_number(Value::text("$-100"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), -100.0);
}

TEST(CoerceToNumberTextCurrency, TrailingDollarRejected) {
  // Only `€` is accepted as a trailing currency marker; a trailing `$` is
  // #VALUE! (matches VALUE("100$")).
  auto r = coerce_to_number(Value::text("100$"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextCurrency, TrailingEuro) {
  auto r = coerce_to_number(Value::text("100\xE2\x82\xAC"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
}

TEST(CoerceToNumberTextCurrency, TrailingYenRejected) {
  // A trailing half-width `¥` is likewise not an accepted suffix -> #VALUE!.
  auto r = coerce_to_number(Value::text("1000\xC2\xA5"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextCurrency, LeadingEuroPadded) {
  // Outer ASCII trim strips the spaces; the currency strip handles `€`.
  auto r =
      coerce_to_number(Value::text(" \xE2\x82\xAC"
                                   "100 "));
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
  // A currency symbol on BOTH ends ("$100" + trailing euro) is rejected —
  // Mac Excel 365 accepts a currency marker on the leading OR trailing side,
  // never both. VALUE("$100€") and "$100€"+0 both surface #VALUE!.
  auto r = coerce_to_number(Value::text("$100\xE2\x82\xAC"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextCurrency, EuroPrefixOrSuffixEachAccepted) {
  // A single-sided euro is accepted whether leading or trailing.
  auto pre =
      coerce_to_number(Value::text("\xE2\x82\xAC"
                                   "100"));
  ASSERT_TRUE(pre.has_value());
  EXPECT_DOUBLE_EQ(pre.value(), 100.0);
  auto post = coerce_to_number(Value::text("100\xE2\x82\xAC"));
  ASSERT_TRUE(post.has_value());
  EXPECT_DOUBLE_EQ(post.value(), 100.0);
}

TEST(CoerceToNumberTextCurrency, SpaceBetweenSymbolAndNumber) {
  // "$ 100" -> 100 (whitespace between the currency symbol and the number).
  auto r = coerce_to_number(Value::text("$ 100"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
}

TEST(CoerceToNumberTextCurrency, SignBeforeOrAfterSymbol) {
  // The sign may sit on either side of the currency symbol.
  auto before = coerce_to_number(Value::text("-$100"));
  ASSERT_TRUE(before.has_value());
  EXPECT_DOUBLE_EQ(before.value(), -100.0);
  auto after = coerce_to_number(Value::text("$-100"));
  ASSERT_TRUE(after.has_value());
  EXPECT_DOUBLE_EQ(after.value(), -100.0);
}

TEST(CoerceToNumberTextCurrency, YenKanjiSuffixRejected) {
  // The kanji "円" (U+5186) is NOT accepted as a currency marker even under
  // ja-JP: "100円" surfaces #VALUE!.
  auto r = coerce_to_number(Value::text("100\xE5\x86\x86"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextCurrency, CentRejected) {
  // `¢` (U+00A2) is outside the ja-JP currency set.
  auto r =
      coerce_to_number(Value::text("\xC2\xA2"
                                   "100"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextCurrency, WonRejected) {
  // `₩` (U+20A9) is outside the ja-JP currency set.
  auto r =
      coerce_to_number(Value::text("\xE2\x82\xA9"
                                   "100"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CoerceToNumberTextCurrency, FullWidthYenAccepted) {
  // The full-width yen sign `￥` (U+FFE5) is accepted: "￥100" -> 100.
  auto r =
      coerce_to_number(Value::text("\xEF\xBF\xA5"
                                   "100"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 100.0);
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

// ---------------------------------------------------------------------------
// Locale-aware numeric coercion shared with VALUE(): thousands grouping,
// accounting parentheses, full-width digits, currency + grouping, and
// surrounding whitespace all coerce successfully wherever a number is
// expected (arithmetic operators, criteria).
// ---------------------------------------------------------------------------

TEST(CoerceToNumberTextLocale, ThousandsGrouped) {
  auto r = coerce_to_number(Value::text("1,000"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 1000.0);
}

TEST(CoerceToNumberTextLocale, AccountingParensNegate) {
  auto r = coerce_to_number(Value::text("(100)"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), -100.0);
}

TEST(CoerceToNumberTextLocale, FullWidthDigits) {
  // "１２３" (U+FF11 U+FF12 U+FF13) folds to 123.
  auto r = coerce_to_number(Value::text("\xEF\xBC\x91\xEF\xBC\x92\xEF\xBC\x93"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 123.0);
}

TEST(CoerceToNumberTextLocale, YenWithGrouping) {
  // "¥1,000" (U+00A5 prefix + grouped digits) coerces to 1000.
  auto r =
      coerce_to_number(Value::text("\xC2\xA5"
                                   "1,000"));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 1000.0);
}

TEST(CoerceToNumberTextLocale, SurroundingWhitespace) {
  auto r = coerce_to_number(Value::text(" 12 "));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 12.0);
}

TEST(CoerceToNumberTextLocale, MalformedGroupingRejected) {
  // "12,34" is not a valid 3-digit grouping and stays #VALUE!.
  auto r = coerce_to_number(Value::text("12,34"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// Locale-independent numeric parsing: the evaluator always uses '.' as the
// decimal separator regardless of the host process's LC_NUMERIC.
// ---------------------------------------------------------------------------

TEST(ParseDoubleCLocale, DecimalIsAlwaysDotUnderCLocale) {
  // Under the C locale a comma never acts as a decimal separator: "1.5"
  // parses fully, and "1,5" stops at the comma. This holds no matter what
  // the host LC_NUMERIC is set to, because the parser swaps to a C locale.
  char* end = nullptr;
  EXPECT_DOUBLE_EQ(parse_double_c_locale("1.5", &end), 1.5);
  EXPECT_EQ(*end, '\0');
  EXPECT_DOUBLE_EQ(parse_double_c_locale("1,5", &end), 1.0);
  EXPECT_EQ(*end, ',');
}

TEST(ParseDoubleCLocale, UnaffectedByCommaDecimalHostLocale) {
  // Install a comma-decimal locale on THIS thread only (uselocale is
  // thread-local, so parallel ctest workers are unaffected). "1.5" must
  // still parse to 1.5 — proving numeric parsing ignores the host locale.
  // If the locale is not installed on the host, the assertion below still
  // holds via the default C locale.
  const locale_t comma_locale = newlocale(LC_NUMERIC_MASK, "de_DE.UTF-8", static_cast<locale_t>(0));
  const locale_t previous =
      (comma_locale != static_cast<locale_t>(0)) ? uselocale(comma_locale) : static_cast<locale_t>(0);

  auto r = coerce_to_number(Value::text("1.5"));
  EXPECT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 1.5);

  if (comma_locale != static_cast<locale_t>(0)) {
    uselocale(previous);
    freelocale(comma_locale);
  }
}

// ---------------------------------------------------------------------------
// matrix_strict_number / matrix_strict_number_cell
// ---------------------------------------------------------------------------
//
// Locks in the LINEST / FORECAST.ETS "numbers-only matrix" contract:
// Number passes through (NaN / Inf included — the strict path is for
// raw passthrough), Bool coerces to 1/0, Error propagates, everything
// else (Blank, Text incl. numeric-looking text, Array, Ref, Lambda)
// is `#VALUE!`. Tests mirror the behaviour of the production helper
// `coerce_strict_numeric` in `forecast_ets_lazy.cpp`.

TEST(MatrixStrictNumber, NumberPassesThrough) {
  auto r = matrix_strict_number(Value::number(2.5));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 2.5);
}

TEST(MatrixStrictNumber, BoolTrueIsOne) {
  auto r = matrix_strict_number(Value::boolean(true));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 1.0);
}

TEST(MatrixStrictNumber, BoolFalseIsZero) {
  auto r = matrix_strict_number(Value::boolean(false));
  ASSERT_TRUE(r.has_value());
  EXPECT_DOUBLE_EQ(r.value(), 0.0);
}

TEST(MatrixStrictNumber, NumericLookingTextRejected) {
  // The whole point of the strict variant: even numeric-looking text
  // is `#VALUE!` so LINEST / FORECAST.ETS do not silently accept
  // stringly-typed cells.
  auto r = matrix_strict_number(Value::text("3.14"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(MatrixStrictNumber, NonNumericTextRejected) {
  auto r = matrix_strict_number(Value::text("hello"));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(MatrixStrictNumber, BlankRejected) {
  // Blank is `#VALUE!` here — counted-as-zero only applies to
  // AVERAGE / SUM / etc., not to matrix-strict callers.
  auto r = matrix_strict_number(Value::blank());
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(MatrixStrictNumber, ErrorPropagates) {
  auto r = matrix_strict_number(Value::error(ErrorCode::NA));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::NA);
}

TEST(MatrixStrictNumber, ErrorDiv0Propagates) {
  auto r = matrix_strict_number(Value::error(ErrorCode::Div0));
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Div0);
}

TEST(MatrixStrictNumber, ArrayRejected) {
  // Matrix-strict callers consume cells one at a time; nested Array
  // at a cell position is `#VALUE!`.
  Arena arena;
  const Value arr = MakeArray(arena, 1U, 2U, {Value::number(1.0), Value::number(2.0)});
  auto r = matrix_strict_number(arr);
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(MatrixStrictNumberCell, ForwardsToScalarVariant) {
  // The `(row, col)` variant has the same semantics today; the
  // context is reserved for future structured-log enrichment. A
  // smoke-test on each kind ensures the row / col arguments do not
  // accidentally alter behaviour.
  auto r_num = matrix_strict_number_cell(Value::number(7.5), 3, 4);
  ASSERT_TRUE(r_num.has_value());
  EXPECT_DOUBLE_EQ(r_num.value(), 7.5);

  auto r_text = matrix_strict_number_cell(Value::text("3.14"), 0, 0);
  ASSERT_FALSE(r_text.has_value());
  EXPECT_EQ(r_text.error(), ErrorCode::Value);

  auto r_err = matrix_strict_number_cell(Value::error(ErrorCode::NA), 99, 99);
  ASSERT_FALSE(r_err.has_value());
  EXPECT_EQ(r_err.error(), ErrorCode::NA);
}

// ---------------------------------------------------------------------------
// collect_numerics(Value, NumericCollectPolicy)
// ---------------------------------------------------------------------------
//
// The default policy models the AVERAGE / SUM family: Numbers kept,
// everything non-Number dropped, Error aborts. Flag combinations
// recover the "A"-family and the SMALL / LARGE direct-scalar variant.

TEST(CollectNumericsDefault, ScalarNumber) {
  NumericCollectPolicy policy;
  auto r = collect_numerics(Value::number(3.0), policy);
  ASSERT_TRUE(r.has_value());
  ASSERT_EQ(r.value().size(), 1U);
  EXPECT_DOUBLE_EQ(r.value()[0], 3.0);
}

TEST(CollectNumericsDefault, ScalarBoolDropped) {
  // Default: Bool is dropped silently.
  NumericCollectPolicy policy;
  auto r = collect_numerics(Value::boolean(true), policy);
  ASSERT_TRUE(r.has_value());
  EXPECT_TRUE(r.value().empty());
}

TEST(CollectNumericsDefault, ScalarTextDropped) {
  // Default: Text is dropped silently, even numeric-looking text.
  NumericCollectPolicy policy;
  auto r = collect_numerics(Value::text("3.14"), policy);
  ASSERT_TRUE(r.has_value());
  EXPECT_TRUE(r.value().empty());
}

TEST(CollectNumericsDefault, ScalarBlankDropped) {
  NumericCollectPolicy policy;
  auto r = collect_numerics(Value::blank(), policy);
  ASSERT_TRUE(r.has_value());
  EXPECT_TRUE(r.value().empty());
}

TEST(CollectNumericsDefault, ScalarErrorPropagates) {
  // Default `error_on_error_cell = true`.
  NumericCollectPolicy policy;
  auto r = collect_numerics(Value::error(ErrorCode::Div0), policy);
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Div0);
}

TEST(CollectNumericsDefault, ArrayMixedDropsNonNumber) {
  // Array: row-major flatten, keep only Numbers.
  Arena arena;
  const Value arr = MakeArray(arena, 2U, 3U,
                              {Value::number(1.0), Value::text("hi"), Value::number(2.0), Value::boolean(true),
                               Value::blank(), Value::number(3.0)});
  NumericCollectPolicy policy;
  auto r = collect_numerics(arr, policy);
  ASSERT_TRUE(r.has_value());
  ASSERT_EQ(r.value().size(), 3U);
  EXPECT_DOUBLE_EQ(r.value()[0], 1.0);
  EXPECT_DOUBLE_EQ(r.value()[1], 2.0);
  EXPECT_DOUBLE_EQ(r.value()[2], 3.0);
}

TEST(CollectNumericsDefault, ArrayWithErrorAborts) {
  // First Error encountered in row-major scan order wins.
  Arena arena;
  const Value arr = MakeArray(arena, 1U, 3U, {Value::number(1.0), Value::error(ErrorCode::NA), Value::number(2.0)});
  NumericCollectPolicy policy;
  auto r = collect_numerics(arr, policy);
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::NA);
}

TEST(CollectNumericsIncludeBool, BoolContributesOneAndZero) {
  // `include_bool` flag effect: Bool -> 1/0 instead of dropped.
  Arena arena;
  const Value arr = MakeArray(arena, 1U, 3U, {Value::boolean(true), Value::boolean(false), Value::number(5.0)});
  NumericCollectPolicy policy;
  policy.include_bool = true;
  auto r = collect_numerics(arr, policy);
  ASSERT_TRUE(r.has_value());
  ASSERT_EQ(r.value().size(), 3U);
  EXPECT_DOUBLE_EQ(r.value()[0], 1.0);
  EXPECT_DOUBLE_EQ(r.value()[1], 0.0);
  EXPECT_DOUBLE_EQ(r.value()[2], 5.0);
}

TEST(CollectNumericsIncludeTextLiteral, ParseableTextContributes) {
  // `include_text_numeric_literal` flag: text that parses via
  // coerce_to_number contributes its parsed value.
  Arena arena;
  const Value arr = MakeArray(arena, 1U, 2U, {Value::text("3.14"), Value::number(2.0)});
  NumericCollectPolicy policy;
  policy.include_text_numeric_literal = true;
  auto r = collect_numerics(arr, policy);
  ASSERT_TRUE(r.has_value());
  ASSERT_EQ(r.value().size(), 2U);
  EXPECT_DOUBLE_EQ(r.value()[0], 3.14);
  EXPECT_DOUBLE_EQ(r.value()[1], 2.0);
}

TEST(CollectNumericsIncludeTextLiteral, UnparseableTextDroppedSilentlyByDefault) {
  // `include_text_numeric_literal = true, error_on_text = false`:
  // unparseable text is silently skipped (SMALL / LARGE rule).
  Arena arena;
  const Value arr = MakeArray(arena, 1U, 2U, {Value::text("hello"), Value::number(2.0)});
  NumericCollectPolicy policy;
  policy.include_text_numeric_literal = true;
  policy.error_on_text = false;
  auto r = collect_numerics(arr, policy);
  ASSERT_TRUE(r.has_value());
  ASSERT_EQ(r.value().size(), 1U);
  EXPECT_DOUBLE_EQ(r.value()[0], 2.0);
}

TEST(CollectNumericsErrorOnText, UnparseableTextAborts) {
  // `error_on_text = true`: unparseable text propagates `#VALUE!`
  // (the "A"-family rule).
  Arena arena;
  const Value arr = MakeArray(arena, 1U, 2U, {Value::text("hello"), Value::number(2.0)});
  NumericCollectPolicy policy;
  policy.include_text_numeric_literal = true;
  policy.error_on_text = true;
  auto r = collect_numerics(arr, policy);
  ASSERT_FALSE(r.has_value());
  EXPECT_EQ(r.error(), ErrorCode::Value);
}

TEST(CollectNumericsErrorOnErrorCellDisabled, ErrorDroppedSilently) {
  // `error_on_error_cell = false`: stray Error cells are dropped so
  // the regression family can run its own ordered error-propagation
  // pass first.
  Arena arena;
  const Value arr = MakeArray(arena, 1U, 3U, {Value::number(1.0), Value::error(ErrorCode::NA), Value::number(2.0)});
  NumericCollectPolicy policy;
  policy.error_on_error_cell = false;
  auto r = collect_numerics(arr, policy);
  ASSERT_TRUE(r.has_value());
  ASSERT_EQ(r.value().size(), 2U);
  EXPECT_DOUBLE_EQ(r.value()[0], 1.0);
  EXPECT_DOUBLE_EQ(r.value()[1], 2.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
