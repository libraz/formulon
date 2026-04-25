// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for TEXT, VALUE, and NUMBERVALUE. Each test parses a
// formula source and evaluates it through the default registry.

#include "value.h"

#include <cmath>
#include <string_view>

#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace formulon {
namespace eval {
namespace {

Value EvalSource(std::string_view src) {
  static thread_local Arena parse_arena;
  static thread_local Arena eval_arena;
  parse_arena.reset();
  eval_arena.reset();
  parser::Parser p(src, parse_arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return evaluate(*root, eval_arena);
}

// ---------------------------------------------------------------------------
// TEXT
// ---------------------------------------------------------------------------

TEST(TextFunctionText, IntegerFormat) {
  const Value v = EvalSource("=TEXT(1234, \"#,##0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "1,234");
}

TEST(TextFunctionText, TwoDecimals) {
  const Value v = EvalSource("=TEXT(3.14159, \"0.00\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "3.14");
}

TEST(TextFunctionText, Percent) {
  const Value v = EvalSource("=TEXT(0.1234, \"0.00%\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "12.34%");
}

TEST(TextFunctionText, Negative) {
  const Value v = EvalSource("=TEXT(-5, \"0.00;(0.00)\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "(5.00)");
}

TEST(TextFunctionText, EmptyFormat) {
  const Value v = EvalSource("=TEXT(42, \"\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "");
}

TEST(TextFunctionText, IsoDate) {
  const Value v = EvalSource("=TEXT(45366, \"yyyy-mm-dd\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "2024-03-15");
}

TEST(TextFunctionText, KanjiDate) {
  const Value v = EvalSource("=TEXT(45366, \"yyyy年m月d日\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "2024年3月15日");
}

TEST(TextFunctionText, NonNumericTextIsValueError) {
  // Mac Excel ja-JP rejects TEXT with a non-numeric text first argument.
  const Value v = EvalSource("=TEXT(\"abc\", \"text is @\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextFunctionText, BoolTrueReturnsUppercaseText) {
  // Bool values ignore the format string entirely.
  const Value v = EvalSource("=TEXT(TRUE, \"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "TRUE");
}

TEST(TextFunctionText, BoolFalseReturnsUppercaseText) {
  const Value v = EvalSource("=TEXT(FALSE, \"0\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "FALSE");
}

TEST(TextFunctionText, NumericTextStillCoerces) {
  // Numeric strings still parse and render through the format.
  const Value v = EvalSource("=TEXT(\"42\", \"0.00\")");
  ASSERT_TRUE(v.is_text());
  EXPECT_EQ(v.as_text(), "42.00");
}

TEST(TextFunctionText, EmptyStringIsValueError) {
  // "" is not coercible to a number; TEXT rejects it.
  const Value v = EvalSource("=TEXT(\"\", \"0\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(TextFunctionText, ErrorPropagates) {
  const Value v = EvalSource("=TEXT(#REF!, \"0\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

// ---------------------------------------------------------------------------
// VALUE
// ---------------------------------------------------------------------------

TEST(ValueFunction, IntegerString) {
  const Value v = EvalSource("=VALUE(\"123\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 123.0);
}

TEST(ValueFunction, DecimalString) {
  const Value v = EvalSource("=VALUE(\"1.5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1.5);
}

TEST(ValueFunction, ScientificString) {
  const Value v = EvalSource("=VALUE(\"1.5e2\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 150.0);
}

TEST(ValueFunction, WithCommaThousands) {
  const Value v = EvalSource("=VALUE(\"1,234.5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1234.5);
}

TEST(ValueFunction, WithPercent) {
  const Value v = EvalSource("=VALUE(\"50%\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(ValueFunction, WithDollarPrefix) {
  const Value v = EvalSource("=VALUE(\"$100\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 100.0);
}

TEST(ValueFunction, WithLeadingWhitespace) {
  const Value v = EvalSource("=VALUE(\"   42   \")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

TEST(ValueFunction, NegativeSign) {
  const Value v = EvalSource("=VALUE(\"-3.14\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), -3.14);
}

TEST(ValueFunction, NumberPassthrough) {
  const Value v = EvalSource("=VALUE(5)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 5.0);
}

TEST(ValueFunction, BoolRejected) {
  const Value v = EvalSource("=VALUE(TRUE)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, NonNumericString) {
  const Value v = EvalSource("=VALUE(\"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, ErrorPropagates) {
  const Value v = EvalSource("=VALUE(#REF!)");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Ref);
}

TEST(ValueFunction, EuroSuffixInteger) {
  // Mac Excel 365 ja-JP accepts Euro as a trailing currency suffix.
  const Value v = EvalSource("=VALUE(\"23\xE2\x82\xAC\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 23.0);
}

TEST(ValueFunction, EuroSuffixSmall) {
  const Value v = EvalSource("=VALUE(\"12\xE2\x82\xAC\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 12.0);
}

TEST(ValueFunction, EuroPrefix) {
  const Value v = EvalSource("=VALUE(\"\xE2\x82\xAC""23\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 23.0);
}

TEST(ValueFunction, DollarSuffixRejected) {
  // Mac Excel ja-JP rejects `$` as a trailing suffix (only Euro is
  // bidirectional).
  const Value v = EvalSource("=VALUE(\"23$\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, YenKanjiSuffixRejected) {
  // The yen kanji (U+5186) is not accepted by Mac Excel.
  const Value v = EvalSource("=VALUE(\"23\xE5\x86\x86\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, InvalidThousandsGrouping) {
  // `"12,34"` is rejected by Mac: first group is 2 digits (OK), but the
  // final group must be exactly 3 digits.
  const Value v = EvalSource("=VALUE(\"12,34\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, ValidThousandsGrouping) {
  const Value v = EvalSource("=VALUE(\"1,234\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1234.0);
}

TEST(ValueFunction, ValidThousandsGroupingMultiple) {
  const Value v = EvalSource("=VALUE(\"1,234,567\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1234567.0);
}

TEST(ValueFunction, FourDigitFinalGroupRejected) {
  const Value v = EvalSource("=VALUE(\"1,2345\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, LeadingGroupSepRejected) {
  const Value v = EvalSource("=VALUE(\",234\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, TrailingGroupSepRejected) {
  const Value v = EvalSource("=VALUE(\"1,\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(ValueFunction, ThousandsWithDecimalAccepted) {
  const Value v = EvalSource("=VALUE(\"1,234.56\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 1234.56);
}

TEST(ValueFunction, IsoDate) {
  const Value v = EvalSource("=VALUE(\"2024-03-15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 45366.0);
}

TEST(ValueFunction, TimeOnly) {
  const Value v = EvalSource("=VALUE(\"13:30\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5625);
}

TEST(ValueFunction, DateAndTime) {
  const Value v = EvalSource("=VALUE(\"2024-03-15 12:00\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 45366.5);
}

// ---------------------------------------------------------------------------
// NUMBERVALUE
// ---------------------------------------------------------------------------

TEST(NumberValueFunction, DefaultSeparators) {
  const Value v = EvalSource("=NUMBERVALUE(\"1,234.5\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1234.5);
}

TEST(NumberValueFunction, CustomDecimalSep) {
  // European-style: decimal is "," and group is "."
  const Value v = EvalSource("=NUMBERVALUE(\"1.234,5\", \",\", \".\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1234.5);
}

TEST(NumberValueFunction, TrailingPercent) {
  const Value v = EvalSource("=NUMBERVALUE(\"50%\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.5);
}

TEST(NumberValueFunction, MultipleTrailingPercents) {
  // Each `%` multiplies by 0.01: 50 * 0.01 * 0.01 = 0.005.
  const Value v = EvalSource("=NUMBERVALUE(\"50%%\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 0.005);
}

TEST(NumberValueFunction, CommaDecimalOnlyNoGroupCollision) {
  // With only `decimal_sep = ","` supplied, grouping is disabled so the
  // default group sep of `,` cannot collide.
  const Value v = EvalSource("=NUMBERVALUE(\"3,14\", \",\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.14);
}

TEST(NumberValueFunction, WhitespaceTrimmed) {
  const Value v = EvalSource("=NUMBERVALUE(\"  3.14  \")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 3.14);
}

TEST(NumberValueFunction, Scientific) {
  const Value v = EvalSource("=NUMBERVALUE(\"1.5E3\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 1500.0);
}

TEST(NumberValueFunction, RejectNonNumeric) {
  const Value v = EvalSource("=NUMBERVALUE(\"abc\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(NumberValueFunction, AcceptsDateString) {
  // Mac Excel ja-JP NUMBERVALUE accepts date strings after the numeric
  // path fails (unlike Microsoft's docs). 2024-03-15 -> serial 45366.
  const Value v = EvalSource("=NUMBERVALUE(\"2024-03-15\")");
  ASSERT_TRUE(v.is_number());
  EXPECT_DOUBLE_EQ(v.as_number(), 45366.0);
}

TEST(NumberValueFunction, SameSeparatorRejected) {
  const Value v = EvalSource("=NUMBERVALUE(\"1.5\", \".\", \".\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(NumberValueFunction, EmptyDecimalSepRejected) {
  const Value v = EvalSource("=NUMBERVALUE(\"1.5\", \"\")");
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Value);
}

TEST(NumberValueFunction, NumericPassthrough) {
  // First arg coerces to text via Formulon's shortest-double form, so this
  // succeeds.
  const Value v = EvalSource("=NUMBERVALUE(42)");
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 42.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
