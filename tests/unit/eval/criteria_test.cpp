// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `parse_criterion` and `matches_criterion`. These helpers
// power `COUNTIF`, `SUMIF`, and `AVERAGEIF`; the tests here exercise them
// in isolation so that criterion-level edge cases (operator prefixes,
// numeric-text probing, wildcard handling, blank-cell rules) are covered
// without going through range resolution.

#include "eval/criteria.h"

#include <string>

#include "gtest/gtest.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// parse_criterion
// ---------------------------------------------------------------------------

TEST(CriteriaParse, BareNumberBecomesEqNumeric) {
  const ParsedCriterion c = parse_criterion(Value::number(5.0));
  EXPECT_EQ(c.op, CriteriaOp::Eq);
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_DOUBLE_EQ(c.rhs_number, 5.0);
  EXPECT_FALSE(c.has_wildcard);
}

TEST(CriteriaParse, BareBoolTrueBecomesEqOne) {
  const ParsedCriterion c = parse_criterion(Value::boolean(true));
  EXPECT_EQ(c.op, CriteriaOp::Eq);
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_DOUBLE_EQ(c.rhs_number, 1.0);
}

TEST(CriteriaParse, BareBoolFalseBecomesEqZero) {
  const ParsedCriterion c = parse_criterion(Value::boolean(false));
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_DOUBLE_EQ(c.rhs_number, 0.0);
}

TEST(CriteriaParse, BlankBecomesEqNumericZero) {
  // A blank criterion (e.g. `COUNTIF(range, B1)` with B1 empty) coerces to
  // the numeric 0 criterion, not the empty-text one. Matches Excel 365's
  // "empty cell reference is 0" rule.
  const ParsedCriterion c = parse_criterion(Value::blank());
  EXPECT_EQ(c.op, CriteriaOp::Eq);
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_DOUBLE_EQ(c.rhs_number, 0.0);
}

TEST(CriteriaParse, TextWithoutPrefixIsEqText) {
  const ParsedCriterion c = parse_criterion(Value::text("hello"));
  EXPECT_EQ(c.op, CriteriaOp::Eq);
  EXPECT_FALSE(c.rhs_is_number);
  EXPECT_EQ(c.rhs_text, "hello");
  EXPECT_FALSE(c.has_wildcard);
}

TEST(CriteriaParse, TextWithEqPrefixIsEqText) {
  const ParsedCriterion c = parse_criterion(Value::text("=foo"));
  EXPECT_EQ(c.op, CriteriaOp::Eq);
  EXPECT_FALSE(c.rhs_is_number);
  EXPECT_EQ(c.rhs_text, "foo");
}

TEST(CriteriaParse, TextWithGtPrefixNumericRhs) {
  const ParsedCriterion c = parse_criterion(Value::text(">5"));
  EXPECT_EQ(c.op, CriteriaOp::Gt);
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_DOUBLE_EQ(c.rhs_number, 5.0);
}

TEST(CriteriaParse, TextWithGtEqPrefixNumericRhs) {
  const ParsedCriterion c = parse_criterion(Value::text(">=10"));
  EXPECT_EQ(c.op, CriteriaOp::GtEq);
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_DOUBLE_EQ(c.rhs_number, 10.0);
}

TEST(CriteriaParse, TextWithLtPrefixNumericRhs) {
  const ParsedCriterion c = parse_criterion(Value::text("<0"));
  EXPECT_EQ(c.op, CriteriaOp::Lt);
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_DOUBLE_EQ(c.rhs_number, 0.0);
}

TEST(CriteriaParse, TextWithLtEqPrefixNumericRhs) {
  const ParsedCriterion c = parse_criterion(Value::text("<=3.14"));
  EXPECT_EQ(c.op, CriteriaOp::LtEq);
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_DOUBLE_EQ(c.rhs_number, 3.14);
}

TEST(CriteriaParse, TextWithNotEqPrefixNumericRhs) {
  const ParsedCriterion c = parse_criterion(Value::text("<>42"));
  EXPECT_EQ(c.op, CriteriaOp::NotEq);
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_DOUBLE_EQ(c.rhs_number, 42.0);
}

TEST(CriteriaParse, TextWithNotEqAndEmptyIsTextNotEq) {
  // "<>" with nothing after is "non-blank", not "not equal to zero".
  const ParsedCriterion c = parse_criterion(Value::text("<>"));
  EXPECT_EQ(c.op, CriteriaOp::NotEq);
  EXPECT_FALSE(c.rhs_is_number);
  EXPECT_EQ(c.rhs_text, std::string_view{});
}

TEST(CriteriaParse, TextWithGtAndTextRhsIsTextOrdering) {
  const ParsedCriterion c = parse_criterion(Value::text(">banana"));
  EXPECT_EQ(c.op, CriteriaOp::Gt);
  EXPECT_FALSE(c.rhs_is_number);
  EXPECT_EQ(c.rhs_text, "banana");
}

TEST(CriteriaParse, WildcardStarSetsFlag) {
  const ParsedCriterion c = parse_criterion(Value::text("A*"));
  EXPECT_TRUE(c.has_wildcard);
}

TEST(CriteriaParse, WildcardQuestionSetsFlag) {
  const ParsedCriterion c = parse_criterion(Value::text("?ed"));
  EXPECT_TRUE(c.has_wildcard);
}

TEST(CriteriaParse, EscapedStarIsLiteral) {
  const ParsedCriterion c = parse_criterion(Value::text("~*"));
  EXPECT_FALSE(c.has_wildcard);
}

TEST(CriteriaParse, EscapedQuestionIsLiteral) {
  const ParsedCriterion c = parse_criterion(Value::text("~?"));
  EXPECT_FALSE(c.has_wildcard);
}

TEST(CriteriaParse, MixedEscapeAndRealWildcardSetsFlag) {
  // "~?*" contains an escaped '?' followed by a real '*', so the flag
  // should still be set.
  const ParsedCriterion c = parse_criterion(Value::text("~?*"));
  EXPECT_TRUE(c.has_wildcard);
}

TEST(CriteriaParse, WildcardInOrderingOpIgnoredForFlag) {
  // Ordering ops treat wildcards as literals — has_wildcard stays false.
  const ParsedCriterion c = parse_criterion(Value::text(">A*"));
  EXPECT_EQ(c.op, CriteriaOp::Gt);
  EXPECT_FALSE(c.rhs_is_number);
  EXPECT_EQ(c.rhs_text, "A*");
  EXPECT_FALSE(c.has_wildcard);
}

TEST(CriteriaParse, IsoDateStringCoercesToNumericSerial) {
  // "2024-07-01" is Excel serial 45474. The criterion should be numeric so
  // COUNTIFS / SUMIFS / D-functions compare against date-serial cells.
  const ParsedCriterion c = parse_criterion(Value::text(">=2024-07-01"));
  EXPECT_EQ(c.op, CriteriaOp::GtEq);
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_DOUBLE_EQ(c.rhs_number, 45474.0);
}

TEST(CriteriaParse, SlashDateStringCoercesToNumericSerial) {
  const ParsedCriterion c = parse_criterion(Value::text("<2024/7/1"));
  EXPECT_EQ(c.op, CriteriaOp::Lt);
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_DOUBLE_EQ(c.rhs_number, 45474.0);
}

TEST(CriteriaParse, DateLikeButWildcardStaysText) {
  // A wildcard in the RHS is a strong signal that the user wants byte-wise
  // text matching; skip the date-serial fallback even when the prefix parses
  // as a date.
  const ParsedCriterion c = parse_criterion(Value::text("2024-*"));
  EXPECT_EQ(c.op, CriteriaOp::Eq);
  EXPECT_FALSE(c.rhs_is_number);
  EXPECT_EQ(c.rhs_text, "2024-*");
  EXPECT_TRUE(c.has_wildcard);
}

TEST(CriteriaParse, UnparseableTextStaysText) {
  // "banana" is neither numeric nor a date; stay on the text path.
  const ParsedCriterion c = parse_criterion(Value::text(">banana"));
  EXPECT_EQ(c.op, CriteriaOp::Gt);
  EXPECT_FALSE(c.rhs_is_number);
  EXPECT_EQ(c.rhs_text, "banana");
}

// ---------------------------------------------------------------------------
// matches_criterion: numeric
// ---------------------------------------------------------------------------

TEST(CriteriaMatchNumeric, EqMatchesExactNumber) {
  const ParsedCriterion c = parse_criterion(Value::number(5.0));
  EXPECT_TRUE(matches_criterion(Value::number(5.0), c));
  EXPECT_FALSE(matches_criterion(Value::number(5.000001), c));
  EXPECT_FALSE(matches_criterion(Value::number(4.0), c));
}

TEST(CriteriaMatchNumeric, GtPrefix) {
  const ParsedCriterion c = parse_criterion(Value::text(">5"));
  EXPECT_TRUE(matches_criterion(Value::number(6.0), c));
  EXPECT_FALSE(matches_criterion(Value::number(5.0), c));
  EXPECT_FALSE(matches_criterion(Value::number(-1.0), c));
}

TEST(CriteriaMatchNumeric, LtEqPrefix) {
  const ParsedCriterion c = parse_criterion(Value::text("<=10"));
  EXPECT_TRUE(matches_criterion(Value::number(10.0), c));
  EXPECT_TRUE(matches_criterion(Value::number(-9999.0), c));
  EXPECT_FALSE(matches_criterion(Value::number(10.000001), c));
}

TEST(CriteriaMatchNumeric, NotEqZero) {
  const ParsedCriterion c = parse_criterion(Value::text("<>0"));
  EXPECT_TRUE(matches_criterion(Value::number(1.0), c));
  EXPECT_TRUE(matches_criterion(Value::number(-0.5), c));
  EXPECT_FALSE(matches_criterion(Value::number(0.0), c));
}

TEST(CriteriaMatchNumeric, NonNumericTextDoesNotMatchNumericCriterion) {
  const ParsedCriterion c = parse_criterion(Value::text(">5"));
  EXPECT_FALSE(matches_criterion(Value::text("banana"), c));
}

TEST(CriteriaMatchNumeric, NumericTextDoesNotMatchNumericCriterion) {
  // Mac Excel 365 is type-strict: a numeric criterion (">5") matches only
  // Number cells. A Text cell whose string happens to parse as a number
  // (e.g. "10") does NOT match. This matches COUNTIF's oracle behaviour.
  const ParsedCriterion c = parse_criterion(Value::text(">5"));
  EXPECT_FALSE(matches_criterion(Value::text("10"), c));
  EXPECT_TRUE(matches_criterion(Value::number(10.0), c));
}

// ---------------------------------------------------------------------------
// matches_criterion: text equality (case-insensitive)
// ---------------------------------------------------------------------------

TEST(CriteriaMatchText, EqCaseInsensitive) {
  const ParsedCriterion c = parse_criterion(Value::text("HELLO"));
  EXPECT_TRUE(matches_criterion(Value::text("hello"), c));
  EXPECT_TRUE(matches_criterion(Value::text("Hello"), c));
  EXPECT_FALSE(matches_criterion(Value::text("world"), c));
}

TEST(CriteriaMatchText, NotEqCaseInsensitive) {
  const ParsedCriterion c = parse_criterion(Value::text("<>apple"));
  EXPECT_FALSE(matches_criterion(Value::text("APPLE"), c));
  EXPECT_TRUE(matches_criterion(Value::text("banana"), c));
}

TEST(CriteriaMatchText, OrderingTextCaseInsensitive) {
  const ParsedCriterion c = parse_criterion(Value::text(">banana"));
  EXPECT_TRUE(matches_criterion(Value::text("cherry"), c));
  EXPECT_TRUE(matches_criterion(Value::text("CHERRY"), c));
  EXPECT_FALSE(matches_criterion(Value::text("apple"), c));
  EXPECT_FALSE(matches_criterion(Value::text("banana"), c));
}

// ---------------------------------------------------------------------------
// matches_criterion: wildcards
// ---------------------------------------------------------------------------

TEST(CriteriaMatchWildcard, StarPrefix) {
  const ParsedCriterion c = parse_criterion(Value::text("*foo"));
  EXPECT_TRUE(matches_criterion(Value::text("foo"), c));
  EXPECT_TRUE(matches_criterion(Value::text("barfoo"), c));
  EXPECT_FALSE(matches_criterion(Value::text("foobar"), c));
}

TEST(CriteriaMatchWildcard, StarSuffix) {
  const ParsedCriterion c = parse_criterion(Value::text("foo*"));
  EXPECT_TRUE(matches_criterion(Value::text("foo"), c));
  EXPECT_TRUE(matches_criterion(Value::text("foobar"), c));
  EXPECT_FALSE(matches_criterion(Value::text("barfoo"), c));
}

TEST(CriteriaMatchWildcard, StarMiddle) {
  const ParsedCriterion c = parse_criterion(Value::text("a*z"));
  EXPECT_TRUE(matches_criterion(Value::text("az"), c));
  EXPECT_TRUE(matches_criterion(Value::text("amz"), c));
  EXPECT_TRUE(matches_criterion(Value::text("abcz"), c));
  EXPECT_FALSE(matches_criterion(Value::text("zza"), c));
}

TEST(CriteriaMatchWildcard, QuestionMatchesOneByte) {
  const ParsedCriterion c = parse_criterion(Value::text("?ed"));
  EXPECT_TRUE(matches_criterion(Value::text("bed"), c));
  EXPECT_TRUE(matches_criterion(Value::text("Ted"), c));
  EXPECT_FALSE(matches_criterion(Value::text("ed"), c));
  EXPECT_FALSE(matches_criterion(Value::text("breed"), c));
}

TEST(CriteriaMatchWildcard, EscapedStarIsLiteral) {
  // "~*foo" means: match a literal '*' followed by "foo".
  const ParsedCriterion c = parse_criterion(Value::text("~*foo"));
  EXPECT_TRUE(matches_criterion(Value::text("*foo"), c));
  EXPECT_FALSE(matches_criterion(Value::text("barfoo"), c));
}

TEST(CriteriaMatchWildcard, EscapedQuestionIsLiteral) {
  const ParsedCriterion c = parse_criterion(Value::text("foo~?"));
  EXPECT_TRUE(matches_criterion(Value::text("foo?"), c));
  EXPECT_FALSE(matches_criterion(Value::text("fooX"), c));
}

TEST(CriteriaMatchWildcard, EscapedTildeIsLiteral) {
  const ParsedCriterion c = parse_criterion(Value::text("~~"));
  EXPECT_TRUE(matches_criterion(Value::text("~"), c));
  EXPECT_FALSE(matches_criterion(Value::text("~~"), c));
}

TEST(CriteriaMatchWildcard, NotEqWithWildcard) {
  const ParsedCriterion c = parse_criterion(Value::text("<>*foo"));
  EXPECT_FALSE(matches_criterion(Value::text("foo"), c));
  EXPECT_FALSE(matches_criterion(Value::text("barfoo"), c));
  EXPECT_TRUE(matches_criterion(Value::text("bar"), c));
}

TEST(CriteriaMatchWildcard, StarAloneMatchesEverythingNonBlank) {
  const ParsedCriterion c = parse_criterion(Value::text("*"));
  EXPECT_TRUE(matches_criterion(Value::text(""), c));
  EXPECT_TRUE(matches_criterion(Value::text("hello"), c));
  // Blank cell doesn't hit the text path unless rhs_text is empty.
  EXPECT_FALSE(matches_criterion(Value::blank(), c));
}

// ---------------------------------------------------------------------------
// matches_criterion: blank-cell rules
// ---------------------------------------------------------------------------

TEST(CriteriaMatchBlank, EqEmptyStringMatchesBlankCell) {
  const ParsedCriterion c = parse_criterion(Value::text(""));
  EXPECT_TRUE(matches_criterion(Value::blank(), c));
  EXPECT_TRUE(matches_criterion(Value::text(""), c));
  EXPECT_FALSE(matches_criterion(Value::text("x"), c));
}

TEST(CriteriaMatchBlank, NotEqEmptyStringRejectsBlankCell) {
  const ParsedCriterion c = parse_criterion(Value::text("<>"));
  EXPECT_FALSE(matches_criterion(Value::blank(), c));
  EXPECT_TRUE(matches_criterion(Value::text("x"), c));
}

TEST(CriteriaMatchBlank, NotEqNonEmptyMatchesBlankCell) {
  // Blank is "not equal to 5".
  const ParsedCriterion c = parse_criterion(Value::text("<>5"));
  EXPECT_TRUE(matches_criterion(Value::blank(), c));
}

TEST(CriteriaMatchBlank, OrderingAgainstBlankFails) {
  const ParsedCriterion c = parse_criterion(Value::text(">0"));
  EXPECT_FALSE(matches_criterion(Value::blank(), c));
}

// ---------------------------------------------------------------------------
// matches_criterion: error cells -- error is unequal to any concrete value
// so NotEq matches, every other op fails.
// ---------------------------------------------------------------------------

TEST(CriteriaMatchError, ErrorCellMatchesNotEqOnly) {
  const ParsedCriterion c_eq = parse_criterion(Value::text(""));
  const ParsedCriterion c_gt = parse_criterion(Value::text(">0"));
  const ParsedCriterion c_neq = parse_criterion(Value::text("<>0"));
  EXPECT_FALSE(matches_criterion(Value::error(ErrorCode::Div0), c_eq));
  EXPECT_FALSE(matches_criterion(Value::error(ErrorCode::Div0), c_gt));
  EXPECT_TRUE(matches_criterion(Value::error(ErrorCode::Div0), c_neq));
  EXPECT_FALSE(matches_criterion(Value::error(ErrorCode::NA), c_eq));
  EXPECT_TRUE(matches_criterion(Value::error(ErrorCode::NA), c_neq));
}

// ---------------------------------------------------------------------------
// matches_criterion: bool folding
// ---------------------------------------------------------------------------

TEST(CriteriaMatchBool, BoolCellDoesNotMatchNumericCriterion) {
  // Type-strict rule: a Number criterion matches only Number cells. A
  // Bool cell (TRUE / FALSE) does not match a Number criterion of 1.0,
  // even though both coerce to 1 numerically.
  const ParsedCriterion c = parse_criterion(Value::number(1.0));
  EXPECT_FALSE(matches_criterion(Value::boolean(true), c));
  EXPECT_FALSE(matches_criterion(Value::boolean(false), c));
  EXPECT_TRUE(matches_criterion(Value::number(1.0), c));
}

TEST(CriteriaMatchNumeric, NotEqNumericCriterionBroadensToOtherTypes) {
  // "<>23" matches anything that isn't a Number equal to 23 — including
  // Text, Bool, and Number cells with a different value. This broadening
  // is specific to NotEq; ordering ops stay type-strict.
  const ParsedCriterion c = parse_criterion(Value::text("<>23"));
  EXPECT_TRUE(matches_criterion(Value::number(1.0), c));
  EXPECT_FALSE(matches_criterion(Value::number(23.0), c));
  EXPECT_TRUE(matches_criterion(Value::text("23"), c));       // type differs
  EXPECT_TRUE(matches_criterion(Value::text("hey"), c));
  EXPECT_TRUE(matches_criterion(Value::boolean(true), c));     // type differs
  EXPECT_TRUE(matches_criterion(Value::boolean(false), c));
}

TEST(CriteriaMatchNumeric, OrderingNumericCriterionStaysTypeStrict) {
  // ">0" still matches only Number cells; Text "5" and Bool TRUE do not.
  const ParsedCriterion c = parse_criterion(Value::text(">0"));
  EXPECT_TRUE(matches_criterion(Value::number(1.0), c));
  EXPECT_FALSE(matches_criterion(Value::text("5"), c));
  EXPECT_FALSE(matches_criterion(Value::boolean(true), c));
}

TEST(CriteriaMatchBool, BoolCellAgainstBoolCriterion) {
  // A Bool criterion matches only Bool cells with the same value.
  const ParsedCriterion c = parse_criterion(Value::boolean(true));
  EXPECT_TRUE(matches_criterion(Value::boolean(true), c));
  EXPECT_FALSE(matches_criterion(Value::boolean(false), c));
  // And does not match a Number cell with value 1.0 (type mismatch).
  EXPECT_FALSE(matches_criterion(Value::number(1.0), c));
}

}  // namespace
}  // namespace eval
}  // namespace formulon
