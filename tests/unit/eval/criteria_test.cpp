//
// Unit tests for `parse_criterion` and `matches_criterion`. These helpers
// power `COUNTIF`, `SUMIF`, and `AVERAGEIF`; the tests here exercise them
// in isolation so that criterion-level edge cases (operator prefixes,
// numeric-text probing, wildcard handling, blank-cell rules) are covered
// without going through range resolution.

#include "eval/criteria.h"

#include <fstream>
#include <string>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "utils/text_ops.h"
#include "value.h"

#ifndef FORMULON_COLLATION_DATA_DIR
#error "FORMULON_COLLATION_DATA_DIR must name tools/collation/data"
#endif

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
  EXPECT_TRUE(c.rhs_from_bool);
  EXPECT_DOUBLE_EQ(c.rhs_number, 1.0);
}

TEST(CriteriaParse, BareBoolFalseBecomesEqZero) {
  const ParsedCriterion c = parse_criterion(Value::boolean(false));
  EXPECT_TRUE(c.rhs_is_number);
  EXPECT_TRUE(c.rhs_from_bool);
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

TEST(CriteriaParse, TextBooleanLiteralsBecomeBoolCriteria) {
  struct TestCase {
    const char* raw;
    CriteriaOp op;
    double number;
  };
  const TestCase cases[] = {
      {"TRUE", CriteriaOp::Eq, 1.0},      {"true", CriteriaOp::Eq, 1.0},       {"=TRUE", CriteriaOp::Eq, 1.0},
      {"FALSE", CriteriaOp::Eq, 0.0},     {"false", CriteriaOp::Eq, 0.0},      {"=FALSE", CriteriaOp::Eq, 0.0},
      {"<>TRUE", CriteriaOp::NotEq, 1.0}, {"<>false", CriteriaOp::NotEq, 0.0},
  };

  for (const TestCase& test : cases) {
    const ParsedCriterion c = parse_criterion(Value::text(test.raw));
    EXPECT_EQ(c.op, test.op) << test.raw;
    EXPECT_TRUE(c.rhs_is_number) << test.raw;
    EXPECT_TRUE(c.rhs_from_bool) << test.raw;
    EXPECT_DOUBLE_EQ(c.rhs_number, test.number) << test.raw;
    EXPECT_TRUE(c.rhs_text.empty()) << test.raw;
  }
}

TEST(CriteriaParse, BooleanLookingWildcardAndOrderingCriteriaStayText) {
  const ParsedCriterion wildcard = parse_criterion(Value::text("TRUE*"));
  EXPECT_EQ(wildcard.op, CriteriaOp::Eq);
  EXPECT_FALSE(wildcard.rhs_is_number);
  EXPECT_FALSE(wildcard.rhs_from_bool);
  EXPECT_TRUE(wildcard.has_wildcard);
  EXPECT_EQ(wildcard.rhs_text, "TRUE*");

  const ParsedCriterion ordering = parse_criterion(Value::text(">TRUE"));
  EXPECT_EQ(ordering.op, CriteriaOp::Gt);
  EXPECT_FALSE(ordering.rhs_is_number);
  EXPECT_FALSE(ordering.rhs_from_bool);
  EXPECT_FALSE(ordering.has_wildcard);
  EXPECT_EQ(ordering.rhs_text, "TRUE");
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

TEST(CriteriaMatchNumeric, ExplicitEqStillCoercesNumericText) {
  const ParsedCriterion c = parse_criterion(Value::text("=23"));
  ASSERT_TRUE(c.rhs_is_number);
  EXPECT_FALSE(c.rhs_from_bool);
  EXPECT_TRUE(matches_criterion(Value::number(23.0), c));
  EXPECT_TRUE(matches_criterion(Value::text("23"), c));
  EXPECT_FALSE(matches_criterion(Value::text("TRUE"), c));
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

TEST(CriteriaMatchWildcard, MalformedTildeMatchesNothing) {
  // Mac Excel 365 treats `~~` (and any other malformed escape) as an
  // un-matchable pattern: `=COUNTIF(range, "~~")` returns 0 even when a
  // cell holds a literal tilde. Excel's `~` only escapes `*` and `?`;
  // every other use renders the entire criterion invalid. Verified via
  // tests/oracle/cases/countif.yaml case `countif_tilde_escapes_tilde`.
  const ParsedCriterion c = parse_criterion(Value::text("~~"));
  EXPECT_TRUE(c.rhs_invalid_wildcard);
  EXPECT_FALSE(matches_criterion(Value::text("~"), c));
  EXPECT_FALSE(matches_criterion(Value::text("~~"), c));
  EXPECT_FALSE(matches_criterion(Value::text("anything"), c));
  // The mirror NotEq matches every non-blank, non-error cell.
  const ParsedCriterion neq = parse_criterion(Value::text("<>~~"));
  EXPECT_TRUE(neq.rhs_invalid_wildcard);
  EXPECT_TRUE(matches_criterion(Value::text("~"), neq));
  EXPECT_TRUE(matches_criterion(Value::text("anything"), neq));
}

TEST(CriteriaMatchWildcard, TildeFollowedByNonMetaIsInvalid) {
  // `~a` is malformed under the strict "tilde escapes only `*` or `?`"
  // rule documented by Excel. Mac has only been probed for `~~` directly,
  // but the consistent behaviour is to reject the entire pattern. Locked
  // in here so future drift is caught by unit tests rather than the slower
  // oracle pipeline. Note: SEARCH / FIND / MATCH still honour the lenient
  // `wildcard_match` decoding (`~a` -> literal `a`) — this strictness
  // applies only to the criteria layer used by COUNTIF / SUMIF / etc.
  const ParsedCriterion c = parse_criterion(Value::text("~a"));
  EXPECT_TRUE(c.rhs_invalid_wildcard);
  EXPECT_FALSE(matches_criterion(Value::text("a"), c));
  EXPECT_FALSE(matches_criterion(Value::text("~a"), c));
  // A trailing bare `~` is likewise malformed.
  const ParsedCriterion trailing = parse_criterion(Value::text("foo~"));
  EXPECT_TRUE(trailing.rhs_invalid_wildcard);
  EXPECT_FALSE(matches_criterion(Value::text("foo"), trailing));
  EXPECT_FALSE(matches_criterion(Value::text("foo~"), trailing));
}

TEST(CriteriaMatchWildcard, NotEqWithWildcard) {
  const ParsedCriterion c = parse_criterion(Value::text("<>*foo"));
  EXPECT_FALSE(matches_criterion(Value::text("foo"), c));
  EXPECT_FALSE(matches_criterion(Value::text("barfoo"), c));
  EXPECT_TRUE(matches_criterion(Value::text("bar"), c));
}

TEST(CriteriaMatchWildcard, StarAloneMatchesTextCellsOnly) {
  // Mac Excel 365 treats `=COUNTIF(range, "*")` as "match any text cell".
  // A zero-length string is one, so it matches even though it has no
  // characters; Number / Bool / Blank cells are excluded by the
  // text-only-cells rule. Verified via tests/oracle/cases/countif.yaml
  // case `countif_wildcard_star_alone_text_only`.
  const ParsedCriterion c = parse_criterion(Value::text("*"));
  EXPECT_TRUE(matches_criterion(Value::text("hello"), c));
  EXPECT_TRUE(matches_criterion(Value::text(""), c));
  EXPECT_FALSE(matches_criterion(Value::number(10.0), c));
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

TEST(CriteriaMatchBlank, BareEqualsMatchesOnlyGenuinelyBlankCells) {
  // `"="` and `""` both parse to Eq with an empty RHS, but Excel keeps
  // them apart: `""` above accepts a zero-length string, `"="` asks
  // strictly whether the cell is empty and a zero-length string is not.
  const ParsedCriterion c = parse_criterion(Value::text("="));
  EXPECT_TRUE(matches_criterion(Value::blank(), c));
  EXPECT_FALSE(matches_criterion(Value::text(""), c));
  EXPECT_FALSE(matches_criterion(Value::text("x"), c));
  EXPECT_FALSE(matches_criterion(Value::number(0.0), c));
}

TEST(CriteriaMatchBlank, NotEqEmptyStringRejectsBlankCell) {
  // The mirror of the bare `"="` probe: every non-blank cell, and a
  // zero-length string counts as populated.
  const ParsedCriterion c = parse_criterion(Value::text("<>"));
  EXPECT_FALSE(matches_criterion(Value::blank(), c));
  EXPECT_TRUE(matches_criterion(Value::text(""), c));
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
  EXPECT_TRUE(matches_criterion(Value::text("23"), c));  // type differs
  EXPECT_TRUE(matches_criterion(Value::text("hey"), c));
  EXPECT_TRUE(matches_criterion(Value::boolean(true), c));  // type differs
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

TEST(CriteriaMatchBool, BoolTrueCriterionDoesNotMatchTextTrueCaseInsensitive) {
  // Excel's Bool Eq criteria are type-strict: Text cells whose display form
  // is "TRUE" do not match a Bool TRUE criterion, regardless of case.
  const ParsedCriterion c = parse_criterion(Value::boolean(true));
  EXPECT_FALSE(matches_criterion(Value::text("TRUE"), c));
  EXPECT_FALSE(matches_criterion(Value::text("true"), c));
  EXPECT_FALSE(matches_criterion(Value::text("True"), c));
}

TEST(CriteriaMatchBool, BoolTrueCriterionRejectsUnrelatedCells) {
  // A TRUE Bool criterion does not match a FALSE or numeric/text value.
  const ParsedCriterion c = parse_criterion(Value::boolean(true));
  EXPECT_FALSE(matches_criterion(Value::text("FALSE"), c));
  EXPECT_FALSE(matches_criterion(Value::text("1"), c));
  EXPECT_FALSE(matches_criterion(Value::number(1.0), c));
}

TEST(CriteriaMatchBool, BoolFalseCriterionDoesNotMatchTextFalseCaseInsensitive) {
  const ParsedCriterion c = parse_criterion(Value::boolean(false));
  EXPECT_FALSE(matches_criterion(Value::text("FALSE"), c));
  EXPECT_FALSE(matches_criterion(Value::text("false"), c));
  EXPECT_FALSE(matches_criterion(Value::text("False"), c));
  // And does NOT match Text "TRUE" or Number 0.
  EXPECT_FALSE(matches_criterion(Value::text("TRUE"), c));
  EXPECT_FALSE(matches_criterion(Value::number(0.0), c));
}

TEST(CriteriaMatchBool, ParsedBoolEqCriteriaAreTypeStrict) {
  struct TestCase {
    const char* raw;
    bool expected;
  };
  const TestCase cases[] = {
      {"TRUE", true}, {"true", true}, {"=TRUE", true}, {"FALSE", false}, {"false", false}, {"=FALSE", false},
  };

  for (const TestCase& test : cases) {
    const ParsedCriterion c = parse_criterion(Value::text(test.raw));
    ASSERT_EQ(c.op, CriteriaOp::Eq) << test.raw;
    ASSERT_TRUE(c.rhs_is_number) << test.raw;
    ASSERT_TRUE(c.rhs_from_bool) << test.raw;
    EXPECT_TRUE(matches_criterion(Value::boolean(test.expected), c)) << test.raw;
    EXPECT_FALSE(matches_criterion(Value::boolean(!test.expected), c)) << test.raw;
    EXPECT_FALSE(matches_criterion(Value::text(test.expected ? "TRUE" : "FALSE"), c)) << test.raw;
    EXPECT_FALSE(matches_criterion(Value::number(test.expected ? 1.0 : 0.0), c)) << test.raw;
  }
}

TEST(CriteriaMatchBool, ParsedBoolNotEqCriteriaBroadenAcrossTypes) {
  struct TestCase {
    const char* raw;
    bool excluded;
  };
  const TestCase cases[] = {{"<>TRUE", true}, {"<>false", false}};

  for (const TestCase& test : cases) {
    const ParsedCriterion c = parse_criterion(Value::text(test.raw));
    ASSERT_EQ(c.op, CriteriaOp::NotEq) << test.raw;
    ASSERT_TRUE(c.rhs_is_number) << test.raw;
    ASSERT_TRUE(c.rhs_from_bool) << test.raw;
    EXPECT_FALSE(matches_criterion(Value::boolean(test.excluded), c)) << test.raw;
    EXPECT_TRUE(matches_criterion(Value::boolean(!test.excluded), c)) << test.raw;
    EXPECT_TRUE(matches_criterion(Value::text(test.excluded ? "TRUE" : "FALSE"), c)) << test.raw;
    EXPECT_TRUE(matches_criterion(Value::number(test.excluded ? 1.0 : 0.0), c)) << test.raw;
  }
}

TEST(CriteriaMatchBool, BooleanLookingWildcardAndOrderingCriteriaRemainText) {
  const ParsedCriterion wildcard = parse_criterion(Value::text("TRUE*"));
  EXPECT_TRUE(matches_criterion(Value::text("TRUE"), wildcard));
  EXPECT_TRUE(matches_criterion(Value::text("TRUE-ish"), wildcard));
  EXPECT_FALSE(matches_criterion(Value::boolean(true), wildcard));

  const ParsedCriterion ordering = parse_criterion(Value::text(">TRUE"));
  EXPECT_TRUE(matches_criterion(Value::text("ZZZ"), ordering));
  EXPECT_FALSE(matches_criterion(Value::text("TRUE"), ordering));
  EXPECT_FALSE(matches_criterion(Value::boolean(true), ordering));
  EXPECT_FALSE(matches_criterion(Value::number(2.0), ordering));
}

TEST(CriteriaMatchBool, BoolCriterionNotEqUnchangedForText) {
  // NotEq + cross-kind returns `true` unconditionally (see
  // NotEqNumericCriterionBroadensToOtherTypes). A parsed bool NotEq
  // criterion therefore excludes only the same Bool value; text and
  // Number cells remain in the result even when their display form/value
  // is TRUE/1.
  const ParsedCriterion c = parse_criterion(Value::text("<>TRUE"));
  ASSERT_TRUE(c.rhs_from_bool);
  EXPECT_TRUE(matches_criterion(Value::text("TRUE"), c));
  EXPECT_TRUE(matches_criterion(Value::text("FALSE"), c));
  EXPECT_TRUE(matches_criterion(Value::number(1.0), c));
  EXPECT_FALSE(matches_criterion(Value::boolean(true), c));
  EXPECT_TRUE(matches_criterion(Value::boolean(false), c));
}

// ---------------------------------------------------------------------------
// ParsedCriterion copy / move: rhs_text must rebind when aliasing rhs_storage.
// Without an explicit special-member set the default move would leave
// `rhs_text` pointing at the moved-from object's SSO buffer, which becomes
// empty after move. This silently breaks `COUNTIFS`/`SUMIFS` when the
// aggregator stores each parsed criterion in a `unique_ptr` via a prvalue
// move, yielding false ">" / "<" matches that propagate sibling errors.
// ---------------------------------------------------------------------------

TEST(ParsedCriterionMove, TextRhsSurvivesMoveAndCopy) {
  ParsedCriterion src = parse_criterion(Value::text(">c"));
  ASSERT_EQ(src.op, CriteriaOp::Gt);
  ASSERT_EQ(src.rhs_text, "c");
  ParsedCriterion moved(std::move(src));
  EXPECT_EQ(moved.op, CriteriaOp::Gt);
  EXPECT_EQ(moved.rhs_text, "c");
  ParsedCriterion assigned;
  assigned = parse_criterion(Value::text("<=5"));
  EXPECT_EQ(assigned.op, CriteriaOp::LtEq);
  EXPECT_TRUE(assigned.rhs_is_number);
  ParsedCriterion copied(moved);
  EXPECT_EQ(copied.rhs_text, "c");
}

// ---------------------------------------------------------------------------
// Wildcard criteria in Mac Excel 365 match only Text cells. A Number or
// Bool cell never satisfies a wildcard pattern, even when its display form
// would literally match. See tests/oracle/cases/countif_edges.yaml case
// `countif_wildcard_matches_text_cells_only`.
// ---------------------------------------------------------------------------

TEST(WildcardCriterion, DoesNotMatchNumberCells) {
  const ParsedCriterion c = parse_criterion(Value::text("12*"));
  ASSERT_TRUE(c.has_wildcard);
  EXPECT_TRUE(matches_criterion(Value::text("12 Oak Street"), c));
  EXPECT_FALSE(matches_criterion(Value::number(1213.0), c));
  EXPECT_FALSE(matches_criterion(Value::number(12.0), c));
  EXPECT_FALSE(matches_criterion(Value::boolean(true), c));
}

TEST(WildcardCriterion, NotEqPatternCountsNonTextCells) {
  const ParsedCriterion c = parse_criterion(Value::text("<>12*"));
  ASSERT_TRUE(c.has_wildcard);
  EXPECT_FALSE(matches_criterion(Value::text("12 Oak"), c));
  EXPECT_TRUE(matches_criterion(Value::text("hello"), c));
  EXPECT_TRUE(matches_criterion(Value::number(1213.0), c));
}

// ---------------------------------------------------------------------------
// Collation-backed text matching.
// ---------------------------------------------------------------------------

ExcelProfile mac_profile(ExcelLocale locale) {
  return ExcelProfile{ExcelHost::kMac365, locale};
}

ExcelProfile win_profile(ExcelLocale locale) {
  return ExcelProfile{ExcelHost::kWin365, locale};
}

// "U+0061 U+0301" -> UTF-8.
std::string decode_codepoints(const std::string& field) {
  std::string out;
  std::size_t i = 0;
  while (i < field.size()) {
    const std::size_t space = field.find(' ', i);
    const std::size_t end = space == std::string::npos ? field.size() : space;
    out += encode_utf8_codepoint(static_cast<std::uint32_t>(std::stoul(field.substr(i + 2, end - i - 2), nullptr, 16)));
    i = end + 1;
  }
  return out;
}

bool countif_hit(const std::string& criterion, const std::string& cell, ExcelProfile profile) {
  const ScopedEvalProfile scope(profile);
  return matches_criterion(Value::text(cell), parse_criterion(Value::text(criterion)), profile);
}

bool dcount_hit(const std::string& criterion, const std::string& cell, ExcelProfile profile) {
  const ScopedEvalProfile scope(profile);
  return matches_criterion(Value::text(cell), parse_criterion_dfunc(Value::text(criterion)), profile);
}

// Drives the countif / countifeq / sumif / dcount columns of the
// equality table. Row a is the criterion and b the cell; rows whose criterion
// is numeric or carries a comparator or wildcard take other criteria paths.
void check_eq_columns(const char* tag, ExcelLocale locale) {
  const ExcelProfile profile = mac_profile(locale);
  std::ifstream in(std::string(FORMULON_COLLATION_DATA_DIR) + "/fixtures/eq_" + tag + ".tsv");
  ASSERT_TRUE(in.good());
  std::string line;
  std::getline(in, line);
  std::vector<std::string> header;
  const auto split = [](const std::string& text) {
    std::vector<std::string> fields;
    std::size_t start = 0;
    for (std::size_t tab = text.find('\t'); tab != std::string::npos; tab = text.find('\t', start)) {
      fields.push_back(text.substr(start, tab - start));
      start = tab + 1;
    }
    fields.push_back(text.substr(start));
    return fields;
  };
  header = split(line);
  int checked = 0;
  while (std::getline(in, line)) {
    if (line.empty()) {
      continue;
    }
    const std::vector<std::string> row = split(line);
    ASSERT_EQ(row.size(), header.size());
    const std::string a = decode_codepoints(row[0]);
    const std::string b = decode_codepoints(row[1]);
    const ParsedCriterion probe = parse_criterion(Value::text(a));
    if (probe.rhs_is_number || probe.has_wildcard || probe.op != CriteriaOp::Eq || a.empty()) {
      continue;
    }
    for (std::size_t col = 2; col < header.size(); ++col) {
      const std::string& fn = header[col];
      bool got = false;
      std::string hit = "1";
      if (fn == "countif" || fn == "countifeq" || fn == "sumif") {
        got = countif_hit(a, b, profile);
        hit = fn == "sumif" ? "2" : "1";
      } else if (fn == "dcount") {
        got = dcount_hit(a, b, profile);
      } else {
        continue;
      }
      ++checked;
      EXPECT_EQ(got, row[col] == hit) << tag << " " << fn << " " << row[0] << " | " << row[1] << " = " << row[col];
    }
  }
  EXPECT_GT(checked, 100);
}

TEST(CriteriaCollation, MacJaJpEqTable) {
  check_eq_columns("ja_JP", ExcelLocale::kJaJP);
}

TEST(CriteriaCollation, MacEnUsEqTable) {
  check_eq_columns("en_US", ExcelLocale::kEnUS);
}

TEST(CriteriaCollation, MacZhCnEqTable) {
  check_eq_columns("zh_CN", ExcelLocale::kZhCN);
}

TEST(CriteriaCollation, MacThThEqTable) {
  check_eq_columns("th_TH", ExcelLocale::kThTH);
}

TEST(CriteriaCollation, DcountPrefixCellsEn) {
  const ExcelProfile en = mac_profile(ExcelLocale::kEnUS);
  EXPECT_FALSE(dcount_hit("\xC3\x9F", "ss", en));  // criterion ß, cell ss
  EXPECT_TRUE(dcount_hit("ss", "\xC3\x9F", en));
  EXPECT_FALSE(dcount_hit("\xC4\xB3", "ij", en));    // criterion ĳ, cell ij
  EXPECT_TRUE(dcount_hit("\xEF\xBC\x91", "1", en));  // full-width 1 against 1
}

TEST(CriteriaCollation, DcountPrefixCellsJa) {
  const ExcelProfile ja = mac_profile(ExcelLocale::kJaJP);
  EXPECT_TRUE(dcount_hit("\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC", ja));  // criterion ｶﾞ, cell ガ
}

TEST(CriteriaCollation, DcountPrefixCellsZh) {
  const ExcelProfile zh = mac_profile(ExcelLocale::kZhCN);
  EXPECT_FALSE(dcount_hit("\xC3\xA6", "ae", zh));  // criterion æ, cell ae
  EXPECT_TRUE(dcount_hit("ss", "\xC3\x9F", zh));
  EXPECT_TRUE(dcount_hit("\xEF\xBC\x91", "1", zh));
}

TEST(CriteriaCollation, JaIterationMarkCountif) {
  const ExcelProfile ja = mac_profile(ExcelLocale::kJaJP);
  EXPECT_TRUE(countif_hit("\xE3\x81\x8B\xE3\x82\x9D", "\xE3\x81\x8B\xE3\x81\x8B", ja));  // かゝ vs かか
  EXPECT_TRUE(countif_hit("\xE3\x81\x8B\xE3\x81\x8B", "\xE3\x81\x8B\xE3\x82\x9D", ja));
}

TEST(CriteriaCollation, MacOrderingUsesCollation) {
  const ExcelProfile en = mac_profile(ExcelLocale::kEnUS);
  const ScopedEvalProfile scope(en);
  const ParsedCriterion gt = parse_criterion(Value::text(">b"));
  EXPECT_TRUE(matches_criterion(Value::text("C"), gt, en));
  EXPECT_FALSE(matches_criterion(Value::text("B"), gt, en));
  EXPECT_FALSE(matches_criterion(Value::text("a"), gt, en));
  const ParsedCriterion ne = parse_criterion(Value::text("<>ss"));
  EXPECT_FALSE(matches_criterion(Value::text("\xC3\x9F"), ne, en));
}

TEST(CriteriaCollation, WindowsJaUsesCollation) {
  const ExcelProfile win = win_profile(ExcelLocale::kJaJP);
  // On Windows ja-JP: width and kana type separate, case and ligatures do not.
  EXPECT_TRUE(countif_hit("\xEF\xBD\xB1", "\xEF\xBD\xB1", win));
  EXPECT_FALSE(countif_hit("\xEF\xBD\xB1", "\xE3\x82\xA2", win));              // ｱ vs ア
  EXPECT_FALSE(countif_hit("\xE3\x81\x82", "\xE3\x82\xA2", win));              // あ vs ア
  EXPECT_FALSE(countif_hit("\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC", win));  // ｶﾞ vs ガ
  EXPECT_TRUE(countif_hit("ABC", "abc", win));
  EXPECT_TRUE(countif_hit("\xEF\xBC\x91\xEF\xBC\x92x", "12x", win));
  EXPECT_TRUE(countif_hit("\xC3\x9F", "ss", win));  // ß
  EXPECT_TRUE(countif_hit("\xC3\xA6", "ae", win));  // æ
  EXPECT_FALSE(dcount_hit("\xC3\x9F", "ss", win));
  const ScopedEvalProfile scope(win);
  const ParsedCriterion gt = parse_criterion(Value::text(">ab"));
  EXPECT_TRUE(matches_criterion(Value::text("a-b"), gt, win));  // hyphen weighs after every letter
}

}  // namespace
}  // namespace eval
}  // namespace formulon
