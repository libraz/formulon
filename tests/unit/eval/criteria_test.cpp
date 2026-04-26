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

TEST(CriteriaMatchWildcard, StarAloneMatchesNonEmptyTextOnly) {
  // Mac Excel 365 treats `=COUNTIF(range, "*")` as "match any non-empty
  // text cell". An empty-text cell `Value::text("")` has no characters
  // and is excluded; Number / Bool / Blank cells are likewise excluded
  // (the text-only-cells rule already in place). Verified via
  // tests/oracle/cases/countif.yaml case
  // `countif_wildcard_star_alone_text_only`.
  const ParsedCriterion c = parse_criterion(Value::text("*"));
  EXPECT_TRUE(matches_criterion(Value::text("hello"), c));
  EXPECT_FALSE(matches_criterion(Value::text(""), c));
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

TEST(CriteriaMatchBool, BoolTrueCriterionMatchesTextTrueCaseInsensitive) {
  // Mac Excel 365 (ja-JP) extends Bool Eq criteria to also match Text
  // cells whose display form is "TRUE" / "FALSE" case-insensitively.
  // Verified via tests/oracle/golden/countif_edges.golden.json cases
  // countif_bool_true_over_mixed and countif_cell_criterion_bool_true_with_text.
  const ParsedCriterion c = parse_criterion(Value::boolean(true));
  EXPECT_TRUE(matches_criterion(Value::text("TRUE"), c));
  EXPECT_TRUE(matches_criterion(Value::text("true"), c));
  EXPECT_TRUE(matches_criterion(Value::text("True"), c));
}

TEST(CriteriaMatchBool, BoolTrueCriterionRejectsUnrelatedCells) {
  // The bool-to-text broadening is narrow: only literal "TRUE" /
  // "FALSE" Text cells match. Text "FALSE", Text "1", and Number 1 do
  // NOT match a TRUE bool criterion.
  const ParsedCriterion c = parse_criterion(Value::boolean(true));
  EXPECT_FALSE(matches_criterion(Value::text("FALSE"), c));
  EXPECT_FALSE(matches_criterion(Value::text("1"), c));
  EXPECT_FALSE(matches_criterion(Value::number(1.0), c));
}

TEST(CriteriaMatchBool, BoolFalseCriterionMatchesTextFalseCaseInsensitive) {
  const ParsedCriterion c = parse_criterion(Value::boolean(false));
  EXPECT_TRUE(matches_criterion(Value::text("FALSE"), c));
  EXPECT_TRUE(matches_criterion(Value::text("false"), c));
  EXPECT_TRUE(matches_criterion(Value::text("False"), c));
  // And does NOT match Text "TRUE" or Number 0.
  EXPECT_FALSE(matches_criterion(Value::text("TRUE"), c));
  EXPECT_FALSE(matches_criterion(Value::number(0.0), c));
}

TEST(CriteriaMatchBool, BoolCriterionNotEqUnchangedForText) {
  // NotEq + cross-kind returns `true` unconditionally (see
  // NotEqNumericCriterionBroadensToOtherTypes). The new Eq-only branch
  // MUST NOT run for NotEq, so a bool-from-cell NotEq against a Text
  // "TRUE" cell still returns true (cell type differs from Bool).
  //
  // `parse_criterion` always produces Eq for a Bool literal, so we build
  // the ParsedCriterion by hand to reach the NotEq + rhs_from_bool path
  // that `*IF` callers can synthesise via argument plumbing.
  ParsedCriterion c;
  c.op = CriteriaOp::NotEq;
  c.rhs_is_number = true;
  c.rhs_from_bool = true;
  c.rhs_number = 1.0;
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
// wildcard_find_dbcs: byte-mode `?` for SEARCHB. `?` matches only codepoints
// whose ja-JP DBCS cost is 1 (ASCII, half-width katakana); kanji/hiragana/
// full-width punctuation/emoji refuse to match. `*` and `~?`/`~*` retain
// their normal semantics.
// ---------------------------------------------------------------------------

TEST(WildcardFindDbcs, QuestionSkipsKanjiToAsciiOffset) {
  // "漢" is a 3-byte UTF-8 codepoint with DBCS cost 2; `?` refuses to match
  // it. The next codepoint is 'A' at byte offset 3.
  EXPECT_EQ(wildcard_find_dbcs("?", "漢ABC"), std::size_t{3});
}

TEST(WildcardFindDbcs, QuestionMatchesAsciiAtOffsetZero) {
  EXPECT_EQ(wildcard_find_dbcs("?", "abc"), std::size_t{0});
}

TEST(WildcardFindDbcs, QuestionFindsNothingInPureKanji) {
  // No SBCS codepoint anywhere; `?` matches nothing.
  EXPECT_EQ(wildcard_find_dbcs("?", "漢"), std::string_view::npos);
}

TEST(WildcardFindDbcs, QuestionWithLiteralTailAfterKanji) {
  // At start=0, "?" against 漢 fails. At start=3 the suffix is "abc" but
  // pattern "?abc" needs 4 chars and only "abc" remains — no match.
  EXPECT_EQ(wildcard_find_dbcs("?abc", "漢abc"), std::string_view::npos);
}

TEST(WildcardFindDbcs, QuestionInsideAsciiHaystack) {
  EXPECT_EQ(wildcard_find_dbcs("a?c", "abc"), std::size_t{0});
}

TEST(WildcardFindDbcs, StarSwallowsKanjiThenQuestionMatchesAscii) {
  // `*` consumes the leading 漢 (2 DBCS bytes / 3 UTF-8 bytes) and `?`
  // then matches 'A'.
  EXPECT_EQ(wildcard_find_dbcs("*?", "漢A"), std::size_t{0});
}

TEST(WildcardFindDbcs, EscapedQuestionIsLiteralRegardlessOfMode) {
  // `~?` is a literal '?' under both wildcard modes; the byte-mode flag
  // only affects unescaped `?`.
  EXPECT_EQ(wildcard_find_dbcs("~?", "x?y"), std::size_t{1});
}

TEST(WildcardFindDbcs, QuestionMatchesHalfWidthKatakana) {
  // ｱ (U+FF71) is half-width katakana, classified as DBCS=1 in ja-JP, so
  // byte-mode `?` matches it at offset 0.
  EXPECT_EQ(wildcard_find_dbcs("?", "ｱABC"), std::size_t{0});
}

}  // namespace
}  // namespace eval
}  // namespace formulon
