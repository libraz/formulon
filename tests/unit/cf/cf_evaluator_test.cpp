// Conditional-format evaluator tests: literal and text predicates.

#include "cf/cf_evaluator.h"

#include "cell.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "cf_evaluator_fixtures.h"
#include "gtest/gtest.h"
#include "utils/error.h"
#include "value.h"
namespace formulon::cf {
namespace {

using test::At;
using test::CFEvalHarness;
using test::MakeRange;
using test::MakeRule;
using test::PopulateLinearPopulation;
using test::RGB;
using test::SqrefContext;
using test::TwoStopMinMax;

TEST(CFEvaluator, ContainsBlanksMatchesBlank) {
  CFRule r = MakeRule(RuleType::ContainsBlanks);
  EXPECT_TRUE(match_rule(r, Value::blank()));
  EXPECT_FALSE(match_rule(r, Value::number(0.0)));
  EXPECT_FALSE(match_rule(r, Value::boolean(false)));
}

// Excel stores/evaluates the Blanks CF rule as LEN(TRIM(A1))=0, so a
// formula-produced empty string or whitespace-only text counts as blank,
// not only a true blank cell.
TEST(CFEvaluator, ContainsBlanksMatchesEmptyAndWhitespaceOnlyText) {
  CFRule r = MakeRule(RuleType::ContainsBlanks);
  EXPECT_TRUE(match_rule(r, Value::text("")));
  EXPECT_TRUE(match_rule(r, Value::text("   ")));
  EXPECT_TRUE(match_rule(r, Value::text("\xE3\x80\x80")));  // U+3000 ideographic space
  EXPECT_FALSE(match_rule(r, Value::text("x")));
  EXPECT_FALSE(match_rule(r, Value::text(" x ")));
}

TEST(CFEvaluator, NotContainsBlanksIsComplementOfContainsBlanks) {
  CFRule r = MakeRule(RuleType::NotContainsBlanks);
  EXPECT_FALSE(match_rule(r, Value::blank()));
  EXPECT_FALSE(match_rule(r, Value::text("")));
  EXPECT_FALSE(match_rule(r, Value::text("   ")));
  EXPECT_TRUE(match_rule(r, Value::number(1.0)));
  EXPECT_TRUE(match_rule(r, Value::text("x")));
  EXPECT_TRUE(match_rule(r, Value::boolean(true)));
  EXPECT_TRUE(match_rule(r, Value::error(ErrorCode::Div0)));
}

TEST(CFEvaluator, ContainsErrorsMatchesAnyError) {
  CFRule r = MakeRule(RuleType::ContainsErrors);
  EXPECT_TRUE(match_rule(r, Value::error(ErrorCode::Div0)));
  EXPECT_TRUE(match_rule(r, Value::error(ErrorCode::Value)));
  EXPECT_TRUE(match_rule(r, Value::error(ErrorCode::NA)));
  EXPECT_FALSE(match_rule(r, Value::number(0.0)));
  EXPECT_FALSE(match_rule(r, Value::blank()));
  EXPECT_FALSE(match_rule(r, Value::text("not error")));
}

TEST(CFEvaluator, NotContainsErrorsIsComplementOfContainsErrors) {
  CFRule r = MakeRule(RuleType::NotContainsErrors);
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::Div0)));
  EXPECT_TRUE(match_rule(r, Value::number(1.0)));
  EXPECT_TRUE(match_rule(r, Value::blank()));
}

TEST(CFEvaluator, RuleTypesNotYetImplementedReturnFalse) {
  // Pinning the staging contract: the value-only overload returns
  // false for any rule kind that needs evaluation context (formula
  // evaluator, today_serial, sqref population) or that lands in a
  // later PR. A test here catches accidental fall-through.
  for (auto t : {RuleType::Expression, RuleType::ColorScale, RuleType::DataBar, RuleType::IconSet, RuleType::Top10,
                 RuleType::AboveAverage, RuleType::TimePeriod, RuleType::DuplicateValues, RuleType::UniqueValues}) {
    CFRule r = MakeRule(t);
    EXPECT_FALSE(match_rule(r, Value::number(1.0))) << "type=" << static_cast<int>(t);
    EXPECT_FALSE(match_rule(r, Value::blank())) << "type=" << static_cast<int>(t);
    EXPECT_FALSE(match_rule(r, Value::text("x"))) << "type=" << static_cast<int>(t);
  }
}
TEST(CFEvaluator, CellIsLessThanNumeric) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::LessThan;
  r.formula1 = "10";
  EXPECT_TRUE(match_rule(r, Value::number(5.0)));
  EXPECT_FALSE(match_rule(r, Value::number(10.0)));
  EXPECT_FALSE(match_rule(r, Value::number(15.0)));
}

TEST(CFEvaluator, CellIsLessThanOrEqualNumeric) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::LessThanOrEqual;
  r.formula1 = "10";
  EXPECT_TRUE(match_rule(r, Value::number(5.0)));
  EXPECT_TRUE(match_rule(r, Value::number(10.0)));
  EXPECT_FALSE(match_rule(r, Value::number(15.0)));
}

TEST(CFEvaluator, CellIsEqualNumeric) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "42";
  EXPECT_TRUE(match_rule(r, Value::number(42.0)));
  EXPECT_FALSE(match_rule(r, Value::number(41.0)));
  EXPECT_FALSE(match_rule(r, Value::number(43.0)));
}

TEST(CFEvaluator, CellIsNotEqualNumeric) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::NotEqual;
  r.formula1 = "42";
  EXPECT_FALSE(match_rule(r, Value::number(42.0)));
  EXPECT_TRUE(match_rule(r, Value::number(41.0)));
  EXPECT_TRUE(match_rule(r, Value::number(43.0)));
}

TEST(CFEvaluator, CellIsGreaterThanOrEqualNumeric) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::GreaterThanOrEqual;
  r.formula1 = "10";
  EXPECT_FALSE(match_rule(r, Value::number(5.0)));
  EXPECT_TRUE(match_rule(r, Value::number(10.0)));
  EXPECT_TRUE(match_rule(r, Value::number(15.0)));
}

TEST(CFEvaluator, CellIsGreaterThanNumeric) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::GreaterThan;
  r.formula1 = "10";
  EXPECT_FALSE(match_rule(r, Value::number(5.0)));
  EXPECT_FALSE(match_rule(r, Value::number(10.0)));
  EXPECT_TRUE(match_rule(r, Value::number(15.0)));
}

TEST(CFEvaluator, CellIsBetweenNumericIsInclusive) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Between;
  r.formula1 = "5";
  r.formula2 = "10";
  EXPECT_FALSE(match_rule(r, Value::number(4.999)));
  EXPECT_TRUE(match_rule(r, Value::number(5.0)));
  EXPECT_TRUE(match_rule(r, Value::number(7.5)));
  EXPECT_TRUE(match_rule(r, Value::number(10.0)));
  EXPECT_FALSE(match_rule(r, Value::number(10.001)));
}

TEST(CFEvaluator, CellIsNotBetweenNumericIsExclusive) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::NotBetween;
  r.formula1 = "5";
  r.formula2 = "10";
  EXPECT_TRUE(match_rule(r, Value::number(4.999)));
  EXPECT_FALSE(match_rule(r, Value::number(5.0)));
  EXPECT_FALSE(match_rule(r, Value::number(7.5)));
  EXPECT_FALSE(match_rule(r, Value::number(10.0)));
  EXPECT_TRUE(match_rule(r, Value::number(10.001)));
}

TEST(CFEvaluator, CellIsAcceptsSignedAndDecimalLiterals) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::LessThan;
  r.formula1 = "-1.5";
  EXPECT_TRUE(match_rule(r, Value::number(-2.0)));
  EXPECT_FALSE(match_rule(r, Value::number(-1.5)));
  EXPECT_FALSE(match_rule(r, Value::number(0.0)));
}

TEST(CFEvaluator, CellIsEqualTextIsCaseInsensitive) {
  // Excel CF cellIs equality on text is case-insensitive (verified
  // against Mac Excel 365). The fold here is ASCII-only; full Unicode
  // case-folding is the broader text-comparison story's problem.
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "\"hello\"";
  EXPECT_TRUE(match_rule(r, Value::text("hello")));
  EXPECT_TRUE(match_rule(r, Value::text("HELLO")));
  EXPECT_TRUE(match_rule(r, Value::text("HeLLo")));
  EXPECT_FALSE(match_rule(r, Value::text("world")));
}

TEST(CFEvaluator, CellIsEqualNonAsciiTextRoutesThroughEngineCompare) {
  // A cellIs equality rule with a non-ASCII (multi-byte UTF-8) operand must
  // evaluate active/inactive exactly like the engine's `=` operator. The
  // comparison now routes through `eval::compare_values` instead of a
  // CF-local ASCII-only path, so an exact byte match fires and a mismatch
  // stays inactive. The ASCII case-fold leaves non-ASCII bytes untouched, so
  // the comparison is byte-exact for them.
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "\"\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E\"";  // "日本語"
  EXPECT_TRUE(match_rule(r, Value::text("\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E")));
  EXPECT_FALSE(match_rule(r, Value::text("\xE6\x97\xA5\xE6\x9C\xAC")));  // "日本"
  // ASCII letters embedded with the non-ASCII run still fold case-insensitively.
  r.formula1 = "\"caf\xC3\xA9 A\"";                          // "café A"
  EXPECT_TRUE(match_rule(r, Value::text("caf\xC3\xA9 a")));  // "café a"
}

TEST(CFEvaluator, CellIsTextLiteralUnescapesDoubledQuotes) {
  // OOXML escapes embedded `"` as `""` inside the formula text. Verify
  // the parser unescapes it before comparison.
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "\"say \"\"hi\"\"\"";  // source = "say ""hi"""  →  say "hi"
  EXPECT_TRUE(match_rule(r, Value::text("say \"hi\"")));
  EXPECT_FALSE(match_rule(r, Value::text("say hi")));
}

TEST(CFEvaluator, CellIsBooleanLiteral) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "TRUE";
  EXPECT_TRUE(match_rule(r, Value::boolean(true)));
  EXPECT_FALSE(match_rule(r, Value::boolean(false)));

  r.formula1 = "false";  // Excel emits uppercase, but be lenient.
  EXPECT_FALSE(match_rule(r, Value::boolean(true)));
  EXPECT_TRUE(match_rule(r, Value::boolean(false)));
}

TEST(CFEvaluator, CellIsBoolAgainstNumberLiteralCoercesToZeroOne) {
  // Excel treats BOOLs as 1/0 in cellIs numeric comparisons.
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "1";
  EXPECT_TRUE(match_rule(r, Value::boolean(true)));
  EXPECT_FALSE(match_rule(r, Value::boolean(false)));

  r.formula1 = "0";
  EXPECT_FALSE(match_rule(r, Value::boolean(true)));
  EXPECT_TRUE(match_rule(r, Value::boolean(false)));
}

TEST(CFEvaluator, CellIsCrossKindReturnsFalse) {
  // PR7 takes the conservative stance that cross-kind comparisons
  // don't fire. Number rule against text/error/blank cells, and text
  // rule against numeric cells, all return false.
  CFRule num_rule = MakeRule(RuleType::CellIs);
  num_rule.op = CellIsOperator::Equal;
  num_rule.formula1 = "10";
  EXPECT_FALSE(match_rule(num_rule, Value::text("10")));
  EXPECT_FALSE(match_rule(num_rule, Value::error(ErrorCode::Div0)));
  EXPECT_FALSE(match_rule(num_rule, Value::blank()));

  CFRule text_rule = MakeRule(RuleType::CellIs);
  text_rule.op = CellIsOperator::Equal;
  text_rule.formula1 = "\"hello\"";
  EXPECT_FALSE(match_rule(text_rule, Value::number(0.0)));
  EXPECT_FALSE(match_rule(text_rule, Value::error(ErrorCode::Value)));
  EXPECT_FALSE(match_rule(text_rule, Value::blank()));
}

TEST(CFEvaluator, CellIsMissingOperatorReturnsFalse) {
  CFRule r = MakeRule(RuleType::CellIs);
  // op is left unset
  r.formula1 = "10";
  EXPECT_FALSE(match_rule(r, Value::number(5.0)));
}

TEST(CFEvaluator, CellIsMissingFormula1ReturnsFalse) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  // formula1 absent
  EXPECT_FALSE(match_rule(r, Value::number(0.0)));
}

TEST(CFEvaluator, CellIsBetweenMissingFormula2ReturnsFalse) {
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Between;
  r.formula1 = "1";
  // formula2 absent
  EXPECT_FALSE(match_rule(r, Value::number(0.5)));
}

TEST(CFEvaluator, CellIsNonLiteralFormulaReturnsFalse) {
  // PR7 only handles literal operands. Anything that needs the formula
  // evaluator (references, arithmetic) lands with PR8 and silently
  // does not match for now.
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "$A$1";
  EXPECT_FALSE(match_rule(r, Value::number(0.0)));

  r.formula1 = "5+5";
  EXPECT_FALSE(match_rule(r, Value::number(10.0)));

  r.formula1 = "AVERAGE(A1:A10)";
  EXPECT_FALSE(match_rule(r, Value::number(5.0)));
}

// ---------------------------------------------------------------------------
// ContainsText / NotContainsText / BeginsWith / EndsWith
// ---------------------------------------------------------------------------

TEST(CFEvaluator, ContainsTextMatchesSubstring) {
  CFRule r = MakeRule(RuleType::ContainsText);
  r.text = "foo";
  EXPECT_TRUE(match_rule(r, Value::text("foobar")));
  EXPECT_TRUE(match_rule(r, Value::text("hello foo world")));
  EXPECT_TRUE(match_rule(r, Value::text("barfoo")));
  EXPECT_FALSE(match_rule(r, Value::text("bar")));
  EXPECT_FALSE(match_rule(r, Value::text("")));
}

TEST(CFEvaluator, ContainsTextIsAsciiCaseInsensitive) {
  CFRule r = MakeRule(RuleType::ContainsText);
  r.text = "FOO";
  EXPECT_TRUE(match_rule(r, Value::text("foobar")));
  EXPECT_TRUE(match_rule(r, Value::text("FoObAr")));
  EXPECT_TRUE(match_rule(r, Value::text("XfOoY")));
}

TEST(CFEvaluator, ContainsTextEmptyNeedleMatchesAnyText) {
  // Every text contains the empty string; matches Excel's SEARCH-based
  // generated formula.
  CFRule r = MakeRule(RuleType::ContainsText);
  r.text = "";
  EXPECT_TRUE(match_rule(r, Value::text("anything")));
  EXPECT_TRUE(match_rule(r, Value::text("")));
}

TEST(CFEvaluator, ContainsTextMissingTextFieldReturnsFalse) {
  CFRule r = MakeRule(RuleType::ContainsText);
  // r.text not set
  EXPECT_FALSE(match_rule(r, Value::text("foobar")));
}

TEST(CFEvaluator, ContainsTextCoercesNonTextCellToDisplayedText) {
  // Excel's text rules search the cell's displayed text. A number is
  // coerced to its General rendering before the substring test, so a
  // needle present in that rendering matches; one absent from it does
  // not. Error cells have no searchable display text and never match.
  CFRule r = MakeRule(RuleType::ContainsText);
  r.text = "foo";
  EXPECT_FALSE(match_rule(r, Value::number(42.0)));  // "42" lacks "foo"
  EXPECT_FALSE(match_rule(r, Value::boolean(true)));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA)));
  EXPECT_FALSE(match_rule(r, Value::blank()));

  CFRule digit = MakeRule(RuleType::ContainsText);
  digit.text = "2";
  EXPECT_TRUE(match_rule(digit, Value::number(42.0)));   // "42" contains "2"
  EXPECT_FALSE(match_rule(digit, Value::number(99.0)));  // "99" lacks "2"

  CFRule rue = MakeRule(RuleType::ContainsText);
  rue.text = "RUE";
  EXPECT_TRUE(match_rule(rue, Value::boolean(true)));  // "TRUE" contains "RUE"
}

TEST(CFEvaluator, NotContainsTextIsComplementOnTextCells) {
  CFRule r = MakeRule(RuleType::NotContainsText);
  r.text = "foo";
  EXPECT_TRUE(match_rule(r, Value::text("bar")));
  EXPECT_TRUE(match_rule(r, Value::text("")));
  EXPECT_FALSE(match_rule(r, Value::text("foobar")));
  EXPECT_FALSE(match_rule(r, Value::text("FOO")));  // case-insensitive
}

TEST(CFEvaluator, NotContainsTextEmptyNeedleNeverMatchesTextCell) {
  // Every text "contains" the empty string, so its complement never
  // matches.
  CFRule r = MakeRule(RuleType::NotContainsText);
  r.text = "";
  EXPECT_FALSE(match_rule(r, Value::text("anything")));
  EXPECT_FALSE(match_rule(r, Value::text("")));
}

TEST(CFEvaluator, NotContainsTextMatchesNonTextCellLackingNeedle) {
  // The negation is the predicate complement over the coerced displayed
  // text. A numeric or blank cell whose rendering does not contain the
  // needle "does not contain" it, so the rule fires and Excel highlights
  // it. A number whose rendering does contain the needle does not match.
  CFRule r = MakeRule(RuleType::NotContainsText);
  r.text = "foo";
  EXPECT_TRUE(match_rule(r, Value::number(42.0)));  // "42" lacks "foo" -> match
  EXPECT_TRUE(match_rule(r, Value::blank()));       // "" lacks "foo" -> match

  CFRule digit = MakeRule(RuleType::NotContainsText);
  digit.text = "2";
  EXPECT_FALSE(match_rule(digit, Value::number(42.0)));  // "42" contains "2"
  EXPECT_TRUE(match_rule(digit, Value::number(99.0)));   // "99" lacks "2"

  // Error cells carry no searchable display text; neither form applies.
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA)));
}

TEST(CFEvaluator, BeginsWithMatchesPrefix) {
  CFRule r = MakeRule(RuleType::BeginsWith);
  r.text = "Hello";
  EXPECT_TRUE(match_rule(r, Value::text("Hello, World")));
  EXPECT_TRUE(match_rule(r, Value::text("hello world")));  // case-insensitive
  EXPECT_FALSE(match_rule(r, Value::text("World, Hello")));
  EXPECT_FALSE(match_rule(r, Value::text("xHello")));
}

TEST(CFEvaluator, BeginsWithEmptyPrefixMatchesAnyText) {
  CFRule r = MakeRule(RuleType::BeginsWith);
  r.text = "";
  EXPECT_TRUE(match_rule(r, Value::text("anything")));
  EXPECT_TRUE(match_rule(r, Value::text("")));
}

TEST(CFEvaluator, BeginsWithLongerPrefixDoesNotMatch) {
  CFRule r = MakeRule(RuleType::BeginsWith);
  r.text = "longerthancell";
  EXPECT_FALSE(match_rule(r, Value::text("short")));
}

TEST(CFEvaluator, BeginsWithCoercesNonTextCellToDisplayedText) {
  CFRule r = MakeRule(RuleType::BeginsWith);
  r.text = "foo";
  EXPECT_FALSE(match_rule(r, Value::number(42.0)));  // "42" lacks the prefix
  EXPECT_FALSE(match_rule(r, Value::blank()));

  CFRule prefix = MakeRule(RuleType::BeginsWith);
  prefix.text = "4";
  EXPECT_TRUE(match_rule(prefix, Value::number(42.0)));  // "42" begins with "4"
}

TEST(CFEvaluator, EndsWithMatchesSuffix) {
  CFRule r = MakeRule(RuleType::EndsWith);
  r.text = "World";
  EXPECT_TRUE(match_rule(r, Value::text("Hello, World")));
  EXPECT_TRUE(match_rule(r, Value::text("hello world")));  // case-insensitive
  EXPECT_FALSE(match_rule(r, Value::text("World, Hello")));
  EXPECT_FALSE(match_rule(r, Value::text("Worldx")));
}

TEST(CFEvaluator, EndsWithEmptySuffixMatchesAnyText) {
  CFRule r = MakeRule(RuleType::EndsWith);
  r.text = "";
  EXPECT_TRUE(match_rule(r, Value::text("anything")));
  EXPECT_TRUE(match_rule(r, Value::text("")));
}

TEST(CFEvaluator, EndsWithLongerSuffixDoesNotMatch) {
  CFRule r = MakeRule(RuleType::EndsWith);
  r.text = "longerthancell";
  EXPECT_FALSE(match_rule(r, Value::text("short")));
}

TEST(CFEvaluator, EndsWithCoercesNonTextCellToDisplayedText) {
  CFRule r = MakeRule(RuleType::EndsWith);
  r.text = "foo";
  EXPECT_FALSE(match_rule(r, Value::number(42.0)));  // "42" lacks the suffix
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA)));

  CFRule suffix = MakeRule(RuleType::EndsWith);
  suffix.text = "2";
  EXPECT_TRUE(match_rule(suffix, Value::number(42.0)));  // "42" ends with "2"
}

TEST(CFEvaluator, MakeMatchPopulatesIdentityFields) {
  CFRule r = MakeRule(RuleType::ContainsBlanks);
  CFMatch m = make_match(r);
  EXPECT_EQ(m.rule_id, "rule-x");
  EXPECT_EQ(m.priority, 5);
  EXPECT_EQ(m.kind, CFMatchKind::DifferentialFormat);
  ASSERT_TRUE(m.dxf_id.has_value());
  EXPECT_EQ(m.dxf_id.value(), 7u);
  EXPECT_FALSE(m.resolved_fill_color.has_value());
  EXPECT_FALSE(m.data_bar_render.has_value());
  EXPECT_FALSE(m.icon_render.has_value());
}

TEST(CFEvaluator, MakeMatchPropagatesEmptyDxf) {
  CFRule r = MakeRule(RuleType::NotContainsErrors);
  r.dxf_id.reset();
  CFMatch m = make_match(r);
  EXPECT_FALSE(m.dxf_id.has_value());
}

}  // namespace
}  // namespace formulon::cf
