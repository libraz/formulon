// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for the conditional-format evaluator. Coverage so far:
// the value-only rule types (ContainsBlanks / NotContainsBlanks /
// ContainsErrors / NotContainsErrors) and `cellIs` against literal
// formula1/formula2. Later PRs add expression, containsText, top10/
// aboveAverage/timePeriod, and the visual rule kinds. The "other rule
// types fall through to false" guarantee is pinned here so the staging
// strategy stays observable.

#include "cf/cf_evaluator.h"

#include <vector>

#include "cell.h"
#include "cf/cf_helpers.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "eval/date_time.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::cf {
namespace {

CFRule MakeRule(RuleType t) {
  CFRule r;
  r.type = t;
  r.priority = 5;
  r.id = "rule-x";
  r.dxf_id = 7u;
  return r;
}

TEST(CFEvaluator, ContainsBlanksMatchesBlank) {
  CFRule r = MakeRule(RuleType::ContainsBlanks);
  EXPECT_TRUE(match_rule(r, Value::blank()));
  EXPECT_FALSE(match_rule(r, Value::number(0.0)));
  EXPECT_FALSE(match_rule(r, Value::text("")));
  EXPECT_FALSE(match_rule(r, Value::boolean(false)));
}

TEST(CFEvaluator, NotContainsBlanksIsComplementOfContainsBlanks) {
  CFRule r = MakeRule(RuleType::NotContainsBlanks);
  EXPECT_FALSE(match_rule(r, Value::blank()));
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

// ---------------------------------------------------------------------------
// Context-aware overload: Expression rules and CellIs with non-literal
// formula1 / formula2.
// ---------------------------------------------------------------------------

CellAddress At(std::uint32_t row, std::uint32_t col) {
  CellAddress addr{};
  addr.row = row;
  addr.col = col;
  return addr;
}

struct CFEvalHarness {
  Sheet sheet{"Sheet1"};
  Arena arena;
  eval::EvalContext eval_ctx{sheet};

  CFEvalContext context(CellAddress anchor, CellAddress target) {
    CFEvalContext ctx;
    ctx.anchor = anchor;
    ctx.target = target;
    ctx.arena = &arena;
    ctx.registry = &eval::default_registry();
    ctx.eval_ctx = &eval_ctx;
    return ctx;
  }
};

TEST(CFEvaluator, ExpressionRuleLiteralTrueMatches) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "TRUE";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleLiteralFalseDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "FALSE";
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleNonZeroNumberMatches) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "1";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleZeroNumberDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "0";
  EXPECT_FALSE(match_rule(r, Value::number(1.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleArithmeticComparison) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(15.0));  // A1 = 15
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "A1>10";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleRelativeReferenceShiftsToTargetRow) {
  // Anchor at A1, target at A4. Formula `=A1>10` shifts to `=A4>10`.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(5.0));   // A1
  harness.sheet.set_cell_value(3, 0, Value::number(15.0));  // A4
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "A1>10";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(3, 0))));
}

TEST(CFEvaluator, ExpressionRuleRelativeReferenceShiftsToTargetColumn) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(5.0));   // A1
  harness.sheet.set_cell_value(0, 2, Value::number(20.0));  // C1
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "A1>10";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 2))));
}

TEST(CFEvaluator, ExpressionRuleAbsoluteReferenceLocksAtAnchor) {
  // `$A$1>10` always reads A1, regardless of target.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(15.0));  // A1
  harness.sheet.set_cell_value(0, 5, Value::number(0.0));   // F1
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "$A$1>10";
  EXPECT_TRUE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 5))));
}

TEST(CFEvaluator, ExpressionRuleErrorResultDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "1/0";  // produces #DIV/0!
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleTextResultDoesNotMatch) {
  // Excel's expression-rule truthiness rejects text — only bool / number trigger.
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "\"yes\"";
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleParseFailureDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "(((";
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleMissingFormulaDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  // formula1 unset
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ExpressionRuleOutOfBoundsShiftDoesNotMatch) {
  // Anchor at A1, target at A1 with delta = (-1, 0) is impossible
  // (target above anchor by one row would land on row 0 = okay).
  // Pick a formula referencing A1 with anchor at B5 and target at A1
  // → delta = (-4, -1). Shift A1 by (-4, -1) goes out of bounds.
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::Expression);
  r.formula1 = "A1>10";
  // The shifted ref becomes #REF!; comparison `(#REF!)>10` propagates
  // an error, which is not truthy.
  EXPECT_FALSE(match_rule(r, Value::number(0.0), harness.context(At(4, 1), At(0, 0))));
}

// ---------------------------------------------------------------------------
// CellIs with non-literal operand
// ---------------------------------------------------------------------------

TEST(CFEvaluator, CellIsEvaluatesReferenceOperand) {
  // CellIs Equal `=A1`, with A1 = 10. Cell value = 10 → matches.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(10.0));  // A1
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "A1";
  EXPECT_TRUE(match_rule(r, Value::number(10.0), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(r, Value::number(11.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, CellIsEvaluatesArithmeticOperand) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::LessThan;
  r.formula1 = "10*2";  // = 20
  EXPECT_TRUE(match_rule(r, Value::number(15.0), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(r, Value::number(25.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, CellIsBetweenWithReferenceOperands) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(5.0));   // A1
  harness.sheet.set_cell_value(0, 1, Value::number(10.0));  // B1
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Between;
  r.formula1 = "A1";
  r.formula2 = "B1";
  EXPECT_TRUE(match_rule(r, Value::number(7.5), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(r, Value::number(11.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, CellIsLiteralPathStillWorksThroughEvaluatorOverload) {
  // Verify the context-aware overload doesn't regress the literal path.
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "42";
  EXPECT_TRUE(match_rule(r, Value::number(42.0), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(r, Value::number(43.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, CellIsReferenceShiftsWithTarget) {
  // CellIs with `A1` operand, anchor at A1, target at A2 → operand is A2.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(1, 0, Value::number(50.0));  // A2
  CFRule r = MakeRule(RuleType::CellIs);
  r.op = CellIsOperator::Equal;
  r.formula1 = "A1";
  EXPECT_TRUE(match_rule(r, Value::number(50.0), harness.context(At(0, 0), At(1, 0))));
}

TEST(CFEvaluator, ContextAwareOverloadDelegatesValueOnlyKinds) {
  // Verify the context-aware overload routes value-only rules to the
  // simple overload. This is a regression guard: if the dispatcher
  // ever forgets to delegate, the rule would silently return false.
  CFEvalHarness harness;
  CFRule blanks = MakeRule(RuleType::ContainsBlanks);
  EXPECT_TRUE(match_rule(blanks, Value::blank(), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(blanks, Value::number(0.0), harness.context(At(0, 0), At(0, 0))));

  CFRule errors = MakeRule(RuleType::ContainsErrors);
  EXPECT_TRUE(match_rule(errors, Value::error(ErrorCode::Div0), harness.context(At(0, 0), At(0, 0))));

  CFRule contains = MakeRule(RuleType::ContainsText);
  contains.text = "foo";
  EXPECT_TRUE(match_rule(contains, Value::text("foobar"), harness.context(At(0, 0), At(0, 0))));
  EXPECT_FALSE(match_rule(contains, Value::text("bar"), harness.context(At(0, 0), At(0, 0))));
}

// ---------------------------------------------------------------------------
// TimePeriod
//
// Anchored at Wednesday 2024-03-13 throughout. Excel weekday semantics
// place that day at WEEKDAY=4 (Sun=1). The week therefore spans
// 2024-03-10 (Sun) through 2024-03-16 (Sat).
// ---------------------------------------------------------------------------

CFEvalContext PinnedContext(CFEvalHarness& harness, double today_serial) {
  CFEvalContext ctx = harness.context(At(0, 0), At(0, 0));
  ctx.today_serial = today_serial;
  return ctx;
}

double Serial(int year, unsigned month, unsigned day) {
  return eval::date_time::serial_from_ymd(year, month, day);
}

TEST(CFEvaluator, TimePeriodWithoutTodaySerialDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Today;
  // ctx.today_serial intentionally left as nullopt.
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 13)), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, TimePeriodMissingBucketDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  // r.time_period not set
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 13)), PinnedContext(harness, Serial(2024, 3, 13))));
}

TEST(CFEvaluator, TimePeriodNonNumericCellDoesNotMatch) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Today;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_FALSE(match_rule(r, Value::text("2024-03-13"), ctx));
  EXPECT_FALSE(match_rule(r, Value::boolean(true), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
}

TEST(CFEvaluator, TimePeriodTodayMatchesSameDay) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Today;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 13)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 12)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 14)), ctx));
}

TEST(CFEvaluator, TimePeriodTodayDropsTimeOfDayFraction) {
  // Cell carrying date + 0.75 (= 18:00) still matches Today.
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Today;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 13) + 0.75), ctx));
}

TEST(CFEvaluator, TimePeriodYesterdayMatchesPreviousDay) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Yesterday;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 12)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 11)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 13)), ctx));
}

TEST(CFEvaluator, TimePeriodTomorrowMatchesNextDay) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Tomorrow;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 14)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 13)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 15)), ctx));
}

TEST(CFEvaluator, TimePeriodLast7DaysSpansSixDaysBackThroughToday) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Last7Days;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 13)), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 7)), ctx));   // today - 6
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 6)), ctx));  // today - 7
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 14)), ctx));
}

TEST(CFEvaluator, TimePeriodThisWeekIsSundayThroughSaturday) {
  // 2024-03-13 is Wed. Week = 2024-03-10 (Sun) .. 2024-03-16 (Sat).
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::ThisWeek;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 10)), ctx));   // Sun
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 13)), ctx));   // Wed
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 16)), ctx));   // Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 9)), ctx));   // prior Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 17)), ctx));  // next Sun
}

TEST(CFEvaluator, TimePeriodLastWeekIsPriorSundayThroughSaturday) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::LastWeek;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 3)), ctx));    // Sun
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 9)), ctx));    // Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 2)), ctx));   // 2 weeks back
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 10)), ctx));  // this Sun
}

TEST(CFEvaluator, TimePeriodNextWeekIsFollowingSundayThroughSaturday) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::NextWeek;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 17)), ctx));   // Sun
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 23)), ctx));   // Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 16)), ctx));  // this Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 24)), ctx));  // 2 weeks ahead
}

TEST(CFEvaluator, TimePeriodThisMonthMatchesSameYearAndMonth) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::ThisMonth;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 1)), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 31)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 2, 29)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 4, 1)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2023, 3, 13)), ctx));  // same month, prior year
}

TEST(CFEvaluator, TimePeriodLastMonthHandlesYearBoundary) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::LastMonth;
  // today = 2024-01-15 → last month = 2023-12.
  const auto ctx = PinnedContext(harness, Serial(2024, 1, 15));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2023, 12, 1)), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2023, 12, 31)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 1, 1)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2023, 11, 30)), ctx));
}

TEST(CFEvaluator, TimePeriodNextMonthHandlesYearBoundary) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::NextMonth;
  // today = 2024-12-15 → next month = 2025-01.
  const auto ctx = PinnedContext(harness, Serial(2024, 12, 15));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2025, 1, 1)), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2025, 1, 31)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 12, 31)), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2025, 2, 1)), ctx));
}

TEST(CFEvaluator, TimePeriodValueOnlyOverloadStillReturnsFalse) {
  // The value-only overload has no today reference, so TimePeriod
  // continues to return false there. Pin the contract.
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::Today;
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 13))));
}

// ---------------------------------------------------------------------------
// DuplicateValues / UniqueValues
// ---------------------------------------------------------------------------

CFCellRange MakeRange(std::uint32_t r1, std::uint32_t c1, std::uint32_t r2, std::uint32_t c2) {
  CFCellRange range{};
  range.first.row = r1;
  range.first.col = c1;
  range.last.row = r2;
  range.last.col = c2;
  return range;
}

CFEvalContext SqrefContext(CFEvalHarness& harness, const std::vector<CFCellRange>& sqref) {
  CFEvalContext ctx = harness.context(At(0, 0), At(0, 0));
  ctx.sqref = &sqref;
  return ctx;
}

TEST(CFEvaluator, DuplicateValuesWithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(7.0));
  harness.sheet.set_cell_value(0, 1, Value::number(7.0));
  CFRule r = MakeRule(RuleType::DuplicateValues);
  // ctx.sqref intentionally nullptr.
  EXPECT_FALSE(match_rule(r, Value::number(7.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, DuplicateValuesMatchesNumberAppearingTwice) {
  CFEvalHarness harness;
  // A1:A3 = [7, 7, 9]. Value 7 appears twice → duplicate; 9 appears once.
  harness.sheet.set_cell_value(0, 0, Value::number(7.0));
  harness.sheet.set_cell_value(1, 0, Value::number(7.0));
  harness.sheet.set_cell_value(2, 0, Value::number(9.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(7.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(9.0), ctx));
}

TEST(CFEvaluator, DuplicateValuesNumbersUseExactEquality) {
  // 1.0 and 1.0000000001 are distinct under IEEE-754 equality, even
  // though they round to the same display string.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(1.0));
  harness.sheet.set_cell_value(1, 0, Value::number(1.0000000001));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 1, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(1.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(1.0000000001), ctx));
}

TEST(CFEvaluator, DuplicateValuesTextIsAsciiCaseInsensitive) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::text("apple"));
  harness.sheet.set_cell_value(1, 0, Value::text("APPLE"));
  harness.sheet.set_cell_value(2, 0, Value::text("orange"));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::text("apple"), ctx));
  EXPECT_TRUE(match_rule(r, Value::text("Apple"), ctx));
  EXPECT_FALSE(match_rule(r, Value::text("orange"), ctx));
}

TEST(CFEvaluator, DuplicateValuesBooleansMatchByIdentity) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::boolean(true));
  harness.sheet.set_cell_value(1, 0, Value::boolean(true));
  harness.sheet.set_cell_value(2, 0, Value::boolean(false));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::boolean(true), ctx));
  EXPECT_FALSE(match_rule(r, Value::boolean(false), ctx));
}

TEST(CFEvaluator, DuplicateValuesIsCrossKindFalse) {
  // Number 1 and text "1" are distinct, even though Excel coerces in
  // some contexts. Mirroring the cellIs cross-kind=false stance.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(1.0));
  harness.sheet.set_cell_value(1, 0, Value::text("1"));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 1, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(1.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::text("1"), ctx));
}

TEST(CFEvaluator, DuplicateValuesErrorsAndBlanksDoNotMatch) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::error(ErrorCode::NA));
  harness.sheet.set_cell_value(1, 0, Value::error(ErrorCode::NA));
  // A3 left blank.
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
}

TEST(CFEvaluator, DuplicateValuesAcrossMultipleSqrefRanges) {
  // sqref unions A1:A2 and C1:C2. Value 5 appears in A1 and C2 → dup.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(5.0));
  harness.sheet.set_cell_value(1, 0, Value::number(99.0));
  harness.sheet.set_cell_value(0, 2, Value::number(11.0));
  harness.sheet.set_cell_value(1, 2, Value::number(5.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 1, 0), MakeRange(0, 2, 1, 2)};
  CFRule r = MakeRule(RuleType::DuplicateValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(5.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(99.0), ctx));
}

TEST(CFEvaluator, UniqueValuesMatchesValueAppearingExactlyOnce) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(7.0));
  harness.sheet.set_cell_value(1, 0, Value::number(7.0));
  harness.sheet.set_cell_value(2, 0, Value::number(9.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::UniqueValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(9.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(7.0), ctx));
}

TEST(CFEvaluator, UniqueValuesNonStorableValueKindDoesNotMatch) {
  // Errors and blanks never match either rule.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(1.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 0, 0)};
  CFRule r = MakeRule(RuleType::UniqueValues);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
}

TEST(CFEvaluator, DuplicateUniqueValueOnlyOverloadStillReturnsFalse) {
  // Pin the staging contract: value-only overload has no sqref access,
  // so DuplicateValues / UniqueValues continue to return false there.
  CFRule dup = MakeRule(RuleType::DuplicateValues);
  EXPECT_FALSE(match_rule(dup, Value::number(1.0)));
  CFRule uniq = MakeRule(RuleType::UniqueValues);
  EXPECT_FALSE(match_rule(uniq, Value::number(1.0)));
}

// ---------------------------------------------------------------------------
// AboveAverage
// ---------------------------------------------------------------------------
//
// All AboveAverage tests use the population [10, 20, 30, 40, 50] in
// A1:A5 — mean = 30, sample std-dev = sqrt(250) ≈ 15.811.

void PopulateLinearPopulation(CFEvalHarness& harness) {
  harness.sheet.set_cell_value(0, 0, Value::number(10.0));
  harness.sheet.set_cell_value(1, 0, Value::number(20.0));
  harness.sheet.set_cell_value(2, 0, Value::number(30.0));
  harness.sheet.set_cell_value(3, 0, Value::number(40.0));
  harness.sheet.set_cell_value(4, 0, Value::number(50.0));
}

TEST(CFEvaluator, AboveAverageWithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule r = MakeRule(RuleType::AboveAverage);
  // ctx.sqref intentionally nullptr.
  EXPECT_FALSE(match_rule(r, Value::number(40.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, AboveAverageMatchesValuesStrictlyAboveMean) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  r.above_average = true;
  r.equal_average = false;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(40.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));  // mean
  EXPECT_FALSE(match_rule(r, Value::number(20.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(10.0), ctx));
}

TEST(CFEvaluator, AboveAverageEqualAverageIncludesMean) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  r.above_average = true;
  r.equal_average = true;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(40.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(30.0), ctx));  // mean inclusive
  EXPECT_FALSE(match_rule(r, Value::number(29.999), ctx));
}

TEST(CFEvaluator, AboveAverageBelowSideMatchesValuesStrictlyBelowMean) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  r.above_average = false;
  r.equal_average = false;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(20.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(40.0), ctx));
}

TEST(CFEvaluator, AboveAverageStdDevShiftsThresholdAboveMean) {
  // mean = 30, sample stddev ≈ 15.811. With std_dev = 1, threshold ≈
  // 45.811 → only 50 (above) matches; 40 falls below.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  r.above_average = true;
  r.equal_average = false;
  r.std_dev = 1.0;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(40.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
}

TEST(CFEvaluator, AboveAverageStdDevShiftsThresholdBelowMeanForBelowSide) {
  // mean = 30, sample stddev ≈ 15.811. With std_dev = 1 on the below
  // side, threshold ≈ 14.189 → only 10 matches.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  r.above_average = false;
  r.equal_average = false;
  r.std_dev = 1.0;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(20.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
}

TEST(CFEvaluator, AboveAverageNonNumericCellDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::text("40"), ctx));
  EXPECT_FALSE(match_rule(r, Value::boolean(true), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
}

TEST(CFEvaluator, AboveAverageEmptyPopulationDoesNotMatch) {
  CFEvalHarness harness;  // sqref points at cells that are all blank.
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(0.0), ctx));
}

TEST(CFEvaluator, AboveAverageBooleansAndTextExcludedFromPopulation) {
  // Population is [10, 50] (numbers only) → mean = 30. Boolean TRUE
  // and text "100" do not contribute.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(10.0));
  harness.sheet.set_cell_value(1, 0, Value::boolean(true));
  harness.sheet.set_cell_value(2, 0, Value::text("100"));
  harness.sheet.set_cell_value(3, 0, Value::number(50.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 3, 0)};
  CFRule r = MakeRule(RuleType::AboveAverage);
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(20.0), ctx));
}

TEST(CFEvaluator, AboveAverageValueOnlyOverloadStillReturnsFalse) {
  // Pin the staging contract: AboveAverage continues to return false on
  // the value-only overload (no sqref access).
  CFRule r = MakeRule(RuleType::AboveAverage);
  EXPECT_FALSE(match_rule(r, Value::number(50.0)));
}

// ---------------------------------------------------------------------------
// Top10
// ---------------------------------------------------------------------------

TEST(CFEvaluator, Top10WithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 2;
  EXPECT_FALSE(match_rule(r, Value::number(50.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, Top10MatchesTopNValuesByRank) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);  // [10, 20, 30, 40, 50]
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 2;
  r.bottom = false;
  r.percent = false;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(40.0), ctx));  // 2nd-largest
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(10.0), ctx));
}

TEST(CFEvaluator, Top10BottomMatchesBottomNValuesByRank) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 2;
  r.bottom = true;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(20.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(50.0), ctx));
}

TEST(CFEvaluator, Top10TiesAtThresholdAreIncluded) {
  // Population is [40, 40, 40, 30, 10]. Top 2 by rank → threshold = 40,
  // and all three 40s match because the comparison is `>= threshold`.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(40.0));
  harness.sheet.set_cell_value(1, 0, Value::number(40.0));
  harness.sheet.set_cell_value(2, 0, Value::number(40.0));
  harness.sheet.set_cell_value(3, 0, Value::number(30.0));
  harness.sheet.set_cell_value(4, 0, Value::number(10.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 2;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(40.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(10.0), ctx));
}

TEST(CFEvaluator, Top10PercentInterpretsRankAsPercentOfPopulation) {
  // Population size 10 → top 30% = floor(10 * 30 / 100) = 3.
  CFEvalHarness harness;
  for (std::uint32_t row = 0; row < 10; ++row) {
    harness.sheet.set_cell_value(row, 0, Value::number(static_cast<double>(row + 1)));
  }
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 9, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 30;
  r.percent = true;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(9.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(8.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(7.0), ctx));
}

TEST(CFEvaluator, Top10PercentClampsToAtLeastOne) {
  // Population size 5, percent = 1 → floor(5 * 1 / 100) = 0; clamps to 1.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 1;
  r.percent = true;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_FALSE(match_rule(r, Value::number(40.0), ctx));
}

TEST(CFEvaluator, Top10RankExceedingPopulationClampsToAll) {
  // Rank 100 with a 5-cell population → threshold = min/max so every
  // cell matches.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 100;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
}

TEST(CFEvaluator, Top10ZeroOrNegativeRankDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 0;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(50.0), ctx));
  r.rank = -3;
  const auto ctx2 = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(50.0), ctx2));
}

TEST(CFEvaluator, Top10NonNumericCellDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 5;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::text("50"), ctx));
  EXPECT_FALSE(match_rule(r, Value::boolean(true), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
}

TEST(CFEvaluator, Top10EmptyPopulationDoesNotMatch) {
  CFEvalHarness harness;
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 2;
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_FALSE(match_rule(r, Value::number(0.0), ctx));
}

TEST(CFEvaluator, Top10DefaultRankIsTen) {
  // No `rank` set → defaults to 10. With only 5 numbers, that clamps to
  // 5 (the entire population), so every value matches.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::Top10);
  // r.rank intentionally unset.
  const auto ctx = SqrefContext(harness, sqref);
  EXPECT_TRUE(match_rule(r, Value::number(50.0), ctx));
  EXPECT_TRUE(match_rule(r, Value::number(10.0), ctx));
}

TEST(CFEvaluator, Top10ValueOnlyOverloadStillReturnsFalse) {
  // Pin the staging contract: Top10 continues to return false on the
  // value-only overload (no sqref access).
  CFRule r = MakeRule(RuleType::Top10);
  r.rank = 5;
  EXPECT_FALSE(match_rule(r, Value::number(50.0)));
}

// ---------------------------------------------------------------------------
// ColorScale
// ---------------------------------------------------------------------------

Color RGB(std::uint8_t red, std::uint8_t green, std::uint8_t blue) {
  Color color{};
  color.r = red;
  color.g = green;
  color.b = blue;
  color.a = 255;
  return color;
}

CfValueObject Cfvo(CfvoType type, std::string value = "") {
  CfValueObject cfvo;
  cfvo.type = type;
  cfvo.value = std::move(value);
  return cfvo;
}

ColorScaleSpec TwoStopMinMax(Color lo, Color hi) {
  ColorScaleSpec spec;
  spec.thresholds = {Cfvo(CfvoType::Min), Cfvo(CfvoType::Max)};
  spec.colors = {lo, hi};
  return spec;
}

ColorScaleSpec ThreeStopMinMidMax(Color lo, Color mid, Color hi, std::string mid_percent = "50") {
  ColorScaleSpec spec;
  spec.thresholds = {Cfvo(CfvoType::Min), Cfvo(CfvoType::Percentile, std::move(mid_percent)), Cfvo(CfvoType::Max)};
  spec.colors = {lo, mid, hi};
  return spec;
}

TEST(CFEvaluator, ColorScaleWithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, ColorScaleTwoStopMinMaxInterpolatesBetweenStops) {
  // Population [10, 20, 30, 40, 50]; min = 10 (red), max = 50 (green).
  // Cell 30 is the midpoint → R=128, G=128, B=0 (linear RGB blend).
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(30.0), ctx);
  EXPECT_EQ(match.kind, CFMatchKind::ColorScale);
  ASSERT_TRUE(match.resolved_fill_color.has_value());
  EXPECT_EQ(match.resolved_fill_color->r, 128);
  EXPECT_EQ(match.resolved_fill_color->g, 128);
  EXPECT_EQ(match.resolved_fill_color->b, 0);
}

TEST(CFEvaluator, ColorScaleClampsToOuterStops) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  // Below min: clamps to red.
  CFMatch lo = make_match(r, Value::number(-100.0), ctx);
  ASSERT_TRUE(lo.resolved_fill_color.has_value());
  EXPECT_EQ(*lo.resolved_fill_color, RGB(255, 0, 0));
  // Above max: clamps to green.
  CFMatch hi = make_match(r, Value::number(1000.0), ctx);
  ASSERT_TRUE(hi.resolved_fill_color.has_value());
  EXPECT_EQ(*hi.resolved_fill_color, RGB(0, 255, 0));
}

TEST(CFEvaluator, ColorScaleAtMinAndMaxReturnsExactStopColors) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch low = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(low.resolved_fill_color.has_value());
  EXPECT_EQ(*low.resolved_fill_color, RGB(255, 0, 0));
  CFMatch high = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(high.resolved_fill_color.has_value());
  EXPECT_EQ(*high.resolved_fill_color, RGB(0, 255, 0));
}

TEST(CFEvaluator, ColorScaleThreeStopUsesMiddleStopAtMedian) {
  // Population [10, 20, 30, 40, 50]; mid = percentile(50) = 30 (white).
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = ThreeStopMinMidMax(RGB(255, 0, 0), RGB(255, 255, 255), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch mid = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(mid.resolved_fill_color.has_value());
  EXPECT_EQ(*mid.resolved_fill_color, RGB(255, 255, 255));
}

TEST(CFEvaluator, ColorScaleThreeStopInterpolatesWithinSegment) {
  // Population [10, 20, 30, 40, 50]; segments 10..30 (red→white) and
  // 30..50 (white→green). Cell 20 is the midpoint of the lower segment.
  // Lower segment blend: (255, 0, 0) → (255, 255, 255), fraction 0.5
  // → (255, 128, 128).
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = ThreeStopMinMidMax(RGB(255, 0, 0), RGB(255, 255, 255), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(20.0), ctx);
  ASSERT_TRUE(match.resolved_fill_color.has_value());
  EXPECT_EQ(match.resolved_fill_color->r, 255);
  EXPECT_EQ(match.resolved_fill_color->g, 128);
  EXPECT_EQ(match.resolved_fill_color->b, 128);
}

TEST(CFEvaluator, ColorScaleNumberCfvoUsesLiteralThreshold) {
  // Force min=0, max=100 via Number CFVOs irrespective of population.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  ColorScaleSpec spec;
  spec.thresholds = {Cfvo(CfvoType::Number, "0"), Cfvo(CfvoType::Number, "100")};
  spec.colors = {RGB(0, 0, 0), RGB(255, 255, 255)};
  r.color_scale = std::move(spec);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(match.resolved_fill_color.has_value());
  // Halfway between black and white = (128, 128, 128).
  EXPECT_EQ(*match.resolved_fill_color, RGB(128, 128, 128));
}

TEST(CFEvaluator, ColorScalePercentCfvoUsesPopulationRangeFraction) {
  // Population [10, 20, 30, 40, 50]; min=10, max=50 → range=40.
  // Percent 25 → 10 + 0.25*40 = 20. Percent 75 → 10 + 0.75*40 = 40.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  ColorScaleSpec spec;
  spec.thresholds = {Cfvo(CfvoType::Percent, "25"), Cfvo(CfvoType::Percent, "75")};
  spec.colors = {RGB(0, 0, 0), RGB(255, 255, 255)};
  r.color_scale = std::move(spec);
  const auto ctx = SqrefContext(harness, sqref);

  // Cell at lower-stop position (20) → black.
  CFMatch lo = make_match(r, Value::number(20.0), ctx);
  ASSERT_TRUE(lo.resolved_fill_color.has_value());
  EXPECT_EQ(*lo.resolved_fill_color, RGB(0, 0, 0));
  // Cell at upper-stop position (40) → white.
  CFMatch hi = make_match(r, Value::number(40.0), ctx);
  ASSERT_TRUE(hi.resolved_fill_color.has_value());
  EXPECT_EQ(*hi.resolved_fill_color, RGB(255, 255, 255));
}

TEST(CFEvaluator, ColorScaleNonNumericCellDoesNotResolve) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::text("middle"), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
  CFMatch match = make_match(r, Value::text("middle"), ctx);
  EXPECT_FALSE(match.resolved_fill_color.has_value());
}

TEST(CFEvaluator, ColorScaleEmptyPopulationDoesNotResolve) {
  CFEvalHarness harness;  // No values populated; sqref is all-blank.
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::number(0.0), ctx));
  CFMatch match = make_match(r, Value::number(0.0), ctx);
  EXPECT_FALSE(match.resolved_fill_color.has_value());
}

TEST(CFEvaluator, ColorScaleDegeneratePopulationCollapsesToFirstStop) {
  // Population is [7, 7, 7, 7]; min == max. The cell value 7 hits the
  // first stop's clamp branch → returns colors[0].
  CFEvalHarness harness;
  for (std::uint32_t row = 0; row < 4; ++row) {
    harness.sheet.set_cell_value(row, 0, Value::number(7.0));
  }
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 3, 0)};
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(7.0), ctx);
  ASSERT_TRUE(match.resolved_fill_color.has_value());
  EXPECT_EQ(*match.resolved_fill_color, RGB(255, 0, 0));
}

TEST(CFEvaluator, ColorScaleFormulaCfvoEvaluatesAtAnchor) {
  // CFVO formulas reference cells; value comes from the formula evaluator.
  // A1 = 0 (min anchor), B1 = 100 (max anchor); cell 50 → mid grey.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(0.0));
  harness.sheet.set_cell_value(0, 1, Value::number(100.0));
  // Population is the union: still 0..100 across [A1, B1].
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 0, 1)};
  CFRule r = MakeRule(RuleType::ColorScale);
  ColorScaleSpec spec;
  spec.thresholds = {Cfvo(CfvoType::Formula, "A1"), Cfvo(CfvoType::Formula, "B1")};
  spec.colors = {RGB(0, 0, 0), RGB(255, 255, 255)};
  r.color_scale = std::move(spec);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(match.resolved_fill_color.has_value());
  EXPECT_EQ(*match.resolved_fill_color, RGB(128, 128, 128));
}

TEST(CFEvaluator, ColorScaleValueOnlyOverloadStillReturnsFalse) {
  // Pin the staging contract: ColorScale returns false on the
  // value-only overload (no sqref / population access). The value-only
  // make_match still returns a DifferentialFormat match.
  CFRule r = MakeRule(RuleType::ColorScale);
  r.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  EXPECT_FALSE(match_rule(r, Value::number(30.0)));

  CFMatch match = make_match(r);
  EXPECT_EQ(match.kind, CFMatchKind::DifferentialFormat);
  EXPECT_FALSE(match.resolved_fill_color.has_value());
}

// ---------------------------------------------------------------------------
// DataBar
// ---------------------------------------------------------------------------

DataBarSpec MakeDataBarSpec(CfValueObject min_cfvo, CfValueObject max_cfvo, Color fill = RGB(0, 128, 255),
                            DataBarAxisPosition axis = DataBarAxisPosition::Automatic) {
  DataBarSpec spec;
  spec.min = std::move(min_cfvo);
  spec.max = std::move(max_cfvo);
  spec.fill = fill;
  spec.negative_fill = RGB(255, 0, 0);
  spec.axis_position = axis;
  spec.min_length_pct = 0;
  spec.max_length_pct = 100;
  return spec;
}

TEST(CFEvaluator, DataBarWithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  EXPECT_FALSE(match_rule(r, Value::number(30.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, DataBarLengthIsLinearBetweenMinAndMax) {
  // Population [10, 20, 30, 40, 50]; min=10, max=50 → range=40.
  // Cell at min → length 0%. Cell at max → length 100%. Cell at 30 → 50%.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch lo = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(lo.data_bar_render.has_value());
  EXPECT_EQ(lo.kind, CFMatchKind::DataBar);
  EXPECT_DOUBLE_EQ(lo.data_bar_render->length_pct, 0.0);

  CFMatch hi = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(hi.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(hi.data_bar_render->length_pct, 100.0);

  CFMatch mid = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(mid.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(mid.data_bar_render->length_pct, 50.0);
}

TEST(CFEvaluator, DataBarClampsValuesOutsideThresholdRange) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch below = make_match(r, Value::number(-100.0), ctx);
  ASSERT_TRUE(below.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(below.data_bar_render->length_pct, 0.0);

  CFMatch above = make_match(r, Value::number(1000.0), ctx);
  ASSERT_TRUE(above.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(above.data_bar_render->length_pct, 100.0);
}

TEST(CFEvaluator, DataBarMinAndMaxLengthAreApplied) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  DataBarSpec spec = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  spec.min_length_pct = 10;
  spec.max_length_pct = 90;
  r.data_bar = std::move(spec);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch lo = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(lo.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(lo.data_bar_render->length_pct, 10.0);

  CFMatch mid = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(mid.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(mid.data_bar_render->length_pct, 50.0);

  CFMatch hi = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(hi.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(hi.data_bar_render->length_pct, 90.0);
}

TEST(CFEvaluator, DataBarAutomaticAxisAtZeroForAllNonNegativePopulation) {
  // Population [10, 20, 30, 40, 50] is all >= 0 → axis at left edge.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->axis_position_pct, 0.0);
  EXPECT_FALSE(match.data_bar_render->is_negative);
}

TEST(CFEvaluator, DataBarAutomaticAxisAtHundredForAllNegativePopulation) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(-50.0));
  harness.sheet.set_cell_value(1, 0, Value::number(-30.0));
  harness.sheet.set_cell_value(2, 0, Value::number(-10.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(-30.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->axis_position_pct, 100.0);
  EXPECT_TRUE(match.data_bar_render->is_negative);
}

TEST(CFEvaluator, DataBarAutomaticAxisProportionalForMixedSignPopulation) {
  // Population [-30, 10, 70] → min = -30, max = 70. Negative span = 30,
  // total span = 100 → axis at 30%.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(-30.0));
  harness.sheet.set_cell_value(1, 0, Value::number(10.0));
  harness.sheet.set_cell_value(2, 0, Value::number(70.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 2, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->axis_position_pct, 30.0);
}

TEST(CFEvaluator, DataBarMixedSignLengthSplitsAtAxisInsteadOfWholeRangeLinear) {
  // Population [-50, -10, 20, 100] → min = -50, max = 100. Axis at
  // |min| / (max - min) = 50 / 150 ≈ 33.33%.
  //
  // Regression: length used to be the whole-range linear map
  // `(cell - min) / (max - min)`, which for e.g. cell=-10 would give
  // `(-10 - -50) / 150 = 26.67%` -- a bar nearly as long as the true
  // min. The correct OOXML semantics split at the axis: positive bars
  // scale by `value / max`, negative bars by `value / min` (both
  // negative, so a positive fraction) -- cell=-10 should be a *short*
  // bar (20% of the negative side), not a long one.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(-50.0));
  harness.sheet.set_cell_value(1, 0, Value::number(-10.0));
  harness.sheet.set_cell_value(2, 0, Value::number(20.0));
  harness.sheet.set_cell_value(3, 0, Value::number(100.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 3, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch axis_probe = make_match(r, Value::number(20.0), ctx);
  ASSERT_TRUE(axis_probe.data_bar_render.has_value());
  EXPECT_NEAR(axis_probe.data_bar_render->axis_position_pct, 33.333333333333336, 1e-9);

  // Most-negative value: full-length bar on the negative side.
  CFMatch most_negative = make_match(r, Value::number(-50.0), ctx);
  ASSERT_TRUE(most_negative.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(most_negative.data_bar_render->length_pct, 100.0);
  EXPECT_TRUE(most_negative.data_bar_render->is_negative);

  // -10 is 20% of the way from 0 to min (-50): a short negative bar.
  CFMatch small_negative = make_match(r, Value::number(-10.0), ctx);
  ASSERT_TRUE(small_negative.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(small_negative.data_bar_render->length_pct, 20.0);
  EXPECT_TRUE(small_negative.data_bar_render->is_negative);

  // 20 is 20% of the way from 0 to max (100): a short positive bar.
  CFMatch small_positive = make_match(r, Value::number(20.0), ctx);
  ASSERT_TRUE(small_positive.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(small_positive.data_bar_render->length_pct, 20.0);
  EXPECT_FALSE(small_positive.data_bar_render->is_negative);

  // Most-positive value (= max): full-length bar on the positive side.
  CFMatch most_positive = make_match(r, Value::number(100.0), ctx);
  ASSERT_TRUE(most_positive.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(most_positive.data_bar_render->length_pct, 100.0);
  EXPECT_FALSE(most_positive.data_bar_render->is_negative);
}

TEST(CFEvaluator, DataBarAllPositiveDataKeepsWholeRangeLinearLength) {
  // Non-regression: same-sign data must keep the original whole-range
  // linear map even though it also uses Automatic axis mode.
  // Population [10, 20, 30, 40, 50] -- min=10, max=50, mirrors
  // `DataBarLengthIsLinearBetweenMinAndMax` but pinned to the mixed-sign
  // code path's guard condition (all non-negative here).
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch mid = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(mid.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(mid.data_bar_render->length_pct, 50.0);
}

TEST(CFEvaluator, DataBarMiddleAxisIsAlwaysFifty) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max), RGB(0, 0, 255), DataBarAxisPosition::Middle);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->axis_position_pct, 50.0);
}

TEST(CFEvaluator, DataBarNoneAxisIsZero) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max), RGB(0, 0, 255), DataBarAxisPosition::None);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->axis_position_pct, 0.0);
}

TEST(CFEvaluator, DataBarSelectsNegativeFillForNegativeValues) {
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(-30.0));
  harness.sheet.set_cell_value(1, 0, Value::number(70.0));
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 1, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  DataBarSpec spec = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  spec.fill = RGB(0, 200, 0);
  spec.negative_fill = RGB(200, 0, 0);
  r.data_bar = std::move(spec);
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch positive = make_match(r, Value::number(70.0), ctx);
  ASSERT_TRUE(positive.data_bar_render.has_value());
  EXPECT_EQ(positive.data_bar_render->fill, RGB(0, 200, 0));
  EXPECT_FALSE(positive.data_bar_render->is_negative);

  CFMatch negative = make_match(r, Value::number(-30.0), ctx);
  ASSERT_TRUE(negative.data_bar_render.has_value());
  EXPECT_EQ(negative.data_bar_render->fill, RGB(200, 0, 0));
  EXPECT_TRUE(negative.data_bar_render->is_negative);
}

TEST(CFEvaluator, DataBarNumberCfvosUseLiteralThresholds) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  // Force min=0, max=100 regardless of population.
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Number, "0"), Cfvo(CfvoType::Number, "100"));
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch match = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(match.data_bar_render.has_value());
  EXPECT_DOUBLE_EQ(match.data_bar_render->length_pct, 50.0);
}

TEST(CFEvaluator, DataBarNonNumericCellDoesNotResolve) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::text("30"), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
  CFMatch match = make_match(r, Value::text("30"), ctx);
  EXPECT_FALSE(match.data_bar_render.has_value());
}

TEST(CFEvaluator, DataBarDegenerateRangeDoesNotResolve) {
  // Population is all 7s → min == max → no meaningful bar length.
  CFEvalHarness harness;
  for (std::uint32_t row = 0; row < 4; ++row) {
    harness.sheet.set_cell_value(row, 0, Value::number(7.0));
  }
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 3, 0)};
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::number(7.0), ctx));
  CFMatch match = make_match(r, Value::number(7.0), ctx);
  EXPECT_FALSE(match.data_bar_render.has_value());
}

TEST(CFEvaluator, DataBarValueOnlyOverloadStillReturnsFalse) {
  CFRule r = MakeRule(RuleType::DataBar);
  r.data_bar = MakeDataBarSpec(Cfvo(CfvoType::Min), Cfvo(CfvoType::Max));
  EXPECT_FALSE(match_rule(r, Value::number(30.0)));
  CFMatch match = make_match(r);
  EXPECT_EQ(match.kind, CFMatchKind::DifferentialFormat);
  EXPECT_FALSE(match.data_bar_render.has_value());
}

// ---------------------------------------------------------------------------
// IconSet
// ---------------------------------------------------------------------------

CfValueObject IconCfvo(CfvoType type, std::string value, bool gte = true) {
  CfValueObject cfvo = Cfvo(type, std::move(value));
  cfvo.gte = gte;
  return cfvo;
}

IconSetSpec ThreeIconNumberSet(std::string lo, std::string hi, IconSetName name = IconSetName::Three_TrafficLights1) {
  IconSetSpec spec;
  spec.name = name;
  spec.thresholds = {IconCfvo(CfvoType::Number, std::move(lo)), IconCfvo(CfvoType::Number, std::move(hi))};
  return spec;
}

TEST(CFEvaluator, IconSetWithoutSqrefDoesNotMatch) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule r = MakeRule(RuleType::IconSet);
  r.icon_set = ThreeIconNumberSet("20", "40");
  EXPECT_FALSE(match_rule(r, Value::number(30.0), harness.context(At(0, 0), At(0, 0))));
}

TEST(CFEvaluator, IconSetThreeIconAssignsBucketByThreshold) {
  // Thresholds 20 / 40 → bucket 0: cell < 20; bucket 1: 20 <= cell < 40;
  // bucket 2: cell >= 40.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  r.icon_set = ThreeIconNumberSet("20", "40");
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch low = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(low.icon_render.has_value());
  EXPECT_EQ(low.kind, CFMatchKind::IconSet);
  EXPECT_EQ(low.icon_render->icon_index, 0);
  EXPECT_EQ(low.icon_render->set_name, IconSetName::Three_TrafficLights1);

  CFMatch mid = make_match(r, Value::number(30.0), ctx);
  ASSERT_TRUE(mid.icon_render.has_value());
  EXPECT_EQ(mid.icon_render->icon_index, 1);

  CFMatch high = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(high.icon_render.has_value());
  EXPECT_EQ(high.icon_render->icon_index, 2);
}

TEST(CFEvaluator, IconSetExactlyAtThresholdHonoursGteFlag) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  // gte=true on threshold 20: cell == 20 belongs to the upper bucket.
  IconSetSpec spec_gte;
  spec_gte.name = IconSetName::Three_TrafficLights1;
  spec_gte.thresholds = {IconCfvo(CfvoType::Number, "20", true), IconCfvo(CfvoType::Number, "40", true)};
  r.icon_set = spec_gte;
  const auto ctx_gte = SqrefContext(harness, sqref);
  CFMatch at_threshold_gte = make_match(r, Value::number(20.0), ctx_gte);
  ASSERT_TRUE(at_threshold_gte.icon_render.has_value());
  EXPECT_EQ(at_threshold_gte.icon_render->icon_index, 1);

  // gte=false on threshold 20: cell == 20 belongs to the lower bucket.
  IconSetSpec spec_gt;
  spec_gt.name = IconSetName::Three_TrafficLights1;
  spec_gt.thresholds = {IconCfvo(CfvoType::Number, "20", false), IconCfvo(CfvoType::Number, "40", false)};
  r.icon_set = spec_gt;
  const auto ctx_gt = SqrefContext(harness, sqref);
  CFMatch at_threshold_gt = make_match(r, Value::number(20.0), ctx_gt);
  ASSERT_TRUE(at_threshold_gt.icon_render.has_value());
  EXPECT_EQ(at_threshold_gt.icon_render->icon_index, 0);
}

TEST(CFEvaluator, IconSetReverseFlipsBucketIndex) {
  // Without reverse: 10 → 0, 50 → 2. With reverse: 10 → 2, 50 → 0.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  IconSetSpec spec = ThreeIconNumberSet("20", "40");
  spec.reverse = true;
  r.icon_set = spec;
  const auto ctx = SqrefContext(harness, sqref);

  CFMatch low = make_match(r, Value::number(10.0), ctx);
  ASSERT_TRUE(low.icon_render.has_value());
  EXPECT_EQ(low.icon_render->icon_index, 2);

  CFMatch high = make_match(r, Value::number(50.0), ctx);
  ASSERT_TRUE(high.icon_render.has_value());
  EXPECT_EQ(high.icon_render->icon_index, 0);
}

TEST(CFEvaluator, IconSetFiveIconAssignsAcrossFourThresholds) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  IconSetSpec spec;
  spec.name = IconSetName::Five_Arrows;
  spec.thresholds = {IconCfvo(CfvoType::Number, "15"), IconCfvo(CfvoType::Number, "25"),
                     IconCfvo(CfvoType::Number, "35"), IconCfvo(CfvoType::Number, "45")};
  r.icon_set = spec;
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_EQ(make_match(r, Value::number(10.0), ctx).icon_render->icon_index, 0);
  EXPECT_EQ(make_match(r, Value::number(20.0), ctx).icon_render->icon_index, 1);
  EXPECT_EQ(make_match(r, Value::number(30.0), ctx).icon_render->icon_index, 2);
  EXPECT_EQ(make_match(r, Value::number(40.0), ctx).icon_render->icon_index, 3);
  EXPECT_EQ(make_match(r, Value::number(50.0), ctx).icon_render->icon_index, 4);
}

TEST(CFEvaluator, IconSetPercentCfvoUsesPopulationRangeFraction) {
  // Population [10, 20, 30, 40, 50]; min=10, max=50 → range=40.
  // Thresholds at 25% (=20) and 75% (=40).
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  IconSetSpec spec;
  spec.name = IconSetName::Three_TrafficLights1;
  spec.thresholds = {IconCfvo(CfvoType::Percent, "25"), IconCfvo(CfvoType::Percent, "75")};
  r.icon_set = spec;
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_EQ(make_match(r, Value::number(10.0), ctx).icon_render->icon_index, 0);
  EXPECT_EQ(make_match(r, Value::number(30.0), ctx).icon_render->icon_index, 1);
  EXPECT_EQ(make_match(r, Value::number(50.0), ctx).icon_render->icon_index, 2);
}

TEST(CFEvaluator, IconSetNonNumericCellDoesNotResolve) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  r.icon_set = ThreeIconNumberSet("20", "40");
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::text("30"), ctx));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA), ctx));
  EXPECT_FALSE(match_rule(r, Value::blank(), ctx));
  CFMatch match = make_match(r, Value::text("30"), ctx);
  EXPECT_FALSE(match.icon_render.has_value());
}

TEST(CFEvaluator, IconSetEmptyPopulationDoesNotResolve) {
  CFEvalHarness harness;
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  r.icon_set = ThreeIconNumberSet("20", "40");
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
  CFMatch match = make_match(r, Value::number(30.0), ctx);
  EXPECT_FALSE(match.icon_render.has_value());
}

TEST(CFEvaluator, IconSetEmptyThresholdsDoesNotResolve) {
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  const std::vector<CFCellRange> sqref{MakeRange(0, 0, 4, 0)};
  CFRule r = MakeRule(RuleType::IconSet);
  IconSetSpec spec;
  spec.name = IconSetName::Three_TrafficLights1;
  // No thresholds.
  r.icon_set = spec;
  const auto ctx = SqrefContext(harness, sqref);

  EXPECT_FALSE(match_rule(r, Value::number(30.0), ctx));
}

TEST(CFEvaluator, IconSetValueOnlyOverloadStillReturnsFalse) {
  CFRule r = MakeRule(RuleType::IconSet);
  r.icon_set = ThreeIconNumberSet("20", "40");
  EXPECT_FALSE(match_rule(r, Value::number(30.0)));
  CFMatch match = make_match(r);
  EXPECT_EQ(match.kind, CFMatchKind::DifferentialFormat);
  EXPECT_FALSE(match.icon_render.has_value());
}

// ---------------------------------------------------------------------------
// evaluate_cf_at — cross-block priority chain + stopIfTrue
// ---------------------------------------------------------------------------

CFHost MakeHost(CFEvalHarness& harness) {
  CFHost host;
  host.arena = &harness.arena;
  host.registry = &eval::default_registry();
  host.eval_ctx = &harness.eval_ctx;
  return host;
}

ConditionalFormat MakeBlock(std::vector<CFCellRange> sqref, std::vector<CFRule> rules) {
  ConditionalFormat block;
  block.sqref = std::move(sqref);
  block.rules = std::move(rules);
  return block;
}

CFRule MakeBlanksRule(std::int32_t priority, std::uint32_t dxf_id, std::string id, bool stop_if_true = false) {
  CFRule rule;
  rule.type = RuleType::ContainsBlanks;
  rule.priority = priority;
  rule.dxf_id = dxf_id;
  rule.id = std::move(id);
  rule.stop_if_true = stop_if_true;
  return rule;
}

TEST(CFEvaluator, EvaluateCfAtWithoutBlocksReturnsEmpty) {
  CFEvalHarness harness;
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  EXPECT_TRUE(matches.empty());
}

TEST(CFEvaluator, EvaluateCfAtIgnoresBlocksWhoseSqrefExcludesTarget) {
  CFEvalHarness harness;
  // A1 is blank; block sqref is C1:D2 → does not contain A1.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 2, 1, 3)}, {MakeBlanksRule(1, 7, "rule-1")}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  EXPECT_TRUE(matches.empty());
}

TEST(CFEvaluator, EvaluateCfAtReturnsMatchingDifferentialFormat) {
  CFEvalHarness harness;
  // A1 is blank by default; ContainsBlanks rule should fire.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 1, 1)}, {MakeBlanksRule(1, 7, "rule-1")}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].rule_id, "rule-1");
  EXPECT_EQ(matches[0].priority, 1);
  EXPECT_EQ(matches[0].kind, CFMatchKind::DifferentialFormat);
  ASSERT_TRUE(matches[0].dxf_id.has_value());
  EXPECT_EQ(*matches[0].dxf_id, 7u);
}

TEST(CFEvaluator, EvaluateCfAtSortsByPriorityAscending) {
  CFEvalHarness harness;
  // Two rules in the same block, declared in reverse priority order.
  // Both fire (cell A1 is blank); evaluator must sort priority=1 first.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 1, 1)},
                {MakeBlanksRule(5, 50, "rule-low-priority"), MakeBlanksRule(1, 10, "rule-high-priority")}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  ASSERT_EQ(matches.size(), 2u);
  EXPECT_EQ(matches[0].rule_id, "rule-high-priority");
  EXPECT_EQ(matches[0].priority, 1);
  EXPECT_EQ(matches[1].rule_id, "rule-low-priority");
  EXPECT_EQ(matches[1].priority, 5);
}

TEST(CFEvaluator, EvaluateCfAtSortsAcrossBlocks) {
  CFEvalHarness harness;
  // Two separate blocks — priority is workbook-global so the
  // priority-2 rule from the second block evaluates between the
  // priority-1 and priority-3 rules of the first block.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 1, 1)}, {MakeBlanksRule(1, 11, "p1"), MakeBlanksRule(3, 13, "p3")}));
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(2, 12, "p2")}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  ASSERT_EQ(matches.size(), 3u);
  EXPECT_EQ(matches[0].rule_id, "p1");
  EXPECT_EQ(matches[1].rule_id, "p2");
  EXPECT_EQ(matches[2].rule_id, "p3");
}

TEST(CFEvaluator, EvaluateCfAtStopIfTrueHaltsEvaluation) {
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 1, 1)},
                {MakeBlanksRule(1, 10, "first", /*stop_if_true=*/true), MakeBlanksRule(2, 20, "second")}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].rule_id, "first");
}

TEST(CFEvaluator, EvaluateCfAtStopIfTrueDoesNotHaltOnNonMatch) {
  // First rule does not match (cell is non-blank), so stop_if_true is
  // never triggered; the second rule still evaluates.
  CFEvalHarness harness;
  harness.sheet.set_cell_value(0, 0, Value::number(7.0));

  CFRule cell_is_rule;
  cell_is_rule.type = RuleType::CellIs;
  cell_is_rule.priority = 1;
  cell_is_rule.dxf_id = 100;
  cell_is_rule.id = "cell-is-zero";
  cell_is_rule.op = CellIsOperator::Equal;
  cell_is_rule.formula1 = "0";
  cell_is_rule.stop_if_true = true;

  CFRule errors_rule;
  errors_rule.type = RuleType::NotContainsErrors;
  errors_rule.priority = 2;
  errors_rule.dxf_id = 200;
  errors_rule.id = "not-error";

  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {cell_is_rule, errors_rule}));
  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].rule_id, "not-error");
}

TEST(CFEvaluator, EvaluateCfAtVisualRulePopulatesRenderPayload) {
  // ColorScale rule on A1:A5; cell A3 is the population midpoint.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  CFRule color_rule = MakeRule(RuleType::ColorScale);
  color_rule.priority = 1;
  color_rule.id = "color-1";
  color_rule.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  harness.sheet.mutable_conditional_formats().push_back(MakeBlock({MakeRange(0, 0, 4, 0)}, {color_rule}));

  const auto host = MakeHost(harness);
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(2, 0), host);
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].kind, CFMatchKind::ColorScale);
  ASSERT_TRUE(matches[0].resolved_fill_color.has_value());
  EXPECT_EQ(matches[0].resolved_fill_color->r, 128);
  EXPECT_EQ(matches[0].resolved_fill_color->g, 128);
  EXPECT_EQ(matches[0].resolved_fill_color->b, 0);
}

TEST(CFEvaluator, EvaluateCfAtMissingHostFieldsReturnsEmpty) {
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(1, 7, "rule-1")}));
  CFHost host;  // arena/registry/eval_ctx all null.
  std::vector<CFMatch> matches = evaluate_cf_at(harness.sheet, At(0, 0), host);
  EXPECT_TRUE(matches.empty());
}

// ---------------------------------------------------------------------------
// evaluate_cf_for_range — viewport-range API
// ---------------------------------------------------------------------------

TEST(CFEvaluator, EvaluateCfForRangeReturnsOneEntryPerMatchedCell) {
  CFEvalHarness harness;
  // A1 and A2 are blank → both match ContainsBlanks. A3 has a value.
  harness.sheet.set_cell_value(2, 0, Value::number(42.0));
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 2, 0)}, {MakeBlanksRule(1, 7, "blanks")}));

  const auto host = MakeHost(harness);
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 2, 0), host);
  ASSERT_EQ(matches.size(), 2u);
  EXPECT_EQ(matches[0].cell.row, 0u);
  EXPECT_EQ(matches[0].cell.col, 0u);
  EXPECT_EQ(matches[1].cell.row, 1u);
  EXPECT_EQ(matches[1].cell.col, 0u);
  ASSERT_EQ(matches[0].matches.size(), 1u);
  EXPECT_EQ(matches[0].matches[0].rule_id, "blanks");
}

TEST(CFEvaluator, EvaluateCfForRangeEmitsRowMajorOrder) {
  CFEvalHarness harness;
  // 2x2 viewport, all blank → all 4 cells match.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 1, 1)}, {MakeBlanksRule(1, 7, "blanks")}));

  const auto host = MakeHost(harness);
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 1, 1), host);
  ASSERT_EQ(matches.size(), 4u);
  // Row-major: (0,0), (0,1), (1,0), (1,1).
  EXPECT_EQ(matches[0].cell.row, 0u);
  EXPECT_EQ(matches[0].cell.col, 0u);
  EXPECT_EQ(matches[1].cell.row, 0u);
  EXPECT_EQ(matches[1].cell.col, 1u);
  EXPECT_EQ(matches[2].cell.row, 1u);
  EXPECT_EQ(matches[2].cell.col, 0u);
  EXPECT_EQ(matches[3].cell.row, 1u);
  EXPECT_EQ(matches[3].cell.col, 1u);
}

TEST(CFEvaluator, EvaluateCfForRangeSkipsCellsWithNoMatches) {
  CFEvalHarness harness;
  // Block applies only to A1; B1 has no rules.
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(1, 7, "blanks")}));
  const auto host = MakeHost(harness);
  // Viewport spans A1:B1 → only A1 should appear in the result.
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 0, 1), host);
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].cell.row, 0u);
  EXPECT_EQ(matches[0].cell.col, 0u);
}

TEST(CFEvaluator, EvaluateCfForRangeSingleCellRangeStillVisited) {
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(1, 7, "blanks")}));
  const auto host = MakeHost(harness);
  // first == last (A1:A1).
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 0, 0), host);
  ASSERT_EQ(matches.size(), 1u);
  EXPECT_EQ(matches[0].cell.row, 0u);
  EXPECT_EQ(matches[0].cell.col, 0u);
}

TEST(CFEvaluator, EvaluateCfForRangeReturnsEmptyForEmptyHost) {
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 0)}, {MakeBlanksRule(1, 7, "blanks")}));
  CFHost host;  // null fields.
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 0, 0), host);
  EXPECT_TRUE(matches.empty());
}

TEST(CFEvaluator, EvaluateCfForRangeAggregatesPriorityOrderPerCell) {
  // Two rules at different priorities; both match every cell. Pin that
  // each cell's match list comes back in priority order.
  CFEvalHarness harness;
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 0, 1)}, {MakeBlanksRule(2, 20, "later"), MakeBlanksRule(1, 10, "earlier")}));
  const auto host = MakeHost(harness);
  std::vector<CFRangeCellMatches> matches = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 0, 1), host);
  ASSERT_EQ(matches.size(), 2u);
  for (const auto& cell : matches) {
    ASSERT_EQ(cell.matches.size(), 2u);
    EXPECT_EQ(cell.matches[0].rule_id, "earlier");
    EXPECT_EQ(cell.matches[1].rule_id, "later");
  }
}

// ---------------------------------------------------------------------------
// Population caching — viewport API must produce results identical to
// per-cell `evaluate_cf_at` calls. This pins the optimisation: caching
// the population across cells in a range may not change rendered output.
// ---------------------------------------------------------------------------

CFRule MakeColorScaleBlock(std::int32_t priority, std::string id) {
  CFRule rule;
  rule.type = RuleType::ColorScale;
  rule.priority = priority;
  rule.id = std::move(id);
  rule.color_scale = TwoStopMinMax(RGB(255, 0, 0), RGB(0, 255, 0));
  return rule;
}

CFRule MakeTop10Rule(std::int32_t priority, std::string id, std::int32_t rank) {
  CFRule rule;
  rule.type = RuleType::Top10;
  rule.priority = priority;
  rule.dxf_id = 5;
  rule.id = std::move(id);
  rule.rank = rank;
  rule.bottom = false;
  rule.percent = false;
  return rule;
}

TEST(CFEvaluator, EvaluateCfForRangeCachedPathMatchesUncached) {
  // ColorScale over A1:A5 = [10, 20, 30, 40, 50]. Every cell is in
  // the sqref so `evaluate_cf_at` and `evaluate_cf_for_range` must
  // resolve identical fill colours.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 4, 0)}, {MakeColorScaleBlock(1, "color-scale")}));

  const auto host = MakeHost(harness);
  std::vector<CFRangeCellMatches> ranged = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 4, 0), host);
  ASSERT_EQ(ranged.size(), 5u);

  for (std::uint32_t row = 0; row < 5; ++row) {
    std::vector<CFMatch> per_cell = evaluate_cf_at(harness.sheet, At(row, 0), host);
    ASSERT_EQ(per_cell.size(), 1u) << "row=" << row;
    ASSERT_EQ(ranged[row].matches.size(), 1u) << "row=" << row;
    const auto& cached = ranged[row].matches[0];
    const auto& fresh = per_cell[0];
    ASSERT_TRUE(cached.resolved_fill_color.has_value()) << "row=" << row;
    ASSERT_TRUE(fresh.resolved_fill_color.has_value()) << "row=" << row;
    EXPECT_EQ(cached.resolved_fill_color->r, fresh.resolved_fill_color->r) << "row=" << row;
    EXPECT_EQ(cached.resolved_fill_color->g, fresh.resolved_fill_color->g) << "row=" << row;
    EXPECT_EQ(cached.resolved_fill_color->b, fresh.resolved_fill_color->b) << "row=" << row;
    EXPECT_EQ(cached.resolved_fill_color->a, fresh.resolved_fill_color->a) << "row=" << row;
  }
}

TEST(CFEvaluator, EvaluateCfForRangeCachedPathPreservesTop10Behavior) {
  // Top-2 over [10, 20, 30, 40, 50] should match rows 3 (40) and 4
  // (50). Cached and uncached paths must agree on which cells fire.
  CFEvalHarness harness;
  PopulateLinearPopulation(harness);
  harness.sheet.mutable_conditional_formats().push_back(
      MakeBlock({MakeRange(0, 0, 4, 0)}, {MakeTop10Rule(1, "top2", 2)}));

  const auto host = MakeHost(harness);
  std::vector<CFRangeCellMatches> ranged = evaluate_cf_for_range(harness.sheet, MakeRange(0, 0, 4, 0), host);

  // Row-major: only rows 3 and 4 should appear.
  ASSERT_EQ(ranged.size(), 2u);
  EXPECT_EQ(ranged[0].cell.row, 3u);
  EXPECT_EQ(ranged[1].cell.row, 4u);

  // Cross-check by walking cell-by-cell with the uncached entry point.
  for (std::uint32_t row = 0; row < 5; ++row) {
    std::vector<CFMatch> per_cell = evaluate_cf_at(harness.sheet, At(row, 0), host);
    const bool expected = row >= 3;
    EXPECT_EQ(per_cell.size(), expected ? 1u : 0u) << "row=" << row;
  }
}

// A rule whose sqref is an explicit multi-million-cell rectangle (not
// whole-column / whole-row notation) must be clamped to the populated
// extent before scanning, so the scan stays bounded and still returns the
// same match/numeric results as the tight rectangle would.
TEST(CFHelpers, ExplicitGiantSqrefIsClampedToPopulatedExtent) {
  Sheet sheet("S");
  sheet.set_cell_cached_value(0, 0, Value::number(5.0));
  sheet.set_cell_cached_value(50, 5, Value::number(5.0));
  sheet.set_cell_cached_value(100, 3, Value::number(5.0));
  sheet.set_cell_cached_value(100, 10, Value::number(9.0));  // populated, non-matching

  // Explicit rectangle far larger than the ~65k clamp threshold, yet not
  // full-column (last.row != kCfMaxRows-1) nor full-row (last.col !=
  // kCfMaxCols-1) — the case the old code walked cell-by-cell.
  CFCellRange giant;
  giant.first = CellAddress{0, 0};
  giant.last = CellAddress{400000, 4000};
  ASSERT_FALSE(giant.is_full_col());
  ASSERT_FALSE(giant.is_full_row());
  const std::vector<CFCellRange> sqref{giant};

  EXPECT_EQ(helpers::count_matches_in_sqref(Value::number(5.0), sqref, sheet), 3u);

  const std::vector<double> numbers = helpers::collect_numeric_values(sqref, sheet);
  EXPECT_EQ(numbers.size(), 4u);  // three 5s + one 9
  double sum = 0.0;
  for (double n : numbers) {
    sum += n;
  }
  EXPECT_DOUBLE_EQ(sum, 24.0);
}

}  // namespace
}  // namespace formulon::cf
