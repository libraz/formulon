// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the conditional-format evaluator. Coverage so far:
// the value-only rule types (ContainsBlanks / NotContainsBlanks /
// ContainsErrors / NotContainsErrors) and `cellIs` against literal
// formula1/formula2. Later PRs add expression, containsText, top10/
// aboveAverage/timePeriod, and the visual rule kinds. The "other rule
// types fall through to false" guarantee is pinned here so the staging
// strategy stays observable.

#include "cf/cf_evaluator.h"

#include "cell.h"
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
  // Pinning the staging contract: the rule types whose evaluator logic
  // lands in subsequent PRs must not silently match anything in the
  // meantime. A test here catches accidental fall-through.
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

TEST(CFEvaluator, ContainsTextNonTextCellDoesNotMatch) {
  // Conservative cross-kind stance, mirroring cellIs. Documented in
  // cf_evaluator.h.
  CFRule r = MakeRule(RuleType::ContainsText);
  r.text = "foo";
  EXPECT_FALSE(match_rule(r, Value::number(42.0)));
  EXPECT_FALSE(match_rule(r, Value::boolean(true)));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA)));
  EXPECT_FALSE(match_rule(r, Value::blank()));
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

TEST(CFEvaluator, NotContainsTextNonTextCellDoesNotMatch) {
  // Symmetric with the positive form: non-text cells never match
  // either variant. Documented in cf_evaluator.h.
  CFRule r = MakeRule(RuleType::NotContainsText);
  r.text = "foo";
  EXPECT_FALSE(match_rule(r, Value::number(42.0)));
  EXPECT_FALSE(match_rule(r, Value::blank()));
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

TEST(CFEvaluator, BeginsWithNonTextCellDoesNotMatch) {
  CFRule r = MakeRule(RuleType::BeginsWith);
  r.text = "foo";
  EXPECT_FALSE(match_rule(r, Value::number(42.0)));
  EXPECT_FALSE(match_rule(r, Value::blank()));
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

TEST(CFEvaluator, EndsWithNonTextCellDoesNotMatch) {
  CFRule r = MakeRule(RuleType::EndsWith);
  r.text = "foo";
  EXPECT_FALSE(match_rule(r, Value::number(42.0)));
  EXPECT_FALSE(match_rule(r, Value::error(ErrorCode::NA)));
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
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 10)), ctx));  // Sun
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 13)), ctx));  // Wed
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 16)), ctx));  // Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 9)), ctx));   // prior Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 17)), ctx));  // next Sun
}

TEST(CFEvaluator, TimePeriodLastWeekIsPriorSundayThroughSaturday) {
  CFEvalHarness harness;
  CFRule r = MakeRule(RuleType::TimePeriod);
  r.time_period = TimePeriod::LastWeek;
  const auto ctx = PinnedContext(harness, Serial(2024, 3, 13));
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 3)), ctx));   // Sun
  EXPECT_TRUE(match_rule(r, Value::number(Serial(2024, 3, 9)), ctx));   // Sat
  EXPECT_FALSE(match_rule(r, Value::number(Serial(2024, 3, 2)), ctx));  // 2 weeks back
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

}  // namespace
}  // namespace formulon::cf
