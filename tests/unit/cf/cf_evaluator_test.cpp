// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the conditional-format evaluator. PR6 covers the
// value-only rule types (ContainsBlanks / NotContainsBlanks /
// ContainsErrors / NotContainsErrors); later PRs add cellIs, expression,
// containsText, top10/aboveAverage/timePeriod, and the visual rule
// kinds. The "other rule types fall through to false" guarantee is
// pinned here so the staging strategy stays observable.

#include "cf/cf_evaluator.h"

#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "gtest/gtest.h"
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
  // Pinning the staging contract: the sixteen rule types whose evaluator
  // logic lands in subsequent PRs must not silently match anything in
  // the meantime. A test here catches accidental fall-through.
  for (auto t : {RuleType::Expression, RuleType::CellIs, RuleType::ColorScale, RuleType::DataBar, RuleType::IconSet,
                 RuleType::Top10, RuleType::AboveAverage, RuleType::ContainsText, RuleType::NotContainsText,
                 RuleType::BeginsWith, RuleType::EndsWith, RuleType::TimePeriod, RuleType::DuplicateValues,
                 RuleType::UniqueValues}) {
    CFRule r = MakeRule(t);
    EXPECT_FALSE(match_rule(r, Value::number(1.0))) << "type=" << static_cast<int>(t);
    EXPECT_FALSE(match_rule(r, Value::blank())) << "type=" << static_cast<int>(t);
    EXPECT_FALSE(match_rule(r, Value::text("x"))) << "type=" << static_cast<int>(t);
  }
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
