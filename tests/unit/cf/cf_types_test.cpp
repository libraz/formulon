// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Default-construction and basic invariant tests for the conditional-
// format data model. These structures are header-only and behaviour-free
// at this stage; the suite locks in the contract that subsequent PRs
// (reader, writer, evaluator, sheet integration) build on.

#include "cf/cf_types.h"

#include "gtest/gtest.h"

namespace formulon::cf {
namespace {

TEST(CFTypes, RuleTypeEnumPinning) {
  EXPECT_EQ(static_cast<std::uint8_t>(RuleType::Expression), 0u);
  EXPECT_EQ(static_cast<std::uint8_t>(RuleType::CellIs), 1u);
  EXPECT_EQ(static_cast<std::uint8_t>(RuleType::ColorScale), 2u);
  EXPECT_EQ(static_cast<std::uint8_t>(RuleType::DataBar), 3u);
  EXPECT_EQ(static_cast<std::uint8_t>(RuleType::IconSet), 4u);
  EXPECT_EQ(static_cast<std::uint8_t>(RuleType::UniqueValues), 17u);
}

TEST(CFTypes, CellIsOperatorEnumPinning) {
  EXPECT_EQ(static_cast<std::uint8_t>(CellIsOperator::LessThan), 0u);
  EXPECT_EQ(static_cast<std::uint8_t>(CellIsOperator::Between), 6u);
  EXPECT_EQ(static_cast<std::uint8_t>(CellIsOperator::NotBetween), 7u);
}

TEST(CFTypes, CfvoTypeEnumPinning) {
  EXPECT_EQ(static_cast<std::uint8_t>(CfvoType::Number), 0u);
  EXPECT_EQ(static_cast<std::uint8_t>(CfvoType::AutoMin), 6u);
  EXPECT_EQ(static_cast<std::uint8_t>(CfvoType::AutoMax), 7u);
}

TEST(CFTypes, IconSetNameEnumPinning) {
  EXPECT_EQ(static_cast<std::uint8_t>(IconSetName::Three_Arrows), 0u);
  EXPECT_EQ(static_cast<std::uint8_t>(IconSetName::Five_Quarters), 16u);
}

TEST(CFTypes, ColorDefaultIsOpaqueBlack) {
  Color c;
  EXPECT_EQ(c.r, 0u);
  EXPECT_EQ(c.g, 0u);
  EXPECT_EQ(c.b, 0u);
  EXPECT_EQ(c.a, 255u);
}

TEST(CFTypes, ColorEquality) {
  Color a{1, 2, 3, 4};
  Color b{1, 2, 3, 4};
  Color c{1, 2, 3, 5};
  EXPECT_EQ(a, b);
  EXPECT_NE(a, c);
}

TEST(CFTypes, CellRangeEquality) {
  CFCellRange r1{{0, 0}, {3, 4}};
  CFCellRange r2{{0, 0}, {3, 4}};
  CFCellRange r3{{0, 0}, {3, 5}};
  EXPECT_EQ(r1, r2);
  EXPECT_NE(r1, r3);
}

TEST(CFTypes, CfValueObjectDefaults) {
  CfValueObject v;
  EXPECT_EQ(v.type, CfvoType::Number);
  EXPECT_TRUE(v.value.empty());
  EXPECT_TRUE(v.gte);
}

TEST(CFTypes, ColorScaleSpecDefaults) {
  ColorScaleSpec s;
  EXPECT_TRUE(s.thresholds.empty());
  EXPECT_TRUE(s.colors.empty());
}

TEST(CFTypes, DataBarSpecDefaults) {
  DataBarSpec d;
  EXPECT_FALSE(d.border.has_value());
  EXPECT_FALSE(d.negative_border.has_value());
  EXPECT_EQ(d.axis_position, DataBarAxisPosition::Automatic);
  EXPECT_EQ(d.axis_color, Color({0, 0, 0, 255}));
  EXPECT_TRUE(d.gradient);
  EXPECT_TRUE(d.show_value);
  EXPECT_EQ(d.min_length_pct, 10u);
  EXPECT_EQ(d.max_length_pct, 90u);
}

TEST(CFTypes, IconSetSpecDefaults) {
  IconSetSpec i;
  EXPECT_EQ(i.name, IconSetName::Three_Arrows);
  EXPECT_TRUE(i.thresholds.empty());
  EXPECT_FALSE(i.reverse);
  EXPECT_TRUE(i.show_value);
  EXPECT_TRUE(i.percent);
}

TEST(CFRule, AllOptionalsEmptyByDefault) {
  CFRule r;
  EXPECT_TRUE(r.id.empty());
  EXPECT_EQ(r.type, RuleType::Expression);
  EXPECT_EQ(r.priority, 1);
  EXPECT_FALSE(r.stop_if_true);
  EXPECT_FALSE(r.dxf_id.has_value());
  EXPECT_FALSE(r.formula1.has_value());
  EXPECT_FALSE(r.formula2.has_value());
  EXPECT_FALSE(r.op.has_value());
  EXPECT_FALSE(r.color_scale.has_value());
  EXPECT_FALSE(r.data_bar.has_value());
  EXPECT_FALSE(r.icon_set.has_value());
  EXPECT_FALSE(r.rank.has_value());
  EXPECT_FALSE(r.percent);
  EXPECT_FALSE(r.bottom);
  EXPECT_TRUE(r.above_average);
  EXPECT_FALSE(r.equal_average);
  EXPECT_FALSE(r.std_dev.has_value());
  EXPECT_FALSE(r.text.has_value());
  EXPECT_FALSE(r.time_period.has_value());
}

TEST(CFRule, CellIsBetweenShape) {
  CFRule r;
  r.type = RuleType::CellIs;
  r.op = CellIsOperator::Between;
  r.formula1 = "10";
  r.formula2 = "20";
  r.dxf_id = 4u;
  r.priority = 3;

  ASSERT_TRUE(r.op.has_value());
  EXPECT_EQ(r.op.value(), CellIsOperator::Between);
  ASSERT_TRUE(r.formula1.has_value());
  ASSERT_TRUE(r.formula2.has_value());
  EXPECT_EQ(r.formula1.value(), "10");
  EXPECT_EQ(r.formula2.value(), "20");
  ASSERT_TRUE(r.dxf_id.has_value());
  EXPECT_EQ(r.dxf_id.value(), 4u);
  EXPECT_EQ(r.priority, 3);
}

TEST(CFRule, ColorScaleThreeStopShape) {
  CFRule r;
  r.type = RuleType::ColorScale;
  ColorScaleSpec spec;
  spec.thresholds.push_back({CfvoType::Min, "", true});
  spec.thresholds.push_back({CfvoType::Percentile, "50", true});
  spec.thresholds.push_back({CfvoType::Max, "", true});
  spec.colors.push_back({255, 0, 0, 255});
  spec.colors.push_back({255, 255, 0, 255});
  spec.colors.push_back({0, 255, 0, 255});
  r.color_scale = std::move(spec);

  ASSERT_TRUE(r.color_scale.has_value());
  EXPECT_EQ(r.color_scale->thresholds.size(), 3u);
  EXPECT_EQ(r.color_scale->colors.size(), 3u);
  EXPECT_EQ(r.color_scale->thresholds[1].type, CfvoType::Percentile);
  EXPECT_EQ(r.color_scale->thresholds[1].value, "50");
  EXPECT_EQ(r.color_scale->colors[0], Color({255, 0, 0, 255}));
}

TEST(CFRule, IconSetReverseFlag) {
  CFRule r;
  r.type = RuleType::IconSet;
  IconSetSpec spec;
  spec.name = IconSetName::Five_Arrows;
  spec.reverse = true;
  spec.thresholds.assign(4, CfValueObject{});
  r.icon_set = std::move(spec);

  ASSERT_TRUE(r.icon_set.has_value());
  EXPECT_EQ(r.icon_set->name, IconSetName::Five_Arrows);
  EXPECT_TRUE(r.icon_set->reverse);
  EXPECT_EQ(r.icon_set->thresholds.size(), 4u);
}

TEST(ConditionalFormat, Defaults) {
  ConditionalFormat cf;
  EXPECT_TRUE(cf.sqref.empty());
  EXPECT_TRUE(cf.rules.empty());
  EXPECT_FALSE(cf.pivot_scope);
}

TEST(ConditionalFormat, AccumulateRulesAndRanges) {
  ConditionalFormat cf;
  cf.sqref.push_back({{0, 0}, {9, 0}});   // A1:A10
  cf.sqref.push_back({{4, 3}, {14, 3}});  // D5:D15
  CFRule r;
  r.type = RuleType::CellIs;
  r.op = CellIsOperator::GreaterThan;
  r.formula1 = "100";
  r.priority = 1;
  cf.rules.push_back(std::move(r));

  EXPECT_EQ(cf.sqref.size(), 2u);
  EXPECT_EQ(cf.sqref[0].first.row, 0u);
  EXPECT_EQ(cf.sqref[1].first.row, 4u);
  ASSERT_EQ(cf.rules.size(), 1u);
  EXPECT_EQ(cf.rules[0].type, RuleType::CellIs);
  ASSERT_TRUE(cf.rules[0].op.has_value());
  EXPECT_EQ(cf.rules[0].op.value(), CellIsOperator::GreaterThan);
}

}  // namespace
}  // namespace formulon::cf
