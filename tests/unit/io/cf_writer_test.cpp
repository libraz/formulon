// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Round-trip unit tests for the OOXML conditional-formatting writer.
// Each test builds a `cf::ConditionalFormat` model in memory, emits the
// XML via `write_conditional_formattings`, wraps the result in a
// `<worksheet>` shell, parses it via `read_conditional_formats`, and
// asserts the parsed list matches the input bit-for-bit.

#include "io/cf_writer.h"

#include <string>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/cf_reader.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {

std::vector<cf::ConditionalFormat> RoundTrip(const std::vector<cf::ConditionalFormat>& input) {
  std::string xml = "<worksheet>";
  xml.append(write_conditional_formattings(input));
  xml.append("</worksheet>");
  pugi::xml_document doc;
  pugi::xml_parse_result rc = doc.load_string(xml.c_str());
  EXPECT_TRUE(rc) << rc.description();
  auto out_or = read_conditional_formats(doc.child("worksheet"));
  EXPECT_TRUE(static_cast<bool>(out_or));
  return std::move(out_or.value());
}

TEST(CFWriter, EmptyListProducesEmptyString) {
  std::vector<cf::ConditionalFormat> input;
  EXPECT_TRUE(write_conditional_formattings(input).empty());
}

TEST(CFWriter, CellIsRuleRoundTrips) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::CellIs;
  r.priority = 1;
  r.stop_if_true = true;
  r.dxf_id = 3u;
  r.op = cf::CellIsOperator::GreaterThan;
  r.formula1 = "10";
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  ASSERT_EQ(out.size(), 1u);
  ASSERT_EQ(out[0].sqref.size(), 1u);
  EXPECT_EQ(out[0].sqref[0], cf::CFCellRange({{0, 0}, {9, 0}}));
  ASSERT_EQ(out[0].rules.size(), 1u);
  EXPECT_EQ(out[0].rules[0].type, cf::RuleType::CellIs);
  EXPECT_EQ(out[0].rules[0].priority, 1);
  EXPECT_TRUE(out[0].rules[0].stop_if_true);
  EXPECT_EQ(out[0].rules[0].dxf_id.value(), 3u);
  EXPECT_EQ(out[0].rules[0].op.value(), cf::CellIsOperator::GreaterThan);
  EXPECT_EQ(out[0].rules[0].formula1.value(), "10");
}

TEST(CFWriter, CellIsBetweenTwoFormulasRoundTrips) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{1, 1}, {4, 1}});
  cf::CFRule r;
  r.type = cf::RuleType::CellIs;
  r.op = cf::CellIsOperator::Between;
  r.formula1 = "10";
  r.formula2 = "20";
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  EXPECT_EQ(out[0].rules[0].formula1.value(), "10");
  EXPECT_EQ(out[0].rules[0].formula2.value(), "20");
  EXPECT_EQ(out[0].rules[0].op.value(), cf::CellIsOperator::Between);
}

TEST(CFWriter, ExpressionFormulaIsXmlEscaped) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {0, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::Expression;
  r.formula1 = "A1>5 & B1<10";  // literal `&` must escape
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  EXPECT_EQ(out[0].rules[0].formula1.value(), "A1>5 & B1<10");
}

TEST(CFWriter, SqrefSingleCellEmitsShortForm) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {0, 0}});
  cf.sqref.push_back({{4, 3}, {14, 3}});
  cf::CFRule r;
  r.type = cf::RuleType::Expression;
  r.formula1 = "A1=0";
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  std::string xml = write_conditional_formattings({cf});
  // Single-cell range emits as plain "A1", not "A1:A1".
  EXPECT_NE(xml.find("sqref=\"A1 D5:D15\""), std::string::npos) << xml;

  auto out = RoundTrip({cf});
  ASSERT_EQ(out[0].sqref.size(), 2u);
  EXPECT_EQ(out[0].sqref[0], cf::CFCellRange({{0, 0}, {0, 0}}));
  EXPECT_EQ(out[0].sqref[1], cf::CFCellRange({{4, 3}, {14, 3}}));
}

TEST(CFWriter, ColorScaleThreeStopRoundTrips) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::ColorScale;
  cf::ColorScaleSpec s;
  s.thresholds.push_back({cf::CfvoType::Min, "", true});
  s.thresholds.push_back({cf::CfvoType::Percentile, "50", true});
  s.thresholds.push_back({cf::CfvoType::Max, "", true});
  s.colors.push_back({255, 0, 0, 255});
  s.colors.push_back({255, 255, 0, 255});
  s.colors.push_back({0, 255, 0, 255});
  r.color_scale = std::move(s);
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  ASSERT_TRUE(out[0].rules[0].color_scale.has_value());
  EXPECT_EQ(out[0].rules[0].color_scale->thresholds.size(), 3u);
  EXPECT_EQ(out[0].rules[0].color_scale->thresholds[1].type, cf::CfvoType::Percentile);
  EXPECT_EQ(out[0].rules[0].color_scale->thresholds[1].value, "50");
  EXPECT_EQ(out[0].rules[0].color_scale->colors[0], cf::Color({255, 0, 0, 255}));
  EXPECT_EQ(out[0].rules[0].color_scale->colors[1], cf::Color({255, 255, 0, 255}));
  EXPECT_EQ(out[0].rules[0].color_scale->colors[2], cf::Color({0, 255, 0, 255}));
}

TEST(CFWriter, DataBarRoundTrips) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 2}, {9, 2}});
  cf::CFRule r;
  r.type = cf::RuleType::DataBar;
  cf::DataBarSpec d;
  d.min = {cf::CfvoType::Min, "", true};
  d.max = {cf::CfvoType::Max, "", true};
  d.fill = {0x63, 0x8E, 0xC6, 255};
  d.min_length_pct = 0;
  d.max_length_pct = 100;
  d.show_value = false;
  r.data_bar = std::move(d);
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  ASSERT_TRUE(out[0].rules[0].data_bar.has_value());
  EXPECT_EQ(out[0].rules[0].data_bar->fill, cf::Color({0x63, 0x8E, 0xC6, 255}));
  EXPECT_EQ(out[0].rules[0].data_bar->min_length_pct, 0u);
  EXPECT_EQ(out[0].rules[0].data_bar->max_length_pct, 100u);
  EXPECT_FALSE(out[0].rules[0].data_bar->show_value);
}

TEST(CFWriter, IconSetReverseAndPercentRoundTrip) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 3}, {9, 3}});
  cf::CFRule r;
  r.type = cf::RuleType::IconSet;
  cf::IconSetSpec i;
  i.name = cf::IconSetName::Three_TrafficLights2;
  i.reverse = true;
  i.percent = false;
  i.thresholds.push_back({cf::CfvoType::Number, "0", true});
  i.thresholds.push_back({cf::CfvoType::Number, "50", true});
  i.thresholds.push_back({cf::CfvoType::Number, "100", false});  // gte=0
  r.icon_set = std::move(i);
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  ASSERT_TRUE(out[0].rules[0].icon_set.has_value());
  EXPECT_EQ(out[0].rules[0].icon_set->name, cf::IconSetName::Three_TrafficLights2);
  EXPECT_TRUE(out[0].rules[0].icon_set->reverse);
  EXPECT_FALSE(out[0].rules[0].icon_set->percent);
  EXPECT_EQ(out[0].rules[0].icon_set->thresholds.size(), 3u);
  EXPECT_TRUE(out[0].rules[0].icon_set->thresholds[1].gte);
  EXPECT_FALSE(out[0].rules[0].icon_set->thresholds[2].gte);
}

TEST(CFWriter, Top10RoundTrip) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {99, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::Top10;
  r.rank = 5;
  r.percent = true;
  r.bottom = true;
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  EXPECT_EQ(out[0].rules[0].type, cf::RuleType::Top10);
  EXPECT_EQ(out[0].rules[0].rank.value(), 5);
  EXPECT_TRUE(out[0].rules[0].percent);
  EXPECT_TRUE(out[0].rules[0].bottom);
}

TEST(CFWriter, AboveAverageRoundTrip) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 1}, {49, 1}});
  cf::CFRule r;
  r.type = cf::RuleType::AboveAverage;
  r.above_average = false;
  r.equal_average = true;
  r.std_dev = 2.0;
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  EXPECT_FALSE(out[0].rules[0].above_average);
  EXPECT_TRUE(out[0].rules[0].equal_average);
  ASSERT_TRUE(out[0].rules[0].std_dev.has_value());
  EXPECT_EQ(out[0].rules[0].std_dev.value(), 2.0);
}

TEST(CFWriter, ContainsTextRoundTrip) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::ContainsText;
  r.text = "error";
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  EXPECT_EQ(out[0].rules[0].type, cf::RuleType::ContainsText);
  EXPECT_EQ(out[0].rules[0].text.value(), "error");
}

TEST(CFWriter, TimePeriodRoundTrip) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {0, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::TimePeriod;
  r.time_period = cf::TimePeriod::LastWeek;
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  EXPECT_EQ(out[0].rules[0].type, cf::RuleType::TimePeriod);
  EXPECT_EQ(out[0].rules[0].time_period.value(), cf::TimePeriod::LastWeek);
}

TEST(CFWriter, IdAttributeRoundTrips) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {0, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::CellIs;
  r.op = cf::CellIsOperator::Equal;
  r.formula1 = "0";
  r.id = "{12345678-90AB-CDEF-1234-567890ABCDEF}";
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  EXPECT_EQ(out[0].rules[0].id, "{12345678-90AB-CDEF-1234-567890ABCDEF}");
}

TEST(CFWriter, PivotScopeRoundTrips) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {9, 0}});
  cf.pivot_scope = true;
  cf::CFRule r;
  r.type = cf::RuleType::Expression;
  r.formula1 = "A1=0";
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  auto out = RoundTrip({cf});
  EXPECT_TRUE(out[0].pivot_scope);
}

TEST(CFWriter, MultipleBlocksAndRulesPreserveOrder) {
  cf::ConditionalFormat a{};
  a.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r1;
  r1.type = cf::RuleType::CellIs;
  r1.op = cf::CellIsOperator::GreaterThan;
  r1.formula1 = "50";
  r1.priority = 1;
  r1.dxf_id = 0u;
  cf::CFRule r2;
  r2.type = cf::RuleType::CellIs;
  r2.op = cf::CellIsOperator::LessThan;
  r2.formula1 = "10";
  r2.priority = 2;
  r2.dxf_id = 1u;
  a.rules.push_back(std::move(r1));
  a.rules.push_back(std::move(r2));

  cf::ConditionalFormat b{};
  b.sqref.push_back({{0, 1}, {9, 1}});
  cf::CFRule r3;
  r3.type = cf::RuleType::Expression;
  r3.formula1 = "$A1=0";
  r3.priority = 3;
  r3.dxf_id = 2u;
  b.rules.push_back(std::move(r3));

  auto out = RoundTrip({a, b});
  ASSERT_EQ(out.size(), 2u);
  ASSERT_EQ(out[0].rules.size(), 2u);
  EXPECT_EQ(out[0].rules[0].priority, 1);
  EXPECT_EQ(out[0].rules[1].priority, 2);
  ASSERT_EQ(out[1].rules.size(), 1u);
  EXPECT_EQ(out[1].rules[0].priority, 3);
  EXPECT_EQ(out[1].rules[0].formula1.value(), "$A1=0");
}

TEST(CFWriter, AllRuleTypesRoundTrip) {
  // One block per rule type; ensures the type-string mapping is
  // correct for every member of the RuleType enum. Some types cannot
  // be tested with `read_conditional_formats` round-trip because they
  // have no payload (containsBlanks / notContainsBlanks /
  // containsErrors / notContainsErrors / duplicateValues /
  // uniqueValues); for those we only verify the substring is present.
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {0, 0}});

  // Run one type at a time.
  for (auto type :
       {cf::RuleType::Expression, cf::RuleType::CellIs, cf::RuleType::ColorScale, cf::RuleType::DataBar,
        cf::RuleType::IconSet, cf::RuleType::Top10, cf::RuleType::AboveAverage, cf::RuleType::ContainsText,
        cf::RuleType::NotContainsText, cf::RuleType::BeginsWith, cf::RuleType::EndsWith, cf::RuleType::ContainsBlanks,
        cf::RuleType::NotContainsBlanks, cf::RuleType::ContainsErrors, cf::RuleType::NotContainsErrors,
        cf::RuleType::TimePeriod, cf::RuleType::DuplicateValues, cf::RuleType::UniqueValues}) {
    cf.rules.clear();
    cf::CFRule r;
    r.type = type;
    r.dxf_id = 0u;
    if (type == cf::RuleType::CellIs) {
      r.op = cf::CellIsOperator::Equal;
      r.formula1 = "0";
    } else if (type == cf::RuleType::TimePeriod) {
      r.time_period = cf::TimePeriod::Today;
    } else if (type == cf::RuleType::ContainsText || type == cf::RuleType::NotContainsText ||
               type == cf::RuleType::BeginsWith || type == cf::RuleType::EndsWith) {
      r.text = "x";
    }
    cf.rules.push_back(std::move(r));

    auto out = RoundTrip({cf});
    ASSERT_EQ(out.size(), 1u) << "type=" << static_cast<int>(type);
    ASSERT_EQ(out[0].rules.size(), 1u) << "type=" << static_cast<int>(type);
    EXPECT_EQ(out[0].rules[0].type, type) << "type=" << static_cast<int>(type);
  }
}

}  // namespace
}  // namespace formulon::io
