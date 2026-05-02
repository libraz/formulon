// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the OOXML conditional-formatting reader.

#include "io/cf_reader.h"

#include <cstring>
#include <string>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {
namespace {

pugi::xml_document Load(const char* xml) {
  pugi::xml_document doc;
  pugi::xml_parse_result rc = doc.load_string(xml);
  EXPECT_TRUE(rc) << rc.description();
  return doc;
}

TEST(CFReader, EmptyWorksheetReturnsNoBlocks) {
  pugi::xml_document doc = Load(R"(<worksheet><sheetData/></worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  EXPECT_TRUE(cfs.value().empty());
}

TEST(CFReader, SingleCellIsRule) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1:A10">
        <cfRule type="cellIs" priority="1" operator="greaterThan" dxfId="3" stopIfTrue="1">
          <formula>10</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");

  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  ASSERT_EQ(cfs.value().size(), 1u);
  const auto& cf = cfs.value()[0];

  ASSERT_EQ(cf.sqref.size(), 1u);
  EXPECT_EQ(cf.sqref[0].first.row, 0u);
  EXPECT_EQ(cf.sqref[0].first.col, 0u);
  EXPECT_EQ(cf.sqref[0].last.row, 9u);
  EXPECT_EQ(cf.sqref[0].last.col, 0u);
  EXPECT_FALSE(cf.pivot_scope);

  ASSERT_EQ(cf.rules.size(), 1u);
  const auto& r = cf.rules[0];
  EXPECT_EQ(r.type, cf::RuleType::CellIs);
  EXPECT_EQ(r.priority, 1);
  EXPECT_TRUE(r.stop_if_true);
  ASSERT_TRUE(r.dxf_id.has_value());
  EXPECT_EQ(r.dxf_id.value(), 3u);
  ASSERT_TRUE(r.op.has_value());
  EXPECT_EQ(r.op.value(), cf::CellIsOperator::GreaterThan);
  ASSERT_TRUE(r.formula1.has_value());
  EXPECT_EQ(r.formula1.value(), "10");
  EXPECT_FALSE(r.formula2.has_value());
}

TEST(CFReader, CellIsBetweenCarriesTwoFormulas) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="B2:B5">
        <cfRule type="cellIs" priority="2" operator="between" dxfId="0">
          <formula>10</formula>
          <formula>20</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  ASSERT_EQ(cfs.value().size(), 1u);
  const auto& r = cfs.value()[0].rules[0];
  EXPECT_EQ(r.op.value(), cf::CellIsOperator::Between);
  ASSERT_TRUE(r.formula1.has_value());
  ASSERT_TRUE(r.formula2.has_value());
  EXPECT_EQ(r.formula1.value(), "10");
  EXPECT_EQ(r.formula2.value(), "20");
}

TEST(CFReader, ExpressionRuleStripsLeadingEquals) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1">
        <cfRule type="expression" priority="1" dxfId="0">
          <formula>=A1&gt;100</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  EXPECT_EQ(r.type, cf::RuleType::Expression);
  ASSERT_TRUE(r.formula1.has_value());
  EXPECT_EQ(r.formula1.value(), "A1>100");
}

TEST(CFReader, MultipleRangesInSqref) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1:A10 D5:D15 G1">
        <cfRule type="expression" priority="1" dxfId="0"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& cf = cfs.value()[0];
  ASSERT_EQ(cf.sqref.size(), 3u);
  EXPECT_EQ(cf.sqref[0].first.row, 0u);
  EXPECT_EQ(cf.sqref[0].last.row, 9u);
  EXPECT_EQ(cf.sqref[1].first.row, 4u);
  EXPECT_EQ(cf.sqref[1].first.col, 3u);
  EXPECT_EQ(cf.sqref[2].first.row, 0u);
  EXPECT_EQ(cf.sqref[2].first.col, 6u);
  EXPECT_EQ(cf.sqref[2].first, cf.sqref[2].last);  // Single-cell.
}

TEST(CFReader, ColorScaleThreeStop) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1:A10">
        <cfRule type="colorScale" priority="1">
          <colorScale>
            <cfvo type="min"/>
            <cfvo type="percentile" val="50"/>
            <cfvo type="max"/>
            <color rgb="FFFF0000"/>
            <color rgb="FFFFFF00"/>
            <color rgb="FF00FF00"/>
          </colorScale>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  EXPECT_EQ(r.type, cf::RuleType::ColorScale);
  ASSERT_TRUE(r.color_scale.has_value());
  EXPECT_EQ(r.color_scale->thresholds.size(), 3u);
  EXPECT_EQ(r.color_scale->colors.size(), 3u);
  EXPECT_EQ(r.color_scale->thresholds[0].type, cf::CfvoType::Min);
  EXPECT_EQ(r.color_scale->thresholds[1].type, cf::CfvoType::Percentile);
  EXPECT_EQ(r.color_scale->thresholds[1].value, "50");
  EXPECT_EQ(r.color_scale->colors[0], cf::Color({255, 0, 0, 255}));
  EXPECT_EQ(r.color_scale->colors[1], cf::Color({255, 255, 0, 255}));
  EXPECT_EQ(r.color_scale->colors[2], cf::Color({0, 255, 0, 255}));
}

TEST(CFReader, DataBarBasic) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="C1:C10">
        <cfRule type="dataBar" priority="1">
          <dataBar minLength="0" maxLength="100" showValue="0">
            <cfvo type="min"/>
            <cfvo type="max"/>
            <color rgb="FF638EC6"/>
          </dataBar>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  EXPECT_EQ(r.type, cf::RuleType::DataBar);
  ASSERT_TRUE(r.data_bar.has_value());
  EXPECT_EQ(r.data_bar->min.type, cf::CfvoType::Min);
  EXPECT_EQ(r.data_bar->max.type, cf::CfvoType::Max);
  EXPECT_EQ(r.data_bar->fill, cf::Color({0x63, 0x8E, 0xC6, 255}));
  EXPECT_EQ(r.data_bar->min_length_pct, 0u);
  EXPECT_EQ(r.data_bar->max_length_pct, 100u);
  EXPECT_FALSE(r.data_bar->show_value);
}

TEST(CFReader, IconSetWithReverseAndPercent) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="D1:D10">
        <cfRule type="iconSet" priority="1">
          <iconSet iconSet="3TrafficLights2" reverse="1" percent="0">
            <cfvo type="num" val="0" gte="1"/>
            <cfvo type="num" val="50" gte="1"/>
            <cfvo type="num" val="100" gte="0"/>
          </iconSet>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  EXPECT_EQ(r.type, cf::RuleType::IconSet);
  ASSERT_TRUE(r.icon_set.has_value());
  EXPECT_EQ(r.icon_set->name, cf::IconSetName::Three_TrafficLights2);
  EXPECT_TRUE(r.icon_set->reverse);
  EXPECT_FALSE(r.icon_set->percent);
  EXPECT_EQ(r.icon_set->thresholds.size(), 3u);
  EXPECT_TRUE(r.icon_set->thresholds[1].gte);
  EXPECT_FALSE(r.icon_set->thresholds[2].gte);
}

TEST(CFReader, Top10AttributesRoundTrip) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1:A100">
        <cfRule type="top10" priority="1" rank="5" percent="1" bottom="1" dxfId="0"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  EXPECT_EQ(r.type, cf::RuleType::Top10);
  ASSERT_TRUE(r.rank.has_value());
  EXPECT_EQ(r.rank.value(), 5);
  EXPECT_TRUE(r.percent);
  EXPECT_TRUE(r.bottom);
}

TEST(CFReader, AboveAverageStdDev) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="B1:B50">
        <cfRule type="aboveAverage" priority="1" aboveAverage="0" equalAverage="1" stdDev="2" dxfId="0"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  EXPECT_EQ(r.type, cf::RuleType::AboveAverage);
  EXPECT_FALSE(r.above_average);
  EXPECT_TRUE(r.equal_average);
  ASSERT_TRUE(r.std_dev.has_value());
  EXPECT_EQ(r.std_dev.value(), 2.0);
}

TEST(CFReader, ContainsTextCarriesText) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1:A10">
        <cfRule type="containsText" priority="1" operator="containsText" text="error" dxfId="0"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  EXPECT_EQ(r.type, cf::RuleType::ContainsText);
  ASSERT_TRUE(r.text.has_value());
  EXPECT_EQ(r.text.value(), "error");
}

TEST(CFReader, TimePeriodEnumDecoded) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1">
        <cfRule type="timePeriod" priority="1" timePeriod="lastWeek" dxfId="0"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  EXPECT_EQ(r.type, cf::RuleType::TimePeriod);
  ASSERT_TRUE(r.time_period.has_value());
  EXPECT_EQ(r.time_period.value(), cf::TimePeriod::LastWeek);
}

TEST(CFReader, MultipleBlocksAndRulesPreserveOrder) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1:A10">
        <cfRule type="cellIs" priority="1" operator="greaterThan" dxfId="0">
          <formula>50</formula>
        </cfRule>
        <cfRule type="cellIs" priority="2" operator="lessThan" dxfId="1">
          <formula>10</formula>
        </cfRule>
      </conditionalFormatting>
      <conditionalFormatting sqref="B1:B10" pivot="1">
        <cfRule type="expression" priority="3" dxfId="2">
          <formula>$A1=0</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  ASSERT_EQ(cfs.value().size(), 2u);
  EXPECT_EQ(cfs.value()[0].rules.size(), 2u);
  EXPECT_EQ(cfs.value()[0].rules[0].priority, 1);
  EXPECT_EQ(cfs.value()[0].rules[1].priority, 2);
  EXPECT_FALSE(cfs.value()[0].pivot_scope);

  EXPECT_TRUE(cfs.value()[1].pivot_scope);
  ASSERT_EQ(cfs.value()[1].rules.size(), 1u);
  EXPECT_EQ(cfs.value()[1].rules[0].priority, 3);
  ASSERT_TRUE(cfs.value()[1].rules[0].formula1.has_value());
  EXPECT_EQ(cfs.value()[1].rules[0].formula1.value(), "$A1=0");
}

TEST(CFReader, ExtLstChildIsIgnored) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1:A10">
        <cfRule type="cellIs" priority="1" operator="equal" dxfId="0">
          <formula>0</formula>
          <extLst>
            <ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}">
              <x14:id>{ABCDEF}</x14:id>
            </ext>
          </extLst>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  ASSERT_TRUE(r.formula1.has_value());
  EXPECT_EQ(r.formula1.value(), "0");
}

TEST(CFReader, MissingSqrefAttributeFails) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting>
        <cfRule type="cellIs" priority="1" operator="equal" dxfId="0">
          <formula>0</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_FALSE(cfs);
  EXPECT_EQ(cfs.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CFReader, EmptySqrefAttributeFails) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="">
        <cfRule type="cellIs" priority="1"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_FALSE(cfs);
  EXPECT_EQ(cfs.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CFReader, UnparseableSqrefTokenFails) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="not_a_range">
        <cfRule type="cellIs" priority="1"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_FALSE(cfs);
  EXPECT_EQ(cfs.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(CFReader, UnknownRuleTypeFoldsToExpression) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1">
        <cfRule type="someFutureExcelType" priority="1" dxfId="0">
          <formula>A1=0</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  EXPECT_EQ(r.type, cf::RuleType::Expression);
  ASSERT_TRUE(r.formula1.has_value());
  EXPECT_EQ(r.formula1.value(), "A1=0");
}

TEST(CFReader, RgbWithoutAlphaIsAccepted) {
  // Some Excel exports drop the alpha byte for fully-opaque colours.
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1">
        <cfRule type="colorScale" priority="1">
          <colorScale>
            <cfvo type="min"/>
            <cfvo type="max"/>
            <color rgb="FF0000"/>
            <color rgb="00FF00"/>
          </colorScale>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  ASSERT_TRUE(r.color_scale.has_value());
  EXPECT_EQ(r.color_scale->colors[0], cf::Color({255, 0, 0, 255}));
  EXPECT_EQ(r.color_scale->colors[1], cf::Color({0, 255, 0, 255}));
}

TEST(CFReader, IdAttributeRoundTrips) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1">
        <cfRule type="cellIs" priority="1" operator="equal" dxfId="0" id="{12345678-90AB-CDEF-1234-567890ABCDEF}">
          <formula>0</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  EXPECT_EQ(cfs.value()[0].rules[0].id, "{12345678-90AB-CDEF-1234-567890ABCDEF}");
}

}  // namespace
}  // namespace formulon::io
