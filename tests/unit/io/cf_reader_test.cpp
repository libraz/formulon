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

TEST(CFReader, FullColumnSqrefAccepted) {
  // A whole-column sqref must not drop the CF block; it is stored at full
  // row extent and classifies as a whole column.
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A:A">
        <cfRule type="expression" priority="1" dxfId="0"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  ASSERT_EQ(cfs.value().size(), 1u);
  const auto& cf = cfs.value()[0];
  ASSERT_EQ(cf.sqref.size(), 1u);
  EXPECT_TRUE(cf.sqref[0].is_full_col());
  EXPECT_EQ(cf.sqref[0].first.col, 0u);
  EXPECT_EQ(cf.sqref[0].last.col, 0u);
  EXPECT_EQ(cf.sqref[0].first.row, 0u);
  EXPECT_EQ(cf.sqref[0].last.row, cf::kCfMaxRows - 1U);
}

TEST(CFReader, MultiColumnFullSqrefAccepted) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A:C">
        <cfRule type="expression" priority="1" dxfId="0"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& cf = cfs.value()[0];
  ASSERT_EQ(cf.sqref.size(), 1u);
  EXPECT_TRUE(cf.sqref[0].is_full_col());
  EXPECT_EQ(cf.sqref[0].first.col, 0u);
  EXPECT_EQ(cf.sqref[0].last.col, 2u);
}

TEST(CFReader, FullRowSqrefAccepted) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="2:2">
        <cfRule type="expression" priority="1" dxfId="0"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& cf = cfs.value()[0];
  ASSERT_EQ(cf.sqref.size(), 1u);
  EXPECT_TRUE(cf.sqref[0].is_full_row());
  EXPECT_EQ(cf.sqref[0].first.row, 1u);
  EXPECT_EQ(cf.sqref[0].last.row, 1u);
  EXPECT_EQ(cf.sqref[0].first.col, 0u);
  EXPECT_EQ(cf.sqref[0].last.col, cf::kCfMaxCols - 1U);
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

TEST(CFReader, DataBarX14OverlayAppliesNegativeAxisAndGradient) {
  // Excel 2010+ always writes the legacy `<dataBar>` for backward
  // compatibility alongside the richer `<x14:dataBar>` extension, cross-
  // referenced by the base `<cfRule id="...">` GUID. The legacy element
  // alone cannot express negative fill / axis colour+position / solid
  // fill — those only exist in the x14 overlay.
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="C1:C10">
        <cfRule type="dataBar" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
          <dataBar minLength="0" maxLength="100">
            <cfvo type="min"/>
            <cfvo type="max"/>
            <color rgb="FF638EC6"/>
          </dataBar>
        </cfRule>
      </conditionalFormatting>
      <extLst>
        <ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
          <x14:conditionalFormattings>
            <x14:conditionalFormatting>
              <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000001}">
                <x14:dataBar minLength="0" maxLength="100" gradient="0" axisPosition="middle">
                  <x14:cfvo type="autoMin"/>
                  <x14:cfvo type="autoMax"/>
                  <x14:negativeFillColor rgb="FFFF0000"/>
                  <x14:negativeBorderColor rgb="FFAA0000"/>
                  <x14:axisColor rgb="FF808080"/>
                  <x14:borderColor rgb="FF000000"/>
                </x14:dataBar>
              </x14:cfRule>
            </x14:conditionalFormatting>
          </x14:conditionalFormattings>
        </ext>
      </extLst>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  ASSERT_TRUE(r.data_bar.has_value());
  // The legacy positive fill is untouched by the overlay.
  EXPECT_EQ(r.data_bar->fill, cf::Color({0x63, 0x8E, 0xC6, 255}));
  EXPECT_EQ(r.data_bar->negative_fill, cf::Color({0xFF, 0x00, 0x00, 255}));
  ASSERT_TRUE(r.data_bar->negative_border.has_value());
  EXPECT_EQ(*r.data_bar->negative_border, cf::Color({0xAA, 0x00, 0x00, 255}));
  EXPECT_EQ(r.data_bar->axis_color, cf::Color({0x80, 0x80, 0x80, 255}));
  EXPECT_EQ(r.data_bar->axis_position, cf::DataBarAxisPosition::Middle);
  EXPECT_FALSE(r.data_bar->gradient);
  ASSERT_TRUE(r.data_bar->border.has_value());
  EXPECT_EQ(*r.data_bar->border, cf::Color({0x00, 0x00, 0x00, 255}));
}

TEST(CFReader, DataBarX14OverlayNegativeSameAsPositiveOverridesExplicitColor) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="C1:C10">
        <cfRule type="dataBar" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
          <dataBar minLength="0" maxLength="100">
            <cfvo type="min"/>
            <cfvo type="max"/>
            <color rgb="FF00B050"/>
          </dataBar>
        </cfRule>
      </conditionalFormatting>
      <extLst>
        <ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}">
          <x14:conditionalFormattings>
            <x14:conditionalFormatting>
              <x14:cfRule type="dataBar" id="{DA7ABA51-AAAA-BBBB-0001-000000000002}">
                <x14:dataBar minLength="0" maxLength="100" negativeBarColorSameAsPositive="1">
                  <x14:cfvo type="autoMin"/>
                  <x14:cfvo type="autoMax"/>
                  <x14:negativeFillColor rgb="FFFF0000"/>
                </x14:dataBar>
              </x14:cfRule>
            </x14:conditionalFormatting>
          </x14:conditionalFormattings>
        </ext>
      </extLst>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  ASSERT_TRUE(r.data_bar.has_value());
  // `negativeBarColorSameAsPositive="1"` wins over the explicit
  // `<x14:negativeFillColor>` sibling, per the CT_DataBar schema.
  EXPECT_EQ(r.data_bar->negative_fill, r.data_bar->fill);
}

TEST(CFReader, DataBarWithoutX14OverlayKeepsLegacyDefaults) {
  // A pre-2010 (or overlay-less) `<dataBar>` has no way to express
  // negative fill / axis position; the reader must not invent one.
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="C1:C10">
        <cfRule type="dataBar" priority="1" id="{DA7ABA51-AAAA-BBBB-0001-000000000003}">
          <dataBar minLength="0" maxLength="100">
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
  ASSERT_TRUE(r.data_bar.has_value());
  EXPECT_EQ(r.data_bar->negative_fill, r.data_bar->fill);
  EXPECT_EQ(r.data_bar->axis_position, cf::DataBarAxisPosition::Automatic);
  EXPECT_TRUE(r.data_bar->gradient);
}

TEST(CFReader, IconSetWithReverseAndPercent) {
  // A 3-icon set carries 3 `<cfvo>` elements in the XML; the first is the
  // floor of the lowest bucket and is dropped, leaving 2 real thresholds
  // in the in-memory model (see `IconSetSpec::thresholds`).
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
  ASSERT_EQ(r.icon_set->thresholds.size(), 2u);
  EXPECT_EQ(r.icon_set->thresholds[0].value, "50");
  EXPECT_TRUE(r.icon_set->thresholds[0].gte);
  EXPECT_EQ(r.icon_set->thresholds[1].value, "100");
  EXPECT_FALSE(r.icon_set->thresholds[1].gte);
}

TEST(CFReader, IconSetFourIconDropsOnlyFloorCfvo) {
  // 4-icon set: 4 `<cfvo>` in XML → 3 real thresholds after the floor is
  // dropped. Regression for the off-by-one that previously kept all N
  // cfvo entries, pushing every bucket boundary index up by one and
  // making the top icon unreachable (index N would exceed [0, N-1]).
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="E1:E10">
        <cfRule type="iconSet" priority="1">
          <iconSet iconSet="4TrafficLights">
            <cfvo type="percent" val="0"/>
            <cfvo type="percent" val="25"/>
            <cfvo type="percent" val="50"/>
            <cfvo type="percent" val="75"/>
          </iconSet>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  ASSERT_TRUE(r.icon_set.has_value());
  ASSERT_EQ(r.icon_set->thresholds.size(), 3u);
  EXPECT_EQ(r.icon_set->thresholds[0].value, "25");
  EXPECT_EQ(r.icon_set->thresholds[1].value, "50");
  EXPECT_EQ(r.icon_set->thresholds[2].value, "75");
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

TEST(CFReader, ExtLstChildIsCapturedVerbatim) {
  // The rule's own <extLst> (typically an `<x14:id>` cross-reference to a
  // richer `<x14:cfRule>` counterpart) does not stop the rest of the rule
  // from parsing, and is captured verbatim rather than silently dropped
  // (see `CFRule::ext_lst_raw`).
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
  ASSERT_TRUE(r.ext_lst_raw.has_value());
  EXPECT_NE(r.ext_lst_raw->find("x14:id"), std::string::npos) << *r.ext_lst_raw;
  EXPECT_NE(r.ext_lst_raw->find("{ABCDEF}"), std::string::npos) << *r.ext_lst_raw;
}

TEST(CFReader, RuleWithoutExtLstLeavesRawFieldEmpty) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1:A10">
        <cfRule type="cellIs" priority="1" operator="equal" dxfId="0">
          <formula>0</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  EXPECT_FALSE(cfs.value()[0].rules[0].ext_lst_raw.has_value());
  EXPECT_FALSE(cfs.value()[0].ext_lst_raw.has_value());
}

TEST(CFReader, DataBarWithX14ExtensionExtLstIsCapturedVerbatim) {
  // A real Excel 2010+ DataBar rule's own `<extLst>` typically carries
  // only the `<x14:id>` cross-reference (the extended negative-fill /
  // axis / gradient properties live in a separate `<x14:cfRule>` at the
  // worksheet level, referenced by this same id -- out of scope for this
  // per-rule capture). The cross-reference itself must still round-trip.
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1:A10">
        <cfRule type="dataBar" priority="1" id="{5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB}">
          <dataBar>
            <cfvo type="min"/>
            <cfvo type="max"/>
            <color rgb="FF638EC6"/>
          </dataBar>
          <extLst>
            <ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"
                 xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
              <x14:id>{5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB}</x14:id>
            </ext>
          </extLst>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  const auto& r = cfs.value()[0].rules[0];
  ASSERT_TRUE(r.data_bar.has_value());
  ASSERT_TRUE(r.ext_lst_raw.has_value());
  EXPECT_NE(r.ext_lst_raw->find("5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB"), std::string::npos) << *r.ext_lst_raw;
}

TEST(CFReader, ConditionalFormattingBlockExtLstIsCapturedVerbatim) {
  // `<conditionalFormatting>`'s own schema-trailing `extLst?` (a sibling
  // of the block's `<cfRule>` children, distinct from any individual
  // rule's own `<extLst>`) is also captured verbatim.
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1:A10">
        <cfRule type="cellIs" priority="1" operator="equal" dxfId="0">
          <formula>0</formula>
        </cfRule>
        <extLst>
          <ext uri="{some-future-extension}">
            <futureThing/>
          </ext>
        </extLst>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  ASSERT_TRUE(cfs.value()[0].ext_lst_raw.has_value());
  EXPECT_NE(cfs.value()[0].ext_lst_raw->find("futureThing"), std::string::npos) << *cfs.value()[0].ext_lst_raw;
}

TEST(CFReader, AbsoluteMarkersInSqrefAreAccepted) {
  // `$A$1:$A$10` is valid OOXML; the absolute markers must be stripped,
  // not rejected. Previously a single `$` token failed the whole load.
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="$A$1:$A$10 $C$5">
        <cfRule type="cellIs" priority="1" operator="greaterThan" dxfId="0">
          <formula>10</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  ASSERT_EQ(cfs.value().size(), 1u);
  const auto& cf = cfs.value()[0];
  ASSERT_EQ(cf.sqref.size(), 2u);
  EXPECT_EQ(cf.sqref[0].first.row, 0u);
  EXPECT_EQ(cf.sqref[0].first.col, 0u);
  EXPECT_EQ(cf.sqref[0].last.row, 9u);
  EXPECT_EQ(cf.sqref[0].last.col, 0u);
  EXPECT_EQ(cf.sqref[1].first.row, 4u);
  EXPECT_EQ(cf.sqref[1].first.col, 2u);
  EXPECT_EQ(cf.sqref[1].first, cf.sqref[1].last);  // Single-cell.
  ASSERT_EQ(cf.rules.size(), 1u);
}

TEST(CFReader, MissingSqrefAttributeSkipsBlock) {
  // A CF block is a presentation overlay; a missing sqref skips just that
  // block, leaving the load successful rather than aborting the workbook.
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting>
        <cfRule type="cellIs" priority="1" operator="equal" dxfId="0">
          <formula>0</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  EXPECT_TRUE(cfs.value().empty());
}

TEST(CFReader, EmptySqrefAttributeSkipsBlock) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="">
        <cfRule type="cellIs" priority="1"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  EXPECT_TRUE(cfs.value().empty());
}

TEST(CFReader, UnparseableSqrefTokenSkipsBlock) {
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="not_a_range">
        <cfRule type="cellIs" priority="1"/>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  EXPECT_TRUE(cfs.value().empty());
}

TEST(CFReader, MalformedBlockSkippedButValidBlockKept) {
  // One malformed block must not take down the sibling valid block; the
  // load continues and only the good block survives.
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="not_a_range">
        <cfRule type="cellIs" priority="1"/>
      </conditionalFormatting>
      <conditionalFormatting sqref="$B$2:$B$5">
        <cfRule type="cellIs" priority="2" operator="lessThan" dxfId="0">
          <formula>3</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  ASSERT_EQ(cfs.value().size(), 1u);
  const auto& cf = cfs.value()[0];
  ASSERT_EQ(cf.sqref.size(), 1u);
  EXPECT_EQ(cf.sqref[0].first.row, 1u);
  EXPECT_EQ(cf.sqref[0].first.col, 1u);
  EXPECT_EQ(cf.sqref[0].last.row, 4u);
  ASSERT_EQ(cf.rules.size(), 1u);
  EXPECT_EQ(cf.rules[0].priority, 2);
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

TEST(CFReader, OutOfRangePriorityFallsBackToDefault) {
  // Regression: priority="9999999999" (> INT32_MAX) used to slip through
  // ParseI32Attr because errno was implementation-defined; now the value
  // is bounds-checked against INT32_MIN..INT32_MAX explicitly.
  pugi::xml_document doc = Load(R"(
    <worksheet>
      <conditionalFormatting sqref="A1">
        <cfRule type="expression" priority="9999999999" dxfId="0">
          <formula>1</formula>
        </cfRule>
      </conditionalFormatting>
    </worksheet>)");
  auto cfs = read_conditional_formats(doc.child("worksheet"));
  ASSERT_TRUE(cfs);
  ASSERT_EQ(cfs.value().size(), 1u);
  ASSERT_EQ(cfs.value()[0].rules.size(), 1u);
  // priority default is 1; the value must NOT be clamped to INT32_MAX
  // and must NOT be wrapped to a negative.
  EXPECT_EQ(cfs.value()[0].rules[0].priority, 1);
}

}  // namespace
}  // namespace formulon::io
