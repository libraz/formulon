//
// Round-trip unit tests for the OOXML conditional-formatting writer.
// Each test builds a `cf::ConditionalFormat` model in memory, emits the
// XML via `write_conditional_formattings`, wraps the result in a
// `<worksheet>` shell, parses it via `read_conditional_formats`, and
// asserts the parsed list matches the input bit-for-bit.

#include "io/cf_writer.h"

#include <cstddef>
#include <string>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/cf_overlay.h"
#include "io/cf_reader.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {

/// Stands in for the package's `<dxfs>` record count. Large enough that
/// every `dxf_id` these tests set resolves, so only the test that targets
/// the bound sees it engage.
constexpr std::size_t kDxfCount = 8;

std::vector<cf::ConditionalFormat> RoundTrip(const std::vector<cf::ConditionalFormat>& input) {
  std::string xml = "<worksheet>";
  xml.append(write_conditional_formattings(input, kDxfCount));
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
  EXPECT_TRUE(write_conditional_formattings(input, kDxfCount).empty());
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

  std::string xml = write_conditional_formattings({cf}, kDxfCount);
  // Single-cell range emits as plain "A1", not "A1:A1".
  EXPECT_NE(xml.find("sqref=\"A1 D5:D15\""), std::string::npos) << xml;

  auto out = RoundTrip({cf});
  ASSERT_EQ(out[0].sqref.size(), 2u);
  EXPECT_EQ(out[0].sqref[0], cf::CFCellRange({{0, 0}, {0, 0}}));
  EXPECT_EQ(out[0].sqref[1], cf::CFCellRange({{4, 3}, {14, 3}}));
}

TEST(CFWriter, FullColumnSqrefEmitsCompactFormAndRoundTrips) {
  cf::ConditionalFormat cf{};
  // Whole column A (full row extent), and whole columns B:C.
  cf.sqref.push_back({{0, 0}, {cf::kCfMaxRows - 1U, 0}});
  cf.sqref.push_back({{0, 1}, {cf::kCfMaxRows - 1U, 2}});
  cf::CFRule r;
  r.type = cf::RuleType::Expression;
  r.formula1 = "A1=0";
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  std::string xml = write_conditional_formattings({cf}, kDxfCount);
  EXPECT_NE(xml.find("sqref=\"A:A B:C\""), std::string::npos) << xml;

  auto out = RoundTrip({cf});
  ASSERT_EQ(out[0].sqref.size(), 2u);
  EXPECT_TRUE(out[0].sqref[0].is_full_col());
  EXPECT_EQ(out[0].sqref[0], cf.sqref[0]);
  EXPECT_EQ(out[0].sqref[1], cf.sqref[1]);
}

TEST(CFWriter, FullRowSqrefEmitsCompactFormAndRoundTrips) {
  cf::ConditionalFormat cf{};
  // Whole row 3 (full column extent).
  cf.sqref.push_back({{2, 0}, {2, cf::kCfMaxCols - 1U}});
  cf::CFRule r;
  r.type = cf::RuleType::Expression;
  r.formula1 = "A1=0";
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  std::string xml = write_conditional_formattings({cf}, kDxfCount);
  EXPECT_NE(xml.find("sqref=\"3:3\""), std::string::npos) << xml;

  auto out = RoundTrip({cf});
  ASSERT_EQ(out[0].sqref.size(), 1u);
  EXPECT_TRUE(out[0].sqref[0].is_full_row());
  EXPECT_EQ(out[0].sqref[0], cf.sqref[0]);
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
  // Three_TrafficLights2 is a 3-icon set: the model carries N-1 = 2 real
  // thresholds; the writer must re-synthesize the floor cfvo so the
  // emitted XML carries N = 3 `<cfvo>` elements (schema-valid, matches
  // what Excel itself emits).
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 3}, {9, 3}});
  cf::CFRule r;
  r.type = cf::RuleType::IconSet;
  cf::IconSetSpec i;
  i.name = cf::IconSetName::Three_TrafficLights2;
  i.reverse = true;
  i.percent = false;
  i.thresholds.push_back({cf::CfvoType::Number, "50", true});
  i.thresholds.push_back({cf::CfvoType::Number, "100", false});  // gte=0
  r.icon_set = std::move(i);
  cf.rules.push_back(std::move(r));

  std::string xml = write_conditional_formattings({cf}, kDxfCount);
  // Floor cfvo (re-synthesized) plus the 2 real thresholds: 3 `<cfvo>`
  // elements for the 3-icon set. `gte` is omitted when true (the default)
  // and emitted as `gte="0"` only when false.
  EXPECT_NE(xml.find(R"(<cfvo type="percent" val="0"/>)"), std::string::npos) << xml;
  EXPECT_NE(xml.find(R"(<cfvo type="num" val="50"/>)"), std::string::npos) << xml;
  EXPECT_NE(xml.find(R"(<cfvo type="num" val="100" gte="0"/>)"), std::string::npos) << xml;

  auto out = RoundTrip({cf});
  ASSERT_TRUE(out[0].rules[0].icon_set.has_value());
  EXPECT_EQ(out[0].rules[0].icon_set->name, cf::IconSetName::Three_TrafficLights2);
  EXPECT_TRUE(out[0].rules[0].icon_set->reverse);
  EXPECT_FALSE(out[0].rules[0].icon_set->percent);
  ASSERT_EQ(out[0].rules[0].icon_set->thresholds.size(), 2u);
  EXPECT_EQ(out[0].rules[0].icon_set->thresholds[0].value, "50");
  EXPECT_TRUE(out[0].rules[0].icon_set->thresholds[0].gte);
  EXPECT_EQ(out[0].rules[0].icon_set->thresholds[1].value, "100");
  EXPECT_FALSE(out[0].rules[0].icon_set->thresholds[1].gte);
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

TEST(CFWriter, ContainsTextWithNewlineIsAttributeEscaped) {
  // `text=` is an attribute value; a literal embedded newline would be
  // normalised away by any conforming XML parser on reload. The writer
  // must emit a character reference instead of the raw byte, and the
  // round trip through the reader must still recover the exact string.
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::ContainsText;
  r.text = "line one\nline two";
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  const std::string xml = write_conditional_formattings({cf}, kDxfCount);
  EXPECT_NE(xml.find("text=\"line one&#10;line two\""), std::string::npos) << xml;

  auto out = RoundTrip({cf});
  EXPECT_EQ(out[0].rules[0].text.value(), "line one\nline two");
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

TEST(CFWriter, IdOnARuleWithNoExtensionPayloadIsNotEmitted) {
  // `CT_CfRule` has no `id` attribute and a rule with nothing in the x14
  // extension has nothing to link to, so the id stays an in-memory
  // handle. Emitting it would put a non-schema attribute in the file
  // that Excel strips on its next save anyway.
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {0, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::CellIs;
  r.op = cf::CellIsOperator::Equal;
  r.formula1 = "0";
  r.id = "{12345678-90AB-CDEF-1234-567890ABCDEF}";
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));

  const std::string xml = write_conditional_formattings({cf}, kDxfCount);
  EXPECT_EQ(xml.find("12345678-90AB-CDEF-1234-567890ABCDEF"), std::string::npos) << xml;
  EXPECT_TRUE(RoundTrip({cf})[0].rules[0].id.empty());
}

TEST(CFWriter, RuleExtLstRoundTrips) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::DataBar;
  cf::DataBarSpec d;
  d.min = {cf::CfvoType::Min, "", true};
  d.max = {cf::CfvoType::Max, "", true};
  d.fill = {0x63, 0x8E, 0xC6, 255};
  r.data_bar = std::move(d);
  r.id = "{5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB}";
  r.ext_lst_raw =
      "<extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
      "<x14:id>{5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB}</x14:id></ext></extLst>";
  cf.rules.push_back(std::move(r));

  std::string xml = write_conditional_formattings({cf}, kDxfCount);
  EXPECT_NE(xml.find("<x14:id>{5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB}</x14:id>"), std::string::npos) << xml;

  auto out = RoundTrip({cf});
  ASSERT_EQ(out.size(), 1u);
  ASSERT_EQ(out[0].rules.size(), 1u);
  ASSERT_TRUE(out[0].rules[0].ext_lst_raw.has_value());
  EXPECT_NE(out[0].rules[0].ext_lst_raw->find("5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB"), std::string::npos)
      << *out[0].rules[0].ext_lst_raw;
}

TEST(CFWriter, ConditionalFormattingBlockExtLstRoundTrips) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::Expression;
  r.formula1 = "A1=0";
  r.dxf_id = 0u;
  cf.rules.push_back(std::move(r));
  cf.ext_lst_raw = "<extLst><ext uri=\"{some-future-extension}\"><futureThing/></ext></extLst>";

  std::string xml = write_conditional_formattings({cf}, kDxfCount);
  EXPECT_NE(xml.find("<futureThing/>"), std::string::npos) << xml;

  auto out = RoundTrip({cf});
  ASSERT_EQ(out.size(), 1u);
  ASSERT_TRUE(out[0].ext_lst_raw.has_value());
  EXPECT_NE(out[0].ext_lst_raw->find("futureThing"), std::string::npos) << *out[0].ext_lst_raw;
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

// Excel treats a `dxfId` its `<dxfs>` table cannot resolve as package
// corruption and repairs the sheet by discarding *all* of its conditional
// formatting. Emitting the attribute is therefore strictly worse than
// dropping it: the rule survives, only its differential format is lost.
TEST(CFWriter, DxfIdBeyondTheDxfsTableIsNotEmitted) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule in_range{};
  in_range.type = cf::RuleType::CellIs;
  in_range.priority = 1;
  in_range.op = cf::CellIsOperator::GreaterThan;
  in_range.formula1 = "5";
  in_range.dxf_id = 1u;
  cf::CFRule dangling{};
  dangling.type = cf::RuleType::CellIs;
  dangling.priority = 2;
  dangling.op = cf::CellIsOperator::LessThan;
  dangling.formula1 = "0";
  dangling.dxf_id = 2u;
  cf.rules.push_back(in_range);
  cf.rules.push_back(dangling);

  const std::string xml = write_conditional_formattings({cf}, /*dxf_count=*/2u);
  EXPECT_NE(xml.find("dxfId=\"1\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("dxfId=\"2\""), std::string::npos) << xml;
  // The rule itself is still written; only its unresolvable format
  // reference is dropped.
  EXPECT_NE(xml.find("operator=\"lessThan\""), std::string::npos) << xml;
}

TEST(CFWriter, EveryEmittedDxfIdResolvesAgainstAnEmptyDxfsTable) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {0, 0}});
  cf::CFRule r{};
  r.type = cf::RuleType::CellIs;
  r.priority = 1;
  r.op = cf::CellIsOperator::Equal;
  r.formula1 = "1";
  r.dxf_id = 0u;
  cf.rules.push_back(r);

  EXPECT_EQ(write_conditional_formattings({cf}, /*dxf_count=*/0u).find("dxfId="), std::string::npos);
  EXPECT_NE(write_conditional_formattings({cf}, /*dxf_count=*/1u).find("dxfId=\"0\""), std::string::npos);
}

// --- x14 data-bar extension ------------------------------------------
//
// The settings below have no legacy `<dataBar>` attribute, so they only
// survive a save if the writer builds the x14 counterpart. `RoundTrip`
// above deliberately omits the worksheet `<extLst>`, which is where that
// counterpart lives; these tests use `RoundTripWithOverlay` so a failure
// to emit either half shows up as a lost setting rather than as XML that
// merely looks right.

/// A data bar carrying every x14-only setting, so one round trip covers
/// all of them.
cf::ConditionalFormat DataBarWithExtensionSettings(std::string id) {
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {4, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::DataBar;
  r.id = std::move(id);
  cf::DataBarSpec d;
  d.min = {cf::CfvoType::Min, "", true};
  d.max = {cf::CfvoType::Max, "", true};
  d.fill = {0x63, 0x8E, 0xC6, 255};
  d.border = cf::Color{0x00, 0x33, 0x66, 255};
  d.negative_fill = cf::Color{0xFF, 0x00, 0x00, 255};
  d.negative_border = cf::Color{0x99, 0x00, 0x00, 255};
  d.axis_position = cf::DataBarAxisPosition::Middle;
  d.axis_color = cf::Color{0x11, 0x22, 0x33, 255};
  d.gradient = false;
  d.min_length_pct = 0;
  d.max_length_pct = 100;
  r.data_bar = std::move(d);
  cf.rules.push_back(std::move(r));
  return cf;
}

/// Round trip through both halves of the save path: the legacy blocks
/// and the worksheet-level `<extLst>` the extension entries are merged
/// into. `existing_ext_lst` stands in for an overlay the source file
/// already had.
std::vector<cf::ConditionalFormat> RoundTripWithOverlay(const std::vector<cf::ConditionalFormat>& input,
                                                        const std::string& existing_ext_lst = std::string()) {
  std::string xml = "<worksheet>";
  xml.append(write_conditional_formattings(input, kDxfCount));
  xml.append(merge_x14_cf_entries(existing_ext_lst, build_x14_cf_overlay_entries(input)));
  xml.append("</worksheet>");
  pugi::xml_document doc;
  pugi::xml_parse_result rc = doc.load_string(xml.c_str());
  EXPECT_TRUE(rc) << rc.description() << "\n" << xml;
  auto out_or = read_conditional_formats(doc.child("worksheet"));
  EXPECT_TRUE(static_cast<bool>(out_or));
  return std::move(out_or.value());
}

TEST(CFWriter, ProgrammaticDataBarExtensionSettingsSurviveASaveCycle) {
  const auto input = DataBarWithExtensionSettings("{FC000000-0000-0000-0000-000000000001}");
  auto out = RoundTripWithOverlay({input});

  ASSERT_EQ(out.size(), 1u);
  ASSERT_EQ(out[0].rules.size(), 1u);
  const cf::CFRule& r = out[0].rules[0];
  EXPECT_EQ(r.id, "{FC000000-0000-0000-0000-000000000001}");
  ASSERT_TRUE(r.data_bar.has_value());
  const cf::DataBarSpec& d = r.data_bar.value();
  EXPECT_EQ(d.fill, (cf::Color{0x63, 0x8E, 0xC6, 255}));
  ASSERT_TRUE(d.border.has_value());
  EXPECT_EQ(d.border.value(), (cf::Color{0x00, 0x33, 0x66, 255}));
  EXPECT_EQ(d.negative_fill, (cf::Color{0xFF, 0x00, 0x00, 255}));
  ASSERT_TRUE(d.negative_border.has_value());
  EXPECT_EQ(d.negative_border.value(), (cf::Color{0x99, 0x00, 0x00, 255}));
  EXPECT_EQ(d.axis_position, cf::DataBarAxisPosition::Middle);
  EXPECT_EQ(d.axis_color, (cf::Color{0x11, 0x22, 0x33, 255}));
  EXPECT_FALSE(d.gradient);
  EXPECT_EQ(d.min_length_pct, 0);
  EXPECT_EQ(d.max_length_pct, 100);
}

TEST(CFWriter, DataBarExtensionIsReachedThroughTheNestedIdNotAnAttribute) {
  const auto input = DataBarWithExtensionSettings("{FC000000-0000-0000-0000-000000000001}");
  const std::string legacy = write_conditional_formattings({input}, kDxfCount);

  EXPECT_NE(legacy.find("<x14:id>{FC000000-0000-0000-0000-000000000001}</x14:id>"), std::string::npos) << legacy;
  EXPECT_NE(legacy.find("uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\""), std::string::npos) << legacy;
  // Excel drops both the attribute and the block it points at when it
  // re-saves a file spelled that way.
  EXPECT_EQ(legacy.find("<cfRule type=\"dataBar\" priority=\"1\" id="), std::string::npos) << legacy;
}

TEST(CFWriter, PlainDataBarProducesNoExtensionContent) {
  // A bar expressible in the legacy element alone must not grow an x14
  // overlay: a file that never had one would gain extension bytes on
  // every save.
  cf::ConditionalFormat cf{};
  cf.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::DataBar;
  r.id = "{FC000000-0000-0000-0000-000000000007}";
  cf::DataBarSpec d;
  d.min = {cf::CfvoType::Min, "", true};
  d.max = {cf::CfvoType::Max, "", true};
  d.fill = {0x63, 0x8E, 0xC6, 255};
  d.negative_fill = d.fill;
  r.data_bar = std::move(d);
  cf.rules.push_back(std::move(r));

  EXPECT_TRUE(build_x14_cf_overlay_entries({cf}).empty());
  EXPECT_EQ(write_conditional_formattings({cf}, kDxfCount).find("x14:id"), std::string::npos);
}

TEST(CFWriter, LoadedOverlayWinsOverARebuiltEntry) {
  // The captured overlay carries the whole `<x14:cfRule>` payload,
  // including parts the model does not represent. Appending a rebuilt
  // entry beside it would duplicate the id and drop those parts.
  const auto input = DataBarWithExtensionSettings("{5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB}");
  const std::string loaded =
      "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
      "<x14:conditionalFormattings><x14:conditionalFormatting>"
      "<x14:cfRule type=\"dataBar\" id=\"{5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB}\">"
      "<x14:dataBar minLength=\"0\" maxLength=\"100\"><x14:cfvo type=\"autoMin\"/>"
      "<x14:cfvo type=\"autoMax\"/><x14:someUnmodelledThing/></x14:dataBar></x14:cfRule>"
      "<xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>";

  const std::string merged = merge_x14_cf_entries(loaded, build_x14_cf_overlay_entries({input}));
  EXPECT_EQ(merged, loaded);
}

TEST(CFWriter, RebuiltEntryJoinsAnOverlayThatBelongsToAnotherRule) {
  const auto input = DataBarWithExtensionSettings("{FC000000-0000-0000-0000-000000000001}");
  const std::string loaded =
      "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
      "<x14:conditionalFormattings><x14:conditionalFormatting>"
      "<x14:cfRule type=\"dataBar\" id=\"{5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB}\">"
      "<x14:dataBar><x14:cfvo type=\"autoMin\"/><x14:cfvo type=\"autoMax\"/></x14:dataBar>"
      "</x14:cfRule><xm:sqref>B1:B5</xm:sqref>"
      "</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>";

  const std::string merged = merge_x14_cf_entries(loaded, build_x14_cf_overlay_entries({input}));
  EXPECT_NE(merged.find("{5A9D8B1C-3E4F-4A2B-9C1D-1234567890AB}"), std::string::npos) << merged;
  EXPECT_NE(merged.find("{FC000000-0000-0000-0000-000000000001}"), std::string::npos) << merged;
  // One `<ext>`, not a second one alongside the existing block.
  EXPECT_EQ(merged.find("{78C0D931-6437-407d-A8EE-F0AAD7539E65}"),
            merged.rfind("{78C0D931-6437-407d-A8EE-F0AAD7539E65}"))
      << merged;
}

TEST(CFWriter, EmptyEntryListLeavesTheOverlayByteIdentical) {
  const std::string loaded = "<extLst><ext uri=\"{some-future-extension}\"><futureThing/></ext></extLst>";
  EXPECT_EQ(merge_x14_cf_entries(loaded, std::string()), loaded);
  EXPECT_TRUE(merge_x14_cf_entries(std::string(), std::string()).empty());
}

TEST(CFWriter, ExtensionEntryReusesTheRuleExtLstWhenTheRuleAlreadyHasOne) {
  // A rule can arrive with an `<extLst>` holding some other extension.
  // `CT_CfRule` allows one, so the x14 link has to join it rather than
  // be emitted as a second element.
  auto input = DataBarWithExtensionSettings("{FC000000-0000-0000-0000-000000000001}");
  input.rules[0].ext_lst_raw = "<extLst><ext uri=\"{some-future-extension}\"><futureThing/></ext></extLst>";

  const std::string legacy = write_conditional_formattings({input}, kDxfCount);
  EXPECT_EQ(legacy.find("<extLst>"), legacy.rfind("<extLst>")) << legacy;
  EXPECT_NE(legacy.find("<futureThing/>"), std::string::npos) << legacy;
  EXPECT_NE(legacy.find("<x14:id>{FC000000-0000-0000-0000-000000000001}</x14:id>"), std::string::npos) << legacy;
}

}  // namespace
}  // namespace formulon::io
