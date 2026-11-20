// Standalone x14 visual rules in XLSB (colour scale, data bar and icon set, stock or custom): the records Excel
// writes for a threshold naming another sheet, read into the model and written back byte for byte, and the shapes the
// writer refuses.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/cf_x14_records_internal.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"
#include "writer_test_helpers.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBeginCondFmt14Id = 1046;
constexpr std::uint16_t kBeginCfRule14Id = 1048;
constexpr std::uint16_t kColor14Id = 1055;
constexpr std::uint16_t kBeginCfs14Id = 1135;
constexpr std::uint16_t kEndCfs14Id = 1136;
constexpr std::uint16_t kBrtBeginCfRuleId = 463;
constexpr std::uint16_t kBrtBeginCondFmtId = 461;
constexpr std::uint16_t kBrtEndCondFmtId = 462;
constexpr std::uint16_t kCfRuleExtId = 1146;
constexpr const char* kSheetPart = "xl/worksheets/sheet1.bin";

using RecordRun = std::vector<std::pair<std::uint16_t, std::vector<std::uint8_t>>>;

std::vector<std::uint8_t> ReadFixture(const std::string& relative) {
  return ReadFixtureFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + relative);
}

Workbook LoadXlsx(const std::vector<std::uint8_t>& bytes) {
  auto r = read_ooxml(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto r = read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> SaveXlsb(const Workbook& wb) {
  auto bytes = write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(bytes)) << (bytes ? "" : bytes.error().message);
  return bytes ? std::move(bytes.value()) : std::vector<std::uint8_t>{};
}

std::vector<std::uint8_t> SheetPart(const std::vector<std::uint8_t>& package, const char* part = kSheetPart) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry(part);
  EXPECT_TRUE(static_cast<bool>(entry)) << part;
  return entry ? std::move(entry.value()) : std::vector<std::uint8_t>{};
}

/// The records of `part` from the first `from` record through the last `to` record.
RecordRun RunBetween(const std::vector<std::uint8_t>& part, std::uint16_t from, std::uint16_t to) {
  const std::vector<FramedRecord> recs = split_records(part);
  std::size_t first = recs.size();
  std::size_t last = recs.size();
  for (std::size_t i = 0; i < recs.size(); ++i) {
    first = first == recs.size() && recs[i].type == from ? i : first;
    last = recs[i].type == to ? i : last;
  }
  RecordRun out;
  for (std::size_t i = first; i < recs.size() && last != recs.size() && i <= last; ++i) {
    out.emplace_back(recs[i].type,
                     std::vector<std::uint8_t>(recs[i].payload.data, recs[i].payload.data + recs[i].payload.size));
  }
  return out;
}

RecordRun X14CfRun(const std::vector<std::uint8_t>& package) {
  return RunBetween(SheetPart(package), kBeginCfs14Id, kEndCfs14Id);
}

std::size_t CountRecords(const std::vector<std::uint8_t>& part, std::uint16_t type) {
  std::size_t n = 0;
  for (const FramedRecord& rec : split_records(part)) {
    n += rec.type == type ? 1U : 0U;
  }
  return n;
}

std::string VisualFixture(int n, const char* suffix) {
  return "x14_visual_cross/V" + std::to_string(n) + suffix;
}

// Authored xlsx -> xlsb, and Excel's xlsb -> xlsb, both write the CF run Excel saved.
TEST(XlsbX14Visual, AuthoredAndExcelSavedBooksWriteExcelsCfRun) {
  for (int n = 1; n <= 12; ++n) {
    SCOPED_TRACE("V" + std::to_string(n));
    const std::vector<std::uint8_t> excel_file = ReadFixture(VisualFixture(n, "_sit.xlsb"));
    const RecordRun excel = RunBetween(SheetPart(excel_file), kBeginCfs14Id, kEndCfs14Id);
    ASSERT_FALSE(excel.empty());
    EXPECT_EQ(X14CfRun(SaveXlsb(LoadXlsx(ReadFixture(VisualFixture(n, ".xlsx"))))), excel);
    EXPECT_EQ(X14CfRun(SaveXlsb(LoadXlsb(excel_file))), excel);
  }
}

const cf::CFRule* FindRule(const Workbook& wb, cf::RuleType type, std::size_t nth = 0) {
  for (const cf::ConditionalFormat& f : wb.sheet(0).conditional_formats()) {
    for (const cf::CFRule& r : f.rules) {
      if (r.type == type && nth-- == 0U) {
        return &r;
      }
    }
  }
  return nullptr;
}

constexpr const char* kCrossRef = "Sheet2!$A$1";
constexpr const char* kDefaultId = "{00000000-000E-0000-0000-000001000000}";

void ExpectThreshold(const cf::CfValueObject& v, cf::CfvoType type, const std::string& value, bool gte = true) {
  EXPECT_EQ(v.type, type);
  EXPECT_EQ(v.value, value);
  EXPECT_EQ(v.gte, gte);
}

const cf::Color kFill{0x63, 0x8E, 0xC6, 255};

TEST(XlsbX14Visual, ExcelSavedDataBarWithShowValueOffReadsAsAModelRule) {
  const Workbook wb = LoadXlsb(ReadFixture(VisualFixture(1, "_sit.xlsb")));
  ASSERT_EQ(wb.sheet(0).conditional_formats().size(), 1U);
  const cf::CFRule* rule = FindRule(wb, cf::RuleType::DataBar);
  ASSERT_NE(rule, nullptr);
  EXPECT_EQ(rule->priority, 1);
  EXPECT_EQ(rule->id, kDefaultId);
  const cf::DataBarSpec& bar = *rule->data_bar;
  ExpectThreshold(bar.min, cf::CfvoType::Formula, kCrossRef);
  ExpectThreshold(bar.max, cf::CfvoType::Max, "");
  EXPECT_FALSE(bar.show_value);
  EXPECT_EQ(bar.fill, kFill);
  EXPECT_EQ(bar.axis_position, cf::DataBarAxisPosition::None);
  EXPECT_TRUE(bar.gradient);
}

TEST(XlsbX14Visual, ExcelSavedColorScalesReadAsModelRules) {
  struct Stop {
    cf::CfvoType type;
    const char* value;
  };
  struct Case {
    int n;
    std::vector<Stop> stops;
  };
  const std::vector<Case> cases = {
      {2, {{cf::CfvoType::Min, ""}, {cf::CfvoType::Formula, kCrossRef}}},
      {3, {{cf::CfvoType::Number, "0"}, {cf::CfvoType::Formula, kCrossRef}}},
      {4, {{cf::CfvoType::Percent, "10"}, {cf::CfvoType::Formula, kCrossRef}}},
      {5, {{cf::CfvoType::Percentile, "10"}, {cf::CfvoType::Formula, kCrossRef}}},
      {6, {{cf::CfvoType::Min, ""}, {cf::CfvoType::Formula, kCrossRef}, {cf::CfvoType::Max, ""}}},
  };
  const std::vector<cf::Color> colors = {{0xF8, 0x69, 0x6B, 255}, {0xFF, 0xEB, 0x84, 255}, {0x63, 0xBE, 0x7B, 255}};
  for (const Case& c : cases) {
    SCOPED_TRACE("V" + std::to_string(c.n));
    const Workbook wb = LoadXlsb(ReadFixture(VisualFixture(c.n, "_sit.xlsb")));
    const cf::CFRule* rule = FindRule(wb, cf::RuleType::ColorScale);
    ASSERT_NE(rule, nullptr);
    EXPECT_EQ(rule->id, kDefaultId);
    const cf::ColorScaleSpec& scale = *rule->color_scale;
    ASSERT_EQ(scale.thresholds.size(), c.stops.size());
    ASSERT_EQ(scale.colors.size(), c.stops.size());
    for (std::size_t i = 0; i < c.stops.size(); ++i) {
      ExpectThreshold(scale.thresholds[i], c.stops[i].type, c.stops[i].value);
      // A two-stop scale runs from the first colour to the last.
      EXPECT_EQ(scale.colors[i], colors[c.stops.size() == 2U && i == 1U ? 2U : i]);
    }
  }
}

TEST(XlsbX14Visual, ExcelSavedDataBarWithEveryOptionalColourReadsAsAModelRule) {
  const Workbook wb = LoadXlsb(ReadFixture(VisualFixture(7, "_sit.xlsb")));
  const cf::CFRule* rule = FindRule(wb, cf::RuleType::DataBar);
  ASSERT_NE(rule, nullptr);
  EXPECT_EQ(rule->id, "{5A1E0C11-0000-4000-8000-000000000007}");
  const cf::DataBarSpec& bar = *rule->data_bar;
  EXPECT_EQ(bar.min_length_pct, 5U);
  EXPECT_EQ(bar.max_length_pct, 95U);
  EXPECT_TRUE(bar.show_value);
  EXPECT_FALSE(bar.gradient);
  EXPECT_EQ(bar.direction, cf::DataBarDirection::RightToLeft);
  EXPECT_EQ(bar.axis_position, cf::DataBarAxisPosition::Automatic);
  EXPECT_EQ(bar.fill, kFill);
  ASSERT_TRUE(bar.border.has_value());
  EXPECT_EQ(*bar.border, (cf::Color{0x11, 0x22, 0x33, 255}));
  EXPECT_EQ(bar.negative_fill, (cf::Color{0x44, 0x55, 0x66, 255}));
  ASSERT_TRUE(bar.negative_border.has_value());
  EXPECT_EQ(*bar.negative_border, (cf::Color{0x77, 0x88, 0x99, 255}));
  EXPECT_EQ(bar.axis_color, (cf::Color{0xAA, 0xBB, 0xCC, 255}));
}

TEST(XlsbX14Visual, ExcelSavedDataBarWithOnlyTheAxisColourReadsAsAModelRule) {
  const Workbook wb = LoadXlsb(ReadFixture(VisualFixture(8, "_sit.xlsb")));
  const cf::CFRule* rule = FindRule(wb, cf::RuleType::DataBar);
  ASSERT_NE(rule, nullptr);
  const cf::DataBarSpec& bar = *rule->data_bar;
  EXPECT_EQ(bar.axis_position, cf::DataBarAxisPosition::Automatic);
  EXPECT_FALSE(bar.border.has_value());
  EXPECT_EQ(bar.negative_fill, bar.fill);
  EXPECT_FALSE(bar.negative_border.has_value());
  EXPECT_EQ(bar.axis_color, (cf::Color{0, 0, 0, 255}));
}

TEST(XlsbX14Visual, ExcelSavedStockIconSetsReadAsModelRules) {
  {
    const Workbook wb = LoadXlsb(ReadFixture(VisualFixture(9, "_sit.xlsb")));
    const cf::CFRule* rule = FindRule(wb, cf::RuleType::IconSet);
    ASSERT_NE(rule, nullptr);
    const cf::IconSetSpec& set = *rule->icon_set;
    EXPECT_EQ(set.name, cf::IconSetName::Three_Arrows);
    EXPECT_TRUE(set.custom_icons.empty());
    ExpectThreshold(set.floor, cf::CfvoType::Percent, "0");
    ASSERT_EQ(set.thresholds.size(), 2U);
    ExpectThreshold(set.thresholds[0], cf::CfvoType::Formula, kCrossRef);
    ExpectThreshold(set.thresholds[1], cf::CfvoType::Percent, "67");
  }
  {
    const Workbook wb = LoadXlsb(ReadFixture(VisualFixture(10, "_sit.xlsb")));
    const cf::CFRule* rule = FindRule(wb, cf::RuleType::IconSet);
    ASSERT_NE(rule, nullptr);
    const cf::IconSetSpec& set = *rule->icon_set;
    ExpectThreshold(set.floor, cf::CfvoType::Number, "0", true);
    ASSERT_EQ(set.thresholds.size(), 2U);
    ExpectThreshold(set.thresholds[0], cf::CfvoType::Formula, kCrossRef, false);
    ExpectThreshold(set.thresholds[1], cf::CfvoType::Number, "5", true);
  }
  {
    const Workbook wb = LoadXlsb(ReadFixture(VisualFixture(11, "_sit.xlsb")));
    const cf::CFRule* rule = FindRule(wb, cf::RuleType::IconSet);
    ASSERT_NE(rule, nullptr);
    const cf::IconSetSpec& set = *rule->icon_set;
    EXPECT_EQ(set.name, cf::IconSetName::Four_Arrows);
    EXPECT_TRUE(set.reverse);
    EXPECT_FALSE(set.show_value);
    ExpectThreshold(set.floor, cf::CfvoType::Percentile, "0");
    ASSERT_EQ(set.thresholds.size(), 3U);
    ExpectThreshold(set.thresholds[0], cf::CfvoType::Formula, kCrossRef);
    ExpectThreshold(set.thresholds[1], cf::CfvoType::Percentile, "50");
    ExpectThreshold(set.thresholds[2], cf::CfvoType::Percentile, "75");
  }
}

TEST(XlsbX14Visual, ExcelSavedMixedBlockReadsAsOneFormatWithTheLegacyRuleFirst) {
  const Workbook wb = LoadXlsb(ReadFixture(VisualFixture(12, "_sit.xlsb")));
  ASSERT_EQ(wb.sheet(0).conditional_formats().size(), 1U);
  const std::vector<cf::CFRule>& rules = wb.sheet(0).conditional_formats()[0].rules;
  ASSERT_EQ(rules.size(), 2U);
  EXPECT_EQ(rules[0].type, cf::RuleType::CellIs);
  EXPECT_EQ(rules[0].priority, 1);
  EXPECT_EQ(rules[1].type, cf::RuleType::DataBar);
  EXPECT_EQ(rules[1].priority, 2);
  EXPECT_EQ(rules[1].id, "{00000000-000E-0000-0000-000002000000}");
}

// A visual rule with a threshold naming another sheet and a legacy rule on one range: the legacy block carries the
// legacy rule alone, the data bar is its own x14 block, and both survive a second save unchanged.
TEST(XlsbX14Visual, MixedBlockSplitsIntoALegacyAndAnX14BlockAndResavesStably) {
  const std::vector<std::uint8_t> first = SaveXlsb(LoadXlsb(ReadFixture(VisualFixture(12, "_sit.xlsb"))));
  const std::vector<std::uint8_t> part = SheetPart(first);
  EXPECT_EQ(CountRecords(part, kBrtBeginCondFmtId), 1U);
  EXPECT_EQ(CountRecords(part, kBrtBeginCfRuleId), 1U);
  EXPECT_EQ(CountRecords(part, kBeginCondFmt14Id), 1U);
  EXPECT_EQ(CountRecords(part, kBeginCfRule14Id), 1U);
  EXPECT_EQ(SheetPart(SaveXlsb(LoadXlsb(first))), part);
}

cf::CFRule CrossDataBar() {
  cf::CFRule rule;
  rule.type = cf::RuleType::DataBar;
  rule.priority = 1;
  cf::DataBarSpec bar;
  bar.min = {cf::CfvoType::Formula, kCrossRef, true};
  bar.max = {cf::CfvoType::Max, "", true};
  bar.fill = kFill;
  bar.negative_fill = kFill;
  bar.axis_position = cf::DataBarAxisPosition::None;
  rule.data_bar = bar;
  return rule;
}

cf::CFRule CrossColorScale() {
  cf::CFRule rule;
  rule.type = cf::RuleType::ColorScale;
  rule.priority = 1;
  cf::ColorScaleSpec scale;
  scale.thresholds = {{cf::CfvoType::Min, "", true}, {cf::CfvoType::Formula, kCrossRef, true}};
  scale.colors = {{0xF8, 0x69, 0x6B, 255}, {0x63, 0xBE, 0x7B, 255}};
  rule.color_scale = scale;
  return rule;
}

cf::CFRule CrossIconSet(cf::IconSetName name = cf::IconSetName::Three_Arrows) {
  cf::CFRule rule;
  rule.type = cf::RuleType::IconSet;
  rule.priority = 1;
  cf::IconSetSpec set;
  set.name = name;
  set.floor = {cf::CfvoType::Percent, "0", true};
  const std::size_t levels = cf::icon_count(name);
  for (std::size_t i = 1; i < levels; ++i) {
    set.thresholds.push_back(i == 1U ? cf::CfValueObject{cf::CfvoType::Formula, kCrossRef, true}
                                     : cf::CfValueObject{cf::CfvoType::Percent, std::to_string(i * 20U), true});
  }
  rule.icon_set = set;
  return rule;
}

Workbook BookWith(const cf::CFRule& rule, bool pivot_scope = false) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  cf::ConditionalFormat format;
  format.sqref = {cf::CFCellRange{{0, 0}, {9, 0}}};
  format.pivot_scope = pivot_scope;
  format.rules.push_back(rule);
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(format));
  return wb;
}

void ExpectRefused(const Workbook& wb) {
  auto written = write_xlsb(wb);
  ASSERT_FALSE(static_cast<bool>(written));
  EXPECT_EQ(written.error().code, FormulonErrorCode::kInvalidArgument) << written.error().message;
}

void ExpectWritten(const cf::CFRule& rule) {
  auto written = write_xlsb(BookWith(rule));
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
}

TEST(XlsbX14Visual, IdLessRulesGetTheIdExcelGivesAndGivenIdsAreKept) {
  // An id-less rule gets the id Excel gives it; a given id is kept.
  for (cf::CFRule rule : {CrossDataBar(), CrossColorScale(), CrossIconSet()}) {
    const Workbook plain = LoadXlsb(SaveXlsb(BookWith(rule)));
    const cf::CFRule* back = FindRule(plain, rule.type);
    ASSERT_NE(back, nullptr);
    EXPECT_EQ(back->id, kDefaultId);
    rule.id = "{FC000000-0000-0000-0000-0000000000B2}";
    const Workbook named = LoadXlsb(SaveXlsb(BookWith(rule)));
    ASSERT_NE(FindRule(named, rule.type), nullptr);
    EXPECT_EQ(FindRule(named, rule.type)->id, rule.id);
  }
}

// Excel drops stopIfTrue from a data bar and a stock icon set stored only in the x14 extension; a colour scale with it
// has no measured record form.
TEST(XlsbX14Visual, StopIfTrueIsDroppedFromDataBarsAndStockIconSets) {
  for (cf::CFRule rule : {CrossDataBar(), CrossIconSet()}) {
    const RecordRun plain = X14CfRun(SaveXlsb(BookWith(rule)));
    rule.stop_if_true = true;
    EXPECT_EQ(X14CfRun(SaveXlsb(BookWith(rule))), plain);
  }
  cf::CFRule scale = CrossColorScale();
  scale.stop_if_true = true;
  ExpectRefused(BookWith(scale));
}

TEST(XlsbX14Visual, RuleExtensionXmlIsRefused) {
  cf::CFRule rule = CrossIconSet();
  rule.ext_lst_raw = "<extLst/>";
  ExpectRefused(BookWith(rule));
}

TEST(XlsbX14Visual, NonPositivePriorityIsRefused) {
  cf::CFRule rule = CrossDataBar();
  rule.priority = 0;
  ExpectRefused(BookWith(rule));
}

TEST(XlsbX14Visual, IdThatIsNotAGuidIsRefused) {
  cf::CFRule rule = CrossColorScale();
  rule.id = "not-a-guid";
  ExpectRefused(BookWith(rule));
}

TEST(XlsbX14Visual, DataBarWithAMiddleAxisIsRefused) {
  cf::CFRule rule = CrossDataBar();
  rule.data_bar->axis_position = cf::DataBarAxisPosition::Middle;
  ExpectRefused(BookWith(rule));
}

// A data bar or colour scale threshold has no gte; gte = false is written as the rule without it, as Excel saves it.
TEST(XlsbX14Visual, GteFalseOnDataBarAndColorScaleThresholdsIsDropped) {
  cf::CFRule bar = CrossDataBar();
  const RecordRun bar_plain = X14CfRun(SaveXlsb(BookWith(bar)));
  bar.data_bar->min.gte = false;
  bar.data_bar->max.gte = false;
  EXPECT_EQ(X14CfRun(SaveXlsb(BookWith(bar))), bar_plain);
  cf::CFRule scale = CrossColorScale();
  const RecordRun scale_plain = X14CfRun(SaveXlsb(BookWith(scale)));
  scale.color_scale->thresholds[1].gte = false;
  EXPECT_EQ(X14CfRun(SaveXlsb(BookWith(scale))), scale_plain);
}

TEST(XlsbX14Visual, DataBarAutoThresholdIsRefused) {
  cf::CFRule rule = CrossDataBar();
  rule.data_bar->max = {cf::CfvoType::AutoMax, "", true};
  ExpectRefused(BookWith(rule));
}

TEST(XlsbX14Visual, ColorScaleAutoThresholdIsRefused) {
  cf::CFRule rule = CrossColorScale();
  rule.color_scale->thresholds[0] = {cf::CfvoType::AutoMin, "", true};
  ExpectRefused(BookWith(rule));
}

TEST(XlsbX14Visual, DataBarNegativeBorderWithoutABorderIsRefused) {
  cf::CFRule rule = CrossDataBar();
  rule.data_bar->negative_border = cf::Color{1, 2, 3, 255};
  ExpectRefused(BookWith(rule));
}

// A bar without an axis keeps no axis colour, as Excel saves it.
TEST(XlsbX14Visual, DataBarAxisColourWithoutAnAxisIsDropped) {
  cf::CFRule rule = CrossDataBar();
  const RecordRun plain = X14CfRun(SaveXlsb(BookWith(rule)));
  rule.data_bar->axis_color = cf::Color{1, 2, 3, 255};
  EXPECT_EQ(X14CfRun(SaveXlsb(BookWith(rule))), plain);
}

TEST(XlsbX14Visual, FiveLevelStockIconSetIsRefused) {
  ExpectRefused(BookWith(CrossIconSet(cf::IconSetName::Five_Arrows)));
}

// A threshold formula with an absolute same-sheet reference: BrtCFVO14's cbFmla is the future-record size (cce, as
// Excel writes it), both alone and beside a reference to another sheet.
TEST(XlsbX14Visual, ThresholdWithASameSheetReferenceWritesExcelsSizeField) {
  for (const char* stem : {"W16", "W17"}) {
    SCOPED_TRACE(stem);
    const std::string base = std::string("x14_visual_cross/") + stem;
    const std::vector<std::uint8_t> excel_file = ReadFixture(base + "_sit.xlsb");
    const RecordRun excel = RunBetween(SheetPart(excel_file), kBeginCfs14Id, kEndCfs14Id);
    ASSERT_FALSE(excel.empty());
    EXPECT_EQ(X14CfRun(SaveXlsb(LoadXlsx(ReadFixture(base + ".xlsx")))), excel);
    EXPECT_EQ(X14CfRun(SaveXlsb(LoadXlsb(excel_file))), excel);
  }
}

TEST(XlsbX14Visual, ThresholdNamingAnotherWorkbookIsRefused) {
  cf::CFRule rule = CrossDataBar();
  rule.data_bar->max = {cf::CfvoType::Formula, "[Other.xlsx]Sheet1!$A$1", true};
  ExpectRefused(BookWith(rule));
}

TEST(XlsbX14Visual, TwoAndThreeStopColorScalesAndStockIconSetsAreWritten) {
  ExpectWritten(CrossColorScale());
  cf::CFRule three = CrossColorScale();
  three.color_scale->thresholds.insert(three.color_scale->thresholds.begin() + 1,
                                       cf::CfValueObject{cf::CfvoType::Number, "1", true});
  three.color_scale->colors.insert(three.color_scale->colors.begin() + 1, cf::Color{0xFF, 0xEB, 0x84, 255});
  ExpectWritten(three);
  ExpectWritten(CrossIconSet(cf::IconSetName::Four_Arrows));
}

// A standalone data bar a reader refuses (here, one colour too many) stays in the tail and is written back as it was.
TEST(XlsbX14Visual, UnreadableStandaloneDataBarSurvivesTheSave) {
  const std::vector<std::uint8_t> source = ReadFixture(VisualFixture(8, "_sit.xlsb"));
  std::vector<std::uint8_t> patched_part;
  bool doubled = false;
  for (const FramedRecord& rec : split_records(SheetPart(source))) {
    emit_record(patched_part, rec.type, rec.payload);
    if (rec.type == kColor14Id && !doubled) {
      emit_record(patched_part, rec.type, rec.payload);
      doubled = true;
    }
  }
  ASSERT_TRUE(doubled);
  const std::vector<std::uint8_t> patched = RepackWithEntries(source, {{kSheetPart, patched_part}});
  const Workbook wb = LoadXlsb(patched);
  EXPECT_EQ(FindRule(wb, cf::RuleType::DataBar), nullptr);
  const std::vector<std::uint8_t> saved = SheetPart(SaveXlsb(wb));
  EXPECT_EQ(CountRecords(saved, kBeginCfRule14Id), 1U);
  EXPECT_EQ(CountRecords(saved, kColor14Id), 3U);
  EXPECT_EQ(SheetPart(SaveXlsb(LoadXlsb(SaveXlsb(wb)))), saved);
}

// Linked data-bar payloads (iPri 0xFFFFFFFF, template 0) are not standalone rules: loading and saving keeps every
// legacy rule's link and every payload record.
TEST(XlsbX14Visual, LinkedDataBarPayloadsAreNotTakenAsStandaloneRules) {
  const std::vector<std::uint8_t> source = ReadFixture("xlsb_feature_x14bars.xlsb");
  const std::vector<std::uint8_t> before = SheetPart(source);
  const Workbook wb = LoadXlsb(source);
  std::size_t bars = 0;
  for (const cf::ConditionalFormat& f : wb.sheet(0).conditional_formats()) {
    for (const cf::CFRule& r : f.rules) {
      bars += r.type == cf::RuleType::DataBar ? 1U : 0U;
    }
  }
  EXPECT_EQ(bars, CountRecords(before, kBeginCfRule14Id));
  const std::vector<std::uint8_t> after = SheetPart(SaveXlsb(wb));
  EXPECT_EQ(CountRecords(after, kBeginCfRule14Id), CountRecords(before, kBeginCfRule14Id));
  EXPECT_EQ(CountRecords(after, kCfRuleExtId), CountRecords(before, kCfRuleExtId));
}

// One linked data bar (a legacy rule with an x14 payload) and one standalone data bar (a threshold on another sheet)
// on one sheet are both kept.
TEST(XlsbX14Visual, LinkedAndStandaloneDataBarsOnOneSheetAreBothKept) {
  Workbook wb = BookWith(CrossDataBar());
  cf::ConditionalFormat linked;
  linked.sqref = {cf::CFCellRange{{0, 2}, {9, 2}}};
  cf::CFRule rule;
  rule.type = cf::RuleType::DataBar;
  rule.priority = 2;
  rule.id = "{FC000000-0000-0000-0000-0000000000C3}";
  cf::DataBarSpec bar;
  bar.min = {cf::CfvoType::Min, "", true};
  bar.max = {cf::CfvoType::Max, "", true};
  bar.fill = kFill;
  bar.negative_fill = cf::Color{0xFF, 0, 0, 255};
  bar.axis_position = cf::DataBarAxisPosition::Automatic;
  rule.data_bar = bar;
  linked.rules.push_back(rule);
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(linked));
  const std::vector<std::uint8_t> first = SaveXlsb(wb);
  const std::vector<std::uint8_t> part = SheetPart(first);
  EXPECT_EQ(CountRecords(part, kBeginCfRule14Id), 2U);
  EXPECT_EQ(CountRecords(part, kBrtBeginCfRuleId), 1U);
  const Workbook back = LoadXlsb(first);
  const cf::CFRule* standalone = nullptr;
  const cf::CFRule* legacy = nullptr;
  for (const cf::ConditionalFormat& f : back.sheet(0).conditional_formats()) {
    for (const cf::CFRule& r : f.rules) {
      (r.priority == 1 ? standalone : legacy) = &r;
    }
  }
  ASSERT_NE(standalone, nullptr);
  ASSERT_NE(legacy, nullptr);
  ExpectThreshold(standalone->data_bar->min, cf::CfvoType::Formula, kCrossRef);
  EXPECT_EQ(legacy->data_bar->negative_fill, (cf::Color{0xFF, 0, 0, 255}));
  const std::vector<std::uint8_t> resaved = SheetPart(SaveXlsb(back));
  EXPECT_EQ(RunBetween(resaved, kBrtBeginCondFmtId, kBrtEndCondFmtId),
            RunBetween(part, kBrtBeginCondFmtId, kBrtEndCondFmtId));
  EXPECT_EQ(RunBetween(resaved, kBeginCfs14Id, kEndCfs14Id), RunBetween(part, kBeginCfs14Id, kEndCfs14Id));
}

/// The GUID bytes of every BrtBeginCFRule14 of `part` whose iPri (+16) is positive, and the count of the others.
std::pair<std::vector<std::vector<std::uint8_t>>, std::size_t> Rule14Guids(const std::vector<std::uint8_t>& part) {
  constexpr std::size_t kPriOffset = 16;
  constexpr std::size_t kGuidOffset = 50;
  constexpr std::size_t kGuidBytes = 16;
  std::vector<std::vector<std::uint8_t>> guids;
  std::size_t non_positive = 0;
  for (const FramedRecord& rec : split_records(part)) {
    if (rec.type != kBeginCfRule14Id || rec.payload.size < kGuidOffset + kGuidBytes) {
      continue;
    }
    const std::uint8_t* p = rec.payload.data;
    const auto pri = static_cast<std::int32_t>(static_cast<std::uint32_t>(p[kPriOffset]) |
                                               (static_cast<std::uint32_t>(p[kPriOffset + 1]) << 8U) |
                                               (static_cast<std::uint32_t>(p[kPriOffset + 2]) << 16U) |
                                               (static_cast<std::uint32_t>(p[kPriOffset + 3]) << 24U));
    if (pri <= 0) {
      ++non_positive;
      continue;
    }
    guids.emplace_back(p + kGuidOffset, p + kGuidOffset + kGuidBytes);
  }
  return {guids, non_positive};
}

// A linked data bar whose threshold is edited to name another sheet is written once, as a standalone rule: its
// retained payload goes with the link.
TEST(XlsbX14Visual, LinkedDataBarEditedToAnOtherSheetThresholdIsWrittenOnce) {
  Workbook wb = BookWith(CrossDataBar());
  cf::ConditionalFormat& format = wb.sheet(0).mutable_conditional_formats().front();
  cf::CFRule& rule = format.rules.front();
  rule.id = "{FC000000-0000-0000-0000-0000000000D4}";
  rule.data_bar->min = {cf::CfvoType::Min, "", true};
  rule.data_bar->negative_fill = cf::Color{0xFF, 0, 0, 255};
  rule.data_bar->axis_position = cf::DataBarAxisPosition::Automatic;
  const std::vector<std::uint8_t> linked_part = SheetPart(SaveXlsb(wb));
  ASSERT_EQ(CountRecords(linked_part, kCfRuleExtId), 1U);
  ASSERT_EQ(Rule14Guids(linked_part).second, 1U);
  Workbook back = LoadXlsb(SaveXlsb(wb));
  cf::CFRule* edited = nullptr;
  for (cf::ConditionalFormat& f : back.sheet(0).mutable_conditional_formats()) {
    for (cf::CFRule& r : f.rules) {
      edited = r.type == cf::RuleType::DataBar ? &r : edited;
    }
  }
  ASSERT_NE(edited, nullptr);
  edited->data_bar->min = {cf::CfvoType::Formula, kCrossRef, true};
  const std::vector<std::uint8_t> saved = SaveXlsb(back);
  const std::vector<std::uint8_t> part = SheetPart(saved);
  const auto [guids, non_positive] = Rule14Guids(part);
  EXPECT_EQ(guids.size(), 1U);
  EXPECT_EQ(non_positive, 0U);
  EXPECT_EQ(CountRecords(part, kCfRuleExtId), 0U);
  EXPECT_EQ(CountRecords(part, kBrtBeginCfRuleId), 0U);
  const Workbook reloaded = LoadXlsb(saved);
  const cf::CFRule* reread = FindRule(reloaded, cf::RuleType::DataBar);
  ASSERT_NE(reread, nullptr);
  EXPECT_EQ(reread->id, rule.id);
  ExpectThreshold(reread->data_bar->min, cf::CfvoType::Formula, kCrossRef);
}

// A generated id avoids the GUID of a rule retained verbatim: an id-less rule added beside Excel's colour scale (made
// unreadable, so it stays retained under the id the generator would give first) gets the next id.
TEST(XlsbX14Visual, GeneratedIdsAvoidRetainedRuleGuids) {
  const std::vector<std::uint8_t> source = ReadFixture(VisualFixture(2, "_sit.xlsb"));
  std::vector<std::uint8_t> patched_part;
  bool doubled = false;
  for (const FramedRecord& rec : split_records(SheetPart(source))) {
    emit_record(patched_part, rec.type, rec.payload);
    if (rec.type == kColor14Id && !doubled) {
      emit_record(patched_part, rec.type, rec.payload);
      doubled = true;
    }
  }
  ASSERT_TRUE(doubled);
  Workbook wb = LoadXlsb(RepackWithEntries(source, {{kSheetPart, patched_part}}));
  ASSERT_EQ(FindRule(wb, cf::RuleType::ColorScale), nullptr);
  cf::ConditionalFormat added;
  added.sqref = {cf::CFCellRange{{0, 3}, {9, 3}}};
  added.rules.push_back(CrossDataBar());
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(added));
  const std::vector<std::uint8_t> saved = SaveXlsb(wb);
  const auto [guids, non_positive] = Rule14Guids(SheetPart(saved));
  ASSERT_EQ(guids.size(), 2U);
  EXPECT_EQ(non_positive, 0U);
  EXPECT_NE(guids[0], guids[1]);
  const Workbook reloaded = LoadXlsb(saved);
  const cf::CFRule* bar = FindRule(reloaded, cf::RuleType::DataBar);
  ASSERT_NE(bar, nullptr);
  EXPECT_NE(bar->id, kDefaultId);
}

// Two stock icon sets of one spec in one format, both stored as x14 records, read back as two rules: the first is not
// the legacy copy of the second.
TEST(XlsbX14Visual, StockIconSetsOfOneSpecAreNotMerged) {
  Workbook wb = BookWith(CrossIconSet());
  cf::CFRule second = CrossIconSet();
  second.priority = 2;
  second.id = "{FC000000-0000-0000-0000-0000000000D4}";
  wb.sheet(0).mutable_conditional_formats()[0].rules.push_back(second);
  const Workbook back = LoadXlsb(SaveXlsb(wb));
  ASSERT_EQ(back.sheet(0).conditional_formats().size(), 1U);
  EXPECT_EQ(back.sheet(0).conditional_formats()[0].rules.size(), 2U);
  EXPECT_NE(FindRule(back, cf::RuleType::IconSet, 0), nullptr);
  EXPECT_NE(FindRule(back, cf::RuleType::IconSet, 1), nullptr);
}

// Two x14-only data bars sharing one id and no payload, as Excel keeps them, write as legacy blocks without a link.
TEST(XlsbX14Visual, LegacyRulesWithAnIdButNoPayloadCarryNoLink) {
  const std::vector<std::uint8_t> excel_file = ReadFixture("cf_id3/b_sit.xlsb");
  const std::vector<std::uint8_t> excel_part = SheetPart(excel_file);
  const RecordRun excel = RunBetween(excel_part, kBrtBeginCondFmtId, kBrtEndCondFmtId);
  ASSERT_FALSE(excel.empty());
  const std::vector<std::uint8_t> ours = SheetPart(SaveXlsb(LoadXlsx(ReadFixture("cf_id3/b.xlsx"))));
  EXPECT_EQ(RunBetween(ours, kBrtBeginCondFmtId, kBrtEndCondFmtId), excel);
  EXPECT_EQ(CountRecords(ours, kCfRuleExtId), 0U);
  EXPECT_EQ(CountRecords(ours, kBeginCfRule14Id), 0U);
  const std::vector<std::uint8_t> resaved = SheetPart(SaveXlsb(LoadXlsb(excel_file)));
  EXPECT_EQ(RunBetween(resaved, kBrtBeginCondFmtId, kBrtEndCondFmtId), excel);
}

/// The `<conditionalFormatting>` elements of a worksheet part, and whether it holds any x14 rule.
std::pair<std::string, bool> XlsxCfBlocks(const std::vector<std::uint8_t>& sheet_xml) {
  const std::string xml(sheet_xml.begin(), sheet_xml.end());
  std::string blocks;
  const std::string open = "<conditionalFormatting ";
  const std::string close = "</conditionalFormatting>";
  for (std::size_t at = xml.find(open); at != std::string::npos; at = xml.find(open, at)) {
    const std::size_t end = xml.find(close, at);
    if (end == std::string::npos) {
      break;
    }
    blocks += xml.substr(at, end + close.size() - at);
    at = end;
  }
  return {blocks, xml.find("<x14:cfRule") != std::string::npos};
}

// The same two data bars saved to xlsx are legacy-only as well, as in Excel's save: one block per range, no id and no
// x14 rule.
TEST(XlsbX14Visual, LegacyRulesWithAnIdButNoPayloadSaveToXlsxLegacyOnly) {
  constexpr const char* kXmlPart = "xl/worksheets/sheet1.xml";
  const auto excel = XlsxCfBlocks(SheetPart(ReadFixture("cf_id3/b_sit.xlsx"), kXmlPart));
  ASSERT_FALSE(excel.first.empty());
  ASSERT_FALSE(excel.second);
  auto bytes = write_ooxml(LoadXlsx(ReadFixture("cf_id3/b.xlsx")));
  ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
  const auto ours = XlsxCfBlocks(SheetPart(bytes.value(), kXmlPart));
  EXPECT_FALSE(ours.second);
  EXPECT_EQ(ours.first.find("<extLst"), std::string::npos);
  EXPECT_EQ(ours.first.find(" id="), std::string::npos);
  const std::size_t a = ours.first.find("sqref=\"A1:A10\"><cfRule type=\"dataBar\" priority=\"1\">");
  const std::size_t b = ours.first.find("sqref=\"B1:B10\"><cfRule type=\"dataBar\" priority=\"2\">");
  ASSERT_NE(b, std::string::npos);
  EXPECT_LT(a, b);
}

// Shapes Excel saves by dropping what has no record form: an axis colour on a bar without an axis, gte = false on a
// data bar or colour scale threshold, stopIfTrue on an x14-only data bar or stock icon set, and a pivot-scoped block.
// The authored xlsx and Excel's own xlsb both write the CF run Excel saved.
TEST(XlsbX14Visual, ShapesExcelSavesLossilyWriteExcelsCfRun) {
  for (const char* stem : {"W6", "W11", "W12", "W13", "W14", "G1_pivot_scope"}) {
    SCOPED_TRACE(stem);
    const std::string base = std::string("x14_visual_cross/") + stem;
    const std::vector<std::uint8_t> excel_file = ReadFixture(base + "_sit.xlsb");
    const RecordRun excel = X14CfRun(excel_file);
    ASSERT_FALSE(excel.empty());
    EXPECT_EQ(X14CfRun(SaveXlsb(LoadXlsx(ReadFixture(base + ".xlsx")))), excel);
    EXPECT_EQ(X14CfRun(SaveXlsb(LoadXlsb(excel_file))), excel);
  }
}

// A linked data bar with showValue off: the x14 payload states showValue as the legacy rule does.
TEST(XlsbX14Visual, LinkedDataBarWithShowValueOffWritesExcelsRecords) {
  const std::vector<std::uint8_t> excel_file = ReadFixture("x14_visual_cross/W15_sit.xlsb");
  const std::vector<std::uint8_t> excel_part = SheetPart(excel_file);
  const RecordRun legacy = RunBetween(excel_part, kBrtBeginCondFmtId, kBrtEndCondFmtId);
  const RecordRun x14 = RunBetween(excel_part, kBeginCfs14Id, kEndCfs14Id);
  ASSERT_FALSE(legacy.empty());
  ASSERT_FALSE(x14.empty());
  for (const std::vector<std::uint8_t>& written :
       {SaveXlsb(LoadXlsx(ReadFixture("x14_visual_cross/W15.xlsx"))), SaveXlsb(LoadXlsb(excel_file))}) {
    const std::vector<std::uint8_t> part = SheetPart(written);
    EXPECT_EQ(RunBetween(part, kBrtBeginCondFmtId, kBrtEndCondFmtId), legacy);
    EXPECT_EQ(RunBetween(part, kBeginCfs14Id, kEndCfs14Id), x14);
  }
  const Workbook wb = LoadXlsb(excel_file);
  const cf::CFRule* rule = FindRule(wb, cf::RuleType::DataBar);
  ASSERT_NE(rule, nullptr);
  EXPECT_FALSE(rule->data_bar->show_value);
  EXPECT_EQ(rule->data_bar->negative_fill, (cf::Color{0xFF, 0x00, 0x00, 255}));
}

// A linked data bar whose minimum and maximum are formulas (FC: numeric control) writes the legacy and x14 record
// runs Excel saved.
TEST(XlsbX14Visual, LinkedDataBarWithFormulaThresholdsWritesExcelsRecords) {
  for (const char* stem : {"databar_formula_threshold_F1", "databar_formula_threshold_FC"}) {
    SCOPED_TRACE(stem);
    const std::string base = std::string("x14_visual_cross/") + stem + "_sit";
    const std::vector<std::uint8_t> excel_file = ReadFixture(base + ".xlsb");
    const std::vector<std::uint8_t> excel_part = SheetPart(excel_file);
    const RecordRun legacy = RunBetween(excel_part, kBrtBeginCondFmtId, kBrtEndCondFmtId);
    const RecordRun x14 = RunBetween(excel_part, kBeginCfs14Id, kEndCfs14Id);
    ASSERT_FALSE(legacy.empty());
    ASSERT_FALSE(x14.empty());
    for (const std::vector<std::uint8_t>& written :
         {SaveXlsb(LoadXlsx(ReadFixture(base + ".xlsx"))), SaveXlsb(LoadXlsb(excel_file))}) {
      const std::vector<std::uint8_t> part = SheetPart(written);
      EXPECT_EQ(RunBetween(part, kBrtBeginCondFmtId, kBrtEndCondFmtId), legacy);
      EXPECT_EQ(RunBetween(part, kBeginCfs14Id, kEndCfs14Id), x14);
    }
  }
}

// Excel's xlsb keeps the pivot scope of an x14-only block in its head; it reads back as a pivot-scoped format beside
// the unscoped control.
TEST(XlsbX14Visual, ExcelSavedPivotScopedBlockReadsAsAPivotScopedFormat) {
  const Workbook wb = LoadXlsb(ReadFixture("x14_visual_cross/G1_pivot_scope_sit.xlsb"));
  std::size_t scoped = 0;
  std::size_t plain = 0;
  for (const cf::ConditionalFormat& f : wb.sheet(0).conditional_formats()) {
    ASSERT_EQ(f.sqref.size(), 1U);
    ASSERT_EQ(f.rules.size(), 1U);
    EXPECT_EQ(f.rules[0].type, cf::RuleType::DataBar);
    if (f.pivot_scope) {
      ++scoped;
      EXPECT_EQ(f.sqref[0].first.row, 3U);
      EXPECT_EQ(f.sqref[0].first.col, 1U);
      EXPECT_EQ(f.rules[0].id, "{5A1E0C13-0000-4000-8000-000000000001}");
    } else {
      ++plain;
      EXPECT_EQ(f.rules[0].id, "{5A1E0C13-0000-4000-8000-000000000002}");
    }
  }
  EXPECT_EQ(scoped, 1U);
  EXPECT_EQ(plain, 1U);
}

std::string XlsxX14Rules(const std::vector<std::uint8_t>& package) {
  const std::vector<std::uint8_t> part = SheetPart(package, "xl/worksheets/sheet1.xml");
  const std::string xml(part.begin(), part.end());
  const std::size_t at = xml.find("<x14:conditionalFormattings>");
  const std::size_t end = xml.find("</x14:conditionalFormattings>");
  return at == std::string::npos || end == std::string::npos ? std::string() : xml.substr(at, end - at);
}

// The xlsx writer drops the same attributes Excel's save drops.
TEST(XlsbX14Visual, XlsxSaveDropsWhatExcelDrops) {
  struct Case {
    const char* stem;
    const char* gone;
  };
  for (const Case& c : {Case{"W6", "<x14:axisColor"}, Case{"W11", " gte="}, Case{"W12", " gte="},
                        Case{"W13", " stopIfTrue="}, Case{"W14", " stopIfTrue="}}) {
    SCOPED_TRACE(c.stem);
    const std::string authored = XlsxX14Rules(ReadFixture(std::string("x14_visual_cross/") + c.stem + ".xlsx"));
    ASSERT_NE(authored.find(c.gone), std::string::npos);
    ASSERT_EQ(XlsxX14Rules(ReadFixture(std::string("x14_visual_cross/") + c.stem + "_sit.xlsx")).find(c.gone),
              std::string::npos);
    auto bytes = write_ooxml(LoadXlsx(ReadFixture(std::string("x14_visual_cross/") + c.stem + ".xlsx")));
    ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
    const std::string ours = XlsxX14Rules(bytes.value());
    ASSERT_NE(ours.find("<x14:cfRule"), std::string::npos);
    EXPECT_EQ(ours.find(c.gone), std::string::npos);
  }
  auto pivot = write_ooxml(LoadXlsx(ReadFixture("x14_visual_cross/G1_pivot_scope.xlsx")));
  ASSERT_TRUE(static_cast<bool>(pivot)) << pivot.error().message;
  EXPECT_NE(XlsxX14Rules(pivot.value()).find("pivot=\"1\""), std::string::npos);
}

// A bar without an axis holds its authored axis colour in memory; Excel's save keeps none, so the xlsx save reloads
// with the default axis colour, as Excel's own save does.
TEST(XlsbX14Visual, AxisColourOfABarWithoutAnAxisReloadsAsExcelsSaveDoes) {
  const auto first_bar = [](const Workbook& wb) {
    for (const cf::ConditionalFormat& format : wb.sheet(0).conditional_formats()) {
      for (const cf::CFRule& rule : format.rules) {
        if (rule.data_bar) {
          return *rule.data_bar;
        }
      }
    }
    ADD_FAILURE() << "no data bar";
    return cf::DataBarSpec{};
  };
  const auto rgba = [](const cf::Color& c) {
    return std::to_string(c.r) + "," + std::to_string(c.g) + "," + std::to_string(c.b) + "," + std::to_string(c.a);
  };
  const Workbook authored = LoadXlsx(ReadFixture("x14_visual_cross/W6.xlsx"));
  const cf::DataBarSpec before = first_bar(authored);
  ASSERT_EQ(before.axis_position, cf::DataBarAxisPosition::None);
  EXPECT_EQ(rgba(before.axis_color), "170,187,204,255");
  const cf::DataBarSpec excel = first_bar(LoadXlsx(ReadFixture("x14_visual_cross/W6_sit.xlsx")));
  EXPECT_EQ(rgba(excel.axis_color), rgba(cf::DataBarSpec{}.axis_color));
  auto bytes = write_ooxml(authored);
  ASSERT_TRUE(static_cast<bool>(bytes)) << bytes.error().message;
  EXPECT_EQ(rgba(first_bar(LoadXlsx(bytes.value())).axis_color), rgba(excel.axis_color));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
