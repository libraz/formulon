// Custom icon sets in XLSB: a rule with per-level icons is a standalone x14
// rule (BrtBeginCFRule14 / BrtBeginIconSet14 / BrtCFVO14 / BrtCFIcon), read
// back into the same model the .xlsx reader fills.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr const char* kIconId = "{FC000000-0000-0000-0000-0000000000A1}";

cf::CFRule IconRule() {
  cf::CFRule rule;
  rule.type = cf::RuleType::IconSet;
  rule.priority = 2;
  rule.id = kIconId;
  cf::IconSetSpec spec;
  spec.name = cf::IconSetName::Three_Arrows;
  spec.floor = {cf::CfvoType::Number, "-5", true};
  spec.thresholds = {{cf::CfvoType::Percent, "33", false}, {cf::CfvoType::Formula, "Sheet1!$C$1", true}};
  spec.show_value = false;
  spec.custom_icons = {{"3Flags", 2}, {"NoIcons", 0}, {"3Stars", 1}};
  rule.icon_set = spec;
  return rule;
}

Workbook BookWithIcons() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  for (std::uint32_t r = 0; r < 5; ++r) {
    wb.sheet(0).set_cell_value(r, 0U, Value::number(r * 10.0));
  }
  cf::ConditionalFormat block;
  block.sqref = {cf::CFCellRange{{0, 0}, {4, 0}}};
  cf::CFRule plain;
  plain.type = cf::RuleType::Expression;
  plain.priority = 1;
  plain.formula1 = "A1>20";
  block.rules.push_back(plain);
  block.rules.push_back(IconRule());
  wb.sheet(0).mutable_conditional_formats().push_back(block);
  return wb;
}

std::vector<std::uint8_t> SaveXlsb(const Workbook& wb) {
  auto bytes = write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(bytes)) << (bytes ? "" : bytes.error().message);
  return bytes ? std::move(bytes.value()) : std::vector<std::uint8_t>{};
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto r = read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint16_t> SheetRecordTypes(const std::vector<std::uint8_t>& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(entry));
  std::vector<std::uint16_t> out;
  if (entry) {
    for (const FramedRecord& rec : split_records(entry.value())) {
      out.push_back(rec.type);
    }
  }
  return out;
}

std::size_t Count(const std::vector<std::uint16_t>& types, std::uint16_t type) {
  std::size_t n = 0;
  for (std::uint16_t t : types) {
    n += t == type ? 1U : 0U;
  }
  return n;
}

const cf::CFRule* FindRule(const Workbook& wb, const std::string& id) {
  for (const cf::ConditionalFormat& f : wb.sheet(0).conditional_formats()) {
    for (const cf::CFRule& r : f.rules) {
      if (r.id == id) {
        return &r;
      }
    }
  }
  return nullptr;
}

void ExpectIconRule(const Workbook& wb) {
  const cf::CFRule* rule = FindRule(wb, kIconId);
  ASSERT_NE(rule, nullptr);
  EXPECT_EQ(rule->type, cf::RuleType::IconSet);
  EXPECT_EQ(rule->priority, 2);
  ASSERT_TRUE(rule->icon_set.has_value());
  const cf::IconSetSpec& spec = *rule->icon_set;
  EXPECT_EQ(spec.name, cf::IconSetName::Three_Arrows);
  EXPECT_FALSE(spec.show_value);
  EXPECT_EQ(spec.floor.type, cf::CfvoType::Number);
  EXPECT_EQ(spec.floor.value, "-5");
  ASSERT_EQ(spec.thresholds.size(), 2U);
  EXPECT_EQ(spec.thresholds[0].type, cf::CfvoType::Percent);
  EXPECT_EQ(spec.thresholds[0].value, "33");
  EXPECT_FALSE(spec.thresholds[0].gte);
  EXPECT_EQ(spec.thresholds[1].type, cf::CfvoType::Formula);
  EXPECT_EQ(spec.thresholds[1].value, "Sheet1!$C$1");
  ASSERT_EQ(spec.custom_icons.size(), 3U);
  EXPECT_EQ(spec.custom_icons[0].icon_set, "3Flags");
  EXPECT_EQ(spec.custom_icons[0].icon_id, 2);
  EXPECT_EQ(spec.custom_icons[1].icon_set, "NoIcons");
  EXPECT_EQ(spec.custom_icons[1].icon_id, 0);
  EXPECT_EQ(spec.custom_icons[2].icon_set, "3Stars");
  EXPECT_EQ(spec.custom_icons[2].icon_id, 1);
}

TEST(XlsbCfIconsFidelity, ModelRuleIsWrittenAsAStandaloneX14Rule) {
  const std::vector<std::uint8_t> bytes = SaveXlsb(BookWithIcons());
  const std::vector<std::uint16_t> types = SheetRecordTypes(bytes);
  // No legacy BrtBeginIconSet; one x14 rule with three cfvos and icons.
  EXPECT_EQ(Count(types, 465U), 0U);
  EXPECT_EQ(Count(types, 463U), 1U);
  EXPECT_EQ(Count(types, 1048U), 1U);
  EXPECT_EQ(Count(types, 1052U), 1U);
  EXPECT_EQ(Count(types, 1050U), 3U);
  EXPECT_EQ(Count(types, 1112U), 3U);
  EXPECT_EQ(Count(types, 1155U), 1U);
  EXPECT_EQ(Count(types, 35U), Count(types, 36U));

  const Workbook back = LoadXlsb(bytes);
  ExpectIconRule(back);
  // The legacy rule keeps its block; the icon rule joins it by range.
  ASSERT_EQ(back.sheet(0).conditional_formats().size(), 1U);
  EXPECT_EQ(back.sheet(0).conditional_formats()[0].rules.size(), 2U);
}

TEST(XlsbCfIconsFidelity, XlsbResaveKeepsTheIcons) {
  const std::vector<std::uint8_t> first = SaveXlsb(BookWithIcons());
  const std::vector<std::uint8_t> second = SaveXlsb(LoadXlsb(first));
  ExpectIconRule(LoadXlsb(second));
  EXPECT_EQ(Count(SheetRecordTypes(second), 1112U), 3U);
}

TEST(XlsbCfIconsFidelity, XlsxRuleSurvivesConversionToXlsb) {
  auto xlsx = write_ooxml(BookWithIcons());
  ASSERT_TRUE(static_cast<bool>(xlsx));
  auto loaded = read_ooxml(SpanOf(xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(loaded));
  ExpectIconRule(loaded.value().workbook);
  ExpectIconRule(LoadXlsb(SaveXlsb(loaded.value().workbook)));
}

TEST(XlsbCfIconsFidelity, EditsOfAnXlsbLoadedRuleAreWritten) {
  Workbook wb = LoadXlsb(SaveXlsb(BookWithIcons()));
  cf::CFRule* rule = nullptr;
  for (cf::ConditionalFormat& f : wb.sheet(0).mutable_conditional_formats()) {
    for (cf::CFRule& r : f.rules) {
      if (r.id == kIconId) {
        rule = &r;
      }
    }
  }
  ASSERT_NE(rule, nullptr);
  rule->icon_set->custom_icons = {{"5Boxes", 4}, {"3Triangles", 0}, {"3Arrows", 1}};
  const Workbook edited = LoadXlsb(SaveXlsb(wb));
  const cf::CFRule* back = FindRule(edited, kIconId);
  ASSERT_NE(back, nullptr);
  ASSERT_EQ(back->icon_set->custom_icons.size(), 3U);
  EXPECT_EQ(back->icon_set->custom_icons[0].icon_set, "5Boxes");
  EXPECT_EQ(back->icon_set->custom_icons[0].icon_id, 4);
  EXPECT_EQ(back->icon_set->custom_icons[1].icon_set, "3Triangles");

  // Without custom icons the rule goes back to a legacy icon set.
  rule->icon_set->custom_icons.clear();
  const std::vector<std::uint8_t> plain = SaveXlsb(wb);
  const std::vector<std::uint16_t> types = SheetRecordTypes(plain);
  EXPECT_EQ(Count(types, 1112U), 0U);
  EXPECT_EQ(Count(types, 465U), 1U);
  const cf::CFRule* legacy = FindRule(LoadXlsb(plain), kIconId);
  ASSERT_NE(legacy, nullptr);
  EXPECT_TRUE(legacy->icon_set->custom_icons.empty());
}

TEST(XlsbCfIconsFidelity, RemovedRuleLeavesNoX14Records) {
  Workbook wb = LoadXlsb(SaveXlsb(BookWithIcons()));
  wb.sheet(0).mutable_conditional_formats()[0].rules.pop_back();
  const std::vector<std::uint16_t> types = SheetRecordTypes(SaveXlsb(wb));
  EXPECT_EQ(Count(types, 1048U) + Count(types, 1046U) + Count(types, 1135U), 0U);
  EXPECT_EQ(Count(types, 463U), 1U);
}

TEST(XlsbCfIconsFidelity, MismatchedIconCountIsRefused) {
  Workbook wb = BookWithIcons();
  wb.sheet(0).mutable_conditional_formats()[0].rules[1].icon_set->custom_icons.pop_back();
  EXPECT_FALSE(static_cast<bool>(write_xlsb(wb)));
}

std::vector<std::uint8_t> ReadFixture(const std::string& file) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + file;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), f) != out.size()) {
    out.clear();
  }
  std::fclose(f);
  return out;
}

Workbook LoadXlsx(const std::vector<std::uint8_t>& bytes) {
  auto r = read_ooxml(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

/// Excel's A1:A10 rule: 3Arrows over 0 / 33 / 67 percent, icons 3Triangles/0,
/// NoIcons/0 and 5Boxes/4.
void ExpectExcelCustomIconSet(const Workbook& wb) {
  const auto& formats = wb.sheet(0).conditional_formats();
  ASSERT_EQ(formats.size(), 1U);
  ASSERT_EQ(formats[0].rules.size(), 1U);
  ASSERT_EQ(formats[0].sqref.size(), 1U);
  EXPECT_EQ(formats[0].sqref[0].last.row, 9U);
  const cf::CFRule& rule = formats[0].rules[0];
  EXPECT_EQ(rule.type, cf::RuleType::IconSet);
  ASSERT_TRUE(rule.icon_set.has_value());
  const cf::IconSetSpec& spec = *rule.icon_set;
  EXPECT_EQ(spec.name, cf::IconSetName::Three_Arrows);
  EXPECT_EQ(spec.floor.type, cf::CfvoType::Percent);
  EXPECT_EQ(spec.floor.value, "0");
  ASSERT_EQ(spec.thresholds.size(), 2U);
  EXPECT_EQ(spec.thresholds[0].type, cf::CfvoType::Percent);
  EXPECT_EQ(spec.thresholds[0].value, "33");
  EXPECT_EQ(spec.thresholds[1].value, "67");
  ASSERT_EQ(spec.custom_icons.size(), 3U);
  EXPECT_EQ(spec.custom_icons[0].icon_set, "3Triangles");
  EXPECT_EQ(spec.custom_icons[0].icon_id, 0);
  EXPECT_EQ(spec.custom_icons[1].icon_set, "NoIcons");
  EXPECT_EQ(spec.custom_icons[1].icon_id, 0);
  EXPECT_EQ(spec.custom_icons[2].icon_set, "5Boxes");
  EXPECT_EQ(spec.custom_icons[2].icon_id, 4);
}

TEST(ExcelCustomIconSet, XlsxOnlyX14RuleLoads) {
  ExpectExcelCustomIconSet(LoadXlsx(ReadFixture("cf_custom_icon_set.xlsx")));
}

TEST(ExcelCustomIconSet, XlsxResaveKeepsOneRule) {
  const Workbook loaded = LoadXlsx(ReadFixture("cf_custom_icon_set.xlsx"));
  auto saved = write_ooxml(loaded);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  ExpectExcelCustomIconSet(LoadXlsx(saved.value()));
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(saved.value()))));
  auto sheet = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(sheet));
  const std::string xml(sheet.value().begin(), sheet.value().end());
  std::size_t x14_rules = 0;
  for (std::size_t at = xml.find("<x14:cfRule"); at != std::string::npos; at = xml.find("<x14:cfRule", at + 1U)) {
    ++x14_rules;
  }
  EXPECT_EQ(x14_rules, 1U) << xml;
  const Workbook via_xlsb = LoadXlsb(SaveXlsb(loaded));
  ExpectExcelCustomIconSet(via_xlsb);
}

TEST(ExcelCustomIconSet, XlsbShadowAndX14RulesLoadAsOneRule) {
  ExpectExcelCustomIconSet(LoadXlsb(ReadFixture("cf_custom_icon_set.xlsb")));
}

TEST(ExcelCustomIconSet, XlsbResaveKeepsOneRule) {
  const Workbook loaded = LoadXlsb(ReadFixture("cf_custom_icon_set.xlsb"));
  ExpectExcelCustomIconSet(LoadXlsb(SaveXlsb(loaded)));
  auto saved = write_ooxml(loaded);
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  ExpectExcelCustomIconSet(LoadXlsx(saved.value()));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
