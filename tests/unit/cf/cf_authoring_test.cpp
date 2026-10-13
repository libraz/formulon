//
// Icon counts and the derived formulas Excel stores for side-field rule
// types, checked against the Excel-saved `xlsb_feature_flags` and
// `xlsb_feature_text_rules` pairs.

#include "cf/cf_authoring.h"

#include <algorithm>
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
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::cf {
namespace {

std::vector<std::uint8_t> ReadFixture(const std::string& name) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
  }
  std::fclose(file);
  return out;
}

Workbook ReadXlsxFixture(const std::string& name) {
  const std::vector<std::uint8_t> bytes = ReadFixture(name);
  auto result = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  return result ? std::move(result.value().workbook) : Workbook::create_empty();
}

bool IsDerivedType(RuleType t) {
  const auto v = static_cast<std::uint8_t>(t);
  return t == RuleType::TimePeriod || (v >= static_cast<std::uint8_t>(RuleType::ContainsText) &&
                                       v <= static_cast<std::uint8_t>(RuleType::NotContainsErrors));
}

// Payload tails (strParam + formula bytes) of every BrtBeginCFRule in `part`;
// the leading 42 bytes carry the template, dxf, priority and formula sizes.
std::vector<std::vector<std::uint8_t>> CfRuleTails(const std::vector<std::uint8_t>& xlsb, const std::string& part) {
  constexpr std::uint16_t kBeginCfRule = 463;
  constexpr std::size_t kFixedHead = 42;
  std::vector<std::vector<std::uint8_t>> out;
  io::ZipReader zip;
  if (!zip.open(io::ByteSpan{xlsb.data(), xlsb.size()})) {
    ADD_FAILURE() << "zip open failed";
    return out;
  }
  auto body = zip.read_entry(part);
  if (!body) {
    ADD_FAILURE() << "missing part " << part;
    return out;
  }
  io::ByteSpan cursor{body.value().data(), body.value().size()};
  while (cursor.size != 0U) {
    auto rec = io::xlsb::read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << "record walk failed";
      break;
    }
    const io::ByteSpan p = rec.value().payload;
    if (rec.value().type == kBeginCfRule && p.size >= kFixedHead) {
      out.emplace_back(p.data + kFixedHead, p.data + p.size);
    }
  }
  return out;
}

TEST(CfAuthoring, IconCountMatchesTheSetNamePrefix) {
  for (std::size_t i = 0; i < kIconSetNames.size(); ++i) {
    const auto expected = static_cast<std::size_t>(kIconSetNames[i].front() - '0');
    EXPECT_EQ(icon_count(static_cast<IconSetName>(i)), expected) << kIconSetNames[i];
  }
}

// Every derived formula Excel saved in the two fixtures is reproduced from the
// rule's side fields and the top-left cell of its first sqref range.
TEST(CfAuthoring, DerivedFormulasMatchExcelsSavedFormulas) {
  std::size_t checked = 0;
  for (const char* name : {"xlsb_feature_flags.xlsx", "xlsb_feature_text_rules.xlsx"}) {
    const Workbook wb = ReadXlsxFixture(name);
    for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
      for (const ConditionalFormat& block : wb.sheet(s).conditional_formats()) {
        for (const CFRule& rule : block.rules) {
          if (!IsDerivedType(rule.type)) {
            continue;
          }
          ASSERT_TRUE(rule.formula1.has_value()) << name;
          const std::optional<std::string> derived = derived_rule_formula(rule, sqref_anchor(block.sqref));
          ASSERT_TRUE(derived.has_value()) << name << " " << *rule.formula1;
          EXPECT_EQ(*derived, *rule.formula1) << name;
          ++checked;
        }
      }
    }
  }
  // 10 periods in flags; 8 text / blanks / errors rules and 1 period in text_rules.
  EXPECT_EQ(checked, 19U);
}

TEST(CfAuthoring, DerivedFormulaDoublesQuotesInTheText) {
  CFRule rule;
  rule.type = RuleType::BeginsWith;
  rule.text = "a\"b";
  EXPECT_EQ(derived_rule_formula(rule, CellAddress{2, 2}),
            std::optional<std::string>("LEFT(C3,LEN(\"a\"\"b\"))=\"a\"\"b\""));
}

TEST(CfAuthoring, NoDerivedFormulaWithoutItsSideField) {
  CFRule text_rule;
  text_rule.type = RuleType::ContainsText;
  EXPECT_FALSE(derived_rule_formula(text_rule, CellAddress{0, 0}).has_value());
  CFRule period_rule;
  period_rule.type = RuleType::TimePeriod;
  EXPECT_FALSE(derived_rule_formula(period_rule, CellAddress{0, 0}).has_value());
  CFRule cell_is;
  cell_is.type = RuleType::CellIs;
  EXPECT_FALSE(derived_rule_formula(cell_is, CellAddress{0, 0}).has_value());
  CFRule blanks;
  blanks.type = RuleType::ContainsBlanks;
  EXPECT_EQ(derived_rule_formula(blanks, CellAddress{0, 27}), std::optional<std::string>("LEN(TRIM(AB1))=0"));
}

// A period rule carrying only the derived formula saves to the formula bytes
// Excel wrote for the same period.
TEST(CfAuthoring, DerivedPeriodFormulasSaveToExcelsXlsbBytes) {
  Workbook wb = Workbook::create();
  for (std::uint8_t p = 0; p <= static_cast<std::uint8_t>(TimePeriod::NextMonth); ++p) {
    ConditionalFormat block;
    CFCellRange range;
    range.first = CellAddress{p, 1};
    range.last = range.first;
    block.sqref.push_back(range);
    CFRule rule;
    rule.type = RuleType::TimePeriod;
    rule.priority = p + 1;
    rule.time_period = static_cast<TimePeriod>(p);
    rule.formula1 = derived_rule_formula(rule, range.first);
    block.rules.push_back(std::move(rule));
    wb.sheet(0).mutable_conditional_formats().push_back(std::move(block));
  }
  auto written = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  const auto ours = CfRuleTails(written.value(), "xl/worksheets/sheet1.bin");
  const auto excel = CfRuleTails(ReadFixture("xlsb_feature_flags.xlsb"), "xl/worksheets/sheet1.bin");
  ASSERT_EQ(ours.size(), 10U);
  for (std::size_t i = 0; i < ours.size(); ++i) {
    EXPECT_NE(std::find(excel.begin(), excel.end(), ours[i]), excel.end()) << "period " << i;
  }
}

}  // namespace
}  // namespace formulon::cf
