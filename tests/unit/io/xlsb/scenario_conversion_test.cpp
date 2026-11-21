//
// Scenarios across formats, against Mac Excel 365 xlsx/xlsb pairs of one book: the `<scenarios>` element and the
// BrtBeginScenMan block hold the same scenarios. The pairs cover scenario flags, comment and user;
// input-cell values and number formats; input-cell deleted / undone; the manager's absent current, show and
// sqref; several sqref ranges; several scenarios with mixed flags; non-ASCII text and a row past 65535.

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/writer.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io::xlsb {
namespace {

namespace tc = test;

constexpr const char* kSheetBin = "xl/worksheets/sheet1.bin";
constexpr const char* kSheetXml = "xl/worksheets/sheet1.xml";

const char* const kBooks[] = {
    "excel/retained_children_edit/children_0_base",
    "excel/scenarios/scenario_flags",
    "excel/scenarios/input_values",
    "excel/scenarios/input_flags",
    "excel/scenarios/input_all_flags",
    "excel/scenarios/no_manager_attrs",
    "excel/scenarios/no_show_two_ranges",
    "excel/scenarios/current_without_sqref",
    "excel/scenarios/locked_and_hidden",
    "excel/scenarios/three_scenarios",
    "excel/scenarios/wide_values",
    "excel/scenarios/show_without_current",
};

std::string Fixture(const std::string& stem, const char* ext) {
  return std::string(FORMULON_FIXTURES_DIR) + "/" + stem + "." + ext;
}

/// The scenario block of `package`'s sheet, with the type of the record before it.
struct ScenarioBlock {
  std::uint16_t before = 0;
  std::vector<tc::IndexedRecord> records;
};

ScenarioBlock ScenariosOf(const tc::Bytes& package) {
  ScenarioBlock out;
  const std::optional<tc::Bytes> part = tc::ReadPart(package, kSheetBin);
  EXPECT_TRUE(part.has_value());
  if (!part) {
    return out;
  }
  const std::vector<tc::IndexedRecord> records = tc::SplitPart(*part);
  for (std::size_t i = 0; i < records.size(); ++i) {
    const std::uint16_t type = records[i].type;
    if (type < kBrtBeginScenMan || type > kBrtSlc) {
      continue;
    }
    if (out.records.empty() && i > 0U) {
      out.before = records[i - 1U].type;
    }
    out.records.push_back(records[i]);
  }
  return out;
}

void ExpectSameBlock(const ScenarioBlock& ours, const ScenarioBlock& excel) {
  ASSERT_GE(excel.records.size(), 3U);
  ASSERT_EQ(ours.records.size(), excel.records.size());
  EXPECT_EQ(ours.before, excel.before);
  for (std::size_t i = 0; i < excel.records.size(); ++i) {
    EXPECT_EQ(ours.records[i].type, excel.records[i].type) << i;
    EXPECT_EQ(ours.records[i].payload, excel.records[i].payload) << i;
  }
}

/// The `<scenarios>` element of a sheet part's text; empty when it has none.
std::string ScenariosElement(const tc::Bytes& part) {
  const std::string text(part.begin(), part.end());
  const std::size_t from = text.find("<scenarios");
  const std::size_t to = text.find("</scenarios>");
  if (from == std::string::npos || to == std::string::npos) {
    return {};
  }
  return text.substr(from, to + std::string("</scenarios>").size() - from);
}

class XlsbScenarioConversion : public ::testing::TestWithParam<const char*> {};

TEST_P(XlsbScenarioConversion, XlsxElementWritesAsExcelsRecords) {
  const std::string stem = GetParam();
  const tc::Converted converted = tc::ConvertToXlsb(std::string(FORMULON_FIXTURES_DIR), stem + ".xlsx");
  ASSERT_TRUE(converted.error.empty()) << converted.error;
  ExpectSameBlock(ScenariosOf(converted.bytes), ScenariosOf(tc::ReadBytes(Fixture(stem, "xlsb"))));
}

TEST_P(XlsbScenarioConversion, XlsbRecordsWriteBackUnchanged) {
  const std::string stem = GetParam();
  const tc::Converted converted = tc::ConvertToXlsb(std::string(FORMULON_FIXTURES_DIR), stem + ".xlsb");
  ASSERT_TRUE(converted.error.empty()) << converted.error;
  ExpectSameBlock(ScenariosOf(converted.bytes), ScenariosOf(tc::ReadBytes(Fixture(stem, "xlsb"))));
}

TEST_P(XlsbScenarioConversion, XlsbRecordsSaveAsExcelsXlsxElement) {
  const std::string stem = GetParam();
  const tc::Bytes xlsb = tc::ReadBytes(Fixture(stem, "xlsb"));
  auto read = read_xlsb(ByteSpan{xlsb.data(), xlsb.size()});
  ASSERT_TRUE(read.has_value()) << read.error().message;
  auto written = write_ooxml(read.value().workbook);
  ASSERT_TRUE(written.has_value()) << written.error().message;
  const std::optional<tc::Bytes> ours = tc::ReadPart(written.value(), kSheetXml);
  const std::optional<tc::Bytes> excel = tc::ReadPart(tc::ReadBytes(Fixture(stem, "xlsx")), kSheetXml);
  ASSERT_TRUE(ours.has_value());
  ASSERT_TRUE(excel.has_value());
  const std::string excel_element = ScenariosElement(*excel);
  ASSERT_FALSE(excel_element.empty());
  EXPECT_EQ(ScenariosElement(*ours), excel_element);
}

INSTANTIATE_TEST_SUITE_P(ExcelPairs, XlsbScenarioConversion, ::testing::ValuesIn(kBooks));

}  // namespace
}  // namespace formulon::io::xlsb
