// Row height provenance: an explicit (custom) height is distinct from an auto
// height that only caches `ht`, in OOXML (`customHeight`) and XLSB
// (`BrtRowHdr` fUnsynced).

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "sheet_layout.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

namespace formulon {
namespace {

/// Saves `wb` as `.xlsx` and reads it back.
::testing::AssertionResult ThroughXlsx(const Workbook& wb, Workbook* out) {
  auto saved = io::write_ooxml(wb);
  if (!saved) {
    return ::testing::AssertionFailure() << "write_ooxml failed: " << saved.error().message;
  }
  auto reloaded = io::read_ooxml(test::span_of(saved.value()));
  if (!reloaded) {
    return ::testing::AssertionFailure() << "read_ooxml failed: " << reloaded.error().message;
  }
  *out = std::move(reloaded.value().workbook);
  return ::testing::AssertionSuccess();
}

/// Saves `wb` as `.xlsb` and reads it back.
::testing::AssertionResult ThroughXlsb(const Workbook& wb, Workbook* out) {
  auto saved = io::xlsb::write_xlsb(wb);
  if (!saved) {
    return ::testing::AssertionFailure() << "write_xlsb failed: " << saved.error().message;
  }
  auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
  if (!reloaded) {
    return ::testing::AssertionFailure() << "read_xlsb failed: " << reloaded.error().message;
  }
  *out = std::move(reloaded.value().workbook);
  return ::testing::AssertionSuccess();
}

constexpr const char* kSheetPart = "xl/worksheets/sheet1.xml";

RowLayout MakeRow(std::uint32_t row, double height, bool custom) {
  RowLayout r;
  r.row = row;
  r.height = height;
  r.has_height = true;
  r.custom_height = custom;
  return r;
}

Workbook MakeWorkbook() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  auto& rows = wb.sheet(0).mutable_layout().row_overrides;
  rows.push_back(MakeRow(0, 30.0, true));
  rows.push_back(MakeRow(1, 18.0, false));
  return wb;
}

const RowLayout* FindRow(const Workbook& wb, std::uint32_t row) {
  for (const RowLayout& r : wb.sheet(0).layout().row_overrides) {
    if (r.row == row) {
      return &r;
    }
  }
  return nullptr;
}

void ExpectCustomAndAuto(const Workbook& wb) {
  const RowLayout* custom = FindRow(wb, 0);
  const RowLayout* autoh = FindRow(wb, 1);
  ASSERT_NE(custom, nullptr);
  ASSERT_NE(autoh, nullptr);
  EXPECT_TRUE(custom->has_height);
  EXPECT_TRUE(custom->custom_height);
  EXPECT_DOUBLE_EQ(custom->height, 30.0);
  EXPECT_TRUE(autoh->has_height);
  EXPECT_FALSE(autoh->custom_height);
  EXPECT_DOUBLE_EQ(autoh->height, 18.0);
}

TEST(RowCustomHeight, WriterEmitsCustomHeightOnlyForCustomRows) {
  Workbook wb = MakeWorkbook();
  auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(saved.has_value());
  std::string xml;
  ASSERT_TRUE(test::extract_part(test::span_of(saved.value()), kSheetPart, &xml));
  EXPECT_NE(xml.find("ht=\"30\" customHeight=\"1\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("ht=\"18\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("ht=\"18\" customHeight"), std::string::npos) << xml;
}

TEST(RowCustomHeight, OoxmlRoundTripKeepsCustomAndAuto) {
  Workbook wb = MakeWorkbook();
  Workbook once = Workbook::create_empty();
  ASSERT_TRUE(ThroughXlsx(wb, &once));
  ExpectCustomAndAuto(once);
  Workbook twice = Workbook::create_empty();
  ASSERT_TRUE(ThroughXlsx(once, &twice));
  ExpectCustomAndAuto(twice);
}

TEST(RowCustomHeight, SaxPathReadsCustomHeight) {
  Workbook wb = MakeWorkbook();
  auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(saved.has_value());
  auto read = io::internal::ReadOoxmlWithSaxThresholdForTesting(test::span_of(saved.value()), 0U);
  ASSERT_TRUE(read.has_value());
  ExpectCustomAndAuto(read.value().workbook);
}

TEST(RowCustomHeight, HtWithoutCustomHeightStaysAuto) {
  Workbook wb = MakeWorkbook();
  Workbook once = Workbook::create_empty();
  ASSERT_TRUE(ThroughXlsx(wb, &once));
  Workbook twice = Workbook::create_empty();
  ASSERT_TRUE(ThroughXlsx(once, &twice));
  const RowLayout* autoh = FindRow(twice, 1);
  ASSERT_NE(autoh, nullptr);
  EXPECT_TRUE(autoh->has_height);
  EXPECT_FALSE(autoh->custom_height);
}

TEST(RowCustomHeight, XlsbUnsyncedRoundTrips) {
  Workbook wb = MakeWorkbook();
  Workbook back = Workbook::create_empty();
  ASSERT_TRUE(ThroughXlsb(wb, &back));
  const RowLayout* custom = FindRow(back, 0);
  ASSERT_NE(custom, nullptr);
  EXPECT_TRUE(custom->has_height);
  EXPECT_TRUE(custom->custom_height);
  EXPECT_DOUBLE_EQ(custom->height, 30.0);
  // fUnsynced clear: the row reads back without a custom override.
  const RowLayout* autoh = FindRow(back, 1);
  EXPECT_TRUE(autoh == nullptr || !autoh->custom_height);
}

}  // namespace
}  // namespace formulon
