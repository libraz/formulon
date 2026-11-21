// BrtDVal carries a validation's input-method mode in flag bits 10-17. Excel numbers it noControl 0, on 1, off 2,
// disabled 3, hiragana 4, fullKatakana 5, halfKatakana 6, fullAlpha 7, halfAlpha 8, fullHangul 9, halfHangul 10, and
// the .xlsx form is the `imeMode` attribute, omitted for noControl.

#include <cstdint>
#include <map>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
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

// Saved by Excel from a book whose B1..B11 validations take the eleven modes in schema order (noControl, off, on,
// disabled, ...), B12 a whole-number rule with hiragana, and B13 no mode (merged with B1).
std::vector<std::uint8_t> Fixture() {
  return ReadFixtureFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_dv_ime_mode.xlsb");
}

std::vector<std::uint8_t> Part(const std::vector<std::uint8_t>& package, const std::string& name) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(entry)) << name;
  return entry ? entry.value() : std::vector<std::uint8_t>{};
}

std::vector<std::uint32_t> DValFlags(const std::vector<std::uint8_t>& package) {
  std::vector<std::uint32_t> out;
  const std::vector<std::uint8_t> sheet = Part(package, "xl/worksheets/sheet1.bin");
  for (const FramedRecord& r : split_records(sheet)) {
    if (r.type == kBrtDVal) {
      out.push_back(load_u32(r.payload.data));
    }
  }
  return out;
}

TEST(XlsbDvImeMode, LoadsEveryModeAndWritesTheSameFlagWords) {
  const std::vector<std::uint8_t> source = Fixture();
  auto loaded = read_xlsb(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& wb = loaded.value().workbook;

  std::map<std::uint32_t, std::string> by_row;
  for (const DataValidation& dv : wb.sheet(0).validations()) {
    ASSERT_FALSE(dv.ranges.empty());
    by_row[dv.ranges.front().first_row] = dv.extra_attrs;
  }
  const std::map<std::uint32_t, std::string> expected{
      {0U, ""},
      {1U, " imeMode=\"off\""},
      {2U, " imeMode=\"on\""},
      {3U, " imeMode=\"disabled\""},
      {4U, " imeMode=\"hiragana\""},
      {5U, " imeMode=\"fullKatakana\""},
      {6U, " imeMode=\"halfKatakana\""},
      {7U, " imeMode=\"fullAlpha\""},
      {8U, " imeMode=\"halfAlpha\""},
      {9U, " imeMode=\"fullHangul\""},
      {10U, " imeMode=\"halfHangul\""},
      {11U, " imeMode=\"hiragana\""},
  };
  EXPECT_EQ(by_row, expected);

  auto xlsx = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx));
  const std::vector<std::uint8_t> sheet_xml = Part(xlsx.value(), "xl/worksheets/sheet1.xml");
  EXPECT_NE(std::string(sheet_xml.begin(), sheet_xml.end()).find("imeMode=\"fullKatakana\""), std::string::npos);

  auto written = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  EXPECT_EQ(DValFlags(written.value()), DValFlags(source));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
