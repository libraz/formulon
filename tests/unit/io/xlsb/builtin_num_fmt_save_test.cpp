//
// Built-in number formats a save declares with their code, against Mac Excel 365 (ja-JP) xlsx/xlsb pairs: Excel's own
// save writes ids 5-8 as `<numFmt>` / BrtFmt with the stored currency code, so both writers keep them in both formats.

#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/styles_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/styles_writer.h"
#include "pugixml.hpp"
#include "styles.h"
#include "xlsb_twin_compare.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io::xlsb {
namespace {

namespace tc = test;

constexpr const char* kStem = "cell_style_save/cs1_C1_all_blank_resaved";
constexpr auto kBrtFmt = static_cast<std::uint16_t>(XlsbRecordType::BrtFmt);

using Formats = std::vector<std::pair<std::uint32_t, std::string>>;

std::string Fixture(const char* ext) {
  return std::string(FORMULON_FIXTURES_DIR) + "/" + kStem + "." + ext;
}

/// The `<numFmts>` entries of a package's styles part, in order.
Formats XmlFormatsOf(const tc::Bytes& package) {
  Formats out;
  const std::optional<tc::Bytes> part = tc::ReadPart(package, "xl/styles.xml");
  EXPECT_TRUE(part.has_value());
  if (!part) {
    return out;
  }
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_buffer(part->data(), part->size()));
  for (const pugi::xml_node fmt : doc.child("styleSheet").child("numFmts").children("numFmt")) {
    out.emplace_back(fmt.attribute("numFmtId").as_uint(), fmt.attribute("formatCode").value());
  }
  return out;
}

/// The BrtFmt records of a package's styles part, in order.
std::vector<tc::IndexedRecord> BinFormatsOf(const tc::Bytes& package) {
  std::vector<tc::IndexedRecord> out;
  const std::optional<tc::Bytes> part = tc::ReadPart(package, "xl/styles.bin");
  EXPECT_TRUE(part.has_value());
  if (!part) {
    return out;
  }
  for (const tc::IndexedRecord& rec : tc::SplitPart(*part)) {
    if (rec.type == kBrtFmt) {
      out.push_back(rec);
    }
  }
  return out;
}

void ExpectSameBinFormats(const std::vector<tc::IndexedRecord>& ours, const std::vector<tc::IndexedRecord>& excel) {
  ASSERT_EQ(excel.size(), 4U);
  ASSERT_EQ(ours.size(), excel.size());
  for (std::size_t i = 0; i < excel.size(); ++i) {
    EXPECT_EQ(ours[i].payload, excel[i].payload) << i;
  }
}

TEST(BuiltinNumFmtSave, XlsxKeepsExcelsDeclaredBuiltins) {
  const tc::Bytes xlsx = tc::ReadBytes(Fixture("xlsx"));
  auto read = read_ooxml(ByteSpan{xlsx.data(), xlsx.size()});
  ASSERT_TRUE(read.has_value()) << read.error().message;
  auto written = write_ooxml(read.value().workbook);
  ASSERT_TRUE(written.has_value()) << written.error().message;
  const Formats excel = XmlFormatsOf(xlsx);
  ASSERT_EQ(excel.size(), 4U);
  EXPECT_EQ(XmlFormatsOf(written.value()), excel);
}

TEST(BuiltinNumFmtSave, XlsbSavesExcelsXlsxFormats) {
  const tc::Bytes xlsb = tc::ReadBytes(Fixture("xlsb"));
  auto read = read_xlsb(ByteSpan{xlsb.data(), xlsb.size()});
  ASSERT_TRUE(read.has_value()) << read.error().message;
  auto written = write_ooxml(read.value().workbook);
  ASSERT_TRUE(written.has_value()) << written.error().message;
  const Formats excel = XmlFormatsOf(tc::ReadBytes(Fixture("xlsx")));
  ASSERT_EQ(excel.size(), 4U);
  EXPECT_EQ(XmlFormatsOf(written.value()), excel);
}

TEST(BuiltinNumFmtSave, XlsbKeepsExcelsDeclaredBuiltins) {
  const tc::Converted converted = tc::ConvertToXlsb(std::string(FORMULON_FIXTURES_DIR), std::string(kStem) + ".xlsb");
  ASSERT_TRUE(converted.error.empty()) << converted.error;
  ExpectSameBinFormats(BinFormatsOf(converted.bytes), BinFormatsOf(tc::ReadBytes(Fixture("xlsb"))));
}

TEST(BuiltinNumFmtSave, XlsxSavesExcelsXlsbRecords) {
  const tc::Converted converted = tc::ConvertToXlsb(std::string(FORMULON_FIXTURES_DIR), std::string(kStem) + ".xlsx");
  ASSERT_TRUE(converted.error.empty()) << converted.error;
  ExpectSameBinFormats(BinFormatsOf(converted.bytes), BinFormatsOf(tc::ReadBytes(Fixture("xlsb"))));
}

TEST(BuiltinNumFmtSave, OtherBuiltinIdsStayBare) {
  StylesTable table;
  table.num_fmt_strings = {"0.00%", "\"$\"#,##0"};
  table.num_fmts = {NumFmtRecord{14U, 0U}, NumFmtRecord{5U, 1U}};
  const std::string xml = write_styles(table);
  EXPECT_EQ(xml.find("numFmtId=\"14\" formatCode"), std::string::npos);
  EXPECT_NE(xml.find("numFmtId=\"5\" formatCode"), std::string::npos);
  std::vector<std::uint16_t> ids;
  const std::vector<std::uint8_t> bin = write_styles_bin(table);
  for (const tc::IndexedRecord& rec : tc::SplitPart(tc::Bytes(bin.begin(), bin.end()))) {
    if (rec.type == kBrtFmt) {
      ids.push_back(static_cast<std::uint16_t>(rec.payload[0] | (rec.payload[1] << 8U)));
    }
  }
  EXPECT_EQ(ids, (std::vector<std::uint16_t>{5U}));
}

}  // namespace
}  // namespace formulon::io::xlsb
