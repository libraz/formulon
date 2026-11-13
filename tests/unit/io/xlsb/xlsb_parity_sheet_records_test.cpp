// The worksheet's BrtPhoneticInfo and closing revision uid, held against Excel 16.115's own xlsb and xlsx saves of the
// same sources (tests/fixtures/excel/xlsb_parity/sheetrecords).

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/sheet_phonetic.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "workbook.h"
#include "writer_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBrtPhoneticInfo = 537;
constexpr std::uint16_t kBrtUid = 3072;
constexpr std::uint16_t kBrtEndSheet = 130;

using Bytes = std::vector<std::uint8_t>;

Bytes ReadFile(const std::string& path) {
  Bytes out;
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
    ADD_FAILURE() << "short read on fixture: " << path;
  }
  std::fclose(f);
  return out;
}

std::string Fixture(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/sheetrecords/" + name;
}

Bytes Part(const Bytes& package, const std::string& path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto bytes = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(bytes)) << path;
  return bytes ? std::move(bytes.value()) : Bytes{};
}

Bytes FromXlsx(const std::string& stem) {
  auto loaded = read_ooxml(SpanOf(ReadFile(Fixture(stem + "_src.xlsx"))));
  EXPECT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  if (!loaded) {
    return {};
  }
  auto written = write_xlsb_with_result(loaded.value().workbook);
  EXPECT_TRUE(static_cast<bool>(written)) << written.error().message;
  return written ? std::move(written.value().bytes) : Bytes{};
}

std::vector<Bytes> Payloads(const Bytes& part, std::uint16_t type) {
  std::vector<Bytes> out;
  for (const FramedRecord& rec : split_records(part)) {
    if (rec.type == type) {
      out.emplace_back(rec.payload.data, rec.payload.data + rec.payload.size);
    }
  }
  return out;
}

/// The BrtUid of the wrapper that closes the worksheet, empty when there is none.
Bytes SheetUid(const Bytes& part) {
  const std::vector<FramedRecord> recs = split_records(part);
  const std::size_t n = recs.size();
  if (n < 3U || recs[n - 1U].type != kBrtEndSheet || recs[n - 3U].type != kBrtUid) {
    return {};
  }
  return Bytes(recs[n - 3U].payload.data, recs[n - 3U].payload.data + recs[n - 3U].payload.size);
}

std::string SheetXml(const Bytes& package, int sheet) {
  const Bytes xml = Part(package, "xl/worksheets/sheet" + std::to_string(sheet) + ".xml");
  return std::string(xml.begin(), xml.end());
}

const char* const kPhoneticCases[] = {
    "phon_font1",         "phon_type_halfwidthKatakana", "phon_type_fullwidthKatakana",
    "phon_type_Hiragana", "phon_type_noConversion",      "phon_align_noControl",
    "phon_align_left",    "phon_align_center",           "phon_align_distributed",
    "phon_both",
};

TEST(XlsbParitySheetRecords, PhoneticInfoMatchesExcel) {
  for (const char* stem : kPhoneticCases) {
    SCOPED_TRACE(stem);
    const auto excel =
        Payloads(Part(ReadFile(Fixture(std::string(stem) + ".xlsb")), "xl/worksheets/sheet1.bin"), kBrtPhoneticInfo);
    ASSERT_EQ(excel.size(), 1U);
    EXPECT_EQ(Payloads(Part(FromXlsx(stem), "xl/worksheets/sheet1.bin"), kBrtPhoneticInfo), excel);
  }
}

// Excel gives a sheet with no phoneticPr a new phonetic font and points the record at it; the writer does not add
// fonts, so the sheet stays without the record and reads back without the element.
TEST(XlsbParitySheetRecords, AbsentPhoneticWritesNoRecord) {
  EXPECT_TRUE(Payloads(Part(FromXlsx("phon_absent"), "xl/worksheets/sheet1.bin"), kBrtPhoneticInfo).empty());
}

// Excel points a phoneticPr whose fontId lies past the font table at font 0 when it loads the book, keeping type and
// alignment; the readers do the same, so both writers match Excel's saves.
TEST(XlsbParitySheetRecords, OutOfRangeFontClampsAtLoad) {
  const Bytes excel_xlsx = ReadFile(Fixture("phon_font_out_of_range.xlsx"));
  const Bytes excel_xlsb = ReadFile(Fixture("phon_font_out_of_range.xlsb"));
  EXPECT_EQ(Payloads(Part(FromXlsx("phon_font_out_of_range"), "xl/worksheets/sheet1.bin"), kBrtPhoneticInfo),
            Payloads(Part(excel_xlsb, "xl/worksheets/sheet1.bin"), kBrtPhoneticInfo));

  auto loaded = read_ooxml(SpanOf(ReadFile(Fixture("phon_font_out_of_range_src.xlsx"))));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  auto xlsx = write_ooxml(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  const auto element = [](const std::string& xml) {
    const std::size_t at = xml.find("<phoneticPr");
    return at == std::string::npos ? std::string() : xml.substr(at, xml.find("/>", at) + 2U - at);
  };
  EXPECT_EQ(element(SheetXml(xlsx.value(), 1)), "<phoneticPr fontId=\"0\" type=\"Hiragana\"/>");
  EXPECT_EQ(element(SheetXml(xlsx.value(), 1)), element(SheetXml(excel_xlsx, 1)));

  // The xlsb reader applies the same rule to a record naming a font the book does not have.
  Workbook& wb = loaded.value().workbook;
  set_sheet_phonetic_properties(wb.sheet(0), PhoneticProperties{9U, 2U, 3U});
  auto xlsb = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  auto reread = read_xlsb(SpanOf(xlsb.value().bytes));
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;
  const auto props = sheet_phonetic_properties(reread.value().workbook.sheet(0));
  ASSERT_TRUE(props.has_value());
  EXPECT_EQ(props->font_id, 0U);
  EXPECT_EQ(props->type, 2U);
  EXPECT_EQ(props->alignment, 3U);
}

// A stated uid is written as stated; a sheet without one gets the uid Excel assigns from its position.
TEST(XlsbParitySheetRecords, SheetUidMatchesExcel) {
  const Bytes ours = FromXlsx("uid_second_only");
  const Bytes excel = ReadFile(Fixture("uid_second_only.xlsb"));
  for (int sheet = 1; sheet <= 3; ++sheet) {
    SCOPED_TRACE(sheet);
    const std::string part = "xl/worksheets/sheet" + std::to_string(sheet) + ".bin";
    ASSERT_EQ(SheetUid(Part(excel, part)).size(), 16U);
    EXPECT_EQ(SheetUid(Part(ours, part)), SheetUid(Part(excel, part)));
  }
}

// Excel's xlsb reads into the model and writes back both records unchanged, and as xlsx carries the same phoneticPr
// and xr:uid that Excel's own xlsx save of the source has.
TEST(XlsbParitySheetRecords, ExcelXlsbRoundTrips) {
  for (const char* stem : {"phon_both", "phon_type_halfwidthKatakana", "uid_second_only"}) {
    SCOPED_TRACE(stem);
    const Bytes excel = ReadFile(Fixture(std::string(stem) + ".xlsb"));
    const Bytes excel_xlsx = ReadFile(Fixture(std::string(stem) + ".xlsx"));
    auto loaded = read_xlsb(SpanOf(excel));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const Workbook& wb = loaded.value().workbook;
    auto xlsb = write_xlsb_with_result(wb);
    ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
    auto xlsx = write_ooxml(wb);
    ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
    for (int sheet = 1; sheet <= static_cast<int>(wb.sheet_count()); ++sheet) {
      SCOPED_TRACE(sheet);
      const std::string part = "xl/worksheets/sheet" + std::to_string(sheet) + ".bin";
      EXPECT_EQ(Payloads(Part(xlsb.value().bytes, part), kBrtPhoneticInfo),
                Payloads(Part(excel, part), kBrtPhoneticInfo));
      EXPECT_EQ(SheetUid(Part(xlsb.value().bytes, part)), SheetUid(Part(excel, part)));
      const std::string ours_xml = SheetXml(xlsx.value(), sheet);
      const std::string excel_xml = SheetXml(excel_xlsx, sheet);
      const auto element = [](const std::string& xml) {
        const std::size_t at = xml.find("<phoneticPr");
        return at == std::string::npos ? std::string() : xml.substr(at, xml.find("/>", at) + 2U - at);
      };
      const auto uid = [](const std::string& xml) {
        const std::size_t at = xml.find(" xr:uid=\"");
        return at == std::string::npos ? std::string() : xml.substr(at, 48U);
      };
      EXPECT_EQ(element(ours_xml), element(excel_xml));
      EXPECT_FALSE(uid(excel_xml).empty());
      EXPECT_EQ(uid(ours_xml), uid(excel_xml));
    }
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
