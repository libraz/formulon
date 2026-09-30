//
// Spilled (non-anchor) cells against a Mac Excel 365-produced workbook,
// saved as both `tests/fixtures/excel/spill_phantom_cells.xlsx` and `.xlsb`.
//
// Sheet1: A1:A4 = 1, "x", TRUE, =1/0; spills of numbers (C1 =SEQUENCE(3)),
// text (E1 ={"a","b";"c","d"}), booleans (H1), a mix with an error (J1),
// a reference over mixed cells (L1 =A1:A4), a volatile anchor
// (N1 =SEQUENCE(2)+RAND()*0), a styled anchor (P1, with P3 styled on its
// own) and a spill over a pre-styled blank (R1, R2 styled).
//
// Each file is loaded, optionally recalculated, and saved in both formats;
// every formula and spilled cell must come out as Excel wrote it. A loaded
// spilled value that turned into a constant would block the recalculated
// spill and show up here as a changed cell.

#include <algorithm>
#include <cstdint>
#include <cstdio>
#include <map>
#include <string>
#include <tuple>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFixture(const char* file) {
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
    ADD_FAILURE() << "short read on fixture: " << path;
    out.clear();
  }
  std::fclose(f);
  return out;
}

Workbook Load(const std::vector<std::uint8_t>& bytes, bool xlsb) {
  if (xlsb) {
    auto r = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
    EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
    return r ? std::move(r.value().workbook) : Workbook::create_empty();
  }
  auto r = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

/// Every `<c>` element of a sheet part keyed by its `r`, attributes sorted,
/// and the text escapes that do not change the XML dropped.
std::map<std::string, std::string> Cells(const std::string& xml) {
  std::map<std::string, std::string> out;
  std::size_t at = 0;
  while ((at = xml.find("<c r=\"", at)) != std::string::npos) {
    const std::size_t end_self = xml.find("/>", at);
    const std::size_t end_open = xml.find('>', at);
    const bool self_closing = end_self == end_open - 1U;
    const std::size_t end = self_closing ? end_open + 1U : xml.find("</c>", at) + 4U;
    std::string tag = xml.substr(at + 3U, end_open - at - 3U - (self_closing ? 1U : 0U));
    std::vector<std::string> attrs;
    std::size_t p = 0;
    while (p < tag.size()) {
      const std::size_t q = tag.find('"', tag.find('"', p) + 1U);
      attrs.push_back(tag.substr(p, q + 1U - p));
      p = q + 2U;
    }
    const std::string address = attrs.front();
    std::sort(attrs.begin(), attrs.end());
    std::string body = self_closing ? std::string() : xml.substr(end_open + 1U, end - end_open - 1U);
    for (const std::string& drop : {std::string(" xml:space=\"preserve\"")}) {
      for (std::size_t d; (d = body.find(drop)) != std::string::npos;) {
        body.erase(d, drop.size());
      }
    }
    for (std::size_t d; (d = body.find("&quot;")) != std::string::npos;) {
      body.replace(d, 6U, "\"");
    }
    std::string key = address.substr(3U, address.size() - 4U);
    std::string value;
    for (const std::string& a : attrs) {
      value += a + " ";
    }
    out[key] = value + body;
    at = end;
  }
  return out;
}

/// Formula-cell (`BrtFmla*`) and `BrtArrFmla` records of a sheet part keyed
/// by address, with the fields Excel leaves uninitialised cleared: the last
/// four bytes of a `PtgArray` and the `u16` of a leading `PtgAttrSemi`.
std::map<std::string, std::vector<std::uint8_t>> FormulaRecords(const std::vector<std::uint8_t>& part) {
  std::map<std::string, std::vector<std::uint8_t>> out;
  io::ByteSpan cursor{part.data(), part.size()};
  std::uint32_t row = 0;
  std::string last;
  auto u32 = [](const std::uint8_t* d) {
    return static_cast<std::uint32_t>(d[0]) | (static_cast<std::uint32_t>(d[1]) << 8) |
           (static_cast<std::uint32_t>(d[2]) << 16) | (static_cast<std::uint32_t>(d[3]) << 24);
  };
  while (cursor.size > 0U) {
    auto record = io::xlsb::read_record(cursor);
    if (!record) {
      ADD_FAILURE() << "unreadable record";
      break;
    }
    const io::ByteSpan p = record.value().payload;
    const std::uint16_t type = record.value().type;
    if (type == static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtRowHdr)) {
      row = u32(p.data);
    } else if (type >= static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtFmlaString) &&
               type <= static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtFmlaError)) {
      last = std::to_string(row) + ":" + std::to_string(u32(p.data));
      out[last] = std::vector<std::uint8_t>(p.data, p.data + p.size);
    } else if (type == static_cast<std::uint16_t>(io::xlsb::XlsbRecordType::BrtArrFmla)) {
      std::vector<std::uint8_t> bytes(p.data, p.data + p.size);
      const std::size_t rgce = 21U;
      if (bytes.size() > rgce + 15U && (bytes[rgce] & 0x1FU) == 0x00U && (bytes[rgce] & 0x60U) != 0U) {
        std::fill(bytes.begin() + rgce + 11, bytes.begin() + rgce + 15, 0);  // PtgArray
      }
      if (bytes.size() > rgce + 4U && bytes[rgce] == 0x19 && bytes[rgce + 1U] == 0x01) {
        bytes[rgce + 2U] = bytes[rgce + 3U] = 0;  // PtgAttrSemi
      }
      out[last + "#arr"] = std::move(bytes);
    }
  }
  return out;
}

class SpillPhantomFixture : public ::testing::TestWithParam<std::tuple<bool, bool>> {};

TEST_P(SpillPhantomFixture, SpilledCellsSaveAsExcelWritesThem) {
  const auto [from_xlsb, recalc] = GetParam();
  Workbook wb = Load(ReadFixture(from_xlsb ? "spill_phantom_cells.xlsb" : "spill_phantom_cells.xlsx"), from_xlsb);
  if (recalc) {
    (void)wb.recalc(eval::default_registry());
  }

  // .xlsx: every cell as Excel wrote it, except A2, whose shared-string
  // index depends on the table's order.
  std::string excel_xml;
  const std::vector<std::uint8_t> excel_xlsx = ReadFixture("spill_phantom_cells.xlsx");
  ASSERT_TRUE(test::extract_part(test::span_of(excel_xlsx), "xl/worksheets/sheet1.xml", &excel_xml));
  auto saved_xlsx = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved_xlsx));
  std::string ours_xml;
  ASSERT_TRUE(test::extract_part(test::span_of(saved_xlsx.value()), "xl/worksheets/sheet1.xml", &ours_xml));
  std::map<std::string, std::string> excel_cells = Cells(excel_xml);
  std::map<std::string, std::string> ours_cells = Cells(ours_xml);
  excel_cells.erase("A2");
  ours_cells.erase("A2");
  EXPECT_EQ(excel_cells.size(), 27U);
  EXPECT_EQ(ours_cells, excel_cells);

  // .xlsb: every formula and spilled cell record as Excel wrote it.
  std::vector<std::uint8_t> excel_part;
  const std::vector<std::uint8_t> excel_xlsb = ReadFixture("spill_phantom_cells.xlsb");
  io::ZipReader excel_zip;
  ASSERT_TRUE(static_cast<bool>(excel_zip.open(io::ByteSpan{excel_xlsb.data(), excel_xlsb.size()})));
  auto excel_sheet = excel_zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(excel_sheet));
  auto saved_xlsb = io::xlsb::write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(saved_xlsb));
  io::ZipReader ours_zip;
  ASSERT_TRUE(
      static_cast<bool>(ours_zip.open(io::ByteSpan{saved_xlsb.value().bytes.data(), saved_xlsb.value().bytes.size()})));
  auto ours_sheet = ours_zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(ours_sheet));
  const auto excel_records = FormulaRecords(excel_sheet.value());
  EXPECT_EQ(excel_records.size(), 33U);
  EXPECT_EQ(FormulaRecords(ours_sheet.value()), excel_records);
}

INSTANTIATE_TEST_SUITE_P(FromBothFormats, SpillPhantomFixture,
                         ::testing::Combine(::testing::Bool(), ::testing::Bool()));

}  // namespace
}  // namespace formulon
