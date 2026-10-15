// The DOM and SAX sheet readers lex `<c cm=>` the same way: a non-numeric or
// out-of-range value degrades to 0.

#include <cstdint>
#include <deque>
#include <string>
#include <tuple>
#include <vector>

#include "gtest/gtest.h"
#include "io/sheet_reader.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

using Metadata = std::vector<std::tuple<std::uint32_t, std::uint32_t, std::uint32_t>>;

std::string SheetXml(const std::string& cm) {
  return "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
         "<row r=\"1\"><c r=\"A1\" cm=\"" +
         cm + "\"><f>1+1</f><v>2</v></c></row></sheetData></worksheet>";
}

Workbook Host() {
  Workbook wb = Workbook::create_empty();
  wb.set_styles(StylesTable{});
  wb.add_sheet("Sheet1");
  return wb;
}

Metadata ViaDom(const std::string& xml) {
  pugi::xml_document doc;
  const std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  EXPECT_TRUE(load_xml_buffer(doc, bytes, "cm_parity", "sheet.xml"));
  Workbook wb = Host();
  SheetReadContext ctx;
  auto rs = read_sheet_data(doc, 0U, wb, ctx, wb.mutable_text_storage());
  EXPECT_TRUE(static_cast<bool>(rs));
  return ctx.cell_metadata;
}

Metadata ViaSax(const std::string& xml) {
  Workbook wb = Host();
  SheetReadContext ctx;
  const ByteSpan span{reinterpret_cast<const std::uint8_t*>(xml.data()), xml.size()};
  auto rs = read_sheet_data_sax(span, 0U, wb, ctx, wb.mutable_text_storage());
  EXPECT_TRUE(static_cast<bool>(rs));
  return ctx.cell_metadata;
}

void ExpectBothRead(const std::string& cm, std::uint32_t expected) {
  const std::string xml = SheetXml(cm);
  const Metadata dom = ViaDom(xml);
  const Metadata sax = ViaSax(xml);
  ASSERT_EQ(dom.size(), 1U) << cm;
  ASSERT_EQ(sax.size(), 1U) << cm;
  EXPECT_EQ(std::get<2>(dom[0]), expected) << "dom cm=" << cm;
  EXPECT_EQ(std::get<2>(sax[0]), expected) << "sax cm=" << cm;
}

TEST(SheetReaderCellMetadataParity, PlainIndexIsKept) {
  ExpectBothRead("1", 1U);
}

TEST(SheetReaderCellMetadataParity, EmbeddedNonDigitDegradesToZero) {
  ExpectBothRead("1a2", 0U);
}

TEST(SheetReaderCellMetadataParity, OutOfRangeValueDegradesToZero) {
  ExpectBothRead("99999999999", 0U);
}

}  // namespace
}  // namespace io
}  // namespace formulon
