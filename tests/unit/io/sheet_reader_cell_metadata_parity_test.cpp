// The DOM and SAX sheet readers lex `<c cm=>` and `<c vm=>` the same way: a
// non-numeric or out-of-range value degrades to 0. `cm` is kept on formula
// cells only, `vm` on any cell.

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

SheetReadContext ContextViaDom(const std::string& xml) {
  pugi::xml_document doc;
  const std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
  EXPECT_TRUE(load_xml_buffer(doc, bytes, "cm_parity", "sheet.xml"));
  Workbook wb = Host();
  SheetReadContext ctx;
  auto rs = read_sheet_data(doc, 0U, wb, ctx, wb.mutable_text_storage());
  EXPECT_TRUE(static_cast<bool>(rs));
  return ctx;
}

SheetReadContext ContextViaSax(const std::string& xml) {
  Workbook wb = Host();
  SheetReadContext ctx;
  const ByteSpan span{reinterpret_cast<const std::uint8_t*>(xml.data()), xml.size()};
  auto rs = read_sheet_data_sax(span, 0U, wb, ctx, wb.mutable_text_storage());
  EXPECT_TRUE(static_cast<bool>(rs));
  return ctx;
}

Metadata ViaDom(const std::string& xml) {
  return ContextViaDom(xml).cell_metadata;
}

Metadata ViaSax(const std::string& xml) {
  return ContextViaSax(xml).cell_metadata;
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

// A blocked formula anchor and a constant (a blocked pivot's anchor) both carry `vm`.
TEST(SheetReaderCellMetadataParity, ValueMetadataIsReadOnFormulasAndConstants) {
  const std::string xml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
      "<row r=\"1\"><c r=\"A1\" t=\"e\" cm=\"1\" vm=\"3\"><f t=\"array\" ref=\"A1\">A5:A6</f><v>#VALUE!</v></c>"
      "<c r=\"B1\" s=\"1\" t=\"e\" vm=\"1\"><v>#VALUE!</v></c><c r=\"C1\" t=\"e\" vm=\"x\"><v>#VALUE!</v></c>"
      "</row></sheetData></worksheet>";
  for (const SheetReadContext& ctx : {ContextViaDom(xml), ContextViaSax(xml)}) {
    ASSERT_EQ(ctx.value_metadata.size(), 3U);
    EXPECT_EQ(ctx.value_metadata[0].col, 0U);
    EXPECT_EQ(ctx.value_metadata[0].vm, 3U);
    EXPECT_EQ(ctx.value_metadata[1].col, 1U);
    EXPECT_EQ(ctx.value_metadata[1].vm, 1U);
    EXPECT_EQ(ctx.value_metadata[2].vm, 0U);
  }
}

}  // namespace
}  // namespace io
}  // namespace formulon
