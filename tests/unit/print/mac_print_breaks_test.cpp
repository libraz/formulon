// Pins Mac pagination against vertical page breaks measured in Mac Excel for one-row print areas.

#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "column_width.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "print/pagination.h"
#include "support/json_cases.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace print {
namespace {

constexpr char kXmlDecl[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
constexpr char kRelNs[] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

std::string FontXml(const std::string& font, int size) {
  if (font == "Calibri") {
    return "<font><sz val=\"" + std::to_string(size) + "\"/><name val=\"Calibri\"/><family val=\"2\"/></font>";
  }
  return "<font><sz val=\"" + std::to_string(size) + "\"/><name val=\"" + font +
         "\"/><family val=\"3\"/><charset val=\"128\"/></font>";
}

/// Builds the one-sheet package Excel opened: font 0 is the named face, columns A:T share one width.
std::vector<std::uint8_t> BuildPackage(const std::string& font, int size, double width) {
  const std::string content_types =
      std::string(kXmlDecl) +
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
      "<Override PartName=\"/xl/styles.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
      "</Types>";
  const std::string rels = std::string(kXmlDecl) +
                           "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                           "<Relationship Id=\"rId1\" Type=\"" +
                           kRelNs + "/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>";
  const std::string workbook = std::string(kXmlDecl) +
                               "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                               "xmlns:r=\"" +
                               kRelNs +
                               "\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
                               "<definedNames><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">"
                               "Sheet1!$A$1:$T$1</definedName></definedNames></workbook>";
  const std::string workbook_rels =
      std::string(kXmlDecl) +
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" Type=\"" +
      kRelNs +
      "/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
      "<Relationship Id=\"rId2\" Type=\"" +
      kRelNs + "/styles\" Target=\"styles.xml\"/></Relationships>";
  const std::string styles =
      std::string(kXmlDecl) +
      "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\">" +
      FontXml(font, size) +
      "</fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill>"
      "<fill><patternFill patternType=\"gray125\"/></fill></fills>"
      "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
      "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
      "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"
      "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>";
  char width_text[32];
  std::snprintf(width_text, sizeof(width_text), "%.2f", width);
  const std::string sheet =
      std::string(kXmlDecl) +
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<dimension ref=\"A1:T1\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>"
      "<cols><col min=\"1\" max=\"20\" width=\"" +
      width_text +
      "\" customWidth=\"1\"/></cols>"
      "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c><c r=\"T1\"><v>2</v></c></row></sheetData>"
      "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>"
      "<pageSetup paperSize=\"9\" scale=\"100\" orientation=\"portrait\"/></worksheet>";

  mz_zip_archive writer{};
  if (mz_zip_writer_init_heap(&writer, 0, 4096) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_writer_init_heap failed";
    return {};
  }
  const std::pair<const char*, const std::string*> parts[] = {
      {"[Content_Types].xml", &content_types},        {"_rels/.rels", &rels},     {"xl/workbook.xml", &workbook},
      {"xl/_rels/workbook.xml.rels", &workbook_rels}, {"xl/styles.xml", &styles}, {"xl/worksheets/sheet1.xml", &sheet},
  };
  for (const auto& part : parts) {
    if (mz_zip_writer_add_mem(&writer, part.first, part.second->data(), part.second->size(),
                              static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)) == MZ_FALSE) {
      ADD_FAILURE() << "mz_zip_writer_add_mem failed for " << part.first;
      mz_zip_writer_end(&writer);
      return {};
    }
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  if (mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_writer_finalize_heap_archive failed";
    mz_zip_writer_end(&writer);
    return {};
  }
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  mz_zip_writer_end(&writer);
  return out;
}

TEST(MacPrintBreaks, MatchMacExcel) {
  test_support::Json json;
  ASSERT_TRUE(
      test_support::LoadRepoJson(__FILE__, "tests/unit/", "tests/fixtures/excel/mac/print_breaks/breaks.json", &json));
  const test_support::Json& cases = json.At("cases");
  ASSERT_EQ(cases.items.size(), 453U);

  for (const test_support::Json& c : cases.items) {
    const std::string font = c.At("font").text;
    const int size = static_cast<int>(c.At("size").number);
    const double width = c.At("width").number;
    SCOPED_TRACE(font + " " + std::to_string(size) + " w=" + std::to_string(width));

    const std::vector<std::uint8_t> bytes = BuildPackage(font, size, width);
    ASSERT_FALSE(bytes.empty());
    auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
    ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
    Workbook wb = std::move(read_or.value().workbook);
    wb.set_excel_profile(mac_365_ja_jp_profile());

    const ColumnWidthModel model = resolve_column_width_model(wb, wb.excel_profile(), GeometryMode::kPrint);
    EXPECT_DOUBLE_EQ(column_chars_to_points(width, model), c.At("width_pt").number);

    auto page_or = paginate(wb, 0U);
    ASSERT_TRUE(static_cast<bool>(page_or)) << page_or.error().message;
    const std::vector<std::uint32_t>& got = page_or.value().v_breaks;

    std::vector<std::uint32_t> want;
    for (const test_support::Json& b : c.At("breaks").items) {
      want.push_back(static_cast<std::uint32_t>(b.number));
    }
    std::vector<std::uint32_t> compared = got;
    if (c.At("truncated").boolean) {
      ASSERT_GE(got.size(), want.size());
      compared.resize(want.size());
    }
    EXPECT_EQ(compared, want);
  }
}

}  // namespace
}  // namespace print
}  // namespace formulon
