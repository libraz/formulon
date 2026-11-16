// Pins Mac column widths and pagination against Mac Excel 16.115: vertical page breaks of one-row print areas, and
// the width and ColumnWidth of a default and two stored columns under each measured Normal face and size.

#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "column_width.h"
#include "excel_profile.h"
#include "font_metrics.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "print/pagination.h"
#include "print/sheet_geometry.h"
#include "support/json_cases.h"
#include "theme.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace print {
namespace {

constexpr char kXmlDecl[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
constexpr char kRelNs[] = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

struct FaceRecord {
  const char* face;
  int family;
  int charset;  ///< 0: no `<charset>` element.
};

/// The `<family>` / `<charset>` the measured books stored for each face.
constexpr FaceRecord kFaceRecords[] = {
    {"Calibri", 2, 0},      {"Aptos", 2, 0},  {"Aptos Narrow", 2, 0},
    {"游ゴシック", 3, 128}, {"等线", 3, 134}, {"맑은 고딕", 3, 129},
    {"新細明體", 1, 136},   {"Tahoma", 2, 0}, {"ＭＳ Ｐゴシック", 3, 128},
    {"Yu Gothic", 3, 128},
};

std::string FontXml(const std::string& font, int size) {
  for (const FaceRecord& r : kFaceRecords) {
    if (font == r.face) {
      std::string xml = "<font><sz val=\"" + std::to_string(size) + "\"/><name val=\"" + font + "\"/><family val=\"" +
                        std::to_string(r.family) + "\"/>";
      if (r.charset != 0) {
        xml += "<charset val=\"" + std::to_string(r.charset) + "\"/>";
      }
      return xml + "</font>";
    }
  }
  ADD_FAILURE() << "no font record for " << font;
  return {};
}

/// `<cols>` giving columns A:T one stored width.
std::string UniformCols(double width) {
  char width_text[32];
  std::snprintf(width_text, sizeof(width_text), "%.2f", width);
  return std::string("<cols><col min=\"1\" max=\"20\" width=\"") + width_text + "\" customWidth=\"1\"/></cols>";
}

/// Builds the one-sheet package Excel opened: font 0 is the named face, `cols` the column widths, and with
/// `print_area` the print area A1:T1.
std::vector<std::uint8_t> BuildPackage(const std::string& font, int size, const std::string& cols, bool print_area) {
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
                               kRelNs + "\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                               (print_area ? "<definedNames><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">"
                                             "Sheet1!$A$1:$T$1</definedName></definedNames>"
                                           : "") +
                               "</workbook>";
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
  const std::string sheet =
      std::string(kXmlDecl) +
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<dimension ref=\"A1:T1\"/><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews>"
      "<sheetFormatPr defaultRowHeight=\"15\"/>" +
      cols +
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

    const std::vector<std::uint8_t> bytes = BuildPackage(font, size, UniformCols(width), true);
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

Workbook ReadMacBook(const std::vector<std::uint8_t>& bytes) {
  auto read_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  Workbook wb = std::move(read_or.value().workbook);
  wb.set_excel_profile(mac_365_ja_jp_profile());
  return wb;
}

struct MeasuredColumn {
  double width_pt;      ///< `Range.Width`.
  double column_width;  ///< `Range.ColumnWidth`, the displayed character count.
};

struct WidthBook {
  const char* face;
  int size;
  double mdw;              ///< (Width of C - Width of B) / 70.
  double padding;          ///< Width of A - its ColumnWidth x MDW.
  double standard_height;  ///< `StandardHeight`.
  MeasuredColumn cols[3];  ///< A (default width), B (stored 30), C (stored 100).
};

// Mac Excel 16.115 (backup probe mdw-2026-11-16): font 0 is the face with no theme link, `<sheetFormatPr
// defaultRowHeight="15"/>`, B and C stored at 30 and 100 characters, A at the default width.
constexpr WidthBook kWidthBooks[] = {
    {"Calibri", 9, 5, 5, 12, {{45, 8}, {150, 29}, {500, 99}}},
    {"Calibri", 11, 6, 5, 15, {{53, 8}, {180, 29.17}, {600, 99.17}}},
    {"Calibri", 16, 8, 5, 21, {{69, 8}, {240, 29.38}, {800, 99.38}}},
    {"Aptos", 8, 4, 3, 11, {{35, 8}, {120, 29.25}, {400, 99.25}}},
    {"Aptos", 9, 5, 5, 13, {{45, 8}, {150, 29}, {500, 99}}},
    {"Aptos", 10, 5, 5, 14, {{45, 8}, {150, 29}, {500, 99}}},
    {"Aptos", 14, 7, 5, 19, {{61, 8}, {210, 29.29}, {700, 99.29}}},
    {"Aptos", 16, 9, 7, 22, {{79, 8}, {270, 29.22}, {900, 99.22}}},
    {"Aptos", 18, 10, 7, 24, {{87, 8}, {300, 29.3}, {1000, 99.3}}},
    {"Aptos", 20, 11, 7, 27, {{95, 8}, {330, 29.36}, {1100, 99.36}}},
    {"Aptos Narrow", 8, 4, 3, 11, {{35, 8}, {120, 29.25}, {400, 99.25}}},
    {"Aptos Narrow", 9, 5, 5, 13, {{45, 8}, {150, 29}, {500, 99}}},
    {"Aptos Narrow", 16, 8, 5, 22, {{69, 8}, {240, 29.38}, {800, 99.38}}},
    {"Aptos Narrow", 18, 9, 7, 24, {{79, 8}, {270, 29.22}, {900, 99.22}}},
    {"Aptos Narrow", 20, 10, 7, 27, {{87, 8}, {300, 29.3}, {1000, 99.3}}},
    {"游ゴシック", 8, 4, 3, 13, {{35, 8}, {120, 29.25}, {400, 99.25}}},
    {"游ゴシック", 9, 5, 5, 16, {{45, 8}, {150, 29}, {500, 99}}},
    {"游ゴシック", 11, 6, 5, 18, {{53, 8}, {180, 29.17}, {600, 99.17}}},
    {"游ゴシック", 16, 9, 7, 27, {{79, 8}, {270, 29.22}, {900, 99.22}}},
    {"游ゴシック", 18, 10, 7, 31, {{87, 8}, {300, 29.3}, {1000, 99.3}}},
    {"游ゴシック", 20, 11, 7, 33, {{95, 8}, {330, 29.36}, {1100, 99.36}}},
    {"等线", 8, 4, 3, 11, {{35, 8}, {120, 29.25}, {400, 99.25}}},
    {"等线", 9, 5, 5, 12, {{45, 8}, {150, 29}, {500, 99}}},
    {"等线", 10, 5, 5, 13, {{45, 8}, {150, 29}, {500, 99}}},
    {"等线", 14, 7, 5, 18, {{61, 8}, {210, 29.29}, {700, 99.29}}},
    {"等线", 16, 8, 5, 21, {{69, 8}, {240, 29.38}, {800, 99.38}}},
    {"等线", 18, 9, 7, 23, {{79, 8}, {270, 29.22}, {900, 99.22}}},
    {"等线", 20, 11, 7, 26, {{95, 8}, {330, 29.36}, {1100, 99.36}}},
    {"맑은 고딕", 8, 4, 3, 13, {{35, 8}, {120, 29.25}, {400, 99.25}}},
    {"맑은 고딕", 9, 5, 5, 14, {{45, 8}, {150, 29}, {500, 99}}},
    {"맑은 고딕", 10, 6, 5, 15, {{53, 8}, {180, 29.17}, {600, 99.17}}},
    {"맑은 고딕", 14, 8, 5, 20, {{69, 8}, {240, 29.38}, {800, 99.38}}},
    {"맑은 고딕", 16, 9, 7, 23, {{79, 8}, {270, 29.22}, {900, 99.22}}},
    {"맑은 고딕", 18, 10, 7, 27, {{87, 8}, {300, 29.3}, {1000, 99.3}}},
    {"맑은 고딕", 20, 11, 7, 30, {{95, 8}, {330, 29.36}, {1100, 99.36}}},
    {"新細明體", 8, 4, 3, 11, {{35, 8}, {120, 29.25}, {400, 99.25}}},
    {"新細明體", 9, 4, 3, 12, {{35, 8}, {120, 29.25}, {400, 99.25}}},
    {"新細明體", 10, 5, 5, 13, {{45, 8}, {150, 29}, {500, 99}}},
    {"新細明體", 14, 7, 5, 20, {{61, 8}, {210, 29.29}, {700, 99.29}}},
    {"新細明體", 16, 8, 5, 22, {{69, 8}, {240, 29.38}, {800, 99.38}}},
    {"新細明體", 18, 8, 5, 25, {{69, 8}, {240, 29.38}, {800, 99.38}}},
    {"新細明體", 20, 9, 7, 27, {{79, 8}, {270, 29.22}, {900, 99.22}}},
    {"Tahoma", 8, 4, 3, 11, {{35, 8}, {120, 29.25}, {400, 99.25}}},
    {"Tahoma", 9, 5, 5, 12, {{45, 8}, {150, 29}, {500, 99}}},
    {"Tahoma", 14, 8, 5, 18, {{69, 8}, {240, 29.38}, {800, 99.38}}},
    {"Tahoma", 16, 9, 7, 20, {{79, 8}, {270, 29.22}, {900, 99.22}}},
    {"Tahoma", 18, 10, 7, 23, {{87, 8}, {300, 29.3}, {1000, 99.3}}},
    {"Tahoma", 20, 11, 7, 25, {{95, 8}, {330, 29.36}, {1100, 99.36}}},
    {"ＭＳ Ｐゴシック", 8, 4, 3, 12, {{35, 8}, {120, 29.25}, {400, 99.25}}},
    {"ＭＳ Ｐゴシック", 9, 5, 5, 13, {{45, 8}, {150, 29}, {500, 99}}},
    {"ＭＳ Ｐゴシック", 10, 5, 5, 14, {{45, 8}, {150, 29}, {500, 99}}},
    {"ＭＳ Ｐゴシック", 14, 7, 5, 17, {{61, 8}, {210, 29.29}, {700, 99.29}}},
    {"ＭＳ Ｐゴシック", 16, 8, 5, 19, {{69, 8}, {240, 29.38}, {800, 99.38}}},
    {"ＭＳ Ｐゴシック", 18, 9, 7, 22, {{79, 8}, {270, 29.22}, {900, 99.22}}},
    {"ＭＳ Ｐゴシック", 20, 10, 7, 24, {{87, 8}, {300, 29.3}, {1000, 99.3}}},
    {"Yu Gothic", 10, 6, 5, 17, {{53, 8}, {180, 29.17}, {600, 99.17}}},
    {"Yu Gothic", 11, 6, 5, 17, {{53, 8}, {180, 29.17}, {600, 99.17}}},
    {"Yu Gothic", 12, 7, 5, 18, {{61, 8}, {210, 29.29}, {700, 99.29}}},
};

TEST(MacColumnWidth, MatchMacExcelPerNormalFace) {
  for (const WidthBook& book : kWidthBooks) {
    SCOPED_TRACE(std::string(book.face) + " " + std::to_string(book.size));
    const std::string cols =
        "<cols><col min=\"2\" max=\"2\" width=\"30\" customWidth=\"1\"/>"
        "<col min=\"3\" max=\"3\" width=\"100\" customWidth=\"1\"/></cols>";
    const Workbook wb = ReadMacBook(BuildPackage(book.face, book.size, cols, false));
    for (const GeometryMode mode : {GeometryMode::kDisplay, GeometryMode::kPrint}) {
      const ColumnWidthModel model = resolve_column_width_model(wb, wb.excel_profile(), mode);
      EXPECT_TRUE(model.calibrated);
      EXPECT_EQ(model.max_digit_width_px, static_cast<int>(book.mdw));
      EXPECT_DOUBLE_EQ(model.points_per_char, book.mdw);
      EXPECT_DOUBLE_EQ(model.padding_pt, book.padding);
      for (std::uint32_t col = 0; col < 3U; ++col) {
        const double chars = effective_column_width_chars(wb.sheet(0), col, model);
        EXPECT_DOUBLE_EQ(column_chars_to_points(chars, model), book.cols[col].width_pt) << "column " << col;
        EXPECT_DOUBLE_EQ(column_display_chars(chars, model), book.cols[col].column_width) << "column " << col;
      }
    }
    const NormalFontMetrics metrics = normal_font_metrics(wb.styles(), default_theme(), wb.excel_profile());
    EXPECT_DOUBLE_EQ(metrics.row_height_pt, book.standard_height);
    EXPECT_TRUE(metrics.row_height_calibrated);
  }
}

struct PrintBook {
  const char* face;
  int size;
  int width;                ///< Stored width of A:T.
  double width_pt;          ///< Width of A.
  std::uint32_t breaks[4];  ///< Vertical breaks, the 0-based first column of each later page.
};

// Mac Excel 16.115 (the same probe): the print-break layout above under further Normal faces.
constexpr PrintBook kPrintBooks[] = {
    {"游ゴシック", 10, 8, 48, {10}},
    {"游ゴシック", 10, 12, 72, {6, 12, 18}},
    {"游ゴシック", 10, 20, 120, {4, 8, 12, 16}},
    {"等线", 11, 8, 48, {10}},
    {"等线", 11, 12, 72, {6, 12, 18}},
    {"等线", 11, 20, 120, {4, 8, 12, 16}},
    {"Aptos Narrow", 12, 8, 48, {10}},
    {"Aptos Narrow", 12, 12, 72, {6, 12, 18}},
    {"Aptos Narrow", 12, 20, 120, {4, 8, 12, 16}},
};

TEST(MacColumnWidth, PrintBreaksPerNormalFace) {
  for (const PrintBook& book : kPrintBooks) {
    SCOPED_TRACE(std::string(book.face) + " " + std::to_string(book.size) + " w=" + std::to_string(book.width));
    const Workbook wb = ReadMacBook(BuildPackage(book.face, book.size, UniformCols(book.width), true));
    const ColumnWidthModel model = resolve_column_width_model(wb, wb.excel_profile(), GeometryMode::kPrint);
    EXPECT_TRUE(model.calibrated);
    EXPECT_DOUBLE_EQ(column_chars_to_points(book.width, model), book.width_pt);
    auto page_or = paginate(wb, 0U);
    ASSERT_TRUE(static_cast<bool>(page_or)) << page_or.error().message;
    std::vector<std::uint32_t> want;
    for (const std::uint32_t b : book.breaks) {
      if (b != 0U) {
        want.push_back(b);
      }
    }
    EXPECT_EQ(page_or.value().v_breaks, want);
  }
}

}  // namespace
}  // namespace print
}  // namespace formulon
