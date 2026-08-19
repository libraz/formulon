//
// End-to-end tests for authoring print settings on an empty workbook and
// saving a package Excel can open.
//
// Two obligations meet here. The first is that a report built entirely
// through the API - paper, margins, fit-to-page, header/footer, print area,
// print titles, manual breaks - survives a save/load cycle intact and
// occupies the ECMA-376 element positions Excel expects. The second is that
// none of that regresses the passthrough route consumers already rely on:
// open an Excel-authored template, change one thing, and keep every
// attribute, relationship and binary part the engine does not model.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

std::vector<std::uint8_t> ReadPart(io::ZipReader& zip, const char* name) {
  auto bytes_or = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << "missing part: " << name;
  if (!bytes_or) {
    return {};
  }
  return bytes_or.value();
}

struct PartFile {
  const char* path;
  std::string_view body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& part : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, part.path, part.body.data(), part.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << "miniz add failed for " << part.path;
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(archive_size);
  if (archive_ptr != nullptr && archive_size > 0) {
    std::memcpy(out.data(), archive_ptr, archive_size);
  }
  mz_free(archive_ptr);
  return out;
}

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct BufferGuard {
  std::uint8_t* data = nullptr;
  std::size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

/// 0-based position of `name` among `worksheet`'s element children, or
/// `-1` when absent.
int ChildPosition(const pugi::xml_node& worksheet, const char* name) {
  int index = 0;
  for (pugi::xml_node child = worksheet.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    if (std::string_view(child.name()) == name) {
      return index;
    }
    ++index;
  }
  return -1;
}

}  // namespace

TEST(OoxmlPrintAuthoring, EmptyWorkbookToFullyConfiguredReport) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "\xE5\xA3\xB2\xE4\xB8\x8A"), 0);  // 売上

  fm_page_setup_t setup{};
  setup.paper_size_engaged = 1;
  setup.paper_size = 9;  // A4
  setup.orientation_engaged = 1;
  setup.orientation = FM_ORIENTATION_PORTRAIT;
  setup.fit_to_page_engaged = 1;
  setup.fit_to_page = 1;
  setup.fit_to_width_engaged = 1;
  setup.fit_to_width = 1;
  setup.fit_to_height_engaged = 1;
  setup.fit_to_height = 0;
  ASSERT_EQ(fm_sheet_set_page_setup(wb.handle, 0, &setup), 0);

  fm_page_margins_t margins{};
  margins.left_engaged = 1;
  margins.left = 0.5;
  margins.right_engaged = 1;
  margins.right = 0.5;
  margins.top_engaged = 1;
  margins.top = 0.8;
  margins.bottom_engaged = 1;
  margins.bottom = 0.8;
  margins.header_engaged = 1;
  margins.header = 0.3;
  margins.footer_engaged = 1;
  margins.footer = 0.3;
  ASSERT_EQ(fm_sheet_set_page_margins(wb.handle, 0, &margins), 0);

  fm_print_options_t options{};
  options.horizontal_centered_engaged = 1;
  options.horizontal_centered = 1;
  ASSERT_EQ(fm_sheet_set_print_options(wb.handle, 0, &options), 0);

  fm_header_footer_t hf{};
  hf.odd_header = "&C\xE6\x9C\x88\xE6\xAC\xA1\xE5\xA0\xB1\xE5\x91\x8A";  // "&C月次報告"
  hf.odd_footer = "&R&P / &N";
  ASSERT_EQ(fm_sheet_set_header_footer(wb.handle, 0, &hf), 0);

  ASSERT_EQ(fm_sheet_set_print_area(wb.handle, 0, "A1:F80"), 0);
  ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, "1:2", ""), 0);
  ASSERT_EQ(fm_sheet_add_row_break(wb.handle, 0, 39, 1), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);

  // --- The package Excel would see -----------------------------------
  std::vector<std::uint8_t> bytes(saved.data, saved.data + saved.len);
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  const std::vector<std::uint8_t> sheet_bytes = ReadPart(zip, "xl/worksheets/sheet1.xml");
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_buffer(sheet_bytes.data(), sheet_bytes.size()));
  const pugi::xml_node worksheet = doc.child("worksheet");
  ASSERT_TRUE(worksheet);

  // ECMA-376 fixes the order of these children. Excel repairs a worksheet
  // whose print elements appear out of sequence, so the positions matter
  // as much as the content.
  const int sheet_pr = ChildPosition(worksheet, "sheetPr");
  const int sheet_data = ChildPosition(worksheet, "sheetData");
  const int print_options = ChildPosition(worksheet, "printOptions");
  const int page_margins = ChildPosition(worksheet, "pageMargins");
  const int page_setup = ChildPosition(worksheet, "pageSetup");
  const int header_footer = ChildPosition(worksheet, "headerFooter");
  const int row_breaks = ChildPosition(worksheet, "rowBreaks");
  ASSERT_GE(sheet_pr, 0);
  ASSERT_GE(row_breaks, 0);
  EXPECT_LT(sheet_pr, sheet_data);
  EXPECT_LT(sheet_data, print_options);
  EXPECT_LT(print_options, page_margins);
  EXPECT_LT(page_margins, page_setup);
  EXPECT_LT(page_setup, header_footer);
  EXPECT_LT(header_footer, row_breaks);

  EXPECT_STREQ(worksheet.child("sheetPr").child("pageSetUpPr").attribute("fitToPage").value(), "true");
  EXPECT_STREQ(worksheet.child("pageSetup").attribute("paperSize").value(), "9");
  EXPECT_STREQ(worksheet.child("pageSetup").attribute("fitToHeight").value(), "0");
  // The header code reaches the file escaped and decodes back to `&C...`.
  EXPECT_EQ(std::string(worksheet.child("headerFooter").child("oddHeader").text().get()),
            "&C\xE6\x9C\x88\xE6\xAC\xA1\xE5\xA0\xB1\xE5\x91\x8A");
  EXPECT_STREQ(worksheet.child("rowBreaks").attribute("count").value(), "1");
  EXPECT_STREQ(worksheet.child("rowBreaks").attribute("manualBreakCount").value(), "1");
  // The break was authored before 0-based row 39; OOXML stores that index
  // as-is, which is how Excel writes it (`id="39"` == the page ends after
  // 39 rows). Emitting 40 put the break one row late in Excel.
  EXPECT_STREQ(worksheet.child("rowBreaks").child("brk").attribute("id").value(), "39");

  // --- Reloading it -------------------------------------------------
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);

  fm_page_setup_t read_setup{};
  ASSERT_EQ(fm_sheet_get_page_setup(reloaded.handle, 0, &read_setup), 0);
  EXPECT_EQ(read_setup.paper_size, 9U);
  EXPECT_EQ(read_setup.orientation, FM_ORIENTATION_PORTRAIT);
  EXPECT_EQ(read_setup.fit_to_page, 1);
  EXPECT_EQ(read_setup.fit_to_width, 1U);
  EXPECT_EQ(read_setup.fit_to_height, 0U);

  fm_page_margins_t read_margins{};
  ASSERT_EQ(fm_sheet_get_page_margins(reloaded.handle, 0, &read_margins), 0);
  EXPECT_DOUBLE_EQ(read_margins.left, 0.5);
  EXPECT_DOUBLE_EQ(read_margins.top, 0.8);

  const char* ranges = nullptr;
  ASSERT_EQ(fm_sheet_get_print_area(reloaded.handle, 0, &ranges), 0);
  EXPECT_STREQ(ranges, "A1:F80");

  const char* rows = nullptr;
  const char* cols = nullptr;
  ASSERT_EQ(fm_sheet_get_print_titles(reloaded.handle, 0, &rows, &cols), 0);
  EXPECT_STREQ(rows, "1:2");
  EXPECT_STREQ(cols, "");

  ASSERT_EQ(fm_sheet_row_break_count(reloaded.handle, 0), 1U);
  fm_page_break_t brk{};
  ASSERT_EQ(fm_sheet_row_break_at(reloaded.handle, 0, 0, &brk), 0);
  EXPECT_EQ(brk.id, 39U);
  EXPECT_EQ(brk.manual, 1);
}

namespace {

/// A worksheet carrying the print content an Excel-authored template has:
/// unmodelled `<pageSetup>` attributes, a printerSettings relationship, a
/// drawing reference, and a legacy header/footer drawing.
constexpr std::string_view kTemplateSheet =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
    "  <sheetPr codeName=\"Sheet1\"><tabColor rgb=\"FF00B050\"/></sheetPr>\n"
    "  <sheetData/>\n"
    "  <printOptions horizontalCentered=\"1\"/>\n"
    "  <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>\n"
    "  <pageSetup paperSize=\"9\" orientation=\"portrait\" horizontalDpi=\"600\" verticalDpi=\"600\" "
    "copies=\"2\" useFirstPageNumber=\"1\" r:id=\"rId2\"/>\n"
    "  <headerFooter><oddHeader>&amp;L&amp;G&amp;C\xE3\x83\xAD\xE3\x82\xB4</oddHeader></headerFooter>\n"
    "  <legacyDrawingHF r:id=\"rId3\"/>\n"
    "  <drawing r:id=\"rId1\"/>\n"
    "</worksheet>\n";

std::vector<std::uint8_t> BuildTemplatePackage() {
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Default Extension=\"png\" ContentType=\"image/png\"/>\n"
      "  <Default Extension=\"bin\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/>\n"
      "  <Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n"
      "  <Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
      "  <Override PartName=\"/xl/drawings/drawing1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>\n"
      "</Types>\n";
  const std::string_view package_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view workbook_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>\n"
      "</workbook>\n";
  const std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view sheet_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" "
      "Target=\"../drawings/drawing1.xml\"/>\n"
      "  <Relationship Id=\"rId2\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings\" "
      "Target=\"../printerSettings/printerSettings1.bin\"/>\n"
      "  <Relationship Id=\"rId3\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
      "Target=\"../media/image1.png\"/>\n"
      "</Relationships>\n";
  const std::string_view drawing_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>\n";
  const std::string_view image_bytes = "\x89PNG\r\n\x1a\n";
  const std::string_view printer_bytes = "\x01\x02\x03\x04";
  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", kTemplateSheet},
      {"xl/worksheets/_rels/sheet1.xml.rels", sheet_rels},
      {"xl/drawings/drawing1.xml", drawing_xml},
      {"xl/media/image1.png", image_bytes},
      {"xl/printerSettings/printerSettings1.bin", printer_bytes},
  });
}

}  // namespace

TEST(OoxmlPrintAuthoring, TypedPatchOnATemplateKeepsEverythingItDoesNotState) {
  const std::vector<std::uint8_t> source = BuildTemplatePackage();
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(source.data(), source.size(), &wb.handle), 0);

  // Change exactly one attribute.
  fm_page_setup_t patch{};
  patch.orientation_engaged = 1;
  patch.orientation = FM_ORIENTATION_LANDSCAPE;
  ASSERT_EQ(fm_sheet_set_page_setup(wb.handle, 0, &patch), 0);

  BufferGuard saved;
  fm_save_diagnostics_t diagnostics{};
  ASSERT_EQ(
      fm_workbook_save_with_diagnostics(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len, &diagnostics), 0);
  // The consumer-visible contract for the template route: rewriting print
  // settings costs no part and no relationship.
  EXPECT_EQ(diagnostics.dropped_part_count, 0U);
  EXPECT_EQ(diagnostics.dropped_relationship_count, 0U);
  EXPECT_EQ(diagnostics.renumbered_part_count, 0U);
  std::vector<std::uint8_t> bytes(saved.data, saved.data + saved.len);

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  const std::vector<std::uint8_t> sheet_bytes = ReadPart(zip, "xl/worksheets/sheet1.xml");
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_buffer(sheet_bytes.data(), sheet_bytes.size()));
  const pugi::xml_node worksheet = doc.child("worksheet");
  const pugi::xml_node page_setup = worksheet.child("pageSetup");

  EXPECT_STREQ(page_setup.attribute("orientation").value(), "landscape");
  // Everything the patch did not name, including the four attributes the
  // engine has no field for, has to survive verbatim - that passthrough is
  // what makes the "edit an Excel template" route usable at all.
  EXPECT_STREQ(page_setup.attribute("paperSize").value(), "9");
  EXPECT_STREQ(page_setup.attribute("horizontalDpi").value(), "600");
  EXPECT_STREQ(page_setup.attribute("verticalDpi").value(), "600");
  EXPECT_STREQ(page_setup.attribute("copies").value(), "2");
  EXPECT_STREQ(page_setup.attribute("useFirstPageNumber").value(), "1");
  // The printerSettings reference must still resolve.
  EXPECT_FALSE(std::string_view(page_setup.attribute("r:id").value()).empty());

  // The template's other content is untouched.
  EXPECT_STREQ(worksheet.child("sheetPr").attribute("codeName").value(), "Sheet1");
  EXPECT_STREQ(worksheet.child("sheetPr").child("tabColor").attribute("rgb").value(), "FF00B050");
  EXPECT_TRUE(worksheet.child("legacyDrawingHF"));
  EXPECT_TRUE(worksheet.child("drawing"));
  EXPECT_TRUE(static_cast<bool>(zip.read_entry("xl/media/image1.png")));
  EXPECT_TRUE(static_cast<bool>(zip.read_entry("xl/drawings/drawing1.xml")));
  EXPECT_TRUE(static_cast<bool>(zip.read_entry("xl/printerSettings/printerSettings1.bin")));
}

TEST(OoxmlPrintAuthoring, FitToPageHelperOnATemplateKeepsTabColorAndCodeName) {
  const std::vector<std::uint8_t> source = BuildTemplatePackage();
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(source.data(), source.size(), &wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_fit_to_page(wb.handle, 0, 1), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  std::vector<std::uint8_t> bytes(saved.data, saved.data + saved.len);
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  const std::vector<std::uint8_t> sheet_bytes = ReadPart(zip, "xl/worksheets/sheet1.xml");
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_buffer(sheet_bytes.data(), sheet_bytes.size()));
  const pugi::xml_node sheet_pr = doc.child("worksheet").child("sheetPr");
  EXPECT_STREQ(sheet_pr.attribute("codeName").value(), "Sheet1");
  EXPECT_STREQ(sheet_pr.child("tabColor").attribute("rgb").value(), "FF00B050");
  EXPECT_STREQ(sheet_pr.child("pageSetUpPr").attribute("fitToPage").value(), "true");
}

TEST(OoxmlPrintAuthoring, ManualBreakCountExcludesAutomaticBreaks) {
  // `count` is every `<brk>`; `manualBreakCount` only those with `man`.
  // Overstating the second makes Excel's page-break preview draw a solid
  // user-break line where the file declares a computed one.
  constexpr std::string_view sheet =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "  <rowBreaks count=\"3\" manualBreakCount=\"3\">\n"
      "    <brk id=\"10\" max=\"16383\" man=\"1\"/>\n"
      "    <brk id=\"20\" max=\"16383\"/>\n"
      "    <brk id=\"30\" max=\"16383\" man=\"1\"/>\n"
      "  </rowBreaks>\n"
      "</worksheet>\n";
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n"
      "  <Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
      "</Types>\n";
  const std::string_view package_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view workbook_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>\n"
      "</workbook>\n";
  const std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "</Relationships>\n";
  const std::vector<std::uint8_t> source = BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet},
  });

  auto load_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(load_or)) << load_or.error().message;
  auto save_or = load_or.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << save_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  const std::vector<std::uint8_t> sheet_bytes = ReadPart(zip, "xl/worksheets/sheet1.xml");
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_buffer(sheet_bytes.data(), sheet_bytes.size()));
  const pugi::xml_node row_breaks = doc.child("worksheet").child("rowBreaks");
  ASSERT_TRUE(row_breaks);
  EXPECT_STREQ(row_breaks.attribute("count").value(), "3");
  EXPECT_STREQ(row_breaks.attribute("manualBreakCount").value(), "2");
}

}  // namespace formulon
