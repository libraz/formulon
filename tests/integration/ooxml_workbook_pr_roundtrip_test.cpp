//
// Integration test: workbook.xml level elements `<workbookPr>` (with the
// `date1904` date-system flag and the VBA `codeName`), `<bookViews>`
// (tab-selection state via `activeTab`), and `<workbookProtection>` must
// round-trip. Before this, the writer regenerated only <sheets> /
// <definedNames> / <calcPr> / <pivotCaches>, so a load->save cycle
// silently dropped all three — most damagingly `date1904`, which shifts
// every date serial by 1462 days when lost.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "pugixml.hpp"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

struct PartFile {
  const char* path;
  std::string_view body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& p : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, p.path, p.body.data(), p.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << "miniz add failed for " << p.path;
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

// Minimal package whose workbook.xml carries workbookPr (date1904 +
// codeName), workbookProtection, and bookViews (activeTab).
std::vector<std::uint8_t> BuildPackage(std::string_view workbook_body) {
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
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
  const std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "</worksheet>\n";
  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_body},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
  });
}

std::vector<std::uint8_t> ReadPart(io::ZipReader& zip, std::string_view name) {
  auto bytes_or = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << "missing part: " << name;
  if (!bytes_or) {
    return {};
  }
  return bytes_or.value();
}

constexpr std::string_view kWorkbookWithAll =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
    "  <fileVersion appName=\"xl\" lastEdited=\"8\"/>\n"
    "  <fileSharing readOnlyRecommended=\"1\"/>\n"
    "  <workbookPr date1904=\"1\" codeName=\"ThisWorkbook\"/>\n"
    "  <workbookProtection lockStructure=\"1\" lockWindows=\"0\"/>\n"
    "  <bookViews>\n"
    "    <workbookView xWindow=\"0\" yWindow=\"0\" activeTab=\"2\"/>\n"
    "  </bookViews>\n"
    "  <sheets>\n"
    "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
    "  </sheets>\n"
    "  <extLst><ext uri=\"urn:test:workbook-extension\"/></extLst>\n"
    "</workbook>\n";

TEST(OoxmlWorkbookPrRoundTrip, Date1904FlagAndRawElementsSurvive) {
  const std::vector<std::uint8_t> source = BuildPackage(kWorkbookWithAll);
  auto load_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;
  const Workbook& wb = load_or.value().workbook;

  // Model-level date1904 flag parsed from <workbookPr date1904="1">.
  EXPECT_TRUE(wb.date1904());
  // Raw captures populated.
  EXPECT_NE(wb.workbook_pr_xml().find("codeName=\"ThisWorkbook\""), std::string::npos);
  EXPECT_NE(wb.workbook_protection_xml().find("lockStructure=\"1\""), std::string::npos);
  EXPECT_NE(wb.book_views_xml().find("activeTab=\"2\""), std::string::npos);
  EXPECT_NE(wb.file_version_xml().find("lastEdited=\"8\""), std::string::npos);
  EXPECT_NE(wb.file_sharing_xml().find("readOnlyRecommended=\"1\""), std::string::npos);
  EXPECT_NE(wb.workbook_ext_lst_xml().find("urn:test:workbook-extension"), std::string::npos);

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  const std::vector<std::uint8_t> out = ReadPart(zip, "xl/workbook.xml");
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_buffer(out.data(), out.size()));
  pugi::xml_node root = doc.child("workbook");
  ASSERT_TRUE(root);

  // date1904 + codeName preserved on workbookPr.
  pugi::xml_node wpr = root.child("workbookPr");
  ASSERT_TRUE(wpr) << "workbookPr dropped on save";
  EXPECT_STREQ(wpr.attribute("date1904").value(), "1");
  EXPECT_STREQ(wpr.attribute("codeName").value(), "ThisWorkbook");

  // workbookProtection preserved (both explicit lock flags).
  pugi::xml_node wp = root.child("workbookProtection");
  ASSERT_TRUE(wp) << "workbookProtection dropped on save";
  EXPECT_STREQ(wp.attribute("lockStructure").value(), "1");
  EXPECT_STREQ(wp.attribute("lockWindows").value(), "0");

  // bookViews / activeTab preserved.
  pugi::xml_node bv = root.child("bookViews");
  ASSERT_TRUE(bv) << "bookViews dropped on save";
  EXPECT_STREQ(bv.child("workbookView").attribute("activeTab").value(), "2");

  // ECMA-376 element order: workbookPr, workbookProtection, bookViews all
  // precede <sheets>.
  int order_wpr = -1;
  int order_wp = -1;
  int order_bv = -1;
  int order_sheets = -1;
  int idx = 0;
  for (pugi::xml_node n = root.first_child(); n; n = n.next_sibling(), ++idx) {
    const std::string_view name = n.name();
    if (name == "workbookPr") {
      order_wpr = idx;
    } else if (name == "workbookProtection") {
      order_wp = idx;
    } else if (name == "bookViews") {
      order_bv = idx;
    } else if (name == "sheets") {
      order_sheets = idx;
    }
  }
  EXPECT_LT(order_wpr, order_wp);
  EXPECT_LT(order_wp, order_bv);
  EXPECT_LT(order_bv, order_sheets);
}

TEST(OoxmlWorkbookPrRoundTrip, DefaultWorkbookOmitsWorkbookPr) {
  // A workbook without <workbookPr> must not gain a synthetic one on save
  // (the 1904 flag is false, so nothing to emit).
  constexpr std::string_view kPlain =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheets>\n"
      "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
      "  </sheets>\n"
      "</workbook>\n";
  auto load_or = io::read_ooxml(SpanOf(BuildPackage(kPlain)));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;
  EXPECT_FALSE(load_or.value().workbook.date1904());

  auto save_or = load_or.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  const std::vector<std::uint8_t> out = ReadPart(zip, "xl/workbook.xml");
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_buffer(out.data(), out.size()));
  EXPECT_FALSE(doc.child("workbook").child("workbookPr")) << "plain workbook must not gain a synthetic <workbookPr>";
}

TEST(OoxmlWorkbookPrRoundTrip, ProgrammaticDate1904SynthesisesWorkbookPr) {
  // When date1904 is set on a workbook that carried no raw <workbookPr>,
  // the writer synthesises a minimal element so the flag is not lost.
  auto load_or = io::read_ooxml(
      SpanOf(BuildPackage("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                          "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                          "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
                          "  <sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>\n"
                          "</workbook>\n")));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;
  Workbook wb = std::move(load_or.value().workbook);
  wb.set_date1904(true);

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  const std::vector<std::uint8_t> out = ReadPart(zip, "xl/workbook.xml");
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_buffer(out.data(), out.size()));
  pugi::xml_node wpr = doc.child("workbook").child("workbookPr");
  ASSERT_TRUE(wpr) << "date1904 flag must synthesise a <workbookPr>";
  EXPECT_STREQ(wpr.attribute("date1904").value(), "1");
}

}  // namespace
}  // namespace formulon
