//
// Integration test: parts declared only through `<Default Extension>` in
// `[Content_Types].xml` (vbaProject.bin, images, drawings, VML, and
// their rels) must round-trip through the passthrough mechanism. Real
// Excel-authored .xlsm / .xlsx packages declare these by extension, not
// with a per-part `<Override>`; without Default-typed capture a
// load->save cycle silently drops macros, images, shapes, and note
// geometry, and leaves dangling relationships that trip Excel's repair
// dialog. These tests also cover the writer's relationship-integrity
// guard (skip rels whose target part is absent) and the worksheet
// `<drawing>` round-trip.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/passthrough_part.h"
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

std::vector<std::uint8_t> ReadPart(io::ZipReader& zip, std::string_view name) {
  auto bytes_or = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(bytes_or)) << "missing part: " << name;
  if (!bytes_or) {
    return {};
  }
  return bytes_or.value();
}

// -- (a) Default-typed vbaProject.bin round-trip -----------------------------

std::vector<std::uint8_t> BuildMacroPackage() {
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>\n"
      "  <Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.ms-excel.sheet.macroEnabled.main+xml\"/>\n"
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
      "  <sheets>\n"
      "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
      "  </sheets>\n"
      "</workbook>\n";
  // The vbaProject rel type is a Microsoft-specific URI the reader does
  // not model, so it flows through unknown-workbook-rels. A second rel
  // deliberately points at a part that does not exist in the archive so
  // we can assert the writer drops it instead of dangling.
  const std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "  <Relationship Id=\"rId2\" "
      "Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" "
      "Target=\"vbaProject.bin\"/>\n"
      "  <Relationship Id=\"rId3\" "
      "Type=\"http://schemas.microsoft.com/office/2006/relationships/customUI\" "
      "Target=\"ghostRibbon.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "</worksheet>\n";
  // Non-UTF-8 binary payload: a marker that must survive verbatim.
  const std::string_view vba_bin = std::string_view("\x00\x01MZVBA\xFF\xFE\x00blob", 13);

  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/vbaProject.bin", vba_bin},
  });
}

TEST(OoxmlDefaultPassthrough, VbaProjectBytesAndDefaultRoundTrip) {
  const std::vector<std::uint8_t> source = BuildMacroPackage();
  auto load_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;

  // The reader must have captured the Default-typed part.
  const std::vector<io::PassthroughPart>& parts = load_or.value().workbook.passthrough_parts();
  bool saw_vba = false;
  for (const io::PassthroughPart& p : parts) {
    if (p.path == "xl/vbaProject.bin") {
      saw_vba = true;
      EXPECT_TRUE(p.content_type.empty()) << "Default-typed part must carry an empty content type";
    }
  }
  EXPECT_TRUE(saw_vba) << "vbaProject.bin not captured as passthrough";

  auto save_or = load_or.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  const std::vector<std::uint8_t>& out = save_or.value();

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(out))));

  // Bytes preserved verbatim.
  ASSERT_TRUE(zip.has_entry("xl/vbaProject.bin"));
  const std::vector<std::uint8_t> vba_bytes = ReadPart(zip, "xl/vbaProject.bin");
  const std::string_view expected = std::string_view("\x00\x01MZVBA\xFF\xFE\x00blob", 13);
  ASSERT_EQ(vba_bytes.size(), expected.size());
  EXPECT_EQ(0, std::memcmp(vba_bytes.data(), expected.data(), expected.size()));

  // `<Default Extension="bin">` round-tripped into content types.
  const std::vector<std::uint8_t> ct = ReadPart(zip, "[Content_Types].xml");
  pugi::xml_document ct_doc;
  ASSERT_TRUE(ct_doc.load_buffer(ct.data(), ct.size()));
  pugi::xml_node ct_root = ct_doc.child("Types");
  ASSERT_TRUE(ct_root);
  bool saw_bin_default = false;
  for (pugi::xml_node node = ct_root.child("Default"); node; node = node.next_sibling("Default")) {
    if (std::string_view(node.attribute("Extension").value()) == "bin") {
      saw_bin_default = true;
      EXPECT_EQ(std::string_view(node.attribute("ContentType").value()), "application/vnd.ms-office.vbaProject");
    }
  }
  EXPECT_TRUE(saw_bin_default) << "Default Extension=bin missing from round-tripped content types";

  // The vbaProject relationship survives; the customUI relationship
  // pointing at a missing part is dropped.
  const std::vector<std::uint8_t> rels = ReadPart(zip, "xl/_rels/workbook.xml.rels");
  pugi::xml_document rels_doc;
  ASSERT_TRUE(rels_doc.load_buffer(rels.data(), rels.size()));
  pugi::xml_node rels_root = rels_doc.child("Relationships");
  ASSERT_TRUE(rels_root);
  bool saw_vba_rel = false;
  bool saw_ghost_rel = false;
  for (pugi::xml_node rel = rels_root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    const std::string_view type = rel.attribute("Type").value();
    if (type == "http://schemas.microsoft.com/office/2006/relationships/vbaProject") {
      saw_vba_rel = true;
      EXPECT_EQ(std::string_view(rel.attribute("Target").value()), "vbaProject.bin");
    } else if (type == "http://schemas.microsoft.com/office/2006/relationships/customUI") {
      saw_ghost_rel = true;
    }
  }
  EXPECT_TRUE(saw_vba_rel) << "vbaProject relationship missing from emitted workbook rels";
  EXPECT_FALSE(saw_ghost_rel) << "relationship to absent part must be dropped, not dangling";
}

// -- (c) Worksheet drawing + image round-trip --------------------------------

std::vector<std::uint8_t> BuildDrawingPackage() {
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Default Extension=\"png\" ContentType=\"image/png\"/>\n"
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
      "  <sheets>\n"
      "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
      "  </sheets>\n"
      "</workbook>\n";
  const std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheetData/>\n"
      "  <drawing r:id=\"rId1\"/>\n"
      "</worksheet>\n";
  const std::string_view sheet_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" "
      "Target=\"../drawings/drawing1.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view drawing_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>\n";
  const std::string_view drawing_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" "
      "Target=\"../media/image1.png\"/>\n"
      "</Relationships>\n";
  const std::string_view image_png = std::string_view("\x89PNG\r\n\x1a\n fake png bytes", 22);

  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/worksheets/_rels/sheet1.xml.rels", sheet_rels},
      {"xl/drawings/drawing1.xml", drawing_xml},
      {"xl/drawings/_rels/drawing1.xml.rels", drawing_rels},
      {"xl/media/image1.png", image_png},
  });
}

TEST(OoxmlDefaultPassthrough, DrawingAndMediaRoundTrip) {
  const std::vector<std::uint8_t> source = BuildDrawingPackage();
  auto load_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;

  // Sheet captured the drawing reference target.
  const Workbook& wb = load_or.value().workbook;
  ASSERT_EQ(wb.sheet_count(), 1U);
  EXPECT_EQ(wb.sheet(0).drawing_rel_target(), "xl/drawings/drawing1.xml");

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  const std::vector<std::uint8_t>& out = save_or.value();

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(out))));

  // Drawing part, its rels, and the media all survive verbatim.
  ASSERT_TRUE(zip.has_entry("xl/drawings/drawing1.xml"));
  ASSERT_TRUE(zip.has_entry("xl/drawings/_rels/drawing1.xml.rels"));
  ASSERT_TRUE(zip.has_entry("xl/media/image1.png"));
  const std::vector<std::uint8_t> png = ReadPart(zip, "xl/media/image1.png");
  const std::string_view expected_png = std::string_view("\x89PNG\r\n\x1a\n fake png bytes", 22);
  ASSERT_EQ(png.size(), expected_png.size());
  EXPECT_EQ(0, std::memcmp(png.data(), expected_png.data(), expected_png.size()));

  // Worksheet body re-emits the <drawing r:id> reference.
  const std::vector<std::uint8_t> sheet = ReadPart(zip, "xl/worksheets/sheet1.xml");
  pugi::xml_document sheet_doc;
  ASSERT_TRUE(sheet_doc.load_buffer(sheet.data(), sheet.size()));
  pugi::xml_node ws = sheet_doc.child("worksheet");
  ASSERT_TRUE(ws);
  pugi::xml_node drawing = ws.child("drawing");
  ASSERT_TRUE(drawing) << "worksheet <drawing> element missing after round-trip";
  const std::string drawing_rid = drawing.attribute("r:id").value();
  ASSERT_FALSE(drawing_rid.empty());

  // Sheet rels carries a drawing relationship whose rId matches the body
  // reference and whose target resolves to the drawing part.
  const std::vector<std::uint8_t> sheet_rels = ReadPart(zip, "xl/worksheets/_rels/sheet1.xml.rels");
  pugi::xml_document sr_doc;
  ASSERT_TRUE(sr_doc.load_buffer(sheet_rels.data(), sheet_rels.size()));
  pugi::xml_node sr_root = sr_doc.child("Relationships");
  ASSERT_TRUE(sr_root);
  bool saw_drawing_rel = false;
  for (pugi::xml_node rel = sr_root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    if (std::string_view(rel.attribute("Type").value()) ==
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing") {
      saw_drawing_rel = true;
      EXPECT_EQ(std::string_view(rel.attribute("Id").value()), drawing_rid);
      EXPECT_EQ(std::string_view(rel.attribute("Target").value()), "../drawings/drawing1.xml");
    }
  }
  EXPECT_TRUE(saw_drawing_rel) << "drawing relationship missing from sheet rels";

  // `<Default Extension="png">` round-tripped so the image resolves.
  const std::vector<std::uint8_t> ct = ReadPart(zip, "[Content_Types].xml");
  pugi::xml_document ct_doc;
  ASSERT_TRUE(ct_doc.load_buffer(ct.data(), ct.size()));
  bool saw_png_default = false;
  for (pugi::xml_node node = ct_doc.child("Types").child("Default"); node; node = node.next_sibling("Default")) {
    if (std::string_view(node.attribute("Extension").value()) == "png") {
      saw_png_default = true;
    }
  }
  EXPECT_TRUE(saw_png_default) << "Default Extension=png missing from round-tripped content types";
}

// -- (d) VML part round-trip -------------------------------------------------

std::vector<std::uint8_t> BuildVmlPackage() {
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>\n"
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
      "  <sheets>\n"
      "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
      "  </sheets>\n"
      "</workbook>\n";
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
  const std::string_view vml = "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:shape id=\"note\"/></xml>\n";

  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/drawings/vmlDrawing1.vml", vml},
  });
}

TEST(OoxmlDefaultPassthrough, VmlPartRoundTrip) {
  const std::vector<std::uint8_t> source = BuildVmlPackage();
  auto load_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;

  auto save_or = load_or.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  const std::vector<std::uint8_t>& out = save_or.value();

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(out))));

  // VML bytes survive verbatim (no comments regenerate the part, so it
  // rides through passthrough).
  ASSERT_TRUE(zip.has_entry("xl/drawings/vmlDrawing1.vml"));
  const std::vector<std::uint8_t> vml_bytes = ReadPart(zip, "xl/drawings/vmlDrawing1.vml");
  const std::string_view expected = "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:shape id=\"note\"/></xml>\n";
  ASSERT_EQ(vml_bytes.size(), expected.size());
  EXPECT_EQ(0, std::memcmp(vml_bytes.data(), expected.data(), expected.size()));

  // `<Default Extension="vml">` round-tripped.
  const std::vector<std::uint8_t> ct = ReadPart(zip, "[Content_Types].xml");
  pugi::xml_document ct_doc;
  ASSERT_TRUE(ct_doc.load_buffer(ct.data(), ct.size()));
  bool saw_vml_default = false;
  for (pugi::xml_node node = ct_doc.child("Types").child("Default"); node; node = node.next_sibling("Default")) {
    if (std::string_view(node.attribute("Extension").value()) == "vml") {
      saw_vml_default = true;
    }
  }
  EXPECT_TRUE(saw_vml_default) << "Default Extension=vml missing from round-tripped content types";
}

}  // namespace
}  // namespace formulon
