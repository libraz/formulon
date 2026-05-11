// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Integration test: workbook-level `<Relationship>` entries whose Type
// URI the reader does not recognise (theme, calcChain, vbaProject,
// customXml, ...) must round-trip through `xl/_rels/workbook.xml.rels`.
// Without this, the matching `<Override>`-listed parts survive in
// passthrough but become orphans in the package graph and Excel opens
// the result in "needs repair" mode.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/passthrough_part.h"
#include "io/unknown_relationship.h"
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

/// Materialises `parts` into a heap-allocated zip archive byte vector
/// via miniz. Mirrors the helper used by `ooxml_metadata_test.cpp`.
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

/// Builds a minimal xlsx package with one worksheet plus theme and
/// calcChain parts; the workbook-rels file references both via
/// relationship types the reader does not consume directly.
std::vector<std::uint8_t> BuildPackageWithThemeAndCalcChain() {
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
      "  <Override PartName=\"/xl/theme/theme1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\n"
      "  <Override PartName=\"/xl/calcChain.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/>\n"
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
      "  <Relationship Id=\"rId2\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" "
      "Target=\"theme/theme1.xml\"/>\n"
      "  <Relationship Id=\"rId3\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain\" "
      "Target=\"calcChain.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "</worksheet>\n";
  // The theme and calcChain bodies just need to be present + well-formed
  // enough that the reader's passthrough capture keeps the bytes; the
  // tests assert on the rels graph, not on the body XML structure.
  const std::string_view theme_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office\"/>\n";
  const std::string_view calc_chain_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<calcChain xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"/>\n";

  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/theme/theme1.xml", theme_xml},
      {"xl/calcChain.xml", calc_chain_xml},
  });
}

TEST(OoxmlPassthroughRels, UnknownWorkbookRelsAreCaptured) {
  const std::vector<std::uint8_t> bytes = BuildPackageWithThemeAndCalcChain();
  auto load_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;
  const std::vector<io::UnknownRelationship>& rels = load_or.value().workbook.unknown_workbook_rels();
  ASSERT_EQ(rels.size(), 2U);

  bool saw_theme = false;
  bool saw_calc_chain = false;
  for (const io::UnknownRelationship& r : rels) {
    EXPECT_FALSE(r.target_external);
    if (r.type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
      saw_theme = true;
      EXPECT_EQ(r.target, "xl/theme/theme1.xml");
    } else if (r.type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain") {
      saw_calc_chain = true;
      EXPECT_EQ(r.target, "xl/calcChain.xml");
    } else {
      ADD_FAILURE() << "unexpected unknown rel type: " << r.type;
    }
  }
  EXPECT_TRUE(saw_theme);
  EXPECT_TRUE(saw_calc_chain);
}

TEST(OoxmlPassthroughRels, UnknownWorkbookRelsRoundTrip) {
  const std::vector<std::uint8_t> source_bytes = BuildPackageWithThemeAndCalcChain();
  auto load_or = io::read_ooxml(SpanOf(source_bytes));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;

  // Save back through the writer and reopen the archive to inspect the
  // raw rels and content-types parts.
  auto save_or = load_or.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  const std::vector<std::uint8_t>& out_bytes = save_or.value();

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(out_bytes))));
  ASSERT_TRUE(zip.has_entry("xl/_rels/workbook.xml.rels"));
  auto rels_bytes_or = zip.read_entry("xl/_rels/workbook.xml.rels");
  ASSERT_TRUE(static_cast<bool>(rels_bytes_or));
  const std::vector<std::uint8_t>& rels_bytes = rels_bytes_or.value();

  pugi::xml_document rels_doc;
  ASSERT_TRUE(rels_doc.load_buffer(rels_bytes.data(), rels_bytes.size()));
  pugi::xml_node root = rels_doc.child("Relationships");
  ASSERT_TRUE(root);

  bool saw_theme_rel = false;
  bool saw_calc_chain_rel = false;
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    const std::string_view type = rel.attribute("Type").value();
    const std::string_view target = rel.attribute("Target").value();
    if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme") {
      saw_theme_rel = true;
      EXPECT_EQ(target, "theme/theme1.xml");
    } else if (type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain") {
      saw_calc_chain_rel = true;
      EXPECT_EQ(target, "calcChain.xml");
    }
  }
  EXPECT_TRUE(saw_theme_rel) << "theme relationship missing from emitted workbook rels";
  EXPECT_TRUE(saw_calc_chain_rel) << "calcChain relationship missing from emitted workbook rels";

  // The matching `<Override>` entries in `[Content_Types].xml` must
  // also survive (existing passthrough behaviour).
  ASSERT_TRUE(zip.has_entry("[Content_Types].xml"));
  auto ct_bytes_or = zip.read_entry("[Content_Types].xml");
  ASSERT_TRUE(static_cast<bool>(ct_bytes_or));
  const std::vector<std::uint8_t>& ct_bytes = ct_bytes_or.value();
  pugi::xml_document ct_doc;
  ASSERT_TRUE(ct_doc.load_buffer(ct_bytes.data(), ct_bytes.size()));
  pugi::xml_node ct_root = ct_doc.child("Types");
  ASSERT_TRUE(ct_root);

  bool saw_theme_override = false;
  bool saw_calc_chain_override = false;
  for (pugi::xml_node node = ct_root.child("Override"); node; node = node.next_sibling("Override")) {
    const std::string_view part_name = node.attribute("PartName").value();
    if (part_name == "/xl/theme/theme1.xml") {
      saw_theme_override = true;
    } else if (part_name == "/xl/calcChain.xml") {
      saw_calc_chain_override = true;
    }
  }
  EXPECT_TRUE(saw_theme_override) << "theme Override missing from [Content_Types].xml";
  EXPECT_TRUE(saw_calc_chain_override) << "calcChain Override missing from [Content_Types].xml";
}

}  // namespace
}  // namespace formulon
