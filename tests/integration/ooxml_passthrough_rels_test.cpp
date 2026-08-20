//
// Integration test: workbook-level `<Relationship>` entries whose Type
// URI the reader does not recognise (theme, calcChain, vbaProject,
// customXml, ...) must round-trip through `xl/_rels/workbook.xml.rels`.
// Without this, the matching `<Override>`-listed parts survive in
// passthrough but become orphans in the package graph and Excel opens
// the result in "needs repair" mode.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_reader.h"
#include "io/passthrough_part.h"
#include "io/tables_reader.h"
#include "io/unknown_relationship.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "value.h"
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

/// A worksheet whose rels contain an extension relationship which no
/// Formulon sheet feature consumes. Its target must remain reachable after
/// save even though the sheet itself is rebuilt.
std::vector<std::uint8_t> BuildPackageWithUnknownSheetRelationship() {
  const std::string_view content_types =
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
      "<Override PartName=\"/xl/controls/control1.xml\" ContentType=\"application/vnd.example.control+xml\"/>"
      "</Types>";
  const std::string_view package_rels =
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>"
      "</Relationships>";
  const std::string_view workbook_xml =
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";
  const std::string_view workbook_rels =
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>"
      "</Relationships>";
  const std::string_view sheet_xml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<sheetData/><controls><control shapeId=\"1\" r:id=\"rId9\"/></controls></worksheet>";
  const std::string_view sheet_rels =
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId9\" Type=\"urn:example:relationships/control\" Target=\"../controls/control1.xml\"/>"
      "</Relationships>";
  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/worksheets/_rels/sheet1.xml.rels", sheet_rels},
      {"xl/controls/control1.xml", "<control id=\"1\"/>"},
  });
}

/// A worksheet next to a chartsheet. Chartsheets appear in workbook.xml's
/// ordinary `<sheets>` sequence but cannot be parsed by the worksheet
/// reader, so this is the minimal regression shape for opaque sheets.
std::vector<std::uint8_t> BuildPackageWithChartSheet() {
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
      "<Override PartName=\"/xl/chartsheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml\"/>"
      "</Types>";
  const std::string_view package_rels =
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>"
      "</Relationships>";
  const std::string_view workbook_xml =
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/>"
      "<sheet name=\"Chart\" sheetId=\"2\" state=\"hidden\" r:id=\"rId2\"/></sheets></workbook>";
  const std::string_view workbook_rels =
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>"
      "<Relationship Id=\"rId2\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet\" "
      "Target=\"chartsheets/sheet1.xml\"/>"
      "</Relationships>";
  const std::string_view sheet_xml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>";
  const std::string_view chart_sheet_xml =
      "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetViews/></chartsheet>";
  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/chartsheets/sheet1.xml", chart_sheet_xml},
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

TEST(OoxmlPassthroughRels, UnknownSheetRelationshipsAndTargetsRoundTrip) {
  auto load_or = io::read_ooxml(SpanOf(BuildPackageWithUnknownSheetRelationship()));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;
  ASSERT_EQ(load_or.value().workbook.sheet(0).unknown_relationships().size(), 1U);
  const io::UnknownRelationship& relationship = load_or.value().workbook.sheet(0).unknown_relationships().front();
  EXPECT_EQ(relationship.id, "rId9");
  EXPECT_EQ(relationship.type, "urn:example:relationships/control");
  EXPECT_EQ(relationship.target, "xl/controls/control1.xml");
  const WorksheetRawExtensions& raw = load_or.value().workbook.sheet(0).raw_extensions();
  const auto controls = std::find_if(raw.begin(), raw.end(), [](const WorksheetRawChild& child) {
    return child.slot == worksheet_child::slot_of("controls");
  });
  ASSERT_NE(controls, raw.end());
  EXPECT_NE(controls->xml.find("r:id=\"rId9\""), std::string::npos);

  auto save_or = load_or.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  ASSERT_TRUE(zip.has_entry("xl/controls/control1.xml"));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  const std::string sheet_xml(sheet_or.value().begin(), sheet_or.value().end());
  EXPECT_NE(sheet_xml.find("<controls><control shapeId=\"1\" r:id=\"rId9\"/></controls>"), std::string::npos);
  auto rels_or = zip.read_entry("xl/worksheets/_rels/sheet1.xml.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  const std::string rels(rels_or.value().begin(), rels_or.value().end());
  EXPECT_NE(rels.find("Id=\"rId9\" Type=\"urn:example:relationships/control\" Target=\"../controls/control1.xml\""),
            std::string::npos);
}

TEST(OoxmlPassthroughRels, ChartSheetIsAcceptedAndRoundTripsAsOpaqueSheet) {
  const std::vector<std::uint8_t> source_bytes = BuildPackageWithChartSheet();
  auto load_or = io::read_ooxml(SpanOf(source_bytes));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;
  const Workbook& wb = load_or.value().workbook;
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(0).name(), "Data");
  EXPECT_EQ(wb.sheet(1).name(), "Chart");
  EXPECT_TRUE(wb.sheet(1).is_opaque_ooxml_sheet());
  EXPECT_TRUE(wb.sheet(1).view().tab_hidden);
  EXPECT_EQ(wb.sheet(1).opaque_ooxml_part_path(), "xl/chartsheets/sheet1.xml");

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  ASSERT_TRUE(zip.has_entry("xl/chartsheets/sheet1.xml"));
  auto chart_or = zip.read_entry("xl/chartsheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(chart_or));
  EXPECT_EQ(
      std::string(chart_or.value().begin(), chart_or.value().end()),
      "<chartsheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetViews/></chartsheet>");

  auto rels_or = zip.read_entry("xl/_rels/workbook.xml.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  const std::string rels(rels_or.value().begin(), rels_or.value().end());
  EXPECT_NE(rels.find("relationships/chartsheet\" Target=\"chartsheets/sheet1.xml\""), std::string::npos);

  auto roundtrip_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(roundtrip_or)) << "second read failed: " << roundtrip_or.error().message;
  ASSERT_EQ(roundtrip_or.value().workbook.sheet_count(), 2U);
  EXPECT_TRUE(roundtrip_or.value().workbook.sheet(1).is_opaque_ooxml_sheet());
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
  // calcChain is deliberately dropped on save: it is a recalculation-order
  // cache that goes stale the moment any cell value changes, and a stale
  // calcChain makes real Excel reject / "repair" the workbook. The reader
  // still captures its rel (see UnknownWorkbookRelsAreCaptured), but the
  // writer must not re-emit it.
  EXPECT_FALSE(saw_calc_chain_rel) << "stale calcChain relationship must be dropped on save";

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
  // calcChain part is dropped, so its Override must not be emitted either.
  EXPECT_FALSE(saw_calc_chain_override) << "dropped calcChain must not leave an Override in [Content_Types].xml";
}

// ---------------------------------------------------------------------------
// Sheet-rels rId integrity: the invariants below hold regardless of which
// relationship types share a sheet's rels file — every r:id the worksheet
// body references must resolve to a `<Relationship>` in that same sheet's
// rels output, every rels file the writer emits must declare at least one
// relationship, and no two zip entries may share a path.
// ---------------------------------------------------------------------------

/// Recursively collects every `r:id` attribute value under `node`, in
/// document order. Used to check the property "every r:id the body
/// references resolves to a Relationship Id" without hardcoding which
/// specific elements carry one.
void CollectRIds(const pugi::xml_node& node, std::vector<std::string>& out) {
  for (pugi::xml_attribute attr = node.first_attribute(); attr; attr = attr.next_attribute()) {
    if (std::string_view(attr.name()) == "r:id") {
      out.emplace_back(attr.value());
    }
  }
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    CollectRIds(child, out);
  }
}

/// Collects every `Id` attribute from a `<Relationships>` document's
/// direct `<Relationship>` children.
std::vector<std::string> CollectRelationshipIds(const pugi::xml_node& relationships_root) {
  std::vector<std::string> ids;
  for (pugi::xml_node rel = relationships_root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    ids.emplace_back(rel.attribute("Id").value());
  }
  return ids;
}

TEST(OoxmlSheetRelsIntegrity, TablePartRidMatchesItsOwnTableRelationship) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0U, 0U, Value::text("Region"));
  s.set_cell_value(1U, 0U, Value::text("North"));

  io::TableMetadata table;
  table.id = 1U;
  table.name = "Regions";
  table.display_name = "Regions";
  table.ref = "A1:A2";
  table.sheet_index = 0U;
  table.header_row = true;
  table.totals_row = false;
  table.columns = {io::TableColumn{1U, "Region", "", "", ""}};
  wb.set_tables({std::move(table)});

  // An unknown relationship reserving the lowest rId ("rId1"): a writer
  // that hardcodes the table's r:id to its positional rId(i+1) would
  // collide with it instead of naming the id this rels file actually
  // assigned to the table relationship.
  io::UnknownRelationship unknown;
  unknown.id = "rId1";
  unknown.type = "urn:example:relationships/control";
  unknown.target = "xl/controls/control1.xml";
  s.set_unknown_relationships({unknown});

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  auto sheet_bytes_or = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(sheet_bytes_or));
  pugi::xml_document sheet_doc;
  ASSERT_TRUE(sheet_doc.load_buffer(sheet_bytes_or.value().data(), sheet_bytes_or.value().size()));
  pugi::xml_node table_part = sheet_doc.child("worksheet").child("tableParts").child("tablePart");
  ASSERT_TRUE(table_part) << "no <tablePart> in the saved worksheet body";
  const std::string table_rid = table_part.attribute("r:id").value();
  ASSERT_FALSE(table_rid.empty());
  // The reserved unknown-relationship id must not have been silently
  // reassigned to the table.
  EXPECT_NE(table_rid, "rId1");

  auto rels_bytes_or = zip.read_entry("xl/worksheets/_rels/sheet1.xml.rels");
  ASSERT_TRUE(static_cast<bool>(rels_bytes_or));
  pugi::xml_document rels_doc;
  ASSERT_TRUE(rels_doc.load_buffer(rels_bytes_or.value().data(), rels_bytes_or.value().size()));
  bool found = false;
  for (pugi::xml_node rel = rels_doc.child("Relationships").child("Relationship"); rel;
       rel = rel.next_sibling("Relationship")) {
    if (std::string_view(rel.attribute("Id").value()) == table_rid) {
      found = true;
      EXPECT_EQ(std::string_view(rel.attribute("Type").value()), io::kRelTable)
          << "the tablePart's r:id must resolve to a table-typed relationship";
    }
  }
  EXPECT_TRUE(found) << "tablePart r:id \"" << table_rid << "\" has no matching Relationship";
}

TEST(OoxmlSheetRelsIntegrity, MissingInternalTargetIsDroppedButValidPassthroughSurvives) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);

  io::UnknownRelationship missing;
  missing.id = "rId7";
  missing.type = "urn:example:relationships/missing-control";
  missing.target = "xl/controls/missing.xml";
  io::UnknownRelationship valid;
  valid.id = "rId8";
  valid.type = "urn:example:relationships/control";
  valid.target = "xl/controls/control1.xml";
  s.set_unknown_relationships({missing, valid});

  io::PassthroughPart control;
  control.path = "xl/controls/control1.xml";
  control.content_type = "application/vnd.example.control+xml";
  control.bytes = {'<', 'c', 'o', 'n', 't', 'r', 'o', 'l', '/', '>'};
  wb.set_passthrough_parts({std::move(control)});

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  ASSERT_TRUE(zip.has_entry("xl/controls/control1.xml"));
  auto rels_or = zip.read_entry("xl/worksheets/_rels/sheet1.xml.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  pugi::xml_document rels_doc;
  ASSERT_TRUE(rels_doc.load_buffer(rels_or.value().data(), rels_or.value().size()));
  pugi::xml_node relationships = rels_doc.child("Relationships");
  ASSERT_TRUE(relationships);
  ASSERT_EQ(CollectRelationshipIds(relationships).size(), 1U);
  pugi::xml_node rel = relationships.child("Relationship");
  ASSERT_TRUE(rel);
  EXPECT_EQ(rel.attribute("Id").value(), std::string("rId8"));
  EXPECT_EQ(rel.attribute("Target").value(), std::string("../controls/control1.xml"));
  EXPECT_EQ(rel.next_sibling("Relationship"), pugi::xml_node());
  EXPECT_EQ(std::string(rels_or.value().begin(), rels_or.value().end()).find("missing.xml"), std::string::npos);
}

TEST(OoxmlSheetRelsIntegrity, OnlyMissingInternalTargetProducesNoRelsPart) {
  Workbook wb = Workbook::create();
  io::UnknownRelationship missing;
  missing.id = "rId7";
  missing.type = "urn:example:relationships/missing-control";
  missing.target = "xl/controls/missing.xml";
  wb.sheet(0).set_unknown_relationships({missing});

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  EXPECT_FALSE(zip.has_entry("xl/worksheets/_rels/sheet1.xml.rels"));
}

TEST(OoxmlSheetRelsIntegrity, ExternalUnknownRelationshipSurvivesWithoutPayload) {
  Workbook wb = Workbook::create();
  io::UnknownRelationship external;
  external.id = "rId9";
  external.type = "urn:example:relationships/external-control";
  external.target = "https://example.test/controls/control1.xml";
  external.target_external = true;
  wb.sheet(0).set_unknown_relationships({external});

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  auto rels_or = zip.read_entry("xl/worksheets/_rels/sheet1.xml.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  pugi::xml_document rels_doc;
  ASSERT_TRUE(rels_doc.load_buffer(rels_or.value().data(), rels_or.value().size()));
  const pugi::xml_node rel = rels_doc.child("Relationships").child("Relationship");
  ASSERT_TRUE(rel);
  EXPECT_EQ(std::string_view(rel.attribute("Id").value()), "rId9");
  EXPECT_EQ(std::string_view(rel.attribute("Type").value()), "urn:example:relationships/external-control");
  EXPECT_EQ(std::string_view(rel.attribute("Target").value()), "https://example.test/controls/control1.xml");
  EXPECT_EQ(std::string_view(rel.attribute("TargetMode").value()), "External");
  EXPECT_EQ(rel.next_sibling("Relationship"), pugi::xml_node());
}

TEST(OoxmlSheetRelsIntegrity, MissingInternalTargetDropsStrayReservedRelsPassthrough) {
  Workbook wb = Workbook::create();
  io::UnknownRelationship missing;
  missing.id = "rId7";
  missing.type = "urn:example:relationships/missing-control";
  missing.target = "xl/controls/missing.xml";
  wb.sheet(0).set_unknown_relationships({missing});

  io::PassthroughPart stray_rels;
  stray_rels.path = "xl/worksheets/_rels/sheet1.xml.rels";
  stray_rels.bytes = {'b', 'o', 'g', 'u', 's'};
  wb.set_passthrough_parts({std::move(stray_rels)});

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  const std::vector<std::string> entries = zip.list_entries();
  EXPECT_EQ(std::count(entries.begin(), entries.end(), std::string("xl/worksheets/_rels/sheet1.xml.rels")), 0);
}

TEST(OoxmlSheetRelsIntegrity, GeneratedStylesCollisionDropsUnknownRelationship) {
  Workbook wb = Workbook::create();
  io::UnknownRelationship unknown;
  unknown.id = "rId11";
  unknown.type = "urn:example:relationships/alternate-styles";
  unknown.target = "xl/styles.xml";
  wb.sheet(0).set_unknown_relationships({unknown});

  io::PassthroughPart colliding_styles;
  colliding_styles.path = "xl/styles.xml";
  colliding_styles.content_type = "application/vnd.example.styles+xml";
  colliding_styles.bytes = {'b', 'o', 'g', 'u', 's'};
  wb.set_passthrough_parts({std::move(colliding_styles)});

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  ASSERT_TRUE(zip.has_entry("xl/styles.xml"));
  auto styles_or = zip.read_entry("xl/styles.xml");
  ASSERT_TRUE(static_cast<bool>(styles_or));
  EXPECT_NE(std::string(styles_or.value().begin(), styles_or.value().end()), "bogus");
  EXPECT_FALSE(zip.has_entry("xl/worksheets/_rels/sheet1.xml.rels"));
}

/// Builds a package with a worksheet carrying two `kRelVmlDrawing`
/// relationships: one bound to `<legacyDrawing>` (comment geometry) and
/// one bound to `<legacyDrawingHF>` (header/footer image), plus a
/// `<comments>` part. `with_comments = false` drops the comments part
/// and the `<legacyDrawing>` element, leaving only the header/footer
/// VML — the "logo without comments" shape.
std::vector<std::uint8_t> BuildPackageWithTwoVmlDrawings(bool with_comments) {
  std::string content_types =
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>"
      "<Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
      "</Types>";
  const std::string_view package_rels =
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>"
      "</Relationships>";
  const std::string_view workbook_xml =
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";
  const std::string_view workbook_rels =
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>"
      "</Relationships>";

  std::string sheet_xml =
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData/>";
  std::string sheet_rels = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
  if (with_comments) {
    sheet_xml += "<legacyDrawing r:id=\"rId2\"/>";
    sheet_rels +=
        "<Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" "
        "Target=\"../comments1.xml\"/>"
        "<Relationship Id=\"rId2\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" "
        "Target=\"../drawings/vmlDrawing1.vml\"/>";
  }
  sheet_xml += "<legacyDrawingHF r:id=\"rId3\"/></worksheet>";
  sheet_rels +=
      "<Relationship Id=\"rId3\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" "
      "Target=\"../drawings/vmlDrawing2.vml\"/></Relationships>";

  const std::string_view comments_xml =
      "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<authors><author>Alice</author></authors>"
      "<commentList><comment ref=\"A1\" authorId=\"0\"><text><t>hi</t></text></comment></commentList>"
      "</comments>";
  const std::string_view comment_vml =
      "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:shape id=\"comment-shape\"/></xml>";
  const std::string_view header_footer_vml =
      "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:shape id=\"hf-shape\"/></xml>";

  std::vector<PartFile> parts = {
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/worksheets/_rels/sheet1.xml.rels", sheet_rels},
      {"xl/drawings/vmlDrawing2.vml", header_footer_vml},
  };
  if (with_comments) {
    parts.push_back({"xl/comments1.xml", comments_xml});
    parts.push_back({"xl/drawings/vmlDrawing1.vml", comment_vml});
  }
  return BuildZip(parts);
}

TEST(OoxmlSheetRelsIntegrity, TwoVmlDrawingRelationshipsBothSurviveRoundTrip) {
  const std::vector<std::uint8_t> source = BuildPackageWithTwoVmlDrawings(/*with_comments=*/true);
  auto load_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;
  const Sheet& loaded = load_or.value().workbook.sheet(0);
  // The comment geometry VML is the one <legacyDrawing> names, not
  // whichever kRelVmlDrawing relationship happened to be read last.
  EXPECT_EQ(loaded.comment_vml_path(), "xl/drawings/vmlDrawing1.vml");
  ASSERT_EQ(loaded.comments().size(), 1U);

  // The header/footer VML relationship is preserved as an unknown
  // relationship instead of being silently dropped.
  bool saw_hf_relationship = false;
  for (const io::UnknownRelationship& rel : loaded.unknown_relationships()) {
    if (rel.id == "rId3") {
      saw_hf_relationship = true;
      EXPECT_EQ(rel.type, io::kRelVmlDrawing);
      EXPECT_EQ(rel.target, "xl/drawings/vmlDrawing2.vml");
    }
  }
  EXPECT_TRUE(saw_hf_relationship) << "legacyDrawingHF's vmlDrawing relationship was dropped instead of preserved";

  auto save_or = load_or.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));

  // Both VML parts round-trip with distinct bytes: comment geometry is
  // not overwritten by (or confused with) the header/footer image.
  ASSERT_TRUE(zip.has_entry("xl/drawings/vmlDrawing1.vml"));
  ASSERT_TRUE(zip.has_entry("xl/drawings/vmlDrawing2.vml"));
  auto comment_vml_or = zip.read_entry("xl/drawings/vmlDrawing1.vml");
  auto hf_vml_or = zip.read_entry("xl/drawings/vmlDrawing2.vml");
  ASSERT_TRUE(static_cast<bool>(comment_vml_or));
  ASSERT_TRUE(static_cast<bool>(hf_vml_or));
  const std::string comment_vml_text(comment_vml_or.value().begin(), comment_vml_or.value().end());
  const std::string hf_vml_text(hf_vml_or.value().begin(), hf_vml_or.value().end());
  EXPECT_NE(comment_vml_text.find("comment-shape"), std::string::npos);
  EXPECT_NE(hf_vml_text.find("hf-shape"), std::string::npos);

  // Property assertion: every r:id in the saved worksheet body resolves
  // to a Relationship Id declared in that sheet's own rels file.
  auto sheet_bytes_or = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(sheet_bytes_or));
  pugi::xml_document sheet_doc;
  ASSERT_TRUE(sheet_doc.load_buffer(sheet_bytes_or.value().data(), sheet_bytes_or.value().size()));
  std::vector<std::string> body_rids;
  CollectRIds(sheet_doc.child("worksheet"), body_rids);
  ASSERT_FALSE(body_rids.empty());

  ASSERT_TRUE(zip.has_entry("xl/worksheets/_rels/sheet1.xml.rels"));
  auto rels_bytes_or = zip.read_entry("xl/worksheets/_rels/sheet1.xml.rels");
  ASSERT_TRUE(static_cast<bool>(rels_bytes_or));
  pugi::xml_document rels_doc;
  ASSERT_TRUE(rels_doc.load_buffer(rels_bytes_or.value().data(), rels_bytes_or.value().size()));
  const std::vector<std::string> rel_ids = CollectRelationshipIds(rels_doc.child("Relationships"));

  for (const std::string& rid : body_rids) {
    EXPECT_NE(std::find(rel_ids.begin(), rel_ids.end(), rid), rel_ids.end())
        << "worksheet body r:id \"" << rid << "\" has no matching Relationship in the sheet's own rels file";
  }
}

TEST(OoxmlSheetRelsIntegrity, HeaderFooterVmlWithoutCommentsSurvivesRoundTrip) {
  const std::vector<std::uint8_t> source = BuildPackageWithTwoVmlDrawings(/*with_comments=*/false);
  auto load_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;
  const Sheet& loaded = load_or.value().workbook.sheet(0);
  // No comments on this sheet, so no relationship is modelled as the
  // comment-VML slot.
  EXPECT_TRUE(loaded.comment_vml_path().empty());
  EXPECT_TRUE(loaded.comments().empty());

  bool saw_hf_relationship = false;
  for (const io::UnknownRelationship& rel : loaded.unknown_relationships()) {
    if (rel.id == "rId3") {
      saw_hf_relationship = true;
      EXPECT_EQ(rel.type, io::kRelVmlDrawing);
      EXPECT_EQ(rel.target, "xl/drawings/vmlDrawing2.vml");
    }
  }
  EXPECT_TRUE(saw_hf_relationship)
      << "the sole vmlDrawing relationship on a comment-less sheet must not be silently dropped";

  auto save_or = load_or.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));

  ASSERT_TRUE(zip.has_entry("xl/worksheets/_rels/sheet1.xml.rels"));
  auto rels_bytes_or = zip.read_entry("xl/worksheets/_rels/sheet1.xml.rels");
  ASSERT_TRUE(static_cast<bool>(rels_bytes_or));
  const std::string rels_text(rels_bytes_or.value().begin(), rels_bytes_or.value().end());
  EXPECT_NE(rels_text.find("Id=\"rId3\""), std::string::npos)
      << "header/footer vmlDrawing relationship missing from the saved rels file";
  EXPECT_NE(rels_text.find("../drawings/vmlDrawing2.vml"), std::string::npos);

  ASSERT_TRUE(zip.has_entry("xl/drawings/vmlDrawing2.vml"));
  auto vml_bytes_or = zip.read_entry("xl/drawings/vmlDrawing2.vml");
  ASSERT_TRUE(static_cast<bool>(vml_bytes_or));
  const std::string vml_text(vml_bytes_or.value().begin(), vml_bytes_or.value().end());
  EXPECT_NE(vml_text.find("hf-shape"), std::string::npos);

  auto sheet_bytes_or = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(sheet_bytes_or));
  const std::string sheet_text(sheet_bytes_or.value().begin(), sheet_bytes_or.value().end());
  EXPECT_NE(sheet_text.find("legacyDrawingHF r:id=\"rId3\""), std::string::npos)
      << "the worksheet body's legacyDrawingHF element must keep the rId its rels entry uses";
}

TEST(OoxmlSheetRelsIntegrity, UnknownRelationshipsOnlySheetDoesNotDuplicateRelsEntry) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  io::UnknownRelationship unknown;
  unknown.id = "rId9";
  unknown.type = "urn:example:relationships/control";
  unknown.target = "xl/controls/control1.xml";
  s.set_unknown_relationships({unknown});

  // A stray passthrough part deliberately placed at the exact path the
  // writer generates for this sheet's own rels file.
  std::vector<io::PassthroughPart> parts = wb.passthrough_parts();
  io::PassthroughPart control;
  control.path = "xl/controls/control1.xml";
  control.content_type = "application/vnd.example.control+xml";
  control.bytes = {'<', 'c', 'o', 'n', 't', 'r', 'o', 'l', '/', '>'};
  parts.push_back(std::move(control));
  io::PassthroughPart bogus;
  bogus.path = "xl/worksheets/_rels/sheet1.xml.rels";
  bogus.bytes = {'b', 'o', 'g', 'u', 's'};
  parts.push_back(bogus);
  wb.set_passthrough_parts(std::move(parts));

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;

  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  const std::vector<std::string> entries = zip.list_entries();
  const std::size_t count = static_cast<std::size_t>(
      std::count(entries.begin(), entries.end(), std::string("xl/worksheets/_rels/sheet1.xml.rels")));
  EXPECT_EQ(count, 1U) << "the sheet rels path must appear exactly once in the saved package";

  // The surviving entry is the writer's own generated rels (the unknown
  // relationship id survives), not the bogus passthrough bytes.
  auto rels_or = zip.read_entry("xl/worksheets/_rels/sheet1.xml.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  const std::string rels(rels_or.value().begin(), rels_or.value().end());
  EXPECT_NE(rels.find("rId9"), std::string::npos);
}

TEST(OoxmlSheetRelsIntegrity, InternalOnlyHyperlinksProduceNoRelsPart) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  Hyperlink h;
  h.row = 0U;
  h.col = 0U;
  h.last_row = 0U;
  h.last_col = 0U;
  h.location = "Sheet1!B2";  // purely internal; target left empty
  h.display = "Go to B2";
  s.mutable_hyperlinks().push_back(h);

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(save_or.value()))));
  EXPECT_FALSE(zip.has_entry("xl/worksheets/_rels/sheet1.xml.rels"))
      << "a sheet whose only hyperlinks are purely internal must get no rels part";

  // The hyperlink itself still round-trips via its inline location=.
  auto reload_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(reload_or)) << "reload failed: " << reload_or.error().message;
  ASSERT_EQ(reload_or.value().workbook.sheet(0).hyperlinks().size(), 1U);
  EXPECT_EQ(reload_or.value().workbook.sheet(0).hyperlinks()[0].location, "Sheet1!B2");
  EXPECT_TRUE(reload_or.value().workbook.sheet(0).hyperlinks()[0].target.empty());
}

}  // namespace
}  // namespace formulon
