//
// Integration tests for Bundle 2.4 metadata (defined names + tables).
// We assemble synthetic in-memory `.xlsx` packages via miniz and feed
// them through `read_ooxml`, then verify the metadata lands on the
// workbook and that consumed table parts no longer appear in
// `unknown_parts`.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/ooxml_reader.h"
#include "io/tables_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
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
/// via miniz. Mirrors the pattern used by the existing roundtrip suite.
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

std::string ExtractEntry(const std::vector<std::uint8_t>& archive_bytes, std::string_view name) {
  mz_zip_archive reader{};
  if (mz_zip_reader_init_mem(&reader, archive_bytes.data(), archive_bytes.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_reader_init_mem failed";
    return {};
  }
  const int index = mz_zip_reader_locate_file(&reader, std::string(name).c_str(), nullptr, 0);
  if (index < 0) {
    ADD_FAILURE() << "entry not found: " << name;
    mz_zip_reader_end(&reader);
    return {};
  }
  std::size_t extracted_size = 0;
  void* extracted = mz_zip_reader_extract_to_heap(&reader, static_cast<mz_uint>(index), &extracted_size, 0);
  if (extracted == nullptr) {
    ADD_FAILURE() << "extract_to_heap failed for: " << name;
    mz_zip_reader_end(&reader);
    return {};
  }
  std::string body(static_cast<const char*>(extracted), extracted_size);
  mz_free(extracted);
  mz_zip_reader_end(&reader);
  return body;
}

// ---------------------------------------------------------------------------
// DefinedName preservation through full read pipeline
// ---------------------------------------------------------------------------

TEST(OoxmlMetadata, DefinedNamesLandOnWorkbook) {
  // Two sheets, two defined names: one workbook-scoped, one
  // sheet-scoped (localSheetId=1).
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
      "  <Override PartName=\"/xl/worksheets/sheet2.xml\" "
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
      "    <sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
      "    <sheet name=\"S2\" sheetId=\"2\" r:id=\"rId2\"/>\n"
      "  </sheets>\n"
      "  <definedNames>\n"
      "    <definedName name=\"Sales\">S1!$A$1:$A$10</definedName>\n"
      "    <definedName name=\"LocalRange\" localSheetId=\"1\" hidden=\"true\">S2!$B$1</definedName>\n"
      "  </definedNames>\n"
      "</workbook>\n";
  const std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "  <Relationship Id=\"rId2\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet2.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view sheet1_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "</worksheet>\n";
  const std::string_view sheet2_xml = sheet1_xml;

  const std::vector<std::uint8_t> bytes = BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet1_xml},
      {"xl/worksheets/sheet2.xml", sheet2_xml},
  });

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& wb = result_or.value().workbook;

  ASSERT_EQ(wb.defined_names().size(), 2U);
  EXPECT_EQ(wb.defined_names()[0].name, "Sales");
  EXPECT_EQ(wb.defined_names()[0].formula, "S1!$A$1:$A$10");
  EXPECT_EQ(wb.defined_names()[0].local_sheet_id, -1);
  EXPECT_FALSE(wb.defined_names()[0].hidden);

  EXPECT_EQ(wb.defined_names()[1].name, "LocalRange");
  EXPECT_EQ(wb.defined_names()[1].formula, "S2!$B$1");
  EXPECT_EQ(wb.defined_names()[1].local_sheet_id, 1);
  EXPECT_TRUE(wb.defined_names()[1].hidden);
}

// ---------------------------------------------------------------------------
// Table preservation and unknown_parts hygiene
// ---------------------------------------------------------------------------

TEST(OoxmlMetadata, TablesLandOnWorkbookAndPartIsConsumed) {
  // Single-sheet workbook with one table at xl/tables/table1.xml,
  // referenced via xl/worksheets/_rels/sheet1.xml.rels using a
  // relative `Target="../tables/table1.xml"`.
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
      "  <Override PartName=\"/xl/tables/table1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>\n"
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
      "  <tableParts count=\"1\">\n"
      "    <tablePart r:id=\"rId1\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>\n"
      "  </tableParts>\n"
      "</worksheet>\n";
  const std::string_view sheet_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" "
      "Target=\"../tables/table1.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view table_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"Sales\" "
      "displayName=\"Sales\" ref=\"A1:C5\" totalsRowCount=\"1\">\n"
      "  <tableColumns count=\"3\">\n"
      "    <tableColumn id=\"1\" name=\"Region\" totalsRowLabel=\"Total\"/>\n"
      "    <tableColumn id=\"2\" name=\"Q1\" totalsRowFunction=\"sum\"/>\n"
      "    <tableColumn id=\"3\" name=\"Q2\" totalsRowFunction=\"sum\"/>\n"
      "  </tableColumns>\n"
      "</table>\n";

  const std::vector<std::uint8_t> bytes = BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/worksheets/_rels/sheet1.xml.rels", sheet_rels},
      {"xl/tables/table1.xml", table_xml},
  });

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const io::OoxmlReadResult& result = result_or.value();
  const Workbook& wb = result.workbook;

  // Table metadata landed.
  ASSERT_EQ(wb.tables().size(), 1U);
  const io::TableMetadata& table = wb.tables()[0];
  EXPECT_EQ(table.id, 1U);
  EXPECT_EQ(table.name, "Sales");
  EXPECT_EQ(table.display_name, "Sales");
  EXPECT_EQ(table.ref, "A1:C5");
  EXPECT_EQ(table.sheet_index, 0U);
  EXPECT_TRUE(table.header_row);
  EXPECT_TRUE(table.totals_row);
  ASSERT_EQ(table.columns.size(), 3U);
  EXPECT_EQ(table.columns[0].name, "Region");
  EXPECT_EQ(table.columns[0].totals_label, "Total");
  EXPECT_EQ(table.columns[1].totals_function, "sum");

  // unknown_parts must NOT include the table part — the reader
  // consumes tables and the writer re-emits them via the metadata
  // path, not the passthrough path.
  const std::vector<io::PassthroughPart>& parts = result.unknown_parts;
  const auto leaked = std::find_if(parts.begin(), parts.end(),
                                   [](const io::PassthroughPart& p) { return p.path == "xl/tables/table1.xml"; });
  EXPECT_EQ(leaked, parts.end()) << "table1.xml leaked into unknown_parts";
}

TEST(OoxmlMetadata, EmptyWorkbookHasNoMetadata) {
  // Sanity: a workbook with neither defined names nor table parts must
  // still produce empty (not "missing", not "errored") metadata vectors.
  Workbook src = Workbook::create();
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  const std::vector<std::uint8_t>& bytes = save_or.value();

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const Workbook& dst = result_or.value().workbook;
  EXPECT_TRUE(dst.defined_names().empty());
  EXPECT_TRUE(dst.tables().empty());
}

// ---------------------------------------------------------------------------
// Writer round-trip: defined names, tables, passthrough.
// ---------------------------------------------------------------------------

/// Helper that tools the `read -> set metadata -> write -> read` loop.
/// Returns the bytes produced by the second read so callers can chain a
/// third round-trip when they want to verify byte-stability.
std::vector<std::uint8_t> SaveOrDie(const Workbook& wb) {
  auto save_or = wb.save();
  EXPECT_TRUE(static_cast<bool>(save_or)) << "save() failed: " << save_or.error().message;
  return save_or.value();
}

TEST(OoxmlMetadata, DefinedNamesRoundTripThroughWriter) {
  Workbook src = Workbook::create_empty();
  src.add_sheet("Alpha");
  src.add_sheet("Beta");

  std::vector<io::DefinedName> names;
  io::DefinedName workbook_scope;
  workbook_scope.name = "WB_SCOPE";
  workbook_scope.formula = "Alpha!$A$1:$A$10";
  // Defaults: local_sheet_id = -1, hidden = false, comment empty.
  names.push_back(std::move(workbook_scope));

  io::DefinedName sheet_scope;
  sheet_scope.name = "Local_Range";
  sheet_scope.formula = "Beta!$B$1:$B$5";
  sheet_scope.local_sheet_id = 1;
  sheet_scope.hidden = true;
  sheet_scope.comment = "Beta-only.";
  names.push_back(std::move(sheet_scope));
  src.set_defined_names(std::move(names));

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& dst = result_or.value().workbook;
  ASSERT_EQ(dst.defined_names().size(), 2U);

  EXPECT_EQ(dst.defined_names()[0].name, "WB_SCOPE");
  EXPECT_EQ(dst.defined_names()[0].formula, "Alpha!$A$1:$A$10");
  EXPECT_EQ(dst.defined_names()[0].local_sheet_id, -1);
  EXPECT_FALSE(dst.defined_names()[0].hidden);
  EXPECT_EQ(dst.defined_names()[0].comment, "");

  EXPECT_EQ(dst.defined_names()[1].name, "Local_Range");
  EXPECT_EQ(dst.defined_names()[1].formula, "Beta!$B$1:$B$5");
  EXPECT_EQ(dst.defined_names()[1].local_sheet_id, 1);
  EXPECT_TRUE(dst.defined_names()[1].hidden);
  EXPECT_EQ(dst.defined_names()[1].comment, "Beta-only.");
}

TEST(OoxmlMetadata, TableRoundTripThroughWriter) {
  Workbook src = Workbook::create();  // single sheet "Sheet1"

  io::TableMetadata table;
  table.id = 1;
  table.name = "Sales";
  table.display_name = "Sales";
  table.ref = "A1:C5";
  table.sheet_index = 0;
  table.header_row = true;
  table.totals_row = true;
  table.columns.push_back(io::TableColumn{1, "Region", "Total", "", ""});
  table.columns.push_back(io::TableColumn{2, "Q1", "", "sum", ""});
  table.columns.push_back(io::TableColumn{3, "Q2", "", "sum", ""});

  std::vector<io::TableMetadata> tables;
  tables.push_back(std::move(table));
  src.set_tables(std::move(tables));

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& dst = result_or.value().workbook;
  ASSERT_EQ(dst.tables().size(), 1U);
  const io::TableMetadata& got = dst.tables()[0];
  EXPECT_EQ(got.id, 1U);
  EXPECT_EQ(got.name, "Sales");
  EXPECT_EQ(got.display_name, "Sales");
  EXPECT_EQ(got.ref, "A1:C5");
  EXPECT_EQ(got.sheet_index, 0U);
  EXPECT_TRUE(got.header_row);
  EXPECT_TRUE(got.totals_row);
  ASSERT_EQ(got.columns.size(), 3U);
  EXPECT_EQ(got.columns[0].name, "Region");
  EXPECT_EQ(got.columns[0].totals_label, "Total");
  EXPECT_EQ(got.columns[1].totals_function, "sum");
  EXPECT_EQ(got.columns[2].totals_function, "sum");
}

/// Builds a synthetic in-memory `.xlsx` package whose only "extra" part
/// is `xl/theme/theme1.xml` carrying a tiny stub. The reader captures
/// the bytes; the writer re-emits them; a final read confirms the part
/// is still present and byte-identical.
std::vector<std::uint8_t> BuildXlsxWithThemePart() {
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
  // Distinctive payload so the round-trip assertion is meaningful.
  const std::string_view theme_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Stub\">\n"
      "  <a:themeElements/>\n"
      "</a:theme>\n";

  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/theme/theme1.xml", theme_xml},
  });
}

TEST(OoxmlMetadata, PassthroughPartRoundTripsBytesAndContentType) {
  const std::vector<std::uint8_t> input = BuildXlsxWithThemePart();

  auto first_or = io::read_ooxml(SpanOf(input));
  ASSERT_TRUE(static_cast<bool>(first_or)) << "read_ooxml: " << first_or.error().message;
  const io::OoxmlReadResult& first = first_or.value();

  // The theme part must surface as a passthrough entry on both views.
  auto find_theme = [](const std::vector<io::PassthroughPart>& parts) {
    return std::find_if(parts.begin(), parts.end(),
                        [](const io::PassthroughPart& p) { return p.path == "xl/theme/theme1.xml"; });
  };
  ASSERT_NE(find_theme(first.unknown_parts), first.unknown_parts.end()) << "theme not on read result";
  ASSERT_NE(find_theme(first.workbook.passthrough_parts()), first.workbook.passthrough_parts().end())
      << "theme not on workbook";

  const io::PassthroughPart& read_back = *find_theme(first.workbook.passthrough_parts());
  EXPECT_EQ(read_back.content_type, "application/vnd.openxmlformats-officedocument.theme+xml");
  ASSERT_FALSE(read_back.bytes.empty());

  // Re-emit and verify the part is still present and byte-stable.
  const std::vector<std::uint8_t> rewritten = SaveOrDie(first.workbook);
  const std::string workbook_rels = ExtractEntry(rewritten, "xl/_rels/workbook.xml.rels");
  EXPECT_NE(workbook_rels.find("relationships/theme"), std::string::npos) << workbook_rels;
  EXPECT_NE(workbook_rels.find("Target=\"theme/theme1.xml\""), std::string::npos) << workbook_rels;

  auto second_or = io::read_ooxml(SpanOf(rewritten));
  ASSERT_TRUE(static_cast<bool>(second_or)) << "second read_ooxml: " << second_or.error().message;
  const io::OoxmlReadResult& second = second_or.value();

  auto theme_it = find_theme(second.unknown_parts);
  ASSERT_NE(theme_it, second.unknown_parts.end()) << "theme dropped on second pass";
  EXPECT_EQ(theme_it->content_type, "application/vnd.openxmlformats-officedocument.theme+xml");
  ASSERT_EQ(theme_it->bytes.size(), read_back.bytes.size());
  EXPECT_TRUE(std::equal(theme_it->bytes.begin(), theme_it->bytes.end(), read_back.bytes.begin()))
      << "theme bytes diverged";
}

TEST(OoxmlMetadata, WellKnownPassthroughPartsGetRelationships) {
  Workbook wb = Workbook::create();
  std::vector<io::PassthroughPart> parts;
  auto add_part = [&parts](std::string path, std::string content_type, std::string_view body) {
    io::PassthroughPart part;
    part.path = std::move(path);
    part.content_type = std::move(content_type);
    part.bytes.assign(body.begin(), body.end());
    parts.push_back(std::move(part));
  };
  add_part("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml", "<core/>");
  add_part("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml", "<app/>");
  add_part("docProps/custom.xml", "application/vnd.openxmlformats-officedocument.custom-properties+xml", "<custom/>");
  add_part("xl/calcChain.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
           "<calcChain/>");
  add_part("xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
           "<sst/>");
  wb.set_passthrough_parts(std::move(parts));

  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);
  const std::string package_rels = ExtractEntry(bytes, "_rels/.rels");
  EXPECT_NE(package_rels.find("metadata/core-properties"), std::string::npos) << package_rels;
  EXPECT_NE(package_rels.find("Target=\"docProps/core.xml\""), std::string::npos) << package_rels;
  EXPECT_NE(package_rels.find("extended-properties"), std::string::npos) << package_rels;
  EXPECT_NE(package_rels.find("Target=\"docProps/app.xml\""), std::string::npos) << package_rels;
  EXPECT_NE(package_rels.find("custom-properties"), std::string::npos) << package_rels;
  EXPECT_NE(package_rels.find("Target=\"docProps/custom.xml\""), std::string::npos) << package_rels;

  const std::string workbook_rels = ExtractEntry(bytes, "xl/_rels/workbook.xml.rels");
  EXPECT_NE(workbook_rels.find("relationships/calcChain"), std::string::npos) << workbook_rels;
  EXPECT_NE(workbook_rels.find("Target=\"calcChain.xml\""), std::string::npos) << workbook_rels;
  EXPECT_NE(workbook_rels.find("relationships/sharedStrings"), std::string::npos) << workbook_rels;
  EXPECT_NE(workbook_rels.find("Target=\"sharedStrings.xml\""), std::string::npos) << workbook_rels;
}

TEST(OoxmlMetadata, PackageLevelPassthroughRelationshipsSurviveReadWrite) {
  Workbook wb = Workbook::create();
  io::PassthroughPart thumbnail;
  thumbnail.path = "docProps/thumbnail.jpeg";
  thumbnail.content_type = "image/jpeg";
  thumbnail.bytes = {0xffU, 0xd8U, 0xffU, 0xd9U};
  wb.set_passthrough_parts({thumbnail});
  wb.set_unknown_package_rels({io::UnknownRelationship{
      "rId9", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail",
      "docProps/thumbnail.jpeg", false}});

  const std::vector<std::uint8_t> first = SaveOrDie(wb);
  auto read_or = io::read_ooxml(SpanOf(first));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message;
  const Workbook& read = read_or.value().workbook;
  ASSERT_EQ(read.unknown_package_rels().size(), 1U);
  EXPECT_EQ(read.unknown_package_rels()[0].target, "docProps/thumbnail.jpeg");

  const std::vector<std::uint8_t> second = SaveOrDie(read);
  const std::string package_rels = ExtractEntry(second, "_rels/.rels");
  EXPECT_NE(package_rels.find("relationships/metadata/thumbnail"), std::string::npos) << package_rels;
  EXPECT_NE(package_rels.find("Target=\"docProps/thumbnail.jpeg\""), std::string::npos) << package_rels;
  EXPECT_EQ(ExtractEntry(second, "docProps/thumbnail.jpeg"), std::string("\xff\xd8\xff\xd9", 4));
}

TEST(OoxmlMetadata, WorksheetPrintSettingsRoundTrip) {
  const std::string printer_payload("FORMULON-PRINTER-SETTINGS\0BIN", 29);
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Default Extension=\"bin\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings\"/>\n"
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
  const std::string_view sheet_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId7\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings\" "
      "Target=\"../printerSettings/printerSettings1.bin\"/>\n"
      "</Relationships>\n";
  const std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>\n"
      "  <sheetData/>\n"
      "  <pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>\n"
      "  <pageSetup paperSize=\"9\" orientation=\"landscape\" r:id=\"rId7\"/>\n"
      "</worksheet>\n";

  const std::vector<std::uint8_t> source = BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/worksheets/_rels/sheet1.xml.rels", sheet_rels},
      {"xl/printerSettings/printerSettings1.bin", std::string_view(printer_payload.data(), printer_payload.size())},
  });

  auto first_or = io::read_ooxml(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(first_or)) << "read_ooxml: " << first_or.error().message;
  const SheetPrintSettings& print = first_or.value().workbook.sheet(0).print_settings();
  EXPECT_NE(print.sheet_pr_xml.find("pageSetUpPr"), std::string::npos) << print.sheet_pr_xml;
  EXPECT_NE(print.page_margins_xml.find("left=\"0.7\""), std::string::npos) << print.page_margins_xml;
  EXPECT_NE(print.page_setup_xml.find("orientation=\"landscape\""), std::string::npos) << print.page_setup_xml;
  EXPECT_EQ(print.printer_settings_rid, "rId7");
  EXPECT_EQ(print.printer_settings_path, "xl/printerSettings/printerSettings1.bin");

  const std::vector<std::uint8_t> rewritten = SaveOrDie(first_or.value().workbook);
  const std::string rewritten_sheet = ExtractEntry(rewritten, "xl/worksheets/sheet1.xml");
  EXPECT_NE(rewritten_sheet.find("<pageMargins"), std::string::npos) << rewritten_sheet;
  EXPECT_NE(rewritten_sheet.find("orientation=\"landscape\""), std::string::npos) << rewritten_sheet;
  EXPECT_NE(rewritten_sheet.find("r:id=\"rId7\""), std::string::npos) << rewritten_sheet;
  const std::string rewritten_rels = ExtractEntry(rewritten, "xl/worksheets/_rels/sheet1.xml.rels");
  EXPECT_NE(rewritten_rels.find("relationships/printerSettings"), std::string::npos) << rewritten_rels;
  EXPECT_NE(rewritten_rels.find("Target=\"../printerSettings/printerSettings1.bin\""), std::string::npos)
      << rewritten_rels;
  EXPECT_EQ(ExtractEntry(rewritten, "xl/printerSettings/printerSettings1.bin"), printer_payload);

  auto second_or = io::read_ooxml(SpanOf(rewritten));
  ASSERT_TRUE(static_cast<bool>(second_or)) << "second read_ooxml: " << second_or.error().message;
  EXPECT_EQ(second_or.value().workbook.sheet(0).print_settings().printer_settings_path,
            "xl/printerSettings/printerSettings1.bin");
}

TEST(OoxmlMetadata, CombinedDefinedNamesTablesAndPassthrough) {
  const std::vector<std::uint8_t> input = BuildXlsxWithThemePart();

  auto first_or = io::read_ooxml(SpanOf(input));
  ASSERT_TRUE(static_cast<bool>(first_or));
  Workbook wb = std::move(first_or.value().workbook);

  // Add defined names + a table to exercise all three round-trip
  // surfaces in one pass.
  std::vector<io::DefinedName> names;
  io::DefinedName n;
  n.name = "Combo";
  n.formula = "Sheet1!$A$1";
  names.push_back(std::move(n));
  wb.set_defined_names(std::move(names));

  io::TableMetadata table;
  table.id = 7;
  table.name = "ComboTable";
  table.display_name = "ComboTable";
  table.ref = "A1:B2";
  table.sheet_index = 0;
  table.columns.push_back(io::TableColumn{1, "X", "", "", ""});
  table.columns.push_back(io::TableColumn{2, "Y", "", "", ""});
  std::vector<io::TableMetadata> tables;
  tables.push_back(std::move(table));
  wb.set_tables(std::move(tables));

  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const io::OoxmlReadResult& result = result_or.value();
  const Workbook& dst = result.workbook;

  // Defined name preserved.
  ASSERT_EQ(dst.defined_names().size(), 1U);
  EXPECT_EQ(dst.defined_names()[0].name, "Combo");

  // Table preserved with its custom id.
  ASSERT_EQ(dst.tables().size(), 1U);
  EXPECT_EQ(dst.tables()[0].id, 7U);
  EXPECT_EQ(dst.tables()[0].name, "ComboTable");
  // The writer used the source id, so the file lives at table7.xml.
  // The reader does not surface the table part as unknown.
  for (const io::PassthroughPart& p : result.unknown_parts) {
    EXPECT_NE(p.path, "xl/tables/table7.xml");
  }

  // Passthrough part still present.
  auto theme_it = std::find_if(result.unknown_parts.begin(), result.unknown_parts.end(),
                               [](const io::PassthroughPart& p) { return p.path == "xl/theme/theme1.xml"; });
  ASSERT_NE(theme_it, result.unknown_parts.end()) << "theme dropped during combined round-trip";
}

TEST(OoxmlMetadata, TableCalculatedColumnFormulaRoundTrip) {
  // Verify that <calculatedColumnFormula> survives a full save -> read
  // cycle without truncation, and that columns without one still emit
  // the self-closing <tableColumn/> form (covered indirectly: column 0
  // reads back with an empty formula string).
  Workbook src = Workbook::create();  // single sheet "Sheet1"

  io::TableMetadata table;
  table.id = 11;
  table.name = "MyTable";
  table.display_name = "MyTable";
  table.ref = "A1:B5";
  table.sheet_index = 0;
  io::TableColumn item;
  item.id = 1;
  item.name = "Item";
  table.columns.push_back(std::move(item));
  io::TableColumn total;
  total.id = 2;
  total.name = "Total";
  total.calculated_column_formula = "SUM(MyTable[Qty])";
  table.columns.push_back(std::move(total));

  std::vector<io::TableMetadata> tables;
  tables.push_back(std::move(table));
  src.set_tables(std::move(tables));

  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  // Crack open the archive directly to assert the element is in the
  // raw XML — round-trip equality on the in-memory model is below.
  mz_zip_archive reader{};
  ASSERT_NE(mz_zip_reader_init_mem(&reader, bytes.data(), bytes.size(), 0), MZ_FALSE);
  const std::string part_name = "xl/tables/table11.xml";
  const int idx = mz_zip_reader_locate_file(&reader, part_name.c_str(), nullptr, 0);
  ASSERT_GE(idx, 0) << "expected " << part_name << " in archive";
  std::size_t extracted_size = 0;
  void* extracted = mz_zip_reader_extract_to_heap(&reader, static_cast<mz_uint>(idx), &extracted_size, 0);
  ASSERT_NE(extracted, nullptr);
  const std::string body(static_cast<const char*>(extracted), extracted_size);
  mz_free(extracted);
  mz_zip_reader_end(&reader);
  EXPECT_NE(body.find("<calculatedColumnFormula>SUM(MyTable[Qty])</calculatedColumnFormula>"), std::string::npos)
      << "calculatedColumnFormula element missing from emitted table part:\n"
      << body;

  // In-memory round-trip equality.
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& dst = result_or.value().workbook;
  ASSERT_EQ(dst.tables().size(), 1U);
  const io::TableMetadata& got = dst.tables()[0];
  ASSERT_EQ(got.columns.size(), 2U);
  EXPECT_EQ(got.columns[0].name, "Item");
  EXPECT_TRUE(got.columns[0].calculated_column_formula.empty());
  EXPECT_EQ(got.columns[1].name, "Total");
  EXPECT_EQ(got.columns[1].calculated_column_formula, "SUM(MyTable[Qty])");
}

TEST(OoxmlMetadata, PassthroughCollisionWithGeneratedPathDropsPassthrough) {
  // Construct a workbook that has a passthrough entry colliding with a
  // generated path (xl/styles.xml). The writer must emit the generated
  // styles part and silently drop the passthrough copy. We verify by
  // re-reading and checking that the styles part is the writer's
  // version, not the stale passthrough payload.
  Workbook wb = Workbook::create();
  std::vector<io::PassthroughPart> parts;
  io::PassthroughPart bogus;
  bogus.path = "xl/styles.xml";
  bogus.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
  // Distinctive bytes; if the writer mistakenly emitted these we'd see
  // the stub instead of the real styles XML on the second read.
  const std::string stub = "<?xml version=\"1.0\"?><stub/>";
  bogus.bytes.assign(stub.begin(), stub.end());
  parts.push_back(std::move(bogus));
  wb.set_passthrough_parts(std::move(parts));

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save() failed: " << save_or.error().message;

  // Extract xl/styles.xml directly from the archive and confirm it is
  // the writer-generated content (contains <styleSheet>), not the stub.
  mz_zip_archive reader{};
  ASSERT_NE(mz_zip_reader_init_mem(&reader, save_or.value().data(), save_or.value().size(), 0), MZ_FALSE);
  const int idx = mz_zip_reader_locate_file(&reader, "xl/styles.xml", nullptr, 0);
  ASSERT_GE(idx, 0);
  std::size_t extracted_size = 0;
  void* extracted = mz_zip_reader_extract_to_heap(&reader, static_cast<mz_uint>(idx), &extracted_size, 0);
  ASSERT_NE(extracted, nullptr);
  const std::string body(static_cast<const char*>(extracted), extracted_size);
  mz_free(extracted);
  mz_zip_reader_end(&reader);

  EXPECT_NE(body.find("<styleSheet"), std::string::npos)
      << "writer emitted passthrough stub instead of generated styles.xml";
  EXPECT_EQ(body.find("<stub/>"), std::string::npos) << "passthrough stub leaked into output";

  // And the resulting archive must still be readable end-to-end.
  auto result_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "round-trip failed after collision";
}

}  // namespace
}  // namespace formulon
