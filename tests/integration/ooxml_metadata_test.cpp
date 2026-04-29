// Copyright 2026 libraz. Licensed under the MIT License.
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

  // unknown_parts must NOT include the table part — Bundle 2.4 wires
  // its consumption.
  const std::vector<std::string>& parts = result.unknown_parts;
  EXPECT_EQ(std::find(parts.begin(), parts.end(), "xl/tables/table1.xml"), parts.end())
      << "table1.xml leaked into unknown_parts";
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

}  // namespace
}  // namespace formulon
