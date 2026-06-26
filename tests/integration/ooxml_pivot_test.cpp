// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Integration test for the OOXML pivot wiring. Constructs a minimal in-memory
// `.xlsx` package whose workbook references a single pivot cache with two
// fields and three records, plus a single pivot table on Sheet2 anchored at
// D1. Drives the bytes through `read_ooxml`, then asserts:
//
//   1. The workbook's pivot-cache list has one entry with the expected
//      `cacheId`, two fields, and three records.
//   2. Sheet2's pivot-table list has one entry whose `pivot_cache_id`
//      matches the cache's id.
//   3. `pivot::evaluate` against the loaded cache + table produces the
//      expected aggregates: "North" -> 400, "South" -> 200, grand
//      total -> 600.
//
// This is the OOXML round-trip companion of the unit-level tests under
// `tests/unit/io/pivot_*_reader_test.cpp` and the GETPIVOTDATA workbook-
// fixture tests under `tests/unit/eval/getpivotdata_lazy_test.cpp`.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "miniz.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/record_access.h"
#include "sheet.h"
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

/// Materialises `parts` into a heap-allocated zip archive byte vector via
/// miniz. Mirrors the helper used by `ooxml_metadata_test.cpp`.
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
// Single-table integration test.
//
// Source rows (Sheet1!A1:B4):
//
//   Region | Amount
//   -------+-------
//   North  | 100
//   South  | 200
//   North  | 300
//
// Pivot table (Sheet2!D1) with Region on the row axis and Sum-of-Amount on
// the data axis. The pivot is wired through a single workbook-level
// pivotCacheDefinition (cacheId=0) plus a matching pivotCacheRecords part.
// ---------------------------------------------------------------------------

TEST(OoxmlPivot, PackageWithPivotCacheAndPivotTableLoads) {
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
      "  <Override PartName=\"/xl/pivotCache/pivotCacheDefinition1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml\"/>\n"
      "  <Override PartName=\"/xl/pivotCache/pivotCacheRecords1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml\"/>\n"
      "  <Override PartName=\"/xl/pivotTables/pivotTable1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml\"/>\n"
      "</Types>\n";

  const std::string_view package_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>\n"
      "</Relationships>\n";

  // workbook.xml: two sheets and a single pivotCache entry whose r:id
  // resolves through xl/_rels/workbook.xml.rels.
  const std::string_view workbook_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheets>\n"
      "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
      "    <sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/>\n"
      "  </sheets>\n"
      "  <pivotCaches>\n"
      "    <pivotCache cacheId=\"0\" r:id=\"rId3\"/>\n"
      "  </pivotCaches>\n"
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
      "  <Relationship Id=\"rId3\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\" "
      "Target=\"pivotCache/pivotCacheDefinition1.xml\"/>\n"
      "</Relationships>\n";

  // Sheet1 hosts the source range A1:B4 but we leave the cells empty
  // (the pivot reads from the cache, not from the live cells, so the
  // test does not need to populate them).
  const std::string_view sheet1_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "</worksheet>\n";

  // Sheet2 hosts the pivot table; its rels file routes a relationship id
  // to the pivot-table part.
  const std::string_view sheet2_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "</worksheet>\n";

  const std::string_view sheet2_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" "
      "Target=\"../pivotTables/pivotTable1.xml\"/>\n"
      "</Relationships>\n";

  // pivotCacheDefinition1.xml: two fields. Region is a discrete shared-
  // items field with two distinct values; Amount is a range-typed numeric
  // field with no shared items (records carry inline `<n>` entries).
  const std::string_view pivot_cache_def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
      "r:id=\"rId1\" recordCount=\"3\">\n"
      "  <cacheSource type=\"worksheet\"/>\n"
      "  <cacheFields count=\"2\">\n"
      "    <cacheField name=\"Region\">\n"
      "      <sharedItems count=\"2\">\n"
      "        <s v=\"North\"/>\n"
      "        <s v=\"South\"/>\n"
      "      </sharedItems>\n"
      "    </cacheField>\n"
      "    <cacheField name=\"Amount\">\n"
      "      <sharedItems containsNumber=\"1\" containsInteger=\"1\" minValue=\"100\" maxValue=\"300\"/>\n"
      "    </cacheField>\n"
      "  </cacheFields>\n"
      "</pivotCacheDefinition>\n";

  // pivotCacheDefinition1.xml.rels: points at the records part.
  const std::string_view pivot_cache_def_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords\" "
      "Target=\"pivotCacheRecords1.xml\"/>\n"
      "</Relationships>\n";

  // pivotCacheRecords1.xml: three records pairing Region indices with
  // inline Amount numbers.
  //   row 0: North, 100
  //   row 1: South, 200
  //   row 2: North, 300
  const std::string_view pivot_cache_records =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\">\n"
      "  <r><x v=\"0\"/><n v=\"100\"/></r>\n"
      "  <r><x v=\"1\"/><n v=\"200\"/></r>\n"
      "  <r><x v=\"0\"/><n v=\"300\"/></r>\n"
      "</pivotCacheRecords>\n";

  // pivotTable1.xml: anchored at D1 (zero-based row 0, col 3) on Sheet2,
  // with Region as the row field (axisRow) and Sum of Amount as the data
  // field. Span 5 rows x 2 cols is large enough for the rendered pivot
  // (header + 2 group rows + grand total).
  const std::string_view pivot_table_def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "name=\"PivotTable1\" cacheId=\"0\">\n"
      "  <location ref=\"D1:E5\"/>\n"
      "  <pivotFields count=\"2\">\n"
      "    <pivotField axis=\"axisRow\" name=\"Region\">\n"
      "      <items count=\"3\">\n"
      "        <item x=\"0\"/>\n"
      "        <item x=\"1\"/>\n"
      "        <item t=\"default\"/>\n"
      "      </items>\n"
      "    </pivotField>\n"
      "    <pivotField dataField=\"1\" name=\"Amount\"/>\n"
      "  </pivotFields>\n"
      "  <rowFields count=\"1\">\n"
      "    <field x=\"0\"/>\n"
      "  </rowFields>\n"
      "  <dataFields count=\"1\">\n"
      "    <dataField name=\"Sum of Amount\" fld=\"1\" subtotal=\"sum\"/>\n"
      "  </dataFields>\n"
      "</pivotTableDefinition>\n";

  const std::vector<std::uint8_t> bytes = BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet1_xml},
      {"xl/worksheets/sheet2.xml", sheet2_xml},
      {"xl/worksheets/_rels/sheet2.xml.rels", sheet2_rels},
      {"xl/pivotCache/pivotCacheDefinition1.xml", pivot_cache_def},
      {"xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", pivot_cache_def_rels},
      {"xl/pivotCache/pivotCacheRecords1.xml", pivot_cache_records},
      {"xl/pivotTables/pivotTable1.xml", pivot_table_def},
  });

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& wb = result_or.value().workbook;

  // ---------- Workbook-level: pivot cache landed --------------------------
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  const pivot::PivotCache* cache = wb.pivot_caches()[0].get();
  ASSERT_NE(cache, nullptr);
  EXPECT_EQ(cache->cache_id(), 0U);
  ASSERT_EQ(cache->fields().size(), 2U);
  EXPECT_EQ(cache->fields()[0].name, "Region");
  ASSERT_EQ(cache->fields()[0].shared_items.size(), 2U);
  EXPECT_EQ(cache->fields()[0].shared_items[0].as_text(), "North");
  EXPECT_EQ(cache->fields()[0].shared_items[1].as_text(), "South");
  EXPECT_EQ(cache->fields()[1].name, "Amount");
  EXPECT_TRUE(cache->fields()[1].shared_items.empty());
  ASSERT_EQ(cache->records().size(), 3U);
  // Records: Region is a shared field, so its cell stores the shared_items
  // index and resolves to text via `cell_value`; Amount carries an inline
  // number.
  EXPECT_EQ(pivot::cell_value(*cache, cache->records()[0], 0).as_text(), "North");
  EXPECT_DOUBLE_EQ(cache->records()[0].cells[1].as_number(), 100.0);
  EXPECT_EQ(pivot::cell_value(*cache, cache->records()[1], 0).as_text(), "South");
  EXPECT_DOUBLE_EQ(cache->records()[1].cells[1].as_number(), 200.0);
  EXPECT_EQ(pivot::cell_value(*cache, cache->records()[2], 0).as_text(), "North");
  EXPECT_DOUBLE_EQ(cache->records()[2].cells[1].as_number(), 300.0);

  // ---------- Sheet-level: pivot table landed on Sheet2 -------------------
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_TRUE(wb.sheet(0).pivot_tables().empty());
  ASSERT_EQ(wb.sheet(1).pivot_tables().size(), 1U);
  const pivot::PivotTable* table = wb.sheet(1).pivot_tables()[0].get();
  ASSERT_NE(table, nullptr);
  EXPECT_EQ(table->name(), "PivotTable1");
  EXPECT_EQ(table->pivot_cache_id(), cache->cache_id());
  EXPECT_EQ(table->anchor_row(), 0U);
  EXPECT_EQ(table->anchor_col(), 3U);
  ASSERT_EQ(table->row_field_order().size(), 1U);
  EXPECT_EQ(table->row_field_order()[0], 0U);
  ASSERT_EQ(table->data_fields().size(), 1U);
  EXPECT_EQ(table->data_fields()[0].name, "Sum of Amount");

  // ---------- End-to-end: evaluator produces the expected aggregates -----
  auto eval_or = pivot::evaluate(*table, *cache);
  ASSERT_TRUE(static_cast<bool>(eval_or)) << "pivot::evaluate: " << eval_or.error().message;
  const pivot::PivotResult& result = eval_or.value();

  // North = 100 + 300 = 400, South = 200, grand total = 600.
  ASSERT_EQ(result.rows.size(), 2U);
  // Row hierarchy preserves source-document order of shared items
  // (North, South) absent any explicit sort.
  EXPECT_EQ(result.rows[0].label, "North");
  EXPECT_EQ(result.rows[1].label, "South");
  ASSERT_EQ(result.values.size(), 2U);
  ASSERT_EQ(result.values[0].size(), 1U);
  ASSERT_EQ(result.values[0][0].size(), 1U);
  ASSERT_EQ(result.values[1].size(), 1U);
  ASSERT_EQ(result.values[1][0].size(), 1U);
  EXPECT_TRUE(result.values[0][0][0].is_number());
  EXPECT_DOUBLE_EQ(result.values[0][0][0].as_number(), 400.0);
  EXPECT_TRUE(result.values[1][0][0].is_number());
  EXPECT_DOUBLE_EQ(result.values[1][0][0].as_number(), 200.0);
  EXPECT_TRUE(result.grand_total.is_number());
  EXPECT_DOUBLE_EQ(result.grand_total.as_number(), 600.0);
}

}  // namespace
}  // namespace formulon
