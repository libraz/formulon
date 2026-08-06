//
// Integration test for the "real Excel" pivot shape: the pivot-table part
// links to its cache purely by index (no `name` attribute on any
// `<pivotField>`, items referenced only via `<item x="N">`). This is the
// form Excel actually writes, and the form the older in-memory pivot unit
// tests never exercised. It drives the bytes through `read_ooxml` and then
// asserts the post-load name-resolution pass wired up in the reader:
//
//   (a) each field's `source_name` is resolved from the cache field, so
//       GETPIVOTDATA can match a field by its source-column name instead
//       of returning #REF!;
//   (b) manual-filter visibility (`<item h="1">`) attaches to the correct
//       item by its resolved name, so the evaluator drops the hidden row;
//   (c) resolved item names surface as the pivot row labels; and
//   (d) an inline decimal record value survives a cache read -> write ->
//       read cycle instead of being floored into an <x> index.
//
// Companion of `ooxml_pivot_test.cpp` (which uses the name-attribute form)
// and the unit-level round-trip tests under `tests/unit/io/pivot_*`.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/pivot_cache_reader.h"
#include "io/pivot_cache_writer.h"
#include "miniz.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/record_access.h"
#include "sheet.h"
#include "utils/arena.h"
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

Value EvalWith(std::string_view src, const eval::EvalContext& ctx) {
  static thread_local Arena arena;
  arena.reset();
  parser::Parser p(src, arena);
  parser::AstNode* root = p.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  return eval::evaluate(*root, arena, eval::default_registry(), ctx);
}

// The package parts are shared by both tests below; a small builder keeps
// the byte-blobs in one place. `hide_south` toggles the manual-filter
// visibility flag on the "South" item so the filtering test can turn it on
// while the name-resolution / GETPIVOTDATA test leaves every item visible.
std::vector<std::uint8_t> BuildPackage(bool hide_south) {
  static const std::string_view content_types =
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

  static const std::string_view package_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>\n"
      "</Relationships>\n";

  static const std::string_view workbook_xml =
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

  static const std::string_view workbook_rels =
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

  static const std::string_view sheet1_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>\n";
  static const std::string_view sheet2_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>\n";
  static const std::string_view sheet2_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\" "
      "Target=\"../pivotTables/pivotTable1.xml\"/>\n"
      "</Relationships>\n";

  // Real Excel form: cache fields carry names, but the pivot part does
  // not repeat them. Amount is range-typed so records carry inline values.
  static const std::string_view pivot_cache_def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
      "r:id=\"rId1\" recordCount=\"3\">\n"
      "  <cacheSource type=\"worksheet\"/>\n"
      "  <cacheFields count=\"2\">\n"
      "    <cacheField name=\"Region\"><sharedItems count=\"2\"><s v=\"North\"/><s v=\"South\"/></sharedItems>"
      "</cacheField>\n"
      "    <cacheField name=\"Amount\"><sharedItems containsNumber=\"1\" minValue=\"100.5\" maxValue=\"300\"/>"
      "</cacheField>\n"
      "  </cacheFields>\n"
      "</pivotCacheDefinition>\n";

  static const std::string_view pivot_cache_def_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords\" "
      "Target=\"pivotCacheRecords1.xml\"/>\n"
      "</Relationships>\n";

  // Note the inline decimal 100.5 on the first record's Amount cell: it
  // must NOT be floored into an <x> index on a save.
  static const std::string_view pivot_cache_records =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\">\n"
      "  <r><x v=\"0\"/><n v=\"100.5\"/></r>\n"
      "  <r><x v=\"1\"/><n v=\"200\"/></r>\n"
      "  <r><x v=\"0\"/><n v=\"300\"/></r>\n"
      "</pivotCacheRecords>\n";

  // pivotTable part in real-Excel form: no `name` on either pivotField,
  // items reference the cache only through `x`. `<pivotTableStyleInfo>` is
  // an unmodelled tail element used to check passthrough survives.
  static const std::string_view pivot_table_visible =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "name=\"PivotTable1\" cacheId=\"0\">\n"
      "  <location ref=\"D1:E5\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>\n"
      "  <pivotFields count=\"2\">\n"
      "    <pivotField axis=\"axisRow\"><items count=\"3\"><item x=\"0\"/><item x=\"1\"/><item t=\"default\"/>"
      "</items></pivotField>\n"
      "    <pivotField dataField=\"1\"/>\n"
      "  </pivotFields>\n"
      "  <rowFields count=\"1\"><field x=\"0\"/></rowFields>\n"
      "  <dataFields count=\"1\"><dataField name=\"Sum of Amount\" fld=\"1\" subtotal=\"sum\"/></dataFields>\n"
      "  <pivotTableStyleInfo name=\"PivotStyleLight16\"/>\n"
      "</pivotTableDefinition>\n";

  static const std::string_view pivot_table_hidden =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "name=\"PivotTable1\" cacheId=\"0\">\n"
      "  <location ref=\"D1:E5\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>\n"
      "  <pivotFields count=\"2\">\n"
      "    <pivotField axis=\"axisRow\"><items count=\"3\"><item x=\"0\"/><item x=\"1\" h=\"1\"/><item t=\"default\"/>"
      "</items></pivotField>\n"
      "    <pivotField dataField=\"1\"/>\n"
      "  </pivotFields>\n"
      "  <rowFields count=\"1\"><field x=\"0\"/></rowFields>\n"
      "  <dataFields count=\"1\"><dataField name=\"Sum of Amount\" fld=\"1\" subtotal=\"sum\"/></dataFields>\n"
      "  <pivotTableStyleInfo name=\"PivotStyleLight16\"/>\n"
      "</pivotTableDefinition>\n";

  return BuildZip({
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
      {"xl/pivotTables/pivotTable1.xml", hide_south ? pivot_table_hidden : pivot_table_visible},
  });
}

// ---------------------------------------------------------------------------
// (a) source_name / item name resolution + GETPIVOTDATA + (d) inline decimal
// ---------------------------------------------------------------------------

TEST(OoxmlPivotNames, ResolvesNamesAndGetPivotDataWorks) {
  const std::vector<std::uint8_t> bytes = BuildPackage(/*hide_south=*/false);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  Workbook& wb = result_or.value().workbook;

  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  const pivot::PivotCache* cache = wb.pivot_caches()[0].get();
  ASSERT_EQ(wb.sheet(1).pivot_tables().size(), 1U);
  const pivot::PivotTable* table = wb.sheet(1).pivot_tables()[0].get();

  // (a) source_name resolved positionally from the cache, even though the
  // pivotField carried no `name` attribute.
  ASSERT_EQ(table->fields().size(), 2U);
  EXPECT_EQ(table->fields()[0].source_name, "Region");
  EXPECT_EQ(table->fields()[1].source_name, "Amount");
  EXPECT_TRUE(table->fields()[0].custom_name.empty());

  // (c) item names resolved from the cache shared_items via the item's
  // captured cache index.
  ASSERT_EQ(table->fields()[0].items.size(), 2U);
  EXPECT_EQ(table->fields()[0].items[0].name, "North");
  EXPECT_EQ(table->fields()[0].items[1].name, "South");

  // (d) the inline decimal Amount survived the load as an inline value,
  // not an index.
  ASSERT_EQ(cache->records().size(), 3U);
  ASSERT_GE(cache->records()[0].cell_is_index.size(), 2U);
  EXPECT_FALSE(cache->records()[0].cell_is_index[1]);
  EXPECT_DOUBLE_EQ(pivot::cell_value(*cache, cache->records()[0], 1).as_number(), 100.5);

  // (a) end-to-end: GETPIVOTDATA matches the field by its resolved source
  // name and returns the leaf aggregate rather than #REF!.
  eval::EvalState state;
  const eval::EvalContext ctx(wb, wb.sheet(1), state);
  const Value north = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", D1, \"Region\", \"North\")", ctx);
  ASSERT_TRUE(north.is_number()) << "GETPIVOTDATA North returned non-number";
  EXPECT_DOUBLE_EQ(north.as_number(), 400.5);  // 100.5 + 300
  const Value south = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", D1, \"Region\", \"South\")", ctx);
  ASSERT_TRUE(south.is_number());
  EXPECT_DOUBLE_EQ(south.as_number(), 200.0);
  const Value grand = EvalWith("=GETPIVOTDATA(\"Sum of Amount\", D1)", ctx);
  ASSERT_TRUE(grand.is_number());
  EXPECT_DOUBLE_EQ(grand.as_number(), 600.5);

  // (d) the inline decimal also survives a cache records write -> read
  // cycle (the floor-to-index bug would turn 100.5 into <x v="100"/>).
  const std::string rec_xml = io::write_pivot_cache_records(*cache);
  EXPECT_NE(rec_xml.find("<n v=\"100.5\"/>"), std::string::npos) << rec_xml;
  EXPECT_EQ(rec_xml.find("<x v=\"100\"/>"), std::string::npos) << rec_xml;
}

// ---------------------------------------------------------------------------
// (b) manual-filter visibility attaches to the item by resolved name
// ---------------------------------------------------------------------------

TEST(OoxmlPivotNames, HiddenItemFiltersByResolvedName) {
  const std::vector<std::uint8_t> bytes = BuildPackage(/*hide_south=*/true);
  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& wb = result_or.value().workbook;
  const pivot::PivotCache* cache = wb.pivot_caches()[0].get();
  const pivot::PivotTable* table = wb.sheet(1).pivot_tables()[0].get();

  // The hidden flag landed on the item whose resolved name is "South".
  ASSERT_EQ(table->fields()[0].items.size(), 2U);
  ASSERT_EQ(table->fields()[0].items[1].name, "South");
  EXPECT_TRUE(table->fields()[0].items[0].visible);
  EXPECT_FALSE(table->fields()[0].items[1].visible);

  // The evaluator applies that visibility by name: only the North row
  // survives, and the grand total drops to North's contribution.
  auto eval_or = pivot::evaluate(*table, *cache);
  ASSERT_TRUE(static_cast<bool>(eval_or)) << "pivot::evaluate: " << eval_or.error().message;
  const pivot::PivotResult& result = eval_or.value();
  ASSERT_EQ(result.rows.size(), 1U);
  EXPECT_EQ(result.rows[0].label, "North");
}

}  // namespace
}  // namespace formulon
