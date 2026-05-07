// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Stable C ABI PivotTable layout tests. The workbook is loaded through the
// C surface from a minimal OOXML package; assertions then use only C ABI
// entry points to count and project the pivot.

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "miniz.h"
#include "utils/error.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct PivotCellsGuard {
  fm_pivot_cells_t* handle = nullptr;
  ~PivotCellsGuard() { fm_pivot_cells_destroy(handle); }
  PivotCellsGuard() = default;
  PivotCellsGuard(const PivotCellsGuard&) = delete;
  PivotCellsGuard& operator=(const PivotCellsGuard&) = delete;
};

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

std::vector<std::uint8_t> BuildPivotWorkbookBytes() {
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

  const std::string_view sheet_xml =
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

  const std::string_view pivot_cache_def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" "
      "r:id=\"rId1\" recordCount=\"3\">\n"
      "  <cacheSource type=\"worksheet\"/>\n"
      "  <cacheFields count=\"2\">\n"
      "    <cacheField name=\"Region\"><sharedItems count=\"2\"><s v=\"North\"/><s "
      "v=\"South\"/></sharedItems></cacheField>\n"
      "    <cacheField name=\"Amount\"><sharedItems containsNumber=\"1\"/></cacheField>\n"
      "  </cacheFields>\n"
      "</pivotCacheDefinition>\n";

  const std::string_view pivot_cache_def_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords\" "
      "Target=\"pivotCacheRecords1.xml\"/>\n"
      "</Relationships>\n";

  const std::string_view pivot_cache_records =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotCacheRecords xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\">\n"
      "  <r><x v=\"0\"/><n v=\"100\"/></r>\n"
      "  <r><x v=\"1\"/><n v=\"200\"/></r>\n"
      "  <r><x v=\"0\"/><n v=\"300\"/></r>\n"
      "</pivotCacheRecords>\n";

  const std::string_view pivot_table_def =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "name=\"PivotTable1\" cacheId=\"0\">\n"
      "  <location ref=\"D1:E5\"/>\n"
      "  <pivotFields count=\"2\">\n"
      "    <pivotField axis=\"axisRow\" name=\"Region\"><items count=\"3\"><item x=\"0\"/><item x=\"1\"/><item "
      "t=\"default\"/></items></pivotField>\n"
      "    <pivotField dataField=\"1\" name=\"Amount\"/>\n"
      "  </pivotFields>\n"
      "  <rowFields count=\"1\"><field x=\"0\"/></rowFields>\n"
      "  <dataFields count=\"1\"><dataField name=\"Sum of Amount\" fld=\"1\" subtotal=\"sum\"/></dataFields>\n"
      "</pivotTableDefinition>\n";

  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/worksheets/sheet2.xml", sheet_xml},
      {"xl/worksheets/_rels/sheet2.xml.rels", sheet2_rels},
      {"xl/pivotCache/pivotCacheDefinition1.xml", pivot_cache_def},
      {"xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels", pivot_cache_def_rels},
      {"xl/pivotCache/pivotCacheRecords1.xml", pivot_cache_records},
      {"xl/pivotTables/pivotTable1.xml", pivot_table_def},
  });
}

const fm_pivot_cell_t* FindCell(const std::vector<fm_pivot_cell_t>& cells, std::uint32_t row, std::uint32_t col) {
  for (const fm_pivot_cell_t& cell : cells) {
    if (cell.row == row && cell.col == col) {
      return &cell;
    }
  }
  return nullptr;
}

}  // namespace

TEST(FormulonCApiPivot, CountAndLayoutLoadedPivot) {
  const std::vector<std::uint8_t> bytes = BuildPivotWorkbookBytes();
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &wb.handle), 0) << fm_last_error_message();
  ASSERT_NE(wb.handle, nullptr);

  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_pivot_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  ASSERT_EQ(fm_workbook_pivot_count(wb.handle, 1, &count), 0);
  EXPECT_EQ(count, 1U);

  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 1, 0, &projected.handle), 0) << fm_last_error_message();
  ASSERT_NE(projected.handle, nullptr);

  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_EQ(fm_pivot_cells_bounds(projected.handle, &top, &left, &rows, &cols), 0);
  EXPECT_EQ(top, 0U);
  EXPECT_EQ(left, 3U);
  EXPECT_EQ(rows, 5U);
  EXPECT_EQ(cols, 3U);

  std::vector<fm_pivot_cell_t> cells;
  const std::size_t cell_count = fm_pivot_cells_count(projected.handle);
  ASSERT_GT(cell_count, 0U);
  cells.resize(cell_count);
  for (std::size_t i = 0; i < cell_count; ++i) {
    ASSERT_EQ(fm_pivot_cells_at(projected.handle, i, &cells[i]), 0);
  }

  const fm_pivot_cell_t* region = FindCell(cells, 1, 3);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->kind, FM_PIVOT_CELL_HEADER);
  EXPECT_EQ(region->value.kind, FM_VAL_TEXT);
  ASSERT_NE(region->value.u.text, nullptr);
  EXPECT_STREQ(region->value.u.text, "Region");

  const fm_pivot_cell_t* north_label = FindCell(cells, 2, 3);
  ASSERT_NE(north_label, nullptr);
  EXPECT_EQ(north_label->kind, FM_PIVOT_CELL_ROW_LABEL);
  EXPECT_EQ(north_label->value.kind, FM_VAL_TEXT);
  EXPECT_STREQ(north_label->value.u.text, "North");

  const fm_pivot_cell_t* north_sum = FindCell(cells, 2, 4);
  ASSERT_NE(north_sum, nullptr);
  EXPECT_EQ(north_sum->kind, FM_PIVOT_CELL_DATA);
  EXPECT_EQ(north_sum->value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(north_sum->value.u.number, 400.0);
  ASSERT_NE(north_sum->field_name, nullptr);
  EXPECT_STREQ(north_sum->field_name, "Sum of Amount");

  const fm_pivot_cell_t* grand_total = FindCell(cells, 4, 5);
  ASSERT_NE(grand_total, nullptr);
  EXPECT_EQ(grand_total->kind, FM_PIVOT_CELL_GRAND_TOTAL);
  EXPECT_EQ(grand_total->value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(grand_total->value.u.number, 600.0);
}

TEST(FormulonCApiPivot, InvalidPivotArgsReturnErrors) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  std::size_t count = 0;
  EXPECT_EQ(fm_workbook_pivot_count(nullptr, 0, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_pivot_count(wb.handle, 99, &count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  fm_pivot_cells_t* cells = nullptr;
  EXPECT_EQ(fm_workbook_pivot_layout(wb.handle, 0, 0, &cells),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(cells, nullptr);

  fm_pivot_cell_t cell{};
  EXPECT_EQ(fm_pivot_cells_at(nullptr, 0, &cell),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}
