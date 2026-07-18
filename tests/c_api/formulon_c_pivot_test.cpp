// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Stable C ABI PivotTable layout tests. The workbook is loaded through the
// C surface from a minimal OOXML package; assertions then use only C ABI
// entry points to count and project the pivot.

#include <cstddef>
#include <cstdint>
#include <cstring>
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

// ----------------------------------------------------------------------------
// Pivot mutation surface
// ----------------------------------------------------------------------------

namespace {

// Builds a pivot cache + table from scratch via the C ABI:
//   * 2 fields: "Region" (text) and "Amount" (numeric).
//   * 4 records: (North, 100), (South, 200), (North, 300), (South, 400).
//   * Pivot table on sheet 0, name "PT", anchor (0, 3), with Region as the
//     row field and Amount aggregated as Sum.
fm_status_t BuildScratchPivot(fm_workbook_t* wb, std::uint32_t* out_cache_id, std::size_t* out_pivot_index) {
  std::uint32_t cache_id = 0;
  fm_status_t st = fm_workbook_pivot_cache_create(wb, 0, &cache_id);
  if (st != 0) {
    return st;
  }
  std::size_t region_idx = 99;
  st = fm_workbook_pivot_cache_field_add(wb, cache_id, "Region", &region_idx);
  if (st != 0) {
    return st;
  }
  std::size_t amount_idx = 99;
  st = fm_workbook_pivot_cache_field_add(wb, cache_id, "Amount", &amount_idx);
  if (st != 0) {
    return st;
  }
  // Shared items for "Region": index 0 = "North", index 1 = "South".
  st = fm_workbook_pivot_cache_field_add_shared_item_text(wb, cache_id, region_idx, "North");
  if (st != 0) {
    return st;
  }
  st = fm_workbook_pivot_cache_field_add_shared_item_text(wb, cache_id, region_idx, "South");
  if (st != 0) {
    return st;
  }
  // 4 records.
  const struct {
    double region_index;
    double amount;
  } rows[] = {{0.0, 100.0}, {1.0, 200.0}, {0.0, 300.0}, {1.0, 400.0}};
  for (const auto& row : rows) {
    std::size_t rec_idx = 99;
    st = fm_workbook_pivot_cache_record_add(wb, cache_id, &rec_idx);
    if (st != 0) {
      return st;
    }
    st = fm_workbook_pivot_cache_record_set_number(wb, cache_id, rec_idx, region_idx, row.region_index);
    if (st != 0) {
      return st;
    }
    st = fm_workbook_pivot_cache_record_set_number(wb, cache_id, rec_idx, amount_idx, row.amount);
    if (st != 0) {
      return st;
    }
  }
  // Pivot table.
  std::size_t pivot_idx = 99;
  st = fm_workbook_pivot_create(wb, 0, "PT", cache_id, /*anchor_row=*/0U, /*anchor_col=*/3U, &pivot_idx);
  if (st != 0) {
    return st;
  }
  // Region (row) + items 0 and 1, plus the OOXML "default" subtotal item.
  fm_pivot_field_spec_t region_spec{};
  region_spec.source_name = "Region";
  region_spec.custom_name = "";
  region_spec.axis = FM_PIVOT_AXIS_ROW;
  region_spec.subtotal_top = 0;
  region_spec.number_format = "";
  std::size_t region_field = 99;
  st = fm_workbook_pivot_field_add(wb, 0, pivot_idx, &region_spec, &region_field);
  if (st != 0) {
    return st;
  }
  st = fm_workbook_pivot_field_add_item(wb, 0, pivot_idx, region_field, "North", 1);
  if (st != 0) {
    return st;
  }
  st = fm_workbook_pivot_field_add_item(wb, 0, pivot_idx, region_field, "South", 1);
  if (st != 0) {
    return st;
  }
  // Amount (value-axis source field).
  fm_pivot_field_spec_t amount_spec{};
  amount_spec.source_name = "Amount";
  amount_spec.custom_name = "";
  amount_spec.axis = FM_PIVOT_AXIS_VALUE;
  amount_spec.subtotal_top = 0;
  amount_spec.number_format = "";
  std::size_t amount_field = 99;
  st = fm_workbook_pivot_field_add(wb, 0, pivot_idx, &amount_spec, &amount_field);
  if (st != 0) {
    return st;
  }
  // Row-field order.
  const std::uint32_t row_order[] = {static_cast<std::uint32_t>(region_field)};
  st = fm_workbook_pivot_set_row_field_order(wb, 0, pivot_idx, row_order, 1U);
  if (st != 0) {
    return st;
  }
  // Data field: Sum of Amount.
  fm_pivot_data_field_spec_t df_spec{};
  df_spec.name = "Sum of Amount";
  df_spec.field_index = static_cast<std::uint32_t>(amount_field);
  df_spec.aggregation = FM_PIVOT_AGG_SUM;
  df_spec.number_format = "";
  df_spec.show_as = FM_PIVOT_SHOW_AS_NORMAL;
  df_spec.show_as_base_field = -1;
  df_spec.show_as_base_item = -1;
  std::size_t df_idx = 99;
  st = fm_workbook_pivot_data_field_add(wb, 0, pivot_idx, &df_spec, &df_idx);
  if (st != 0) {
    return st;
  }
  if (out_cache_id != nullptr) {
    *out_cache_id = cache_id;
  }
  if (out_pivot_index != nullptr) {
    *out_pivot_index = pivot_idx;
  }
  return 0;
}

std::vector<fm_pivot_cell_t> CollectCells(fm_pivot_cells_t* handle) {
  const std::size_t n = fm_pivot_cells_count(handle);
  std::vector<fm_pivot_cell_t> out(n);
  for (std::size_t i = 0; i < n; ++i) {
    EXPECT_EQ(fm_pivot_cells_at(handle, i, &out[i]), 0);
  }
  return out;
}

}  // namespace

TEST(FormulonCApiPivot, CreatePivotFromScratch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  std::size_t cache_count = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_count(wb.handle, &cache_count), 0);
  EXPECT_EQ(cache_count, 1U);

  std::size_t pivot_count = 0;
  ASSERT_EQ(fm_workbook_pivot_count(wb.handle, 0, &pivot_count), 0);
  EXPECT_EQ(pivot_count, 1U);

  std::size_t records = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_record_count(wb.handle, cache_id, &records), 0);
  EXPECT_EQ(records, 4U);

  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();

  const std::vector<fm_pivot_cell_t> cells = CollectCells(projected.handle);
  ASSERT_FALSE(cells.empty());

  // North (regions 0, 2) sum = 400; South (regions 1, 3) sum = 600.
  // Grand total = 1000.
  bool saw_north = false;
  bool saw_south = false;
  bool saw_grand = false;
  for (const fm_pivot_cell_t& c : cells) {
    if (c.kind == FM_PIVOT_CELL_DATA && c.value.kind == FM_VAL_NUMBER) {
      if (c.value.u.number == 400.0) {
        saw_north = true;
      } else if (c.value.u.number == 600.0) {
        saw_south = true;
      }
    }
    if (c.kind == FM_PIVOT_CELL_GRAND_TOTAL && c.value.kind == FM_VAL_NUMBER && c.value.u.number == 1000.0) {
      saw_grand = true;
    }
  }
  EXPECT_TRUE(saw_north);
  EXPECT_TRUE(saw_south);
  EXPECT_TRUE(saw_grand);
}

TEST(FormulonCApiPivot, PivotCacheMutationInvalidatesMemoisedLayout) {
  // fm_workbook_pivot_layout memoises the pivot's evaluated result. A cache
  // mutation must invalidate that memo so a re-layout reflects the change
  // instead of returning the stale projection.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  auto has_data_value = [](fm_pivot_cells_t* handle, double want) -> bool {
    for (const fm_pivot_cell_t& c : CollectCells(handle)) {
      if (c.kind == FM_PIVOT_CELL_DATA && c.value.kind == FM_VAL_NUMBER && c.value.u.number == want) {
        return true;
      }
    }
    return false;
  };

  // Baseline: North (records 0, 2 = 100 + 300) sums to 400.
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
    EXPECT_TRUE(has_data_value(projected.handle, 400.0));
  }

  // Bump record 0's amount (field index 1) from 100 to 1100 (+1000).
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, 1, 1100.0), 0) << fm_last_error_message();

  // Re-layout must reflect the mutated cache: North is now 1100 + 300 = 1400,
  // and the stale 400 must be gone.
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
    EXPECT_TRUE(has_data_value(projected.handle, 1400.0)) << "layout returned a stale memoised projection";
    EXPECT_FALSE(has_data_value(projected.handle, 400.0)) << "stale North value survived the cache mutation";
  }
}

TEST(FormulonCApiPivot, PivotCacheSharedItemsAcceptErrorValues) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &count), 0);
  EXPECT_EQ(count, 2U);

  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_error(wb.handle, cache_id, 0,
                                                                1),  // ErrorCode::Div0
            0)
      << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &count), 0);
  EXPECT_EQ(count, 3U);
}

TEST(FormulonCApiPivot, PivotCacheWorksheetSourceRoundTripsThroughApi) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  int32_t present = -1;
  const char* ref = nullptr;
  const char* sheet = nullptr;
  const char* name = nullptr;
  ASSERT_EQ(fm_workbook_pivot_cache_get_worksheet_source(wb.handle, cache_id, &present, &ref, &sheet, &name), 0);
  EXPECT_EQ(present, 0);
  EXPECT_STREQ(ref, "");
  EXPECT_STREQ(sheet, "");
  EXPECT_STREQ(name, "");

  ASSERT_EQ(fm_workbook_pivot_cache_set_worksheet_source(wb.handle, cache_id, 1, "$A$1:$C$5", "Data", nullptr), 0)
      << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_cache_get_worksheet_source(wb.handle, cache_id, &present, &ref, &sheet, &name), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(ref, "$A$1:$C$5");
  EXPECT_STREQ(sheet, "Data");
  EXPECT_STREQ(name, "");

  ASSERT_EQ(fm_workbook_pivot_cache_set_worksheet_source(wb.handle, cache_id, 0, "ignored", "ignored", "ignored"), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_get_worksheet_source(wb.handle, cache_id, &present, &ref, &sheet, &name), 0);
  EXPECT_EQ(present, 0);
  EXPECT_STREQ(ref, "");
  EXPECT_STREQ(sheet, "");
  EXPECT_STREQ(name, "");
}

TEST(FormulonCApiPivot, PivotReportLayoutRoundTripsThroughApi) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_layout_t layout = FM_PIVOT_LAYOUT_TABULAR;
  ASSERT_EQ(fm_workbook_pivot_get_layout(wb.handle, 0, pivot_idx, &layout), 0);
  EXPECT_EQ(layout, FM_PIVOT_LAYOUT_COMPACT);

  ASSERT_EQ(fm_workbook_pivot_set_layout(wb.handle, 0, pivot_idx, FM_PIVOT_LAYOUT_TABULAR), 0);
  ASSERT_EQ(fm_workbook_pivot_get_layout(wb.handle, 0, pivot_idx, &layout), 0);
  EXPECT_EQ(layout, FM_PIVOT_LAYOUT_TABULAR);

  ASSERT_EQ(fm_workbook_pivot_set_layout(wb.handle, 0, pivot_idx, FM_PIVOT_LAYOUT_OUTLINE), 0);
  ASSERT_EQ(fm_workbook_pivot_get_layout(wb.handle, 0, pivot_idx, &layout), 0);
  EXPECT_EQ(layout, FM_PIVOT_LAYOUT_OUTLINE);

  // Domain is {0, 1, 2}; build an out-of-range value via memcpy because a
  // direct `static_cast<fm_pivot_layout_t>(99)` is unspecified per the
  // standard (99 sits outside the enum's representable range, which GCC's
  // -Wconversion correctly flags). The C ABI accepts the raw byte value
  // regardless and routes it through the validation switch.
  fm_pivot_layout_t bad_layout{};
  const int raw = 99;
  std::memcpy(&bad_layout, &raw, sizeof(bad_layout));
  EXPECT_EQ(fm_workbook_pivot_set_layout(wb.handle, 0, pivot_idx, bad_layout),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiPivot, PivotCacheErrorSettersRejectInvalidErrorCodes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  EXPECT_EQ(fm_workbook_pivot_cache_field_add_shared_item_error(wb.handle, cache_id, 0, -1),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_pivot_cache_field_add_shared_item_error(wb.handle, cache_id, 0, 999),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_error(wb.handle, cache_id, 0, 1, -1),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_error(wb.handle, cache_id, 0, 1, 999),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiPivot, PivotCacheRecordSettersRejectOutOfRangeFieldIndex) {
  // BuildScratchPivot declares two cache fields (Region=0, Amount=1), so
  // field_idx 2 is the first out-of-range value. A large / wrapping field_idx
  // must be rejected before `grow_record_cells` resizes `field_idx + 1`
  // cells — otherwise SIZE_MAX wraps to 0 and the subsequent cell write lands
  // out of bounds.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  const auto kInvalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  const std::size_t kFieldCount = 2;
  const std::size_t kMax = static_cast<std::size_t>(-1);

  // In-range write still succeeds (guards against an over-tight bound).
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, kFieldCount - 1, 1.0), 0);

  // field_idx == field_count and beyond are rejected on every setter.
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, kFieldCount, 1.0), kInvalid);
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, kMax, 1.0), kInvalid);
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_text(wb.handle, cache_id, 0, kMax, "x"), kInvalid);
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_bool(wb.handle, cache_id, 0, kMax, 1), kInvalid);
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_blank(wb.handle, cache_id, 0, kMax), kInvalid);
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_error(wb.handle, cache_id, 0, kMax, 0), kInvalid);
}

TEST(FormulonCApiPivot, MutateExistingPivotFilter) {
  // Build a scratch pivot — the OOXML reader populates only `custom_name`
  // from `<pivotField name="...">`, while the filter resolver matches on
  // `source_name`, so a filter applied to a loaded workbook would silently
  // no-op. Building from scratch lets us exercise the same surface against
  // a pivot whose `source_name` is set the way the resolver expects.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // Baseline projection contains "South" because no filter is active.
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0);
    bool saw_south = false;
    for (const fm_pivot_cell_t& c : CollectCells(projected.handle)) {
      if (c.kind == FM_PIVOT_CELL_ROW_LABEL && c.value.kind == FM_VAL_TEXT && c.value.u.text != nullptr &&
          std::string_view(c.value.u.text) == "South") {
        saw_south = true;
      }
    }
    EXPECT_TRUE(saw_south);
  }

  // Add a label filter that only retains "North".
  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  spec.value_text = "N";
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), 0) << fm_last_error_message();

  std::size_t filter_count = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &filter_count), 0);
  EXPECT_EQ(filter_count, 1U);

  // Re-project — "South" must be dropped.
  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0);
  bool saw_south = false;
  bool saw_north = false;
  for (const fm_pivot_cell_t& c : CollectCells(projected.handle)) {
    if (c.kind == FM_PIVOT_CELL_ROW_LABEL && c.value.kind == FM_VAL_TEXT && c.value.u.text != nullptr) {
      if (std::string_view(c.value.u.text) == "South") {
        saw_south = true;
      }
      if (std::string_view(c.value.u.text) == "North") {
        saw_north = true;
      }
    }
  }
  EXPECT_TRUE(saw_north);
  EXPECT_FALSE(saw_south);

  // Verify remove_at clears the filter.
  ASSERT_EQ(fm_workbook_pivot_filter_remove_at(wb.handle, 0, pivot_idx, 0), 0);
  std::size_t after_count = 99;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &after_count), 0);
  EXPECT_EQ(after_count, 0U);
}

TEST(FormulonCApiPivot, MutateShowValuesAs) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // Re-set the data field to PercentOfRow.
  std::size_t df_count = 0;
  ASSERT_EQ(fm_workbook_pivot_data_field_count(wb.handle, 0, pivot_idx, &df_count), 0);
  ASSERT_EQ(df_count, 1U);

  fm_pivot_data_field_spec_t df_spec{};
  df_spec.name = "Sum of Amount";
  df_spec.field_index = 1U;  // Amount is the second pivot field.
  df_spec.aggregation = FM_PIVOT_AGG_SUM;
  df_spec.number_format = "";
  df_spec.show_as = FM_PIVOT_SHOW_AS_PERCENT_OF_ROW;
  df_spec.show_as_base_field = -1;
  df_spec.show_as_base_item = -1;
  ASSERT_EQ(fm_workbook_pivot_data_field_set(wb.handle, 0, pivot_idx, 0, &df_spec), 0) << fm_last_error_message();

  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();

  // Each FM_PIVOT_CELL_DATA cell should equal 1.0 (single data column ->
  // each row's percent-of-row is 100% of itself).
  std::size_t data_count = 0;
  for (const fm_pivot_cell_t& c : CollectCells(projected.handle)) {
    if (c.kind == FM_PIVOT_CELL_DATA && c.value.kind == FM_VAL_NUMBER) {
      EXPECT_DOUBLE_EQ(c.value.u.number, 1.0);
      ++data_count;
    }
  }
  EXPECT_GT(data_count, 0U);
}

TEST(FormulonCApiPivot, RemovePivotAndCache) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  std::size_t before_pivots = 0;
  std::size_t before_caches = 0;
  ASSERT_EQ(fm_workbook_pivot_count(wb.handle, 0, &before_pivots), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_count(wb.handle, &before_caches), 0);
  EXPECT_EQ(before_pivots, 1U);
  EXPECT_EQ(before_caches, 1U);

  // Remove the pivot first, then the cache.
  ASSERT_EQ(fm_workbook_pivot_remove(wb.handle, 0, pivot_idx), 0);
  std::size_t after_remove_pivots = 99;
  ASSERT_EQ(fm_workbook_pivot_count(wb.handle, 0, &after_remove_pivots), 0);
  EXPECT_EQ(after_remove_pivots, 0U);

  ASSERT_EQ(fm_workbook_pivot_cache_remove(wb.handle, cache_id), 0);
  std::size_t after_caches = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_count(wb.handle, &after_caches), 0);
  EXPECT_EQ(after_caches, 0U);
}

TEST(FormulonCApiPivot, CacheRemoveBlockedByPivot) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  EXPECT_EQ(fm_workbook_pivot_cache_remove(wb.handle, cache_id),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // Removing a non-existent cache id must also surface kInvalidArgument.
  EXPECT_EQ(fm_workbook_pivot_cache_remove(wb.handle, 9999U),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiPivot, NullPointerArgumentsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_create(wb.handle, 0, &cache_id), 0);
  std::size_t field_idx = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_field_add(wb.handle, cache_id, "F", &field_idx), 0);

  std::size_t out_count = 0;
  EXPECT_EQ(fm_workbook_pivot_cache_count(nullptr, &out_count),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_pivot_cache_count(wb.handle, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  std::size_t fc = 0;
  EXPECT_EQ(fm_workbook_pivot_cache_field_count(nullptr, cache_id, &fc),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  EXPECT_EQ(fm_workbook_pivot_cache_field_add(wb.handle, cache_id, nullptr, &field_idx),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  EXPECT_EQ(fm_workbook_pivot_cache_field_add_shared_item_text(wb.handle, cache_id, field_idx, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  std::size_t pivot_idx = 99;
  EXPECT_EQ(fm_workbook_pivot_create(wb.handle, 0, nullptr, cache_id, 0, 0, &pivot_idx),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  // spec arguments
  EXPECT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, 0, nullptr, &pivot_idx),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  fm_pivot_filter_spec_t bad_filter{};
  bad_filter.field_name = nullptr;
  bad_filter.type = FM_PIVOT_FILTER_LABEL_CONTAINS;
  bad_filter.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  bad_filter.value_text = "X";
  bad_filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, 0, &bad_filter),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiPivot, OutOfGridAnchorRejected) {
  // Excel grid ceilings; kept as literals so this test needs no core
  // header. Mirrors Sheet::kMaxRows / kMaxCols.
  constexpr std::uint32_t kMaxRows = 1'048'576U;
  constexpr std::uint32_t kMaxCols = 16'384U;

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_create(wb.handle, 0, &cache_id), 0);
  std::size_t field_idx = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_field_add(wb.handle, cache_id, "F", &field_idx), 0);

  // create() with an anchor past the grid must be rejected, not stored.
  std::size_t pivot_idx = 99;
  EXPECT_EQ(fm_workbook_pivot_create(wb.handle, 0, "PT", cache_id, kMaxRows, 0U, &pivot_idx),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_pivot_create(wb.handle, 0, "PT", cache_id, 0U, kMaxCols, &pivot_idx),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));

  // A valid create succeeds, then set_anchor must reject a span whose far
  // corner leaves the grid (the wrap the audit flagged) and a zero span.
  ASSERT_EQ(fm_workbook_pivot_create(wb.handle, 0, "PT", cache_id, 0U, 0U, &pivot_idx), 0);
  EXPECT_EQ(
      fm_workbook_pivot_set_anchor(wb.handle, 0, pivot_idx, kMaxRows - 1U, 0U, /*span_rows=*/2U, /*span_cols=*/1U),
      static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_pivot_set_anchor(wb.handle, 0, pivot_idx, 0U, 0U, /*span_rows=*/0U, /*span_cols=*/1U),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  // An in-grid span still succeeds.
  EXPECT_EQ(fm_workbook_pivot_set_anchor(wb.handle, 0, pivot_idx, 0U, 0U, /*span_rows=*/3U, /*span_cols=*/2U), 0);
}
