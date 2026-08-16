//
// Stable C ABI PivotTable layout tests. The workbook is loaded through the
// C surface from a minimal OOXML package; assertions then use only C ABI
// entry points to count and project the pivot.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "gtest/gtest.h"
#include "miniz.h"
#include "pivot/pivot_table.h"
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

struct BufferGuard {
  std::uint8_t* data = nullptr;
  std::size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

struct PartFile {
  const char* path;
  std::string_view body;
};

// A foreign-ABI caller can put any int-sized value into an enum field, which is
// exactly what the C surface has to reject. Writing the bytes reproduces that
// without a constant conversion: an enumerator outside the enum's implied value
// range is unspecified, and GCC rejects it under -Wconversion.
template <typename Enum>
Enum RawEnumValue(std::uint32_t raw) {
  static_assert(sizeof(Enum) == sizeof(std::uint32_t), "C ABI enums are int-sized");
  Enum value{};
  std::memcpy(&value, &raw, sizeof(value));
  return value;
}

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

std::string ExtractZipEntry(const std::vector<std::uint8_t>& archive_bytes, std::string_view path) {
  mz_zip_archive reader{};
  if (mz_zip_reader_init_mem(&reader, archive_bytes.data(), archive_bytes.size(), 0) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_reader_init_mem failed";
    return {};
  }
  const int index = mz_zip_reader_locate_file(&reader, std::string(path).c_str(), nullptr, 0);
  if (index < 0) {
    ADD_FAILURE() << "entry not found: " << path;
    mz_zip_reader_end(&reader);
    return {};
  }
  std::size_t extracted_size = 0;
  void* extracted = mz_zip_reader_extract_to_heap(&reader, static_cast<mz_uint>(index), &extracted_size, 0);
  if (extracted == nullptr) {
    ADD_FAILURE() << "extract_to_heap failed for: " << path;
    mz_zip_reader_end(&reader);
    return {};
  }
  std::string body(static_cast<const char*>(extracted), extracted_size);
  mz_free(extracted);
  mz_zip_reader_end(&reader);
  return body;
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
  EXPECT_EQ(rows, 4U);
  EXPECT_EQ(cols, 2U);

  std::vector<fm_pivot_cell_t> cells;
  const std::size_t cell_count = fm_pivot_cells_count(projected.handle);
  ASSERT_GT(cell_count, 0U);
  cells.resize(cell_count);
  for (std::size_t i = 0; i < cell_count; ++i) {
    ASSERT_EQ(fm_pivot_cells_at(projected.handle, i, &cells[i]), 0);
  }

  const fm_pivot_cell_t* row_labels = FindCell(cells, 0, 3);
  ASSERT_NE(row_labels, nullptr);
  EXPECT_EQ(row_labels->kind, FM_PIVOT_CELL_HEADER);
  EXPECT_EQ(row_labels->value.kind, FM_VAL_TEXT);
  ASSERT_NE(row_labels->value.u.text, nullptr);
  EXPECT_STREQ(row_labels->value.u.text, "行ラベル");

  const fm_pivot_cell_t* north_label = FindCell(cells, 1, 3);
  ASSERT_NE(north_label, nullptr);
  EXPECT_EQ(north_label->kind, FM_PIVOT_CELL_ROW_LABEL);
  EXPECT_EQ(north_label->value.kind, FM_VAL_TEXT);
  EXPECT_STREQ(north_label->value.u.text, "North");

  const fm_pivot_cell_t* north_sum = FindCell(cells, 1, 4);
  ASSERT_NE(north_sum, nullptr);
  EXPECT_EQ(north_sum->kind, FM_PIVOT_CELL_DATA);
  EXPECT_EQ(north_sum->value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(north_sum->value.u.number, 400.0);
  ASSERT_NE(north_sum->field_name, nullptr);
  EXPECT_STREQ(north_sum->field_name, "Sum of Amount");

  const fm_pivot_cell_t* grand_total = FindCell(cells, 3, 4);
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

TEST(FormulonCApiPivot, SavedScratchPivotEmitsLocationRequiredDefaults) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  ASSERT_EQ(fm_workbook_pivot_set_anchor(wb.handle, 0, pivot_idx, 0U, 3U, 5U, 2U), 0) << fm_last_error_message();

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> package(saved.data, saved.data + saved.len);
  const std::string pivot_xml = ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml");
  ASSERT_FALSE(pivot_xml.empty());
  EXPECT_NE(pivot_xml.find("<location ref=\"D1:E5\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>"),
            std::string::npos)
      << "xml=" << pivot_xml;
}

TEST(FormulonCApiPivot, SavedPivotLocationRefCoversTheProjectedGrid) {
  // `fm_workbook_pivot_create` installs a 1x1 placeholder span and nothing
  // revises it as fields are added, so a pivot built purely through the C
  // surface used to save a `ref` describing a single cell. Excel opens such
  // a file cleanly and recognises the pivot, then terminates the moment the
  // report is refreshed -- the defect is invisible to a round trip and has
  // to be pinned on the emitted bytes.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // Deliberately no `fm_workbook_pivot_set_anchor`: this is the path a host
  // takes when it never reasons about the report's extent at all.
  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_EQ(fm_pivot_cells_bounds(projected.handle, &top, &left, &rows, &cols), 0);
  ASSERT_GT(rows, 1U) << "projection must be larger than the placeholder for this test to mean anything";
  ASSERT_GT(cols, 1U);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> package(saved.data, saved.data + saved.len);
  const std::string pivot_xml = ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml");
  ASSERT_FALSE(pivot_xml.empty());

  // Anchored at D1, projected 4 rows x 2 cols -> D1:E4.
  EXPECT_NE(pivot_xml.find("<location ref=\"D1:E4\""), std::string::npos) << "xml=" << pivot_xml;
  EXPECT_EQ(pivot_xml.find("<location ref=\"D1:D1\""), std::string::npos) << "xml=" << pivot_xml;
}

TEST(FormulonCApiPivot, AddingAFieldAfterLoadReprojectsTheLocationRef) {
  // The loaded package declares `ref="D1:E5"`, which described the report
  // as Excel left it. Adding a column field widens the grid, so re-emitting
  // the authored span would now under-size it -- the same crash, reached by
  // loading rather than by building.
  WorkbookGuard wb;
  const std::vector<std::uint8_t> package = BuildPivotWorkbookBytes();
  ASSERT_EQ(fm_workbook_load(package.data(), package.size(), &wb.handle), 0) << fm_last_error_message();

  fm_pivot_field_spec_t region_spec{};
  region_spec.source_name = "Region";
  region_spec.custom_name = "";
  region_spec.axis = FM_PIVOT_AXIS_COL;
  region_spec.subtotal_top = 0;
  region_spec.number_format = "";
  std::size_t added = 99;
  ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 1, 0, &region_spec, &added), 0) << fm_last_error_message();
  const std::uint32_t col_order[] = {static_cast<std::uint32_t>(added)};
  ASSERT_EQ(fm_workbook_pivot_set_col_field_order(wb.handle, 1, 0, col_order, 1), 0) << fm_last_error_message();

  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 1, 0, &projected.handle), 0) << fm_last_error_message();
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_EQ(fm_pivot_cells_bounds(projected.handle, &top, &left, &rows, &cols), 0);
  ASSERT_GT(cols, 2U) << "the added column field must widen the report";

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> written(saved.data, saved.data + saved.len);
  const std::string pivot_xml = ExtractZipEntry(written, "xl/pivotTables/pivotTable1.xml");
  ASSERT_FALSE(pivot_xml.empty());
  EXPECT_EQ(pivot_xml.find("<location ref=\"D1:E5\""), std::string::npos)
      << "the authored span no longer describes the report; xml=" << pivot_xml;
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

TEST(FormulonCApiPivot, PivotCacheSharedItemIndexToleratesOutOfDomainNumbers) {
  // `BuildScratchPivot` leaves `cell_is_index` empty, so "Region" (field 0)
  // reads its numeric cells as indices into its two shared items. The record
  // setter accepts any double, so the layout below is the point where an
  // out-of-domain index would be narrowed. It must resolve to a blank label
  // rather than reading past the shared items or trapping the instance.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  const double bad_indices[] = {
      std::numeric_limits<double>::quiet_NaN(),
      std::numeric_limits<double>::infinity(),
      -std::numeric_limits<double>::infinity(),
      -1.0,
      1e30,
      4.3e9,  // Past the wasm32 `size_t` range.
      2.0,    // One past the last shared item.
  };
  for (const double bad : bad_indices) {
    ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, 0, bad), 0) << fm_last_error_message();
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
    EXPECT_FALSE(CollectCells(projected.handle).empty());
  }

  // A valid index still resolves after all of that.
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, 0, 0.0), 0) << fm_last_error_message();
  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
  bool saw_north_total = false;
  for (const fm_pivot_cell_t& c : CollectCells(projected.handle)) {
    if (c.kind == FM_PIVOT_CELL_DATA && c.value.kind == FM_VAL_NUMBER && c.value.u.number == 400.0) {
      saw_north_total = true;
    }
  }
  EXPECT_TRUE(saw_north_total);
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

  // The setter accepts a raw int32_t so FFI callers can probe the complete
  // invalid domain without constructing an out-of-domain C enum.
  const std::int32_t invalid_layouts[] = {
      99,
      std::numeric_limits<std::int32_t>::min(),
      std::numeric_limits<std::int32_t>::max(),
  };
  for (const std::int32_t raw : invalid_layouts) {
    EXPECT_EQ(fm_workbook_pivot_set_layout(wb.handle, 0, pivot_idx, raw),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << raw;
  }
}

TEST(FormulonCApiPivot, ScalarEnumMutatorsRejectRawValuesWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  BufferGuard before;
  ASSERT_EQ(fm_workbook_save(wb.handle, &before.data, &before.len), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> snapshot(before.data, before.data + before.len);
  const std::int32_t invalid[] = {99, std::numeric_limits<std::int32_t>::min(),
                                  std::numeric_limits<std::int32_t>::max()};
  const fm_status_t expected = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  for (const std::int32_t raw : invalid) {
    EXPECT_EQ(fm_workbook_pivot_field_set_axis(wb.handle, 0, pivot_idx, 0, raw), expected) << raw;
    EXPECT_EQ(fm_workbook_pivot_field_add_aggregation(wb.handle, 0, pivot_idx, 1, raw), expected) << raw;
    EXPECT_EQ(fm_workbook_pivot_field_add_subtotal_fn(wb.handle, 0, pivot_idx, 0, raw), expected) << raw;
    EXPECT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, 0, raw, 0, -1, -1), expected) << raw;
    EXPECT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, 0, 3, raw, -1, -1), expected) << raw;
  }

  BufferGuard after;
  ASSERT_EQ(fm_workbook_save(wb.handle, &after.data, &after.len), 0) << fm_last_error_message();
  EXPECT_EQ(std::vector<std::uint8_t>(after.data, after.data + after.len), snapshot);
}

TEST(FormulonCApiPivot, PivotProjectionUsesWorkbookLocaleAndReportLayout) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  ASSERT_EQ(fm_workbook_pivot_set_layout(wb.handle, 0, pivot_idx, FM_PIVOT_LAYOUT_TABULAR), 0);

  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_EQ(fm_pivot_cells_bounds(projected.handle, &top, &left, &rows, &cols), 0);
  EXPECT_EQ(cols, 2U);  // Tabular exposes its two row-field columns.

  bool saw_japanese_total = false;
  const std::size_t cell_count = fm_pivot_cells_count(projected.handle);
  for (std::size_t i = 0; i < cell_count; ++i) {
    fm_pivot_cell_t cell{};
    ASSERT_EQ(fm_pivot_cells_at(projected.handle, i, &cell), 0);
    if (cell.value.kind == FM_VAL_TEXT && cell.value.u.text != nullptr &&
        std::string_view(cell.value.u.text) == "総計") {
      saw_japanese_total = true;
    }
  }
  EXPECT_TRUE(saw_japanese_total);
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

TEST(FormulonCApiPivot, InspectPivotFiltersByIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t int_filter{};
  int_filter.axis = FM_PIVOT_AXIS_VALUE;
  int_filter.field_name = "Amount";
  int_filter.type = FM_PIVOT_FILTER_VALUE_TOP_10;
  int_filter.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  int_filter.value_int = 3;
  int_filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  int_filter.data_field_index = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &int_filter), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t double_filter{};
  double_filter.axis = FM_PIVOT_AXIS_VALUE;
  double_filter.field_name = "Amount";
  double_filter.type = FM_PIVOT_FILTER_VALUE_GREATER_THAN;
  double_filter.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
  double_filter.value_double = 50.5;
  double_filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  double_filter.data_field_index = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &double_filter), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t text_filter{};
  text_filter.axis = FM_PIVOT_AXIS_ROW;
  text_filter.field_name = "Region";
  text_filter.type = FM_PIVOT_FILTER_LABEL_CONTAINS;
  text_filter.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  text_filter.value_text = "o";
  text_filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &text_filter), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t range_int_double{};
  range_int_double.axis = FM_PIVOT_AXIS_VALUE;
  range_int_double.field_name = "Amount";
  range_int_double.type = FM_PIVOT_FILTER_VALUE_BETWEEN;
  range_int_double.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  range_int_double.value_int = 100;
  range_int_double.value_high_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
  range_int_double.value_high_double = 500.25;
  range_int_double.data_field_index = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &range_int_double), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t range_double_int{};
  range_double_int.axis = FM_PIVOT_AXIS_VALUE;
  range_double_int.field_name = "Amount";
  range_double_int.type = FM_PIVOT_FILTER_VALUE_BETWEEN;
  range_double_int.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
  range_double_int.value_double = 99.75;
  range_double_int.value_high_kind = FM_PIVOT_FILTER_VALUE_INT;
  range_double_int.value_high_int = 450;
  range_double_int.data_field_index = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &range_double_int), 0) << fm_last_error_message();

  std::size_t filter_count = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &filter_count), 0);
  ASSERT_EQ(filter_count, 5U);

  // Establish a projected result before inspection; a read must not change
  // either the active-filter count or the memoised evaluation state.
  PivotCellsGuard before_projection;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &before_projection.handle), 0) << fm_last_error_message();
  const std::vector<fm_pivot_cell_t> before_cells = CollectCells(before_projection.handle);

  fm_pivot_filter_spec_t got{};
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 0, &got), 0);
  EXPECT_EQ(got.axis, FM_PIVOT_AXIS_VALUE);
  ASSERT_NE(got.field_name, nullptr);
  EXPECT_STREQ(got.field_name, "Amount");
  EXPECT_EQ(got.type, FM_PIVOT_FILTER_VALUE_TOP_10);
  EXPECT_EQ(got.value_kind, FM_PIVOT_FILTER_VALUE_INT);
  EXPECT_EQ(got.value_int, 3);
  EXPECT_DOUBLE_EQ(got.value_double, 0.0);
  EXPECT_EQ(got.value_text, nullptr);
  EXPECT_EQ(got.value_high_kind, FM_PIVOT_FILTER_VALUE_NONE);
  EXPECT_EQ(got.value_high_int, 0);
  EXPECT_DOUBLE_EQ(got.value_high_double, 0.0);
  EXPECT_EQ(got.data_field_index, 0U);

  got = {};
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 1, &got), 0);
  EXPECT_EQ(got.axis, FM_PIVOT_AXIS_VALUE);
  ASSERT_NE(got.field_name, nullptr);
  EXPECT_STREQ(got.field_name, "Amount");
  EXPECT_EQ(got.type, FM_PIVOT_FILTER_VALUE_GREATER_THAN);
  EXPECT_EQ(got.value_kind, FM_PIVOT_FILTER_VALUE_DOUBLE);
  EXPECT_EQ(got.value_int, 0);
  EXPECT_DOUBLE_EQ(got.value_double, 50.5);
  EXPECT_EQ(got.value_text, nullptr);
  EXPECT_EQ(got.value_high_kind, FM_PIVOT_FILTER_VALUE_NONE);
  EXPECT_EQ(got.value_high_int, 0);
  EXPECT_DOUBLE_EQ(got.value_high_double, 0.0);
  EXPECT_EQ(got.data_field_index, 0U);

  got = {};
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 2, &got), 0);
  EXPECT_EQ(got.axis, FM_PIVOT_AXIS_ROW);
  ASSERT_NE(got.field_name, nullptr);
  EXPECT_STREQ(got.field_name, "Region");
  EXPECT_EQ(got.type, FM_PIVOT_FILTER_LABEL_CONTAINS);
  EXPECT_EQ(got.value_kind, FM_PIVOT_FILTER_VALUE_TEXT);
  EXPECT_EQ(got.value_int, 0);
  EXPECT_DOUBLE_EQ(got.value_double, 0.0);
  ASSERT_NE(got.value_text, nullptr);
  EXPECT_STREQ(got.value_text, "o");
  EXPECT_EQ(got.value_high_kind, FM_PIVOT_FILTER_VALUE_NONE);
  EXPECT_EQ(got.value_high_int, 0);
  EXPECT_DOUBLE_EQ(got.value_high_double, 0.0);
  EXPECT_EQ(got.data_field_index, 0U);

  got = {};
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 3, &got), 0);
  EXPECT_EQ(got.axis, FM_PIVOT_AXIS_VALUE);
  ASSERT_NE(got.field_name, nullptr);
  EXPECT_STREQ(got.field_name, "Amount");
  EXPECT_EQ(got.type, FM_PIVOT_FILTER_VALUE_BETWEEN);
  EXPECT_EQ(got.value_kind, FM_PIVOT_FILTER_VALUE_INT);
  EXPECT_EQ(got.value_int, 100);
  EXPECT_DOUBLE_EQ(got.value_double, 0.0);
  EXPECT_EQ(got.value_text, nullptr);
  EXPECT_EQ(got.value_high_kind, FM_PIVOT_FILTER_VALUE_DOUBLE);
  EXPECT_EQ(got.value_high_int, 0);
  EXPECT_DOUBLE_EQ(got.value_high_double, 500.25);
  EXPECT_EQ(got.data_field_index, 0U);

  got = {};
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 4, &got), 0);
  EXPECT_EQ(got.axis, FM_PIVOT_AXIS_VALUE);
  ASSERT_NE(got.field_name, nullptr);
  EXPECT_STREQ(got.field_name, "Amount");
  EXPECT_EQ(got.type, FM_PIVOT_FILTER_VALUE_BETWEEN);
  EXPECT_EQ(got.value_kind, FM_PIVOT_FILTER_VALUE_DOUBLE);
  EXPECT_EQ(got.value_int, 0);
  EXPECT_DOUBLE_EQ(got.value_double, 99.75);
  EXPECT_EQ(got.value_text, nullptr);
  EXPECT_EQ(got.value_high_kind, FM_PIVOT_FILTER_VALUE_INT);
  EXPECT_EQ(got.value_high_int, 450);
  EXPECT_DOUBLE_EQ(got.value_high_double, 0.0);
  EXPECT_EQ(got.data_field_index, 0U);

  std::size_t unchanged_count = 0;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &unchanged_count), 0);
  EXPECT_EQ(unchanged_count, filter_count);
  PivotCellsGuard after_projection;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &after_projection.handle), 0) << fm_last_error_message();
  const std::vector<fm_pivot_cell_t> after_cells = CollectCells(after_projection.handle);
  ASSERT_EQ(after_cells.size(), before_cells.size());
  for (std::size_t i = 0; i < before_cells.size(); ++i) {
    EXPECT_EQ(after_cells[i].row, before_cells[i].row);
    EXPECT_EQ(after_cells[i].col, before_cells[i].col);
    EXPECT_EQ(after_cells[i].kind, before_cells[i].kind);
    EXPECT_EQ(after_cells[i].value.kind, before_cells[i].value.kind);
    if (before_cells[i].value.kind == FM_VAL_NUMBER) {
      EXPECT_DOUBLE_EQ(after_cells[i].value.u.number, before_cells[i].value.u.number);
    } else if (before_cells[i].value.kind == FM_VAL_TEXT) {
      ASSERT_NE(before_cells[i].value.u.text, nullptr);
      ASSERT_NE(after_cells[i].value.u.text, nullptr);
      EXPECT_STREQ(after_cells[i].value.u.text, before_cells[i].value.u.text);
    }
  }

  // Adding a filter invalidates the old model-backed views. Do not inspect
  // those pointers after the mutation; reacquire the entry instead.
  fm_pivot_filter_spec_t mutation{};
  mutation.axis = FM_PIVOT_AXIS_ROW;
  mutation.field_name = "Region";
  mutation.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  mutation.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  mutation.value_text = "N";
  mutation.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &mutation), 0) << fm_last_error_message();
  fm_pivot_filter_spec_t reacquired{};
  ASSERT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 2, &reacquired), 0);
  EXPECT_EQ(reacquired.value_kind, FM_PIVOT_FILTER_VALUE_TEXT);
  ASSERT_NE(reacquired.value_text, nullptr);
  EXPECT_STREQ(reacquired.value_text, "o");

  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  const auto make_sentinel = [] {
    fm_pivot_filter_spec_t sentinel{};
    sentinel.axis = FM_PIVOT_AXIS_PAGE;
    sentinel.field_name = "sentinel-field";
    sentinel.type = FM_PIVOT_FILTER_LABEL_DATE;
    sentinel.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
    sentinel.value_int = -123456789;
    sentinel.value_double = -9876.5;
    sentinel.value_text = "sentinel-text";
    sentinel.value_high_kind = FM_PIVOT_FILTER_VALUE_INT;
    sentinel.value_high_int = 13579;
    sentinel.value_high_double = 24680.5;
    sentinel.data_field_index = 0xDEADBEEFU;
    return sentinel;
  };
  const auto expect_unchanged = [](const fm_pivot_filter_spec_t& expected, const fm_pivot_filter_spec_t& actual) {
    EXPECT_EQ(actual.axis, expected.axis);
    EXPECT_EQ(actual.field_name, expected.field_name);
    EXPECT_EQ(actual.type, expected.type);
    EXPECT_EQ(actual.value_kind, expected.value_kind);
    EXPECT_EQ(actual.value_int, expected.value_int);
    EXPECT_DOUBLE_EQ(actual.value_double, expected.value_double);
    EXPECT_EQ(actual.value_text, expected.value_text);
    EXPECT_EQ(actual.value_high_kind, expected.value_high_kind);
    EXPECT_EQ(actual.value_high_int, expected.value_high_int);
    EXPECT_DOUBLE_EQ(actual.value_high_double, expected.value_high_double);
    EXPECT_EQ(actual.data_field_index, expected.data_field_index);
  };
  const auto expect_error_preserves_output = [&](const auto& invoke, fm_status_t expected_status) {
    fm_pivot_filter_spec_t invalid_out = make_sentinel();
    const fm_pivot_filter_spec_t before = invalid_out;
    EXPECT_EQ(invoke(&invalid_out), expected_status);
    expect_unchanged(before, invalid_out);
  };
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 6, out); },
      invalid);
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_workbook_pivot_filter_at(wb.handle, 1, pivot_idx, 0, out); },
      invalid);
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx + 1, 0, out); },
      invalid);
  expect_error_preserves_output(
      [&](fm_pivot_filter_spec_t* out) { return fm_workbook_pivot_filter_at(nullptr, 0, pivot_idx, 0, out); },
      static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(FormulonCApiPivot, GetterRejectsUnrepresentableModelEnumsWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  auto* table = wb.handle->workbook().sheet(0).mutable_pivot_tables()[pivot_idx].get();
  ASSERT_NE(table, nullptr);
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  const auto make_sentinel = [] {
    fm_pivot_filter_spec_t sentinel{};
    sentinel.axis = FM_PIVOT_AXIS_PAGE;
    sentinel.field_name = "sentinel-field";
    sentinel.type = FM_PIVOT_FILTER_LABEL_DATE;
    sentinel.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
    sentinel.value_int = -123456789;
    sentinel.value_double = -9876.5;
    sentinel.value_text = "sentinel-text";
    sentinel.value_high_kind = FM_PIVOT_FILTER_VALUE_INT;
    sentinel.value_high_int = 13579;
    sentinel.value_high_double = 24680.5;
    sentinel.data_field_index = 0xDEADBEEFU;
    return sentinel;
  };
  const auto inject_and_expect_rejected = [&](formulon::pivot::PivotAxis axis, formulon::pivot::FilterType type) {
    table->mutable_active_filters().clear();
    formulon::pivot::PivotFilter filter;
    filter.axis = axis;
    filter.field_name = "Region";
    filter.type = type;
    filter.value = std::string("N");
    table->mutable_active_filters().push_back(filter);

    fm_pivot_filter_spec_t out = make_sentinel();
    const fm_pivot_filter_spec_t before = out;
    EXPECT_EQ(fm_workbook_pivot_filter_at(wb.handle, 0, pivot_idx, 0, &out), invalid);
    EXPECT_EQ(std::memcmp(&out, &before, sizeof(out)), 0);
  };

  inject_and_expect_rejected(formulon::pivot::PivotAxis::None, formulon::pivot::FilterType::LabelContains);
  inject_and_expect_rejected(formulon::pivot::PivotAxis::Row, static_cast<formulon::pivot::FilterType>(0xFF));
}

TEST(FormulonCApiPivot, PivotFilterAddRejectsRawAxisAndTypeWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = RawEnumValue<fm_pivot_axis_t>(99);
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  spec.value_text = "N";
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 0U);

  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.type = RawEnumValue<fm_pivot_filter_type_t>(99);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPivot, PivotFilterRejectsInvalidDataFieldWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_VALUE_TOP_10;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  spec.value_int = 1;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  spec.data_field_index = 1;  // BuildScratchPivot has only slot 0.
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPivot, PivotFilterRejectsSelectorWhenNoDataFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_data_field_clear(wb.handle, 0, pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_VALUE_TOP_10;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_INT;
  spec.value_int = 1;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  spec.data_field_index = 0;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiPivot, PivotFilterRejectsUnknownFieldNameWithoutMutation) {
  // A label filter naming no pivot field would be a silent no-op inside the
  // engine, so the mutator rejects it instead of reporting success.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "NoSuchField";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  spec.value_text = "N";
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);

  std::size_t count = 99;
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 0U);

  // The data field's display name resolves through to its source field, so
  // the same call shape is accepted for it.
  spec.field_name = "Sum of Amount";
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_filter_count(wb.handle, 0, pivot_idx, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST(FormulonCApiPivot, FilterErrorUsesApiName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_filter_spec_t spec{};
  spec.axis = FM_PIVOT_AXIS_ROW;
  spec.field_name = "Region";
  spec.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  spec.value_kind = FM_PIVOT_FILTER_VALUE_NONE;
  spec.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &spec), invalid);
  EXPECT_EQ(std::string_view(fm_last_error_message()).find("fm_workbook_pivot_filter_add:"), 0U);
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

TEST(FormulonCApiPivot, MutationAndClearExportsCompleteTheirLifecycles) {
  // Exercise the C ABI setters that are not needed by the projected-layout
  // examples above.  Each mutation is followed by the corresponding clear
  // where one exists, which also verifies that the handle remains usable
  // across each invalidation boundary.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  std::uint32_t id_at_zero = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_id_at(wb.handle, 0, &id_at_zero), 0);
  EXPECT_EQ(id_at_zero, cache_id);

  const char* field_name = nullptr;
  ASSERT_EQ(fm_workbook_pivot_cache_field_name(wb.handle, cache_id, 0, &field_name), 0);
  EXPECT_STREQ(field_name, "Region");

  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_number(wb.handle, cache_id, 0, 7.0), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_bool(wb.handle, cache_id, 0, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_blank(wb.handle, cache_id, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_clear_shared_items(wb.handle, cache_id, 0), 0);
  std::size_t shared_count = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &shared_count), 0);
  EXPECT_EQ(shared_count, 0U);

  std::size_t record_idx = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_record_add(wb.handle, cache_id, &record_idx), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_text(wb.handle, cache_id, record_idx, 0, "North"), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_bool(wb.handle, cache_id, record_idx, 0, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_blank(wb.handle, cache_id, record_idx, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_error(wb.handle, cache_id, record_idx, 0, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_record_clear(wb.handle, cache_id), 0);
  std::size_t record_count = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_record_count(wb.handle, cache_id, &record_count), 0);
  EXPECT_EQ(record_count, 0U);

  ASSERT_EQ(fm_workbook_pivot_set_name(wb.handle, 0, pivot_idx, "Renamed"), 0);
  ASSERT_EQ(fm_workbook_pivot_set_grand_totals(wb.handle, 0, pivot_idx, 0, 1), 0);
  std::size_t field_count = 99;
  ASSERT_EQ(fm_workbook_pivot_field_count(wb.handle, 0, pivot_idx, &field_count), 0);
  ASSERT_EQ(field_count, 2U);

  ASSERT_EQ(fm_workbook_pivot_field_set_axis(wb.handle, 0, pivot_idx, 0, FM_PIVOT_AXIS_ROW), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_sort(wb.handle, 0, pivot_idx, 0, 0, "Amount"), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_subtotal_top(wb.handle, 0, pivot_idx, 0, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_aggregation(wb.handle, 0, pivot_idx, 1, FM_PIVOT_AGG_COUNT), 0);
  ASSERT_EQ(fm_workbook_pivot_field_clear_aggregations(wb.handle, 0, pivot_idx, 1), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_item_visible(wb.handle, 0, pivot_idx, 0, 0, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_clear_items(wb.handle, 0, pivot_idx, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_subtotal_fn(wb.handle, 0, pivot_idx, 0, FM_PIVOT_AGG_COUNT), 0);
  ASSERT_EQ(fm_workbook_pivot_field_clear_subtotal_fns(wb.handle, 0, pivot_idx, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, 0, FM_PIVOT_DATE_YEAR,
                                                   FM_PIVOT_CALENDAR_GREGORIAN, 2020, 2026),
            0);
  ASSERT_EQ(fm_workbook_pivot_field_clear_date_group(wb.handle, 0, pivot_idx, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_number_format(wb.handle, 0, pivot_idx, 1, "#,##0.00"), 0);
  const std::uint32_t col_order[] = {1U};
  ASSERT_EQ(fm_workbook_pivot_set_col_field_order(wb.handle, 0, pivot_idx, col_order, 1U), 0);

  ASSERT_EQ(fm_workbook_pivot_data_field_clear(wb.handle, 0, pivot_idx), 0);
  std::size_t data_field_count = 99;
  ASSERT_EQ(fm_workbook_pivot_data_field_count(wb.handle, 0, pivot_idx, &data_field_count), 0);
  EXPECT_EQ(data_field_count, 0U);

  fm_pivot_filter_spec_t filter{};
  filter.axis = FM_PIVOT_AXIS_ROW;
  filter.field_name = "Region";
  filter.type = FM_PIVOT_FILTER_LABEL_BEGINS_WITH;
  filter.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
  filter.value_text = "N";
  filter.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  ASSERT_EQ(fm_workbook_pivot_filter_add(wb.handle, 0, pivot_idx, &filter), 0);
  ASSERT_EQ(fm_workbook_pivot_filter_clear(wb.handle, 0, pivot_idx), 0);

  ASSERT_EQ(fm_workbook_pivot_field_clear(wb.handle, 0, pivot_idx), 0);
  ASSERT_EQ(fm_workbook_pivot_field_count(wb.handle, 0, pivot_idx, &field_count), 0);
  EXPECT_EQ(field_count, 0U);
  ASSERT_EQ(fm_workbook_pivot_cache_field_clear(wb.handle, cache_id), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_count(wb.handle, cache_id, &field_count), 0);
  EXPECT_EQ(field_count, 0U);
}

namespace {

// Region shared items are `[0] = "North"`, `[1] = blank`; the two records
// carry one of each. The blank shared item deliberately does NOT sit at index
// 0, which is what a name-addressed item silently binds to.
fm_status_t BuildBlankItemPivot(fm_workbook_t* wb, std::uint32_t* out_cache_id, std::size_t* out_pivot_index) {
  std::uint32_t cache_id = 0;
  if (fm_status_t st = fm_workbook_pivot_cache_create(wb, 0U, &cache_id); st != 0) {
    return st;
  }
  std::size_t region_idx = 99;
  std::size_t amount_idx = 99;
  if (fm_status_t st = fm_workbook_pivot_cache_field_add(wb, cache_id, "Region", &region_idx); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_field_add(wb, cache_id, "Amount", &amount_idx); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_field_add_shared_item_text(wb, cache_id, region_idx, "North"); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_field_add_shared_item_blank(wb, cache_id, region_idx); st != 0) {
    return st;
  }

  std::size_t rec_idx = 99;
  if (fm_status_t st = fm_workbook_pivot_cache_record_add(wb, cache_id, &rec_idx); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_record_set_number(wb, cache_id, rec_idx, region_idx, 0.0); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_record_set_number(wb, cache_id, rec_idx, amount_idx, 100.0); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_record_add(wb, cache_id, &rec_idx); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_record_set_blank(wb, cache_id, rec_idx, region_idx); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_record_set_number(wb, cache_id, rec_idx, amount_idx, 200.0); st != 0) {
    return st;
  }

  std::size_t pivot_idx = 99;
  if (fm_status_t st = fm_workbook_pivot_create(wb, 0, "PT", cache_id, 0U, 0U, &pivot_idx); st != 0) {
    return st;
  }
  fm_pivot_field_spec_t region_spec{};
  region_spec.source_name = "Region";
  region_spec.custom_name = "";
  region_spec.axis = FM_PIVOT_AXIS_ROW;
  region_spec.subtotal_top = 0;
  region_spec.number_format = "";
  std::size_t region_field = 99;
  if (fm_status_t st = fm_workbook_pivot_field_add(wb, 0, pivot_idx, &region_spec, &region_field); st != 0) {
    return st;
  }
  fm_pivot_field_spec_t amount_spec{};
  amount_spec.source_name = "Amount";
  amount_spec.custom_name = "";
  amount_spec.axis = FM_PIVOT_AXIS_VALUE;
  amount_spec.subtotal_top = 0;
  amount_spec.number_format = "";
  std::size_t amount_field = 99;
  if (fm_status_t st = fm_workbook_pivot_field_add(wb, 0, pivot_idx, &amount_spec, &amount_field); st != 0) {
    return st;
  }
  const std::uint32_t row_order[] = {static_cast<std::uint32_t>(region_field)};
  if (fm_status_t st = fm_workbook_pivot_set_row_field_order(wb, 0, pivot_idx, row_order, 1U); st != 0) {
    return st;
  }
  fm_pivot_data_field_spec_t df_spec{};
  df_spec.name = "Sum of Amount";
  df_spec.field_index = static_cast<std::uint32_t>(amount_field);
  df_spec.aggregation = FM_PIVOT_AGG_SUM;
  df_spec.number_format = "";
  df_spec.show_as = FM_PIVOT_SHOW_AS_NORMAL;
  df_spec.show_as_base_field = -1;
  df_spec.show_as_base_item = -1;
  std::size_t df_idx = 99;
  if (fm_status_t st = fm_workbook_pivot_data_field_add(wb, 0, pivot_idx, &df_spec, &df_idx); st != 0) {
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

// Sum of every `FM_PIVOT_CELL_DATA` cell in the projected layout.
double SumDataCells(fm_pivot_cells_t* handle) {
  double total = 0.0;
  const std::size_t n = fm_pivot_cells_count(handle);
  for (std::size_t i = 0; i < n; ++i) {
    fm_pivot_cell_t cell{};
    EXPECT_EQ(fm_pivot_cells_at(handle, i, &cell), 0);
    if (cell.kind == FM_PIVOT_CELL_DATA && cell.value.kind == FM_VAL_NUMBER) {
      total += cell.value.u.number;
    }
  }
  return total;
}

}  // namespace

TEST(FormulonCApiPivot, AddItemAtBindsTheBlankItemByCacheIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildBlankItemPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // Both source rows contribute before any manual filter exists.
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
    EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 300.0);
  }

  // The blank shared item is at cache index 1, so the index-addressed adder is
  // the only way to name it: an empty label carries no binding of its own.
  ASSERT_EQ(fm_workbook_pivot_field_add_item(wb.handle, 0, pivot_idx, 0, "North", 1), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_item_at(wb.handle, 0, pivot_idx, 0, /*cache_index=*/1U, /*visible=*/0), 0);
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
    EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 100.0);
  }
}

TEST(FormulonCApiPivot, AddItemWithEmptyNameCannotHideTheBlankRow) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildBlankItemPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // A name-addressed item defaults to cache index 0, which here holds
  // "North". The item is unlabelled, so it is matched by its binding, and that
  // binding is not blank -- it filters nothing at all. This is the gap
  // `fm_workbook_pivot_field_add_item_at` closes.
  ASSERT_EQ(fm_workbook_pivot_field_add_item(wb.handle, 0, pivot_idx, 0, "North", 1), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_item(wb.handle, 0, pivot_idx, 0, "", 0), 0);
  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
  EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 300.0);
}
