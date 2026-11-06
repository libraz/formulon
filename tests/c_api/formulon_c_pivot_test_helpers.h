#pragma once

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_count_of.h"
#include "formulon_c_out_index.h"
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
  ~BufferGuard() { fm_buffer_destroy(data); }
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
[[maybe_unused]] inline Enum RawEnumValue(std::uint32_t raw) {
  static_assert(sizeof(Enum) == sizeof(std::uint32_t), "C ABI enums are int-sized");
  Enum value{};
  std::memcpy(&value, &raw, sizeof(value));
  return value;
}

[[maybe_unused]] inline std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
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

[[maybe_unused]] inline std::string ExtractZipEntry(const std::vector<std::uint8_t>& archive_bytes,
                                                    std::string_view path) {
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

[[maybe_unused]] inline std::vector<fm_pivot_cell_t> CollectCells(fm_pivot_cells_t* handle) {
  const std::size_t n = count_of(fm_pivot_cells_get_cell_count, handle);
  std::vector<fm_pivot_cell_t> out(n);
  for (std::size_t i = 0; i < n; ++i) {
    EXPECT_EQ(fm_pivot_cells_get_cell_at(handle, i, &out[i]), 0);
  }
  return out;
}

/// The layout's `(top, left, rows, cols)` extent read through `fm_pivot_cells_get_bounds`; an empty
/// layout reads as all zeros.
[[maybe_unused]] inline fm_status_t PivotLayoutExtent(const fm_pivot_cells_t* cells, std::uint32_t* top,
                                                      std::uint32_t* left, std::uint32_t* rows, std::uint32_t* cols) {
  std::int32_t present = 0;
  fm_cell_rect_t rect{};
  const fm_status_t rc = fm_pivot_cells_get_bounds(cells, &present, &rect);
  *top = rect.first_row;
  *left = rect.first_col;
  *rows = present != 0 ? rect.last_row - rect.first_row + 1U : 0U;
  *cols = present != 0 ? rect.last_col - rect.first_col + 1U : 0U;
  return rc;
}

[[maybe_unused]] inline std::size_t CountOccurrences(const std::string& haystack, std::string_view needle) {
  std::size_t count = 0;
  for (std::size_t at = haystack.find(needle); at != std::string::npos; at = haystack.find(needle, at + 1U)) {
    ++count;
  }
  return count;
}

[[maybe_unused]] inline std::string SavedPivotXml(fm_workbook_t* wb) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  return ExtractZipEntry(std::vector<std::uint8_t>(saved.data, saved.data + saved.len),
                         "xl/pivotTables/pivotTable1.xml");
}

// True when the projected layout holds a text cell spelling `text`.
[[maybe_unused]] inline bool LayoutHasText(fm_workbook_t* wb, std::size_t pivot_index, std::string_view text) {
  PivotCellsGuard projected;
  EXPECT_EQ(fm_sheet_get_pivot_layout(wb, 0, pivot_index, &projected.handle), 0) << fm_get_last_error_message();
  for (const fm_pivot_cell_t& cell : CollectCells(projected.handle)) {
    if (cell.value.kind == FM_VAL_TEXT && cell.value.u.text != nullptr && text == cell.value.u.text) {
      return true;
    }
  }
  return false;
}

// Builds a pivot cache + table from scratch via the C ABI:
//   * 2 fields: "Region" (text) and "Amount" (numeric).
//   * 4 records: (North, 100), (South, 200), (North, 300), (South, 400).
//   * Pivot table on sheet 0, name "PT", anchor (0, 3), with Region as the
//     row field and Amount aggregated as Sum.
[[maybe_unused]] inline fm_status_t BuildScratchPivot(fm_workbook_t* wb, std::uint32_t* out_cache_id,
                                                      std::size_t* out_pivot_index) {
  std::uint32_t cache_id = 0;
  fm_status_t st = fm_workbook_add_pivot_cache(wb, 0, &cache_id);
  if (st != 0) {
    return st;
  }
  // Excel refuses a package whose cache declares no source, so a cache
  // that is going to be saved needs one. The range only has to describe
  // where the records came from; it is not required to hold data, since
  // the records themselves carry the values.
  st = fm_workbook_set_pivot_cache_worksheet_source(wb, cache_id, /*present=*/1, "A1:B5", "Sheet1", nullptr);
  if (st != 0) {
    return st;
  }
  std::size_t region_index = 99;
  st = fm_workbook_add_pivot_cache_field(wb, cache_id, "Region", &region_index);
  if (st != 0) {
    return st;
  }
  std::size_t amount_index = 99;
  st = fm_workbook_add_pivot_cache_field(wb, cache_id, "Amount", &amount_index);
  if (st != 0) {
    return st;
  }
  // Shared items for "Region": index 0 = "North", index 1 = "South".
  st = fm_workbook_add_pivot_cache_field_shared_item_text(wb, cache_id, region_index, "North");
  if (st != 0) {
    return st;
  }
  st = fm_workbook_add_pivot_cache_field_shared_item_text(wb, cache_id, region_index, "South");
  if (st != 0) {
    return st;
  }
  // 4 records.
  const struct {
    double region_index;
    double amount;
  } rows[] = {{0.0, 100.0}, {1.0, 200.0}, {0.0, 300.0}, {1.0, 400.0}};
  for (const auto& row : rows) {
    std::size_t rec_index = 99;
    st = fm_workbook_add_pivot_cache_record(wb, cache_id, &rec_index);
    if (st != 0) {
      return st;
    }
    st = fm_workbook_set_pivot_cache_record_number(wb, cache_id, rec_index, region_index, row.region_index);
    if (st != 0) {
      return st;
    }
    st = fm_workbook_set_pivot_cache_record_number(wb, cache_id, rec_index, amount_index, row.amount);
    if (st != 0) {
      return st;
    }
  }
  // Pivot table.
  std::size_t pivot_index = 99;
  st = fm_sheet_add_pivot(wb, 0, "PT", cache_id, /*anchor_row=*/0U, /*anchor_col=*/3U, &pivot_index);
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
  st = fm_sheet_add_pivot_field(wb, 0, pivot_index, &region_spec, &region_field);
  if (st != 0) {
    return st;
  }
  st = fm_sheet_add_pivot_field_item(wb, 0, pivot_index, region_field, "North", 1, ignored_index());
  if (st != 0) {
    return st;
  }
  st = fm_sheet_add_pivot_field_item(wb, 0, pivot_index, region_field, "South", 1, ignored_index());
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
  st = fm_sheet_add_pivot_field(wb, 0, pivot_index, &amount_spec, &amount_field);
  if (st != 0) {
    return st;
  }
  // Row-field order.
  const std::uint32_t row_order[] = {static_cast<std::uint32_t>(region_field)};
  st = fm_sheet_set_pivot_row_field_order(wb, 0, pivot_index, row_order, 1U);
  if (st != 0) {
    return st;
  }
  // Data field: Sum of Amount.
  fm_pivot_data_field_spec_t df_spec{};
  df_spec.name = "Sum of Amount";
  df_spec.field_index = static_cast<std::uint32_t>(amount_field);
  df_spec.aggregation = FM_PIVOT_AGG_SUM;
  df_spec.number_format = "";
  df_spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  df_spec.show_as_base_field = -1;
  df_spec.show_as_base_item = -1;
  std::size_t df_index = 99;
  st = fm_sheet_add_pivot_data_field(wb, 0, pivot_index, &df_spec, &df_index);
  if (st != 0) {
    return st;
  }
  if (out_cache_id != nullptr) {
    *out_cache_id = cache_id;
  }
  if (out_pivot_index != nullptr) {
    *out_pivot_index = pivot_index;
  }
  return 0;
}

}  // namespace
