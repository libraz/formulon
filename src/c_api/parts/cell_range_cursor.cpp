//
// C ABI - range cursor: paged enumeration of the cells in a rectangle
// (`fm_sheet_get_cells_in_range`) and the `fm_cell_range_t` page accessors.

#include <cstddef>
#include <cstdint>
#include <memory>

#include "c_api/formulon_c.h"
#include "c_api/parts/cell_range.h"
#include "c_api/parts/common.h"
#include "data_table.h"
#include "sheet.h"
#include "workbook.h"

using formulon::c_api::parts::array_member;
using formulon::c_api::parts::cell_range_page_size;
using formulon::c_api::parts::check_cell_cursor;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_sheet;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::ensure_calculated;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::SheetUse;

extern "C" fm_status_t fm_sheet_get_cells_in_range(const fm_workbook_t* wb, size_t sheet_index,
                                                   const fm_cell_rect_t* range, uint64_t cursor, uint32_t limit,
                                                   fm_cell_range_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  *out = nullptr;
  if (range == nullptr) {
    return null_argument("range");
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheetOrChartsheet); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(*range); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_cursor(cursor); rc != 0) {
    return rc;
  }
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  auto handle = std::unique_ptr<fm_cell_range_t>(new fm_cell_range_t{});
  handle->next_cursor =
      wb->workbook()
          .sheet(sheet_index)
          .cells_in_range(
              range->first_row, range->first_col, range->last_row, range->last_col, cursor, cell_range_page_size(limit),
              [](const formulon::Sheet::RangeCell& cell, void* ctx) {
                static_cast<fm_cell_range_t*>(ctx)->append(cell);
              },
              handle.get());
  // A data table's result cells show its `TABLE()` formula, as `fm_sheet_get_formula` does.
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  for (fm_cell_range_t::Entry& entry : handle->entries) {
    if (entry.formula != nullptr) {
      continue;
    }
    if (const formulon::DataTable* table = sheet.data_table_covering(entry.row, entry.col); table != nullptr) {
      entry.formula = handle->text.emplace_back(formulon::data_table_formula_text(*table)).c_str();
    } else if (const auto block = array_member(wb->workbook(), sheet_index, entry.row, entry.col)) {
      entry.formula = handle->text.emplace_back(block->formula).c_str();
    }
  }
  handle->spell_formulas(wb->workbook(), sheet_index);
  *out = handle.release();
  return 0;
}

extern "C" fm_status_t fm_cell_range_get_cell_count(const fm_cell_range_t* range, size_t* out_count) {
  clear_last_error();
  if (range == nullptr || out_count == nullptr) {
    return null_argument();
  }
  *out_count = range->entries.size();
  return 0;
}

extern "C" fm_status_t fm_cell_range_get_cell_at(const fm_cell_range_t* range, size_t idx, uint32_t* out_row,
                                                 uint32_t* out_col, const char** out_formula, fm_value_t* out_value) {
  clear_last_error();
  if (range == nullptr || out_row == nullptr || out_col == nullptr || out_value == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(idx, range->entries.size(), "index"); rc != 0) {
    return rc;
  }
  const fm_cell_range_t::Entry& entry = range->entries[idx];
  *out_row = entry.row;
  *out_col = entry.col;
  if (out_formula != nullptr) {
    *out_formula = entry.formula;
  }
  *out_value = entry.value;
  return 0;
}

extern "C" fm_status_t fm_cell_range_next_cursor(const fm_cell_range_t* range, uint64_t* out_cursor) {
  clear_last_error();
  if (range == nullptr || out_cursor == nullptr) {
    return null_argument();
  }
  *out_cursor = range->next_cursor;
  return 0;
}

extern "C" void fm_cell_range_destroy(fm_cell_range_t* range) {
  delete range;
}
