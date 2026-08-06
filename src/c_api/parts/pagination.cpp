//
// C ABI - print pagination result handle.
//

#include "print/pagination.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "utils/error.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

struct fm_pagination {
  formulon::print::PaginationResult result;
};

extern "C" fm_status_t fm_workbook_paginate(const fm_workbook_t* wb, size_t sheet_index, fm_pagination_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_paginate: NULL out");
  }
  *out = nullptr;
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_paginate"); rc != 0) {
    return rc;
  }
  auto paginated = formulon::print::paginate(wb->workbook(), static_cast<std::uint32_t>(sheet_index));
  if (!paginated) {
    return set_last_error(paginated.error());
  }
  auto handle = std::unique_ptr<fm_pagination_t>(new fm_pagination_t{});
  handle->result = std::move(paginated.value());
  *out = handle.release();
  return 0;
}

extern "C" void fm_pagination_destroy(fm_pagination_t* pagination) {
  delete pagination;
}

extern "C" uint32_t fm_pagination_page_count(const fm_pagination_t* pagination) {
  return pagination == nullptr ? 0U : pagination->result.page_count;
}

extern "C" size_t fm_pagination_print_area_count(const fm_pagination_t* pagination) {
  return pagination == nullptr ? 0U : pagination->result.print_area.size();
}

extern "C" fm_status_t fm_pagination_print_area_at(const fm_pagination_t* pagination, size_t index,
                                                   fm_print_range_t* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_pagination_print_area_at: NULL argument");
  }
  if (index >= pagination->result.print_area.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_pagination_print_area_at: index out of range");
  }
  const formulon::print::CellRange& range = pagination->result.print_area[index];
  *out = fm_print_range_t{range.first_row, range.first_col, range.last_row, range.last_col};
  return 0;
}

extern "C" size_t fm_pagination_horizontal_break_count(const fm_pagination_t* pagination) {
  return pagination == nullptr ? 0U : pagination->result.h_breaks.size();
}

extern "C" fm_status_t fm_pagination_horizontal_break_at(const fm_pagination_t* pagination, size_t index,
                                                         uint32_t* out_row) {
  clear_last_error();
  if (pagination == nullptr || out_row == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_pagination_horizontal_break_at: NULL argument");
  }
  if (index >= pagination->result.h_breaks.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_pagination_horizontal_break_at: index out of range");
  }
  *out_row = pagination->result.h_breaks[index];
  return 0;
}

extern "C" size_t fm_pagination_vertical_break_count(const fm_pagination_t* pagination) {
  return pagination == nullptr ? 0U : pagination->result.v_breaks.size();
}

extern "C" fm_status_t fm_pagination_vertical_break_at(const fm_pagination_t* pagination, size_t index,
                                                       uint32_t* out_col) {
  clear_last_error();
  if (pagination == nullptr || out_col == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_pagination_vertical_break_at: NULL argument");
  }
  if (index >= pagination->result.v_breaks.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_pagination_vertical_break_at: index out of range");
  }
  *out_col = pagination->result.v_breaks[index];
  return 0;
}
