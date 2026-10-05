//
// C ABI - print pagination result handle.
//

#include "print/pagination.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <utility>
#include <vector>

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

namespace {

fm_status_t break_at(const fm_pagination_t* pagination, size_t index, uint32_t* out, bool horizontal,
                     const char* null_message, const char* range_message) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, null_message);
  }
  const std::vector<std::uint32_t>& breaks = horizontal ? pagination->result.h_breaks : pagination->result.v_breaks;
  if (index >= breaks.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, range_message);
  }
  *out = breaks[index];
  return 0;
}

}  // namespace

extern "C" size_t fm_pagination_horizontal_break_count(const fm_pagination_t* pagination) {
  return pagination == nullptr ? 0U : pagination->result.h_breaks.size();
}

extern "C" fm_status_t fm_pagination_horizontal_break_at(const fm_pagination_t* pagination, size_t index,
                                                         uint32_t* out_row) {
  return break_at(pagination, index, out_row, /*horizontal=*/true, "fm_pagination_horizontal_break_at: NULL argument",
                  "fm_pagination_horizontal_break_at: index out of range");
}

extern "C" size_t fm_pagination_vertical_break_count(const fm_pagination_t* pagination) {
  return pagination == nullptr ? 0U : pagination->result.v_breaks.size();
}

extern "C" fm_status_t fm_pagination_vertical_break_at(const fm_pagination_t* pagination, size_t index,
                                                       uint32_t* out_col) {
  return break_at(pagination, index, out_col, /*horizontal=*/false, "fm_pagination_vertical_break_at: NULL argument",
                  "fm_pagination_vertical_break_at: index out of range");
}

extern "C" fm_status_t fm_pagination_paper(const fm_pagination_t* pagination, fm_paper_info* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_pagination_paper: NULL argument");
  }
  const formulon::print::PaperInfo& paper = pagination->result.paper;
  *out = fm_paper_info{};
  out->width_pt = paper.width_pt;
  out->height_pt = paper.height_pt;
  out->landscape = paper.landscape ? 1 : 0;
  out->known = paper.known ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_pagination_margins(const fm_pagination_t* pagination, fm_margins_pt* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_pagination_margins: NULL argument");
  }
  const formulon::print::MarginsPt& margins = pagination->result.margins;
  *out = fm_margins_pt{margins.left, margins.right, margins.top, margins.bottom, margins.header, margins.footer};
  return 0;
}

extern "C" fm_status_t fm_pagination_printable(const fm_pagination_t* pagination, fm_rect_pt* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_pagination_printable: NULL argument");
  }
  const formulon::print::RectPt& rect = pagination->result.printable;
  *out = fm_rect_pt{rect.x, rect.y, rect.width, rect.height};
  return 0;
}

extern "C" fm_status_t fm_pagination_scale(const fm_pagination_t* pagination, double* out_scale) {
  clear_last_error();
  if (pagination == nullptr || out_scale == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_pagination_scale: NULL argument");
  }
  *out_scale = pagination->result.scale;
  return 0;
}

extern "C" fm_status_t fm_pagination_page_order(const fm_pagination_t* pagination, int32_t* out_order) {
  clear_last_error();
  if (pagination == nullptr || out_order == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_pagination_page_order: NULL argument");
  }
  *out_order = static_cast<int32_t>(pagination->result.page_order);
  return 0;
}

extern "C" fm_status_t fm_pagination_print_titles(const fm_pagination_t* pagination, fm_print_titles* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_pagination_print_titles: NULL argument");
  }
  const formulon::print::PrintTitleSpan& titles = pagination->result.print_titles;
  *out = fm_print_titles{};
  out->has_rows = titles.has_rows ? 1 : 0;
  out->first_row = titles.first_row;
  out->last_row = titles.last_row;
  out->has_cols = titles.has_cols ? 1 : 0;
  out->first_col = titles.first_col;
  out->last_col = titles.last_col;
  return 0;
}

extern "C" fm_status_t fm_pagination_page_at(const fm_pagination_t* pagination, size_t index, fm_page_layout* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_pagination_page_at: NULL argument");
  }
  if (index >= pagination->result.pages.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_pagination_page_at: index out of range");
  }
  const formulon::print::PageLayout& page = pagination->result.pages[index];
  *out = fm_page_layout{};
  out->area_index = page.area_index;
  out->first_row = page.first_row;
  out->last_row = page.last_row;
  out->first_col = page.first_col;
  out->last_col = page.last_col;
  out->origin_x_pt = page.origin_x_pt;
  out->origin_y_pt = page.origin_y_pt;
  out->width_pt = page.width_pt;
  out->height_pt = page.height_pt;
  return 0;
}

namespace {

fm_status_t break_is_manual(const fm_pagination_t* pagination, size_t index, int32_t* out_manual, bool horizontal,
                            const char* null_message, const char* range_message) {
  clear_last_error();
  if (pagination == nullptr || out_manual == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, null_message);
  }
  const std::vector<bool>& manual = horizontal ? pagination->result.h_break_manual : pagination->result.v_break_manual;
  if (index >= manual.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, range_message);
  }
  *out_manual = manual[index] ? 1 : 0;
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_pagination_horizontal_break_is_manual(const fm_pagination_t* pagination, size_t index,
                                                                int32_t* out_manual) {
  return break_is_manual(pagination, index, out_manual, /*horizontal=*/true,
                         "fm_pagination_horizontal_break_is_manual: NULL argument",
                         "fm_pagination_horizontal_break_is_manual: index out of range");
}

extern "C" fm_status_t fm_pagination_vertical_break_is_manual(const fm_pagination_t* pagination, size_t index,
                                                              int32_t* out_manual) {
  return break_is_manual(pagination, index, out_manual, /*horizontal=*/false,
                         "fm_pagination_vertical_break_is_manual: NULL argument",
                         "fm_pagination_vertical_break_is_manual: index out of range");
}
