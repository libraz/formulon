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

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

struct fm_pagination {
  formulon::print::PaginationResult result;
};

extern "C" fm_status_t fm_workbook_get_pagination(const fm_workbook_t* wb, size_t sheet_index, fm_pagination_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_get_pagination: NULL out");
  }
  *out = nullptr;
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_get_pagination"); rc != 0) {
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

extern "C" uint32_t fm_pagination_get_page_count(const fm_pagination_t* pagination) {
  clear_last_error();
  return pagination == nullptr ? 0U : pagination->result.page_count;
}

extern "C" size_t fm_pagination_get_print_area_count(const fm_pagination_t* pagination) {
  clear_last_error();
  return pagination == nullptr ? 0U : pagination->result.print_area.size();
}

extern "C" fm_status_t fm_pagination_get_print_area_at(const fm_pagination_t* pagination, size_t index,
                                                       fm_print_range_t* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return null_argument("fm_pagination_get_print_area_at");
  }
  if (auto rc = check_index(index, pagination->result.print_area.size(), "fm_pagination_get_print_area_at", "index");
      rc != 0) {
    return rc;
  }
  const formulon::print::CellRange& range = pagination->result.print_area[index];
  *out = fm_print_range_t{range.first_row, range.first_col, range.last_row, range.last_col};
  return 0;
}

namespace {

// Shared body of the break getters. `manual` reads the manual-flag list
// instead of the position list; the value is returned through `value`.
fm_status_t break_entry(const fm_pagination_t* pagination, size_t index, const void* out, bool horizontal, bool manual,
                        const char* fn, uint32_t* value) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return null_argument(fn);
  }
  const auto& result = pagination->result;
  const std::size_t size = manual ? (horizontal ? result.h_break_manual.size() : result.v_break_manual.size())
                                  : (horizontal ? result.h_breaks.size() : result.v_breaks.size());
  if (auto rc = check_index(index, size, fn, "index"); rc != 0) {
    return rc;
  }
  *value = manual ? static_cast<uint32_t>((horizontal ? result.h_break_manual : result.v_break_manual)[index])
                  : (horizontal ? result.h_breaks : result.v_breaks)[index];
  return 0;
}

}  // namespace

extern "C" size_t fm_pagination_get_row_break_count(const fm_pagination_t* pagination) {
  clear_last_error();
  return pagination == nullptr ? 0U : pagination->result.h_breaks.size();
}

extern "C" fm_status_t fm_pagination_get_row_break_at(const fm_pagination_t* pagination, size_t index,
                                                      uint32_t* out_row) {
  uint32_t value = 0;
  if (auto rc = break_entry(pagination, index, out_row, true, false, "fm_pagination_get_row_break_at", &value);
      rc != 0) {
    return rc;
  }
  *out_row = value;
  return 0;
}

extern "C" size_t fm_pagination_get_col_break_count(const fm_pagination_t* pagination) {
  clear_last_error();
  return pagination == nullptr ? 0U : pagination->result.v_breaks.size();
}

extern "C" fm_status_t fm_pagination_get_col_break_at(const fm_pagination_t* pagination, size_t index,
                                                      uint32_t* out_col) {
  uint32_t value = 0;
  if (auto rc = break_entry(pagination, index, out_col, false, false, "fm_pagination_get_col_break_at", &value);
      rc != 0) {
    return rc;
  }
  *out_col = value;
  return 0;
}

extern "C" fm_status_t fm_pagination_get_paper(const fm_pagination_t* pagination, fm_paper_info_t* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return null_argument("fm_pagination_get_paper");
  }
  const formulon::print::PaperInfo& paper = pagination->result.paper;
  *out = fm_paper_info_t{};
  out->width_pt = paper.width_pt;
  out->height_pt = paper.height_pt;
  out->landscape = paper.landscape ? 1 : 0;
  out->known = paper.known ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_pagination_get_margins(const fm_pagination_t* pagination, fm_margins_pt_t* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return null_argument("fm_pagination_get_margins");
  }
  const formulon::print::MarginsPt& margins = pagination->result.margins;
  *out = fm_margins_pt_t{margins.left, margins.right, margins.top, margins.bottom, margins.header, margins.footer};
  return 0;
}

extern "C" fm_status_t fm_pagination_get_printable(const fm_pagination_t* pagination, fm_rect_pt_t* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return null_argument("fm_pagination_get_printable");
  }
  const formulon::print::RectPt& rect = pagination->result.printable;
  *out = fm_rect_pt_t{rect.x, rect.y, rect.width, rect.height};
  return 0;
}

extern "C" fm_status_t fm_pagination_get_scale(const fm_pagination_t* pagination, double* out_scale) {
  clear_last_error();
  if (pagination == nullptr || out_scale == nullptr) {
    return null_argument("fm_pagination_get_scale");
  }
  *out_scale = pagination->result.scale;
  return 0;
}

extern "C" fm_status_t fm_pagination_get_page_order(const fm_pagination_t* pagination, int32_t* out_order) {
  clear_last_error();
  if (pagination == nullptr || out_order == nullptr) {
    return null_argument("fm_pagination_get_page_order");
  }
  *out_order = static_cast<int32_t>(pagination->result.page_order);
  return 0;
}

extern "C" fm_status_t fm_pagination_get_print_titles(const fm_pagination_t* pagination, fm_print_titles_t* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return null_argument("fm_pagination_get_print_titles");
  }
  const formulon::print::PrintTitleSpan& titles = pagination->result.print_titles;
  *out = fm_print_titles_t{};
  out->has_rows = titles.has_rows ? 1 : 0;
  out->first_row = titles.first_row;
  out->last_row = titles.last_row;
  out->has_cols = titles.has_cols ? 1 : 0;
  out->first_col = titles.first_col;
  out->last_col = titles.last_col;
  return 0;
}

extern "C" fm_status_t fm_pagination_get_page_at(const fm_pagination_t* pagination, size_t index,
                                                 fm_page_layout_t* out) {
  clear_last_error();
  if (pagination == nullptr || out == nullptr) {
    return null_argument("fm_pagination_get_page_at");
  }
  if (auto rc = check_index(index, pagination->result.pages.size(), "fm_pagination_get_page_at", "index"); rc != 0) {
    return rc;
  }
  const formulon::print::PageLayout& page = pagination->result.pages[index];
  *out = fm_page_layout_t{};
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

extern "C" fm_status_t fm_pagination_is_row_break_manual(const fm_pagination_t* pagination, size_t index,
                                                         int32_t* out_manual) {
  uint32_t value = 0;
  if (auto rc = break_entry(pagination, index, out_manual, true, true, "fm_pagination_is_row_break_manual", &value);
      rc != 0) {
    return rc;
  }
  *out_manual = static_cast<int32_t>(value);
  return 0;
}

extern "C" fm_status_t fm_pagination_is_col_break_manual(const fm_pagination_t* pagination, size_t index,
                                                         int32_t* out_manual) {
  uint32_t value = 0;
  if (auto rc = break_entry(pagination, index, out_manual, false, true, "fm_pagination_is_col_break_manual", &value);
      rc != 0) {
    return rc;
  }
  *out_manual = static_cast<int32_t>(value);
  return 0;
}
