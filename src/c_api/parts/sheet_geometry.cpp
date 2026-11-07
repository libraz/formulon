//
// C ABI - sheet geometry in points (cell rectangles, column widths, row
// heights, and the column width model).

#include "print/sheet_geometry.h"

#include <cstdint>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "column_width.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_column_span;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_finite_non_negative;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::check_sheet_row;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::publish_scratch;

namespace {

// Validates the shared `(sheet, mode)` prefix of the geometry entry points.
fm_status_t check_geometry_args(const fm_workbook_t* wb, size_t sheet_index, int32_t mode) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return check_enum_domain(mode, 0, FM_GEOMETRY_PRINT, "mode");
}

formulon::ColumnWidthModel width_model_for(const fm_workbook_t* wb, int32_t mode) {
  return formulon::resolve_column_width_model(wb->workbook().styles(), static_cast<formulon::GeometryMode>(mode));
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_cell_rect_pt(const fm_workbook_t* wb, size_t sheet_index,
                                                 const fm_cell_rect_t* range, int32_t mode, fm_rect_pt_t* out) {
  clear_last_error();
  if (range == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_geometry_args(wb, sheet_index, mode); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(*range); rc != 0) {
    return rc;
  }
  const std::uint32_t first_row = range->first_row;
  const std::uint32_t first_col = range->first_col;
  const std::uint32_t last_row = range->last_row;
  const std::uint32_t last_col = range->last_col;
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  const formulon::ColumnWidthModel model = width_model_for(wb, mode);
  double width = 0.0;
  for (const double chars : formulon::print::column_widths_chars(sheet, first_col, last_col, model)) {
    width += formulon::column_chars_to_points(chars, model);
  }
  const double y = formulon::print::row_offset_pt(sheet, first_row);
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_rect_pt_t{};
  out->struct_size = struct_size;
  out->x = formulon::print::column_offset_pt(sheet, first_col, model);
  out->y = y;
  out->width = width;
  out->height = first_row == last_row ? formulon::print::effective_row_height_pt(sheet, first_row)
                                      : formulon::print::row_offset_pt(sheet, last_row + 1U) - y;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_column_width_pt(const fm_workbook_t* wb, size_t sheet_index, uint32_t col,
                                                    int32_t mode, double* out_pt) {
  clear_last_error();
  if (out_pt == nullptr) {
    return null_argument();
  }
  if (auto rc = check_geometry_args(wb, sheet_index, mode); rc != 0) {
    return rc;
  }
  if (auto rc = check_column_span(col, col); rc != 0) {
    return rc;
  }
  *out_pt =
      formulon::print::effective_column_width_pt(wb->workbook().sheet(sheet_index), col, width_model_for(wb, mode));
  return 0;
}

extern "C" fm_status_t fm_sheet_get_row_height_pt(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                  double* out_pt) {
  clear_last_error();
  if (out_pt == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_row(wb, sheet_index, row); rc != 0) {
    return rc;
  }
  *out_pt = formulon::print::effective_row_height_pt(wb->workbook().sheet(sheet_index), row);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_width_model(const fm_workbook_t* wb, size_t sheet_index, int32_t mode,
                                                fm_width_model_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_geometry_args(wb, sheet_index, mode); rc != 0) {
    return rc;
  }
  const formulon::ColumnWidthModel model = width_model_for(wb, mode);
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_width_model_t{};
  out->struct_size = struct_size;
  out->points_per_char = model.points_per_char;
  out->padding_pt = model.padding_pt;
  out->normal_font_size = model.normal_font_size;
  out->calibrated = model.calibrated ? 1 : 0;
  out->normal_font_name = publish_scratch(wb, model.normal_font_name);
  out->platform = model.platform;
  return 0;
}

extern "C" fm_status_t fm_sheet_convert_column_chars_to_pt(const fm_workbook_t* wb, size_t sheet_index, int32_t mode,
                                                           double chars, double* out_pt) {
  clear_last_error();
  if (out_pt == nullptr) {
    return null_argument();
  }
  if (auto rc = check_geometry_args(wb, sheet_index, mode); rc != 0) {
    return rc;
  }
  if (auto rc = check_finite_non_negative(chars, "chars"); rc != 0) {
    return rc;
  }
  *out_pt = formulon::column_chars_to_points(chars, width_model_for(wb, mode));
  return 0;
}

extern "C" fm_status_t fm_sheet_convert_column_pt_to_chars(const fm_workbook_t* wb, size_t sheet_index, int32_t mode,
                                                           double pt, double* out_chars) {
  clear_last_error();
  if (out_chars == nullptr) {
    return null_argument();
  }
  if (auto rc = check_geometry_args(wb, sheet_index, mode); rc != 0) {
    return rc;
  }
  if (auto rc = check_finite_non_negative(pt, "pt"); rc != 0) {
    return rc;
  }
  *out_chars = formulon::column_points_to_chars(pt, width_model_for(wb, mode));
  return 0;
}
