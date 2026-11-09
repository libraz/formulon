#pragma once

#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"

// `fm_sheet_update_view` with the named fields SET and every other field KEEP.

template <typename T>
fm_status_t UpdateViewField(fm_workbook_t* wb, std::size_t sheet, fm_patch_op_t fm_sheet_view_patch_t::*op,
                            T fm_sheet_view_patch_t::*field, T value) {
  fm_sheet_view_patch_t patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  patch.*op = FM_PATCH_OP_SET;
  patch.*field = value;
  return fm_sheet_update_view(wb, sheet, &patch);
}

template <typename T>
fm_status_t UpdateViewPair(fm_workbook_t* wb, std::size_t sheet, fm_patch_op_t fm_sheet_view_patch_t::*first_op,
                           T fm_sheet_view_patch_t::*first, T first_value,
                           fm_patch_op_t fm_sheet_view_patch_t::*second_op, T fm_sheet_view_patch_t::*second,
                           T second_value) {
  fm_sheet_view_patch_t patch = FM_STRUCT_INIT(fm_sheet_view_patch_t);
  patch.*first_op = FM_PATCH_OP_SET;
  patch.*first = first_value;
  patch.*second_op = FM_PATCH_OP_SET;
  patch.*second = second_value;
  return fm_sheet_update_view(wb, sheet, &patch);
}

inline fm_status_t SetViewZoom(fm_workbook_t* wb, std::size_t sheet, std::uint32_t zoom_pct) {
  return UpdateViewField(wb, sheet, &fm_sheet_view_patch_t::zoom_pct_op, &fm_sheet_view_patch_t::zoom_pct, zoom_pct);
}

inline fm_status_t SetViewFreeze(fm_workbook_t* wb, std::size_t sheet, std::uint32_t rows, std::uint32_t cols) {
  return UpdateViewPair(wb, sheet, &fm_sheet_view_patch_t::freeze_rows_op, &fm_sheet_view_patch_t::freeze_rows, rows,
                        &fm_sheet_view_patch_t::freeze_cols_op, &fm_sheet_view_patch_t::freeze_cols, cols);
}

inline fm_status_t SetViewSplit(fm_workbook_t* wb, std::size_t sheet, double x_twips, double y_twips) {
  return UpdateViewPair(wb, sheet, &fm_sheet_view_patch_t::split_x_op, &fm_sheet_view_patch_t::split_x, x_twips,
                        &fm_sheet_view_patch_t::split_y_op, &fm_sheet_view_patch_t::split_y, y_twips);
}

inline fm_status_t SetViewTopLeft(fm_workbook_t* wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  return UpdateViewPair(wb, sheet, &fm_sheet_view_patch_t::top_left_row_op, &fm_sheet_view_patch_t::top_left_row, row,
                        &fm_sheet_view_patch_t::top_left_col_op, &fm_sheet_view_patch_t::top_left_col, col);
}

inline fm_status_t SetViewPaneTopLeft(fm_workbook_t* wb, std::size_t sheet, std::uint32_t row, std::uint32_t col) {
  return UpdateViewPair(wb, sheet, &fm_sheet_view_patch_t::pane_top_left_row_op,
                        &fm_sheet_view_patch_t::pane_top_left_row, row, &fm_sheet_view_patch_t::pane_top_left_col_op,
                        &fm_sheet_view_patch_t::pane_top_left_col, col);
}

inline fm_status_t SetViewVisibility(fm_workbook_t* wb, std::size_t sheet, std::int32_t visibility) {
  return UpdateViewField(wb, sheet, &fm_sheet_view_patch_t::visibility_op, &fm_sheet_view_patch_t::visibility,
                         visibility);
}

inline fm_status_t SetViewMode(fm_workbook_t* wb, std::size_t sheet, std::int32_t mode) {
  return UpdateViewField(wb, sheet, &fm_sheet_view_patch_t::view_mode_op, &fm_sheet_view_patch_t::view_mode, mode);
}

// One 0/1 display flag: `show_grid_lines`, `right_to_left`, `tab_selected`, ...
inline fm_status_t SetViewFlag(fm_workbook_t* wb, std::size_t sheet, fm_patch_op_t fm_sheet_view_patch_t::*op,
                               std::int32_t fm_sheet_view_patch_t::*flag, std::int32_t value) {
  return UpdateViewField(wb, sheet, op, flag, value);
}
