//
// C ABI - sheet view state (zoom, freeze, split, selection, visibility, and
// the per-sheet display flags). Bridges `Sheet::view()` over the stable C ABI.

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "io/sqref_text.h"
#include "sheet.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_keeps_visible_sheet;
using formulon::c_api::parts::check_limit;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;

extern "C" fm_status_t fm_sheet_set_zoom(fm_workbook_t* wb, size_t sheet_index, uint32_t zoom_scale) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  // Clamp to the OOXML-valid `[10, 400]` interval; out-of-range values
  // are rounded to the nearest endpoint rather than rejected so JS
  // callers do not have to mirror the bound.
  std::uint32_t clamped = zoom_scale;
  if (clamped < 10U) {
    clamped = 10U;
  } else if (clamped > 400U) {
    clamped = 400U;
  }
  wb->workbook().sheet(sheet_index).mutable_view().zoom_scale = clamped;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_freeze(fm_workbook_t* wb, size_t sheet_index, uint32_t freeze_rows,
                                           uint32_t freeze_cols) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_limit(freeze_rows, formulon::Sheet::kMaxRows, "freeze_rows"); rc != 0) {
    return rc;
  }
  if (auto rc = check_limit(freeze_cols, formulon::Sheet::kMaxCols, "freeze_cols"); rc != 0) {
    return rc;
  }
  formulon::SheetView& view = wb->workbook().sheet(sheet_index).mutable_view();
  view.set_freeze(freeze_rows, freeze_cols);
  return 0;
}

namespace {

// `<sheetView view>` text of an `fm_sheet_view_mode_t`; the normal view is
// the absent attribute.
std::string_view view_mode_from_c(std::int32_t mode) {
  switch (mode) {
    case FM_SHEET_VIEW_MODE_PAGE_BREAK_PREVIEW:
      return "pageBreakPreview";
    case FM_SHEET_VIEW_MODE_PAGE_LAYOUT:
      return "pageLayout";
    default:
      return "";
  }
}

// An unrecognised `view` value reads as the normal view; the stored text
// still round-trips on save.
std::int32_t view_mode_to_c(std::string_view mode) {
  if (mode == "pageBreakPreview") {
    return FM_SHEET_VIEW_MODE_PAGE_BREAK_PREVIEW;
  }
  if (mode == "pageLayout") {
    return FM_SHEET_VIEW_MODE_PAGE_LAYOUT;
  }
  return FM_SHEET_VIEW_MODE_NORMAL;
}

// Hiding the active sheet hands the active tab on as Excel does. Kept out of
// the Workbook setters so a load replaying a hidden active sheet keeps it.
void activate_successor_if_hidden(formulon::Workbook& workbook, std::uint32_t index) {
  if (workbook.sheet(index).view().visibility() != formulon::SheetVisibility::kVisible &&
      workbook.active_sheet() == index) {
    if (const std::uint32_t successor = workbook.nearest_visible_sheet(index); successor != index) {
      workbook.set_active_sheet(successor);
    }
  }
}

}  // namespace

extern "C" fm_status_t fm_sheet_set_visibility(fm_workbook_t* wb, size_t sheet_index, int32_t visibility) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (visibility < static_cast<int32_t>(FM_SHEET_VISIBLE) || visibility > static_cast<int32_t>(FM_SHEET_VERY_HIDDEN)) {
    return invalid_argument("unknown visibility", "visibility=" + std::to_string(visibility));
  }
  if (visibility != static_cast<int32_t>(FM_SHEET_VISIBLE)) {
    if (auto rc = check_keeps_visible_sheet(wb->workbook(), static_cast<std::uint32_t>(sheet_index)); rc != 0) {
      return rc;
    }
  }
  formulon::SheetView& view = wb->workbook().sheet(sheet_index).mutable_view();
  view.set_visibility(static_cast<formulon::SheetVisibility>(visibility));
  activate_successor_if_hidden(wb->workbook(), static_cast<std::uint32_t>(sheet_index));
  return 0;
}

extern "C" fm_status_t fm_sheet_get_view(const fm_workbook_t* wb, size_t sheet_index, fm_sheet_view_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const formulon::SheetView& v = wb->workbook().sheet(sheet_index).view();
  out->zoom_scale = v.zoom_scale;
  out->freeze_rows = v.freeze_rows;
  out->freeze_cols = v.freeze_cols;
  // `visibility()` resolves the model's two bits, including the
  // combination no writer should produce.
  out->visibility = static_cast<int32_t>(v.visibility());
  out->show_grid_lines = v.show_grid_lines ? 1 : 0;
  out->show_row_col_headers = v.show_row_col_headers ? 1 : 0;
  out->show_zeros = v.show_zeros ? 1 : 0;
  out->right_to_left = v.right_to_left ? 1 : 0;
  out->tab_selected = v.tab_selected ? 1 : 0;
  out->view_mode = view_mode_to_c(v.view_mode);
  out->top_left_row = 0;
  out->top_left_col = 0;
  (void)formulon::a1::parse_a1_ref(v.top_left_cell, &out->top_left_row, &out->top_left_col);
  formulon::io::pane_origin(v, &out->pane_top_left_row, &out->pane_top_left_col);
  out->split_x = v.is_frozen() ? 0.0 : v.split_x;
  out->split_y = v.is_frozen() ? 0.0 : v.split_y;
  out->show_formulas = v.show_formulas ? 1 : 0;
  out->show_outline_symbols = v.show_outline_symbols ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_top_left(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().top_left_cell =
      row == 0U && col == 0U ? std::string() : formulon::a1::encode_a1(row, col);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_pane_top_left(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  formulon::SheetView& view = wb->workbook().sheet(sheet_index).mutable_view();
  // A pane-less view has one scrolled region, so its default origin is accepted.
  if (!view.divides_rows() && !view.divides_cols() && row == view.freeze_rows && col == view.freeze_cols) {
    view.pane_top_left_cell.clear();
    return 0;
  }
  if ((!view.divides_rows() && !view.divides_cols()) || row < view.freeze_rows || col < view.freeze_cols ||
      !formulon::Sheet::coord_in_grid(row, col)) {
    return invalid_argument("cell is not in the unfrozen area");
  }
  view.pane_top_left_cell =
      row == view.freeze_rows && col == view.freeze_cols ? std::string() : formulon::a1::encode_a1(row, col);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_split(fm_workbook_t* wb, size_t sheet_index, double x_twips, double y_twips) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (!wb->workbook().sheet(sheet_index).mutable_view().set_split(x_twips, y_twips)) {
    return invalid_argument("split must be finite and non-negative");
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_selection(const fm_workbook_t* wb, size_t sheet_index, uint32_t* out_active_row,
                                              uint32_t* out_active_col, fm_cell_rect_t* out_ranges, size_t capacity,
                                              size_t* out_total) {
  clear_last_error();
  if (out_active_row == nullptr || out_active_col == nullptr || out_total == nullptr ||
      (out_ranges == nullptr && capacity != 0U)) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  // The active pane's selection as it would be saved.
  const formulon::SheetView& view = wb->workbook().sheet(sheet_index).view();
  const std::string_view active_pane = formulon::SheetView::pane_key(view.effective_active_pane());
  std::vector<formulon::io::ResolvedSelection> resolved = formulon::io::resolve_view_selections(view);
  const formulon::io::ResolvedSelection* selection = &resolved.front();
  for (const formulon::io::ResolvedSelection& candidate : resolved) {
    if (candidate.pane == active_pane) {
      selection = &candidate;
    }
  }
  *out_active_row = selection->active_row;
  *out_active_col = selection->active_col;
  const std::vector<formulon::MergeRange>& ranges = selection->ranges;
  *out_total = ranges.size();
  for (std::size_t i = 0; i < ranges.size() && i < capacity; ++i) {
    out_ranges[i] = {ranges[i].first_row, ranges[i].first_col, ranges[i].last_row, ranges[i].last_col};
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_set_selection(fm_workbook_t* wb, size_t sheet_index, uint32_t active_row,
                                              uint32_t active_col, const fm_cell_rect_t* ranges, size_t range_count) {
  clear_last_error();
  if (range_count != 0U && ranges == nullptr) {
    return null_argument("ranges");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(active_row, active_col); rc != 0) {
    return rc;
  }
  std::vector<formulon::MergeRange> selected;
  for (std::size_t i = 0; i < range_count; ++i) {
    const fm_cell_rect_t& r = ranges[i];
    if (!formulon::Sheet::coord_in_grid(r.first_row, r.first_col) ||
        !formulon::Sheet::coord_in_grid(r.last_row, r.last_col) || r.first_row > r.last_row ||
        r.first_col > r.last_col) {
      return invalid_argument("range out of grid or inverted");
    }
    selected.push_back({r.first_row, r.first_col, r.last_row, r.last_col});
  }
  formulon::SheetView& view = wb->workbook().sheet(sheet_index).mutable_view();
  formulon::SheetSelection replacement{std::string(formulon::SheetView::pane_key(view.effective_active_pane())),
                                       formulon::a1::encode_a1(active_row, active_col),
                                       formulon::io::format_sqref_text(selected)};
  for (formulon::SheetSelection& selection : view.selections) {
    if (formulon::SheetView::pane_key(selection.pane) == replacement.pane) {
      selection = std::move(replacement);
      return 0;
    }
  }
  view.selections.push_back(std::move(replacement));
  return 0;
}

namespace {

fm_status_t set_view_flag(fm_workbook_t* wb, size_t sheet_index, bool formulon::SheetView::*member, int32_t value) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().*member = (value != 0);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_set_show_grid_lines(fm_workbook_t* wb, size_t sheet_index, int32_t show) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, &formulon::SheetView::show_grid_lines, show);
}

extern "C" fm_status_t fm_sheet_set_show_row_col_headers(fm_workbook_t* wb, size_t sheet_index, int32_t show) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, &formulon::SheetView::show_row_col_headers, show);
}

extern "C" fm_status_t fm_sheet_set_show_zeros(fm_workbook_t* wb, size_t sheet_index, int32_t show) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, &formulon::SheetView::show_zeros, show);
}

extern "C" fm_status_t fm_sheet_set_show_formulas(fm_workbook_t* wb, size_t sheet_index, int32_t show) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, &formulon::SheetView::show_formulas, show);
}

extern "C" fm_status_t fm_sheet_set_show_outline_symbols(fm_workbook_t* wb, size_t sheet_index, int32_t show) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, &formulon::SheetView::show_outline_symbols, show);
}

extern "C" fm_status_t fm_sheet_set_right_to_left(fm_workbook_t* wb, size_t sheet_index, int32_t right_to_left) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, &formulon::SheetView::right_to_left, right_to_left);
}

extern "C" fm_status_t fm_sheet_set_tab_selected(fm_workbook_t* wb, size_t sheet_index, int32_t selected) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, &formulon::SheetView::tab_selected, selected);
}

extern "C" fm_status_t fm_sheet_set_view_mode(fm_workbook_t* wb, size_t sheet_index, int32_t mode) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(mode, 0, FM_SHEET_VIEW_MODE_PAGE_LAYOUT, "mode"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().view_mode = std::string(view_mode_from_c(mode));
  return 0;
}
