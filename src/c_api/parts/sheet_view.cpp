//
// C ABI - sheet view state (zoom, freeze, split, selection, visibility, and
// the per-sheet display flags). Bridges `Sheet::view()` over the stable C ABI.

#include <cmath>
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
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok_without_clear;

namespace {

// `<sheetView zoomScale>` is a percentage Excel accepts only inside this band.
constexpr std::uint32_t kMinZoomPercent = 10U;
constexpr std::uint32_t kMaxZoomPercent = 400U;

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

extern "C" fm_status_t fm_sheet_get_view(const fm_workbook_t* wb, size_t sheet_index, fm_sheet_view_info_t* out) {
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
  out->zoom_pct = v.zoom_scale;
  out->freeze_rows = v.freeze_rows;
  out->freeze_cols = v.freeze_cols;
  out->show_grid_lines = v.show_grid_lines ? 1 : 0;
  out->show_row_column_headers = v.show_row_col_headers ? 1 : 0;
  out->show_zeros = v.show_zeros ? 1 : 0;
  out->right_to_left = v.right_to_left ? 1 : 0;
  out->tab_selected = v.tab_selected ? 1 : 0;
  out->view_mode = view_mode_to_c(v.view_mode);
  out->top_left_row = 0;
  out->top_left_col = 0;
  (void)formulon::a1::parse_a1_ref(v.top_left_cell, &out->top_left_row, &out->top_left_col);
  formulon::io::pane_origin(v, &out->pane_top_left_row, &out->pane_top_left_col);
  out->split_x_twips = v.is_frozen() ? 0.0 : v.split_x;
  out->split_y_twips = v.is_frozen() ? 0.0 : v.split_y;
  out->show_formulas = v.show_formulas ? 1 : 0;
  out->show_outline_symbols = v.show_outline_symbols ? 1 : 0;
  return 0;
}

namespace {

// One 0/1 display flag of `fm_sheet_view_patch_t` and its model member.
struct ViewFlag {
  fm_patch_op_t op;
  std::int32_t value;
  bool formulon::SheetView::*member;
};

// Whether a stated value of `op` / `value` is non-zero.
template <typename T>
bool states_non_zero(fm_patch_op_t op, T value) {
  return op == FM_PATCH_OP_SET && value != T{};
}

// The patched value of one field: `value` on SET, `current` otherwise.
template <typename T>
T patched(fm_patch_op_t op, T value, T current) {
  return op == FM_PATCH_OP_SET ? value : current;
}

}  // namespace

extern "C" fm_status_t fm_sheet_update_view(fm_workbook_t* wb, size_t sheet_index, const fm_sheet_view_patch_t* patch) {
  clear_last_error();
  if (patch == nullptr) {
    return null_argument("patch");
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const ViewFlag flags[] = {
      {patch->show_grid_lines_op, patch->show_grid_lines, &formulon::SheetView::show_grid_lines},
      {patch->show_row_column_headers_op, patch->show_row_column_headers, &formulon::SheetView::show_row_col_headers},
      {patch->show_zeros_op, patch->show_zeros, &formulon::SheetView::show_zeros},
      {patch->show_formulas_op, patch->show_formulas, &formulon::SheetView::show_formulas},
      {patch->show_outline_symbols_op, patch->show_outline_symbols, &formulon::SheetView::show_outline_symbols},
      {patch->right_to_left_op, patch->right_to_left, &formulon::SheetView::right_to_left},
      {patch->tab_selected_op, patch->tab_selected, &formulon::SheetView::tab_selected},
  };
  const fm_patch_op_t ops[] = {
      patch->zoom_pct_op,      patch->freeze_rows_op,   patch->freeze_cols_op,       patch->view_mode_op,
      patch->top_left_row_op,  patch->top_left_col_op,  patch->pane_top_left_row_op, patch->pane_top_left_col_op,
      patch->split_x_twips_op, patch->split_y_twips_op,
  };
  for (const fm_patch_op_t op : ops) {
    if (!patch_op_ok_without_clear(op)) {
      return invalid_argument("op is not a patch op", "op=" + std::to_string(op));
    }
  }
  for (const ViewFlag& flag : flags) {
    if (!patch_op_ok_without_clear(flag.op)) {
      return invalid_argument("op is not a patch op", "op=" + std::to_string(flag.op));
    }
  }
  if (patch->zoom_pct_op == FM_PATCH_OP_SET &&
      (patch->zoom_pct < kMinZoomPercent || patch->zoom_pct > kMaxZoomPercent)) {
    return invalid_argument(
        ("zoom_pct must be in " + std::to_string(kMinZoomPercent) + ".." + std::to_string(kMaxZoomPercent)).c_str(),
        "zoom_pct=" + std::to_string(patch->zoom_pct));
  }
  if (patch->view_mode_op == FM_PATCH_OP_SET) {
    if (auto rc = check_enum_domain(patch->view_mode, 0, FM_SHEET_VIEW_MODE_PAGE_LAYOUT, "view_mode"); rc != 0) {
      return rc;
    }
  }
  // A freeze replaces a split and a split a freeze, so one patch states one.
  if ((states_non_zero(patch->freeze_rows_op, patch->freeze_rows) ||
       states_non_zero(patch->freeze_cols_op, patch->freeze_cols)) &&
      (states_non_zero(patch->split_x_twips_op, patch->split_x_twips) ||
       states_non_zero(patch->split_y_twips_op, patch->split_y_twips))) {
    return invalid_argument("split_x_twips cannot be combined with freeze_rows");
  }

  formulon::Workbook& workbook = wb->workbook();
  const formulon::SheetView& current = workbook.sheet(sheet_index).view();
  formulon::SheetView next = current;

  const std::uint32_t freeze_rows = patched(patch->freeze_rows_op, patch->freeze_rows, current.freeze_rows);
  const std::uint32_t freeze_cols = patched(patch->freeze_cols_op, patch->freeze_cols, current.freeze_cols);
  if (patch->freeze_rows_op == FM_PATCH_OP_SET || patch->freeze_cols_op == FM_PATCH_OP_SET) {
    // The first unfrozen cell must lie inside the grid.
    if (!formulon::Sheet::coord_in_grid(freeze_rows, freeze_cols)) {
      return invalid_argument(
          "freeze_rows and freeze_cols must leave the first unfrozen cell inside the grid",
          "freeze_rows=" + std::to_string(freeze_rows) + " freeze_cols=" + std::to_string(freeze_cols));
    }
    next.set_freeze(freeze_rows, freeze_cols);
  }
  if (patch->split_x_twips_op == FM_PATCH_OP_SET || patch->split_y_twips_op == FM_PATCH_OP_SET) {
    // The split a frozen view reports is (0, 0).
    const double split_x =
        patched(patch->split_x_twips_op, patch->split_x_twips, current.is_frozen() ? 0.0 : current.split_x);
    const double split_y =
        patched(patch->split_y_twips_op, patch->split_y_twips, current.is_frozen() ? 0.0 : current.split_y);
    if (!std::isfinite(split_x) || !std::isfinite(split_y) || split_x < 0.0 || split_y < 0.0) {
      return invalid_argument("split_x_twips and split_y_twips must be finite and not negative");
    }
    // Clearing the split leaves a freeze in place; any other split replaces it.
    if (!(next.is_frozen() && split_x == 0.0 && split_y == 0.0)) {
      next.set_split(split_x, split_y);
    }
  }
  if (patch->pane_top_left_row_op == FM_PATCH_OP_SET || patch->pane_top_left_col_op == FM_PATCH_OP_SET) {
    std::uint32_t origin_row = 0;
    std::uint32_t origin_col = 0;
    formulon::io::pane_origin(next, &origin_row, &origin_col);
    const std::uint32_t row = patched(patch->pane_top_left_row_op, patch->pane_top_left_row, origin_row);
    const std::uint32_t col = patched(patch->pane_top_left_col_op, patch->pane_top_left_col, origin_col);
    const bool has_panes = next.divides_rows() || next.divides_cols();
    // A pane-less view has one scrolled region, so only its default origin is accepted.
    if (!has_panes) {
      if (row != next.freeze_rows || col != next.freeze_cols) {
        return invalid_argument("pane_top_left_row and pane_top_left_col must be the default on a view without panes");
      }
      next.pane_top_left_cell.clear();
    } else {
      if (row < next.freeze_rows || col < next.freeze_cols || !formulon::Sheet::coord_in_grid(row, col)) {
        return invalid_argument("pane_top_left_row and pane_top_left_col must be in the unfrozen area");
      }
      next.pane_top_left_cell =
          row == next.freeze_rows && col == next.freeze_cols ? std::string() : formulon::a1::encode_a1(row, col);
    }
  }
  if (patch->top_left_row_op == FM_PATCH_OP_SET || patch->top_left_col_op == FM_PATCH_OP_SET) {
    std::uint32_t current_row = 0;
    std::uint32_t current_col = 0;
    (void)formulon::a1::parse_a1_ref(current.top_left_cell, &current_row, &current_col);
    const std::uint32_t row = patched(patch->top_left_row_op, patch->top_left_row, current_row);
    const std::uint32_t col = patched(patch->top_left_col_op, patch->top_left_col, current_col);
    if (auto rc = check_cell_coord(row, col); rc != 0) {
      return rc;
    }
    next.top_left_cell = row == 0U && col == 0U ? std::string() : formulon::a1::encode_a1(row, col);
  }

  // Every check passed; nothing below can fail.
  if (patch->zoom_pct_op == FM_PATCH_OP_SET) {
    next.zoom_scale = patch->zoom_pct;
  }
  if (patch->view_mode_op == FM_PATCH_OP_SET) {
    next.view_mode = std::string(view_mode_from_c(patch->view_mode));
  }
  for (const ViewFlag& flag : flags) {
    if (flag.op == FM_PATCH_OP_SET) {
      next.*flag.member = flag.value != 0;
    }
  }
  workbook.sheet(sheet_index).mutable_view() = std::move(next);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_visibility(const fm_workbook_t* wb, size_t sheet_index, int32_t* out_visibility) {
  clear_last_error();
  if (out_visibility == nullptr) {
    return null_argument("out_visibility");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  // `visibility()` resolves the model's two bits, including the
  // combination no writer should produce.
  *out_visibility = static_cast<int32_t>(wb->workbook().sheet(sheet_index).view().visibility());
  return 0;
}

extern "C" fm_status_t fm_sheet_set_visibility(fm_workbook_t* wb, size_t sheet_index, int32_t visibility) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc =
          check_enum_domain(visibility, FM_SHEET_VISIBILITY_VISIBLE, FM_SHEET_VISIBILITY_VERY_HIDDEN, "visibility");
      rc != 0) {
    return rc;
  }
  formulon::Workbook& workbook = wb->workbook();
  const auto index = static_cast<std::uint32_t>(sheet_index);
  if (visibility != FM_SHEET_VISIBILITY_VISIBLE) {
    if (auto rc = check_keeps_visible_sheet(workbook, index); rc != 0) {
      return rc;
    }
  }
  workbook.sheet(sheet_index).mutable_view().set_visibility(static_cast<formulon::SheetVisibility>(visibility));
  activate_successor_if_hidden(workbook, index);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_selection(const fm_workbook_t* wb, size_t sheet_index, uint32_t* out_active_row,
                                              uint32_t* out_active_col, fm_cell_rect_t* out, size_t capacity,
                                              size_t* out_count) {
  clear_last_error();
  if (out_active_row == nullptr || out_active_col == nullptr || out_count == nullptr ||
      (out == nullptr && capacity != 0U)) {
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
  *out_count = ranges.size();
  for (std::size_t i = 0; i < ranges.size() && i < capacity; ++i) {
    out[i] = {ranges[i].first_row, ranges[i].first_col, ranges[i].last_row, ranges[i].last_col};
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
      return invalid_argument(("ranges[" + std::to_string(i) + "] is outside the grid or inverted").c_str());
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
