//
// C ABI - sheet view state (zoom, freeze, split, selection, visibility, and
// the per-sheet display flags). Bridges `Sheet::view()` over the stable C ABI.

#include <cmath>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "drawing/drawing_xml.h"
#include "io/passthrough_part_index.h"
#include "io/sqref_text.h"
#include "io/xlsb/chartsheet_records.h"
#include "sheet.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "workbook.h"

using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_keeps_visible_sheet;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::refuse_opaque_change;
using formulon::c_api::parts::reset_record;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::SheetUse;
using formulon::c_api::parts::store_cstr_or_null;

namespace {

// `<sheetView zoomScale>` is a percentage Excel accepts only inside this band.
constexpr std::uint32_t kMinZoomPercent = 10U;
// @limit: excel
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

// The fields of `out` after `struct_size`, as `fm_sheet_get_view` reports `v`.
void fill_view_info(const formulon::SheetView& v, fm_sheet_view_info_t* out) {
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
  out->zoom_to_fit = v.zoom_to_fit ? 1 : 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_view(const fm_workbook_t* wb, size_t sheet_index, fm_sheet_view_info_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheetOrChartsheet); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  fill_view_info(wb->workbook().sheet(sheet_index).view(), out);
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

// Whether `op` / `value` sets a field to something other than its reported `current`.
template <typename T>
bool changes(fm_patch_op_t op, T value, T current) {
  return op == FM_PATCH_OP_SET && value != current;
}

// Whether `patch` changes any field of a view reported as `current` other than the zoom and the selected tab.
bool changes_beyond_zoom_and_tab(const fm_sheet_view_patch_t& patch, const fm_sheet_view_info_t& current) {
  return changes(patch.freeze_rows_op, patch.freeze_rows, current.freeze_rows) ||
         changes(patch.freeze_cols_op, patch.freeze_cols, current.freeze_cols) ||
         changes(patch.view_mode_op, patch.view_mode, current.view_mode) ||
         changes(patch.top_left_row_op, patch.top_left_row, current.top_left_row) ||
         changes(patch.top_left_col_op, patch.top_left_col, current.top_left_col) ||
         changes(patch.pane_top_left_row_op, patch.pane_top_left_row, current.pane_top_left_row) ||
         changes(patch.pane_top_left_col_op, patch.pane_top_left_col, current.pane_top_left_col) ||
         changes(patch.split_x_twips_op, patch.split_x_twips, current.split_x_twips) ||
         changes(patch.split_y_twips_op, patch.split_y_twips, current.split_y_twips) ||
         changes(patch.show_grid_lines_op, patch.show_grid_lines != 0, current.show_grid_lines != 0) ||
         changes(patch.show_row_column_headers_op, patch.show_row_column_headers != 0,
                 current.show_row_column_headers != 0) ||
         changes(patch.show_zeros_op, patch.show_zeros != 0, current.show_zeros != 0) ||
         changes(patch.show_formulas_op, patch.show_formulas != 0, current.show_formulas != 0) ||
         changes(patch.show_outline_symbols_op, patch.show_outline_symbols != 0, current.show_outline_symbols != 0) ||
         changes(patch.right_to_left_op, patch.right_to_left != 0, current.right_to_left != 0);
}

}  // namespace

extern "C" fm_status_t fm_sheet_update_view(fm_workbook_t* wb, size_t sheet_index, const fm_sheet_view_patch_t* patch) {
  clear_last_error();
  if (patch == nullptr) {
    return null_argument("patch");
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheetOrChartsheet); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
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
  // A chartsheet keeps its stored view except for a new zoom (which ends zoom-to-fit) and its selected tab.
  if (workbook.sheet(sheet_index).is_opaque_ooxml_sheet()) {
    fm_sheet_view_info_t reported{};
    fill_view_info(current, &reported);
    if (changes_beyond_zoom_and_tab(*patch, reported)) {
      return refuse_opaque_change("patch", "a field other than zoom_pct or tab_selected", sheet_index);
    }
    formulon::SheetView& view = workbook.sheet(sheet_index).mutable_view();
    if (changes(patch->zoom_pct_op, patch->zoom_pct, reported.zoom_pct)) {
      view.zoom_scale = patch->zoom_pct;
      view.zoom_to_fit = false;
    }
    if (patch->tab_selected_op == FM_PATCH_OP_SET) {
      view.tab_selected = patch->tab_selected != 0;
    }
    return 0;
  }
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

extern "C" fm_status_t fm_sheet_get_kind(const fm_workbook_t* wb, size_t sheet_index, int32_t* out_kind) {
  clear_last_error();
  if (out_kind == nullptr) {
    return null_argument("out_kind");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_kind = static_cast<int32_t>(wb->workbook().sheet(sheet_index).kind());
  return 0;
}

extern "C" fm_status_t fm_sheet_get_chartsheet_chart(const fm_workbook_t* wb, size_t sheet_index, int32_t* out_present,
                                                     uint32_t* out_object_id) {
  clear_last_error();
  if (out_present == nullptr) {
    return null_argument("out_present");
  }
  if (out_object_id == nullptr) {
    return null_argument("out_object_id");
  }
  *out_present = 0;
  *out_object_id = 0;
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (wb->workbook().sheet(sheet_index).kind() != formulon::SheetKind::kChartsheet) {
    return 0;
  }
  auto parts = formulon::list_chart_parts(wb->workbook(), sheet_index);
  if (!parts) {
    return set_last_error(parts.error());
  }
  if (!parts.value().empty()) {
    *out_present = 1;
    *out_object_id = parts.value().front().object_id;
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_chartsheet_protection(const fm_workbook_t* wb, size_t sheet_index,
                                                          int32_t* out_present, fm_chartsheet_protection_t* out) {
  clear_last_error();
  if (out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  reset_record(out);
  *out_present = 0;
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  if (sheet.kind() != formulon::SheetKind::kChartsheet) {
    return 0;
  }
  const std::optional<formulon::ChartsheetProtection>& stored = sheet.chartsheet_protection();
  if (!stored) {
    return 0;
  }
  // Several strings are live at once, so they share one scratch generation.
  formulon::c_api::parts::TextStore& store = fresh_read_scratch(wb);
  const auto text = [&store](const std::string& value) { return store_cstr_or_null(store, value); };
  out->algorithm_name = text(stored->algorithm_name);
  out->hash_value = text(stored->hash_value);
  out->salt_value = text(stored->salt_value);
  out->spin_count = stored->spin_count;
  out->legacy_password_hash = text(stored->legacy_password_hash);
  out->allow_edit_content = stored->content_locked ? 0 : 1;
  out->allow_edit_objects = stored->objects_locked ? 0 : 1;
  *out_present = 1;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_chartsheet_protection(fm_workbook_t* wb, size_t sheet_index,
                                                          const fm_chartsheet_protection_t* in) {
  clear_last_error();
  if (in != nullptr) {
    if (auto rc = check_record_size(*in); rc != 0) {
      return rc;
    }
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kChartsheet); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  std::optional<formulon::ChartsheetProtection> wanted;
  if (in != nullptr) {
    const std::pair<const char*, const char*> texts[] = {{in->algorithm_name, "algorithm_name"},
                                                         {in->hash_value, "hash_value"},
                                                         {in->salt_value, "salt_value"},
                                                         {in->legacy_password_hash, "legacy_password_hash"}};
    for (const auto& [value, field] : texts) {
      if (value != nullptr && value[0] == '\0') {
        return invalid_argument((std::string(field) + " must not be empty").c_str(), std::string("field=") + field);
      }
    }
    if (auto rc = check_utf8({in->algorithm_name, in->hash_value, in->salt_value, in->legacy_password_hash}); rc != 0) {
      return rc;
    }
    const std::pair<std::int32_t, const char*> flags[] = {{in->allow_edit_content, "allow_edit_content"},
                                                          {in->allow_edit_objects, "allow_edit_objects"}};
    for (const auto& [value, field] : flags) {
      if (value != 0 && value != 1) {
        return invalid_argument((std::string(field) + " is out of range").c_str());
      }
    }
    formulon::ChartsheetProtection model;
    model.algorithm_name = in->algorithm_name != nullptr ? in->algorithm_name : "";
    model.hash_value = in->hash_value != nullptr ? in->hash_value : "";
    model.salt_value = in->salt_value != nullptr ? in->salt_value : "";
    model.spin_count = in->spin_count;
    model.legacy_password_hash = in->legacy_password_hash != nullptr ? in->legacy_password_hash : "";
    model.content_locked = in->allow_edit_content == 0;
    model.objects_locked = in->allow_edit_objects == 0;
    wanted = std::move(model);
  }
  const std::optional<formulon::ChartsheetProtection>& current = sheet.chartsheet_protection();
  if (wanted) {
    // A record that protects nothing is not a protection to store: it restates the current one or none.
    const formulon::ChartsheetProtection none;
    if (*wanted != none) {
      sheet.set_chartsheet_protection(std::move(wanted));
      return 0;
    }
    if (!current || *current == none) {
      return 0;
    }
  } else if (!current) {
    return 0;
  }
  return invalid_argument("Chartsheet protection cannot be removed", "sheet_index=" + std::to_string(sheet_index));
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
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
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
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
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
