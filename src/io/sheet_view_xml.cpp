#include "io/sheet_view_xml.h"

#include <cmath>
#include <cstdint>
#include <string>
#include <string_view>

#include "io/sqref_text.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/a1_ref.h"

namespace formulon {
namespace io {

void read_sheet_view_element(const pugi::xml_node& sheet_view, SheetView& view) {
  if (sheet_view.attribute("zoomScale")) {
    const std::int32_t raw = attr_i32(sheet_view, "zoomScale", 0);
    if (raw >= 10 && raw <= 400) {
      view.zoom_scale = static_cast<std::uint32_t>(raw);
    }
  }
  // Display attributes. Three default to true in the schema, so absence
  // means "shown"; the tri-state reader applies each attribute's real
  // default rather than a blanket false.
  view.show_grid_lines = read_xsd_bool(sheet_view, "showGridLines", true);
  view.show_row_col_headers = read_xsd_bool(sheet_view, "showRowColHeaders", true);
  view.show_zeros = read_xsd_bool(sheet_view, "showZeros", true);
  view.right_to_left = read_xsd_bool(sheet_view, "rightToLeft", false);
  view.tab_selected = read_xsd_bool(sheet_view, "tabSelected", false);
  view.window_protection = read_xsd_bool(sheet_view, "windowProtection", false);
  view.show_formulas = read_xsd_bool(sheet_view, "showFormulas", false);
  view.show_ruler = read_xsd_bool(sheet_view, "showRuler", true);
  view.show_outline_symbols = read_xsd_bool(sheet_view, "showOutlineSymbols", true);
  view.default_grid_color = read_xsd_bool(sheet_view, "defaultGridColor", true);
  view.show_white_space = read_xsd_bool(sheet_view, "showWhiteSpace", true);
  view.color_id = attr_u32(sheet_view, "colorId", SheetView::kDefaultColorId);
  // A scale outside 10..400 is not a usable zoom; 0 stands for absent.
  const auto read_zoom = [&](const char* name) {
    const std::int32_t raw = attr_i32(sheet_view, name, 0);
    return raw >= 10 && raw <= 400 ? static_cast<std::uint32_t>(raw) : 0U;
  };
  view.zoom_scale_normal = read_zoom("zoomScaleNormal");
  view.zoom_scale_sheet_layout_view = read_zoom("zoomScaleSheetLayoutView");
  view.zoom_scale_page_layout_view = read_zoom("zoomScalePageLayoutView");
  if (pugi::xml_attribute v = sheet_view.attribute("view"); v) {
    const std::string_view mode = v.value();
    // "normal" is the schema default; keep the model empty for it so a
    // default sheet stays byte-clean on re-emit.
    if (mode != "normal") {
      view.view_mode.assign(mode);
    }
  }
  view.top_left_cell = attr_str(sheet_view, "topLeftCell");
  for (pugi::xml_node sel = sheet_view.child("selection"); sel; sel = sel.next_sibling("selection")) {
    const std::string_view pane_name = attr_str(sel, "pane");
    view.selections.push_back(SheetSelection{pane_name.empty() ? std::string("topLeft") : std::string(pane_name),
                                             std::string(attr_str(sel, "activeCell")),
                                             std::string(attr_str(sel, "sqref")), attr_u32(sel, "activeCellId", 0U)});
  }
  for (pugi::xml_node child = sheet_view.child("pivotSelection"); child; child = child.next_sibling("pivotSelection")) {
    append_raw_xml(view.tail_xml, child);
  }
  if (pugi::xml_node ext = sheet_view.child("extLst"); ext) {
    append_raw_xml(view.tail_xml, ext);
  }
  pugi::xml_node pane = sheet_view.child("pane");
  if (pane) {
    // `ST_PaneState` has three values; two of them freeze. `frozenSplit`
    // is a frozen pane that also remembers a movable split position, so
    // its xSplit/ySplit bind the frozen extent exactly like `frozen`.
    const std::string_view state = attr_str(pane, "state");
    if (state == "frozen" || state == "frozenSplit") {
      const std::int32_t y = attr_i32(pane, "ySplit", 0);
      const std::int32_t x = attr_i32(pane, "xSplit", 0);
      if (y > 0) {
        view.freeze_rows = static_cast<std::uint32_t>(y);
      }
      if (x > 0) {
        view.freeze_cols = static_cast<std::uint32_t>(x);
      }
      view.frozen_split = state == "frozenSplit";
    } else if (state.empty() || state == "split") {
      const double y = attr_f64(pane, "ySplit", 0.0);
      const double x = attr_f64(pane, "xSplit", 0.0);
      view.split_y = std::isfinite(y) && y > 0.0 ? y : 0.0;
      view.split_x = std::isfinite(x) && x > 0.0 ? x : 0.0;
    }
    if (view.divides_rows() || view.divides_cols()) {
      view.pane_top_left_cell = attr_str(pane, "topLeftCell");
      // An absent `activePane` is the schema default, `topLeft`.
      const std::string_view active = attr_str(pane, "activePane");
      view.active_pane = active.empty() ? std::string("topLeft") : std::string(active);
    } else {
      view.frozen_split = false;
    }
  }
}

bool sheet_view_element_is_default(const SheetView& view) {
  return view.zoom_scale == SheetView::kDefaultZoomScale && !view.divides_rows() && !view.divides_cols() &&
         view.show_grid_lines && view.show_row_col_headers && view.show_zeros && !view.right_to_left &&
         !view.tab_selected && view.view_mode.empty() && !view.window_protection && !view.show_formulas &&
         view.show_ruler && view.show_outline_symbols && view.default_grid_color && view.show_white_space &&
         view.color_id == SheetView::kDefaultColorId && view.zoom_scale_normal == 0U &&
         view.zoom_scale_sheet_layout_view == 0U && view.zoom_scale_page_layout_view == 0U && view.tail_xml.empty() &&
         view.top_left_cell.empty() && view.selections.empty();
}

std::string build_sheet_view_element_xml(const SheetView& view, std::uint32_t workbook_view_id) {
  std::string children;
  if (view.divides_rows() || view.divides_cols()) {
    children.append("<pane");
    if (view.is_frozen()) {
      if (view.freeze_cols != 0U) {
        append_xml_attr_uint(children, "xSplit", view.freeze_cols);
      }
      if (view.freeze_rows != 0U) {
        append_xml_attr_uint(children, "ySplit", view.freeze_rows);
      }
    } else {
      if (view.split_x > 0.0) {
        append_xml_attr_number(children, "xSplit", view.split_x);
      }
      if (view.split_y > 0.0) {
        append_xml_attr_number(children, "ySplit", view.split_y);
      }
    }
    std::uint32_t pane_row = 0;
    std::uint32_t pane_col = 0;
    pane_origin(view, &pane_row, &pane_col);
    append_xml_attr(children, "topLeftCell", a1::encode_a1(pane_row, pane_col));
    // `topLeft` is the schema default, which Excel leaves unspelled.
    if (view.effective_active_pane() != "topLeft") {
      append_xml_attr(children, "activePane", view.effective_active_pane());
    }
    if (view.is_frozen()) {
      append_xml_attr(children, "state", view.frozen_split ? "frozenSplit" : "frozen");
    }
    children.append("/>");
  }
  // Excel omits an A1 active cell, an A1 sqref, and a top-left selection that is both.
  for (const ResolvedSelection& selection : resolve_view_selections(view)) {
    const bool a1_active = selection.active_row == 0U && selection.active_col == 0U;
    const bool a1_ranges = selection.ranges.size() == 1U && selection.ranges.front() == MergeRange{};
    if (selection.pane == "topLeft" && a1_active && a1_ranges) {
      continue;
    }
    children.append("<selection");
    if (selection.pane != "topLeft") {
      append_xml_attr(children, "pane", selection.pane);
    }
    if (!a1_active) {
      append_xml_attr(children, "activeCell", a1::encode_a1(selection.active_row, selection.active_col));
    }
    if (selection.active_cell_id != 0U) {
      append_xml_attr_uint(children, "activeCellId", selection.active_cell_id);
    }
    if (!a1_ranges) {
      append_xml_attr(children, "sqref", format_sqref_text(selection.ranges));
    }
    children.append("/>");
  }
  children.append(view.tail_xml);
  std::string out;
  out.reserve(192);
  out.append("<sheetView");
  // Display attributes precede workbookViewId in Excel's own emission
  // order. Emit each only when it differs from its schema default so a
  // near-default sheet stays compact.
  if (view.window_protection) {
    emit_xsd_bool_attr(out, "windowProtection", true);
  }
  if (view.show_formulas) {
    emit_xsd_bool_attr(out, "showFormulas", true);
  }
  if (!view.show_grid_lines) {
    emit_xsd_bool_attr(out, "showGridLines", false);
  }
  if (!view.show_row_col_headers) {
    emit_xsd_bool_attr(out, "showRowColHeaders", false);
  }
  if (!view.show_zeros) {
    emit_xsd_bool_attr(out, "showZeros", false);
  }
  if (view.right_to_left) {
    emit_xsd_bool_attr(out, "rightToLeft", true);
  }
  if (view.tab_selected) {
    emit_xsd_bool_attr(out, "tabSelected", true);
  }
  if (!view.show_ruler) {
    emit_xsd_bool_attr(out, "showRuler", false);
  }
  if (!view.show_outline_symbols) {
    emit_xsd_bool_attr(out, "showOutlineSymbols", false);
  }
  if (!view.default_grid_color) {
    emit_xsd_bool_attr(out, "defaultGridColor", false);
  }
  if (!view.show_white_space) {
    emit_xsd_bool_attr(out, "showWhiteSpace", false);
  }
  if (!view.view_mode.empty()) {
    append_xml_attr(out, "view", view.view_mode);
  }
  if (!view.top_left_cell.empty()) {
    append_xml_attr(out, "topLeftCell", view.top_left_cell);
  }
  if (view.color_id != SheetView::kDefaultColorId) {
    append_xml_attr_uint(out, "colorId", view.color_id);
  }
  append_xml_attr_uint(out, "workbookViewId", workbook_view_id);
  if (view.zoom_scale != SheetView::kDefaultZoomScale) {
    append_xml_attr_uint(out, "zoomScale", view.zoom_scale);
  }
  if (view.zoom_scale_normal != 0U) {
    append_xml_attr_uint(out, "zoomScaleNormal", view.zoom_scale_normal);
  }
  if (view.zoom_scale_sheet_layout_view != 0U) {
    append_xml_attr_uint(out, "zoomScaleSheetLayoutView", view.zoom_scale_sheet_layout_view);
  }
  if (view.zoom_scale_page_layout_view != 0U) {
    append_xml_attr_uint(out, "zoomScalePageLayoutView", view.zoom_scale_page_layout_view);
  }
  if (children.empty()) {
    out.append("/>");
  } else {
    out.push_back('>');
    out.append(children);
    out.append("</sheetView>");
  }
  return out;
}

}  // namespace io
}  // namespace formulon
