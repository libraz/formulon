//
// C ABI - sheet view / layout (columns, rows, view, zoom, freeze).
//
// Bridges `Sheet::view()` / `Sheet::layout()` over the stable C ABI.
// The `set_*` mutators upsert into the underlying `SheetLayout` vectors
// via small private helpers so concurrent column-span edits stay
// internally consistent (e.g. setting the width on a span that
// overlaps an existing entry replaces only the overlapping portion).

#include <algorithm>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/xml_fragment.h"
#include "io/auto_filter_xml.h"
#include "print/sheet_geometry.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "style_region.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::drop_stateless_columns;
using formulon::overlay_column_span;
using formulon::row_override_is_empty;
using formulon::c_api::parts::check_column_span;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_finite_non_negative;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_row_index;
using formulon::c_api::parts::check_row_span;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::copy_owned_bytes;
using formulon::c_api::parts::FragmentValidation;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::validate_single_element_fragment;

namespace {

fm_status_t check_sheet_column_span(const fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                    const char* fn) {
  if (auto rc = check_sheet_index(wb, sheet_index, fn); rc != 0) {
    return rc;
  }
  return check_column_span(first, last, fn);
}

fm_status_t check_sheet_row_span(const fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                 const char* fn) {
  if (auto rc = check_sheet_index(wb, sheet_index, fn); rc != 0) {
    return rc;
  }
  return check_row_span(first, last, fn);
}

fm_status_t check_cell_xf_index(const fm_workbook_t* wb, std::uint32_t xf_index, const char* fn) {
  const std::size_t count = wb->workbook().styles().cell_xfs.size();
  if (static_cast<std::size_t>(xf_index) < count) {
    return 0;
  }
  const std::string message = std::string(fn) + ": xf_index out of range";
  return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, message.c_str(),
                           "xf_index=" + std::to_string(xf_index) + " cell_xfs_count=" + std::to_string(count));
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_column_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_column_count: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_column_count"); rc != 0) {
    return rc;
  }
  *out_count = wb->workbook().sheet(sheet_index).layout().columns.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_column(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                           fm_column_layout_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_column: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_column"); rc != 0) {
    return rc;
  }
  const auto& cols = wb->workbook().sheet(sheet_index).layout().columns;
  if (auto rc = check_index(idx, cols.size(), "fm_sheet_get_column", "idx"); rc != 0) {
    return rc;
  }
  *out = fm_column_layout_t{};
  out->first = cols[idx].first;
  out->last = cols[idx].last;
  out->width = cols[idx].width;
  out->hidden = cols[idx].hidden ? 1 : 0;
  out->outline_level = cols[idx].outline_level;
  out->has_width = formulon::HasExplicitColumnWidth(cols[idx]) ? 1 : 0;
  out->has_style = cols[idx].has_style ? 1 : 0;
  out->style_xf = cols[idx].style_xf;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_row_override_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_row_override_count: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_row_override_count"); rc != 0) {
    return rc;
  }
  *out_count = wb->workbook().sheet(sheet_index).layout().row_overrides.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_row_override(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                 fm_row_layout_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_row_override: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_row_override"); rc != 0) {
    return rc;
  }
  const auto& rows = wb->workbook().sheet(sheet_index).layout().row_overrides;
  if (auto rc = check_index(idx, rows.size(), "fm_sheet_get_row_override", "idx"); rc != 0) {
    return rc;
  }
  *out = fm_row_layout_t{};
  out->row = rows[idx].row;
  out->height = rows[idx].height;
  out->hidden = rows[idx].hidden ? 1 : 0;
  out->outline_level = rows[idx].outline_level;
  out->has_style = rows[idx].has_style ? 1 : 0;
  out->style_xf = rows[idx].style_xf;
  out->has_height = rows[idx].has_height ? 1 : 0;
  out->custom_height = rows[idx].custom_height ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                 double width) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last, "fm_sheet_set_column_width"); rc != 0) {
    return rc;
  }
  if (auto rc = check_finite_non_negative(width, "fm_sheet_set_column_width", "width"); rc != 0) {
    return rc;
  }
  overlay_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last,
                      [width](formulon::ColumnLayout& entry) {
                        entry.width = width;
                        entry.has_width = true;
                      });
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_column_width(fm_workbook_t* wb, size_t sheet_index, uint32_t first,
                                                   uint32_t last) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last, "fm_sheet_clear_column_width"); rc != 0) {
    return rc;
  }
  formulon::SheetLayout& layout = wb->workbook().sheet(sheet_index).mutable_layout();
  overlay_column_span(layout, first, last, [](formulon::ColumnLayout& entry) {
    entry.width = 0.0;
    entry.has_width = false;
  });
  drop_stateless_columns(layout);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                  int32_t hidden) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last, "fm_sheet_set_column_hidden"); rc != 0) {
    return rc;
  }
  overlay_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last,
                      [hidden](formulon::ColumnLayout& entry) { entry.hidden = (hidden != 0); });
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                   uint8_t level) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last, "fm_sheet_set_column_outline"); rc != 0) {
    return rc;
  }
  overlay_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last,
                      [level](formulon::ColumnLayout& entry) { entry.outline_level = level; });
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_style(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                                 uint32_t xf_index) {
  static constexpr const char* kFn = "fm_sheet_set_column_style";
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last, kFn); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_xf_index(wb, xf_index, kFn); rc != 0) {
    return rc;
  }
  formulon::set_column_style_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last, true, xf_index);
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_column_style(fm_workbook_t* wb, size_t sheet_index, uint32_t first,
                                                   uint32_t last) {
  static constexpr const char* kFn = "fm_sheet_clear_column_style";
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last, kFn); rc != 0) {
    return rc;
  }
  formulon::set_column_style_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last, false, 0U);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row, double height) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_row_height"); rc != 0) {
    return rc;
  }
  if (auto rc = check_row_index(row, "fm_sheet_set_row_height"); rc != 0) {
    return rc;
  }
  if (auto rc = check_finite_non_negative(height, "fm_sheet_set_row_height", "height"); rc != 0) {
    return rc;
  }
  formulon::RowLayout* entry = &formulon::upsert_row_override(wb->workbook().sheet(sheet_index).mutable_layout(), row);
  entry->height = height;
  entry->has_height = true;
  entry->custom_height = true;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t hidden) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_row_hidden"); rc != 0) {
    return rc;
  }
  if (auto rc = check_row_index(row, "fm_sheet_set_row_hidden"); rc != 0) {
    return rc;
  }
  formulon::RowLayout* entry = &formulon::upsert_row_override(wb->workbook().sheet(sheet_index).mutable_layout(), row);
  const bool next_hidden = (hidden != 0);
  if (entry->hidden == next_hidden) {
    return 0;
  }
  entry->hidden = next_hidden;
  // SUBTOTAL(101..111, ...) and AGGREGATE's hidden-row-skipping option read
  // this row's hidden state without that state appearing anywhere in the
  // dep graph's cell/range edges, so the ordinary dirty propagation a
  // cell-value write triggers cannot reach them.
  wb->workbook().mark_row_visibility_dependents_dirty();
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint8_t level) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_row_outline"); rc != 0) {
    return rc;
  }
  if (auto rc = check_row_index(row, "fm_sheet_set_row_outline"); rc != 0) {
    return rc;
  }
  formulon::RowLayout* entry = &formulon::upsert_row_override(wb->workbook().sheet(sheet_index).mutable_layout(), row);
  entry->outline_level = level;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_style(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last,
                                              uint32_t xf_index) {
  static constexpr const char* kFn = "fm_sheet_set_row_style";
  clear_last_error();
  if (auto rc = check_sheet_row_span(wb, sheet_index, first, last, kFn); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_xf_index(wb, xf_index, kFn); rc != 0) {
    return rc;
  }
  formulon::set_row_style_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last, xf_index);
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_row_style(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last) {
  static constexpr const char* kFn = "fm_sheet_clear_row_style";
  clear_last_error();
  if (auto rc = check_sheet_row_span(wb, sheet_index, first, last, kFn); rc != 0) {
    return rc;
  }
  formulon::clear_row_style_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last);
  return 0;
}

namespace {

// Validates a selection rect, normalising its corners like
// `fm_sheet_set_range_xf_index`.
fm_status_t resolve_style_rect(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row, uint32_t first_col,
                               uint32_t last_row, uint32_t last_col, formulon::StyleRect* out, const char* fn) {
  if (auto rc = check_sheet_index(wb, sheet_index, fn); rc != 0) {
    return rc;
  }
  out->first_row = std::min(first_row, last_row);
  out->last_row = std::max(first_row, last_row);
  out->first_col = std::min(first_col, last_col);
  out->last_col = std::max(first_col, last_col);
  if (out->last_row >= formulon::Sheet::kMaxRows || out->last_col >= formulon::Sheet::kMaxCols) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, fn);
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_range_style_bases(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                                  uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                                  uint32_t* out_xfs, size_t cap, size_t* out_len) {
  static constexpr const char* kFn = "fm_sheet_range_style_bases";
  clear_last_error();
  if (out_len == nullptr || (out_xfs == nullptr && cap > 0U)) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn);
  }
  *out_len = 0U;
  formulon::StyleRect rect;
  if (auto rc = resolve_style_rect(wb, sheet_index, first_row, first_col, last_row, last_col, &rect, kFn); rc != 0) {
    return rc;
  }
  auto bases = wb->workbook().range_style_bases(sheet_index, rect);
  if (!bases) {
    return set_last_error(bases.error());
  }
  *out_len = bases.value().size();
  std::copy_n(bases.value().begin(), std::min(cap, bases.value().size()), out_xfs);
  return 0;
}

extern "C" fm_status_t fm_sheet_remap_range_style(fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                                  uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                                  const uint32_t* from_xfs, const uint32_t* to_xfs, size_t count) {
  static constexpr const char* kFn = "fm_sheet_remap_range_style";
  clear_last_error();
  if (count > 0U && (from_xfs == nullptr || to_xfs == nullptr)) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn);
  }
  formulon::StyleRect rect;
  if (auto rc = resolve_style_rect(wb, sheet_index, first_row, first_col, last_row, last_col, &rect, kFn); rc != 0) {
    return rc;
  }
  std::vector<formulon::XfRemap> map(count);
  for (size_t i = 0; i < count; ++i) {
    map[i] = {from_xfs[i], to_xfs[i]};
  }
  auto r = wb->workbook().remap_range_style(sheet_index, rect, std::move(map));
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_snapshot_range_style(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                                     uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                                     uint8_t** out_bytes, size_t* out_len) {
  static constexpr const char* kFn = "fm_sheet_snapshot_range_style";
  clear_last_error();
  if (out_bytes != nullptr) {
    *out_bytes = nullptr;
  }
  if (out_len != nullptr) {
    *out_len = 0U;
  }
  if (out_bytes == nullptr || out_len == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn);
  }
  formulon::StyleRect rect;
  if (auto rc = resolve_style_rect(wb, sheet_index, first_row, first_col, last_row, last_col, &rect, kFn); rc != 0) {
    return rc;
  }
  auto bytes = wb->workbook().snapshot_range_style(sheet_index, rect);
  if (!bytes) {
    return set_last_error(bytes.error());
  }
  copy_owned_bytes(bytes.value(), out_bytes, out_len);
  return 0;
}

extern "C" fm_status_t fm_sheet_restore_range_style(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes,
                                                    size_t len) {
  static constexpr const char* kFn = "fm_sheet_restore_range_style";
  clear_last_error();
  if (bytes == nullptr && len > 0U) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, kFn);
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kFn); rc != 0) {
    return rc;
  }
  auto r = wb->workbook().restore_range_style(sheet_index, bytes, len);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_clear_row_height"); rc != 0) {
    return rc;
  }
  if (auto rc = check_row_index(row, "fm_sheet_clear_row_height"); rc != 0) {
    return rc;
  }
  auto& overrides = wb->workbook().sheet(sheet_index).mutable_layout().row_overrides;
  for (auto it = overrides.begin(); it != overrides.end(); ++it) {
    if (it->row != row) {
      continue;
    }
    it->height = 0.0;
    it->has_height = false;
    it->custom_height = false;
    if (row_override_is_empty(*it)) {
      overrides.erase(it);
    }
    break;
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_set_zoom(fm_workbook_t* wb, size_t sheet_index, uint32_t zoom_scale) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_zoom"); rc != 0) {
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
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_freeze"); rc != 0) {
    return rc;
  }
  formulon::SheetView& view = wb->workbook().sheet(sheet_index).mutable_view();
  view.freeze_rows = freeze_rows;
  view.freeze_cols = freeze_cols;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_tab_hidden(fm_workbook_t* wb, size_t sheet_index, int32_t hidden) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_tab_hidden"); rc != 0) {
    return rc;
  }
  formulon::SheetView& view = wb->workbook().sheet(sheet_index).mutable_view();
  if (hidden != 0) {
    // A sheet already loaded as very-hidden stays that way: this bool
    // cannot express the difference, so asking for "hidden" on a sheet
    // that is hidden says nothing new and must not quietly demote it.
    view.tab_hidden = true;
  } else {
    // Showing a sheet has to clear both bits, or the workbook would save
    // a state the caller just asked to leave.
    view.set_visibility(formulon::SheetVisibility::kVisible);
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_set_visibility(fm_workbook_t* wb, size_t sheet_index, int32_t visibility) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_visibility"); rc != 0) {
    return rc;
  }
  if (visibility < static_cast<int32_t>(FM_SHEET_VISIBLE) || visibility > static_cast<int32_t>(FM_SHEET_VERY_HIDDEN)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_set_visibility: unknown visibility", "visibility=" + std::to_string(visibility));
  }
  formulon::SheetView& view = wb->workbook().sheet(sheet_index).mutable_view();
  view.set_visibility(static_cast<formulon::SheetVisibility>(visibility));
  return 0;
}

extern "C" fm_status_t fm_sheet_get_view(const fm_workbook_t* wb, size_t sheet_index, fm_sheet_view_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_view: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_view"); rc != 0) {
    return rc;
  }
  const formulon::SheetView& v = wb->workbook().sheet(sheet_index).view();
  out->zoom_scale = v.zoom_scale;
  out->freeze_rows = v.freeze_rows;
  out->freeze_cols = v.freeze_cols;
  // `visibility()` resolves the model's two bits, so the pair handed out
  // here agrees even for the combination no writer should produce.
  const formulon::SheetVisibility state = v.visibility();
  out->tab_hidden = state != formulon::SheetVisibility::kVisible ? 1 : 0;
  out->visibility = static_cast<int32_t>(state);
  out->show_grid_lines = v.show_grid_lines ? 1 : 0;
  out->show_row_col_headers = v.show_row_col_headers ? 1 : 0;
  out->show_zeros = v.show_zeros ? 1 : 0;
  out->right_to_left = v.right_to_left ? 1 : 0;
  out->tab_selected = v.tab_selected ? 1 : 0;
  out->view_mode = v.view_mode.c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_auto_filter_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml) {
  clear_last_error();
  if (out_xml == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_auto_filter_xml: NULL out_xml");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_auto_filter_xml"); rc != 0) {
    return rc;
  }
  fm_workbook_t* mutable_wb = const_cast<fm_workbook_t*>(wb);
  mutable_wb->read_scratch.clear();
  mutable_wb->read_scratch.emplace_back(formulon::io::auto_filter_xml(wb->workbook().sheet(sheet_index).auto_filter()));
  *out_xml = mutable_wb->read_scratch.back().c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_set_auto_filter_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  clear_last_error();
  if (xml == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_set_auto_filter_xml: NULL xml");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_auto_filter_xml"); rc != 0) {
    return rc;
  }
  const std::string fragment(xml);
  if (!fragment.empty()) {
    const FragmentValidation validation = validate_single_element_fragment(fragment, "autoFilter");
    if (!validation.valid) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_sheet_set_auto_filter_xml: invalid autoFilter XML fragment", validation.context);
    }
  }
  // Routed through the workbook so `_FilterDatabase` and SUBTOTAL/AGGREGATE
  // dirty state follow the change.
  if (auto result = wb->workbook().set_sheet_auto_filter_xml(sheet_index, fragment); !result) {
    return set_last_error(result.error());
  }
  return 0;
}

namespace {

fm_status_t set_view_flag(fm_workbook_t* wb, size_t sheet_index, const char* fn, bool formulon::SheetView::*member,
                          int32_t value) {
  if (auto rc = check_sheet_index(wb, sheet_index, fn); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().*member = (value != 0);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_set_show_grid_lines(fm_workbook_t* wb, size_t sheet_index, int32_t show) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, "fm_sheet_set_show_grid_lines", &formulon::SheetView::show_grid_lines, show);
}

extern "C" fm_status_t fm_sheet_set_show_row_col_headers(fm_workbook_t* wb, size_t sheet_index, int32_t show) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, "fm_sheet_set_show_row_col_headers", &formulon::SheetView::show_row_col_headers,
                       show);
}

extern "C" fm_status_t fm_sheet_set_show_zeros(fm_workbook_t* wb, size_t sheet_index, int32_t show) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, "fm_sheet_set_show_zeros", &formulon::SheetView::show_zeros, show);
}

extern "C" fm_status_t fm_sheet_set_right_to_left(fm_workbook_t* wb, size_t sheet_index, int32_t right_to_left) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, "fm_sheet_set_right_to_left", &formulon::SheetView::right_to_left,
                       right_to_left);
}

extern "C" fm_status_t fm_sheet_set_tab_selected(fm_workbook_t* wb, size_t sheet_index, int32_t selected) {
  clear_last_error();
  return set_view_flag(wb, sheet_index, "fm_sheet_set_tab_selected", &formulon::SheetView::tab_selected, selected);
}

extern "C" fm_status_t fm_sheet_set_view_mode(fm_workbook_t* wb, size_t sheet_index, const char* mode) {
  clear_last_error();
  if (mode == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_set_view_mode: NULL mode");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_view_mode"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_view().view_mode = mode;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_format_defaults(const fm_workbook_t* wb, size_t sheet_index,
                                                    fm_sheet_format_defaults* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_format_defaults: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_format_defaults"); rc != 0) {
    return rc;
  }
  const formulon::SheetFormatDefaults& defaults = wb->workbook().sheet(sheet_index).format_defaults();
  *out = fm_sheet_format_defaults{};
  out->default_col_width = defaults.default_col_width;
  out->default_row_height = defaults.default_row_height;
  out->base_col_width = defaults.base_col_width;
  out->has_default_col_width = defaults.has_default_col_width ? 1 : 0;
  out->has_default_row_height = defaults.has_default_row_height ? 1 : 0;
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  out->effective_default_col_width = formulon::print::default_column_width_chars(sheet);
  out->effective_default_row_height = formulon::print::default_row_height_pt(sheet);
  return 0;
}

extern "C" fm_status_t fm_sheet_set_format_defaults(fm_workbook_t* wb, size_t sheet_index,
                                                    const fm_sheet_format_defaults* defaults) {
  clear_last_error();
  if (defaults == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_set_format_defaults: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_format_defaults"); rc != 0) {
    return rc;
  }
  constexpr const char* kApi = "fm_sheet_set_format_defaults";
  const bool has_col_width = defaults->has_default_col_width != 0;
  const bool has_row_height = defaults->has_default_row_height != 0;
  if (has_col_width) {
    if (auto rc = check_finite_non_negative(defaults->default_col_width, kApi, "default_col_width"); rc != 0) {
      return rc;
    }
  }
  if (has_row_height) {
    if (auto rc = check_finite_non_negative(defaults->default_row_height, kApi, "default_row_height"); rc != 0) {
      return rc;
    }
  }
  if (auto rc = check_finite_non_negative(defaults->base_col_width, kApi, "base_col_width"); rc != 0) {
    return rc;
  }
  formulon::SheetFormatDefaults& target = wb->workbook().sheet(sheet_index).mutable_format_defaults();
  target.default_col_width = has_col_width ? defaults->default_col_width : 0.0;
  target.default_row_height = has_row_height ? defaults->default_row_height : 0.0;
  target.base_col_width = defaults->base_col_width;
  target.has_default_col_width = has_col_width;
  target.has_default_row_height = has_row_height;
  return 0;
}

namespace {

// Validates the shared `(sheet, mode)` prefix of the geometry entry points.
fm_status_t check_geometry_args(const fm_workbook_t* wb, size_t sheet_index, int32_t mode, const char* api) {
  if (auto rc = check_sheet_index(wb, sheet_index, api); rc != 0) {
    return rc;
  }
  return check_enum_domain(mode, FM_GEOMETRY_PRINT, api, "mode");
}

formulon::print::ColumnWidthModel width_model_for(const fm_workbook_t* wb, int32_t mode) {
  return formulon::print::resolve_column_width_model(wb->workbook().styles(),
                                                     static_cast<formulon::print::GeometryMode>(mode));
}

}  // namespace

extern "C" fm_status_t fm_sheet_cell_rect_pt(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                             uint32_t first_col, uint32_t last_row, uint32_t last_col, int32_t mode,
                                             fm_rect_pt* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cell_rect_pt: NULL argument");
  }
  if (auto rc = check_geometry_args(wb, sheet_index, mode, "fm_sheet_cell_rect_pt"); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(first_row, first_col, last_row, last_col, "fm_sheet_cell_rect_pt"); rc != 0) {
    return rc;
  }
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  const formulon::print::ColumnWidthModel model = width_model_for(wb, mode);
  double width = 0.0;
  for (const double chars : formulon::print::column_widths_chars(sheet, first_col, last_col)) {
    width += formulon::print::column_chars_to_points(chars, model);
  }
  const double y = formulon::print::row_offset_pt(sheet, first_row);
  *out = fm_rect_pt{};
  out->x = formulon::print::column_offset_pt(sheet, first_col, model);
  out->y = y;
  out->width = width;
  out->height = first_row == last_row ? formulon::print::effective_row_height_pt(sheet, first_row)
                                      : formulon::print::row_offset_pt(sheet, last_row + 1U) - y;
  return 0;
}

extern "C" fm_status_t fm_sheet_column_width_pt(const fm_workbook_t* wb, size_t sheet_index, uint32_t col, int32_t mode,
                                                double* out_pt) {
  clear_last_error();
  if (out_pt == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_column_width_pt: NULL argument");
  }
  if (auto rc = check_geometry_args(wb, sheet_index, mode, "fm_sheet_column_width_pt"); rc != 0) {
    return rc;
  }
  if (auto rc = check_column_span(col, col, "fm_sheet_column_width_pt"); rc != 0) {
    return rc;
  }
  *out_pt =
      formulon::print::effective_column_width_pt(wb->workbook().sheet(sheet_index), col, width_model_for(wb, mode));
  return 0;
}

extern "C" fm_status_t fm_sheet_row_height_pt(const fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                              double* out_pt) {
  clear_last_error();
  if (out_pt == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_row_height_pt: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_row_height_pt"); rc != 0) {
    return rc;
  }
  if (auto rc = check_row_index(row, "fm_sheet_row_height_pt"); rc != 0) {
    return rc;
  }
  *out_pt = formulon::print::effective_row_height_pt(wb->workbook().sheet(sheet_index), row);
  return 0;
}

extern "C" fm_status_t fm_sheet_width_model(const fm_workbook_t* wb, size_t sheet_index, int32_t mode,
                                            fm_width_model* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_width_model: NULL argument");
  }
  if (auto rc = check_geometry_args(wb, sheet_index, mode, "fm_sheet_width_model"); rc != 0) {
    return rc;
  }
  const formulon::print::ColumnWidthModel model = width_model_for(wb, mode);
  fm_workbook_t* mutable_wb = const_cast<fm_workbook_t*>(wb);
  mutable_wb->read_scratch.clear();
  mutable_wb->read_scratch.emplace_back(model.normal_font_name);
  *out = fm_width_model{};
  out->points_per_char = model.points_per_char;
  out->padding_pt = model.padding_pt;
  out->normal_font_size = model.normal_font_size;
  out->calibrated = model.calibrated ? 1 : 0;
  out->normal_font_name = mutable_wb->read_scratch.back().c_str();
  out->platform = model.platform;
  return 0;
}

extern "C" fm_status_t fm_sheet_column_chars_to_pt(const fm_workbook_t* wb, size_t sheet_index, int32_t mode,
                                                   double chars, double* out_pt) {
  clear_last_error();
  if (out_pt == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_column_chars_to_pt: NULL argument");
  }
  if (auto rc = check_geometry_args(wb, sheet_index, mode, "fm_sheet_column_chars_to_pt"); rc != 0) {
    return rc;
  }
  if (auto rc = check_finite_non_negative(chars, "fm_sheet_column_chars_to_pt", "chars"); rc != 0) {
    return rc;
  }
  *out_pt = formulon::print::column_chars_to_points(chars, width_model_for(wb, mode));
  return 0;
}

extern "C" fm_status_t fm_sheet_column_pt_to_chars(const fm_workbook_t* wb, size_t sheet_index, int32_t mode, double pt,
                                                   double* out_chars) {
  clear_last_error();
  if (out_chars == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_column_pt_to_chars: NULL argument");
  }
  if (auto rc = check_geometry_args(wb, sheet_index, mode, "fm_sheet_column_pt_to_chars"); rc != 0) {
    return rc;
  }
  if (auto rc = check_finite_non_negative(pt, "fm_sheet_column_pt_to_chars", "pt"); rc != 0) {
    return rc;
  }
  *out_chars = formulon::print::column_points_to_chars(pt, width_model_for(wb, mode));
  return 0;
}
