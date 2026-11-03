//
// C ABI - sheet layout (columns, rows, range styles, format defaults).
//
// Bridges `Sheet::layout()` over the stable C ABI.
// The `set_*` mutators upsert into the underlying `SheetLayout` vectors
// via small private helpers so concurrent column-span edits stay
// internally consistent (e.g. setting the width on a span that
// overlaps an existing entry replaces only the overlapping portion).

#include <algorithm>
#include <cstdint>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "column_width.h"
#include "print/sheet_geometry.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "style_region.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::drop_stateless_columns;
using formulon::overlay_column_span;
using formulon::c_api::parts::check_cell_xf_index;
using formulon::c_api::parts::check_column_span;
using formulon::c_api::parts::check_finite_non_negative;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_row_span;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::check_sheet_row;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::copy_owned_bytes;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::kMaxOutlineLevel;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

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
  out->collapsed = cols[idx].collapsed ? 1 : 0;
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
  out->collapsed = rows[idx].collapsed ? 1 : 0;
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
  if (auto rc = check_index(level, kMaxOutlineLevel + 1U, "fm_sheet_set_column_outline", "level"); rc != 0) {
    return rc;
  }
  overlay_column_span(wb->workbook().sheet(sheet_index).mutable_layout(), first, last,
                      [level](formulon::ColumnLayout& entry) { entry.outline_level = level; });
  return 0;
}

extern "C" fm_status_t fm_sheet_set_column_collapsed(fm_workbook_t* wb, size_t sheet_index, uint32_t first,
                                                     uint32_t last, int32_t collapsed) {
  clear_last_error();
  if (auto rc = check_sheet_column_span(wb, sheet_index, first, last, "fm_sheet_set_column_collapsed"); rc != 0) {
    return rc;
  }
  formulon::SheetLayout& layout = wb->workbook().sheet(sheet_index).mutable_layout();
  overlay_column_span(layout, first, last,
                      [collapsed](formulon::ColumnLayout& entry) { entry.collapsed = (collapsed != 0); });
  if (collapsed == 0) {
    drop_stateless_columns(layout);
  }
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

namespace {

// Applies `clear` to the existing override of `row` and drops the override
// once it carries nothing else. A row without an override is left alone.
void clear_row_flag(formulon::Sheet& sheet, uint32_t row, void (*clear)(formulon::RowLayout&)) {
  formulon::SheetLayout& layout = sheet.mutable_layout();
  const auto it = formulon::row_override_lower_bound(layout, row);
  if (it != layout.row_overrides.end() && it->row == row) {
    clear(*it);
    if (formulon::should_erase_row_override(*it, sheet.format_defaults().zero_height, true)) {
      layout.row_overrides.erase(it);
    }
  }
}

// Stores `height` on `row`, as a manual override when `custom` and as an
// automatic height's cached value otherwise; every other row attribute is kept.
fm_status_t store_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row, double height, bool custom,
                             const char* fn) {
  clear_last_error();
  if (auto rc = check_sheet_row(wb, sheet_index, row, fn); rc != 0) {
    return rc;
  }
  if (auto rc = check_finite_non_negative(height, fn, "height"); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  formulon::RowLayout* entry =
      &formulon::upsert_row_override(sheet.mutable_layout(), row, sheet.format_defaults().zero_height);
  entry->height = height;
  entry->has_height = true;
  entry->custom_height = custom;
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_set_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row, double height) {
  return store_row_height(wb, sheet_index, row, height, true, "fm_sheet_set_row_height");
}

extern "C" fm_status_t fm_sheet_set_row_auto_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                    double height) {
  return store_row_height(wb, sheet_index, row, height, false, "fm_sheet_set_row_auto_height");
}

extern "C" fm_status_t fm_sheet_set_row_hidden(fm_workbook_t* wb, size_t sheet_index, uint32_t row, int32_t hidden) {
  clear_last_error();
  if (auto rc = check_sheet_row(wb, sheet_index, row, "fm_sheet_set_row_hidden"); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  const bool default_hidden = sheet.format_defaults().zero_height;
  const formulon::RowLayout* previous = formulon::find_row_override(sheet.layout(), row);
  const bool was_hidden = previous == nullptr ? default_hidden : previous->hidden;
  formulon::RowLayout* entry = &formulon::upsert_row_override(sheet.mutable_layout(), row, default_hidden);
  const bool next_hidden = (hidden != 0);
  entry->hidden = next_hidden;
  // SUBTOTAL(101..111, ...) and AGGREGATE's hidden-row-skipping option read
  // this row's hidden state without that state appearing anywhere in the
  // dep graph's cell/range edges, so the ordinary dirty propagation a
  // cell-value write triggers cannot reach them.
  if (was_hidden != next_hidden) {
    wb->workbook().mark_row_visibility_dependents_dirty();
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_outline(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint8_t level) {
  clear_last_error();
  if (auto rc = check_sheet_row(wb, sheet_index, row, "fm_sheet_set_row_outline"); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(level, kMaxOutlineLevel + 1U, "fm_sheet_set_row_outline", "level"); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  formulon::RowLayout* entry =
      &formulon::upsert_row_override(sheet.mutable_layout(), row, sheet.format_defaults().zero_height);
  entry->outline_level = level;
  return 0;
}

extern "C" fm_status_t fm_sheet_set_row_collapsed(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                                  int32_t collapsed) {
  clear_last_error();
  if (auto rc = check_sheet_row(wb, sheet_index, row, "fm_sheet_set_row_collapsed"); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  if (collapsed != 0) {
    formulon::upsert_row_override(sheet.mutable_layout(), row, sheet.format_defaults().zero_height).collapsed = true;
    return 0;
  }
  clear_row_flag(sheet, row, [](formulon::RowLayout& entry) { entry.collapsed = false; });
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
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  formulon::set_row_style_span(sheet.mutable_layout(), first, last, xf_index, sheet.format_defaults().zero_height);
  return 0;
}

extern "C" fm_status_t fm_sheet_clear_row_style(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last) {
  static constexpr const char* kFn = "fm_sheet_clear_row_style";
  clear_last_error();
  if (auto rc = check_sheet_row_span(wb, sheet_index, first, last, kFn); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  formulon::clear_row_style_span(sheet.mutable_layout(), first, last, sheet.format_defaults().zero_height);
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
  return check_sheet_rect(out->first_row, out->first_col, out->last_row, out->last_col, fn);
}

}  // namespace

namespace {

fm_status_t range_style_bases_into(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row, uint32_t first_col,
                                   uint32_t last_row, uint32_t last_col, const uint32_t* excluded_rows,
                                   const uint32_t* excluded_cols, size_t excluded_count, uint32_t* out_xfs, size_t cap,
                                   size_t* out_len, const char* fn) {
  clear_last_error();
  if (out_len == nullptr || (out_xfs == nullptr && cap > 0U) ||
      (excluded_count > 0U && (excluded_rows == nullptr || excluded_cols == nullptr))) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn);
  }
  *out_len = 0U;
  formulon::StyleRect rect;
  if (auto rc = resolve_style_rect(wb, sheet_index, first_row, first_col, last_row, last_col, &rect, fn); rc != 0) {
    return rc;
  }
  std::vector<formulon::StyleCell> excluded(excluded_count);
  for (size_t i = 0; i < excluded_count; ++i) {
    excluded[i] = {excluded_rows[i], excluded_cols[i]};
  }
  auto bases = wb->workbook().range_style_bases(sheet_index, rect, std::move(excluded));
  if (!bases) {
    return set_last_error(bases.error());
  }
  *out_len = bases.value().size();
  std::copy_n(bases.value().begin(), std::min(cap, bases.value().size()), out_xfs);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_range_style_bases(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                                  uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                                  uint32_t* out_xfs, size_t cap, size_t* out_len) {
  return range_style_bases_into(wb, sheet_index, first_row, first_col, last_row, last_col, nullptr, nullptr, 0U,
                                out_xfs, cap, out_len, "fm_sheet_range_style_bases");
}

extern "C" fm_status_t fm_sheet_range_style_bases_excluding(const fm_workbook_t* wb, size_t sheet_index,
                                                            uint32_t first_row, uint32_t first_col, uint32_t last_row,
                                                            uint32_t last_col, const uint32_t* excluded_rows,
                                                            const uint32_t* excluded_cols, size_t excluded_count,
                                                            uint32_t* out_xfs, size_t cap, size_t* out_len) {
  return range_style_bases_into(wb, sheet_index, first_row, first_col, last_row, last_col, excluded_rows, excluded_cols,
                                excluded_count, out_xfs, cap, out_len, "fm_sheet_range_style_bases_excluding");
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
  return finish_status(wb->workbook().remap_range_style(sheet_index, rect, std::move(map)));
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
  return finish_status(wb->workbook().restore_range_style(sheet_index, bytes, len));
}

extern "C" fm_status_t fm_sheet_clear_row_height(fm_workbook_t* wb, size_t sheet_index, uint32_t row) {
  clear_last_error();
  if (auto rc = check_sheet_row(wb, sheet_index, row, "fm_sheet_clear_row_height"); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  clear_row_flag(sheet, row, [](formulon::RowLayout& entry) {
    entry.height = 0.0;
    entry.has_height = false;
    entry.custom_height = false;
  });
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
  out->custom_height = defaults.custom_height ? 1 : 0;
  out->zero_height = defaults.zero_height ? 1 : 0;
  const formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  out->effective_default_col_width = formulon::default_column_width_chars(
      sheet, formulon::resolve_column_width_model(wb->workbook().styles(), formulon::GeometryMode::kDisplay));
  out->effective_default_row_height = defaults.zero_height ? 0.0 : formulon::print::default_row_height_pt(sheet);
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
  target.custom_height = defaults->custom_height != 0;
  // Hidden-row state (SUBTOTAL / AGGREGATE) depends on the default.
  const bool zero_height = defaults->zero_height != 0;
  const bool zero_height_changed = target.zero_height != zero_height;
  target.zero_height = zero_height;
  if (zero_height_changed) {
    wb->workbook().mark_row_visibility_dependents_dirty();
  }
  return 0;
}
