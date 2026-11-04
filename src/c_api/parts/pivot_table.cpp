//
// C ABI - PivotTable mutation surface (tables, field order, data fields).
// Fields and filters live in `pivot_field.cpp` / `pivot_filter.cpp`; shared
// helpers come from `pivot_internal.h`.

#include "pivot/pivot_table.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/pivot_internal.h"
#include "pivot/field_lookup.h"
#include "pivot/pivot_locale.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/utf8_length.h"
#include "workbook.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_num_fmt_id_text;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::find_cache;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalidate_pivot_result;
using formulon::c_api::parts::pivot_agg_from_fm;
using formulon::c_api::parts::pivot_count;
using formulon::c_api::parts::pivot_show_as_from_fm;
using formulon::c_api::parts::resolve_pivot_checked;
using formulon::c_api::parts::resolve_pivot_mut;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::validate;

namespace {

// Longest grand total caption, in UTF-16 code units.
constexpr std::uint32_t kMaxGrandTotalCaptionUnits = 255U;

fm_pivot_layout_t pivot_layout_to_fm(formulon::pivot::PivotLayout layout) {
  switch (layout) {
    case formulon::pivot::PivotLayout::Tabular:
      return FM_PIVOT_LAYOUT_TABULAR;
    case formulon::pivot::PivotLayout::Outline:
      return FM_PIVOT_LAYOUT_OUTLINE;
    case formulon::pivot::PivotLayout::Compact:
    default:
      return FM_PIVOT_LAYOUT_COMPACT;
  }
}

formulon::pivot::PivotLayout pivot_layout_from_fm(std::int32_t layout) {
  switch (layout) {
    case FM_PIVOT_LAYOUT_TABULAR:
      return formulon::pivot::PivotLayout::Tabular;
    case FM_PIVOT_LAYOUT_OUTLINE:
      return formulon::pivot::PivotLayout::Outline;
    case FM_PIVOT_LAYOUT_COMPACT:
    default:
      return formulon::pivot::PivotLayout::Compact;
  }
}

}  // namespace

// ---------------------------------------------------------------------------
// PivotTable lifecycle / metadata
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_pivot_create(fm_workbook_t* wb, std::size_t sheet_index, const char* utf8_name,
                                                std::uint32_t cache_id, std::uint32_t anchor_row,
                                                std::uint32_t anchor_col, std::size_t* out_pivot_index) {
  clear_last_error();
  if (utf8_name == nullptr || out_pivot_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_create: NULL argument");
  }
  if (auto rc = check_utf8("fm_workbook_pivot_create", {utf8_name}); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_pivot_create"); rc != 0) {
    return rc;
  }
  if (find_cache(wb->workbook(), cache_id) == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_create: cache_id not found", "cache_id=" + std::to_string(cache_id));
  }
  if (!formulon::Sheet::coord_in_grid(anchor_row, anchor_col)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_create: anchor out of grid",
                             "row=" + std::to_string(anchor_row) + " col=" + std::to_string(anchor_col));
  }
  auto table = std::make_unique<formulon::pivot::PivotTable>();
  table->set_name(utf8_name);
  table->set_pivot_cache_id(cache_id);
  // Default span of 1x1 anchored at the requested cell; callers can
  // adjust via `fm_workbook_pivot_set_anchor`.
  table->set_anchor(anchor_row, anchor_col, 1U, 1U);
  // Excel stores the Values caption in the creating UI's wording and keeps it
  // when the file is opened under another locale.
  table->set_data_caption(formulon::pivot::pivot_layout_options_for(wb->workbook().excel_profile()).values_label);
  // The default style the desktop spreadsheet gives a new PivotTable.
  formulon::pivot::PivotStyleInfo style;
  style.name = "PivotStyleLight16";
  style.show_row_headers = true;
  style.show_col_headers = true;
  style.show_last_column = true;
  table->set_style_info(std::move(style));
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  sheet.add_pivot_table(std::move(table));
  *out_pivot_index = sheet.pivot_tables().size() - 1U;
  wb->workbook().mark_all_formulas_dirty();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_remove(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_pivot_remove"); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  auto& pivots = sheet.mutable_pivot_tables();
  if (auto rc = check_index(pivot_index, pivots.size(), "fm_workbook_pivot_remove", "pivot_index"); rc != 0) {
    return rc;
  }
  pivots.erase(pivots.begin() + static_cast<std::ptrdiff_t>(pivot_index));
  wb->workbook().mark_all_formulas_dirty();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_set_name(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                  const char* utf8_name) {
  clear_last_error();
  if (wb == nullptr || utf8_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_set_name: NULL argument");
  }
  if (auto rc = check_utf8("fm_workbook_pivot_set_name", {utf8_name}); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, "fm_workbook_pivot_set_name", &table);
      rc != 0) {
    return rc;
  }
  table->set_name(utf8_name);
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_set_anchor(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                    std::uint32_t anchor_row, std::uint32_t anchor_col,
                                                    std::uint32_t span_rows, std::uint32_t span_cols) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, "fm_workbook_pivot_set_anchor", &table);
      rc != 0) {
    return rc;
  }
  // A span must be at least 1x1 and its far corner must stay inside the
  // grid; reject before `set_anchor` so downstream layout never iterates
  // an out-of-grid or wrapped rectangle. The far corner is computed in 64
  // bits and saturated, so `anchor + span` cannot wrap back into the grid.
  if (span_rows == 0U || span_cols == 0U) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_set_anchor: span out of grid",
                             "span_rows=" + std::to_string(span_rows) + " span_cols=" + std::to_string(span_cols));
  }
  const auto far_edge = [](std::uint32_t anchor, std::uint32_t span) {
    return static_cast<std::uint32_t>(std::min<std::uint64_t>(static_cast<std::uint64_t>(anchor) + span - 1U,
                                                              std::numeric_limits<std::uint32_t>::max()));
  };
  if (auto rc = check_sheet_rect(anchor_row, anchor_col, far_edge(anchor_row, span_rows),
                                 far_edge(anchor_col, span_cols), "fm_workbook_pivot_set_anchor");
      rc != 0) {
    return rc;
  }
  table->set_anchor(anchor_row, anchor_col, span_rows, span_cols);
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_set_grand_totals(fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index, std::int32_t rows_enabled,
                                                          std::int32_t cols_enabled) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc =
          resolve_pivot_checked(wb, sheet_index, pivot_index, "fm_workbook_pivot_set_grand_totals", &table);
      rc != 0) {
    return rc;
  }
  table->set_grand_totals(rows_enabled != 0, cols_enabled != 0);
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_set_grand_total_caption(fm_workbook_t* wb, std::size_t sheet_index,
                                                                 std::size_t pivot_index, const char* caption) {
  clear_last_error();
  constexpr const char* kApi = "fm_workbook_pivot_set_grand_total_caption";
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, kApi, &table); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8(kApi, {caption}); rc != 0) {
    return rc;
  }
  if (caption == nullptr || caption[0] == '\0') {
    table->set_grand_total_caption(std::nullopt);
  } else {
    if (formulon::utf16_units_in(caption) > kMaxGrandTotalCaptionUnits) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               (std::string(kApi) + ": caption longer than 255 UTF-16 code units").c_str());
    }
    table->set_grand_total_caption(std::string(caption));
  }
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_set_auto_formatting(fm_workbook_t* wb, std::size_t sheet_index,
                                                             std::size_t pivot_index, std::int32_t on) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc =
          resolve_pivot_checked(wb, sheet_index, pivot_index, "fm_workbook_pivot_set_auto_formatting", &table);
      rc != 0) {
    return rc;
  }
  table->set_auto_formatting(on != 0);
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_show_details(fm_workbook_t* wb, std::size_t sheet_index,
                                                      std::size_t pivot_index, std::uint32_t row, std::uint32_t col,
                                                      std::size_t* out_sheet_index) {
  clear_last_error();
  if (wb == nullptr || out_sheet_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_show_details: NULL argument");
  }
  formulon::Expected<std::size_t, formulon::Error> created =
      wb->workbook().show_pivot_details(sheet_index, pivot_index, row, col);
  if (!created) {
    return set_last_error(created.error());
  }
  *out_sheet_index = created.value();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_set_style_info(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, const char* utf8_style_name,
                                                        std::int32_t show_row_headers, std::int32_t show_col_headers,
                                                        std::int32_t show_row_stripes, std::int32_t show_col_stripes,
                                                        std::int32_t show_last_column) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, "fm_workbook_pivot_set_style_info", &table);
      rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8("fm_workbook_pivot_set_style_info", {utf8_style_name}); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotStyleInfo info;
  if (utf8_style_name != nullptr) {
    info.name = utf8_style_name;
  }
  info.show_row_headers = show_row_headers != 0;
  info.show_col_headers = show_col_headers != 0;
  info.show_row_stripes = show_row_stripes != 0;
  info.show_col_stripes = show_col_stripes != 0;
  info.show_last_column = show_last_column != 0;
  table->set_style_info(std::move(info));
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_get_layout(const fm_workbook_t* wb, std::size_t sheet_index,
                                                    std::size_t pivot_index, fm_pivot_layout_t* out_layout) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc =
          resolve_pivot_checked(wb, {out_layout}, sheet_index, pivot_index, "fm_workbook_pivot_get_layout", &table);
      rc != 0) {
    return rc;
  }
  *out_layout = pivot_layout_to_fm(table->layout());
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_set_layout(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                    std::int32_t layout) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_set_layout: wb is NULL");
  }
  if (layout != FM_PIVOT_LAYOUT_COMPACT && layout != FM_PIVOT_LAYOUT_TABULAR && layout != FM_PIVOT_LAYOUT_OUTLINE) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_set_layout: layout out of range",
                             "layout=" + std::to_string(static_cast<int>(layout)));
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, "fm_workbook_pivot_set_layout", &table);
      rc != 0) {
    return rc;
  }
  table->set_layout(pivot_layout_from_fm(layout));
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

// PivotTable field order (row / col)
// ---------------------------------------------------------------------------

namespace {

// The axis the Values level sits on, as the layout resolves it: the rows when
// they record a position, else the columns. Meaningful only with two or more
// data fields; fewer leave no Values level at all.
bool values_on_rows(const formulon::pivot::PivotTable& table) {
  return table.row_values_position().has_value();
}

bool has_values_level(const formulon::pivot::PivotTable& table) {
  return table.data_fields().size() >= 2U;
}

// Common backing for `set_row_field_order` / `set_col_field_order`.
fm_status_t set_field_order(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                            const std::uint32_t* indices, std::size_t count, bool row, const char* fn) {
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(fn) + ": wb is NULL").c_str());
  }
  if (count > 0 && indices == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(fn) + ": indices is NULL").c_str());
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, fn);
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  const std::size_t field_count = table->fields().size();
  constexpr std::uint8_t kNotRequested = 0U;
  constexpr std::uint8_t kRequested = 1U;
  constexpr std::uint8_t kAlreadyRemoved = 2U;
  std::vector<std::uint8_t> requested_field(field_count, kNotRequested);
  // The getter exposes the axis vector's storage, so copy the request before
  // updating either order in case `indices` aliases one of them. The Values
  // level is split off: `values_slot` is its index among the fields.
  std::vector<std::uint32_t> requested_order;
  std::optional<std::size_t> values_slot;
  for (std::size_t i = 0; i < count; ++i) {
    if (indices[i] == FM_PIVOT_VALUES_FIELD) {
      if (values_slot.has_value()) {
        return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                                 (std::string(fn) + ": indices lists the Values field twice").c_str(),
                                 "i=" + std::to_string(i));
      }
      values_slot = requested_order.size();
      continue;
    }
    if (indices[i] >= field_count) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               (std::string(fn) + ": indices[i] out of range").c_str(),
                               "i=" + std::to_string(i) + " value=" + std::to_string(indices[i]));
    }
    if (requested_field[indices[i]] != kNotRequested) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               (std::string(fn) + ": indices lists a field twice").c_str(),
                               "i=" + std::to_string(i) + " value=" + std::to_string(indices[i]));
    }
    requested_field[indices[i]] = kRequested;
    requested_order.push_back(indices[i]);
  }
  const bool holds_values = has_values_level(*table) && values_on_rows(*table) == row;
  if (values_slot.has_value() && !has_values_level(*table)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(fn) + ": the Values field needs at least two data fields").c_str(),
                             "data_fields=" + std::to_string(table->data_fields().size()));
  }
  if (holds_values && !values_slot.has_value()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(fn) + ": the Values field is on this axis and must stay listed").c_str());
  }
  // The list is the axis's whole membership: a field left out comes off the
  // axis, and a listed field moves onto it from wherever it was.
  const formulon::pivot::PivotAxis axis = row ? formulon::pivot::PivotAxis::Row : formulon::pivot::PivotAxis::Col;
  std::vector<std::uint32_t>& target = row ? table->mutable_row_field_order() : table->mutable_col_field_order();
  for (const std::uint32_t field : target) {
    if (field < field_count && requested_field[field] == kNotRequested) {
      table->mutable_fields()[field].axis = formulon::pivot::PivotAxis::None;
    }
  }
  for (const std::uint32_t field : requested_order) {
    table->mutable_fields()[field].axis = axis;
  }

  // Remove each requested field's first occurrence from the opposite order,
  // preserving its stable order even if a loaded document contains duplicates.
  std::vector<std::uint32_t>& opposite = row ? table->mutable_col_field_order() : table->mutable_row_field_order();
  const std::optional<std::size_t> opposite_values_position =
      row ? table->col_values_position() : table->row_values_position();
  std::size_t removed_before_values = 0U;
  std::size_t kept = 0U;
  for (std::size_t read = 0U; read < opposite.size(); ++read) {
    const std::uint32_t field = opposite[read];
    if (field < field_count && requested_field[field] == kRequested) {
      requested_field[field] = kAlreadyRemoved;
      if (opposite_values_position.has_value() && read < *opposite_values_position) {
        ++removed_before_values;
      }
      continue;
    }
    if (kept != read) {
      opposite[kept] = field;
    }
    ++kept;
  }
  opposite.resize(kept);
  if (opposite_values_position.has_value() && removed_before_values > 0U) {
    if (row) {
      table->set_col_values_position(*opposite_values_position - removed_before_values);
    } else {
      table->set_row_values_position(*opposite_values_position - removed_before_values);
    }
  }
  target.assign(requested_order.begin(), requested_order.end());
  if (values_slot.has_value()) {
    // The Values level moves here; the other axis gives it up.
    if (row) {
      table->set_row_values_position(*values_slot);
      table->set_col_values_position(std::nullopt);
    } else {
      table->set_col_values_position(*values_slot);
      table->set_row_values_position(std::nullopt);
    }
  } else {
    // Keep a Values marker within the new order's slots.
    const std::optional<std::size_t> values_position =
        row ? table->row_values_position() : table->col_values_position();
    if (values_position.has_value() && *values_position > requested_order.size()) {
      if (row) {
        table->set_row_values_position(requested_order.size());
      } else {
        table->set_col_values_position(requested_order.size());
      }
    }
  }
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_pivot_set_row_field_order(fm_workbook_t* wb, std::size_t sheet_index,
                                                             std::size_t pivot_index, const std::uint32_t* indices,
                                                             std::size_t count) {
  clear_last_error();
  return set_field_order(wb, sheet_index, pivot_index, indices, count, true, "fm_workbook_pivot_set_row_field_order");
}

extern "C" fm_status_t fm_workbook_pivot_set_col_field_order(fm_workbook_t* wb, std::size_t sheet_index,
                                                             std::size_t pivot_index, const std::uint32_t* indices,
                                                             std::size_t count) {
  clear_last_error();
  return set_field_order(wb, sheet_index, pivot_index, indices, count, false, "fm_workbook_pivot_set_col_field_order");
}

// ---------------------------------------------------------------------------
// PivotTable data fields
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_pivot_data_field_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index, std::size_t* out_count) {
  clear_last_error();
  return pivot_count(wb, sheet_index, pivot_index, out_count, "fm_workbook_pivot_data_field_count",
                     [](const formulon::pivot::PivotTable& table) { return table.data_fields().size(); });
}

namespace {

// Materialises a `fm_pivot_data_field_spec_t` into a `PivotDataField`.
// The spec is already validated by `validate(*spec, api)` and
// `check_num_fmt_id_text` at the entry point. `field_index` and a set
// `show_as_base_field` must index `table.fields()`, else the binding error
// is set and the status returned. A NULL or empty name is derived from the
// workbook profile.
fm_status_t fill_data_field(const formulon::Workbook& workbook, const formulon::pivot::PivotTable& table,
                            const fm_pivot_data_field_spec_t& spec, const char* api,
                            formulon::pivot::PivotDataField* out) {
  if (auto rc = check_index(spec.field_index, table.fields().size(), api, "field_index"); rc != 0) {
    return rc;
  }
  // `-1` is the only unset spelling; any other negative wraps past every field.
  if (spec.show_as_base_field != -1) {
    if (auto rc = check_index(static_cast<std::uint32_t>(spec.show_as_base_field), table.fields().size(), api,
                              "show_as_base_field");
        rc != 0) {
      return rc;
    }
  }
  out->aggregation = pivot_agg_from_fm(spec.aggregation);
  if (spec.name != nullptr && spec.name[0] != '\0') {
    out->name = spec.name;
  } else {
    out->name = formulon::pivot::data_field_display_name(
        out->aggregation, formulon::pivot::pivot_field_display_name(table.fields()[spec.field_index]),
        workbook.excel_profile());
  }
  out->field_index = spec.field_index;
  out->number_format = spec.number_format != nullptr ? spec.number_format : "";
  out->show_as = pivot_show_as_from_fm(spec.show_as);
  if (spec.show_as_base_field >= 0) {
    out->show_as_base_field = static_cast<std::uint32_t>(spec.show_as_base_field);
  } else {
    out->show_as_base_field.reset();
  }
  if (spec.show_as_base_item >= 0) {
    out->show_as_base_item = static_cast<std::uint32_t>(spec.show_as_base_item);
  } else {
    out->show_as_base_item.reset();
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_pivot_data_field_add(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, const fm_pivot_data_field_spec_t* spec,
                                                        std::size_t* out_idx) {
  clear_last_error();
  if (wb == nullptr || spec == nullptr || out_idx == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_data_field_add: NULL argument");
  }
  if (auto rc = validate(*spec, "fm_workbook_pivot_data_field_add"); rc != 0) {
    return rc;
  }
  if (auto rc = check_num_fmt_id_text(wb->workbook(), spec->number_format, "fm_workbook_pivot_data_field_add");
      rc != 0) {
    return rc;
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_data_field_add");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  formulon::pivot::PivotDataField df;
  if (auto rc = fill_data_field(wb->workbook(), *table, *spec, "fm_workbook_pivot_data_field_add", &df); rc != 0) {
    return rc;
  }
  table->mutable_data_fields().push_back(std::move(df));
  *out_idx = table->data_fields().size() - 1U;
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_data_field_clear(fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc =
          resolve_pivot_checked(wb, sheet_index, pivot_index, "fm_workbook_pivot_data_field_clear", &table);
      rc != 0) {
    return rc;
  }
  table->mutable_data_fields().clear();
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_data_field_set(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::size_t data_field_idx,
                                                        const fm_pivot_data_field_spec_t* spec) {
  clear_last_error();
  if (wb == nullptr || spec == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_data_field_set: NULL argument");
  }
  if (auto rc = validate(*spec, "fm_workbook_pivot_data_field_set"); rc != 0) {
    return rc;
  }
  if (auto rc = check_num_fmt_id_text(wb->workbook(), spec->number_format, "fm_workbook_pivot_data_field_set");
      rc != 0) {
    return rc;
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_data_field_set");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  if (auto rc = check_index(data_field_idx, table->data_fields().size(), "fm_workbook_pivot_data_field_set",
                            "data_field_idx");
      rc != 0) {
    return rc;
  }
  formulon::pivot::PivotDataField df;
  if (auto rc = fill_data_field(wb->workbook(), *table, *spec, "fm_workbook_pivot_data_field_set", &df); rc != 0) {
    return rc;
  }
  table->mutable_data_fields()[data_field_idx] = std::move(df);
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

// ---------------------------------------------------------------------------
// PivotTable definition read-back
// ---------------------------------------------------------------------------

namespace {

fm_status_t get_field_order(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                            const std::uint32_t** out_indices, std::size_t* out_count, bool row, const char* fn) {
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out_indices, out_count}, sheet_index, pivot_index, fn, &table); rc != 0) {
    return rc;
  }
  const std::vector<std::uint32_t>& order = row ? table->row_field_order() : table->col_field_order();
  // The Values level is listed on the axis that holds it; the columns hold it,
  // at the end, when the table records no position.
  std::optional<std::size_t> values_slot;
  if (has_values_level(*table) && values_on_rows(*table) == row) {
    values_slot = row ? table->row_values_position() : table->col_values_position().value_or(order.size());
  }
  if (!values_slot.has_value()) {
    *out_indices = order.data();
    *out_count = order.size();
    return 0;
  }
  // The listing is not the model's own vector, so it lives in the read scratch.
  // A string shorter than its inline buffer would not be 4-byte aligned, hence
  // the padding that forces a heap block.
  const std::size_t listed = order.size() + 1U;
  std::string& buffer =
      fresh_read_scratch(wb).emplace_back(std::max<std::size_t>(listed * sizeof(std::uint32_t), 64U), '\0');
  auto* listing = reinterpret_cast<std::uint32_t*>(buffer.data());
  const std::size_t slot = std::min(*values_slot, order.size());
  std::copy(order.begin(), order.begin() + static_cast<std::ptrdiff_t>(slot), listing);
  listing[slot] = FM_PIVOT_VALUES_FIELD;
  std::copy(order.begin() + static_cast<std::ptrdiff_t>(slot), order.end(), listing + slot + 1U);
  *out_indices = listing;
  *out_count = listed;
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_pivot_get_cache_id(const fm_workbook_t* wb, std::size_t sheet_index,
                                                      std::size_t pivot_index, std::uint32_t* out_cache_id) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc =
          resolve_pivot_checked(wb, {out_cache_id}, sheet_index, pivot_index, "fm_workbook_pivot_get_cache_id", &table);
      rc != 0) {
    return rc;
  }
  *out_cache_id = table->pivot_cache_id();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_get_name(const fm_workbook_t* wb, std::size_t sheet_index,
                                                  std::size_t pivot_index, const char** out_name) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out_name}, sheet_index, pivot_index, "fm_workbook_pivot_get_name", &table);
      rc != 0) {
    return rc;
  }
  *out_name = table->name().c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_get_grand_totals(const fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index, std::int32_t* out_rows,
                                                          std::int32_t* out_cols) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out_rows, out_cols}, sheet_index, pivot_index,
                                      "fm_workbook_pivot_get_grand_totals", &table);
      rc != 0) {
    return rc;
  }
  *out_rows = table->grand_totals_rows() ? 1 : 0;
  *out_cols = table->grand_totals_cols() ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_get_grand_total_caption(const fm_workbook_t* wb, std::size_t sheet_index,
                                                                 std::size_t pivot_index, const char** out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, "fm_workbook_pivot_get_grand_total_caption",
                                      &table);
      rc != 0) {
    return rc;
  }
  const std::optional<std::string>& caption = table->grand_total_caption();
  *out = caption.has_value() ? caption->c_str() : nullptr;
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_get_auto_formatting(const fm_workbook_t* wb, std::size_t sheet_index,
                                                             std::size_t pivot_index, std::int32_t* out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc =
          resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, "fm_workbook_pivot_get_auto_formatting", &table);
      rc != 0) {
    return rc;
  }
  *out = table->auto_formatting() ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_get_style_info(const fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, fm_pivot_style_info_t* out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, "fm_workbook_pivot_get_style_info", &table);
      rc != 0) {
    return rc;
  }
  fm_pivot_style_info_t result{};
  result.name = "";
  if (const auto& info = table->style_info(); info.has_value()) {
    result.name = info->name.c_str();
    result.present = 1;
    result.show_row_headers = info->show_row_headers ? 1 : 0;
    result.show_col_headers = info->show_col_headers ? 1 : 0;
    result.show_row_stripes = info->show_row_stripes ? 1 : 0;
    result.show_col_stripes = info->show_col_stripes ? 1 : 0;
    result.show_last_column = info->show_last_column ? 1 : 0;
  }
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_data_field_at(const fm_workbook_t* wb, std::size_t sheet_index,
                                                       std::size_t pivot_index, std::size_t data_field_idx,
                                                       fm_pivot_data_field_spec_t* out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, "fm_workbook_pivot_data_field_at", &table);
      rc != 0) {
    return rc;
  }
  if (auto rc =
          check_index(data_field_idx, table->data_fields().size(), "fm_workbook_pivot_data_field_at", "data_field_idx");
      rc != 0) {
    return rc;
  }
  const formulon::pivot::PivotDataField& df = table->data_fields()[data_field_idx];
  fm_pivot_data_field_spec_t result{};
  result.name = df.name.c_str();
  result.field_index = df.field_index;
  result.aggregation = static_cast<fm_pivot_aggregation_t>(df.aggregation);
  result.number_format = df.number_format.c_str();
  result.show_as = static_cast<fm_pivot_show_as_t>(df.show_as);
  result.show_as_base_field =
      df.show_as_base_field.has_value() ? static_cast<std::int32_t>(*df.show_as_base_field) : -1;
  result.show_as_base_item = df.show_as_base_item.has_value() ? static_cast<std::int32_t>(*df.show_as_base_item) : -1;
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_get_row_field_order(const fm_workbook_t* wb, std::size_t sheet_index,
                                                             std::size_t pivot_index, const std::uint32_t** out_indices,
                                                             std::size_t* out_count) {
  clear_last_error();
  return get_field_order(wb, sheet_index, pivot_index, out_indices, out_count, true,
                         "fm_workbook_pivot_get_row_field_order");
}

extern "C" fm_status_t fm_workbook_pivot_get_col_field_order(const fm_workbook_t* wb, std::size_t sheet_index,
                                                             std::size_t pivot_index, const std::uint32_t** out_indices,
                                                             std::size_t* out_count) {
  clear_last_error();
  return get_field_order(wb, sheet_index, pivot_index, out_indices, out_count, false,
                         "fm_workbook_pivot_get_col_field_order");
}
