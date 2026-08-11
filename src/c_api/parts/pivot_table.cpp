//
// C ABI - PivotTable mutation surface (tables, fields, items,
// data fields, filters). Shares helpers with `pivot_cache.cpp` via
// `pivot_internal.h`.

#include "pivot/pivot_table.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/pivot_internal.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::find_cache;
using formulon::c_api::parts::invalidate_pivot_result;
using formulon::c_api::parts::pivot_agg_from_fm;
using formulon::c_api::parts::pivot_axis_from_fm;
using formulon::c_api::parts::pivot_calendar_from_fm;
using formulon::c_api::parts::pivot_date_grouping_from_fm;
using formulon::c_api::parts::pivot_filter_type_from_fm;
using formulon::c_api::parts::pivot_show_as_from_fm;
using formulon::c_api::parts::pivot_subtotal_from_fm;
using formulon::c_api::parts::resolve_pivot;
using formulon::c_api::parts::resolve_pivot_mut;
using formulon::c_api::parts::set_binding_error;

namespace {

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

formulon::pivot::PivotLayout pivot_layout_from_fm(fm_pivot_layout_t layout) {
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
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  sheet.add_pivot_table(std::move(table));
  *out_pivot_index = sheet.pivot_tables().size() - 1U;
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_remove(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_pivot_remove"); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  auto& pivots = sheet.mutable_pivot_tables();
  if (pivot_index >= pivots.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_remove: pivot_index out of range",
                             "pivot_index=" + std::to_string(pivot_index));
  }
  pivots.erase(pivots.begin() + static_cast<std::ptrdiff_t>(pivot_index));
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_set_name(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                  const char* utf8_name) {
  clear_last_error();
  if (wb == nullptr || utf8_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_set_name: NULL argument");
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_set_name");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  table->set_name(utf8_name);
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_set_anchor(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                    std::uint32_t anchor_row, std::uint32_t anchor_col,
                                                    std::uint32_t span_rows, std::uint32_t span_cols) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_set_anchor: wb is NULL");
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_set_anchor");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  // A span must be at least 1x1 and its far corner must stay inside the
  // grid; reject before `set_anchor` so downstream layout never iterates
  // an out-of-grid or wrapped rectangle. 64-bit arithmetic avoids the
  // `anchor + span` wrap the audit flagged.
  if (span_rows == 0U || span_cols == 0U ||
      static_cast<std::uint64_t>(anchor_row) + span_rows > formulon::Sheet::kMaxRows ||
      static_cast<std::uint64_t>(anchor_col) + span_cols > formulon::Sheet::kMaxCols) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_set_anchor: span out of grid",
                             "row=" + std::to_string(anchor_row) + " col=" + std::to_string(anchor_col) +
                                 " span_rows=" + std::to_string(span_rows) + " span_cols=" + std::to_string(span_cols));
  }
  table->set_anchor(anchor_row, anchor_col, span_rows, span_cols);
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_set_grand_totals(fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index, std::int32_t rows_enabled,
                                                          std::int32_t cols_enabled) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_set_grand_totals: wb is NULL");
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_set_grand_totals");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  table->set_grand_totals(rows_enabled != 0, cols_enabled != 0);
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_get_layout(const fm_workbook_t* wb, std::size_t sheet_index,
                                                    std::size_t pivot_index, fm_pivot_layout_t* out_layout) {
  clear_last_error();
  if (wb == nullptr || out_layout == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_get_layout: NULL argument");
  }
  const auto* table = resolve_pivot(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_get_layout");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  *out_layout = pivot_layout_to_fm(table->layout());
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_set_layout(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                    fm_pivot_layout_t layout) {
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
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_set_layout");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  table->set_layout(pivot_layout_from_fm(layout));
  invalidate_pivot_result(*table);
  return 0;
}

// ---------------------------------------------------------------------------
// PivotTable fields
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_pivot_field_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::size_t pivot_index, std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_count: NULL argument");
  }
  const auto* table = resolve_pivot(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_field_count");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  *out_count = table->fields().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_add(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                   const fm_pivot_field_spec_t* spec, std::size_t* out_field_idx) {
  clear_last_error();
  if (wb == nullptr || spec == nullptr || out_field_idx == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_add: NULL argument");
  }
  if (spec->source_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_add: spec->source_name is NULL");
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_field_add");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  formulon::pivot::PivotField field;
  field.source_name = spec->source_name;
  field.custom_name = spec->custom_name != nullptr ? spec->custom_name : "";
  field.axis = pivot_axis_from_fm(spec->axis);
  field.subtotal_top = spec->subtotal_top != 0;
  field.number_format = spec->number_format != nullptr ? spec->number_format : "";
  table->mutable_fields().push_back(std::move(field));
  *out_field_idx = table->fields().size() - 1U;
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_clear(fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::size_t pivot_index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_clear: wb is NULL");
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_field_clear");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  table->mutable_fields().clear();
  invalidate_pivot_result(*table);
  return 0;
}

namespace {

// Look up a pivot field for the `field_set_*` / `field_add_*` family.
// Sets the binding error and returns NULL on miss.
formulon::pivot::PivotField* lookup_pivot_field_mut(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                    std::size_t field_idx, const char* fn,
                                                    formulon::pivot::PivotTable** out_table) {
  if (wb == nullptr) {
    set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, (std::string(fn) + ": wb is NULL").c_str());
    return nullptr;
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, fn);
  if (table == nullptr) {
    return nullptr;
  }
  if (field_idx >= table->fields().size()) {
    set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                      (std::string(fn) + ": field_idx out of range").c_str(), "field_idx=" + std::to_string(field_idx));
    return nullptr;
  }
  if (out_table != nullptr) {
    *out_table = table;
  }
  return &table->mutable_fields()[field_idx];
}

}  // namespace

extern "C" fm_status_t fm_workbook_pivot_field_set_axis(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::size_t field_idx,
                                                        fm_pivot_axis_t axis) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  auto* field =
      lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_set_axis", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->axis = pivot_axis_from_fm(axis);
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_set_sort(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::size_t field_idx,
                                                        std::int32_t ascending, const char* by_field) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  auto* field =
      lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_set_sort", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->sort.ascending = ascending != 0;
  field->sort.by_field = by_field != nullptr ? by_field : "";
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_set_subtotal_top(fm_workbook_t* wb, std::size_t sheet_index,
                                                                std::size_t pivot_index, std::size_t field_idx,
                                                                std::int32_t top) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  auto* field = lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx,
                                       "fm_workbook_pivot_field_set_subtotal_top", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->subtotal_top = top != 0;
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_add_aggregation(fm_workbook_t* wb, std::size_t sheet_index,
                                                               std::size_t pivot_index, std::size_t field_idx,
                                                               fm_pivot_aggregation_t agg) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  auto* field = lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx,
                                       "fm_workbook_pivot_field_add_aggregation", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->aggregations.push_back(pivot_agg_from_fm(agg));
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_clear_aggregations(fm_workbook_t* wb, std::size_t sheet_index,
                                                                  std::size_t pivot_index, std::size_t field_idx) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  auto* field = lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx,
                                       "fm_workbook_pivot_field_clear_aggregations", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->aggregations.clear();
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_add_item(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::size_t field_idx,
                                                        const char* utf8_name, std::int32_t visible) {
  clear_last_error();
  if (utf8_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_add_item: utf8_name is NULL");
  }
  formulon::pivot::PivotTable* table = nullptr;
  auto* field =
      lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_add_item", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  formulon::pivot::PivotItem item;
  item.name = utf8_name;
  item.visible = visible != 0;
  field->items.push_back(std::move(item));
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_clear_items(fm_workbook_t* wb, std::size_t sheet_index,
                                                           std::size_t pivot_index, std::size_t field_idx) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  auto* field =
      lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_clear_items", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->items.clear();
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_set_item_visible(fm_workbook_t* wb, std::size_t sheet_index,
                                                                std::size_t pivot_index, std::size_t field_idx,
                                                                std::size_t item_idx, std::int32_t visible) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  auto* field = lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx,
                                       "fm_workbook_pivot_field_set_item_visible", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  if (item_idx >= field->items.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_field_set_item_visible: item_idx out of range",
                             "item_idx=" + std::to_string(item_idx));
  }
  field->items[item_idx].visible = visible != 0;
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_add_subtotal_fn(fm_workbook_t* wb, std::size_t sheet_index,
                                                               std::size_t pivot_index, std::size_t field_idx,
                                                               fm_pivot_aggregation_t agg) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  auto* field = lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx,
                                       "fm_workbook_pivot_field_add_subtotal_fn", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->subtotal_fns.push_back(pivot_subtotal_from_fm(agg));
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_clear_subtotal_fns(fm_workbook_t* wb, std::size_t sheet_index,
                                                                  std::size_t pivot_index, std::size_t field_idx) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  auto* field = lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx,
                                       "fm_workbook_pivot_field_clear_subtotal_fns", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->subtotal_fns.clear();
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_set_date_group(fm_workbook_t* wb, std::size_t sheet_index,
                                                              std::size_t pivot_index, std::size_t field_idx,
                                                              fm_pivot_date_grouping_t granularity,
                                                              fm_pivot_calendar_t calendar,
                                                              std::int32_t start_year_or_neg1,
                                                              std::int32_t end_year_or_neg1) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  auto* field =
      lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_set_date_group", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  formulon::pivot::PivotDateGroup grp;
  grp.granularity = pivot_date_grouping_from_fm(granularity);
  grp.calendar = pivot_calendar_from_fm(calendar);
  if (start_year_or_neg1 != -1) {
    grp.start_year = start_year_or_neg1;
  }
  if (end_year_or_neg1 != -1) {
    grp.end_year = end_year_or_neg1;
  }
  field->date_group = std::move(grp);
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_clear_date_group(fm_workbook_t* wb, std::size_t sheet_index,
                                                                std::size_t pivot_index, std::size_t field_idx) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  auto* field = lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx,
                                       "fm_workbook_pivot_field_clear_date_group", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->date_group.reset();
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_set_number_format(fm_workbook_t* wb, std::size_t sheet_index,
                                                                 std::size_t pivot_index, std::size_t field_idx,
                                                                 const char* utf8) {
  clear_last_error();
  if (utf8 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_set_number_format: utf8 is NULL");
  }
  formulon::pivot::PivotTable* table = nullptr;
  auto* field = lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx,
                                       "fm_workbook_pivot_field_set_number_format", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->number_format = utf8;
  invalidate_pivot_result(*table);
  return 0;
}

// ---------------------------------------------------------------------------
// PivotTable field order (row / col)
// ---------------------------------------------------------------------------

namespace {

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
  for (std::size_t i = 0; i < count; ++i) {
    if (indices[i] >= field_count) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               (std::string(fn) + ": indices[i] out of range").c_str(),
                               "i=" + std::to_string(i) + " value=" + std::to_string(indices[i]));
    }
  }
  std::vector<std::uint32_t>& target = row ? table->mutable_row_field_order() : table->mutable_col_field_order();
  target.assign(indices, indices + count);
  invalidate_pivot_result(*table);
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
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_data_field_count: NULL argument");
  }
  const auto* table = resolve_pivot(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_data_field_count");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  *out_count = table->data_fields().size();
  return 0;
}

namespace {

// Materialises a `fm_pivot_data_field_spec_t` into a `PivotDataField`,
// validating the required string fields. On failure writes the binding
// error and returns false.
bool fill_data_field(const fm_pivot_data_field_spec_t& spec, formulon::pivot::PivotDataField* out, const char* fn) {
  if (spec.name == nullptr) {
    set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                      (std::string(fn) + ": spec->name is NULL").c_str());
    return false;
  }
  out->name = spec.name;
  out->field_index = spec.field_index;
  out->aggregation = pivot_agg_from_fm(spec.aggregation);
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
  return true;
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
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_data_field_add");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  formulon::pivot::PivotDataField df;
  if (!fill_data_field(*spec, &df, "fm_workbook_pivot_data_field_add")) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  }
  table->mutable_data_fields().push_back(std::move(df));
  *out_idx = table->data_fields().size() - 1U;
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_data_field_clear(fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_data_field_clear: wb is NULL");
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_data_field_clear");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  table->mutable_data_fields().clear();
  invalidate_pivot_result(*table);
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
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_data_field_set");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  if (data_field_idx >= table->data_fields().size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_data_field_set: data_field_idx out of range",
                             "data_field_idx=" + std::to_string(data_field_idx));
  }
  formulon::pivot::PivotDataField df;
  if (!fill_data_field(*spec, &df, "fm_workbook_pivot_data_field_set")) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  }
  table->mutable_data_fields()[data_field_idx] = std::move(df);
  invalidate_pivot_result(*table);
  return 0;
}

// ---------------------------------------------------------------------------
// PivotTable filters
// ---------------------------------------------------------------------------

namespace {

fm_status_t add_pivot_filter_impl(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                  const fm_pivot_filter_spec_ex_t* spec, bool validate_data_field_index,
                                  const char* api) {
  if (wb == nullptr || spec == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(api) + ": NULL argument").c_str());
  }
  if (spec->field_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(api) + ": spec->field_name is NULL").c_str());
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, api);
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  const bool is_value_filter = spec->type == FM_PIVOT_FILTER_VALUE_TOP_10 ||
                               spec->type == FM_PIVOT_FILTER_VALUE_GREATER_THAN ||
                               spec->type == FM_PIVOT_FILTER_VALUE_BETWEEN;
  if (validate_data_field_index && is_value_filter && spec->data_field_index >= table->data_fields().size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(api) + ": data_field_index out of range").c_str(),
                             "data_field_index=" + std::to_string(spec->data_field_index));
  }
  formulon::pivot::PivotFilter filter;
  filter.axis = pivot_axis_from_fm(spec->axis);
  filter.field_name = spec->field_name;
  filter.type = pivot_filter_type_from_fm(spec->type);
  filter.data_field_index = spec->data_field_index;
  switch (spec->value_kind) {
    case FM_PIVOT_FILTER_VALUE_INT:
      filter.value = static_cast<int>(spec->value_int);
      break;
    case FM_PIVOT_FILTER_VALUE_DOUBLE:
      filter.value = spec->value_double;
      break;
    case FM_PIVOT_FILTER_VALUE_TEXT:
      if (spec->value_text == nullptr) {
        return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                                 (std::string(api) + ": spec->value_text is NULL").c_str());
      }
      filter.value = std::string(spec->value_text);
      break;
    case FM_PIVOT_FILTER_VALUE_NONE:
    default:
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               (std::string(api) + ": spec->value_kind is unset").c_str(),
                               "value_kind=" + std::to_string(static_cast<int>(spec->value_kind)));
  }
  switch (spec->value_high_kind) {
    case FM_PIVOT_FILTER_VALUE_INT:
      filter.value_high = static_cast<int>(spec->value_high_int);
      break;
    case FM_PIVOT_FILTER_VALUE_DOUBLE:
      filter.value_high = spec->value_high_double;
      break;
    case FM_PIVOT_FILTER_VALUE_NONE:
      // Leave default monostate.
      break;
    case FM_PIVOT_FILTER_VALUE_TEXT:
    default:
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               (std::string(api) + ": spec->value_high_kind not supported").c_str(),
                               "value_high_kind=" + std::to_string(static_cast<int>(spec->value_high_kind)));
  }
  table->mutable_active_filters().push_back(std::move(filter));
  invalidate_pivot_result(*table);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_pivot_filter_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                      std::size_t pivot_index, std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_filter_count: NULL argument");
  }
  const auto* table = resolve_pivot(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_filter_count");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  *out_count = table->active_filters().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_filter_add_ex(fm_workbook_t* wb, std::size_t sheet_index,
                                                       std::size_t pivot_index, const fm_pivot_filter_spec_ex_t* spec) {
  clear_last_error();
  return add_pivot_filter_impl(wb, sheet_index, pivot_index, spec, true, "fm_workbook_pivot_filter_add_ex");
}

extern "C" fm_status_t fm_workbook_pivot_filter_add(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                    const fm_pivot_filter_spec_t* spec) {
  clear_last_error();
  fm_pivot_filter_spec_ex_t ex{};
  if (spec != nullptr) {
    ex.axis = spec->axis;
    ex.field_name = spec->field_name;
    ex.type = spec->type;
    ex.value_kind = spec->value_kind;
    ex.value_int = spec->value_int;
    ex.value_double = spec->value_double;
    ex.value_text = spec->value_text;
    ex.value_high_kind = spec->value_high_kind;
    ex.value_high_int = spec->value_high_int;
    ex.value_high_double = spec->value_high_double;
  }
  ex.data_field_index = 0;
  return add_pivot_filter_impl(wb, sheet_index, pivot_index, spec == nullptr ? nullptr : &ex, false,
                               "fm_workbook_pivot_filter_add");
}

extern "C" fm_status_t fm_workbook_pivot_filter_clear(fm_workbook_t* wb, std::size_t sheet_index,
                                                      std::size_t pivot_index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_filter_clear: wb is NULL");
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_filter_clear");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  table->mutable_active_filters().clear();
  invalidate_pivot_result(*table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_filter_remove_at(fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index, std::size_t filter_idx) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_filter_remove_at: wb is NULL");
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_filter_remove_at");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  auto& filters = table->mutable_active_filters();
  if (filter_idx >= filters.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_filter_remove_at: filter_idx out of range",
                             "filter_idx=" + std::to_string(filter_idx));
  }
  filters.erase(filters.begin() + static_cast<std::ptrdiff_t>(filter_idx));
  invalidate_pivot_result(*table);
  return 0;
}
