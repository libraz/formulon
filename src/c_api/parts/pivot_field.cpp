//
// C ABI - PivotTable field surface (fields, items, date grouping,
// read-back). Shares helpers with `pivot_table.cpp` via `pivot_internal.h`.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/pivot_internal.h"
#include "pivot/date_serial.h"
#include "pivot/field_lookup.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/number_text.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_num_fmt_id_text;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::find_cache;
using formulon::c_api::parts::invalidate_pivot_result;
using formulon::c_api::parts::pivot_axis_from_fm;
using formulon::c_api::parts::pivot_axis_to_fm;
using formulon::c_api::parts::pivot_calendar_from_fm;
using formulon::c_api::parts::pivot_count;
using formulon::c_api::parts::pivot_date_grouping_from_fm;
using formulon::c_api::parts::pivot_subtotal_from_fm;
using formulon::c_api::parts::resolve_pivot;
using formulon::c_api::parts::resolve_pivot_checked;
using formulon::c_api::parts::resolve_pivot_mut;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::validate;
using formulon::pivot::is_valid_date_serial;

namespace {

// Drops `field` from the row (or column) order, shifting a Values marker
// that sat after it.
void remove_from_field_order(formulon::pivot::PivotTable& table, bool row, std::uint32_t field) {
  std::vector<std::uint32_t>& order = row ? table.mutable_row_field_order() : table.mutable_col_field_order();
  const auto it = std::find(order.begin(), order.end(), field);
  if (it == order.end()) {
    return;
  }
  const auto index = static_cast<std::size_t>(it - order.begin());
  order.erase(it);
  const std::optional<std::size_t> values_position = row ? table.row_values_position() : table.col_values_position();
  if (values_position.has_value() && index < *values_position) {
    if (row) {
      table.set_row_values_position(*values_position - 1U);
    } else {
      table.set_col_values_position(*values_position - 1U);
    }
  }
}

// Keeps the row / column orders in step with `field`'s axis: a Row or Col
// field is listed once in that axis's order (appended when new) and in no
// other order.
void sync_field_order(formulon::pivot::PivotTable& table, std::uint32_t field) {
  const formulon::pivot::PivotAxis axis = table.fields()[field].axis;
  for (const bool row : {true, false}) {
    if (axis != (row ? formulon::pivot::PivotAxis::Row : formulon::pivot::PivotAxis::Col)) {
      remove_from_field_order(table, row, field);
      continue;
    }
    std::vector<std::uint32_t>& order = row ? table.mutable_row_field_order() : table.mutable_col_field_order();
    if (std::find(order.begin(), order.end(), field) == order.end()) {
      order.push_back(field);
    }
  }
}

constexpr std::int32_t kPivotAxisNone = -1;

// A field taken off the report leaves no data fields behind, and the value
// filters scoring a dropped data field go with them; the others follow their
// data field's new position. Filters naming the field itself are removed too.
void drop_field_aggregates(formulon::pivot::PivotTable& table, std::size_t field_idx) {
  std::vector<formulon::pivot::PivotDataField>& data_fields = table.mutable_data_fields();
  std::vector<std::size_t> new_index(data_fields.size());
  std::size_t kept = 0;
  for (std::size_t i = 0; i < data_fields.size(); ++i) {
    new_index[i] = data_fields[i].field_index == field_idx ? data_fields.size() : kept++;
  }
  const formulon::pivot::PivotField& field = table.fields()[field_idx];
  table.erase_active_filters_if([&](formulon::pivot::PivotFilter& filter) {
    if (formulon::pivot::pivot_field_has_name(field, filter.field_name)) {
      return true;
    }
    const bool is_value = filter.type == formulon::pivot::FilterType::ValueTop10 ||
                          filter.type == formulon::pivot::FilterType::ValueGreaterThan ||
                          filter.type == formulon::pivot::FilterType::ValueBetween;
    if (!is_value || filter.data_field_index >= new_index.size()) {
      return false;
    }
    const std::size_t moved = new_index[filter.data_field_index];
    filter.data_field_index = static_cast<decltype(filter.data_field_index)>(moved);
    return moved == data_fields.size();
  });
  data_fields.erase(
      std::remove_if(data_fields.begin(), data_fields.end(),
                     [field_idx](const formulon::pivot::PivotDataField& df) { return df.field_index == field_idx; }),
      data_fields.end());
}

// Resolves `field_idx` on a const pivot. Sets the binding error and returns
// NULL on miss.
const formulon::pivot::PivotField* lookup_pivot_field(const fm_workbook_t* wb, std::size_t sheet_index,
                                                      std::size_t pivot_index, std::size_t field_idx, const char* fn,
                                                      const formulon::pivot::PivotTable** out_table) {
  const auto* table = resolve_pivot(wb->workbook(), sheet_index, pivot_index, fn);
  if (table == nullptr) {
    return nullptr;
  }
  if (check_index(field_idx, table->fields().size(), fn, "field_idx") != 0) {
    return nullptr;
  }
  if (out_table != nullptr) {
    *out_table = table;
  }
  return &table->fields()[field_idx];
}

}  // namespace

// ---------------------------------------------------------------------------
// PivotTable fields
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_pivot_field_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::size_t pivot_index, std::size_t* out_count) {
  clear_last_error();
  return pivot_count(wb, sheet_index, pivot_index, out_count, "fm_workbook_pivot_field_count",
                     [](const formulon::pivot::PivotTable& table) { return table.fields().size(); });
}

extern "C" fm_status_t fm_workbook_pivot_field_add(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                   const fm_pivot_field_spec_t* spec, std::size_t* out_field_idx) {
  clear_last_error();
  if (wb == nullptr || spec == nullptr || out_field_idx == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_add: NULL argument");
  }
  if (auto rc = validate(*spec, "fm_workbook_pivot_field_add"); rc != 0) {
    return rc;
  }
  if (auto rc = check_num_fmt_id_text(wb->workbook(), spec->number_format, "fm_workbook_pivot_field_add"); rc != 0) {
    return rc;
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_field_add");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  // The pivot-field list is positional over the bound cache's fields, so a
  // field named in the cache is configured in place at the cache field's
  // position, after materialising a bare field for every cache column. A
  // name the cache does not carry is appended.
  std::vector<formulon::pivot::PivotField>& fields = table->mutable_fields();
  std::size_t index = fields.size();
  if (const auto* cache = find_cache(wb->workbook(), table->pivot_cache_id()); cache != nullptr) {
    const auto& cache_fields = cache->fields();
    for (std::size_t i = 0; i < cache_fields.size(); ++i) {
      if (cache_fields[i].name == spec->source_name) {
        index = i;
        break;
      }
    }
    if (index != fields.size()) {
      for (std::size_t i = fields.size(); i < cache_fields.size(); ++i) {
        formulon::pivot::PivotField bare;
        bare.source_name = cache_fields[i].name;
        bare.axis = formulon::pivot::PivotAxis::None;
        fields.push_back(std::move(bare));
      }
    }
  }
  if (index == fields.size()) {
    fields.emplace_back();
  }
  formulon::pivot::PivotField& field = fields[index];
  field.source_name = spec->source_name;
  field.custom_name = spec->custom_name != nullptr ? spec->custom_name : "";
  field.axis = pivot_axis_from_fm(spec->axis);
  field.subtotal_top = spec->subtotal_top != 0;
  field.number_format = spec->number_format != nullptr ? spec->number_format : "";
  *out_field_idx = index;
  sync_field_order(*table, static_cast<std::uint32_t>(index));
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_clear(fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::size_t pivot_index) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, "fm_workbook_pivot_field_clear", &table);
      rc != 0) {
    return rc;
  }
  table->mutable_fields().clear();
  table->mutable_row_field_order().clear();
  table->mutable_col_field_order().clear();
  // A Values marker can only lead an empty order.
  if (table->row_values_position().has_value()) {
    table->set_row_values_position(0U);
  }
  if (table->col_values_position().has_value()) {
    table->set_col_values_position(0U);
  }
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

namespace {

// Mutable variant for the `field_set_*` / `field_add_*` family. Sets the
// binding error and returns NULL on miss.
formulon::pivot::PivotField* lookup_pivot_field_mut(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                    std::size_t field_idx, const char* fn,
                                                    formulon::pivot::PivotTable** out_table) {
  if (wb == nullptr) {
    set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, (std::string(fn) + ": wb is NULL").c_str());
    return nullptr;
  }
  const formulon::pivot::PivotTable* table = nullptr;
  if (lookup_pivot_field(wb, sheet_index, pivot_index, field_idx, fn, &table) == nullptr) {
    return nullptr;
  }
  // The pivot tables are owned non-const objects; only the lookup path is shared.
  auto* mutable_table = const_cast<formulon::pivot::PivotTable*>(table);
  if (out_table != nullptr) {
    *out_table = mutable_table;
  }
  return &mutable_table->mutable_fields()[field_idx];
}

// Status-returning wrapper over `lookup_pivot_field_mut` for the setters below.
fm_status_t resolve_pivot_field(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                std::size_t field_idx, const char* fn, formulon::pivot::PivotTable** table,
                                formulon::pivot::PivotField** field) {
  *field = lookup_pivot_field_mut(wb, sheet_index, pivot_index, field_idx, fn, table);
  return *field != nullptr ? 0 : static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
}

template <typename Mutation>
fm_status_t mutate_pivot_field(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                               std::size_t field_idx, const char* fn, Mutation&& mutation) {
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_idx, fn, &table, &field); rc != 0) {
    return rc;
  }
  std::forward<Mutation>(mutation)(*field);
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_pivot_field_set_axis(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::size_t field_idx,
                                                        std::int32_t axis) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_set_axis: wb is NULL");
  }
  // -1 takes the field off every axis.
  fm_status_t enum_status =
      axis == kPivotAxisNone ? 0 : check_enum_domain(axis, 3, "fm_workbook_pivot_field_set_axis", "axis");
  if (enum_status != 0) {
    return enum_status;
  }
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_set_axis",
                                           &table, &field);
      rc != 0) {
    return rc;
  }
  if (axis == kPivotAxisNone) {
    field->axis = formulon::pivot::PivotAxis::None;
    drop_field_aggregates(*table, field_idx);
  } else {
    field->axis = pivot_axis_from_fm(static_cast<fm_pivot_axis_t>(axis));
  }
  sync_field_order(*table, static_cast<std::uint32_t>(field_idx));
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_set_sort(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::size_t field_idx,
                                                        std::int32_t order, const char* by_field) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_set_sort: wb is NULL");
  }
  fm_status_t enum_status = check_enum_domain(order, FM_PIVOT_SORT_MANUAL, "fm_workbook_pivot_field_set_sort", "order");
  if (enum_status != 0) {
    return enum_status;
  }
  if (auto rc = check_utf8("fm_workbook_pivot_field_set_sort", {by_field}); rc != 0) {
    return rc;
  }
  const bool manual = order == FM_PIVOT_SORT_MANUAL;
  if (manual && by_field != nullptr && by_field[0] != '\0') {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_field_set_sort: manual order takes no by_field key");
  }
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_set_sort",
                            [order, manual, by_field](formulon::pivot::PivotField& field) {
                              field.sort.manual = manual;
                              field.sort.stated = true;
                              // Manual keeps the ascending form the OOXML reader yields for it.
                              field.sort.ascending = order != FM_PIVOT_SORT_DESCENDING;
                              field.sort.by_field = by_field != nullptr ? by_field : "";
                            });
}

extern "C" fm_status_t fm_workbook_pivot_field_set_subtotal_top(fm_workbook_t* wb, std::size_t sheet_index,
                                                                std::size_t pivot_index, std::size_t field_idx,
                                                                std::int32_t top) {
  clear_last_error();
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_set_subtotal_top",
                            [top](formulon::pivot::PivotField& field) { field.subtotal_top = top != 0; });
}

extern "C" fm_status_t fm_workbook_pivot_field_add_item(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::size_t field_idx,
                                                        const char* utf8_name, std::int32_t visible) {
  clear_last_error();
  if (utf8_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_add_item: utf8_name is NULL");
  }
  if (auto rc = check_utf8("fm_workbook_pivot_field_add_item", {utf8_name}); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_add_item",
                                           &table, &field);
      rc != 0) {
    return rc;
  }
  formulon::pivot::PivotItem item;
  item.name = utf8_name;
  item.visible = visible != 0;
  // Bind to the shared item carrying this label so the writer emits the matching `x` and the item
  // keeps its place in a manual order across a save. A label with no match stays unbound.
  if (!item.name.empty()) {
    if (const auto* cache = find_cache(wb->workbook(), table->pivot_cache_id());
        cache != nullptr && field_idx < cache->fields().size()) {
      formulon::pivot::PivotItem probe;
      const std::size_t shared_count = cache->fields()[field_idx].shared_items.size();
      for (std::size_t i = 0; i < shared_count; ++i) {
        probe.cache_index = static_cast<std::uint32_t>(i);
        if (formulon::pivot::pivot_item_label(*cache, field_idx, probe) == item.name) {
          item.has_cache_index = true;
          item.cache_index = probe.cache_index;
          break;
        }
      }
    }
  }
  field->items.push_back(std::move(item));
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_add_item_at(fm_workbook_t* wb, std::size_t sheet_index,
                                                           std::size_t pivot_index, std::size_t field_idx,
                                                           std::uint32_t cache_index, std::int32_t visible) {
  clear_last_error();
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_add_item_at",
                            [cache_index, visible](formulon::pivot::PivotField& field) {
                              // The label is left empty: evaluation derives it from the bound shared item
                              // (`pivot_item_label`). A binding that resolves to a blank value has no
                              // label, which is what makes this the constructor for the blank item.
                              formulon::pivot::PivotItem item;
                              item.visible = visible != 0;
                              item.has_cache_index = true;
                              item.cache_index = cache_index;
                              field.items.push_back(std::move(item));
                            });
}

extern "C" fm_status_t fm_workbook_pivot_field_clear_items(fm_workbook_t* wb, std::size_t sheet_index,
                                                           std::size_t pivot_index, std::size_t field_idx) {
  clear_last_error();
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_clear_items",
                            [](formulon::pivot::PivotField& field) {
                              field.items.clear();
                              field.marker_items.clear();
                            });
}

extern "C" fm_status_t fm_workbook_pivot_field_set_item_visible(fm_workbook_t* wb, std::size_t sheet_index,
                                                                std::size_t pivot_index, std::size_t field_idx,
                                                                std::size_t item_idx, std::int32_t visible) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_idx,
                                           "fm_workbook_pivot_field_set_item_visible", &table, &field);
      rc != 0) {
    return rc;
  }
  if (auto rc = check_index(item_idx, field->items.size(), "fm_workbook_pivot_field_set_item_visible", "item_idx");
      rc != 0) {
    return rc;
  }
  field->items[item_idx].visible = visible != 0;
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_add_subtotal_fn(fm_workbook_t* wb, std::size_t sheet_index,
                                                               std::size_t pivot_index, std::size_t field_idx,
                                                               std::int32_t agg) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_add_subtotal_fn: wb is NULL");
  }
  fm_status_t enum_status = check_enum_domain(agg, 10, "fm_workbook_pivot_field_add_subtotal_fn", "aggregation");
  if (enum_status != 0) {
    return enum_status;
  }
  return mutate_pivot_field(
      wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_add_subtotal_fn",
      [agg](formulon::pivot::PivotField& field) {
        field.subtotal_fns.push_back(pivot_subtotal_from_fm(static_cast<fm_pivot_aggregation_t>(agg)));
      });
}

extern "C" fm_status_t fm_workbook_pivot_field_clear_subtotal_fns(fm_workbook_t* wb, std::size_t sheet_index,
                                                                  std::size_t pivot_index, std::size_t field_idx) {
  clear_last_error();
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_clear_subtotal_fns",
                            [](formulon::pivot::PivotField& field) { field.subtotal_fns.clear(); });
}

extern "C" fm_status_t fm_workbook_pivot_field_set_date_group(
    fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index, std::size_t field_idx,
    std::int32_t granularity, std::int32_t calendar, std::int32_t start_year_or_neg1, std::int32_t end_year_or_neg1,
    std::uint32_t interval_days, double start_serial_or_neg1, double end_serial_or_neg1) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_set_date_group: wb is NULL");
  }
  fm_status_t enum_status = check_enum_domain(granularity, 7, "fm_workbook_pivot_field_set_date_group", "granularity");
  if (enum_status != 0) {
    return enum_status;
  }
  enum_status = check_enum_domain(calendar, 1, "fm_workbook_pivot_field_set_date_group", "calendar");
  if (enum_status != 0) {
    return enum_status;
  }
  const bool has_start = start_serial_or_neg1 != -1.0;
  const bool has_end = end_serial_or_neg1 != -1.0;
  if (granularity == FM_PIVOT_DATE_DAYS) {
    if (interval_days == 0) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "interval_days must be non-zero",
                               "fm_workbook_pivot_field_set_date_group: interval_days=0");
    }
    if (has_start && has_end && start_serial_or_neg1 > end_serial_or_neg1) {
      return set_binding_error(
          formulon::FormulonErrorCode::kInvalidArgument, "start_serial_or_neg1 must not exceed end_serial_or_neg1",
          "fm_workbook_pivot_field_set_date_group: start=" + formulon::fixed_string(start_serial_or_neg1, 6) +
              " end=" + formulon::fixed_string(end_serial_or_neg1, 6));
    }
    if ((has_start && !is_valid_date_serial(start_serial_or_neg1, wb->workbook().date1904())) ||
        (has_end && !is_valid_date_serial(end_serial_or_neg1, wb->workbook().date1904()))) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "date-group serial bounds must be finite dates in the supported calendar",
                               "fm_workbook_pivot_field_set_date_group: invalid serial bound");
    }
  }
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_set_date_group",
                            [=](formulon::pivot::PivotField& field) {
                              formulon::pivot::PivotDateGroup grp;
                              grp.granularity =
                                  pivot_date_grouping_from_fm(static_cast<fm_pivot_date_grouping_t>(granularity));
                              grp.calendar = pivot_calendar_from_fm(static_cast<fm_pivot_calendar_t>(calendar));
                              if (start_year_or_neg1 != -1) {
                                grp.start_year = start_year_or_neg1;
                              }
                              if (end_year_or_neg1 != -1) {
                                grp.end_year = end_year_or_neg1;
                              }
                              grp.interval_days = interval_days;
                              if (has_start) {
                                grp.start_serial = start_serial_or_neg1;
                              }
                              if (has_end) {
                                grp.end_serial = end_serial_or_neg1;
                              }
                              field.date_group = std::move(grp);
                            });
}

extern "C" fm_status_t fm_workbook_pivot_field_clear_date_group(fm_workbook_t* wb, std::size_t sheet_index,
                                                                std::size_t pivot_index, std::size_t field_idx) {
  clear_last_error();
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_clear_date_group",
                            [](formulon::pivot::PivotField& field) { field.date_group.reset(); });
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
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_idx,
                                           "fm_workbook_pivot_field_set_number_format", &table, &field);
      rc != 0) {
    return rc;
  }
  if (auto rc = check_num_fmt_id_text(wb->workbook(), utf8, "fm_workbook_pivot_field_set_number_format"); rc != 0) {
    return rc;
  }
  field->number_format = utf8;
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_at(const fm_workbook_t* wb, std::size_t sheet_index,
                                                  std::size_t pivot_index, std::size_t field_idx,
                                                  fm_pivot_field_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_at: NULL argument");
  }
  const formulon::pivot::PivotTable* table = nullptr;
  const formulon::pivot::PivotField* field =
      lookup_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_at", &table);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  fm_pivot_field_info_t result{};
  result.source_name = field->source_name.c_str();
  result.custom_name = field->custom_name.c_str();
  const auto axis = pivot_axis_to_fm(field->axis);
  result.axis = axis.has_value() ? static_cast<std::int32_t>(*axis) : -1;
  result.subtotal_top = field->subtotal_top ? 1 : 0;
  result.number_format = field->number_format.c_str();
  const auto* cache = find_cache(wb->workbook(), table->pivot_cache_id());
  result.cache_field_index =
      cache != nullptr && field_idx < cache->fields().size() ? static_cast<std::int32_t>(field_idx) : -1;
  result.sort_order = field->sort.manual      ? FM_PIVOT_SORT_MANUAL
                      : field->sort.ascending ? FM_PIVOT_SORT_ASCENDING
                                              : FM_PIVOT_SORT_DESCENDING;
  result.sort_by_field = field->sort.by_field.c_str();
  result.date_start_year = -1;
  result.date_end_year = -1;
  result.date_start_serial = -1.0;
  result.date_end_serial = -1.0;
  if (field->date_group.has_value()) {
    const formulon::pivot::PivotDateGroup& group = *field->date_group;
    result.has_date_group = 1;
    result.date_granularity = static_cast<std::int32_t>(group.granularity);
    result.date_calendar = static_cast<std::int32_t>(group.calendar);
    result.date_start_year = group.start_year.value_or(-1);
    result.date_end_year = group.end_year.value_or(-1);
    result.date_interval_days = group.interval_days;
    result.date_start_serial = group.start_serial.value_or(-1.0);
    result.date_end_serial = group.end_serial.value_or(-1.0);
  }
  // Custom functions replace the automatic subtotal, as the writer and the
  // layout treat them, so the flag reads the effective state.
  result.default_subtotal = field->default_subtotal && field->subtotal_fns.empty() ? 1 : 0;
  for (const formulon::pivot::SubtotalFn fn : field->subtotal_fns) {
    result.subtotal_fns |= 1U << static_cast<unsigned>(fn);
  }
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_item_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index, std::size_t field_idx,
                                                          std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_item_count: NULL argument");
  }
  const formulon::pivot::PivotField* field =
      lookup_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_item_count", nullptr);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  *out_count = field->items.size();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_item_at(const fm_workbook_t* wb, std::size_t sheet_index,
                                                       std::size_t pivot_index, std::size_t field_idx,
                                                       std::size_t item_idx, fm_pivot_item_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_item_at: NULL argument");
  }
  const formulon::pivot::PivotField* field =
      lookup_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_item_at", nullptr);
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  if (auto rc = check_index(item_idx, field->items.size(), "fm_workbook_pivot_field_item_at", "item_idx"); rc != 0) {
    return rc;
  }
  const formulon::pivot::PivotItem& item = field->items[item_idx];
  fm_pivot_item_info_t result{};
  result.name = item.name.c_str();
  result.cache_index = item.has_cache_index ? static_cast<std::int32_t>(item.cache_index) : -1;
  result.visible = item.visible ? 1 : 0;
  *out = result;
  return 0;
}
