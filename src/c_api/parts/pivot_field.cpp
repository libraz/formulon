//
// C ABI - PivotTable field surface (fields, items, date grouping,
// read-back). Shares helpers with `pivot_table.cpp` via `pivot_internal.h`.

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/pivot_internal.h"
#include "eval/collation.h"
#include "pivot/field_grouping.h"
#include "pivot/field_lookup.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/utf8_length.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_num_fmt_id_text;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::find_cache;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::invalidate_pivot_result;
using formulon::c_api::parts::pivot_axis_from_fm;
using formulon::c_api::parts::pivot_axis_to_fm;
using formulon::c_api::parts::pivot_calendar_from_fm;
using formulon::c_api::parts::pivot_count;
using formulon::c_api::parts::pivot_subtotal_from_fm;
using formulon::c_api::parts::resolve_pivot;
using formulon::c_api::parts::resolve_pivot_checked;
using formulon::c_api::parts::resolve_pivot_mut;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::validate;

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

// Longest item caption, in UTF-16 code units.
constexpr std::uint32_t kMaxItemCaptionUnits = 255U;

// True when `a` and `b` are one item under the profile's text collation, the
// comparison the evaluator merges items with.
bool same_item_label(const formulon::Workbook& workbook, std::string_view a, std::string_view b) {
  return formulon::eval::collate(a, b, workbook.excel_profile(), formulon::eval::CollationStrength::kDynamic) == 0;
}

// The label `item` takes from its value alone: its own name, else the bound
// cache value's text.
std::string item_value_label(const formulon::Workbook& workbook, const formulon::pivot::PivotTable& table,
                             std::size_t field_idx, const formulon::pivot::PivotItem& item) {
  if (const auto* cache = find_cache(workbook, table.pivot_cache_id()); cache != nullptr) {
    return formulon::pivot::pivot_item_label(*cache, field_idx, item);
  }
  return item.name;
}

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
    if (!formulon::pivot::is_value_filter_type(filter.type) || filter.data_field_index >= new_index.size()) {
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
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(fn) + ": wb is NULL").c_str());
  }
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

extern "C" fm_status_t fm_workbook_pivot_field_set_default_subtotal(fm_workbook_t* wb, std::size_t sheet_index,
                                                                    std::size_t pivot_index, std::size_t field_idx,
                                                                    std::int32_t on) {
  clear_last_error();
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_set_default_subtotal",
                            [on](formulon::pivot::PivotField& field) { field.default_subtotal = on != 0; });
}

extern "C" fm_status_t fm_workbook_pivot_field_set_insert_blank_row(fm_workbook_t* wb, std::size_t sheet_index,
                                                                    std::size_t pivot_index, std::size_t field_idx,
                                                                    std::int32_t on) {
  clear_last_error();
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_set_insert_blank_row",
                            [on](formulon::pivot::PivotField& field) { field.insert_blank_row = on != 0; });
}

extern "C" fm_status_t fm_workbook_pivot_field_set_fill_down_labels(fm_workbook_t* wb, std::size_t sheet_index,
                                                                    std::size_t pivot_index, std::size_t field_idx,
                                                                    std::int32_t on) {
  clear_last_error();
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_set_fill_down_labels",
                            [on](formulon::pivot::PivotField& field) { field.fill_down_labels = on != 0; });
}

extern "C" fm_status_t fm_workbook_pivot_field_set_show_detail(fm_workbook_t* wb, std::size_t sheet_index,
                                                               std::size_t pivot_index, std::size_t field_idx,
                                                               std::int32_t on) {
  clear_last_error();
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_idx, "fm_workbook_pivot_field_set_show_detail",
                            [on](formulon::pivot::PivotField& field) {
                              for (formulon::pivot::PivotItem& item : field.items) {
                                item.show_detail = on != 0;
                              }
                              for (formulon::pivot::PivotMarkerItem& marker : field.marker_items) {
                                if (marker.type == "default") {
                                  marker.show_detail = on != 0;
                                }
                              }
                            });
}

extern "C" fm_status_t fm_workbook_pivot_field_set_item_show_detail(fm_workbook_t* wb, std::size_t sheet_index,
                                                                    std::size_t pivot_index, std::size_t field_idx,
                                                                    std::size_t item_idx, std::int32_t on) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_idx,
                                           "fm_workbook_pivot_field_set_item_show_detail", &table, &field);
      rc != 0) {
    return rc;
  }
  if (auto rc = check_index(item_idx, field->items.size(), "fm_workbook_pivot_field_set_item_show_detail", "item_idx");
      rc != 0) {
    return rc;
  }
  field->items[item_idx].show_detail = on != 0;
  invalidate_pivot_result(wb->workbook(), *table);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_field_set_item_caption(fm_workbook_t* wb, std::size_t sheet_index,
                                                                std::size_t pivot_index, std::size_t field_idx,
                                                                std::size_t item_idx, const char* caption) {
  clear_last_error();
  constexpr const char* kApi = "fm_workbook_pivot_field_set_item_caption";
  if (auto rc = check_utf8(kApi, {caption}); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_idx, kApi, &table, &field); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(item_idx, field->items.size(), kApi, "item_idx"); rc != 0) {
    return rc;
  }
  const std::string_view text = caption != nullptr ? std::string_view(caption) : std::string_view();
  if (formulon::utf16_units_in(text) > kMaxItemCaptionUnits) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(kApi) + ": caption longer than 255 UTF-16 code units").c_str());
  }
  formulon::pivot::PivotItem& item = field->items[item_idx];
  if (text.empty() ||
      same_item_label(wb->workbook(), text, item_value_label(wb->workbook(), *table, field_idx, item))) {
    item.caption.clear();
  } else {
    for (std::size_t i = 0; i < field->items.size(); ++i) {
      const formulon::pivot::PivotItem& other = field->items[i];
      const std::string shown =
          !other.caption.empty() ? other.caption : item_value_label(wb->workbook(), *table, field_idx, other);
      if (i != item_idx && same_item_label(wb->workbook(), text, shown)) {
        return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                                 (std::string(kApi) + ": caption matches another item of the field").c_str(),
                                 "item_idx=" + std::to_string(i));
      }
    }
    item.caption = std::string(text);
  }
  invalidate_pivot_result(wb->workbook(), *table);
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

extern "C" fm_status_t fm_workbook_pivot_field_set_date_group(fm_workbook_t* wb, std::size_t sheet_index,
                                                              std::size_t pivot_index, std::size_t field_idx,
                                                              std::uint32_t periods_mask, std::int32_t calendar,
                                                              std::uint32_t interval_days, double start_or_nan,
                                                              double end_or_nan) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_set_date_group: wb is NULL");
  }
  if (fm_status_t rc = check_enum_domain(calendar, 1, "fm_workbook_pivot_field_set_date_group", "calendar"); rc != 0) {
    return rc;
  }
  const auto bound = [](double serial) { return std::isnan(serial) ? std::nullopt : std::optional<double>(serial); };
  return finish_status(
      formulon::pivot::set_date_group(wb->workbook(), sheet_index, pivot_index, field_idx, periods_mask,
                                      pivot_calendar_from_fm(static_cast<fm_pivot_calendar_t>(calendar)), interval_days,
                                      bound(start_or_nan), bound(end_or_nan)));
}

extern "C" fm_status_t fm_workbook_pivot_field_clear_date_group(fm_workbook_t* wb, std::size_t sheet_index,
                                                                std::size_t pivot_index, std::size_t field_idx) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_field_clear_date_group: wb is NULL");
  }
  return finish_status(formulon::pivot::clear_date_group(wb->workbook(), sheet_index, pivot_index, field_idx));
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
  result.group_base_index = -1;
  if (result.cache_field_index >= 0) {
    const bool date1904 = wb->workbook().date1904();
    const std::size_t base = formulon::pivot::group_base_index(*cache, field_idx);
    if (base != field_idx) {
      result.group_base_index = static_cast<std::int32_t>(base);
    } else {
      result.date_periods = formulon::pivot::date_group_periods(*cache, field_idx, date1904);
      result.date_calendar =
          static_cast<std::int32_t>(formulon::pivot::date_group_calendar(*cache, field_idx, date1904));
    }
  }
  // Custom functions replace the automatic subtotal, as the writer and the
  // layout treat them, so the flag reads the effective state.
  result.default_subtotal = field->default_subtotal && field->subtotal_fns.empty() ? 1 : 0;
  for (const formulon::pivot::SubtotalFn fn : field->subtotal_fns) {
    result.subtotal_fns |= 1U << static_cast<unsigned>(fn);
  }
  result.insert_blank_row = field->insert_blank_row ? 1 : 0;
  result.fill_down_labels = field->fill_down_labels ? 1 : 0;
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
  result.caption = item.caption.empty() ? nullptr : item.caption.c_str();
  result.show_detail = item.show_detail ? 1 : 0;
  *out = result;
  return 0;
}
