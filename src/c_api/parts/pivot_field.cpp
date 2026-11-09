//
// C ABI - PivotTable field surface (fields, items, date grouping,
// read-back). Shares helpers with `pivot_table.cpp` via `pivot_internal.h`.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/pivot_internal.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "pivot/field_grouping.h"
#include "pivot/field_lookup.h"
#include "pivot/field_patch.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/expected.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_finite;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_limit;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::find_cache;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::mark_cache_pivots_pending;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::pivot_axis_from_fm;
using formulon::c_api::parts::pivot_axis_to_fm;
using formulon::c_api::parts::pivot_calendar_from_fm;
using formulon::c_api::parts::pivot_count;
using formulon::c_api::parts::pivot_num_fmt_code;
using formulon::c_api::parts::pivot_num_fmt_id_text;
using formulon::c_api::parts::pivot_subtotal_from_fm;
using formulon::c_api::parts::PivotEdit;
using formulon::c_api::parts::resolve_pivot;
using formulon::c_api::parts::resolve_pivot_checked;
using formulon::c_api::parts::resolve_pivot_mut;
using formulon::c_api::parts::set_values_position;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::text_or_null;
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
    set_values_position(table, row, *values_position - 1U);
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

// Resolves `field_index` on a const pivot. Returns 0 with `*out_field` (and
// `*out_table` when non-NULL) set, or the error status with the binding error
// populated.
fm_status_t lookup_pivot_field(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                               std::size_t field_index, const formulon::pivot::PivotTable** out_table,
                               const formulon::pivot::PivotField** out_field) {
  *out_field = nullptr;
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(field_index, table->fields().size(), "field_index"); rc != 0) {
    return rc;
  }
  if (out_table != nullptr) {
    *out_table = table;
  }
  *out_field = &table->fields()[field_index];
  return 0;
}

}  // namespace

// ---------------------------------------------------------------------------
// PivotTable fields
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_set_pivot_page_field_order(fm_workbook_t* wb, std::size_t sheet_index,
                                                           std::size_t pivot_index, const std::uint32_t* indices,
                                                           std::size_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (count > 0 && indices == nullptr) {
    return null_argument("indices");
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_mut(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  const std::size_t field_count = table->fields().size();
  std::vector<bool> requested_field(field_count, false);
  // Every entry is validated before the first change.
  std::vector<std::uint32_t> requested_order;
  requested_order.reserve(count);
  for (std::size_t i = 0; i < count; ++i) {
    if (indices[i] == FM_PIVOT_VALUES_FIELD) {
      return invalid_argument("the Values field is never a page field", "i=" + std::to_string(i));
    }
    if (auto rc = check_index(indices[i], field_count, "indices[i]"); rc != 0) {
      return rc;
    }
    if (requested_field[indices[i]]) {
      return invalid_argument("indices lists a field twice",
                              "i=" + std::to_string(i) + " value=" + std::to_string(indices[i]));
    }
    requested_field[indices[i]] = true;
    requested_order.push_back(indices[i]);
  }
  // The list is the axis's whole membership: a page field left out comes off
  // the axis, and a listed field moves onto it from wherever it was.
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  for (const std::uint32_t field : table->page_field_order()) {
    if (!requested_field[field]) {
      table->set_field_axis(field, formulon::pivot::PivotAxis::None);
    }
  }
  std::vector<formulon::pivot::PivotPageField> pages;
  pages.reserve(requested_order.size());
  for (const std::uint32_t field : requested_order) {
    if (const formulon::pivot::PivotPageField* kept = table->find_page_field(field); kept != nullptr) {
      pages.push_back(*kept);
    } else {
      formulon::pivot::PivotPageField added;
      added.field_index = field;
      pages.push_back(added);
    }
    table->set_field_axis(field, formulon::pivot::PivotAxis::Page);
    remove_from_field_order(*table, true, field);
    remove_from_field_order(*table, false, field);
  }
  table->mutable_page_fields() = std::move(pages);
  return edit.commit();
}

extern "C" fm_status_t fm_sheet_get_pivot_page_field_order(const fm_workbook_t* wb, std::size_t sheet_index,
                                                           std::size_t pivot_index, std::uint32_t* out,
                                                           std::size_t capacity, std::size_t* out_count) {
  clear_last_error();
  if (out == nullptr && capacity > 0U) {
    return null_argument("out");
  }
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out_count}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  const std::vector<std::uint32_t> order = table->page_field_order();
  std::copy_n(order.begin(), std::min(capacity, order.size()), out);
  *out_count = order.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_pivot_field_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                      std::size_t pivot_index, std::size_t* out_count) {
  clear_last_error();
  return pivot_count(wb, sheet_index, pivot_index, out_count,
                     [](const formulon::pivot::PivotTable& table) { return table.fields().size(); });
}

extern "C" fm_status_t fm_sheet_add_pivot_field(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                const fm_pivot_field_spec_t* spec, std::size_t* out_field_index) {
  clear_last_error();
  if (wb == nullptr || spec == nullptr || out_field_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*spec); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(spec->subtotal_top, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "subtotal_top"); rc != 0) {
    return rc;
  }
  if (auto rc = validate(*spec); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_mut(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  std::string number_format;
  if (auto rc = pivot_num_fmt_id_text(wb->workbook(), spec->number_format, &number_format); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
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
  field.custom_name = text_or_empty(spec->custom_name);
  field.subtotal_top = spec->subtotal_top != FM_TRISTATE_FALSE;
  field.number_format = std::move(number_format);
  table->set_field_axis(static_cast<std::uint32_t>(index),
                        pivot_axis_from_fm(static_cast<fm_pivot_axis_t>(spec->axis)));
  sync_field_order(*table, static_cast<std::uint32_t>(index));
  const fm_status_t status = edit.commit();
  if (status == 0) {
    *out_field_index = index;
  }
  return status;
}

extern "C" fm_status_t fm_sheet_clear_pivot_fields(fm_workbook_t* wb, std::size_t sheet_index,
                                                   std::size_t pivot_index) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
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
  return edit.commit();
}

namespace {

// Mutable field lookup for the `field_set_*` / `field_add_*` family: rejects a
// NULL `wb`, then resolves like `lookup_pivot_field`.
fm_status_t resolve_pivot_field(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                std::size_t field_index, formulon::pivot::PivotTable** table,
                                formulon::pivot::PivotField** field) {
  *field = nullptr;
  if (wb == nullptr) {
    return null_argument("wb");
  }
  const formulon::pivot::PivotTable* found = nullptr;
  const formulon::pivot::PivotField* found_field = nullptr;
  if (auto rc = lookup_pivot_field(wb, sheet_index, pivot_index, field_index, &found, &found_field); rc != 0) {
    return rc;
  }
  // The pivot tables are owned non-const objects; only the lookup path is shared.
  auto* mutable_table = const_cast<formulon::pivot::PivotTable*>(found);
  if (table != nullptr) {
    *table = mutable_table;
  }
  *field = &mutable_table->mutable_fields()[field_index];
  return 0;
}

// Grouping edits the shared cache, so every pivot drawing from it is left
// pending rather than projected.
fm_status_t finish_cache_grouping(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                  const formulon::Expected<void, formulon::Error>& grouped) {
  if (grouped) {
    mark_cache_pivots_pending(wb->workbook(),
                              wb->workbook().sheet(sheet_index).pivot_tables()[pivot_index]->pivot_cache_id());
  }
  return finish_status(grouped);
}

template <typename Mutation>
fm_status_t mutate_pivot_field(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                               std::size_t field_index, Mutation&& mutation) {
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_index, &table, &field); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  std::forward<Mutation>(mutation)(*field);
  return edit.commit();
}

// Resolves a field and bounds-checks one of its items.
fm_status_t resolve_pivot_item(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                               std::size_t field_index, std::size_t item_index, formulon::pivot::PivotTable** table,
                               formulon::pivot::PivotField** field) {
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_index, table, field); rc != 0) {
    return rc;
  }
  return check_index(item_index, (*field)->items.size(), "item_index");
}

// Reads a C patch into the engine's optional keys, keeping the first failure: an op outside
// KEEP / CLEAR / SET, CLEAR on a key that refuses it, a flag other than 0 / 1, or a NULL or
// malformed string under SET.
class FieldPatchReader {
 public:
  fm_status_t status() const { return rc_; }

  bool accepts(fm_patch_op_t op, bool clearable, const char* name) {
    if (rc_ == 0 && !patch_op_ok(op)) {
      rc_ = invalid_argument("patch op out of range", name);
    }
    if (rc_ == 0 && !clearable && !patch_op_ok_without_clear(op)) {
      rc_ = invalid_argument("CLEAR is not valid for this field", name);
    }
    return rc_ == 0 && op != FM_PATCH_OP_KEEP;
  }

  void flag(fm_patch_op_t op, std::int32_t value, std::optional<bool>& out, const char* name) {
    if (accepts(op, false, name) && (rc_ = check_enum_domain(value, 0, 1, name)) == 0) {
      out = value == 1;
    }
  }

  // A clearable text key: CLEAR writes the empty string, which the engine reads as removal.
  void text(fm_patch_op_t op, const char* value, std::optional<std::string>& out, const char* name) {
    if (!accepts(op, true, name)) {
      return;
    }
    if (op == FM_PATCH_OP_CLEAR) {
      out = std::string();
      return;
    }
    if (value == nullptr) {
      rc_ = invalid_argument("string must not be NULL under SET", name);
      return;
    }
    if ((rc_ = check_utf8(value, "field", name)) == 0) {
      out = std::string(value);
    }
  }

 private:
  fm_status_t rc_ = 0;
};

formulon::pivot::FieldSortOrder field_sort_order_from_fm(std::int32_t order) {
  switch (order) {
    case FM_PIVOT_SORT_ORDER_ASCENDING:
      return formulon::pivot::FieldSortOrder::kAscending;
    case FM_PIVOT_SORT_ORDER_DESCENDING:
      return formulon::pivot::FieldSortOrder::kDescending;
    case FM_PIVOT_SORT_ORDER_MANUAL:
    default:
      return formulon::pivot::FieldSortOrder::kManual;
  }
}

}  // namespace

extern "C" fm_status_t fm_sheet_set_pivot_field_axis(fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::size_t pivot_index, std::size_t field_index,
                                                     std::int32_t axis) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  // -1 takes the field off every axis.
  fm_status_t enum_status = axis == FM_PIVOT_AXIS_ABSENT ? 0 : check_enum_domain(axis, 0, FM_PIVOT_AXIS_PAGE, "axis");
  if (enum_status != 0) {
    return enum_status;
  }
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_index, &table, &field); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  if (axis == FM_PIVOT_AXIS_ABSENT) {
    table->set_field_axis(static_cast<std::uint32_t>(field_index), formulon::pivot::PivotAxis::None);
    formulon::pivot::drop_field_aggregates(*table, field_index);
  } else {
    table->set_field_axis(static_cast<std::uint32_t>(field_index),
                          pivot_axis_from_fm(static_cast<fm_pivot_axis_t>(axis)));
  }
  sync_field_order(*table, static_cast<std::uint32_t>(field_index));
  return edit.commit();
}

extern "C" fm_status_t fm_sheet_update_pivot_field(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                   std::size_t field_index, const fm_pivot_field_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || patch == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_index, &table, &field); rc != 0) {
    return rc;
  }
  FieldPatchReader reader;
  formulon::pivot::PivotFieldPatch next;
  reader.text(patch->custom_name_op, patch->custom_name, next.custom_name, "custom_name");
  // The format code is checked here and interned only once the engine has accepted the rest.
  std::optional<std::string> number_format_code;
  reader.text(patch->number_format_op, patch->number_format, number_format_code, "number_format");
  reader.flag(patch->default_subtotal_op, patch->default_subtotal, next.default_subtotal, "default_subtotal");
  reader.flag(patch->insert_blank_row_op, patch->insert_blank_row, next.insert_blank_row, "insert_blank_row");
  reader.flag(patch->fill_down_labels_op, patch->fill_down_labels, next.fill_down_labels, "fill_down_labels");
  reader.flag(patch->show_all_op, patch->show_all, next.show_all, "show_all");
  reader.flag(patch->subtotal_top_op, patch->subtotal_top, next.subtotal_top, "subtotal_top");
  if (reader.accepts(patch->sort_order_op, false, "sort_order")) {
    if (auto rc = check_enum_domain(patch->sort_order, FM_PIVOT_SORT_ORDER_MANUAL, FM_PIVOT_SORT_ORDER_DESCENDING,
                                    "sort_order");
        rc != 0) {
      return rc;
    }
    next.sort_order = field_sort_order_from_fm(patch->sort_order);
  }
  reader.text(patch->sort_by_field_op, patch->sort_by_field, next.sort_by_field, "sort_by_field");
  reader.flag(patch->show_detail_op, patch->show_detail, next.show_detail, "show_detail");
  if (reader.status() != 0) {
    return reader.status();
  }
  if (auto checked = formulon::pivot::check_field_patch(wb->workbook(), *table, field_index, next); !checked) {
    return finish_status(checked);
  }
  const bool format_changes = number_format_code.has_value();
  if (format_changes) {
    std::string id_text;
    if (auto rc = pivot_num_fmt_id_text(wb->workbook(), number_format_code->c_str(), &id_text); rc != 0) {
      return rc;
    }
    next.number_format = std::move(id_text);
  }
  const bool changes = format_changes || next.custom_name || next.default_subtotal || next.insert_blank_row ||
                       next.fill_down_labels || next.show_all || next.subtotal_top || next.sort_order ||
                       next.sort_by_field || next.show_detail;
  if (!changes) {
    return 0;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  if (auto applied = formulon::pivot::apply_field_patch(wb->workbook(), *table, field_index, next); !applied) {
    return finish_status(applied);
  }
  return edit.commit();
}

extern "C" fm_status_t fm_sheet_update_pivot_field_item(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::size_t field_index,
                                                        std::size_t item_index, const fm_pivot_item_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || patch == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_item(wb, sheet_index, pivot_index, field_index, item_index, &table, &field);
      rc != 0) {
    return rc;
  }
  FieldPatchReader reader;
  formulon::pivot::PivotItemPatch next;
  reader.flag(patch->visible_op, patch->visible, next.visible, "visible");
  reader.flag(patch->show_detail_op, patch->show_detail, next.show_detail, "show_detail");
  reader.text(patch->caption_op, patch->caption, next.caption, "caption");
  if (reader.status() != 0) {
    return reader.status();
  }
  if (auto checked = formulon::pivot::check_item_patch(wb->workbook(), *table, field_index, item_index, next);
      !checked) {
    return finish_status(checked);
  }
  if (!next.visible && !next.show_detail && !next.caption) {
    return 0;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  if (auto applied = formulon::pivot::apply_item_patch(wb->workbook(), *table, field_index, item_index, next);
      !applied) {
    return finish_status(applied);
  }
  return edit.commit();
}

extern "C" fm_status_t fm_sheet_add_pivot_field_item(fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::size_t pivot_index, std::size_t field_index,
                                                     const char* utf8_name, std::int32_t visible,
                                                     std::size_t* out_index) {
  clear_last_error();
  if (utf8_name == nullptr || out_index == nullptr) {
    return null_argument(utf8_name == nullptr ? "utf8_name" : "out_index");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({utf8_name}); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotField* field = nullptr;
  if (fm_status_t rc = resolve_pivot_field(wb, sheet_index, pivot_index, field_index, &table, &field); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotItem item;
  item.name = utf8_name;
  item.visible = visible != 0;
  // Bind to the shared item carrying this label so the writer emits the matching `x` and the item
  // keeps its place in a manual order across a save. A label with no match stays unbound.
  if (!item.name.empty()) {
    if (const auto* cache = find_cache(wb->workbook(), table->pivot_cache_id());
        cache != nullptr && field_index < cache->fields().size()) {
      formulon::pivot::PivotItem probe;
      const std::size_t shared_count = cache->fields()[field_index].shared_items.size();
      for (std::size_t i = 0; i < shared_count; ++i) {
        probe.cache_index = static_cast<std::uint32_t>(i);
        if (formulon::pivot::pivot_item_label(*cache, field_index, probe) == item.name) {
          item.has_cache_index = true;
          item.cache_index = probe.cache_index;
          break;
        }
      }
    }
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  field->items.push_back(std::move(item));
  const std::size_t added = field->items.size() - 1U;
  const fm_status_t status = edit.commit();
  if (status == 0) {
    *out_index = added;
  }
  return status;
}

extern "C" fm_status_t fm_sheet_add_pivot_field_item_by_cache_index(fm_workbook_t* wb, std::size_t sheet_index,
                                                                    std::size_t pivot_index, std::size_t field_index,
                                                                    std::size_t cache_index, std::int32_t visible,
                                                                    std::size_t* out_index) {
  clear_last_error();
  if (out_index == nullptr) {
    return null_argument("out_index");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  // A shared item is stored as a 32-bit cache index; a wider one names no item.
  if (cache_index > std::numeric_limits<std::uint32_t>::max()) {
    return check_limit(cache_index, static_cast<std::size_t>(std::numeric_limits<std::uint32_t>::max()) + 1U,
                       "cache_index");
  }
  std::size_t added = 0;
  const fm_status_t status = mutate_pivot_field(
      wb, sheet_index, pivot_index, field_index,
      [cache_index = static_cast<std::uint32_t>(cache_index), visible, &added](formulon::pivot::PivotField& field) {
        // The label is left empty: evaluation derives it from the bound shared item
        // (`pivot_item_label`). A binding that resolves to a blank value has no
        // label, which is what makes this the constructor for the blank item.
        formulon::pivot::PivotItem item;
        item.visible = visible != 0;
        item.has_cache_index = true;
        item.cache_index = cache_index;
        field.items.push_back(std::move(item));
        added = field.items.size() - 1U;
      });
  if (status == 0) {
    *out_index = added;
  }
  return status;
}

extern "C" fm_status_t fm_sheet_clear_pivot_field_items(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::size_t field_index) {
  clear_last_error();
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_index, [](formulon::pivot::PivotField& field) {
    field.items.clear();
    field.marker_items.clear();
  });
}

extern "C" fm_status_t fm_sheet_set_pivot_field_subtotal_functions(fm_workbook_t* wb, std::size_t sheet_index,
                                                                   std::size_t pivot_index, std::size_t field_index,
                                                                   const std::int32_t* aggregations,
                                                                   std::size_t count) {
  clear_last_error();
  if (wb == nullptr || (aggregations == nullptr && count > 0)) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  std::vector<formulon::pivot::SubtotalFn> functions;
  functions.reserve(count);
  std::uint32_t seen = 0;
  for (std::size_t i = 0; i < count; ++i) {
    if (auto rc = check_enum_domain(aggregations[i], 0, FM_PIVOT_AGGREGATION_VARP, "aggregations"); rc != 0) {
      return rc;
    }
    const std::uint32_t bit = 1U << static_cast<unsigned>(aggregations[i]);
    if ((seen & bit) != 0U) {
      return invalid_argument("aggregations lists one function twice", "index=" + std::to_string(i));
    }
    seen |= bit;
    functions.push_back(pivot_subtotal_from_fm(static_cast<fm_pivot_aggregation_t>(aggregations[i])));
  }
  return mutate_pivot_field(wb, sheet_index, pivot_index, field_index,
                            [&functions](formulon::pivot::PivotField& field) { field.subtotal_fns = functions; });
}

extern "C" fm_status_t fm_sheet_set_pivot_field_date_group(fm_workbook_t* wb, std::size_t sheet_index,
                                                           std::size_t pivot_index, std::size_t field_index,
                                                           const fm_pivot_date_group_spec_t* group) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (group == nullptr) {
    return finish_cache_grouping(
        wb, sheet_index, pivot_index,
        formulon::pivot::clear_date_group(wb->workbook(), sheet_index, pivot_index, field_index));
  }
  if (auto rc = check_record_size(*group); rc != 0) {
    return rc;
  }
  if (fm_status_t rc = check_enum_domain(group->calendar, 0, FM_PIVOT_CALENDAR_JAPANESE, "calendar"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(group->has_start, 0, 1, "has_start"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(group->has_end, 0, 1, "has_end"); rc != 0) {
    return rc;
  }
  std::optional<double> start;
  std::optional<double> end;
  if (group->has_start == 1) {
    if (auto rc = check_finite(group->start, "start"); rc != 0) {
      return rc;
    }
    start = group->start;
  }
  if (group->has_end == 1) {
    if (auto rc = check_finite(group->end, "end"); rc != 0) {
      return rc;
    }
    end = group->end;
  }
  return finish_cache_grouping(
      wb, sheet_index, pivot_index,
      formulon::pivot::set_date_group(wb->workbook(), sheet_index, pivot_index, field_index, group->periods,
                                      pivot_calendar_from_fm(static_cast<fm_pivot_calendar_t>(group->calendar)),
                                      group->interval_days, start, end));
}

extern "C" fm_status_t fm_sheet_get_pivot_field_at(const fm_workbook_t* wb, std::size_t sheet_index,
                                                   std::size_t pivot_index, std::size_t field_index,
                                                   fm_pivot_field_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotTable* table = nullptr;
  const formulon::pivot::PivotField* field = nullptr;
  if (auto rc = lookup_pivot_field(wb, sheet_index, pivot_index, field_index, &table, &field); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  fm_pivot_field_info_t result{};
  result.struct_size = out->struct_size;
  result.source_name = field->source_name.c_str();
  result.custom_name = text_or_null(field->custom_name.c_str());
  const auto axis = pivot_axis_to_fm(field->axis);
  result.axis = axis.has_value() ? static_cast<std::int32_t>(*axis) : FM_PIVOT_AXIS_ABSENT;
  result.subtotal_top = field->subtotal_top ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  result.number_format = text_or_null(pivot_num_fmt_code(wb->workbook(), field->number_format));
  const auto* cache = find_cache(wb->workbook(), table->pivot_cache_id());
  result.has_cache_field_index = cache != nullptr && field_index < cache->fields().size() ? 1 : 0;
  result.cache_field_index = result.has_cache_field_index != 0 ? static_cast<std::uint32_t>(field_index) : 0U;
  result.sort_order = field->sort.manual      ? FM_PIVOT_SORT_ORDER_MANUAL
                      : field->sort.ascending ? FM_PIVOT_SORT_ORDER_ASCENDING
                                              : FM_PIVOT_SORT_ORDER_DESCENDING;
  result.sort_by_field = text_or_null(field->sort.by_field.c_str());
  if (result.has_cache_field_index != 0) {
    const bool date1904 = wb->workbook().date1904();
    const std::size_t base = formulon::pivot::group_base_index(*cache, field_index);
    if (base != field_index) {
      result.has_group_base_index = 1;
      result.group_base_index = static_cast<std::uint32_t>(base);
    } else {
      result.date_periods = formulon::pivot::date_group_periods(*cache, field_index, date1904);
      result.date_calendar =
          static_cast<std::int32_t>(formulon::pivot::date_group_calendar(*cache, field_index, date1904));
    }
  }
  // Custom functions replace the automatic subtotal, as the writer and the
  // layout treat them, so the flag reads the effective state.
  result.default_subtotal = field->default_subtotal && field->subtotal_fns.empty() ? 1 : 0;
  for (const formulon::pivot::SubtotalFn fn : field->subtotal_fns) {
    result.subtotal_functions |= 1U << static_cast<unsigned>(fn);
  }
  result.insert_blank_row = field->insert_blank_row ? 1 : 0;
  result.fill_down_labels = field->fill_down_labels ? 1 : 0;
  result.show_all = field->show_all ? 1 : 0;
  result.show_detail = formulon::pivot::field_show_detail(*field) ? 1 : 0;
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_pivot_field_item_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                           std::size_t pivot_index, std::size_t field_index,
                                                           std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotField* field = nullptr;
  if (auto rc = lookup_pivot_field(wb, sheet_index, pivot_index, field_index, nullptr, &field); rc != 0) {
    return rc;
  }
  *out_count = field->items.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_pivot_field_item_at(const fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::size_t field_index,
                                                        std::size_t item_index, fm_pivot_item_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotField* field = nullptr;
  if (auto rc = lookup_pivot_field(wb, sheet_index, pivot_index, field_index, nullptr, &field); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(item_index, field->items.size(), "item_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const formulon::pivot::PivotItem& item = field->items[item_index];
  fm_pivot_item_info_t result{};
  result.struct_size = out->struct_size;
  result.name = text_or_null(item.name.c_str());
  result.has_cache_index = item.has_cache_index ? 1 : 0;
  result.cache_index = item.has_cache_index ? item.cache_index : 0U;
  result.visible = item.visible ? 1 : 0;
  result.caption = item.caption.empty() ? nullptr : item.caption.c_str();
  result.show_detail = item.show_detail ? 1 : 0;
  *out = result;
  return 0;
}
