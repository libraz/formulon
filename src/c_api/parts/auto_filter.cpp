//
// C ABI - typed sheet and table AutoFilter: read, replace, remove, and the
// apply / clear / evaluate operations over the filtered rows.

#include "auto_filter.h"

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
#include "cf/auto_filter_eval.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

namespace {

static_assert(sizeof(fm_merge_range) == sizeof(formulon::MergeRange), "fm_merge_range / MergeRange size mismatch");

// `fm_filter_column::dxf_id` of a colour filter that names no dxf.
constexpr std::uint32_t kAbsentDxfId = std::numeric_limits<std::uint32_t>::max();

// The AutoFilter an entry point operates on: the sheet's own, or a table's.
struct Target {
  std::size_t sheet = 0;
  std::optional<std::size_t> table;
  const formulon::AutoFilter* filter = nullptr;
};

fm_status_t resolve_sheet(const fm_workbook_t* wb, std::size_t sheet_index, const char* api, Target* out) {
  if (auto rc = check_sheet_index(wb, sheet_index, api); rc != 0) {
    return rc;
  }
  out->sheet = sheet_index;
  out->filter = wb->workbook().sheet(sheet_index).auto_filter();
  return 0;
}

fm_status_t resolve_table(const fm_workbook_t* wb, std::size_t table_index, const char* api, Target* out) {
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, api, "wb=NULL");
  }
  const auto& tables = wb->workbook().tables();
  if (table_index >= tables.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(api) + ": table index out of range").c_str(),
                             "table_index=" + std::to_string(table_index) + " count=" + std::to_string(tables.size()));
  }
  out->sheet = tables[table_index].sheet_index;
  out->table = table_index;
  out->filter = tables[table_index].auto_filter_xml.get();
  return 0;
}

// Rejects an AutoFilter the typed surface cannot operate on: absent, or kept
// opaque because it did not parse into the model.
fm_status_t require_typed(const Target& target, const char* api) {
  if (target.filter == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kNotFound, (std::string(api) + ": no AutoFilter").c_str());
  }
  if (target.filter->is_opaque()) {
    return set_binding_error(formulon::FormulonErrorCode::kAutoFilterInvalid,
                             (std::string(api) + ": AutoFilter did not parse into the typed model").c_str());
  }
  return 0;
}

fm_merge_range to_fm(const formulon::MergeRange& r) {
  return fm_merge_range{r.first_row, r.first_col, r.last_row, r.last_col};
}

formulon::MergeRange from_fm(const fm_merge_range& r) {
  formulon::MergeRange out;
  out.first_row = r.first_row;
  out.first_col = r.first_col;
  out.last_row = r.last_row;
  out.last_col = r.last_col;
  return out;
}

std::string text_or_empty(const char* text) {
  return text != nullptr ? std::string(text) : std::string();
}

// --- model -> record --------------------------------------------------------

// Publishes `filter` into `out`, with every string in the handle's read
// scratch and every array in the handle's AutoFilter arenas.
void publish(const fm_workbook_t* wb, const formulon::AutoFilter& filter, fm_auto_filter* out) {
  fm_workbook_t* scratch = const_cast<fm_workbook_t*>(wb);
  scratch->read_scratch.clear();
  scratch->filter_column_scratch.clear();
  scratch->filter_value_scratch.clear();
  scratch->date_group_scratch.clear();
  scratch->sort_condition_scratch.clear();
  auto keep = [scratch](const std::string& text) {
    scratch->read_scratch.emplace_back(text);
    return scratch->read_scratch.back().c_str();
  };

  std::vector<fm_filter_column> columns;
  columns.reserve(filter.columns.size());
  for (const formulon::FilterColumn& in : filter.columns) {
    fm_filter_column c{};
    c.col_id = in.col_id;
    c.hidden_button = in.hidden_button ? 1 : 0;
    c.show_button = in.show_button ? 1 : 0;
    c.kind = static_cast<int32_t>(in.kind);
    switch (in.kind) {
      case formulon::FilterKind::kValues: {
        c.filter_blank = in.values.blank ? 1 : 0;
        std::vector<const char*> values;
        values.reserve(in.values.values.size());
        for (const std::string& v : in.values.values) {
          values.push_back(keep(v));
        }
        c.value_count = static_cast<uint32_t>(values.size());
        c.values = scratch->filter_value_scratch.adopt(std::move(values));
        std::vector<fm_date_group_item> groups;
        groups.reserve(in.values.date_groups.size());
        for (const formulon::DateGroupItem& g : in.values.date_groups) {
          groups.push_back(
              fm_date_group_item{g.year, g.month, g.day, g.hour, g.minute, g.second, static_cast<uint8_t>(g.grouping)});
        }
        c.date_group_count = static_cast<uint32_t>(groups.size());
        c.date_groups = scratch->date_group_scratch.adopt(std::move(groups));
        break;
      }
      case formulon::FilterKind::kCustom:
        c.custom_and = in.custom.and_join ? 1 : 0;
        c.custom_count = static_cast<int32_t>(in.custom.filters.size());
        if (!in.custom.filters.empty()) {
          c.op1 = static_cast<int32_t>(in.custom.filters[0].op);
          c.val1 = keep(in.custom.filters[0].val);
        }
        if (in.custom.filters.size() > 1U) {
          c.op2 = static_cast<int32_t>(in.custom.filters[1].op);
          c.val2 = keep(in.custom.filters[1].val);
        }
        break;
      case formulon::FilterKind::kTop10:
        c.top = in.top10.top ? 1 : 0;
        c.percent = in.top10.percent ? 1 : 0;
        c.top_val = in.top10.val;
        c.has_filter_val = in.top10.filter_val.has_value() ? 1 : 0;
        c.filter_val = in.top10.filter_val.value_or(0.0);
        break;
      case formulon::FilterKind::kDynamic:
        c.dynamic_type = static_cast<int32_t>(in.dynamic.type);
        c.has_dyn_val = in.dynamic.val.has_value() ? 1 : 0;
        c.dyn_val = in.dynamic.val.value_or(0.0);
        c.has_dyn_max_val = in.dynamic.max_val.has_value() ? 1 : 0;
        c.dyn_max_val = in.dynamic.max_val.value_or(0.0);
        c.val_iso = keep(in.dynamic.val_iso);
        c.max_val_iso = keep(in.dynamic.max_val_iso);
        break;
      case formulon::FilterKind::kColor:
        c.dxf_id = in.color.dxf_id.value_or(kAbsentDxfId);
        c.cell_color = in.color.cell_color ? 1 : 0;
        break;
      case formulon::FilterKind::kIcon:
        c.icon_set = static_cast<int32_t>(in.icon.icon_set);
        c.has_icon_id = in.icon.icon_id.has_value() ? 1 : 0;
        c.icon_id = static_cast<int32_t>(in.icon.icon_id.value_or(0U));
        break;
      case formulon::FilterKind::kNone:
        break;
    }
    columns.push_back(c);
  }

  *out = fm_auto_filter{};
  out->range = to_fm(filter.range);
  out->column_count = static_cast<uint32_t>(columns.size());
  out->columns = scratch->filter_column_scratch.adopt(std::move(columns));
  if (filter.sort.has_value()) {
    const formulon::SortState& sort = *filter.sort;
    out->has_sort = 1;
    out->sort_ref = to_fm(sort.ref);
    out->column_sort = sort.column_sort ? 1 : 0;
    out->case_sensitive = sort.case_sensitive ? 1 : 0;
    out->sort_method = static_cast<int32_t>(sort.sort_method);
    std::vector<fm_sort_condition> conditions;
    conditions.reserve(sort.conditions.size());
    for (const formulon::SortCondition& in : sort.conditions) {
      fm_sort_condition c{};
      c.ref = to_fm(in.ref);
      c.descending = in.descending ? 1 : 0;
      c.sort_by = static_cast<int32_t>(in.sort_by);
      c.custom_list = keep(in.custom_list);
      c.has_dxf_id = in.dxf_id.has_value() ? 1 : 0;
      c.dxf_id = in.dxf_id.value_or(0U);
      c.icon_set = static_cast<int32_t>(in.icon_set);
      c.has_icon_id = in.icon_id.has_value() ? 1 : 0;
      c.icon_id = static_cast<int32_t>(in.icon_id.value_or(0U));
      conditions.push_back(c);
    }
    out->condition_count = static_cast<uint32_t>(conditions.size());
    out->conditions = scratch->sort_condition_scratch.adopt(std::move(conditions));
  }
}

// --- record -> model --------------------------------------------------------

fm_status_t check_array(const void* data, uint32_t count, const char* api, const char* field) {
  if (data == nullptr && count != 0U) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(api) + ": NULL array with a non-zero count").c_str(),
                             std::string("field=") + field + " count=" + std::to_string(count));
  }
  return 0;
}

fm_status_t column_from_fm(const fm_filter_column& in, const char* api, formulon::FilterColumn* out) {
  if (auto rc = check_enum_domain(in.kind, static_cast<int64_t>(formulon::FilterKind::kIcon), api, "kind"); rc != 0) {
    return rc;
  }
  out->col_id = in.col_id;
  out->hidden_button = in.hidden_button != 0;
  out->show_button = in.show_button != 0;
  out->kind = static_cast<formulon::FilterKind>(in.kind);
  switch (out->kind) {
    case formulon::FilterKind::kValues: {
      if (auto rc = check_array(in.values, in.value_count, api, "values"); rc != 0) {
        return rc;
      }
      if (auto rc = check_array(in.date_groups, in.date_group_count, api, "date_groups"); rc != 0) {
        return rc;
      }
      out->values.blank = in.filter_blank != 0;
      for (uint32_t i = 0; i < in.value_count; ++i) {
        out->values.values.push_back(text_or_empty(in.values[i]));
      }
      for (uint32_t i = 0; i < in.date_group_count; ++i) {
        const fm_date_group_item& g = in.date_groups[i];
        if (auto rc = check_enum_domain(g.grouping, static_cast<int64_t>(formulon::DateTimeGrouping::kSecond), api,
                                        "date_groups.grouping");
            rc != 0) {
          return rc;
        }
        formulon::DateGroupItem item;
        item.year = g.year;
        item.month = g.month;
        item.day = g.day;
        item.hour = g.hour;
        item.minute = g.minute;
        item.second = g.second;
        item.grouping = static_cast<formulon::DateTimeGrouping>(g.grouping);
        out->values.date_groups.push_back(item);
      }
      return 0;
    }
    case formulon::FilterKind::kCustom: {
      if (in.custom_count < 0 || in.custom_count > 2) {
        return set_binding_error(formulon::FormulonErrorCode::kAutoFilterInvalid,
                                 (std::string(api) + ": custom filter needs one or two conditions").c_str(),
                                 "custom_count=" + std::to_string(in.custom_count));
      }
      constexpr auto kLastOp = static_cast<int64_t>(formulon::FilterOperator::kGreaterThan);
      out->custom.and_join = in.custom_and != 0;
      if (in.custom_count >= 1) {
        if (auto rc = check_enum_domain(in.op1, kLastOp, api, "op1"); rc != 0) {
          return rc;
        }
        out->custom.filters.push_back(
            formulon::CustomFilter{static_cast<formulon::FilterOperator>(in.op1), text_or_empty(in.val1)});
      }
      if (in.custom_count == 2) {
        if (auto rc = check_enum_domain(in.op2, kLastOp, api, "op2"); rc != 0) {
          return rc;
        }
        out->custom.filters.push_back(
            formulon::CustomFilter{static_cast<formulon::FilterOperator>(in.op2), text_or_empty(in.val2)});
      }
      return 0;
    }
    case formulon::FilterKind::kTop10:
      out->top10.top = in.top != 0;
      out->top10.percent = in.percent != 0;
      out->top10.val = in.top_val;
      if (in.has_filter_val != 0) {
        out->top10.filter_val = in.filter_val;
      }
      return 0;
    case formulon::FilterKind::kDynamic:
      if (auto rc = check_enum_domain(in.dynamic_type, static_cast<int64_t>(formulon::DynamicFilterType::kM12), api,
                                      "dynamic_type");
          rc != 0) {
        return rc;
      }
      out->dynamic.type = static_cast<formulon::DynamicFilterType>(in.dynamic_type);
      if (in.has_dyn_val != 0) {
        out->dynamic.val = in.dyn_val;
      }
      if (in.has_dyn_max_val != 0) {
        out->dynamic.max_val = in.dyn_max_val;
      }
      out->dynamic.val_iso = text_or_empty(in.val_iso);
      out->dynamic.max_val_iso = text_or_empty(in.max_val_iso);
      return 0;
    case formulon::FilterKind::kColor:
      if (in.dxf_id != kAbsentDxfId) {
        out->color.dxf_id = in.dxf_id;
      }
      out->color.cell_color = in.cell_color != 0;
      return 0;
    case formulon::FilterKind::kIcon:
      if (auto rc = check_enum_domain(in.icon_set, static_cast<int64_t>(formulon::cf::IconSetName::Five_Quarters), api,
                                      "icon_set");
          rc != 0) {
        return rc;
      }
      out->icon.icon_set = static_cast<formulon::cf::IconSetName>(in.icon_set);
      if (in.has_icon_id != 0) {
        if (auto rc = check_enum_domain(in.icon_id, std::numeric_limits<int32_t>::max(), api, "icon_id"); rc != 0) {
          return rc;
        }
        out->icon.icon_id = static_cast<uint32_t>(in.icon_id);
      }
      return 0;
    case formulon::FilterKind::kNone:
      return 0;
  }
  return 0;
}

fm_status_t sort_condition_from_fm(const fm_sort_condition& in, const char* api, formulon::SortCondition* out) {
  if (auto rc = check_enum_domain(in.sort_by, static_cast<int64_t>(formulon::SortBy::kIcon), api, "sort_by"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(in.icon_set, static_cast<int64_t>(formulon::cf::IconSetName::Five_Quarters), api,
                                  "conditions.icon_set");
      rc != 0) {
    return rc;
  }
  out->ref = from_fm(in.ref);
  out->descending = in.descending != 0;
  out->sort_by = static_cast<formulon::SortBy>(in.sort_by);
  out->custom_list = text_or_empty(in.custom_list);
  if (in.has_dxf_id != 0) {
    out->dxf_id = in.dxf_id;
  }
  out->icon_set = static_cast<formulon::cf::IconSetName>(in.icon_set);
  if (in.has_icon_id != 0) {
    if (auto rc = check_enum_domain(in.icon_id, std::numeric_limits<int32_t>::max(), api, "conditions.icon_id");
        rc != 0) {
      return rc;
    }
    out->icon_id = static_cast<uint32_t>(in.icon_id);
  }
  return 0;
}

fm_status_t filter_from_fm(const fm_auto_filter& in, const char* api, formulon::AutoFilter* out) {
  if (auto rc = check_array(in.columns, in.column_count, api, "columns"); rc != 0) {
    return rc;
  }
  out->range = from_fm(in.range);
  for (uint32_t i = 0; i < in.column_count; ++i) {
    formulon::FilterColumn column;
    if (auto rc = column_from_fm(in.columns[i], api, &column); rc != 0) {
      return rc;
    }
    out->columns.push_back(std::move(column));
  }
  if (in.has_sort == 0) {
    return 0;
  }
  if (auto rc = check_array(in.conditions, in.condition_count, api, "conditions"); rc != 0) {
    return rc;
  }
  if (auto rc =
          check_enum_domain(in.sort_method, static_cast<int64_t>(formulon::SortMethod::kStroke), api, "sort_method");
      rc != 0) {
    return rc;
  }
  formulon::SortState sort;
  sort.ref = from_fm(in.sort_ref);
  sort.column_sort = in.column_sort != 0;
  sort.case_sensitive = in.case_sensitive != 0;
  sort.sort_method = static_cast<formulon::SortMethod>(in.sort_method);
  for (uint32_t i = 0; i < in.condition_count; ++i) {
    formulon::SortCondition condition;
    if (auto rc = sort_condition_from_fm(in.conditions[i], api, &condition); rc != 0) {
      return rc;
    }
    sort.conditions.push_back(std::move(condition));
  }
  out->sort = std::move(sort);
  return 0;
}

// Copies from `existing` what the C records have no slot for: unmodelled
// attributes and extensions of the filter, its sort state and each column
// (matched by `col_id`), and a values filter's `calendarType`.
void carry_unmodelled(const formulon::AutoFilter& existing, formulon::AutoFilter* model) {
  model->extra_attrs = existing.extra_attrs;
  model->ext_lst_xml = existing.ext_lst_xml;
  if (model->sort.has_value() && existing.sort.has_value()) {
    model->sort->extra_attrs = existing.sort->extra_attrs;
    model->sort->ext_lst_xml = existing.sort->ext_lst_xml;
  }
  for (formulon::FilterColumn& column : model->columns) {
    for (const formulon::FilterColumn& old : existing.columns) {
      if (old.col_id != column.col_id) {
        continue;
      }
      column.extra_attrs = old.extra_attrs;
      column.extra_xml = old.extra_xml;
      column.ext_xml = old.ext_xml;
      column.values.calendar_type = old.values.calendar_type;
      break;
    }
  }
}

// Drops every column criterion. A column that still carries a button
// setting or retained markup stays, without a criterion.
void drop_criteria(formulon::AutoFilter& filter) {
  std::vector<formulon::FilterColumn> kept;
  for (formulon::FilterColumn& column : filter.columns) {
    formulon::FilterColumn bare;
    bare.col_id = column.col_id;
    bare.hidden_button = column.hidden_button;
    bare.show_button = column.show_button;
    bare.extra_attrs = std::move(column.extra_attrs);
    bare.extra_xml = std::move(column.extra_xml);
    bare.ext_xml = std::move(column.ext_xml);
    if (bare.hidden_button || !bare.show_button || !bare.extra_attrs.empty() || !bare.extra_xml.empty() ||
        !bare.ext_xml.empty()) {
      kept.push_back(std::move(bare));
    }
  }
  filter.columns = std::move(kept);
}

// --- shared bodies ----------------------------------------------------------

fm_status_t get_body(const fm_workbook_t* wb, const Target& target, fm_auto_filter* out, int32_t* out_present,
                     const char* api) {
  if (target.filter == nullptr) {
    *out = fm_auto_filter{};
    *out_present = 0;
    return 0;
  }
  if (auto rc = require_typed(target, api); rc != 0) {
    return rc;
  }
  publish(wb, *target.filter, out);
  *out_present = 1;
  return 0;
}

fm_status_t store(fm_workbook_t* wb, const Target& target, formulon::AutoFilter filter) {
  formulon::Workbook& book = wb->workbook();
  auto result = target.table.has_value() ? book.set_table_auto_filter(*target.table, std::move(filter))
                                         : book.set_sheet_auto_filter(target.sheet, std::move(filter));
  return result ? 0 : set_last_error(result.error());
}

fm_status_t set_body(fm_workbook_t* wb, const Target& target, const fm_auto_filter* filter, const char* api) {
  if (filter == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(api) + ": NULL filter").c_str());
  }
  formulon::AutoFilter model;
  if (auto rc = filter_from_fm(*filter, api, &model); rc != 0) {
    return rc;
  }
  if (target.filter != nullptr && !target.filter->is_opaque()) {
    carry_unmodelled(*target.filter, &model);
  }
  return store(wb, target, std::move(model));
}

fm_status_t remove_body(fm_workbook_t* wb, const Target& target) {
  formulon::Workbook& book = wb->workbook();
  auto result = target.table.has_value() ? book.remove_table_auto_filter(*target.table)
                                         : book.remove_sheet_auto_filter(target.sheet);
  return result ? 0 : set_last_error(result.error());
}

fm_status_t apply_body(fm_workbook_t* wb, const Target& target, const char* api) {
  if (auto rc = require_typed(target, api); rc != 0) {
    return rc;
  }
  // Copied: the row edits must not read through storage the workbook owns.
  const formulon::AutoFilter filter = *target.filter;
  auto result = formulon::apply_auto_filter(wb->workbook(), target.sheet, filter);
  return result ? 0 : set_last_error(result.error());
}

fm_status_t clear_body(fm_workbook_t* wb, const Target& target, const char* api) {
  if (auto rc = require_typed(target, api); rc != 0) {
    return rc;
  }
  formulon::AutoFilter filter = *target.filter;
  drop_criteria(filter);
  if (auto rc = store(wb, target, filter); rc != 0) {
    return rc;
  }
  auto result = formulon::apply_auto_filter(wb->workbook(), target.sheet, filter);
  return result ? 0 : set_last_error(result.error());
}

fm_status_t evaluate_body(const fm_workbook_t* wb, const Target& target, uint8_t* out_match, size_t cap,
                          size_t* out_len, uint32_t* out_first_row, const char* api) {
  if (auto rc = require_typed(target, api); rc != 0) {
    return rc;
  }
  const formulon::Workbook& book = wb->workbook();
  auto matches = formulon::evaluate_auto_filter(book, book.sheet(target.sheet), *target.filter);
  if (!matches) {
    return set_last_error(matches.error());
  }
  const std::vector<bool>& flags = matches.value();
  const std::size_t written = std::min(cap, flags.size());
  for (std::size_t i = 0; i < written; ++i) {
    out_match[i] = flags[i] ? 1U : 0U;
  }
  *out_len = flags.size();
  *out_first_row = target.filter->range.first_row + 1U;
  return 0;
}

bool evaluate_args_missing(const uint8_t* out_match, size_t cap, const size_t* out_len, const uint32_t* out_first_row) {
  return out_len == nullptr || out_first_row == nullptr || (out_match == nullptr && cap > 0U);
}

}  // namespace

// --- sheet ------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_get_auto_filter(const fm_workbook_t* wb, size_t sheet_index, fm_auto_filter* out,
                                                int32_t* out_present) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_get_auto_filter";
  if (out == nullptr || out_present == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_get_auto_filter: NULL out");
  }
  Target target;
  if (auto rc = resolve_sheet(wb, sheet_index, kApi, &target); rc != 0) {
    return rc;
  }
  return get_body(wb, target, out, out_present, kApi);
}

extern "C" fm_status_t fm_sheet_set_auto_filter(fm_workbook_t* wb, size_t sheet_index, const fm_auto_filter* filter) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_set_auto_filter";
  Target target;
  if (auto rc = resolve_sheet(wb, sheet_index, kApi, &target); rc != 0) {
    return rc;
  }
  return set_body(wb, target, filter, kApi);
}

extern "C" fm_status_t fm_sheet_remove_auto_filter(fm_workbook_t* wb, size_t sheet_index) {
  clear_last_error();
  Target target;
  if (auto rc = resolve_sheet(wb, sheet_index, "fm_sheet_remove_auto_filter", &target); rc != 0) {
    return rc;
  }
  return remove_body(wb, target);
}

extern "C" fm_status_t fm_sheet_apply_auto_filter(fm_workbook_t* wb, size_t sheet_index) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_apply_auto_filter";
  Target target;
  if (auto rc = resolve_sheet(wb, sheet_index, kApi, &target); rc != 0) {
    return rc;
  }
  return apply_body(wb, target, kApi);
}

extern "C" fm_status_t fm_sheet_clear_auto_filter(fm_workbook_t* wb, size_t sheet_index) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_clear_auto_filter";
  Target target;
  if (auto rc = resolve_sheet(wb, sheet_index, kApi, &target); rc != 0) {
    return rc;
  }
  return clear_body(wb, target, kApi);
}

extern "C" fm_status_t fm_sheet_evaluate_auto_filter(const fm_workbook_t* wb, size_t sheet_index, uint8_t* out_match,
                                                     size_t cap, size_t* out_len, uint32_t* out_first_row) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_evaluate_auto_filter";
  if (evaluate_args_missing(out_match, cap, out_len, out_first_row)) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_evaluate_auto_filter: NULL argument");
  }
  Target target;
  if (auto rc = resolve_sheet(wb, sheet_index, kApi, &target); rc != 0) {
    return rc;
  }
  return evaluate_body(wb, target, out_match, cap, out_len, out_first_row, kApi);
}

// --- table ------------------------------------------------------------------

extern "C" fm_status_t fm_table_get_auto_filter(const fm_workbook_t* wb, size_t table_index, fm_auto_filter* out,
                                                int32_t* out_present) {
  clear_last_error();
  constexpr const char* kApi = "fm_table_get_auto_filter";
  if (out == nullptr || out_present == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_table_get_auto_filter: NULL out");
  }
  Target target;
  if (auto rc = resolve_table(wb, table_index, kApi, &target); rc != 0) {
    return rc;
  }
  return get_body(wb, target, out, out_present, kApi);
}

extern "C" fm_status_t fm_table_set_auto_filter(fm_workbook_t* wb, size_t table_index, const fm_auto_filter* filter) {
  clear_last_error();
  constexpr const char* kApi = "fm_table_set_auto_filter";
  Target target;
  if (auto rc = resolve_table(wb, table_index, kApi, &target); rc != 0) {
    return rc;
  }
  return set_body(wb, target, filter, kApi);
}

extern "C" fm_status_t fm_table_remove_auto_filter(fm_workbook_t* wb, size_t table_index) {
  clear_last_error();
  Target target;
  if (auto rc = resolve_table(wb, table_index, "fm_table_remove_auto_filter", &target); rc != 0) {
    return rc;
  }
  return remove_body(wb, target);
}

extern "C" fm_status_t fm_table_apply_auto_filter(fm_workbook_t* wb, size_t table_index) {
  clear_last_error();
  constexpr const char* kApi = "fm_table_apply_auto_filter";
  Target target;
  if (auto rc = resolve_table(wb, table_index, kApi, &target); rc != 0) {
    return rc;
  }
  return apply_body(wb, target, kApi);
}

extern "C" fm_status_t fm_table_clear_auto_filter(fm_workbook_t* wb, size_t table_index) {
  clear_last_error();
  constexpr const char* kApi = "fm_table_clear_auto_filter";
  Target target;
  if (auto rc = resolve_table(wb, table_index, kApi, &target); rc != 0) {
    return rc;
  }
  return clear_body(wb, target, kApi);
}

extern "C" fm_status_t fm_table_evaluate_auto_filter(const fm_workbook_t* wb, size_t table_index, uint8_t* out_match,
                                                     size_t cap, size_t* out_len, uint32_t* out_first_row) {
  clear_last_error();
  constexpr const char* kApi = "fm_table_evaluate_auto_filter";
  if (evaluate_args_missing(out_match, cap, out_len, out_first_row)) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_table_evaluate_auto_filter: NULL argument");
  }
  Target target;
  if (auto rc = resolve_table(wb, table_index, kApi, &target); rc != 0) {
    return rc;
  }
  return evaluate_body(wb, target, out_match, cap, out_len, out_first_row, kApi);
}
