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
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "c_api/parts/xml_fragment.h"
#include "cf/auto_filter_eval.h"
#include "io/auto_filter_xml.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::ensure_calculated;
using formulon::c_api::parts::filter_operator_from_c;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::FragmentValidation;
using formulon::c_api::parts::fresh_arena;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::null_field;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::resolve_sheet_table;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::to_c;
using formulon::c_api::parts::to_cell_rect;
using formulon::c_api::parts::to_merge_range;
using formulon::c_api::parts::validate_single_element_fragment;

namespace {

static_assert(sizeof(fm_cell_rect_t) == sizeof(formulon::MergeRange), "fm_cell_rect_t / MergeRange size mismatch");

// The AutoFilter an entry point operates on: the sheet's own, or a table's.
struct Target {
  std::size_t sheet = 0;
  std::optional<std::size_t> table;
  const formulon::AutoFilter* filter = nullptr;
};

fm_status_t resolve_sheet(const fm_workbook_t* wb, std::size_t sheet_index, Target* out) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  out->sheet = sheet_index;
  out->filter = wb->workbook().sheet(sheet_index).auto_filter();
  return 0;
}

fm_status_t resolve_table(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t table_index, Target* out) {
  std::size_t global = 0;
  if (auto rc = resolve_sheet_table(wb, sheet_index, table_index, &global); rc != 0) {
    return rc;
  }
  out->sheet = sheet_index;
  out->table = global;
  out->filter = wb->workbook().tables()[global].auto_filter_xml.get();
  return 0;
}

// Rejects an AutoFilter the typed surface cannot operate on: absent, or kept
// opaque because it did not parse into the model.
fm_status_t require_typed(const Target& target) {
  if (target.filter == nullptr) {
    return not_found("no AutoFilter");
  }
  if (target.filter->is_opaque()) {
    return set_binding_error(formulon::FormulonErrorCode::kAutoFilterInvalid,
                             "AutoFilter did not parse into the typed model");
  }
  return 0;
}

// --- model -> record --------------------------------------------------------

// Publishes `filter` into `out`, with every string in the handle's read
// scratch and every array in the handle's AutoFilter arenas.
void publish(const fm_workbook_t* wb, const formulon::AutoFilter& filter, fm_auto_filter_t* out) {
  fm_workbook_t* scratch = const_cast<fm_workbook_t*>(wb);
  fresh_read_scratch(wb);
  fresh_arena(wb, &fm_workbook_t::filter_column_scratch);
  fresh_arena(wb, &fm_workbook_t::filter_value_scratch);
  fresh_arena(wb, &fm_workbook_t::date_group_scratch);
  fresh_arena(wb, &fm_workbook_t::filter_condition_scratch);
  fresh_arena(wb, &fm_workbook_t::sort_condition_scratch);
  auto keep = [scratch](const std::string& text) { return store_cstr(scratch->read_scratch, text); };

  std::vector<fm_filter_column_t> columns;
  columns.reserve(filter.columns.size());
  for (const formulon::FilterColumn& in : filter.columns) {
    fm_filter_column_t c{};
    c.column_offset = in.col_id;
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
        std::vector<fm_date_group_item_t> groups;
        groups.reserve(in.values.date_groups.size());
        for (const formulon::DateGroupItem& g : in.values.date_groups) {
          groups.push_back(fm_date_group_item_t{g.year, g.month, g.day, g.hour, g.minute, g.second,
                                                static_cast<int32_t>(g.grouping)});
        }
        c.date_group_count = static_cast<uint32_t>(groups.size());
        c.date_groups = scratch->date_group_scratch.adopt(std::move(groups));
        break;
      }
      case formulon::FilterKind::kCustom:
        c.custom_and = in.custom.and_join ? 1 : 0;
        {
          std::vector<fm_filter_condition_t> conditions;
          conditions.reserve(in.custom.filters.size());
          for (const formulon::CustomFilter& f : in.custom.filters) {
            conditions.push_back(fm_filter_condition_t{to_c(f.op), keep(f.val)});
          }
          c.condition_count = static_cast<uint32_t>(conditions.size());
          c.conditions = scratch->filter_condition_scratch.adopt(std::move(conditions));
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
        c.has_dxf_index = in.color.dxf_id.has_value() ? 1 : 0;
        c.dxf_index = in.color.dxf_id.value_or(0U);
        c.cell_color = in.color.cell_color ? 1 : 0;
        break;
      case formulon::FilterKind::kIcon:
        c.icon_set = static_cast<int32_t>(in.icon.icon_set);
        c.has_icon_index = in.icon.icon_id.has_value() ? 1 : 0;
        c.icon_index = static_cast<int32_t>(in.icon.icon_id.value_or(0U));
        break;
      case formulon::FilterKind::kNone:
        break;
    }
    columns.push_back(c);
  }

  const std::uint32_t struct_size = out->struct_size;
  *out = FM_STRUCT_INIT(fm_auto_filter_t);
  out->struct_size = struct_size;
  out->range = to_cell_rect(filter.range);
  out->column_count = static_cast<uint32_t>(columns.size());
  out->columns = scratch->filter_column_scratch.adopt(std::move(columns));
  if (filter.sort.has_value()) {
    const formulon::SortState& sort = *filter.sort;
    out->has_sort = 1;
    out->sort_ref = to_cell_rect(sort.ref);
    out->column_sort = sort.column_sort ? 1 : 0;
    out->case_sensitive = sort.case_sensitive ? 1 : 0;
    out->sort_method = static_cast<int32_t>(sort.sort_method);
    std::vector<fm_sort_condition_t> conditions;
    conditions.reserve(sort.conditions.size());
    for (const formulon::SortCondition& in : sort.conditions) {
      fm_sort_condition_t c{};
      c.ref = to_cell_rect(in.ref);
      c.descending = in.descending ? 1 : 0;
      c.sort_by = static_cast<int32_t>(in.sort_by);
      c.custom_list = keep(in.custom_list);
      c.has_dxf_index = in.dxf_id.has_value() ? 1 : 0;
      c.dxf_index = in.dxf_id.value_or(0U);
      c.icon_set = static_cast<int32_t>(in.icon_set);
      c.has_icon_index = in.icon_id.has_value() ? 1 : 0;
      c.icon_index = static_cast<int32_t>(in.icon_id.value_or(0U));
      conditions.push_back(c);
    }
    out->condition_count = static_cast<uint32_t>(conditions.size());
    out->conditions = scratch->sort_condition_scratch.adopt(std::move(conditions));
  }
}

// --- record -> model --------------------------------------------------------

fm_status_t check_array(const void* data, uint32_t count, const char* field) {
  if (data == nullptr && count != 0U) {
    return null_field(field);
  }
  return 0;
}

fm_status_t column_from_fm(const fm_filter_column_t& in, formulon::FilterColumn* out) {
  if (auto rc = check_enum_domain(in.kind, 0, static_cast<int64_t>(formulon::FilterKind::kIcon), "kind"); rc != 0) {
    return rc;
  }
  out->col_id = in.column_offset;
  out->hidden_button = in.hidden_button != 0;
  out->show_button = in.show_button != 0;
  out->kind = static_cast<formulon::FilterKind>(in.kind);
  switch (out->kind) {
    case formulon::FilterKind::kValues: {
      if (auto rc = check_array(in.values, in.value_count, "values"); rc != 0) {
        return rc;
      }
      if (auto rc = check_array(in.date_groups, in.date_group_count, "date_groups"); rc != 0) {
        return rc;
      }
      out->values.blank = in.filter_blank != 0;
      for (uint32_t i = 0; i < in.value_count; ++i) {
        out->values.values.push_back(text_or_empty(in.values[i]));
      }
      for (uint32_t i = 0; i < in.date_group_count; ++i) {
        const fm_date_group_item_t& g = in.date_groups[i];
        if (auto rc = check_enum_domain(g.grouping, 0, static_cast<int64_t>(formulon::DateTimeGrouping::kSecond),
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
      if (auto rc = check_array(in.conditions, in.condition_count, "conditions"); rc != 0) {
        return rc;
      }
      if (in.condition_count < 1U || in.condition_count > 2U) {
        return set_binding_error(formulon::FormulonErrorCode::kAutoFilterInvalid,
                                 "custom filter needs one or two conditions",
                                 "condition_count=" + std::to_string(in.condition_count));
      }
      // A condition always states a relational operator; the range members are not one.
      constexpr int64_t kFirstOp = FM_COMPARISON_OPERATOR_EQUAL;
      constexpr int64_t kLastOp = FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL;
      out->custom.and_join = in.custom_and != 0;
      for (uint32_t i = 0; i < in.condition_count; ++i) {
        if (auto rc = check_enum_domain(in.conditions[i].op, kFirstOp, kLastOp, "conditions.op"); rc != 0) {
          return rc;
        }
        out->custom.filters.push_back(
            formulon::CustomFilter{filter_operator_from_c(in.conditions[i].op), text_or_empty(in.conditions[i].value)});
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
      if (auto rc = check_enum_domain(in.dynamic_type, 0, static_cast<int64_t>(formulon::DynamicFilterType::kM12),
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
      if (in.has_dxf_index != 0) {
        out->color.dxf_id = in.dxf_index;
      }
      out->color.cell_color = in.cell_color != 0;
      return 0;
    case formulon::FilterKind::kIcon:
      if (auto rc = check_enum_domain(in.icon_set, 0, static_cast<int64_t>(formulon::cf::IconSetName::Five_Quarters),
                                      "icon_set");
          rc != 0) {
        return rc;
      }
      out->icon.icon_set = static_cast<formulon::cf::IconSetName>(in.icon_set);
      if (in.has_icon_index != 0) {
        if (auto rc = check_enum_domain(in.icon_index, 0, std::numeric_limits<int32_t>::max(), "icon_id"); rc != 0) {
          return rc;
        }
        out->icon.icon_id = static_cast<uint32_t>(in.icon_index);
      }
      return 0;
    case formulon::FilterKind::kNone:
      return 0;
  }
  return 0;
}

fm_status_t sort_condition_from_fm(const fm_sort_condition_t& in, formulon::SortCondition* out) {
  if (auto rc = check_enum_domain(in.sort_by, 0, static_cast<int64_t>(formulon::SortBy::kIcon), "sort_by"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(in.icon_set, 0, static_cast<int64_t>(formulon::cf::IconSetName::Five_Quarters),
                                  "conditions.icon_set");
      rc != 0) {
    return rc;
  }
  out->ref = to_merge_range(in.ref);
  out->descending = in.descending != 0;
  out->sort_by = static_cast<formulon::SortBy>(in.sort_by);
  out->custom_list = text_or_empty(in.custom_list);
  if (in.has_dxf_index != 0) {
    out->dxf_id = in.dxf_index;
  }
  out->icon_set = static_cast<formulon::cf::IconSetName>(in.icon_set);
  if (in.has_icon_index != 0) {
    if (auto rc = check_enum_domain(in.icon_index, 0, std::numeric_limits<int32_t>::max(), "conditions.icon_id");
        rc != 0) {
      return rc;
    }
    out->icon_id = static_cast<uint32_t>(in.icon_index);
  }
  return 0;
}

fm_status_t filter_from_fm(const fm_auto_filter_t& in, formulon::AutoFilter* out) {
  if (auto rc = check_array(in.columns, in.column_count, "columns"); rc != 0) {
    return rc;
  }
  out->range = to_merge_range(in.range);
  for (uint32_t i = 0; i < in.column_count; ++i) {
    formulon::FilterColumn column;
    if (auto rc = column_from_fm(in.columns[i], &column); rc != 0) {
      return rc;
    }
    out->columns.push_back(std::move(column));
  }
  if (in.has_sort == 0) {
    return 0;
  }
  if (auto rc = check_array(in.conditions, in.condition_count, "conditions"); rc != 0) {
    return rc;
  }
  if (auto rc =
          check_enum_domain(in.sort_method, 0, static_cast<int64_t>(formulon::SortMethod::kStroke), "sort_method");
      rc != 0) {
    return rc;
  }
  formulon::SortState sort;
  sort.ref = to_merge_range(in.sort_ref);
  sort.column_sort = in.column_sort != 0;
  sort.case_sensitive = in.case_sensitive != 0;
  sort.sort_method = static_cast<formulon::SortMethod>(in.sort_method);
  for (uint32_t i = 0; i < in.condition_count; ++i) {
    formulon::SortCondition condition;
    if (auto rc = sort_condition_from_fm(in.conditions[i], &condition); rc != 0) {
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

fm_status_t get_body(const fm_workbook_t* wb, const Target& target, int32_t* out_present, fm_auto_filter_t* out) {
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (target.filter == nullptr) {
    const std::uint32_t struct_size = out->struct_size;
    *out = FM_STRUCT_INIT(fm_auto_filter_t);
    out->struct_size = struct_size;
    *out_present = 0;
    return 0;
  }
  if (auto rc = require_typed(target); rc != 0) {
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
  return finish_status(result);
}

fm_status_t set_body(fm_workbook_t* wb, const Target& target, const fm_auto_filter_t* filter) {
  if (auto rc = check_record_size(*filter); rc != 0) {
    return rc;
  }
  formulon::AutoFilter model;
  if (auto rc = filter_from_fm(*filter, &model); rc != 0) {
    return rc;
  }
  if (target.filter != nullptr && !target.filter->is_opaque()) {
    carry_unmodelled(*target.filter, &model);
  }
  return store(wb, target, std::move(model));
}

fm_status_t remove_body(fm_workbook_t* wb, const Target& target) {
  if (target.filter == nullptr) {
    return 0;
  }
  formulon::Workbook& book = wb->workbook();
  auto result = target.table.has_value() ? book.remove_table_auto_filter(*target.table)
                                         : book.remove_sheet_auto_filter(target.sheet);
  return finish_status(result);
}

fm_status_t apply_body(fm_workbook_t* wb, const Target& target) {
  if (auto rc = require_typed(target); rc != 0) {
    return rc;
  }
  // Copied: the row edits must not read through storage the workbook owns.
  const formulon::AutoFilter filter = *target.filter;
  auto result = formulon::apply_auto_filter(wb->workbook(), target.sheet, filter);
  return finish_status(result);
}

fm_status_t clear_body(fm_workbook_t* wb, const Target& target) {
  if (auto rc = require_typed(target); rc != 0) {
    return rc;
  }
  formulon::AutoFilter filter = *target.filter;
  drop_criteria(filter);
  if (auto rc = store(wb, target, filter); rc != 0) {
    return rc;
  }
  auto result = formulon::apply_auto_filter(wb->workbook(), target.sheet, filter);
  return finish_status(result);
}

fm_status_t evaluate_body(const fm_workbook_t* wb, const Target& target, uint8_t* out, size_t capacity,
                          size_t* out_count, uint32_t* out_first_row) {
  if (auto rc = require_typed(target); rc != 0) {
    return rc;
  }
  const formulon::Workbook& book = wb->workbook();
  auto matches = formulon::evaluate_auto_filter(book, book.sheet(target.sheet), *target.filter);
  if (!matches) {
    return set_last_error(matches.error());
  }
  const std::vector<bool>& flags = matches.value();
  const std::size_t written = std::min(capacity, flags.size());
  for (std::size_t i = 0; i < written; ++i) {
    out[i] = flags[i] ? 1U : 0U;
  }
  *out_count = flags.size();
  *out_first_row = target.filter->range.first_row + 1U;
  return 0;
}

bool evaluate_args_missing(const uint8_t* out, size_t capacity, const size_t* out_count,
                           const uint32_t* out_first_row) {
  return out_count == nullptr || out_first_row == nullptr || (out == nullptr && capacity > 0U);
}

}  // namespace

// --- sheet ------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_get_auto_filter(const fm_workbook_t* wb, size_t sheet_index, int32_t* out_present,
                                                fm_auto_filter_t* out) {
  clear_last_error();
  if (out == nullptr || out_present == nullptr) {
    return null_argument("out");
  }
  Target target;
  if (auto rc = resolve_sheet(wb, sheet_index, &target); rc != 0) {
    return rc;
  }
  return get_body(wb, target, out_present, out);
}

extern "C" fm_status_t fm_sheet_set_auto_filter(fm_workbook_t* wb, size_t sheet_index, const fm_auto_filter_t* filter) {
  clear_last_error();
  Target target;
  if (auto rc = resolve_sheet(wb, sheet_index, &target); rc != 0) {
    return rc;
  }
  return filter == nullptr ? remove_body(wb, target) : set_body(wb, target, filter);
}

extern "C" fm_status_t fm_sheet_apply_auto_filter(fm_workbook_t* wb, size_t sheet_index) {
  clear_last_error();
  Target target;
  if (auto rc = resolve_sheet(wb, sheet_index, &target); rc != 0) {
    return rc;
  }
  return apply_body(wb, target);
}

extern "C" fm_status_t fm_sheet_clear_auto_filter_criteria(fm_workbook_t* wb, size_t sheet_index) {
  clear_last_error();
  Target target;
  if (auto rc = resolve_sheet(wb, sheet_index, &target); rc != 0) {
    return rc;
  }
  return clear_body(wb, target);
}

extern "C" fm_status_t fm_sheet_evaluate_auto_filter(const fm_workbook_t* wb, size_t sheet_index, uint8_t* out,
                                                     size_t capacity, size_t* out_count, uint32_t* out_first_row) {
  clear_last_error();
  if (evaluate_args_missing(out, capacity, out_count, out_first_row)) {
    return null_argument();
  }
  Target target;
  if (auto rc = resolve_sheet(wb, sheet_index, &target); rc != 0) {
    return rc;
  }
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  return evaluate_body(wb, target, out, capacity, out_count, out_first_row);
}

// --- table ------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_get_table_auto_filter(const fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                      int32_t* out_present, fm_auto_filter_t* out) {
  clear_last_error();
  if (out == nullptr || out_present == nullptr) {
    return null_argument("out");
  }
  Target target;
  if (auto rc = resolve_table(wb, sheet_index, table_index, &target); rc != 0) {
    return rc;
  }
  return get_body(wb, target, out_present, out);
}

extern "C" fm_status_t fm_sheet_set_table_auto_filter(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                                      const fm_auto_filter_t* filter) {
  clear_last_error();
  Target target;
  if (auto rc = resolve_table(wb, sheet_index, table_index, &target); rc != 0) {
    return rc;
  }
  return filter == nullptr ? remove_body(wb, target) : set_body(wb, target, filter);
}

extern "C" fm_status_t fm_sheet_apply_table_auto_filter(fm_workbook_t* wb, size_t sheet_index, size_t table_index) {
  clear_last_error();
  Target target;
  if (auto rc = resolve_table(wb, sheet_index, table_index, &target); rc != 0) {
    return rc;
  }
  return apply_body(wb, target);
}

extern "C" fm_status_t fm_sheet_clear_table_auto_filter_criteria(fm_workbook_t* wb, size_t sheet_index,
                                                                 size_t table_index) {
  clear_last_error();
  Target target;
  if (auto rc = resolve_table(wb, sheet_index, table_index, &target); rc != 0) {
    return rc;
  }
  return clear_body(wb, target);
}

extern "C" fm_status_t fm_sheet_evaluate_table_auto_filter(const fm_workbook_t* wb, size_t sheet_index,
                                                           size_t table_index, uint8_t* out, size_t capacity,
                                                           size_t* out_count, uint32_t* out_first_row) {
  clear_last_error();
  if (evaluate_args_missing(out, capacity, out_count, out_first_row)) {
    return null_argument();
  }
  Target target;
  if (auto rc = resolve_table(wb, sheet_index, table_index, &target); rc != 0) {
    return rc;
  }
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  return evaluate_body(wb, target, out, capacity, out_count, out_first_row);
}

extern "C" fm_status_t fm_sheet_get_auto_filter_xml(const fm_workbook_t* wb, size_t sheet_index, const char** out_xml) {
  clear_last_error();
  if (out_xml == nullptr) {
    return null_argument("out_xml");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_xml = publish_scratch(wb, formulon::io::auto_filter_xml(wb->workbook().sheet(sheet_index).auto_filter()));
  return 0;
}

extern "C" fm_status_t fm_sheet_set_auto_filter_xml(fm_workbook_t* wb, size_t sheet_index, const char* xml) {
  clear_last_error();
  if (xml == nullptr) {
    return null_argument("xml");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const std::string fragment(xml);
  if (!fragment.empty()) {
    if (auto rc = check_utf8(fragment, "element=autoFilter"); rc != 0) {
      return rc;
    }
    const FragmentValidation validation = validate_single_element_fragment(fragment, "autoFilter");
    if (!validation.valid) {
      return invalid_argument("xml is not a valid autoFilter element", validation.context);
    }
  }
  // Routed through the workbook so `_FilterDatabase` and SUBTOTAL/AGGREGATE
  // dirty state follow the change.
  if (auto result = wb->workbook().set_sheet_auto_filter_xml(sheet_index, fragment); !result) {
    return set_last_error(result.error());
  }
  return 0;
}
