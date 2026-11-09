//
// C ABI - sparkline groups of a sheet's worksheet extension list: reading,
// adding, updating, replacing the sparklines of, and removing groups.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "io/sparkline_groups.h"
#include "utils/error.h"
#include "utils/formula_text.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::color_spec_from_c;
using formulon::c_api::parts::color_spec_to_c;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::sparkline_empty_cells_from_c;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::to_c;
using formulon::c_api::parts::with_formula_prefix;

namespace {

namespace io = formulon::io;

io::SparklineGroup group_from_c(const fm_sparkline_group_t& g) {
  io::SparklineGroup out;
  out.type = static_cast<io::SparklineType>(g.type);
  out.color_series = color_spec_from_c(g.color_series);
  out.color_negative = color_spec_from_c(g.color_negative);
  out.color_axis = color_spec_from_c(g.color_axis);
  out.color_markers = color_spec_from_c(g.color_markers);
  out.color_first = color_spec_from_c(g.color_first);
  out.color_last = color_spec_from_c(g.color_last);
  out.color_high = color_spec_from_c(g.color_high);
  out.color_low = color_spec_from_c(g.color_low);
  out.markers = g.markers != 0;
  out.high = g.high != 0;
  out.low = g.low != 0;
  out.first = g.first != 0;
  out.last = g.last != 0;
  out.negative = g.negative != 0;
  out.display_x_axis = g.display_x_axis != 0;
  out.display_hidden = g.display_hidden != 0;
  out.min_axis_type = static_cast<io::SparklineAxisType>(g.min_axis_type);
  out.max_axis_type = static_cast<io::SparklineAxisType>(g.max_axis_type);
  out.has_manual_min = g.has_manual_min != 0;
  out.has_manual_max = g.has_manual_max != 0;
  out.manual_min = g.manual_min;
  out.manual_max = g.manual_max;
  out.date_axis = g.date_axis != 0;
  out.date_axis_ref = std::string(formulon::strip_formula_prefix(text_or_empty(g.date_axis_ref)));
  out.line_weight = g.line_weight;
  out.empty_cells_as = sparkline_empty_cells_from_c(g.empty_cells_as);
  return out;
}

// A group always states how it draws empty cells, so `ABSENT` is not a value here.
fm_status_t check_group_enums(const fm_sparkline_group_t& g) {
  return check_enum_domain(g.empty_cells_as, FM_BLANKS_AS_ZERO, FM_BLANKS_AS_SPAN, "empty_cells_as");
}

// Rejects a sparkline whose formula or location is malformed UTF-8.
fm_status_t check_items_utf8(const fm_sparkline_t* items, size_t count) {
  for (size_t i = 0; i < count; ++i) {
    if (auto rc = check_utf8({items[i].formula, items[i].location}); rc != 0) {
      return rc;
    }
  }
  return 0;
}

std::vector<io::SparklineItem> items_from_c(const fm_sparkline_t* items, size_t count) {
  std::vector<io::SparklineItem> out;
  out.reserve(count);
  for (size_t i = 0; i < count; ++i) {
    out.push_back({std::string(formulon::strip_formula_prefix(text_or_empty(items[i].formula))),
                   text_or_empty(items[i].location)});
  }
  return out;
}

// Merges the SET and CLEAR fields of `p` into `g`, checking every op first.
fm_status_t apply_patch(const fm_sparkline_group_patch_t& p, io::SparklineGroup& g) {
  struct Op {
    fm_patch_op_t op;
    const char* name;
    bool clearable;
  };
  const Op ops[] = {{p.color_series_op, "color_series", true},
                    {p.color_negative_op, "color_negative", true},
                    {p.color_axis_op, "color_axis", true},
                    {p.color_markers_op, "color_markers", true},
                    {p.color_first_op, "color_first", true},
                    {p.color_last_op, "color_last", true},
                    {p.color_high_op, "color_high", true},
                    {p.color_low_op, "color_low", true},
                    {p.manual_min_op, "manual_min", true},
                    {p.manual_max_op, "manual_max", true},
                    {p.line_weight_op, "line_weight", false},
                    {p.type_op, "type", false},
                    {p.markers_op, "markers", false},
                    {p.high_op, "high", false},
                    {p.low_op, "low", false},
                    {p.first_op, "first", false},
                    {p.last_op, "last", false},
                    {p.negative_op, "negative", false},
                    {p.display_x_axis_op, "display_x_axis", false},
                    {p.display_hidden_op, "display_hidden", false},
                    {p.min_axis_type_op, "min_axis_type", false},
                    {p.max_axis_type_op, "max_axis_type", false},
                    {p.date_axis_op, "date_axis", false},
                    {p.empty_cells_as_op, "empty_cells_as", false},
                    {p.date_axis_ref_op, "date_axis_ref", true},
                    {p.items_op, "items", false}};
  for (const Op& f : ops) {
    if (!(f.clearable ? patch_op_ok(f.op) : patch_op_ok_without_clear(f.op))) {
      return invalid_argument(patch_op_ok(f.op) ? "field cannot be cleared" : "op must be KEEP, CLEAR or SET",
                              std::string("field=") + f.name);
    }
  }
  const auto set = [](fm_patch_op_t op) { return op == FM_PATCH_OP_SET; };
  const std::pair<fm_patch_op_t, const fm_color_t*> colors[] = {
      {p.color_series_op, &p.color_series}, {p.color_negative_op, &p.color_negative},
      {p.color_axis_op, &p.color_axis},     {p.color_markers_op, &p.color_markers},
      {p.color_first_op, &p.color_first},   {p.color_last_op, &p.color_last},
      {p.color_high_op, &p.color_high},     {p.color_low_op, &p.color_low}};
  formulon::ColorSpec* const targets[] = {&g.color_series, &g.color_negative, &g.color_axis, &g.color_markers,
                                          &g.color_first,  &g.color_last,     &g.color_high, &g.color_low};
  for (std::size_t i = 0; i < 8; ++i) {
    if (set(colors[i].first)) {
      *targets[i] = color_spec_from_c(*colors[i].second);
    } else if (colors[i].first == FM_PATCH_OP_CLEAR) {
      *targets[i] = formulon::ColorSpec{};
    }
  }
  const std::pair<fm_patch_op_t, std::int32_t> flags[] = {{p.markers_op, p.markers},
                                                          {p.high_op, p.high},
                                                          {p.low_op, p.low},
                                                          {p.first_op, p.first},
                                                          {p.last_op, p.last},
                                                          {p.negative_op, p.negative},
                                                          {p.display_x_axis_op, p.display_x_axis},
                                                          {p.display_hidden_op, p.display_hidden},
                                                          {p.date_axis_op, p.date_axis}};
  bool* const flag_targets[] = {&g.markers,        &g.high,           &g.low,      &g.first, &g.last, &g.negative,
                                &g.display_x_axis, &g.display_hidden, &g.date_axis};
  for (std::size_t i = 0; i < 9; ++i) {
    if (set(flags[i].first)) {
      *flag_targets[i] = flags[i].second != 0;
    }
  }
  if (set(p.type_op)) {
    if (auto rc = check_enum_domain(p.type, FM_SPARKLINE_TYPE_LINE, FM_SPARKLINE_TYPE_STACKED, "type"); rc != 0) {
      return rc;
    }
    g.type = static_cast<io::SparklineType>(p.type);
  }
  if (set(p.min_axis_type_op)) {
    if (auto rc = check_enum_domain(p.min_axis_type, FM_SPARKLINE_AXIS_TYPE_INDIVIDUAL, FM_SPARKLINE_AXIS_TYPE_CUSTOM,
                                    "min_axis_type");
        rc != 0) {
      return rc;
    }
    g.min_axis_type = static_cast<io::SparklineAxisType>(p.min_axis_type);
  }
  if (set(p.max_axis_type_op)) {
    if (auto rc = check_enum_domain(p.max_axis_type, FM_SPARKLINE_AXIS_TYPE_INDIVIDUAL, FM_SPARKLINE_AXIS_TYPE_CUSTOM,
                                    "max_axis_type");
        rc != 0) {
      return rc;
    }
    g.max_axis_type = static_cast<io::SparklineAxisType>(p.max_axis_type);
  }
  if (set(p.empty_cells_as_op)) {
    if (auto rc = check_enum_domain(p.empty_cells_as, FM_BLANKS_AS_ZERO, FM_BLANKS_AS_SPAN, "empty_cells_as");
        rc != 0) {
      return rc;
    }
    g.empty_cells_as = sparkline_empty_cells_from_c(p.empty_cells_as);
  }
  if (set(p.line_weight_op)) {
    g.line_weight = p.line_weight;
  }
  if (p.manual_min_op != FM_PATCH_OP_KEEP) {
    g.has_manual_min = set(p.manual_min_op);
    g.manual_min = set(p.manual_min_op) ? p.manual_min : 0.0;
  }
  if (p.manual_max_op != FM_PATCH_OP_KEEP) {
    g.has_manual_max = set(p.manual_max_op);
    g.manual_max = set(p.manual_max_op) ? p.manual_max : 0.0;
  }
  if (set(p.date_axis_ref_op)) {
    if (p.date_axis_ref == nullptr) {
      return invalid_argument("a SET string cannot be NULL", "field=date_axis_ref");
    }
    if (auto rc = check_utf8(p.date_axis_ref, "field", "date_axis_ref"); rc != 0) {
      return rc;
    }
    g.date_axis_ref = std::string(formulon::strip_formula_prefix(p.date_axis_ref));
  } else if (p.date_axis_ref_op == FM_PATCH_OP_CLEAR) {
    g.date_axis_ref.clear();
  }
  return 0;
}

// Sheet `sheet_index` for a write. Sparkline groups of an `.xlsb` sheet are read into the
// extension XML; groups the reader could not decode stay binary and are kept beside it.
fm_status_t writable_sheet(fm_workbook_t* wb, size_t sheet_index, formulon::Sheet** sheet) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *sheet = &wb->workbook().sheet(sheet_index);
  return 0;
}

// `idx` against the sheet's sparkline groups.
fm_status_t check_group_index(const fm_workbook_t* wb, size_t sheet_index, size_t idx) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return check_index(idx, io::read_sparkline_groups(wb->workbook().sheet(sheet_index).ext_lst_xml()).size(),
                     "group_index");
}

// Runs `edit` on a copy of the sheet's extension list and keeps it on success.
template <typename Edit>
fm_status_t edit_ext_lst(fm_workbook_t* wb, size_t sheet_index, Edit edit) {
  formulon::Sheet* sheet = nullptr;
  if (auto rc = writable_sheet(wb, sheet_index, &sheet); rc != 0) {
    return rc;
  }
  std::string xml = sheet->ext_lst_xml();
  auto edited = edit(xml);
  if (!edited) {
    return set_last_error(edited.error());
  }
  sheet->set_ext_lst_xml(std::move(xml));
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_sparkline_group_count(const fm_workbook_t* wb, size_t sheet_index,
                                                          size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument("out_count");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_count = io::read_sparkline_groups(wb->workbook().sheet(sheet_index).ext_lst_xml()).size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_sparkline_group_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                       fm_sparkline_group_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const std::vector<io::SparklineGroup> groups =
      io::read_sparkline_groups(wb->workbook().sheet(sheet_index).ext_lst_xml());
  if (auto rc = check_index(idx, groups.size(), "group_index"); rc != 0) {
    return rc;
  }
  const io::SparklineGroup& g = groups[idx];
  TextStore& store = fresh_read_scratch(wb);
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_sparkline_group_t{};
  out->struct_size = struct_size;
  out->color_series = color_spec_to_c(g.color_series);
  out->color_negative = color_spec_to_c(g.color_negative);
  out->color_axis = color_spec_to_c(g.color_axis);
  out->color_markers = color_spec_to_c(g.color_markers);
  out->color_first = color_spec_to_c(g.color_first);
  out->color_last = color_spec_to_c(g.color_last);
  out->color_high = color_spec_to_c(g.color_high);
  out->color_low = color_spec_to_c(g.color_low);
  out->manual_min = g.manual_min;
  out->manual_max = g.manual_max;
  out->line_weight = g.line_weight;
  out->type = static_cast<int32_t>(g.type);
  out->markers = g.markers ? 1 : 0;
  out->high = g.high ? 1 : 0;
  out->low = g.low ? 1 : 0;
  out->first = g.first ? 1 : 0;
  out->last = g.last ? 1 : 0;
  out->negative = g.negative ? 1 : 0;
  out->display_x_axis = g.display_x_axis ? 1 : 0;
  out->display_hidden = g.display_hidden ? 1 : 0;
  out->min_axis_type = static_cast<int32_t>(g.min_axis_type);
  out->max_axis_type = static_cast<int32_t>(g.max_axis_type);
  out->has_manual_min = g.has_manual_min ? 1 : 0;
  out->has_manual_max = g.has_manual_max ? 1 : 0;
  out->date_axis = g.date_axis ? 1 : 0;
  out->empty_cells_as = to_c(g.empty_cells_as);
  out->date_axis_ref = g.date_axis_ref.empty() ? nullptr : store_cstr(store, with_formula_prefix(g.date_axis_ref));
  out->sparkline_count = static_cast<std::uint32_t>(g.items.size());
  return 0;
}

extern "C" fm_status_t fm_sheet_get_sparkline_at(const fm_workbook_t* wb, size_t sheet_index, size_t group_index,
                                                 size_t idx, fm_sparkline_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const std::vector<io::SparklineGroup> groups =
      io::read_sparkline_groups(wb->workbook().sheet(sheet_index).ext_lst_xml());
  if (auto rc = check_index(group_index, groups.size(), "group_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(idx, groups[group_index].items.size(), "sparkline_index"); rc != 0) {
    return rc;
  }
  const io::SparklineItem& item = groups[group_index].items[idx];
  TextStore& store = fresh_read_scratch(wb);
  out->formula = store_cstr(store, with_formula_prefix(item.formula));
  out->location = store_cstr(store, item.location);
  return 0;
}

extern "C" fm_status_t fm_sheet_add_sparkline_group(fm_workbook_t* wb, size_t sheet_index,
                                                    const fm_sparkline_group_t* group, const fm_sparkline_t* items,
                                                    size_t count, size_t* out_index) {
  clear_last_error();
  if (group == nullptr || items == nullptr || out_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*group); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({group->date_axis_ref}); rc != 0) {
    return rc;
  }
  if (auto rc = check_group_enums(*group); rc != 0) {
    return rc;
  }
  if (auto rc = check_items_utf8(items, count); rc != 0) {
    return rc;
  }
  io::SparklineGroup model = group_from_c(*group);
  model.items = items_from_c(items, count);
  return edit_ext_lst(wb, sheet_index, [&model, out_index](std::string& xml) {
    auto added = io::add_sparkline_group(xml, model);
    if (added) {
      *out_index = io::read_sparkline_groups(xml).size() - 1U;
    }
    return added;
  });
}

extern "C" fm_status_t fm_sheet_update_sparkline_group(fm_workbook_t* wb, size_t sheet_index, size_t group_index,
                                                       const fm_sparkline_group_patch_t* patch) {
  clear_last_error();
  if (patch == nullptr) {
    return null_argument("patch");
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const std::vector<io::SparklineGroup> groups =
      io::read_sparkline_groups(wb->workbook().sheet(sheet_index).ext_lst_xml());
  if (auto rc = check_index(group_index, groups.size(), "group_index"); rc != 0) {
    return rc;
  }
  io::SparklineGroup model = groups[group_index];
  if (auto rc = apply_patch(*patch, model); rc != 0) {
    return rc;
  }
  const bool replace_items = patch->items_op == FM_PATCH_OP_SET;
  std::vector<io::SparklineItem> replacement;
  if (replace_items) {
    if (patch->items == nullptr && patch->items_count > 0U) {
      return null_argument("items");
    }
    if (patch->items_count == 0U) {
      return invalid_argument("a SET items needs at least one sparkline", "field=items");
    }
    if (auto rc = check_items_utf8(patch->items, patch->items_count); rc != 0) {
      return rc;
    }
    replacement = items_from_c(patch->items, patch->items_count);
  }
  return edit_ext_lst(wb, sheet_index, [&](std::string& xml) -> formulon::Expected<void, formulon::Error> {
    if (replace_items) {
      if (auto replaced = io::set_group_sparklines(xml, group_index, replacement); !replaced) {
        return replaced;
      }
    }
    return io::update_sparkline_group(xml, group_index, model);
  });
}

extern "C" fm_status_t fm_sheet_remove_sparkline_group(fm_workbook_t* wb, size_t sheet_index, size_t idx) {
  clear_last_error();
  if (auto rc = check_group_index(wb, sheet_index, idx); rc != 0) {
    return rc;
  }
  return edit_ext_lst(wb, sheet_index, [idx](std::string& xml) { return io::remove_sparkline_group(xml, idx); });
}
