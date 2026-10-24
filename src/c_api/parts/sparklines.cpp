//
// C ABI - sparkline groups of a sheet's worksheet extension list: reading,
// adding, updating, replacing the sparklines of, and removing groups.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "io/sparkline_groups.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;

namespace {

using formulon::FormulonErrorCode;
namespace io = formulon::io;

fm_color_spec color_to_c(const formulon::ColorSpec& src) {
  fm_color_spec out{};
  out.tint = src.tint;
  out.rgb = src.rgb;
  out.theme = src.theme;
  out.indexed = src.indexed;
  out.kind = static_cast<uint8_t>(src.kind);
  return out;
}

formulon::ColorSpec color_from_c(const fm_color_spec& src) {
  formulon::ColorSpec out;
  out.kind = static_cast<formulon::ColorSpec::Kind>(src.kind);
  out.rgb = src.rgb;
  out.theme = src.theme;
  out.tint = src.tint;
  out.indexed = src.indexed;
  return out;
}

io::SparklineGroup group_from_c(const fm_sparkline_group& g) {
  io::SparklineGroup out;
  out.type = static_cast<io::SparklineType>(g.type);
  out.color_series = color_from_c(g.color_series);
  out.color_negative = color_from_c(g.color_negative);
  out.color_axis = color_from_c(g.color_axis);
  out.color_markers = color_from_c(g.color_markers);
  out.color_first = color_from_c(g.color_first);
  out.color_last = color_from_c(g.color_last);
  out.color_high = color_from_c(g.color_high);
  out.color_low = color_from_c(g.color_low);
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
  out.date_axis_ref = g.date_axis_ref != nullptr ? g.date_axis_ref : "";
  out.line_weight = g.line_weight;
  out.empty_cells_as = static_cast<io::SparklineEmptyCells>(g.empty_cells_as);
  return out;
}

// Rejects a sparkline whose formula or location is malformed UTF-8.
fm_status_t check_items_utf8(const fm_sparkline* items, size_t count, const char* api) {
  for (size_t i = 0; i < count; ++i) {
    if (auto rc = check_utf8(api, {items[i].formula, items[i].location}); rc != 0) {
      return rc;
    }
  }
  return 0;
}

std::vector<io::SparklineItem> items_from_c(const fm_sparkline* items, size_t count) {
  std::vector<io::SparklineItem> out;
  out.reserve(count);
  for (size_t i = 0; i < count; ++i) {
    out.push_back(
        {items[i].formula != nullptr ? items[i].formula : "", items[i].location != nullptr ? items[i].location : ""});
  }
  return out;
}

// Sheet `sheet_index` for a write. Sparkline groups of an `.xlsb` sheet are read into the
// extension XML; groups the reader could not decode stay binary and are kept beside it.
fm_status_t writable_sheet(fm_workbook_t* wb, size_t sheet_index, const char* api, formulon::Sheet** sheet) {
  if (auto rc = check_sheet_index(wb, sheet_index, api); rc != 0) {
    return rc;
  }
  *sheet = &wb->workbook().sheet(sheet_index);
  return 0;
}

// Runs `edit` on a copy of the sheet's extension list and keeps it on success.
template <typename Edit>
fm_status_t edit_ext_lst(fm_workbook_t* wb, size_t sheet_index, const char* api, Edit edit) {
  formulon::Sheet* sheet = nullptr;
  if (auto rc = writable_sheet(wb, sheet_index, api, &sheet); rc != 0) {
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

extern "C" fm_status_t fm_sheet_sparkline_group_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(FormulonErrorCode::kBindingNullPointer, "fm_sheet_sparkline_group_count: NULL out_count");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_sparkline_group_count"); rc != 0) {
    return rc;
  }
  *out_count = io::read_sparkline_groups(wb->workbook().sheet(sheet_index).ext_lst_xml()).size();
  return 0;
}

extern "C" fm_status_t fm_sheet_sparkline_group_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                   fm_sparkline_group* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(FormulonErrorCode::kBindingNullPointer, "fm_sheet_sparkline_group_at: NULL out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_sparkline_group_at"); rc != 0) {
    return rc;
  }
  const std::vector<io::SparklineGroup> groups =
      io::read_sparkline_groups(wb->workbook().sheet(sheet_index).ext_lst_xml());
  if (idx >= groups.size()) {
    return set_binding_error(FormulonErrorCode::kInvalidArgument, "fm_sheet_sparkline_group_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(groups.size()));
  }
  const io::SparklineGroup& g = groups[idx];
  fm_workbook_t* scratch = const_cast<fm_workbook_t*>(wb);
  scratch->read_scratch.clear();
  *out = fm_sparkline_group{};
  out->color_series = color_to_c(g.color_series);
  out->color_negative = color_to_c(g.color_negative);
  out->color_axis = color_to_c(g.color_axis);
  out->color_markers = color_to_c(g.color_markers);
  out->color_first = color_to_c(g.color_first);
  out->color_last = color_to_c(g.color_last);
  out->color_high = color_to_c(g.color_high);
  out->color_low = color_to_c(g.color_low);
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
  out->empty_cells_as = static_cast<int32_t>(g.empty_cells_as);
  out->date_axis_ref = store_cstr(scratch->read_scratch, g.date_axis_ref);
  out->sparkline_count = g.items.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_sparkline_at(const fm_workbook_t* wb, size_t sheet_index, size_t group_idx, size_t idx,
                                             fm_sparkline* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(FormulonErrorCode::kBindingNullPointer, "fm_sheet_sparkline_at: NULL out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_sparkline_at"); rc != 0) {
    return rc;
  }
  const std::vector<io::SparklineGroup> groups =
      io::read_sparkline_groups(wb->workbook().sheet(sheet_index).ext_lst_xml());
  if (group_idx >= groups.size() || idx >= groups[group_idx].items.size()) {
    return set_binding_error(FormulonErrorCode::kInvalidArgument, "fm_sheet_sparkline_at: index out of range",
                             "group_idx=" + std::to_string(group_idx) + " idx=" + std::to_string(idx));
  }
  const io::SparklineItem& item = groups[group_idx].items[idx];
  fm_workbook_t* scratch = const_cast<fm_workbook_t*>(wb);
  scratch->read_scratch.clear();
  out->formula = store_cstr(scratch->read_scratch, item.formula);
  out->location = store_cstr(scratch->read_scratch, item.location);
  return 0;
}

extern "C" fm_status_t fm_sheet_add_sparkline_group(fm_workbook_t* wb, size_t sheet_index,
                                                    const fm_sparkline_group* group, const fm_sparkline* items,
                                                    size_t count) {
  clear_last_error();
  if (group == nullptr || items == nullptr) {
    return set_binding_error(FormulonErrorCode::kBindingNullPointer, "fm_sheet_add_sparkline_group: NULL argument");
  }
  if (auto rc = check_utf8("fm_sheet_add_sparkline_group", {group->date_axis_ref}); rc != 0) {
    return rc;
  }
  if (auto rc = check_items_utf8(items, count, "fm_sheet_add_sparkline_group"); rc != 0) {
    return rc;
  }
  io::SparklineGroup model = group_from_c(*group);
  model.items = items_from_c(items, count);
  return edit_ext_lst(wb, sheet_index, "fm_sheet_add_sparkline_group",
                      [&model](std::string& xml) { return io::add_sparkline_group(xml, model); });
}

extern "C" fm_status_t fm_sheet_update_sparkline_group(fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                       const fm_sparkline_group* group) {
  clear_last_error();
  if (group == nullptr) {
    return set_binding_error(FormulonErrorCode::kBindingNullPointer, "fm_sheet_update_sparkline_group: NULL group");
  }
  if (auto rc = check_utf8("fm_sheet_update_sparkline_group", {group->date_axis_ref}); rc != 0) {
    return rc;
  }
  const io::SparklineGroup model = group_from_c(*group);
  return edit_ext_lst(wb, sheet_index, "fm_sheet_update_sparkline_group",
                      [&](std::string& xml) { return io::update_sparkline_group(xml, idx, model); });
}

extern "C" fm_status_t fm_sheet_set_sparkline_group_sparklines(fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                               const fm_sparkline* items, size_t count) {
  clear_last_error();
  if (items == nullptr) {
    return set_binding_error(FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_set_sparkline_group_sparklines: NULL items");
  }
  if (auto rc = check_items_utf8(items, count, "fm_sheet_set_sparkline_group_sparklines"); rc != 0) {
    return rc;
  }
  const std::vector<io::SparklineItem> model = items_from_c(items, count);
  return edit_ext_lst(wb, sheet_index, "fm_sheet_set_sparkline_group_sparklines",
                      [&](std::string& xml) { return io::set_group_sparklines(xml, idx, model); });
}

extern "C" fm_status_t fm_sheet_remove_sparkline_group(fm_workbook_t* wb, size_t sheet_index, size_t idx) {
  clear_last_error();
  return edit_ext_lst(wb, sheet_index, "fm_sheet_remove_sparkline_group",
                      [idx](std::string& xml) { return io::remove_sparkline_group(xml, idx); });
}
