//
// C ABI - the table slicers of a worksheet: listing, adding, updating and
// removing them.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "drawing/slicer_edit.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::fill_anchor;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::resolve_sheet_table;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::slicer_sort_order_from_c;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::to_c;

namespace {

bool is_set(fm_patch_op_t op) {
  return op == FM_PATCH_OP_SET;
}

// Converts `p` to the model spec. `stored` is the slicer an update merges into (NULL on add); a non-zero
// return is the status to report. On add every field may be CLEAR (Excel's default); on update only `style`.
fm_status_t to_spec(const fm_slicer_patch_t& p, const formulon::SlicerInfo* stored, formulon::SlicerSpec& spec) {
  namespace field = formulon::slicer_field;
  const bool add = stored == nullptr;
  const std::pair<fm_patch_op_t, const char*> ops[] = {{p.column_count_op, "column_count"},
                                                       {p.sort_order_op, "sort_order"},
                                                       {p.cross_filter_op, "cross_filter"},
                                                       {p.from_row_op, "from_row"},
                                                       {p.from_col_op, "from_col"},
                                                       {p.to_row_op, "to_row"},
                                                       {p.to_col_op, "to_col"},
                                                       {p.from_row_offset_emu_op, "from_row_offset_emu"},
                                                       {p.from_col_offset_emu_op, "from_col_offset_emu"},
                                                       {p.to_row_offset_emu_op, "to_row_offset_emu"},
                                                       {p.to_col_offset_emu_op, "to_col_offset_emu"},
                                                       {p.row_height_emu_op, "row_height_emu"},
                                                       {p.name_op, "name"},
                                                       {p.caption_op, "caption"},
                                                       {p.style_op, "style"}};
  for (const auto& [op, name] : ops) {
    const bool clearable = add || std::string_view(name) == "style";
    if (!(clearable ? patch_op_ok(op) : patch_op_ok_without_clear(op))) {
      return invalid_argument(patch_op_ok(op) ? "field cannot be cleared" : "op must be KEEP, CLEAR or SET",
                              std::string("field=") + name);
    }
  }
  const std::pair<fm_patch_op_t, const char*> texts[] = {
      {p.name_op, p.name}, {p.caption_op, p.caption}, {p.style_op, p.style}};
  const char* const text_names[] = {"name", "caption", "style"};
  for (std::size_t i = 0; i < 3; ++i) {
    if (!is_set(texts[i].first)) {
      continue;
    }
    if (texts[i].second == nullptr) {
      return invalid_argument("a SET string cannot be NULL", std::string("field=") + text_names[i]);
    }
    if (auto rc = check_utf8(texts[i].second, "field", text_names[i]); rc != 0) {
      return rc;
    }
  }
  if (!add && is_set(p.name_op)) {
    return invalid_argument("a slicer cannot be renamed", "field=name");
  }

  spec.fields = 0;
  if (is_set(p.name_op)) {
    spec.fields |= field::kName;
    spec.name = p.name;
  }
  if (is_set(p.caption_op)) {
    spec.fields |= field::kCaption;
    spec.caption = p.caption;
  }
  if (is_set(p.style_op) || (!add && p.style_op == FM_PATCH_OP_CLEAR)) {
    spec.fields |= field::kStyle;
    spec.style = is_set(p.style_op) ? std::string(p.style) : std::string();
  }
  if (is_set(p.column_count_op)) {
    spec.fields |= field::kColumnCount;
    spec.column_count = p.column_count;
  }
  if (is_set(p.row_height_emu_op)) {
    spec.fields |= field::kRowHeight;
    spec.row_height = p.row_height_emu;
  }
  if (is_set(p.sort_order_op)) {
    if (auto rc = check_enum_domain(p.sort_order, FM_SLICER_SORT_ORDER_ASCENDING, FM_SLICER_SORT_ORDER_DESCENDING,
                                    "sort_order");
        rc != 0) {
      return rc;
    }
    spec.fields |= field::kSortOrder;
    spec.sort_order = slicer_sort_order_from_c(p.sort_order);
  }
  if (is_set(p.cross_filter_op)) {
    if (auto rc = check_enum_domain(p.cross_filter, 0, FM_SLICER_CROSS_FILTER_NONE, "cross_filter"); rc != 0) {
      return rc;
    }
    spec.fields |= field::kCrossFilter;
    spec.cross_filter = static_cast<formulon::SlicerCrossFilter>(p.cross_filter);
  }

  // The four anchor cells are SET together; an offset alone moves the anchor within its current cells.
  const int cells = static_cast<int>(is_set(p.from_row_op)) + static_cast<int>(is_set(p.from_col_op)) +
                    static_cast<int>(is_set(p.to_row_op)) + static_cast<int>(is_set(p.to_col_op));
  if (cells != 0 && cells != 4) {
    return invalid_argument("the four anchor cells are set together", "field=from_row");
  }
  if (add && cells != 4) {
    return invalid_argument("a new slicer needs its four anchor cells SET", "field=from_row");
  }
  const bool offsets = is_set(p.from_row_offset_emu_op) || is_set(p.from_col_offset_emu_op) ||
                       is_set(p.to_row_offset_emu_op) || is_set(p.to_col_offset_emu_op);
  if (cells == 4 || offsets) {
    spec.fields |= field::kAnchor;
    if (stored != nullptr) {
      spec.from = stored->object.from;
      spec.to = stored->object.to;
    }
    if (cells == 4) {
      spec.from.row = p.from_row;
      spec.from.col = p.from_col;
      spec.to.row = p.to_row;
      spec.to.col = p.to_col;
    }
    const auto take = [](fm_patch_op_t op, std::int64_t value, std::int64_t& out) {
      if (is_set(op)) {
        out = value;
      }
    };
    take(p.from_row_offset_emu_op, p.from_row_offset_emu, spec.from.row_off);
    take(p.from_col_offset_emu_op, p.from_col_offset_emu, spec.from.col_off);
    take(p.to_row_offset_emu_op, p.to_row_offset_emu, spec.to.row_off);
    take(p.to_col_offset_emu_op, p.to_col_offset_emu, spec.to.col_off);
  }
  return 0;
}

// Finds the position of the slicer whose frame id is `object_id` in `list_slicers`. A non-zero
// return is the status to report; `*found` is false when no slicer carries the id.
fm_status_t find_slicer(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t* index, bool* found) {
  auto slicers = formulon::list_slicers(wb->workbook(), sheet_index);
  if (!slicers) {
    return set_last_error(slicers.error());
  }
  const std::vector<formulon::SlicerInfo>& list = slicers.value();
  *found = false;
  for (size_t i = 0; object_id != 0U && i < list.size(); ++i) {
    if (list[i].object.object_id == object_id) {
      *index = i;
      *found = true;
      break;
    }
  }
  return 0;
}

// Resolves `object_id` for a mutator; an id no slicer holds is not found.
fm_status_t require_slicer(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t* index) {
  bool found = false;
  if (auto rc = find_slicer(wb, sheet_index, object_id, index, &found); rc != 0) {
    return rc;
  }
  if (!found) {
    return not_found("no slicer has that object_id", "object_id=" + std::to_string(object_id));
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_slicer_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument("out_count");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto slicers = formulon::list_slicers(wb->workbook(), sheet_index);
  if (!slicers) {
    return set_last_error(slicers.error());
  }
  *out_count = slicers.value().size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_slicer_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                              fm_slicer_info_t* out) {
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
  auto slicers = formulon::list_slicers(wb->workbook(), sheet_index);
  if (!slicers) {
    return set_last_error(slicers.error());
  }
  const std::vector<formulon::SlicerInfo>& list = slicers.value();
  if (auto rc = check_index(idx, list.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::SlicerInfo& info = list[idx];
  const formulon::DrawingObject& obj = info.object;
  TextStore& store = fresh_read_scratch(wb);
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_slicer_info_t{};
  out->struct_size = struct_size;
  out->object_id = obj.object_id;
  out->column_count = info.column_count;
  out->sort_order = to_c(info.sort_order);
  out->cross_filter = static_cast<int32_t>(info.cross_filter);
  fill_anchor(obj, out);
  out->row_height_emu = info.row_height;
  out->name = store_cstr(store, info.name);
  out->caption = store_cstr(store, info.caption);
  out->cache_name = store_cstr(store, info.cache_name);
  out->table_name = store_cstr(store, info.table_name);
  out->column_name = store_cstr(store, info.column_name);
  out->style = info.style.empty() ? nullptr : store_cstr(store, info.style);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_slicer(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                           int32_t* out_present, fm_slicer_info_t* out) {
  clear_last_error();
  if (out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_present = 0;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_slicer_info_t{};
  out->struct_size = struct_size;
  size_t index = 0;
  bool found = false;
  if (auto rc = find_slicer(wb, sheet_index, object_id, &index, &found); rc != 0 || !found) {
    return rc;
  }
  *out_present = 1;
  return fm_sheet_get_slicer_at(wb, sheet_index, index, out);
}

extern "C" fm_status_t fm_sheet_add_slicer(fm_workbook_t* wb, size_t sheet_index, size_t table_sheet_index,
                                           size_t table_index, size_t column_index, const fm_slicer_patch_t* spec,
                                           uint32_t* out_object_id) {
  clear_last_error();
  if (spec == nullptr || out_object_id == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*spec); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  // The engine addresses a table by its workbook-wide position.
  size_t global = 0;
  if (auto rc = resolve_sheet_table(wb, table_sheet_index, table_index, &global); rc != 0) {
    return rc;
  }
  formulon::SlicerSpec model;
  if (auto rc = to_spec(*spec, nullptr, model); rc != 0) {
    return rc;
  }
  auto id = formulon::add_slicer(wb->workbook(), sheet_index, global, column_index, model);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_update_slicer(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                              const fm_slicer_patch_t* patch) {
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
  size_t idx = 0;
  if (auto rc = require_slicer(wb, sheet_index, object_id, &idx); rc != 0) {
    return rc;
  }
  auto slicers = formulon::list_slicers(wb->workbook(), sheet_index);
  if (!slicers) {
    return set_last_error(slicers.error());
  }
  formulon::SlicerSpec model;
  if (auto rc = to_spec(*patch, &slicers.value()[idx], model); rc != 0) {
    return rc;
  }
  return finish_status(formulon::update_slicer(wb->workbook(), sheet_index, idx, model));
}

extern "C" fm_status_t fm_sheet_remove_slicer(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  size_t idx = 0;
  bool found = false;
  if (auto rc = find_slicer(wb, sheet_index, object_id, &idx, &found); rc != 0) {
    return rc;
  }
  if (!found) {
    return not_found("no slicer has that object_id", "object_id=" + std::to_string(object_id));
  }
  return finish_status(formulon::remove_slicer(wb->workbook(), sheet_index, idx));
}
