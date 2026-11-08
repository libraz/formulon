//
// C ABI - the table slicers of a worksheet: listing, adding, updating and
// removing them.

#include <cstddef>
#include <cstdint>
#include <string>
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
using formulon::c_api::parts::resolve_sheet_table;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::slicer_sort_order_from_c;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::to_c;

namespace {

// Converts `slicer` to the model spec; a non-zero return is the status to report.
fm_status_t to_spec(const fm_slicer_t& slicer, formulon::SlicerSpec& spec) {
  spec.fields = slicer.fields;
  if ((slicer.fields & FM_SLICER_FIELD_SORT_ORDER) != 0U) {
    if (auto rc = check_enum_domain(slicer.sort_order, FM_SLICER_SORT_ORDER_ASCENDING, FM_SLICER_SORT_ORDER_DESCENDING,
                                    "sort_order");
        rc != 0) {
      return rc;
    }
  }
  if ((slicer.fields & FM_SLICER_FIELD_CROSS_FILTER) != 0U) {
    if (auto rc = check_enum_domain(slicer.cross_filter, 0, FM_SLICER_CROSS_FILTER_NONE, "cross_filter"); rc != 0) {
      return rc;
    }
  }
  if (auto rc = check_utf8({slicer.name, slicer.caption, slicer.style}); rc != 0) {
    return rc;
  }
  spec.name = text_or_empty(slicer.name);
  spec.caption = text_or_empty(slicer.caption);
  spec.style = text_or_empty(slicer.style);
  spec.column_count = slicer.column_count;
  spec.row_height = slicer.row_height_emu;
  spec.from =
      formulon::AnchorPoint{slicer.from_row, slicer.from_col, slicer.from_row_offset_emu, slicer.from_col_offset_emu};
  spec.to = formulon::AnchorPoint{slicer.to_row, slicer.to_col, slicer.to_row_offset_emu, slicer.to_col_offset_emu};
  spec.sort_order = slicer_sort_order_from_c(slicer.sort_order);
  spec.cross_filter = static_cast<formulon::SlicerCrossFilter>(slicer.cross_filter);
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

// Resolves `object_id` for a mutator, reporting an unknown id as an invalid argument.
fm_status_t require_slicer(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t* index) {
  bool found = false;
  if (auto rc = find_slicer(wb, sheet_index, object_id, index, &found); rc != 0) {
    return rc;
  }
  if (!found) {
    return invalid_argument("no slicer has that object_id", "object_id=" + std::to_string(object_id));
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
                                              fm_slicer_t* out) {
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
  *out = fm_slicer_t{};
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
  out->style = store_cstr(store, info.style);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_slicer(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                           int32_t* out_present, fm_slicer_t* out) {
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
  *out = fm_slicer_t{};
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
                                           size_t table_index, size_t column_index, const fm_slicer_t* spec,
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
  if (auto rc = to_spec(*spec, model); rc != 0) {
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
                                              const fm_slicer_t* spec) {
  clear_last_error();
  if (spec == nullptr) {
    return null_argument("spec");
  }
  if (auto rc = check_record_size(*spec); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  formulon::SlicerSpec model;
  if (auto rc = to_spec(*spec, model); rc != 0) {
    return rc;
  }
  size_t idx = 0;
  if (auto rc = require_slicer(wb, sheet_index, object_id, &idx); rc != 0) {
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
