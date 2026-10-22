//
// C ABI - the table slicers of a worksheet: listing, adding, updating and
// removing them.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "drawing/slicer_edit.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;

namespace {

// Converts `slicer` to the model spec; a non-zero return is the status to report.
fm_status_t to_spec(const fm_slicer& slicer, const char* api, formulon::SlicerSpec& spec) {
  spec.fields = slicer.fields;
  if ((slicer.fields & FM_SLICER_FIELD_SORT_ORDER) != 0U) {
    if (auto rc = check_enum_domain(slicer.sort_order, FM_SLICER_SORT_DESCENDING, api, "sort_order"); rc != 0) {
      return rc;
    }
  }
  if ((slicer.fields & FM_SLICER_FIELD_CROSS_FILTER) != 0U) {
    if (auto rc = check_enum_domain(slicer.cross_filter, FM_SLICER_CROSS_NONE, api, "cross_filter"); rc != 0) {
      return rc;
    }
  }
  spec.name = slicer.name != nullptr ? slicer.name : "";
  spec.caption = slicer.caption != nullptr ? slicer.caption : "";
  spec.style = slicer.style != nullptr ? slicer.style : "";
  spec.column_count = slicer.column_count;
  spec.row_height = slicer.row_height;
  spec.from = formulon::AnchorPoint{slicer.from_row, slicer.from_col, slicer.from_row_off, slicer.from_col_off};
  spec.to = formulon::AnchorPoint{slicer.to_row, slicer.to_col, slicer.to_row_off, slicer.to_col_off};
  spec.sort_order = static_cast<formulon::SlicerSortOrder>(slicer.sort_order);
  spec.cross_filter = static_cast<formulon::SlicerCrossFilter>(slicer.cross_filter);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_slicer_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_slicer_count: NULL out_count");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_slicer_count"); rc != 0) {
    return rc;
  }
  auto slicers = formulon::list_slicers(wb->workbook(), sheet_index);
  if (!slicers) {
    return set_last_error(slicers.error());
  }
  *out_count = slicers.value().size();
  return 0;
}

extern "C" fm_status_t fm_sheet_slicer_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx, fm_slicer* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_slicer_at: NULL out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_slicer_at"); rc != 0) {
    return rc;
  }
  auto slicers = formulon::list_slicers(wb->workbook(), sheet_index);
  if (!slicers) {
    return set_last_error(slicers.error());
  }
  const std::vector<formulon::SlicerInfo>& list = slicers.value();
  if (idx >= list.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_slicer_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(list.size()));
  }
  const formulon::SlicerInfo& info = list[idx];
  const formulon::DrawingObject& obj = info.object;
  fm_workbook_t* scratch = const_cast<fm_workbook_t*>(wb);
  scratch->read_scratch.clear();
  *out = fm_slicer{};
  out->object_id = obj.object_id;
  out->column_count = info.column_count;
  out->sort_order = static_cast<int32_t>(info.sort_order);
  out->cross_filter = static_cast<int32_t>(info.cross_filter);
  out->anchor_kind = static_cast<int32_t>(obj.anchor_kind);
  out->from_row = obj.from.row;
  out->from_col = obj.from.col;
  out->to_row = obj.to.row;
  out->to_col = obj.to.col;
  out->from_row_off = obj.from.row_off;
  out->from_col_off = obj.from.col_off;
  out->to_row_off = obj.to.row_off;
  out->to_col_off = obj.to.col_off;
  out->row_height = info.row_height;
  out->name = store_cstr(scratch->read_scratch, info.name);
  out->caption = store_cstr(scratch->read_scratch, info.caption);
  out->cache_name = store_cstr(scratch->read_scratch, info.cache_name);
  out->table_name = store_cstr(scratch->read_scratch, info.table_name);
  out->column_name = store_cstr(scratch->read_scratch, info.column_name);
  out->style = store_cstr(scratch->read_scratch, info.style);
  return 0;
}

extern "C" fm_status_t fm_sheet_add_slicer(fm_workbook_t* wb, size_t sheet_index, size_t table_index,
                                           size_t column_index, const fm_slicer* spec, uint32_t* out_object_id) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_add_slicer";
  if (spec == nullptr || out_object_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_add_slicer: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  formulon::SlicerSpec model;
  if (auto rc = to_spec(*spec, kApi, model); rc != 0) {
    return rc;
  }
  auto id = formulon::add_slicer(wb->workbook(), sheet_index, table_index, column_index, model);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_update_slicer(fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                              const fm_slicer* spec) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_update_slicer";
  if (spec == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_update_slicer: NULL spec");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  formulon::SlicerSpec model;
  if (auto rc = to_spec(*spec, kApi, model); rc != 0) {
    return rc;
  }
  auto updated = formulon::update_slicer(wb->workbook(), sheet_index, idx, model);
  return updated ? 0 : set_last_error(updated.error());
}

extern "C" fm_status_t fm_sheet_remove_slicer(fm_workbook_t* wb, size_t sheet_index, size_t idx) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_remove_slicer"); rc != 0) {
    return rc;
  }
  auto removed = formulon::remove_slicer(wb->workbook(), sheet_index, idx);
  return removed ? 0 : set_last_error(removed.error());
}
