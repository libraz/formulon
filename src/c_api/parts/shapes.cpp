//
// C ABI - the shapes, connectors and groups of a sheet's drawing: listing,
// inserting, updating and removing them.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "drawing/shape_edit.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;

namespace {

// Converts `shape` to the model spec; a non-zero return is the status to report.
fm_status_t to_spec(const fm_shape& shape, const char* api, formulon::ShapeSpec& spec) {
  using formulon::AnchorKind;
  using formulon::EditAs;
  spec.fields = shape.fields;
  if ((shape.fields & ~formulon::shape_field::kAll) != 0U) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(api) + ": unknown field bit").c_str());
  }
  if (shape.name != nullptr) {
    spec.name = shape.name;
  }
  if (shape.prst != nullptr) {
    spec.prst = shape.prst;
  }
  if (shape.text != nullptr) {
    spec.text = shape.text;
  }
  spec.text_box = shape.text_box != 0;
  if ((shape.fields & FM_SHAPE_FIELD_ANCHOR) != 0U) {
    if (auto rc =
            check_enum_domain(shape.anchor_kind, static_cast<std::int64_t>(AnchorKind::kAbsolute), api, "anchor_kind");
        rc != 0) {
      return rc;
    }
    if (auto rc = check_enum_domain(shape.edit_as, static_cast<std::int64_t>(EditAs::kAbsolute), api, "edit_as");
        rc != 0) {
      return rc;
    }
  }
  spec.anchor_kind = static_cast<AnchorKind>(shape.anchor_kind);
  spec.edit_as = static_cast<EditAs>(shape.edit_as);
  spec.from = formulon::AnchorPoint{shape.from_row, shape.from_col, shape.from_row_off, shape.from_col_off};
  spec.to = formulon::AnchorPoint{shape.to_row, shape.to_col, shape.to_row_off, shape.to_col_off};
  spec.cx = shape.cx;
  spec.cy = shape.cy;
  spec.fill = static_cast<formulon::ShapePaint>(shape.fill_kind);
  spec.fill_rgb = shape.fill_rgb;
  spec.line = static_cast<formulon::ShapePaint>(shape.line_kind);
  spec.line_rgb = shape.line_rgb;
  spec.line_width_emu = shape.line_width_emu;
  spec.rotation = shape.rotation;
  spec.hidden = shape.hidden != 0;
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_shape_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_shape_count: NULL out_count");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_shape_count"); rc != 0) {
    return rc;
  }
  auto shapes = formulon::list_shapes(wb->workbook(), sheet_index);
  if (!shapes) {
    return set_last_error(shapes.error());
  }
  *out_count = shapes.value().size();
  return 0;
}

extern "C" fm_status_t fm_sheet_shape_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx, fm_shape* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_shape_at: NULL out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_shape_at"); rc != 0) {
    return rc;
  }
  auto shapes = formulon::list_shapes(wb->workbook(), sheet_index);
  if (!shapes) {
    return set_last_error(shapes.error());
  }
  const std::vector<formulon::ShapeInfo>& list = shapes.value();
  if (idx >= list.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_shape_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(list.size()));
  }
  const formulon::ShapeInfo& info = list[idx];
  const formulon::DrawingObject& obj = info.object;
  fm_workbook_t* scratch = const_cast<fm_workbook_t*>(wb);
  scratch->read_scratch.clear();
  *out = fm_shape{};
  out->object_id = obj.object_id;
  out->kind = static_cast<int32_t>(obj.kind);
  out->anchor_kind = static_cast<int32_t>(obj.anchor_kind);
  out->edit_as = static_cast<int32_t>(obj.edit_as);
  out->from_row = obj.from.row;
  out->from_col = obj.from.col;
  out->from_row_off = obj.from.row_off;
  out->from_col_off = obj.from.col_off;
  out->to_row = obj.to.row;
  out->to_col = obj.to.col;
  out->to_row_off = obj.to.row_off;
  out->to_col_off = obj.to.col_off;
  out->cx = obj.cx;
  out->cy = obj.cy;
  out->fill_kind = static_cast<int32_t>(info.fill);
  out->fill_rgb = info.fill_rgb;
  out->line_kind = static_cast<int32_t>(info.line);
  out->line_rgb = info.line_rgb;
  out->line_width_emu = info.line_width_emu;
  out->rotation = info.rotation;
  out->hidden = info.hidden ? 1 : 0;
  out->text_box = info.text_box ? 1 : 0;
  out->name = store_cstr(scratch->read_scratch, obj.name);
  out->prst = store_cstr(scratch->read_scratch, info.prst);
  out->text = store_cstr(scratch->read_scratch, info.text);
  return 0;
}

extern "C" fm_status_t fm_sheet_insert_shape(fm_workbook_t* wb, size_t sheet_index, const fm_shape* spec,
                                             uint32_t* out_object_id) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_insert_shape";
  if (spec == nullptr || out_object_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_insert_shape: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  formulon::ShapeSpec model;
  if (auto rc = to_spec(*spec, kApi, model); rc != 0) {
    return rc;
  }
  auto id = formulon::insert_shape(wb->workbook(), sheet_index, model);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_update_shape(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                             const fm_shape* spec) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_update_shape";
  if (spec == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_update_shape: NULL spec");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  formulon::ShapeSpec model;
  if (auto rc = to_spec(*spec, kApi, model); rc != 0) {
    return rc;
  }
  auto updated = formulon::update_shape(wb->workbook(), sheet_index, object_id, model);
  return updated ? 0 : set_last_error(updated.error());
}

extern "C" fm_status_t fm_sheet_remove_shape(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_remove_shape"); rc != 0) {
    return rc;
  }
  auto removed = formulon::remove_shape(wb->workbook(), sheet_index, object_id);
  return removed ? 0 : set_last_error(removed.error());
}
