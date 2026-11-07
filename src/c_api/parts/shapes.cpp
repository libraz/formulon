//
// C ABI - the shapes, connectors and groups of a sheet's drawing: listing,
// inserting, updating and removing them.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "drawing/shape_edit.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::anchor_kind_from_c;
using formulon::c_api::parts::check_anchor_enums;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::color_spec_from_c;
using formulon::c_api::parts::color_spec_to_c;
using formulon::c_api::parts::edit_as_from_c;
using formulon::c_api::parts::fill_anchor;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::shape_paint_from_c;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::to_c;

namespace {

// A shape paint stores 0xRRGGBB; the C colour carries it as opaque ARGB.
constexpr std::uint32_t kOpaque = 0xFF000000U;
constexpr std::uint32_t kRgbMask = 0x00FFFFFFU;

// The colour of a shape paint: Rgb for a solid paint, None otherwise.
fm_color_spec_t paint_color_to_c(formulon::ShapePaint paint, std::uint32_t rgb) {
  formulon::ColorSpec spec;
  if (paint == formulon::ShapePaint::kSolid) {
    spec.kind = formulon::ColorSpec::Kind::kRgb;
    spec.rgb = kOpaque | (rgb & kRgbMask);
  }
  return color_spec_to_c(spec);
}

std::uint32_t paint_rgb_from_c(const fm_color_spec_t& color) {
  std::uint32_t rgb = 0U;
  color_spec_from_c(color, &rgb);
  return rgb & kRgbMask;
}

// Converts `shape` to the model spec; a non-zero return is the status to report.
fm_status_t to_spec(const fm_shape_t& shape, formulon::ShapeSpec& spec) {
  spec.fields = shape.fields;
  if ((shape.fields & ~formulon::shape_field::kAll) != 0U) {
    return invalid_argument("unknown field bit");
  }
  if (auto rc = check_utf8({shape.name, shape.preset_geometry, shape.text}); rc != 0) {
    return rc;
  }
  if (shape.name != nullptr) {
    spec.name = shape.name;
  }
  if (shape.preset_geometry != nullptr) {
    spec.prst = shape.preset_geometry;
  }
  if (shape.text != nullptr) {
    spec.text = shape.text;
  }
  spec.text_box = shape.text_box != 0;
  if ((shape.fields & FM_SHAPE_FIELD_ANCHOR) != 0U) {
    if (auto rc = check_anchor_enums(shape.anchor_kind, shape.edit_as); rc != 0) {
      return rc;
    }
  }
  spec.anchor_kind = anchor_kind_from_c(shape.anchor_kind);
  spec.edit_as = edit_as_from_c(shape.edit_as);
  spec.from =
      formulon::AnchorPoint{shape.from_row, shape.from_col, shape.from_row_offset_emu, shape.from_col_offset_emu};
  spec.to = formulon::AnchorPoint{shape.to_row, shape.to_col, shape.to_row_offset_emu, shape.to_col_offset_emu};
  spec.cx = shape.width_emu;
  spec.cy = shape.height_emu;
  spec.fill = shape_paint_from_c(shape.fill_kind);
  spec.fill_rgb = paint_rgb_from_c(shape.fill_color);
  spec.line = shape_paint_from_c(shape.line_kind);
  spec.line_rgb = paint_rgb_from_c(shape.line_color);
  spec.line_width_emu = shape.line_width_emu;
  spec.rotation = shape.rotation;
  spec.hidden = shape.hidden != 0;
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_shape_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument("out_count");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto shapes = formulon::list_shapes(wb->workbook(), sheet_index);
  if (!shapes) {
    return set_last_error(shapes.error());
  }
  *out_count = shapes.value().size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_shape_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx, fm_shape_t* out) {
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
  auto shapes = formulon::list_shapes(wb->workbook(), sheet_index);
  if (!shapes) {
    return set_last_error(shapes.error());
  }
  const std::vector<formulon::ShapeInfo>& list = shapes.value();
  if (auto rc = check_index(idx, list.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::ShapeInfo& info = list[idx];
  const formulon::DrawingObject& obj = info.object;
  TextStore& store = fresh_read_scratch(wb);
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_shape_t{};
  out->struct_size = struct_size;
  out->object_id = obj.object_id;
  out->kind = static_cast<int32_t>(obj.kind);
  out->edit_as = to_c(obj.edit_as);
  fill_anchor(obj, out);
  out->width_emu = obj.cx;
  out->height_emu = obj.cy;
  out->fill_kind = to_c(info.fill);
  out->fill_color = paint_color_to_c(info.fill, info.fill_rgb);
  out->line_kind = to_c(info.line);
  out->line_color = paint_color_to_c(info.line, info.line_rgb);
  out->line_width_emu = info.line_width_emu;
  out->rotation = info.rotation;
  out->hidden = info.hidden ? 1 : 0;
  out->text_box = info.text_box ? 1 : 0;
  out->name = store_cstr(store, obj.name);
  out->preset_geometry = store_cstr(store, info.prst);
  out->text = store_cstr(store, info.text);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_shape(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                          int32_t* out_present, fm_shape_t* out) {
  clear_last_error();
  if (out_present == nullptr || out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_present = 0;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_shape_t{};
  out->struct_size = struct_size;
  auto shapes = formulon::list_shapes(wb->workbook(), sheet_index);
  if (!shapes) {
    return set_last_error(shapes.error());
  }
  const std::vector<formulon::ShapeInfo>& list = shapes.value();
  for (size_t i = 0; i < list.size(); ++i) {
    if (list[i].object.object_id == object_id) {
      *out_present = 1;
      return fm_sheet_get_shape_at(wb, sheet_index, i, out);
    }
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_add_shape(fm_workbook_t* wb, size_t sheet_index, const fm_shape_t* spec,
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
  formulon::ShapeSpec model;
  if (auto rc = to_spec(*spec, model); rc != 0) {
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
                                             const fm_shape_t* spec) {
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
  formulon::ShapeSpec model;
  if (auto rc = to_spec(*spec, model); rc != 0) {
    return rc;
  }
  return finish_status(formulon::update_shape(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_remove_shape(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto shapes = formulon::list_shapes(wb->workbook(), sheet_index);
  if (!shapes) {
    return set_last_error(shapes.error());
  }
  const std::vector<formulon::ShapeInfo>& list = shapes.value();
  if (std::none_of(list.begin(), list.end(),
                   [&](const formulon::ShapeInfo& s) { return s.object.object_id == object_id; })) {
    return not_found("no shape has that object_id", "object_id=" + std::to_string(object_id));
  }
  return finish_status(formulon::remove_shape(wb->workbook(), sheet_index, object_id));
}
