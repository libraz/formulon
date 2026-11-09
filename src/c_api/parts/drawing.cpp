//
// C ABI - a sheet's drawing: probing images, listing the drawing's
// objects, reading and inserting pictures, and moving, reordering,
// capturing, restoring and removing pictures and charts.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/drawing_anchor.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_geometry.h"
#include "drawing/drawing_xml.h"
#include "drawing/image_header.h"
#include "drawing/shape_edit.h"
#include "drawing/slicer_edit.h"
#include "io/ooxml/part_dom.h"
#include "passthrough_part.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::anchor_kind_from_c;
using formulon::c_api::parts::AnchorPlan;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::copy_owned_bytes;
using formulon::c_api::parts::drawing_patch_from_c;
using formulon::c_api::parts::edit_as_from_c;
using formulon::c_api::parts::fill_drawing_anchor;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::null_field;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::patch_op_refusal;
using formulon::c_api::parts::plan_add_anchor;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::to_c;

namespace {

// Fills `out` without touching its `struct_size`.
void fill_image_info(const formulon::ImageInfo& info, fm_image_info_t* out) {
  out->format = static_cast<int32_t>(info.format);
  out->width_px = info.px_width;
  out->height_px = info.px_height;
}

// Offsets and sizes of an anchor lie in [0, 2^31) EMU.
constexpr std::int64_t kEmuLimit = std::int64_t{1} << 31;

bool in_emu_range(std::int64_t value) {
  return value >= 0 && value < kEmuLimit;
}

bool is_set(fm_patch_op_t op) {
  return op == FM_PATCH_OP_SET;
}

bool cell_outside_sheet(const formulon::AnchorPoint& p) {
  return p.row >= formulon::Sheet::kMaxRows || p.col >= formulon::Sheet::kMaxCols;
}

bool point_is_zero(const formulon::AnchorPoint& p) {
  return p.row == 0U && p.col == 0U && p.row_off == 0 && p.col_off == 0;
}

// The object's own size; a two-cell anchor's size is its box.
void current_size(const formulon::Geometry& geo, const formulon::DrawingObject& obj, std::int64_t* cx,
                  std::int64_t* cy) {
  if (obj.anchor_kind == formulon::AnchorKind::kTwoCell) {
    *cx = geo.span(false, obj.from, obj.to);
    *cy = geo.span(true, obj.from, obj.to);
    return;
  }
  *cx = obj.cx;
  *cy = obj.cy;
}

// Reads the ops of `p` into the engine's partial anchor; the merge itself is the engine's.
fm_status_t anchor_patch_from_c(const fm_drawing_anchor_patch_t& p, formulon::DrawingAnchorPatch* out) {
  struct OpField {
    fm_patch_op_t op;
    const char* name;
  };
  const OpField ops[] = {{p.kind_op, "kind"},
                         {p.edit_as_op, "edit_as"},
                         {p.from_row_op, "from_row"},
                         {p.from_col_op, "from_col"},
                         {p.from_row_offset_emu_op, "from_row_offset_emu"},
                         {p.from_col_offset_emu_op, "from_col_offset_emu"},
                         {p.to_row_op, "to_row"},
                         {p.to_col_op, "to_col"},
                         {p.to_row_offset_emu_op, "to_row_offset_emu"},
                         {p.to_col_offset_emu_op, "to_col_offset_emu"},
                         {p.width_emu_op, "width_emu"},
                         {p.height_emu_op, "height_emu"}};
  for (const OpField& f : ops) {
    if (!patch_op_ok_without_clear(f.op)) {
      return patch_op_refusal(f.op, std::string("anchor.") + f.name);
    }
  }
  if (is_set(p.kind_op)) {
    if (p.kind != FM_ANCHOR_KIND_ONE_CELL && p.kind != FM_ANCHOR_KIND_TWO_CELL) {
      return invalid_argument("anchor.kind is not a one-cell or two-cell kind",
                              "anchor.kind=" + std::to_string(p.kind));
    }
    out->kind = formulon::c_api::parts::anchor_kind_from_c(p.kind, formulon::AnchorKind::kOneCell);
  }
  if (is_set(p.edit_as_op)) {
    if (auto rc = check_enum_domain(p.edit_as, FM_ANCHOR_EDIT_AS_DEFAULT, FM_ANCHOR_EDIT_AS_ABSOLUTE, "anchor.edit_as");
        rc != 0) {
      return rc;
    }
    out->edit_as = formulon::c_api::parts::edit_as_from_c(p.edit_as);
  }
  const auto take = [](fm_patch_op_t op, auto value, auto& field) {
    if (is_set(op)) {
      field = value;
    }
  };
  take(p.from_row_op, p.from_row, out->from_row);
  take(p.from_col_op, p.from_col, out->from_col);
  take(p.from_row_offset_emu_op, p.from_row_offset_emu, out->from_row_off);
  take(p.from_col_offset_emu_op, p.from_col_offset_emu, out->from_col_off);
  take(p.to_row_op, p.to_row, out->to_row);
  take(p.to_col_op, p.to_col, out->to_col);
  take(p.to_row_offset_emu_op, p.to_row_offset_emu, out->to_row_off);
  take(p.to_col_offset_emu_op, p.to_col_offset_emu, out->to_col_off);
  take(p.width_emu_op, p.width_emu, out->width_emu);
  take(p.height_emu_op, p.height_emu, out->height_emu);
  return 0;
}

}  // namespace

namespace formulon {
namespace c_api {
namespace parts {

fm_status_t plan_add_anchor(const Workbook& wb, std::size_t sheet, const fm_drawing_anchor_t& a,
                            AnchorKind default_kind, bool size_required, AnchorPlan* out) {
  if (auto rc = check_enum_domain(a.kind, FM_ANCHOR_KIND_DEFAULT, FM_ANCHOR_KIND_ABSOLUTE, "anchor.kind"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(a.edit_as, FM_ANCHOR_EDIT_AS_DEFAULT, FM_ANCHOR_EDIT_AS_ABSOLUTE, "anchor.edit_as");
      rc != 0) {
    return rc;
  }
  AnchorPlan plan;
  plan.kind = anchor_kind_from_c(a.kind, default_kind);
  if (plan.kind == AnchorKind::kAbsolute) {
    return invalid_argument("anchor.kind must be one-cell or two-cell", "anchor.kind=" + std::to_string(a.kind));
  }
  plan.edit_as = edit_as_from_c(a.edit_as);
  plan.from = AnchorPoint{a.from_row, a.from_col, a.from_row_offset_emu, a.from_col_offset_emu};
  plan.to = AnchorPoint{a.to_row, a.to_col, a.to_row_offset_emu, a.to_col_offset_emu};
  if (!in_emu_range(plan.from.row_off) || !in_emu_range(plan.from.col_off) || !in_emu_range(plan.to.row_off) ||
      !in_emu_range(plan.to.col_off) || !in_emu_range(a.width_emu) || !in_emu_range(a.height_emu)) {
    return invalid_argument("anchor offset or size is out of range (maximum 2147483647)", "field=anchor");
  }
  if (cell_outside_sheet(plan.from) || cell_outside_sheet(plan.to)) {
    return invalid_argument("anchor is outside the grid", "field=anchor");
  }
  const Geometry geo(wb, wb.sheet(sheet));
  plan.cx = a.width_emu;
  plan.cy = a.height_emu;
  if (plan.kind == AnchorKind::kTwoCell) {
    if (!point_is_zero(plan.to)) {
      plan.cx = geo.span(false, plan.from, plan.to);
      plan.cy = geo.span(true, plan.from, plan.to);
      if (plan.to.row < plan.from.row || plan.to.col < plan.from.col || plan.cx < 0 || plan.cy < 0) {
        return invalid_argument("anchor.to_row must not lie before anchor.from_row", "field=anchor.to_row");
      }
    } else {
      plan.to = plan.from;
      if (!geo.advance(true, plan.to.row, plan.to.row_off, plan.cy) ||
          !geo.advance(false, plan.to.col, plan.to.col_off, plan.cx)) {
        return invalid_argument("anchor extends past the grid", "field=anchor");
      }
    }
  } else {
    plan.to = AnchorPoint{};
  }
  if (size_required && plan.cx <= 0 && plan.cy <= 0) {
    return invalid_argument("anchor needs a width or a height", "field=anchor");
  }
  *out = plan;
  return 0;
}

fm_status_t apply_placement_extras(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                   const fm_drawing_placement_spec_t& placement, bool with_description) {
  DrawingCommonPatch patch;
  if (placement.hidden == FM_TRISTATE_TRUE) {
    patch.hidden = true;
  }
  if (with_description && placement.description != nullptr && placement.description[0] != '\0') {
    patch.description = std::string(placement.description);
  }
  if (!patch.hidden && !patch.description) {
    return 0;
  }
  return finish_status(update_drawing(wb, sheet, object_id, patch));
}

void fill_drawing_anchor(const Workbook& wb, std::size_t sheet, const DrawingObject& obj, fm_drawing_anchor_t* out) {
  const Geometry geo(wb, wb.sheet(sheet));
  std::int64_t cx = 0;
  std::int64_t cy = 0;
  current_size(geo, obj, &cx, &cy);
  *out = fm_drawing_anchor_t{};
  out->kind = to_c(obj.anchor_kind);
  out->edit_as = to_c(obj.anchor_kind == AnchorKind::kTwoCell ? obj.edit_as : EditAs::kTwoCell);
  out->from_row = obj.from.row;
  out->from_col = obj.from.col;
  out->from_row_offset_emu = obj.from.row_off;
  out->from_col_offset_emu = obj.from.col_off;
  if (obj.anchor_kind == AnchorKind::kTwoCell) {
    out->to_row = obj.to.row;
    out->to_col = obj.to.col;
    out->to_row_offset_emu = obj.to.row_off;
    out->to_col_offset_emu = obj.to.col_off;
  }
  out->width_emu = cx;
  out->height_emu = cy;
}

fm_status_t drawing_patch_from_c(const fm_drawing_patch_t& patch, formulon::DrawingCommonPatch* out) {
  if (auto rc = check_record_size(patch); rc != 0) {
    return rc;
  }
  const fm_patch_op_t plain_ops[] = {patch.name_op, patch.hidden_op};
  for (const fm_patch_op_t op : plain_ops) {
    if (!patch_op_ok_without_clear(op)) {
      return patch_op_refusal(op, op == patch.name_op ? "name" : "hidden");
    }
  }
  if (!patch_op_ok(patch.description_op)) {
    return invalid_argument("description_op is not a patch op",
                            "description_op=" + std::to_string(patch.description_op));
  }
  if (is_set(patch.name_op) && (patch.name == nullptr || patch.name[0] == '\0')) {
    return invalid_argument("name must not be empty", "name=");
  }
  if (is_set(patch.description_op) && patch.description == nullptr) {
    return null_field("description");
  }
  if (auto rc = check_utf8(
          {is_set(patch.name_op) ? patch.name : nullptr, is_set(patch.description_op) ? patch.description : nullptr});
      rc != 0) {
    return rc;
  }
  if (is_set(patch.hidden_op)) {
    if (auto rc = check_enum_domain(patch.hidden, 0, 1, "hidden"); rc != 0) {
      return rc;
    }
  }

  if (is_set(patch.name_op)) {
    out->name = std::string(patch.name);
  }
  if (is_set(patch.description_op)) {
    out->description = std::string(patch.description);
  } else if (patch.description_op == FM_PATCH_OP_CLEAR) {
    out->description = std::string();
  }
  if (is_set(patch.hidden_op)) {
    out->hidden = patch.hidden != 0;
  }
  if (auto rc = anchor_patch_from_c(patch.anchor, &out->anchor); rc != 0) {
    return rc;
  }
  return 0;
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

extern "C" fm_status_t fm_get_image_info(const uint8_t* bytes, size_t len, fm_image_info_t* out) {
  clear_last_error();
  if (bytes == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  auto info = formulon::probe_image(bytes, len);
  if (!info) {
    return set_last_error(info.error());
  }
  fill_image_info(info.value(), out);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_drawing_object_count(const fm_workbook_t* wb, size_t sheet_index,
                                                         size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "out_count must not be NULL");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto objects = formulon::list_drawing_objects(wb->workbook(), sheet_index);
  if (!objects) {
    return set_last_error(objects.error());
  }
  *out_count = objects.value().size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_drawing_object_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                      fm_drawing_object_info_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "out must not be NULL");
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto objects = formulon::list_drawing_objects(wb->workbook(), sheet_index);
  if (!objects) {
    return set_last_error(objects.error());
  }
  const std::vector<formulon::DrawingObject>& list = objects.value();
  if (auto rc = check_index(idx, list.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::DrawingObject& obj = list[idx];
  TextStore& store = fresh_read_scratch(wb);
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_drawing_object_info_t{};
  out->struct_size = struct_size;
  out->object_id = obj.object_id;
  out->kind = static_cast<int32_t>(obj.kind);
  out->hidden = obj.hidden ? 1 : 0;
  fill_drawing_anchor(wb->workbook(), sheet_index, obj, &out->anchor);
  out->name = store_cstr(store, obj.name);
  out->description = store_cstr(store, obj.descr);
  out->media_path = obj.kind == formulon::DrawingObjectKind::kPicture ? store_cstr(store, obj.media_path) : nullptr;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_drawing_object(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                   int32_t* out_present, fm_drawing_object_info_t* out) {
  clear_last_error();
  if (out_present == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "out_present and out must not be NULL");
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  *out_present = 0;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_drawing_object_info_t{};
  out->struct_size = struct_size;
  auto objects = formulon::list_drawing_objects(wb->workbook(), sheet_index);
  if (!objects) {
    return set_last_error(objects.error());
  }
  const std::vector<formulon::DrawingObject>& list = objects.value();
  for (size_t i = 0; i < list.size(); ++i) {
    if (list[i].object_id == object_id) {
      *out_present = 1;
      return fm_sheet_get_drawing_object_at(wb, sheet_index, i, out);
    }
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_get_picture(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                            int32_t* out_present, const uint8_t** out_bytes, size_t* out_len,
                                            fm_image_info_t* out_info) {
  clear_last_error();
  if (out_present == nullptr || out_bytes == nullptr || out_len == nullptr || out_info == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "out_present, out_bytes, out_len and out_info must not be NULL");
  }
  if (auto rc = check_record_size(*out_info); rc != 0) {
    return rc;
  }
  *out_present = 0;
  *out_bytes = nullptr;
  *out_len = 0;
  {
    const std::uint32_t struct_size = out_info->struct_size;
    *out_info = fm_image_info_t{};
    out_info->struct_size = struct_size;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto objects = formulon::list_drawing_objects(wb->workbook(), sheet_index);
  if (!objects) {
    return set_last_error(objects.error());
  }
  for (const formulon::DrawingObject& obj : objects.value()) {
    if (obj.object_id != object_id || obj.kind != formulon::DrawingObjectKind::kPicture) {
      continue;
    }
    const formulon::PassthroughPart* media = formulon::io::ooxml::find_passthrough_part(wb->workbook(), obj.media_path);
    if (media == nullptr) {
      break;
    }
    auto info = formulon::probe_image(media->bytes.data(), media->bytes.size());
    if (!info) {
      return set_last_error(info.error());
    }
    *out_bytes = media->bytes.data();
    *out_len = media->bytes.size();
    fill_image_info(info.value(), out_info);
    *out_present = 1;
    return 0;
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_add_picture(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes, size_t len,
                                            const fm_drawing_placement_spec_t* opts, uint32_t* out_object_id) {
  clear_last_error();
  if (bytes == nullptr || opts == nullptr || out_object_id == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*opts); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({opts->name, opts->description}); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(opts->hidden, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "hidden"); rc != 0) {
    return rc;
  }
  AnchorPlan plan;
  if (auto rc =
          plan_add_anchor(wb->workbook(), sheet_index, opts->anchor, formulon::AnchorKind::kOneCell, false, &plan);
      rc != 0) {
    return rc;
  }
  formulon::ImageInsertOptions model;
  model.name = text_or_empty(opts->name);
  model.descr = text_or_empty(opts->description);
  model.anchor_kind = plan.kind;
  model.edit_as = plan.edit_as;
  model.row = plan.from.row;
  model.col = plan.from.col;
  model.row_off = plan.from.row_off;
  model.col_off = plan.from.col_off;
  model.width_emu = plan.cx;
  model.height_emu = plan.cy;
  auto id = formulon::insert_image(wb->workbook(), sheet_index, bytes, len, model);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return formulon::c_api::parts::apply_placement_extras(wb->workbook(), sheet_index, id.value(), *opts, false);
}

extern "C" fm_status_t fm_sheet_remove_drawing(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return finish_status(formulon::remove_drawing(wb->workbook(), sheet_index, object_id));
}

extern "C" fm_status_t fm_sheet_update_drawing(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                               const fm_drawing_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || patch == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  formulon::DrawingCommonPatch model;
  if (auto rc = drawing_patch_from_c(*patch, &model); rc != 0) {
    return rc;
  }
  return finish_status(formulon::update_drawing(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_set_drawing_z_order(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                    size_t z_order_index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return finish_status(formulon::set_drawing_z_order(wb->workbook(), sheet_index, object_id,
                                                     formulon::c_api::parts::saturate_index(z_order_index)));
}

extern "C" fm_status_t fm_sheet_snapshot_drawing(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                 uint8_t** out_bytes, size_t* out_len) {
  clear_last_error();
  if (out_bytes != nullptr) {
    *out_bytes = nullptr;
  }
  if (out_len != nullptr) {
    *out_len = 0;
  }
  if (out_bytes == nullptr || out_len == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto snapshot = formulon::snapshot_drawing(wb->workbook(), sheet_index, object_id);
  if (!snapshot) {
    return set_last_error(snapshot.error());
  }
  const std::vector<std::uint8_t>& bytes = snapshot.value();
  copy_owned_bytes(bytes, out_bytes, out_len);
  return 0;
}

extern "C" fm_status_t fm_sheet_restore_drawing(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes, size_t len,
                                                uint32_t flags, uint32_t* out_object_id) {
  clear_last_error();
  if (bytes == nullptr || out_object_id == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto id = formulon::restore_drawing(wb->workbook(), sheet_index, bytes, len, flags);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return 0;
}
