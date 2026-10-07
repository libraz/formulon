//
// C ABI - images in a sheet's drawing: probing, listing the drawing's
// objects, reading, inserting, moving, reordering, capturing, restoring and
// removing pictures.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "drawing/image_header.h"
#include "io/ooxml/part_dom.h"
#include "passthrough_part.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

namespace {

constexpr const char* kNullPointer = "NULL argument";

fm_image_info to_fm(const formulon::ImageInfo& info) {
  return fm_image_info{static_cast<int32_t>(info.format), info.px_width, info.px_height};
}

const char* keep(fm_workbook_t* scratch, const std::string& text) {
  scratch->read_scratch.emplace_back(text);
  return scratch->read_scratch.back().c_str();
}

}  // namespace

extern "C" fm_status_t fm_workbook_probe_image(const fm_workbook_t* wb, const uint8_t* bytes, size_t len,
                                               fm_image_info* out) {
  clear_last_error();
  if (wb == nullptr || bytes == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_probe_image: NULL argument");
  }
  auto info = formulon::probe_image(bytes, len);
  if (!info) {
    return set_last_error(info.error());
  }
  *out = to_fm(info.value());
  return 0;
}

extern "C" fm_status_t fm_sheet_drawing_object_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_drawing_object_count: NULL out_count");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_drawing_object_count"); rc != 0) {
    return rc;
  }
  auto objects = formulon::list_drawing_objects(wb->workbook(), sheet_index);
  if (!objects) {
    return set_last_error(objects.error());
  }
  *out_count = objects.value().size();
  return 0;
}

extern "C" fm_status_t fm_sheet_drawing_object_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                                  fm_drawing_object* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_drawing_object_at: NULL out");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_drawing_object_at"); rc != 0) {
    return rc;
  }
  auto objects = formulon::list_drawing_objects(wb->workbook(), sheet_index);
  if (!objects) {
    return set_last_error(objects.error());
  }
  const std::vector<formulon::DrawingObject>& list = objects.value();
  if (idx >= list.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_drawing_object_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(list.size()));
  }
  const formulon::DrawingObject& obj = list[idx];
  fm_workbook_t* scratch = const_cast<fm_workbook_t*>(wb);
  scratch->read_scratch.clear();
  *out = fm_drawing_object{};
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
  out->image_format = static_cast<int32_t>(obj.image_format);
  out->name = keep(scratch, obj.name);
  out->descr = keep(scratch, obj.descr);
  out->media_path = keep(scratch, obj.media_path);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_image(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                          const uint8_t** out_bytes, size_t* out_len, fm_image_info* out_info) {
  clear_last_error();
  if (out_bytes == nullptr || out_len == nullptr || out_info == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_get_image: NULL output pointer");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_get_image"); rc != 0) {
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
    *out_info = to_fm(info.value());
    return 0;
  }
  return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                           "fm_sheet_get_image: no picture with that object id",
                           "sheet_index=" + std::to_string(sheet_index) + " object_id=" + std::to_string(object_id));
}

extern "C" fm_status_t fm_sheet_insert_image(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes, size_t len,
                                             const fm_image_insert* opts, uint32_t* out_object_id) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_insert_image";
  if (bytes == nullptr || opts == nullptr || out_object_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(kApi) + ": " + kNullPointer).c_str());
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(opts->anchor_kind, static_cast<std::int64_t>(formulon::AnchorKind::kAbsolute), kApi,
                                  "anchor_kind");
      rc != 0) {
    return rc;
  }
  if (auto rc =
          check_enum_domain(opts->edit_as, static_cast<std::int64_t>(formulon::EditAs::kAbsolute), kApi, "edit_as");
      rc != 0) {
    return rc;
  }
  formulon::ImageInsertOptions model;
  model.name = opts->name != nullptr ? opts->name : "";
  model.descr = opts->descr != nullptr ? opts->descr : "";
  model.anchor_kind = static_cast<formulon::AnchorKind>(opts->anchor_kind);
  model.edit_as = static_cast<formulon::EditAs>(opts->edit_as);
  model.row = opts->row;
  model.col = opts->col;
  model.row_off = opts->row_off_emu;
  model.col_off = opts->col_off_emu;
  model.width_emu = opts->width_emu;
  model.height_emu = opts->height_emu;
  auto id = formulon::insert_image(wb->workbook(), sheet_index, bytes, len, model);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_image(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_remove_image"); rc != 0) {
    return rc;
  }
  auto removed = formulon::remove_image(wb->workbook(), sheet_index, object_id);
  return removed ? 0 : set_last_error(removed.error());
}

extern "C" fm_status_t fm_sheet_set_image_anchor(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                 const fm_image_anchor* anchor) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_set_image_anchor";
  if (anchor == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(kApi) + ": " + kNullPointer).c_str());
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(anchor->anchor_kind, static_cast<std::int64_t>(formulon::AnchorKind::kAbsolute), kApi,
                                  "anchor_kind");
      rc != 0) {
    return rc;
  }
  if (auto rc =
          check_enum_domain(anchor->edit_as, static_cast<std::int64_t>(formulon::EditAs::kAbsolute), kApi, "edit_as");
      rc != 0) {
    return rc;
  }
  formulon::ImageAnchor model;
  model.anchor_kind = static_cast<formulon::AnchorKind>(anchor->anchor_kind);
  model.edit_as = static_cast<formulon::EditAs>(anchor->edit_as);
  model.row = anchor->row;
  model.col = anchor->col;
  model.row_off = anchor->row_off_emu;
  model.col_off = anchor->col_off_emu;
  model.width_emu = anchor->width_emu;
  model.height_emu = anchor->height_emu;
  auto moved = formulon::set_image_anchor(wb->workbook(), sheet_index, object_id, model);
  return moved ? 0 : set_last_error(moved.error());
}

extern "C" fm_status_t fm_sheet_set_image_z_order(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                  uint32_t index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_set_image_z_order"); rc != 0) {
    return rc;
  }
  auto moved = formulon::set_image_z_order(wb->workbook(), sheet_index, object_id, index);
  return moved ? 0 : set_last_error(moved.error());
}

extern "C" fm_status_t fm_sheet_snapshot_image(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                               uint8_t** out_bytes, size_t* out_len) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_snapshot_image";
  if (out_bytes != nullptr) {
    *out_bytes = nullptr;
  }
  if (out_len != nullptr) {
    *out_len = 0;
  }
  if (out_bytes == nullptr || out_len == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(kApi) + ": " + kNullPointer).c_str());
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  auto snapshot = formulon::snapshot_image(wb->workbook(), sheet_index, object_id);
  if (!snapshot) {
    return set_last_error(snapshot.error());
  }
  const std::vector<std::uint8_t>& bytes = snapshot.value();
  auto* buffer = new uint8_t[bytes.size()];
  if (!bytes.empty()) {
    std::memcpy(buffer, bytes.data(), bytes.size());
  }
  *out_bytes = buffer;
  *out_len = bytes.size();
  return 0;
}

extern "C" fm_status_t fm_sheet_restore_image(fm_workbook_t* wb, size_t sheet_index, const uint8_t* bytes, size_t len,
                                              uint32_t flags, uint32_t* out_object_id) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_restore_image";
  if (bytes == nullptr || out_object_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(kApi) + ": " + kNullPointer).c_str());
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  auto id = formulon::restore_image(wb->workbook(), sheet_index, bytes, len, flags);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return 0;
}
