//
// C ABI - a sheet's drawing: probing images, listing the drawing's
// objects, reading and inserting pictures, and moving, reordering,
// capturing, restoring and removing pictures and charts.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "drawing/image_header.h"
#include "io/ooxml/part_dom.h"
#include "passthrough_part.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::anchor_kind_from_c;
using formulon::c_api::parts::check_anchor_enums;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::copy_owned_bytes;
using formulon::c_api::parts::edit_as_from_c;
using formulon::c_api::parts::fill_anchor;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::to_c;

namespace {

// Fills `out` without touching its `struct_size`.
void fill_image_info(const formulon::ImageInfo& info, fm_image_info_t* out) {
  out->format = static_cast<int32_t>(info.format);
  out->width_px = info.px_width;
  out->height_px = info.px_height;
}

}  // namespace

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
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "out_count is NULL");
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
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "out is NULL");
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
  out->edit_as = to_c(obj.edit_as);
  fill_anchor(obj, out);
  out->width_emu = obj.cx;
  out->height_emu = obj.cy;
  out->image_format = static_cast<int32_t>(obj.image_format);
  out->name = store_cstr(store, obj.name);
  out->description = store_cstr(store, obj.descr);
  out->media_path = obj.kind == formulon::DrawingObjectKind::kPicture ? store_cstr(store, obj.media_path) : nullptr;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_drawing_object(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                   int32_t* out_present, fm_drawing_object_info_t* out) {
  clear_last_error();
  if (out_present == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "out is NULL");
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
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "output pointer is NULL");
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
                                            const fm_picture_insert_spec_t* opts, uint32_t* out_object_id) {
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
  const fm_drawing_anchor_t& anchor = opts->anchor;
  if (auto rc = check_anchor_enums(anchor.anchor_kind, anchor.edit_as); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({opts->name, opts->description}); rc != 0) {
    return rc;
  }
  formulon::ImageInsertOptions model;
  model.name = opts->name != nullptr ? opts->name : "";
  model.descr = opts->description != nullptr ? opts->description : "";
  model.anchor_kind = anchor_kind_from_c(anchor.anchor_kind);
  model.edit_as = edit_as_from_c(anchor.edit_as);
  model.row = anchor.row;
  model.col = anchor.col;
  model.row_off = anchor.row_offset_emu;
  model.col_off = anchor.col_offset_emu;
  model.width_emu = anchor.width_emu;
  model.height_emu = anchor.height_emu;
  auto id = formulon::insert_image(wb->workbook(), sheet_index, bytes, len, model);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_drawing(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return finish_status(formulon::remove_drawing(wb->workbook(), sheet_index, object_id));
}

extern "C" fm_status_t fm_sheet_set_drawing_anchor(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                   const fm_drawing_anchor_t* anchor) {
  clear_last_error();
  if (anchor == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_anchor_enums(anchor->anchor_kind, anchor->edit_as); rc != 0) {
    return rc;
  }
  formulon::DrawingAnchor model;
  model.anchor_kind = anchor_kind_from_c(anchor->anchor_kind);
  model.edit_as = edit_as_from_c(anchor->edit_as);
  model.row = anchor->row;
  model.col = anchor->col;
  model.row_off = anchor->row_offset_emu;
  model.col_off = anchor->col_offset_emu;
  model.width_emu = anchor->width_emu;
  model.height_emu = anchor->height_emu;
  return finish_status(formulon::set_drawing_anchor(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_set_drawing_name(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                 const char* name) {
  clear_last_error();
  if (name == nullptr) {
    return null_argument("name");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  return finish_status(formulon::set_drawing_name(wb->workbook(), sheet_index, object_id, name));
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
