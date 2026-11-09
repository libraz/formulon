//
// C ABI - the shapes, connectors and groups of a sheet's drawing: listing,
// inserting, updating and removing them.

#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/chart_patch_reader.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "drawing/drawing_geometry.h"
#include "drawing/shape_edit.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::anchor_kind_from_c;
using formulon::c_api::parts::check_enum_domain;
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
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::shape_paint_from_c;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::TextStore;
using formulon::c_api::parts::to_c;

namespace {

// A shape's Rgb colour is opaque; the alpha byte of the C colour is not kept.
constexpr std::uint32_t kOpaque = 0xFF000000U;
constexpr std::uint32_t kRgbMask = 0x00FFFFFFU;

// The colour of a shape paint: Rgb or Theme for a solid paint, None otherwise.
fm_color_t paint_color_to_c(formulon::ShapePaint paint, const formulon::ColorSpec& color) {
  return color_spec_to_c(paint == formulon::ShapePaint::kSolid ? color : formulon::ColorSpec{});
}

// The colour of a solid paint; a None colour is black. A non-solid paint must carry None.
fm_status_t paint_color_from_c(formulon::ShapePaint paint, const fm_color_t& color, const char* field,
                               formulon::ColorSpec& out) {
  out = formulon::ColorSpec{};
  if (paint != formulon::ShapePaint::kSolid) {
    if (color.kind != FM_COLOR_KIND_NONE) {
      return invalid_argument("colour needs a solid paint", std::string("field=") + field);
    }
    return 0;
  }
  if (const char* why = formulon::c_api::parts::color_spec_problem(color); why != nullptr) {
    return invalid_argument(why, std::string("field=") + field);
  }
  out = color_spec_from_c(color);
  if (out.kind == formulon::ColorSpec::Kind::kNone) {
    out.kind = formulon::ColorSpec::Kind::kRgb;
  }
  if (out.kind == formulon::ColorSpec::Kind::kRgb) {
    out.rgb = kOpaque | (out.rgb & kRgbMask);
  }
  return 0;
}

/// `a:xfrm/@rot` counts 60000ths of a degree.
constexpr double kRotationUnitsPerDegree = 60000.0;
constexpr double kMaxRotationDeg = 3600.0;
constexpr double kFullTurnDeg = 360.0;
constexpr double kMaxLineWidthPt = 1584.0;

// The stored paint of an update whose id names no shape; the engine refuses the id before using it.
constexpr formulon::ShapePaint kUnsetPaint = formulon::ShapePaint::kUnset;
const formulon::ColorSpec kNoColor{};

bool is_set(fm_patch_op_t op) {
  return op == FM_PATCH_OP_SET;
}

struct OpField {
  fm_patch_op_t op;
  const char* name;
  bool paint;  // One of the four fields an update may clear.
};

// Every op of `p` with its field name, in declaration order.
std::array<OpField, 23> ops_of(const fm_shape_patch_t& p) {
  return {{{p.anchor_kind_op, "anchor_kind", false},
           {p.edit_as_op, "edit_as", false},
           {p.from_row_op, "from_row", false},
           {p.from_col_op, "from_col", false},
           {p.to_row_op, "to_row", false},
           {p.to_col_op, "to_col", false},
           {p.from_row_offset_emu_op, "from_row_offset_emu", false},
           {p.from_col_offset_emu_op, "from_col_offset_emu", false},
           {p.to_row_offset_emu_op, "to_row_offset_emu", false},
           {p.to_col_offset_emu_op, "to_col_offset_emu", false},
           {p.width_emu_op, "width_emu", false},
           {p.height_emu_op, "height_emu", false},
           {p.fill_kind_op, "fill_kind", true},
           {p.fill_color_op, "fill_color", true},
           {p.line_kind_op, "line_kind", true},
           {p.line_color_op, "line_color", true},
           {p.line_width_pt_op, "line_width_pt", false},
           {p.rotation_deg_op, "rotation_deg", false},
           {p.hidden_op, "hidden", false},
           {p.text_box_op, "text_box", false},
           {p.name_op, "name", false},
           {p.preset_geometry_op, "preset_geometry", false},
           {p.text_op, "text", false}}};
}

// The ops, strings and UTF-8 of `p`. On add every field may be CLEAR (Excel's default); on update only the
// paint fields.
fm_status_t check_patch(const fm_shape_patch_t& p, bool add) {
  for (const OpField& f : ops_of(p)) {
    const bool ok = add || f.paint ? patch_op_ok(f.op) : patch_op_ok_without_clear(f.op);
    if (!ok) {
      return invalid_argument(patch_op_ok(f.op) ? "field cannot be cleared" : "op must be KEEP, CLEAR or SET",
                              std::string("field=") + f.name);
    }
  }
  const std::array<std::pair<fm_patch_op_t, const char*>, 3> texts = {
      {{p.name_op, p.name}, {p.preset_geometry_op, p.preset_geometry}, {p.text_op, p.text}}};
  const char* const names[] = {"name", "preset_geometry", "text"};
  for (std::size_t i = 0; i < texts.size(); ++i) {
    if (!is_set(texts[i].first)) {
      continue;
    }
    if (texts[i].second == nullptr) {
      return invalid_argument("a SET string cannot be NULL", std::string("field=") + names[i]);
    }
    if (auto rc = check_utf8(texts[i].second, "field", names[i]); rc != 0) {
      return rc;
    }
  }
  return 0;
}

fm_status_t check_flag(std::int32_t value, const char* field) {
  return check_enum_domain(value, 0, 1, field);
}

// One paint (fill or outline) of a patch, merged with the stored paint (`stored` is NULL on add). `apply`
// says whether the patch writes it.
fm_status_t paint_from_patch(fm_patch_op_t kind_op, std::int32_t kind, fm_patch_op_t color_op, const fm_color_t& color,
                             const formulon::ShapePaint* stored_kind, const formulon::ColorSpec* stored_color,
                             const char* kind_name, const char* color_name, formulon::ShapePaint& out_paint,
                             formulon::ColorSpec& out_color, bool& apply) {
  apply = false;
  const bool clear = stored_kind != nullptr && (kind_op == FM_PATCH_OP_CLEAR || color_op == FM_PATCH_OP_CLEAR);
  if (clear) {
    if (is_set(kind_op) || is_set(color_op)) {
      return invalid_argument("one of a paint pair cannot be cleared while the other is set",
                              std::string("field=") + kind_name);
    }
    out_paint = formulon::ShapePaint::kUnset;
    out_color = formulon::ColorSpec{};
    apply = true;
    return 0;
  }
  if (!is_set(kind_op) && !is_set(color_op)) {
    return 0;
  }
  if (is_set(kind_op) && kind != FM_SHAPE_PAINT_SOLID && kind != FM_SHAPE_PAINT_NONE) {
    return invalid_argument("a paint is solid or none", std::string("field=") + kind_name);
  }
  // Setting only a colour makes the paint solid.
  out_paint = is_set(kind_op) ? shape_paint_from_c(kind) : formulon::ShapePaint::kSolid;
  if (is_set(color_op)) {
    if (auto rc = paint_color_from_c(out_paint, color, color_name, out_color); rc != 0) {
      return rc;
    }
  } else if (out_paint == formulon::ShapePaint::kSolid && stored_kind != nullptr &&
             *stored_kind == formulon::ShapePaint::kSolid) {
    out_color = *stored_color;
  } else {
    fm_color_t none{};
    none.kind = FM_COLOR_KIND_NONE;
    if (auto rc = paint_color_from_c(out_paint, none, color_name, out_color); rc != 0) {
      return rc;
    }
  }
  apply = true;
  return 0;
}

// Converts `p` to the model spec. `stored` is the shape an update merges into (NULL on add, and on an update
// whose id names no shape, which the engine then refuses); a non-zero return is the status to report.
fm_status_t to_spec(const fm_shape_patch_t& p, bool add, const formulon::ShapeInfo* stored, formulon::ShapeSpec& spec) {
  namespace field = formulon::shape_field;
  if (auto rc = check_patch(p, add); rc != 0) {
    return rc;
  }
  spec.fields = 0;
  if (is_set(p.name_op)) {
    spec.fields |= field::kName;
    spec.name = p.name;
  }
  if (is_set(p.preset_geometry_op)) {
    spec.fields |= field::kPrst;
    spec.prst = p.preset_geometry;
  }
  if (is_set(p.text_op)) {
    spec.fields |= field::kText;
    spec.text = p.text;
  }
  if (is_set(p.text_box_op)) {
    if (auto rc = check_flag(p.text_box, "text_box"); rc != 0) {
      return rc;
    }
    spec.fields |= field::kTextBox;
    spec.text_box = p.text_box != 0;
  }
  if (is_set(p.hidden_op)) {
    if (auto rc = check_flag(p.hidden, "hidden"); rc != 0) {
      return rc;
    }
    spec.fields |= field::kHidden;
    spec.hidden = p.hidden != 0;
  }

  const fm_patch_op_t anchor_ops[] = {p.anchor_kind_op,
                                      p.edit_as_op,
                                      p.from_row_op,
                                      p.from_col_op,
                                      p.to_row_op,
                                      p.to_col_op,
                                      p.from_row_offset_emu_op,
                                      p.from_col_offset_emu_op,
                                      p.to_row_offset_emu_op,
                                      p.to_col_offset_emu_op,
                                      p.width_emu_op,
                                      p.height_emu_op};
  const bool any_anchor = std::any_of(std::begin(anchor_ops), std::end(anchor_ops), is_set);
  if (add && !is_set(p.anchor_kind_op)) {
    return invalid_argument("a new shape needs anchor_kind SET", "field=anchor_kind");
  }
  if (add || any_anchor) {
    spec.fields |= field::kAnchor;
    if (stored != nullptr) {
      const formulon::DrawingObject& obj = stored->object;
      spec.anchor_kind = obj.anchor_kind;
      spec.edit_as = obj.edit_as;
      spec.from = obj.from;
      spec.to = obj.to;
      spec.cx = obj.cx;
      spec.cy = obj.cy;
    }
    if (is_set(p.anchor_kind_op)) {
      if (auto rc = check_enum_domain(p.anchor_kind, 0, FM_ANCHOR_KIND_ABSOLUTE, "anchor_kind"); rc != 0) {
        return rc;
      }
      const formulon::AnchorKind kind = anchor_kind_from_c(p.anchor_kind);
      if (add && kind == formulon::AnchorKind::kAbsolute) {
        return invalid_argument("a new shape is anchored one-cell or two-cell", "field=anchor_kind");
      }
      if (stored != nullptr && kind != stored->object.anchor_kind) {
        return invalid_argument("an update cannot change the anchor kind", "field=anchor_kind");
      }
      spec.anchor_kind = kind;
      if (add) {
        spec.edit_as = kind == formulon::AnchorKind::kOneCell ? formulon::EditAs::kOneCell : formulon::EditAs::kTwoCell;
      }
    }
    if (is_set(p.edit_as_op)) {
      if (auto rc = check_enum_domain(p.edit_as, 0, FM_ANCHOR_EDIT_AS_ABSOLUTE, "edit_as"); rc != 0) {
        return rc;
      }
      spec.edit_as = edit_as_from_c(p.edit_as);
    }
    const auto take = [](fm_patch_op_t op, auto value, auto& out) {
      if (is_set(op)) {
        out = value;
      }
    };
    take(p.from_row_op, p.from_row, spec.from.row);
    take(p.from_col_op, p.from_col, spec.from.col);
    take(p.to_row_op, p.to_row, spec.to.row);
    take(p.to_col_op, p.to_col, spec.to.col);
    take(p.from_row_offset_emu_op, p.from_row_offset_emu, spec.from.row_off);
    take(p.from_col_offset_emu_op, p.from_col_offset_emu, spec.from.col_off);
    take(p.to_row_offset_emu_op, p.to_row_offset_emu, spec.to.row_off);
    take(p.to_col_offset_emu_op, p.to_col_offset_emu, spec.to.col_off);
    take(p.width_emu_op, p.width_emu, spec.cx);
    take(p.height_emu_op, p.height_emu, spec.cy);
  }

  bool apply = false;
  if (auto rc = paint_from_patch(p.fill_kind_op, p.fill_kind, p.fill_color_op, p.fill_color,
                                 add ? nullptr : (stored != nullptr ? &stored->fill : &kUnsetPaint),
                                 stored != nullptr ? &stored->fill_color : &kNoColor, "fill_kind", "fill_color",
                                 spec.fill, spec.fill_color, apply);
      rc != 0) {
    return rc;
  }
  if (apply) {
    spec.fields |= field::kFill;
  }
  if (auto rc = paint_from_patch(p.line_kind_op, p.line_kind, p.line_color_op, p.line_color,
                                 add ? nullptr : (stored != nullptr ? &stored->line : &kUnsetPaint),
                                 stored != nullptr ? &stored->line_color : &kNoColor, "line_kind", "line_color",
                                 spec.line, spec.line_color, apply);
      rc != 0) {
    return rc;
  }
  if (apply) {
    spec.fields |= field::kLine;
  }

  if (is_set(p.line_width_pt_op)) {
    if (!std::isfinite(p.line_width_pt) || p.line_width_pt < 0.0 || p.line_width_pt > kMaxLineWidthPt) {
      return invalid_argument("line width is outside its range", "field=line_width_pt");
    }
    spec.fields |= field::kLineWidth;
    spec.line_width_emu = std::llround(p.line_width_pt * static_cast<double>(formulon::kEmuPerPoint));
  }
  if (is_set(p.rotation_deg_op)) {
    if (!std::isfinite(p.rotation_deg) || std::fabs(p.rotation_deg) > kMaxRotationDeg) {
      return invalid_argument("rotation is outside its range", "field=rotation_deg");
    }
    double turned = std::fmod(p.rotation_deg, kFullTurnDeg);
    if (turned < 0.0) {
      turned += kFullTurnDeg;
    }
    const std::int64_t units = std::llround(turned * kRotationUnitsPerDegree);
    spec.fields |= field::kRotation;
    spec.rotation =
        static_cast<std::int32_t>(units % static_cast<std::int64_t>(kFullTurnDeg * kRotationUnitsPerDegree));
  }
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

extern "C" fm_status_t fm_sheet_get_shape_at(const fm_workbook_t* wb, size_t sheet_index, size_t idx,
                                             fm_shape_info_t* out) {
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
  *out = fm_shape_info_t{};
  out->struct_size = struct_size;
  out->object_id = obj.object_id;
  out->kind = static_cast<int32_t>(obj.kind);
  out->edit_as = to_c(obj.edit_as);
  fill_anchor(obj, out);
  out->width_emu = obj.cx;
  out->height_emu = obj.cy;
  out->fill_kind = to_c(info.fill);
  out->fill_color = paint_color_to_c(info.fill, info.fill_color);
  out->line_kind = to_c(info.line);
  out->line_color = paint_color_to_c(info.line, info.line_color);
  out->line_width_pt = static_cast<double>(info.line_width_emu) / static_cast<double>(formulon::kEmuPerPoint);
  out->rotation_deg = static_cast<double>(info.rotation) / kRotationUnitsPerDegree;
  out->hidden = info.hidden ? 1 : 0;
  out->text_box = info.text_box ? 1 : 0;
  out->name = store_cstr(store, obj.name);
  out->preset_geometry = store_cstr(store, info.prst);
  out->text = store_cstr(store, info.text);
  return 0;
}

extern "C" fm_status_t fm_sheet_get_shape(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                          int32_t* out_present, fm_shape_info_t* out) {
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
  *out = fm_shape_info_t{};
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

extern "C" fm_status_t fm_sheet_add_shape(fm_workbook_t* wb, size_t sheet_index, const fm_shape_patch_t* spec,
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
  if (auto rc = to_spec(*spec, true, nullptr, model); rc != 0) {
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
                                             const fm_shape_patch_t* patch) {
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
  auto shapes = formulon::list_shapes(wb->workbook(), sheet_index);
  if (!shapes) {
    return set_last_error(shapes.error());
  }
  const std::vector<formulon::ShapeInfo>& list = shapes.value();
  const auto it = std::find_if(list.begin(), list.end(),
                               [&](const formulon::ShapeInfo& s) { return s.object.object_id == object_id; });
  formulon::ShapeSpec model;
  if (auto rc = to_spec(*patch, false, it != list.end() ? &*it : nullptr, model); rc != 0) {
    return rc;
  }
  return finish_status(formulon::update_shape(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_remove_shape(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return finish_status(formulon::remove_shape(wb->workbook(), sheet_index, object_id));
}
