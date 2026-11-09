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
#include "c_api/parts/drawing_anchor.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "drawing/drawing_geometry.h"
#include "drawing/shape_edit.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::AnchorPlan;
using formulon::c_api::parts::apply_placement_extras;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::color_spec_from_c;
using formulon::c_api::parts::color_spec_to_c;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::plan_add_anchor;
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
};

// Every op of `p` with its field name, in declaration order.
std::array<OpField, 7> ops_of(const fm_shape_patch_t& p) {
  return {{{p.fill_kind_op, "fill_kind"},
           {p.fill_color_op, "fill_color"},
           {p.line_kind_op, "line_kind"},
           {p.line_color_op, "line_color"},
           {p.line_width_pt_op, "line_width_pt"},
           {p.rotation_deg_op, "rotation_deg"},
           {p.text_box_op, "text_box"}}};
}

// The ops, strings and UTF-8 of `p`; no field can be cleared.
fm_status_t check_patch(const fm_shape_patch_t& p) {
  for (const OpField& f : ops_of(p)) {
    if (!patch_op_ok_without_clear(f.op)) {
      return invalid_argument(patch_op_ok(f.op) ? "field cannot be cleared" : "op must be KEEP or SET",
                              std::string("field=") + f.name);
    }
  }
  const std::array<std::pair<fm_patch_op_t, const char*>, 2> texts = {
      {{p.preset_geometry_op, p.preset_geometry}, {p.text_op, p.text}}};
  const char* const names[] = {"preset_geometry", "text"};
  for (std::size_t i = 0; i < texts.size(); ++i) {
    if (!patch_op_ok_without_clear(texts[i].first)) {
      return invalid_argument(patch_op_ok(texts[i].first) ? "field cannot be cleared" : "op must be KEEP or SET",
                              std::string("field=") + names[i]);
    }
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
  if (!is_set(kind_op) && !is_set(color_op)) {
    return 0;
  }
  if (is_set(kind_op) && kind != FM_SHAPE_PAINT_DEFAULT && kind != FM_SHAPE_PAINT_SOLID &&
      kind != FM_SHAPE_PAINT_NONE) {
    return invalid_argument("a paint is default, solid or none", std::string("field=") + kind_name);
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

// Converts `p` to the model spec's shape fields. `stored` is the shape an update merges into (NULL on add, and
// on an update whose id names no shape, which the engine then refuses); a non-zero return is the status to report.
fm_status_t to_spec(const fm_shape_patch_t& p, bool add, const formulon::ShapeInfo* stored, formulon::ShapeSpec& spec) {
  namespace field = formulon::shape_field;
  if (auto rc = check_patch(p); rc != 0) {
    return rc;
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
  for (const formulon::ShapeInfo& info : shapes.value()) {
    if (info.object.object_id != object_id) {
      continue;
    }
    TextStore& store = fresh_read_scratch(wb);
    out->object_id = info.object.object_id;
    out->kind = static_cast<int32_t>(info.object.kind);
    out->fill_kind = to_c(info.fill);
    out->fill_color = paint_color_to_c(info.fill, info.fill_color);
    out->line_kind = to_c(info.line);
    out->line_color = paint_color_to_c(info.line, info.line_color);
    out->line_width_pt = static_cast<double>(info.line_width_emu) / static_cast<double>(formulon::kEmuPerPoint);
    out->rotation_deg = static_cast<double>(info.rotation) / kRotationUnitsPerDegree;
    out->text_box = info.text_box ? 1 : 0;
    out->preset_geometry = store_cstr(store, info.prst);
    out->text = store_cstr(store, info.text);
    *out_present = 1;
    break;
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_add_shape(fm_workbook_t* wb, size_t sheet_index, const fm_shape_spec_t* spec,
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
  if (auto rc = check_record_size(spec->placement); rc != 0) {
    return rc;
  }
  const fm_drawing_placement_spec_t& placement = spec->placement;
  if (auto rc = check_utf8({placement.name, placement.description}); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(placement.hidden, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "placement.hidden");
      rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(spec->fill_kind, FM_SHAPE_PAINT_DEFAULT, FM_SHAPE_PAINT_NONE, "fill_kind"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(spec->line_kind, FM_SHAPE_PAINT_DEFAULT, FM_SHAPE_PAINT_NONE, "line_kind"); rc != 0) {
    return rc;
  }
  AnchorPlan plan;
  if (auto rc =
          plan_add_anchor(wb->workbook(), sheet_index, placement.anchor, formulon::AnchorKind::kTwoCell, true, &plan);
      rc != 0) {
    return rc;
  }
  // Every field of the spec is a value: the patch that carries them all.
  fm_shape_patch_t fields{};
  fields.struct_size = sizeof(fields);
  fields.fill_kind_op = fields.fill_color_op = FM_PATCH_OP_SET;
  fields.fill_kind = spec->fill_kind;
  fields.fill_color = spec->fill_color;
  fields.line_kind_op = fields.line_color_op = FM_PATCH_OP_SET;
  fields.line_kind = spec->line_kind;
  fields.line_color = spec->line_color;
  fields.line_width_pt_op = fields.rotation_deg_op = FM_PATCH_OP_SET;
  fields.line_width_pt = spec->line_width_pt;
  fields.rotation_deg = spec->rotation_deg;
  // A connector or group holds no text, so the flag is written only when it is on.
  if (spec->text_box != 0) {
    fields.text_box_op = FM_PATCH_OP_SET;
    fields.text_box = spec->text_box;
  }
  if (spec->preset_geometry != nullptr && spec->preset_geometry[0] != '\0') {
    fields.preset_geometry_op = FM_PATCH_OP_SET;
    fields.preset_geometry = spec->preset_geometry;
  }
  if (spec->text != nullptr) {
    fields.text_op = FM_PATCH_OP_SET;
    fields.text = spec->text;
  }
  formulon::ShapeSpec model;
  if (auto rc = to_spec(fields, true, nullptr, model); rc != 0) {
    return rc;
  }
  namespace field = formulon::shape_field;
  model.fields |= field::kAnchor;
  model.anchor_kind = plan.kind;
  model.edit_as = plan.edit_as;
  model.from = plan.from;
  model.to = plan.to;
  model.cx = plan.cx;
  model.cy = plan.cy;
  if (placement.name != nullptr && placement.name[0] != '\0') {
    model.fields |= field::kName;
    model.name = placement.name;
  }
  if (placement.hidden == FM_TRISTATE_TRUE) {
    model.fields |= field::kHidden;
    model.hidden = true;
  }
  auto id = formulon::insert_shape(wb->workbook(), sheet_index, model);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  fm_drawing_placement_spec_t description_only{};
  description_only.description = placement.description;
  return apply_placement_extras(wb->workbook(), sheet_index, id.value(), description_only, true);
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
