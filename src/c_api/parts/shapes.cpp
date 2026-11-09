//
// C ABI - the shapes, connectors and groups of a sheet's drawing: listing,
// inserting, updating and removing them.

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
#include "c_api/parts/struct_v1_sizes.h"
#include "drawing/drawing_edit.h"
#include "drawing/shape_edit.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::AnchorPlan;
using formulon::c_api::parts::apply_placement_extras;
using formulon::c_api::parts::chart_fill_to_c;
using formulon::c_api::parts::chart_line_to_c;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::drawing_patch_from_c;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::PatchReader;
using formulon::c_api::parts::plan_add_anchor;
using formulon::c_api::parts::read_fill;
using formulon::c_api::parts::read_line;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::TextStore;

namespace {

/// `a:xfrm/@rot` counts 60000ths of a degree.
constexpr double kRotationUnitsPerDegree = 60000.0;
constexpr double kMaxRotationDeg = 3600.0;
constexpr double kFullTurnDeg = 360.0;

bool is_set(fm_patch_op_t op) {
  return op == FM_PATCH_OP_SET;
}

// The ops, strings and UTF-8 of the non-paint fields of `p`; none of them can be cleared.
fm_status_t check_patch(const fm_shape_patch_t& p) {
  const std::array<std::pair<fm_patch_op_t, const char*>, 2> plain = {
      {{p.rotation_deg_op, "rotation_deg"}, {p.text_box_op, "text_box"}}};
  for (const auto& [op, name] : plain) {
    if (!patch_op_ok_without_clear(op)) {
      return invalid_argument(patch_op_ok(op) ? "field cannot be cleared" : "op must be KEEP or SET",
                              std::string("field=") + name);
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

// Converts `p` to the model spec's shape fields; a non-zero return is the status to report.
fm_status_t to_spec(const fm_shape_patch_t& p, formulon::ShapeSpec& spec) {
  namespace field = formulon::shape_field;
  if (auto rc = check_patch(p); rc != 0) {
    return rc;
  }
  PatchReader r;
  spec.paint_fill = read_fill(r, p.fill);
  spec.paint_line = read_line(r, p.line);
  if (r.status() != 0) {
    return r.status();
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

// The fill record of a spec as the patch that writes it: each stated member is SET; all-zero states nothing.
fm_status_t fill_patch_of(const fm_chart_fill_t& f, fm_chart_fill_patch_t& out) {
  if (f.other != 0) {
    return invalid_argument("other is a read-only fill state", "field=fill.other");
  }
  if (auto rc = check_flag(f.has_transparency_pct, "fill.has_transparency_pct"); rc != 0) {
    return rc;
  }
  if (auto rc = check_flag(f.none, "fill.none"); rc != 0) {
    return rc;
  }
  out = fm_chart_fill_patch_t{};
  if (f.color.kind != FM_COLOR_KIND_NONE) {
    out.color_op = FM_PATCH_OP_SET;
    out.color = f.color;
  }
  if (f.has_transparency_pct != 0) {
    out.transparency_pct_op = FM_PATCH_OP_SET;
    out.transparency_pct = f.transparency_pct;
  }
  if (f.none != 0) {
    out.none_op = FM_PATCH_OP_SET;
    out.none = 1;
  }
  return 0;
}

// The outline record of a spec as the patch that writes it, as `fill_patch_of`.
fm_status_t line_patch_of(const fm_chart_line_t& l, fm_chart_line_patch_t& out) {
  if (l.other != 0) {
    return invalid_argument("other is a read-only outline state", "field=line.other");
  }
  const std::array<std::pair<std::int32_t, const char*>, 4> flags = {
      {{l.has_transparency_pct, "line.has_transparency_pct"},
       {l.has_width_pt, "line.has_width_pt"},
       {l.has_dash, "line.has_dash"},
       {l.none, "line.none"}}};
  for (const auto& [value, name] : flags) {
    if (auto rc = check_flag(value, name); rc != 0) {
      return rc;
    }
  }
  out = fm_chart_line_patch_t{};
  if (l.color.kind != FM_COLOR_KIND_NONE) {
    out.color_op = FM_PATCH_OP_SET;
    out.color = l.color;
  }
  if (l.has_transparency_pct != 0) {
    out.transparency_pct_op = FM_PATCH_OP_SET;
    out.transparency_pct = l.transparency_pct;
  }
  if (l.has_width_pt != 0) {
    out.width_pt_op = FM_PATCH_OP_SET;
    out.width_pt = l.width_pt;
  }
  if (l.has_dash != 0) {
    out.dash_op = FM_PATCH_OP_SET;
    out.dash = l.dash;
  }
  if (l.none != 0) {
    out.none_op = FM_PATCH_OP_SET;
    out.none = 1;
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
    out->fill = chart_fill_to_c(info.paint_fill);
    out->line = chart_line_to_c(info.paint_line);
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
  AnchorPlan plan;
  if (auto rc =
          plan_add_anchor(wb->workbook(), sheet_index, placement.anchor, formulon::AnchorKind::kTwoCell, true, &plan);
      rc != 0) {
    return rc;
  }
  // Every field of the spec is a value: the patch that carries them all.
  fm_shape_patch_t fields{};
  fields.struct_size = sizeof(fields);
  if (auto rc = fill_patch_of(spec->fill, fields.fill); rc != 0) {
    return rc;
  }
  if (auto rc = line_patch_of(spec->line, fields.line); rc != 0) {
    return rc;
  }
  fields.rotation_deg_op = FM_PATCH_OP_SET;
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
  if (auto rc = to_spec(fields, model); rc != 0) {
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
                                             const fm_drawing_patch_t* common, const fm_shape_patch_t* shape) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  formulon::DrawingCommonPatch common_model;
  if (common != nullptr) {
    if (auto rc = drawing_patch_from_c(*common, &common_model); rc != 0) {
      return rc;
    }
  }
  formulon::ShapeSpec shape_model;
  if (shape != nullptr) {
    if (auto rc = check_record_size(*shape); rc != 0) {
      return rc;
    }
    if (auto rc = to_spec(*shape, shape_model); rc != 0) {
      return rc;
    }
  }
  return finish_status(formulon::update_shape(wb->workbook(), sheet_index, object_id,
                                              common != nullptr ? &common_model : nullptr,
                                              shape != nullptr ? &shape_model : nullptr));
}
