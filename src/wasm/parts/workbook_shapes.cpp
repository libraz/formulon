//
// JsWorkbook drawing-shape and sparkline-group surface: listing, inserting,
// updating and removing the shapes of a sheet's drawing, and the sparkline
// groups of its worksheet extension list.

#include <emscripten/val.h>

#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

constexpr const char* kSparklineColorKeys[] = {"colorSeries", "colorNegative", "colorAxis", "colorMarkers",
                                               "colorFirst",  "colorLast",     "colorHigh", "colorLow"};

/// True when `owner[key]` is neither undefined nor null.
bool present(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* key) {
  return reader.ok() && js_value_present(reader.value(owner, key, key));
}

emscripten::val shape_to_val(const fm_shape_info_t& s) {
  emscripten::val o = emscripten::val::object();
  o.set("objectId", s.object_id);
  o.set("kind", s.kind);
  o.set("anchorKind", s.anchor_kind);
  o.set("editAs", s.edit_as);
  o.set("fromRow", s.from_row);
  o.set("fromCol", s.from_col);
  o.set("fromRowOffsetEmu", static_cast<double>(s.from_row_offset_emu));
  o.set("fromColOffsetEmu", static_cast<double>(s.from_col_offset_emu));
  o.set("toRow", s.to_row);
  o.set("toCol", s.to_col);
  o.set("toRowOffsetEmu", static_cast<double>(s.to_row_offset_emu));
  o.set("toColOffsetEmu", static_cast<double>(s.to_col_offset_emu));
  o.set("widthEmu", static_cast<double>(s.width_emu));
  o.set("heightEmu", static_cast<double>(s.height_emu));
  js_set_cstr_fields(o, {{"name", s.name}, {"presetGeometry", s.preset_geometry}, {"text", s.text}});
  o.set("textBox", s.text_box != 0);
  js_set_index_or_null(o, "fillKind", s.fill_kind);
  o.set("fillColor", js_color_spec_or_null(s.fill_color));
  js_set_index_or_null(o, "lineKind", s.line_kind);
  o.set("lineColor", js_color_spec_or_null(s.line_color));
  o.set("lineWidthPt", s.line_width_pt);
  o.set("rotationDeg", s.rotation_deg);
  o.set("hidden", s.hidden != 0);
  return o;
}

struct ShapeText {
  std::string name;
  std::string preset_geometry;
  std::string text;
};

/// Reads the fields `spec` carries; an omitted field is `KEEP`, `null` is `CLEAR` (accepted on the four paint fields).
fm_shape_patch_t pull_shape(JsNarrowNumericReader& reader, const emscripten::val& spec, ShapeText& store) {
  fm_shape_patch_t s = FM_STRUCT_INIT(fm_shape_patch_t);
  const auto text = [&](const char* key, std::string& storage, fm_patch_op_t* op, const char** value) {
    const std::string field = std::string("shape.") + key;
    js_patch_string(reader, spec, key, field.c_str(), false, storage, op, value);
  };
  const auto u32 = [&](const char* key, fm_patch_op_t* op, uint32_t* value) {
    const std::string field = std::string("shape.") + key;
    js_patch_value(reader, spec, key, field.c_str(), op, value,
                   [&](const emscripten::val& v) { return reader.u32_value(v, 0U, field.c_str()); });
  };
  const auto i32 = [&](const char* key, fm_patch_op_t* op, int32_t* value) {
    const std::string field = std::string("shape.") + key;
    js_patch_value(reader, spec, key, field.c_str(), op, value,
                   [&](const emscripten::val& v) { return reader.i32_value(v, 0, field.c_str()); });
  };
  const auto i64 = [&](const char* key, fm_patch_op_t* op, int64_t* value) {
    const std::string field = std::string("shape.") + key;
    js_patch_value(reader, spec, key, field.c_str(), op, value,
                   [&](const emscripten::val&) { return reader.i64(spec, key, 0, field.c_str()); });
  };
  const auto f64 = [&](const char* key, fm_patch_op_t* op, double* value) {
    const std::string field = std::string("shape.") + key;
    js_patch_value(reader, spec, key, field.c_str(), op, value,
                   [&](const emscripten::val& v) { return reader.number_value(v, 0.0, field.c_str()); });
  };
  const auto color = [&](const char* key, fm_patch_op_t* op, fm_color_t* value) {
    const std::string field = std::string("shape.") + key;
    js_patch_value(reader, spec, key, field.c_str(), op, value,
                   [&](const emscripten::val& v) { return js_pull_color_spec_value(v, field.c_str(), &reader); });
  };
  text("name", store.name, &s.name_op, &s.name);
  text("presetGeometry", store.preset_geometry, &s.preset_geometry_op, &s.preset_geometry);
  text("text", store.text, &s.text_op, &s.text);
  js_patch_value(reader, spec, "textBox", "shape.textBox", &s.text_box_op, &s.text_box,
                 [&](const emscripten::val& v) { return reader.boolean_value(v, false) ? 1 : 0; });
  i32("anchorKind", &s.anchor_kind_op, &s.anchor_kind);
  i32("editAs", &s.edit_as_op, &s.edit_as);
  u32("fromRow", &s.from_row_op, &s.from_row);
  u32("fromCol", &s.from_col_op, &s.from_col);
  i64("fromRowOffsetEmu", &s.from_row_offset_emu_op, &s.from_row_offset_emu);
  i64("fromColOffsetEmu", &s.from_col_offset_emu_op, &s.from_col_offset_emu);
  u32("toRow", &s.to_row_op, &s.to_row);
  u32("toCol", &s.to_col_op, &s.to_col);
  i64("toRowOffsetEmu", &s.to_row_offset_emu_op, &s.to_row_offset_emu);
  i64("toColOffsetEmu", &s.to_col_offset_emu_op, &s.to_col_offset_emu);
  i64("widthEmu", &s.width_emu_op, &s.width_emu);
  i64("heightEmu", &s.height_emu_op, &s.height_emu);
  i32("fillKind", &s.fill_kind_op, &s.fill_kind);
  color("fillColor", &s.fill_color_op, &s.fill_color);
  i32("lineKind", &s.line_kind_op, &s.line_kind);
  color("lineColor", &s.line_color_op, &s.line_color);
  f64("lineWidthPt", &s.line_width_pt_op, &s.line_width_pt);
  f64("rotationDeg", &s.rotation_deg_op, &s.rotation_deg);
  js_patch_value(reader, spec, "hidden", "shape.hidden", &s.hidden_op, &s.hidden,
                 [&](const emscripten::val& v) { return reader.boolean_value(v, false) ? 1 : 0; });
  return s;
}

fm_color_t* group_color(fm_sparkline_group_t& g, std::size_t i) {
  fm_color_t* const colors[] = {&g.color_series, &g.color_negative, &g.color_axis, &g.color_markers,
                                &g.color_first,  &g.color_last,     &g.color_high, &g.color_low};
  return colors[i];
}

emscripten::val sparkline_group_to_val(fm_sparkline_group_t g, emscripten::val sparklines) {
  emscripten::val o = emscripten::val::object();
  for (std::size_t i = 0; i < 8U; ++i) {
    o.set(kSparklineColorKeys[i], js_color_spec_or_null(*group_color(g, i)));
  }
  o.set("type", g.type);
  o.set("markers", g.markers != 0);
  o.set("high", g.high != 0);
  o.set("low", g.low != 0);
  o.set("first", g.first != 0);
  o.set("last", g.last != 0);
  o.set("negative", g.negative != 0);
  o.set("displayXAxis", g.display_x_axis != 0);
  o.set("displayHidden", g.display_hidden != 0);
  o.set("minAxisType", g.min_axis_type);
  o.set("maxAxisType", g.max_axis_type);
  o.set("manualMin", g.has_manual_min != 0 ? emscripten::val(g.manual_min) : emscripten::val::null());
  o.set("manualMax", g.has_manual_max != 0 ? emscripten::val(g.manual_max) : emscripten::val::null());
  o.set("dateAxis", g.date_axis != 0);
  js_set_cstr_or_null(o, "dateAxisRef", g.date_axis_ref);
  o.set("lineWeight", g.line_weight);
  o.set("emptyCellsAs", g.empty_cells_as);
  o.set("sparklines", sparklines);
  return o;
}

/// Reads a group spec; absent keys take the defaults of a new line group.
fm_sparkline_group_t pull_group(JsNarrowNumericReader& reader, const emscripten::val& spec, std::string& ref_store) {
  fm_sparkline_group_t g = FM_STRUCT_INIT(fm_sparkline_group_t);
  for (std::size_t i = 0; i < 8U; ++i) {
    *group_color(g, i) = js_pull_color_spec(spec, kSparklineColorKeys[i], &reader);
  }
  g.type = reader.i32(spec, "type", FM_SPARKLINE_TYPE_LINE, "group.type");
  g.markers = reader.boolean(spec, "markers", false, "group.markers") ? 1 : 0;
  g.high = reader.boolean(spec, "high", false, "group.high") ? 1 : 0;
  g.low = reader.boolean(spec, "low", false, "group.low") ? 1 : 0;
  g.first = reader.boolean(spec, "first", false, "group.first") ? 1 : 0;
  g.last = reader.boolean(spec, "last", false, "group.last") ? 1 : 0;
  g.negative = reader.boolean(spec, "negative", false, "group.negative") ? 1 : 0;
  g.display_x_axis = reader.boolean(spec, "displayXAxis", false, "group.displayXAxis") ? 1 : 0;
  g.display_hidden = reader.boolean(spec, "displayHidden", false, "group.displayHidden") ? 1 : 0;
  g.min_axis_type = reader.i32(spec, "minAxisType", FM_SPARKLINE_AXIS_TYPE_INDIVIDUAL, "group.minAxisType");
  g.max_axis_type = reader.i32(spec, "maxAxisType", FM_SPARKLINE_AXIS_TYPE_INDIVIDUAL, "group.maxAxisType");
  if (present(reader, spec, "manualMin")) {
    g.has_manual_min = 1;
    g.manual_min = reader.number(spec, "manualMin", 0.0, "group.manualMin");
  }
  if (present(reader, spec, "manualMax")) {
    g.has_manual_max = 1;
    g.manual_max = reader.number(spec, "manualMax", 0.0, "group.manualMax");
  }
  g.date_axis = reader.boolean(spec, "dateAxis", false, "group.dateAxis") ? 1 : 0;
  g.date_axis_ref = reader.optional_string(spec, "dateAxisRef", ref_store, "group.dateAxisRef");
  g.line_weight = reader.number(spec, "lineWeight", 0.75, "group.lineWeight");
  g.empty_cells_as = reader.i32(spec, "emptyCellsAs", FM_BLANKS_AS_GAP, "group.emptyCellsAs");
  return g;
}

/// Reads `[{ formula, location }]`; `texts` keeps the strings `items` points into.
bool pull_sparklines(JsNarrowNumericReader& reader, const emscripten::val& array, std::vector<std::string>& texts,
                     std::vector<fm_sparkline_t>& items) {
  if (!reader.is_array(array, "sparklines")) {
    return false;
  }
  const uint32_t count = reader.length(array, "sparklines");
  texts.reserve(2U * count);
  for (uint32_t i = 0; reader.ok() && i < count; ++i) {
    const emscripten::val item = reader.array_element(array, i, "sparklines");
    texts.push_back(reader.string(item, "formula", "sparkline.formula"));
    texts.push_back(reader.string(item, "location", "sparkline.location"));
  }
  for (std::size_t i = 0; reader.ok() && i < texts.size(); i += 2U) {
    items.push_back(fm_sparkline_t{texts[i].c_str(), texts[i + 1U].c_str()});
  }
  return reader.ok();
}

fm_status_t shape_item(const fm_workbook_t* handle, size_t sheet, size_t index, emscripten::val* out) {
  fm_shape_info_t s = FM_STRUCT_INIT(fm_shape_info_t);
  const fm_status_t rc = fm_sheet_get_shape_at(handle, sheet, index, &s);
  if (rc == 0) {
    *out = shape_to_val(s);
  }
  return rc;
}

// Reads sparkline group `index` with its sparklines nested under `sparklines`;
// a failed read leaves `*out` untouched.
fm_status_t sparkline_group_item(const fm_workbook_t* handle, size_t sheet, size_t index, emscripten::val* out) {
  fm_sparkline_group_t g = FM_STRUCT_INIT(fm_sparkline_group_t);
  fm_status_t rc = fm_sheet_get_sparkline_group_at(handle, sheet, index, &g);
  if (rc != 0) {
    return rc;
  }
  emscripten::val group = sparkline_group_to_val(g, emscripten::val::array());
  for (size_t k = 0; rc == 0 && k < g.sparkline_count; ++k) {
    fm_sparkline_t item{};
    rc = fm_sheet_get_sparkline_at(handle, sheet, index, k, &item);
    if (rc == 0) {
      emscripten::val o = emscripten::val::object();
      js_set_cstr_fields(o, {{"formula", item.formula}, {"location", item.location}});
      group["sparklines"].call<void>("push", o);
    }
  }
  if (rc == 0) {
    *out = group;
  }
  return rc;
}

}  // namespace

emscripten::val JsWorkbook::listShapes(uint32_t sheet) const {
  emscripten::val arr = emscripten::val::array();
  if (!succeeded(js_collect_list(handle_, sheet, &fm_sheet_get_shape_count, &shape_item, arr))) {
    return emscripten::val::undefined();
  }
  return arr;
}

uint32_t JsWorkbook::insertShape(uint32_t sheet, emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  JsNarrowNumericReader reader("Sheet.addShape");
  ShapeText store;
  const fm_shape_patch_t shape = pull_shape(reader, spec, store);
  if (!reader.ok()) {
    return 0U;
  }
  uint32_t id = 0;
  if (!succeeded(fm_sheet_add_shape(handle_, sheet, &shape, &id))) {
    return 0U;
  }
  return id;
}

void JsWorkbook::updateShape(uint32_t sheet, uint32_t objectId, emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Shape.update");
  ShapeText store;
  const fm_shape_patch_t shape = pull_shape(reader, spec, store);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_shape(handle_, sheet, objectId, &shape));
}

emscripten::val JsWorkbook::getShape(uint32_t sheet, uint32_t objectId) const {
  fm_shape_info_t s = FM_STRUCT_INIT(fm_shape_info_t);
  int32_t present = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_shape(handle_, sheet, objectId, &present, &s) : kBindingInvalidHandle;
  return optional_or_null(rc, present, rc == 0 && present != 0 ? shape_to_val(s) : emscripten::val::null());
}

void JsWorkbook::removeShape(uint32_t sheet, uint32_t objectId) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_shape(handle_, sheet, objectId) : kBindingInvalidHandle);
}

double JsWorkbook::getSparklineGroupCount(uint32_t sheet) const {
  return js_count(handle_, sheet, &fm_sheet_get_sparkline_group_count);
}

emscripten::val JsWorkbook::getSparklineGroup(uint32_t sheet, uint32_t sparklineGroupIndex) const {
  emscripten::val group = emscripten::val::undefined();
  const fm_status_t rc =
      handle_ != nullptr ? sparkline_group_item(handle_, sheet, sparklineGroupIndex, &group) : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  return group;
}

uint32_t JsWorkbook::addSparklineGroup(uint32_t sheet, emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  JsNarrowNumericReader reader("Sheet.addSparklineGroup");
  std::string ref_store;
  std::vector<std::string> texts;
  std::vector<fm_sparkline_t> items;
  const fm_sparkline_group_t group = pull_group(reader, spec, ref_store);
  if (reader.ok()) {
    pull_sparklines(reader, reader.value(spec, "sparklines", "group.sparklines"), texts, items);
  }
  if (!reader.ok()) {
    return 0U;
  }
  // An empty array reaches the core, which rejects it; keep a valid pointer to pass.
  fm_sparkline_t none{};
  std::size_t out = 0;
  if (!succeeded(fm_sheet_add_sparkline_group(handle_, sheet, &group, items.empty() ? &none : items.data(),
                                              items.size(), &out))) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

void JsWorkbook::updateSparklineGroup(uint32_t sheet, uint32_t index, emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Sheet.updateSparklineGroup");
  std::string ref_store;
  std::vector<std::string> texts;
  std::vector<fm_sparkline_t> items;
  fm_sparkline_group_patch_t p = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  fm_patch_op_t* const color_ops[] = {&p.color_series_op, &p.color_negative_op, &p.color_axis_op, &p.color_markers_op,
                                      &p.color_first_op,  &p.color_last_op,     &p.color_high_op, &p.color_low_op};
  fm_color_t* const colors[] = {&p.color_series, &p.color_negative, &p.color_axis, &p.color_markers,
                                &p.color_first,  &p.color_last,     &p.color_high, &p.color_low};
  for (std::size_t i = 0; i < 8U; ++i) {
    const std::string field = std::string("group.") + kSparklineColorKeys[i];
    js_patch_value(reader, spec, kSparklineColorKeys[i], field.c_str(), color_ops[i], colors[i],
                   [&](const emscripten::val& v) { return js_pull_color_spec_value(v, field.c_str(), &reader); });
  }
  const auto num = [&](const char* key, fm_patch_op_t* op, double* value) {
    const std::string field = std::string("group.") + key;
    js_patch_value(reader, spec, key, field.c_str(), op, value,
                   [&](const emscripten::val& v) { return reader.number_value(v, 0.0, field.c_str()); });
  };
  const auto i32 = [&](const char* key, fm_patch_op_t* op, int32_t* value) {
    const std::string field = std::string("group.") + key;
    js_patch_value(reader, spec, key, field.c_str(), op, value,
                   [&](const emscripten::val& v) { return reader.i32_value(v, 0, field.c_str()); });
  };
  const auto flag = [&](const char* key, fm_patch_op_t* op, int32_t* value) {
    const std::string field = std::string("group.") + key;
    js_patch_value(reader, spec, key, field.c_str(), op, value,
                   [&](const emscripten::val& v) { return reader.boolean_value(v, false) ? 1 : 0; });
  };
  num("manualMin", &p.manual_min_op, &p.manual_min);
  num("manualMax", &p.manual_max_op, &p.manual_max);
  num("lineWeight", &p.line_weight_op, &p.line_weight);
  i32("type", &p.type_op, &p.type);
  flag("markers", &p.markers_op, &p.markers);
  flag("high", &p.high_op, &p.high);
  flag("low", &p.low_op, &p.low);
  flag("first", &p.first_op, &p.first);
  flag("last", &p.last_op, &p.last);
  flag("negative", &p.negative_op, &p.negative);
  flag("displayXAxis", &p.display_x_axis_op, &p.display_x_axis);
  flag("displayHidden", &p.display_hidden_op, &p.display_hidden);
  i32("minAxisType", &p.min_axis_type_op, &p.min_axis_type);
  i32("maxAxisType", &p.max_axis_type_op, &p.max_axis_type);
  flag("dateAxis", &p.date_axis_op, &p.date_axis);
  i32("emptyCellsAs", &p.empty_cells_as_op, &p.empty_cells_as);
  js_patch_string(reader, spec, "dateAxisRef", "group.dateAxisRef", true, ref_store, &p.date_axis_ref_op,
                  &p.date_axis_ref);
  const emscripten::val sparklines = reader.value(spec, "sparklines", "group.sparklines");
  if (reader.ok() && js_value_present(sparklines) && pull_sparklines(reader, sparklines, texts, items)) {
    p.items_op = FM_PATCH_OP_SET;
    p.items = items.data();
    p.items_count = static_cast<uint32_t>(items.size());
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_sparkline_group(handle_, sheet, index, &p));
}

void JsWorkbook::removeSparklineGroup(uint32_t sheet, uint32_t index) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_sparkline_group(handle_, sheet, index) : kBindingInvalidHandle);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
