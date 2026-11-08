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

emscripten::val shape_to_val(const fm_shape_t& s) {
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

/// Reads the fields `spec` carries; the `fields` mask names exactly those.
fm_shape_t pull_shape(JsNarrowNumericReader& reader, const emscripten::val& spec, ShapeText& store) {
  fm_shape_t s = FM_STRUCT_INIT(fm_shape_t);
  if (present(reader, spec, "name")) {
    s.fields |= FM_SHAPE_FIELD_NAME;
    s.name = reader.optional_string(spec, "name", store.name, "shape.name");
  }
  if (present(reader, spec, "presetGeometry")) {
    s.fields |= FM_SHAPE_FIELD_PRESET_GEOMETRY;
    s.preset_geometry = reader.optional_string(spec, "presetGeometry", store.preset_geometry, "shape.presetGeometry");
  }
  if (present(reader, spec, "text")) {
    s.fields |= FM_SHAPE_FIELD_TEXT;
    s.text = reader.optional_string(spec, "text", store.text, "shape.text");
  }
  if (present(reader, spec, "textBox")) {
    s.fields |= FM_SHAPE_FIELD_TEXT_BOX;
    s.text_box = reader.boolean(spec, "textBox", false, "shape.textBox") ? 1 : 0;
  }
  if (present(reader, spec, "anchorKind")) {
    s.fields |= FM_SHAPE_FIELD_ANCHOR;
    s.anchor_kind = reader.i32(spec, "anchorKind", 0, "shape.anchorKind");
    s.edit_as = reader.i32(spec, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL, "shape.editAs");
    s.from_row = reader.u32(spec, "fromRow", 0U, "shape.fromRow");
    s.from_col = reader.u32(spec, "fromCol", 0U, "shape.fromCol");
    s.from_row_offset_emu = reader.i64(spec, "fromRowOffsetEmu", 0, "shape.fromRowOffsetEmu");
    s.from_col_offset_emu = reader.i64(spec, "fromColOffsetEmu", 0, "shape.fromColOffsetEmu");
    s.to_row = reader.u32(spec, "toRow", 0U, "shape.toRow");
    s.to_col = reader.u32(spec, "toCol", 0U, "shape.toCol");
    s.to_row_offset_emu = reader.i64(spec, "toRowOffsetEmu", 0, "shape.toRowOffsetEmu");
    s.to_col_offset_emu = reader.i64(spec, "toColOffsetEmu", 0, "shape.toColOffsetEmu");
    s.width_emu = reader.i64(spec, "widthEmu", 0, "shape.widthEmu");
    s.height_emu = reader.i64(spec, "heightEmu", 0, "shape.heightEmu");
  }
  if (present(reader, spec, "fillKind")) {
    s.fields |= FM_SHAPE_FIELD_FILL;
    s.fill_kind = reader.i32(spec, "fillKind", 0, "shape.fillKind");
    s.fill_color = js_pull_color_spec(spec, "fillColor", &reader);
  }
  if (present(reader, spec, "lineKind")) {
    s.fields |= FM_SHAPE_FIELD_LINE;
    s.line_kind = reader.i32(spec, "lineKind", 0, "shape.lineKind");
    s.line_color = js_pull_color_spec(spec, "lineColor", &reader);
  }
  if (present(reader, spec, "lineWidthPt")) {
    s.fields |= FM_SHAPE_FIELD_LINE_WIDTH;
    s.line_width_pt = reader.number(spec, "lineWidthPt", 0.0, "shape.lineWidthPt");
  }
  if (present(reader, spec, "rotationDeg")) {
    s.fields |= FM_SHAPE_FIELD_ROTATION;
    s.rotation_deg = reader.number(spec, "rotationDeg", 0.0, "shape.rotationDeg");
  }
  if (present(reader, spec, "hidden")) {
    s.fields |= FM_SHAPE_FIELD_HIDDEN;
    s.hidden = reader.boolean(spec, "hidden", false, "shape.hidden") ? 1 : 0;
  }
  return s;
}

fm_color_spec_t* group_color(fm_sparkline_group_t& g, std::size_t i) {
  fm_color_spec_t* const colors[] = {&g.color_series, &g.color_negative, &g.color_axis, &g.color_markers,
                                     &g.color_first,  &g.color_last,     &g.color_high, &g.color_low};
  return colors[i];
}

emscripten::val sparkline_group_to_val(fm_sparkline_group_t g, emscripten::val sparklines) {
  emscripten::val o = emscripten::val::object();
  for (std::size_t i = 0; i < 8U; ++i) {
    o.set(kSparklineColorKeys[i], js_color_spec(*group_color(g, i)));
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
  g.empty_cells_as = reader.i32(spec, "emptyCellsAs", FM_SPARKLINE_EMPTY_CELLS_GAP, "group.emptyCellsAs");
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
  fm_shape_t s = FM_STRUCT_INIT(fm_shape_t);
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
  fm_status_t rc = fm_sheet_get_sparkline_group(handle, sheet, index, &g);
  if (rc != 0) {
    return rc;
  }
  emscripten::val group = sparkline_group_to_val(g, emscripten::val::array());
  for (size_t k = 0; rc == 0 && k < g.sparkline_count; ++k) {
    fm_sparkline_t item{};
    rc = fm_sheet_get_sparkline(handle, sheet, index, k, &item);
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
  JsNarrowNumericReader reader("addShape");
  ShapeText store;
  const fm_shape_t shape = pull_shape(reader, spec, store);
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
  JsNarrowNumericReader reader("updateShape");
  ShapeText store;
  const fm_shape_t shape = pull_shape(reader, spec, store);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_shape(handle_, sheet, objectId, &shape));
}

emscripten::val JsWorkbook::getShape(uint32_t sheet, uint32_t objectId) const {
  fm_shape_t s = FM_STRUCT_INIT(fm_shape_t);
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
  JsNarrowNumericReader reader("addSparklineGroup");
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
  JsNarrowNumericReader reader("updateSparklineGroup");
  std::string ref_store;
  std::vector<std::string> texts;
  std::vector<fm_sparkline_t> items;
  const fm_sparkline_group_t group = pull_group(reader, spec, ref_store);
  const emscripten::val sparklines = reader.value(spec, "sparklines", "group.sparklines");
  const bool replace = js_value_present(sparklines);
  if (reader.ok() && replace) {
    pull_sparklines(reader, sparklines, texts, items);
  }
  if (!reader.ok()) {
    return;
  }
  // Without `sparklines` the group keeps its own; an empty array reaches the core, which rejects it.
  fm_sparkline_t none{};
  const fm_sparkline_t* replacement = !replace ? nullptr : items.empty() ? &none : items.data();
  succeeded(fm_sheet_update_sparkline_group(handle_, sheet, index, &group, replacement, items.size()));
}

void JsWorkbook::removeSparklineGroup(uint32_t sheet, uint32_t index) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_sparkline_group(handle_, sheet, index) : kBindingInvalidHandle);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
