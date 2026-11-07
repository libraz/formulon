// Drawing-shape and sparkline-group bindings: list, insert, update and remove
// the shapes of a sheet's drawing, and the sparkline groups of its worksheet
// extension list.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

constexpr const char* kSparklineColorKeys[] = {"colorSeries", "colorNegative", "colorAxis", "colorMarkers",
                                               "colorFirst",  "colorLast",     "colorHigh", "colorLow"};

fm_color_spec_t* GroupColor(fm_sparkline_group_t& g, size_t i) {
  fm_color_spec_t* const colors[] = {&g.color_series, &g.color_negative, &g.color_axis, &g.color_markers,
                                     &g.color_first,  &g.color_last,     &g.color_high, &g.color_low};
  return colors[i];
}

struct ShapeText {
  std::string name;
  std::string preset_geometry;
  std::string text;
};

// Reads the fields `spec` carries; the `fields` mask names exactly those.
fm_shape_t PullShape(CheckedSpecReader& reader, const Napi::Object& spec, ShapeText& store) {
  fm_shape_t s = FM_STRUCT_INIT(fm_shape_t);
  bool present = false;
  if (reader.String(spec, "name", &store.name)) {
    s.fields |= FM_SHAPE_FIELD_NAME;
    s.name = store.name.c_str();
  }
  if (reader.String(spec, "presetGeometry", &store.preset_geometry)) {
    s.fields |= FM_SHAPE_FIELD_PRESET_GEOMETRY;
    s.preset_geometry = store.preset_geometry.c_str();
  }
  if (reader.String(spec, "text", &store.text)) {
    s.fields |= FM_SHAPE_FIELD_TEXT;
    s.text = store.text.c_str();
  }
  const bool text_box = reader.Bool(spec, "textBox", false, &present);
  if (present) {
    s.fields |= FM_SHAPE_FIELD_TEXT_BOX;
    s.text_box = text_box ? 1 : 0;
  }
  s.anchor_kind = reader.I32(spec, "anchorKind", 0, &present);
  if (present) {
    s.fields |= FM_SHAPE_FIELD_ANCHOR;
    s.edit_as = reader.I32(spec, "editAs", FM_ANCHOR_EDIT_AS_TWO_CELL);
    s.from_row = reader.U32(spec, "fromRow", 0U);
    s.from_col = reader.U32(spec, "fromCol", 0U);
    s.from_row_offset_emu = reader.I64(spec, "fromRowOffsetEmu", 0);
    s.from_col_offset_emu = reader.I64(spec, "fromColOffsetEmu", 0);
    s.to_row = reader.U32(spec, "toRow", 0U);
    s.to_col = reader.U32(spec, "toCol", 0U);
    s.to_row_offset_emu = reader.I64(spec, "toRowOffsetEmu", 0);
    s.to_col_offset_emu = reader.I64(spec, "toColOffsetEmu", 0);
    s.width_emu = reader.I64(spec, "widthEmu", 0);
    s.height_emu = reader.I64(spec, "heightEmu", 0);
  }
  s.fill_kind = reader.I32(spec, "fillKind", 0, &present);
  if (present) {
    s.fields |= FM_SHAPE_FIELD_FILL;
    s.fill_color = PullColorSpec(reader, spec, "fillColor");
  }
  s.line_kind = reader.I32(spec, "lineKind", 0, &present);
  if (present) {
    s.fields |= FM_SHAPE_FIELD_LINE;
    s.line_color = PullColorSpec(reader, spec, "lineColor");
  }
  s.line_width_emu = reader.I64(spec, "lineWidthEmu", 0, &present);
  s.fields |= present ? FM_SHAPE_FIELD_LINE_WIDTH : 0U;
  s.rotation = reader.I32(spec, "rotation", 0, &present);
  s.fields |= present ? FM_SHAPE_FIELD_ROTATION : 0U;
  const bool hidden = reader.Bool(spec, "hidden", false, &present);
  if (present) {
    s.fields |= FM_SHAPE_FIELD_HIDDEN;
    s.hidden = hidden ? 1 : 0;
  }
  return s;
}

// Reads a group spec; absent keys take the defaults of a new line group.
fm_sparkline_group_t PullGroup(CheckedSpecReader& reader, const Napi::Object& spec, std::string& ref_store) {
  fm_sparkline_group_t g = FM_STRUCT_INIT(fm_sparkline_group_t);
  for (size_t i = 0; i < 8U; ++i) {
    *GroupColor(g, i) = PullColorSpec(reader, spec, kSparklineColorKeys[i]);
  }
  g.type = reader.I32(spec, "type", FM_SPARKLINE_LINE);
  g.markers = reader.Bool(spec, "markers", false) ? 1 : 0;
  g.high = reader.Bool(spec, "high", false) ? 1 : 0;
  g.low = reader.Bool(spec, "low", false) ? 1 : 0;
  g.first = reader.Bool(spec, "first", false) ? 1 : 0;
  g.last = reader.Bool(spec, "last", false) ? 1 : 0;
  g.negative = reader.Bool(spec, "negative", false) ? 1 : 0;
  g.display_x_axis = reader.Bool(spec, "displayXAxis", false) ? 1 : 0;
  g.display_hidden = reader.Bool(spec, "displayHidden", false) ? 1 : 0;
  g.min_axis_type = reader.I32(spec, "minAxisType", FM_SPARKLINE_AXIS_INDIVIDUAL);
  g.max_axis_type = reader.I32(spec, "maxAxisType", FM_SPARKLINE_AXIS_INDIVIDUAL);
  bool present = false;
  g.manual_min = reader.Double(spec, "manualMin", 0.0, &present);
  g.has_manual_min = present ? 1 : 0;
  g.manual_max = reader.Double(spec, "manualMax", 0.0, &present);
  g.has_manual_max = present ? 1 : 0;
  g.date_axis = reader.Bool(spec, "dateAxis", false) ? 1 : 0;
  if (reader.String(spec, "dateAxisRef", &ref_store)) {
    g.date_axis_ref = ref_store.c_str();
  }
  g.line_weight = reader.Double(spec, "lineWeight", 0.75);
  g.empty_cells_as = reader.I32(spec, "emptyCellsAs", FM_SPARKLINE_EMPTY_GAP);
  return g;
}

// Reads `[{ formula, location }]`; `texts` keeps the strings `items` points into.
bool PullSparklines(CheckedSpecReader& reader, const Napi::Array& array, std::vector<std::string>& texts,
                    std::vector<fm_sparkline_t>& items) {
  const uint32_t count = array.Length();
  texts.reserve(2U * count);
  for (uint32_t i = 0; i < count; ++i) {
    Napi::Value element;
    Napi::Object item;
    std::string formula;
    std::string location;
    if (!reader.ArrayElement(array, i, &element) || !element.IsObject()) {
      return false;
    }
    item = element.As<Napi::Object>();
    reader.String(item, "formula", &formula);
    reader.String(item, "location", &location);
    texts.push_back(std::move(formula));
    texts.push_back(std::move(location));
  }
  for (size_t i = 0; i < texts.size(); i += 2U) {
    items.push_back(fm_sparkline_t{texts[i].c_str(), texts[i + 1U].c_str()});
  }
  return reader.ok();
}

// The JS `Shape` record of `s`.
Napi::Object ShapeToJs(Napi::Env env, const fm_shape_t& s) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("objectId", JsNumber(env, s.object_id));
  item.Set("kind", JsNumber(env, s.kind));
  item.Set("anchorKind", JsNumber(env, s.anchor_kind));
  item.Set("editAs", JsNumber(env, s.edit_as));
  item.Set("fromRow", JsNumber(env, s.from_row));
  item.Set("fromCol", JsNumber(env, s.from_col));
  item.Set("fromRowOffsetEmu", JsNumber(env, static_cast<double>(s.from_row_offset_emu)));
  item.Set("fromColOffsetEmu", JsNumber(env, static_cast<double>(s.from_col_offset_emu)));
  item.Set("toRow", JsNumber(env, s.to_row));
  item.Set("toCol", JsNumber(env, s.to_col));
  item.Set("toRowOffsetEmu", JsNumber(env, static_cast<double>(s.to_row_offset_emu)));
  item.Set("toColOffsetEmu", JsNumber(env, static_cast<double>(s.to_col_offset_emu)));
  item.Set("widthEmu", JsNumber(env, static_cast<double>(s.width_emu)));
  item.Set("heightEmu", JsNumber(env, static_cast<double>(s.height_emu)));
  item.Set("name", JsString(env, s.name));
  item.Set("presetGeometry", JsString(env, s.preset_geometry));
  item.Set("text", JsString(env, s.text));
  item.Set("textBox", Napi::Boolean::New(env, s.text_box != 0));
  item.Set("fillKind", JsIndexOrNull(env, s.fill_kind));
  item.Set("fillColor", ColorSpecOrNull(env, s.fill_color));
  item.Set("lineKind", JsIndexOrNull(env, s.line_kind));
  item.Set("lineColor", ColorSpecOrNull(env, s.line_color));
  item.Set("lineWidthEmu", JsNumber(env, static_cast<double>(s.line_width_emu)));
  item.Set("rotation", JsNumber(env, s.rotation));
  item.Set("hidden", Napi::Boolean::New(env, s.hidden != 0));
  return item;
}

}  // namespace

Napi::Value Workbook::ListShapes(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Array arr = Napi::Array::New(env);
  const uint32_t sheet = ArgU32(info, 0);
  const fm_status_t rc = CollectList<fm_shape_t>(
      arr, [&](size_t* n) { return fm_sheet_get_shape_count(handle_, sheet, n); },
      [&](size_t i, fm_shape_t* s) { return fm_sheet_get_shape_at(handle_, sheet, i, s); },
      [&](const fm_shape_t& s) { return ShapeToJs(env, s); });
  return ValueOrThrow(env, rc, arr);
}

Napi::Value Workbook::GetShape(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_shape_t s = FM_STRUCT_INIT(fm_shape_t);
  int32_t present = 0;
  const fm_status_t rc = fm_sheet_get_shape(handle_, ArgU32(info, 0), ArgU32(info, 1), &present, &s);
  return OptionalOrNull(env, rc, present, rc == 0 && present != 0 ? Napi::Value(ShapeToJs(env, s)) : env.Null());
}

Napi::Value Workbook::InsertShape(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  ShapeText store;
  const fm_shape_t shape = PullShape(reader, ArgObject(info, 1), store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_add_shape(handle_, ArgU32(info, 0), &shape, &object_id);
  return ValueOrThrow(env, rc, JsNumber(env, object_id));
}

Napi::Value Workbook::UpdateShape(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  ShapeText store;
  const fm_shape_t shape = PullShape(reader, ArgObject(info, 2), store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_shape(handle_, ArgU32(info, 0), ArgU32(info, 1), &shape), env.Undefined());
}

Napi::Value Workbook::RemoveShape(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_shape);
}

Napi::Value Workbook::GetSparklineGroupCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_sheet_get_sparkline_group_count(handle_, ArgU32(info, 0), &count);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(count)));
}

Napi::Value Workbook::GetSparklineGroup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = ArgU32(info, 0);
  const std::size_t index = ArgU32(info, 1);
  fm_sparkline_group_t g = FM_STRUCT_INIT(fm_sparkline_group_t);
  fm_status_t rc = fm_sheet_get_sparkline_group(handle_, sheet, index, &g);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object item = Napi::Object::New(env);
  for (size_t c = 0; c < 8U; ++c) {
    item.Set(kSparklineColorKeys[c], ColorSpecToJs(env, *GroupColor(g, c)));
  }
  item.Set("type", JsNumber(env, g.type));
  item.Set("markers", Napi::Boolean::New(env, g.markers != 0));
  item.Set("high", Napi::Boolean::New(env, g.high != 0));
  item.Set("low", Napi::Boolean::New(env, g.low != 0));
  item.Set("first", Napi::Boolean::New(env, g.first != 0));
  item.Set("last", Napi::Boolean::New(env, g.last != 0));
  item.Set("negative", Napi::Boolean::New(env, g.negative != 0));
  item.Set("displayXAxis", Napi::Boolean::New(env, g.display_x_axis != 0));
  item.Set("displayHidden", Napi::Boolean::New(env, g.display_hidden != 0));
  item.Set("minAxisType", JsNumber(env, g.min_axis_type));
  item.Set("maxAxisType", JsNumber(env, g.max_axis_type));
  item.Set("manualMin", g.has_manual_min != 0 ? JsNumber(env, g.manual_min) : env.Null());
  item.Set("manualMax", g.has_manual_max != 0 ? JsNumber(env, g.manual_max) : env.Null());
  item.Set("dateAxis", Napi::Boolean::New(env, g.date_axis != 0));
  item.Set("dateAxisRef", JsStringOrNull(env, g.date_axis_ref));
  item.Set("lineWeight", JsNumber(env, g.line_weight));
  item.Set("emptyCellsAs", JsNumber(env, g.empty_cells_as));
  Napi::Array sparklines = Napi::Array::New(env);
  for (size_t k = 0; rc == 0 && k < g.sparkline_count; ++k) {
    fm_sparkline_t sp{};
    rc = fm_sheet_get_sparkline(handle_, sheet, index, k, &sp);
    if (rc == 0) {
      Napi::Object entry = Napi::Object::New(env);
      entry.Set("formula", JsString(env, sp.formula));
      entry.Set("location", JsString(env, sp.location));
      sparklines.Set(static_cast<uint32_t>(k), entry);
    }
  }
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  item.Set("sparklines", sparklines);
  return item;
}

Napi::Value Workbook::AddSparklineGroup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = ArgObject(info, 1);
  std::string ref_store;
  std::vector<std::string> texts;
  std::vector<fm_sparkline_t> items;
  const fm_sparkline_group_t group = PullGroup(reader, spec, ref_store);
  Napi::Array array;
  if (reader.Array(spec, "sparklines", &array)) {
    PullSparklines(reader, array, texts, items);
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  // No array, or an empty one, reaches the core as zero sparklines and is rejected there.
  fm_sparkline_t none{};
  std::size_t out = 0;
  const fm_status_t rc = fm_sheet_add_sparkline_group(handle_, ArgU32(info, 0), &group,
                                                      items.empty() ? &none : items.data(), items.size(), &out);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(out)));
}

Napi::Value Workbook::UpdateSparklineGroup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = ArgObject(info, 2);
  std::string ref_store;
  std::vector<std::string> texts;
  std::vector<fm_sparkline_t> items;
  const fm_sparkline_group_t group = PullGroup(reader, spec, ref_store);
  Napi::Array array;
  const bool replace = reader.Array(spec, "sparklines", &array);
  if (replace) {
    PullSparklines(reader, array, texts, items);
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  // Without `sparklines` the group keeps its own; an empty array reaches the core, which rejects it.
  fm_sparkline_t none{};
  const fm_sparkline_t* replacement = !replace ? nullptr : items.empty() ? &none : items.data();
  return ValueOrThrow(
      env,
      fm_sheet_update_sparkline_group(handle_, ArgU32(info, 0), ArgU32(info, 1), &group, replacement, items.size()),
      env.Undefined());
}

Napi::Value Workbook::RemoveSparklineGroup(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_sparkline_group);
}

}  // namespace formulon_node
