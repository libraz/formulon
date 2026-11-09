// Drawing-shape and sparkline-group bindings: read, insert and update the
// shapes of a sheet's drawing, and the sparkline groups of its worksheet
// extension list.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "node_addon/parts/chart_internal.h"
#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

constexpr const char* kSparklineColorKeys[] = {"colorSeries", "colorNegative", "colorAxis", "colorMarkers",
                                               "colorFirst",  "colorLast",     "colorHigh", "colorLow"};

fm_color_t* GroupColor(fm_sparkline_group_t& g, size_t i) {
  fm_color_t* const colors[] = {&g.color_series, &g.color_negative, &g.color_axis, &g.color_markers,
                                &g.color_first,  &g.color_last,     &g.color_high, &g.color_low};
  return colors[i];
}

struct ShapeText {
  std::string name;
  std::string description;
  std::string preset_geometry;
  std::string text;
};

// Reads a shape to add: the common placement keys plus the shape keys; an omitted key takes the default, and an
// omitted or `null` `fill` / `line` is the shape style's paint.
fm_shape_spec_t PullShapeSpec(CheckedSpecReader& reader, const Napi::Object& spec, ShapeText& store) {
  fm_shape_spec_t s = FM_STRUCT_INIT(fm_shape_spec_t);
  s.placement = PullPlacement(reader, spec, store.name, store.description);
  s.fill = ReadFillSpec(reader, spec);
  s.line = ReadLineSpec(reader, spec);
  s.rotation_deg = reader.Double(spec, "rotationDeg", 0.0);
  s.text_box = reader.Bool(spec, "textBox", false) ? 1 : 0;
  s.preset_geometry =
      reader.String(spec, "presetGeometry", &store.preset_geometry) ? store.preset_geometry.c_str() : nullptr;
  s.text = reader.String(spec, "text", &store.text) ? store.text.c_str() : nullptr;
  return s;
}

// Reads the shape fields `spec` carries into a patch: an omitted key keeps, and `null` on `fill` / `line` clears
// it to the shape style's paint.
fm_shape_patch_t PullShapePatch(CheckedSpecReader& reader, const Napi::Object& spec, ShapeText& store) {
  fm_shape_patch_t s = FM_STRUCT_INIT(fm_shape_patch_t);
  ReadFillPatch(reader, spec, s.fill);
  ReadLinePatch(reader, spec, s.line);
  s.rotation_deg_op = PatchDouble(reader, spec, "rotationDeg", &s.rotation_deg);
  s.text_box_op = PatchBool(reader, spec, "textBox", &s.text_box);
  s.preset_geometry_op = PatchString(reader, spec, "presetGeometry", false, &store.preset_geometry, &s.preset_geometry);
  s.text_op = PatchString(reader, spec, "text", false, &store.text, &s.text);
  return s;
}

// Reads a group spec; absent keys take the defaults of a new line group.
fm_sparkline_group_t PullGroup(CheckedSpecReader& reader, const Napi::Object& spec, std::string& ref_store) {
  fm_sparkline_group_t g = FM_STRUCT_INIT(fm_sparkline_group_t);
  for (size_t i = 0; i < 8U; ++i) {
    *GroupColor(g, i) = PullColorSpec(reader, spec, kSparklineColorKeys[i]);
  }
  g.type = reader.I32(spec, "type", FM_SPARKLINE_TYPE_LINE);
  g.markers = reader.Bool(spec, "markers", false) ? 1 : 0;
  g.high = reader.Bool(spec, "high", false) ? 1 : 0;
  g.low = reader.Bool(spec, "low", false) ? 1 : 0;
  g.first = reader.Bool(spec, "first", false) ? 1 : 0;
  g.last = reader.Bool(spec, "last", false) ? 1 : 0;
  g.negative = reader.Bool(spec, "negative", false) ? 1 : 0;
  g.display_x_axis = reader.Bool(spec, "displayXAxis", false) ? 1 : 0;
  g.display_hidden = reader.Bool(spec, "displayHidden", false) ? 1 : 0;
  g.min_axis_type = reader.I32(spec, "minAxisType", FM_SPARKLINE_AXIS_TYPE_INDIVIDUAL);
  g.max_axis_type = reader.I32(spec, "maxAxisType", FM_SPARKLINE_AXIS_TYPE_INDIVIDUAL);
  bool present = false;
  g.manual_min = reader.Double(spec, "manualMin", 0.0, &present);
  g.has_manual_min = present ? 1 : 0;
  g.manual_max = reader.Double(spec, "manualMax", 0.0, &present);
  g.has_manual_max = present ? 1 : 0;
  g.date_axis = reader.Bool(spec, "dateAxis", false) ? 1 : 0;
  if (reader.String(spec, "dateAxisRef", &ref_store)) {
    g.date_axis_ref = ref_store.c_str();
  }
  g.line_width_pt = reader.Double(spec, "lineWidthPt", 0.75);
  g.empty_cells_as = reader.I32(spec, "emptyCellsAs", FM_BLANKS_AS_GAP);
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

// The JS `ShapeProperties` record of `s`: the shape-specific fields only; a `null` `fill` / `line` is the shape
// style's paint.
Napi::Object ShapeToJs(Napi::Env env, const fm_shape_info_t& s) {
  Napi::Object item = Napi::Object::New(env);
  item.Set("objectId", JsNumber(env, s.object_id));
  item.Set("kind", JsNumber(env, s.kind));
  item.Set("presetGeometry", JsString(env, s.preset_geometry));
  item.Set("text", JsString(env, s.text));
  item.Set("textBox", Napi::Boolean::New(env, s.text_box != 0));
  item.Set("fill", FillOrNull(env, s.fill));
  item.Set("line", LineOrNull(env, s.line));
  item.Set("rotationDeg", JsNumber(env, s.rotation_deg));
  return item;
}

}  // namespace

Napi::Value Workbook::GetShape(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_shape_info_t s = FM_STRUCT_INIT(fm_shape_info_t);
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
  const fm_shape_spec_t shape = PullShapeSpec(reader, ArgObject(info, 1), store);
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
  // One patch object carries both halves; the C call checks both before it writes either.
  CheckedSpecReader reader(env);
  const Napi::Object spec = ArgObject(info, 2);
  ShapeText store;
  const fm_drawing_patch_t common = PullDrawingPatch(reader, spec, store.name, store.description);
  const fm_shape_patch_t shape = PullShapePatch(reader, spec, store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_shape(handle_, ArgU32(info, 0), ArgU32(info, 1), &common, &shape),
                      env.Undefined());
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
  const Napi::Value key = info[1];
  std::size_t index = 0;
  if (key.IsString()) {
    // A string key is a group id: absent -> null, else its position (the sparklines are read by position).
    const std::string id = ArgString(info, 1);
    fm_sparkline_group_t found = FM_STRUCT_INIT(fm_sparkline_group_t);
    int32_t present = 0;
    const fm_status_t found_rc = fm_sheet_get_sparkline_group(handle_, sheet, id.c_str(), &present, &found);
    if (found_rc != 0 || present == 0) {
      return OptionalOrNull(env, found_rc, present, env.Null());
    }
    size_t count = 0;
    fm_status_t count_rc = fm_sheet_get_sparkline_group_count(handle_, sheet, &count);
    for (; count_rc == 0 && index < count; ++index) {
      fm_sparkline_group_t probe = FM_STRUCT_INIT(fm_sparkline_group_t);
      count_rc = fm_sheet_get_sparkline_group_at(handle_, sheet, index, &probe);
      if (count_rc == 0 && id == probe.id) {
        break;
      }
    }
    if (count_rc != 0) {
      return ThrowStatus(env, count_rc);
    }
  } else {
    CheckedSpecReader key_reader(env);
    if (!key.IsNumber()) {
      key_reader.FailType("sparklineGroup", "a number or a string");
      return env.Undefined();
    }
    const uint32_t position = key_reader.U32(key, "sparklineGroup", 0U);
    if (!key_reader.ok()) {
      return env.Undefined();
    }
    index = position;
  }
  fm_sparkline_group_t g = FM_STRUCT_INIT(fm_sparkline_group_t);
  fm_status_t rc = fm_sheet_get_sparkline_group_at(handle_, sheet, index, &g);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object item = Napi::Object::New(env);
  item.Set("id", JsString(env, g.id));
  for (size_t c = 0; c < 8U; ++c) {
    item.Set(kSparklineColorKeys[c], ColorSpecOrNull(env, *GroupColor(g, c)));
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
  item.Set("lineWidthPt", JsNumber(env, g.line_width_pt));
  item.Set("emptyCellsAs", JsNumber(env, g.empty_cells_as));
  Napi::Array sparklines = Napi::Array::New(env);
  for (size_t k = 0; rc == 0 && k < g.sparkline_count; ++k) {
    fm_sparkline_t sp{};
    rc = fm_sheet_get_sparkline_at(handle_, sheet, index, k, &sp);
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
  const char* id = nullptr;
  const fm_status_t rc = fm_sheet_add_sparkline_group(handle_, ArgU32(info, 0), &group,
                                                      items.empty() ? &none : items.data(), items.size(), &id);
  return ValueOrThrow(env, rc, JsString(env, id));
}

namespace {

// Reads a group patch: an omitted key keeps, and `null` clears `manualMin`, `manualMax` and `dateAxisRef`. A
// `sparklines` array replaces the group's sparklines.
fm_sparkline_group_patch_t PullGroupPatch(CheckedSpecReader& reader, const Napi::Object& spec, std::string& ref_store,
                                          std::vector<std::string>& texts, std::vector<fm_sparkline_t>& items) {
  fm_sparkline_group_patch_t g = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  fm_color_t* const colors[] = {&g.color_series, &g.color_negative, &g.color_axis, &g.color_markers,
                                &g.color_first,  &g.color_last,     &g.color_high, &g.color_low};
  fm_patch_op_t* const color_ops[] = {&g.color_series_op, &g.color_negative_op, &g.color_axis_op, &g.color_markers_op,
                                      &g.color_first_op,  &g.color_last_op,     &g.color_high_op, &g.color_low_op};
  for (size_t i = 0; i < 8U; ++i) {
    *color_ops[i] = PatchColor(reader, spec, kSparklineColorKeys[i], false, colors[i]);
  }
  g.type_op = PatchI32(reader, spec, "type", &g.type);
  g.markers_op = PatchBool(reader, spec, "markers", &g.markers);
  g.high_op = PatchBool(reader, spec, "high", &g.high);
  g.low_op = PatchBool(reader, spec, "low", &g.low);
  g.first_op = PatchBool(reader, spec, "first", &g.first);
  g.last_op = PatchBool(reader, spec, "last", &g.last);
  g.negative_op = PatchBool(reader, spec, "negative", &g.negative);
  g.display_x_axis_op = PatchBool(reader, spec, "displayXAxis", &g.display_x_axis);
  g.display_hidden_op = PatchBool(reader, spec, "displayHidden", &g.display_hidden);
  g.min_axis_type_op = PatchI32(reader, spec, "minAxisType", &g.min_axis_type);
  g.max_axis_type_op = PatchI32(reader, spec, "maxAxisType", &g.max_axis_type);
  g.manual_min_op = PatchNullableDouble(reader, spec, "manualMin", &g.manual_min);
  g.manual_max_op = PatchNullableDouble(reader, spec, "manualMax", &g.manual_max);
  g.date_axis_op = PatchBool(reader, spec, "dateAxis", &g.date_axis);
  g.date_axis_ref_op = PatchString(reader, spec, "dateAxisRef", true, &ref_store, &g.date_axis_ref);
  g.line_width_pt_op = PatchDouble(reader, spec, "lineWidthPt", &g.line_width_pt);
  g.empty_cells_as_op = PatchI32(reader, spec, "emptyCellsAs", &g.empty_cells_as);
  Napi::Array array;
  if (reader.Array(spec, "sparklines", &array)) {
    PullSparklines(reader, array, texts, items);
    g.items_op = FM_PATCH_OP_SET;
    g.items = items.empty() ? nullptr : items.data();
    g.items_count = static_cast<uint32_t>(items.size());
  }
  return g;
}

}  // namespace

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
  const fm_sparkline_group_patch_t patch = PullGroupPatch(reader, spec, ref_store, texts, items);
  if (!reader.ok()) {
    return env.Undefined();
  }
  const std::string id = ArgString(info, 1);
  return ValueOrThrow(env, fm_sheet_update_sparkline_group(handle_, ArgU32(info, 0), id.c_str(), &patch),
                      env.Undefined());
}

Napi::Value Workbook::RemoveSparklineGroup(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string id = ArgString(info, 1);
    return fm_sheet_remove_sparkline_group(handle, ArgU32(info, 0), id.c_str());
  });
}

}  // namespace formulon_node
