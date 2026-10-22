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

fm_color_spec* GroupColor(fm_sparkline_group& g, size_t i) {
  fm_color_spec* const colors[] = {&g.color_series, &g.color_negative, &g.color_axis, &g.color_markers,
                                   &g.color_first,  &g.color_last,     &g.color_high, &g.color_low};
  return colors[i];
}

Napi::Object ColorToJs(Napi::Env env, const fm_color_spec& spec) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("kind", JsNumber(env, spec.kind));
  out.Set("rgb", JsNumber(env, spec.rgb));
  out.Set("theme", JsNumber(env, spec.theme));
  out.Set("tint", JsNumber(env, spec.tint));
  out.Set("indexed", JsNumber(env, spec.indexed));
  return out;
}

fm_color_spec PullColor(CheckedSpecReader& reader, const Napi::Object& owner, const char* key) {
  fm_color_spec spec{};
  Napi::Object v;
  if (!reader.Object(owner, key, &v)) {
    return spec;
  }
  spec.kind = reader.U8(v, "kind", 0U);
  spec.rgb = reader.U32(v, "rgb", 0U);
  spec.theme = reader.U32(v, "theme", 0U);
  spec.tint = reader.Double(v, "tint", 0.0);
  spec.indexed = reader.U32(v, "indexed", 0U);
  return spec;
}

struct ShapeText {
  std::string name;
  std::string prst;
  std::string text;
};

// Reads the fields `spec` carries; the `fields` mask names exactly those.
fm_shape PullShape(CheckedSpecReader& reader, const Napi::Object& spec, ShapeText& store) {
  fm_shape s{};
  bool present = false;
  if (reader.String(spec, "name", &store.name)) {
    s.fields |= FM_SHAPE_FIELD_NAME;
    s.name = store.name.c_str();
  }
  if (reader.String(spec, "prst", &store.prst)) {
    s.fields |= FM_SHAPE_FIELD_PRST;
    s.prst = store.prst.c_str();
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
    s.from_row_off = reader.I64(spec, "fromRowOff", 0);
    s.from_col_off = reader.I64(spec, "fromColOff", 0);
    s.to_row = reader.U32(spec, "toRow", 0U);
    s.to_col = reader.U32(spec, "toCol", 0U);
    s.to_row_off = reader.I64(spec, "toRowOff", 0);
    s.to_col_off = reader.I64(spec, "toColOff", 0);
    s.cx = reader.I64(spec, "cx", 0);
    s.cy = reader.I64(spec, "cy", 0);
  }
  s.fill_kind = reader.I32(spec, "fillKind", 0, &present);
  if (present) {
    s.fields |= FM_SHAPE_FIELD_FILL;
    s.fill_rgb = reader.U32(spec, "fillRgb", 0U);
  }
  s.line_kind = reader.I32(spec, "lineKind", 0, &present);
  if (present) {
    s.fields |= FM_SHAPE_FIELD_LINE;
    s.line_rgb = reader.U32(spec, "lineRgb", 0U);
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
fm_sparkline_group PullGroup(CheckedSpecReader& reader, const Napi::Object& spec, std::string& ref_store) {
  fm_sparkline_group g{};
  for (size_t i = 0; i < 8U; ++i) {
    *GroupColor(g, i) = PullColor(reader, spec, kSparklineColorKeys[i]);
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
                    std::vector<fm_sparkline>& items) {
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
    items.push_back(fm_sparkline{texts[i].c_str(), texts[i + 1U].c_str()});
  }
  return reader.ok();
}

}  // namespace

Napi::Value Workbook::ListShapes(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const uint32_t sheet = ArgU32(info, 0);
  size_t count = 0;
  fm_status_t rc = fm_sheet_shape_count(handle_, sheet, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_shape s{};
    rc = fm_sheet_shape_at(handle_, sheet, i, &s);
    if (rc != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    item.Set("objectId", JsNumber(env, s.object_id));
    item.Set("kind", JsNumber(env, s.kind));
    item.Set("anchorKind", JsNumber(env, s.anchor_kind));
    item.Set("editAs", JsNumber(env, s.edit_as));
    item.Set("fromRow", JsNumber(env, s.from_row));
    item.Set("fromCol", JsNumber(env, s.from_col));
    item.Set("fromRowOff", JsNumber(env, static_cast<double>(s.from_row_off)));
    item.Set("fromColOff", JsNumber(env, static_cast<double>(s.from_col_off)));
    item.Set("toRow", JsNumber(env, s.to_row));
    item.Set("toCol", JsNumber(env, s.to_col));
    item.Set("toRowOff", JsNumber(env, static_cast<double>(s.to_row_off)));
    item.Set("toColOff", JsNumber(env, static_cast<double>(s.to_col_off)));
    item.Set("cx", JsNumber(env, static_cast<double>(s.cx)));
    item.Set("cy", JsNumber(env, static_cast<double>(s.cy)));
    item.Set("name", JsString(env, s.name));
    item.Set("prst", JsString(env, s.prst));
    item.Set("text", JsString(env, s.text));
    item.Set("textBox", Napi::Boolean::New(env, s.text_box != 0));
    item.Set("fillKind", JsNumber(env, s.fill_kind));
    item.Set("fillRgb", JsNumber(env, s.fill_rgb));
    item.Set("lineKind", JsNumber(env, s.line_kind));
    item.Set("lineRgb", JsNumber(env, s.line_rgb));
    item.Set("lineWidthEmu", JsNumber(env, static_cast<double>(s.line_width_emu)));
    item.Set("rotation", JsNumber(env, s.rotation));
    item.Set("hidden", Napi::Boolean::New(env, s.hidden != 0));
    arr.Set(static_cast<uint32_t>(i), item);
  }
  return FinishListResult(env, arr, rc);
}

Napi::Value Workbook::InsertShape(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("objectId", JsNumber(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  if (info.Length() < 2 || !info[1].IsObject()) {
    out.Set("status", MakeBindingArgumentError(env, "insertShape expects (sheet:number, spec:object)"));
    return out;
  }
  CheckedSpecReader reader(env);
  ShapeText store;
  const fm_shape shape = PullShape(reader, info[1].As<Napi::Object>(), store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_insert_shape(handle_, ArgU32(info, 0), &shape, &object_id);
  out.Set("objectId", JsNumber(env, rc == 0 ? object_id : 0));
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::UpdateShape(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeBindingArgumentError(env, "updateShape expects (sheet:number, objectId:number, spec:object)");
  }
  CheckedSpecReader reader(env);
  ShapeText store;
  const fm_shape shape = PullShape(reader, info[2].As<Napi::Object>(), store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_update_shape(handle_, ArgU32(info, 0), ArgU32(info, 1), &shape));
}

Napi::Value Workbook::RemoveShape(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_shape);
}

Napi::Value Workbook::ListSparklineGroups(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Array arr = Napi::Array::New(env);
  if (handle_ == nullptr) {
    return FinishListResult(env, arr, kBindingInvalidHandle);
  }
  const uint32_t sheet = ArgU32(info, 0);
  size_t count = 0;
  fm_status_t rc = fm_sheet_sparkline_group_count(handle_, sheet, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    fm_sparkline_group g{};
    rc = fm_sheet_sparkline_group_at(handle_, sheet, i, &g);
    if (rc != 0) {
      break;
    }
    Napi::Object item = Napi::Object::New(env);
    for (size_t c = 0; c < 8U; ++c) {
      item.Set(kSparklineColorKeys[c], ColorToJs(env, *GroupColor(g, c)));
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
    item.Set("dateAxisRef", JsString(env, g.date_axis_ref));
    item.Set("lineWeight", JsNumber(env, g.line_weight));
    item.Set("emptyCellsAs", JsNumber(env, g.empty_cells_as));
    Napi::Array sparklines = Napi::Array::New(env);
    const size_t n = g.sparkline_count;
    for (size_t k = 0; rc == 0 && k < n; ++k) {
      fm_sparkline sp{};
      rc = fm_sheet_sparkline_at(handle_, sheet, i, k, &sp);
      if (rc == 0) {
        Napi::Object entry = Napi::Object::New(env);
        entry.Set("formula", JsString(env, sp.formula));
        entry.Set("location", JsString(env, sp.location));
        sparklines.Set(static_cast<uint32_t>(k), entry);
      }
    }
    item.Set("sparklines", sparklines);
    arr.Set(static_cast<uint32_t>(i), item);
  }
  return FinishListResult(env, arr, rc);
}

Napi::Value Workbook::AddSparklineGroup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 2 || !info[1].IsObject()) {
    return MakeBindingArgumentError(env, "addSparklineGroup expects (sheet:number, spec:object)");
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = info[1].As<Napi::Object>();
  std::string ref_store;
  std::vector<std::string> texts;
  std::vector<fm_sparkline> items;
  const fm_sparkline_group group = PullGroup(reader, spec, ref_store);
  Napi::Array array;
  if (reader.Array(spec, "sparklines", &array)) {
    PullSparklines(reader, array, texts, items);
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  // No array, or an empty one, reaches the core as zero sparklines and is rejected there.
  fm_sparkline none{};
  return MakeStatus(env, fm_sheet_add_sparkline_group(handle_, ArgU32(info, 0), &group,
                                                      items.empty() ? &none : items.data(), items.size()));
}

Napi::Value Workbook::UpdateSparklineGroup(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 3 || !info[2].IsObject()) {
    return MakeBindingArgumentError(env, "updateSparklineGroup expects (sheet:number, index:number, spec:object)");
  }
  CheckedSpecReader reader(env);
  std::string ref_store;
  const fm_sparkline_group group = PullGroup(reader, info[2].As<Napi::Object>(), ref_store);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_update_sparkline_group(handle_, ArgU32(info, 0), ArgU32(info, 1), &group));
}

Napi::Value Workbook::SetSparklineGroupSparklines(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  if (info.Length() < 3 || !info[2].IsArray()) {
    return MakeBindingArgumentError(
        env, "setSparklineGroupSparklines expects (sheet:number, index:number, sparklines:array)");
  }
  CheckedSpecReader reader(env);
  std::vector<std::string> texts;
  std::vector<fm_sparkline> items;
  if (!PullSparklines(reader, info[2].As<Napi::Array>(), texts, items) && !reader.ok()) {
    return env.Undefined();
  }
  fm_sparkline none{};
  return MakeStatus(env, fm_sheet_set_sparkline_group_sparklines(handle_, ArgU32(info, 0), ArgU32(info, 1),
                                                                 items.empty() ? &none : items.data(), items.size()));
}

Napi::Value Workbook::RemoveSparklineGroup(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_remove_sparkline_group);
}

}  // namespace formulon_node
