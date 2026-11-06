// Chart patch bindings: chart creation and the typed edits. Patch objects map
// `undefined` to keep, `null` to clear and a value to set.

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/chart_internal.h"
#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace detail {

Field Probe(CheckedSpecReader& reader, const Napi::Object& owner, const char* key) {
  Field f;
  if (!reader.ok()) {
    return f;
  }
  const Napi::Value v = owner.Get(key);
  if (!reader.ok() || v.IsUndefined()) {
    return f;
  }
  if (v.IsNull()) {
    f.op = FM_CHART_PATCH_CLEAR;
    return f;
  }
  f.op = FM_CHART_PATCH_SET;
  f.value = v;
  return f;
}

/// Reads the patch object argument `info[idx]`; absent reads as `{}`. False after
/// throwing when a value other than a plain object was supplied.
bool PatchObjectArg(const Napi::CallbackInfo& info, size_t idx, const char* api, Napi::Object* out) {
  if (info[idx].IsUndefined()) {
    *out = Napi::Object::New(info.Env());
    return true;
  }
  if (info[idx].IsArray() || !ObjectArg(info, idx, out)) {
    Napi::TypeError::New(info.Env(), std::string(api) + " expects an object argument").ThrowAsJavaScriptException();
    return false;
  }
  return true;
}

}  // namespace detail

namespace {

/// Probes `key` and, when it is set, stores `read(f.value)` into `out`.
template <typename T, typename Read>
void ReadOp(CheckedSpecReader& r, const Napi::Object& owner, const char* key, int32_t& op, T& out, Read&& read) {
  const Field f = Probe(r, owner, key);
  op = f.op;
  if (f.op == FM_CHART_PATCH_SET) {
    out = read(f.value);
  }
}

void ReadU32Op(CheckedSpecReader& r, const Napi::Object& owner, const char* key, int32_t& op, uint32_t& out) {
  ReadOp(r, owner, key, op, out, [&](const Napi::Value& v) { return r.U32(v, key, 0U); });
}

void ReadI32Op(CheckedSpecReader& r, const Napi::Object& owner, const char* key, int32_t& op, int32_t& out) {
  ReadOp(r, owner, key, op, out, [&](const Napi::Value& v) { return r.I32(v, key, 0); });
}

void ReadDoubleOp(CheckedSpecReader& r, const Napi::Object& owner, const char* key, int32_t& op, double& out) {
  ReadOp(r, owner, key, op, out, [&](const Napi::Value& v) { return r.Double(v, key, 0.0); });
}

void ReadBoolOp(CheckedSpecReader& r, const Napi::Object& owner, const char* key, int32_t& op, int32_t& out) {
  ReadOp(r, owner, key, op, out, [&](const Napi::Value& v) { return r.Bool(v, key, false) ? 1 : 0; });
}

/// Keeps the std::string bytes behind the `const char*` fields of a patch alive.
using Arena = std::deque<std::string>;

const char* Intern(Arena& arena, std::string text) {
  arena.push_back(std::move(text));
  return arena.back().c_str();
}

void ReadStringOp(CheckedSpecReader& r, const Napi::Object& owner, const char* key, Arena& arena, int32_t& op,
                  const char*& out) {
  const Field f = Probe(r, owner, key);
  op = f.op;
  if (f.op == FM_CHART_PATCH_SET) {
    std::string text;
    r.String(f.value, key, &text);
    out = Intern(arena, std::move(text));
  }
}

/// A title given as its text, or in the read-back shape `{ kind, text, formula }`. An out-of-domain (absent)
/// kind is left for the C ABI to reject.
fm_chart_title_t ReadTitle(CheckedSpecReader& r, const Napi::Value& v, const char* key, Arena& arena) {
  fm_chart_title_t t{};
  std::string text;
  if (v.IsObject()) {
    const Napi::Object spec = v.As<Napi::Object>();
    std::string formula;
    t.kind = r.I32(spec, "kind", -1);
    r.String(spec, "text", &text);
    r.String(spec, "formula", &formula);
    t.formula = Intern(arena, std::move(formula));
  } else {
    t.kind = FM_CHART_TITLE_TEXT;
    r.String(v, key, &text);
  }
  t.text = Intern(arena, std::move(text));
  return t;
}

/// A series reference: absent keeps (`NULL`), `null` or "" removes, a string replaces.
void ReadRef(CheckedSpecReader& r, const Napi::Object& owner, const char* key, Arena& arena, const char*& out) {
  const Field f = Probe(r, owner, key);
  if (f.op == FM_CHART_PATCH_CLEAR) {
    out = "";
  } else if (f.op == FM_CHART_PATCH_SET) {
    std::string text;
    r.String(f.value, key, &text);
    out = Intern(arena, std::move(text));
  }
}

bool ReadChild(CheckedSpecReader& r, const Napi::Object& owner, const char* key, Napi::Object* out) {
  if (!r.ok()) {
    return false;
  }
  const Napi::Value v = owner.Get(key);
  return r.ok() && r.Object(v, key, out);
}

void ReadFillPatch(CheckedSpecReader& r, const Napi::Object& owner, fm_chart_fill_patch_t& out) {
  Napi::Object o;
  if (!ReadChild(r, owner, "fill", &o)) {
    return;
  }
  ReadU32Op(r, o, "color", out.color_op, out.color);
  ReadDoubleOp(r, o, "transparency", out.transparency_op, out.transparency_pct);
}

void ReadLinePatch(CheckedSpecReader& r, const Napi::Object& owner, fm_chart_line_patch_t& out) {
  Napi::Object o;
  if (!ReadChild(r, owner, "line", &o)) {
    return;
  }
  ReadU32Op(r, o, "color", out.color_op, out.color);
  ReadDoubleOp(r, o, "transparency", out.transparency_op, out.transparency_pct);
  ReadDoubleOp(r, o, "widthPt", out.width_op, out.width_pt);
  ReadI32Op(r, o, "dash", out.dash_op, out.dash);
  ReadBoolOp(r, o, "none", out.none_op, out.none);
}

void ReadMarkerPatch(CheckedSpecReader& r, const Napi::Object& owner, int32_t& op, fm_chart_marker_patch_t& out) {
  const Field f = Probe(r, owner, "marker");
  op = f.op;
  Napi::Object o;
  if (f.op != FM_CHART_PATCH_SET || !r.Object(f.value, "marker", &o)) {
    return;
  }
  ReadI32Op(r, o, "symbol", out.symbol_op, out.symbol);
  ReadDoubleOp(r, o, "size", out.size_op, out.size);
  ReadFillPatch(r, o, out.fill);
  ReadLinePatch(r, o, out.line);
}

void ReadLabelsPatch(CheckedSpecReader& r, const Napi::Object& owner, Arena& arena, int32_t& op,
                     fm_chart_labels_patch_t& out) {
  const Field f = Probe(r, owner, "labels");
  op = f.op;
  Napi::Object o;
  if (f.op != FM_CHART_PATCH_SET || !r.Object(f.value, "labels", &o)) {
    return;
  }
  ReadBoolOp(r, o, "showValue", out.show_value_op, out.show_value);
  ReadBoolOp(r, o, "showCategory", out.show_category_op, out.show_category);
  ReadBoolOp(r, o, "showSeries", out.show_series_op, out.show_series);
  ReadBoolOp(r, o, "showPercent", out.show_percent_op, out.show_percent);
  ReadBoolOp(r, o, "showLegendKey", out.show_legend_key_op, out.show_legend_key);
  ReadStringOp(r, o, "separator", arena, out.separator_op, out.separator);
  ReadI32Op(r, o, "position", out.position_op, out.position);
  ReadStringOp(r, o, "numberFormat", arena, out.number_format_op, out.number_format);
  ReadU32Op(r, o, "fontColor", out.font_color_op, out.font_color);
  ReadDoubleOp(r, o, "fontSizePt", out.font_size_op, out.font_size_pt);
}

void ReadSeriesPatch(CheckedSpecReader& r, const Napi::Object& spec, Arena& arena, fm_chart_series_patch_t& out) {
  ReadRef(r, spec, "name", arena, out.name);
  ReadRef(r, spec, "categories", arena, out.categories);
  ReadRef(r, spec, "values", arena, out.values);
  ReadRef(r, spec, "x", arena, out.x);
  ReadRef(r, spec, "y", arena, out.y);
  ReadRef(r, spec, "size", arena, out.size);
  ReadFillPatch(r, spec, out.fill);
  ReadLinePatch(r, spec, out.line);
  ReadMarkerPatch(r, spec, out.marker_op, out.marker);
  ReadBoolOp(r, spec, "smooth", out.smooth_op, out.smooth);
  ReadLabelsPatch(r, spec, arena, out.labels_op, out.labels);
}

void ReadPointPatch(CheckedSpecReader& r, const Napi::Object& spec, fm_chart_point_patch_t& out) {
  ReadFillPatch(r, spec, out.fill);
  ReadLinePatch(r, spec, out.line);
  ReadMarkerPatch(r, spec, out.marker_op, out.marker);
}

/// A `boolean | null | undefined` header flag as an `fm_chart_header_mode_t`.
int32_t HeaderMode(CheckedSpecReader& r, const Napi::Value& v, const char* key) {
  if (v.IsUndefined() || v.IsNull()) {
    return FM_CHART_HEADER_AUTO;
  }
  return r.Bool(v, key, false) ? FM_CHART_HEADER_YES : FM_CHART_HEADER_NO;
}

int32_t HeaderModeArg(CheckedSpecReader& r, const Napi::CallbackInfo& info, size_t idx, const char* key) {
  return info.Length() > idx ? HeaderMode(r, info[idx], key) : FM_CHART_HEADER_AUTO;
}

/// Reads a `number[]` into `out`; absent reads as empty. False after throwing otherwise.
bool ReadIndexList(CheckedSpecReader& r, const Napi::CallbackInfo& info, size_t idx, const char* key,
                   std::vector<uint32_t>* out) {
  if (info.Length() <= idx || info[idx].IsUndefined() || info[idx].IsNull()) {
    return true;
  }
  Napi::Array list;
  if (!r.Array(info[idx], key, &list)) {
    return r.ok();
  }
  const uint32_t n = list.Length();
  for (uint32_t i = 0; i < n && r.ok(); ++i) {
    Napi::Value v;
    if (r.ArrayElement(list, i, &v)) {
      out->push_back(r.U32(v, key, 0U));
    } else if (r.ok()) {
      Napi::TypeError::New(info.Env(), std::string(key) + " must contain numbers").ThrowAsJavaScriptException();
      return false;
    }
  }
  return r.ok();
}

}  // namespace

// ---- Creation and edits ---------------------------------------------

Napi::Value Workbook::InsertChart(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = ArgObject(info, 1);
  std::string source;
  std::string name;
  std::string descr;
  std::string title;
  reader.String(spec, "source", &source);
  reader.String(spec, "name", &name);
  reader.String(spec, "description", &descr);
  bool has_title = false;
  {
    const Field f = Probe(reader, spec, "title");
    if (f.op == FM_CHART_PATCH_SET) {
      reader.String(f.value, "title", &title);
      has_title = true;
    }
  }
  fm_chart_insert_t opts{};
  opts.kind = reader.I32(spec, "kind", FM_CHART_INSERT_COLUMN);
  opts.grouping = reader.I32(spec, "grouping", FM_CHART_GROUPING_CLUSTERED);
  opts.source = source.c_str();
  opts.direction = reader.I32(spec, "direction", FM_CHART_DIRECTION_AUTO);
  opts.header_row = HeaderMode(reader, reader.ok() ? spec.Get("headerRow") : env.Undefined(), "headerRow");
  opts.header_column = HeaderMode(reader, reader.ok() ? spec.Get("headerColumn") : env.Undefined(), "headerColumn");
  opts.title = has_title ? title.c_str() : nullptr;
  opts.legend = reader.I32(spec, "legend", FM_CHART_LEGEND_BOTTOM);
  opts.placement.name = name.c_str();
  opts.placement.descr = descr.c_str();
  Napi::Object placement;
  if (ReadChild(reader, spec, "placement", &placement)) {
    PullPlacement(reader, placement, opts.placement, FM_ANCHOR_KIND_TWO_CELL);
  } else {
    opts.placement.anchor_kind = FM_ANCHOR_KIND_TWO_CELL;
    opts.placement.edit_as = FM_ANCHOR_EDIT_AS_TWO_CELL;
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_add_chart(handle_, ArgU32(info, 0), &opts, &object_id);
  return ValueOrThrow(env, rc, JsNumber(env, object_id));
}

Napi::Value Workbook::SetChartType(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const int32_t grouping = info.Length() > 4 && !info[4].IsUndefined() ? ArgI32(info, 4) : FM_CHART_GROUPING_CLUSTERED;
  return ValueOrThrow(
      env,
      fm_sheet_set_chart_type(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), ArgI32(info, 3), grouping),
      env.Undefined());
}

Napi::Value Workbook::SetChartTitle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // The registration guard requires `title`; a wrong type throws as the WASM
  // reader does.
  const Napi::Value arg = info[2];
  if (!arg.IsNull() && !arg.IsString() && !arg.IsObject()) {
    Napi::TypeError::New(env, "title must be a string, null or an object").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  Arena arena;
  fm_chart_title_t title{};
  title.kind = FM_CHART_TITLE_NONE;
  if (!arg.IsNull()) {
    title = ReadTitle(reader, arg, "title", arena);
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_chart_title(handle_, ArgU32(info, 0), ArgU32(info, 1), &title),
                      env.Undefined());
}

Napi::Value Workbook::SetChartLegend(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  std::vector<uint32_t> hidden;
  const bool has_hidden = info.Length() > 4 && !info[4].IsUndefined() && !info[4].IsNull();
  if (!ReadIndexList(reader, info, 4, "hiddenSeries", &hidden)) {
    return env.Undefined();
  }
  fm_chart_legend_patch_t legend{};
  legend.position = ArgI32(info, 2);
  legend.overlay = ArgBool(info, 3) ? 1 : 0;
  if (has_hidden) {
    legend.hidden_series_op = FM_CHART_PATCH_SET;
    legend.hidden_series = hidden.empty() ? nullptr : hidden.data();
    legend.hidden_series_count = hidden.size();
  }
  return ValueOrThrow(env, fm_sheet_set_chart_legend(handle_, ArgU32(info, 0), ArgU32(info, 1), &legend),
                      env.Undefined());
}

Napi::Value Workbook::SetChartAxis(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Object spec;
  if (!PatchObjectArg(info, 3, "setChartAxis", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  Arena arena;
  fm_chart_axis_patch_t patch{};
  ReadBoolOp(reader, spec, "visible", patch.visible_op, patch.visible);
  const Field title = Probe(reader, spec, "title");
  patch.title_op = title.op;
  if (title.op == FM_CHART_PATCH_SET) {
    patch.title = ReadTitle(reader, title.value, "title", arena);
  }
  ReadDoubleOp(reader, spec, "min", patch.min_op, patch.min);
  ReadDoubleOp(reader, spec, "max", patch.max_op, patch.max);
  ReadDoubleOp(reader, spec, "majorUnit", patch.major_unit_op, patch.major_unit);
  ReadStringOp(reader, spec, "numberFormat", arena, patch.number_format_op, patch.number_format);
  ReadBoolOp(reader, spec, "majorGridlines", patch.major_gridlines_op, patch.major_gridlines);
  ReadBoolOp(reader, spec, "reversed", patch.reversed_op, patch.reversed);
  ReadDoubleOp(reader, spec, "logBase", patch.log_base_op, patch.log_base);
  ReadI32Op(reader, spec, "crosses", patch.crosses_op, patch.crosses);
  patch.crosses_at = reader.Double(spec, "crossesAt", 0.0);
  ReadStringOp(reader, spec, "tickLabelPosition", arena, patch.tick_label_pos_op, patch.tick_label_pos);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_chart_axis(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &patch),
                      env.Undefined());
}

Napi::Value Workbook::SetChartDisplay(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Object spec;
  if (!PatchObjectArg(info, 2, "setChartDisplay", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  fm_chart_display_patch_t patch{};
  ReadI32Op(reader, spec, "blanksAs", patch.blanks_as_op, patch.blanks_as);
  ReadBoolOp(reader, spec, "plotVisibleOnly", patch.plot_vis_only_op, patch.plot_vis_only);
  ReadBoolOp(reader, spec, "naAsBlank", patch.na_as_blank_op, patch.na_as_blank);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_chart_display(handle_, ArgU32(info, 0), ArgU32(info, 1), &patch),
                      env.Undefined());
}

Napi::Value Workbook::SetChartSeries(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Object spec;
  if (!PatchObjectArg(info, 4, "setChartSeries", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  Arena arena;
  fm_chart_series_patch_t patch{};
  ReadSeriesPatch(reader, spec, arena, patch);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(
      env,
      fm_sheet_set_chart_series(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), ArgU32(info, 3), &patch),
      env.Undefined());
}

Napi::Value Workbook::SetChartPoint(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // A null or empty patch removes the point's formatting.
  Napi::Object spec = Napi::Object::New(env);
  if (info.Length() > 5 && !info[5].IsNull() && !PatchObjectArg(info, 5, "setChartPoint", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  fm_chart_point_patch_t patch{};
  ReadPointPatch(reader, spec, patch);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env,
                      fm_sheet_set_chart_point(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2),
                                               ArgU32(info, 3), ArgU32(info, 4), &patch),
                      env.Undefined());
}

Napi::Value Workbook::SetChartPlot(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Object spec;
  if (!PatchObjectArg(info, 3, "setChartPlot", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  fm_chart_plot_patch_t patch{};
  ReadU32Op(reader, spec, "gapWidth", patch.gap_width_op, patch.gap_width);
  ReadI32Op(reader, spec, "overlap", patch.overlap_op, patch.overlap);
  ReadU32Op(reader, spec, "firstSliceAngle", patch.first_slice_angle_op, patch.first_slice_angle);
  ReadBoolOp(reader, spec, "varyColors", patch.vary_colors_op, patch.vary_colors);
  ReadU32Op(reader, spec, "holeSize", patch.hole_size_op, patch.hole_size);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_chart_plot(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &patch),
                      env.Undefined());
}

Napi::Value Workbook::SetChartFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  std::vector<uint32_t> series;
  std::vector<uint32_t> categories;
  if (!ReadIndexList(reader, info, 2, "hiddenSeries", &series) ||
      !ReadIndexList(reader, info, 3, "hiddenCategories", &categories)) {
    return env.Undefined();
  }
  return ValueOrThrow(
      env,
      fm_sheet_set_chart_filter(handle_, ArgU32(info, 0), ArgU32(info, 1), series.empty() ? nullptr : series.data(),
                                series.size(), categories.empty() ? nullptr : categories.data(), categories.size()),
      env.Undefined());
}

Napi::Value Workbook::AddChartSeries(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Object spec;
  if (!PatchObjectArg(info, 3, "addChartSeries", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  Arena arena;
  fm_chart_series_patch_t patch{};
  ReadSeriesPatch(reader, spec, arena, patch);
  if (!reader.ok()) {
    return env.Undefined();
  }
  std::size_t series_index = 0;
  const fm_status_t rc =
      fm_sheet_add_chart_series(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &patch, &series_index);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(series_index)));
}

Napi::Value Workbook::RemoveChartSeries(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  return ValueOrThrow(
      env, fm_sheet_remove_chart_series(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), ArgU32(info, 3)),
      env.Undefined());
}

Napi::Value Workbook::SetChartSource(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  CheckedSpecReader reader(env);
  const int32_t header_row = HeaderModeArg(reader, info, 4, "headerRow");
  const int32_t header_column = HeaderModeArg(reader, info, 5, "headerColumn");
  if (!reader.ok()) {
    return env.Undefined();
  }
  const std::string source = ArgString(info, 2);
  return ValueOrThrow(env,
                      fm_sheet_set_chart_source(handle_, ArgU32(info, 0), ArgU32(info, 1), source.c_str(),
                                                ArgI32(info, 3), header_row, header_column),
                      env.Undefined());
}

}  // namespace formulon_node
