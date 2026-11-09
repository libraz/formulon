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
    f.op = FM_PATCH_OP_CLEAR;
    return f;
  }
  f.op = FM_PATCH_OP_SET;
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
void ReadOp(CheckedSpecReader& r, const Napi::Object& owner, const char* key, fm_patch_op_t& op, T& out, Read&& read) {
  const Field f = Probe(r, owner, key);
  op = f.op;
  if (f.op == FM_PATCH_OP_SET) {
    out = read(f.value);
  }
}

void ReadU32Op(CheckedSpecReader& r, const Napi::Object& owner, const char* key, fm_patch_op_t& op, uint32_t& out) {
  ReadOp(r, owner, key, op, out, [&](const Napi::Value& v) { return r.U32(v, key, 0U); });
}

void ReadColorOp(CheckedSpecReader& r, const Napi::Object& owner, const char* key, fm_patch_op_t& op, fm_color_t& out) {
  ReadOp(r, owner, key, op, out, [&](const Napi::Value& v) { return ReadColorSpec(r, v, key); });
}

void ReadI32Op(CheckedSpecReader& r, const Napi::Object& owner, const char* key, fm_patch_op_t& op, int32_t& out) {
  ReadOp(r, owner, key, op, out, [&](const Napi::Value& v) { return r.I32(v, key, 0); });
}

void ReadDoubleOp(CheckedSpecReader& r, const Napi::Object& owner, const char* key, fm_patch_op_t& op, double& out) {
  ReadOp(r, owner, key, op, out, [&](const Napi::Value& v) { return r.Double(v, key, 0.0); });
}

void ReadBoolOp(CheckedSpecReader& r, const Napi::Object& owner, const char* key, fm_patch_op_t& op, int32_t& out) {
  ReadOp(r, owner, key, op, out, [&](const Napi::Value& v) { return r.Bool(v, key, false) ? 1 : 0; });
}

/// Keeps the std::string bytes and literal point lists behind the pointer fields of a patch alive.
struct Arena {
  std::deque<std::string> strings;
  std::deque<std::vector<fm_chart_literal_info_t>> literals;
};

const char* Intern(Arena& arena, std::string text) {
  arena.strings.push_back(std::move(text));
  return arena.strings.back().c_str();
}

void ReadStringOp(CheckedSpecReader& r, const Napi::Object& owner, const char* key, Arena& arena, fm_patch_op_t& op,
                  const char*& out) {
  const Field f = Probe(r, owner, key);
  op = f.op;
  if (f.op == FM_PATCH_OP_SET) {
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
    t.kind = FM_CHART_TITLE_KIND_TEXT;
    r.String(v, key, &text);
  }
  t.text = Intern(arena, std::move(text));
  return t;
}

/// The points of a literal source: a number is a number point, a string a text point; `idx` is the position.
void ReadLiteral(CheckedSpecReader& r, const Napi::Value& value, const char* key, Arena& arena,
                 fm_chart_data_source_spec_t& out) {
  Napi::Array list;
  if (!r.Array(value, key, &list)) {
    return;
  }
  arena.literals.emplace_back();
  std::vector<fm_chart_literal_info_t>& points = arena.literals.back();
  for (uint32_t i = 0; i < list.Length() && r.ok(); ++i) {
    Napi::Value element;
    if (!r.ArrayElement(list, i, &element)) {
      continue;
    }
    fm_chart_literal_info_t point = FM_STRUCT_INIT(fm_chart_literal_info_t);
    point.idx = i;
    if (element.IsString()) {
      std::string text;
      r.String(element, key, &text);
      point.is_text = 1;
      point.text = Intern(arena, std::move(text));
    } else {
      point.number = r.Double(element, key, 0.0);
    }
    points.push_back(point);
  }
  out.literal = points.empty() ? nullptr : points.data();
  out.literal_count = static_cast<uint32_t>(points.size());
}

/// A series input: absent keeps, `null` removes, a string is a reference (an empty one removes), and an object is
/// the read-back shape `{ kind, formula, literal }`.
void ReadSource(CheckedSpecReader& r, const Napi::Object& owner, const char* key, Arena& arena, fm_patch_op_t& op,
                fm_chart_data_source_spec_t& out) {
  const Field f = Probe(r, owner, key);
  op = f.op;
  if (f.op != FM_PATCH_OP_SET) {
    return;
  }
  if (f.value.IsString()) {
    std::string text;
    r.String(f.value, key, &text);
    if (text.empty()) {
      out.kind = FM_CHART_SOURCE_KIND_NONE;
    } else {
      out.kind = FM_CHART_SOURCE_KIND_REF;
      out.formula = Intern(arena, std::move(text));
    }
    return;
  }
  Napi::Object spec;
  if (!r.Object(f.value, key, &spec)) {
    r.FailType(key, "string, null or an object");
    return;
  }
  out.kind = r.I32(spec, "kind", FM_CHART_SOURCE_KIND_NONE);
  std::string formula;
  if (r.String(spec, "formula", &formula)) {
    out.formula = Intern(arena, std::move(formula));
  }
  if (out.kind == FM_CHART_SOURCE_KIND_LITERAL) {
    ReadLiteral(r, r.Get(spec, "literal"), key, arena, out);
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
  ReadColorOp(r, o, "color", out.color_op, out.color);
  ReadDoubleOp(r, o, "transparencyPct", out.transparency_pct_op, out.transparency_pct);
}

void ReadLinePatch(CheckedSpecReader& r, const Napi::Object& owner, fm_chart_line_patch_t& out) {
  Napi::Object o;
  if (!ReadChild(r, owner, "line", &o)) {
    return;
  }
  ReadColorOp(r, o, "color", out.color_op, out.color);
  ReadDoubleOp(r, o, "transparencyPct", out.transparency_pct_op, out.transparency_pct);
  ReadDoubleOp(r, o, "widthPt", out.width_pt_op, out.width_pt);
  ReadI32Op(r, o, "dash", out.dash_op, out.dash);
  ReadBoolOp(r, o, "none", out.none_op, out.none);
}

void ReadMarkerPatch(CheckedSpecReader& r, const Napi::Object& owner, fm_patch_op_t& op, fm_chart_marker_patch_t& out) {
  const Field f = Probe(r, owner, "marker");
  op = f.op;
  Napi::Object o;
  if (f.op != FM_PATCH_OP_SET || !r.Object(f.value, "marker", &o)) {
    return;
  }
  ReadI32Op(r, o, "symbol", out.symbol_op, out.symbol);
  ReadDoubleOp(r, o, "sizePt", out.size_pt_op, out.size_pt);
  ReadFillPatch(r, o, out.fill);
  ReadLinePatch(r, o, out.line);
}

void ReadLabelsPatch(CheckedSpecReader& r, const Napi::Object& owner, Arena& arena, fm_patch_op_t& op,
                     fm_chart_labels_patch_t& out) {
  const Field f = Probe(r, owner, "labels");
  op = f.op;
  Napi::Object o;
  if (f.op != FM_PATCH_OP_SET || !r.Object(f.value, "labels", &o)) {
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
  ReadColorOp(r, o, "fontColor", out.font_color_op, out.font_color);
  ReadDoubleOp(r, o, "fontSizePt", out.font_size_pt_op, out.font_size_pt);
  ReadBoolOp(r, o, "sourceLinked", out.source_linked_op, out.source_linked);
  ReadBoolOp(r, o, "deleted", out.deleted_op, out.deleted);
}

void ReadSeriesPatch(CheckedSpecReader& r, const Napi::Object& spec, Arena& arena, fm_chart_series_patch_t& out) {
  ReadSource(r, spec, "name", arena, out.name_op, out.name);
  ReadSource(r, spec, "categories", arena, out.categories_op, out.categories);
  ReadSource(r, spec, "values", arena, out.values_op, out.values);
  ReadSource(r, spec, "x", arena, out.x_op, out.x);
  ReadSource(r, spec, "y", arena, out.y_op, out.y);
  ReadSource(r, spec, "size", arena, out.size_op, out.size);
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

/// Turns every member a replacement record left out (KEEP) into CLEAR. A marker symbol cannot be cleared, so it
/// stays KEEP.
void ClearOmittedMembers(fm_chart_point_patch_t& patch) {
  const auto clear = [](fm_patch_op_t& op) {
    if (op == FM_PATCH_OP_KEEP) {
      op = FM_PATCH_OP_CLEAR;
    }
  };
  const auto clear_fill = [&](fm_chart_fill_patch_t& fill) {
    clear(fill.color_op);
    clear(fill.transparency_pct_op);
  };
  const auto clear_line = [&](fm_chart_line_patch_t& line) {
    clear(line.color_op);
    clear(line.transparency_pct_op);
    clear(line.width_pt_op);
    clear(line.dash_op);
    clear(line.none_op);
  };
  clear_fill(patch.fill);
  clear_line(patch.line);
  clear(patch.marker_op);
  if (patch.marker_op == FM_PATCH_OP_SET) {
    clear(patch.marker.size_pt_op);
    clear_fill(patch.marker.fill);
    clear_line(patch.marker.line);
  }
}

int32_t HeaderModeArg(CheckedSpecReader& r, const Napi::CallbackInfo& info, size_t idx, const char* key) {
  return info.Length() > idx ? r.Tristate(info[idx], key) : FM_TRISTATE_DEFAULT;
}

/// Reads the `number[]` `value` into `out`; absent reads as empty.
void ReadUint32List(CheckedSpecReader& r, const Napi::Value& value, const char* key, std::vector<uint32_t>* out) {
  Napi::Array list;
  if (!r.Array(value, key, &list)) {
    return;
  }
  for (uint32_t i = 0; i < list.Length() && r.ok(); ++i) {
    Napi::Value element;
    if (r.ArrayElement(list, i, &element)) {
      out->push_back(r.U32(element, key, 0U));
    }
  }
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
  std::string description;
  std::string title;
  reader.String(spec, "source", &source);
  bool has_title = false;
  {
    const Field f = Probe(reader, spec, "title");
    if (f.op == FM_PATCH_OP_SET) {
      reader.String(f.value, "title", &title);
      has_title = true;
    }
  }
  fm_chart_insert_spec_t opts = FM_STRUCT_INIT(fm_chart_insert_spec_t);
  opts.type = reader.I32(spec, "type", FM_CHART_PLOT_TYPE_BAR);
  opts.bar_direction = reader.I32(spec, "barDirection", FM_CHART_BAR_DIRECTION_COLUMN);
  opts.grouping = reader.I32(spec, "grouping", FM_CHART_GROUPING_CLUSTERED);
  opts.source = source.c_str();
  opts.direction = reader.I32(spec, "direction", FM_CHART_SERIES_DIRECTION_AUTO);
  opts.header_row = reader.Tristate(spec, "headerRow");
  opts.header_column = reader.Tristate(spec, "headerColumn");
  opts.title = has_title ? title.c_str() : nullptr;
  // An omitted legend is Excel's default bottom legend; `null` is none.
  const Napi::Value legend = reader.Get(spec, "legend");
  opts.legend =
      legend.IsNull() ? FM_CHART_LEGEND_POSITION_ABSENT : reader.I32(legend, "legend", FM_CHART_LEGEND_POSITION_BOTTOM);
  opts.placement = PullImageInsert(reader, reader.Get(spec, "placement"), FM_ANCHOR_KIND_TWO_CELL, name, description);
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
  // `options` is optional: `{ barDirection?, grouping? }`.
  int32_t bar_direction = FM_CHART_BAR_DIRECTION_COLUMN;
  int32_t grouping = FM_CHART_GROUPING_CLUSTERED;
  if (info.Length() > 4 && !info[4].IsUndefined() && !info[4].IsNull()) {
    Napi::Object options;
    if (!PatchObjectArg(info, 4, "setChartType", &options)) {
      return env.Undefined();
    }
    CheckedSpecReader reader(env);
    bar_direction = reader.I32(options, "barDirection", FM_CHART_BAR_DIRECTION_COLUMN);
    grouping = reader.I32(options, "grouping", FM_CHART_GROUPING_CLUSTERED);
    if (!reader.ok()) {
      return env.Undefined();
    }
  }
  return ValueOrThrow(env,
                      fm_sheet_set_chart_type(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2),
                                              ArgI32(info, 3), bar_direction, grouping),
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
  title.kind = FM_CHART_TITLE_KIND_NONE;
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
  // `null` removes the legend.
  if (info[2].IsNull()) {
    return ValueOrThrow(env, fm_sheet_set_chart_legend(handle_, ArgU32(info, 0), ArgU32(info, 1), nullptr),
                        env.Undefined());
  }
  Napi::Object spec;
  if (!PatchObjectArg(info, 2, "setChartLegend", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  std::vector<uint32_t> hidden;
  ReadUint32List(reader, reader.Get(spec, "hiddenEntries"), "hiddenEntries", &hidden);
  fm_chart_legend_t legend = FM_STRUCT_INIT(fm_chart_legend_t);
  legend.position = reader.I32(spec, "position", FM_CHART_LEGEND_POSITION_RIGHT);
  legend.overlay = reader.Bool(spec, "overlay", false) ? 1 : 0;
  legend.hidden_entries = hidden.empty() ? nullptr : hidden.data();
  legend.hidden_entries_count = static_cast<uint32_t>(hidden.size());
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_set_chart_legend(handle_, ArgU32(info, 0), ArgU32(info, 1), &legend),
                      env.Undefined());
}

Napi::Value Workbook::UpdateChartLegend(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Object spec;
  if (!PatchObjectArg(info, 2, "updateChartLegend", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  fm_chart_legend_patch_t patch = FM_STRUCT_INIT(fm_chart_legend_patch_t);
  ReadI32Op(reader, spec, "position", patch.position_op, patch.position);
  ReadBoolOp(reader, spec, "overlay", patch.overlay_op, patch.overlay);
  std::vector<uint32_t> hidden;
  const Field hidden_field = Probe(reader, spec, "hiddenEntries");
  if (hidden_field.op == FM_PATCH_OP_SET) {
    ReadUint32List(reader, hidden_field.value, "hiddenEntries", &hidden);
    patch.hidden_entries_op = FM_PATCH_OP_SET;
    patch.hidden_entries = hidden.empty() ? nullptr : hidden.data();
    patch.hidden_entries_count = static_cast<uint32_t>(hidden.size());
  } else {
    patch.hidden_entries_op = hidden_field.op;
  }
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_chart_legend(handle_, ArgU32(info, 0), ArgU32(info, 1), &patch),
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
  fm_chart_axis_patch_t patch = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  ReadBoolOp(reader, spec, "visible", patch.visible_op, patch.visible);
  const Field title = Probe(reader, spec, "title");
  patch.title_op = title.op;
  if (title.op == FM_PATCH_OP_SET) {
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
  ReadDoubleOp(reader, spec, "crossesAt", patch.crosses_at_op, patch.crosses_at);
  ReadI32Op(reader, spec, "tickLabelPosition", patch.tick_label_position_op, patch.tick_label_position);
  ReadDoubleOp(reader, spec, "minorUnit", patch.minor_unit_op, patch.minor_unit);
  ReadBoolOp(reader, spec, "minorGridlines", patch.minor_gridlines_op, patch.minor_gridlines);
  ReadI32Op(reader, spec, "position", patch.position_op, patch.position);
  ReadBoolOp(reader, spec, "sourceLinked", patch.source_linked_op, patch.source_linked);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env,
                      fm_sheet_update_chart_axis(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &patch),
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
  fm_chart_display_patch_t patch = FM_STRUCT_INIT(fm_chart_display_patch_t);
  ReadI32Op(reader, spec, "blanksAs", patch.blanks_as_op, patch.blanks_as);
  ReadBoolOp(reader, spec, "plotVisibleOnly", patch.plot_visible_only_op, patch.plot_visible_only);
  ReadBoolOp(reader, spec, "naAsBlank", patch.na_as_blank_op, patch.na_as_blank);
  ReadI32Op(reader, spec, "style", patch.style_op, patch.style);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_sheet_update_chart_display(handle_, ArgU32(info, 0), ArgU32(info, 1), &patch),
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
  fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  ReadSeriesPatch(reader, spec, arena, patch);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(
      env,
      fm_sheet_update_chart_series(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), ArgU32(info, 3), &patch),
      env.Undefined());
}

Napi::Value Workbook::UpdateChartPoint(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Object spec;
  if (!PatchObjectArg(info, 5, "updateChartPoint", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  fm_chart_point_patch_t patch = FM_STRUCT_INIT(fm_chart_point_patch_t);
  ReadPointPatch(reader, spec, patch);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env,
                      fm_sheet_update_chart_point(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2),
                                                  ArgU32(info, 3), ArgU32(info, 4), &patch),
                      env.Undefined());
}

// The record replaces the point's formatting: a member it carries is set and a member it omits is cleared;
// `null` clears every member, which removes the override.
Napi::Value Workbook::SetChartPoint(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_chart_point_patch_t patch = FM_STRUCT_INIT(fm_chart_point_patch_t);
  if (!info[5].IsNull()) {
    Napi::Object spec;
    if (!PatchObjectArg(info, 5, "setChartPoint", &spec)) {
      return env.Undefined();
    }
    CheckedSpecReader reader(env);
    ReadPointPatch(reader, spec, patch);
    if (!reader.ok()) {
      return env.Undefined();
    }
  }
  ClearOmittedMembers(patch);
  return ValueOrThrow(env,
                      fm_sheet_update_chart_point(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2),
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
  fm_chart_plot_patch_t patch = FM_STRUCT_INIT(fm_chart_plot_patch_t);
  ReadU32Op(reader, spec, "gapWidthPct", patch.gap_width_pct_op, patch.gap_width_pct);
  ReadI32Op(reader, spec, "overlap", patch.overlap_op, patch.overlap);
  ReadU32Op(reader, spec, "firstSliceAngleDeg", patch.first_slice_angle_deg_op, patch.first_slice_angle_deg);
  ReadBoolOp(reader, spec, "varyColors", patch.vary_colors_op, patch.vary_colors);
  ReadU32Op(reader, spec, "holeSize", patch.hole_size_op, patch.hole_size);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env,
                      fm_sheet_update_chart_plot(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &patch),
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
  fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  ReadSeriesPatch(reader, spec, arena, patch);
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t series_index = 0;
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
