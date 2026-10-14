// Chart bindings: the typed chart view and resolved series data, creation,
// and the typed edits. Reads copy the handle-owned C data onto the JS heap and
// release the handle; patch objects map `undefined` to keep, `null` to clear
// and a value to set.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <memory>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

struct ChartDeleter {
  void operator()(fm_chart_t* chart) const { fm_chart_destroy(chart); }
};
struct ChartDataDeleter {
  void operator()(fm_chart_data_t* data) const { fm_chart_data_destroy(data); }
};
using ChartHandle = std::unique_ptr<fm_chart_t, ChartDeleter>;
using ChartDataHandle = std::unique_ptr<fm_chart_data_t, ChartDataDeleter>;

// ---- Reads ----------------------------------------------------------

/// Records the first accessor failure; later accessors are skipped.
bool Step(fm_status_t& rc, fm_status_t next) {
  if (rc == 0) {
    rc = next;
  }
  return rc == 0;
}

Napi::Value NumberOrNull(Napi::Env env, double v) {
  return std::isnan(v) ? env.Null() : JsNumber(env, v);
}

Napi::Value OptionalNumber(Napi::Env env, int32_t has, double v) {
  return has != 0 ? JsNumber(env, v) : env.Null();
}

Napi::Value StringOrNull(Napi::Env env, const char* s) {
  return s != nullptr ? JsString(env, s) : env.Null();
}

Napi::Value TriToJs(Napi::Env env, int32_t tri) {
  return tri == FM_CHART_TRI_ABSENT ? env.Null() : Napi::Boolean::New(env, tri == FM_CHART_TRI_TRUE);
}

Napi::Object TitleToJs(Napi::Env env, const fm_chart_title_t& t) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("kind", JsNumber(env, t.kind));
  o.Set("text", JsString(env, t.text));
  o.Set("formula", JsString(env, t.formula));
  return o;
}

Napi::Object FillToJs(Napi::Env env, const fm_chart_fill_t& f) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("color", OptionalNumber(env, f.has_color, f.color));
  o.Set("transparency", NumberOrNull(env, f.transparency_pct));
  return o;
}

Napi::Object LineToJs(Napi::Env env, const fm_chart_line_t& l) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("color", OptionalNumber(env, l.has_color, l.color));
  o.Set("transparency", NumberOrNull(env, l.transparency_pct));
  o.Set("widthPt", NumberOrNull(env, l.width_pt));
  o.Set("dash", OptionalNumber(env, l.has_dash, l.dash));
  o.Set("none", JsBool(env, l.none));
  return o;
}

Napi::Object MarkerToJs(Napi::Env env, const fm_chart_marker_t& m) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("symbol", JsNumber(env, m.symbol));
  o.Set("size", OptionalNumber(env, m.has_size, m.size));
  o.Set("fill", FillToJs(env, m.fill));
  o.Set("line", LineToJs(env, m.line));
  return o;
}

Napi::Object LabelsToJs(Napi::Env env, const fm_chart_labels_t& l) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("showValue", JsBool(env, l.show_value));
  o.Set("showCategory", JsBool(env, l.show_category));
  o.Set("showSeries", JsBool(env, l.show_series));
  o.Set("showPercent", JsBool(env, l.show_percent));
  o.Set("showLegendKey", JsBool(env, l.show_legend_key));
  o.Set("separator", StringOrNull(env, l.separator));
  o.Set("position", OptionalNumber(env, l.has_position, l.position));
  o.Set("numberFormat", StringOrNull(env, l.number_format));
  o.Set("sourceLinked", JsBool(env, l.source_linked));
  o.Set("fontColor", OptionalNumber(env, l.has_font_color, l.font_color));
  o.Set("fontSizePt", NumberOrNull(env, l.font_size_pt));
  o.Set("deleted", JsBool(env, l.deleted));
  return o;
}

const char* AxisPositionName(int32_t position) {
  switch (position) {
    case FM_CHART_AXIS_POS_TOP:
      return "t";
    case FM_CHART_AXIS_POS_LEFT:
      return "l";
    case FM_CHART_AXIS_POS_RIGHT:
      return "r";
    default:
      return "b";
  }
}

Napi::Object DataSourceToJs(Napi::Env env, const fm_chart_t* chart, size_t plot, size_t series, int32_t role,
                            const fm_chart_data_source_t& src, fm_status_t& rc) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("kind", JsString(env, src.kind == FM_CHART_SOURCE_REF       ? "ref"
                              : src.kind == FM_CHART_SOURCE_LITERAL ? "literal"
                                                                    : "none"));
  o.Set("formula", JsString(env, src.formula));
  Napi::Array literal = Napi::Array::New(env);
  for (size_t i = 0; i < src.literal_count && rc == 0; ++i) {
    fm_chart_literal_t lit{};
    if (!Step(rc, fm_chart_series_literal_at(chart, plot, series, role, i, &lit))) {
      break;
    }
    literal.Set(static_cast<uint32_t>(i), lit.is_text != 0 ? JsString(env, lit.text) : JsNumber(env, lit.number));
  }
  o.Set("literal", literal);
  return o;
}

Napi::Object SeriesToJs(Napi::Env env, const fm_chart_t* chart, size_t plot, size_t index, fm_status_t& rc) {
  Napi::Object o = Napi::Object::New(env);
  fm_chart_series_t s{};
  if (!Step(rc, fm_chart_series_at(chart, plot, index, &s))) {
    return o;
  }
  o.Set("idx", JsNumber(env, s.idx));
  o.Set("order", JsNumber(env, s.order));
  o.Set("filtered", JsBool(env, s.filtered));
  o.Set("name", DataSourceToJs(env, chart, plot, index, FM_CHART_ROLE_NAME, s.name, rc));
  o.Set("categories", DataSourceToJs(env, chart, plot, index, FM_CHART_ROLE_CATEGORIES, s.categories, rc));
  o.Set("values", DataSourceToJs(env, chart, plot, index, FM_CHART_ROLE_VALUES, s.values, rc));
  o.Set("x", DataSourceToJs(env, chart, plot, index, FM_CHART_ROLE_X, s.x, rc));
  o.Set("y", DataSourceToJs(env, chart, plot, index, FM_CHART_ROLE_Y, s.y, rc));
  o.Set("size", DataSourceToJs(env, chart, plot, index, FM_CHART_ROLE_SIZE, s.size, rc));
  o.Set("fill", FillToJs(env, s.fill));
  o.Set("line", LineToJs(env, s.line));
  o.Set("marker", s.has_marker != 0 ? Napi::Value(MarkerToJs(env, s.marker)) : env.Null());
  o.Set("smooth", TriToJs(env, s.smooth));
  Napi::Array points = Napi::Array::New(env);
  for (size_t i = 0; i < s.point_count && rc == 0; ++i) {
    fm_chart_point_t p{};
    if (!Step(rc, fm_chart_series_point_at(chart, plot, index, i, &p))) {
      break;
    }
    Napi::Object po = Napi::Object::New(env);
    po.Set("idx", JsNumber(env, p.idx));
    po.Set("fill", FillToJs(env, p.fill));
    po.Set("line", LineToJs(env, p.line));
    po.Set("marker", p.has_marker != 0 ? Napi::Value(MarkerToJs(env, p.marker)) : env.Null());
    points.Set(static_cast<uint32_t>(i), po);
  }
  o.Set("points", points);
  o.Set("labels", s.has_labels != 0 ? Napi::Value(LabelsToJs(env, s.labels)) : env.Null());
  return o;
}

Napi::Object PlotToJs(Napi::Env env, const fm_chart_t* chart, size_t index, fm_status_t& rc) {
  Napi::Object o = Napi::Object::New(env);
  fm_chart_plot_t p{};
  if (!Step(rc, fm_chart_plot_at(chart, index, &p))) {
    return o;
  }
  o.Set("type", JsNumber(env, p.type));
  o.Set("chartexLayout",
        p.chartex_layout != nullptr && p.chartex_layout[0] != '\0' ? JsString(env, p.chartex_layout) : env.Null());
  o.Set("barDirection", p.bar_dir == FM_CHART_BAR_DIR_COL   ? Napi::Value(JsString(env, "col"))
                        : p.bar_dir == FM_CHART_BAR_DIR_BAR ? Napi::Value(JsString(env, "bar"))
                                                            : env.Null());
  o.Set("grouping", OptionalNumber(env, p.has_grouping, p.grouping));
  o.Set("varyColors", Napi::Boolean::New(env, p.vary_colors == FM_CHART_TRI_TRUE));
  o.Set("gapWidth", OptionalNumber(env, p.has_gap_width, p.gap_width));
  o.Set("overlap", OptionalNumber(env, p.has_overlap, p.overlap));
  o.Set("holeSize", OptionalNumber(env, p.has_hole_size, p.hole_size));
  o.Set("firstSliceAngle", OptionalNumber(env, p.has_first_slice_angle, p.first_slice_angle));
  Napi::Array axis_ids = Napi::Array::New(env);
  for (size_t i = 0; i < p.axis_id_count && rc == 0; ++i) {
    uint32_t id = 0;
    if (Step(rc, fm_chart_plot_axis_id_at(chart, index, i, &id))) {
      axis_ids.Set(static_cast<uint32_t>(i), JsNumber(env, id));
    }
  }
  o.Set("axisIds", axis_ids);
  Napi::Array series = Napi::Array::New(env);
  for (size_t i = 0; i < p.series_count && rc == 0; ++i) {
    series.Set(static_cast<uint32_t>(i), SeriesToJs(env, chart, index, i, rc));
  }
  o.Set("series", series);
  return o;
}

Napi::Object AxisToJs(Napi::Env env, const fm_chart_axis_t& a) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("id", JsNumber(env, a.id));
  o.Set("kind", JsNumber(env, a.kind));
  o.Set("position", JsString(env, AxisPositionName(a.position)));
  o.Set("deleted", JsBool(env, a.deleted));
  o.Set("crossAxis", JsNumber(env, a.cross_axis));
  o.Set("crosses", JsNumber(env, a.crosses));
  o.Set("crossesAt", NumberOrNull(env, a.crosses_at));
  o.Set("min", NumberOrNull(env, a.min));
  o.Set("max", NumberOrNull(env, a.max));
  o.Set("majorUnit", NumberOrNull(env, a.major_unit));
  o.Set("minorUnit", NumberOrNull(env, a.minor_unit));
  o.Set("logBase", NumberOrNull(env, a.log_base));
  o.Set("reversed", JsBool(env, a.reversed));
  o.Set("numberFormat", StringOrNull(env, a.number_format));
  o.Set("sourceLinked", JsBool(env, a.source_linked));
  o.Set("title", TitleToJs(env, a.title));
  o.Set("majorGridlines", JsBool(env, a.major_gridlines));
  o.Set("minorGridlines", JsBool(env, a.minor_gridlines));
  o.Set("tickLabelPosition", StringOrNull(env, a.tick_label_pos));
  return o;
}

Napi::Object ChartToJs(Napi::Env env, const fm_chart_t* chart, fm_status_t& rc) {
  Napi::Object o = Napi::Object::New(env);
  fm_chart_info_t info{};
  if (!Step(rc, fm_chart_info(chart, &info))) {
    return o;
  }
  o.Set("objectId", JsNumber(env, info.object_id));
  o.Set("name", JsString(env, info.name));
  o.Set("description", JsString(env, info.descr));
  o.Set("isChartex", JsBool(env, info.is_chartex));
  o.Set("isPivotChart", JsBool(env, info.is_pivot_chart));
  o.Set("onChartsheet", JsBool(env, info.on_chartsheet));
  o.Set("title", TitleToJs(env, info.title));
  o.Set("autoTitleDeleted", TriToJs(env, info.auto_title_deleted));

  Napi::Object legend = Napi::Object::New(env);
  legend.Set("position", JsNumber(env, info.legend_position));
  legend.Set("overlay", JsBool(env, info.legend_overlay));
  Napi::Array hidden = Napi::Array::New(env);
  for (size_t i = 0; i < info.legend_hidden_count && rc == 0; ++i) {
    uint32_t idx = 0;
    if (Step(rc, fm_chart_legend_hidden_at(chart, i, &idx))) {
      hidden.Set(static_cast<uint32_t>(i), JsNumber(env, idx));
    }
  }
  legend.Set("hiddenSeries", hidden);
  o.Set("legend", legend);

  o.Set("blanksAs", JsNumber(env, info.blanks_as));
  o.Set("plotVisibleOnly", TriToJs(env, info.plot_vis_only));
  o.Set("date1904", JsBool(env, info.date1904));
  o.Set("style", OptionalNumber(env, info.has_style, info.style));

  Napi::Array plots = Napi::Array::New(env);
  for (size_t i = 0; i < info.plot_count && rc == 0; ++i) {
    plots.Set(static_cast<uint32_t>(i), PlotToJs(env, chart, i, rc));
  }
  o.Set("plots", plots);
  Napi::Array axes = Napi::Array::New(env);
  for (size_t i = 0; i < info.axis_count && rc == 0; ++i) {
    fm_chart_axis_t axis{};
    if (Step(rc, fm_chart_axis_at(chart, i, &axis))) {
      axes.Set(static_cast<uint32_t>(i), AxisToJs(env, axis));
    }
  }
  o.Set("axes", axes);
  Napi::Array hidden_categories = Napi::Array::New(env);
  for (size_t i = 0; i < info.hidden_category_count && rc == 0; ++i) {
    uint32_t idx = 0;
    if (Step(rc, fm_chart_hidden_category_at(chart, i, &idx))) {
      hidden_categories.Set(static_cast<uint32_t>(i), JsNumber(env, idx));
    }
  }
  o.Set("hiddenCategories", hidden_categories);

  if (info.source_range != nullptr) {
    Napi::Object source = Napi::Object::New(env);
    source.Set("range", JsString(env, info.source_range));
    source.Set("direction", JsNumber(env, info.source_direction));
    source.Set("headerRow", JsBool(env, info.source_header_row));
    source.Set("headerColumn", JsBool(env, info.source_header_column));
    o.Set("source", source);
  } else {
    o.Set("source", env.Null());
  }
  return o;
}

Napi::Float64Array CopyDoubles(Napi::Env env, const double* src, size_t n) {
  Napi::Float64Array out = Napi::Float64Array::New(env, n);
  for (size_t i = 0; i < n; ++i) {
    out[i] = src[i];
  }
  return out;
}

Napi::Object SheetRangeToJs(Napi::Env env, const fm_chart_sheet_range_t& r) {
  Napi::Object o = Napi::Object::New(env);
  o.Set("sheet", JsNumber(env, static_cast<double>(r.sheet)));
  o.Set("top", JsNumber(env, r.top));
  o.Set("left", JsNumber(env, r.left));
  o.Set("bottom", JsNumber(env, r.bottom));
  o.Set("right", JsNumber(env, r.right));
  return o;
}

Napi::Object DataSeriesToJs(Napi::Env env, const fm_chart_data_t* data, size_t index, fm_status_t& rc) {
  static const char* const kRoleKeys[] = {"name", "categories", "values", "x", "y", "size"};
  Napi::Object o = Napi::Object::New(env);
  fm_chart_data_series_t s{};
  fm_chart_data_points_t pts{};
  if (!Step(rc, fm_chart_data_series_at(data, index, &s)) ||
      !Step(rc, fm_chart_data_series_points(data, index, &pts))) {
    return o;
  }
  o.Set("plotIndex", JsNumber(env, static_cast<double>(s.plot_index)));
  o.Set("idx", JsNumber(env, s.idx));
  o.Set("name", StringOrNull(env, s.name));
  o.Set("formatCode", JsString(env, s.format_code));
  o.Set("xFormatCode", JsString(env, s.x_format_code));
  o.Set("filtered", JsBool(env, s.filtered));
  o.Set("values", CopyDoubles(env, pts.values, pts.count));
  o.Set("states", CopyBytes(env, pts.states, pts.count));
  o.Set("x", pts.x != nullptr ? Napi::Value(CopyDoubles(env, pts.x, pts.count)) : env.Null());
  o.Set("sizes", pts.sizes != nullptr ? Napi::Value(CopyDoubles(env, pts.sizes, pts.count)) : env.Null());
  Napi::Object ranges = Napi::Object::New(env);
  for (int32_t role = 0; role < 6 && rc == 0; ++role) {
    Napi::Array list = Napi::Array::New(env);
    for (size_t i = 0; i < s.range_count[role] && rc == 0; ++i) {
      fm_chart_sheet_range_t r{};
      if (Step(rc, fm_chart_data_range_at(data, index, role, i, &r))) {
        list.Set(static_cast<uint32_t>(i), SheetRangeToJs(env, r));
      }
    }
    ranges.Set(kRoleKeys[role], list);
  }
  o.Set("ranges", ranges);
  return o;
}

Napi::Object ChartDataToJs(Napi::Env env, const fm_chart_data_t* data, fm_status_t& rc) {
  Napi::Object o = Napi::Object::New(env);
  fm_chart_data_info_t info{};
  if (!Step(rc, fm_chart_data_info(data, &info))) {
    return o;
  }
  o.Set("domain", JsString(env, info.is_xy != 0 ? "xy" : "category"));
  o.Set("pointCount", JsNumber(env, static_cast<double>(info.point_count)));

  Napi::Object categories = Napi::Object::New(env);
  Napi::Array labels = Napi::Array::New(env);
  const bool numeric = info.category_value_count != 0;
  Napi::Float64Array values = Napi::Float64Array::New(env, numeric ? info.category_label_count : 0);
  for (size_t i = 0; i < info.category_label_count && rc == 0; ++i) {
    const char* label = nullptr;
    double value = 0.0;
    if (Step(rc, fm_chart_data_category_at(data, i, &label, &value))) {
      labels.Set(static_cast<uint32_t>(i), JsString(env, label));
      if (numeric) {
        values[i] = value;
      }
    }
  }
  Napi::Array levels = Napi::Array::New(env);
  for (size_t level = 0; level < info.category_level_count && rc == 0; ++level) {
    size_t count = 0;
    Napi::Array row = Napi::Array::New(env);
    if (Step(rc, fm_chart_data_category_level_size(data, level, &count))) {
      for (size_t i = 0; i < count && rc == 0; ++i) {
        const char* label = nullptr;
        if (Step(rc, fm_chart_data_category_level_at(data, level, i, &label))) {
          row.Set(static_cast<uint32_t>(i), JsString(env, label));
        }
      }
    }
    levels.Set(static_cast<uint32_t>(level), row);
  }
  categories.Set("labels", labels);
  categories.Set("levels", levels);
  categories.Set("values", numeric ? Napi::Value(values) : env.Null());
  categories.Set("formatCode", JsString(env, info.category_format_code));
  categories.Set("generated", JsBool(env, info.categories_generated));
  o.Set("categories", categories);

  Napi::Array series = Napi::Array::New(env);
  for (size_t i = 0; i < info.series_count && rc == 0; ++i) {
    series.Set(static_cast<uint32_t>(i), DataSeriesToJs(env, data, i, rc));
  }
  o.Set("series", series);
  o.Set("titleText", StringOrNull(env, info.title_text));

  Napi::Array axis_titles = Napi::Array::New(env);
  for (size_t i = 0; i < info.axis_title_count && rc == 0; ++i) {
    fm_chart_axis_title_text_t t{};
    if (Step(rc, fm_chart_data_axis_title_at(data, i, &t))) {
      Napi::Object item = Napi::Object::New(env);
      item.Set("axisId", JsNumber(env, t.axis_id));
      item.Set("text", JsString(env, t.text));
      axis_titles.Set(static_cast<uint32_t>(i), item);
    }
  }
  o.Set("axisTitleTexts", axis_titles);
  Napi::Array unresolved = Napi::Array::New(env);
  for (size_t i = 0; i < info.unresolved_count && rc == 0; ++i) {
    fm_chart_unresolved_t u{};
    if (Step(rc, fm_chart_data_unresolved_at(data, i, &u))) {
      Napi::Object item = Napi::Object::New(env);
      item.Set("seriesIdx", JsNumber(env, u.series_idx));
      item.Set("source", JsString(env, u.source));
      item.Set("axisId", JsNumber(env, u.axis_id));
      unresolved.Set(static_cast<uint32_t>(i), item);
    }
  }
  o.Set("unresolved", unresolved);
  return o;
}

/// `{ status, <field>: value | null }`; the value is dropped on failure.
Napi::Object ReadResult(Napi::Env env, fm_status_t rc, const char* field, Napi::Value value) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, rc));
  out.Set(field, rc == 0 ? value : env.Null());
  return out;
}

// ---- Patch input ----------------------------------------------------

/// A patch field: absent keeps, `null` clears, anything else sets `value`.
struct Field {
  int32_t op = FM_CHART_PATCH_KEEP;
  Napi::Value value;
};

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

void ReadU32Op(CheckedSpecReader& r, const Napi::Object& owner, const char* key, int32_t& op, uint32_t& out) {
  const Field f = Probe(r, owner, key);
  op = f.op;
  if (f.op == FM_CHART_PATCH_SET) {
    out = r.U32(f.value, key, 0U);
  }
}

void ReadI32Op(CheckedSpecReader& r, const Napi::Object& owner, const char* key, int32_t& op, int32_t& out) {
  const Field f = Probe(r, owner, key);
  op = f.op;
  if (f.op == FM_CHART_PATCH_SET) {
    out = r.I32(f.value, key, 0);
  }
}

void ReadDoubleOp(CheckedSpecReader& r, const Napi::Object& owner, const char* key, int32_t& op, double& out) {
  const Field f = Probe(r, owner, key);
  op = f.op;
  if (f.op == FM_CHART_PATCH_SET) {
    out = r.Double(f.value, key, 0.0);
  }
}

void ReadBoolOp(CheckedSpecReader& r, const Napi::Object& owner, const char* key, int32_t& op, int32_t& out) {
  const Field f = Probe(r, owner, key);
  op = f.op;
  if (f.op == FM_CHART_PATCH_SET) {
    out = r.Bool(f.value, key, false) ? 1 : 0;
  }
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

void ReadFillPatch(CheckedSpecReader& r, const Napi::Object& owner, fm_chart_fill_patch& out) {
  Napi::Object o;
  if (!ReadChild(r, owner, "fill", &o)) {
    return;
  }
  ReadU32Op(r, o, "color", out.color_op, out.color);
  ReadDoubleOp(r, o, "transparency", out.transparency_op, out.transparency_pct);
}

void ReadLinePatch(CheckedSpecReader& r, const Napi::Object& owner, fm_chart_line_patch& out) {
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

void ReadMarkerPatch(CheckedSpecReader& r, const Napi::Object& owner, int32_t& op, fm_chart_marker_patch& out) {
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
                     fm_chart_labels_patch& out) {
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

void ReadSeriesPatch(CheckedSpecReader& r, const Napi::Object& spec, Arena& arena, fm_chart_series_patch& out) {
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

void ReadPointPatch(CheckedSpecReader& r, const Napi::Object& spec, fm_chart_point_patch& out) {
  ReadFillPatch(r, spec, out.fill);
  ReadLinePatch(r, spec, out.line);
  ReadMarkerPatch(r, spec, out.marker_op, out.marker);
}

/// Reads the optional object argument `info[idx]`; absent reads as `{}`. False after
/// throwing when a value other than an object was supplied.
bool ObjectArg(const Napi::CallbackInfo& info, size_t idx, const char* api, Napi::Object* out) {
  if (info.Length() <= idx || info[idx].IsUndefined()) {
    *out = Napi::Object::New(info.Env());
    return true;
  }
  if (!info[idx].IsObject() || info[idx].IsArray()) {
    Napi::TypeError::New(info.Env(), std::string(api) + " expects an object argument").ThrowAsJavaScriptException();
    return false;
  }
  *out = info[idx].As<Napi::Object>();
  return true;
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

// ---- Reads ----------------------------------------------------------

Napi::Value Workbook::GetChart(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ReadResult(env, kBindingInvalidHandle, "chart", env.Null());
  }
  fm_chart_t* raw = nullptr;
  fm_status_t rc = fm_sheet_chart_open(handle_, ArgU32(info, 0), ArgU32(info, 1), &raw);
  const ChartHandle chart(raw);
  Napi::Value value = env.Null();
  if (rc == 0) {
    value = ChartToJs(env, chart.get(), rc);
  }
  return ReadResult(env, rc, "chart", value);
}

Napi::Value Workbook::GetChartData(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ReadResult(env, kBindingInvalidHandle, "data", env.Null());
  }
  fm_chart_data_t* raw = nullptr;
  fm_status_t rc = fm_sheet_chart_data_open(handle_, ArgU32(info, 0), ArgU32(info, 1), &raw);
  const ChartDataHandle data(raw);
  Napi::Value value = env.Null();
  if (rc == 0) {
    value = ChartDataToJs(env, data.get(), rc);
  }
  return ReadResult(env, rc, "data", value);
}

Napi::Value Workbook::PreviewChartData(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ReadResult(env, kBindingInvalidHandle, "data", env.Null());
  }
  Napi::Object spec;
  if (!ObjectArg(info, 2, "previewChartData", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  std::string source;
  const bool has_source = reader.String(spec, "source", &source);
  fm_chart_preview_overrides overrides{};
  overrides.source = has_source ? source.c_str() : nullptr;
  const auto set = [&](const char* key, int32_t& op, int32_t& out, bool is_bool) {
    const Field f = Probe(reader, spec, key);
    if (f.op == FM_CHART_PATCH_CLEAR) {
      // `null` leaves the override unset, as `undefined` does.
      return;
    }
    if (f.op == FM_CHART_PATCH_SET) {
      op = FM_CHART_PATCH_SET;
      out = is_bool ? (reader.Bool(f.value, key, false) ? 1 : 0) : reader.I32(f.value, key, 0);
    }
  };
  set("direction", overrides.direction_op, overrides.direction, false);
  set("headerRow", overrides.header_row_op, overrides.header_row, true);
  set("headerColumn", overrides.header_column_op, overrides.header_column, true);
  set("blanksAs", overrides.blanks_as_op, overrides.blanks_as, false);
  set("plotVisibleOnly", overrides.plot_vis_only_op, overrides.plot_vis_only, true);
  if (!reader.ok()) {
    return env.Undefined();
  }
  fm_chart_data_t* raw = nullptr;
  fm_status_t rc = fm_sheet_chart_preview_open(handle_, ArgU32(info, 0), ArgU32(info, 1), &overrides, &raw);
  const ChartDataHandle data(raw);
  Napi::Value value = env.Null();
  if (rc == 0) {
    value = ChartDataToJs(env, data.get(), rc);
  }
  return ReadResult(env, rc, "data", value);
}

// ---- Creation and edits ---------------------------------------------

Napi::Value Workbook::InsertChart(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("objectId", JsNumber(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  if (info.Length() < 2 || !info[1].IsObject()) {
    out.Set("status", MakeBindingArgumentError(env, "insertChart expects (sheet:number, options:object)"));
    return out;
  }
  CheckedSpecReader reader(env);
  const Napi::Object spec = info[1].As<Napi::Object>();
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
  fm_chart_insert opts{};
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
  const fm_status_t rc = fm_sheet_insert_chart(handle_, ArgU32(info, 0), &opts, &object_id);
  out.Set("objectId", JsNumber(env, rc == 0 ? object_id : 0));
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::SetChartType(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const int32_t grouping = info.Length() > 4 && !info[4].IsUndefined() ? ArgI32(info, 4) : FM_CHART_GROUPING_CLUSTERED;
  return MakeStatus(env, fm_sheet_set_chart_type(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2),
                                                 ArgI32(info, 3), grouping));
}

Napi::Value Workbook::SetChartTitle(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const Napi::Value arg = info.Length() > 2 ? info[2] : env.Null();
  if (!arg.IsNull() && !arg.IsString() && !arg.IsObject()) {
    return MakeBindingArgumentError(
        env, "setChartTitle expects (sheet:number, objectId:number, title:string|null|{kind,text?,formula?})");
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
  return MakeStatus(env, fm_sheet_set_chart_title(handle_, ArgU32(info, 0), ArgU32(info, 1), &title));
}

Napi::Value Workbook::SetChartLegend(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  CheckedSpecReader reader(env);
  std::vector<uint32_t> hidden;
  const bool has_hidden = info.Length() > 4 && !info[4].IsUndefined() && !info[4].IsNull();
  if (!ReadIndexList(reader, info, 4, "hiddenSeries", &hidden)) {
    return env.Undefined();
  }
  fm_chart_legend_patch legend{};
  legend.position = ArgI32(info, 2);
  legend.overlay = ArgBool(info, 3) ? 1 : 0;
  if (has_hidden) {
    legend.hidden_series_op = FM_CHART_PATCH_SET;
    legend.hidden_series = hidden.empty() ? nullptr : hidden.data();
    legend.hidden_series_count = hidden.size();
  }
  return MakeStatus(env, fm_sheet_set_chart_legend(handle_, ArgU32(info, 0), ArgU32(info, 1), &legend));
}

Napi::Value Workbook::SetChartAxis(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  Napi::Object spec;
  if (!ObjectArg(info, 3, "setChartAxis", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  Arena arena;
  fm_chart_axis_patch patch{};
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
  return MakeStatus(env, fm_sheet_set_chart_axis(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &patch));
}

Napi::Value Workbook::SetChartDisplay(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  // `plotVisibleOnly` null or omitted removes the element.
  int32_t plot_vis_only = FM_CHART_TRI_ABSENT;
  if (info.Length() > 3 && !info[3].IsUndefined() && !info[3].IsNull()) {
    plot_vis_only = ArgBool(info, 3) ? FM_CHART_TRI_TRUE : FM_CHART_TRI_FALSE;
  }
  return MakeStatus(
      env, fm_sheet_set_chart_display(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgI32(info, 2), plot_vis_only));
}

Napi::Value Workbook::SetChartSeries(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  Napi::Object spec;
  if (!ObjectArg(info, 4, "setChartSeries", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  Arena arena;
  fm_chart_series_patch patch{};
  ReadSeriesPatch(reader, spec, arena, patch);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_set_chart_series(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2),
                                                   ArgU32(info, 3), &patch));
}

Napi::Value Workbook::SetChartPoint(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  // A null or empty patch removes the point's formatting.
  Napi::Object spec = Napi::Object::New(env);
  if (info.Length() > 5 && !info[5].IsNull() && !ObjectArg(info, 5, "setChartPoint", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  fm_chart_point_patch patch{};
  ReadPointPatch(reader, spec, patch);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_set_chart_point(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2),
                                                  ArgU32(info, 3), ArgU32(info, 4), &patch));
}

Napi::Value Workbook::SetChartPlot(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  Napi::Object spec;
  if (!ObjectArg(info, 3, "setChartPlot", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  fm_chart_plot_patch patch{};
  ReadU32Op(reader, spec, "gapWidth", patch.gap_width_op, patch.gap_width);
  ReadI32Op(reader, spec, "overlap", patch.overlap_op, patch.overlap);
  ReadU32Op(reader, spec, "firstSliceAngle", patch.first_slice_angle_op, patch.first_slice_angle);
  ReadBoolOp(reader, spec, "varyColors", patch.vary_colors_op, patch.vary_colors);
  ReadU32Op(reader, spec, "holeSize", patch.hole_size_op, patch.hole_size);
  if (!reader.ok()) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_set_chart_plot(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &patch));
}

Napi::Value Workbook::SetChartFilter(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  CheckedSpecReader reader(env);
  std::vector<uint32_t> series;
  std::vector<uint32_t> categories;
  if (!ReadIndexList(reader, info, 2, "hiddenSeries", &series) ||
      !ReadIndexList(reader, info, 3, "hiddenCategories", &categories)) {
    return env.Undefined();
  }
  return MakeStatus(env, fm_sheet_set_chart_filter(
                             handle_, ArgU32(info, 0), ArgU32(info, 1), series.empty() ? nullptr : series.data(),
                             series.size(), categories.empty() ? nullptr : categories.data(), categories.size()));
}

Napi::Value Workbook::AddChartSeries(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("seriesIdx", JsNumber(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  Napi::Object spec;
  if (!ObjectArg(info, 3, "addChartSeries", &spec)) {
    return env.Undefined();
  }
  CheckedSpecReader reader(env);
  Arena arena;
  fm_chart_series_patch patch{};
  ReadSeriesPatch(reader, spec, arena, patch);
  if (!reader.ok()) {
    return env.Undefined();
  }
  uint32_t series_idx = 0;
  const fm_status_t rc =
      fm_sheet_add_chart_series(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), &patch, &series_idx);
  out.Set("seriesIdx", JsNumber(env, rc == 0 ? series_idx : 0));
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::RemoveChartSeries(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  return MakeStatus(
      env, fm_sheet_remove_chart_series(handle_, ArgU32(info, 0), ArgU32(info, 1), ArgU32(info, 2), ArgU32(info, 3)));
}

Napi::Value Workbook::SetChartSource(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  CheckedSpecReader reader(env);
  const int32_t header_row = HeaderModeArg(reader, info, 4, "headerRow");
  const int32_t header_column = HeaderModeArg(reader, info, 5, "headerColumn");
  if (!reader.ok()) {
    return env.Undefined();
  }
  const std::string source = ArgString(info, 2);
  return MakeStatus(env, fm_sheet_set_chart_source(handle_, ArgU32(info, 0), ArgU32(info, 1), source.c_str(),
                                                   ArgI32(info, 3), header_row, header_column));
}

}  // namespace formulon_node
