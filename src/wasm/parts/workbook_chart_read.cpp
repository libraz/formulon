//
// JsWorkbook chart reads: the typed chart view and the resolved series data.
// Reads copy the handle-owned C data onto the JS heap and release the handle.
// The output is driven by field tables over the C structs.

#include <emscripten/val.h>

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <memory>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

using emscripten::val;

struct ChartDeleter {
  void operator()(fm_chart_t* chart) const { fm_chart_destroy(chart); }
};
struct ChartDataDeleter {
  void operator()(fm_chart_data_t* data) const { fm_chart_data_destroy(data); }
};

template <class T>
T load(const void* base, size_t off) {
  T v;
  std::memcpy(&v, static_cast<const char*>(base) + off, sizeof(T));
  return v;
}

void push(val& a, const val& v) {
  a.call<void>("push", v);
}

/// Records the first accessor failure; later accessors are skipped.
bool step(fm_status_t& rc, fm_status_t next) {
  if (rc == 0) {
    rc = next;
  }
  return rc == 0;
}

// ---- Reads ----------------------------------------------------------

enum class Rd : uint8_t {
  kI32,
  kU32,
  kF64,
  kBool,
  kStr,
  kStrOrNull,
  kTextOrNull,
  kEnumOrNull,
  kOptF64,
  kOptI32,
  kOptU32,
  kTri,
  kColor,
  kObj,
  kObjOpt,
  kTitle,
};

/// One output property: where the value sits in the C struct, the `has_` flag for the optional kinds.
struct RdField {
  const char* key;
  uint16_t off;
  uint16_t has;
  Rd kind;
  const RdField* sub;
};

// clang-format off
#define RD(key, T, member, kind) {key, offsetof(T, member), 0, Rd::kind, nullptr}
#define RD_OPT(key, T, has, member, kind) {key, offsetof(T, member), offsetof(T, has), Rd::kind, nullptr}
#define RD_OBJ(key, T, member, table) {key, offsetof(T, member), 0, Rd::kObj, table}
#define RD_OBJ_OPT(key, T, has, member, table) {key, offsetof(T, member), offsetof(T, has), Rd::kObjOpt, table}
#define RD_TITLE(key, T, member, table) {key, offsetof(T, member), 0, Rd::kTitle, table}
#define RD_END {nullptr, 0, 0, Rd::kI32, nullptr}
// clang-format on

constexpr RdField kTitleRd[] = {
    RD("kind", fm_chart_title_t, kind, kI32),
    RD("text", fm_chart_title_t, text, kTextOrNull),
    RD("formula", fm_chart_title_t, formula, kTextOrNull),
    RD_END,
};

constexpr RdField kFillRd[] = {
    RD("color", fm_chart_fill_t, color, kColor),
    RD_OPT("transparencyPct", fm_chart_fill_t, has_transparency_pct, transparency_pct, kOptF64),
    RD_END,
};

// clang-format off
constexpr RdField kLineRd[] = {
    RD("color", fm_chart_line_t, color, kColor),
    RD_OPT("transparencyPct", fm_chart_line_t, has_transparency_pct, transparency_pct, kOptF64),
    RD_OPT("widthPt", fm_chart_line_t, has_width_pt, width_pt, kOptF64),
    RD_OPT("dash", fm_chart_line_t, has_dash, dash, kOptI32),
    RD("none", fm_chart_line_t, none, kBool),
    RD_END,
};
// clang-format on

constexpr RdField kMarkerRd[] = {
    RD("symbol", fm_chart_marker_t, symbol, kI32),
    RD_OPT("sizePt", fm_chart_marker_t, has_size_pt, size_pt, kOptU32),
    RD_OBJ("fill", fm_chart_marker_t, fill, kFillRd),
    RD_OBJ("line", fm_chart_marker_t, line, kLineRd),
    RD_END,
};

constexpr RdField kLabelsRd[] = {
    RD("showValue", fm_chart_labels_t, show_value, kBool),
    RD("showCategory", fm_chart_labels_t, show_category, kBool),
    RD("showSeries", fm_chart_labels_t, show_series, kBool),
    RD("showPercent", fm_chart_labels_t, show_percent, kBool),
    RD("showLegendKey", fm_chart_labels_t, show_legend_key, kBool),
    RD("separator", fm_chart_labels_t, separator, kStrOrNull),
    RD_OPT("position", fm_chart_labels_t, has_position, position, kOptI32),
    RD("numberFormat", fm_chart_labels_t, number_format, kStrOrNull),
    RD("sourceLinked", fm_chart_labels_t, source_linked, kBool),
    RD("fontColor", fm_chart_labels_t, font_color, kColor),
    RD_OPT("fontSizePt", fm_chart_labels_t, has_font_size_pt, font_size_pt, kOptF64),
    RD("deleted", fm_chart_labels_t, deleted, kBool),
    RD_END,
};

constexpr RdField kPointRd[] = {
    RD("index", fm_chart_point_info_t, idx, kU32),
    RD_OBJ("fill", fm_chart_point_info_t, fill, kFillRd),
    RD_OBJ("line", fm_chart_point_info_t, line, kLineRd),
    RD_OBJ_OPT("marker", fm_chart_point_info_t, has_marker, marker, kMarkerRd),
    RD_END,
};

constexpr RdField kSeriesRd[] = {
    RD("seriesId", fm_chart_series_info_t, series_id, kU32),
    RD("order", fm_chart_series_info_t, order, kU32),
    RD("filtered", fm_chart_series_info_t, filtered, kBool),
    RD_OBJ("fill", fm_chart_series_info_t, fill, kFillRd),
    RD_OBJ("line", fm_chart_series_info_t, line, kLineRd),
    RD_OBJ_OPT("marker", fm_chart_series_info_t, has_marker, marker, kMarkerRd),
    RD("smooth", fm_chart_series_info_t, smooth, kTri),
    RD_OBJ_OPT("labels", fm_chart_series_info_t, has_labels, labels, kLabelsRd),
    RD_END,
};

constexpr RdField kPlotRd[] = {
    RD("type", fm_chart_plot_info_t, type, kI32),
    RD_OPT("grouping", fm_chart_plot_info_t, has_grouping, grouping, kOptI32),
    RD_OPT("gapWidthPct", fm_chart_plot_info_t, has_gap_width_pct, gap_width_pct, kOptU32),
    RD_OPT("overlap", fm_chart_plot_info_t, has_overlap, overlap, kOptI32),
    RD_OPT("holeSize", fm_chart_plot_info_t, has_hole_size, hole_size, kOptU32),
    RD_OPT("firstSliceAngleDeg", fm_chart_plot_info_t, has_first_slice_angle_deg, first_slice_angle_deg, kOptU32),
    RD_END,
};

constexpr RdField kAxisRd[] = {
    RD("id", fm_chart_axis_info_t, id, kU32),
    RD("kind", fm_chart_axis_info_t, kind, kI32),
    RD("visible", fm_chart_axis_info_t, visible, kBool),
    RD("crossAxis", fm_chart_axis_info_t, cross_axis, kU32),
    RD("crosses", fm_chart_axis_info_t, crosses, kEnumOrNull),
    RD_OPT("crossesAt", fm_chart_axis_info_t, has_crosses_at, crosses_at, kOptF64),
    RD_OPT("min", fm_chart_axis_info_t, has_min, min, kOptF64),
    RD_OPT("max", fm_chart_axis_info_t, has_max, max, kOptF64),
    RD_OPT("majorUnit", fm_chart_axis_info_t, has_major_unit, major_unit, kOptF64),
    RD_OPT("minorUnit", fm_chart_axis_info_t, has_minor_unit, minor_unit, kOptF64),
    RD_OPT("logBase", fm_chart_axis_info_t, has_log_base, log_base, kOptF64),
    RD("reversed", fm_chart_axis_info_t, reversed, kBool),
    RD("numberFormat", fm_chart_axis_info_t, number_format, kStrOrNull),
    RD("sourceLinked", fm_chart_axis_info_t, source_linked, kBool),
    RD_TITLE("title", fm_chart_axis_info_t, title, kTitleRd),
    RD("majorGridlines", fm_chart_axis_info_t, major_gridlines, kBool),
    RD("minorGridlines", fm_chart_axis_info_t, minor_gridlines, kBool),
    RD("tickLabelPosition", fm_chart_axis_info_t, tick_label_position, kEnumOrNull),
    RD_END,
};

constexpr RdField kInfoRd[] = {
    RD("objectId", fm_chart_info_t, object_id, kU32),
    RD("isExtended", fm_chart_info_t, is_extended, kBool),
    RD("isPivotChart", fm_chart_info_t, is_pivot_chart, kBool),
    RD("onChartsheet", fm_chart_info_t, on_chartsheet, kBool),
    RD_TITLE("title", fm_chart_info_t, title, kTitleRd),
    RD("autoTitleDeleted", fm_chart_info_t, auto_title_deleted, kTri),
    RD("blanksAs", fm_chart_info_t, blanks_as, kEnumOrNull),
    RD("plotVisibleOnly", fm_chart_info_t, plot_visible_only, kTri),
    RD("naAsBlank", fm_chart_info_t, na_as_blank, kTri),
    RD("date1904", fm_chart_info_t, date1904, kBool),
    RD_OPT("style", fm_chart_info_t, has_style, style, kOptI32),
    RD_END,
};

val fields_to_val(const void* base, const RdField* f) {
  val o = val::object();
  for (; f->key != nullptr; ++f) {
    val v = val::null();
    switch (f->kind) {
      case Rd::kI32:
        v = val(load<int32_t>(base, f->off));
        break;
      case Rd::kU32:
        v = val(load<uint32_t>(base, f->off));
        break;
      case Rd::kF64:
        v = val(load<double>(base, f->off));
        break;
      case Rd::kBool:
        v = val(load<int32_t>(base, f->off) != 0);
        break;
      case Rd::kStr:
        v = val(string_from_cstr(load<const char*>(base, f->off)));
        break;
      case Rd::kStrOrNull:
        if (const char* s = load<const char*>(base, f->off)) {
          v = val(std::string(s));
        }
        break;
      case Rd::kTextOrNull:
        v = cstr_or_null(load<const char*>(base, f->off));
        break;
      case Rd::kEnumOrNull:
        if (const int32_t e = load<int32_t>(base, f->off); e >= 0) {
          v = val(e);
        }
        break;
      case Rd::kOptF64:
        if (load<int32_t>(base, f->has) != 0) {
          v = val(load<double>(base, f->off));
        }
        break;
      case Rd::kOptI32:
        if (load<int32_t>(base, f->has) != 0) {
          v = val(load<int32_t>(base, f->off));
        }
        break;
      case Rd::kOptU32:
        if (load<int32_t>(base, f->has) != 0) {
          v = val(load<uint32_t>(base, f->off));
        }
        break;
      case Rd::kTri:
        if (const int32_t tri = load<int32_t>(base, f->off); tri != FM_TRISTATE_DEFAULT) {
          v = val(tri == FM_TRISTATE_TRUE);
        }
        break;
      case Rd::kColor:
        v = js_color_spec_or_null(load<fm_color_t>(base, f->off));
        break;
      case Rd::kObj:
        v = fields_to_val(static_cast<const char*>(base) + f->off, f->sub);
        break;
      case Rd::kObjOpt:
        if (load<int32_t>(base, f->has) != 0) {
          v = fields_to_val(static_cast<const char*>(base) + f->off, f->sub);
        }
        break;
      case Rd::kTitle:
        // A title of kind NONE is no title.
        if (load<int32_t>(base, f->off + offsetof(fm_chart_title_t, kind)) != FM_CHART_TITLE_KIND_NONE) {
          v = fields_to_val(static_cast<const char*>(base) + f->off, f->sub);
        }
        break;
    }
    o.set(f->key, v);
  }
  return o;
}

// The six series inputs, in `fm_chart_role_t` order.
constexpr const char* kRoleKeys[] = {"name", "categories", "values", "x", "y", "size"};
constexpr int32_t kRoleCount = FM_CHART_ROLE_SIZE + 1;
static_assert(sizeof(kRoleKeys) / sizeof(kRoleKeys[0]) == kRoleCount, "role keys must cover every fm_chart_role_t");

val source_to_val(const fm_chart_t* chart, size_t plot, size_t series, int32_t role, const fm_chart_data_source_t& src,
                  fm_status_t& rc) {
  val o = val::object();
  o.set("kind", src.kind);
  o.set("formula", val(string_from_cstr(src.formula)));
  val literal = val::array();
  for (size_t i = 0; i < src.literal_count; ++i) {
    fm_chart_literal_info_t lit = FM_STRUCT_INIT(fm_chart_literal_info_t);
    if (!step(rc, fm_chart_get_series_literal_at(chart, plot, series, role, i, &lit))) {
      break;
    }
    push(literal, lit.is_text != 0 ? val(string_from_cstr(lit.text)) : val(lit.number));
  }
  o.set("literal", literal);
  return o;
}

val series_to_val(const fm_chart_t* chart, size_t plot, size_t index, fm_status_t& rc) {
  fm_chart_series_info_t s = FM_STRUCT_INIT(fm_chart_series_info_t);
  if (!step(rc, fm_chart_get_series_at(chart, plot, index, &s))) {
    return val::object();
  }
  val o = fields_to_val(&s, kSeriesRd);
  const fm_chart_data_source_t* const sources[] = {&s.name, &s.categories, &s.values, &s.x, &s.y, &s.size};
  for (int32_t role = 0; role < kRoleCount; ++role) {
    o.set(kRoleKeys[role], source_to_val(chart, plot, index, role, *sources[role], rc));
  }
  val points = val::array();
  for (size_t i = 0; i < s.point_count; ++i) {
    fm_chart_point_info_t p = FM_STRUCT_INIT(fm_chart_point_info_t);
    if (!step(rc, fm_chart_get_series_point_at(chart, plot, index, i, &p))) {
      break;
    }
    push(points, fields_to_val(&p, kPointRd));
  }
  o.set("points", points);
  return o;
}

val plot_to_val(const fm_chart_t* chart, size_t index, fm_status_t& rc) {
  fm_chart_plot_info_t p = FM_STRUCT_INIT(fm_chart_plot_info_t);
  if (!step(rc, fm_chart_get_plot_at(chart, index, &p))) {
    return val::object();
  }
  val o = fields_to_val(&p, kPlotRd);
  o.set("extendedLayout", p.extended_layout != FM_CHART_EXTENDED_LAYOUT_ABSENT ? val(p.extended_layout) : val::null());
  o.set("barDirection", p.bar_direction == FM_CHART_BAR_DIRECTION_ABSENT ? val::null() : val(p.bar_direction));
  o.set("varyColors", val(p.vary_colors == FM_TRISTATE_TRUE));
  val axis_ids = val::array();
  for (size_t i = 0; i < p.axis_id_count; ++i) {
    uint32_t id = 0;
    if (!step(rc, fm_chart_get_plot_axis_id_at(chart, index, i, &id))) {
      break;
    }
    push(axis_ids, val(id));
  }
  o.set("axisIds", axis_ids);
  val series = val::array();
  for (size_t i = 0; i < p.series_count && rc == 0; ++i) {
    push(series, series_to_val(chart, index, i, rc));
  }
  o.set("series", series);
  return o;
}

/// Reads `count` indices through `fetch` into a JS array.
template <class Fetch>
val indices_to_val(size_t count, fm_status_t& rc, Fetch fetch) {
  val a = val::array();
  for (size_t i = 0; i < count; ++i) {
    uint32_t idx = 0;
    if (!step(rc, fetch(i, &idx))) {
      break;
    }
    push(a, val(idx));
  }
  return a;
}

val chart_to_val(const fm_chart_t* chart, fm_status_t& rc) {
  fm_chart_info_t info = FM_STRUCT_INIT(fm_chart_info_t);
  if (!step(rc, fm_chart_get_info(chart, &info))) {
    return val::object();
  }
  val o = fields_to_val(&info, kInfoRd);

  val legend = val::null();
  if (info.legend_position != FM_CHART_LEGEND_POSITION_ABSENT) {
    legend = val::object();
    legend.set("position", val(info.legend_position));
    legend.set("overlay", val(info.legend_overlay != 0));
    legend.set("hiddenEntries", indices_to_val(info.legend_hidden_count, rc, [&](size_t i, uint32_t* out) {
                 return fm_chart_get_legend_hidden_at(chart, i, out);
               }));
  }
  o.set("legend", legend);

  val plots = val::array();
  for (size_t i = 0; i < info.plot_count && rc == 0; ++i) {
    push(plots, plot_to_val(chart, i, rc));
  }
  o.set("plots", plots);
  val axes = val::array();
  for (size_t i = 0; i < info.axis_count; ++i) {
    fm_chart_axis_info_t axis = FM_STRUCT_INIT(fm_chart_axis_info_t);
    if (!step(rc, fm_chart_get_axis_at(chart, i, &axis))) {
      break;
    }
    val a = fields_to_val(&axis, kAxisRd);
    a.set("position", val(axis.position));
    push(axes, a);
  }
  o.set("axes", axes);
  // The filter: the id of every filtered series across the plots, and the hidden categories.
  val filter = val::object();
  val hidden_series = val::array();
  for (size_t i = 0; i < info.plot_count && rc == 0; ++i) {
    fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
    if (!step(rc, fm_chart_get_plot_at(chart, i, &plot))) {
      break;
    }
    for (size_t j = 0; j < plot.series_count; ++j) {
      fm_chart_series_info_t series = FM_STRUCT_INIT(fm_chart_series_info_t);
      if (!step(rc, fm_chart_get_series_at(chart, i, j, &series))) {
        break;
      }
      if (series.filtered != 0) {
        push(hidden_series, val(series.series_id));
      }
    }
  }
  filter.set("hiddenSeries", hidden_series);
  filter.set("hiddenCategories", indices_to_val(info.hidden_category_count, rc, [&](size_t i, uint32_t* out) {
               return fm_chart_get_hidden_category_at(chart, i, out);
             }));
  o.set("filter", filter);

  val source = val::null();
  if (info.source_range != nullptr) {
    source = val::object();
    source.set("range", val(std::string(info.source_range)));
    source.set("direction", val(info.source_direction));
    source.set("headerRow", val(info.source_header_row != 0));
    source.set("headerColumn", val(info.source_header_column != 0));
  }
  o.set("source", source);
  return o;
}

/// A standalone Float64Array copy of `n` doubles.
val copy_doubles(const double* data, size_t n) {
  val out = val::global("Float64Array").new_(n);
  if (n != 0) {
    out.call<void>("set", val(emscripten::typed_memory_view(n, data)));
  }
  return out;
}

val range_to_val(const fm_sheet_rect_t& r) {
  val o = val::object();
  o.set("sheet", val(r.sheet_index));
  o.set("rect", cell_rect_to_val(r.rect));
  return o;
}

val data_series_to_val(const fm_chart_data_t* data, size_t index, fm_status_t& rc) {
  fm_chart_data_series_info_t s = FM_STRUCT_INIT(fm_chart_data_series_info_t);
  fm_chart_data_points_info_t pts = FM_STRUCT_INIT(fm_chart_data_points_info_t);
  if (!step(rc, fm_chart_data_get_series_at(data, index, &s)) ||
      !step(rc, fm_chart_data_get_series_points(data, index, &pts))) {
    return val::object();
  }
  val o = val::object();
  o.set("plotIndex", val(s.plot_index));
  o.set("seriesId", val(s.series_id));
  o.set("name", s.name != nullptr ? val(std::string(s.name)) : val::null());
  o.set("numberFormat", val(string_from_cstr(s.format_code)));
  o.set("xNumberFormat", val(string_from_cstr(s.x_format_code)));
  o.set("filtered", val(s.filtered != 0));
  o.set("values", copy_doubles(pts.values, pts.count));
  o.set("states", bytes_to_val(pts.states, pts.count));
  o.set("x", pts.x != nullptr ? copy_doubles(pts.x, pts.count) : val::null());
  o.set("sizes", pts.sizes != nullptr ? copy_doubles(pts.sizes, pts.count) : val::null());
  val ranges = val::object();
  for (int32_t role = 0; role < kRoleCount; ++role) {
    val list = val::array();
    for (size_t i = 0; i < s.range_count[role]; ++i) {
      fm_sheet_rect_t r{};
      if (!step(rc, fm_chart_data_get_range_at(data, index, role, i, &r))) {
        break;
      }
      push(list, range_to_val(r));
    }
    ranges.set(kRoleKeys[role], list);
  }
  o.set("ranges", ranges);
  return o;
}

val chart_data_to_val(const fm_chart_data_t* data, fm_status_t& rc) {
  fm_chart_data_info_t info = FM_STRUCT_INIT(fm_chart_data_info_t);
  if (!step(rc, fm_chart_data_get_info(data, &info))) {
    return val::object();
  }
  val o = val::object();
  o.set("domain", val(info.is_xy != 0 ? 1 : 0));
  o.set("pointCount", val(info.point_count));

  const bool numeric = info.category_value_count != 0;
  std::vector<double> numbers;
  val labels = val::array();
  for (size_t i = 0; i < info.category_label_count; ++i) {
    const char* label = nullptr;
    int32_t has_value = 0;
    double value = 0.0;
    if (!step(rc, fm_chart_data_get_category_at(data, i, &label, &has_value, &value))) {
      break;
    }
    push(labels, val(string_from_cstr(label)));
    // A text category among numeric ones reads as NaN in the values array.
    numbers.push_back(has_value != 0 ? value : std::numeric_limits<double>::quiet_NaN());
  }
  val levels = val::array();
  for (size_t level = 0; level < info.category_level_count && rc == 0; ++level) {
    size_t count = 0;
    val row = val::array();
    if (step(rc, fm_chart_data_get_category_level_count(data, level, &count))) {
      for (size_t i = 0; i < count && rc == 0; ++i) {
        const char* label = nullptr;
        if (step(rc, fm_chart_data_get_category_level_at(data, level, i, &label))) {
          push(row, val(string_from_cstr(label)));
        }
      }
    }
    push(levels, row);
  }
  val categories = val::object();
  categories.set("labels", labels);
  categories.set("levels", levels);
  categories.set("values", numeric ? copy_doubles(numbers.data(), numbers.size()) : val::null());
  categories.set("numberFormat", val(string_from_cstr(info.category_format_code)));
  categories.set("generated", val(info.categories_generated != 0));
  o.set("categories", categories);

  val series = val::array();
  for (size_t i = 0; i < info.series_count && rc == 0; ++i) {
    push(series, data_series_to_val(data, i, rc));
  }
  o.set("series", series);
  o.set("titleText", info.title_text != nullptr ? val(std::string(info.title_text)) : val::null());

  val axis_titles = val::array();
  for (size_t i = 0; i < info.axis_title_count; ++i) {
    fm_chart_axis_title_text_info_t t = FM_STRUCT_INIT(fm_chart_axis_title_text_info_t);
    if (!step(rc, fm_chart_data_get_axis_title_at(data, i, &t))) {
      break;
    }
    val item = val::object();
    item.set("axisId", val(t.axis_id));
    item.set("text", val(string_from_cstr(t.text)));
    push(axis_titles, item);
  }
  o.set("axisTitleTexts", axis_titles);
  val unresolved = val::array();
  for (size_t i = 0; i < info.unresolved_count; ++i) {
    fm_chart_unresolved_info_t u = FM_STRUCT_INIT(fm_chart_unresolved_info_t);
    if (!step(rc, fm_chart_data_get_unresolved_at(data, i, &u))) {
      break;
    }
    val item = val::object();
    item.set("seriesId", val(u.series_id));
    item.set("source", val(u.source));
    item.set("axisId", val(u.axis_id));
    push(unresolved, item);
  }
  o.set("unresolved", unresolved);
  return o;
}

val open_chart_data_val(fm_status_t rc, int32_t present, fm_chart_data_t* raw) {
  const std::unique_ptr<fm_chart_data_t, ChartDataDeleter> data(raw);
  if (!succeeded(rc)) {
    return val::undefined();
  }
  if (present == 0) {
    return val::null();
  }
  const val out = chart_data_to_val(data.get(), rc);
  if (!succeeded(rc)) {
    return val::undefined();
  }
  return out;
}
}  // namespace

// ---- Reads ----------------------------------------------------------

val JsWorkbook::getChart(uint32_t sheet, uint32_t objectId) const {
  int32_t present = 0;
  fm_chart_t* raw = nullptr;
  fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_chart(handle_, sheet, objectId, &present, &raw) : kBindingInvalidHandle;
  const std::unique_ptr<fm_chart_t, ChartDeleter> chart(raw);
  if (rc != 0 || present == 0) {
    return optional_or_null(rc, present, val::null());
  }
  // A failed sub-read inside the conversion replaces `rc`.
  const val info = chart_to_val(chart.get(), rc);
  return optional_or_null(rc, present, info);
}

val JsWorkbook::getChartData(uint32_t sheet, uint32_t objectId) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return val::undefined();
  }
  fm_chart_data_t* raw = nullptr;
  int32_t present = 0;
  const fm_status_t rc = fm_sheet_get_chart_data(handle_, sheet, objectId, &present, &raw);
  return open_chart_data_val(rc, present, raw);
}

val JsWorkbook::previewChartData(uint32_t sheet, uint32_t objectId, val overrides) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return val::undefined();
  }
  JsNarrowNumericReader reader("Chart.getPreview");
  const val spec = js_object_or_empty(overrides);
  std::string source_store;
  fm_chart_preview_overrides_spec_t o = FM_STRUCT_INIT(fm_chart_preview_overrides_spec_t);
  // A `null` or omitted override stays unset.
  const val source = reader.value(spec, "source", "source");
  if (reader.ok() && js_value_present(source)) {
    o.source_op = FM_PATCH_OP_SET;
    o.source = reader.optional_string_value(source, source_store, "source");
  }
  const auto set = [&](const char* key, fm_patch_op_t& op, int32_t& out, bool is_bool) {
    const val v = reader.value(spec, key, key);
    if (js_value_present(v)) {
      op = FM_PATCH_OP_SET;
      out = is_bool ? (reader.boolean_value(v, false) ? 1 : 0) : reader.i32_value(v, 0, key);
    }
  };
  set("direction", o.direction_op, o.direction, false);
  set("headerRow", o.header_row_op, o.header_row, true);
  set("headerColumn", o.header_column_op, o.header_column, true);
  set("blanksAs", o.blanks_as_op, o.blanks_as, false);
  set("plotVisibleOnly", o.plot_visible_only_op, o.plot_visible_only, true);
  set("naAsBlank", o.na_as_blank_op, o.na_as_blank, true);
  if (!reader.ok()) {
    return val::undefined();
  }
  fm_chart_data_t* raw = nullptr;
  int32_t present = 0;
  const fm_status_t rc = fm_sheet_get_chart_preview(handle_, sheet, objectId, &o, &present, &raw);
  return open_chart_data_val(rc, present, raw);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
