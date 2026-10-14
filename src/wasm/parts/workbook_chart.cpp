//
// JsWorkbook chart surface: the typed chart view, resolved series data,
// creation and the typed edits. Reads copy the handle-owned C data onto the
// JS heap and release the handle. Both directions are driven by field tables
// over the C structs; a patch field maps `undefined` to keep, `null` to clear
// and a value to set.

#include <emscripten/val.h>

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <memory>
#include <string>
#include <utility>
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

template <class T>
void store(void* base, size_t off, T v) {
  std::memcpy(static_cast<char*>(base) + off, &v, sizeof(T));
}

void put(val& o, const char* key, const val& v) {
  o.set(key, v);
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

/// `{ status, <field>: value }`.
val envelope(JsStatus status, const char* field, const val& value) {
  val o = val::object();
  put(o, "status", val(std::move(status)));
  put(o, field, value);
  return o;
}

/// `{ status, <field>: value | null }`; the value is dropped on failure.
val read_envelope(fm_status_t rc, const char* field, const val& value) {
  return envelope(status_from_rc(rc), field, rc == 0 ? value : val::null());
}

/// `{ status, <field>: id }` for the calls that create something; `id` is zero on failure.
val id_envelope(JsStatus status, const char* field, uint32_t id) {
  return envelope(std::move(status), field, val(id));
}

JsStatus argument_error(const JsNarrowNumericReader& reader) {
  return binding_error_status(kInvalidArgument, reader.message().c_str());
}

// ---- Reads ----------------------------------------------------------

enum class Rd : uint8_t { kI32, kU32, kF64, kBool, kStr, kStrOrNull, kNanNull, kOptI32, kOptU32, kTri, kObj, kObjOpt };

/// One output property: where the value sits in the C struct, the `has_` flag for the optional kinds.
struct RdField {
  const char* key;
  uint16_t off;
  uint16_t has;
  Rd kind;
  const RdField* sub;
};

#define RD(key, T, member, kind) {key, offsetof(T, member), 0, Rd::kind, nullptr}
#define RD_OPT(key, T, has, member, kind) {key, offsetof(T, member), offsetof(T, has), Rd::kind, nullptr}
#define RD_OBJ(key, T, member, table) {key, offsetof(T, member), 0, Rd::kObj, table}
#define RD_OBJ_OPT(key, T, has, member, table) {key, offsetof(T, member), offsetof(T, has), Rd::kObjOpt, table}
#define RD_END {nullptr, 0, 0, Rd::kI32, nullptr}

constexpr RdField kTitleRd[] = {
    RD("kind", fm_chart_title_t, kind, kI32),
    RD("text", fm_chart_title_t, text, kStr),
    RD("formula", fm_chart_title_t, formula, kStr),
    RD_END,
};

constexpr RdField kFillRd[] = {
    RD_OPT("color", fm_chart_fill_t, has_color, color, kOptU32),
    RD("transparency", fm_chart_fill_t, transparency_pct, kNanNull),
    RD_END,
};

constexpr RdField kLineRd[] = {
    RD_OPT("color", fm_chart_line_t, has_color, color, kOptU32),
    RD("transparency", fm_chart_line_t, transparency_pct, kNanNull),
    RD("widthPt", fm_chart_line_t, width_pt, kNanNull),
    RD_OPT("dash", fm_chart_line_t, has_dash, dash, kOptI32),
    RD("none", fm_chart_line_t, none, kBool),
    RD_END,
};

constexpr RdField kMarkerRd[] = {
    RD("symbol", fm_chart_marker_t, symbol, kI32),
    RD_OPT("size", fm_chart_marker_t, has_size, size, kOptU32),
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
    RD_OPT("fontColor", fm_chart_labels_t, has_font_color, font_color, kOptU32),
    RD("fontSizePt", fm_chart_labels_t, font_size_pt, kNanNull),
    RD("deleted", fm_chart_labels_t, deleted, kBool),
    RD_END,
};

constexpr RdField kPointRd[] = {
    RD("idx", fm_chart_point_t, idx, kU32),
    RD_OBJ("fill", fm_chart_point_t, fill, kFillRd),
    RD_OBJ("line", fm_chart_point_t, line, kLineRd),
    RD_OBJ_OPT("marker", fm_chart_point_t, has_marker, marker, kMarkerRd),
    RD_END,
};

constexpr RdField kSeriesRd[] = {
    RD("idx", fm_chart_series_t, idx, kU32),
    RD("order", fm_chart_series_t, order, kU32),
    RD("filtered", fm_chart_series_t, filtered, kBool),
    RD_OBJ("fill", fm_chart_series_t, fill, kFillRd),
    RD_OBJ("line", fm_chart_series_t, line, kLineRd),
    RD_OBJ_OPT("marker", fm_chart_series_t, has_marker, marker, kMarkerRd),
    RD("smooth", fm_chart_series_t, smooth, kTri),
    RD_OBJ_OPT("labels", fm_chart_series_t, has_labels, labels, kLabelsRd),
    RD_END,
};

constexpr RdField kPlotRd[] = {
    RD("type", fm_chart_plot_t, type, kI32),
    RD_OPT("grouping", fm_chart_plot_t, has_grouping, grouping, kOptI32),
    RD_OPT("gapWidth", fm_chart_plot_t, has_gap_width, gap_width, kOptU32),
    RD_OPT("overlap", fm_chart_plot_t, has_overlap, overlap, kOptI32),
    RD_OPT("holeSize", fm_chart_plot_t, has_hole_size, hole_size, kOptU32),
    RD_OPT("firstSliceAngle", fm_chart_plot_t, has_first_slice_angle, first_slice_angle, kOptU32),
    RD_END,
};

constexpr RdField kAxisRd[] = {
    RD("id", fm_chart_axis_t, id, kU32),
    RD("kind", fm_chart_axis_t, kind, kI32),
    RD("deleted", fm_chart_axis_t, deleted, kBool),
    RD("crossAxis", fm_chart_axis_t, cross_axis, kU32),
    RD("crosses", fm_chart_axis_t, crosses, kI32),
    RD("crossesAt", fm_chart_axis_t, crosses_at, kNanNull),
    RD("min", fm_chart_axis_t, min, kNanNull),
    RD("max", fm_chart_axis_t, max, kNanNull),
    RD("majorUnit", fm_chart_axis_t, major_unit, kNanNull),
    RD("minorUnit", fm_chart_axis_t, minor_unit, kNanNull),
    RD("logBase", fm_chart_axis_t, log_base, kNanNull),
    RD("reversed", fm_chart_axis_t, reversed, kBool),
    RD("numberFormat", fm_chart_axis_t, number_format, kStrOrNull),
    RD("sourceLinked", fm_chart_axis_t, source_linked, kBool),
    RD_OBJ("title", fm_chart_axis_t, title, kTitleRd),
    RD("majorGridlines", fm_chart_axis_t, major_gridlines, kBool),
    RD("minorGridlines", fm_chart_axis_t, minor_gridlines, kBool),
    RD("tickLabelPosition", fm_chart_axis_t, tick_label_pos, kStrOrNull),
    RD_END,
};

constexpr RdField kInfoRd[] = {
    RD("objectId", fm_chart_info_t, object_id, kU32),
    RD("name", fm_chart_info_t, name, kStr),
    RD("description", fm_chart_info_t, descr, kStr),
    RD("isChartex", fm_chart_info_t, is_chartex, kBool),
    RD("isPivotChart", fm_chart_info_t, is_pivot_chart, kBool),
    RD("onChartsheet", fm_chart_info_t, on_chartsheet, kBool),
    RD_OBJ("title", fm_chart_info_t, title, kTitleRd),
    RD("autoTitleDeleted", fm_chart_info_t, auto_title_deleted, kTri),
    RD("blanksAs", fm_chart_info_t, blanks_as, kI32),
    RD("plotVisibleOnly", fm_chart_info_t, plot_vis_only, kTri),
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
      case Rd::kNanNull:
        if (const double d = load<double>(base, f->off); !std::isnan(d)) {
          v = val(d);
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
        if (const int32_t tri = load<int32_t>(base, f->off); tri != FM_CHART_TRI_ABSENT) {
          v = val(tri == FM_CHART_TRI_TRUE);
        }
        break;
      case Rd::kObj:
        v = fields_to_val(static_cast<const char*>(base) + f->off, f->sub);
        break;
      case Rd::kObjOpt:
        if (load<int32_t>(base, f->has) != 0) {
          v = fields_to_val(static_cast<const char*>(base) + f->off, f->sub);
        }
        break;
    }
    put(o, f->key, v);
  }
  return o;
}

// The six series inputs, in `fm_chart_role_t` order.
constexpr const char* kRoleKeys[] = {"name", "categories", "values", "x", "y", "size"};
constexpr const char* kSourceKinds[] = {"none", "ref", "literal"};
constexpr const char* kAxisPositions[] = {"b", "t", "l", "r"};

val source_to_val(const fm_chart_t* chart, size_t plot, size_t series, int32_t role, const fm_chart_data_source_t& src,
                  fm_status_t& rc) {
  val o = val::object();
  put(o, "kind", val(std::string(kSourceKinds[src.kind])));
  put(o, "formula", val(string_from_cstr(src.formula)));
  val literal = val::array();
  for (size_t i = 0; i < src.literal_count; ++i) {
    fm_chart_literal_t lit{};
    if (!step(rc, fm_chart_series_literal_at(chart, plot, series, role, i, &lit))) {
      break;
    }
    push(literal, lit.is_text != 0 ? val(string_from_cstr(lit.text)) : val(lit.number));
  }
  put(o, "literal", literal);
  return o;
}

val series_to_val(const fm_chart_t* chart, size_t plot, size_t index, fm_status_t& rc) {
  fm_chart_series_t s{};
  if (!step(rc, fm_chart_series_at(chart, plot, index, &s))) {
    return val::object();
  }
  val o = fields_to_val(&s, kSeriesRd);
  const fm_chart_data_source_t* const sources[] = {&s.name, &s.categories, &s.values, &s.x, &s.y, &s.size};
  for (int32_t role = 0; role < 6; ++role) {
    put(o, kRoleKeys[role], source_to_val(chart, plot, index, role, *sources[role], rc));
  }
  val points = val::array();
  for (size_t i = 0; i < s.point_count; ++i) {
    fm_chart_point_t p{};
    if (!step(rc, fm_chart_series_point_at(chart, plot, index, i, &p))) {
      break;
    }
    push(points, fields_to_val(&p, kPointRd));
  }
  put(o, "points", points);
  return o;
}

val plot_to_val(const fm_chart_t* chart, size_t index, fm_status_t& rc) {
  fm_chart_plot_t p{};
  if (!step(rc, fm_chart_plot_at(chart, index, &p))) {
    return val::object();
  }
  val o = fields_to_val(&p, kPlotRd);
  const bool has_layout = p.chartex_layout != nullptr && p.chartex_layout[0] != '\0';
  put(o, "chartexLayout", has_layout ? val(std::string(p.chartex_layout)) : val::null());
  put(o, "barDirection",
      p.bar_dir == FM_CHART_BAR_DIR_COL   ? val(std::string("col"))
      : p.bar_dir == FM_CHART_BAR_DIR_BAR ? val(std::string("bar"))
                                          : val::null());
  put(o, "varyColors", val(p.vary_colors == FM_CHART_TRI_TRUE));
  val axis_ids = val::array();
  for (size_t i = 0; i < p.axis_id_count; ++i) {
    uint32_t id = 0;
    if (!step(rc, fm_chart_plot_axis_id_at(chart, index, i, &id))) {
      break;
    }
    push(axis_ids, val(id));
  }
  put(o, "axisIds", axis_ids);
  val series = val::array();
  for (size_t i = 0; i < p.series_count && rc == 0; ++i) {
    push(series, series_to_val(chart, index, i, rc));
  }
  put(o, "series", series);
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
  fm_chart_info_t info{};
  if (!step(rc, fm_chart_info(chart, &info))) {
    return val::object();
  }
  val o = fields_to_val(&info, kInfoRd);

  val legend = val::object();
  put(legend, "position", val(info.legend_position));
  put(legend, "overlay", val(info.legend_overlay != 0));
  put(legend, "hiddenSeries", indices_to_val(info.legend_hidden_count, rc, [&](size_t i, uint32_t* out) {
        return fm_chart_legend_hidden_at(chart, i, out);
      }));
  put(o, "legend", legend);

  val plots = val::array();
  for (size_t i = 0; i < info.plot_count && rc == 0; ++i) {
    push(plots, plot_to_val(chart, i, rc));
  }
  put(o, "plots", plots);
  val axes = val::array();
  for (size_t i = 0; i < info.axis_count; ++i) {
    fm_chart_axis_t axis{};
    if (!step(rc, fm_chart_axis_at(chart, i, &axis))) {
      break;
    }
    val a = fields_to_val(&axis, kAxisRd);
    put(a, "position", val(std::string(kAxisPositions[axis.position])));
    push(axes, a);
  }
  put(o, "axes", axes);
  put(o, "hiddenCategories", indices_to_val(info.hidden_category_count, rc, [&](size_t i, uint32_t* out) {
        return fm_chart_hidden_category_at(chart, i, out);
      }));

  val source = val::null();
  if (info.source_range != nullptr) {
    source = val::object();
    put(source, "range", val(std::string(info.source_range)));
    put(source, "direction", val(info.source_direction));
    put(source, "headerRow", val(info.source_header_row != 0));
    put(source, "headerColumn", val(info.source_header_column != 0));
  }
  put(o, "source", source);
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

val range_to_val(const fm_chart_sheet_range_t& r) {
  val o = val::object();
  put(o, "sheet", val(r.sheet));
  put(o, "top", val(r.top));
  put(o, "left", val(r.left));
  put(o, "bottom", val(r.bottom));
  put(o, "right", val(r.right));
  return o;
}

val data_series_to_val(const fm_chart_data_t* data, size_t index, fm_status_t& rc) {
  fm_chart_data_series_t s{};
  fm_chart_data_points_t pts{};
  if (!step(rc, fm_chart_data_series_at(data, index, &s)) ||
      !step(rc, fm_chart_data_series_points(data, index, &pts))) {
    return val::object();
  }
  val o = val::object();
  put(o, "plotIndex", val(s.plot_index));
  put(o, "idx", val(s.idx));
  put(o, "name", s.name != nullptr ? val(std::string(s.name)) : val::null());
  put(o, "formatCode", val(string_from_cstr(s.format_code)));
  put(o, "xFormatCode", val(string_from_cstr(s.x_format_code)));
  put(o, "filtered", val(s.filtered != 0));
  put(o, "values", copy_doubles(pts.values, pts.count));
  put(o, "states", bytes_to_val(pts.states, pts.count));
  put(o, "x", pts.x != nullptr ? copy_doubles(pts.x, pts.count) : val::null());
  put(o, "sizes", pts.sizes != nullptr ? copy_doubles(pts.sizes, pts.count) : val::null());
  val ranges = val::object();
  for (int32_t role = 0; role < 6; ++role) {
    val list = val::array();
    for (size_t i = 0; i < s.range_count[role]; ++i) {
      fm_chart_sheet_range_t r{};
      if (!step(rc, fm_chart_data_range_at(data, index, role, i, &r))) {
        break;
      }
      push(list, range_to_val(r));
    }
    put(ranges, kRoleKeys[role], list);
  }
  put(o, "ranges", ranges);
  return o;
}

val chart_data_to_val(const fm_chart_data_t* data, fm_status_t& rc) {
  fm_chart_data_info_t info{};
  if (!step(rc, fm_chart_data_info(data, &info))) {
    return val::object();
  }
  val o = val::object();
  put(o, "domain", val(std::string(info.is_xy != 0 ? "xy" : "category")));
  put(o, "pointCount", val(info.point_count));

  const bool numeric = info.category_value_count != 0;
  std::vector<double> numbers;
  val labels = val::array();
  for (size_t i = 0; i < info.category_label_count; ++i) {
    const char* label = nullptr;
    double value = 0.0;
    if (!step(rc, fm_chart_data_category_at(data, i, &label, &value))) {
      break;
    }
    push(labels, val(string_from_cstr(label)));
    numbers.push_back(value);
  }
  val levels = val::array();
  for (size_t level = 0; level < info.category_level_count && rc == 0; ++level) {
    size_t count = 0;
    val row = val::array();
    if (step(rc, fm_chart_data_category_level_size(data, level, &count))) {
      for (size_t i = 0; i < count && rc == 0; ++i) {
        const char* label = nullptr;
        if (step(rc, fm_chart_data_category_level_at(data, level, i, &label))) {
          push(row, val(string_from_cstr(label)));
        }
      }
    }
    push(levels, row);
  }
  val categories = val::object();
  put(categories, "labels", labels);
  put(categories, "levels", levels);
  put(categories, "values", numeric ? copy_doubles(numbers.data(), numbers.size()) : val::null());
  put(categories, "formatCode", val(string_from_cstr(info.category_format_code)));
  put(categories, "generated", val(info.categories_generated != 0));
  put(o, "categories", categories);

  val series = val::array();
  for (size_t i = 0; i < info.series_count && rc == 0; ++i) {
    push(series, data_series_to_val(data, i, rc));
  }
  put(o, "series", series);
  put(o, "titleText", info.title_text != nullptr ? val(std::string(info.title_text)) : val::null());

  val axis_titles = val::array();
  for (size_t i = 0; i < info.axis_title_count; ++i) {
    fm_chart_axis_title_text_t t{};
    if (!step(rc, fm_chart_data_axis_title_at(data, i, &t))) {
      break;
    }
    val item = val::object();
    put(item, "axisId", val(t.axis_id));
    put(item, "text", val(string_from_cstr(t.text)));
    push(axis_titles, item);
  }
  put(o, "axisTitleTexts", axis_titles);
  val unresolved = val::array();
  for (size_t i = 0; i < info.unresolved_count; ++i) {
    fm_chart_unresolved_t u{};
    if (!step(rc, fm_chart_data_unresolved_at(data, i, &u))) {
      break;
    }
    val item = val::object();
    put(item, "seriesIdx", val(u.series_idx));
    put(item, "source", val(string_from_cstr(u.source)));
    put(item, "axisId", val(u.axis_id));
    push(unresolved, item);
  }
  put(o, "unresolved", unresolved);
  return o;
}

val open_chart_data_val(fm_status_t rc, fm_chart_data_t* raw) {
  const std::unique_ptr<fm_chart_data_t, ChartDataDeleter> data(raw);
  return read_envelope(rc, "data", rc == 0 ? chart_data_to_val(data.get(), rc) : val::null());
}

// ---- Patch input ----------------------------------------------------

enum class Wr : uint8_t { kU32, kI32, kF64, kBool, kStr, kTitle, kRef, kPlainF64, kObj, kObjOp };

/// One input property: the `_op` and value slots in the C patch struct. A `kObj` nests without an op,
/// a `kObjOp` carries one.
struct WrField {
  const char* key;
  uint16_t op;
  uint16_t value;
  Wr kind;
  const WrField* sub;
};

#define WR(key, T, kind, op, value) {key, offsetof(T, op), offsetof(T, value), Wr::kind, nullptr}
#define WR_REF(key, T, member) {key, 0, offsetof(T, member), Wr::kRef, nullptr}
#define WR_PLAIN(key, T, member) {key, 0, offsetof(T, member), Wr::kPlainF64, nullptr}
#define WR_OBJ(key, T, member, table) {key, 0, offsetof(T, member), Wr::kObj, table}
#define WR_OBJ_OP(key, T, op, member, table) {key, offsetof(T, op), offsetof(T, member), Wr::kObjOp, table}
#define WR_END {nullptr, 0, 0, Wr::kU32, nullptr}

constexpr WrField kFillWr[] = {
    WR("color", fm_chart_fill_patch, kU32, color_op, color),
    WR("transparency", fm_chart_fill_patch, kF64, transparency_op, transparency_pct),
    WR_END,
};

constexpr WrField kLineWr[] = {
    WR("color", fm_chart_line_patch, kU32, color_op, color),
    WR("transparency", fm_chart_line_patch, kF64, transparency_op, transparency_pct),
    WR("widthPt", fm_chart_line_patch, kF64, width_op, width_pt),
    WR("dash", fm_chart_line_patch, kI32, dash_op, dash),
    WR("none", fm_chart_line_patch, kBool, none_op, none),
    WR_END,
};

constexpr WrField kMarkerWr[] = {
    WR("symbol", fm_chart_marker_patch, kI32, symbol_op, symbol),
    WR("size", fm_chart_marker_patch, kF64, size_op, size),
    WR_OBJ("fill", fm_chart_marker_patch, fill, kFillWr),
    WR_OBJ("line", fm_chart_marker_patch, line, kLineWr),
    WR_END,
};

constexpr WrField kLabelsWr[] = {
    WR("showValue", fm_chart_labels_patch, kBool, show_value_op, show_value),
    WR("showCategory", fm_chart_labels_patch, kBool, show_category_op, show_category),
    WR("showSeries", fm_chart_labels_patch, kBool, show_series_op, show_series),
    WR("showPercent", fm_chart_labels_patch, kBool, show_percent_op, show_percent),
    WR("showLegendKey", fm_chart_labels_patch, kBool, show_legend_key_op, show_legend_key),
    WR("separator", fm_chart_labels_patch, kStr, separator_op, separator),
    WR("position", fm_chart_labels_patch, kI32, position_op, position),
    WR("numberFormat", fm_chart_labels_patch, kStr, number_format_op, number_format),
    WR("fontColor", fm_chart_labels_patch, kU32, font_color_op, font_color),
    WR("fontSizePt", fm_chart_labels_patch, kF64, font_size_op, font_size_pt),
    WR_END,
};

constexpr WrField kSeriesWr[] = {
    WR_REF("name", fm_chart_series_patch, name),
    WR_REF("categories", fm_chart_series_patch, categories),
    WR_REF("values", fm_chart_series_patch, values),
    WR_REF("x", fm_chart_series_patch, x),
    WR_REF("y", fm_chart_series_patch, y),
    WR_REF("size", fm_chart_series_patch, size),
    WR_OBJ("fill", fm_chart_series_patch, fill, kFillWr),
    WR_OBJ("line", fm_chart_series_patch, line, kLineWr),
    WR_OBJ_OP("marker", fm_chart_series_patch, marker_op, marker, kMarkerWr),
    WR("smooth", fm_chart_series_patch, kBool, smooth_op, smooth),
    WR_OBJ_OP("labels", fm_chart_series_patch, labels_op, labels, kLabelsWr),
    WR_END,
};

constexpr WrField kPointWr[] = {
    WR_OBJ("fill", fm_chart_point_patch, fill, kFillWr),
    WR_OBJ("line", fm_chart_point_patch, line, kLineWr),
    WR_OBJ_OP("marker", fm_chart_point_patch, marker_op, marker, kMarkerWr),
    WR_END,
};

constexpr WrField kPlotWr[] = {
    WR("gapWidth", fm_chart_plot_patch, kU32, gap_width_op, gap_width),
    WR("overlap", fm_chart_plot_patch, kI32, overlap_op, overlap),
    WR("firstSliceAngle", fm_chart_plot_patch, kU32, first_slice_angle_op, first_slice_angle),
    WR("varyColors", fm_chart_plot_patch, kBool, vary_colors_op, vary_colors),
    WR("holeSize", fm_chart_plot_patch, kU32, hole_size_op, hole_size),
    WR_END,
};

constexpr WrField kAxisWr[] = {
    WR("visible", fm_chart_axis_patch, kBool, visible_op, visible),
    WR("title", fm_chart_axis_patch, kTitle, title_op, title),
    WR("min", fm_chart_axis_patch, kF64, min_op, min),
    WR("max", fm_chart_axis_patch, kF64, max_op, max),
    WR("majorUnit", fm_chart_axis_patch, kF64, major_unit_op, major_unit),
    WR("numberFormat", fm_chart_axis_patch, kStr, number_format_op, number_format),
    WR("majorGridlines", fm_chart_axis_patch, kBool, major_gridlines_op, major_gridlines),
    WR("reversed", fm_chart_axis_patch, kBool, reversed_op, reversed),
    WR("logBase", fm_chart_axis_patch, kF64, log_base_op, log_base),
    WR("crosses", fm_chart_axis_patch, kI32, crosses_op, crosses),
    WR_PLAIN("crossesAt", fm_chart_axis_patch, crosses_at),
    WR("tickLabelPosition", fm_chart_axis_patch, kStr, tick_label_pos_op, tick_label_pos),
    WR_END,
};

/// Keeps the bytes behind the `const char*` fields of a patch alive.
using Arena = std::deque<std::string>;

/// A title given as its text, or in the read-back shape `{ kind, text, formula }`. An out-of-domain (absent)
/// kind is left for the C ABI to reject.
fm_chart_title_t read_title(JsNarrowNumericReader& r, Arena& arena, const val& v, const char* field) {
  fm_chart_title_t t{};
  if (v.typeOf().strictlyEquals(val("object"))) {
    t.kind = r.i32(v, "kind", -1, "kind");
    arena.push_back(r.string(v, "text", "text"));
    t.text = arena.back().c_str();
    arena.push_back(r.string(v, "formula", "formula"));
    t.formula = arena.back().c_str();
  } else {
    t.kind = FM_CHART_TITLE_TEXT;
    arena.push_back(r.string_value(v, field));
    t.text = arena.back().c_str();
  }
  return t;
}

/// Fills the zeroed patch struct at `base` from the properties of `owner`.
void read_patch(JsNarrowNumericReader& r, Arena& arena, const val& owner, void* base, const WrField* f) {
  for (; f->key != nullptr && r.ok(); ++f) {
    const val v = r.value(owner, f->key, f->key);
    if (!r.ok() || v.isUndefined()) {
      continue;
    }
    const bool clear = v.isNull();
    const int32_t op = clear ? FM_CHART_PATCH_CLEAR : FM_CHART_PATCH_SET;
    switch (f->kind) {
      case Wr::kRef:
        // A reference is kept as NULL, removed as "" and replaced by a string.
        arena.push_back(clear ? std::string() : r.string_value(v, f->key));
        store(base, f->value, arena.back().c_str());
        break;
      case Wr::kPlainF64:
        if (!clear) {
          store(base, f->value, r.number_value(v, 0.0, f->key));
        }
        break;
      case Wr::kObj:
        if (!clear) {
          read_patch(r, arena, v, static_cast<char*>(base) + f->value, f->sub);
        }
        break;
      case Wr::kObjOp:
        store(base, f->op, op);
        if (!clear) {
          read_patch(r, arena, v, static_cast<char*>(base) + f->value, f->sub);
        }
        break;
      default:
        store(base, f->op, op);
        if (clear) {
          break;
        }
        switch (f->kind) {
          case Wr::kU32:
            store(base, f->value, r.u32_value(v, 0U, f->key));
            break;
          case Wr::kI32:
            store(base, f->value, r.i32_value(v, 0, f->key));
            break;
          case Wr::kF64:
            store(base, f->value, r.number_value(v, 0.0, f->key));
            break;
          case Wr::kBool:
            store(base, f->value, static_cast<int32_t>(r.boolean_value(v, false, f->key)));
            break;
          case Wr::kTitle:
            store(base, f->value, read_title(r, arena, v, f->key));
            break;
          default:
            arena.push_back(r.string_value(v, f->key));
            store(base, f->value, arena.back().c_str());
            break;
        }
        break;
    }
  }
}

/// Reads a `number[]`; absent reads as empty, a non-array is rejected.
std::vector<uint32_t> read_indices(JsNarrowNumericReader& r, const val& list, const char* field) {
  std::vector<uint32_t> out;
  if (!js_value_present(list)) {
    return out;
  }
  // `array_element` reports a non-array as a type error.
  const uint32_t n = r.is_array(list, field) ? r.length(list, field) : 1U;
  for (uint32_t i = 0; i < n && r.ok(); ++i) {
    out.push_back(r.u32_value(r.array_element(list, i, field), 0U, field));
  }
  return out;
}

/// A `boolean | null | undefined` header flag as an `fm_chart_header_mode_t`.
int32_t header_mode(JsNarrowNumericReader& r, const val& v, const char* field) {
  if (!js_value_present(v)) {
    return FM_CHART_HEADER_AUTO;
  }
  return r.boolean_value(v, false, field) ? FM_CHART_HEADER_YES : FM_CHART_HEADER_NO;
}

const uint32_t* data_or_null(const std::vector<uint32_t>& v) {
  return v.empty() ? nullptr : v.data();
}

/// Reads `patch` (absent reads as `{}`) into `base`; false after an argument rejection.
bool read_patch_arg(JsNarrowNumericReader& r, const val& patch, void* base, const WrField* table, Arena& arena) {
  read_patch(r, arena, js_value_present(patch) ? patch : val::object(), base, table);
  return r.ok();
}

}  // namespace

// ---- Reads ----------------------------------------------------------

val JsWorkbook::getChart(uint32_t sheet, uint32_t objectId) const {
  if (handle_ == nullptr) {
    return read_envelope(kBindingInvalidHandle, "chart", val::null());
  }
  fm_chart_t* raw = nullptr;
  fm_status_t rc = fm_sheet_chart_open(handle_, sheet, objectId, &raw);
  const std::unique_ptr<fm_chart_t, ChartDeleter> chart(raw);
  return read_envelope(rc, "chart", rc == 0 ? chart_to_val(chart.get(), rc) : val::null());
}

val JsWorkbook::getChartData(uint32_t sheet, uint32_t objectId) const {
  if (handle_ == nullptr) {
    return read_envelope(kBindingInvalidHandle, "data", val::null());
  }
  fm_chart_data_t* raw = nullptr;
  const fm_status_t rc = fm_sheet_chart_data_open(handle_, sheet, objectId, &raw);
  return open_chart_data_val(rc, raw);
}

val JsWorkbook::previewChartData(uint32_t sheet, uint32_t objectId, val overrides) const {
  if (handle_ == nullptr) {
    return read_envelope(kBindingInvalidHandle, "data", val::null());
  }
  JsNarrowNumericReader reader("previewChartData");
  const val spec = js_value_present(overrides) ? overrides : val::object();
  std::string source_store;
  fm_chart_preview_overrides o{};
  o.source = reader.optional_string(spec, "source", source_store, "source");
  // A `null` or omitted override stays unset.
  const auto set = [&](const char* key, int32_t& op, int32_t& out, bool is_bool) {
    const val v = reader.value(spec, key, key);
    if (js_value_present(v)) {
      op = FM_CHART_PATCH_SET;
      out = is_bool ? (reader.boolean_value(v, false, key) ? 1 : 0) : reader.i32_value(v, 0, key);
    }
  };
  set("direction", o.direction_op, o.direction, false);
  set("headerRow", o.header_row_op, o.header_row, true);
  set("headerColumn", o.header_column_op, o.header_column, true);
  set("blanksAs", o.blanks_as_op, o.blanks_as, false);
  set("plotVisibleOnly", o.plot_vis_only_op, o.plot_vis_only, true);
  if (!reader.ok()) {
    return envelope(argument_error(reader), "data", val::null());
  }
  fm_chart_data_t* raw = nullptr;
  const fm_status_t rc = fm_sheet_chart_preview_open(handle_, sheet, objectId, &o, &raw);
  return open_chart_data_val(rc, raw);
}

// ---- Creation and edits ---------------------------------------------

val JsWorkbook::insertChart(uint32_t sheet, val options) {
  if (handle_ == nullptr) {
    return id_envelope(error_status(kBindingInvalidHandle), "objectId", 0U);
  }
  if (!options.typeOf().strictlyEquals(val("object"))) {
    return id_envelope(binding_error_status(kInvalidArgument, "insertChart expects (sheet:number, options:object)"),
                       "objectId", 0U);
  }
  JsNarrowNumericReader reader("insertChart");
  std::string name_store;
  std::string descr_store;
  std::string title_store;
  fm_chart_insert opts{};
  opts.kind = reader.i32(options, "kind", FM_CHART_INSERT_COLUMN, "kind");
  opts.grouping = reader.i32(options, "grouping", FM_CHART_GROUPING_CLUSTERED, "grouping");
  const std::string source = reader.string(options, "source", "source");
  opts.source = source.c_str();
  opts.direction = reader.i32(options, "direction", FM_CHART_DIRECTION_AUTO, "direction");
  opts.header_row = header_mode(reader, reader.value(options, "headerRow", "headerRow"), "headerRow");
  opts.header_column = header_mode(reader, reader.value(options, "headerColumn", "headerColumn"), "headerColumn");
  opts.title = reader.optional_string(options, "title", title_store, "title");
  opts.legend = reader.i32(options, "legend", FM_CHART_LEGEND_BOTTOM, "legend");
  opts.placement.name = reader.optional_string(options, "name", name_store, "name");
  opts.placement.descr = reader.optional_string(options, "description", descr_store, "description");
  const val placement = reader.value(options, "placement", "placement");
  set_insert_placement(
      opts.placement,
      pull_drawing_anchor(reader, js_value_present(placement) ? placement : val::object(), FM_ANCHOR_KIND_TWO_CELL));
  if (!reader.ok()) {
    return id_envelope(argument_error(reader), "objectId", 0U);
  }
  uint32_t object_id = 0;
  const fm_status_t rc = fm_sheet_insert_chart(handle_, sheet, &opts, &object_id);
  return id_envelope(status_from_rc(rc), "objectId", rc == 0 ? object_id : 0U);
}

JsStatus JsWorkbook::setChartType(uint32_t sheet, uint32_t objectId, uint32_t plot, int32_t kind, val grouping) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setChartType");
  const int32_t group = reader.i32_value(grouping, FM_CHART_GROUPING_CLUSTERED, "grouping");
  if (!reader.ok()) {
    return argument_error(reader);
  }
  return status_from_rc(fm_sheet_set_chart_type(handle_, sheet, objectId, plot, kind, group));
}

JsStatus JsWorkbook::setChartTitle(uint32_t sheet, uint32_t objectId, val title) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setChartTitle");
  Arena arena;
  fm_chart_title_t t{};
  t.kind = FM_CHART_TITLE_NONE;
  if (js_value_present(title)) {
    t = read_title(reader, arena, title, "title");
  }
  if (!reader.ok()) {
    return argument_error(reader);
  }
  return status_from_rc(fm_sheet_set_chart_title(handle_, sheet, objectId, &t));
}

JsStatus JsWorkbook::setChartLegend(uint32_t sheet, uint32_t objectId, int32_t position, bool overlay,
                                    val hiddenSeries) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setChartLegend");
  const std::vector<uint32_t> hidden = read_indices(reader, hiddenSeries, "hiddenSeries");
  if (!reader.ok()) {
    return argument_error(reader);
  }
  fm_chart_legend_patch legend{};
  legend.position = position;
  legend.overlay = overlay ? 1 : 0;
  if (js_value_present(hiddenSeries)) {
    legend.hidden_series_op = FM_CHART_PATCH_SET;
    legend.hidden_series = data_or_null(hidden);
    legend.hidden_series_count = hidden.size();
  }
  return status_from_rc(fm_sheet_set_chart_legend(handle_, sheet, objectId, &legend));
}

JsStatus JsWorkbook::setChartAxis(uint32_t sheet, uint32_t objectId, uint32_t axisId, val patch) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setChartAxis");
  Arena arena;
  fm_chart_axis_patch p{};
  if (!read_patch_arg(reader, patch, &p, kAxisWr, arena)) {
    return argument_error(reader);
  }
  return status_from_rc(fm_sheet_set_chart_axis(handle_, sheet, objectId, axisId, &p));
}

JsStatus JsWorkbook::setChartDisplay(uint32_t sheet, uint32_t objectId, int32_t blanksAs, val plotVisibleOnly) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setChartDisplay");
  // `null` or omitted removes `c:plotVisOnly`.
  int32_t plot_vis_only = FM_CHART_TRI_ABSENT;
  if (js_value_present(plotVisibleOnly)) {
    plot_vis_only =
        reader.boolean_value(plotVisibleOnly, false, "plotVisibleOnly") ? FM_CHART_TRI_TRUE : FM_CHART_TRI_FALSE;
  }
  return status_from_rc(fm_sheet_set_chart_display(handle_, sheet, objectId, blanksAs, plot_vis_only));
}

JsStatus JsWorkbook::setChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesIdx, val patch) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setChartSeries");
  Arena arena;
  fm_chart_series_patch p{};
  if (!read_patch_arg(reader, patch, &p, kSeriesWr, arena)) {
    return argument_error(reader);
  }
  return status_from_rc(fm_sheet_set_chart_series(handle_, sheet, objectId, plot, seriesIdx, &p));
}

JsStatus JsWorkbook::setChartPoint(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesIdx,
                                   uint32_t pointIdx, val patch) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setChartPoint");
  Arena arena;
  fm_chart_point_patch p{};
  // A null or empty patch removes the point's formatting.
  if (!read_patch_arg(reader, patch, &p, kPointWr, arena)) {
    return argument_error(reader);
  }
  return status_from_rc(fm_sheet_set_chart_point(handle_, sheet, objectId, plot, seriesIdx, pointIdx, &p));
}

JsStatus JsWorkbook::setChartPlot(uint32_t sheet, uint32_t objectId, uint32_t plot, val patch) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setChartPlot");
  Arena arena;
  fm_chart_plot_patch p{};
  if (!read_patch_arg(reader, patch, &p, kPlotWr, arena)) {
    return argument_error(reader);
  }
  return status_from_rc(fm_sheet_set_chart_plot(handle_, sheet, objectId, plot, &p));
}

JsStatus JsWorkbook::setChartFilter(uint32_t sheet, uint32_t objectId, val hiddenSeries, val hiddenCategories) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setChartFilter");
  const std::vector<uint32_t> series = read_indices(reader, hiddenSeries, "hiddenSeries");
  const std::vector<uint32_t> categories = read_indices(reader, hiddenCategories, "hiddenCategories");
  if (!reader.ok()) {
    return argument_error(reader);
  }
  return status_from_rc(fm_sheet_set_chart_filter(handle_, sheet, objectId, data_or_null(series), series.size(),
                                                  data_or_null(categories), categories.size()));
}

val JsWorkbook::addChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, val patch) {
  if (handle_ == nullptr) {
    return id_envelope(error_status(kBindingInvalidHandle), "seriesIdx", 0U);
  }
  JsNarrowNumericReader reader("addChartSeries");
  Arena arena;
  fm_chart_series_patch p{};
  if (!read_patch_arg(reader, patch, &p, kSeriesWr, arena)) {
    return id_envelope(argument_error(reader), "seriesIdx", 0U);
  }
  uint32_t series_idx = 0;
  const fm_status_t rc = fm_sheet_add_chart_series(handle_, sheet, objectId, plot, &p, &series_idx);
  return id_envelope(status_from_rc(rc), "seriesIdx", rc == 0 ? series_idx : 0U);
}

JsStatus JsWorkbook::removeChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesIdx) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  return status_from_rc(fm_sheet_remove_chart_series(handle_, sheet, objectId, plot, seriesIdx));
}

JsStatus JsWorkbook::setChartSource(uint32_t sheet, uint32_t objectId, const std::string& source, int32_t direction,
                                    val headerRow, val headerColumn) {
  if (handle_ == nullptr) {
    return error_status(kBindingInvalidHandle);
  }
  JsNarrowNumericReader reader("setChartSource");
  const int32_t row = header_mode(reader, headerRow, "headerRow");
  const int32_t column = header_mode(reader, headerColumn, "headerColumn");
  if (!reader.ok()) {
    return argument_error(reader);
  }
  return status_from_rc(fm_sheet_set_chart_source(handle_, sheet, objectId, source.c_str(), direction, row, column));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
