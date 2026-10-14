//
// Chart oracle support: the one reader of chart parts on the oracle side.
//
// `chart_model_dump <file>` prints every chart of an .xlsx / .xlsb as JSON
// built from `get_chart`, so the Python capture never parses chart XML
// itself. `--view` selects the projections a track compares:
//
//   shape   plot types, bar direction, grouping and series count
//   refs    every reference site (series sources, title)
//   cache   ptCount / pt idx / value / formatCode of each series source
//   format  the editable formatting surface (title, legend, axes, plot
//           options, per-series fill / line / marker / labels / points)
//
// `chart_model_dump --build <cases.json> --case <id> --out <file.xlsx>`
// writes Formulon's workbook for one chart-oracle case spec, which the
// capture opens in Excel. `formulon_oracle_chart_tests` compiles this file
// with FORMULON_CHART_MODEL_DUMP_NO_MAIN and calls the same builder and
// projections, so the verifier and the capture agree on both by
// construction.
//
// Absent elements whose Excel reading is measured are normalised in the
// format view: dispBlanksAs reads as zero, plotVisOnly and
// autoTitleDeleted as true, an axis's crosses as autoZero.

#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <string>
#include <utility>
#include <vector>

#include "chart/chart_edit.h"
#include "chart/chart_model.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "sheet_layout_edit.h"
#include "styles.h"
#include "tests/oracle/json_reader.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {
namespace tests {
namespace oracle {
namespace chart_oracle {

namespace {

Error invalid(std::string message) {
  return make_error(FormulonErrorCode::kInvalidArgument, std::move(message));
}

// ---------------------------------------------------------------------------
// JSON output
// ---------------------------------------------------------------------------

void put_string(std::string& out, const std::string& s) {
  out += '"';
  for (const char ch : s) {
    const unsigned char c = static_cast<unsigned char>(ch);
    switch (ch) {
      case '"':
        out += "\\\"";
        break;
      case '\\':
        out += "\\\\";
        break;
      case '\n':
        out += "\\n";
        break;
      case '\r':
        out += "\\r";
        break;
      case '\t':
        out += "\\t";
        break;
      default:
        if (c < 0x20) {
          char buf[8];
          std::snprintf(buf, sizeof buf, "\\u%04x", c);
          out += buf;
        } else {
          out += ch;
        }
    }
  }
  out += '"';
}

void put_number(std::string& out, double v) {
  if (std::isnan(v) || std::isinf(v)) {
    out += "null";
    return;
  }
  char buf[40];
  if (v == std::floor(v) && std::fabs(v) < 1e15) {
    std::snprintf(buf, sizeof buf, "%.0f", v);
  } else {
    std::snprintf(buf, sizeof buf, "%.17g", v);
  }
  out += buf;
}

/// Comma-separated object / array writer.
class Json {
 public:
  explicit Json(std::string& out) : out_(out) {}
  void open(char c) {
    sep();
    out_ += c;
    first_ = true;
  }
  void close(char c) {
    out_ += c;
    first_ = false;
  }
  void key(const char* k) {
    sep();
    put_string(out_, k);
    out_ += ':';
    first_ = true;
  }
  void str(const std::string& s) {
    sep();
    put_string(out_, s);
  }
  void num(double v) {
    sep();
    put_number(out_, v);
  }
  void boolean(bool b) {
    sep();
    out_ += b ? "true" : "false";
  }
  void null() {
    sep();
    out_ += "null";
  }
  void opt_str(const std::string& s) { s.empty() ? null() : str(s); }
  template <typename T>
  void opt_num(const std::optional<T>& v) {
    v ? num(static_cast<double>(*v)) : null();
  }
  void tri(Tri t) { t == Tri::kAbsent ? null() : boolean(t == Tri::kTrue); }
  void color(const std::optional<std::uint32_t>& c) {
    if (!c) {
      null();
      return;
    }
    char buf[12];
    std::snprintf(buf, sizeof buf, "%08X", static_cast<unsigned>(*c));
    str(buf);
  }

 private:
  void sep() {
    if (!first_) {
      out_ += ',';
    }
    first_ = false;
  }
  std::string& out_;
  bool first_ = true;
};

const char* plot_type_name(ChartPlotType t) {
  static const char* const kNames[] = {"bar",   "bar3D", "line",     "line3D",    "area",    "area3D",
                                       "pie",   "pie3D", "doughnut", "ofPie",     "scatter", "bubble",
                                       "radar", "stock", "surface",  "surface3D", "chartex"};
  const auto i = static_cast<std::size_t>(t);
  return i < sizeof kNames / sizeof kNames[0] ? kNames[i] : "unknown";
}

const char* grouping_name(ChartGrouping g) {
  static const char* const kNames[] = {"clustered", "stacked", "percentStacked", "standard"};
  return kNames[static_cast<std::size_t>(g)];
}

const char* bar_dir_name(ChartBarDirection d) {
  return d == ChartBarDirection::kCol ? "col" : d == ChartBarDirection::kBar ? "bar" : nullptr;
}

const char* title_kind_name(ChartTitleKind k) {
  static const char* const kNames[] = {"none", "text", "ref", "auto"};
  return kNames[static_cast<std::size_t>(k)];
}

const char* legend_pos_name(ChartLegendPosition p) {
  static const char* const kNames[] = {"none", "b", "t", "l", "r", "tr"};
  return kNames[static_cast<std::size_t>(p)];
}

const char* dash_name(ChartLineDash d) {
  static const char* const kNames[] = {"solid", "dash", "dot", "dashDot", "other"};
  return kNames[static_cast<std::size_t>(d)];
}

const char* marker_name(ChartMarkerSymbol s) {
  static const char* const kNames[] = {"none", "auto", "circle", "square", "diamond", "triangle", "other"};
  return kNames[static_cast<std::size_t>(s)];
}

const char* label_pos_name(ChartLabelPosition p) {
  static const char* const kNames[] = {"bestFit", "ctr", "inBase", "inEnd", "outEnd", "t", "b", "l", "r"};
  return kNames[static_cast<std::size_t>(p)];
}

const char* axis_kind_name(ChartAxisKind k) {
  static const char* const kNames[] = {"cat", "val", "date", "ser"};
  return kNames[static_cast<std::size_t>(k)];
}

const char* axis_pos_name(ChartAxisPosition p) {
  static const char* const kNames[] = {"b", "t", "l", "r"};
  return kNames[static_cast<std::size_t>(p)];
}

const char* crosses_name(ChartAxisCrosses c) {
  static const char* const kNames[] = {nullptr, "autoZero", "min", "max", "at"};
  return kNames[static_cast<std::size_t>(c)];
}

const char* blanks_name(ChartBlanksAs b) {
  static const char* const kNames[] = {"zero", "gap", "zero", "span"};
  return kNames[static_cast<std::size_t>(b)];
}

void put_cstr(Json& j, const char* s) {
  s == nullptr ? j.null() : j.str(s);
}

// -- shape -------------------------------------------------------------------

void write_shape(Json& j, const ChartModel& m) {
  j.open('[');
  for (const ChartPlot& p : m.plots) {
    j.open('{');
    j.key("type");
    j.str(plot_type_name(p.type));
    j.key("bar_dir");
    put_cstr(j, bar_dir_name(p.bar_dir));
    j.key("grouping");
    p.grouping ? j.str(grouping_name(*p.grouping)) : j.null();
    j.key("series");
    j.num(static_cast<double>(p.series.size()));
    j.close('}');
  }
  j.close(']');
}

// -- refs --------------------------------------------------------------------

void put_formula(Json& j, const ChartDataSource& s) {
  s.kind == ChartDataSource::Kind::kRef ? j.str(s.formula) : j.null();
}

void write_refs(Json& j, const ChartModel& m) {
  j.open('{');
  j.key("series");
  j.open('[');
  for (std::size_t pi = 0; pi < m.plots.size(); ++pi) {
    for (const ChartSeries& s : m.plots[pi].series) {
      j.open('{');
      j.key("plot");
      j.num(static_cast<double>(pi));
      j.key("idx");
      j.num(s.idx);
      j.key("filtered");
      j.boolean(s.filtered);
      const std::pair<const char*, const ChartDataSource*> sources[] = {
          {"name", &s.name}, {"cat", &s.cat}, {"val", &s.val}, {"x", &s.x}, {"y", &s.y}, {"size", &s.size}};
      for (const auto& [key, src] : sources) {
        if (src->kind != ChartDataSource::Kind::kNone) {
          j.key(key);
          put_formula(j, *src);
        }
      }
      j.close('}');
    }
  }
  j.close(']');
  j.key("title");
  m.title.kind == ChartTitleKind::kRef ? j.str(m.title.ref_formula) : j.null();
  j.key("axis_titles");
  j.open('[');
  for (const ChartAxis& a : m.axes) {
    if (a.title.kind == ChartTitleKind::kRef) {
      j.str(a.title.ref_formula);
    }
  }
  j.close(']');
  j.close('}');
}

// -- cache -------------------------------------------------------------------

void write_points(Json& j, const std::vector<ChartValue>& pts) {
  j.open('[');
  for (const ChartValue& v : pts) {
    j.open('[');
    j.num(v.idx);
    switch (v.kind) {
      case ChartValue::Kind::kNumber:
        j.str("n");
        j.num(v.number);
        break;
      case ChartValue::Kind::kText:
        j.str("s");
        j.str(v.text);
        break;
      case ChartValue::Kind::kBlank:
        j.str("b");
        j.null();
        break;
    }
    j.opt_str(v.format_code);
    j.close(']');
  }
  j.close(']');
}

void write_source_cache(Json& j, const ChartDataSource& s) {
  if (s.kind == ChartDataSource::Kind::kNone) {
    j.null();
    return;
  }
  static const char* const kElement[] = {nullptr, "num", "str", "multi", "cx"};
  const bool ref = s.kind == ChartDataSource::Kind::kRef;
  j.open('{');
  j.key("type");
  j.str(ref ? "ref" : "literal");
  j.key("element");
  put_cstr(j, kElement[static_cast<std::size_t>(s.ref_element)]);
  j.key("count");
  j.num(s.point_count);
  j.key("format");
  j.opt_str(s.cache_format_code);
  j.key("pts");
  write_points(j, ref ? s.cache : s.literal);
  j.key("levels");
  j.open('[');
  for (const std::vector<ChartValue>& level : s.cache_levels) {
    write_points(j, level);
  }
  j.close(']');
  j.close('}');
}

void write_cache(Json& j, const ChartModel& m) {
  j.open('[');
  for (std::size_t pi = 0; pi < m.plots.size(); ++pi) {
    for (const ChartSeries& s : m.plots[pi].series) {
      j.open('{');
      j.key("plot");
      j.num(static_cast<double>(pi));
      j.key("idx");
      j.num(s.idx);
      const std::pair<const char*, const ChartDataSource*> sources[] = {
          {"name", &s.name}, {"cat", &s.cat}, {"val", &s.val}, {"x", &s.x}, {"y", &s.y}, {"size", &s.size}};
      for (const auto& [key, src] : sources) {
        if (src->kind != ChartDataSource::Kind::kNone) {
          j.key(key);
          write_source_cache(j, *src);
        }
      }
      j.close('}');
    }
  }
  j.close(']');
}

// -- format ------------------------------------------------------------------

void write_fill(Json& j, const ChartFill& f) {
  j.open('{');
  j.key("color");
  j.color(f.color);
  j.key("alpha");
  j.opt_num(f.alpha);
  j.close('}');
}

void write_line(Json& j, const ChartLine& l) {
  j.open('{');
  j.key("none");
  j.boolean(l.none);
  j.key("color");
  j.color(l.color);
  j.key("alpha");
  j.opt_num(l.alpha);
  j.key("width");
  j.opt_num(l.width_emu);
  j.key("dash");
  l.dash ? j.str(*l.dash == ChartLineDash::kOther ? l.dash_name : dash_name(*l.dash)) : j.null();
  j.close('}');
}

void write_marker(Json& j, const std::optional<ChartMarker>& m) {
  if (!m) {
    j.null();
    return;
  }
  j.open('{');
  j.key("symbol");
  j.str(m->symbol == ChartMarkerSymbol::kOther ? m->symbol_name : marker_name(m->symbol));
  j.key("size");
  j.opt_num(m->size);
  j.key("fill");
  write_fill(j, m->fill);
  j.key("line");
  write_line(j, m->line);
  j.close('}');
}

void write_labels(Json& j, const std::optional<ChartLabels>& l) {
  if (!l) {
    j.null();
    return;
  }
  j.open('{');
  j.key("deleted");
  j.boolean(l->deleted);
  j.key("val");
  j.boolean(l->show_val);
  j.key("cat");
  j.boolean(l->show_cat);
  j.key("ser");
  j.boolean(l->show_ser);
  j.key("pct");
  j.boolean(l->show_pct);
  j.key("legend_key");
  j.boolean(l->show_legend_key);
  j.key("separator");
  l->separator ? j.str(*l->separator) : j.null();
  j.key("position");
  l->position ? j.str(label_pos_name(*l->position)) : j.null();
  j.key("num_fmt");
  l->num_fmt ? j.str(l->num_fmt->code) : j.null();
  j.key("font_color");
  j.color(l->font_color);
  j.key("font_size");
  j.opt_num(l->font_size_hundredths_pt);
  j.close('}');
}

void write_title(Json& j, const ChartTitle& t) {
  j.open('{');
  j.key("kind");
  j.str(title_kind_name(t.kind));
  j.key("text");
  t.kind == ChartTitleKind::kText ? j.str(t.text) : j.null();
  j.key("ref");
  t.kind == ChartTitleKind::kRef ? j.str(t.ref_formula) : j.null();
  j.close('}');
}

void write_format(Json& j, const ChartModel& m) {
  j.open('{');
  j.key("title");
  write_title(j, m.title);
  j.key("auto_title_deleted");
  j.boolean(m.auto_title_deleted != Tri::kFalse);
  j.key("legend");
  if (!m.legend.present) {
    j.null();
  } else {
    j.open('{');
    j.key("position");
    j.str(legend_pos_name(m.legend.position));
    j.key("overlay");
    j.boolean(m.legend.overlay);
    j.key("hidden");
    j.open('[');
    for (const std::uint32_t idx : m.legend.hidden_series_idx) {
      j.num(idx);
    }
    j.close(']');
    j.close('}');
  }
  j.key("blanks_as");
  j.str(blanks_name(m.disp_blanks_as));
  j.key("plot_vis_only");
  j.boolean(m.plot_vis_only != Tri::kFalse);
  j.key("category_filter");
  j.open('[');
  for (const std::uint32_t idx : m.category_filter) {
    j.num(idx);
  }
  j.close(']');
  j.key("plots");
  j.open('[');
  for (const ChartPlot& p : m.plots) {
    j.open('{');
    j.key("type");
    j.str(plot_type_name(p.type));
    j.key("bar_dir");
    put_cstr(j, bar_dir_name(p.bar_dir));
    j.key("grouping");
    p.grouping ? j.str(grouping_name(*p.grouping)) : j.null();
    j.key("vary_colors");
    j.tri(p.vary_colors);
    j.key("gap_width");
    j.opt_num(p.gap_width);
    j.key("overlap");
    j.opt_num(p.overlap);
    j.key("hole_size");
    j.opt_num(p.hole_size);
    j.key("first_slice_angle");
    j.opt_num(p.first_slice_angle);
    j.key("labels");
    write_labels(j, p.labels);
    j.key("series");
    j.open('[');
    for (const ChartSeries& s : p.series) {
      j.open('{');
      j.key("idx");
      j.num(s.idx);
      j.key("filtered");
      j.boolean(s.filtered);
      j.key("fill");
      write_fill(j, s.fill);
      j.key("line");
      write_line(j, s.line);
      j.key("marker");
      write_marker(j, s.marker);
      j.key("smooth");
      j.tri(s.smooth);
      j.key("labels");
      write_labels(j, s.labels);
      j.key("points");
      j.open('[');
      for (const ChartPoint& pt : s.points) {
        j.open('{');
        j.key("idx");
        j.num(pt.idx);
        j.key("fill");
        write_fill(j, pt.fill);
        j.key("line");
        write_line(j, pt.line);
        j.key("marker");
        write_marker(j, pt.marker);
        j.close('}');
      }
      j.close(']');
      j.close('}');
    }
    j.close(']');
    j.close('}');
  }
  j.close(']');
  j.key("axes");
  j.open('[');
  for (const ChartAxis& a : m.axes) {
    j.open('{');
    j.key("kind");
    j.str(axis_kind_name(a.kind));
    j.key("position");
    j.str(axis_pos_name(a.position));
    j.key("deleted");
    j.boolean(a.deleted);
    j.key("crosses");
    put_cstr(j, crosses_name(a.crosses == ChartAxisCrosses::kAbsent ? ChartAxisCrosses::kAutoZero : a.crosses));
    j.key("crosses_at");
    j.num(a.crosses_at);
    j.key("reversed");
    j.boolean(a.reversed);
    j.key("log_base");
    j.num(a.log_base);
    j.key("min");
    j.num(a.min);
    j.key("max");
    j.num(a.max);
    j.key("major_unit");
    j.num(a.major_unit);
    j.key("num_fmt");
    a.num_fmt ? j.str(a.num_fmt->code) : j.null();
    j.key("title");
    write_title(j, a.title);
    j.key("major_gridlines");
    j.boolean(a.major_gridlines);
    j.key("tick_label_pos");
    j.opt_str(a.tick_label_pos);
    j.close('}');
  }
  j.close(']');
  j.close('}');
}

bool has_view(const std::vector<std::string>& views, const char* name) {
  for (const std::string& v : views) {
    if (v == name || v == "all") {
      return true;
    }
  }
  return false;
}

// ---------------------------------------------------------------------------
// Case builder
// ---------------------------------------------------------------------------

const JsonValue* member(const JsonValue& obj, const char* key) {
  return obj.is_object() ? obj.find(key) : nullptr;
}

std::string str_or(const JsonValue& obj, const char* key, const std::string& fallback) {
  const JsonValue* v = member(obj, key);
  return v != nullptr && v->is_string() ? v->as_string() : fallback;
}

bool bool_or(const JsonValue& obj, const char* key, bool fallback) {
  const JsonValue* v = member(obj, key);
  return v != nullptr && v->is_bool() ? v->as_bool() : fallback;
}

Expected<std::uint32_t, Error> uint_at(const JsonValue& obj, const char* key) {
  const JsonValue* v = member(obj, key);
  if (v == nullptr || !v->is_number() || v->as_number() < 0) {
    return invalid(std::string("expected a non-negative number at '") + key + "'");
  }
  return static_cast<std::uint32_t>(v->as_number());
}

Expected<std::uint32_t, Error> column_index(const std::string& letters) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  if (!a1::parse_a1_ref(letters + "1", &row, &col)) {
    return invalid("bad column letters '" + letters + "'");
  }
  return col;
}

Expected<std::size_t, Error> sheet_index(const Workbook& wb, const std::string& name) {
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (wb.sheet(i).name() == name) {
      return i;
    }
  }
  return invalid("unknown sheet '" + name + "'");
}

/// xf index carrying number format `code` (or built-in `id` when given).
std::uint32_t xf_for_format(Workbook& wb, const std::string& code, const JsonValue* id_v) {
  StylesTable& styles = wb.mutable_styles();
  std::uint16_t id = 0;
  bool found = false;
  if (id_v != nullptr && id_v->is_number()) {
    id = static_cast<std::uint16_t>(id_v->as_number());
    found = true;
  }
  for (std::uint16_t i = 0; i < 164 && !found; ++i) {
    if (code == builtin_num_fmt(i)) {
      id = i;
      found = true;
    }
  }
  for (std::size_t i = 0; i < styles.num_fmts.size() && !found; ++i) {
    if (styles.num_fmt_strings[styles.num_fmts[i].format_string_index] == code) {
      id = styles.num_fmts[i].id;
      found = true;
    }
  }
  if (!found) {
    std::uint16_t next = 164;
    for (const NumFmtRecord& r : styles.num_fmts) {
      next = std::max<std::uint16_t>(next, static_cast<std::uint16_t>(r.id + 1));
    }
    styles.num_fmt_strings.push_back(code);
    styles.num_fmts.push_back(NumFmtRecord{next, static_cast<std::uint32_t>(styles.num_fmt_strings.size() - 1U)});
    id = next;
  }
  for (std::size_t i = 0; i < styles.cell_xfs.size(); ++i) {
    if (styles.cell_xfs[i].num_fmt_id == id && (id == 0 || styles.cell_xfs[i].apply_number_format)) {
      return static_cast<std::uint32_t>(i);
    }
  }
  CellXf xf = styles.cell_xfs.empty() ? CellXf{} : styles.cell_xfs[0];
  xf.num_fmt_id = id;
  xf.apply_number_format = true;
  styles.cell_xfs.push_back(xf);
  return static_cast<std::uint32_t>(styles.cell_xfs.size() - 1U);
}

Expected<void, Error> write_cell(Workbook& wb, std::size_t sheet, std::uint32_t row, std::uint32_t col,
                                 const JsonValue& rec, bool* wrote_formula) {
  const JsonValue* value = &rec;
  if (rec.is_object()) {
    value = member(rec, "value");
    const JsonValue* fmt = member(rec, "format");
    if (fmt != nullptr && fmt->is_string()) {
      RETURN_IF_ERROR(
          wb.set_cell_xf_index(sheet, row, col, xf_for_format(wb, fmt->as_string(), member(rec, "format_id"))));
    }
    if (value == nullptr) {
      return {};
    }
  }
  Sheet& sh = wb.sheet(sheet);
  if (value->is_number()) {
    sh.set_cell_value(row, col, Value::number(value->as_number()));
  } else if (value->is_bool()) {
    sh.set_cell_value(row, col, Value::boolean(value->as_bool()));
  } else if (value->is_string()) {
    const std::string& s = value->as_string();
    if (!s.empty() && s[0] == '=') {
      RETURN_IF_ERROR(wb.set_cell_formula(sheet, row, col, s));
      *wrote_formula = true;
    } else {
      sh.set_cell_value(row, col, Value::text(wb.intern_text(s)));
    }
  } else if (!value->is_null()) {
    return invalid("unsupported cell value");
  }
  return {};
}

Expected<void, Error> build_sheets(Workbook& wb, const JsonValue& spec, bool* wrote_formula) {
  const JsonValue* sheets = member(spec, "sheets");
  if (sheets == nullptr || !sheets->is_array() || sheets->as_array().empty()) {
    return invalid("spec needs a non-empty 'sheets' list");
  }
  for (const JsonValue& sh : sheets->as_array()) {
    const std::size_t index = wb.add_sheet(str_or(sh, "name", ""));
    const JsonValue* cells = member(sh, "cells");
    if (cells == nullptr) {
      continue;
    }
    for (const auto& [addr, rec] : cells->as_object()) {
      std::uint32_t row = 0;
      std::uint32_t col = 0;
      if (!a1::parse_a1_ref(addr, &row, &col)) {
        return invalid("bad cell address '" + addr + "'");
      }
      RETURN_IF_ERROR(write_cell(wb, index, row, col, rec, wrote_formula));
    }
  }
  return {};
}

/// Hidden rows / columns / sheets; applied after the charts as the capture does.
Expected<void, Error> apply_visibility(Workbook& wb, const JsonValue& spec) {
  for (const JsonValue& sh : member(spec, "sheets")->as_array()) {
    ASSIGN_OR_RETURN(const std::size_t index, sheet_index(wb, str_or(sh, "name", "")));
    SheetLayout& layout = wb.sheet(index).mutable_layout();
    if (const JsonValue* rows = member(sh, "hidden_rows"); rows != nullptr) {
      for (const JsonValue& r : rows->as_array()) {
        upsert_row_override(layout, static_cast<std::uint32_t>(r.as_number()) - 1U).hidden = true;
      }
      wb.mark_row_visibility_dependents_dirty();
    }
    if (const JsonValue* cols = member(sh, "hidden_cols"); cols != nullptr) {
      for (const JsonValue& c : cols->as_array()) {
        ASSIGN_OR_RETURN(const std::uint32_t col, column_index(c.as_string()));
        overlay_column_span(layout, col, col, [](ColumnLayout& e) { e.hidden = true; });
      }
    }
    if (bool_or(sh, "hidden", false)) {
      wb.sheet(index).mutable_view().set_visibility(SheetVisibility::kHidden);
    }
  }
  return {};
}

Expected<void, Error> define_names(Workbook& wb, const JsonValue& spec) {
  const JsonValue* names = member(spec, "names");
  if (names == nullptr) {
    return {};
  }
  for (const JsonValue& n : names->as_array()) {
    std::string formula = str_or(n, "refers_to", "");
    if (!formula.empty() && formula[0] == '=') {
      formula.erase(0, 1);
    }
    const std::string scope = str_or(n, "scope", "");
    std::int32_t local = -1;
    if (!scope.empty()) {
      ASSIGN_OR_RETURN(const std::size_t index, sheet_index(wb, scope));
      local = static_cast<std::int32_t>(index);
    }
    RETURN_IF_ERROR(wb.set_defined_name_scoped(str_or(n, "name", ""), formula, local));
  }
  return {};
}

Expected<ChartInsertKind, Error> insert_kind(const std::string& s) {
  static const char* const kNames[] = {"column", "bar", "line", "area", "pie", "scatter", "doughnut"};
  for (std::size_t i = 0; i < 7; ++i) {
    if (s == kNames[i]) {
      return static_cast<ChartInsertKind>(i);
    }
  }
  return invalid("unknown chart kind '" + s + "'");
}

Expected<ChartGrouping, Error> grouping_of(const std::string& s) {
  if (s.empty() || s == "clustered") {
    return ChartGrouping::kClustered;
  }
  if (s == "stacked") {
    return ChartGrouping::kStacked;
  }
  if (s == "percentStacked") {
    return ChartGrouping::kPercentStacked;
  }
  if (s == "standard") {
    return ChartGrouping::kStandard;
  }
  return invalid("unknown grouping '" + s + "'");
}

Expected<ChartSeriesDirection, Error> direction_of(const std::string& s) {
  if (s.empty() || s == "auto") {
    return ChartSeriesDirection::kAuto;
  }
  if (s == "columns") {
    return ChartSeriesDirection::kColumns;
  }
  if (s == "rows") {
    return ChartSeriesDirection::kRows;
  }
  return invalid("unknown direction '" + s + "'");
}

Expected<ChartBlanksAs, Error> blanks_of(const std::string& s) {
  if (s == "gap") {
    return ChartBlanksAs::kGap;
  }
  if (s == "zero") {
    return ChartBlanksAs::kZero;
  }
  if (s == "span") {
    return ChartBlanksAs::kSpan;
  }
  return invalid("unknown blanks_as '" + s + "'");
}

constexpr const char* kScratchSheet = "Scratch";

std::string column_letters(std::uint32_t col) {
  std::string a1 = a1::encode_a1(0, col);
  a1.pop_back();
  return a1;
}

std::string scratch_ref(std::size_t columns) {
  return std::string(kScratchSheet) + "!$A$1:$" + column_letters(static_cast<std::uint32_t>(columns - 1U)) + "$2";
}

Expected<void, Error> set_display(Workbook& wb, std::size_t sheet, std::uint32_t id, const JsonValue& obj) {
  const JsonValue* blanks = member(obj, "blanks_as");
  const JsonValue* vis = member(obj, "plot_vis_only");
  if (blanks == nullptr && vis == nullptr) {
    return {};
  }
  ASSIGN_OR_RETURN(const ChartModel model, get_chart(wb, sheet, id));
  ChartBlanksAs b = model.disp_blanks_as;
  Tri v = model.plot_vis_only;
  if (blanks != nullptr) {
    ASSIGN_OR_RETURN(b, blanks_of(blanks->as_string()));
  }
  if (vis != nullptr) {
    v = vis->as_bool() ? Tri::kTrue : Tri::kFalse;
  }
  return set_chart_display(wb, sheet, id, b, v);
}

/// One chart of a data / structural / create case. Series-listing charts
/// are created over a scratch block and then pointed at their references,
/// the same route the capture takes through `SERIES` formulas.
Expected<void, Error> build_chart(Workbook& wb, const JsonValue& c, std::size_t ordinal) {
  ASSIGN_OR_RETURN(const std::size_t host, sheet_index(wb, str_or(c, "host", wb.sheet(0).name())));
  ChartInsertOptions o;
  ASSIGN_OR_RETURN(o.kind, insert_kind(str_or(c, "kind", "")));
  ASSIGN_OR_RETURN(o.grouping, grouping_of(str_or(c, "grouping", "")));
  o.name = str_or(c, "id", "");
  o.placement.row = static_cast<std::uint32_t>(1U + 16U * ordinal);
  o.placement.col = 20;
  o.placement.width_emu = 4572000;
  o.placement.height_emu = 2743200;
  const JsonValue* series = member(c, "series");
  const bool scatter = o.kind == ChartInsertKind::kScatter;
  if (series != nullptr) {
    const std::size_t n = series->as_array().size();
    o.source = scratch_ref(scatter ? n + 1U : n);
    o.direction = ChartSeriesDirection::kColumns;
    o.header_row = ChartHeaderMode::kNo;
    o.header_column = ChartHeaderMode::kNo;
  } else {
    o.source = str_or(c, "source", "");
    ASSIGN_OR_RETURN(o.direction, direction_of(str_or(c, "direction", "")));
  }
  ASSIGN_OR_RETURN(const std::uint32_t id, insert_chart(wb, host, o));
  if (series != nullptr) {
    ASSIGN_OR_RETURN(const ChartModel model, get_chart(wb, host, id));
    const std::vector<ChartSeries>& made = model.plots.at(0).series;
    if (made.size() != series->as_array().size()) {
      return invalid("chart '" + o.name + "': scratch insert made " + std::to_string(made.size()) + " series");
    }
    for (std::size_t i = 0; i < made.size(); ++i) {
      const JsonValue& s = series->as_array()[i];
      ChartSeriesPatch p;
      p.name = str_or(s, "name", "");
      if (scatter) {
        p.x = str_or(s, "x", "");
        p.y = str_or(s, "y", "");
      } else {
        p.cat = str_or(s, "cat", "");
        p.val = str_or(s, "val", "");
      }
      RETURN_IF_ERROR(set_chart_series(wb, host, id, 0, made[i].idx, p));
    }
  }
  if (const std::string ref = str_or(c, "title_ref", ""); !ref.empty()) {
    ChartTitle t;
    t.kind = ChartTitleKind::kRef;
    t.ref_formula = ref;
    RETURN_IF_ERROR(set_chart_title(wb, host, id, t));
  }
  return set_display(wb, host, id, c);
}

Expected<void, Error> apply_op(Workbook& wb, const JsonValue& op) {
  const std::string kind = str_or(op, "op", "");
  ASSIGN_OR_RETURN(const std::size_t sheet, sheet_index(wb, str_or(op, "sheet", "")));
  if (kind == "rename_sheet") {
    return wb.rename_sheet(static_cast<std::uint32_t>(sheet), str_or(op, "to", ""));
  }
  if (kind == "delete_sheet") {
    return wb.remove_sheet(static_cast<std::uint32_t>(sheet));
  }
  const std::uint32_t count =
      member(op, "count") != nullptr ? static_cast<std::uint32_t>(member(op, "count")->as_number()) : 1U;
  if (kind == "insert_rows" || kind == "delete_rows") {
    ASSIGN_OR_RETURN(const std::uint32_t at, uint_at(op, "at"));
    return kind == "insert_rows" ? wb.insert_rows(sheet, at - 1U, count) : wb.delete_rows(sheet, at - 1U, count);
  }
  if (kind == "insert_cols" || kind == "delete_cols") {
    ASSIGN_OR_RETURN(const std::uint32_t at, column_index(str_or(op, "at", "")));
    return kind == "insert_cols" ? wb.insert_cols(sheet, at, count) : wb.delete_cols(sheet, at, count);
  }
  return invalid("unknown op '" + kind + "'");
}

// -- edits -------------------------------------------------------------------

Expected<std::uint32_t, Error> argb(const JsonValue& v) {
  if (!v.is_string() || v.as_string().size() != 6) {
    return invalid("colors are RRGGBB strings");
  }
  return 0xFF000000U | static_cast<std::uint32_t>(std::stoul(v.as_string(), nullptr, 16));
}

template <typename T, typename F>
Expected<void, Error> patch_field(const JsonValue& obj, const char* key, ChartPatchField<T>& field, F convert) {
  const JsonValue* v = member(obj, key);
  if (v == nullptr) {
    return {};
  }
  if (v->is_null()) {
    field = ChartPatchField<T>::Clear();
    return {};
  }
  ASSIGN_OR_RETURN(T value, convert(*v));
  field = ChartPatchField<T>::Set(std::move(value));
  return {};
}

Expected<double, Error> as_double(const JsonValue& v) {
  if (!v.is_number()) {
    return invalid("expected a number");
  }
  return v.as_number();
}

Expected<std::string, Error> as_text(const JsonValue& v) {
  if (!v.is_string()) {
    return invalid("expected a string");
  }
  return v.as_string();
}

Expected<ChartLineDash, Error> dash_of(const JsonValue& v) {
  static const char* const kNames[] = {"solid", "dash", "dot", "dashDot"};
  for (std::size_t i = 0; i < 4; ++i) {
    if (v.is_string() && v.as_string() == kNames[i]) {
      return static_cast<ChartLineDash>(i);
    }
  }
  return invalid("unknown dash");
}

Expected<ChartLabelPosition, Error> label_pos_of(const JsonValue& v) {
  for (std::size_t i = 0; i < 9; ++i) {
    if (v.is_string() && v.as_string() == label_pos_name(static_cast<ChartLabelPosition>(i))) {
      return static_cast<ChartLabelPosition>(i);
    }
  }
  return invalid("unknown label position");
}

Expected<ChartFillPatch, Error> fill_patch(const JsonValue* obj) {
  ChartFillPatch p;
  if (obj != nullptr) {
    RETURN_IF_ERROR(patch_field(*obj, "color", p.color, argb));
    RETURN_IF_ERROR(patch_field(*obj, "transparency", p.transparency_pct, as_double));
  }
  return p;
}

Expected<ChartLinePatch, Error> line_patch(const JsonValue* obj) {
  ChartLinePatch p;
  if (obj != nullptr) {
    RETURN_IF_ERROR(patch_field(*obj, "color", p.color, argb));
    RETURN_IF_ERROR(patch_field(*obj, "transparency", p.transparency_pct, as_double));
    RETURN_IF_ERROR(patch_field(*obj, "width_pt", p.width_pt, as_double));
    RETURN_IF_ERROR(patch_field(*obj, "dash", p.dash, dash_of));
    if (const JsonValue* none = member(*obj, "none"); none != nullptr) {
      p.none = none->as_bool();
    }
  }
  return p;
}

Expected<ChartMarkerPatch, Error> marker_patch(const JsonValue& obj) {
  ChartMarkerPatch p;
  if (const JsonValue* sym = member(obj, "symbol"); sym != nullptr) {
    bool found = false;
    for (std::size_t i = 0; i < 6; ++i) {
      if (sym->as_string() == marker_name(static_cast<ChartMarkerSymbol>(i))) {
        p.symbol = static_cast<ChartMarkerSymbol>(i);
        found = true;
      }
    }
    if (!found) {
      return invalid("unknown marker symbol");
    }
  }
  RETURN_IF_ERROR(patch_field(obj, "size", p.size, as_double));
  ASSIGN_OR_RETURN(p.fill, fill_patch(member(obj, "fill")));
  ASSIGN_OR_RETURN(p.line, line_patch(member(obj, "line")));
  return p;
}

Expected<ChartLabelsPatch, Error> labels_patch(const JsonValue& obj) {
  ChartLabelsPatch p;
  const std::pair<const char*, std::optional<bool>*> flags[] = {{"val", &p.show_val},
                                                                {"cat", &p.show_cat},
                                                                {"ser", &p.show_ser},
                                                                {"pct", &p.show_pct},
                                                                {"legend_key", &p.show_legend_key}};
  for (const auto& [key, field] : flags) {
    if (const JsonValue* v = member(obj, key); v != nullptr) {
      *field = v->as_bool();
    }
  }
  RETURN_IF_ERROR(patch_field(obj, "separator", p.separator, as_text));
  RETURN_IF_ERROR(patch_field(obj, "position", p.position, label_pos_of));
  RETURN_IF_ERROR(patch_field(obj, "number_format", p.number_format, as_text));
  RETURN_IF_ERROR(patch_field(obj, "font_color", p.font_color, argb));
  RETURN_IF_ERROR(patch_field(obj, "font_size", p.font_size_pt, as_double));
  return p;
}

/// The value axis (scatter: the vertical one) or the category axis (scatter: the horizontal one).
Expected<std::uint32_t, Error> axis_for_role(const ChartModel& m, const std::string& role) {
  const bool scatter = !m.plots.empty() && m.plots[0].type == ChartPlotType::kScatter;
  for (const ChartAxis& a : m.axes) {
    const bool horizontal = a.position == ChartAxisPosition::kBottom || a.position == ChartAxisPosition::kTop;
    const bool category = scatter ? horizontal : (a.kind == ChartAxisKind::kCategory || a.kind == ChartAxisKind::kDate);
    const bool value = scatter ? !horizontal : a.kind == ChartAxisKind::kValue;
    if ((role == "category" && category) || (role == "value" && value)) {
      return a.id;
    }
  }
  return invalid("chart has no " + role + " axis");
}

Expected<void, Error> apply_series_edit(Workbook& wb, std::size_t sheet, std::uint32_t id, const ChartModel& m,
                                        const JsonValue& e) {
  ChartSeriesPatch p;
  ASSIGN_OR_RETURN(p.fill, fill_patch(member(e, "fill")));
  ASSIGN_OR_RETURN(p.line, line_patch(member(e, "line")));
  if (const JsonValue* mk = member(e, "marker"); mk != nullptr) {
    if (mk->is_null()) {
      p.marker = ChartPatchField<ChartMarkerPatch>::Clear();
    } else {
      ASSIGN_OR_RETURN(ChartMarkerPatch mp, marker_patch(*mk));
      p.marker = ChartPatchField<ChartMarkerPatch>::Set(std::move(mp));
    }
  }
  if (const JsonValue* lb = member(e, "labels"); lb != nullptr) {
    if (lb->is_null()) {
      p.labels = ChartPatchField<ChartLabelsPatch>::Clear();
    } else {
      ASSIGN_OR_RETURN(ChartLabelsPatch lp, labels_patch(*lb));
      p.labels = ChartPatchField<ChartLabelsPatch>::Set(std::move(lp));
    }
  }
  const JsonValue* which = member(e, "series");
  for (const ChartSeries& s : m.plots.at(0).series) {
    const bool all = which != nullptr && which->is_string() && which->as_string() == "all";
    if (all ? !s.filtered : (which != nullptr && which->is_number() && s.idx == which->as_number())) {
      RETURN_IF_ERROR(set_chart_series(wb, sheet, id, 0, s.idx, p));
    }
  }
  return {};
}

Expected<void, Error> apply_axis_edit(Workbook& wb, std::size_t sheet, std::uint32_t id, const ChartModel& m,
                                      const JsonValue& e) {
  ASSIGN_OR_RETURN(const std::uint32_t axis, axis_for_role(m, str_or(e, "axis", "")));
  ChartAxisPatch p;
  if (const JsonValue* v = member(e, "visible"); v != nullptr) {
    p.visible = v->as_bool();
  }
  RETURN_IF_ERROR(patch_field(e, "title", p.title, as_text));
  RETURN_IF_ERROR(patch_field(e, "min", p.min, as_double));
  RETURN_IF_ERROR(patch_field(e, "max", p.max, as_double));
  RETURN_IF_ERROR(patch_field(e, "major_unit", p.major_unit, as_double));
  RETURN_IF_ERROR(patch_field(e, "number_format", p.number_format, as_text));
  if (const JsonValue* v = member(e, "major_gridlines"); v != nullptr) {
    p.major_gridlines = v->as_bool();
  }
  if (const JsonValue* v = member(e, "reversed"); v != nullptr) {
    p.reversed = v->as_bool();
  }
  RETURN_IF_ERROR(patch_field(e, "log_base", p.log_base, as_double));
  if (const JsonValue* v = member(e, "crosses"); v != nullptr) {
    const std::string s = v->as_string();
    p.crosses = s == "autoZero" ? ChartAxisCrosses::kAutoZero
                : s == "min"    ? ChartAxisCrosses::kMin
                : s == "max"    ? ChartAxisCrosses::kMax
                                : ChartAxisCrosses::kAt;
    if (const JsonValue* at = member(e, "crosses_at"); at != nullptr) {
      p.crosses_at = at->as_number();
    }
  }
  RETURN_IF_ERROR(patch_field(e, "tick_label_pos", p.tick_label_pos, as_text));
  return set_chart_axis(wb, sheet, id, axis, p);
}

std::vector<std::uint32_t> uint_list(const JsonValue* v) {
  std::vector<std::uint32_t> out;
  if (v != nullptr && v->is_array()) {
    for (const JsonValue& x : v->as_array()) {
      out.push_back(static_cast<std::uint32_t>(x.as_number()));
    }
  }
  return out;
}

Expected<void, Error> apply_edit(Workbook& wb, std::size_t sheet, std::uint32_t id, const JsonValue& e) {
  ASSIGN_OR_RETURN(const ChartModel m, get_chart(wb, sheet, id));
  const std::string op = str_or(e, "op", "");
  if (op == "title") {
    ChartTitle t;
    const std::string kind = str_or(e, "kind", "");
    t.kind = kind == "text" ? ChartTitleKind::kText : kind == "auto" ? ChartTitleKind::kAuto : ChartTitleKind::kNone;
    t.text = str_or(e, "text", "");
    return set_chart_title(wb, sheet, id, t);
  }
  if (op == "legend") {
    ChartLegendPatch p;
    const std::string pos = str_or(e, "position", "r");
    bool found = false;
    for (std::size_t i = 0; i < 6; ++i) {
      if (pos == legend_pos_name(static_cast<ChartLegendPosition>(i))) {
        p.position = static_cast<ChartLegendPosition>(i);
        found = true;
      }
    }
    if (!found) {
      return invalid("unknown legend position '" + pos + "'");
    }
    p.overlay = bool_or(e, "overlay", false);
    if (const JsonValue* hidden = member(e, "hidden"); hidden != nullptr) {
      p.hidden_series_idx = uint_list(hidden);
    }
    return set_chart_legend(wb, sheet, id, p);
  }
  if (op == "series") {
    return apply_series_edit(wb, sheet, id, m, e);
  }
  if (op == "point") {
    ChartPointPatch p;
    ASSIGN_OR_RETURN(p.fill, fill_patch(member(e, "fill")));
    ASSIGN_OR_RETURN(p.line, line_patch(member(e, "line")));
    ASSIGN_OR_RETURN(const std::uint32_t series, uint_at(e, "series"));
    ASSIGN_OR_RETURN(const std::uint32_t point, uint_at(e, "point"));
    return set_chart_point(wb, sheet, id, 0, series, point, p);
  }
  if (op == "axis") {
    return apply_axis_edit(wb, sheet, id, m, e);
  }
  if (op == "plot") {
    ChartPlotPatch p;
    if (const JsonValue* v = member(e, "gap_width"); v != nullptr) {
      p.gap_width = static_cast<std::uint32_t>(v->as_number());
    }
    if (const JsonValue* v = member(e, "overlap"); v != nullptr) {
      p.overlap = static_cast<std::int32_t>(v->as_number());
    }
    if (const JsonValue* v = member(e, "first_slice_angle"); v != nullptr) {
      p.first_slice_angle = static_cast<std::uint32_t>(v->as_number());
    }
    if (const JsonValue* v = member(e, "vary_colors"); v != nullptr) {
      p.vary_colors = v->as_bool();
    }
    if (const JsonValue* v = member(e, "hole_size"); v != nullptr) {
      p.hole_size = static_cast<std::uint32_t>(v->as_number());
    }
    return set_chart_plot(wb, sheet, id, 0, p);
  }
  if (op == "filter") {
    return set_chart_filter(wb, sheet, id, uint_list(member(e, "hidden_series")),
                            uint_list(member(e, "hidden_categories")));
  }
  if (op == "display") {
    return set_display(wb, sheet, id, e);
  }
  if (op == "type") {
    ASSIGN_OR_RETURN(const ChartInsertKind kind, insert_kind(str_or(e, "kind", "")));
    ASSIGN_OR_RETURN(const ChartGrouping grouping, grouping_of(str_or(e, "grouping", "")));
    return set_chart_type(wb, sheet, id, 0, kind, grouping);
  }
  return invalid("unknown edit '" + op + "'");
}

Expected<std::uint32_t, Error> find_chart(const Workbook& wb, std::size_t sheet, const std::string& name) {
  ASSIGN_OR_RETURN(const std::vector<DrawingObject> objects, list_drawing_objects(wb, sheet));
  for (const DrawingObject& o : objects) {
    if (o.kind == DrawingObjectKind::kChart && o.name == name) {
      return o.object_id;
    }
  }
  return invalid("no chart named '" + name + "'");
}

Expected<std::vector<std::uint8_t>, Error> read_file(const std::string& path) {
  std::FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    return invalid("cannot open '" + path + "'");
  }
  std::vector<std::uint8_t> bytes;
  std::uint8_t buf[65536];
  std::size_t n = 0;
  while ((n = std::fread(buf, 1, sizeof buf, f)) > 0) {
    bytes.insert(bytes.end(), buf, buf + n);
  }
  std::fclose(f);
  return bytes;
}

bool ends_with(const std::string& s, const char* suffix) {
  const std::size_t n = std::strlen(suffix);
  return s.size() >= n && s.compare(s.size() - n, n, suffix) == 0;
}

}  // namespace

Expected<Workbook, Error> read_workbook_bytes(const std::vector<std::uint8_t>& bytes, bool xlsb) {
  const io::ByteSpan span{bytes.data(), bytes.size()};
  Workbook wb = Workbook::create_empty();
  if (xlsb) {
    ASSIGN_OR_RETURN(io::xlsb::XlsbReadResult read, io::xlsb::read_xlsb(span));
    wb = std::move(read.workbook);
  } else {
    ASSIGN_OR_RETURN(io::OoxmlReadResult read, io::read_ooxml(span));
    wb = std::move(read.workbook);
  }
  wb.set_excel_profile(mac_365_ja_jp_profile());
  return wb;
}

Expected<Workbook, Error> load_workbook_file(const std::string& path) {
  ASSIGN_OR_RETURN(const std::vector<std::uint8_t> bytes, read_file(path));
  return read_workbook_bytes(bytes, ends_with(path, ".xlsb"));
}

/// Builds Formulon's workbook for one case spec (see the chart track in tests/oracle/README.md).
Expected<Workbook, Error> build_chart_case(const JsonValue& spec, const std::string& fixtures_dir) {
  if (const std::string fixture = str_or(spec, "fixture", ""); !fixture.empty()) {
    ASSIGN_OR_RETURN(Workbook wb, load_workbook_file(fixtures_dir + "/" + fixture));
    const JsonValue* target = member(spec, "target");
    if (target == nullptr) {
      return invalid("an edit case needs a 'target'");
    }
    ASSIGN_OR_RETURN(const std::size_t sheet, sheet_index(wb, str_or(*target, "sheet", "Sheet1")));
    ASSIGN_OR_RETURN(const std::uint32_t id, find_chart(wb, sheet, str_or(*target, "chart", "")));
    if (const JsonValue* edits = member(spec, "edits"); edits != nullptr) {
      for (const JsonValue& e : edits->as_array()) {
        RETURN_IF_ERROR(apply_edit(wb, sheet, id, e));
      }
    }
    return wb;
  }

  Workbook wb = Workbook::create_empty();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  bool wrote_formula = false;
  RETURN_IF_ERROR(build_sheets(wb, spec, &wrote_formula));
  RETURN_IF_ERROR(define_names(wb, spec));
  const JsonValue* charts = member(spec, "charts");
  bool scratch = false;
  if (charts != nullptr) {
    std::size_t widest = 0;
    for (const JsonValue& c : charts->as_array()) {
      if (const JsonValue* s = member(c, "series"); s != nullptr) {
        widest = std::max(widest, s->as_array().size() + 1U);
      }
    }
    if (widest > 0) {
      const std::size_t index = wb.add_sheet(kScratchSheet);
      for (std::uint32_t col = 0; col < widest; ++col) {
        wb.sheet(index).set_cell_value(0, col, Value::number(1.0 + col));
        wb.sheet(index).set_cell_value(1, col, Value::number(2.0 + col));
      }
      scratch = true;
    }
  }
  if (wrote_formula || scratch) {
    RETURN_IF_ERROR(wb.recalc(eval::default_registry()));
  }
  if (charts != nullptr) {
    std::size_t ordinal = 0;
    for (const JsonValue& c : charts->as_array()) {
      RETURN_IF_ERROR(build_chart(wb, c, ordinal++));
    }
  }
  if (scratch) {
    ASSIGN_OR_RETURN(const std::size_t index, sheet_index(wb, kScratchSheet));
    RETURN_IF_ERROR(wb.remove_sheet(static_cast<std::uint32_t>(index)));
  }
  RETURN_IF_ERROR(apply_visibility(wb, spec));
  if (const JsonValue* ops = member(spec, "ops"); ops != nullptr) {
    for (const JsonValue& op : ops->as_array()) {
      RETURN_IF_ERROR(apply_op(wb, op));
    }
  }
  return wb;
}

/// `{"charts": [...]}` with the requested views of every chart, sheets in
/// workbook order and charts in drawing order.
Expected<std::string, Error> dump_charts_json(const Workbook& wb, const std::vector<std::string>& views) {
  ASSIGN_OR_RETURN(const std::vector<SheetChartPartRef> parts, all_chart_parts(wb));
  std::string out;
  Json j(out);
  j.open('{');
  j.key("charts");
  j.open('[');
  for (const SheetChartPartRef& part : parts) {
    ASSIGN_OR_RETURN(const ChartModel m, get_chart(wb, part.sheet, part.ref.object_id));
    j.open('{');
    j.key("sheet");
    j.str(wb.sheet(part.sheet).name());
    j.key("name");
    j.str(m.name);
    j.key("chartex");
    j.boolean(m.is_chartex);
    if (has_view(views, "shape")) {
      j.key("shape");
      write_shape(j, m);
    }
    if (has_view(views, "refs")) {
      j.key("refs");
      write_refs(j, m);
    }
    if (has_view(views, "cache")) {
      j.key("cache");
      write_cache(j, m);
    }
    if (has_view(views, "format")) {
      j.key("format");
      write_format(j, m);
    }
    j.close('}');
  }
  j.close(']');
  j.close('}');
  out += '\n';
  return out;
}

}  // namespace chart_oracle
}  // namespace oracle
}  // namespace tests
}  // namespace formulon

#ifndef FORMULON_CHART_MODEL_DUMP_NO_MAIN

namespace {

int usage() {
  std::fputs(
      "usage: chart_model_dump [--view shape,refs,cache,format|all] <file.xlsx|.xlsb>\n"
      "       chart_model_dump --build <cases.json> --case <id> --out <file.xlsx> [--fixtures <dir>]\n",
      stderr);
  return 2;
}

int fail(const formulon::Error& e) {
  std::fprintf(stderr, "chart_model_dump: %s\n", e.message.c_str());
  return 1;
}

std::vector<std::string> split_views(const std::string& s) {
  std::vector<std::string> out;
  std::size_t start = 0;
  while (start <= s.size()) {
    const std::size_t comma = s.find(',', start);
    out.push_back(s.substr(start, comma == std::string::npos ? std::string::npos : comma - start));
    if (comma == std::string::npos) {
      break;
    }
    start = comma + 1;
  }
  return out;
}

}  // namespace

int main(int argc, char** argv) {
  namespace co = formulon::tests::oracle::chart_oracle;
  std::string views = "all";
  std::string build;
  std::string case_id;
  std::string out_path;
  std::string fixtures = FORMULON_CHART_FIXTURES_DIR;
  std::string input;
  for (int i = 1; i < argc; ++i) {
    const std::string arg = argv[i];
    const bool has_value = i + 1 < argc;
    if (arg == "--view" && has_value) {
      views = argv[++i];
    } else if (arg == "--build" && has_value) {
      build = argv[++i];
    } else if (arg == "--case" && has_value) {
      case_id = argv[++i];
    } else if (arg == "--out" && has_value) {
      out_path = argv[++i];
    } else if (arg == "--fixtures" && has_value) {
      fixtures = argv[++i];
    } else if (!arg.empty() && arg[0] != '-' && input.empty()) {
      input = arg;
    } else {
      return usage();
    }
  }

  if (!build.empty()) {
    if (case_id.empty() || out_path.empty()) {
      return usage();
    }
    auto doc = formulon::tests::oracle::parse_json_file(build);
    if (!doc) {
      std::fprintf(stderr, "chart_model_dump: %s\n", doc.error().message.c_str());
      return 1;
    }
    const formulon::tests::oracle::JsonValue* cases = doc.value().find("cases");
    if (cases == nullptr || !cases->is_array()) {
      std::fputs("chart_model_dump: no 'cases' list\n", stderr);
      return 1;
    }
    for (const auto& c : cases->as_array()) {
      const auto* id = c.find("id");
      const auto* spec = c.find("spec");
      if (id == nullptr || !id->is_string() || id->as_string() != case_id || spec == nullptr) {
        continue;
      }
      auto wb = co::build_chart_case(*spec, fixtures);
      if (!wb) {
        return fail(wb.error());
      }
      auto bytes = wb.value().save();
      if (!bytes) {
        return fail(bytes.error());
      }
      std::FILE* f = std::fopen(out_path.c_str(), "wb");
      if (f == nullptr || std::fwrite(bytes.value().data(), 1, bytes.value().size(), f) != bytes.value().size()) {
        std::fprintf(stderr, "chart_model_dump: cannot write '%s'\n", out_path.c_str());
        if (f != nullptr) {
          std::fclose(f);
        }
        return 1;
      }
      std::fclose(f);
      return 0;
    }
    std::fprintf(stderr, "chart_model_dump: no case '%s'\n", case_id.c_str());
    return 1;
  }

  if (input.empty()) {
    return usage();
  }
  auto wb = co::load_workbook_file(input);
  if (!wb) {
    return fail(wb.error());
  }
  auto json = co::dump_charts_json(wb.value(), split_views(views));
  if (!json) {
    return fail(json.error());
  }
  std::fputs(json.value().c_str(), stdout);
  return 0;
}

#endif  // FORMULON_CHART_MODEL_DUMP_NO_MAIN
