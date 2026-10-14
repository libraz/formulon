//
// C ABI - charts: the typed chart view and the resolved series data behind
// owned handles, chart creation, and the typed edits.
//
// Both handles copy the C++ result at open time and hand out pointers into
// that copy, so every string and array stays valid until the handle is
// destroyed. Patch structs are translated field by field into the C++
// patch types; the edits themselves live in `chart/chart_edit`.

#include <cstddef>
#include <cstdint>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "chart/chart_edit.h"
#include "chart/chart_model.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

struct fm_chart {
  formulon::ChartModel model;
};

struct fm_chart_data {
  /// De-interleaved point arrays of one series.
  struct Points {
    std::vector<double> values;
    std::vector<std::uint8_t> states;
    std::vector<double> x;
  };
  formulon::ChartData data;
  std::vector<Points> points;
};

namespace {

using formulon::ChartPatchField;
using formulon::FormulonErrorCode;

constexpr double kNaN = std::numeric_limits<double>::quiet_NaN();

fm_status_t null_argument(const char* api) {
  return set_binding_error(FormulonErrorCode::kBindingNullPointer, (std::string(api) + ": NULL argument").c_str());
}

const char* opt_cstr(const std::optional<std::string>& s) {
  return s ? s->c_str() : nullptr;
}

const char* nonempty_cstr(const std::string& s) {
  return s.empty() ? nullptr : s.c_str();
}

std::int32_t tri(formulon::Tri t) {
  return static_cast<std::int32_t>(t);
}

double transparency_of(const std::optional<std::uint32_t>& alpha) {
  return alpha ? 100.0 - static_cast<double>(*alpha) / 1000.0 : kNaN;
}

// ---------------------------------------------------------------------------
// Model -> C
// ---------------------------------------------------------------------------

fm_chart_title_t to_fm(const formulon::ChartTitle& t) {
  return fm_chart_title_t{static_cast<std::int32_t>(t.kind), t.text.c_str(), t.ref_formula.c_str()};
}

fm_chart_fill_t to_fm(const formulon::ChartFill& f) {
  return fm_chart_fill_t{f.color ? 1 : 0, f.color.value_or(0U), transparency_of(f.alpha)};
}

fm_chart_line_t to_fm(const formulon::ChartLine& l) {
  fm_chart_line_t out{};
  out.has_color = l.color ? 1 : 0;
  out.color = l.color.value_or(0U);
  out.transparency_pct = transparency_of(l.alpha);
  out.width_pt = l.width_emu ? static_cast<double>(*l.width_emu) / 12700.0 : kNaN;
  out.has_dash = l.dash ? 1 : 0;
  out.dash = l.dash ? static_cast<std::int32_t>(*l.dash) : 0;
  out.none = l.none ? 1 : 0;
  return out;
}

fm_chart_marker_t to_fm(const formulon::ChartMarker& m) {
  fm_chart_marker_t out{};
  out.symbol = static_cast<std::int32_t>(m.symbol);
  out.has_size = m.size ? 1 : 0;
  out.size = m.size.value_or(0U);
  out.fill = to_fm(m.fill);
  out.line = to_fm(m.line);
  return out;
}

fm_chart_labels_t to_fm(const formulon::ChartLabels& l) {
  fm_chart_labels_t out{};
  out.show_value = l.show_val ? 1 : 0;
  out.show_category = l.show_cat ? 1 : 0;
  out.show_series = l.show_ser ? 1 : 0;
  out.show_percent = l.show_pct ? 1 : 0;
  out.show_legend_key = l.show_legend_key ? 1 : 0;
  out.separator = opt_cstr(l.separator);
  out.has_position = l.position ? 1 : 0;
  out.position = l.position ? static_cast<std::int32_t>(*l.position) : 0;
  out.number_format = l.num_fmt ? l.num_fmt->code.c_str() : nullptr;
  out.source_linked = l.num_fmt && l.num_fmt->source_linked ? 1 : 0;
  out.has_font_color = l.font_color ? 1 : 0;
  out.font_color = l.font_color.value_or(0U);
  out.font_size_pt = l.font_size_hundredths_pt ? static_cast<double>(*l.font_size_hundredths_pt) / 100.0 : kNaN;
  out.deleted = l.deleted ? 1 : 0;
  return out;
}

fm_chart_data_source_t to_fm(const formulon::ChartDataSource& d) {
  return fm_chart_data_source_t{static_cast<std::int32_t>(d.kind), d.formula.c_str(), d.literal.size()};
}

const formulon::ChartDataSource* source_of(const formulon::ChartSeries& s, std::int32_t role) {
  switch (role) {
    case FM_CHART_ROLE_NAME:
      return &s.name;
    case FM_CHART_ROLE_CATEGORIES:
      return &s.cat;
    case FM_CHART_ROLE_VALUES:
      return &s.val;
    case FM_CHART_ROLE_X:
      return &s.x;
    case FM_CHART_ROLE_Y:
      return &s.y;
    case FM_CHART_ROLE_SIZE:
      return &s.size;
    default:
      return nullptr;
  }
}

const std::vector<formulon::SheetRect>* ranges_of(const formulon::ChartDataSeries& s, std::int32_t role) {
  switch (role) {
    case FM_CHART_ROLE_NAME:
      return &s.name_ranges;
    case FM_CHART_ROLE_CATEGORIES:
      return &s.cat_ranges;
    case FM_CHART_ROLE_VALUES:
      return &s.val_ranges;
    case FM_CHART_ROLE_X:
      return &s.x_ranges;
    case FM_CHART_ROLE_Y:
      return &s.y_ranges;
    case FM_CHART_ROLE_SIZE:
      return &s.size_ranges;
    default:
      return nullptr;
  }
}

/// Validates `plot` / `series` against the handle and returns the series.
fm_status_t find_series(const fm_chart_t* chart, std::size_t plot, std::size_t series, const char* api,
                        const formulon::ChartSeries** out) {
  if (auto rc = check_index(plot, chart->model.plots.size(), api, "plot"); rc != 0) {
    return rc;
  }
  const std::vector<formulon::ChartSeries>& list = chart->model.plots[plot].series;
  if (auto rc = check_index(series, list.size(), api, "series"); rc != 0) {
    return rc;
  }
  *out = &list[series];
  return 0;
}

std::unique_ptr<fm_chart_data_t> make_data_handle(formulon::ChartData data) {
  auto handle = std::unique_ptr<fm_chart_data_t>(new fm_chart_data_t{});
  handle->data = std::move(data);
  const bool xy = handle->data.domain == formulon::ChartData::Domain::kXY;
  handle->points.resize(handle->data.series.size());
  for (std::size_t i = 0; i < handle->data.series.size(); ++i) {
    const formulon::ChartDataSeries& s = handle->data.series[i];
    fm_chart_data::Points& p = handle->points[i];
    if (xy) {
      for (const formulon::ChartXYPoint& pt : s.xy) {
        p.x.push_back(pt.x);
        p.values.push_back(pt.y);
        p.states.push_back(static_cast<std::uint8_t>(pt.state));
      }
    } else {
      for (const formulon::ChartDataPoint& pt : s.points) {
        p.values.push_back(pt.value);
        p.states.push_back(static_cast<std::uint8_t>(pt.state));
      }
    }
  }
  return handle;
}

// ---------------------------------------------------------------------------
// C patches -> model
// ---------------------------------------------------------------------------

/// Translates patch fields, keeping the first validation failure.
class PatchReader {
 public:
  explicit PatchReader(const char* api) : api_(api) {}

  fm_status_t status() const { return rc_; }

  template <typename T>
  void field(std::int32_t op, T value, ChartPatchField<T>& out, const char* name) {
    if (!op_ok(op, name)) {
      return;
    }
    if (op == FM_CHART_PATCH_CLEAR) {
      out = ChartPatchField<T>::Clear();
    } else if (op == FM_CHART_PATCH_SET) {
      out = ChartPatchField<T>::Set(std::move(value));
    }
  }

  template <typename E>
  void enum_field(std::int32_t op, std::int32_t value, E max, ChartPatchField<E>& out, const char* name) {
    if (op == FM_CHART_PATCH_SET && !enum_ok(value, max, name)) {
      return;
    }
    field(op, static_cast<E>(value), out, name);
  }

  void text_field(std::int32_t op, const char* value, ChartPatchField<std::string>& out, const char* name) {
    field(op, std::string(value != nullptr ? value : ""), out, name);
  }

  template <typename T>
  void optional(std::int32_t op, T value, std::optional<T>& out, const char* name) {
    if (!op_ok(op, name)) {
      return;
    }
    if (op == FM_CHART_PATCH_CLEAR) {
      fail(name, "cannot be cleared");
    } else if (op == FM_CHART_PATCH_SET) {
      out = std::move(value);
    }
  }

  void flag(std::int32_t op, std::int32_t value, std::optional<bool>& out, const char* name) {
    optional(op, value != 0, out, name);
  }

  template <typename E>
  void enum_optional(std::int32_t op, std::int32_t value, E max, std::optional<E>& out, const char* name) {
    if (op == FM_CHART_PATCH_SET && !enum_ok(value, max, name)) {
      return;
    }
    optional(op, static_cast<E>(value), out, name);
  }

  bool enum_ok(std::int64_t value, std::int64_t max, const char* name) {
    if (rc_ == 0) {
      rc_ = check_enum_domain(value, max, api_, name);
    }
    return rc_ == 0;
  }

  template <typename E>
  bool enum_ok(std::int64_t value, E max, const char* name) {
    return enum_ok(value, static_cast<std::int64_t>(max), name);
  }

  void fail(const char* name, const char* why) {
    if (rc_ == 0) {
      rc_ =
          set_binding_error(FormulonErrorCode::kInvalidArgument, (std::string(api_) + ": " + name + " " + why).c_str());
    }
  }

 private:
  bool op_ok(std::int32_t op, const char* name) {
    return enum_ok(op, static_cast<std::int64_t>(FM_CHART_PATCH_SET), name);
  }

  const char* api_;
  fm_status_t rc_ = 0;
};

formulon::ChartFillPatch read_fill(PatchReader& r, const fm_chart_fill_patch& p) {
  formulon::ChartFillPatch out;
  r.field(p.color_op, p.color, out.color, "fill.color_op");
  r.field(p.transparency_op, p.transparency_pct, out.transparency_pct, "fill.transparency_op");
  return out;
}

formulon::ChartLinePatch read_line(PatchReader& r, const fm_chart_line_patch& p) {
  formulon::ChartLinePatch out;
  r.field(p.color_op, p.color, out.color, "line.color_op");
  r.field(p.transparency_op, p.transparency_pct, out.transparency_pct, "line.transparency_op");
  r.field(p.width_op, p.width_pt, out.width_pt, "line.width_op");
  r.enum_field(p.dash_op, p.dash, formulon::ChartLineDash::kOther, out.dash, "line.dash");
  r.flag(p.none_op, p.none, out.none, "line.none_op");
  return out;
}

formulon::ChartMarkerPatch read_marker(PatchReader& r, const fm_chart_marker_patch& p) {
  formulon::ChartMarkerPatch out;
  r.enum_optional(p.symbol_op, p.symbol, formulon::ChartMarkerSymbol::kDash, out.symbol, "marker.symbol");
  r.field(p.size_op, p.size, out.size, "marker.size_op");
  out.fill = read_fill(r, p.fill);
  out.line = read_line(r, p.line);
  return out;
}

formulon::ChartLabelsPatch read_labels(PatchReader& r, const fm_chart_labels_patch& p) {
  formulon::ChartLabelsPatch out;
  r.flag(p.show_value_op, p.show_value, out.show_val, "labels.show_value_op");
  r.flag(p.show_category_op, p.show_category, out.show_cat, "labels.show_category_op");
  r.flag(p.show_series_op, p.show_series, out.show_ser, "labels.show_series_op");
  r.flag(p.show_percent_op, p.show_percent, out.show_pct, "labels.show_percent_op");
  r.flag(p.show_legend_key_op, p.show_legend_key, out.show_legend_key, "labels.show_legend_key_op");
  r.text_field(p.separator_op, p.separator, out.separator, "labels.separator_op");
  r.enum_field(p.position_op, p.position, formulon::ChartLabelPosition::kRight, out.position, "labels.position");
  r.text_field(p.number_format_op, p.number_format, out.number_format, "labels.number_format_op");
  r.field(p.font_color_op, p.font_color, out.font_color, "labels.font_color_op");
  r.field(p.font_size_op, p.font_size_pt, out.font_size_pt, "labels.font_size_op");
  return out;
}

void read_ref(const char* text, std::optional<std::string>& out) {
  if (text != nullptr) {
    out = std::string(text);
  }
}

formulon::ChartSeriesPatch read_series(PatchReader& r, const fm_chart_series_patch& p) {
  formulon::ChartSeriesPatch out;
  read_ref(p.name, out.name);
  read_ref(p.categories, out.cat);
  read_ref(p.values, out.val);
  read_ref(p.x, out.x);
  read_ref(p.y, out.y);
  read_ref(p.size, out.size);
  out.fill = read_fill(r, p.fill);
  out.line = read_line(r, p.line);
  r.field(p.marker_op, p.marker_op == FM_CHART_PATCH_SET ? read_marker(r, p.marker) : formulon::ChartMarkerPatch{},
          out.marker, "marker_op");
  r.flag(p.smooth_op, p.smooth, out.smooth, "smooth_op");
  r.field(p.labels_op, p.labels_op == FM_CHART_PATCH_SET ? read_labels(r, p.labels) : formulon::ChartLabelsPatch{},
          out.labels, "labels_op");
  return out;
}

/// Shared entry checks of the edit functions.
fm_status_t edit_entry(const fm_workbook_t* wb, std::size_t sheet_index, const void* required, const char* api) {
  clear_last_error();
  if (required == nullptr) {
    return null_argument(api);
  }
  return check_sheet_index(wb, sheet_index, api);
}

fm_status_t finish(const formulon::Expected<void, formulon::Error>& result) {
  return result ? 0 : set_last_error(result.error());
}

}  // namespace

// ---------------------------------------------------------------------------
// Chart model reads
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_chart_open(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                           fm_chart_t** out) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_chart_open";
  if (out == nullptr) {
    return null_argument(kApi);
  }
  *out = nullptr;
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  auto model = formulon::get_chart(wb->workbook(), sheet_index, object_id);
  if (!model) {
    return set_last_error(model.error());
  }
  auto handle = std::unique_ptr<fm_chart_t>(new fm_chart_t{});
  handle->model = std::move(model.value());
  *out = handle.release();
  return 0;
}

extern "C" void fm_chart_destroy(fm_chart_t* chart) {
  delete chart;
}

extern "C" fm_status_t fm_chart_info(const fm_chart_t* chart, fm_chart_info_t* out) {
  clear_last_error();
  if (chart == nullptr || out == nullptr) {
    return null_argument("fm_chart_info");
  }
  const formulon::ChartModel& m = chart->model;
  *out = fm_chart_info_t{};
  out->object_id = m.object_id;
  out->name = m.name.c_str();
  out->descr = m.descr.c_str();
  out->is_chartex = m.is_chartex ? 1 : 0;
  out->is_pivot_chart = m.is_pivot_chart ? 1 : 0;
  out->on_chartsheet = m.on_chartsheet ? 1 : 0;
  out->anchor_kind = static_cast<std::int32_t>(m.anchor_kind);
  out->edit_as = static_cast<std::int32_t>(m.edit_as);
  out->from_row = m.from.row;
  out->from_col = m.from.col;
  out->from_row_off = m.from.row_off;
  out->from_col_off = m.from.col_off;
  out->to_row = m.to.row;
  out->to_col = m.to.col;
  out->to_row_off = m.to.row_off;
  out->to_col_off = m.to.col_off;
  out->cx = m.cx;
  out->cy = m.cy;
  out->title = to_fm(m.title);
  out->auto_title_deleted = tri(m.auto_title_deleted);
  out->legend_position =
      m.legend.present ? static_cast<std::int32_t>(m.legend.position) : static_cast<std::int32_t>(FM_CHART_LEGEND_NONE);
  out->legend_overlay = m.legend.overlay ? 1 : 0;
  out->legend_hidden_count = m.legend.hidden_series_idx.size();
  out->blanks_as = static_cast<std::int32_t>(m.disp_blanks_as);
  out->plot_vis_only = tri(m.plot_vis_only);
  out->date1904 = m.date1904 ? 1 : 0;
  out->has_style = m.style ? 1 : 0;
  out->style = m.style.value_or(0);
  out->plot_count = m.plots.size();
  out->axis_count = m.axes.size();
  out->hidden_category_count = m.category_filter.size();
  if (m.source) {
    out->source_range = m.source->range.c_str();
    out->source_direction = static_cast<std::int32_t>(m.source->direction);
    out->source_header_row = m.source->header_row ? 1 : 0;
    out->source_header_column = m.source->header_column ? 1 : 0;
  }
  return 0;
}

extern "C" fm_status_t fm_chart_legend_hidden_at(const fm_chart_t* chart, size_t idx, uint32_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_legend_hidden_at";
  if (chart == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  const std::vector<std::uint32_t>& list = chart->model.legend.hidden_series_idx;
  if (auto rc = check_index(idx, list.size(), kApi, "idx"); rc != 0) {
    return rc;
  }
  *out = list[idx];
  return 0;
}

extern "C" fm_status_t fm_chart_hidden_category_at(const fm_chart_t* chart, size_t idx, uint32_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_hidden_category_at";
  if (chart == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  const std::vector<std::uint32_t>& list = chart->model.category_filter;
  if (auto rc = check_index(idx, list.size(), kApi, "idx"); rc != 0) {
    return rc;
  }
  *out = list[idx];
  return 0;
}

extern "C" fm_status_t fm_chart_plot_at(const fm_chart_t* chart, size_t plot, fm_chart_plot_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_plot_at";
  if (chart == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  if (auto rc = check_index(plot, chart->model.plots.size(), kApi, "plot"); rc != 0) {
    return rc;
  }
  const formulon::ChartPlot& p = chart->model.plots[plot];
  *out = fm_chart_plot_t{};
  out->type = static_cast<std::int32_t>(p.type);
  out->chartex_layout = p.chartex_layout.c_str();
  out->bar_dir = static_cast<std::int32_t>(p.bar_dir);
  out->has_grouping = p.grouping ? 1 : 0;
  out->grouping = p.grouping ? static_cast<std::int32_t>(*p.grouping) : 0;
  out->vary_colors = tri(p.vary_colors);
  out->has_gap_width = p.gap_width ? 1 : 0;
  out->gap_width = p.gap_width.value_or(0U);
  out->has_overlap = p.overlap ? 1 : 0;
  out->overlap = p.overlap.value_or(0);
  out->has_hole_size = p.hole_size ? 1 : 0;
  out->hole_size = p.hole_size.value_or(0U);
  out->has_first_slice_angle = p.first_slice_angle ? 1 : 0;
  out->first_slice_angle = p.first_slice_angle.value_or(0U);
  out->axis_id_count = p.axis_ids.size();
  out->series_count = p.series.size();
  return 0;
}

extern "C" fm_status_t fm_chart_plot_axis_id_at(const fm_chart_t* chart, size_t plot, size_t idx, uint32_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_plot_axis_id_at";
  if (chart == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  if (auto rc = check_index(plot, chart->model.plots.size(), kApi, "plot"); rc != 0) {
    return rc;
  }
  const std::vector<std::uint32_t>& ids = chart->model.plots[plot].axis_ids;
  if (auto rc = check_index(idx, ids.size(), kApi, "idx"); rc != 0) {
    return rc;
  }
  *out = ids[idx];
  return 0;
}

extern "C" fm_status_t fm_chart_series_at(const fm_chart_t* chart, size_t plot, size_t series, fm_chart_series_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_series_at";
  if (chart == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  const formulon::ChartSeries* s = nullptr;
  if (auto rc = find_series(chart, plot, series, kApi, &s); rc != 0) {
    return rc;
  }
  *out = fm_chart_series_t{};
  out->idx = s->idx;
  out->order = s->order;
  out->filtered = s->filtered ? 1 : 0;
  out->name = to_fm(s->name);
  out->categories = to_fm(s->cat);
  out->values = to_fm(s->val);
  out->x = to_fm(s->x);
  out->y = to_fm(s->y);
  out->size = to_fm(s->size);
  out->fill = to_fm(s->fill);
  out->line = to_fm(s->line);
  if (s->marker) {
    out->has_marker = 1;
    out->marker = to_fm(*s->marker);
  }
  out->smooth = tri(s->smooth);
  out->point_count = s->points.size();
  if (s->labels) {
    out->has_labels = 1;
    out->labels = to_fm(*s->labels);
  }
  return 0;
}

extern "C" fm_status_t fm_chart_series_literal_at(const fm_chart_t* chart, size_t plot, size_t series, int32_t role,
                                                  size_t idx, fm_chart_literal_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_series_literal_at";
  if (chart == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  const formulon::ChartSeries* s = nullptr;
  if (auto rc = find_series(chart, plot, series, kApi, &s); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(role, FM_CHART_ROLE_SIZE, kApi, "role"); rc != 0) {
    return rc;
  }
  const std::vector<formulon::ChartValue>& points = source_of(*s, role)->literal;
  if (auto rc = check_index(idx, points.size(), kApi, "idx"); rc != 0) {
    return rc;
  }
  const formulon::ChartValue& v = points[idx];
  const bool text = v.kind == formulon::ChartValue::Kind::kText;
  *out = fm_chart_literal_t{v.idx, text ? 1 : 0, text ? 0.0 : v.number, v.text.c_str()};
  return 0;
}

extern "C" fm_status_t fm_chart_series_point_at(const fm_chart_t* chart, size_t plot, size_t series, size_t idx,
                                                fm_chart_point_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_series_point_at";
  if (chart == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  const formulon::ChartSeries* s = nullptr;
  if (auto rc = find_series(chart, plot, series, kApi, &s); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(idx, s->points.size(), kApi, "idx"); rc != 0) {
    return rc;
  }
  const formulon::ChartPoint& p = s->points[idx];
  *out = fm_chart_point_t{};
  out->idx = p.idx;
  out->fill = to_fm(p.fill);
  out->line = to_fm(p.line);
  if (p.marker) {
    out->has_marker = 1;
    out->marker = to_fm(*p.marker);
  }
  return 0;
}

extern "C" fm_status_t fm_chart_axis_at(const fm_chart_t* chart, size_t idx, fm_chart_axis_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_axis_at";
  if (chart == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  if (auto rc = check_index(idx, chart->model.axes.size(), kApi, "idx"); rc != 0) {
    return rc;
  }
  const formulon::ChartAxis& a = chart->model.axes[idx];
  *out = fm_chart_axis_t{};
  out->id = a.id;
  out->kind = static_cast<std::int32_t>(a.kind);
  out->position = static_cast<std::int32_t>(a.position);
  out->deleted = a.deleted ? 1 : 0;
  out->cross_axis = a.cross_ax;
  out->crosses = static_cast<std::int32_t>(a.crosses);
  out->crosses_at = a.crosses_at;
  out->reversed = a.reversed ? 1 : 0;
  out->log_base = a.log_base;
  out->min = a.min;
  out->max = a.max;
  out->major_unit = a.major_unit;
  out->minor_unit = a.minor_unit;
  out->number_format = a.num_fmt ? a.num_fmt->code.c_str() : nullptr;
  out->source_linked = a.num_fmt && a.num_fmt->source_linked ? 1 : 0;
  out->title = to_fm(a.title);
  out->major_gridlines = a.major_gridlines ? 1 : 0;
  out->minor_gridlines = a.minor_gridlines ? 1 : 0;
  out->tick_label_pos = nonempty_cstr(a.tick_label_pos);
  return 0;
}

// ---------------------------------------------------------------------------
// Resolved chart data
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_chart_data_open(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                fm_chart_data_t** out) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_chart_data_open";
  if (out == nullptr) {
    return null_argument(kApi);
  }
  *out = nullptr;
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  auto data = formulon::get_chart_data(wb->workbook(), sheet_index, object_id);
  if (!data) {
    return set_last_error(data.error());
  }
  *out = make_data_handle(std::move(data.value())).release();
  return 0;
}

extern "C" fm_status_t fm_sheet_chart_preview_open(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                   const fm_chart_preview_overrides* overrides, fm_chart_data_t** out) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_chart_preview_open";
  if (out != nullptr) {
    *out = nullptr;
  }
  if (out == nullptr || overrides == nullptr) {
    return null_argument(kApi);
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  PatchReader r(kApi);
  formulon::ChartPreviewOverrides model;
  if (overrides->source != nullptr) {
    model.source = std::string(overrides->source);
  }
  r.enum_optional(overrides->direction_op, overrides->direction, formulon::ChartSeriesDirection::kRows, model.direction,
                  "direction");
  r.flag(overrides->header_row_op, overrides->header_row, model.header_row, "header_row_op");
  r.flag(overrides->header_column_op, overrides->header_column, model.header_column, "header_column_op");
  r.enum_optional(overrides->blanks_as_op, overrides->blanks_as, formulon::ChartBlanksAs::kSpan, model.blanks_as,
                  "blanks_as");
  r.flag(overrides->plot_vis_only_op, overrides->plot_vis_only, model.plot_vis_only, "plot_vis_only_op");
  if (r.status() != 0) {
    return r.status();
  }
  auto data = formulon::preview_chart_data(wb->workbook(), sheet_index, object_id, model);
  if (!data) {
    return set_last_error(data.error());
  }
  *out = make_data_handle(std::move(data.value())).release();
  return 0;
}

extern "C" void fm_chart_data_destroy(fm_chart_data_t* data) {
  delete data;
}

extern "C" fm_status_t fm_chart_data_info(const fm_chart_data_t* data, fm_chart_data_info_t* out) {
  clear_last_error();
  if (data == nullptr || out == nullptr) {
    return null_argument("fm_chart_data_info");
  }
  const formulon::ChartData& d = data->data;
  *out = fm_chart_data_info_t{};
  out->is_xy = d.domain == formulon::ChartData::Domain::kXY ? 1 : 0;
  out->point_count = d.point_count;
  out->series_count = d.series.size();
  out->category_label_count = d.categories.labels.size();
  out->category_value_count = d.categories.values.size();
  out->category_level_count = d.categories.levels.size();
  out->category_format_code = d.categories.format_code.c_str();
  out->categories_generated = d.categories.generated ? 1 : 0;
  out->title_text = opt_cstr(d.title_text);
  out->axis_title_count = d.axis_title_texts.size();
  out->unresolved_count = d.unresolved.size();
  return 0;
}

extern "C" fm_status_t fm_chart_data_category_at(const fm_chart_data_t* data, size_t idx, const char** out_label,
                                                 double* out_value) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_data_category_at";
  if (data == nullptr || out_label == nullptr || out_value == nullptr) {
    return null_argument(kApi);
  }
  const formulon::ChartDataCategories& c = data->data.categories;
  if (auto rc = check_index(idx, c.labels.size(), kApi, "idx"); rc != 0) {
    return rc;
  }
  *out_label = c.labels[idx].c_str();
  *out_value = idx < c.values.size() ? c.values[idx] : kNaN;
  return 0;
}

extern "C" fm_status_t fm_chart_data_category_level_size(const fm_chart_data_t* data, size_t level, size_t* out_count) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_data_category_level_size";
  if (data == nullptr || out_count == nullptr) {
    return null_argument(kApi);
  }
  const std::vector<std::vector<std::string>>& levels = data->data.categories.levels;
  if (auto rc = check_index(level, levels.size(), kApi, "level"); rc != 0) {
    return rc;
  }
  *out_count = levels[level].size();
  return 0;
}

extern "C" fm_status_t fm_chart_data_category_level_at(const fm_chart_data_t* data, size_t level, size_t idx,
                                                       const char** out_label) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_data_category_level_at";
  if (data == nullptr || out_label == nullptr) {
    return null_argument(kApi);
  }
  const std::vector<std::vector<std::string>>& levels = data->data.categories.levels;
  if (auto rc = check_index(level, levels.size(), kApi, "level"); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(idx, levels[level].size(), kApi, "idx"); rc != 0) {
    return rc;
  }
  *out_label = levels[level][idx].c_str();
  return 0;
}

extern "C" fm_status_t fm_chart_data_series_at(const fm_chart_data_t* data, size_t series,
                                               fm_chart_data_series_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_data_series_at";
  if (data == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  if (auto rc = check_index(series, data->data.series.size(), kApi, "series"); rc != 0) {
    return rc;
  }
  const formulon::ChartDataSeries& s = data->data.series[series];
  *out = fm_chart_data_series_t{};
  out->plot_index = s.plot_index;
  out->idx = s.idx;
  out->name = opt_cstr(s.name);
  out->format_code = s.format_code.c_str();
  out->x_format_code = s.x_format_code.c_str();
  out->filtered = s.filtered ? 1 : 0;
  out->point_count = data->points[series].values.size();
  out->has_sizes = s.sizes.empty() ? 0 : 1;
  for (std::int32_t role = FM_CHART_ROLE_NAME; role <= FM_CHART_ROLE_SIZE; ++role) {
    out->range_count[role] = ranges_of(s, role)->size();
  }
  return 0;
}

extern "C" fm_status_t fm_chart_data_series_points(const fm_chart_data_t* data, size_t series,
                                                   fm_chart_data_points_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_data_series_points";
  if (data == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  if (auto rc = check_index(series, data->data.series.size(), kApi, "series"); rc != 0) {
    return rc;
  }
  const fm_chart_data::Points& p = data->points[series];
  const std::vector<double>& sizes = data->data.series[series].sizes;
  *out = fm_chart_data_points_t{};
  out->count = p.values.size();
  out->values = p.values.data();
  out->states = p.states.data();
  out->x = data->data.domain == formulon::ChartData::Domain::kXY ? p.x.data() : nullptr;
  out->sizes = sizes.empty() ? nullptr : sizes.data();
  return 0;
}

extern "C" fm_status_t fm_chart_data_range_at(const fm_chart_data_t* data, size_t series, int32_t role, size_t idx,
                                              fm_chart_sheet_range_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_data_range_at";
  if (data == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  if (auto rc = check_index(series, data->data.series.size(), kApi, "series"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(role, FM_CHART_ROLE_SIZE, kApi, "role"); rc != 0) {
    return rc;
  }
  const std::vector<formulon::SheetRect>& ranges = *ranges_of(data->data.series[series], role);
  if (auto rc = check_index(idx, ranges.size(), kApi, "idx"); rc != 0) {
    return rc;
  }
  const formulon::SheetRect& r = ranges[idx];
  *out = fm_chart_sheet_range_t{r.sheet, r.top, r.left, r.bottom, r.right};
  return 0;
}

extern "C" fm_status_t fm_chart_data_axis_title_at(const fm_chart_data_t* data, size_t idx,
                                                   fm_chart_axis_title_text_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_data_axis_title_at";
  if (data == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  const std::vector<formulon::ChartAxisTitleText>& list = data->data.axis_title_texts;
  if (auto rc = check_index(idx, list.size(), kApi, "idx"); rc != 0) {
    return rc;
  }
  *out = fm_chart_axis_title_text_t{list[idx].axis_id, list[idx].text.c_str()};
  return 0;
}

extern "C" fm_status_t fm_chart_data_unresolved_at(const fm_chart_data_t* data, size_t idx,
                                                   fm_chart_unresolved_t* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_chart_data_unresolved_at";
  if (data == nullptr || out == nullptr) {
    return null_argument(kApi);
  }
  const std::vector<formulon::ChartUnresolved>& list = data->data.unresolved;
  if (auto rc = check_index(idx, list.size(), kApi, "idx"); rc != 0) {
    return rc;
  }
  *out = fm_chart_unresolved_t{list[idx].series_idx, list[idx].source.c_str(), list[idx].axis_id};
  return 0;
}

// ---------------------------------------------------------------------------
// Creation and edits
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_insert_chart(fm_workbook_t* wb, size_t sheet_index, const fm_chart_insert* opts,
                                             uint32_t* out_object_id) {
  constexpr const char* kApi = "fm_sheet_insert_chart";
  if (auto rc = edit_entry(wb, sheet_index, opts, kApi); rc != 0) {
    return rc;
  }
  if (opts->source == nullptr || out_object_id == nullptr) {
    return null_argument(kApi);
  }
  PatchReader r(kApi);
  r.enum_ok(opts->kind, formulon::ChartInsertKind::kDoughnut, "kind");
  r.enum_ok(opts->grouping, formulon::ChartGrouping::kStandard, "grouping");
  r.enum_ok(opts->direction, formulon::ChartSeriesDirection::kRows, "direction");
  r.enum_ok(opts->header_row, formulon::ChartHeaderMode::kNo, "header_row");
  r.enum_ok(opts->header_column, formulon::ChartHeaderMode::kNo, "header_column");
  r.enum_ok(opts->legend, formulon::ChartLegendPosition::kTopRight, "legend");
  r.enum_ok(opts->placement.anchor_kind, formulon::AnchorKind::kAbsolute, "placement.anchor_kind");
  r.enum_ok(opts->placement.edit_as, formulon::EditAs::kAbsolute, "placement.edit_as");
  if (r.status() != 0) {
    return r.status();
  }
  formulon::ChartInsertOptions model;
  model.kind = static_cast<formulon::ChartInsertKind>(opts->kind);
  model.grouping = static_cast<formulon::ChartGrouping>(opts->grouping);
  model.source = opts->source;
  model.direction = static_cast<formulon::ChartSeriesDirection>(opts->direction);
  model.header_row = static_cast<formulon::ChartHeaderMode>(opts->header_row);
  model.header_column = static_cast<formulon::ChartHeaderMode>(opts->header_column);
  if (opts->title != nullptr) {
    model.title = std::string(opts->title);
  }
  model.legend = static_cast<formulon::ChartLegendPosition>(opts->legend);
  const fm_image_insert& p = opts->placement;
  model.placement.anchor_kind = static_cast<formulon::AnchorKind>(p.anchor_kind);
  model.placement.edit_as = static_cast<formulon::EditAs>(p.edit_as);
  model.placement.row = p.row;
  model.placement.col = p.col;
  model.placement.row_off = p.row_off_emu;
  model.placement.col_off = p.col_off_emu;
  model.placement.width_emu = p.width_emu;
  model.placement.height_emu = p.height_emu;
  model.name = p.name != nullptr ? p.name : "";
  model.descr = p.descr != nullptr ? p.descr : "";
  auto id = formulon::insert_chart(wb->workbook(), sheet_index, model);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_set_chart_type(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                               int32_t kind, int32_t grouping) {
  constexpr const char* kApi = "fm_sheet_set_chart_type";
  if (auto rc = edit_entry(wb, sheet_index, wb, kApi); rc != 0) {
    return rc;
  }
  PatchReader r(kApi);
  r.enum_ok(kind, formulon::ChartInsertKind::kDoughnut, "kind");
  r.enum_ok(grouping, formulon::ChartGrouping::kStandard, "grouping");
  if (r.status() != 0) {
    return r.status();
  }
  return finish(formulon::set_chart_type(wb->workbook(), sheet_index, object_id, plot,
                                         static_cast<formulon::ChartInsertKind>(kind),
                                         static_cast<formulon::ChartGrouping>(grouping)));
}

extern "C" fm_status_t fm_sheet_set_chart_title(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                const fm_chart_title_t* title) {
  constexpr const char* kApi = "fm_sheet_set_chart_title";
  if (auto rc = edit_entry(wb, sheet_index, title, kApi); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(title->kind, FM_CHART_TITLE_AUTO, kApi, "kind"); rc != 0) {
    return rc;
  }
  formulon::ChartTitle model;
  model.kind = static_cast<formulon::ChartTitleKind>(title->kind);
  model.text = title->text != nullptr ? title->text : "";
  model.ref_formula = title->formula != nullptr ? title->formula : "";
  return finish(formulon::set_chart_title(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_set_chart_legend(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                 const fm_chart_legend_patch* legend) {
  constexpr const char* kApi = "fm_sheet_set_chart_legend";
  if (auto rc = edit_entry(wb, sheet_index, legend, kApi); rc != 0) {
    return rc;
  }
  if (legend->hidden_series == nullptr && legend->hidden_series_count != 0) {
    return null_argument(kApi);
  }
  PatchReader r(kApi);
  r.enum_ok(legend->position, formulon::ChartLegendPosition::kTopRight, "position");
  formulon::ChartLegendPatch model;
  model.position = static_cast<formulon::ChartLegendPosition>(legend->position);
  model.overlay = legend->overlay != 0;
  std::vector<std::uint32_t> hidden;
  if (legend->hidden_series_count != 0) {
    hidden.assign(legend->hidden_series, legend->hidden_series + legend->hidden_series_count);
  }
  r.optional(legend->hidden_series_op, std::move(hidden), model.hidden_series_idx, "hidden_series_op");
  if (r.status() != 0) {
    return r.status();
  }
  return finish(formulon::set_chart_legend(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_set_chart_axis(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                               uint32_t axis_id, const fm_chart_axis_patch* patch) {
  constexpr const char* kApi = "fm_sheet_set_chart_axis";
  if (auto rc = edit_entry(wb, sheet_index, patch, kApi); rc != 0) {
    return rc;
  }
  PatchReader r(kApi);
  formulon::ChartAxisPatch model;
  r.flag(patch->visible_op, patch->visible, model.visible, "visible_op");
  r.text_field(patch->title_op, patch->title, model.title, "title_op");
  r.field(patch->min_op, patch->min, model.min, "min_op");
  r.field(patch->max_op, patch->max, model.max, "max_op");
  r.field(patch->major_unit_op, patch->major_unit, model.major_unit, "major_unit_op");
  r.text_field(patch->number_format_op, patch->number_format, model.number_format, "number_format_op");
  r.flag(patch->major_gridlines_op, patch->major_gridlines, model.major_gridlines, "major_gridlines_op");
  r.flag(patch->reversed_op, patch->reversed, model.reversed, "reversed_op");
  r.field(patch->log_base_op, patch->log_base, model.log_base, "log_base_op");
  r.enum_optional(patch->crosses_op, patch->crosses, formulon::ChartAxisCrosses::kAt, model.crosses, "crosses");
  model.crosses_at = patch->crosses_at;
  r.text_field(patch->tick_label_pos_op, patch->tick_label_pos, model.tick_label_pos, "tick_label_pos_op");
  if (r.status() != 0) {
    return r.status();
  }
  return finish(formulon::set_chart_axis(wb->workbook(), sheet_index, object_id, axis_id, model));
}

extern "C" fm_status_t fm_sheet_set_chart_display(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                  int32_t blanks_as, int32_t plot_vis_only) {
  constexpr const char* kApi = "fm_sheet_set_chart_display";
  if (auto rc = edit_entry(wb, sheet_index, wb, kApi); rc != 0) {
    return rc;
  }
  PatchReader r(kApi);
  r.enum_ok(blanks_as, formulon::ChartBlanksAs::kSpan, "blanks_as");
  r.enum_ok(plot_vis_only, formulon::Tri::kTrue, "plot_vis_only");
  if (r.status() != 0) {
    return r.status();
  }
  return finish(formulon::set_chart_display(wb->workbook(), sheet_index, object_id,
                                            static_cast<formulon::ChartBlanksAs>(blanks_as),
                                            static_cast<formulon::Tri>(plot_vis_only)));
}

extern "C" fm_status_t fm_sheet_set_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                                 uint32_t series_idx, const fm_chart_series_patch* patch) {
  constexpr const char* kApi = "fm_sheet_set_chart_series";
  if (auto rc = edit_entry(wb, sheet_index, patch, kApi); rc != 0) {
    return rc;
  }
  PatchReader r(kApi);
  const formulon::ChartSeriesPatch model = read_series(r, *patch);
  if (r.status() != 0) {
    return r.status();
  }
  return finish(formulon::set_chart_series(wb->workbook(), sheet_index, object_id, plot, series_idx, model));
}

extern "C" fm_status_t fm_sheet_set_chart_point(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                                uint32_t series_idx, uint32_t point_idx,
                                                const fm_chart_point_patch* patch) {
  constexpr const char* kApi = "fm_sheet_set_chart_point";
  if (auto rc = edit_entry(wb, sheet_index, patch, kApi); rc != 0) {
    return rc;
  }
  PatchReader r(kApi);
  formulon::ChartPointPatch model;
  model.fill = read_fill(r, patch->fill);
  model.line = read_line(r, patch->line);
  r.field(patch->marker_op,
          patch->marker_op == FM_CHART_PATCH_SET ? read_marker(r, patch->marker) : formulon::ChartMarkerPatch{},
          model.marker, "marker_op");
  if (r.status() != 0) {
    return r.status();
  }
  return finish(formulon::set_chart_point(wb->workbook(), sheet_index, object_id, plot, series_idx, point_idx, model));
}

extern "C" fm_status_t fm_sheet_set_chart_plot(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                               const fm_chart_plot_patch* patch) {
  constexpr const char* kApi = "fm_sheet_set_chart_plot";
  if (auto rc = edit_entry(wb, sheet_index, patch, kApi); rc != 0) {
    return rc;
  }
  PatchReader r(kApi);
  formulon::ChartPlotPatch model;
  r.optional(patch->gap_width_op, patch->gap_width, model.gap_width, "gap_width_op");
  r.optional(patch->overlap_op, patch->overlap, model.overlap, "overlap_op");
  r.optional(patch->first_slice_angle_op, patch->first_slice_angle, model.first_slice_angle, "first_slice_angle_op");
  r.flag(patch->vary_colors_op, patch->vary_colors, model.vary_colors, "vary_colors_op");
  r.optional(patch->hole_size_op, patch->hole_size, model.hole_size, "hole_size_op");
  if (r.status() != 0) {
    return r.status();
  }
  return finish(formulon::set_chart_plot(wb->workbook(), sheet_index, object_id, plot, model));
}

extern "C" fm_status_t fm_sheet_set_chart_filter(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                 const uint32_t* hidden_series, size_t hidden_series_count,
                                                 const uint32_t* hidden_categories, size_t hidden_category_count) {
  constexpr const char* kApi = "fm_sheet_set_chart_filter";
  if (auto rc = edit_entry(wb, sheet_index, wb, kApi); rc != 0) {
    return rc;
  }
  if ((hidden_series == nullptr && hidden_series_count != 0) ||
      (hidden_categories == nullptr && hidden_category_count != 0)) {
    return null_argument(kApi);
  }
  std::vector<std::uint32_t> series;
  std::vector<std::uint32_t> categories;
  if (hidden_series_count != 0) {
    series.assign(hidden_series, hidden_series + hidden_series_count);
  }
  if (hidden_category_count != 0) {
    categories.assign(hidden_categories, hidden_categories + hidden_category_count);
  }
  return finish(formulon::set_chart_filter(wb->workbook(), sheet_index, object_id, series, categories));
}

extern "C" fm_status_t fm_sheet_add_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                                 const fm_chart_series_patch* patch, uint32_t* out_series_idx) {
  constexpr const char* kApi = "fm_sheet_add_chart_series";
  if (auto rc = edit_entry(wb, sheet_index, patch, kApi); rc != 0) {
    return rc;
  }
  if (out_series_idx == nullptr) {
    return null_argument(kApi);
  }
  PatchReader r(kApi);
  const formulon::ChartSeriesPatch model = read_series(r, *patch);
  if (r.status() != 0) {
    return r.status();
  }
  auto idx = formulon::add_chart_series(wb->workbook(), sheet_index, object_id, plot, model);
  if (!idx) {
    return set_last_error(idx.error());
  }
  *out_series_idx = idx.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                    size_t plot, uint32_t series_idx) {
  constexpr const char* kApi = "fm_sheet_remove_chart_series";
  if (auto rc = edit_entry(wb, sheet_index, wb, kApi); rc != 0) {
    return rc;
  }
  return finish(formulon::remove_chart_series(wb->workbook(), sheet_index, object_id, plot, series_idx));
}

extern "C" fm_status_t fm_sheet_set_chart_source(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                 const char* source, int32_t direction, int32_t header_row,
                                                 int32_t header_column) {
  constexpr const char* kApi = "fm_sheet_set_chart_source";
  if (auto rc = edit_entry(wb, sheet_index, source, kApi); rc != 0) {
    return rc;
  }
  PatchReader r(kApi);
  r.enum_ok(direction, formulon::ChartSeriesDirection::kRows, "direction");
  r.enum_ok(header_row, formulon::ChartHeaderMode::kNo, "header_row");
  r.enum_ok(header_column, formulon::ChartHeaderMode::kNo, "header_column");
  if (r.status() != 0) {
    return r.status();
  }
  return finish(formulon::set_chart_source(
      wb->workbook(), sheet_index, object_id, source, static_cast<formulon::ChartSeriesDirection>(direction),
      static_cast<formulon::ChartHeaderMode>(header_row), static_cast<formulon::ChartHeaderMode>(header_column)));
}
