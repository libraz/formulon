//
// C ABI - charts: the typed chart view and the resolved series data behind
// owned handles. Creation and the typed edits live in `chart_mutate.cpp`.
//
// Both handles copy the C++ result at open time and hand out pointers into
// that copy, so every string and array stays valid until the handle is
// destroyed.

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/chart_patch_reader.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "chart/chart_edit.h"
#include "chart/chart_model.h"
#include "drawing/drawing_xml.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::chart_blanks_as_from_c;
using formulon::c_api::parts::chart_fill_to_c;
using formulon::c_api::parts::chart_line_to_c;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::color_spec_to_c;
using formulon::c_api::parts::ensure_calculated;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::kTickLabelPositionTokens;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::ordinal_of;
using formulon::c_api::parts::PatchReader;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::to_c;

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

constexpr double kNaN = std::numeric_limits<double>::quiet_NaN();

const char* opt_cstr(const std::optional<std::string>& s) {
  return s ? s->c_str() : nullptr;
}

std::int32_t tri(formulon::Tri t) {
  return static_cast<std::int32_t>(t);
}

// Colour of a chart fill, line or label; None when absent.
fm_color_t chart_color_to_c(const std::optional<formulon::ColorSpec>& color) {
  return color_spec_to_c(color ? *color : formulon::ColorSpec{});
}

// Validates the caller's `struct_size`, then zeroes `*out` keeping it.
template <typename T>
fm_status_t begin_record(T* out) {
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const std::uint32_t struct_size = out->struct_size;
  *out = T{};
  out->struct_size = struct_size;
  return 0;
}

// Writes an optional model number (NaN when absent) as a has_ flag and its value (0 when absent).
void put_optional(double value, std::int32_t& has, double& out) {
  has = std::isnan(value) ? 0 : 1;
  out = std::isnan(value) ? 0.0 : value;
}

// ---------------------------------------------------------------------------
// Model -> C
// ---------------------------------------------------------------------------

fm_chart_title_t to_fm(const formulon::ChartTitle& t) {
  const bool is_text = t.kind == formulon::ChartTitleKind::kText;
  const bool is_ref = t.kind == formulon::ChartTitleKind::kRef;
  return fm_chart_title_t{to_c(t.kind), is_text ? t.text.c_str() : nullptr, is_ref ? t.ref_formula.c_str() : nullptr};
}

fm_chart_fill_t to_fm(const formulon::ChartFill& f) {
  return chart_fill_to_c(f);
}

fm_chart_line_t to_fm(const formulon::ChartLine& l) {
  return chart_line_to_c(l);
}

fm_chart_marker_t to_fm(const formulon::ChartMarker& m) {
  fm_chart_marker_t out{};
  out.symbol = static_cast<std::int32_t>(m.symbol);
  out.has_size_pt = m.size ? 1 : 0;
  out.size_pt = m.size.value_or(0U);
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
  out.font_color = chart_color_to_c(l.font_color);
  out.has_font_size_pt = l.font_size_hundredths_pt ? 1 : 0;
  out.font_size_pt = l.font_size_hundredths_pt ? static_cast<double>(*l.font_size_hundredths_pt) / 100.0 : 0.0;
  out.deleted = l.deleted ? 1 : 0;
  return out;
}

fm_chart_data_source_t to_fm(const formulon::ChartDataSource& d) {
  return fm_chart_data_source_t{static_cast<std::int32_t>(d.kind), d.formula.c_str(),
                                static_cast<std::uint32_t>(d.literal.size())};
}

/// Respells every formula `model` carries as the workbook's getters hand formulas out.
void spell_formulas(const formulon::Workbook& book, formulon::ChartModel& model) {
  const auto spell = [&book](std::string& text) {
    if (book.spells_self_book_names(text)) {
      text = book.spell_self_book_names(text);
    }
  };
  spell(model.title.ref_formula);
  for (formulon::ChartAxis& axis : model.axes) {
    spell(axis.title.ref_formula);
  }
  for (formulon::ChartPlot& plot : model.plots) {
    for (formulon::ChartSeries& series : plot.series) {
      for (formulon::ChartDataSource* source :
           {&series.name, &series.cat, &series.val, &series.x, &series.y, &series.size}) {
        spell(source->formula);
      }
    }
  }
}

/// Extended-chart layout of `p`; ABSENT for a plot that is not extended or has an unknown layout.
std::int32_t extended_layout_from(const formulon::ChartPlot& p) {
  if (p.type != formulon::ChartPlotType::kChartex) {
    return FM_CHART_EXTENDED_LAYOUT_ABSENT;
  }
  static const char* const kLayouts[] = {"boxWhisker", "clusteredColumn", "funnel",  "paretoLine",
                                         "regionMap",  "sunburst",        "treemap", "waterfall"};
  for (std::int32_t i = 0; i < 8; ++i) {
    if (p.chartex_layout == kLayouts[i]) {
      return i;
    }
  }
  return FM_CHART_EXTENDED_LAYOUT_ABSENT;
}

/// Role of an unresolved input, from the engine's input label.
std::int32_t unresolved_role(const std::string& source) {
  static const struct {
    const char* label;
    std::int32_t role;
  } kRoles[] = {{"name", FM_CHART_ROLE_NAME},   {"cat", FM_CHART_ROLE_CATEGORIES},
                {"val", FM_CHART_ROLE_VALUES},  {"x", FM_CHART_ROLE_X},
                {"y", FM_CHART_ROLE_Y},         {"size", FM_CHART_ROLE_SIZE},
                {"title", FM_CHART_ROLE_TITLE}, {"axis_title", FM_CHART_ROLE_AXIS_TITLE}};
  for (const auto& r : kRoles) {
    if (source == r.label) {
      return r.role;
    }
  }
  return FM_CHART_ROLE_VALUES;
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
fm_status_t find_series(const fm_chart_t* chart, std::size_t plot, std::size_t series,
                        const formulon::ChartSeries** out) {
  if (auto rc = check_index(plot, chart->model.plots.size(), "plot"); rc != 0) {
    return rc;
  }
  const std::vector<formulon::ChartSeries>& list = chart->model.plots[plot].series;
  if (auto rc = check_index(series, list.size(), "series"); rc != 0) {
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

}  // namespace

// ---------------------------------------------------------------------------
// Chart model reads
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_get_chart(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                          int32_t* out_present, fm_chart_t** out) {
  clear_last_error();
  if (out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  *out_present = 0;
  *out = nullptr;
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto model = formulon::get_chart(wb->workbook(), sheet_index, object_id);
  // An unknown id is the only not-found outcome of a chart read.
  if (!model && model.error().code == formulon::FormulonErrorCode::kNotFound) {
    return 0;
  }
  if (!model) {
    return set_last_error(model.error());
  }
  auto handle = std::unique_ptr<fm_chart_t>(new fm_chart_t{});
  handle->model = std::move(model.value());
  spell_formulas(wb->workbook(), handle->model);
  *out = handle.release();
  *out_present = 1;
  return 0;
}

extern "C" void fm_chart_destroy(fm_chart_t* chart) {
  delete chart;
}

extern "C" fm_status_t fm_chart_get_info(const fm_chart_t* chart, fm_chart_info_t* out) {
  clear_last_error();
  if (chart == nullptr || out == nullptr) {
    return null_argument();
  }
  const formulon::ChartModel& m = chart->model;
  if (auto rc = begin_record(out); rc != 0) {
    return rc;
  }
  out->object_id = m.object_id;
  out->is_extended = m.is_chartex ? 1 : 0;
  out->is_pivot_chart = m.is_pivot_chart ? 1 : 0;
  out->on_chartsheet = m.on_chartsheet ? 1 : 0;
  out->title = to_fm(m.title);
  out->auto_title_deleted = tri(m.auto_title_deleted);
  out->legend_position =
      m.legend.present ? to_c(m.legend.position) : static_cast<std::int32_t>(FM_CHART_LEGEND_POSITION_ABSENT);
  out->legend_overlay = m.legend.overlay ? 1 : 0;
  out->legend_hidden_count = static_cast<std::uint32_t>(m.legend.hidden_series_idx.size());
  out->blanks_as = to_c(m.disp_blanks_as);
  out->plot_visible_only = tri(m.plot_vis_only);
  out->date1904 = m.date1904 ? 1 : 0;
  out->has_style = m.style ? 1 : 0;
  out->style = m.style.value_or(0);
  out->plot_count = static_cast<std::uint32_t>(m.plots.size());
  out->axis_count = static_cast<std::uint32_t>(m.axes.size());
  out->hidden_category_count = static_cast<std::uint32_t>(m.category_filter.size());
  if (m.source) {
    out->source_range = m.source->range.c_str();
    out->source_direction = static_cast<std::int32_t>(m.source->direction);
    out->source_header_row = m.source->header_row ? 1 : 0;
    out->source_header_column = m.source->header_column ? 1 : 0;
  }
  out->na_as_blank = tri(m.disp_na_as_blank);
  return 0;
}

namespace {

// Reads `list[idx]` into `*out` after the NULL and range checks shared by the index-list accessors.
fm_status_t index_list_at(const fm_chart_t* chart, const std::vector<std::uint32_t>* list, size_t idx,
                          std::uint32_t* out) {
  if (chart == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(idx, list->size(), "index"); rc != 0) {
    return rc;
  }
  *out = (*list)[idx];
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_chart_get_legend_hidden_at(const fm_chart_t* chart, size_t idx, uint32_t* out) {
  clear_last_error();
  return index_list_at(chart, chart != nullptr ? &chart->model.legend.hidden_series_idx : nullptr, idx, out);
}

extern "C" fm_status_t fm_chart_get_hidden_category_at(const fm_chart_t* chart, size_t idx, uint32_t* out) {
  clear_last_error();
  return index_list_at(chart, chart != nullptr ? &chart->model.category_filter : nullptr, idx, out);
}

extern "C" fm_status_t fm_chart_get_plot_at(const fm_chart_t* chart, size_t plot_index, fm_chart_plot_info_t* out) {
  clear_last_error();
  if (chart == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(plot_index, chart->model.plots.size(), "plot_index"); rc != 0) {
    return rc;
  }
  const formulon::ChartPlot& p = chart->model.plots[plot_index];
  if (auto rc = begin_record(out); rc != 0) {
    return rc;
  }
  out->type = static_cast<std::int32_t>(p.type);
  out->extended_layout = extended_layout_from(p);
  out->bar_direction = to_c(p.bar_dir);
  out->has_grouping = p.grouping ? 1 : 0;
  out->grouping = p.grouping ? static_cast<std::int32_t>(*p.grouping) : 0;
  out->vary_colors = tri(p.vary_colors);
  out->has_gap_width_pct = p.gap_width ? 1 : 0;
  out->gap_width_pct = p.gap_width.value_or(0U);
  out->has_overlap_pct = p.overlap ? 1 : 0;
  out->overlap_pct = p.overlap.value_or(0);
  out->has_hole_size_pct = p.hole_size ? 1 : 0;
  out->hole_size_pct = p.hole_size.value_or(0U);
  out->has_first_slice_angle_deg = p.first_slice_angle ? 1 : 0;
  out->first_slice_angle_deg = p.first_slice_angle.value_or(0U);
  out->axis_id_count = static_cast<std::uint32_t>(p.axis_ids.size());
  out->series_count = static_cast<std::uint32_t>(p.series.size());
  return 0;
}

extern "C" fm_status_t fm_chart_get_plot_axis_id_at(const fm_chart_t* chart, size_t plot_index, size_t idx,
                                                    uint32_t* out) {
  clear_last_error();
  if (chart == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(plot_index, chart->model.plots.size(), "plot_index"); rc != 0) {
    return rc;
  }
  const std::vector<std::uint32_t>& ids = chart->model.plots[plot_index].axis_ids;
  if (auto rc = check_index(idx, ids.size(), "index"); rc != 0) {
    return rc;
  }
  *out = ids[idx];
  return 0;
}

extern "C" fm_status_t fm_chart_get_series_at(const fm_chart_t* chart, size_t plot_index, size_t series,
                                              fm_chart_series_info_t* out) {
  clear_last_error();
  if (chart == nullptr || out == nullptr) {
    return null_argument();
  }
  const formulon::ChartSeries* s = nullptr;
  if (auto rc = find_series(chart, plot_index, series, &s); rc != 0) {
    return rc;
  }
  if (auto rc = begin_record(out); rc != 0) {
    return rc;
  }
  out->series_id = s->idx;
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
  out->point_count = static_cast<std::uint32_t>(s->points.size());
  if (s->labels) {
    out->has_labels = 1;
    out->labels = to_fm(*s->labels);
  }
  return 0;
}

extern "C" fm_status_t fm_chart_get_series_literal_at(const fm_chart_t* chart, size_t plot_index, size_t series,
                                                      int32_t role, size_t idx, fm_chart_literal_info_t* out) {
  clear_last_error();
  if (chart == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const formulon::ChartSeries* s = nullptr;
  if (auto rc = find_series(chart, plot_index, series, &s); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(role, 0, FM_CHART_ROLE_SIZE, "role"); rc != 0) {
    return rc;
  }
  const std::vector<formulon::ChartValue>& points = source_of(*s, role)->literal;
  if (auto rc = check_index(idx, points.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::ChartValue& v = points[idx];
  const bool text = v.kind == formulon::ChartValue::Kind::kText;
  out->idx = v.idx;
  out->is_text = text ? 1 : 0;
  out->number = text ? 0.0 : v.number;
  out->text = v.text.c_str();
  return 0;
}

extern "C" fm_status_t fm_chart_get_series_point_at(const fm_chart_t* chart, size_t plot_index, size_t series,
                                                    size_t idx, fm_chart_point_info_t* out) {
  clear_last_error();
  if (chart == nullptr || out == nullptr) {
    return null_argument();
  }
  const formulon::ChartSeries* s = nullptr;
  if (auto rc = find_series(chart, plot_index, series, &s); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(idx, s->points.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::ChartPoint& p = s->points[idx];
  if (auto rc = begin_record(out); rc != 0) {
    return rc;
  }
  out->idx = p.idx;
  out->fill = to_fm(p.fill);
  out->line = to_fm(p.line);
  if (p.marker) {
    out->has_marker = 1;
    out->marker = to_fm(*p.marker);
  }
  return 0;
}

extern "C" fm_status_t fm_chart_get_axis_at(const fm_chart_t* chart, size_t idx, fm_chart_axis_info_t* out) {
  clear_last_error();
  if (chart == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(idx, chart->model.axes.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::ChartAxis& a = chart->model.axes[idx];
  if (auto rc = begin_record(out); rc != 0) {
    return rc;
  }
  out->id = a.id;
  out->kind = static_cast<std::int32_t>(a.kind);
  out->position = to_c(a.position);
  out->visible = a.deleted ? 0 : 1;
  out->cross_axis = a.cross_ax;
  out->crosses = to_c(a.crosses);
  put_optional(a.crosses_at, out->has_crosses_at, out->crosses_at);
  out->reversed = a.reversed ? 1 : 0;
  put_optional(a.log_base, out->has_log_base, out->log_base);
  put_optional(a.min, out->has_min, out->min);
  put_optional(a.max, out->has_max, out->max);
  put_optional(a.major_unit, out->has_major_unit, out->major_unit);
  put_optional(a.minor_unit, out->has_minor_unit, out->minor_unit);
  out->number_format = a.num_fmt ? a.num_fmt->code.c_str() : nullptr;
  out->source_linked = a.num_fmt && a.num_fmt->source_linked ? 1 : 0;
  out->title = to_fm(a.title);
  out->major_gridlines = a.major_gridlines ? 1 : 0;
  out->minor_gridlines = a.minor_gridlines ? 1 : 0;
  out->tick_label_position = a.tick_label_pos.empty() ? FM_CHART_TICK_LABEL_POSITION_ABSENT
                                                      : ordinal_of(kTickLabelPositionTokens, a.tick_label_pos,
                                                                   FM_CHART_TICK_LABEL_POSITION_ABSENT);
  return 0;
}

// ---------------------------------------------------------------------------
// Resolved chart data
// ---------------------------------------------------------------------------

namespace {

// Sets `*present` to whether the sheet's drawing has a chart frame `object_id`; a non-zero return is the
// drawing's own status when it does not parse.
fm_status_t find_chart(const fm_workbook_t* wb, size_t sheet_index, std::uint32_t object_id, bool* present) {
  auto parts = formulon::list_chart_parts(wb->workbook(), sheet_index);
  if (!parts) {
    return set_last_error(parts.error());
  }
  const std::vector<formulon::ChartPartRef>& list = parts.value();
  *present = std::any_of(list.begin(), list.end(),
                         [object_id](const formulon::ChartPartRef& p) { return p.object_id == object_id; });
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_chart_data(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                               int32_t* out_present, fm_chart_data_t** out) {
  clear_last_error();
  if (out != nullptr) {
    *out = nullptr;
  }
  if (out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  *out_present = 0;
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  bool present = false;
  if (auto rc = find_chart(wb, sheet_index, object_id, &present); rc != 0 || !present) {
    return rc;
  }
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  auto data = formulon::get_chart_data(wb->workbook(), sheet_index, object_id);
  if (!data) {
    return set_last_error(data.error());
  }
  *out = make_data_handle(std::move(data.value())).release();
  *out_present = 1;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_chart_preview(const fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                  const fm_chart_preview_overrides_spec_t* overrides,
                                                  int32_t* out_present, fm_chart_data_t** out) {
  clear_last_error();
  if (out != nullptr) {
    *out = nullptr;
  }
  if (out_present == nullptr || out == nullptr || overrides == nullptr) {
    return null_argument();
  }
  *out_present = 0;
  if (auto rc = check_record_size(*overrides); rc != 0) {
    return rc;
  }
  if (overrides->source_op == FM_PATCH_OP_SET && overrides->source == nullptr) {
    return invalid_argument("SET source is NULL", "field=source");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  bool present = false;
  if (auto rc = find_chart(wb, sheet_index, object_id, &present); rc != 0 || !present) {
    return rc;
  }
  PatchReader r;
  formulon::ChartPreviewOverrides model;
  if (r.text_ok(overrides->source_op == FM_PATCH_OP_SET ? overrides->source : nullptr, "source")) {
    r.optional(overrides->source_op, text_or_empty(overrides->source), model.source, "source");
  }
  r.enum_optional(overrides->direction_op, overrides->direction, formulon::ChartSeriesDirection::kRows, model.direction,
                  "direction");
  r.flag(overrides->header_row_op, overrides->header_row, model.header_row, "header_row_op");
  r.flag(overrides->header_column_op, overrides->header_column, model.header_column, "header_column_op");
  r.enum_optional(overrides->blanks_as_op, overrides->blanks_as, FM_BLANKS_AS_ABSENT, FM_BLANKS_AS_SPAN,
                  chart_blanks_as_from_c, model.blanks_as, "blanks_as");
  r.flag(overrides->plot_visible_only_op, overrides->plot_visible_only, model.plot_vis_only, "plot_visible_only_op");
  r.flag(overrides->na_as_blank_op, overrides->na_as_blank, model.na_as_blank, "na_as_blank_op");
  if (r.status() != 0) {
    return r.status();
  }
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  auto data = formulon::preview_chart_data(wb->workbook(), sheet_index, object_id, model);
  if (!data) {
    return set_last_error(data.error());
  }
  *out = make_data_handle(std::move(data.value())).release();
  *out_present = 1;
  return 0;
}

extern "C" void fm_chart_data_destroy(fm_chart_data_t* data) {
  delete data;
}

extern "C" fm_status_t fm_chart_data_get_info(const fm_chart_data_t* data, fm_chart_data_info_t* out) {
  clear_last_error();
  if (data == nullptr || out == nullptr) {
    return null_argument();
  }
  const formulon::ChartData& d = data->data;
  if (auto rc = begin_record(out); rc != 0) {
    return rc;
  }
  out->is_xy = d.domain == formulon::ChartData::Domain::kXY ? 1 : 0;
  out->point_count = static_cast<std::uint32_t>(d.point_count);
  out->series_count = static_cast<std::uint32_t>(d.series.size());
  out->category_label_count = static_cast<std::uint32_t>(d.categories.labels.size());
  out->category_value_count = static_cast<std::uint32_t>(d.categories.values.size());
  out->category_level_count = static_cast<std::uint32_t>(d.categories.levels.size());
  out->category_number_format = d.categories.format_code.c_str();
  out->categories_generated = d.categories.generated ? 1 : 0;
  out->title_text = opt_cstr(d.title_text);
  out->axis_title_count = static_cast<std::uint32_t>(d.axis_title_texts.size());
  out->unresolved_count = static_cast<std::uint32_t>(d.unresolved.size());
  return 0;
}

extern "C" fm_status_t fm_chart_data_get_category_at(const fm_chart_data_t* data, size_t idx, const char** out_label,
                                                     int32_t* out_has_value, double* out_value) {
  clear_last_error();
  if (data == nullptr || out_label == nullptr || out_has_value == nullptr || out_value == nullptr) {
    return null_argument();
  }
  const formulon::ChartDataCategories& c = data->data.categories;
  if (auto rc = check_index(idx, c.labels.size(), "index"); rc != 0) {
    return rc;
  }
  *out_label = c.labels[idx].c_str();
  put_optional(idx < c.values.size() ? c.values[idx] : kNaN, *out_has_value, *out_value);
  return 0;
}

extern "C" fm_status_t fm_chart_data_get_category_level_count(const fm_chart_data_t* data, size_t level,
                                                              size_t* out_count) {
  clear_last_error();
  if (data == nullptr || out_count == nullptr) {
    return null_argument();
  }
  const std::vector<std::vector<std::string>>& levels = data->data.categories.levels;
  if (auto rc = check_index(level, levels.size(), "level"); rc != 0) {
    return rc;
  }
  *out_count = levels[level].size();
  return 0;
}

extern "C" fm_status_t fm_chart_data_get_category_level_at(const fm_chart_data_t* data, size_t level, size_t idx,
                                                           const char** out_label) {
  clear_last_error();
  if (data == nullptr || out_label == nullptr) {
    return null_argument();
  }
  const std::vector<std::vector<std::string>>& levels = data->data.categories.levels;
  if (auto rc = check_index(level, levels.size(), "level"); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(idx, levels[level].size(), "index"); rc != 0) {
    return rc;
  }
  *out_label = levels[level][idx].c_str();
  return 0;
}

extern "C" fm_status_t fm_chart_data_get_series_at(const fm_chart_data_t* data, size_t series,
                                                   fm_chart_data_series_info_t* out) {
  clear_last_error();
  if (data == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(series, data->data.series.size(), "series"); rc != 0) {
    return rc;
  }
  const formulon::ChartDataSeries& s = data->data.series[series];
  if (auto rc = begin_record(out); rc != 0) {
    return rc;
  }
  out->plot_index = static_cast<std::uint32_t>(s.plot_index);
  out->series_id = s.idx;
  out->name = opt_cstr(s.name);
  out->number_format = s.format_code.c_str();
  out->x_number_format = s.x_format_code.c_str();
  out->filtered = s.filtered ? 1 : 0;
  out->point_count = static_cast<std::uint32_t>(data->points[series].values.size());
  out->has_sizes = s.sizes.empty() ? 0 : 1;
  for (std::int32_t role = FM_CHART_ROLE_NAME; role <= FM_CHART_ROLE_SIZE; ++role) {
    out->range_count[role] = static_cast<std::uint32_t>(ranges_of(s, role)->size());
  }
  return 0;
}

extern "C" fm_status_t fm_chart_data_get_series_points(const fm_chart_data_t* data, size_t series,
                                                       fm_chart_data_points_info_t* out) {
  clear_last_error();
  if (data == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(series, data->data.series.size(), "series"); rc != 0) {
    return rc;
  }
  const fm_chart_data::Points& p = data->points[series];
  const std::vector<double>& sizes = data->data.series[series].sizes;
  if (auto rc = begin_record(out); rc != 0) {
    return rc;
  }
  out->count = static_cast<std::uint32_t>(p.values.size());
  out->values = p.values.data();
  out->states = p.states.data();
  out->x = data->data.domain == formulon::ChartData::Domain::kXY ? p.x.data() : nullptr;
  out->sizes = sizes.empty() ? nullptr : sizes.data();
  return 0;
}

extern "C" fm_status_t fm_chart_data_get_range_at(const fm_chart_data_t* data, size_t series, int32_t role, size_t idx,
                                                  fm_sheet_rect_t* out) {
  clear_last_error();
  if (data == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(series, data->data.series.size(), "series"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(role, 0, FM_CHART_ROLE_SIZE, "role"); rc != 0) {
    return rc;
  }
  const std::vector<formulon::SheetRect>& ranges = *ranges_of(data->data.series[series], role);
  if (auto rc = check_index(idx, ranges.size(), "index"); rc != 0) {
    return rc;
  }
  const formulon::SheetRect& r = ranges[idx];
  *out = fm_sheet_rect_t{static_cast<std::uint32_t>(r.sheet), fm_cell_rect_t{r.top, r.left, r.bottom, r.right}};
  return 0;
}

extern "C" fm_status_t fm_chart_data_get_axis_title_at(const fm_chart_data_t* data, size_t idx,
                                                       fm_chart_axis_title_text_info_t* out) {
  clear_last_error();
  if (data == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const std::vector<formulon::ChartAxisTitleText>& list = data->data.axis_title_texts;
  if (auto rc = check_index(idx, list.size(), "index"); rc != 0) {
    return rc;
  }
  out->axis_id = list[idx].axis_id;
  out->text = list[idx].text.c_str();
  return 0;
}

extern "C" fm_status_t fm_chart_data_get_unresolved_at(const fm_chart_data_t* data, size_t idx,
                                                       fm_chart_unresolved_info_t* out) {
  clear_last_error();
  if (data == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const std::vector<formulon::ChartUnresolved>& list = data->data.unresolved;
  if (auto rc = check_index(idx, list.size(), "index"); rc != 0) {
    return rc;
  }
  out->series_id = list[idx].series_idx;
  out->role = unresolved_role(list[idx].source);
  out->axis_id = list[idx].axis_id;
  return 0;
}
