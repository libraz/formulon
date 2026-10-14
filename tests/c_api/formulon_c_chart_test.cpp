//
// Stable C ABI tests for charts: the typed view and resolved data behind
// owned handles, creation, every typed edit, the preview, capture and
// restore through the drawing operations, and the refusals.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "drawing/drawing_xml.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_chart_title_t) == (kWasm32 ? 12U : 24U), "fm_chart_title_t ABI layout changed");
static_assert(sizeof(fm_chart_fill_t) == 16U, "fm_chart_fill_t ABI layout changed");
static_assert(sizeof(fm_chart_line_t) == 40U, "fm_chart_line_t ABI layout changed");
static_assert(sizeof(fm_chart_marker_t) == 72U, "fm_chart_marker_t ABI layout changed");
static_assert(sizeof(fm_chart_labels_t) == (kWasm32 ? 64U : 80U), "fm_chart_labels_t ABI layout changed");
static_assert(sizeof(fm_chart_info_t) == (kWasm32 ? 176U : 232U), "fm_chart_info_t ABI layout changed");
static_assert(sizeof(fm_chart_plot_t) == (kWasm32 ? 64U : 80U), "fm_chart_plot_t ABI layout changed");
static_assert(sizeof(fm_chart_data_source_t) == (kWasm32 ? 12U : 24U), "fm_chart_data_source_t ABI layout changed");
static_assert(sizeof(fm_chart_series_t) == (kWasm32 ? 304U : 400U), "fm_chart_series_t ABI layout changed");
static_assert(sizeof(fm_chart_literal_t) == 24U, "fm_chart_literal_t ABI layout changed");
static_assert(sizeof(fm_chart_point_t) == 144U, "fm_chart_point_t ABI layout changed");
static_assert(sizeof(fm_chart_axis_t) == (kWasm32 ? 112U : 136U), "fm_chart_axis_t ABI layout changed");
static_assert(sizeof(fm_chart_preview_overrides) == (kWasm32 ? 44U : 48U),
              "fm_chart_preview_overrides ABI layout changed");
static_assert(sizeof(fm_chart_data_info_t) == (kWasm32 ? 44U : 88U), "fm_chart_data_info_t ABI layout changed");
static_assert(sizeof(fm_chart_data_series_t) == (kWasm32 ? 52U : 104U), "fm_chart_data_series_t ABI layout changed");
static_assert(sizeof(fm_chart_data_points_t) == (kWasm32 ? 20U : 40U), "fm_chart_data_points_t ABI layout changed");
static_assert(sizeof(fm_chart_sheet_range_t) == (kWasm32 ? 20U : 24U), "fm_chart_sheet_range_t ABI layout changed");
static_assert(sizeof(fm_chart_axis_title_text_t) == (kWasm32 ? 8U : 16U),
              "fm_chart_axis_title_text_t ABI layout changed");
static_assert(sizeof(fm_chart_unresolved_t) == (kWasm32 ? 12U : 24U), "fm_chart_unresolved_t ABI layout changed");
static_assert(sizeof(fm_chart_insert) == (kWasm32 ? 88U : 112U), "fm_chart_insert ABI layout changed");
static_assert(sizeof(fm_chart_fill_patch) == 24U, "fm_chart_fill_patch ABI layout changed");
static_assert(sizeof(fm_chart_line_patch) == 56U, "fm_chart_line_patch ABI layout changed");
static_assert(sizeof(fm_chart_marker_patch) == 104U, "fm_chart_marker_patch ABI layout changed");
static_assert(sizeof(fm_chart_labels_patch) == (kWasm32 ? 88U : 104U), "fm_chart_labels_patch ABI layout changed");
static_assert(sizeof(fm_chart_series_patch) == (kWasm32 ? 320U : 360U), "fm_chart_series_patch ABI layout changed");
static_assert(sizeof(fm_chart_point_patch) == 192U, "fm_chart_point_patch ABI layout changed");
static_assert(sizeof(fm_chart_plot_patch) == 40U, "fm_chart_plot_patch ABI layout changed");
static_assert(sizeof(fm_chart_axis_patch) == (kWasm32 ? 128U : 152U), "fm_chart_axis_patch ABI layout changed");
static_assert(sizeof(fm_chart_legend_patch) == (kWasm32 ? 20U : 32U), "fm_chart_legend_patch ABI layout changed");

constexpr fm_status_t Code(formulon::FormulonErrorCode c) {
  return static_cast<fm_status_t>(c);
}
constexpr fm_status_t kInvalidArgument = Code(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = Code(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kInvalidReference = Code(formulon::FormulonErrorCode::kIoChartInvalidReference);
constexpr fm_status_t kUnsupportedKind = Code(formulon::FormulonErrorCode::kIoChartUnsupportedKind);

struct ChartGuard {
  fm_chart_t* handle = nullptr;
  ~ChartGuard() { fm_chart_destroy(handle); }
};

struct DataGuard {
  fm_chart_data_t* handle = nullptr;
  ~DataGuard() { fm_chart_data_destroy(handle); }
};

/// Sheet1: A1 blank, B1:D1 names, A2:A5 text categories, B2:D5 numbers, F1:G5 scatter pairs.
void FillGrid(fm_workbook_t* wb) {
  const char* names[] = {"S1", "S2", "S3"};
  const char* cats[] = {"a", "b", "c", "d"};
  for (std::uint32_t c = 0; c < 3; ++c) {
    ASSERT_EQ(fm_workbook_set_text(wb, 0, 0, c + 1, names[c]), 0);
  }
  for (std::uint32_t r = 0; r < 4; ++r) {
    ASSERT_EQ(fm_workbook_set_text(wb, 0, r + 1, 0, cats[r]), 0);
    for (std::uint32_t c = 0; c < 3; ++c) {
      ASSERT_EQ(fm_workbook_set_number(wb, 0, r + 1, c + 1, static_cast<double>((r + 1) * (c + 2))), 0);
    }
    ASSERT_EQ(fm_workbook_set_number(wb, 0, r + 1, 5, static_cast<double>(r * 3)), 0);
    ASSERT_EQ(fm_workbook_set_number(wb, 0, r + 1, 6, static_cast<double>(r * r)), 0);
  }
  ASSERT_EQ(fm_workbook_set_text(wb, 0, 0, 5, "X"), 0);
  ASSERT_EQ(fm_workbook_set_text(wb, 0, 0, 6, "Y"), 0);
}

fm_chart_insert Options(int32_t kind, const char* source) {
  fm_chart_insert o{};
  o.kind = kind;
  o.source = source;
  o.legend = FM_CHART_LEGEND_RIGHT;
  o.placement.anchor_kind = FM_ANCHOR_KIND_TWO_CELL;
  o.placement.row = 6;
  return o;
}

std::uint32_t Insert(fm_workbook_t* wb, int32_t kind, const char* source = "Sheet1!$A$1:$D$5") {
  const fm_chart_insert o = Options(kind, source);
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_insert_chart(wb, 0, &o, &id), 0) << fm_last_error_message();
  return id;
}

fm_chart_info_t Info(const ChartGuard& chart) {
  fm_chart_info_t info{};
  EXPECT_EQ(fm_chart_info(chart.handle, &info), 0);
  return info;
}

fm_chart_series_t Series(const ChartGuard& chart, size_t plot, size_t series) {
  fm_chart_series_t s{};
  EXPECT_EQ(fm_chart_series_at(chart.handle, plot, series, &s), 0) << fm_last_error_message();
  return s;
}

/// Id of the first axis of `kind`.
uint32_t AxisId(const ChartGuard& chart, int32_t kind) {
  const fm_chart_info_t info = Info(chart);
  for (size_t i = 0; i < info.axis_count; ++i) {
    fm_chart_axis_t a{};
    EXPECT_EQ(fm_chart_axis_at(chart.handle, i, &a), 0);
    if (a.kind == kind) {
      return a.id;
    }
  }
  ADD_FAILURE() << "no axis of kind " << kind;
  return 0;
}

fm_chart_axis_t Axis(const ChartGuard& chart, uint32_t id) {
  const fm_chart_info_t info = Info(chart);
  for (size_t i = 0; i < info.axis_count; ++i) {
    fm_chart_axis_t a{};
    EXPECT_EQ(fm_chart_axis_at(chart.handle, i, &a), 0);
    if (a.id == id) {
      return a;
    }
  }
  ADD_FAILURE() << "no axis " << id;
  return {};
}

void Open(fm_workbook_t* wb, uint32_t id, ChartGuard* out, size_t sheet = 0) {
  fm_chart_destroy(out->handle);
  out->handle = nullptr;
  ASSERT_EQ(fm_sheet_chart_open(wb, sheet, id, &out->handle), 0) << fm_last_error_message();
}

void OpenData(fm_workbook_t* wb, uint32_t id, DataGuard* out) {
  fm_chart_data_destroy(out->handle);
  out->handle = nullptr;
  ASSERT_EQ(fm_sheet_chart_data_open(wb, 0, id, &out->handle), 0) << fm_last_error_message();
}

std::vector<double> Values(const DataGuard& data, size_t series) {
  fm_chart_data_points_t p{};
  EXPECT_EQ(fm_chart_data_series_points(data.handle, series, &p), 0);
  return std::vector<double>(p.values, p.values + p.count);
}

void LoadFixture(const std::string& name, WorkbookGuard* out) {
  std::string dir = __FILE__;
  dir = dir.substr(0, dir.find_last_of("/\\"));
  std::ifstream in(dir + "/../fixtures/excel/" + name, std::ios::binary);
  ASSERT_TRUE(in.good()) << name;
  const Bytes bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out->handle), 0) << fm_last_error_message();
}

struct Located {
  size_t sheet = 0;
  uint32_t id = 0;
};

std::vector<Located> AllCharts(fm_workbook_t* wb) {
  std::vector<Located> out;
  auto all = formulon::all_chart_parts(wb->workbook());
  EXPECT_TRUE(static_cast<bool>(all));
  if (all) {
    for (const formulon::SheetChartPartRef& e : all.value()) {
      out.push_back(Located{e.sheet, e.ref.object_id});
    }
  }
  return out;
}

TEST(FormulonCApiChart, EnumOrdinals) {
  EXPECT_EQ(FM_CHART_PLOT_BAR, 0);
  EXPECT_EQ(FM_CHART_PLOT_DOUGHNUT, 8);
  EXPECT_EQ(FM_CHART_PLOT_SCATTER, 10);
  EXPECT_EQ(FM_CHART_PLOT_CHARTEX, 16);
  EXPECT_EQ(FM_CHART_INSERT_SCATTER, 5);
  EXPECT_EQ(FM_CHART_INSERT_DOUGHNUT, 6);
  EXPECT_EQ(FM_CHART_GROUPING_STANDARD, 3);
  EXPECT_EQ(FM_CHART_BLANKS_SPAN, 3);
  EXPECT_EQ(FM_CHART_POINT_CONNECT, 2);
  EXPECT_EQ(FM_CHART_AXIS_SERIES, 3);
  EXPECT_EQ(FM_CHART_AXIS_POS_RIGHT, 3);
  EXPECT_EQ(FM_CHART_CROSSES_AT, 4);
  EXPECT_EQ(FM_CHART_TITLE_AUTO, 3);
  EXPECT_EQ(FM_CHART_DIRECTION_ROWS, 2);
  EXPECT_EQ(FM_CHART_HEADER_NO, 2);
  EXPECT_EQ(FM_CHART_LEGEND_TOP_RIGHT, 5);
  EXPECT_EQ(FM_CHART_DASH_OTHER, 4);
  EXPECT_EQ(FM_CHART_MARKER_OTHER, 6);
  EXPECT_EQ(FM_CHART_LABEL_RIGHT, 8);
  EXPECT_EQ(FM_CHART_BAR_DIR_BAR, 2);
  EXPECT_EQ(FM_CHART_TRI_TRUE, 2);
  EXPECT_EQ(FM_CHART_ROLE_SIZE, 5);
  EXPECT_EQ(FM_CHART_PATCH_SET, 2);
  EXPECT_EQ(FM_CHART_SOURCE_LITERAL, 2);
}

TEST(FormulonCApiChart, InsertAndReadThroughAccessors) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, FM_CHART_INSERT_COLUMN);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_drawing_object_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_drawing_object obj{};
  ASSERT_EQ(fm_sheet_drawing_object_at(wb.handle, 0, 0, &obj), 0);
  EXPECT_EQ(obj.kind, FM_DRAWING_OBJECT_CHART);
  EXPECT_EQ(obj.object_id, id);

  ChartGuard chart;
  Open(wb.handle, id, &chart);
  const fm_chart_info_t info = Info(chart);
  EXPECT_EQ(info.object_id, id);
  EXPECT_EQ(info.anchor_kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(info.from_row, 6U);
  EXPECT_EQ(info.is_chartex, 0);
  EXPECT_EQ(info.is_pivot_chart, 0);
  EXPECT_EQ(info.title.kind, FM_CHART_TITLE_AUTO);
  EXPECT_EQ(info.legend_position, FM_CHART_LEGEND_RIGHT);
  EXPECT_EQ(info.legend_hidden_count, 0U);
  ASSERT_EQ(info.plot_count, 1U);
  EXPECT_EQ(info.axis_count, 2U);
  EXPECT_EQ(info.hidden_category_count, 0U);
  ASSERT_NE(info.source_range, nullptr);
  EXPECT_STREQ(info.source_range, "Sheet1!$A$1:$D$5");
  EXPECT_EQ(info.source_direction, FM_CHART_DIRECTION_COLUMNS);
  EXPECT_EQ(info.source_header_row, 1);
  EXPECT_EQ(info.source_header_column, 1);

  fm_chart_plot_t plot{};
  ASSERT_EQ(fm_chart_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.type, FM_CHART_PLOT_BAR);
  EXPECT_EQ(plot.bar_dir, FM_CHART_BAR_DIR_COL);
  EXPECT_EQ(plot.has_grouping, 1);
  EXPECT_EQ(plot.grouping, FM_CHART_GROUPING_CLUSTERED);
  EXPECT_STREQ(plot.chartex_layout, "");
  ASSERT_EQ(plot.series_count, 3U);
  ASSERT_EQ(plot.axis_id_count, 2U);
  uint32_t axis_id = 0;
  ASSERT_EQ(fm_chart_plot_axis_id_at(chart.handle, 0, 1, &axis_id), 0);
  EXPECT_EQ(Axis(chart, axis_id).id, axis_id);

  const fm_chart_series_t s = Series(chart, 0, 1);
  EXPECT_EQ(s.idx, 1U);
  EXPECT_EQ(s.order, 1U);
  EXPECT_EQ(s.filtered, 0);
  EXPECT_EQ(s.name.kind, FM_CHART_SOURCE_REF);
  EXPECT_STREQ(s.name.formula, "Sheet1!$C$1");
  EXPECT_STREQ(s.categories.formula, "Sheet1!$A$2:$A$5");
  EXPECT_STREQ(s.values.formula, "Sheet1!$C$2:$C$5");
  EXPECT_EQ(s.x.kind, FM_CHART_SOURCE_NONE);
  EXPECT_EQ(s.size.kind, FM_CHART_SOURCE_NONE);
  EXPECT_EQ(s.point_count, 0U);

  const fm_chart_axis_t cat = Axis(chart, AxisId(chart, FM_CHART_AXIS_CATEGORY));
  const fm_chart_axis_t val = Axis(chart, AxisId(chart, FM_CHART_AXIS_VALUE));
  EXPECT_EQ(cat.cross_axis, val.id);
  EXPECT_EQ(val.cross_axis, cat.id);
  EXPECT_EQ(cat.position, FM_CHART_AXIS_POS_BOTTOM);
  EXPECT_TRUE(std::isnan(val.min));
  EXPECT_TRUE(std::isnan(val.max));

  // The handle is a snapshot: a later edit does not reach it.
  const fm_chart_title_t title{FM_CHART_TITLE_TEXT, "Later", nullptr};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &title), 0);
  EXPECT_EQ(Info(chart).title.kind, FM_CHART_TITLE_AUTO);
}

TEST(FormulonCApiChart, ResolvedDataAccessors) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, FM_CHART_INSERT_COLUMN);

  DataGuard data;
  OpenData(wb.handle, id, &data);
  fm_chart_data_info_t info{};
  ASSERT_EQ(fm_chart_data_info(data.handle, &info), 0);
  EXPECT_EQ(info.is_xy, 0);
  EXPECT_EQ(info.point_count, 4U);
  ASSERT_EQ(info.series_count, 3U);
  ASSERT_EQ(info.category_label_count, 4U);
  EXPECT_EQ(info.category_value_count, 0U);
  EXPECT_EQ(info.category_level_count, 0U);
  EXPECT_EQ(info.categories_generated, 0);
  EXPECT_EQ(info.unresolved_count, 0U);

  const char* label = nullptr;
  double value = 0.0;
  ASSERT_EQ(fm_chart_data_category_at(data.handle, 2, &label, &value), 0);
  EXPECT_STREQ(label, "c");
  EXPECT_TRUE(std::isnan(value));
  EXPECT_EQ(fm_chart_data_category_at(data.handle, 4, &label, &value), kInvalidArgument);

  fm_chart_data_series_t s{};
  ASSERT_EQ(fm_chart_data_series_at(data.handle, 0, &s), 0);
  EXPECT_EQ(s.plot_index, 0U);
  EXPECT_EQ(s.idx, 0U);
  ASSERT_NE(s.name, nullptr);
  EXPECT_STREQ(s.name, "S1");
  EXPECT_EQ(s.point_count, 4U);
  EXPECT_EQ(s.has_sizes, 0);
  EXPECT_EQ(s.range_count[FM_CHART_ROLE_NAME], 1U);
  EXPECT_EQ(s.range_count[FM_CHART_ROLE_VALUES], 1U);
  EXPECT_EQ(s.range_count[FM_CHART_ROLE_X], 0U);

  fm_chart_data_points_t points{};
  ASSERT_EQ(fm_chart_data_series_points(data.handle, 0, &points), 0);
  ASSERT_EQ(points.count, 4U);
  EXPECT_EQ(points.x, nullptr);
  EXPECT_EQ(points.sizes, nullptr);
  EXPECT_EQ(Values(data, 0), (std::vector<double>{2, 4, 6, 8}));
  for (size_t i = 0; i < points.count; ++i) {
    EXPECT_EQ(points.states[i], FM_CHART_POINT_VALUE);
  }

  fm_chart_sheet_range_t r{};
  ASSERT_EQ(fm_chart_data_range_at(data.handle, 0, FM_CHART_ROLE_VALUES, 0, &r), 0);
  EXPECT_EQ(r.sheet, 0U);
  EXPECT_EQ(r.top, 1U);
  EXPECT_EQ(r.left, 1U);
  EXPECT_EQ(r.bottom, 4U);
  EXPECT_EQ(r.right, 1U);
  EXPECT_EQ(fm_chart_data_range_at(data.handle, 0, FM_CHART_ROLE_X, 0, &r), kInvalidArgument);
  EXPECT_EQ(fm_chart_data_range_at(data.handle, 0, 6, 0, &r), kInvalidArgument);
  EXPECT_EQ(fm_chart_data_range_at(data.handle, 3, FM_CHART_ROLE_VALUES, 0, &r), kInvalidArgument);

  // A reference to an empty sheet range, then a value change, reach the next open.
  fm_chart_series_patch patch{};
  patch.values = "Sheet1!$H$2:$H$5";
  ASSERT_EQ(fm_sheet_set_chart_series(wb.handle, 0, id, 0, 2, &patch), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 1, 1, 42.0), 0);
  OpenData(wb.handle, id, &data);
  EXPECT_EQ(Values(data, 0)[0], 42.0);
  fm_chart_data_points_t gaps{};
  ASSERT_EQ(fm_chart_data_series_points(data.handle, 2, &gaps), 0);
  ASSERT_EQ(gaps.count, 4U);
  EXPECT_NE(gaps.states[0], FM_CHART_POINT_VALUE);

  // Title text and axis titles resolve through the data handle.
  const fm_chart_title_t title{FM_CHART_TITLE_REF, nullptr, "Sheet1!$B$1"};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &title), 0) << fm_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  fm_chart_axis_patch axis{};
  axis.title_op = FM_CHART_PATCH_SET;
  axis.title = "Amount";
  const uint32_t val_id = AxisId(chart, FM_CHART_AXIS_VALUE);
  ASSERT_EQ(fm_sheet_set_chart_axis(wb.handle, 0, id, val_id, &axis), 0) << fm_last_error_message();
  OpenData(wb.handle, id, &data);
  ASSERT_EQ(fm_chart_data_info(data.handle, &info), 0);
  ASSERT_NE(info.title_text, nullptr);
  EXPECT_STREQ(info.title_text, "S1");
  ASSERT_EQ(info.axis_title_count, 1U);
  fm_chart_axis_title_text_t at{};
  ASSERT_EQ(fm_chart_data_axis_title_at(data.handle, 0, &at), 0);
  EXPECT_EQ(at.axis_id, val_id);
  EXPECT_STREQ(at.text, "Amount");
  EXPECT_EQ(fm_chart_data_axis_title_at(data.handle, 1, &at), kInvalidArgument);
  fm_chart_unresolved_t un{};
  EXPECT_EQ(fm_chart_data_unresolved_at(data.handle, 0, &un), kInvalidArgument);
}

TEST(FormulonCApiChart, ScatterDataCarriesX) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, FM_CHART_INSERT_SCATTER, "Sheet1!$F$1:$G$5");
  DataGuard data;
  OpenData(wb.handle, id, &data);
  fm_chart_data_info_t info{};
  ASSERT_EQ(fm_chart_data_info(data.handle, &info), 0);
  EXPECT_EQ(info.is_xy, 1);
  ASSERT_EQ(info.series_count, 1U);
  fm_chart_data_points_t p{};
  ASSERT_EQ(fm_chart_data_series_points(data.handle, 0, &p), 0);
  ASSERT_EQ(p.count, 4U);
  ASSERT_NE(p.x, nullptr);
  EXPECT_EQ(std::vector<double>(p.x, p.x + p.count), (std::vector<double>{0, 3, 6, 9}));
  EXPECT_EQ(std::vector<double>(p.values, p.values + p.count), (std::vector<double>{0, 1, 4, 9}));

  ChartGuard chart;
  Open(wb.handle, id, &chart);
  fm_chart_plot_t plot{};
  ASSERT_EQ(fm_chart_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.type, FM_CHART_PLOT_SCATTER);
  const fm_chart_series_t s = Series(chart, 0, 0);
  EXPECT_STREQ(s.x.formula, "Sheet1!$F$2:$F$5");
  EXPECT_STREQ(s.y.formula, "Sheet1!$G$2:$G$5");
}

TEST(FormulonCApiChart, TypeTitleLegendDisplayPlotEdits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, FM_CHART_INSERT_COLUMN);
  ChartGuard chart;

  fm_chart_plot_patch plot{};
  plot.gap_width_op = FM_CHART_PATCH_SET;
  plot.gap_width = 50;
  plot.overlap_op = FM_CHART_PATCH_SET;
  plot.overlap = -10;
  plot.vary_colors_op = FM_CHART_PATCH_SET;
  plot.vary_colors = 1;
  ASSERT_EQ(fm_sheet_set_chart_plot(wb.handle, 0, id, 0, &plot), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_plot_t p{};
  ASSERT_EQ(fm_chart_plot_at(chart.handle, 0, &p), 0);
  EXPECT_EQ(p.has_gap_width, 1);
  EXPECT_EQ(p.gap_width, 50U);
  EXPECT_EQ(p.has_overlap, 1);
  EXPECT_EQ(p.overlap, -10);
  EXPECT_EQ(p.vary_colors, FM_CHART_TRI_TRUE);
  // A plot field cannot be cleared, and an unknown op is rejected.
  fm_chart_plot_patch bad{};
  bad.gap_width_op = FM_CHART_PATCH_CLEAR;
  EXPECT_EQ(fm_sheet_set_chart_plot(wb.handle, 0, id, 0, &bad), kInvalidArgument);
  bad.gap_width_op = 3;
  EXPECT_EQ(fm_sheet_set_chart_plot(wb.handle, 0, id, 0, &bad), kInvalidArgument);
  bad = fm_chart_plot_patch{};
  bad.hole_size_op = FM_CHART_PATCH_SET;
  bad.hole_size = 50;
  EXPECT_EQ(fm_sheet_set_chart_plot(wb.handle, 0, id, 0, &bad), kInvalidArgument);  // not a doughnut
  EXPECT_EQ(fm_sheet_set_chart_plot(wb.handle, 0, id, 1, &plot), kInvalidArgument);

  const fm_chart_title_t title{FM_CHART_TITLE_TEXT, "Sales", nullptr};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &title), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).title.kind, FM_CHART_TITLE_TEXT);
  EXPECT_STREQ(Info(chart).title.text, "Sales");
  const fm_chart_title_t none{FM_CHART_TITLE_NONE, nullptr, nullptr};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &none), 0);
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).title.kind, FM_CHART_TITLE_NONE);
  const fm_chart_title_t bad_kind{4, nullptr, nullptr};
  EXPECT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &bad_kind), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, nullptr), kBindingNullPointer);

  const uint32_t hidden[] = {1U};
  fm_chart_legend_patch legend{FM_CHART_LEGEND_BOTTOM, 1, FM_CHART_PATCH_SET, hidden, 1};
  ASSERT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_info_t info = Info(chart);
  EXPECT_EQ(info.legend_position, FM_CHART_LEGEND_BOTTOM);
  EXPECT_EQ(info.legend_overlay, 1);
  ASSERT_EQ(info.legend_hidden_count, 1U);
  uint32_t entry = 0;
  ASSERT_EQ(fm_chart_legend_hidden_at(chart.handle, 0, &entry), 0);
  EXPECT_EQ(entry, 1U);
  EXPECT_EQ(fm_chart_legend_hidden_at(chart.handle, 1, &entry), kInvalidArgument);
  legend = fm_chart_legend_patch{FM_CHART_LEGEND_NONE, 0, FM_CHART_PATCH_KEEP, nullptr, 0};
  ASSERT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), 0);
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).legend_position, FM_CHART_LEGEND_NONE);
  legend.hidden_series_count = 2;
  EXPECT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), kBindingNullPointer);
  legend = fm_chart_legend_patch{6, 0, FM_CHART_PATCH_KEEP, nullptr, 0};
  EXPECT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), kInvalidArgument);

  ASSERT_EQ(fm_sheet_set_chart_display(wb.handle, 0, id, FM_CHART_BLANKS_ZERO, FM_CHART_TRI_FALSE), 0)
      << fm_last_error_message();
  Open(wb.handle, id, &chart);
  info = Info(chart);
  EXPECT_EQ(info.blanks_as, FM_CHART_BLANKS_ZERO);
  EXPECT_EQ(info.plot_vis_only, FM_CHART_TRI_FALSE);
  EXPECT_EQ(fm_sheet_set_chart_display(wb.handle, 0, id, 4, FM_CHART_TRI_FALSE), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_display(wb.handle, 0, id, FM_CHART_BLANKS_GAP, 3), kInvalidArgument);

  ASSERT_EQ(fm_sheet_set_chart_type(wb.handle, 0, id, 0, FM_CHART_INSERT_LINE, FM_CHART_GROUPING_STANDARD), 0)
      << fm_last_error_message();
  Open(wb.handle, id, &chart);
  ASSERT_EQ(fm_chart_plot_at(chart.handle, 0, &p), 0);
  EXPECT_EQ(p.type, FM_CHART_PLOT_LINE);
  EXPECT_EQ(p.series_count, 3U);
  EXPECT_STREQ(Series(chart, 0, 2).values.formula, "Sheet1!$D$2:$D$5");
  EXPECT_EQ(fm_sheet_set_chart_type(wb.handle, 0, id, 0, 7, FM_CHART_GROUPING_STANDARD), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_type(wb.handle, 0, id, 0, FM_CHART_INSERT_LINE, 4), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_type(wb.handle, 0, 99, 0, FM_CHART_INSERT_LINE, 0), kInvalidArgument);
}

TEST(FormulonCApiChart, AxisEdits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, FM_CHART_INSERT_COLUMN);
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  const uint32_t val_id = AxisId(chart, FM_CHART_AXIS_VALUE);
  const uint32_t cat_id = AxisId(chart, FM_CHART_AXIS_CATEGORY);

  fm_chart_axis_patch patch{};
  patch.min_op = FM_CHART_PATCH_SET;
  patch.min = 0.0;
  patch.max_op = FM_CHART_PATCH_SET;
  patch.max = 20.0;
  patch.major_unit_op = FM_CHART_PATCH_SET;
  patch.major_unit = 5.0;
  patch.number_format_op = FM_CHART_PATCH_SET;
  patch.number_format = "0.0";
  patch.title_op = FM_CHART_PATCH_SET;
  patch.title = "Amount";
  patch.major_gridlines_op = FM_CHART_PATCH_SET;
  patch.major_gridlines = 0;
  patch.reversed_op = FM_CHART_PATCH_SET;
  patch.reversed = 1;
  patch.log_base_op = FM_CHART_PATCH_SET;
  patch.log_base = 10.0;
  patch.tick_label_pos_op = FM_CHART_PATCH_SET;
  patch.tick_label_pos = "low";
  ASSERT_EQ(fm_sheet_set_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_axis_t a = Axis(chart, val_id);
  EXPECT_EQ(a.min, 0.0);
  EXPECT_EQ(a.max, 20.0);
  EXPECT_EQ(a.major_unit, 5.0);
  EXPECT_EQ(a.log_base, 10.0);
  ASSERT_NE(a.number_format, nullptr);
  EXPECT_STREQ(a.number_format, "0.0");
  EXPECT_EQ(a.source_linked, 0);
  EXPECT_EQ(a.title.kind, FM_CHART_TITLE_TEXT);
  EXPECT_STREQ(a.title.text, "Amount");
  EXPECT_EQ(a.major_gridlines, 0);
  EXPECT_EQ(a.reversed, 1);
  ASSERT_NE(a.tick_label_pos, nullptr);
  EXPECT_STREQ(a.tick_label_pos, "low");

  patch = fm_chart_axis_patch{};
  patch.max_op = FM_CHART_PATCH_CLEAR;
  patch.log_base_op = FM_CHART_PATCH_CLEAR;
  patch.title_op = FM_CHART_PATCH_CLEAR;
  patch.visible_op = FM_CHART_PATCH_SET;
  patch.visible = 0;
  ASSERT_EQ(fm_sheet_set_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  a = Axis(chart, val_id);
  EXPECT_TRUE(std::isnan(a.max));
  EXPECT_TRUE(std::isnan(a.log_base));
  EXPECT_EQ(a.title.kind, FM_CHART_TITLE_NONE);
  EXPECT_EQ(a.deleted, 1);

  // The crossing point is written on the cross axis.
  patch = fm_chart_axis_patch{};
  patch.crosses_op = FM_CHART_PATCH_SET;
  patch.crosses = FM_CHART_CROSSES_AT;
  patch.crosses_at = 3.0;
  ASSERT_EQ(fm_sheet_set_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  const fm_chart_axis_t crossing = Axis(chart, cat_id);
  const fm_chart_axis_t crossed = Axis(chart, val_id);
  EXPECT_TRUE((crossing.crosses == FM_CHART_CROSSES_AT && crossing.crosses_at == 3.0) ||
              (crossed.crosses == FM_CHART_CROSSES_AT && crossed.crosses_at == 3.0));

  patch = fm_chart_axis_patch{};
  patch.visible_op = FM_CHART_PATCH_CLEAR;
  EXPECT_EQ(fm_sheet_set_chart_axis(wb.handle, 0, id, val_id, &patch), kInvalidArgument);
  patch = fm_chart_axis_patch{};
  patch.crosses_op = FM_CHART_PATCH_SET;
  patch.crosses = 5;
  EXPECT_EQ(fm_sheet_set_chart_axis(wb.handle, 0, id, val_id, &patch), kInvalidArgument);
  const fm_chart_axis_patch keep{};
  EXPECT_EQ(fm_sheet_set_chart_axis(wb.handle, 0, id, 999, &keep), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_axis(wb.handle, 0, id, val_id, nullptr), kBindingNullPointer);
}

TEST(FormulonCApiChart, SeriesPointFilterAndSourceEdits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, FM_CHART_INSERT_LINE);
  ChartGuard chart;

  fm_chart_series_patch patch{};
  patch.line.color_op = FM_CHART_PATCH_SET;
  patch.line.color = 0xFFFF0000U;
  patch.line.transparency_op = FM_CHART_PATCH_SET;
  patch.line.transparency_pct = 25.0;
  patch.line.width_op = FM_CHART_PATCH_SET;
  patch.line.width_pt = 2.0;
  patch.line.dash_op = FM_CHART_PATCH_SET;
  patch.line.dash = FM_CHART_DASH_DASH;
  patch.marker_op = FM_CHART_PATCH_SET;
  patch.marker.symbol_op = FM_CHART_PATCH_SET;
  patch.marker.symbol = FM_CHART_MARKER_CIRCLE;
  patch.marker.size_op = FM_CHART_PATCH_SET;
  patch.marker.size = 7.0;
  patch.marker.fill.color_op = FM_CHART_PATCH_SET;
  patch.marker.fill.color = 0xFF00FF00U;
  patch.smooth_op = FM_CHART_PATCH_SET;
  patch.smooth = 1;
  patch.labels_op = FM_CHART_PATCH_SET;
  patch.labels.show_value_op = FM_CHART_PATCH_SET;
  patch.labels.show_value = 1;
  patch.labels.number_format_op = FM_CHART_PATCH_SET;
  patch.labels.number_format = "0";
  patch.labels.font_size_op = FM_CHART_PATCH_SET;
  patch.labels.font_size_pt = 9.0;
  patch.labels.font_color_op = FM_CHART_PATCH_SET;
  patch.labels.font_color = 0xFF0000FFU;
  patch.labels.separator_op = FM_CHART_PATCH_SET;
  patch.labels.separator = "; ";
  ASSERT_EQ(fm_sheet_set_chart_series(wb.handle, 0, id, 0, 0, &patch), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_series_t s = Series(chart, 0, 0);
  EXPECT_EQ(s.line.has_color, 1);
  EXPECT_EQ(s.line.color, 0xFFFF0000U);
  EXPECT_DOUBLE_EQ(s.line.transparency_pct, 25.0);
  EXPECT_DOUBLE_EQ(s.line.width_pt, 2.0);
  EXPECT_EQ(s.line.has_dash, 1);
  EXPECT_EQ(s.line.dash, FM_CHART_DASH_DASH);
  ASSERT_EQ(s.has_marker, 1);
  EXPECT_EQ(s.marker.symbol, FM_CHART_MARKER_CIRCLE);
  EXPECT_EQ(s.marker.has_size, 1);
  EXPECT_EQ(s.marker.size, 7U);
  EXPECT_EQ(s.marker.fill.color, 0xFF00FF00U);
  EXPECT_EQ(s.smooth, FM_CHART_TRI_TRUE);
  ASSERT_EQ(s.has_labels, 1);
  EXPECT_EQ(s.labels.show_value, 1);
  EXPECT_EQ(s.labels.show_category, 0);
  ASSERT_NE(s.labels.number_format, nullptr);
  EXPECT_STREQ(s.labels.number_format, "0");
  EXPECT_DOUBLE_EQ(s.labels.font_size_pt, 9.0);
  EXPECT_EQ(s.labels.has_font_color, 1);
  EXPECT_EQ(s.labels.font_color, 0xFF0000FFU);
  ASSERT_NE(s.labels.separator, nullptr);
  EXPECT_STREQ(s.labels.separator, "; ");

  patch = fm_chart_series_patch{};
  patch.marker_op = FM_CHART_PATCH_CLEAR;
  patch.labels_op = FM_CHART_PATCH_CLEAR;
  patch.line.width_op = FM_CHART_PATCH_CLEAR;
  patch.name = "";
  ASSERT_EQ(fm_sheet_set_chart_series(wb.handle, 0, id, 0, 0, &patch), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  s = Series(chart, 0, 0);
  EXPECT_EQ(s.has_marker, 0);
  EXPECT_EQ(s.has_labels, 0);
  EXPECT_TRUE(std::isnan(s.line.width_pt));
  EXPECT_EQ(s.name.kind, FM_CHART_SOURCE_NONE);

  patch = fm_chart_series_patch{};
  patch.values = "";
  EXPECT_EQ(fm_sheet_set_chart_series(wb.handle, 0, id, 0, 0, &patch), kInvalidArgument);
  patch.values = "Nope!$B$2:$B$5";
  EXPECT_EQ(fm_sheet_set_chart_series(wb.handle, 0, id, 0, 0, &patch), kInvalidReference);
  patch = fm_chart_series_patch{};
  patch.smooth_op = FM_CHART_PATCH_CLEAR;
  EXPECT_EQ(fm_sheet_set_chart_series(wb.handle, 0, id, 0, 0, &patch), kInvalidArgument);
  patch = fm_chart_series_patch{};
  patch.marker_op = FM_CHART_PATCH_SET;
  patch.marker.symbol_op = FM_CHART_PATCH_SET;
  patch.marker.symbol = 7;
  EXPECT_EQ(fm_sheet_set_chart_series(wb.handle, 0, id, 0, 0, &patch), kInvalidArgument);
  // A nested patch whose op is KEEP is not read.
  patch.marker_op = FM_CHART_PATCH_KEEP;
  EXPECT_EQ(fm_sheet_set_chart_series(wb.handle, 0, id, 0, 0, &patch), 0) << fm_last_error_message();
  EXPECT_EQ(fm_sheet_set_chart_series(wb.handle, 0, id, 0, 9, &patch), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_series(wb.handle, 0, id, 0, 0, nullptr), kBindingNullPointer);

  fm_chart_point_patch point{};
  point.line.color_op = FM_CHART_PATCH_SET;
  point.line.color = 0xFF123456U;
  point.marker_op = FM_CHART_PATCH_SET;
  point.marker.symbol_op = FM_CHART_PATCH_SET;
  point.marker.symbol = FM_CHART_MARKER_TRIANGLE;
  ASSERT_EQ(fm_sheet_set_chart_point(wb.handle, 0, id, 0, 1, 2, &point), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  s = Series(chart, 0, 1);
  ASSERT_EQ(s.point_count, 1U);
  fm_chart_point_t pt{};
  ASSERT_EQ(fm_chart_series_point_at(chart.handle, 0, 1, 0, &pt), 0);
  EXPECT_EQ(pt.idx, 2U);
  EXPECT_EQ(pt.line.color, 0xFF123456U);
  ASSERT_EQ(pt.has_marker, 1);
  EXPECT_EQ(pt.marker.symbol, FM_CHART_MARKER_TRIANGLE);
  EXPECT_EQ(fm_chart_series_point_at(chart.handle, 0, 1, 1, &pt), kInvalidArgument);
  const fm_chart_point_patch empty{};
  ASSERT_EQ(fm_sheet_set_chart_point(wb.handle, 0, id, 0, 1, 2, &empty), 0);
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Series(chart, 0, 1).point_count, 0U);
  EXPECT_EQ(fm_sheet_set_chart_point(wb.handle, 0, id, 0, 1, 2, nullptr), kBindingNullPointer);

  const uint32_t hidden_series[] = {1U};
  const uint32_t hidden_categories[] = {0U, 2U};
  ASSERT_EQ(fm_sheet_set_chart_filter(wb.handle, 0, id, hidden_series, 1, hidden_categories, 2), 0)
      << fm_last_error_message();
  Open(wb.handle, id, &chart);
  ASSERT_EQ(Info(chart).hidden_category_count, 2U);
  uint32_t cat = 0;
  ASSERT_EQ(fm_chart_hidden_category_at(chart.handle, 1, &cat), 0);
  EXPECT_EQ(cat, 2U);
  EXPECT_EQ(fm_chart_hidden_category_at(chart.handle, 2, &cat), kInvalidArgument);
  bool filtered = false;
  fm_chart_plot_t plot{};
  ASSERT_EQ(fm_chart_plot_at(chart.handle, 0, &plot), 0);
  for (size_t i = 0; i < plot.series_count; ++i) {
    const fm_chart_series_t each = Series(chart, 0, i);
    filtered = filtered || (each.idx == 1U && each.filtered == 1);
  }
  EXPECT_TRUE(filtered);
  ASSERT_EQ(fm_sheet_set_chart_filter(wb.handle, 0, id, nullptr, 0, nullptr, 0), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).hidden_category_count, 0U);
  EXPECT_EQ(fm_sheet_set_chart_filter(wb.handle, 0, id, nullptr, 1, nullptr, 0), kBindingNullPointer);

  patch = fm_chart_series_patch{};
  patch.name = "Sheet1!$D$1";
  patch.categories = "Sheet1!$A$2:$A$5";
  patch.values = "Sheet1!$D$2:$D$5";
  uint32_t added = 0;
  ASSERT_EQ(fm_sheet_add_chart_series(wb.handle, 0, id, 0, &patch, &added), 0) << fm_last_error_message();
  EXPECT_EQ(added, 3U);
  Open(wb.handle, id, &chart);
  ASSERT_EQ(fm_chart_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.series_count, 4U);
  EXPECT_EQ(fm_sheet_add_chart_series(wb.handle, 0, id, 0, &patch, nullptr), kBindingNullPointer);
  ASSERT_EQ(fm_sheet_remove_chart_series(wb.handle, 0, id, 0, added), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  ASSERT_EQ(fm_chart_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.series_count, 3U);
  EXPECT_EQ(fm_sheet_remove_chart_series(wb.handle, 0, id, 0, added), kInvalidArgument);

  ASSERT_EQ(fm_sheet_set_chart_source(wb.handle, 0, id, "Sheet1!$A$1:$C$5", FM_CHART_DIRECTION_ROWS,
                                      FM_CHART_HEADER_AUTO, FM_CHART_HEADER_AUTO),
            0)
      << fm_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_info_t info = Info(chart);
  ASSERT_EQ(fm_chart_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.series_count, 4U);
  ASSERT_NE(info.source_range, nullptr);
  EXPECT_STREQ(info.source_range, "Sheet1!$A$1:$C$5");
  EXPECT_EQ(info.source_direction, FM_CHART_DIRECTION_ROWS);
  EXPECT_EQ(fm_sheet_set_chart_source(wb.handle, 0, id, "Nope!$A$1:$B$5", FM_CHART_DIRECTION_AUTO, FM_CHART_HEADER_AUTO,
                                      FM_CHART_HEADER_AUTO),
            kInvalidReference);
  EXPECT_EQ(
      fm_sheet_set_chart_source(wb.handle, 0, id, "Sheet1!$A$1:$C$5", 3, FM_CHART_HEADER_AUTO, FM_CHART_HEADER_AUTO),
      kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_source(wb.handle, 0, id, nullptr, 0, 0, 0), kBindingNullPointer);
}

TEST(FormulonCApiChart, PreviewLeavesWorkbookUnchanged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, FM_CHART_INSERT_COLUMN);

  fm_chart_preview_overrides o{};
  o.source = "Sheet1!$A$1:$B$5";
  DataGuard preview;
  ASSERT_EQ(fm_sheet_chart_preview_open(wb.handle, 0, id, &o, &preview.handle), 0) << fm_last_error_message();
  fm_chart_data_info_t info{};
  ASSERT_EQ(fm_chart_data_info(preview.handle, &info), 0);
  EXPECT_EQ(info.series_count, 1U);

  o = fm_chart_preview_overrides{};
  o.direction_op = FM_CHART_PATCH_SET;
  o.direction = FM_CHART_DIRECTION_ROWS;
  DataGuard rows;
  ASSERT_EQ(fm_sheet_chart_preview_open(wb.handle, 0, id, &o, &rows.handle), 0) << fm_last_error_message();
  ASSERT_EQ(fm_chart_data_info(rows.handle, &info), 0);
  EXPECT_EQ(info.series_count, 4U);
  EXPECT_EQ(info.point_count, 3U);

  // Applying the same change for real gives the same data.
  ASSERT_EQ(fm_sheet_set_chart_source(wb.handle, 0, id, "Sheet1!$A$1:$D$5", FM_CHART_DIRECTION_ROWS,
                                      FM_CHART_HEADER_AUTO, FM_CHART_HEADER_AUTO),
            0);
  DataGuard applied;
  OpenData(wb.handle, id, &applied);
  for (size_t i = 0; i < 4; ++i) {
    EXPECT_EQ(Values(rows, i), Values(applied, i));
  }

  DataGuard current;
  o = fm_chart_preview_overrides{};
  o.source = "Sheet1!$A$1:$B$5";
  ASSERT_EQ(fm_sheet_chart_preview_open(wb.handle, 0, id, &o, &current.handle), 0);
  DataGuard after;
  OpenData(wb.handle, id, &after);
  ASSERT_EQ(fm_chart_data_info(after.handle, &info), 0);
  EXPECT_EQ(info.series_count, 4U);

  o = fm_chart_preview_overrides{};
  o.blanks_as_op = FM_CHART_PATCH_CLEAR;
  fm_chart_data_t* raw = reinterpret_cast<fm_chart_data_t*>(&o);  // reset to NULL on failure
  EXPECT_EQ(fm_sheet_chart_preview_open(wb.handle, 0, id, &o, &raw), kInvalidArgument);
  EXPECT_EQ(raw, nullptr);
  o = fm_chart_preview_overrides{};
  o.source = "Nope!$A$1:$B$5";
  EXPECT_EQ(fm_sheet_chart_preview_open(wb.handle, 0, id, &o, &raw), kInvalidReference);
  EXPECT_EQ(fm_sheet_chart_preview_open(wb.handle, 0, id, nullptr, &raw), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_chart_preview_open(wb.handle, 0, id, &o, nullptr), kBindingNullPointer);
}

TEST(FormulonCApiChart, DoughnutHoleSize) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, FM_CHART_INSERT_DOUGHNUT, "Sheet1!$A$1:$B$5");
  fm_chart_plot_patch patch{};
  patch.hole_size_op = FM_CHART_PATCH_SET;
  patch.hole_size = 60;
  ASSERT_EQ(fm_sheet_set_chart_plot(wb.handle, 0, id, 0, &patch), 0) << fm_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  fm_chart_plot_t plot{};
  ASSERT_EQ(fm_chart_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.type, FM_CHART_PLOT_DOUGHNUT);
  EXPECT_EQ(plot.has_hole_size, 1);
  EXPECT_EQ(plot.hole_size, 60U);
  EXPECT_EQ(Info(chart).axis_count, 0U);
}

TEST(FormulonCApiChart, CaptureRestoreAndRemoveThroughDrawingOperations) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, FM_CHART_INSERT_COLUMN);
  const fm_chart_title_t before{FM_CHART_TITLE_TEXT, "Before", nullptr};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &before), 0);

  BufferGuard snap;
  ASSERT_EQ(fm_sheet_snapshot_drawing(wb.handle, 0, id, &snap.data, &snap.len), 0) << fm_last_error_message();
  const fm_chart_title_t after{FM_CHART_TITLE_TEXT, "After", nullptr};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &after), 0);

  // Same id: the chart part's bytes come back.
  uint32_t restored = 0;
  ASSERT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len, 0, &restored), 0) << fm_last_error_message();
  EXPECT_EQ(restored, id);
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  EXPECT_STREQ(Info(chart).title.text, "Before");

  // Move and reorder like any drawing object.
  fm_drawing_anchor anchor{};
  anchor.anchor_kind = FM_ANCHOR_KIND_TWO_CELL;
  anchor.row = 10;
  anchor.col = 2;
  ASSERT_EQ(fm_sheet_set_drawing_anchor(wb.handle, 0, id, &anchor), 0) << fm_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).from_row, 10U);
  EXPECT_EQ(Info(chart).from_col, 2U);

  // A copy with a fresh id reads as its own chart.
  uint32_t copy = 0;
  ASSERT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len, FM_DRAWING_RESTORE_NEW_ID, &copy), 0)
      << fm_last_error_message();
  EXPECT_NE(copy, id);
  ChartGuard copied;
  Open(wb.handle, copy, &copied);
  EXPECT_STREQ(Info(copied).title.text, "Before");
  ASSERT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, copy, 0), 0) << fm_last_error_message();

  ASSERT_EQ(fm_sheet_remove_drawing(wb.handle, 0, id), 0) << fm_last_error_message();
  fm_chart_t* gone = nullptr;
  EXPECT_EQ(fm_sheet_chart_open(wb.handle, 0, id, &gone), kInvalidArgument);
  EXPECT_EQ(gone, nullptr);
  ASSERT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len, 0, &restored), 0) << fm_last_error_message();
  EXPECT_EQ(restored, id);
  Open(wb.handle, id, &chart);
  EXPECT_STREQ(Info(chart).title.text, "Before");

  // The edited chart survives a save and reload.
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_last_error_message();
  ChartGuard again;
  Open(reloaded.handle, id, &again);
  EXPECT_STREQ(Info(again).title.text, "Before");
}

TEST(FormulonCApiChart, OpenAndAccessorsRejectBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, FM_CHART_INSERT_COLUMN);

  fm_chart_t* chart = reinterpret_cast<fm_chart_t*>(&wb);
  EXPECT_EQ(fm_sheet_chart_open(wb.handle, 0, 99, &chart), kInvalidArgument);
  EXPECT_EQ(chart, nullptr);
  EXPECT_EQ(fm_sheet_chart_open(wb.handle, 4, id, &chart), kInvalidArgument);
  EXPECT_EQ(fm_sheet_chart_open(wb.handle, 0, id, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_chart_open(nullptr, 0, id, &chart), kBindingNullPointer);
  fm_chart_data_t* data = nullptr;
  EXPECT_EQ(fm_sheet_chart_data_open(wb.handle, 0, 99, &data), kInvalidArgument);
  EXPECT_EQ(data, nullptr);
  EXPECT_EQ(fm_sheet_chart_data_open(wb.handle, 0, id, nullptr), kBindingNullPointer);
  fm_chart_destroy(nullptr);
  fm_chart_data_destroy(nullptr);

  ChartGuard guard;
  Open(wb.handle, id, &guard);
  fm_chart_info_t info{};
  EXPECT_EQ(fm_chart_info(nullptr, &info), kBindingNullPointer);
  EXPECT_EQ(fm_chart_info(guard.handle, nullptr), kBindingNullPointer);
  fm_chart_plot_t plot{};
  EXPECT_EQ(fm_chart_plot_at(guard.handle, 1, &plot), kInvalidArgument);
  uint32_t u = 0;
  EXPECT_EQ(fm_chart_plot_axis_id_at(guard.handle, 0, 2, &u), kInvalidArgument);
  EXPECT_EQ(fm_chart_plot_axis_id_at(guard.handle, 1, 0, &u), kInvalidArgument);
  fm_chart_series_t s{};
  EXPECT_EQ(fm_chart_series_at(guard.handle, 0, 3, &s), kInvalidArgument);
  fm_chart_literal_t lit{};
  EXPECT_EQ(fm_chart_series_literal_at(guard.handle, 0, 0, FM_CHART_ROLE_VALUES, 0, &lit), kInvalidArgument);
  EXPECT_EQ(fm_chart_series_literal_at(guard.handle, 0, 0, 6, 0, &lit), kInvalidArgument);
  EXPECT_EQ(fm_chart_series_literal_at(guard.handle, 0, 0, -1, 0, &lit), kInvalidArgument);
  fm_chart_point_t pt{};
  EXPECT_EQ(fm_chart_series_point_at(guard.handle, 0, 0, 0, &pt), kInvalidArgument);
  fm_chart_axis_t axis{};
  EXPECT_EQ(fm_chart_axis_at(guard.handle, 2, &axis), kInvalidArgument);
  EXPECT_EQ(fm_chart_axis_at(guard.handle, 0, nullptr), kBindingNullPointer);

  DataGuard dg;
  OpenData(wb.handle, id, &dg);
  fm_chart_data_series_t ds{};
  EXPECT_EQ(fm_chart_data_series_at(dg.handle, 3, &ds), kInvalidArgument);
  fm_chart_data_points_t pts{};
  EXPECT_EQ(fm_chart_data_series_points(dg.handle, 3, &pts), kInvalidArgument);
  EXPECT_EQ(fm_chart_data_series_points(nullptr, 0, &pts), kBindingNullPointer);
  size_t n = 0;
  EXPECT_EQ(fm_chart_data_category_level_size(dg.handle, 0, &n), kInvalidArgument);
  const char* text = nullptr;
  EXPECT_EQ(fm_chart_data_category_level_at(dg.handle, 0, 0, &text), kInvalidArgument);
  double v = 0.0;
  EXPECT_EQ(fm_chart_data_category_at(dg.handle, 0, nullptr, &v), kBindingNullPointer);
}

TEST(FormulonCApiChart, InsertRejectsBadOptions) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  uint32_t id = 0;
  fm_chart_insert o = Options(FM_CHART_INSERT_COLUMN, "Sheet1!$A$1:$D$5");
  o.kind = 7;
  EXPECT_EQ(fm_sheet_insert_chart(wb.handle, 0, &o, &id), kInvalidArgument);
  o = Options(FM_CHART_INSERT_COLUMN, "Sheet1!$A$1:$D$5");
  o.legend = 6;
  EXPECT_EQ(fm_sheet_insert_chart(wb.handle, 0, &o, &id), kInvalidArgument);
  o = Options(FM_CHART_INSERT_COLUMN, "Sheet1!$A$1:$D$5");
  o.header_row = 3;
  EXPECT_EQ(fm_sheet_insert_chart(wb.handle, 0, &o, &id), kInvalidArgument);
  o = Options(FM_CHART_INSERT_COLUMN, "Sheet1!$A$1:$D$5");
  o.placement.anchor_kind = 3;
  EXPECT_EQ(fm_sheet_insert_chart(wb.handle, 0, &o, &id), kInvalidArgument);
  o = Options(FM_CHART_INSERT_COLUMN, "Nope!$A$1:$D$5");
  EXPECT_EQ(fm_sheet_insert_chart(wb.handle, 0, &o, &id), kInvalidReference);
  o = Options(FM_CHART_INSERT_COLUMN, nullptr);
  EXPECT_EQ(fm_sheet_insert_chart(wb.handle, 0, &o, &id), kBindingNullPointer);
  o = Options(FM_CHART_INSERT_COLUMN, "Sheet1!$A$1:$D$5");
  EXPECT_EQ(fm_sheet_insert_chart(wb.handle, 0, &o, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_insert_chart(wb.handle, 0, nullptr, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_insert_chart(wb.handle, 3, &o, &id), kInvalidArgument);
  size_t count = 9;
  ASSERT_EQ(fm_sheet_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  // An empty title is no title; a given name and description are kept.
  o.title = "";
  o.placement.name = "Sales chart";
  o.placement.descr = "Quarterly";
  ASSERT_EQ(fm_sheet_insert_chart(wb.handle, 0, &o, &id), 0) << fm_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  const fm_chart_info_t info = Info(chart);
  EXPECT_EQ(info.title.kind, FM_CHART_TITLE_NONE);
  EXPECT_EQ(info.auto_title_deleted, FM_CHART_TRI_TRUE);
  EXPECT_STREQ(info.name, "Sales chart");
  EXPECT_STREQ(info.descr, "Quarterly");
}

TEST(FormulonCApiChart, ChartexAndPivotChartsRefuseEdits) {
  WorkbookGuard cx;
  LoadFixture("chart_chartex.xlsx", &cx);
  const std::vector<Located> cx_charts = AllCharts(cx.handle);
  ASSERT_FALSE(cx_charts.empty());
  const Located c = cx_charts[0];
  ChartGuard chart;
  Open(cx.handle, c.id, &chart, c.sheet);
  EXPECT_EQ(Info(chart).is_chartex, 1);
  fm_chart_plot_t plot{};
  ASSERT_EQ(fm_chart_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.type, FM_CHART_PLOT_CHARTEX);
  EXPECT_STRNE(plot.chartex_layout, "");
  const fm_chart_title_t title{FM_CHART_TITLE_TEXT, "T", nullptr};
  EXPECT_EQ(fm_sheet_set_chart_title(cx.handle, c.sheet, c.id, &title), kUnsupportedKind);
  EXPECT_EQ(fm_sheet_set_chart_display(cx.handle, c.sheet, c.id, FM_CHART_BLANKS_GAP, FM_CHART_TRI_TRUE),
            kUnsupportedKind);
  fm_chart_preview_overrides o{};
  fm_chart_data_t* preview = nullptr;
  EXPECT_EQ(fm_sheet_chart_preview_open(cx.handle, c.sheet, c.id, &o, &preview), kUnsupportedKind);
  DataGuard data;
  ASSERT_EQ(fm_sheet_chart_data_open(cx.handle, c.sheet, c.id, &data.handle), 0) << fm_last_error_message();

  WorkbookGuard pv;
  LoadFixture("chart_pivot.xlsx", &pv);
  const std::vector<Located> pv_charts = AllCharts(pv.handle);
  ASSERT_FALSE(pv_charts.empty());
  const Located p = pv_charts[0];
  Open(pv.handle, p.id, &chart, p.sheet);
  EXPECT_EQ(Info(chart).is_pivot_chart, 1);
  fm_chart_series_patch rename{};
  rename.name = "Sheet1!$A$1";
  EXPECT_EQ(fm_sheet_set_chart_series(pv.handle, p.sheet, p.id, 0, 0, &rename), kUnsupportedKind);
  uint32_t added = 0;
  EXPECT_EQ(fm_sheet_add_chart_series(pv.handle, p.sheet, p.id, 0, &rename, &added), kUnsupportedKind);
  EXPECT_EQ(fm_sheet_remove_chart_series(pv.handle, p.sheet, p.id, 0, 0), kUnsupportedKind);
  const uint32_t hidden[] = {0U};
  EXPECT_EQ(fm_sheet_set_chart_filter(pv.handle, p.sheet, p.id, hidden, 1, nullptr, 0), kUnsupportedKind);
  EXPECT_EQ(fm_sheet_set_chart_source(pv.handle, p.sheet, p.id, "Sheet1!$A$1:$B$3", FM_CHART_DIRECTION_AUTO,
                                      FM_CHART_HEADER_AUTO, FM_CHART_HEADER_AUTO),
            kUnsupportedKind);
  EXPECT_EQ(fm_sheet_set_chart_title(pv.handle, p.sheet, p.id, &title), 0) << fm_last_error_message();
}

TEST(FormulonCApiChart, LiteralsAndMultiLevelCategoriesFromExcelFiles) {
  WorkbookGuard lit;
  LoadFixture("chart_cache.xlsx", &lit);
  size_t literals = 0;
  for (const Located& c : AllCharts(lit.handle)) {
    ChartGuard chart;
    Open(lit.handle, c.id, &chart, c.sheet);
    const fm_chart_info_t info = Info(chart);
    for (size_t p = 0; p < info.plot_count; ++p) {
      fm_chart_plot_t plot{};
      ASSERT_EQ(fm_chart_plot_at(chart.handle, p, &plot), 0);
      for (size_t i = 0; i < plot.series_count; ++i) {
        const fm_chart_series_t s = Series(chart, p, i);
        const fm_chart_data_source_t* sources[] = {&s.name, &s.categories, &s.values, &s.x, &s.y, &s.size};
        for (int32_t role = FM_CHART_ROLE_NAME; role <= FM_CHART_ROLE_SIZE; ++role) {
          const fm_chart_data_source_t& src = *sources[role];
          if (src.kind != FM_CHART_SOURCE_LITERAL) {
            continue;
          }
          ASSERT_GT(src.literal_count, 0U);
          for (size_t k = 0; k < src.literal_count; ++k) {
            fm_chart_literal_t l{};
            ASSERT_EQ(fm_chart_series_literal_at(chart.handle, p, i, role, k, &l), 0);
            ASSERT_NE(l.text, nullptr);
            ++literals;
          }
          fm_chart_literal_t l{};
          EXPECT_EQ(fm_chart_series_literal_at(chart.handle, p, i, role, src.literal_count, &l), kInvalidArgument);
        }
      }
    }
  }
  EXPECT_GT(literals, 0U);

  WorkbookGuard multi;
  LoadFixture("chart_data.xlsx", &multi);
  bool found = false;
  for (const Located& c : AllCharts(multi.handle)) {
    DataGuard data;
    ASSERT_EQ(fm_sheet_chart_data_open(multi.handle, c.sheet, c.id, &data.handle), 0) << fm_last_error_message();
    fm_chart_data_info_t info{};
    ASSERT_EQ(fm_chart_data_info(data.handle, &info), 0);
    for (size_t level = 0; level < info.category_level_count; ++level) {
      size_t n = 0;
      ASSERT_EQ(fm_chart_data_category_level_size(data.handle, level, &n), 0);
      for (size_t k = 0; k < n; ++k) {
        const char* label = nullptr;
        ASSERT_EQ(fm_chart_data_category_level_at(data.handle, level, k, &label), 0);
        ASSERT_NE(label, nullptr);
      }
      const char* label = nullptr;
      EXPECT_EQ(fm_chart_data_category_level_at(data.handle, level, n, &label), kInvalidArgument);
      found = true;
    }
  }
  EXPECT_TRUE(found);
}

}  // namespace
