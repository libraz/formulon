//
// Stable C ABI tests for charts: the typed view and resolved data behind
// owned handles, creation, every typed edit, the preview, capture and
// restore through the drawing operations, and the refusals.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <limits>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "drawing/drawing_xml.h"
#include "formulon_c_out_index.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "passthrough_part.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_chart_title_t) == (kWasm32 ? 12U : 24U), "fm_chart_title_t ABI layout changed");
static_assert(sizeof(fm_chart_fill_t) == 40U, "fm_chart_fill_t ABI layout changed");
static_assert(sizeof(fm_chart_line_t) == 72U, "fm_chart_line_t ABI layout changed");
static_assert(sizeof(fm_chart_marker_t) == 128U, "fm_chart_marker_t ABI layout changed");
static_assert(sizeof(fm_chart_labels_t) == (kWasm32 ? 88U : 104U), "fm_chart_labels_t ABI layout changed");
static_assert(sizeof(fm_chart_info_t) == (kWasm32 ? 100U : 120U), "fm_chart_info_t ABI layout changed");
static_assert(sizeof(fm_chart_plot_info_t) == 68U, "fm_chart_plot_info_t ABI layout changed");
static_assert(sizeof(fm_chart_data_source_t) == (kWasm32 ? 12U : 24U), "fm_chart_data_source_t ABI layout changed");
static_assert(sizeof(fm_chart_series_info_t) == (kWasm32 ? 440U : 528U), "fm_chart_series_info_t ABI layout changed");
static_assert(sizeof(fm_chart_literal_info_t) == 32U, "fm_chart_literal_info_t ABI layout changed");
static_assert(sizeof(fm_chart_point_info_t) == 256U, "fm_chart_point_info_t ABI layout changed");
static_assert(sizeof(fm_chart_axis_info_t) == (kWasm32 ? 152U : 176U), "fm_chart_axis_info_t ABI layout changed");
static_assert(sizeof(fm_chart_preview_overrides_spec_t) == (kWasm32 ? 60U : 64U),
              "fm_chart_preview_overrides_spec_t ABI layout changed");
static_assert(sizeof(fm_chart_data_info_t) == (kWasm32 ? 48U : 64U), "fm_chart_data_info_t ABI layout changed");
static_assert(sizeof(fm_chart_data_series_info_t) == (kWasm32 ? 60U : 80U),
              "fm_chart_data_series_info_t ABI layout changed");
static_assert(offsetof(fm_chart_data_series_info_t, x_format_code) == (kWasm32 ? 56U : 72U),
              "fm_chart_data_series_info_t x_format_code offset changed");
static_assert(sizeof(fm_chart_data_points_info_t) == (kWasm32 ? 24U : 40U),
              "fm_chart_data_points_info_t ABI layout changed");
static_assert(sizeof(fm_sheet_rect_t) == 20U, "fm_sheet_rect_t ABI layout changed");
static_assert(sizeof(fm_chart_axis_title_text_info_t) == (kWasm32 ? 12U : 16U),
              "fm_chart_axis_title_text_info_t ABI layout changed");
static_assert(sizeof(fm_chart_unresolved_info_t) == 16U, "fm_chart_unresolved_info_t ABI layout changed");
static_assert(sizeof(fm_chart_insert_spec_t) == (kWasm32 ? 136U : 176U), "fm_chart_insert_spec_t ABI layout changed");
static_assert(sizeof(fm_chart_fill_patch_t) == 48U, "fm_chart_fill_patch_t ABI layout changed");
static_assert(sizeof(fm_chart_line_patch_t) == 80U, "fm_chart_line_patch_t ABI layout changed");
static_assert(sizeof(fm_chart_marker_patch_t) == 152U, "fm_chart_marker_patch_t ABI layout changed");
static_assert(sizeof(fm_chart_labels_patch_t) == (kWasm32 ? 128U : 144U), "fm_chart_labels_patch_t ABI layout changed");
static_assert(sizeof(fm_chart_series_patch_t) == (kWasm32 ? 560U : 688U), "fm_chart_series_patch_t ABI layout changed");
static_assert(sizeof(fm_chart_point_patch_t) == 296U, "fm_chart_point_patch_t ABI layout changed");
static_assert(sizeof(fm_chart_plot_patch_t) == 44U, "fm_chart_plot_patch_t ABI layout changed");
static_assert(sizeof(fm_chart_axis_patch_t) == (kWasm32 ? 184U : 208U), "fm_chart_axis_patch_t ABI layout changed");
static_assert(sizeof(fm_chart_legend_patch_t) == (kWasm32 ? 32U : 40U), "fm_chart_legend_patch_t ABI layout changed");
static_assert(sizeof(fm_chart_legend_t) == (kWasm32 ? 20U : 32U), "fm_chart_legend_t ABI layout changed");
static_assert(sizeof(fm_chart_display_patch_t) == 36U, "fm_chart_display_patch_t ABI layout changed");
static_assert(offsetof(fm_chart_info_t, na_as_blank) == (kWasm32 ? 96U : 116U), "fm_chart_info_t na_as_blank offset");

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
    ASSERT_EQ(fm_sheet_set_text(wb, 0, 0, c + 1, names[c]), 0);
  }
  for (std::uint32_t r = 0; r < 4; ++r) {
    ASSERT_EQ(fm_sheet_set_text(wb, 0, r + 1, 0, cats[r]), 0);
    for (std::uint32_t c = 0; c < 3; ++c) {
      ASSERT_EQ(fm_sheet_set_number(wb, 0, r + 1, c + 1, static_cast<double>((r + 1) * (c + 2))), 0);
    }
    ASSERT_EQ(fm_sheet_set_number(wb, 0, r + 1, 5, static_cast<double>(r * 3)), 0);
    ASSERT_EQ(fm_sheet_set_number(wb, 0, r + 1, 6, static_cast<double>(r * r)), 0);
  }
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 0, 5, "X"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 0, 6, "Y"), 0);
}

/// Writes `formula` as a reference input of a series patch member.
void SetRefSource(fm_patch_op_t* op, fm_chart_data_source_spec_t* spec, const char* formula) {
  *op = FM_PATCH_OP_SET;
  *spec = fm_chart_data_source_spec_t{FM_CHART_SOURCE_KIND_REF, formula, nullptr, 0};
}

/// A writable plot type with the bar direction it needs.
struct InsertKind {
  int32_t type;
  int32_t bar_direction;
};
constexpr InsertKind kColumn{FM_CHART_PLOT_TYPE_BAR, FM_CHART_BAR_DIRECTION_COLUMN};
constexpr InsertKind kLine{FM_CHART_PLOT_TYPE_LINE, FM_CHART_BAR_DIRECTION_COLUMN};
constexpr InsertKind kPie{FM_CHART_PLOT_TYPE_PIE, FM_CHART_BAR_DIRECTION_COLUMN};
constexpr InsertKind kScatter{FM_CHART_PLOT_TYPE_SCATTER, FM_CHART_BAR_DIRECTION_COLUMN};
constexpr InsertKind kDoughnut{FM_CHART_PLOT_TYPE_DOUGHNUT, FM_CHART_BAR_DIRECTION_COLUMN};

fm_chart_insert_spec_t Options(InsertKind kind, const char* source) {
  fm_chart_insert_spec_t o = FM_STRUCT_INIT(fm_chart_insert_spec_t);
  o.type = kind.type;
  o.bar_direction = kind.bar_direction;
  o.source = source;
  o.legend = FM_CHART_LEGEND_POSITION_RIGHT;
  o.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  o.placement.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  o.placement.anchor.edit_as = FM_ANCHOR_EDIT_AS_TWO_CELL;
  o.placement.anchor.from_row = 6;
  return o;
}

std::uint32_t Insert(fm_workbook_t* wb, InsertKind kind, const char* source = "Sheet1!$A$1:$D$5") {
  const fm_chart_insert_spec_t o = Options(kind, source);
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_chart(wb, 0, &o, &id), 0) << fm_get_last_error_message();
  return id;
}

fm_chart_info_t Info(const ChartGuard& chart) {
  fm_chart_info_t info = FM_STRUCT_INIT(fm_chart_info_t);
  EXPECT_EQ(fm_chart_get_info(chart.handle, &info), 0);
  return info;
}

fm_chart_series_info_t Series(const ChartGuard& chart, size_t plot, size_t series) {
  fm_chart_series_info_t s = FM_STRUCT_INIT(fm_chart_series_info_t);
  EXPECT_EQ(fm_chart_get_series_at(chart.handle, plot, series, &s), 0) << fm_get_last_error_message();
  return s;
}

/// Id of the first axis of `kind`.
uint32_t AxisId(const ChartGuard& chart, int32_t kind) {
  const fm_chart_info_t info = Info(chart);
  for (size_t i = 0; i < info.axis_count; ++i) {
    fm_chart_axis_info_t a = FM_STRUCT_INIT(fm_chart_axis_info_t);
    EXPECT_EQ(fm_chart_get_axis_at(chart.handle, i, &a), 0);
    if (a.kind == kind) {
      return a.id;
    }
  }
  ADD_FAILURE() << "no axis of kind " << kind;
  return 0;
}

fm_chart_axis_info_t Axis(const ChartGuard& chart, uint32_t id) {
  const fm_chart_info_t info = Info(chart);
  for (size_t i = 0; i < info.axis_count; ++i) {
    fm_chart_axis_info_t a = FM_STRUCT_INIT(fm_chart_axis_info_t);
    EXPECT_EQ(fm_chart_get_axis_at(chart.handle, i, &a), 0);
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
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_chart(wb, sheet, id, &present, &out->handle), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
}

void OpenData(fm_workbook_t* wb, uint32_t id, DataGuard* out) {
  fm_chart_data_destroy(out->handle);
  out->handle = nullptr;
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_chart_data(wb, 0, id, &present, &out->handle), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
}

std::vector<double> Values(const DataGuard& data, size_t series) {
  fm_chart_data_points_info_t p = FM_STRUCT_INIT(fm_chart_data_points_info_t);
  EXPECT_EQ(fm_chart_data_get_series_points(data.handle, series, &p), 0);
  return std::vector<double>(p.values, p.values + p.count);
}

void LoadFixture(const std::string& name, WorkbookGuard* out) {
  std::string dir = __FILE__;
  dir = dir.substr(0, dir.find_last_of("/\\"));
  std::ifstream in(dir + "/../fixtures/excel/" + name, std::ios::binary);
  ASSERT_TRUE(in.good()) << name;
  const Bytes bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out->handle), 0) << fm_get_last_error_message();
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
  EXPECT_EQ(FM_CHART_PLOT_TYPE_BAR, 0);
  EXPECT_EQ(FM_CHART_PLOT_TYPE_DOUGHNUT, 8);
  EXPECT_EQ(FM_CHART_PLOT_TYPE_SCATTER, 10);
  EXPECT_EQ(FM_CHART_PLOT_TYPE_EXTENDED, 16);
  EXPECT_EQ(FM_CHART_PLOT_TYPE_AREA, 4);
  EXPECT_EQ(FM_CHART_PLOT_TYPE_PIE, 6);
  EXPECT_EQ(FM_CHART_BAR_DIRECTION_COLUMN, 0);
  EXPECT_EQ(FM_CHART_BAR_DIRECTION_BAR, 1);
  EXPECT_EQ(FM_CHART_SOURCE_KIND_NONE, 0);
  EXPECT_EQ(FM_CHART_SOURCE_KIND_REF, 1);
  EXPECT_EQ(FM_CHART_SOURCE_KIND_LITERAL, 2);
  EXPECT_EQ(FM_CHART_GROUPING_STANDARD, 3);
  EXPECT_EQ(FM_CHART_POINT_STATE_CONNECT, 2);
  EXPECT_EQ(FM_CHART_AXIS_KIND_SERIES, 3);
  EXPECT_EQ(FM_CHART_AXIS_POSITION_RIGHT, 3);
  EXPECT_EQ(FM_CHART_AXIS_CROSSES_AUTO_ZERO, 0);
  EXPECT_EQ(FM_CHART_AXIS_CROSSES_AT, 3);
  EXPECT_EQ(FM_CHART_TITLE_KIND_AUTO, 3);
  EXPECT_EQ(FM_CHART_SERIES_DIRECTION_ROWS, 2);
  EXPECT_EQ(FM_TRISTATE_DEFAULT, 0);
  EXPECT_EQ(FM_TRISTATE_FALSE, 1);
  EXPECT_EQ(FM_TRISTATE_TRUE, 2);
  EXPECT_EQ(FM_CHART_LINE_DASH_OTHER, 4);
  EXPECT_EQ(FM_CHART_MARKER_SYMBOL_PICTURE, 11);
  EXPECT_EQ(FM_CHART_LABEL_POSITION_RIGHT, 8);
  EXPECT_EQ(FM_CHART_BAR_DIRECTION_COLUMN, 0);
  EXPECT_EQ(FM_CHART_BAR_DIRECTION_BAR, 1);
  // A member meaning "the element is absent" is -1.
  EXPECT_EQ(FM_BLANKS_AS_ABSENT, -1);
  EXPECT_EQ(FM_CHART_TICK_LABEL_POSITION_ABSENT, -1);
  EXPECT_EQ(FM_CHART_AXIS_CROSSES_ABSENT, -1);
  EXPECT_EQ(FM_CHART_BAR_DIRECTION_ABSENT, -1);
  EXPECT_EQ(FM_CHART_LEGEND_POSITION_ABSENT, -1);
  EXPECT_EQ(FM_TRISTATE_TRUE, 2);
  EXPECT_EQ(FM_CHART_ROLE_SIZE, 5);
  EXPECT_EQ(FM_PATCH_OP_SET, 2);
  EXPECT_EQ(FM_CHART_SOURCE_KIND_LITERAL, 2);
}

// Zero is the OOXML default `r`; the rest follow `ST_LegendPos`.
TEST(FormulonCApiChart, LegendPositionOrdinalsStartAtTheSchemaDefault) {
  EXPECT_EQ(FM_CHART_LEGEND_POSITION_ABSENT, -1);
  EXPECT_EQ(FM_CHART_LEGEND_POSITION_RIGHT, 0);
  EXPECT_EQ(FM_CHART_LEGEND_POSITION_BOTTOM, 1);
  EXPECT_EQ(FM_CHART_LEGEND_POSITION_TOP_RIGHT, 2);
  EXPECT_EQ(FM_CHART_LEGEND_POSITION_LEFT, 3);
  EXPECT_EQ(FM_CHART_LEGEND_POSITION_TOP, 4);
}

// One table for chart and sparkline blanks; zero is the OOXML default `zero`.
TEST(FormulonCApiChart, BlanksAsOrdinalsStartAtTheSchemaDefault) {
  EXPECT_EQ(FM_BLANKS_AS_ABSENT, -1);
  EXPECT_EQ(FM_BLANKS_AS_ZERO, 0);
  EXPECT_EQ(FM_BLANKS_AS_GAP, 1);
  EXPECT_EQ(FM_BLANKS_AS_SPAN, 2);
}

// Zero is the OOXML default `nextTo`; the rest follow `ST_TickLblPos`.
TEST(FormulonCApiChart, TickLabelPositionOrdinalsStartAtTheSchemaDefault) {
  EXPECT_EQ(FM_CHART_TICK_LABEL_POSITION_ABSENT, -1);
  EXPECT_EQ(FM_CHART_TICK_LABEL_POSITION_NEXT_TO, 0);
  EXPECT_EQ(FM_CHART_TICK_LABEL_POSITION_HIGH, 1);
  EXPECT_EQ(FM_CHART_TICK_LABEL_POSITION_LOW, 2);
  EXPECT_EQ(FM_CHART_TICK_LABEL_POSITION_NONE, 3);
}

// An element the chart does not carry reads as -1: no legend, and no
TEST(FormulonCApiChart, InsertTitleFollowsHasTitle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const auto insert = [&](int32_t has_title, fm_chart_title_t title, fm_status_t* status) {
    fm_chart_insert_spec_t o = Options(kLine, "Sheet1!$A$1:$D$5");
    o.has_title = has_title;
    o.title = title;
    uint32_t id = 0;
    *status = fm_sheet_add_chart(wb.handle, 0, &o, &id);
    return id;
  };
  fm_status_t rc = 0;
  uint32_t id = insert(0, {FM_CHART_TITLE_KIND_TEXT, "Ignored", nullptr}, &rc);
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).title.kind, FM_CHART_TITLE_KIND_AUTO);

  id = insert(1, {FM_CHART_TITLE_KIND_NONE, nullptr, nullptr}, &rc);
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();
  ChartGuard none;
  Open(wb.handle, id, &none);
  EXPECT_EQ(Info(none).title.kind, FM_CHART_TITLE_KIND_NONE);

  id = insert(1, {FM_CHART_TITLE_KIND_TEXT, "Sales", nullptr}, &rc);
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();
  ChartGuard text;
  Open(wb.handle, id, &text);
  EXPECT_EQ(Info(text).title.kind, FM_CHART_TITLE_KIND_TEXT);
  EXPECT_STREQ(Info(text).title.text, "Sales");

  insert(1, {FM_CHART_TITLE_KIND_TEXT, "", nullptr}, &rc);
  EXPECT_EQ(rc, kInvalidArgument);
  insert(1, {FM_CHART_TITLE_KIND_TEXT, nullptr, nullptr}, &rc);
  EXPECT_EQ(rc, kInvalidArgument);
  insert(1, {FM_CHART_TITLE_KIND_AUTO + 1, nullptr, nullptr}, &rc);
  EXPECT_EQ(rc, kInvalidArgument);
  insert(2, {FM_CHART_TITLE_KIND_NONE, nullptr, nullptr}, &rc);
  EXPECT_EQ(rc, kInvalidArgument);
}

TEST(FormulonCApiChart, InsertWithBadReferenceTitleLeavesNoChart) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  fm_chart_insert_spec_t o = Options(kLine, "Sheet1!$A$1:$D$5");
  o.has_title = 1;
  uint32_t id = 0;
  for (const char* bad : {"not a reference", "Sheet1!$A$0", ""}) {
    o.title = fm_chart_title_t{FM_CHART_TITLE_KIND_REF, nullptr, bad};
    EXPECT_NE(fm_sheet_add_chart(wb.handle, 0, &o, &id), 0) << bad;
    size_t count = 9;
    ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
    EXPECT_EQ(count, 0U) << bad;
  }
  o.title = fm_chart_title_t{FM_CHART_TITLE_KIND_REF, nullptr, "Sheet1!$B$1"};
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, &id), 0) << fm_get_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).title.kind, FM_CHART_TITLE_KIND_REF);
}

// `c:barDir` on a plot that is not a bar plot.
TEST(FormulonCApiChart, AbsentElementsReadAsMinusOne) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  fm_chart_insert_spec_t o = Options(kLine, "Sheet1!$A$1:$D$5");
  o.legend = FM_CHART_LEGEND_POSITION_ABSENT;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, &id), 0) << fm_get_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).legend_position, -1);
  fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.type, FM_CHART_PLOT_TYPE_LINE);
  EXPECT_EQ(plot.bar_direction, -1);
}

TEST(FormulonCApiChart, InsertAndReadThroughAccessors) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_drawing_object_info_t obj = FM_STRUCT_INIT(fm_drawing_object_info_t);
  ASSERT_EQ(fm_sheet_get_drawing_object_at(wb.handle, 0, 0, &obj), 0);
  EXPECT_EQ(obj.kind, FM_DRAWING_OBJECT_KIND_CHART);
  EXPECT_EQ(obj.object_id, id);

  ChartGuard chart;
  Open(wb.handle, id, &chart);
  const fm_chart_info_t info = Info(chart);
  EXPECT_EQ(info.object_id, id);
  EXPECT_EQ(obj.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(obj.anchor.from_row, 6U);
  EXPECT_EQ(info.is_extended, 0);
  EXPECT_EQ(info.is_pivot_chart, 0);
  EXPECT_EQ(info.title.kind, FM_CHART_TITLE_KIND_AUTO);
  EXPECT_EQ(info.legend_position, FM_CHART_LEGEND_POSITION_RIGHT);
  EXPECT_EQ(info.legend_hidden_count, 0U);
  ASSERT_EQ(info.plot_count, 1U);
  EXPECT_EQ(info.axis_count, 2U);
  EXPECT_EQ(info.hidden_category_count, 0U);
  ASSERT_NE(info.source_range, nullptr);
  EXPECT_STREQ(info.source_range, "Sheet1!$A$1:$D$5");
  EXPECT_EQ(info.source_direction, FM_CHART_SERIES_DIRECTION_COLUMNS);
  EXPECT_EQ(info.source_header_row, 1);
  EXPECT_EQ(info.source_header_column, 1);

  fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.type, FM_CHART_PLOT_TYPE_BAR);
  EXPECT_EQ(plot.bar_direction, FM_CHART_BAR_DIRECTION_COLUMN);
  EXPECT_EQ(plot.has_grouping, 1);
  EXPECT_EQ(plot.grouping, FM_CHART_GROUPING_CLUSTERED);
  EXPECT_EQ(plot.extended_layout, FM_CHART_EXTENDED_LAYOUT_ABSENT);
  ASSERT_EQ(plot.series_count, 3U);
  ASSERT_EQ(plot.axis_id_count, 2U);
  uint32_t axis_id = 0;
  ASSERT_EQ(fm_chart_get_plot_axis_id_at(chart.handle, 0, 1, &axis_id), 0);
  EXPECT_EQ(Axis(chart, axis_id).id, axis_id);

  const fm_chart_series_info_t s = Series(chart, 0, 1);
  EXPECT_EQ(s.series_id, 1U);
  EXPECT_EQ(s.order, 1U);
  EXPECT_EQ(s.filtered, 0);
  EXPECT_EQ(s.name.kind, FM_CHART_SOURCE_KIND_REF);
  EXPECT_STREQ(s.name.formula, "Sheet1!$C$1");
  EXPECT_STREQ(s.categories.formula, "Sheet1!$A$2:$A$5");
  EXPECT_STREQ(s.values.formula, "Sheet1!$C$2:$C$5");
  EXPECT_EQ(s.x.kind, FM_CHART_SOURCE_KIND_NONE);
  EXPECT_EQ(s.size.kind, FM_CHART_SOURCE_KIND_NONE);
  EXPECT_EQ(s.point_count, 0U);

  const fm_chart_axis_info_t cat = Axis(chart, AxisId(chart, FM_CHART_AXIS_KIND_CATEGORY));
  const fm_chart_axis_info_t val = Axis(chart, AxisId(chart, FM_CHART_AXIS_KIND_VALUE));
  EXPECT_EQ(cat.cross_axis, val.id);
  EXPECT_EQ(val.cross_axis, cat.id);
  EXPECT_EQ(cat.position, FM_CHART_AXIS_POSITION_BOTTOM);
  EXPECT_EQ(val.has_min, 0);
  EXPECT_EQ(val.has_max, 0);

  // The handle is a snapshot: a later edit does not reach it.
  const fm_chart_title_t title{FM_CHART_TITLE_KIND_TEXT, "Later", nullptr};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &title), 0);
  EXPECT_EQ(Info(chart).title.kind, FM_CHART_TITLE_KIND_AUTO);
}

TEST(FormulonCApiChart, ResolvedDataAccessors) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);

  DataGuard data;
  OpenData(wb.handle, id, &data);
  fm_chart_data_info_t info = FM_STRUCT_INIT(fm_chart_data_info_t);
  ASSERT_EQ(fm_chart_data_get_info(data.handle, &info), 0);
  EXPECT_EQ(info.is_xy, 0);
  EXPECT_EQ(info.point_count, 4U);
  ASSERT_EQ(info.series_count, 3U);
  ASSERT_EQ(info.category_label_count, 4U);
  EXPECT_EQ(info.category_value_count, 0U);
  EXPECT_EQ(info.category_level_count, 0U);
  EXPECT_EQ(info.categories_generated, 0);
  EXPECT_EQ(info.unresolved_count, 0U);

  const char* label = nullptr;
  int32_t has_value = -1;
  double value = -1.0;
  ASSERT_EQ(fm_chart_data_get_category_at(data.handle, 2, &label, &has_value, &value), 0);
  EXPECT_STREQ(label, "c");
  EXPECT_EQ(has_value, 0);
  EXPECT_EQ(value, 0.0);
  EXPECT_EQ(fm_chart_data_get_category_at(data.handle, 4, &label, &has_value, &value), FM_STATUS_INDEX_OUT_OF_RANGE);

  fm_chart_data_series_info_t s = FM_STRUCT_INIT(fm_chart_data_series_info_t);
  ASSERT_EQ(fm_chart_data_get_series_at(data.handle, 0, &s), 0);
  EXPECT_EQ(s.plot_index, 0U);
  EXPECT_EQ(s.series_id, 0U);
  ASSERT_NE(s.name, nullptr);
  EXPECT_STREQ(s.name, "S1");
  EXPECT_EQ(s.point_count, 4U);
  EXPECT_EQ(s.has_sizes, 0);
  EXPECT_STREQ(s.x_format_code, "");
  EXPECT_EQ(s.range_count[FM_CHART_ROLE_NAME], 1U);
  EXPECT_EQ(s.range_count[FM_CHART_ROLE_VALUES], 1U);
  EXPECT_EQ(s.range_count[FM_CHART_ROLE_X], 0U);

  fm_chart_data_points_info_t points = FM_STRUCT_INIT(fm_chart_data_points_info_t);
  ASSERT_EQ(fm_chart_data_get_series_points(data.handle, 0, &points), 0);
  ASSERT_EQ(points.count, 4U);
  EXPECT_EQ(points.x, nullptr);
  EXPECT_EQ(points.sizes, nullptr);
  EXPECT_EQ(Values(data, 0), (std::vector<double>{2, 4, 6, 8}));
  for (size_t i = 0; i < points.count; ++i) {
    EXPECT_EQ(points.states[i], FM_CHART_POINT_STATE_VALUE);
  }

  fm_sheet_rect_t r{};
  ASSERT_EQ(fm_chart_data_get_range_at(data.handle, 0, FM_CHART_ROLE_VALUES, 0, &r), 0);
  EXPECT_EQ(r.sheet_index, 0U);
  EXPECT_EQ(r.rect.first_row, 1U);
  EXPECT_EQ(r.rect.first_col, 1U);
  EXPECT_EQ(r.rect.last_row, 4U);
  EXPECT_EQ(r.rect.last_col, 1U);
  EXPECT_EQ(fm_chart_data_get_range_at(data.handle, 0, FM_CHART_ROLE_X, 0, &r), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_chart_data_get_range_at(data.handle, 0, 6, 0, &r), kInvalidArgument);
  EXPECT_EQ(fm_chart_data_get_range_at(data.handle, 3, FM_CHART_ROLE_VALUES, 0, &r), FM_STATUS_INDEX_OUT_OF_RANGE);

  // A reference to an empty sheet range, then a value change, reach the next open.
  fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  SetRefSource(&patch.values_op, &patch.values, "Sheet1!$H$2:$H$5");
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 2, &patch), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 42.0), 0);
  OpenData(wb.handle, id, &data);
  EXPECT_EQ(Values(data, 0)[0], 42.0);
  fm_chart_data_points_info_t gaps = FM_STRUCT_INIT(fm_chart_data_points_info_t);
  ASSERT_EQ(fm_chart_data_get_series_points(data.handle, 2, &gaps), 0);
  ASSERT_EQ(gaps.count, 4U);
  EXPECT_NE(gaps.states[0], FM_CHART_POINT_STATE_VALUE);

  // Title text and axis titles resolve through the data handle.
  const fm_chart_title_t title{FM_CHART_TITLE_KIND_REF, nullptr, "Sheet1!$B$1"};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &title), 0) << fm_get_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  fm_chart_axis_patch_t axis = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  axis.title_op = FM_PATCH_OP_SET;
  axis.title = fm_chart_title_t{FM_CHART_TITLE_KIND_TEXT, "Amount", nullptr};
  const uint32_t val_id = AxisId(chart, FM_CHART_AXIS_KIND_VALUE);
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &axis), 0) << fm_get_last_error_message();
  OpenData(wb.handle, id, &data);
  ASSERT_EQ(fm_chart_data_get_info(data.handle, &info), 0);
  ASSERT_NE(info.title_text, nullptr);
  EXPECT_STREQ(info.title_text, "S1");
  ASSERT_EQ(info.axis_title_count, 1U);
  fm_chart_axis_title_text_info_t at = FM_STRUCT_INIT(fm_chart_axis_title_text_info_t);
  ASSERT_EQ(fm_chart_data_get_axis_title_at(data.handle, 0, &at), 0);
  EXPECT_EQ(at.axis_id, val_id);
  EXPECT_STREQ(at.text, "Amount");
  EXPECT_EQ(fm_chart_data_get_axis_title_at(data.handle, 1, &at), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_chart_unresolved_info_t un = FM_STRUCT_INIT(fm_chart_unresolved_info_t);
  EXPECT_EQ(fm_chart_data_get_unresolved_at(data.handle, 0, &un), FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiChart, ScatterDataCarriesX) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  fm_cell_xf_t x_xf{};
  x_xf.num_fmt_id = 2;  // built-in 0.00, on the x column only
  uint32_t x_xf_index = 0;
  ASSERT_EQ(fm_workbook_add_cell_xf(wb.handle, &x_xf, &x_xf_index), 0);
  ASSERT_EQ(fm_sheet_set_range_xf_index(wb.handle, 0, record_arg(fm_cell_rect_t{1, 5, 4, 5}), x_xf_index), 0);
  const uint32_t id = Insert(wb.handle, kScatter, "Sheet1!$F$1:$G$5");
  DataGuard data;
  OpenData(wb.handle, id, &data);
  fm_chart_data_info_t info = FM_STRUCT_INIT(fm_chart_data_info_t);
  ASSERT_EQ(fm_chart_data_get_info(data.handle, &info), 0);
  EXPECT_EQ(info.is_xy, 1);
  ASSERT_EQ(info.series_count, 1U);
  fm_chart_data_points_info_t p = FM_STRUCT_INIT(fm_chart_data_points_info_t);
  ASSERT_EQ(fm_chart_data_get_series_points(data.handle, 0, &p), 0);
  ASSERT_EQ(p.count, 4U);
  fm_chart_data_series_info_t ds = FM_STRUCT_INIT(fm_chart_data_series_info_t);
  ASSERT_EQ(fm_chart_data_get_series_at(data.handle, 0, &ds), 0);
  EXPECT_STREQ(ds.x_format_code, "0.00");
  EXPECT_STREQ(ds.format_code, "General");
  ASSERT_NE(p.x, nullptr);
  EXPECT_EQ(std::vector<double>(p.x, p.x + p.count), (std::vector<double>{0, 3, 6, 9}));
  EXPECT_EQ(std::vector<double>(p.values, p.values + p.count), (std::vector<double>{0, 1, 4, 9}));

  ChartGuard chart;
  Open(wb.handle, id, &chart);
  fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.type, FM_CHART_PLOT_TYPE_SCATTER);
  const fm_chart_series_info_t s = Series(chart, 0, 0);
  EXPECT_STREQ(s.x.formula, "Sheet1!$F$2:$F$5");
  EXPECT_STREQ(s.y.formula, "Sheet1!$G$2:$G$5");
}

TEST(FormulonCApiChart, TypeTitleLegendDisplayPlotEdits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);
  ChartGuard chart;

  fm_chart_plot_patch_t plot = FM_STRUCT_INIT(fm_chart_plot_patch_t);
  plot.gap_width_pct_op = FM_PATCH_OP_SET;
  plot.gap_width_pct = 50;
  plot.overlap_op = FM_PATCH_OP_SET;
  plot.overlap = -10;
  plot.vary_colors_op = FM_PATCH_OP_SET;
  plot.vary_colors = 1;
  ASSERT_EQ(fm_sheet_update_chart_plot(wb.handle, 0, id, 0, &plot), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_plot_info_t p = FM_STRUCT_INIT(fm_chart_plot_info_t);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &p), 0);
  EXPECT_EQ(p.has_gap_width_pct, 1);
  EXPECT_EQ(p.gap_width_pct, 50U);
  EXPECT_EQ(p.has_overlap, 1);
  EXPECT_EQ(p.overlap, -10);
  EXPECT_EQ(p.vary_colors, FM_TRISTATE_TRUE);
  // A plot field cannot be cleared, and an unknown op is rejected.
  fm_chart_plot_patch_t bad = FM_STRUCT_INIT(fm_chart_plot_patch_t);
  bad.gap_width_pct_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_chart_plot(wb.handle, 0, id, 0, &bad), kInvalidArgument);
  bad.gap_width_pct_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(fm_sheet_update_chart_plot(wb.handle, 0, id, 0, &bad), kInvalidArgument);
  bad = fm_chart_plot_patch_t FM_STRUCT_INIT(fm_chart_plot_patch_t);
  bad.hole_size_op = FM_PATCH_OP_SET;
  bad.hole_size = 50;
  EXPECT_EQ(fm_sheet_update_chart_plot(wb.handle, 0, id, 0, &bad), kInvalidArgument);  // not a doughnut
  EXPECT_EQ(fm_sheet_update_chart_plot(wb.handle, 0, id, 1, &plot), FM_STATUS_INDEX_OUT_OF_RANGE);

  const fm_chart_title_t title{FM_CHART_TITLE_KIND_TEXT, "Sales", nullptr};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &title), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).title.kind, FM_CHART_TITLE_KIND_TEXT);
  EXPECT_STREQ(Info(chart).title.text, "Sales");
  EXPECT_EQ(Info(chart).title.formula, nullptr);
  const fm_chart_title_t none{FM_CHART_TITLE_KIND_NONE, nullptr, nullptr};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &none), 0);
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).title.kind, FM_CHART_TITLE_KIND_NONE);
  EXPECT_EQ(Info(chart).title.text, nullptr);
  EXPECT_EQ(Info(chart).title.formula, nullptr);
  const fm_chart_title_t automatic{FM_CHART_TITLE_KIND_AUTO, nullptr, nullptr};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &automatic), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).title.kind, FM_CHART_TITLE_KIND_AUTO);
  const fm_chart_title_t bad_kind{4, nullptr, nullptr};
  EXPECT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &bad_kind), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, nullptr), kBindingNullPointer);

  const uint32_t hidden[] = {1U};
  fm_chart_legend_patch_t legend{sizeof(fm_chart_legend_patch_t),
                                 FM_PATCH_OP_SET,
                                 FM_CHART_LEGEND_POSITION_BOTTOM,
                                 FM_PATCH_OP_SET,
                                 1,
                                 FM_PATCH_OP_SET,
                                 hidden,
                                 1};
  ASSERT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, id, &legend), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_info_t info = Info(chart);
  EXPECT_EQ(info.legend_position, FM_CHART_LEGEND_POSITION_BOTTOM);
  EXPECT_EQ(info.legend_overlay, 1);
  ASSERT_EQ(info.legend_hidden_count, 1U);
  uint32_t entry = 0;
  ASSERT_EQ(fm_chart_get_legend_hidden_at(chart.handle, 0, &entry), 0);
  EXPECT_EQ(entry, 1U);
  EXPECT_EQ(fm_chart_get_legend_hidden_at(chart.handle, 1, &entry), FM_STATUS_INDEX_OUT_OF_RANGE);
  legend = fm_chart_legend_patch_t{sizeof(fm_chart_legend_patch_t),
                                   FM_PATCH_OP_SET,
                                   FM_CHART_LEGEND_POSITION_ABSENT,
                                   FM_PATCH_OP_KEEP,
                                   0,
                                   FM_PATCH_OP_KEEP,
                                   nullptr,
                                   0};
  EXPECT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, id, &legend), kInvalidArgument);
  ASSERT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, nullptr), 0);
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).legend_position, FM_CHART_LEGEND_POSITION_ABSENT);
  legend.hidden_entries_count = 2;
  EXPECT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, id, &legend), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, id, nullptr), kBindingNullPointer);
  legend = fm_chart_legend_patch_t{
      sizeof(fm_chart_legend_patch_t), FM_PATCH_OP_SET, 6, FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_KEEP, nullptr, 0};
  EXPECT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, id, &legend), kInvalidArgument);

  fm_chart_display_patch_t display{sizeof(fm_chart_display_patch_t),
                                   FM_PATCH_OP_SET,
                                   FM_BLANKS_AS_ZERO,
                                   FM_PATCH_OP_SET,
                                   0,
                                   FM_PATCH_OP_KEEP,
                                   0,
                                   FM_PATCH_OP_KEEP,
                                   0};
  ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &display), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  info = Info(chart);
  EXPECT_EQ(info.blanks_as, FM_BLANKS_AS_ZERO);
  EXPECT_EQ(info.plot_visible_only, FM_TRISTATE_FALSE);
  display = fm_chart_display_patch_t{sizeof(fm_chart_display_patch_t),
                                     FM_PATCH_OP_SET,
                                     4,
                                     FM_PATCH_OP_KEEP,
                                     0,
                                     FM_PATCH_OP_KEEP,
                                     0,
                                     FM_PATCH_OP_KEEP,
                                     0};
  EXPECT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &display), kInvalidArgument);
  display = fm_chart_display_patch_t{sizeof(fm_chart_display_patch_t),
                                     FM_PATCH_OP_KEEP,
                                     0,
                                     FM_PATCH_OP_SET,
                                     3,
                                     FM_PATCH_OP_KEEP,
                                     0,
                                     FM_PATCH_OP_KEEP,
                                     0};
  EXPECT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &display), kInvalidArgument);

  ASSERT_EQ(fm_sheet_set_chart_type(wb.handle, 0, id, 0, FM_CHART_PLOT_TYPE_LINE, FM_CHART_BAR_DIRECTION_COLUMN,
                                    FM_CHART_GROUPING_STANDARD),
            0)
      << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &p), 0);
  EXPECT_EQ(p.type, FM_CHART_PLOT_TYPE_LINE);
  EXPECT_EQ(p.series_count, 3U);
  EXPECT_STREQ(Series(chart, 0, 2).values.formula, "Sheet1!$D$2:$D$5");
  EXPECT_EQ(fm_sheet_set_chart_type(wb.handle, 0, id, 0, 99, FM_CHART_BAR_DIRECTION_COLUMN, FM_CHART_GROUPING_STANDARD),
            kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_type(wb.handle, 0, id, 0, FM_CHART_PLOT_TYPE_LINE_3D, FM_CHART_BAR_DIRECTION_COLUMN,
                                    FM_CHART_GROUPING_STANDARD),
            kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_type(wb.handle, 0, id, 0, FM_CHART_PLOT_TYPE_LINE, FM_CHART_BAR_DIRECTION_COLUMN, 4),
            kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_type(wb.handle, 0, 99, 0, FM_CHART_PLOT_TYPE_LINE, FM_CHART_BAR_DIRECTION_COLUMN, 0),
            FM_STATUS_NOT_FOUND);
}

TEST(FormulonCApiChart, AxisEdits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  const uint32_t val_id = AxisId(chart, FM_CHART_AXIS_KIND_VALUE);
  const uint32_t cat_id = AxisId(chart, FM_CHART_AXIS_KIND_CATEGORY);

  fm_chart_axis_patch_t patch = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  patch.min_op = FM_PATCH_OP_SET;
  patch.min = 0.0;
  patch.max_op = FM_PATCH_OP_SET;
  patch.max = 20.0;
  patch.major_unit_op = FM_PATCH_OP_SET;
  patch.major_unit = 5.0;
  patch.number_format_op = FM_PATCH_OP_SET;
  patch.number_format = "0.0";
  patch.title_op = FM_PATCH_OP_SET;
  patch.title = fm_chart_title_t{FM_CHART_TITLE_KIND_TEXT, "Amount", nullptr};
  patch.major_gridlines_op = FM_PATCH_OP_SET;
  patch.major_gridlines = 0;
  patch.reversed_op = FM_PATCH_OP_SET;
  patch.reversed = 1;
  patch.log_base_op = FM_PATCH_OP_SET;
  patch.log_base = 10.0;
  patch.tick_label_position_op = FM_PATCH_OP_SET;
  patch.tick_label_position = FM_CHART_TICK_LABEL_POSITION_LOW;
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_axis_info_t a = Axis(chart, val_id);
  EXPECT_EQ(a.min, 0.0);
  EXPECT_EQ(a.max, 20.0);
  EXPECT_EQ(a.major_unit, 5.0);
  EXPECT_EQ(a.log_base, 10.0);
  ASSERT_NE(a.number_format, nullptr);
  EXPECT_STREQ(a.number_format, "0.0");
  EXPECT_EQ(a.source_linked, 0);
  EXPECT_EQ(a.title.kind, FM_CHART_TITLE_KIND_TEXT);
  EXPECT_STREQ(a.title.text, "Amount");
  EXPECT_EQ(a.major_gridlines, 0);
  EXPECT_EQ(a.reversed, 1);
  EXPECT_EQ(a.tick_label_position, FM_CHART_TICK_LABEL_POSITION_LOW);

  patch = fm_chart_axis_patch_t FM_STRUCT_INIT(fm_chart_axis_patch_t);
  patch.max_op = FM_PATCH_OP_CLEAR;
  patch.log_base_op = FM_PATCH_OP_CLEAR;
  patch.title_op = FM_PATCH_OP_CLEAR;
  patch.visible_op = FM_PATCH_OP_SET;
  patch.visible = 0;
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  a = Axis(chart, val_id);
  EXPECT_EQ(a.has_max, 0);
  EXPECT_EQ(a.has_log_base, 0);
  EXPECT_EQ(a.title.kind, FM_CHART_TITLE_KIND_NONE);
  EXPECT_EQ(a.visible, 0);

  // The axis title takes the chart title kinds; AUTO leaves the text to the viewer.
  patch = fm_chart_axis_patch_t FM_STRUCT_INIT(fm_chart_axis_patch_t);
  patch.title_op = FM_PATCH_OP_SET;
  patch.title = fm_chart_title_t{FM_CHART_TITLE_KIND_AUTO, nullptr, nullptr};
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Axis(chart, val_id).title.kind, FM_CHART_TITLE_KIND_AUTO);
  patch.title.kind = FM_CHART_TITLE_KIND_AUTO + 1;
  EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), kInvalidArgument);

  // The crossing point is written on the cross axis.
  patch = fm_chart_axis_patch_t FM_STRUCT_INIT(fm_chart_axis_patch_t);
  patch.crosses_op = FM_PATCH_OP_SET;
  patch.crosses = FM_CHART_AXIS_CROSSES_AT;
  patch.crosses_at_op = FM_PATCH_OP_SET;
  patch.crosses_at = 3.0;
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  const fm_chart_axis_info_t crossing = Axis(chart, cat_id);
  const fm_chart_axis_info_t crossed = Axis(chart, val_id);
  EXPECT_TRUE(
      (crossing.crosses == FM_CHART_AXIS_CROSSES_AT && crossing.has_crosses_at == 1 && crossing.crosses_at == 3.0) ||
      (crossed.crosses == FM_CHART_AXIS_CROSSES_AT && crossed.has_crosses_at == 1 && crossed.crosses_at == 3.0));
  // CLEAR is refused on the crossing value.
  patch.crosses_at_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), kInvalidArgument);

  patch = fm_chart_axis_patch_t FM_STRUCT_INIT(fm_chart_axis_patch_t);
  patch.visible_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), kInvalidArgument);
  patch = fm_chart_axis_patch_t FM_STRUCT_INIT(fm_chart_axis_patch_t);
  patch.crosses_op = FM_PATCH_OP_SET;
  patch.crosses = FM_CHART_AXIS_CROSSES_AT + 1;
  EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), kInvalidArgument);
  // ABSENT (-1) is inside the domain; nothing below it is.
  patch.crosses = FM_CHART_AXIS_CROSSES_ABSENT - 1;
  EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), kInvalidArgument);
  patch.crosses = FM_CHART_AXIS_CROSSES_ABSENT;
  EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_get_last_error_message();

  // A set tick-label position is a member other than ABSENT; CLEAR restores `nextTo`.
  patch = fm_chart_axis_patch_t FM_STRUCT_INIT(fm_chart_axis_patch_t);
  patch.tick_label_position_op = FM_PATCH_OP_SET;
  patch.tick_label_position = FM_CHART_TICK_LABEL_POSITION_ABSENT;
  EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), kInvalidArgument);
  patch.tick_label_position = FM_CHART_TICK_LABEL_POSITION_NONE + 1;
  EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), kInvalidArgument);
  patch.tick_label_position = FM_CHART_TICK_LABEL_POSITION_NONE;
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Axis(chart, val_id).tick_label_position, FM_CHART_TICK_LABEL_POSITION_NONE);
  patch.tick_label_position_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Axis(chart, val_id).tick_label_position, FM_CHART_TICK_LABEL_POSITION_NEXT_TO);

  const fm_chart_axis_patch_t keep = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, 999, &keep), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, nullptr), kBindingNullPointer);
}

TEST(FormulonCApiChart, SeriesPointFilterAndSourceEdits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kLine);
  ChartGuard chart;

  fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.line.color_op = FM_PATCH_OP_SET;
  patch.line.color.kind = FM_COLOR_KIND_RGB;
  patch.line.color.argb = 0xFFFF0000U;
  patch.line.transparency_pct_op = FM_PATCH_OP_SET;
  patch.line.transparency_pct = 25.0;
  patch.line.width_pt_op = FM_PATCH_OP_SET;
  patch.line.width_pt = 2.0;
  patch.line.dash_op = FM_PATCH_OP_SET;
  patch.line.dash = FM_CHART_LINE_DASH_DASH;
  patch.marker_op = FM_PATCH_OP_SET;
  patch.marker.symbol_op = FM_PATCH_OP_SET;
  patch.marker.symbol = FM_CHART_MARKER_SYMBOL_CIRCLE;
  patch.marker.size_pt_op = FM_PATCH_OP_SET;
  patch.marker.size_pt = 7.0;
  patch.marker.fill.color_op = FM_PATCH_OP_SET;
  patch.marker.fill.color.kind = FM_COLOR_KIND_RGB;
  patch.marker.fill.color.argb = 0xFF00FF00U;
  patch.smooth_op = FM_PATCH_OP_SET;
  patch.smooth = 1;
  patch.labels_op = FM_PATCH_OP_SET;
  patch.labels.show_value_op = FM_PATCH_OP_SET;
  patch.labels.show_value = 1;
  patch.labels.number_format_op = FM_PATCH_OP_SET;
  patch.labels.number_format = "0";
  patch.labels.font_size_pt_op = FM_PATCH_OP_SET;
  patch.labels.font_size_pt = 9.0;
  patch.labels.font_color_op = FM_PATCH_OP_SET;
  patch.labels.font_color.kind = FM_COLOR_KIND_RGB;
  patch.labels.font_color.argb = 0xFF0000FFU;
  patch.labels.separator_op = FM_PATCH_OP_SET;
  patch.labels.separator = "; ";
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_series_info_t s = Series(chart, 0, 0);
  EXPECT_EQ(s.line.color.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(s.line.color.argb, 0xFFFF0000U);
  EXPECT_DOUBLE_EQ(s.line.transparency_pct, 25.0);
  EXPECT_DOUBLE_EQ(s.line.width_pt, 2.0);
  EXPECT_EQ(s.line.has_dash, 1);
  EXPECT_EQ(s.line.dash, FM_CHART_LINE_DASH_DASH);
  ASSERT_EQ(s.has_marker, 1);
  EXPECT_EQ(s.marker.symbol, FM_CHART_MARKER_SYMBOL_CIRCLE);
  EXPECT_EQ(s.marker.has_size_pt, 1);
  EXPECT_EQ(s.marker.size_pt, 7U);
  EXPECT_EQ(s.marker.fill.color.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(s.marker.fill.color.argb, 0xFF00FF00U);
  EXPECT_EQ(s.smooth, FM_TRISTATE_TRUE);

  // A picture marker is read-only: it needs a blip a patch cannot carry.
  fm_chart_series_patch_t picture = FM_STRUCT_INIT(fm_chart_series_patch_t);
  picture.marker_op = FM_PATCH_OP_SET;
  picture.marker.symbol_op = FM_PATCH_OP_SET;
  picture.marker.symbol = FM_CHART_MARKER_SYMBOL_PICTURE;
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &picture), kInvalidArgument);
  ASSERT_EQ(s.has_labels, 1);
  EXPECT_EQ(s.labels.show_value, 1);
  EXPECT_EQ(s.labels.show_category, 0);
  ASSERT_NE(s.labels.number_format, nullptr);
  EXPECT_STREQ(s.labels.number_format, "0");
  EXPECT_DOUBLE_EQ(s.labels.font_size_pt, 9.0);
  EXPECT_EQ(s.labels.font_color.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(s.labels.font_color.argb, 0xFF0000FFU);
  ASSERT_NE(s.labels.separator, nullptr);
  EXPECT_STREQ(s.labels.separator, "; ");

  patch = fm_chart_series_patch_t FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.marker_op = FM_PATCH_OP_CLEAR;
  patch.labels_op = FM_PATCH_OP_CLEAR;
  patch.line.width_pt_op = FM_PATCH_OP_CLEAR;
  patch.name_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  s = Series(chart, 0, 0);
  EXPECT_EQ(s.has_marker, 0);
  EXPECT_EQ(s.has_labels, 0);
  EXPECT_EQ(s.line.has_width_pt, 0);
  EXPECT_EQ(s.name.kind, FM_CHART_SOURCE_KIND_NONE);

  patch = fm_chart_series_patch_t FM_STRUCT_INIT(fm_chart_series_patch_t);
  SetRefSource(&patch.values_op, &patch.values, "");
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), kInvalidArgument);
  SetRefSource(&patch.values_op, &patch.values, "Nope!$B$2:$B$5");
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), kInvalidReference);
  patch = fm_chart_series_patch_t FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.smooth_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), kInvalidArgument);
  patch = fm_chart_series_patch_t FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.marker_op = FM_PATCH_OP_SET;
  patch.marker.symbol_op = FM_PATCH_OP_SET;
  patch.marker.symbol = FM_CHART_MARKER_SYMBOL_PICTURE + 1;
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), kInvalidArgument);
  // A nested patch whose op is KEEP is not read.
  patch.marker_op = FM_PATCH_OP_KEEP;
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), 0) << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 9, &patch), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, nullptr), kBindingNullPointer);

  fm_chart_point_patch_t point = FM_STRUCT_INIT(fm_chart_point_patch_t);
  point.line.color_op = FM_PATCH_OP_SET;
  point.line.color.kind = FM_COLOR_KIND_RGB;
  point.line.color.argb = 0xFF123456U;
  point.marker_op = FM_PATCH_OP_SET;
  point.marker.symbol_op = FM_PATCH_OP_SET;
  point.marker.symbol = FM_CHART_MARKER_SYMBOL_TRIANGLE;
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 1, 2, &point), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  s = Series(chart, 0, 1);
  ASSERT_EQ(s.point_count, 1U);
  fm_chart_point_info_t pt = FM_STRUCT_INIT(fm_chart_point_info_t);
  ASSERT_EQ(fm_chart_get_series_point_at(chart.handle, 0, 1, 0, &pt), 0);
  EXPECT_EQ(pt.idx, 2U);
  EXPECT_EQ(pt.line.color.kind, FM_COLOR_KIND_RGB);
  EXPECT_EQ(pt.line.color.argb, 0xFF123456U);
  ASSERT_EQ(pt.has_marker, 1);
  EXPECT_EQ(pt.marker.symbol, FM_CHART_MARKER_SYMBOL_TRIANGLE);
  EXPECT_EQ(fm_chart_get_series_point_at(chart.handle, 0, 1, 1, &pt), FM_STATUS_INDEX_OUT_OF_RANGE);
  // A patch that changes nothing leaves the override; clearing every member removes it.
  const fm_chart_point_patch_t keep = FM_STRUCT_INIT(fm_chart_point_patch_t);
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 1, 2, &keep), 0);
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Series(chart, 0, 1).point_count, 1U);
  fm_chart_point_patch_t clear = FM_STRUCT_INIT(fm_chart_point_patch_t);
  clear.fill.color_op = FM_PATCH_OP_CLEAR;
  clear.fill.transparency_pct_op = FM_PATCH_OP_CLEAR;
  clear.line.color_op = FM_PATCH_OP_CLEAR;
  clear.line.transparency_pct_op = FM_PATCH_OP_CLEAR;
  clear.line.width_pt_op = FM_PATCH_OP_CLEAR;
  clear.line.dash_op = FM_PATCH_OP_CLEAR;
  clear.line.none_op = FM_PATCH_OP_CLEAR;
  clear.marker_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 1, 2, &clear), 0);
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Series(chart, 0, 1).point_count, 0U);
  EXPECT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 1, 2, nullptr), kBindingNullPointer);

  const uint32_t hidden_series[] = {1U};
  const uint32_t hidden_categories[] = {0U, 2U};
  ASSERT_EQ(fm_sheet_set_chart_filter(wb.handle, 0, id, hidden_series, 1, hidden_categories, 2), 0)
      << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  ASSERT_EQ(Info(chart).hidden_category_count, 2U);
  uint32_t cat = 0;
  ASSERT_EQ(fm_chart_get_hidden_category_at(chart.handle, 1, &cat), 0);
  EXPECT_EQ(cat, 2U);
  EXPECT_EQ(fm_chart_get_hidden_category_at(chart.handle, 2, &cat), FM_STATUS_INDEX_OUT_OF_RANGE);
  bool filtered = false;
  fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  for (size_t i = 0; i < plot.series_count; ++i) {
    const fm_chart_series_info_t each = Series(chart, 0, i);
    filtered = filtered || (each.series_id == 1U && each.filtered == 1);
  }
  EXPECT_TRUE(filtered);
  ASSERT_EQ(fm_sheet_set_chart_filter(wb.handle, 0, id, nullptr, 0, nullptr, 0), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).hidden_category_count, 0U);
  EXPECT_EQ(fm_sheet_set_chart_filter(wb.handle, 0, id, nullptr, 1, nullptr, 0), kBindingNullPointer);

  patch = fm_chart_series_patch_t FM_STRUCT_INIT(fm_chart_series_patch_t);
  SetRefSource(&patch.name_op, &patch.name, "Sheet1!$D$1");
  SetRefSource(&patch.categories_op, &patch.categories, "Sheet1!$A$2:$A$5");
  SetRefSource(&patch.values_op, &patch.values, "Sheet1!$D$2:$D$5");
  uint32_t added = 0;
  ASSERT_EQ(fm_sheet_add_chart_series(wb.handle, 0, id, 0, &patch, &added), 0) << fm_get_last_error_message();
  EXPECT_EQ(added, 3U);
  Open(wb.handle, id, &chart);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.series_count, 4U);
  EXPECT_EQ(fm_sheet_add_chart_series(wb.handle, 0, id, 0, &patch, nullptr), kBindingNullPointer);
  ASSERT_EQ(fm_sheet_remove_chart_series(wb.handle, 0, id, 0, added), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.series_count, 3U);
  EXPECT_EQ(fm_sheet_remove_chart_series(wb.handle, 0, id, 0, added), FM_STATUS_NOT_FOUND);

  ASSERT_EQ(fm_sheet_set_chart_source(wb.handle, 0, id, "Sheet1!$A$1:$C$5", FM_CHART_SERIES_DIRECTION_ROWS,
                                      FM_TRISTATE_DEFAULT, FM_TRISTATE_DEFAULT),
            0)
      << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_info_t info = Info(chart);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.series_count, 4U);
  ASSERT_NE(info.source_range, nullptr);
  EXPECT_STREQ(info.source_range, "Sheet1!$A$1:$C$5");
  EXPECT_EQ(info.source_direction, FM_CHART_SERIES_DIRECTION_ROWS);
  EXPECT_EQ(fm_sheet_set_chart_source(wb.handle, 0, id, "Nope!$A$1:$B$5", FM_CHART_SERIES_DIRECTION_AUTO,
                                      FM_TRISTATE_DEFAULT, FM_TRISTATE_DEFAULT),
            kInvalidReference);
  EXPECT_EQ(
      fm_sheet_set_chart_source(wb.handle, 0, id, "Sheet1!$A$1:$C$5", 3, FM_TRISTATE_DEFAULT, FM_TRISTATE_DEFAULT),
      kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_chart_source(wb.handle, 0, id, nullptr, 0, 0, 0), kBindingNullPointer);
}

TEST(FormulonCApiChart, ThemeColorsRoundTripAndAreValidated) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kLine);

  fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.line.color_op = FM_PATCH_OP_SET;
  patch.line.color.kind = FM_COLOR_KIND_THEME;
  patch.line.color.theme = 4;
  patch.line.color.tint = 0.4;
  patch.labels_op = FM_PATCH_OP_SET;
  patch.labels.font_color_op = FM_PATCH_OP_SET;
  patch.labels.font_color.kind = FM_COLOR_KIND_THEME;
  patch.labels.font_color.theme = 1;
  patch.labels.font_color.tint = -0.25;
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), 0) << fm_get_last_error_message();

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_get_last_error_message();
  ChartGuard chart;
  Open(reloaded.handle, id, &chart);
  const fm_chart_series_info_t s = Series(chart, 0, 0);
  EXPECT_EQ(s.line.color.kind, FM_COLOR_KIND_THEME);
  EXPECT_EQ(s.line.color.theme, 4U);
  EXPECT_DOUBLE_EQ(s.line.color.tint, 0.4);
  EXPECT_EQ(s.labels.font_color.kind, FM_COLOR_KIND_THEME);
  EXPECT_EQ(s.labels.font_color.theme, 1U);
  EXPECT_DOUBLE_EQ(s.labels.font_color.tint, -0.25);

  struct Bad {
    uint8_t kind;
    uint32_t theme;
    double tint;
  };
  const Bad bad[] = {
      {FM_COLOR_KIND_INDEXED, 0, 0.0},
      {FM_COLOR_KIND_AUTO, 0, 0.0},
      {FM_COLOR_KIND_THEME, 12, 0.0},
      {FM_COLOR_KIND_THEME, 4, 1.5},
      {FM_COLOR_KIND_THEME, 4, -1.5},
      {FM_COLOR_KIND_THEME, 4, std::numeric_limits<double>::quiet_NaN()},
      {FM_COLOR_KIND_THEME, 4, std::numeric_limits<double>::infinity()},
  };
  for (const Bad& b : bad) {
    fm_chart_series_patch_t p = FM_STRUCT_INIT(fm_chart_series_patch_t);
    p.line.color_op = FM_PATCH_OP_SET;
    p.line.color.kind = b.kind;
    p.line.color.theme = b.theme;
    p.line.color.tint = b.tint;
    EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &p), kInvalidArgument)
        << "kind=" << int(b.kind) << " theme=" << b.theme << " tint=" << b.tint;
  }
}

TEST(FormulonCApiChart, PreviewLeavesWorkbookUnchanged) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);

  fm_chart_preview_overrides_spec_t o = FM_STRUCT_INIT(fm_chart_preview_overrides_spec_t);
  o.source_op = FM_PATCH_OP_SET;
  o.source = "Sheet1!$A$1:$B$5";
  DataGuard preview;
  ASSERT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, id, &o, &present, &preview.handle), 0)
      << fm_get_last_error_message();
  fm_chart_data_info_t info = FM_STRUCT_INIT(fm_chart_data_info_t);
  ASSERT_EQ(fm_chart_data_get_info(preview.handle, &info), 0);
  EXPECT_EQ(info.series_count, 1U);

  o = fm_chart_preview_overrides_spec_t FM_STRUCT_INIT(fm_chart_preview_overrides_spec_t);
  o.direction_op = FM_PATCH_OP_SET;
  o.direction = FM_CHART_SERIES_DIRECTION_ROWS;
  DataGuard rows;
  ASSERT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, id, &o, &present, &rows.handle), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_chart_data_get_info(rows.handle, &info), 0);
  EXPECT_EQ(info.series_count, 4U);
  EXPECT_EQ(info.point_count, 3U);

  // Applying the same change for real gives the same data.
  ASSERT_EQ(fm_sheet_set_chart_source(wb.handle, 0, id, "Sheet1!$A$1:$D$5", FM_CHART_SERIES_DIRECTION_ROWS,
                                      FM_TRISTATE_DEFAULT, FM_TRISTATE_DEFAULT),
            0);
  DataGuard applied;
  OpenData(wb.handle, id, &applied);
  for (size_t i = 0; i < 4; ++i) {
    EXPECT_EQ(Values(rows, i), Values(applied, i));
  }

  DataGuard current;
  o = fm_chart_preview_overrides_spec_t FM_STRUCT_INIT(fm_chart_preview_overrides_spec_t);
  o.source_op = FM_PATCH_OP_SET;
  o.source = "Sheet1!$A$1:$B$5";
  ASSERT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, id, &o, &present, &current.handle), 0);
  DataGuard after;
  OpenData(wb.handle, id, &after);
  ASSERT_EQ(fm_chart_data_get_info(after.handle, &info), 0);
  EXPECT_EQ(info.series_count, 4U);

  o = fm_chart_preview_overrides_spec_t FM_STRUCT_INIT(fm_chart_preview_overrides_spec_t);
  o.blanks_as_op = FM_PATCH_OP_CLEAR;
  fm_chart_data_t* raw = reinterpret_cast<fm_chart_data_t*>(&o);  // reset to NULL on failure
  EXPECT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, id, &o, &present, &raw), kInvalidArgument);
  EXPECT_EQ(raw, nullptr);
  o = fm_chart_preview_overrides_spec_t FM_STRUCT_INIT(fm_chart_preview_overrides_spec_t);
  o.source_op = FM_PATCH_OP_SET;
  o.source = "Nope!$A$1:$B$5";
  EXPECT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, id, &o, &present, &raw), kInvalidReference);
  o.source = nullptr;  // a SET string that is NULL is an invalid argument
  EXPECT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, id, &o, &present, &raw), kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, id, nullptr, &present, &raw), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, id, &o, &present, nullptr), kBindingNullPointer);
}

TEST(FormulonCApiChart, DoughnutHoleSize) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kDoughnut, "Sheet1!$A$1:$B$5");
  fm_chart_plot_patch_t patch = FM_STRUCT_INIT(fm_chart_plot_patch_t);
  patch.hole_size_op = FM_PATCH_OP_SET;
  patch.hole_size = 60;
  ASSERT_EQ(fm_sheet_update_chart_plot(wb.handle, 0, id, 0, &patch), 0) << fm_get_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.type, FM_CHART_PLOT_TYPE_DOUGHNUT);
  EXPECT_EQ(plot.has_hole_size, 1);
  EXPECT_EQ(plot.hole_size, 60U);
  EXPECT_EQ(Info(chart).axis_count, 0U);
}

TEST(FormulonCApiChart, CaptureRestoreAndRemoveThroughDrawingOperations) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);
  const fm_chart_title_t before{FM_CHART_TITLE_KIND_TEXT, "Before", nullptr};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &before), 0);

  BufferGuard snap;
  ASSERT_EQ(fm_sheet_snapshot_drawing(wb.handle, 0, id, &snap.data, &snap.len), 0) << fm_get_last_error_message();
  const fm_chart_title_t after{FM_CHART_TITLE_KIND_TEXT, "After", nullptr};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &after), 0);

  // Same id: the chart part's bytes come back.
  uint32_t restored = 0;
  ASSERT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len, 0, &restored), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(restored, id);
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  EXPECT_STREQ(Info(chart).title.text, "Before");

  // Move and reorder like any drawing object.
  fm_drawing_patch_t move = FM_STRUCT_INIT(fm_drawing_patch_t);
  move.anchor.from_row_op = FM_PATCH_OP_SET;
  move.anchor.from_row = 10;
  move.anchor.from_col_op = FM_PATCH_OP_SET;
  move.anchor.from_col = 2;
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &move), 0) << fm_get_last_error_message();
  fm_drawing_object_info_t moved = FM_STRUCT_INIT(fm_drawing_object_info_t);
  int32_t moved_present = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object(wb.handle, 0, id, &moved_present, &moved), 0);
  ASSERT_EQ(moved_present, 1);
  EXPECT_EQ(moved.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(moved.anchor.from_row, 10U);
  EXPECT_EQ(moved.anchor.from_col, 2U);

  // A copy with a fresh id reads as its own chart.
  uint32_t copy = 0;
  ASSERT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len, FM_DRAWING_RESTORE_NEW_ID, &copy), 0)
      << fm_get_last_error_message();
  EXPECT_NE(copy, id);
  ChartGuard copied;
  Open(wb.handle, copy, &copied);
  EXPECT_STREQ(Info(copied).title.text, "Before");
  ASSERT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, copy, 0), 0) << fm_get_last_error_message();

  ASSERT_EQ(fm_sheet_remove_drawing(wb.handle, 0, id), 0) << fm_get_last_error_message();
  fm_chart_t* gone = nullptr;
  int32_t present = 1;
  EXPECT_EQ(fm_sheet_get_chart(wb.handle, 0, id, &present, &gone), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(gone, nullptr);
  ASSERT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len, 0, &restored), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(restored, id);
  Open(wb.handle, id, &chart);
  EXPECT_STREQ(Info(chart).title.text, "Before");

  // The edited chart survives a save and reload.
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_get_last_error_message();
  ChartGuard again;
  Open(reloaded.handle, id, &again);
  EXPECT_STREQ(Info(again).title.text, "Before");
}

TEST(FormulonCApiChart, OpenAndAccessorsRejectBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);

  fm_chart_t* chart = reinterpret_cast<fm_chart_t*>(&wb);
  int32_t present = 1;
  EXPECT_EQ(fm_sheet_get_chart(wb.handle, 0, 99, &present, &chart), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(chart, nullptr);
  EXPECT_EQ(fm_sheet_get_chart(wb.handle, 4, id, &present, &chart), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_chart(wb.handle, 0, id, &present, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_chart(wb.handle, 0, id, nullptr, &chart), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_chart(nullptr, 0, id, &present, &chart), kBindingNullPointer);
  fm_chart_data_t* data = nullptr;
  present = 1;
  EXPECT_EQ(fm_sheet_get_chart_data(wb.handle, 0, 99, &present, &data), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(data, nullptr);
  EXPECT_EQ(fm_sheet_get_chart_data(wb.handle, 0, id, &present, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_chart_data(wb.handle, 0, id, nullptr, &data), kBindingNullPointer);
  fm_chart_destroy(nullptr);
  fm_chart_data_destroy(nullptr);

  ChartGuard guard;
  Open(wb.handle, id, &guard);
  fm_chart_info_t info = FM_STRUCT_INIT(fm_chart_info_t);
  EXPECT_EQ(fm_chart_get_info(nullptr, &info), kBindingNullPointer);
  EXPECT_EQ(fm_chart_get_info(guard.handle, nullptr), kBindingNullPointer);
  fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
  EXPECT_EQ(fm_chart_get_plot_at(guard.handle, 1, &plot), FM_STATUS_INDEX_OUT_OF_RANGE);
  uint32_t u = 0;
  EXPECT_EQ(fm_chart_get_plot_axis_id_at(guard.handle, 0, 2, &u), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_chart_get_plot_axis_id_at(guard.handle, 1, 0, &u), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_chart_series_info_t s = FM_STRUCT_INIT(fm_chart_series_info_t);
  EXPECT_EQ(fm_chart_get_series_at(guard.handle, 0, 3, &s), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_chart_literal_info_t lit = FM_STRUCT_INIT(fm_chart_literal_info_t);
  EXPECT_EQ(fm_chart_get_series_literal_at(guard.handle, 0, 0, FM_CHART_ROLE_VALUES, 0, &lit),
            FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_chart_get_series_literal_at(guard.handle, 0, 0, 6, 0, &lit), kInvalidArgument);
  EXPECT_EQ(fm_chart_get_series_literal_at(guard.handle, 0, 0, -1, 0, &lit), kInvalidArgument);
  fm_chart_point_info_t pt = FM_STRUCT_INIT(fm_chart_point_info_t);
  EXPECT_EQ(fm_chart_get_series_point_at(guard.handle, 0, 0, 0, &pt), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_chart_axis_info_t axis = FM_STRUCT_INIT(fm_chart_axis_info_t);
  EXPECT_EQ(fm_chart_get_axis_at(guard.handle, 2, &axis), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_chart_get_axis_at(guard.handle, 0, nullptr), kBindingNullPointer);

  DataGuard dg;
  OpenData(wb.handle, id, &dg);
  fm_chart_data_series_info_t ds = FM_STRUCT_INIT(fm_chart_data_series_info_t);
  EXPECT_EQ(fm_chart_data_get_series_at(dg.handle, 3, &ds), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_chart_data_points_info_t pts = FM_STRUCT_INIT(fm_chart_data_points_info_t);
  EXPECT_EQ(fm_chart_data_get_series_points(dg.handle, 3, &pts), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_chart_data_get_series_points(nullptr, 0, &pts), kBindingNullPointer);
  size_t n = 0;
  EXPECT_EQ(fm_chart_data_get_category_level_count(dg.handle, 0, &n), FM_STATUS_INDEX_OUT_OF_RANGE);
  const char* text = nullptr;
  EXPECT_EQ(fm_chart_data_get_category_level_at(dg.handle, 0, 0, &text), FM_STATUS_INDEX_OUT_OF_RANGE);
  int32_t has_v = 0;
  double v = 0.0;
  EXPECT_EQ(fm_chart_data_get_category_at(dg.handle, 0, nullptr, &has_v, &v), kBindingNullPointer);
  EXPECT_EQ(fm_chart_data_get_category_at(dg.handle, 0, &text, nullptr, &v), kBindingNullPointer);
}

TEST(FormulonCApiChart, InsertRejectsBadOptions) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  uint32_t id = 0;
  fm_chart_insert_spec_t o = Options(kColumn, "Sheet1!$A$1:$D$5");
  o.type = 99;
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, &id), kInvalidArgument);
  o = Options(kColumn, "Sheet1!$A$1:$D$5");
  o.legend = 6;
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, &id), kInvalidArgument);
  o = Options(kColumn, "Sheet1!$A$1:$D$5");
  o.header_row = 3;  // outside fm_tristate_t
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, &id), kInvalidArgument);
  o = Options(kColumn, "Sheet1!$A$1:$D$5");
  o.placement.anchor.kind = 4;  // outside fm_anchor_kind_t
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, &id), kInvalidArgument);
  o = Options(kColumn, "Nope!$A$1:$D$5");
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, &id), kInvalidReference);
  o = Options(kColumn, nullptr);
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, &id), kBindingNullPointer);
  o = Options(kColumn, "Sheet1!$A$1:$D$5");
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, nullptr, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 3, &o, &id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  size_t count = 9;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);

  // A NONE title is no title; a given name and description are kept.
  o.has_title = 1;
  o.title.kind = FM_CHART_TITLE_KIND_NONE;
  o.placement.name = "Sales chart";
  o.placement.description = "Quarterly";
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, &id), 0) << fm_get_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  const fm_chart_info_t info = Info(chart);
  EXPECT_EQ(info.title.kind, FM_CHART_TITLE_KIND_NONE);
  EXPECT_EQ(info.auto_title_deleted, FM_TRISTATE_TRUE);
  fm_drawing_object_info_t obj = FM_STRUCT_INIT(fm_drawing_object_info_t);
  int32_t obj_present = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object(wb.handle, 0, id, &obj_present, &obj), 0);
  ASSERT_EQ(obj_present, 1);
  EXPECT_STREQ(obj.name, "Sales chart");
  EXPECT_STREQ(obj.description, "Quarterly");
}

TEST(FormulonCApiChart, ChartexAndPivotChartsRefuseEdits) {
  int32_t present = 0;
  WorkbookGuard cx;
  LoadFixture("chart_chartex.xlsx", &cx);
  const std::vector<Located> cx_charts = AllCharts(cx.handle);
  ASSERT_FALSE(cx_charts.empty());
  const Located c = cx_charts[0];
  ChartGuard chart;
  Open(cx.handle, c.id, &chart, c.sheet);
  EXPECT_EQ(Info(chart).is_extended, 1);
  fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.type, FM_CHART_PLOT_TYPE_EXTENDED);
  EXPECT_EQ(plot.extended_layout, FM_CHART_EXTENDED_LAYOUT_WATERFALL);
  const fm_chart_title_t title{FM_CHART_TITLE_KIND_TEXT, "T", nullptr};
  EXPECT_EQ(fm_sheet_set_chart_title(cx.handle, c.sheet, c.id, &title), kUnsupportedKind);
  const fm_chart_display_patch_t display{sizeof(fm_chart_display_patch_t),
                                         FM_PATCH_OP_SET,
                                         FM_BLANKS_AS_GAP,
                                         FM_PATCH_OP_SET,
                                         1,
                                         FM_PATCH_OP_KEEP,
                                         0,
                                         FM_PATCH_OP_KEEP,
                                         0};
  EXPECT_EQ(fm_sheet_update_chart_display(cx.handle, c.sheet, c.id, &display), kUnsupportedKind);
  fm_chart_preview_overrides_spec_t o = FM_STRUCT_INIT(fm_chart_preview_overrides_spec_t);
  fm_chart_data_t* preview = nullptr;
  EXPECT_EQ(fm_sheet_get_chart_preview(cx.handle, c.sheet, c.id, &o, &present, &preview), kUnsupportedKind);
  DataGuard data;
  ASSERT_EQ(fm_sheet_get_chart_data(cx.handle, c.sheet, c.id, &present, &data.handle), 0)
      << fm_get_last_error_message();

  WorkbookGuard pv;
  LoadFixture("chart_pivot.xlsx", &pv);
  const std::vector<Located> pv_charts = AllCharts(pv.handle);
  ASSERT_FALSE(pv_charts.empty());
  const Located p = pv_charts[0];
  Open(pv.handle, p.id, &chart, p.sheet);
  EXPECT_EQ(Info(chart).is_pivot_chart, 1);
  fm_chart_series_patch_t rename = FM_STRUCT_INIT(fm_chart_series_patch_t);
  SetRefSource(&rename.name_op, &rename.name, "Sheet1!$A$1");
  EXPECT_EQ(fm_sheet_update_chart_series(pv.handle, p.sheet, p.id, 0, 0, &rename), kUnsupportedKind);
  uint32_t added = 0;
  EXPECT_EQ(fm_sheet_add_chart_series(pv.handle, p.sheet, p.id, 0, &rename, &added), kUnsupportedKind);
  EXPECT_EQ(fm_sheet_remove_chart_series(pv.handle, p.sheet, p.id, 0, 0), kUnsupportedKind);
  const uint32_t hidden[] = {0U};
  EXPECT_EQ(fm_sheet_set_chart_filter(pv.handle, p.sheet, p.id, hidden, 1, nullptr, 0), kUnsupportedKind);
  EXPECT_EQ(fm_sheet_set_chart_source(pv.handle, p.sheet, p.id, "Sheet1!$A$1:$B$3", FM_CHART_SERIES_DIRECTION_AUTO,
                                      FM_TRISTATE_DEFAULT, FM_TRISTATE_DEFAULT),
            kUnsupportedKind);
  EXPECT_EQ(fm_sheet_set_chart_title(pv.handle, p.sheet, p.id, &title), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiChart, LiteralsAndMultiLevelCategoriesFromExcelFiles) {
  int32_t present = 0;
  WorkbookGuard lit;
  LoadFixture("chart_cache.xlsx", &lit);
  size_t literals = 0;
  for (const Located& c : AllCharts(lit.handle)) {
    ChartGuard chart;
    Open(lit.handle, c.id, &chart, c.sheet);
    const fm_chart_info_t info = Info(chart);
    for (size_t p = 0; p < info.plot_count; ++p) {
      fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
      ASSERT_EQ(fm_chart_get_plot_at(chart.handle, p, &plot), 0);
      for (size_t i = 0; i < plot.series_count; ++i) {
        const fm_chart_series_info_t s = Series(chart, p, i);
        const fm_chart_data_source_t* sources[] = {&s.name, &s.categories, &s.values, &s.x, &s.y, &s.size};
        for (int32_t role = FM_CHART_ROLE_NAME; role <= FM_CHART_ROLE_SIZE; ++role) {
          const fm_chart_data_source_t& src = *sources[role];
          if (src.kind != FM_CHART_SOURCE_KIND_LITERAL) {
            continue;
          }
          ASSERT_GT(src.literal_count, 0U);
          for (size_t k = 0; k < src.literal_count; ++k) {
            fm_chart_literal_info_t l = FM_STRUCT_INIT(fm_chart_literal_info_t);
            ASSERT_EQ(fm_chart_get_series_literal_at(chart.handle, p, i, role, k, &l), 0);
            ASSERT_NE(l.text, nullptr);
            ++literals;
          }
          fm_chart_literal_info_t l = FM_STRUCT_INIT(fm_chart_literal_info_t);
          EXPECT_EQ(fm_chart_get_series_literal_at(chart.handle, p, i, role, src.literal_count, &l),
                    FM_STATUS_INDEX_OUT_OF_RANGE);
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
    ASSERT_EQ(fm_sheet_get_chart_data(multi.handle, c.sheet, c.id, &present, &data.handle), 0)
        << fm_get_last_error_message();
    fm_chart_data_info_t info = FM_STRUCT_INIT(fm_chart_data_info_t);
    ASSERT_EQ(fm_chart_data_get_info(data.handle, &info), 0);
    for (size_t level = 0; level < info.category_level_count; ++level) {
      size_t n = 0;
      ASSERT_EQ(fm_chart_data_get_category_level_count(data.handle, level, &n), 0);
      for (size_t k = 0; k < n; ++k) {
        const char* label = nullptr;
        ASSERT_EQ(fm_chart_data_get_category_level_at(data.handle, level, k, &label), 0);
        ASSERT_NE(label, nullptr);
      }
      const char* label = nullptr;
      EXPECT_EQ(fm_chart_data_get_category_level_at(data.handle, level, n, &label), FM_STATUS_INDEX_OUT_OF_RANGE);
      found = true;
    }
  }
  EXPECT_TRUE(found);
}

/// A display patch touching only `c16r3:dispNaAsBlank`.
fm_chart_display_patch_t NaPatch(fm_patch_op_t op, int32_t value) {
  return fm_chart_display_patch_t{
      sizeof(fm_chart_display_patch_t), FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_KEEP, 0, op, value, FM_PATCH_OP_KEEP, 0};
}

int32_t NaAsBlank(fm_workbook_t* wb, uint32_t id, size_t sheet = 0) {
  ChartGuard chart;
  Open(wb, id, &chart, sheet);
  return Info(chart).na_as_blank;
}

/// The part path of chart `id` on sheet 0.
std::string ChartPath(fm_workbook_t* wb, uint32_t id) {
  auto all = formulon::all_chart_parts(wb->workbook());
  EXPECT_TRUE(static_cast<bool>(all));
  if (all) {
    for (const formulon::SheetChartPartRef& e : all.value()) {
      if (e.sheet == 0 && e.ref.object_id == id) {
        return e.ref.part_path;
      }
    }
  }
  ADD_FAILURE() << "no chart " << id;
  return {};
}

TEST(FormulonCApiChart, NaAsBlankReadsFromTemplatesAndExcelFiles) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kLine);
  EXPECT_EQ(NaAsBlank(wb.handle, id), FM_TRISTATE_TRUE);

  WorkbookGuard line;
  LoadFixture("chart_line.xlsx", &line);
  const std::vector<Located> charts = AllCharts(line.handle);
  ASSERT_FALSE(charts.empty());
  EXPECT_EQ(NaAsBlank(line.handle, charts[0].id, charts[0].sheet), FM_TRISTATE_TRUE);

  // A file without the extension reads ABSENT.
  const fm_chart_display_patch_t clear = NaPatch(FM_PATCH_OP_CLEAR, 0);
  ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &clear), 0) << fm_get_last_error_message();
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_get_last_error_message();
  EXPECT_EQ(NaAsBlank(reloaded.handle, id), FM_TRISTATE_DEFAULT);

  // `val="0"`, which Excel never writes, reads FALSE.
  const fm_chart_display_patch_t on = NaPatch(FM_PATCH_OP_SET, 1);
  ASSERT_EQ(fm_sheet_update_chart_display(reloaded.handle, 0, id, &on), 0) << fm_get_last_error_message();
  const std::string path = ChartPath(reloaded.handle, id);
  const formulon::PassthroughPart* part = formulon::io::ooxml::find_passthrough_part(reloaded.handle->workbook(), path);
  ASSERT_NE(part, nullptr);
  std::string xml(part->bytes.begin(), part->bytes.end());
  const std::string on_text = "dispNaAsBlank val=\"1\"";
  const std::size_t at = xml.find(on_text);
  ASSERT_NE(at, std::string::npos);
  xml.replace(at, on_text.size(), "dispNaAsBlank val=\"0\"");
  ASSERT_TRUE(static_cast<bool>(
      reloaded.handle->workbook().replace_passthrough_part(path, std::vector<std::uint8_t>(xml.begin(), xml.end()))));
  EXPECT_EQ(NaAsBlank(reloaded.handle, id), FM_TRISTATE_FALSE);
}

// Each member reaches the part as its `c:dispBlanksAs` token and reads back as
// the same member after a save and load; a removed element reads ABSENT.
TEST(FormulonCApiChart, BlanksAsWritesItsSchemaTokenAndRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kLine);
  const struct {
    int32_t blanks_as;
    const char* token;
  } kCases[] = {{FM_BLANKS_AS_ZERO, "zero"}, {FM_BLANKS_AS_GAP, "gap"}, {FM_BLANKS_AS_SPAN, "span"}};
  for (const auto& c : kCases) {
    fm_chart_display_patch_t patch = FM_STRUCT_INIT(fm_chart_display_patch_t);
    patch.blanks_as_op = FM_PATCH_OP_SET;
    patch.blanks_as = c.blanks_as;
    ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();
    const formulon::PassthroughPart* part =
        formulon::io::ooxml::find_passthrough_part(wb.handle->workbook(), ChartPath(wb.handle, id));
    ASSERT_NE(part, nullptr);
    const std::string xml(part->bytes.begin(), part->bytes.end());
    EXPECT_NE(xml.find(std::string("dispBlanksAs val=\"") + c.token + "\""), std::string::npos) << c.token;

    BufferGuard saved;
    ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
    WorkbookGuard reloaded;
    ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_get_last_error_message();
    ChartGuard chart;
    Open(reloaded.handle, id, &chart);
    EXPECT_EQ(Info(chart).blanks_as, c.blanks_as) << c.token;
  }
}

TEST(FormulonCApiChart, NaAsBlankPatchSetClearAndKeep) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kLine);

  // Setting false drops the extension, as Excel writes it, so it reads ABSENT.
  fm_chart_display_patch_t patch = NaPatch(FM_PATCH_OP_SET, 0);
  ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();
  EXPECT_EQ(NaAsBlank(wb.handle, id), FM_TRISTATE_DEFAULT);
  patch = NaPatch(FM_PATCH_OP_SET, 1);
  ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();
  EXPECT_EQ(NaAsBlank(wb.handle, id), FM_TRISTATE_TRUE);

  // A zeroed patch changes nothing.
  patch = fm_chart_display_patch_t FM_STRUCT_INIT(fm_chart_display_patch_t);
  ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  fm_chart_info_t info = Info(chart);
  EXPECT_EQ(info.na_as_blank, FM_TRISTATE_TRUE);
  EXPECT_EQ(info.blanks_as, FM_BLANKS_AS_GAP);
  EXPECT_EQ(info.plot_visible_only, FM_TRISTATE_TRUE);

  // Each field moves on its own.
  patch = fm_chart_display_patch_t{sizeof(fm_chart_display_patch_t),
                                   FM_PATCH_OP_CLEAR,
                                   0,
                                   FM_PATCH_OP_KEEP,
                                   0,
                                   FM_PATCH_OP_KEEP,
                                   0,
                                   FM_PATCH_OP_KEEP,
                                   0};
  ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  info = Info(chart);
  EXPECT_EQ(info.blanks_as, FM_BLANKS_AS_ABSENT);
  EXPECT_EQ(info.plot_visible_only, FM_TRISTATE_TRUE);
  EXPECT_EQ(info.na_as_blank, FM_TRISTATE_TRUE);

  patch = NaPatch(FM_PATCH_OP_CLEAR, 0);
  ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();
  EXPECT_EQ(NaAsBlank(wb.handle, id), FM_TRISTATE_DEFAULT);
}

TEST(FormulonCApiChart, DisplayPatchRejectsBadFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kLine);

  EXPECT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, nullptr), kBindingNullPointer);
  const fm_chart_display_patch_t bad[] = {
      // An op outside KEEP / CLEAR / SET.
      {sizeof(fm_chart_display_patch_t), static_cast<fm_patch_op_t>(3), FM_BLANKS_AS_GAP, FM_PATCH_OP_KEEP, 0,
       FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_KEEP, 0},
      {sizeof(fm_chart_display_patch_t), FM_PATCH_OP_KEEP, 0, static_cast<fm_patch_op_t>(-1), 0, FM_PATCH_OP_KEEP, 0,
       FM_PATCH_OP_KEEP, 0},
      {sizeof(fm_chart_display_patch_t), FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_KEEP, 0, static_cast<fm_patch_op_t>(3), 1,
       FM_PATCH_OP_KEEP, 0},
      // ABSENT is not a value; CLEAR removes the element.
      {sizeof(fm_chart_display_patch_t), FM_PATCH_OP_SET, FM_BLANKS_AS_ABSENT, FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_KEEP, 0,
       FM_PATCH_OP_KEEP, 0},
      {sizeof(fm_chart_display_patch_t), FM_PATCH_OP_SET, 4, FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_KEEP,
       0},
      {sizeof(fm_chart_display_patch_t), FM_PATCH_OP_SET, -1, FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_KEEP, 0,
       FM_PATCH_OP_KEEP, 0},
      // Flags are 0/1.
      {sizeof(fm_chart_display_patch_t), FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_SET, 2, FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_KEEP,
       0},
      {sizeof(fm_chart_display_patch_t), FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_SET, 2, FM_PATCH_OP_KEEP,
       0},
      {sizeof(fm_chart_display_patch_t), FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_KEEP, 0, FM_PATCH_OP_SET, -1,
       FM_PATCH_OP_KEEP, 0},
  };
  for (const fm_chart_display_patch_t& p : bad) {
    EXPECT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &p), kInvalidArgument)
        << p.blanks_as_op << "/" << p.blanks_as << " " << p.plot_visible_only_op << "/" << p.plot_visible_only << " "
        << p.na_as_blank_op << "/" << p.na_as_blank;
  }
  // A rejected patch leaves the chart unchanged.
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  const fm_chart_info_t info = Info(chart);
  EXPECT_EQ(info.blanks_as, FM_BLANKS_AS_GAP);
  EXPECT_EQ(info.plot_visible_only, FM_TRISTATE_TRUE);
  EXPECT_EQ(info.na_as_blank, FM_TRISTATE_TRUE);
}

TEST(FormulonCApiChart, PreviewNaAsBlankOverrideTurnsNaIntoBlank) {
  int32_t present = 0;
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2, 1, "=NA()"), 0);  // B3, series S1 point 1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  const uint32_t id = Insert(wb.handle, kLine);
  const fm_chart_display_patch_t off = NaPatch(FM_PATCH_OP_CLEAR, 0);
  ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &off), 0) << fm_get_last_error_message();

  fm_chart_preview_overrides_spec_t o = FM_STRUCT_INIT(fm_chart_preview_overrides_spec_t);
  o.blanks_as_op = FM_PATCH_OP_SET;
  o.blanks_as = FM_BLANKS_AS_ZERO;
  o.na_as_blank_op = FM_PATCH_OP_SET;
  o.na_as_blank = 1;
  DataGuard blank;
  ASSERT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, id, &o, &present, &blank.handle), 0)
      << fm_get_last_error_message();
  fm_chart_data_points_info_t p = FM_STRUCT_INIT(fm_chart_data_points_info_t);
  ASSERT_EQ(fm_chart_data_get_series_points(blank.handle, 0, &p), 0);
  ASSERT_GT(p.count, 1U);
  EXPECT_EQ(p.states[1], FM_CHART_POINT_STATE_VALUE);
  EXPECT_EQ(p.values[1], 0.0);

  o.na_as_blank = 0;
  DataGuard na;
  ASSERT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, id, &o, &present, &na.handle), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_chart_data_get_series_points(na.handle, 0, &p), 0);
  ASSERT_GT(p.count, 1U);
  EXPECT_NE(p.states[1], FM_CHART_POINT_STATE_VALUE);

  // The preview leaves the chart's own setting alone; CLEAR is rejected.
  EXPECT_EQ(NaAsBlank(wb.handle, id), FM_TRISTATE_DEFAULT);
  o.na_as_blank_op = FM_PATCH_OP_CLEAR;
  fm_chart_data_t* raw = nullptr;
  EXPECT_EQ(fm_sheet_get_chart_preview(wb.handle, 0, id, &o, &present, &raw), kInvalidArgument);
  EXPECT_EQ(raw, nullptr);
}

TEST(FormulonCApiChart, SetLegendReplacesWholeRecordAndNullRemoves) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);
  ChartGuard chart;

  const uint32_t hidden[] = {1U, 2U};
  fm_chart_legend_t legend = FM_STRUCT_INIT(fm_chart_legend_t);
  legend.position = FM_CHART_LEGEND_POSITION_TOP;
  legend.overlay = 1;
  legend.hidden_entries = hidden;
  legend.hidden_entries_count = 2;
  ASSERT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_info_t info = Info(chart);
  EXPECT_EQ(info.legend_position, FM_CHART_LEGEND_POSITION_TOP);
  EXPECT_EQ(info.legend_overlay, 1);
  EXPECT_EQ(info.legend_hidden_count, 2U);

  // A second record replaces the whole legend, the hidden entries included.
  legend = FM_STRUCT_INIT(fm_chart_legend_t);
  legend.position = FM_CHART_LEGEND_POSITION_BOTTOM;
  ASSERT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  info = Info(chart);
  EXPECT_EQ(info.legend_position, FM_CHART_LEGEND_POSITION_BOTTOM);
  EXPECT_EQ(info.legend_overlay, 0);
  EXPECT_EQ(info.legend_hidden_count, 0U);

  // NULL removes the legend; a second NULL is a no-op.
  ASSERT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, nullptr), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).legend_position, FM_CHART_LEGEND_POSITION_ABSENT);
  EXPECT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, nullptr), 0) << fm_get_last_error_message();

  // A record adds a legend to a chart without one.
  legend.position = FM_CHART_LEGEND_POSITION_LEFT;
  ASSERT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).legend_position, FM_CHART_LEGEND_POSITION_LEFT);
}

// Each member reaches the part as its own `c:legendPos` token and reads back
// as the same member after a save and load.
TEST(FormulonCApiChart, LegendPositionWritesItsSchemaTokenAndRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);
  const struct {
    int32_t position;
    const char* token;
  } kCases[] = {{FM_CHART_LEGEND_POSITION_RIGHT, "r"},
                {FM_CHART_LEGEND_POSITION_BOTTOM, "b"},
                {FM_CHART_LEGEND_POSITION_TOP_RIGHT, "tr"},
                {FM_CHART_LEGEND_POSITION_LEFT, "l"},
                {FM_CHART_LEGEND_POSITION_TOP, "t"}};
  for (const auto& c : kCases) {
    fm_chart_legend_t legend = FM_STRUCT_INIT(fm_chart_legend_t);
    legend.position = c.position;
    ASSERT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), 0) << fm_get_last_error_message();
    const formulon::PassthroughPart* part =
        formulon::io::ooxml::find_passthrough_part(wb.handle->workbook(), ChartPath(wb.handle, id));
    ASSERT_NE(part, nullptr);
    const std::string xml(part->bytes.begin(), part->bytes.end());
    EXPECT_NE(xml.find(std::string("legendPos val=\"") + c.token + "\""), std::string::npos) << c.token;

    BufferGuard saved;
    ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
    WorkbookGuard reloaded;
    ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_get_last_error_message();
    ChartGuard chart;
    Open(reloaded.handle, id, &chart);
    EXPECT_EQ(Info(chart).legend_position, c.position) << c.token;
  }
}

TEST(FormulonCApiChart, SetLegendRejectsInvalidRecords) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);
  ChartGuard chart;

  fm_chart_legend_t legend = FM_STRUCT_INIT(fm_chart_legend_t);
  legend.position = FM_CHART_LEGEND_POSITION_ABSENT;
  EXPECT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), kInvalidArgument);
  legend.position = 6;
  EXPECT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), kInvalidArgument);
  legend.position = FM_CHART_LEGEND_POSITION_TOP;
  legend.overlay = 2;
  EXPECT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), kInvalidArgument);
  legend.overlay = 0;
  legend.hidden_entries = nullptr;
  legend.hidden_entries_count = 2;
  EXPECT_EQ(fm_sheet_set_chart_legend(wb.handle, 0, id, &legend), kBindingNullPointer);

  // A refused record leaves the legend as it was.
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).legend_position, FM_CHART_LEGEND_POSITION_RIGHT);
}

TEST(FormulonCApiChart, UpdateLegendMergesIntoExistingLegend) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);
  ChartGuard chart;

  const uint32_t hidden[] = {1U};
  fm_chart_legend_patch_t patch = FM_STRUCT_INIT(fm_chart_legend_patch_t);
  patch.position_op = FM_PATCH_OP_SET;
  patch.position = FM_CHART_LEGEND_POSITION_RIGHT;
  patch.hidden_entries_op = FM_PATCH_OP_SET;
  patch.hidden_entries = hidden;
  patch.hidden_entries_count = 1;
  ASSERT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();

  // A KEEP on the hidden entries leaves them in place while the position moves.
  patch = FM_STRUCT_INIT(fm_chart_legend_patch_t);
  patch.position_op = FM_PATCH_OP_SET;
  patch.position = FM_CHART_LEGEND_POSITION_TOP;
  patch.overlay_op = FM_PATCH_OP_SET;
  patch.overlay = 1;
  ASSERT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  const fm_chart_info_t info = Info(chart);
  EXPECT_EQ(info.legend_position, FM_CHART_LEGEND_POSITION_TOP);
  EXPECT_EQ(info.legend_overlay, 1);
  ASSERT_EQ(info.legend_hidden_count, 1U);
  uint32_t entry = 0;
  ASSERT_EQ(fm_chart_get_legend_hidden_at(chart.handle, 0, &entry), 0);
  EXPECT_EQ(entry, 1U);
}

TEST(FormulonCApiChart, UpdatePointMergesIntoExistingPointFormat) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kLine);
  ChartGuard chart;

  fm_chart_point_patch_t color = FM_STRUCT_INIT(fm_chart_point_patch_t);
  color.line.color_op = FM_PATCH_OP_SET;
  color.line.color.kind = FM_COLOR_KIND_RGB;
  color.line.color.argb = 0xFF112233U;
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 1, 2, &color), 0) << fm_get_last_error_message();

  fm_chart_point_patch_t marker = FM_STRUCT_INIT(fm_chart_point_patch_t);
  marker.marker_op = FM_PATCH_OP_SET;
  marker.marker.symbol_op = FM_PATCH_OP_SET;
  marker.marker.symbol = FM_CHART_MARKER_SYMBOL_DIAMOND;
  marker.marker.size_pt_op = FM_PATCH_OP_SET;
  marker.marker.size_pt = 9.0;
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 1, 2, &marker), 0) << fm_get_last_error_message();

  Open(wb.handle, id, &chart);
  ASSERT_EQ(Series(chart, 0, 1).point_count, 1U);
  fm_chart_point_info_t pt = FM_STRUCT_INIT(fm_chart_point_info_t);
  ASSERT_EQ(fm_chart_get_series_point_at(chart.handle, 0, 1, 0, &pt), 0);
  EXPECT_EQ(pt.idx, 2U);
  EXPECT_EQ(pt.line.color.argb, 0xFF112233U);
  ASSERT_EQ(pt.has_marker, 1);
  EXPECT_EQ(pt.marker.symbol, FM_CHART_MARKER_SYMBOL_DIAMOND);
  EXPECT_EQ(pt.marker.has_size_pt, 1);
  EXPECT_EQ(pt.marker.size_pt, 9U);
}

TEST(FormulonCApiChart, AxisVisibleReadsOneAndRoundTripsThroughPatch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  const uint32_t val_id = AxisId(chart, FM_CHART_AXIS_KIND_VALUE);
  EXPECT_EQ(Axis(chart, val_id).visible, 1);

  fm_chart_axis_patch_t patch = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  patch.visible_op = FM_PATCH_OP_SET;
  patch.visible = 0;
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Axis(chart, val_id).visible, 0);

  patch.visible = 1;
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, val_id, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Axis(chart, val_id).visible, 1);
}

TEST(FormulonCApiChart, PlotPercentAndDegreeFieldsRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kPie, "Sheet1!$A$1:$B$5");
  ChartGuard chart;

  fm_chart_plot_patch_t patch = FM_STRUCT_INIT(fm_chart_plot_patch_t);
  patch.first_slice_angle_deg_op = FM_PATCH_OP_SET;
  patch.first_slice_angle_deg = 90;
  ASSERT_EQ(fm_sheet_update_chart_plot(wb.handle, 0, id, 0, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.has_first_slice_angle_deg, 1);
  EXPECT_EQ(plot.first_slice_angle_deg, 90U);
  EXPECT_EQ(plot.extended_layout, FM_CHART_EXTENDED_LAYOUT_ABSENT);

  patch = FM_STRUCT_INIT(fm_chart_plot_patch_t);
  patch.first_slice_angle_deg_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_chart_plot(wb.handle, 0, id, 0, &patch), kInvalidArgument);

  const uint32_t column = Insert(wb.handle, kColumn);
  patch = FM_STRUCT_INIT(fm_chart_plot_patch_t);
  patch.gap_width_pct_op = FM_PATCH_OP_SET;
  patch.gap_width_pct = 120;
  ASSERT_EQ(fm_sheet_update_chart_plot(wb.handle, 0, column, 0, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, column, &chart);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.has_gap_width_pct, 1);
  EXPECT_EQ(plot.gap_width_pct, 120U);
}

TEST(FormulonCApiChart, FillTransparencyPctRoundTripsThroughSeriesPatch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);
  ChartGuard chart;

  fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.fill.color_op = FM_PATCH_OP_SET;
  patch.fill.color.kind = FM_COLOR_KIND_RGB;
  patch.fill.color.argb = 0xFF336699U;
  patch.fill.transparency_pct_op = FM_PATCH_OP_SET;
  patch.fill.transparency_pct = 40.0;
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), 0) << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_DOUBLE_EQ(Series(chart, 0, 0).fill.transparency_pct, 40.0);
}

TEST(FormulonCApiChart, ExtendedChartReadsLayoutAndIsExtended) {
  WorkbookGuard cx;
  LoadFixture("chart_chartex.xlsx", &cx);
  const std::vector<Located> charts = AllCharts(cx.handle);
  ASSERT_FALSE(charts.empty());
  ChartGuard chart;
  Open(cx.handle, charts[0].id, &chart, charts[0].sheet);
  EXPECT_EQ(Info(chart).is_extended, 1);
  fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.type, FM_CHART_PLOT_TYPE_EXTENDED);
  EXPECT_EQ(plot.extended_layout, FM_CHART_EXTENDED_LAYOUT_WATERFALL);
}

TEST(FormulonCApiChart, ExtendedLayoutOrdinalsFollowTheFileFormatOrder) {
  EXPECT_EQ(FM_CHART_EXTENDED_LAYOUT_BOX_WHISKER, 0);
  EXPECT_EQ(FM_CHART_EXTENDED_LAYOUT_CLUSTERED_COLUMN, 1);
  EXPECT_EQ(FM_CHART_EXTENDED_LAYOUT_FUNNEL, 2);
  EXPECT_EQ(FM_CHART_EXTENDED_LAYOUT_PARETO_LINE, 3);
  EXPECT_EQ(FM_CHART_EXTENDED_LAYOUT_REGION_MAP, 4);
  EXPECT_EQ(FM_CHART_EXTENDED_LAYOUT_SUNBURST, 5);
  EXPECT_EQ(FM_CHART_EXTENDED_LAYOUT_TREEMAP, 6);
  EXPECT_EQ(FM_CHART_EXTENDED_LAYOUT_WATERFALL, 7);
  EXPECT_EQ(FM_CHART_ROLE_TITLE, 6);
  EXPECT_EQ(FM_CHART_ROLE_AXIS_TITLE, 7);
}

TEST(FormulonCApiChart, HeaderTristateSelectsHeaderRowAndColumn) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  fm_chart_insert_spec_t o = Options(kColumn, "Sheet1!$A$1:$D$5");
  o.header_row = FM_TRISTATE_TRUE;
  o.header_column = FM_TRISTATE_TRUE;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, &id), 0) << fm_get_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).source_header_row, 1);
  EXPECT_EQ(Info(chart).source_header_column, 1);
  fm_chart_plot_info_t plot = FM_STRUCT_INIT(fm_chart_plot_info_t);
  ASSERT_EQ(fm_chart_get_plot_at(chart.handle, 0, &plot), 0);
  EXPECT_EQ(plot.series_count, 3U);

  // FALSE reads the first row as data, so the series are not named from it.
  ASSERT_EQ(fm_sheet_set_chart_source(wb.handle, 0, id, "Sheet1!$A$1:$D$5", FM_CHART_SERIES_DIRECTION_COLUMNS,
                                      FM_TRISTATE_FALSE, FM_TRISTATE_FALSE),
            0)
      << fm_get_last_error_message();
  Open(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).source_header_row, 0);
  EXPECT_EQ(Info(chart).source_header_column, 0);
  EXPECT_EQ(fm_sheet_set_chart_source(wb.handle, 0, id, "Sheet1!$A$1:$D$5", FM_CHART_SERIES_DIRECTION_COLUMNS, 3,
                                      FM_TRISTATE_DEFAULT),
            kInvalidArgument);
}

TEST(FormulonCApiChart, UnresolvedInputsReportTheirRole) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  size_t second = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Other", append_position(wb.handle), &second), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(second, 1U);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 1, 0, 0, "T"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 1, 0, 1, "S"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 1, 1, 0, "x"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 1, 1, 1, 5.0), 0);
  const uint32_t id = Insert(wb.handle, kColumn, "Other!$A$1:$B$2");
  const fm_chart_title_t title{FM_CHART_TITLE_KIND_REF, nullptr, "Other!$A$1"};
  ASSERT_EQ(fm_sheet_set_chart_title(wb.handle, 0, id, &title), 0) << fm_get_last_error_message();
  ChartGuard chart;
  Open(wb.handle, id, &chart);
  fm_chart_axis_patch_t axis = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  axis.title_op = FM_PATCH_OP_SET;
  axis.title = fm_chart_title_t{FM_CHART_TITLE_KIND_REF, nullptr, "Other!$A$1"};
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, AxisId(chart, FM_CHART_AXIS_KIND_VALUE), &axis), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, 1), 0) << fm_get_last_error_message();

  DataGuard data;
  OpenData(wb.handle, id, &data);
  fm_chart_data_info_t info = FM_STRUCT_INIT(fm_chart_data_info_t);
  ASSERT_EQ(fm_chart_data_get_info(data.handle, &info), 0);
  ASSERT_GT(info.unresolved_count, 0U);
  bool title_seen = false;
  bool axis_title_seen = false;
  bool values_seen = false;
  for (size_t i = 0; i < info.unresolved_count; ++i) {
    fm_chart_unresolved_info_t un = FM_STRUCT_INIT(fm_chart_unresolved_info_t);
    ASSERT_EQ(fm_chart_data_get_unresolved_at(data.handle, i, &un), 0);
    title_seen = title_seen || un.source == FM_CHART_ROLE_TITLE;
    axis_title_seen = axis_title_seen || un.source == FM_CHART_ROLE_AXIS_TITLE;
    values_seen = values_seen || un.source == FM_CHART_ROLE_VALUES;
    EXPECT_GE(un.source, FM_CHART_ROLE_NAME);
    EXPECT_LE(un.source, FM_CHART_ROLE_AXIS_TITLE);
  }
  EXPECT_TRUE(title_seen);
  EXPECT_TRUE(axis_title_seen);
  EXPECT_TRUE(values_seen);
}

TEST(FormulonCApiChart, CategoryLevelCountReportsEntriesPerLevel) {
  int32_t present = 0;
  // Levels count the upper levels of a multi-level category; a plain category has none.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  const uint32_t id = Insert(wb.handle, kColumn);
  DataGuard data;
  OpenData(wb.handle, id, &data);
  fm_chart_data_info_t info = FM_STRUCT_INIT(fm_chart_data_info_t);
  ASSERT_EQ(fm_chart_data_get_info(data.handle, &info), 0);
  EXPECT_EQ(info.category_level_count, 0U);
  size_t n = 0;
  EXPECT_EQ(fm_chart_data_get_category_level_count(data.handle, 0, &n), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_chart_data_get_category_level_count(nullptr, 0, &n), kBindingNullPointer);
  EXPECT_EQ(fm_chart_data_get_category_level_count(data.handle, 0, nullptr), kBindingNullPointer);

  WorkbookGuard multi;
  LoadFixture("chart_data.xlsx", &multi);
  bool found = false;
  for (const Located& c : AllCharts(multi.handle)) {
    DataGuard levels;
    ASSERT_EQ(fm_sheet_get_chart_data(multi.handle, c.sheet, c.id, &present, &levels.handle), 0)
        << fm_get_last_error_message();
    fm_chart_data_info_t mi = FM_STRUCT_INIT(fm_chart_data_info_t);
    ASSERT_EQ(fm_chart_data_get_info(levels.handle, &mi), 0);
    for (size_t level = 0; level < mi.category_level_count; ++level) {
      ASSERT_EQ(fm_chart_data_get_category_level_count(levels.handle, level, &n), 0);
      EXPECT_GT(n, 0U);
      found = true;
    }
    EXPECT_EQ(fm_chart_data_get_category_level_count(levels.handle, mi.category_level_count, &n),
              FM_STATUS_INDEX_OUT_OF_RANGE);
  }
  EXPECT_TRUE(found);
}

}  // namespace
