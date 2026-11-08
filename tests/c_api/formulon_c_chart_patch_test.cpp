// C ABI chart edits through patch records: per-role series inputs (reference
// and literal), the axis / labels / display / legend members, the plot-type
// vocabulary with its bar direction, and the point override rules.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

struct ChartGuard {
  fm_chart_t* handle = nullptr;
  ~ChartGuard() { fm_chart_destroy(handle); }
};

/// Sheet1: B1:D1 names, A2:A5 text categories, B2:D5 numbers.
void FillGrid(fm_workbook_t* wb) {
  const char* names[] = {"S1", "S2", "S3"};
  const char* cats[] = {"a", "b", "c", "d"};
  for (std::uint32_t c = 0; c < 3; ++c) {
    ASSERT_EQ(fm_sheet_set_text(wb, 0, 0, c + 1, names[c]), FM_STATUS_OK);
  }
  for (std::uint32_t r = 0; r < 4; ++r) {
    ASSERT_EQ(fm_sheet_set_text(wb, 0, r + 1, 0, cats[r]), FM_STATUS_OK);
    for (std::uint32_t c = 0; c < 3; ++c) {
      ASSERT_EQ(fm_sheet_set_number(wb, 0, r + 1, c + 1, static_cast<double>((r + 1) * (c + 2))), FM_STATUS_OK);
    }
  }
}

fm_chart_insert_t InsertOptions(int32_t type, int32_t bar_direction) {
  fm_chart_insert_t o = FM_STRUCT_INIT(fm_chart_insert_t);
  o.type = type;
  o.bar_direction = bar_direction;
  o.source = "Sheet1!$A$1:$D$5";
  o.legend = FM_CHART_LEGEND_POSITION_RIGHT;
  o.placement.anchor.anchor_kind = FM_ANCHOR_KIND_TWO_CELL;
  o.placement.anchor.edit_as = FM_ANCHOR_EDIT_AS_TWO_CELL;
  o.placement.anchor.row = 6;
  return o;
}

std::uint32_t InsertChart(fm_workbook_t* wb, int32_t type = FM_CHART_PLOT_TYPE_BAR,
                          int32_t bar_direction = FM_CHART_BAR_DIRECTION_COLUMN) {
  const fm_chart_insert_t o = InsertOptions(type, bar_direction);
  std::uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_chart(wb, 0, &o, &id), FM_STATUS_OK) << fm_get_last_error_message();
  return id;
}

void OpenChart(fm_workbook_t* wb, std::uint32_t id, ChartGuard* out) {
  fm_chart_destroy(out->handle);
  out->handle = nullptr;
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_chart(wb, 0, id, &present, &out->handle), FM_STATUS_OK) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
}

fm_chart_info_t Info(const ChartGuard& chart) {
  fm_chart_info_t info = FM_STRUCT_INIT(fm_chart_info_t);
  EXPECT_EQ(fm_chart_get_info(chart.handle, &info), FM_STATUS_OK);
  return info;
}

fm_chart_series_t Series(const ChartGuard& chart, std::size_t plot, std::size_t series) {
  fm_chart_series_t s = FM_STRUCT_INIT(fm_chart_series_t);
  EXPECT_EQ(fm_chart_get_series_at(chart.handle, plot, series, &s), FM_STATUS_OK) << fm_get_last_error_message();
  return s;
}

fm_chart_plot_t Plot(const ChartGuard& chart, std::size_t plot) {
  fm_chart_plot_t p = FM_STRUCT_INIT(fm_chart_plot_t);
  EXPECT_EQ(fm_chart_get_plot_at(chart.handle, plot, &p), FM_STATUS_OK);
  return p;
}

std::uint32_t AxisId(const ChartGuard& chart, std::int32_t kind) {
  const fm_chart_info_t info = Info(chart);
  for (std::size_t i = 0; i < info.axis_count; ++i) {
    fm_chart_axis_t a = FM_STRUCT_INIT(fm_chart_axis_t);
    EXPECT_EQ(fm_chart_get_axis_at(chart.handle, i, &a), FM_STATUS_OK);
    if (a.kind == kind) {
      return a.id;
    }
  }
  ADD_FAILURE() << "no axis of kind " << kind;
  return 0;
}

fm_chart_axis_t Axis(const ChartGuard& chart, std::uint32_t id) {
  const fm_chart_info_t info = Info(chart);
  for (std::size_t i = 0; i < info.axis_count; ++i) {
    fm_chart_axis_t a = FM_STRUCT_INIT(fm_chart_axis_t);
    EXPECT_EQ(fm_chart_get_axis_at(chart.handle, i, &a), FM_STATUS_OK);
    if (a.id == id) {
      return a;
    }
  }
  ADD_FAILURE() << "no axis with id " << id;
  return FM_STRUCT_INIT(fm_chart_axis_t);
}

void SaveAndReload(const WorkbookGuard& source, WorkbookGuard* out) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(source.handle, &saved.data, &saved.len), FM_STATUS_OK) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &out->handle), FM_STATUS_OK) << fm_get_last_error_message();
}

fm_chart_data_source_spec_t Ref(const char* formula) {
  return fm_chart_data_source_spec_t{FM_CHART_SOURCE_KIND_REF, formula, nullptr, 0};
}

fm_chart_data_source_spec_t Literal(const std::vector<fm_chart_literal_t>& points) {
  return fm_chart_data_source_spec_t{FM_CHART_SOURCE_KIND_LITERAL, nullptr, points.data(),
                                     static_cast<uint32_t>(points.size())};
}

fm_chart_literal_t NumberPoint(std::uint32_t idx, double number) {
  fm_chart_literal_t p = FM_STRUCT_INIT(fm_chart_literal_t);
  p.idx = idx;
  p.is_text = 0;
  p.number = number;
  return p;
}

fm_chart_literal_t TextPoint(std::uint32_t idx, const char* text) {
  fm_chart_literal_t p = FM_STRUCT_INIT(fm_chart_literal_t);
  p.idx = idx;
  p.is_text = 1;
  p.text = text;
  return p;
}

// ---------------------------------------------------------------------------
// Series inputs
// ---------------------------------------------------------------------------

TEST(FormulonCApiChartPatch, SeriesPatchWritesLiteralNameAndValuesAndTheyRoundTrip) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertChart(wb.handle);

  const std::vector<fm_chart_literal_t> name = {TextPoint(0, "Total")};
  const std::vector<fm_chart_literal_t> values = {NumberPoint(0, 1.5), NumberPoint(1, 2.5), NumberPoint(2, 4.0)};
  fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.name_op = FM_CHART_PATCH_OP_SET;
  patch.name = Literal(name);
  patch.values_op = FM_CHART_PATCH_OP_SET;
  patch.values = Literal(values);
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_OK) << fm_get_last_error_message();

  const auto check = [&](const ChartGuard& chart) {
    const fm_chart_series_t s = Series(chart, 0, 0);
    EXPECT_EQ(s.name.kind, FM_CHART_SOURCE_KIND_LITERAL);
    ASSERT_EQ(s.name.literal_count, 1U);
    fm_chart_literal_t point = FM_STRUCT_INIT(fm_chart_literal_t);
    ASSERT_EQ(fm_chart_get_series_literal_at(chart.handle, 0, 0, FM_CHART_ROLE_NAME, 0, &point), FM_STATUS_OK);
    EXPECT_EQ(point.is_text, 1);
    EXPECT_STREQ(point.text, "Total");

    EXPECT_EQ(s.values.kind, FM_CHART_SOURCE_KIND_LITERAL);
    ASSERT_EQ(s.values.literal_count, 3U);
    const double expected[] = {1.5, 2.5, 4.0};
    for (std::size_t i = 0; i < 3U; ++i) {
      ASSERT_EQ(fm_chart_get_series_literal_at(chart.handle, 0, 0, FM_CHART_ROLE_VALUES, i, &point), FM_STATUS_OK);
      EXPECT_EQ(point.is_text, 0) << i;
      EXPECT_EQ(point.number, expected[i]) << i;
    }
    // The other series keep their references.
    EXPECT_EQ(Series(chart, 0, 1).values.kind, FM_CHART_SOURCE_KIND_REF);
  };

  ChartGuard chart;
  OpenChart(wb.handle, id, &chart);
  check(chart);

  WorkbookGuard reloaded;
  SaveAndReload(wb, &reloaded);
  ChartGuard reloaded_chart;
  OpenChart(reloaded.handle, id, &reloaded_chart);
  check(reloaded_chart);
}

TEST(FormulonCApiChartPatch, SeriesPatchOpsKeepSetAndRemovePerRole) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertChart(wb.handle);
  ChartGuard chart;
  OpenChart(wb.handle, id, &chart);
  ASSERT_EQ(Series(chart, 0, 0).name.kind, FM_CHART_SOURCE_KIND_REF);
  ASSERT_EQ(Series(chart, 0, 0).categories.kind, FM_CHART_SOURCE_KIND_REF);

  // KEEP leaves every role alone whatever the spec beside it holds.
  fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.values = Ref("Nope!$A$1:$A$2");
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_OK) << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_STREQ(Series(chart, 0, 0).values.formula, "Sheet1!$B$2:$B$5");

  // CLEAR removes name; SET with NONE removes categories.
  patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.name_op = FM_CHART_PATCH_OP_CLEAR;
  patch.categories_op = FM_CHART_PATCH_OP_SET;
  patch.categories = fm_chart_data_source_spec_t{FM_CHART_SOURCE_KIND_NONE, nullptr, nullptr, 0};
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_OK) << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Series(chart, 0, 0).name.kind, FM_CHART_SOURCE_KIND_NONE);
  EXPECT_EQ(Series(chart, 0, 0).categories.kind, FM_CHART_SOURCE_KIND_NONE);
  EXPECT_EQ(Series(chart, 0, 0).values.kind, FM_CHART_SOURCE_KIND_REF);

  // A required role cannot be removed; nothing changes.
  for (const std::int32_t op : {FM_CHART_PATCH_OP_CLEAR, FM_CHART_PATCH_OP_SET}) {
    patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
    patch.values_op = op;
    patch.values = fm_chart_data_source_spec_t{FM_CHART_SOURCE_KIND_NONE, nullptr, nullptr, 0};
    EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_INVALID_ARGUMENT) << op;
  }
  OpenChart(wb.handle, id, &chart);
  EXPECT_STREQ(Series(chart, 0, 0).values.formula, "Sheet1!$B$2:$B$5");

  // A reference input needs a formula; a literal input needs points.
  patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.values_op = FM_CHART_PATCH_OP_SET;
  patch.values = Ref(nullptr);
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_INVALID_ARGUMENT);
  patch.values = Ref("");
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_INVALID_ARGUMENT);
  patch.values = fm_chart_data_source_spec_t{FM_CHART_SOURCE_KIND_LITERAL, nullptr, nullptr, 0};
  EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_INVALID_ARGUMENT);
  OpenChart(wb.handle, id, &chart);
  EXPECT_STREQ(Series(chart, 0, 0).values.formula, "Sheet1!$B$2:$B$5");
}

TEST(FormulonCApiChartPatch, AddedSeriesTakesALiteralValuesInput) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertChart(wb.handle);
  ChartGuard chart;
  OpenChart(wb.handle, id, &chart);
  const std::uint32_t before = Plot(chart, 0).series_count;

  const std::vector<fm_chart_literal_t> values = {NumberPoint(0, 7.0), NumberPoint(1, 8.0)};
  fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.values_op = FM_CHART_PATCH_OP_SET;
  patch.values = Literal(values);
  std::size_t added = 0;
  ASSERT_EQ(fm_sheet_add_chart_series(wb.handle, 0, id, 0, &patch, &added), FM_STATUS_OK)
      << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Plot(chart, 0).series_count, before + 1U);
  const fm_chart_series_t s = Series(chart, 0, before);
  EXPECT_EQ(s.values.kind, FM_CHART_SOURCE_KIND_LITERAL);
  EXPECT_EQ(s.values.literal_count, 2U);

  // A series without values is refused and adds nothing.
  fm_chart_series_patch_t empty = FM_STRUCT_INIT(fm_chart_series_patch_t);
  EXPECT_NE(fm_sheet_add_chart_series(wb.handle, 0, id, 0, &empty, &added), FM_STATUS_OK);
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Plot(chart, 0).series_count, before + 1U);
}

// ---------------------------------------------------------------------------
// Axis, labels, display, legend
// ---------------------------------------------------------------------------

TEST(FormulonCApiChartPatch, AxisPatchSetsMinorUnitGridlinesPositionAndSourceLinking) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertChart(wb.handle);
  ChartGuard chart;
  OpenChart(wb.handle, id, &chart);
  const std::uint32_t value_axis = AxisId(chart, FM_CHART_AXIS_KIND_VALUE);

  fm_chart_axis_patch_t patch = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  patch.minor_unit_op = FM_CHART_PATCH_OP_SET;
  patch.minor_unit = 2.5;
  patch.minor_gridlines_op = FM_CHART_PATCH_OP_SET;
  patch.minor_gridlines = 1;
  patch.position_op = FM_CHART_PATCH_OP_SET;
  patch.position = FM_CHART_AXIS_POSITION_RIGHT;
  patch.source_linked_op = FM_CHART_PATCH_OP_SET;
  patch.source_linked = 1;
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, value_axis, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  fm_chart_axis_t a = Axis(chart, value_axis);
  EXPECT_EQ(a.minor_unit, 2.5);
  EXPECT_EQ(a.minor_gridlines, 1);
  EXPECT_EQ(a.position, FM_CHART_AXIS_POSITION_RIGHT);
  EXPECT_EQ(a.source_linked, 1);

  // KEEP leaves them; CLEAR on the unit makes it automatic.
  patch = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  patch.minor_unit_op = FM_CHART_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, value_axis, &patch), FM_STATUS_OK)
      << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  a = Axis(chart, value_axis);
  EXPECT_TRUE(std::isnan(a.minor_unit));
  EXPECT_EQ(a.minor_gridlines, 1);
  EXPECT_EQ(a.position, FM_CHART_AXIS_POSITION_RIGHT);
  EXPECT_EQ(a.source_linked, 1);

  // The three that always have a value reject CLEAR and leave the axis as it was.
  for (int member = 0; member < 3; ++member) {
    patch = FM_STRUCT_INIT(fm_chart_axis_patch_t);
    patch.minor_unit_op = FM_CHART_PATCH_OP_SET;
    patch.minor_unit = 9.0;
    (member == 0   ? patch.minor_gridlines_op
     : member == 1 ? patch.position_op
                   : patch.source_linked_op) = FM_CHART_PATCH_OP_CLEAR;
    EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, value_axis, &patch), FM_STATUS_INVALID_ARGUMENT) << member;
  }
  OpenChart(wb.handle, id, &chart);
  EXPECT_TRUE(std::isnan(Axis(chart, value_axis).minor_unit));

  // A value outside the enumeration is refused as well.
  patch = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  patch.position_op = FM_CHART_PATCH_OP_SET;
  patch.position = FM_CHART_AXIS_POSITION_RIGHT + 1;
  EXPECT_EQ(fm_sheet_update_chart_axis(wb.handle, 0, id, value_axis, &patch), FM_STATUS_INVALID_ARGUMENT);
}

TEST(FormulonCApiChartPatch, LabelsPatchSetsSourceLinkedAndDeleted) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertChart(wb.handle);

  fm_chart_series_patch_t patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.labels_op = FM_CHART_PATCH_OP_SET;
  patch.labels.show_value_op = FM_CHART_PATCH_OP_SET;
  patch.labels.show_value = 1;
  patch.labels.source_linked_op = FM_CHART_PATCH_OP_SET;
  patch.labels.source_linked = 0;
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_OK) << fm_get_last_error_message();
  ChartGuard chart;
  OpenChart(wb.handle, id, &chart);
  fm_chart_series_t s = Series(chart, 0, 0);
  ASSERT_EQ(s.has_labels, 1);
  EXPECT_EQ(s.labels.source_linked, 0);
  EXPECT_EQ(s.labels.deleted, 0);

  patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.labels_op = FM_CHART_PATCH_OP_SET;
  patch.labels.source_linked_op = FM_CHART_PATCH_OP_SET;
  patch.labels.source_linked = 1;
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_OK) << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  s = Series(chart, 0, 0);
  EXPECT_EQ(s.labels.source_linked, 1);
  EXPECT_EQ(s.labels.deleted, 0);

  // Deleted labels keep nothing beside `c:delete`, the number format included.
  patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
  patch.labels_op = FM_CHART_PATCH_OP_SET;
  patch.labels.deleted_op = FM_CHART_PATCH_OP_SET;
  patch.labels.deleted = 1;
  ASSERT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_OK) << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  s = Series(chart, 0, 0);
  EXPECT_EQ(s.labels.source_linked, 0);
  EXPECT_EQ(s.labels.deleted, 1);

  // CLEAR is refused on both flags and leaves the labels as they were.
  for (int member = 0; member < 2; ++member) {
    patch = FM_STRUCT_INIT(fm_chart_series_patch_t);
    patch.labels_op = FM_CHART_PATCH_OP_SET;
    (member == 0 ? patch.labels.source_linked_op : patch.labels.deleted_op) = FM_CHART_PATCH_OP_CLEAR;
    EXPECT_EQ(fm_sheet_update_chart_series(wb.handle, 0, id, 0, 0, &patch), FM_STATUS_INVALID_ARGUMENT) << member;
  }
  OpenChart(wb.handle, id, &chart);
  s = Series(chart, 0, 0);
  EXPECT_EQ(s.labels.deleted, 1);
}

TEST(FormulonCApiChartPatch, DisplayPatchSetsAndRemovesTheChartStyle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertChart(wb.handle);

  fm_chart_display_patch_t patch = FM_STRUCT_INIT(fm_chart_display_patch_t);
  patch.style_op = FM_CHART_PATCH_OP_SET;
  patch.style = 26;
  ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &patch), FM_STATUS_OK) << fm_get_last_error_message();
  ChartGuard chart;
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).has_style, 1);
  EXPECT_EQ(Info(chart).style, 26);

  WorkbookGuard reloaded;
  SaveAndReload(wb, &reloaded);
  ChartGuard reloaded_chart;
  OpenChart(reloaded.handle, id, &reloaded_chart);
  EXPECT_EQ(Info(reloaded_chart).style, 26);

  // ST_Style is 1..48.
  for (const std::int32_t bad : {0, 49}) {
    patch = FM_STRUCT_INIT(fm_chart_display_patch_t);
    patch.style_op = FM_CHART_PATCH_OP_SET;
    patch.style = bad;
    EXPECT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &patch), FM_STATUS_INVALID_ARGUMENT) << bad;
  }
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).style, 26);

  patch = FM_STRUCT_INIT(fm_chart_display_patch_t);
  patch.style_op = FM_CHART_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_chart_display(wb.handle, 0, id, &patch), FM_STATUS_OK) << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).has_style, 0);
}

TEST(FormulonCApiChartPatch, LegendPatchHidesEntriesIndependentlyOfTheChartFilter) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertChart(wb.handle);

  const std::uint32_t entries[] = {1U, 2U};
  fm_chart_legend_patch_t patch = FM_STRUCT_INIT(fm_chart_legend_patch_t);
  patch.position = FM_CHART_LEGEND_POSITION_TOP;
  patch.hidden_entries_op = FM_CHART_PATCH_OP_SET;
  patch.hidden_entries = entries;
  patch.hidden_entry_count = 2;
  ASSERT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, id, &patch), FM_STATUS_OK) << fm_get_last_error_message();
  ChartGuard chart;
  OpenChart(wb.handle, id, &chart);
  ASSERT_EQ(Info(chart).legend_hidden_count, 2U);
  std::uint32_t entry = 0;
  ASSERT_EQ(fm_chart_get_legend_hidden_at(chart.handle, 0, &entry), FM_STATUS_OK);
  EXPECT_EQ(entry, 1U);
  ASSERT_EQ(fm_chart_get_legend_hidden_at(chart.handle, 1, &entry), FM_STATUS_OK);
  EXPECT_EQ(entry, 2U);
  EXPECT_EQ(Info(chart).hidden_category_count, 0U);

  // The filter hides a series and a category without touching the legend entries.
  const std::uint32_t hidden_series[] = {0U};
  const std::uint32_t hidden_categories[] = {3U};
  ASSERT_EQ(fm_sheet_set_chart_filter(wb.handle, 0, id, hidden_series, 1, hidden_categories, 1), FM_STATUS_OK)
      << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).legend_hidden_count, 2U);
  EXPECT_EQ(Info(chart).hidden_category_count, 1U);

  // KEEP leaves the entries; a non-zero count needs an array.
  patch = FM_STRUCT_INIT(fm_chart_legend_patch_t);
  patch.position = FM_CHART_LEGEND_POSITION_LEFT;
  ASSERT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, id, &patch), FM_STATUS_OK) << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Info(chart).legend_position, FM_CHART_LEGEND_POSITION_LEFT);
  EXPECT_EQ(Info(chart).legend_hidden_count, 2U);
  patch.hidden_entries_op = FM_CHART_PATCH_OP_SET;
  patch.hidden_entries = nullptr;
  patch.hidden_entry_count = 1;
  EXPECT_EQ(fm_sheet_update_chart_legend(wb.handle, 0, id, &patch), FM_STATUS_BINDING_NULL_POINTER);
}

// ---------------------------------------------------------------------------
// Plot type and bar direction
// ---------------------------------------------------------------------------

TEST(FormulonCApiChartPatch, InsertAndSetTypeCarryTheBarDirection) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertChart(wb.handle, FM_CHART_PLOT_TYPE_BAR, FM_CHART_BAR_DIRECTION_BAR);
  ChartGuard chart;
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Plot(chart, 0).type, FM_CHART_PLOT_TYPE_BAR);
  EXPECT_EQ(Plot(chart, 0).bar_direction, FM_CHART_BAR_DIRECTION_BAR);

  // Back to columns, then to a line plot: zero is accepted for a type without a direction.
  ASSERT_EQ(fm_sheet_set_chart_type(wb.handle, 0, id, 0, FM_CHART_PLOT_TYPE_BAR, FM_CHART_BAR_DIRECTION_COLUMN,
                                    FM_CHART_GROUPING_STACKED),
            FM_STATUS_OK)
      << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Plot(chart, 0).type, FM_CHART_PLOT_TYPE_BAR);
  EXPECT_EQ(Plot(chart, 0).bar_direction, FM_CHART_BAR_DIRECTION_COLUMN);
  EXPECT_EQ(Plot(chart, 0).grouping, FM_CHART_GROUPING_STACKED);

  ASSERT_EQ(fm_sheet_set_chart_type(wb.handle, 0, id, 0, FM_CHART_PLOT_TYPE_LINE, 0, FM_CHART_GROUPING_STANDARD),
            FM_STATUS_OK)
      << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Plot(chart, 0).type, FM_CHART_PLOT_TYPE_LINE);
  EXPECT_EQ(Plot(chart, 0).bar_direction, FM_CHART_BAR_DIRECTION_ABSENT);

  // A bar direction of BAR on any other type is refused and changes nothing.
  EXPECT_EQ(fm_sheet_set_chart_type(wb.handle, 0, id, 0, FM_CHART_PLOT_TYPE_AREA, FM_CHART_BAR_DIRECTION_BAR,
                                    FM_CHART_GROUPING_STANDARD),
            FM_STATUS_INVALID_ARGUMENT);
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Plot(chart, 0).type, FM_CHART_PLOT_TYPE_LINE);

  // A round trip of the read plot through the setter changes nothing.
  const fm_chart_plot_t read = Plot(chart, 0);
  ASSERT_EQ(fm_sheet_set_chart_type(wb.handle, 0, id, 0, read.type,
                                    read.bar_direction == FM_CHART_BAR_DIRECTION_BAR ? 1 : 0, read.grouping),
            FM_STATUS_OK)
      << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Plot(chart, 0).type, read.type);
  EXPECT_EQ(Plot(chart, 0).grouping, read.grouping);
  EXPECT_EQ(Plot(chart, 0).series_count, read.series_count);
}

TEST(FormulonCApiChartPatch, InsertRefusesABarDirectionOnAnotherTypeAndCreatesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const fm_chart_insert_t o = InsertOptions(FM_CHART_PLOT_TYPE_PIE, FM_CHART_BAR_DIRECTION_BAR);
  std::uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &o, &id), FM_STATUS_INVALID_ARGUMENT);
  const fm_chart_insert_t zero = InsertOptions(FM_CHART_PLOT_TYPE_PIE, 0);
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &zero, &id), FM_STATUS_OK) << fm_get_last_error_message();
  const fm_chart_insert_t stock = InsertOptions(FM_CHART_PLOT_TYPE_STOCK, 0);
  EXPECT_EQ(fm_sheet_add_chart(wb.handle, 0, &stock, &id), FM_STATUS_INVALID_ARGUMENT);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), FM_STATUS_OK);
  EXPECT_EQ(count, 1U);
}

// ---------------------------------------------------------------------------
// Point override
// ---------------------------------------------------------------------------

fm_chart_point_patch_t SetPointColor() {
  fm_chart_point_patch_t point = FM_STRUCT_INIT(fm_chart_point_patch_t);
  point.line.color_op = FM_CHART_PATCH_OP_SET;
  point.line.color.kind = FM_COLOR_KIND_RGB;
  point.line.color.argb = 0xFF123456U;
  point.marker_op = FM_CHART_PATCH_OP_SET;
  point.marker.symbol_op = FM_CHART_PATCH_OP_SET;
  point.marker.symbol = FM_CHART_MARKER_SYMBOL_TRIANGLE;
  return point;
}

fm_chart_point_patch_t ClearEveryMember() {
  fm_chart_point_patch_t point = FM_STRUCT_INIT(fm_chart_point_patch_t);
  point.fill.color_op = FM_CHART_PATCH_OP_CLEAR;
  point.fill.transparency_pct_op = FM_CHART_PATCH_OP_CLEAR;
  point.line.color_op = FM_CHART_PATCH_OP_CLEAR;
  point.line.transparency_pct_op = FM_CHART_PATCH_OP_CLEAR;
  point.line.width_pt_op = FM_CHART_PATCH_OP_CLEAR;
  point.line.dash_op = FM_CHART_PATCH_OP_CLEAR;
  point.line.none_op = FM_CHART_PATCH_OP_CLEAR;
  point.marker_op = FM_CHART_PATCH_OP_CLEAR;
  return point;
}

TEST(FormulonCApiChartPatch, AllKeepPointPatchLeavesTheOverrideAndClearingEveryMemberRemovesIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertChart(wb.handle, FM_CHART_PLOT_TYPE_LINE);

  const fm_chart_point_patch_t set = SetPointColor();
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 1, 2, &set), FM_STATUS_OK) << fm_get_last_error_message();
  ChartGuard chart;
  OpenChart(wb.handle, id, &chart);
  ASSERT_EQ(Series(chart, 0, 1).point_count, 1U);

  const fm_chart_point_patch_t keep = FM_STRUCT_INIT(fm_chart_point_patch_t);
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 1, 2, &keep), FM_STATUS_OK) << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  ASSERT_EQ(Series(chart, 0, 1).point_count, 1U);
  fm_chart_point_t point = FM_STRUCT_INIT(fm_chart_point_t);
  ASSERT_EQ(fm_chart_get_series_point_at(chart.handle, 0, 1, 0, &point), FM_STATUS_OK);
  EXPECT_EQ(point.idx, 2U);
  EXPECT_EQ(point.line.color.argb, 0xFF123456U);
  EXPECT_EQ(point.has_marker, 1);
  EXPECT_EQ(point.marker.symbol, FM_CHART_MARKER_SYMBOL_TRIANGLE);

  const fm_chart_point_patch_t clear = ClearEveryMember();
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 1, 2, &clear), FM_STATUS_OK)
      << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Series(chart, 0, 1).point_count, 0U);

  // Removing the override of a point that has none does nothing.
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 1, 3, &clear), FM_STATUS_OK)
      << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Series(chart, 0, 1).point_count, 0U);
  // So does an all-KEEP patch on it.
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 1, 3, &keep), FM_STATUS_OK) << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Series(chart, 0, 1).point_count, 0U);
}

TEST(FormulonCApiChartPatch, PointOverrideEmptiedByClearsDisappearsOnlyWhenNothingRemains) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  FillGrid(wb.handle);
  const std::uint32_t id = InsertChart(wb.handle, FM_CHART_PLOT_TYPE_LINE);
  const fm_chart_point_patch_t set = SetPointColor();
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 0, 1, &set), FM_STATUS_OK) << fm_get_last_error_message();

  // Clearing the marker alone leaves the line colour, so the override stays.
  fm_chart_point_patch_t partial = FM_STRUCT_INIT(fm_chart_point_patch_t);
  partial.marker_op = FM_CHART_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 0, 1, &partial), FM_STATUS_OK)
      << fm_get_last_error_message();
  ChartGuard chart;
  OpenChart(wb.handle, id, &chart);
  ASSERT_EQ(Series(chart, 0, 0).point_count, 1U);
  fm_chart_point_t point = FM_STRUCT_INIT(fm_chart_point_t);
  ASSERT_EQ(fm_chart_get_series_point_at(chart.handle, 0, 0, 0, &point), FM_STATUS_OK);
  EXPECT_EQ(point.has_marker, 0);
  EXPECT_EQ(point.line.color.argb, 0xFF123456U);

  // A new point repeats the series line, so its width stays after the colour goes.
  fm_chart_point_patch_t colour = FM_STRUCT_INIT(fm_chart_point_patch_t);
  colour.line.color_op = FM_CHART_PATCH_OP_CLEAR;
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 0, 1, &colour), FM_STATUS_OK)
      << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  ASSERT_EQ(Series(chart, 0, 0).point_count, 1U);

  // Clearing the remaining fill and line members empties the override.
  fm_chart_point_patch_t rest = ClearEveryMember();
  rest.marker_op = FM_CHART_PATCH_OP_KEEP;
  ASSERT_EQ(fm_sheet_update_chart_point(wb.handle, 0, id, 0, 0, 1, &rest), FM_STATUS_OK) << fm_get_last_error_message();
  OpenChart(wb.handle, id, &chart);
  EXPECT_EQ(Series(chart, 0, 0).point_count, 0U);
}

}  // namespace
