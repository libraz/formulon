// A remove addressed by a natural key (name, id or cell) that matches nothing
// is FM_STATUS_NOT_FOUND, and leaves the workbook as it was.

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_setters.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

constexpr const char* kAbsentId = "{00000000-0000-0000-0000-000000000000}";

class FormulonCApiRemoveMissingKey : public ::testing::Test {
 protected:
  void SetUp() override { ASSERT_EQ(fm_workbook_create_default(&wb_.handle), 0); }
  fm_workbook_t* wb() { return wb_.handle; }

  void ExpectNotFound(fm_status_t rc, const char* what) {
    EXPECT_EQ(rc, FM_STATUS_NOT_FOUND) << what;
    EXPECT_STRNE(fm_get_last_error_message(), "") << what;
  }

 private:
  WorkbookGuard wb_;
};

}  // namespace

TEST_F(FormulonCApiRemoveMissingKey, WorkbookScopedNames) {
  ExpectNotFound(fm_workbook_remove_defined_name(wb(), "nope", -1), "defined name");
  ExpectNotFound(fm_workbook_remove_person(wb(), kAbsentId), "person");
  ExpectNotFound(fm_workbook_remove_pivot_cache(wb(), 9999U), "pivot cache");
  ExpectNotFound(fm_workbook_remove_table_style(wb(), "nope"), "table style");
  ExpectNotFound(fm_workbook_remove_cell_style(wb(), "nope"), "cell style");
}

TEST_F(FormulonCApiRemoveMissingKey, SheetScopedKeys) {
  ExpectNotFound(fm_sheet_remove_hyperlink(wb(), 0, 3, 3), "hyperlink");
  const fm_cell_rect_t rect{10, 10, 12, 12};
  ExpectNotFound(fm_sheet_remove_merge(wb(), 0, &rect), "merge");
  ExpectNotFound(fm_sheet_remove_threaded_comment(wb(), 0, kAbsentId), "threaded comment");
  ExpectNotFound(fm_sheet_remove_sparkline_group(wb(), 0, "nope"), "sparkline group");
  ExpectNotFound(fm_sheet_remove_data_table(wb(), 0, 0, 0), "data table");
  ExpectNotFound(fm_sheet_remove_drawing(wb(), 0, 42U), "drawing");
  ExpectNotFound(fm_sheet_remove_cf_rule(wb(), 0, kAbsentId), "cf rule");
  ExpectNotFound(fm_sheet_remove_protected_range(wb(), 0, "nope"), "protected range");
}

// The chart readers are keyed reads: an absent id succeeds, reports the miss
// through out_present and hands out no snapshot.
TEST_F(FormulonCApiRemoveMissingKey, ChartReadersOfAnAbsentIdReportTheMissThroughOutPresent) {
  fm_chart_data_t* data = nullptr;
  std::int32_t present = 99;
  EXPECT_EQ(fm_sheet_get_chart_data(wb(), 0, 42U, &present, &data), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 0);
  EXPECT_EQ(data, nullptr);
  const fm_chart_preview_overrides_spec_t overrides = FM_STRUCT_INIT(fm_chart_preview_overrides_spec_t);
  present = 99;
  EXPECT_EQ(fm_sheet_get_chart_preview(wb(), 0, 42U, &overrides, &present, &data), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 0);
  EXPECT_EQ(data, nullptr);
}

// A mutation keyed by an id that matches nothing is FM_STATUS_NOT_FOUND too.
TEST_F(FormulonCApiRemoveMissingKey, UpdatingAnAbsentThreadedCommentIsNotFound) {
  fm_threaded_comment_patch_t patch = FM_STRUCT_INIT(fm_threaded_comment_patch_t);
  patch.done_op = FM_PATCH_OP_SET;
  patch.done = 1;
  ExpectNotFound(fm_sheet_update_threaded_comment(wb(), 0, kAbsentId, &patch), "threaded comment update");
}

TEST_F(FormulonCApiRemoveMissingKey, UpdatingAnAbsentCfRuleOrProtectedRangeIsNotFound) {
  fm_cf_rule_patch_t cf = FM_STRUCT_INIT(fm_cf_rule_patch_t);
  cf.stop_if_true_op = FM_PATCH_OP_SET;
  cf.stop_if_true = 1;
  ExpectNotFound(fm_sheet_update_cf_rule(wb(), 0, kAbsentId, &cf), "cf rule update");
  fm_protected_range_patch_t range = FM_STRUCT_INIT(fm_protected_range_patch_t);
  range.spin_count_op = FM_PATCH_OP_SET;
  range.spin_count = 1U;
  ExpectNotFound(fm_sheet_update_protected_range(wb(), 0, "nope", &range), "protected range update");
}

// Every edit of a drawing object or chart addressed by an id no object holds
// is FM_STATUS_NOT_FOUND.
TEST_F(FormulonCApiRemoveMissingKey, EditingAnAbsentDrawingOrChartIdIsNotFound) {
  constexpr std::uint32_t kId = 42U;
  fm_drawing_patch_t drawing = FM_STRUCT_INIT(fm_drawing_patch_t);
  drawing.anchor.from_row_op = FM_PATCH_OP_SET;
  drawing.anchor.from_row = 1;
  const fm_sparkline_group_patch_t sparkline_group = FM_STRUCT_INIT(fm_sparkline_group_patch_t);
  const fm_shape_patch_t shape = FM_STRUCT_INIT(fm_shape_patch_t);
  const fm_slicer_patch_t slicer = FM_STRUCT_INIT(fm_slicer_patch_t);
  const fm_chart_title_t title{FM_CHART_TITLE_KIND_TEXT, "t", nullptr};
  const fm_chart_legend_patch_t legend = FM_STRUCT_INIT(fm_chart_legend_patch_t);
  const fm_chart_axis_patch_t axis = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  const fm_chart_display_patch_t display = FM_STRUCT_INIT(fm_chart_display_patch_t);
  const fm_chart_series_patch_t series = FM_STRUCT_INIT(fm_chart_series_patch_t);
  const fm_chart_point_patch_t point = FM_STRUCT_INIT(fm_chart_point_patch_t);
  const fm_chart_plot_patch_t plot = FM_STRUCT_INIT(fm_chart_plot_patch_t);
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  std::uint32_t new_series = 0;
  const struct {
    const char* what;
    fm_status_t rc;
  } cases[] = {
      {"update_drawing", fm_sheet_update_drawing(wb(), 0, kId, &drawing)},
      {"set_drawing_z_order", fm_sheet_set_drawing_z_order(wb(), 0, kId, 0)},
      {"snapshot_drawing", fm_sheet_snapshot_drawing(wb(), 0, kId, &bytes, &len)},
      {"update_shape", fm_sheet_update_shape(wb(), 0, kId, &shape)},
      {"update_slicer", fm_sheet_update_slicer(wb(), 0, kId, &slicer)},
      {"update_sparkline_group", fm_sheet_update_sparkline_group(wb(), 0, "nope", &sparkline_group)},
      {"set_chart_type", fm_sheet_set_chart_type(wb(), 0, kId, 0, FM_CHART_PLOT_TYPE_BAR, FM_CHART_BAR_DIRECTION_COLUMN,
                                                 FM_CHART_GROUPING_CLUSTERED)},
      {"set_chart_title", fm_sheet_set_chart_title(wb(), 0, kId, &title)},
      {"update_chart_legend", fm_sheet_update_chart_legend(wb(), 0, kId, &legend)},
      {"update_chart_axis", fm_sheet_update_chart_axis(wb(), 0, kId, 1U, &axis)},
      {"update_chart_display", fm_sheet_update_chart_display(wb(), 0, kId, &display)},
      {"update_chart_series", fm_sheet_update_chart_series(wb(), 0, kId, 0, 0U, &series)},
      {"update_chart_point", fm_sheet_update_chart_point(wb(), 0, kId, 0, 0U, 0, &point)},
      {"update_chart_plot", fm_sheet_update_chart_plot(wb(), 0, kId, 0, &plot)},
      {"set_chart_filter", fm_sheet_set_chart_filter(wb(), 0, kId, nullptr, 0, nullptr, 0)},
      {"add_chart_series", fm_sheet_add_chart_series(wb(), 0, kId, 0, &series, &new_series)},
      {"remove_chart_series", fm_sheet_remove_chart_series(wb(), 0, kId, 0, 0U)},
      {"set_chart_source", fm_sheet_set_chart_source(wb(), 0, kId, "A1:B2", FM_CHART_SERIES_DIRECTION_AUTO,
                                                     FM_TRISTATE_DEFAULT, FM_TRISTATE_DEFAULT)},
  };
  for (const auto& c : cases) {
    EXPECT_EQ(c.rc, FM_STATUS_NOT_FOUND) << c.what;
  }
  EXPECT_EQ(bytes, nullptr);
  EXPECT_EQ(len, 0U);
}

// A keyed read that matches nothing succeeds and reports the absence instead.
TEST_F(FormulonCApiRemoveMissingKey, KeyedReadsReportAMissThroughOutPresent) {
  std::int32_t present = 99;
  std::int32_t visible = 99;
  EXPECT_EQ(fm_sheet_get_comment_visible(wb(), 0, 3, 3, &present, &visible), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(visible, 0);

  present = 99;
  fm_civil_time_t now{1, 1, 1, 1, 1, 1};
  EXPECT_EQ(fm_workbook_get_pinned_now(wb(), &present, &now), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(now.year, 0);

  present = 99;
  const char* file_name = "stale";
  EXPECT_EQ(fm_workbook_get_file_name(wb(), &present, &file_name), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(file_name, nullptr);

  present = 99;
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  EXPECT_EQ(fm_sheet_get_cf_rule(wb(), 0, kAbsentId, &present, &rule), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(rule.id, nullptr);

  present = 99;
  fm_protected_range_t range = FM_STRUCT_INIT(fm_protected_range_t);
  EXPECT_EQ(fm_sheet_get_protected_range(wb(), 0, "nope", &present, &range), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(range.name, nullptr);

  present = 99;
  fm_auto_filter_t filter = FM_STRUCT_INIT(fm_auto_filter_t);
  EXPECT_EQ(fm_sheet_get_auto_filter(wb(), 0, &present, &filter), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(filter.columns, nullptr);

  ASSERT_EQ(fm_sheet_set_number(wb(), 0, 0, 0, 1.0), 0);
  present = 99;
  fm_parse_error_info_t parse_error = FM_STRUCT_INIT(fm_parse_error_info_t);
  EXPECT_EQ(fm_get_last_parse_error(&present, &parse_error), 0);
  EXPECT_EQ(present, 0);
}

TEST_F(FormulonCApiRemoveMissingKey, TheLastParseErrorIsPresentAfterASyntaxFailure) {
  EXPECT_NE(fm_sheet_set_formula(wb(), 0, 0, 0, "=SUM(1,"), 0);
  std::int32_t present = 99;
  fm_parse_error_info_t parse_error = FM_STRUCT_INIT(fm_parse_error_info_t);
  ASSERT_EQ(fm_get_last_parse_error(&present, &parse_error), 0);
  EXPECT_EQ(present, 1);
  EXPECT_GE(parse_error.line, 1U);
}

// A function name no locale table knows is a key that matches nothing; an
// unknown profile stays an invalid argument.
TEST_F(FormulonCApiRemoveMissingKey, AnUnknownFunctionNameIsNotFound) {
  const char* out = nullptr;
  ExpectNotFound(fm_localize_function_name("NOPE", "mac-365-de_DE", &out), "localize");
  ExpectNotFound(fm_canonicalize_function_name("NOPE", "mac-365-de_DE", &out), "canonicalize");
  EXPECT_EQ(fm_localize_function_name("SUM", "xx-365-zz_ZZ", &out), FM_STATUS_INVALID_ARGUMENT);
}

TEST_F(FormulonCApiRemoveMissingKey, AnAbsentSheetIsStillASheetIndexError) {
  EXPECT_EQ(fm_sheet_remove_hyperlink(wb(), 9, 0, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_row_break(wb(), 9, 7, nullptr), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_remove_data_table(wb(), 9, 0, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST_F(FormulonCApiRemoveMissingKey, APresentKeyRemovesOnceThenIsNotFound) {
  ASSERT_EQ(SetDefinedName(wb(), "Rate", "0.5"), 0);
  EXPECT_EQ(fm_workbook_remove_defined_name(wb(), "rate", -1), 0);
  ExpectNotFound(fm_workbook_remove_defined_name(wb(), "rate", -1), "defined name again");

  fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
  link.rect = fm_cell_rect_t{3, 3, 3, 3};
  link.target = "https://example.com/";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb(), 0, 3, 3, &link), 0);
  EXPECT_EQ(fm_sheet_remove_hyperlink(wb(), 0, 3, 3), 0);
  ExpectNotFound(fm_sheet_remove_hyperlink(wb(), 0, 3, 3), "hyperlink again");

  const fm_cell_rect_t merged{0, 0, 1, 1};
  ASSERT_EQ(fm_sheet_add_merge(wb(), 0, &merged), 0);
  EXPECT_EQ(fm_sheet_remove_merge(wb(), 0, &merged), 0);
  ExpectNotFound(fm_sheet_remove_merge(wb(), 0, &merged), "merge again");
}

TEST_F(FormulonCApiRemoveMissingKey, AFailedRemoveLeavesTheOtherEntriesAlone) {
  ASSERT_EQ(SetDefinedName(wb(), "Keep", "1"), 0);
  ExpectNotFound(fm_workbook_remove_defined_name(wb(), "Other", -1), "other name");
  size_t count = 0;
  ASSERT_EQ(fm_workbook_get_defined_name_count(wb(), &count), 0);
  EXPECT_EQ(count, 1U);
}
