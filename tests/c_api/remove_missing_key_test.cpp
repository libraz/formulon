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
  ExpectNotFound(fm_sheet_remove_shape(wb(), 0, 42U), "shape");
  ExpectNotFound(fm_sheet_remove_slicer(wb(), 0, 42U), "slicer");
  ExpectNotFound(fm_sheet_remove_data_table(wb(), 0, 0, 0), "data table");
  ExpectNotFound(fm_sheet_remove_drawing(wb(), 0, 42U), "drawing");
}

// The chart readers address a chart by its object id as well: an absent id is
// FM_STATUS_NOT_FOUND and hands out no snapshot.
TEST_F(FormulonCApiRemoveMissingKey, ChartReadersOfAnAbsentId) {
  fm_chart_data_t* data = nullptr;
  ExpectNotFound(fm_sheet_get_chart_data(wb(), 0, 42U, &data), "chart data");
  EXPECT_EQ(data, nullptr);
  const fm_chart_preview_overrides_t overrides = FM_STRUCT_INIT(fm_chart_preview_overrides_t);
  ExpectNotFound(fm_sheet_get_chart_preview(wb(), 0, 42U, &overrides, &data), "chart preview");
  EXPECT_EQ(data, nullptr);
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
