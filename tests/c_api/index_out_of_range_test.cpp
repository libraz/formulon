// Index addressing a counted collection: `>= count` is
// FM_STATUS_INDEX_OUT_OF_RANGE (a sheet index keeps its own status), while an
// enum ordinal or a limit stays FM_STATUS_INVALID_ARGUMENT.

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

class FormulonCApiIndexOutOfRange : public ::testing::Test {
 protected:
  void SetUp() override { ASSERT_EQ(fm_workbook_create_default(&wb_.handle), 0); }
  fm_workbook_t* wb() { return wb_.handle; }

 private:
  WorkbookGuard wb_;
};

}  // namespace

TEST_F(FormulonCApiIndexOutOfRange, TheIndexEqualToTheCountIsOutOfRangeAndTheLastIndexIsNot) {
  size_t count = 0;
  ASSERT_EQ(fm_workbook_get_font_count(wb(), &count), 0);
  ASSERT_GT(count, 0U);
  fm_font_record_t font{};
  EXPECT_EQ(fm_workbook_get_font(wb(), count - 1U, &font), 0);
  EXPECT_EQ(fm_workbook_get_font(wb(), count, &font), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(std::string(fm_get_last_error_context()),
            "font_index=" + std::to_string(count) + " count=" + std::to_string(count));
}

TEST_F(FormulonCApiIndexOutOfRange, StyleCollectionsReportTheSameStatus) {
  fm_fill_record_t fill{};
  EXPECT_EQ(fm_workbook_get_fill(wb(), 999, &fill), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_border_record_t border{};
  EXPECT_EQ(fm_workbook_get_border(wb(), 999, &border), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_cell_xf_t xf{};
  EXPECT_EQ(fm_workbook_get_cell_xf(wb(), 999, &xf), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_dxf_record_t dxf{};
  EXPECT_EQ(fm_workbook_get_dxf(wb(), 999, &dxf), FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST_F(FormulonCApiIndexOutOfRange, TablePivotAndSparklineIndicesReportTheSameStatus) {
  fm_table_info_t table = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_EQ(fm_workbook_get_table_info(wb(), 0, &table), FM_STATUS_INDEX_OUT_OF_RANGE);
  uint32_t cache_id = 0;
  EXPECT_EQ(fm_sheet_get_pivot_cache_id(wb(), 0, 0, &cache_id), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_sparkline_group_t group = FM_STRUCT_INIT(fm_sparkline_group_t);
  EXPECT_EQ(fm_sheet_get_sparkline_group(wb(), 0, 0, &group), FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST_F(FormulonCApiIndexOutOfRange, MutatorsRefuseAnIndexPastTheCountBeforeChangingAnything) {
  fm_font_record_t font{};
  ASSERT_EQ(fm_workbook_get_font(wb(), 0, &font), 0);
  EXPECT_EQ(fm_workbook_set_font(wb(), 999, &font), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_remove_sparkline_group(wb(), 0, 0), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_remove_merge_at(wb(), 0, 0), FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST_F(FormulonCApiIndexOutOfRange, ASheetIndexKeepsItsOwnStatus) {
  const char* name = nullptr;
  EXPECT_EQ(fm_workbook_get_sheet_name(wb(), 5, &name), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  size_t count = 0;
  EXPECT_EQ(fm_sheet_get_cell_count(wb(), 5, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

TEST_F(FormulonCApiIndexOutOfRange, EnumOrdinalsAndLimitsStayInvalidArguments) {
  EXPECT_EQ(fm_sheet_set_error(wb(), 0, 0, 0, 9999), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_set_row_outline(wb(), 0, 0, 0, 200), FM_STATUS_INVALID_ARGUMENT);
}
