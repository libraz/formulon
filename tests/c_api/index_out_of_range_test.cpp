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
  fm_font_t font{};
  EXPECT_EQ(fm_workbook_get_font_at(wb(), count - 1U, &font), 0);
  EXPECT_EQ(fm_workbook_get_font_at(wb(), count, &font), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(std::string(fm_get_last_error_context()),
            "font_index=" + std::to_string(count) + " count=" + std::to_string(count));
}

TEST_F(FormulonCApiIndexOutOfRange, StyleCollectionsReportTheSameStatus) {
  fm_fill_t fill{};
  EXPECT_EQ(fm_workbook_get_fill_at(wb(), 999, &fill), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_border_t border{};
  EXPECT_EQ(fm_workbook_get_border_at(wb(), 999, &border), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_cell_xf_t xf{};
  EXPECT_EQ(fm_workbook_get_cell_xf_at(wb(), 999, &xf), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_dxf_t dxf{};
  EXPECT_EQ(fm_workbook_get_dxf_at(wb(), 999, &dxf), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_get_cell_style_xf_at(wb(), 999, &xf), FM_STATUS_INDEX_OUT_OF_RANGE);
  size_t xf_count = 0;
  ASSERT_EQ(fm_workbook_get_cell_xf_count(wb(), &xf_count), 0);
  const auto past = static_cast<uint32_t>(xf_count);
  EXPECT_EQ(fm_sheet_set_column_xf_index(wb(), 0, 0, 0, past), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_row_xf_index(wb(), 0, 0, 0, past), FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST_F(FormulonCApiIndexOutOfRange, LocaleErrorNamePastTheCountIsOutOfRange) {
  size_t count = 0;
  ASSERT_EQ(fm_get_locale_error_name_count(&count), 0);
  const char* canonical = nullptr;
  const char* localized = nullptr;
  int32_t measured = -1;
  EXPECT_EQ(fm_get_locale_error_name_at("win-365-en_US", count - 1U, &canonical, &localized, &measured), 0);
  EXPECT_EQ(fm_get_locale_error_name_at("win-365-en_US", count, &canonical, &localized, &measured),
            FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST_F(FormulonCApiIndexOutOfRange, TablePivotAndSparklineIndicesReportTheSameStatus) {
  fm_table_info_t table = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_EQ(fm_sheet_get_table_at(wb(), 0, 0, &table), FM_STATUS_INDEX_OUT_OF_RANGE);
  uint32_t cache_id = 0;
  EXPECT_EQ(fm_sheet_get_pivot_cache_id(wb(), 0, 0, &cache_id), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_sparkline_group_t group = FM_STRUCT_INIT(fm_sparkline_group_t);
  EXPECT_EQ(fm_sheet_get_sparkline_group_at(wb(), 0, 0, &group), FM_STATUS_INDEX_OUT_OF_RANGE);
}

// The positional reads of keyed collections stay positional.
TEST_F(FormulonCApiIndexOutOfRange, PositionalReadsOfKeyedCollectionsReportTheSameStatus) {
  fm_cf_rule_t rule = FM_STRUCT_INIT(fm_cf_rule_t);
  EXPECT_EQ(fm_sheet_get_cf_rule_at(wb(), 0, 0, &rule), FM_STATUS_INDEX_OUT_OF_RANGE);
  fm_protected_range_t range = FM_STRUCT_INIT(fm_protected_range_t);
  EXPECT_EQ(fm_sheet_get_protected_range_at(wb(), 0, 0, &range), FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST_F(FormulonCApiIndexOutOfRange, MutatorsRefuseAnIndexPastTheCountBeforeChangingAnything) {
  fm_font_t font{};
  ASSERT_EQ(fm_workbook_get_font_at(wb(), 0, &font), 0);
  EXPECT_EQ(fm_workbook_set_font(wb(), 999, &font), FM_STATUS_INDEX_OUT_OF_RANGE);
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
