// C ABI pivot cache refresh and the sharedItems type flags: refresh rebuilds a
// cache from its worksheet source and re-projects the pivots drawing from it,
// leaves everything unchanged when it fails, and the type getter reports
// effective flag values.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <string>

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "pivot/pivot_cache.h"
#include "sheet.h"
#include "workbook.h"

namespace {

constexpr fm_status_t kNotImplemented = static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotImplemented);
constexpr fm_status_t kNotFound = static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound);
constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);
constexpr fm_status_t kIndexOutOfRange = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kSecurityResourceLimit =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kSecurityResourceLimit);

constexpr std::size_t kRegionField = 0U;
constexpr std::size_t kAmountField = 1U;

// The scratch pivot over Sheet1!A1:B5, with the source cells filled in to match the cache records.
struct RefreshFixture {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;

  void Build() {
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, nullptr), 0) << fm_get_last_error_message();
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 0U, 0U, "Region"), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 0U, 1U, "Amount"), 0);
    const char* const regions[] = {"North", "South", "North", "South"};
    const double amounts[] = {100.0, 200.0, 300.0, 400.0};
    for (std::uint32_t i = 0; i < 4U; ++i) {
      ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, i + 1U, 0U, regions[i]), 0);
      ASSERT_EQ(fm_sheet_set_number(wb.handle, 0U, i + 1U, 1U, amounts[i]), 0);
    }
    ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, 0U), 0) << fm_get_last_error_message();
  }

  std::size_t SharedItemCount(std::size_t field) const {
    std::size_t count = 0;
    EXPECT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(wb.handle, cache_id, field, &count), 0);
    return count;
  }

  std::string SharedItemText(std::size_t field, std::size_t item) const {
    fm_value_t v{};
    EXPECT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(wb.handle, cache_id, field, item, &v), 0);
    return v.kind == FM_VALUE_KIND_TEXT ? std::string(v.u.text) : std::string();
  }

  // Whether the text `needle` sits in the pivot's neighbourhood, D1:F12.
  bool ReportHasText(const char* needle) const {
    for (std::uint32_t r = 0; r < 12U; ++r) {
      for (std::uint32_t c = 3U; c < 6U; ++c) {
        fm_value_t v{};
        if (fm_sheet_get_value(wb.handle, 0U, r, c, &v) == 0 && v.kind == FM_VALUE_KIND_TEXT &&
            std::strcmp(v.u.text, needle) == 0) {
          return true;
        }
      }
    }
    return false;
  }
};

fm_pivot_cache_field_types_t FieldTypes(const fm_workbook_t* wb, std::uint32_t cache_id, std::size_t field) {
  fm_pivot_cache_field_types_t types = FM_STRUCT_INIT(fm_pivot_cache_field_types_t);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_types(wb, cache_id, field, &types), 0) << fm_get_last_error_message();
  return types;
}

TEST(CApiPivotRefresh, RebuildsTheCacheAndTheReportFromTheSource) {
  RefreshFixture f;
  f.Build();
  ASSERT_EQ(f.SharedItemCount(kRegionField), 2U);
  ASSERT_FALSE(f.ReportHasText("East"));

  ASSERT_EQ(fm_sheet_set_text(f.wb.handle, 0U, 4U, 0U, "East"), 0);
  ASSERT_EQ(fm_sheet_set_number(f.wb.handle, 0U, 1U, 1U, 1000.0), 0);
  ASSERT_EQ(fm_workbook_refresh_pivot_cache(f.wb.handle, f.cache_id), 0) << fm_get_last_error_message();

  // Shared items follow the source order: North, South, East.
  ASSERT_EQ(f.SharedItemCount(kRegionField), 3U);
  EXPECT_EQ(f.SharedItemText(kRegionField, 0U), "North");
  EXPECT_EQ(f.SharedItemText(kRegionField, 2U), "East");
  std::size_t records = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_record_count(f.wb.handle, f.cache_id, &records), 0);
  EXPECT_EQ(records, 4U);
  fm_value_t amount{};
  ASSERT_EQ(fm_workbook_get_pivot_cache_record_value(f.wb.handle, f.cache_id, 0U, kAmountField, &amount), 0);
  EXPECT_EQ(amount.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(amount.u.number, 1000.0);
  EXPECT_TRUE(f.ReportHasText("East"));
}

TEST(CApiPivotRefresh, ReportsStatusesForBadArguments) {
  RefreshFixture f;
  f.Build();
  EXPECT_EQ(fm_workbook_refresh_pivot_cache(nullptr, f.cache_id), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_refresh_pivot_cache(f.wb.handle, f.cache_id + 100U), kNotFound);
  EXPECT_EQ(fm_workbook_refresh_pivot_cache(f.wb.handle, 0U), kNotFound);
}

TEST(CApiPivotRefresh, AnUnresolvedSourceChangesNothing) {
  RefreshFixture f;
  f.Build();
  ASSERT_EQ(SetPivotCacheWorksheetSource(f.wb.handle, f.cache_id, /*present=*/1, "A1:B5", "NoSuchSheet", nullptr), 0);
  ASSERT_EQ(fm_sheet_set_text(f.wb.handle, 0U, 4U, 0U, "East"), 0);
  EXPECT_EQ(fm_workbook_refresh_pivot_cache(f.wb.handle, f.cache_id), kPreconditionFailed);
  EXPECT_EQ(f.SharedItemCount(kRegionField), 2U);
  EXPECT_FALSE(f.ReportHasText("East"));

  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(f.wb.handle, f.cache_id, nullptr), 0);
  EXPECT_EQ(fm_workbook_refresh_pivot_cache(f.wb.handle, f.cache_id), kPreconditionFailed);
  EXPECT_EQ(f.SharedItemCount(kRegionField), 2U);
}

TEST(CApiPivotRefresh, AReportThatCannotGrowIsBlockedUntilItsRangeIsFree) {
  RefreshFixture f;
  f.Build();
  // The report fills D1:E4; a new item makes it one row longer, into D5.
  for (std::uint32_t r = 4U; r < 9U; ++r) {
    ASSERT_EQ(fm_sheet_set_text(f.wb.handle, 0U, r, 3U, "keep"), 0);
  }
  ASSERT_EQ(fm_sheet_set_text(f.wb.handle, 0U, 4U, 0U, "East"), 0);
  ASSERT_EQ(fm_workbook_refresh_pivot_cache(f.wb.handle, f.cache_id), 0) << fm_get_last_error_message();
  // The cache is refreshed; the report is withheld behind `#SPILL!` and the cells stay.
  EXPECT_EQ(f.SharedItemCount(kRegionField), 3U);
  EXPECT_FALSE(f.ReportHasText("East"));
  fm_value_t anchor{};
  ASSERT_EQ(fm_sheet_get_value(f.wb.handle, 0U, 0U, 3U, &anchor), 0);
  ASSERT_EQ(anchor.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(anchor.u.error_code, static_cast<std::int32_t>(formulon::ErrorCode::Spill));
  EXPECT_TRUE(f.ReportHasText("keep"));
  std::int32_t present = 0;
  fm_spill_info_t spill = FM_STRUCT_INIT(fm_spill_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_spill(f.wb.handle, 0U, 0U, &present, &spill), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
  EXPECT_EQ(spill.blocked, 1);
  EXPECT_EQ(spill.cause, FM_SPILL_CAUSE_OCCUPIED);
  EXPECT_EQ(spill.rect.first_row, 0U);
  EXPECT_EQ(spill.rect.first_col, 3U);
  EXPECT_EQ(spill.rect.last_row, 4U);
  EXPECT_EQ(spill.rect.last_col, 4U);

  // Clearing the cells renders the report on the next read, over the range it needed.
  for (std::uint32_t r = 4U; r < 9U; ++r) {
    ASSERT_EQ(fm_sheet_set_blank(f.wb.handle, 0U, r, 3U), 0);
  }
  EXPECT_TRUE(f.ReportHasText("East"));
  ASSERT_EQ(fm_sheet_get_pivot_spill(f.wb.handle, 0U, 0U, &present, &spill), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  EXPECT_EQ(spill.blocked, 0);
  EXPECT_EQ(spill.cause, FM_SPILL_CAUSE_ABSENT);
  EXPECT_EQ(spill.rect.last_row, 4U);
}

TEST(CApiPivotRefresh, ASourceOverTheRecordCellLimitIsRefused) {
  RefreshFixture f;
  f.Build();
  // 1048575 data rows x 33 columns is past 2^25 cells.
  ASSERT_EQ(SetPivotCacheWorksheetSource(f.wb.handle, f.cache_id, /*present=*/1, "A1:AG1048576", "Sheet1", nullptr), 0);
  EXPECT_EQ(fm_workbook_refresh_pivot_cache(f.wb.handle, f.cache_id), kSecurityResourceLimit);
  EXPECT_EQ(f.SharedItemCount(kRegionField), 2U);
}

TEST(CApiPivotCacheFieldTypes, ReadsDerivedFlagsOfAWorkbookBuiltCache) {
  RefreshFixture f;
  f.Build();
  const fm_pivot_cache_field_types_t region = FieldTypes(f.wb.handle, f.cache_id, kRegionField);
  EXPECT_EQ(region.contains_string, 1);
  EXPECT_EQ(region.contains_number, 0);
  EXPECT_EQ(region.contains_blank, 0);
  EXPECT_EQ(region.contains_date, 0);
  EXPECT_EQ(region.contains_mixed_types, 0);
  EXPECT_EQ(region.contains_semi_mixed_types, 1);
  EXPECT_EQ(region.contains_non_date, 1);

  const fm_pivot_cache_field_types_t amount = FieldTypes(f.wb.handle, f.cache_id, kAmountField);
  EXPECT_EQ(amount.contains_number, 1);
  EXPECT_EQ(amount.contains_integer, 1);
  EXPECT_EQ(amount.contains_string, 0);
  EXPECT_EQ(amount.contains_blank, 0);
  EXPECT_EQ(amount.contains_semi_mixed_types, 0);
  EXPECT_EQ(amount.contains_non_date, 1);
}

TEST(CApiPivotCacheFieldTypes, ReadsRefreshedFlags) {
  RefreshFixture f;
  f.Build();
  ASSERT_EQ(fm_sheet_set_number(f.wb.handle, 0U, 1U, 1U, 1.5), 0);
  ASSERT_EQ(fm_sheet_set_blank(f.wb.handle, 0U, 2U, 1U), 0);
  ASSERT_EQ(fm_workbook_refresh_pivot_cache(f.wb.handle, f.cache_id), 0) << fm_get_last_error_message();
  const fm_pivot_cache_field_types_t amount = FieldTypes(f.wb.handle, f.cache_id, kAmountField);
  EXPECT_EQ(amount.contains_number, 1);
  EXPECT_EQ(amount.contains_integer, 0);
  EXPECT_EQ(amount.contains_blank, 1);
  EXPECT_EQ(amount.contains_semi_mixed_types, 1);
  EXPECT_EQ(amount.contains_string, 0);
}

TEST(CApiPivotCacheFieldTypes, ReportsStatusesForBadArguments) {
  RefreshFixture f;
  f.Build();
  fm_pivot_cache_field_types_t types = FM_STRUCT_INIT(fm_pivot_cache_field_types_t);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_types(nullptr, f.cache_id, 0U, &types), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_types(f.wb.handle, f.cache_id, 0U, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_types(f.wb.handle, f.cache_id + 100U, 0U, &types), kNotFound);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_types(f.wb.handle, f.cache_id, 2U, &types), kIndexOutOfRange);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_types(f.wb.handle, f.cache_id, std::size_t{1} << 31, &types),
            kIndexOutOfRange);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_types(f.wb.handle, f.cache_id, SIZE_MAX, &types), kIndexOutOfRange);
}

// Numeric grouping of the Amount field over the values 20, 150, 250 and 300.
struct NumericGroupFixture : RefreshFixture {
  void BuildNumbers() {
    Build();
    const double amounts[] = {20.0, 150.0, 250.0, 300.0};
    for (std::size_t i = 0; i < 4U; ++i) {
      ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, i, kAmountField, amounts[i]), 0);
    }
  }

  fm_status_t Group(const fm_pivot_numeric_group_spec_t* spec) {
    return fm_sheet_set_pivot_field_numeric_group(wb.handle, 0U, 0U, kAmountField, spec);
  }

  bool Read(std::size_t field, fm_pivot_numeric_group_spec_t* out) const {
    std::int32_t present = -1;
    *out = FM_STRUCT_INIT(fm_pivot_numeric_group_spec_t);
    EXPECT_EQ(fm_sheet_get_pivot_field_numeric_group(wb.handle, 0U, 0U, field, &present, out), 0)
        << fm_get_last_error_message();
    return present == 1;
  }

  std::size_t FieldCount() const {
    std::size_t count = 0;
    EXPECT_EQ(fm_sheet_get_pivot_field_count(wb.handle, 0U, 0U, &count), 0);
    return count;
  }
};

TEST(CApiPivotNumericGroup, AutoBoundsGroupTheValuesAndReadBack) {
  NumericGroupFixture f;
  f.BuildNumbers();
  fm_pivot_numeric_group_spec_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_spec_t);
  spec.interval = 100.0;
  ASSERT_EQ(f.Group(&spec), 0) << fm_get_last_error_message();
  ASSERT_EQ(f.FieldCount(), 3U);

  fm_pivot_numeric_group_spec_t read{};
  ASSERT_TRUE(f.Read(kAmountField, &read));
  EXPECT_EQ(read.has_start, 0);
  EXPECT_EQ(read.has_end, 0);
  EXPECT_EQ(read.start, 20.0);
  EXPECT_EQ(read.end, 300.0);
  EXPECT_EQ(read.interval, 100.0);
  // The group field reads the same grouping.
  fm_pivot_numeric_group_spec_t from_group{};
  ASSERT_TRUE(f.Read(2U, &from_group));
  EXPECT_EQ(from_group.interval, 100.0);

  // Items: <20, 20-119, 120-219, 220-319, >320.
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(f.wb.handle, 0U, 0U, 2U, FM_PIVOT_AXIS_ROW), 0);
  ASSERT_EQ(fm_sheet_project_pivot(f.wb.handle, 0U, 0U), 0) << fm_get_last_error_message();
  std::size_t items = 0;
  ASSERT_EQ(fm_sheet_get_pivot_field_item_count(f.wb.handle, 0U, 0U, 2U, &items), 0);
  EXPECT_EQ(items, 5U);
  EXPECT_TRUE(f.ReportHasText("20-119"));
  EXPECT_TRUE(f.ReportHasText("120-219"));
  EXPECT_TRUE(f.ReportHasText("220-319"));

  // A value read can be written back unchanged.
  ASSERT_EQ(f.Group(&read), 0) << fm_get_last_error_message();
  fm_pivot_numeric_group_spec_t again{};
  ASSERT_TRUE(f.Read(kAmountField, &again));
  EXPECT_EQ(again.has_start, 0);
  EXPECT_EQ(again.start, 20.0);
  EXPECT_EQ(again.interval, 100.0);
}

TEST(CApiPivotNumericGroup, ExplicitBoundsAreReadAsSet) {
  NumericGroupFixture f;
  f.BuildNumbers();
  fm_pivot_numeric_group_spec_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_spec_t);
  spec.has_start = 1;
  spec.start = 0.0;
  spec.has_end = 1;
  spec.end = 400.0;
  spec.interval = 100.0;
  ASSERT_EQ(f.Group(&spec), 0) << fm_get_last_error_message();
  fm_pivot_numeric_group_spec_t read{};
  ASSERT_TRUE(f.Read(kAmountField, &read));
  EXPECT_EQ(read.has_start, 1);
  EXPECT_EQ(read.start, 0.0);
  EXPECT_EQ(read.has_end, 1);
  EXPECT_EQ(read.end, 400.0);
}

TEST(CApiPivotNumericGroup, ANullGroupRemovesTheGrouping) {
  NumericGroupFixture f;
  f.BuildNumbers();
  fm_pivot_numeric_group_spec_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_spec_t);
  spec.interval = 100.0;
  ASSERT_EQ(f.Group(&spec), 0);
  ASSERT_EQ(f.FieldCount(), 3U);
  ASSERT_EQ(f.Group(nullptr), 0) << fm_get_last_error_message();
  EXPECT_EQ(f.FieldCount(), 2U);
  fm_pivot_numeric_group_spec_t read{};
  EXPECT_FALSE(f.Read(kAmountField, &read));
  EXPECT_EQ(read.interval, 0.0);
}

TEST(CApiPivotNumericGroup, RefusesWhatItCannotGroup) {
  NumericGroupFixture f;
  f.BuildNumbers();
  fm_pivot_numeric_group_spec_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_spec_t);
  spec.interval = 1.5;
  EXPECT_EQ(f.Group(&spec), kNotImplemented);
  spec.interval = 100.0;
  spec.has_start = 1;
  spec.start = 0.5;
  EXPECT_EQ(f.Group(&spec), kNotImplemented);
  spec.has_start = 2;
  EXPECT_EQ(f.Group(&spec), kInvalidArgument);
  spec.has_start = 0;
  spec.interval = 0.0;
  EXPECT_EQ(f.Group(&spec), kInvalidArgument);
  spec.interval = std::numeric_limits<double>::quiet_NaN();
  EXPECT_EQ(f.Group(&spec), kInvalidArgument);
  EXPECT_EQ(f.FieldCount(), 2U);
}

TEST(CApiPivotNumericGroup, ReportsStatusesForBadArguments) {
  NumericGroupFixture f;
  f.BuildNumbers();
  std::int32_t present = -1;
  fm_pivot_numeric_group_spec_t numeric = FM_STRUCT_INIT(fm_pivot_numeric_group_spec_t);
  fm_pivot_date_group_spec_t date = FM_STRUCT_INIT(fm_pivot_date_group_spec_t);
  EXPECT_EQ(fm_sheet_get_pivot_field_numeric_group(nullptr, 0U, 0U, 0U, &present, &numeric), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_pivot_field_numeric_group(f.wb.handle, 0U, 0U, 0U, nullptr, &numeric), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_pivot_field_numeric_group(f.wb.handle, 0U, 0U, 0U, &present, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_pivot_field_numeric_group(f.wb.handle, 0U, 0U, 9U, &present, &numeric), kIndexOutOfRange);
  EXPECT_EQ(fm_sheet_get_pivot_field_date_group(nullptr, 0U, 0U, 0U, &present, &date), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_pivot_field_date_group(f.wb.handle, 0U, 0U, 0U, nullptr, &date), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_pivot_field_date_group(f.wb.handle, 0U, 0U, 0U, &present, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_pivot_field_date_group(f.wb.handle, 0U, 0U, 9U, &present, &date), kIndexOutOfRange);
  EXPECT_EQ(fm_sheet_set_pivot_field_numeric_group(nullptr, 0U, 0U, 0U, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_pivot_field_date_group(f.wb.handle, 0U, 0U, 0U, &present, &date), 0);
  EXPECT_EQ(present, 0);
}

}  // namespace
