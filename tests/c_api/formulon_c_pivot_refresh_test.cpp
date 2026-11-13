// C ABI pivot cache refresh and the sharedItems type flags: refresh rebuilds a
// cache from its worksheet source and re-projects the pivots drawing from it,
// leaves everything unchanged when it fails, and the type getter reports
// effective flag values.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <string>
#include <utility>
#include <vector>

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_resource_limit.h"
#include "pivot/pivot_cache.h"
#include "sheet.h"
#include "workbook.h"

namespace {

constexpr fm_status_t kNotFound = static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound);
constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);
constexpr fm_status_t kIndexOutOfRange = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIndexOutOfRange);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

constexpr std::size_t kRegionField = 0U;
constexpr std::size_t kAmountField = 1U;

// The scratch pivot over Sheet1!A1:B5, with the source cells filled in to match the cache records.
struct RefreshFixture {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;

  void Build() {
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, nullptr), 0) << fm_get_last_error_message();
    FillScratchSource(wb.handle);
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

fm_pivot_cache_field_types_info_t FieldTypes(const fm_workbook_t* wb, std::uint32_t cache_id, std::size_t field) {
  fm_pivot_cache_field_types_info_t types = FM_STRUCT_INIT(fm_pivot_cache_field_types_info_t);
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
  const std::string context =
      ExpectResourceLimit(fm_workbook_refresh_pivot_cache(f.wb.handle, f.cache_id), "pivot_cache_record_cells");
  EXPECT_NE(context.find(" limit=33554432"), std::string::npos) << context;
  EXPECT_EQ(f.SharedItemCount(kRegionField), 2U);
}

// --- Refresh on load ---------------------------------------------------------
// A cache built in memory is saved with refreshOnLoad, so the reloaded book refreshes it from the
// source before anything else reads or writes it, as an Excel open does.

std::vector<std::uint8_t> SaveAs(fm_workbook_t* wb, fm_workbook_format_t format) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, format, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  return std::vector<std::uint8_t>(saved.data, saved.data + saved.len);
}

void Load(const std::vector<std::uint8_t>& bytes, WorkbookGuard& out) {
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out.handle), 0) << fm_get_last_error_message();
}

double RecordNumber(const fm_workbook_t* wb, std::uint32_t cache_id, std::size_t record, std::size_t field) {
  fm_value_t v{};
  EXPECT_EQ(fm_workbook_get_pivot_cache_record_value(wb, cache_id, record, field, &v), 0)
      << fm_get_last_error_message();
  return v.kind == FM_VALUE_KIND_NUMBER ? v.u.number : std::numeric_limits<double>::quiet_NaN();
}

TEST(CApiPivotRefreshOnLoad, AnEditBeforeTheFirstEvaluationIsKept) {
  RefreshFixture f;
  f.Build();
  WorkbookGuard loaded;
  Load(SaveAs(f.wb.handle, FM_WORKBOOK_FORMAT_XLSX), loaded);
  std::size_t record = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_record(loaded.handle, f.cache_id, &record), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(loaded.handle, f.cache_id, record, kRegionField, 0.0), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(loaded.handle, f.cache_id, record, kAmountField, 7.0), 0);
  PivotCellsGuard layout;
  ASSERT_EQ(fm_sheet_get_pivot_layout(loaded.handle, 0, 0, &layout.handle), 0) << fm_get_last_error_message();
  std::size_t records = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_record_count(loaded.handle, f.cache_id, &records), 0);
  EXPECT_EQ(records, 5U);
  EXPECT_EQ(RecordNumber(loaded.handle, f.cache_id, 4U, kAmountField), 7.0);
}

TEST(CApiPivotRefreshOnLoad, TheProfileSetAfterTheLoadNamesBlankHeaders) {
  RefreshFixture f;
  f.Build();
  // A third source column with a blank header joins the cache as a field the profile names.
  for (std::uint32_t r = 1U; r < 5U; ++r) {
    ASSERT_EQ(fm_sheet_set_number(f.wb.handle, 0U, r, 2U, 1.0), 0);
  }
  ASSERT_EQ(SetPivotCacheWorksheetSource(f.wb.handle, f.cache_id, /*present=*/1, "A1:C5", "Sheet1", nullptr), 0);
  const std::vector<std::uint8_t> saved = SaveAs(f.wb.handle, FM_WORKBOOK_FORMAT_XLSX);
  for (const auto& [profile, word] : {std::pair<const char*, const char*>{"win-365-ja_JP", "フィールド3"},
                                      std::pair<const char*, const char*>{"win-365-de_DE", "Feld3"}}) {
    SCOPED_TRACE(profile);
    WorkbookGuard loaded;
    Load(saved, loaded);
    ASSERT_EQ(fm_workbook_set_profile_id(loaded.handle, profile), 0);
    std::size_t fields = 0;
    ASSERT_EQ(fm_workbook_get_pivot_cache_field_count(loaded.handle, f.cache_id, &fields), 0);
    ASSERT_EQ(fields, 3U);
    const char* name = nullptr;
    ASSERT_EQ(fm_workbook_get_pivot_cache_field_name(loaded.handle, f.cache_id, 2U, &name), 0);
    EXPECT_STREQ(name, word);
  }
}

TEST(CApiPivotRefreshOnLoad, ASaveWithNoEvaluationWritesTheRefreshedRecords) {
  RefreshFixture f;
  f.Build();
  // The stored records keep 100; the source says 1000.
  ASSERT_EQ(fm_sheet_set_number(f.wb.handle, 0U, 1U, 1U, 1000.0), 0);
  WorkbookGuard loaded;
  Load(SaveAs(f.wb.handle, FM_WORKBOOK_FORMAT_XLSX), loaded);
  const std::vector<std::uint8_t> xlsx = SaveAs(loaded.handle, FM_WORKBOOK_FORMAT_XLSX);
  const std::string records = ExtractZipEntry(xlsx, "xl/pivotCache/pivotCacheRecords1.xml");
  EXPECT_NE(records.find("<n v=\"1000\"/>"), std::string::npos) << records;
  EXPECT_EQ(records.find("<n v=\"100\"/>"), std::string::npos) << records;
}

TEST(CApiPivotCacheFieldTypes, ReadsDerivedFlagsOfAWorkbookBuiltCache) {
  RefreshFixture f;
  f.Build();
  const fm_pivot_cache_field_types_info_t region = FieldTypes(f.wb.handle, f.cache_id, kRegionField);
  EXPECT_EQ(region.contains_string, 1);
  EXPECT_EQ(region.contains_number, 0);
  EXPECT_EQ(region.contains_blank, 0);
  EXPECT_EQ(region.contains_date, 0);
  EXPECT_EQ(region.contains_mixed_types, 0);
  EXPECT_EQ(region.contains_semi_mixed_types, 1);
  EXPECT_EQ(region.contains_non_date, 1);

  const fm_pivot_cache_field_types_info_t amount = FieldTypes(f.wb.handle, f.cache_id, kAmountField);
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
  const fm_pivot_cache_field_types_info_t amount = FieldTypes(f.wb.handle, f.cache_id, kAmountField);
  EXPECT_EQ(amount.contains_number, 1);
  EXPECT_EQ(amount.contains_integer, 0);
  EXPECT_EQ(amount.contains_blank, 1);
  EXPECT_EQ(amount.contains_semi_mixed_types, 1);
  EXPECT_EQ(amount.contains_string, 0);
}

TEST(CApiPivotCacheFieldTypes, ReportsStatusesForBadArguments) {
  RefreshFixture f;
  f.Build();
  fm_pivot_cache_field_types_info_t types = FM_STRUCT_INIT(fm_pivot_cache_field_types_info_t);
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

  fm_status_t Group(const fm_pivot_numeric_group_t* spec) {
    return fm_sheet_set_pivot_field_numeric_group(wb.handle, 0U, 0U, kAmountField, spec);
  }

  bool Read(std::size_t field, fm_pivot_numeric_group_t* out) const {
    std::int32_t present = -1;
    *out = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
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
  fm_pivot_numeric_group_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  spec.interval = 100.0;
  ASSERT_EQ(f.Group(&spec), 0) << fm_get_last_error_message();
  ASSERT_EQ(f.FieldCount(), 3U);

  fm_pivot_numeric_group_t read{};
  ASSERT_TRUE(f.Read(kAmountField, &read));
  EXPECT_EQ(read.auto_start, FM_TRISTATE_TRUE);
  EXPECT_EQ(read.auto_end, FM_TRISTATE_TRUE);
  EXPECT_EQ(read.start, 20.0);
  EXPECT_EQ(read.end, 300.0);
  EXPECT_EQ(read.interval, 100.0);
  // The group field reads the same grouping.
  fm_pivot_numeric_group_t from_group{};
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
  fm_pivot_numeric_group_t again{};
  ASSERT_TRUE(f.Read(kAmountField, &again));
  EXPECT_EQ(again.auto_start, FM_TRISTATE_TRUE);
  EXPECT_EQ(again.auto_end, FM_TRISTATE_TRUE);
  EXPECT_EQ(again.start, 20.0);
  EXPECT_EQ(again.end, 300.0);
  EXPECT_EQ(again.interval, 100.0);
}

TEST(CApiPivotNumericGroup, ExplicitBoundsAreReadAsSetAndWriteBackUnchanged) {
  NumericGroupFixture f;
  f.BuildNumbers();
  fm_pivot_numeric_group_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  spec.auto_start = FM_TRISTATE_FALSE;
  spec.start = 0.0;
  spec.auto_end = FM_TRISTATE_FALSE;
  spec.end = 400.0;
  spec.interval = 100.0;
  ASSERT_EQ(f.Group(&spec), 0) << fm_get_last_error_message();
  fm_pivot_numeric_group_t read{};
  ASSERT_TRUE(f.Read(kAmountField, &read));
  EXPECT_EQ(read.auto_start, FM_TRISTATE_FALSE);
  EXPECT_EQ(read.start, 0.0);
  EXPECT_EQ(read.auto_end, FM_TRISTATE_FALSE);
  EXPECT_EQ(read.end, 400.0);

  ASSERT_EQ(f.Group(&read), 0) << fm_get_last_error_message();
  fm_pivot_numeric_group_t again{};
  ASSERT_TRUE(f.Read(kAmountField, &again));
  EXPECT_EQ(again.auto_start, FM_TRISTATE_FALSE);
  EXPECT_EQ(again.start, 0.0);
  EXPECT_EQ(again.auto_end, FM_TRISTATE_FALSE);
  EXPECT_EQ(again.end, 400.0);
  EXPECT_EQ(again.interval, 100.0);
}

TEST(CApiPivotNumericGroup, ABoundThatFollowsTheDataIsNotRead) {
  NumericGroupFixture f;
  f.BuildNumbers();
  fm_pivot_numeric_group_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  spec.auto_start = FM_TRISTATE_TRUE;
  spec.start = std::numeric_limits<double>::quiet_NaN();
  spec.auto_end = FM_TRISTATE_DEFAULT;
  spec.end = -std::numeric_limits<double>::infinity();
  spec.interval = 100.0;
  ASSERT_EQ(f.Group(&spec), 0) << fm_get_last_error_message();
  fm_pivot_numeric_group_t read{};
  ASSERT_TRUE(f.Read(kAmountField, &read));
  EXPECT_EQ(read.auto_start, FM_TRISTATE_TRUE);
  EXPECT_EQ(read.start, 20.0);
  EXPECT_EQ(read.auto_end, FM_TRISTATE_TRUE);
  EXPECT_EQ(read.end, 300.0);
}

TEST(CApiPivotNumericGroup, ANullGroupRemovesTheGrouping) {
  NumericGroupFixture f;
  f.BuildNumbers();
  fm_pivot_numeric_group_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  spec.interval = 100.0;
  ASSERT_EQ(f.Group(&spec), 0);
  ASSERT_EQ(f.FieldCount(), 3U);
  ASSERT_EQ(f.Group(nullptr), 0) << fm_get_last_error_message();
  EXPECT_EQ(f.FieldCount(), 2U);
  fm_pivot_numeric_group_t read{};
  EXPECT_FALSE(f.Read(kAmountField, &read));
  EXPECT_EQ(read.struct_size, sizeof(fm_pivot_numeric_group_t));
  EXPECT_EQ(read.auto_start, 0);
  EXPECT_EQ(read.start, 0.0);
  EXPECT_EQ(read.auto_end, 0);
  EXPECT_EQ(read.end, 0.0);
  EXPECT_EQ(read.interval, 0.0);
}

TEST(CApiPivotNumericGroup, RefusesBadSettingsAndGroupsNothing) {
  NumericGroupFixture f;
  f.BuildNumbers();
  const double nan = std::numeric_limits<double>::quiet_NaN();
  const double inf = std::numeric_limits<double>::infinity();
  const auto base = [] {
    fm_pivot_numeric_group_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
    spec.interval = 100.0;
    return spec;
  };
  fm_pivot_numeric_group_t spec = base();
  spec.auto_start = 3;
  EXPECT_EQ(f.Group(&spec), kInvalidArgument);
  spec = base();
  spec.auto_end = -1;
  EXPECT_EQ(f.Group(&spec), kInvalidArgument);
  for (const double bad : {nan, inf, -inf}) {
    spec = base();
    spec.auto_start = FM_TRISTATE_FALSE;
    spec.start = bad;
    EXPECT_EQ(f.Group(&spec), kInvalidArgument) << bad;
    spec = base();
    spec.auto_end = FM_TRISTATE_FALSE;
    spec.end = bad;
    EXPECT_EQ(f.Group(&spec), kInvalidArgument) << bad;
    spec = base();
    spec.interval = bad;
    EXPECT_EQ(f.Group(&spec), kInvalidArgument) << bad;
  }
  spec = base();
  spec.interval = 0.0;
  EXPECT_EQ(f.Group(&spec), kInvalidArgument);
  spec.interval = -1.5;
  EXPECT_EQ(f.Group(&spec), kInvalidArgument);
  spec = base();
  spec.auto_start = FM_TRISTATE_FALSE;
  spec.start = 300.0;
  spec.auto_end = FM_TRISTATE_FALSE;
  spec.end = 20.0;
  EXPECT_EQ(f.Group(&spec), kInvalidArgument);
  spec = base();
  spec.interval = 1e-300;
  EXPECT_EQ(f.Group(&spec), kInvalidArgument);
  EXPECT_EQ(f.FieldCount(), 2U);
  fm_pivot_numeric_group_t read{};
  EXPECT_FALSE(f.Read(kAmountField, &read));
}

TEST(CApiPivotNumericGroup, AFieldOfTextOrBlanksOrACalculatedFieldIsAPrecondition) {
  NumericGroupFixture f;
  f.BuildNumbers();
  fm_pivot_numeric_group_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  spec.interval = 100.0;
  EXPECT_EQ(fm_sheet_set_pivot_field_numeric_group(f.wb.handle, 0U, 0U, kRegionField, &spec), kPreconditionFailed);
  std::size_t calculated = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache_calculated_field(f.wb.handle, f.cache_id, "Twice", "Amount*2", &calculated), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_pivot_field_numeric_group(f.wb.handle, 0U, 0U, calculated, &spec), kPreconditionFailed);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_blank(f.wb.handle, f.cache_id, 1U, kAmountField), 0);
  EXPECT_EQ(f.Group(&spec), kPreconditionFailed);
  fm_pivot_numeric_group_t read{};
  EXPECT_FALSE(f.Read(kAmountField, &read));
  EXPECT_FALSE(f.Read(kRegionField, &read));
}

TEST(CApiPivotNumericGroup, ARefusalLeavesThePreviousGroupingInPlace) {
  NumericGroupFixture f;
  f.BuildNumbers();
  fm_pivot_numeric_group_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  spec.interval = 100.0;
  ASSERT_EQ(f.Group(&spec), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(f.wb.handle, 0U, 0U, 2U, FM_PIVOT_AXIS_ROW), 0);
  std::uint32_t order[4] = {};
  std::size_t order_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(f.wb.handle, 0U, 0U, order, 4U, &order_count), 0);
  const std::vector<std::uint32_t> before(order, order + order_count);

  spec.auto_start = 3;
  EXPECT_EQ(f.Group(&spec), kInvalidArgument);
  spec = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  spec.interval = 1e-300;
  EXPECT_EQ(f.Group(&spec), kInvalidArgument);
  spec.interval = 50.0;
  EXPECT_EQ(fm_sheet_set_pivot_field_numeric_group(f.wb.handle, 0U, 0U, kRegionField, &spec), kPreconditionFailed);

  EXPECT_EQ(f.FieldCount(), 3U);
  fm_pivot_numeric_group_t read{};
  ASSERT_TRUE(f.Read(kAmountField, &read));
  EXPECT_EQ(read.interval, 100.0);
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(f.wb.handle, 0U, 0U, order, 4U, &order_count), 0);
  EXPECT_EQ(std::vector<std::uint32_t>(order, order + order_count), before);
}

TEST(CApiPivotNumericGroup, ReportsStatusesForBadArguments) {
  NumericGroupFixture f;
  f.BuildNumbers();
  std::int32_t present = -1;
  fm_pivot_numeric_group_t numeric = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  fm_pivot_date_group_t date = FM_STRUCT_INIT(fm_pivot_date_group_t);
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

// A pivot at D1 summing Amount over a source on Sheet1 whose Amount column
// holds `amounts` (Mac Excel's measured bases), refreshed into a new cache.
// With `region_rows` the Region column holds the amounts too and sits on the
// rows; otherwise it holds text and no field is on an axis.
struct FractionalFixture {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;

  void Build(const std::vector<double>& amounts, const char* profile = nullptr, bool region_rows = false) {
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    if (profile != nullptr) {
      ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, profile), 0) << fm_get_last_error_message();
    }
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 0U, 0U, "Region"), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 0U, 1U, "Amount"), 0);
    for (std::uint32_t i = 0; i < amounts.size(); ++i) {
      if (region_rows) {
        ASSERT_EQ(fm_sheet_set_number(wb.handle, 0U, i + 1U, 0U, amounts[i]), 0);
      } else {
        ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, i + 1U, 0U, "North"), 0);
      }
      ASSERT_EQ(fm_sheet_set_number(wb.handle, 0U, i + 1U, 1U, amounts[i]), 0);
    }
    ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0U, &cache_id), 0);
    const std::string ref = "A1:B" + std::to_string(amounts.size() + 1U);
    ASSERT_EQ(SetPivotCacheWorksheetSource(wb.handle, cache_id, /*present=*/1, ref.c_str(), "Sheet1", nullptr), 0);
    ASSERT_EQ(fm_workbook_refresh_pivot_cache(wb.handle, cache_id), 0) << fm_get_last_error_message();
    ASSERT_EQ(AddPivot(wb.handle, 0U, "PT", cache_id, /*anchor_row=*/0U, /*anchor_col=*/3U, ignored_index()), 0)
        << fm_get_last_error_message();
    ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0U, 0U, kAmountField, FM_PIVOT_AXIS_VALUE), 0);
    fm_pivot_data_field_t sum = FM_STRUCT_INIT(fm_pivot_data_field_t);
    sum.name = "Sum of Amount";
    sum.field_index = static_cast<std::uint32_t>(kAmountField);
    sum.aggregation = FM_PIVOT_AGGREGATION_SUM;
    sum.number_format = "";
    sum.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
    ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0U, 0U, &sum, ignored_index()), 0)
        << fm_get_last_error_message();
    if (region_rows) {
      ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0U, 0U, kRegionField, FM_PIVOT_AXIS_ROW), 0)
          << fm_get_last_error_message();
    }
  }

  // Groups Amount; a NaN bound follows the data. The group field joins the rows.
  fm_status_t Group(double start_or_nan, double end_or_nan, double interval) {
    fm_pivot_numeric_group_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
    spec.auto_start = std::isnan(start_or_nan) ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
    spec.start = start_or_nan;
    spec.auto_end = std::isnan(end_or_nan) ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
    spec.end = end_or_nan;
    spec.interval = interval;
    const fm_status_t rc = fm_sheet_set_pivot_field_numeric_group(wb.handle, 0U, 0U, kAmountField, &spec);
    if (rc != 0) {
      return rc;
    }
    return fm_sheet_set_pivot_field_axis(wb.handle, 0U, 0U, kGroupField, FM_PIVOT_AXIS_ROW);
  }

  // The labels of the group field's items, in item order.
  std::vector<std::string> ItemLabels() const {
    std::size_t count = 0;
    EXPECT_EQ(fm_sheet_get_pivot_field_item_count(wb.handle, 0U, 0U, kGroupField, &count), 0);
    std::vector<std::string> labels;
    for (std::size_t i = 0; i < count; ++i) {
      fm_pivot_item_info_t item = FM_STRUCT_INIT(fm_pivot_item_info_t);
      EXPECT_EQ(fm_sheet_get_pivot_field_item_at(wb.handle, 0U, 0U, kGroupField, i, &item), 0);
      labels.emplace_back(item.name != nullptr ? item.name : "");
    }
    return labels;
  }

  // The rendered rows as (label, Sum of Amount); the grand total row is a cell kind of its own.
  std::vector<std::pair<std::string, double>> Rows() const {
    PivotCellsGuard projected;
    EXPECT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0U, 0U, &projected.handle), 0) << fm_get_last_error_message();
    std::vector<std::pair<std::string, double>> rows;
    std::vector<double> sums;
    for (const fm_pivot_cell_info_t& cell : CollectCells(projected.handle)) {
      if (cell.kind == FM_PIVOT_CELL_KIND_ROW_LABEL && cell.value.kind == FM_VALUE_KIND_TEXT) {
        rows.emplace_back(cell.value.u.text, 0.0);
      } else if (cell.kind == FM_PIVOT_CELL_KIND_DATA && cell.value.kind == FM_VALUE_KIND_NUMBER) {
        sums.push_back(cell.value.u.number);
      }
    }
    EXPECT_EQ(rows.size(), sums.size());
    for (std::size_t i = 0; i < rows.size() && i < sums.size(); ++i) {
      rows[i].second = sums[i];
    }
    return rows;
  }

  static constexpr std::size_t kGroupField = 2U;
};

const std::vector<double> kBaseA = {0.3, 0.6, 0.9, 1.2, 1.5, 1.8, 2.1, 2.4, 2.7, 3.0, 3.3, 3.6};
const std::vector<double> kBaseB = {-2.5, -2.0, -1.5, -1.0, -0.5, 0.0, 0.5, 1.0, 1.5, 2.0, 2.5};

void ExpectRows(const std::vector<std::pair<std::string, double>>& actual,
                const std::vector<std::pair<std::string, double>>& expected, const char* name) {
  ASSERT_EQ(actual.size(), expected.size()) << name;
  for (std::size_t i = 0; i < expected.size(); ++i) {
    EXPECT_EQ(actual[i].first, expected[i].first) << name << " row " << i;
    EXPECT_NEAR(actual[i].second, expected[i].second, 1e-9) << name << " row " << i;
  }
}

TEST(CApiPivotNumericGroup, FractionalGroupsLabelAndSumAsExcel) {
  const double kAuto = std::numeric_limits<double>::quiet_NaN();
  struct Case {
    const char* name;
    const std::vector<double>& base;
    double start;
    double end;
    double interval;
    std::vector<std::pair<std::string, double>> rows;
  };
  const std::vector<Case> cases = {
      {"c01",
       kBaseA,
       kAuto,
       kAuto,
       0.5,
       {{"0.3-0.8", 0.9},
        {"0.8-1.3", 2.1},
        {"1.3-1.8", 1.5},
        {"1.8-2.3", 3.9},
        {"2.3-2.8", 5.1},
        {"2.8-3.3", 3.0},
        {"3.3-3.8", 6.9}}},
      {"c03", kBaseA, kAuto, kAuto, 1.5, {{"0.3-1.8", 4.5}, {"1.8-3.3", 12.0}, {"3.3-4.8", 6.9}}},
      {"c05", kBaseA, 0.0, 4.0, 1.5, {{"0-1.5", 3.0}, {"1.5-3", 10.5}, {"3-4.5", 9.9}}},
      {"c06",
       kBaseB,
       kAuto,
       kAuto,
       0.5,
       {{"-2.5--2", -2.5},
        {"-2--1.5", -2.0},
        {"-1.5--1", -1.5},
        {"-1--0.5", -1.0},
        {"-0.5-0", -0.5},
        {"0-0.5", 0.0},
        {"0.5-1", 0.5},
        {"1-1.5", 1.0},
        {"1.5-2", 1.5},
        {"2-2.5", 4.5}}},
      {"c07",
       kBaseB,
       -2.25,
       2.25,
       0.75,
       {{"<-2.25", -2.5},
        {"-2.25--1.5", -2.0},
        {"-1.5--0.75", -2.5},
        {"-0.75-0", -0.5},
        {"0-0.75", 0.5},
        {"0.75-1.5", 1.0},
        {"1.5-2.25", 3.5},
        {">2.25", 2.5}}},
      {"c09",
       kBaseA,
       1.0,
       3.0,
       0.5,
       {{"<1", 1.8}, {"1-1.5", 1.2}, {"1.5-2", 3.3}, {"2-2.5", 4.5}, {"2.5-3", 5.7}, {">3", 6.9}}},
  };
  for (const Case& c : cases) {
    FractionalFixture f;
    f.Build(c.base);
    ASSERT_EQ(f.Group(c.start, c.end, c.interval), 0) << c.name << ": " << fm_get_last_error_message();
    ExpectRows(f.Rows(), c.rows, c.name);
  }
}

TEST(CApiPivotNumericGroup, FractionalGroupItemsIncludeTheEmptyRanges) {
  FractionalFixture f;
  f.Build(kBaseA);
  ASSERT_EQ(f.Group(std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(), 0.5), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(f.ItemLabels(), (std::vector<std::string>{"<0.3", "0.3-0.8", "0.8-1.3", "1.3-1.8", "1.8-2.3", "2.3-2.8",
                                                      "2.8-3.3", "3.3-3.8", ">3.8"}));
  fm_pivot_numeric_group_t read = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_pivot_field_numeric_group(f.wb.handle, 0U, 0U, kAmountField, &present, &read), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(read.start, 0.3);
  EXPECT_EQ(read.end, 3.6);
  EXPECT_EQ(read.interval, 0.5);
}

TEST(CApiPivotNumericGroup, FractionalLabelsUseTheProfileDecimalSeparator) {
  FractionalFixture a;
  a.Build(kBaseA, "win-365-de_DE");
  ASSERT_EQ(a.Group(std::numeric_limits<double>::quiet_NaN(), std::numeric_limits<double>::quiet_NaN(), 0.5), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(a.ItemLabels(), (std::vector<std::string>{"<0,3", "0,3-0,8", "0,8-1,3", "1,3-1,8", "1,8-2,3", "2,3-2,8",
                                                      "2,8-3,3", "3,3-3,8", ">3,8"}));
  ExpectRows(a.Rows(),
             {{"0,3-0,8", 0.9},
              {"0,8-1,3", 2.1},
              {"1,3-1,8", 1.5},
              {"1,8-2,3", 3.9},
              {"2,3-2,8", 5.1},
              {"2,8-3,3", 3.0},
              {"3,3-3,8", 6.9}},
             "de c01");
  FractionalFixture b;
  b.Build(kBaseB, "win-365-fr_FR");
  ASSERT_EQ(b.Group(-2.25, 2.25, 0.75), 0) << fm_get_last_error_message();
  EXPECT_EQ(b.ItemLabels(), (std::vector<std::string>{"<-2,25", "-2,25--1,5", "-1,5--0,75", "-0,75-0", "0-0,75",
                                                      "0,75-1,5", "1,5-2,25", ">2,25"}));
}

TEST(CApiPivotNumericGroup, AGroupedReportThatCannotBePlacedIsBlocked) {
  // Region holds base A on the rows; grouping it puts its group field in its place.
  FractionalFixture f;
  f.Build(kBaseA, nullptr, /*region_rows=*/true);
  fm_pivot_numeric_group_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  spec.interval = 1.5;
  ASSERT_EQ(fm_sheet_set_pivot_field_numeric_group(f.wb.handle, 0U, 0U, kRegionField, &spec), 0)
      << fm_get_last_error_message();
  // Three ranges fill D1:E5; seven need D1:E9, which D7 blocks.
  ASSERT_EQ(f.Rows().size(), 3U);
  for (std::uint32_t r = 6U; r < 12U; ++r) {
    ASSERT_EQ(fm_sheet_set_text(f.wb.handle, 0U, r, 3U, "keep"), 0);
  }
  spec.interval = 0.5;
  ASSERT_EQ(fm_sheet_set_pivot_field_numeric_group(f.wb.handle, 0U, 0U, kRegionField, &spec), 0)
      << fm_get_last_error_message();
  fm_pivot_numeric_group_t read = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_pivot_field_numeric_group(f.wb.handle, 0U, 0U, kRegionField, &present, &read), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(read.interval, 0.5);
  fm_value_t anchor{};
  ASSERT_EQ(fm_sheet_get_value(f.wb.handle, 0U, 0U, 3U, &anchor), 0);
  ASSERT_EQ(anchor.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(anchor.u.error_code, static_cast<std::int32_t>(formulon::ErrorCode::Spill));
  fm_value_t kept{};
  ASSERT_EQ(fm_sheet_get_value(f.wb.handle, 0U, 6U, 3U, &kept), 0);
  ASSERT_EQ(kept.kind, FM_VALUE_KIND_TEXT);
  EXPECT_STREQ(kept.u.text, "keep");
}

}  // namespace
