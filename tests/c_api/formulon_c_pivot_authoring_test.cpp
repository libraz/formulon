// Stable C ABI PivotTable authoring and cache tests.

#include <iterator>
#include <limits>

#include "formulon_c_pivot_test_helpers.h"
#include "utils/date_time.h"

namespace {

// A NaN date-group bound leaves it to the data.
constexpr double kAutoBound = std::numeric_limits<double>::quiet_NaN();

}  // namespace

TEST(FormulonCApiPivot, CreatePivotFromScratch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  std::size_t cache_count = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_count(wb.handle, &cache_count), 0);
  EXPECT_EQ(cache_count, 1U);

  std::size_t pivot_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_count(wb.handle, 0, &pivot_count), 0);
  EXPECT_EQ(pivot_count, 1U);

  std::size_t records = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_record_count(wb.handle, cache_id, &records), 0);
  EXPECT_EQ(records, 4U);

  PivotCellsGuard projected;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_get_last_error_message();

  const std::vector<fm_pivot_cell_t> cells = CollectCells(projected.handle);
  ASSERT_FALSE(cells.empty());

  // North (regions 0, 2) sum = 400; South (regions 1, 3) sum = 600.
  // Grand total = 1000.
  bool saw_north = false;
  bool saw_south = false;
  bool saw_grand = false;
  for (const fm_pivot_cell_t& c : cells) {
    if (c.kind == FM_PIVOT_CELL_DATA && c.value.kind == FM_VAL_NUMBER) {
      if (c.value.u.number == 400.0) {
        saw_north = true;
      } else if (c.value.u.number == 600.0) {
        saw_south = true;
      }
    }
    if (c.kind == FM_PIVOT_CELL_GRAND_TOTAL && c.value.kind == FM_VAL_NUMBER && c.value.u.number == 1000.0) {
      saw_grand = true;
    }
  }
  EXPECT_TRUE(saw_north);
  EXPECT_TRUE(saw_south);
  EXPECT_TRUE(saw_grand);
}

TEST(FormulonCApiPivot, SaveRejectsAPivotCacheWithNoWorksheetSource) {
  // The failure this replaces was silent: the package saved, validated
  // against the schema, and round-tripped through our own reader, and only
  // Excel -- offering to repair it -- ever said otherwise. A host has no
  // way to discover that from the API, so the save has to say it.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache_id, /*present=*/0, nullptr, nullptr, nullptr),
            0)
      << fm_get_last_error_message();

  BufferGuard saved;
  EXPECT_NE(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << "a cache with no source must not save silently";
  EXPECT_NE(std::string(fm_get_last_error_message()).find("worksheet source"), std::string::npos)
      << "message must name what is missing: " << fm_get_last_error_message();

  // Restoring the source makes the same workbook savable again, so the
  // check is about the cache's state and not about some other damage.
  ASSERT_EQ(
      fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache_id, /*present=*/1, "A1:B5", "Sheet1", nullptr), 0);
  BufferGuard retry;
  EXPECT_EQ(fm_workbook_save(wb.handle, &retry.data, &retry.len), 0) << fm_get_last_error_message();
}

TEST(FormulonCApiPivot, SavedScratchPivotEmitsLocationRequiredDefaults) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "win-365-ja_JP"), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  ASSERT_EQ(fm_sheet_set_pivot_anchor(wb.handle, 0, pivot_idx, 0U, 3U, 5U, 2U), 0) << fm_get_last_error_message();

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  const std::vector<std::uint8_t> package(saved.data, saved.data + saved.len);
  const std::string pivot_xml = ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml");
  ASSERT_FALSE(pivot_xml.empty());
  EXPECT_NE(pivot_xml.find("<location ref=\"D1:E5\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>"),
            std::string::npos)
      << "xml=" << pivot_xml;
}

TEST(FormulonCApiPivot, SavedPivotLocationRefCoversTheProjectedGrid) {
  // `fm_sheet_add_pivot` installs a 1x1 placeholder span and nothing
  // revises it as fields are added, so a pivot built purely through the C
  // surface used to save a `ref` describing a single cell. Excel opens such
  // a file cleanly and recognises the pivot, then terminates the moment the
  // report is refreshed -- the defect is invisible to a round trip and has
  // to be pinned on the emitted bytes.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "win-365-ja_JP"), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  // Deliberately no `fm_sheet_set_pivot_anchor`: this is the path a host
  // takes when it never reasons about the report's extent at all.
  PivotCellsGuard projected;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_get_last_error_message();
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_EQ(fm_pivot_cells_get_bounds(projected.handle, &top, &left, &rows, &cols), 0);
  ASSERT_GT(rows, 1U) << "projection must be larger than the placeholder for this test to mean anything";
  ASSERT_GT(cols, 1U);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  const std::vector<std::uint8_t> package(saved.data, saved.data + saved.len);
  const std::string pivot_xml = ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml");
  ASSERT_FALSE(pivot_xml.empty());

  // Anchored at D1, projected 4 rows x 2 cols -> D1:E4.
  EXPECT_NE(pivot_xml.find("<location ref=\"D1:E4\""), std::string::npos) << "xml=" << pivot_xml;
  EXPECT_EQ(pivot_xml.find("<location ref=\"D1:D1\""), std::string::npos) << "xml=" << pivot_xml;
}

TEST(FormulonCApiPivot, PivotCacheMutationInvalidatesMemoisedLayout) {
  // fm_sheet_get_pivot_layout memoises the pivot's evaluated result. A cache
  // mutation must invalidate that memo so a re-layout reflects the change
  // instead of returning the stale projection.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  auto has_data_value = [](fm_pivot_cells_t* handle, double want) -> bool {
    for (const fm_pivot_cell_t& c : CollectCells(handle)) {
      if (c.kind == FM_PIVOT_CELL_DATA && c.value.kind == FM_VAL_NUMBER && c.value.u.number == want) {
        return true;
      }
    }
    return false;
  };

  // Baseline: North (records 0, 2 = 100 + 300) sums to 400.
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_get_last_error_message();
    EXPECT_TRUE(has_data_value(projected.handle, 400.0));
  }

  // Bump record 0's amount (field index 1) from 100 to 1100 (+1000).
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 0, 1, 1100.0), 0)
      << fm_get_last_error_message();

  // Re-layout must reflect the mutated cache: North is now 1100 + 300 = 1400,
  // and the stale 400 must be gone.
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_get_last_error_message();
    EXPECT_TRUE(has_data_value(projected.handle, 1400.0)) << "layout returned a stale memoised projection";
    EXPECT_FALSE(has_data_value(projected.handle, 400.0)) << "stale North value survived the cache mutation";
  }
}

TEST(FormulonCApiPivot, PivotMutationsInvalidateGetPivotDataFormula) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  // A1 is outside the report at D1, so this exercises the normal formula
  // dependency path rather than reading a projected pivot cell directly.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=GETPIVOTDATA(\"Sum of Amount\",D1,\"Region\",\"North\")"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  auto get_a1 = [&]() {
    fm_value_t value{};
    EXPECT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 0, &value), 0);
    return value;
  };
  fm_value_t value = get_a1();
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 400.0);

  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 0, 1, 1100.0), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  value = get_a1();
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 1400.0);

  fm_pivot_data_field_spec_t average_spec{};
  average_spec.name = "Sum of Amount";
  average_spec.field_index = 1U;
  average_spec.aggregation = FM_PIVOT_AGG_AVERAGE;
  average_spec.number_format = "";
  average_spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  average_spec.show_as_base_field = -1;
  average_spec.show_as_base_item = -1;
  ASSERT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot_idx, 0, &average_spec), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  value = get_a1();
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 700.0);

  // Keeping this formula's result warm before removal proves that the
  // successful lifecycle mutation itself dirties formula cells.
  ASSERT_EQ(fm_sheet_remove_pivot(wb.handle, 0, pivot_idx), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  value = get_a1();
  ASSERT_EQ(value.kind, FM_VAL_ERROR);
  EXPECT_EQ(value.u.error_code, static_cast<std::int32_t>(formulon::ErrorCode::Ref));
}

TEST(FormulonCApiPivot, LoadedPivotCacheMutationReprojectsAuthoredLocation) {
  WorkbookGuard original;
  ASSERT_EQ(fm_workbook_create_default(&original.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(original.handle, "win-365-ja_JP"), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(original.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(original.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_excel_profile_id(loaded.handle, "win-365-ja_JP"), 0);
  EXPECT_NE(SavedPivotXml(loaded.handle).find("<location ref=\"D1:E4\""), std::string::npos);

  // The loaded location is marked authored. Add a new shared item and record
  // so the evaluated report grows by one row; saving must then use the new
  // projected extent instead of preserving the stale authored ref.
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(loaded.handle, cache_id, 0, "West"), 0)
      << fm_get_last_error_message();
  std::size_t new_record = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_record(loaded.handle, cache_id, &new_record), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(loaded.handle, cache_id, new_record, 0, 2.0), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(loaded.handle, cache_id, new_record, 1, 500.0), 0)
      << fm_get_last_error_message();

  PivotCellsGuard projected;
  ASSERT_EQ(fm_sheet_get_pivot_layout(loaded.handle, 0, 0, &projected.handle), 0) << fm_get_last_error_message();
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_EQ(fm_pivot_cells_get_bounds(projected.handle, &top, &left, &rows, &cols), 0);
  EXPECT_EQ(top, 0U);
  EXPECT_EQ(left, 3U);
  EXPECT_EQ(rows, 5U);
  EXPECT_EQ(cols, 2U);

  const std::string pivot_xml = SavedPivotXml(loaded.handle);
  EXPECT_NE(pivot_xml.find("<location ref=\"D1:E5\""), std::string::npos) << pivot_xml;
}

TEST(FormulonCApiPivot, PivotCacheSharedItemsAcceptErrorValues) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &count), 0);
  EXPECT_EQ(count, 2U);

  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_error(wb.handle, cache_id, 0,
                                                                1),  // ErrorCode::Div0
            0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &count), 0);
  EXPECT_EQ(count, 3U);
}

TEST(FormulonCApiPivot, PivotCacheSharedItemValueReadsEachKind) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(wb.handle, cache_id, 0, 45292.0), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_bool(wb.handle, cache_id, 0, 1), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_blank(wb.handle, cache_id, 0), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_error(wb.handle, cache_id, 0, 1), 0);  // ErrorCode::Div0

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(wb.handle, cache_id, 0, 0, &v), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(v.kind, FM_VAL_TEXT);
  EXPECT_STREQ(v.u.text, "North");
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(wb.handle, cache_id, 0, 1, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_TEXT);
  EXPECT_STREQ(v.u.text, "South");
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(wb.handle, cache_id, 0, 2, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_EQ(v.u.number, 45292.0);
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(wb.handle, cache_id, 0, 3, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_BOOL);
  EXPECT_EQ(v.u.boolean, 1);
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(wb.handle, cache_id, 0, 4, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_BLANK);
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(wb.handle, cache_id, 0, 5, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_ERROR);
  EXPECT_EQ(v.u.error_code, 1);
}

TEST(FormulonCApiPivot, PivotCacheSharedItemValueRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  const auto invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  const auto null_ptr = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
  fm_value_t v{};
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(wb.handle, cache_id, 0, 2, &v), invalid);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(wb.handle, cache_id, 99, 0, &v), invalid);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(wb.handle, cache_id, 0, 0, nullptr), null_ptr);
  EXPECT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(nullptr, cache_id, 0, 0, &v), null_ptr);
}

TEST(FormulonCApiPivot, PivotCacheSharedItemIndexToleratesOutOfDomainNumbers) {
  // `BuildScratchPivot` leaves `cell_is_index` empty, so "Region" (field 0)
  // reads its numeric cells as indices into its two shared items. The record
  // setter accepts any finite double, so the layout below is the point where
  // an out-of-domain index would be narrowed. It must resolve to a blank label
  // rather than reading past the shared items or trapping the instance.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  const double bad_indices[] = {
      -1.0, 1e30,
      4.3e9,  // Past the wasm32 `size_t` range.
      2.0,    // One past the last shared item.
  };
  for (const double bad : bad_indices) {
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 0, 0, bad), 0)
        << fm_get_last_error_message();
    PivotCellsGuard projected;
    ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_get_last_error_message();
    EXPECT_FALSE(CollectCells(projected.handle).empty());
  }

  // A valid index still resolves after all of that.
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 0, 0, 0.0), 0)
      << fm_get_last_error_message();
  PivotCellsGuard projected;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_get_last_error_message();
  bool saw_north_total = false;
  for (const fm_pivot_cell_t& c : CollectCells(projected.handle)) {
    if (c.kind == FM_PIVOT_CELL_DATA && c.value.kind == FM_VAL_NUMBER && c.value.u.number == 400.0) {
      saw_north_total = true;
    }
  }
  EXPECT_TRUE(saw_north_total);
}

TEST(FormulonCApiPivot, PivotCacheWorksheetSourceRoundTripsThroughApi) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  int32_t present = -1;
  const char* ref = nullptr;
  const char* sheet = nullptr;
  const char* name = nullptr;
  // `BuildScratchPivot` installs a source because a cache without one
  // cannot be saved. Clear it here so the absent state this test is about
  // is established by the test rather than inherited from the helper.
  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache_id, 0, nullptr, nullptr, nullptr), 0);
  ASSERT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache_id, &present, &ref, &sheet, &name), 0);
  EXPECT_EQ(present, 0);
  EXPECT_STREQ(ref, "");
  EXPECT_STREQ(sheet, "");
  EXPECT_STREQ(name, "");

  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache_id, 1, "$A$1:$C$5", "Data", nullptr), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache_id, &present, &ref, &sheet, &name), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(ref, "$A$1:$C$5");
  EXPECT_STREQ(sheet, "Data");
  EXPECT_STREQ(name, "");

  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb.handle, cache_id, 0, "ignored", "ignored", "ignored"), 0);
  ASSERT_EQ(fm_workbook_get_pivot_cache_worksheet_source(wb.handle, cache_id, &present, &ref, &sheet, &name), 0);
  EXPECT_EQ(present, 0);
  EXPECT_STREQ(ref, "");
  EXPECT_STREQ(sheet, "");
  EXPECT_STREQ(name, "");
}

TEST(FormulonCApiPivot, PivotReportLayoutRoundTripsThroughApi) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  int32_t layout = FM_PIVOT_REPORT_LAYOUT_TABULAR;
  ASSERT_EQ(fm_sheet_get_pivot_report_layout(wb.handle, 0, pivot_idx, &layout), 0);
  EXPECT_EQ(layout, FM_PIVOT_REPORT_LAYOUT_COMPACT);

  ASSERT_EQ(fm_sheet_set_pivot_report_layout(wb.handle, 0, pivot_idx, FM_PIVOT_REPORT_LAYOUT_TABULAR), 0);
  ASSERT_EQ(fm_sheet_get_pivot_report_layout(wb.handle, 0, pivot_idx, &layout), 0);
  EXPECT_EQ(layout, FM_PIVOT_REPORT_LAYOUT_TABULAR);

  ASSERT_EQ(fm_sheet_set_pivot_report_layout(wb.handle, 0, pivot_idx, FM_PIVOT_REPORT_LAYOUT_OUTLINE), 0);
  ASSERT_EQ(fm_sheet_get_pivot_report_layout(wb.handle, 0, pivot_idx, &layout), 0);
  EXPECT_EQ(layout, FM_PIVOT_REPORT_LAYOUT_OUTLINE);

  // The setter accepts a raw int32_t so FFI callers can probe the complete
  // invalid domain without constructing an out-of-domain C enum.
  const std::int32_t invalid_layouts[] = {
      99,
      std::numeric_limits<std::int32_t>::min(),
      std::numeric_limits<std::int32_t>::max(),
  };
  for (const std::int32_t raw : invalid_layouts) {
    EXPECT_EQ(fm_sheet_set_pivot_report_layout(wb.handle, 0, pivot_idx, raw),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << raw;
  }
}

TEST(FormulonCApiPivot, ScalarEnumMutatorsRejectRawValuesWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  BufferGuard before;
  ASSERT_EQ(fm_workbook_save(wb.handle, &before.data, &before.len), 0) << fm_get_last_error_message();
  const std::vector<std::uint8_t> snapshot(before.data, before.data + before.len);
  const std::int32_t invalid[] = {99, std::numeric_limits<std::int32_t>::min(),
                                  std::numeric_limits<std::int32_t>::max()};
  const fm_status_t expected = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  for (const std::int32_t raw : invalid) {
    EXPECT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot_idx, 0, raw), expected) << raw;
    EXPECT_EQ(fm_sheet_add_pivot_field_subtotal_fn(wb.handle, 0, pivot_idx, 0, raw), expected) << raw;
    EXPECT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0, pivot_idx, 0, raw, nullptr), expected) << raw;
    EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_idx, 0, static_cast<std::uint32_t>(raw),
                                                  FM_PIVOT_CALENDAR_GREGORIAN, 1, kAutoBound, kAutoBound),
              expected)
        << raw;
    EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_idx, 0, FM_PIVOT_PERIOD_YEARS, raw, 1, kAutoBound,
                                                  kAutoBound),
              expected)
        << raw;
  }

  BufferGuard after;
  ASSERT_EQ(fm_workbook_save(wb.handle, &after.data, &after.len), 0) << fm_get_last_error_message();
  EXPECT_EQ(std::vector<std::uint8_t>(after.data, after.data + after.len), snapshot);
}

namespace {

// Row-label texts of the projected layout, in row order.
std::vector<std::string> RowLabels(fm_workbook_t* wb, std::size_t pivot_idx) {
  PivotCellsGuard projected;
  EXPECT_EQ(fm_sheet_get_pivot_layout(wb, 0, pivot_idx, &projected.handle), 0) << fm_get_last_error_message();
  std::vector<std::string> labels;
  for (const fm_pivot_cell_t& cell : CollectCells(projected.handle)) {
    if (cell.kind == FM_PIVOT_CELL_ROW_LABEL && cell.value.kind == FM_VAL_TEXT) {
      labels.emplace_back(cell.value.u.text);
    }
  }
  return labels;
}

}  // namespace

TEST(FormulonCApiPivot, ManualSortKeepsItemOrderThroughSaveAndReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "win-365-ja_JP"), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  // Shared items are North, South; the items are added in the opposite order.
  ASSERT_EQ(fm_sheet_clear_pivot_field_items(wb.handle, 0, pivot_idx, 0), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item(wb.handle, 0, pivot_idx, 0, "South", 1), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item(wb.handle, 0, pivot_idx, 0, "North", 1), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0, pivot_idx, 0, FM_PIVOT_SORT_ASCENDING, nullptr), 0);
  EXPECT_EQ(RowLabels(wb.handle, pivot_idx), (std::vector<std::string>{"North", "South"}));

  ASSERT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0, pivot_idx, 0, FM_PIVOT_SORT_MANUAL, nullptr), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(RowLabels(wb.handle, pivot_idx), (std::vector<std::string>{"South", "North"}));

  const std::string xml = SavedPivotXml(wb.handle);
  EXPECT_EQ(xml.find("sortType"), std::string::npos) << xml;
  EXPECT_NE(xml.find("<item x=\"1\"/><item x=\"0\"/>"), std::string::npos) << xml;

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_excel_profile_id(loaded.handle, "win-365-ja_JP"), 0);
  EXPECT_EQ(RowLabels(loaded.handle, 0), (std::vector<std::string>{"South", "North"}));

  // Leaving manual order returns to label order and spells it out.
  ASSERT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0, pivot_idx, 0, FM_PIVOT_SORT_ASCENDING, nullptr), 0);
  EXPECT_EQ(RowLabels(wb.handle, pivot_idx), (std::vector<std::string>{"North", "South"}));
  EXPECT_NE(SavedPivotXml(wb.handle).find("sortType=\"ascending\""), std::string::npos);
}

TEST(FormulonCApiPivot, ManualSortRejectsAByFieldKeyWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  BufferGuard before;
  ASSERT_EQ(fm_workbook_save(wb.handle, &before.data, &before.len), 0) << fm_get_last_error_message();
  const std::vector<std::uint8_t> snapshot(before.data, before.data + before.len);
  const fm_status_t expected = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0, pivot_idx, 0, FM_PIVOT_SORT_MANUAL, "Sum of Amount"), expected);
  EXPECT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0, pivot_idx, 0, 3, nullptr), expected);
  EXPECT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0, pivot_idx, 0, -1, nullptr), expected);
  // An empty key is the same as no key.
  EXPECT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0, pivot_idx, 0, FM_PIVOT_SORT_MANUAL, ""), 0);

  BufferGuard after;
  ASSERT_EQ(fm_sheet_set_pivot_field_sort(wb.handle, 0, pivot_idx, 0, FM_PIVOT_SORT_ASCENDING, nullptr), 0);
  ASSERT_EQ(fm_workbook_save(wb.handle, &after.data, &after.len), 0) << fm_get_last_error_message();
  EXPECT_EQ(std::vector<std::uint8_t>(after.data, after.data + after.len), snapshot);
}

namespace {

// Date / Amount records (2024-01-01, 2025-01-01, 2024-06-01, 2025-06-01 with
// 100..400), shown by a pivot with Date on rows and the sum of Amount.
void BuildDatePivot(fm_workbook_t* wb, bool date1904, std::uint32_t* out_cache_id, std::size_t* out_pivot_idx,
                    std::size_t* out_date_field) {
  std::uint32_t& cache_id = *out_cache_id;
  std::size_t& pivot_idx = *out_pivot_idx;
  const double dates[] = {
      formulon::date_time::serial_from_ymd(2024, 1U, 1U, date1904),
      formulon::date_time::serial_from_ymd(2025, 1U, 1U, date1904),
      formulon::date_time::serial_from_ymd(2024, 6U, 1U, date1904),
      formulon::date_time::serial_from_ymd(2025, 6U, 1U, date1904),
  };
  const double amounts[] = {100.0, 200.0, 300.0, 400.0};

  ASSERT_EQ(fm_workbook_add_pivot_cache(wb, 0, &cache_id), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb, cache_id, 1, "A1:B5", "Sheet1", nullptr), 0)
      << fm_get_last_error_message();
  std::size_t date_cache_field = 99;
  std::size_t amount_cache_field = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb, cache_id, "Date", &date_cache_field), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb, cache_id, "Amount", &amount_cache_field), 0)
      << fm_get_last_error_message();
  for (std::size_t record = 0; record < std::size(dates); ++record) {
    std::size_t record_idx = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb, cache_id, &record_idx), 0) << fm_get_last_error_message();
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb, cache_id, record_idx, date_cache_field, dates[record]), 0)
        << fm_get_last_error_message();
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb, cache_id, record_idx, amount_cache_field, amounts[record]),
              0)
        << fm_get_last_error_message();
  }

  ASSERT_EQ(fm_sheet_add_pivot(wb, 0, "PT", cache_id, 0U, 3U, &pivot_idx), 0) << fm_get_last_error_message();
  fm_pivot_field_spec_t date_spec{};
  date_spec.source_name = "Date";
  date_spec.custom_name = "";
  date_spec.axis = FM_PIVOT_AXIS_ROW;
  date_spec.subtotal_top = 0;
  date_spec.number_format = "";
  std::size_t& date_field = *out_date_field;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb, 0, pivot_idx, &date_spec, &date_field), 0) << fm_get_last_error_message();
  const std::uint32_t row_order[] = {static_cast<std::uint32_t>(date_field)};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb, 0, pivot_idx, row_order, 1U), 0) << fm_get_last_error_message();
  fm_pivot_field_spec_t amount_spec{};
  amount_spec.source_name = "Amount";
  amount_spec.custom_name = "";
  amount_spec.axis = FM_PIVOT_AXIS_VALUE;
  amount_spec.subtotal_top = 0;
  amount_spec.number_format = "";
  std::size_t amount_field = 99;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb, 0, pivot_idx, &amount_spec, &amount_field), 0) << fm_get_last_error_message();
  fm_pivot_data_field_spec_t data_spec{};
  data_spec.name = "Sum of Amount";
  data_spec.field_index = static_cast<std::uint32_t>(amount_field);
  data_spec.aggregation = FM_PIVOT_AGG_SUM;
  data_spec.number_format = "";
  data_spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  data_spec.show_as_base_field = -1;
  data_spec.show_as_base_item = -1;
  std::size_t data_field = 99;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb, 0, pivot_idx, &data_spec, &data_field), 0) << fm_get_last_error_message();
}

}  // namespace

TEST(FormulonCApiPivot, DateGroupRejectsBadArgumentsWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  std::size_t date_field = 0;
  BuildDatePivot(wb.handle, /*date1904=*/false, &cache_id, &pivot_idx, &date_field);
  if (HasFatalFailure()) {
    return;
  }

  BufferGuard before;
  ASSERT_EQ(fm_workbook_save(wb.handle, &before.data, &before.len), 0) << fm_get_last_error_message();
  const std::vector<std::uint8_t> snapshot(before.data, before.data + before.len);
  const fm_status_t expected = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  const double jan1 = formulon::date_time::serial_from_ymd(2024, 1U, 1U);

  // No period, or a bit past the years.
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_idx, date_field, 0U, FM_PIVOT_CALENDAR_GREGORIAN, 1,
                                                kAutoBound, kAutoBound),
            expected);
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_idx, date_field, FM_PIVOT_PERIOD_YEARS | 128U,
                                                FM_PIVOT_CALENDAR_GREGORIAN, 1, kAutoBound, kAutoBound),
            expected);
  // A zero interval, and an interval with a period besides days.
  EXPECT_EQ(
      fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_idx, date_field, FM_PIVOT_PERIOD_DAYS,
                                          FM_PIVOT_CALENDAR_GREGORIAN, /*interval_days=*/0, kAutoBound, kAutoBound),
      expected);
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(
                wb.handle, 0, pivot_idx, date_field, FM_PIVOT_PERIOD_DAYS | FM_PIVOT_PERIOD_MONTHS,
                FM_PIVOT_CALENDAR_GREGORIAN, /*interval_days=*/7, kAutoBound, kAutoBound),
            expected);
  // A start after the end, compared before flooring.
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_idx, date_field, FM_PIVOT_PERIOD_DAYS,
                                                FM_PIVOT_CALENDAR_GREGORIAN, 7, jan1 + 1.0, jan1),
            expected);
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_idx, date_field, FM_PIVOT_PERIOD_DAYS,
                                                FM_PIVOT_CALENDAR_GREGORIAN, 7, jan1 + 0.9, jan1 + 0.1),
            expected);
  // A field holding anything but dates.
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_text(wb.handle, cache_id, 0, 0, "n/a"), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_idx, date_field, FM_PIVOT_PERIOD_YEARS,
                                                FM_PIVOT_CALENDAR_GREGORIAN, 1, kAutoBound, kAutoBound),
            expected);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 0, 0, jan1), 0)
      << fm_get_last_error_message();

  BufferGuard after;
  ASSERT_EQ(fm_workbook_save(wb.handle, &after.data, &after.len), 0) << fm_get_last_error_message();
  EXPECT_EQ(std::vector<std::uint8_t>(after.data, after.data + after.len), snapshot);
}

TEST(FormulonCApiPivot, DateGroupRejectsInvalidSerialBoundsWithoutMutation) {
  const fm_status_t expected = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  for (const bool date1904 : {false, true}) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    if (date1904) {
      wb.handle->workbook().set_date1904(true);
    }
    std::uint32_t cache_id = 0;
    std::size_t pivot_idx = 0;
    std::size_t date_field = 0;
    BuildDatePivot(wb.handle, date1904, &cache_id, &pivot_idx, &date_field);
    if (HasFatalFailure()) {
      return;
    }
    ASSERT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_idx, date_field, FM_PIVOT_PERIOD_YEARS,
                                                  FM_PIVOT_CALENDAR_GREGORIAN, 1, kAutoBound, kAutoBound),
              0)
        << fm_get_last_error_message();

    auto has_data_value = [&](fm_pivot_cells_t* cells, double want) {
      for (const fm_pivot_cell_t& cell : CollectCells(cells)) {
        if (cell.kind == FM_PIVOT_CELL_DATA && cell.value.kind == FM_VAL_NUMBER && cell.value.u.number == want) {
          return true;
        }
      }
      return false;
    };
    {
      PivotCellsGuard projected;
      ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0)
          << fm_get_last_error_message();
      EXPECT_TRUE(LayoutHasText(wb.handle, pivot_idx, "2024"));
      EXPECT_TRUE(LayoutHasText(wb.handle, pivot_idx, "2025"));
      EXPECT_TRUE(has_data_value(projected.handle, 400.0));
      EXPECT_TRUE(has_data_value(projected.handle, 600.0));
    }

    BufferGuard before;
    ASSERT_EQ(fm_workbook_save(wb.handle, &before.data, &before.len), 0) << fm_get_last_error_message();
    const std::vector<std::uint8_t> snapshot(before.data, before.data + before.len);
    const double last_valid = formulon::date_time::serial_from_ymd(9999, 12U, 31U, date1904);
    const double invalid[] = {
        -2.0,
        -0.1,
        std::numeric_limits<double>::infinity(),
        -std::numeric_limits<double>::infinity(),
        last_valid + 1.0,
        std::numeric_limits<double>::max(),
    };
    for (const double raw : invalid) {
      EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_idx, date_field, FM_PIVOT_PERIOD_DAYS,
                                                    FM_PIVOT_CALENDAR_GREGORIAN, 7, raw, 70.0),
                expected)
          << "invalid start=" << raw << " date1904=" << date1904;
      EXPECT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_idx, date_field, FM_PIVOT_PERIOD_DAYS,
                                                    FM_PIVOT_CALENDAR_GREGORIAN, 7, 0.0, raw),
                expected)
          << "invalid end=" << raw << " date1904=" << date1904;
    }

    // Force the memoised result through the cache mutation path before
    // re-evaluating. The successful no-op write also proves the rejected
    // setters did not leave the field in an invalid state.
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 0, 1, 100.0), 0)
        << fm_get_last_error_message();
    {
      PivotCellsGuard projected;
      ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0)
          << fm_get_last_error_message();
      EXPECT_TRUE(LayoutHasText(wb.handle, pivot_idx, "2024"));
      EXPECT_TRUE(LayoutHasText(wb.handle, pivot_idx, "2025"));
      EXPECT_TRUE(has_data_value(projected.handle, 400.0));
      EXPECT_TRUE(has_data_value(projected.handle, 600.0));
    }

    BufferGuard after;
    ASSERT_EQ(fm_workbook_save(wb.handle, &after.data, &after.len), 0) << fm_get_last_error_message();
    EXPECT_EQ(std::vector<std::uint8_t>(after.data, after.data + after.len), snapshot);
  }
}

TEST(FormulonCApiPivot, StructEnumMutatorsRejectRawValuesWithoutMutation) {
  // The struct-taking twins of the scalar mutators above must enforce the
  // same domains. Without that, an out-of-domain `axis` collapses onto
  // `PivotAxis::Row` and an out-of-domain `show_as` onto `Normal`, and
  // the call reports success on a pivot that means something else.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  BufferGuard before;
  ASSERT_EQ(fm_workbook_save(wb.handle, &before.data, &before.len), 0) << fm_get_last_error_message();
  const std::vector<std::uint8_t> snapshot(before.data, before.data + before.len);
  const fm_status_t expected = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

  fm_pivot_field_spec_t field_spec{};
  field_spec.source_name = "Region";
  field_spec.axis = RawEnumValue<fm_pivot_axis_t>(99);
  std::size_t field_idx = 99U;
  EXPECT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot_idx, &field_spec, &field_idx), expected);

  fm_pivot_data_field_spec_t df_spec{};
  df_spec.name = "Total";
  df_spec.field_index = 1;
  df_spec.aggregation = RawEnumValue<fm_pivot_aggregation_t>(99);
  std::size_t df_idx = 99U;
  EXPECT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot_idx, &df_spec, &df_idx), expected);
  EXPECT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot_idx, 0, &df_spec), expected);

  df_spec.aggregation = FM_PIVOT_AGG_SUM;
  df_spec.show_as = RawEnumValue<fm_pivot_show_values_as_t>(99);
  EXPECT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0, pivot_idx, &df_spec, &df_idx), expected);
  EXPECT_EQ(fm_sheet_set_pivot_data_field(wb.handle, 0, pivot_idx, 0, &df_spec), expected);

  BufferGuard after;
  ASSERT_EQ(fm_workbook_save(wb.handle, &after.data, &after.len), 0) << fm_get_last_error_message();
  EXPECT_EQ(std::vector<std::uint8_t>(after.data, after.data + after.len), snapshot);
}

TEST(FormulonCApiPivot, PivotProjectionUsesWorkbookLocaleAndReportLayout) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "win-365-ja_JP"), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  ASSERT_EQ(fm_sheet_set_pivot_report_layout(wb.handle, 0, pivot_idx, FM_PIVOT_REPORT_LAYOUT_TABULAR), 0);

  PivotCellsGuard projected;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_get_last_error_message();
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_EQ(fm_pivot_cells_get_bounds(projected.handle, &top, &left, &rows, &cols), 0);
  EXPECT_EQ(cols, 2U);  // Tabular exposes its two row-field columns.

  bool saw_japanese_total = false;
  const std::size_t cell_count = count_of(fm_pivot_cells_get_cell_count, projected.handle);
  for (std::size_t i = 0; i < cell_count; ++i) {
    fm_pivot_cell_t cell{};
    ASSERT_EQ(fm_pivot_cells_get_cell_at(projected.handle, i, &cell), 0);
    if (cell.value.kind == FM_VAL_TEXT && cell.value.u.text != nullptr &&
        std::string_view(cell.value.u.text) == "総計") {
      saw_japanese_total = true;
    }
  }
  EXPECT_TRUE(saw_japanese_total);
}

TEST(FormulonCApiPivot, PivotCacheErrorSettersRejectInvalidErrorCodes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_error(wb.handle, cache_id, 0, -1),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_error(wb.handle, cache_id, 0, 999),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_error(wb.handle, cache_id, 0, 1, -1),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_error(wb.handle, cache_id, 0, 1, 999),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiPivot, PivotCacheRecordSettersRejectOutOfRangeFieldIndex) {
  // BuildScratchPivot declares two cache fields (Region=0, Amount=1), so
  // field_idx 2 is the first out-of-range value. A large / wrapping field_idx
  // must be rejected before `grow_record_cells` resizes `field_idx + 1`
  // cells — otherwise SIZE_MAX wraps to 0 and the subsequent cell write lands
  // out of bounds.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();

  const auto kInvalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  const std::size_t kFieldCount = 2;
  const std::size_t kMax = static_cast<std::size_t>(-1);

  // In-range write still succeeds (guards against an over-tight bound).
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 0, kFieldCount - 1, 1.0), 0);

  // field_idx == field_count and beyond are rejected on every setter.
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 0, kFieldCount, 1.0), kInvalid);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 0, kMax, 1.0), kInvalid);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_text(wb.handle, cache_id, 0, kMax, "x"), kInvalid);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_bool(wb.handle, cache_id, 0, kMax, 1), kInvalid);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_blank(wb.handle, cache_id, 0, kMax), kInvalid);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_error(wb.handle, cache_id, 0, kMax, 0), kInvalid);
}

TEST(FormulonCApiPivot, FieldAxisKeepsRowAndColumnOrdersInStep) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "win-365-ja_JP"), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_get_last_error_message();
  const std::string kRegionRows = "<rowFields count=\"1\"><field x=\"0\"/></rowFields>";
  const std::string kRegionCols = "<colFields count=\"1\"><field x=\"0\"/></colFields>";
  // The helper adds Region on the row axis and then restates the same order.
  std::string xml = SavedPivotXml(wb.handle);
  EXPECT_NE(xml.find(kRegionRows), std::string::npos) << xml;

  // A field added on the row axis is placed without an explicit order.
  ASSERT_EQ(fm_sheet_clear_pivot_fields(wb.handle, 0, pivot_idx), 0);
  fm_pivot_field_spec_t region_spec{};
  region_spec.source_name = "Region";
  region_spec.axis = FM_PIVOT_AXIS_ROW;
  region_spec.subtotal_top = 1;
  std::size_t region = 99;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot_idx, &region_spec, &region), 0) << fm_get_last_error_message();
  fm_pivot_field_spec_t amount_spec{};
  amount_spec.source_name = "Amount";
  amount_spec.axis = FM_PIVOT_AXIS_VALUE;
  amount_spec.subtotal_top = 1;
  std::size_t amount = 99;
  ASSERT_EQ(fm_sheet_add_pivot_field(wb.handle, 0, pivot_idx, &amount_spec, &amount), 0) << fm_get_last_error_message();
  EXPECT_EQ(RowLabels(wb.handle, pivot_idx), (std::vector<std::string>{"North", "South"}));

  // Moving the field, twice, lists it once in the new order only.
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot_idx, region, FM_PIVOT_AXIS_COL), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot_idx, region, FM_PIVOT_AXIS_COL), 0);
  xml = SavedPivotXml(wb.handle);
  EXPECT_EQ(xml.find("<rowFields"), std::string::npos) << xml;
  EXPECT_NE(xml.find(kRegionCols), std::string::npos) << xml;

  // Off both axes, then back onto the rows.
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot_idx, region, FM_PIVOT_AXIS_VALUE), 0);
  xml = SavedPivotXml(wb.handle);
  EXPECT_EQ(xml.find("<rowFields"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("<colFields"), std::string::npos) << xml;
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot_idx, region, FM_PIVOT_AXIS_ROW), 0);
  EXPECT_NE(SavedPivotXml(wb.handle).find(kRegionRows), std::string::npos);

  // Setting an order moves its fields onto that axis and drops the rest.
  const auto region_index = static_cast<std::uint32_t>(region);
  ASSERT_EQ(fm_sheet_set_pivot_col_field_order(wb.handle, 0, pivot_idx, &region_index, 1U), 0);
  xml = SavedPivotXml(wb.handle);
  EXPECT_EQ(xml.find("<rowFields"), std::string::npos) << xml;
  EXPECT_NE(xml.find(kRegionCols), std::string::npos) << xml;
  ASSERT_EQ(fm_sheet_set_pivot_col_field_order(wb.handle, 0, pivot_idx, nullptr, 0U), 0);
  xml = SavedPivotXml(wb.handle);
  EXPECT_EQ(xml.find("<colFields"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("axis=\"axisCol\""), std::string::npos) << xml;

  const std::uint32_t twice[] = {region_index, region_index};
  EXPECT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0, pivot_idx, twice, 2U),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(SavedPivotXml(wb.handle).find("<rowFields"), std::string::npos);
}

// ---------------------------------------------------------------------------
// Date cache items
// ---------------------------------------------------------------------------

namespace {

constexpr fm_status_t kInvalidArgumentStatus = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNullPointerStatus = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);

// A cache of "When" (row axis) and "Amt" (value axis) with one pivot, so it saves.
void BuildDatePivot(fm_workbook_t* wb, std::uint32_t* out_cache_id) {
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb, 0, out_cache_id), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_pivot_cache_worksheet_source(wb, *out_cache_id, 1, "A1:B5", "Sheet1", nullptr), 0);
  for (const char* name : {"When", "Amt"}) {
    std::size_t idx = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb, *out_cache_id, name, &idx), 0) << fm_get_last_error_message();
  }
  std::size_t pivot = 99;
  ASSERT_EQ(fm_sheet_add_pivot(wb, 0, "PT", *out_cache_id, 0U, 3U, &pivot), 0) << fm_get_last_error_message();
  const char* const names[] = {"When", "Amt"};
  const fm_pivot_axis_t axes[] = {FM_PIVOT_AXIS_ROW, FM_PIVOT_AXIS_VALUE};
  for (std::size_t i = 0; i < 2; ++i) {
    fm_pivot_field_spec_t spec{};
    spec.source_name = names[i];
    spec.custom_name = "";
    spec.axis = axes[i];
    spec.number_format = "";
    std::size_t field = 99;
    ASSERT_EQ(fm_sheet_add_pivot_field(wb, 0, pivot, &spec, &field), 0) << fm_get_last_error_message();
  }
  const std::uint32_t rows[] = {0U};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb, 0, pivot, rows, 1U), 0);
  fm_pivot_data_field_spec_t df{};
  df.name = "Sum of Amt";
  df.field_index = 1U;
  df.aggregation = FM_PIVOT_AGG_SUM;
  df.number_format = "";
  df.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  df.show_as_base_field = -1;
  df.show_as_base_item = -1;
  std::size_t df_idx = 99;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb, 0, pivot, &df, &df_idx), 0) << fm_get_last_error_message();
}

std::string SavedCachePart(fm_workbook_t* wb, const char* path) {
  BufferGuard saved;
  EXPECT_EQ(fm_workbook_save(wb, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  return ExtractZipEntry(std::vector<std::uint8_t>(saved.data, saved.data + saved.len), path);
}

}  // namespace

TEST(FormulonCApiPivot, PivotCacheDateSharedItemsReadBackAsDates) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  BuildDatePivot(wb.handle, &cache_id);

  // A plain number before the date, one after it, and a text item.
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(wb.handle, cache_id, 0, 45292.0), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(wb.handle, cache_id, 0, 45292.5), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(wb.handle, cache_id, 0, 45293.0), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache_id, 0, "x"), 0);

  const std::int32_t expected[] = {0, 1, 0, 0};
  for (std::size_t i = 0; i < 4; ++i) {
    std::int32_t is_date = -1;
    ASSERT_EQ(fm_workbook_is_pivot_cache_field_shared_item_date(wb.handle, cache_id, 0, i, &is_date), 0)
        << fm_get_last_error_message();
    EXPECT_EQ(is_date, expected[i]) << i;
  }
  fm_value_t v{};
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(wb.handle, cache_id, 0, 1, &v), 0);
  ASSERT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_EQ(v.u.number, 45292.5);

  std::int32_t is_date = -1;
  EXPECT_EQ(fm_workbook_is_pivot_cache_field_shared_item_date(wb.handle, cache_id, 0, 4, &is_date),
            kInvalidArgumentStatus);
  EXPECT_EQ(fm_workbook_is_pivot_cache_field_shared_item_date(wb.handle, cache_id, 9, 0, &is_date),
            kInvalidArgumentStatus);
  EXPECT_EQ(fm_workbook_is_pivot_cache_field_shared_item_date(wb.handle, 9999U, 0, 0, &is_date),
            kInvalidArgumentStatus);
  EXPECT_EQ(fm_workbook_is_pivot_cache_field_shared_item_date(wb.handle, cache_id, 0, 0, nullptr), kNullPointerStatus);
  EXPECT_EQ(fm_workbook_is_pivot_cache_field_shared_item_date(nullptr, cache_id, 0, 0, &is_date), kNullPointerStatus);
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(nullptr, cache_id, 0, 45292.0), kNullPointerStatus);
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(wb.handle, cache_id, 9, 45292.0),
            kInvalidArgumentStatus);
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(wb.handle, 9999U, 0, 45292.0), kInvalidArgumentStatus);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_date(nullptr, cache_id, 0, 0, 45292.0), kNullPointerStatus);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_date(wb.handle, cache_id, 0, 0, 45292.0), kInvalidArgumentStatus);
  size_t record = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache_id, &record), 0);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_date(wb.handle, cache_id, record, 9, 45292.0), kInvalidArgumentStatus);
  EXPECT_EQ(fm_workbook_set_pivot_cache_record_date(wb.handle, cache_id, record + 1, 0, 45292.0),
            kInvalidArgumentStatus);
  size_t count = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &count), 0);
  EXPECT_EQ(count, 4U);
}

TEST(FormulonCApiPivot, PivotCacheDateSerialBoundsAreValidated) {
  struct System {
    int32_t date1904;
    double last_day;       // largest accepted floor
    double just_past_end;  // smallest rejected value
    bool has_sixty;
  };
  const System systems[] = {{0, 2958465.99, 2958466.0, true}, {1, 2957003.5, 2957004.0, false}};
  for (const System& sys : systems) {
    SCOPED_TRACE(sys.date1904);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ASSERT_EQ(fm_workbook_set_date1904(wb.handle, sys.date1904), 0);
    std::uint32_t cache_id = 0;
    BuildDatePivot(wb.handle, &cache_id);
    size_t record = 0;
    ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache_id, &record), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, record, 0, 7.0), 0);

    const double bad[] = {std::numeric_limits<double>::quiet_NaN(),
                          std::numeric_limits<double>::infinity(),
                          -std::numeric_limits<double>::infinity(),
                          -1.0,
                          -0.5,
                          sys.just_past_end,
                          sys.just_past_end + 1.0};
    for (const double serial : bad) {
      SCOPED_TRACE(serial);
      EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(wb.handle, cache_id, 0, serial),
                kInvalidArgumentStatus);
      EXPECT_EQ(fm_workbook_set_pivot_cache_record_date(wb.handle, cache_id, record, 0, serial),
                kInvalidArgumentStatus);
    }
    // Nothing was mutated by the rejections.
    size_t count = 99;
    ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &count), 0);
    EXPECT_EQ(count, 0U);
    // The axis field's items are derived at save, so the 7 shows up in the definition.
    const std::string def = SavedCachePart(wb.handle, "xl/pivotCache/pivotCacheDefinition1.xml");
    const std::string records = SavedCachePart(wb.handle, "xl/pivotCache/pivotCacheRecords1.xml");
    EXPECT_NE(def.find("<n v=\"7\"/>"), std::string::npos) << def;
    EXPECT_EQ(def.find("<d "), std::string::npos) << def;
    EXPECT_EQ(records.find("<d "), std::string::npos) << records;

    std::vector<double> good = {0.0, -0.0, 0.5, 1.0, 59.0, 61.0, sys.last_day};
    if (sys.has_sixty) {
      good.push_back(60.0);
    }
    for (const double serial : good) {
      SCOPED_TRACE(serial);
      EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(wb.handle, cache_id, 0, serial), 0)
          << fm_get_last_error_message();
      EXPECT_EQ(fm_workbook_set_pivot_cache_record_date(wb.handle, cache_id, record, 0, serial), 0)
          << fm_get_last_error_message();
    }
    ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &count), 0);
    EXPECT_EQ(count, good.size());
  }
}

TEST(FormulonCApiPivot, PivotCacheRecordDateOnSharedFieldIsInline) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  BuildDatePivot(wb.handle, &cache_id);

  // Record 0 carries the date before the field has any shared item.
  size_t r0 = 0;
  size_t r1 = 0;
  size_t r2 = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache_id, &r0), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_date(wb.handle, cache_id, r0, 0, 45292.0), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, r0, 1, 1.0), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(wb.handle, cache_id, 0, 10.0), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(wb.handle, cache_id, 0, 20.0), 0);
  // Record 1 indexes a shared item; record 2 sets a date after the items exist.
  ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache_id, &r1), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, r1, 0, 1.0), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, r1, 1, 2.0), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache_id, &r2), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, r2, 0, 0.0), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_date(wb.handle, cache_id, r2, 0, 45296.0), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, r2, 1, 3.0), 0);

  const std::string records = SavedCachePart(wb.handle, "xl/pivotCache/pivotCacheRecords1.xml");
  EXPECT_NE(records.find("<r><d v=\"2024-01-01T00:00:00\"/><n v=\"1\"/></r>"), std::string::npos) << records;
  EXPECT_NE(records.find("<r><x v=\"1\"/><n v=\"2\"/></r>"), std::string::npos) << records;
  EXPECT_NE(records.find("<r><d v=\"2024-01-05T00:00:00\"/><n v=\"3\"/></r>"), std::string::npos) << records;
}

TEST(FormulonCApiPivot, PivotCacheDateItemsSaveWithExcelHints) {
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    std::uint32_t cache_id = 0;
    BuildDatePivot(wb.handle, &cache_id);
    for (const double serial : {45300.0, 45292.0, 45296.0}) {
      ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(wb.handle, cache_id, 0, serial), 0);
    }
    const std::string def = SavedCachePart(wb.handle, "xl/pivotCache/pivotCacheDefinition1.xml");
    EXPECT_NE(def.find("<sharedItems containsSemiMixedTypes=\"0\" containsNonDate=\"0\" containsDate=\"1\" "
                       "containsString=\"0\" minDate=\"2024-01-01T00:00:00\" maxDate=\"2024-01-10T00:00:00\" "
                       "count=\"3\"><d v=\"2024-01-09T00:00:00\"/><d v=\"2024-01-01T00:00:00\"/>"
                       "<d v=\"2024-01-05T00:00:00\"/></sharedItems>"),
              std::string::npos)
        << def;
  }
  {
    // Dates beside plain numbers: no date bounds, and the field is mixed.
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    std::uint32_t cache_id = 0;
    BuildDatePivot(wb.handle, &cache_id);
    ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(wb.handle, cache_id, 0, 45292.0), 0);
    ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(wb.handle, cache_id, 0, 2.5), 0);
    const std::string def = SavedCachePart(wb.handle, "xl/pivotCache/pivotCacheDefinition1.xml");
    EXPECT_NE(def.find("<sharedItems containsSemiMixedTypes=\"0\" containsDate=\"1\" containsString=\"0\" "
                       "containsMixedTypes=\"1\" count=\"2\"><d v=\"2024-01-01T00:00:00\"/><n v=\"2.5\"/>"),
              std::string::npos)
        << def;
  }
}
