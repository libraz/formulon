// Stable C ABI PivotTable authoring and cache tests.

#include <iterator>
#include <limits>

#include "formulon_c_pivot_test_helpers.h"
#include "utils/date_time.h"

TEST(FormulonCApiPivot, CreatePivotFromScratch) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  std::size_t cache_count = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_count(wb.handle, &cache_count), 0);
  EXPECT_EQ(cache_count, 1U);

  std::size_t pivot_count = 0;
  ASSERT_EQ(fm_workbook_pivot_count(wb.handle, 0, &pivot_count), 0);
  EXPECT_EQ(pivot_count, 1U);

  std::size_t records = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_record_count(wb.handle, cache_id, &records), 0);
  EXPECT_EQ(records, 4U);

  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();

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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  ASSERT_EQ(fm_workbook_pivot_cache_set_worksheet_source(wb.handle, cache_id, /*present=*/0, nullptr, nullptr, nullptr),
            0)
      << fm_last_error_message();

  BufferGuard saved;
  EXPECT_NE(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << "a cache with no source must not save silently";
  EXPECT_NE(std::string(fm_last_error_message()).find("worksheet source"), std::string::npos)
      << "message must name what is missing: " << fm_last_error_message();

  // Restoring the source makes the same workbook savable again, so the
  // check is about the cache's state and not about some other damage.
  ASSERT_EQ(
      fm_workbook_pivot_cache_set_worksheet_source(wb.handle, cache_id, /*present=*/1, "A1:B5", "Sheet1", nullptr), 0);
  BufferGuard retry;
  EXPECT_EQ(fm_workbook_save(wb.handle, &retry.data, &retry.len), 0) << fm_last_error_message();
}

TEST(FormulonCApiPivot, SavedScratchPivotEmitsLocationRequiredDefaults) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  ASSERT_EQ(fm_workbook_pivot_set_anchor(wb.handle, 0, pivot_idx, 0U, 3U, 5U, 2U), 0) << fm_last_error_message();

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> package(saved.data, saved.data + saved.len);
  const std::string pivot_xml = ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml");
  ASSERT_FALSE(pivot_xml.empty());
  EXPECT_NE(pivot_xml.find("<location ref=\"D1:E5\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>"),
            std::string::npos)
      << "xml=" << pivot_xml;
}

TEST(FormulonCApiPivot, SavedPivotLocationRefCoversTheProjectedGrid) {
  // `fm_workbook_pivot_create` installs a 1x1 placeholder span and nothing
  // revises it as fields are added, so a pivot built purely through the C
  // surface used to save a `ref` describing a single cell. Excel opens such
  // a file cleanly and recognises the pivot, then terminates the moment the
  // report is refreshed -- the defect is invisible to a round trip and has
  // to be pinned on the emitted bytes.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // Deliberately no `fm_workbook_pivot_set_anchor`: this is the path a host
  // takes when it never reasons about the report's extent at all.
  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_EQ(fm_pivot_cells_bounds(projected.handle, &top, &left, &rows, &cols), 0);
  ASSERT_GT(rows, 1U) << "projection must be larger than the placeholder for this test to mean anything";
  ASSERT_GT(cols, 1U);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> package(saved.data, saved.data + saved.len);
  const std::string pivot_xml = ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml");
  ASSERT_FALSE(pivot_xml.empty());

  // Anchored at D1, projected 4 rows x 2 cols -> D1:E4.
  EXPECT_NE(pivot_xml.find("<location ref=\"D1:E4\""), std::string::npos) << "xml=" << pivot_xml;
  EXPECT_EQ(pivot_xml.find("<location ref=\"D1:D1\""), std::string::npos) << "xml=" << pivot_xml;
}

TEST(FormulonCApiPivot, PivotCacheMutationInvalidatesMemoisedLayout) {
  // fm_workbook_pivot_layout memoises the pivot's evaluated result. A cache
  // mutation must invalidate that memo so a re-layout reflects the change
  // instead of returning the stale projection.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

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
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
    EXPECT_TRUE(has_data_value(projected.handle, 400.0));
  }

  // Bump record 0's amount (field index 1) from 100 to 1100 (+1000).
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, 1, 1100.0), 0) << fm_last_error_message();

  // Re-layout must reflect the mutated cache: North is now 1100 + 300 = 1400,
  // and the stale 400 must be gone.
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
    EXPECT_TRUE(has_data_value(projected.handle, 1400.0)) << "layout returned a stale memoised projection";
    EXPECT_FALSE(has_data_value(projected.handle, 400.0)) << "stale North value survived the cache mutation";
  }
}

TEST(FormulonCApiPivot, PivotMutationsInvalidateGetPivotDataFormula) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // A1 is outside the report at D1, so this exercises the normal formula
  // dependency path rather than reading a projected pivot cell directly.
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=GETPIVOTDATA(\"Sum of Amount\",D1,\"Region\",\"North\")"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  auto get_a1 = [&]() {
    fm_value_t value{};
    EXPECT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &value), 0);
    return value;
  };
  fm_value_t value = get_a1();
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 400.0);

  ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, 1, 1100.0), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  value = get_a1();
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 1400.0);

  fm_pivot_data_field_spec_t average_spec{};
  average_spec.name = "Sum of Amount";
  average_spec.field_index = 1U;
  average_spec.aggregation = FM_PIVOT_AGG_AVERAGE;
  average_spec.number_format = "";
  average_spec.show_as = FM_PIVOT_SHOW_AS_NORMAL;
  average_spec.show_as_base_field = -1;
  average_spec.show_as_base_item = -1;
  ASSERT_EQ(fm_workbook_pivot_data_field_set(wb.handle, 0, pivot_idx, 0, &average_spec), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  value = get_a1();
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 700.0);

  // Keeping this formula's result warm before removal proves that the
  // successful lifecycle mutation itself dirties formula cells.
  ASSERT_EQ(fm_workbook_pivot_remove(wb.handle, 0, pivot_idx), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  value = get_a1();
  ASSERT_EQ(value.kind, FM_VAL_ERROR);
  EXPECT_EQ(value.u.error_code, static_cast<std::int32_t>(formulon::ErrorCode::Ref));
}

TEST(FormulonCApiPivot, LoadedPivotCacheMutationReprojectsAuthoredLocation) {
  WorkbookGuard original;
  ASSERT_EQ(fm_workbook_create(&original.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(original.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(original.handle, &saved.data, &saved.len), 0) << fm_last_error_message();

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0) << fm_last_error_message();
  EXPECT_NE(SavedPivotXml(loaded.handle).find("<location ref=\"D1:E4\""), std::string::npos);

  // The loaded location is marked authored. Add a new shared item and record
  // so the evaluated report grows by one row; saving must then use the new
  // projected extent instead of preserving the stale authored ref.
  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_text(loaded.handle, cache_id, 0, "West"), 0)
      << fm_last_error_message();
  std::size_t new_record = 99;
  ASSERT_EQ(fm_workbook_pivot_cache_record_add(loaded.handle, cache_id, &new_record), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(loaded.handle, cache_id, new_record, 0, 2.0), 0)
      << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(loaded.handle, cache_id, new_record, 1, 500.0), 0)
      << fm_last_error_message();

  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(loaded.handle, 0, 0, &projected.handle), 0) << fm_last_error_message();
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_EQ(fm_pivot_cells_bounds(projected.handle, &top, &left, &rows, &cols), 0);
  EXPECT_EQ(top, 0U);
  EXPECT_EQ(left, 3U);
  EXPECT_EQ(rows, 5U);
  EXPECT_EQ(cols, 2U);

  const std::string pivot_xml = SavedPivotXml(loaded.handle);
  EXPECT_NE(pivot_xml.find("<location ref=\"D1:E5\""), std::string::npos) << pivot_xml;
}

TEST(FormulonCApiPivot, PivotCacheSharedItemsAcceptErrorValues) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &count), 0);
  EXPECT_EQ(count, 2U);

  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_error(wb.handle, cache_id, 0,
                                                                1),  // ErrorCode::Div0
            0)
      << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0, &count), 0);
  EXPECT_EQ(count, 3U);
}

TEST(FormulonCApiPivot, PivotCacheSharedItemIndexToleratesOutOfDomainNumbers) {
  // `BuildScratchPivot` leaves `cell_is_index` empty, so "Region" (field 0)
  // reads its numeric cells as indices into its two shared items. The record
  // setter accepts any double, so the layout below is the point where an
  // out-of-domain index would be narrowed. It must resolve to a blank label
  // rather than reading past the shared items or trapping the instance.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  const double bad_indices[] = {
      std::numeric_limits<double>::quiet_NaN(),
      std::numeric_limits<double>::infinity(),
      -std::numeric_limits<double>::infinity(),
      -1.0,
      1e30,
      4.3e9,  // Past the wasm32 `size_t` range.
      2.0,    // One past the last shared item.
  };
  for (const double bad : bad_indices) {
    ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, 0, bad), 0) << fm_last_error_message();
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
    EXPECT_FALSE(CollectCells(projected.handle).empty());
  }

  // A valid index still resolves after all of that.
  ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, 0, 0.0), 0) << fm_last_error_message();
  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
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
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  int32_t present = -1;
  const char* ref = nullptr;
  const char* sheet = nullptr;
  const char* name = nullptr;
  // `BuildScratchPivot` installs a source because a cache without one
  // cannot be saved. Clear it here so the absent state this test is about
  // is established by the test rather than inherited from the helper.
  ASSERT_EQ(fm_workbook_pivot_cache_set_worksheet_source(wb.handle, cache_id, 0, nullptr, nullptr, nullptr), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_get_worksheet_source(wb.handle, cache_id, &present, &ref, &sheet, &name), 0);
  EXPECT_EQ(present, 0);
  EXPECT_STREQ(ref, "");
  EXPECT_STREQ(sheet, "");
  EXPECT_STREQ(name, "");

  ASSERT_EQ(fm_workbook_pivot_cache_set_worksheet_source(wb.handle, cache_id, 1, "$A$1:$C$5", "Data", nullptr), 0)
      << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_cache_get_worksheet_source(wb.handle, cache_id, &present, &ref, &sheet, &name), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(ref, "$A$1:$C$5");
  EXPECT_STREQ(sheet, "Data");
  EXPECT_STREQ(name, "");

  ASSERT_EQ(fm_workbook_pivot_cache_set_worksheet_source(wb.handle, cache_id, 0, "ignored", "ignored", "ignored"), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_get_worksheet_source(wb.handle, cache_id, &present, &ref, &sheet, &name), 0);
  EXPECT_EQ(present, 0);
  EXPECT_STREQ(ref, "");
  EXPECT_STREQ(sheet, "");
  EXPECT_STREQ(name, "");
}

TEST(FormulonCApiPivot, PivotReportLayoutRoundTripsThroughApi) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  fm_pivot_layout_t layout = FM_PIVOT_LAYOUT_TABULAR;
  ASSERT_EQ(fm_workbook_pivot_get_layout(wb.handle, 0, pivot_idx, &layout), 0);
  EXPECT_EQ(layout, FM_PIVOT_LAYOUT_COMPACT);

  ASSERT_EQ(fm_workbook_pivot_set_layout(wb.handle, 0, pivot_idx, FM_PIVOT_LAYOUT_TABULAR), 0);
  ASSERT_EQ(fm_workbook_pivot_get_layout(wb.handle, 0, pivot_idx, &layout), 0);
  EXPECT_EQ(layout, FM_PIVOT_LAYOUT_TABULAR);

  ASSERT_EQ(fm_workbook_pivot_set_layout(wb.handle, 0, pivot_idx, FM_PIVOT_LAYOUT_OUTLINE), 0);
  ASSERT_EQ(fm_workbook_pivot_get_layout(wb.handle, 0, pivot_idx, &layout), 0);
  EXPECT_EQ(layout, FM_PIVOT_LAYOUT_OUTLINE);

  // The setter accepts a raw int32_t so FFI callers can probe the complete
  // invalid domain without constructing an out-of-domain C enum.
  const std::int32_t invalid_layouts[] = {
      99,
      std::numeric_limits<std::int32_t>::min(),
      std::numeric_limits<std::int32_t>::max(),
  };
  for (const std::int32_t raw : invalid_layouts) {
    EXPECT_EQ(fm_workbook_pivot_set_layout(wb.handle, 0, pivot_idx, raw),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << raw;
  }
}

TEST(FormulonCApiPivot, ScalarEnumMutatorsRejectRawValuesWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  BufferGuard before;
  ASSERT_EQ(fm_workbook_save(wb.handle, &before.data, &before.len), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> snapshot(before.data, before.data + before.len);
  const std::int32_t invalid[] = {99, std::numeric_limits<std::int32_t>::min(),
                                  std::numeric_limits<std::int32_t>::max()};
  const fm_status_t expected = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  for (const std::int32_t raw : invalid) {
    EXPECT_EQ(fm_workbook_pivot_field_set_axis(wb.handle, 0, pivot_idx, 0, raw), expected) << raw;
    EXPECT_EQ(fm_workbook_pivot_field_add_subtotal_fn(wb.handle, 0, pivot_idx, 0, raw), expected) << raw;
    EXPECT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, 0, raw, 0, -1, -1, 1, -1.0, -1.0),
              expected)
        << raw;
    EXPECT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, 0, 3, raw, -1, -1, 1, -1.0, -1.0),
              expected)
        << raw;
  }

  BufferGuard after;
  ASSERT_EQ(fm_workbook_save(wb.handle, &after.data, &after.len), 0) << fm_last_error_message();
  EXPECT_EQ(std::vector<std::uint8_t>(after.data, after.data + after.len), snapshot);
}

TEST(FormulonCApiPivot, DateGroupRejectsZeroIntervalAndInvertedWindowWithoutMutation) {
  // FM_PIVOT_DATE_DAYS-only validation: a zero interval and an explicit
  // Start past an explicit End are both nonsensical groupings, so
  // neither may reach the model.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  BufferGuard before;
  ASSERT_EQ(fm_workbook_save(wb.handle, &before.data, &before.len), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> snapshot(before.data, before.data + before.len);
  const fm_status_t expected = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

  // interval_days == 0.
  EXPECT_EQ(
      fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, 0, FM_PIVOT_DATE_DAYS,
                                             FM_PIVOT_CALENDAR_GREGORIAN, -1, -1, /*interval_days=*/0, -1.0, -1.0),
      expected);
  // start_serial_or_neg1 > end_serial_or_neg1 (both explicit).
  EXPECT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, 0, FM_PIVOT_DATE_DAYS,
                                                   FM_PIVOT_CALENDAR_GREGORIAN, -1, -1, /*interval_days=*/7,
                                                   /*start_serial=*/61.0, /*end_serial=*/60.0),
            expected);
  // Compare the raw values before flooring: 0.9 > 0.1 is still inverted.
  EXPECT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, 0, FM_PIVOT_DATE_DAYS,
                                                   FM_PIVOT_CALENDAR_GREGORIAN, -1, -1, /*interval_days=*/7,
                                                   /*start_serial=*/0.9, /*end_serial=*/0.1),
            expected);

  BufferGuard after;
  ASSERT_EQ(fm_workbook_save(wb.handle, &after.data, &after.len), 0) << fm_last_error_message();
  EXPECT_EQ(std::vector<std::uint8_t>(after.data, after.data + after.len), snapshot);
}

TEST(FormulonCApiPivot, DateGroupRejectsInvalidSerialBoundsWithoutMutation) {
  const fm_status_t expected = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  for (const bool date1904 : {false, true}) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    if (date1904) {
      wb.handle->workbook().set_date1904(true);
    }
    std::uint32_t cache_id = 0;
    std::size_t pivot_idx = 0;
    const double dates[] = {
        formulon::date_time::serial_from_ymd(2024, 1U, 1U, date1904),
        formulon::date_time::serial_from_ymd(2025, 1U, 1U, date1904),
        formulon::date_time::serial_from_ymd(2024, 6U, 1U, date1904),
        formulon::date_time::serial_from_ymd(2025, 6U, 1U, date1904),
    };
    const double amounts[] = {100.0, 200.0, 300.0, 400.0};

    ASSERT_EQ(fm_workbook_pivot_cache_create(wb.handle, 0, &cache_id), 0) << fm_last_error_message();
    ASSERT_EQ(fm_workbook_pivot_cache_set_worksheet_source(wb.handle, cache_id, 1, "A1:B5", "Sheet1", nullptr), 0)
        << fm_last_error_message();
    std::size_t date_cache_field = 99;
    std::size_t amount_cache_field = 99;
    ASSERT_EQ(fm_workbook_pivot_cache_field_add(wb.handle, cache_id, "Date", &date_cache_field), 0)
        << fm_last_error_message();
    ASSERT_EQ(fm_workbook_pivot_cache_field_add(wb.handle, cache_id, "Amount", &amount_cache_field), 0)
        << fm_last_error_message();
    for (std::size_t record = 0; record < std::size(dates); ++record) {
      std::size_t record_idx = 99;
      ASSERT_EQ(fm_workbook_pivot_cache_record_add(wb.handle, cache_id, &record_idx), 0) << fm_last_error_message();
      ASSERT_EQ(
          fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, record_idx, date_cache_field, dates[record]),
          0)
          << fm_last_error_message();
      ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, record_idx, amount_cache_field,
                                                          amounts[record]),
                0)
          << fm_last_error_message();
    }

    ASSERT_EQ(fm_workbook_pivot_create(wb.handle, 0, "PT", cache_id, 0U, 3U, &pivot_idx), 0) << fm_last_error_message();
    fm_pivot_field_spec_t date_spec{};
    date_spec.source_name = "Date";
    date_spec.custom_name = "";
    date_spec.axis = FM_PIVOT_AXIS_ROW;
    date_spec.subtotal_top = 0;
    date_spec.number_format = "";
    std::size_t date_field = 99;
    ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot_idx, &date_spec, &date_field), 0)
        << fm_last_error_message();
    const std::uint32_t row_order[] = {static_cast<std::uint32_t>(date_field)};
    ASSERT_EQ(fm_workbook_pivot_set_row_field_order(wb.handle, 0, pivot_idx, row_order, 1U), 0)
        << fm_last_error_message();
    fm_pivot_field_spec_t amount_spec{};
    amount_spec.source_name = "Amount";
    amount_spec.custom_name = "";
    amount_spec.axis = FM_PIVOT_AXIS_VALUE;
    amount_spec.subtotal_top = 0;
    amount_spec.number_format = "";
    std::size_t amount_field = 99;
    ASSERT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot_idx, &amount_spec, &amount_field), 0)
        << fm_last_error_message();
    fm_pivot_data_field_spec_t data_spec{};
    data_spec.name = "Sum of Amount";
    data_spec.field_index = static_cast<std::uint32_t>(amount_field);
    data_spec.aggregation = FM_PIVOT_AGG_SUM;
    data_spec.number_format = "";
    data_spec.show_as = FM_PIVOT_SHOW_AS_NORMAL;
    data_spec.show_as_base_field = -1;
    data_spec.show_as_base_item = -1;
    std::size_t data_field = 99;
    ASSERT_EQ(fm_workbook_pivot_data_field_add(wb.handle, 0, pivot_idx, &data_spec, &data_field), 0)
        << fm_last_error_message();
    ASSERT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, date_field, FM_PIVOT_DATE_YEAR,
                                                     FM_PIVOT_CALENDAR_GREGORIAN, -1, -1, 1, -1.0, -1.0),
              0)
        << fm_last_error_message();

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
      ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
      EXPECT_TRUE(LayoutHasText(wb.handle, pivot_idx, "2024"));
      EXPECT_TRUE(LayoutHasText(wb.handle, pivot_idx, "2025"));
      EXPECT_TRUE(has_data_value(projected.handle, 400.0));
      EXPECT_TRUE(has_data_value(projected.handle, 600.0));
    }

    BufferGuard before;
    ASSERT_EQ(fm_workbook_save(wb.handle, &before.data, &before.len), 0) << fm_last_error_message();
    const std::vector<std::uint8_t> snapshot(before.data, before.data + before.len);
    const double last_valid = formulon::date_time::serial_from_ymd(9999, 12U, 31U, date1904);
    const double invalid[] = {
        -2.0,
        -0.1,
        std::numeric_limits<double>::quiet_NaN(),
        std::numeric_limits<double>::infinity(),
        -std::numeric_limits<double>::infinity(),
        last_valid + 1.0,
        std::numeric_limits<double>::max(),
    };
    for (const double raw : invalid) {
      EXPECT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, date_field, FM_PIVOT_DATE_DAYS,
                                                       FM_PIVOT_CALENDAR_GREGORIAN, -1, -1, 7, raw, 70.0),
                expected)
          << "invalid start=" << raw << " date1904=" << date1904;
      EXPECT_EQ(fm_workbook_pivot_field_set_date_group(wb.handle, 0, pivot_idx, date_field, FM_PIVOT_DATE_DAYS,
                                                       FM_PIVOT_CALENDAR_GREGORIAN, -1, -1, 7, 0.0, raw),
                expected)
          << "invalid end=" << raw << " date1904=" << date1904;
    }

    // Force the memoised result through the cache mutation path before
    // re-evaluating. The successful no-op write also proves the rejected
    // setters did not leave the field in an invalid state.
    ASSERT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, 1, 100.0), 0)
        << fm_last_error_message();
    {
      PivotCellsGuard projected;
      ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
      EXPECT_TRUE(LayoutHasText(wb.handle, pivot_idx, "2024"));
      EXPECT_TRUE(LayoutHasText(wb.handle, pivot_idx, "2025"));
      EXPECT_TRUE(has_data_value(projected.handle, 400.0));
      EXPECT_TRUE(has_data_value(projected.handle, 600.0));
    }

    BufferGuard after;
    ASSERT_EQ(fm_workbook_save(wb.handle, &after.data, &after.len), 0) << fm_last_error_message();
    EXPECT_EQ(std::vector<std::uint8_t>(after.data, after.data + after.len), snapshot);
  }
}

TEST(FormulonCApiPivot, StructEnumMutatorsRejectRawValuesWithoutMutation) {
  // The struct-taking twins of the scalar mutators above must enforce the
  // same domains. Without that, an out-of-domain `axis` collapses onto
  // `PivotAxis::Row` and an out-of-domain `show_as` onto `Normal`, and
  // the call reports success on a pivot that means something else.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  BufferGuard before;
  ASSERT_EQ(fm_workbook_save(wb.handle, &before.data, &before.len), 0) << fm_last_error_message();
  const std::vector<std::uint8_t> snapshot(before.data, before.data + before.len);
  const fm_status_t expected = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

  fm_pivot_field_spec_t field_spec{};
  field_spec.source_name = "Region";
  field_spec.axis = RawEnumValue<fm_pivot_axis_t>(99);
  std::size_t field_idx = 99U;
  EXPECT_EQ(fm_workbook_pivot_field_add(wb.handle, 0, pivot_idx, &field_spec, &field_idx), expected);

  fm_pivot_data_field_spec_t df_spec{};
  df_spec.name = "Total";
  df_spec.field_index = 1;
  df_spec.aggregation = RawEnumValue<fm_pivot_aggregation_t>(99);
  std::size_t df_idx = 99U;
  EXPECT_EQ(fm_workbook_pivot_data_field_add(wb.handle, 0, pivot_idx, &df_spec, &df_idx), expected);
  EXPECT_EQ(fm_workbook_pivot_data_field_set(wb.handle, 0, pivot_idx, 0, &df_spec), expected);

  df_spec.aggregation = FM_PIVOT_AGG_SUM;
  df_spec.show_as = RawEnumValue<fm_pivot_show_as_t>(99);
  EXPECT_EQ(fm_workbook_pivot_data_field_add(wb.handle, 0, pivot_idx, &df_spec, &df_idx), expected);
  EXPECT_EQ(fm_workbook_pivot_data_field_set(wb.handle, 0, pivot_idx, 0, &df_spec), expected);

  BufferGuard after;
  ASSERT_EQ(fm_workbook_save(wb.handle, &after.data, &after.len), 0) << fm_last_error_message();
  EXPECT_EQ(std::vector<std::uint8_t>(after.data, after.data + after.len), snapshot);
}

TEST(FormulonCApiPivot, PivotProjectionUsesWorkbookLocaleAndReportLayout) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  ASSERT_EQ(fm_workbook_pivot_set_layout(wb.handle, 0, pivot_idx, FM_PIVOT_LAYOUT_TABULAR), 0);

  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  ASSERT_EQ(fm_pivot_cells_bounds(projected.handle, &top, &left, &rows, &cols), 0);
  EXPECT_EQ(cols, 2U);  // Tabular exposes its two row-field columns.

  bool saw_japanese_total = false;
  const std::size_t cell_count = fm_pivot_cells_count(projected.handle);
  for (std::size_t i = 0; i < cell_count; ++i) {
    fm_pivot_cell_t cell{};
    ASSERT_EQ(fm_pivot_cells_at(projected.handle, i, &cell), 0);
    if (cell.value.kind == FM_VAL_TEXT && cell.value.u.text != nullptr &&
        std::string_view(cell.value.u.text) == "総計") {
      saw_japanese_total = true;
    }
  }
  EXPECT_TRUE(saw_japanese_total);
}

TEST(FormulonCApiPivot, PivotCacheErrorSettersRejectInvalidErrorCodes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  EXPECT_EQ(fm_workbook_pivot_cache_field_add_shared_item_error(wb.handle, cache_id, 0, -1),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_pivot_cache_field_add_shared_item_error(wb.handle, cache_id, 0, 999),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_error(wb.handle, cache_id, 0, 1, -1),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_error(wb.handle, cache_id, 0, 1, 999),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiPivot, PivotCacheRecordSettersRejectOutOfRangeFieldIndex) {
  // BuildScratchPivot declares two cache fields (Region=0, Amount=1), so
  // field_idx 2 is the first out-of-range value. A large / wrapping field_idx
  // must be rejected before `grow_record_cells` resizes `field_idx + 1`
  // cells — otherwise SIZE_MAX wraps to 0 and the subsequent cell write lands
  // out of bounds.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  const auto kInvalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  const std::size_t kFieldCount = 2;
  const std::size_t kMax = static_cast<std::size_t>(-1);

  // In-range write still succeeds (guards against an over-tight bound).
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, kFieldCount - 1, 1.0), 0);

  // field_idx == field_count and beyond are rejected on every setter.
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, kFieldCount, 1.0), kInvalid);
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_number(wb.handle, cache_id, 0, kMax, 1.0), kInvalid);
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_text(wb.handle, cache_id, 0, kMax, "x"), kInvalid);
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_bool(wb.handle, cache_id, 0, kMax, 1), kInvalid);
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_blank(wb.handle, cache_id, 0, kMax), kInvalid);
  EXPECT_EQ(fm_workbook_pivot_cache_record_set_error(wb.handle, cache_id, 0, kMax, 0), kInvalid);
}
