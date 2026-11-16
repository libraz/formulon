// A pivot assembled through the C ABI (cache fields, inline record values,
// fields placed on axes without items) must save as a package Excel opens: every
// axis field lists its items and every discrete cache field carries shared items.
// Its report lives in the sheet cells, written on every table edit, blocked
// where it would grow over another structure, and cleared on removal.

#include <cstddef>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <string>
#include <utility>
#include <vector>

#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"

namespace {

struct SourceRow {
  const char* region;  // nullptr is a blank cell.
  const char* product;
  double qty;
  double amt;
};

constexpr SourceRow kRows[] = {{"East", "Pen", 3, 30},  {"East", "Ink", 5, 100}, {"West", "Pen", 4, 40},
                               {"West", "Pen", 6, 60},  {"West", "Ink", 2, 50},  {"North", "Pen", 1, 10},
                               {nullptr, "Ink", 7, 70}, {"East", "Pen", 2, 20}};

void SetText(fm_workbook_t* wb, std::uint32_t cache_id, std::size_t record, std::size_t field, const char* text) {
  if (text == nullptr) {
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_blank(wb, cache_id, record, field), 0) << fm_get_last_error_message();
  } else {
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_text(wb, cache_id, record, field, text), 0)
        << fm_get_last_error_message();
  }
}

// The sequence a host drives: cache fields and inline records, one pivot field
// per cache field, row / column orders, then two data fields.
void BuildHostPivot(fm_workbook_t* wb) {
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb, 0, &cache_id), 0) << fm_get_last_error_message();
  ASSERT_EQ(SetPivotCacheWorksheetSource(wb, cache_id, 1, "A1:D9", "Sheet1", nullptr), 0);
  const char* const headers[] = {"Region", "Product", "Qty", "Amt"};
  for (std::uint32_t c = 0; c < 4U; ++c) {
    std::size_t idx = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb, cache_id, headers[c], &idx), 0) << fm_get_last_error_message();
    ASSERT_EQ(fm_sheet_set_text(wb, 0U, 0U, c, headers[c]), 0);
  }
  // The source holds the records, so a reload's refresh reads the same cache.
  for (std::uint32_t r = 0; r < 8U; ++r) {
    if (kRows[r].region != nullptr) {
      ASSERT_EQ(fm_sheet_set_text(wb, 0U, r + 1U, 0U, kRows[r].region), 0);
    }
    ASSERT_EQ(fm_sheet_set_text(wb, 0U, r + 1U, 1U, kRows[r].product), 0);
    ASSERT_EQ(fm_sheet_set_number(wb, 0U, r + 1U, 2U, kRows[r].qty), 0);
    ASSERT_EQ(fm_sheet_set_number(wb, 0U, r + 1U, 3U, kRows[r].amt), 0);
  }
  for (const SourceRow& row : kRows) {
    std::size_t record = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb, cache_id, &record), 0);
    SetText(wb, cache_id, record, 0, row.region);
    SetText(wb, cache_id, record, 1, row.product);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb, cache_id, record, 2, row.qty), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb, cache_id, record, 3, row.amt), 0);
  }
  std::size_t pivot = 99;
  ASSERT_EQ(AddPivot(wb, 0, "PT1", cache_id, 2U, 5U, &pivot), 0) << fm_get_last_error_message();
  const fm_pivot_axis_t axes[] = {FM_PIVOT_AXIS_ROW, FM_PIVOT_AXIS_COLUMN, FM_PIVOT_AXIS_VALUE, FM_PIVOT_AXIS_VALUE};
  for (std::size_t i = 0; i < 4; ++i) {
    // The pivot already carries one field per cache field, in cache order.
    ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb, 0, pivot, i, axes[i]), 0) << fm_get_last_error_message();
  }
  const std::uint32_t rows[] = {0U};
  const std::uint32_t cols[] = {1U};
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb, 0, pivot, rows, 1U), 0);
  ASSERT_EQ(fm_sheet_set_pivot_column_field_order(wb, 0, pivot, cols, 1U), 0);
  const struct {
    const char* name;
    std::uint32_t field;
  } values[] = {{"Amt sum", 3U}, {"Qty sum", 2U}};
  for (const auto& value : values) {
    fm_pivot_data_field_t df = FM_STRUCT_INIT(fm_pivot_data_field_t);
    df.name = value.name;
    df.field_index = value.field;
    df.aggregation = FM_PIVOT_AGGREGATION_SUM;
    df.number_format = "";
    df.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
    df.has_show_as_base_field = 0;
    df.has_show_as_base_item = 0;
    std::size_t idx = 99;
    ASSERT_EQ(fm_sheet_add_pivot_data_field(wb, 0, pivot, &df, &idx), 0) << fm_get_last_error_message();
  }
  ASSERT_EQ(fm_sheet_set_pivot_report_layout(wb, 0, pivot, FM_PIVOT_REPORT_LAYOUT_COMPACT), 0);
}

struct Cell {
  std::uint32_t row;
  std::uint32_t col;
  int32_t kind;
  int32_t value_kind;
  std::string text;
  double number;
  bool operator==(const Cell& o) const {
    return row == o.row && col == o.col && kind == o.kind && value_kind == o.value_kind && text == o.text &&
           number == o.number;
  }
};

std::vector<Cell> ProjectedCells(fm_workbook_t* wb) {
  PivotCellsGuard projected;
  EXPECT_EQ(fm_sheet_get_pivot_layout(wb, 0, 0, &projected.handle), 0) << fm_get_last_error_message();
  std::vector<Cell> out;
  for (const fm_pivot_cell_info_t& c : CollectCells(projected.handle)) {
    Cell cell{c.row, c.col, c.kind, c.value.kind, {}, 0.0};
    if (c.value.kind == FM_VALUE_KIND_TEXT && c.value.u.text != nullptr) {
      cell.text = c.value.u.text;
    } else if (c.value.kind == FM_VALUE_KIND_NUMBER) {
      cell.number = c.value.u.number;
    }
    out.push_back(std::move(cell));
  }
  return out;
}

}  // namespace

TEST(FormulonCApiPivotSaveShape, HostBuiltPivotSavesWithItemsAndSharedItems) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  BuildHostPivot(wb.handle);
  const std::vector<Cell> before = ProjectedCells(wb.handle);
  ASSERT_FALSE(before.empty());

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
  const std::vector<std::uint8_t> package(saved.data, saved.data + saved.len);
  const std::string table = ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml");
  const std::string def = ExtractZipEntry(package, "xl/pivotCache/pivotCacheDefinition1.xml");
  const std::string rec = ExtractZipEntry(package, "xl/pivotCache/pivotCacheRecords1.xml");

  // Region and Product sit on axes: each lists one item per shared item plus the default subtotal.
  // Region has East, West, North and blank; Product has Pen and Ink.
  EXPECT_EQ(CountOccurrences(table, "<items count=\"5\">"), 1U) << table;
  EXPECT_EQ(CountOccurrences(table, "<items count=\"3\">"), 1U) << table;
  EXPECT_EQ(CountOccurrences(table, "<items"), 2U) << table;
  // East, North, West ascending, blank last; Ink before Pen.
  EXPECT_NE(table.find("<item x=\"0\"/><item x=\"2\"/><item x=\"1\"/><item x=\"3\"/><item t=\"default\"/>"),
            std::string::npos)
      << table;
  EXPECT_NE(table.find("<item x=\"1\"/><item x=\"0\"/><item t=\"default\"/>"), std::string::npos) << table;

  // Discrete fields carry shared items, never the numeric placeholder.
  EXPECT_NE(def.find("<cacheField name=\"Region\"><sharedItems containsBlank=\"1\" count=\"4\"><s v=\"East\"/>"
                     "<s v=\"West\"/><s v=\"North\"/><m/></sharedItems>"),
            std::string::npos)
      << def;
  EXPECT_NE(def.find("<cacheField name=\"Product\"><sharedItems count=\"2\"><s v=\"Pen\"/><s v=\"Ink\"/>"),
            std::string::npos)
      << def;
  // Qty (1..7) and Amt (10..100) are range-typed fields with derived numeric hints.
  EXPECT_NE(def.find("<cacheField name=\"Qty\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" "
                     "containsNumber=\"1\" containsInteger=\"1\" minValue=\"1\" maxValue=\"7\"/>"),
            std::string::npos)
      << def;
  EXPECT_NE(def.find("<cacheField name=\"Amt\"><sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" "
                     "containsNumber=\"1\" containsInteger=\"1\" minValue=\"10\" maxValue=\"100\"/>"),
            std::string::npos)
      << def;

  // Records reference the derived fields by index and hold no inline text for them.
  EXPECT_EQ(rec.find("<s v="), std::string::npos) << rec;
  EXPECT_EQ(rec.find("<m/>"), std::string::npos) << rec;
  EXPECT_NE(rec.find("<r><x v=\"0\"/><x v=\"0\"/><n v=\"3\"/><n v=\"30\"/></r>"), std::string::npos) << rec;
  EXPECT_NE(rec.find("<r><x v=\"3\"/><x v=\"1\"/><n v=\"7\"/><n v=\"70\"/></r>"), std::string::npos) << rec;

  // Reloading evaluates to the report the host saw.
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &loaded.handle), 0) << fm_get_last_error_message();
  EXPECT_EQ(ProjectedCells(loaded.handle), before);
}

// ---------------------------------------------------------------------------
// The report in the sheet cells
// ---------------------------------------------------------------------------

namespace {

constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);

// BuildScratchPivot's report at D1:E4: a header row, North 400, South 600 and
// the grand total 1000 in column E.
constexpr std::uint32_t kValueCol = 4U;
constexpr std::uint32_t kLabelCol = 3U;

fm_value_t SheetValue(fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  fm_value_t value{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0U, row, col, &value), 0) << fm_get_last_error_message();
  return value;
}

void ExpectSheetNumber(fm_workbook_t* wb, std::uint32_t row, std::uint32_t col, double expected) {
  const fm_value_t value = SheetValue(wb, row, col);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER) << "row=" << row << " col=" << col;
  EXPECT_DOUBLE_EQ(value.u.number, expected) << "row=" << row << " col=" << col;
}

void ExpectSheetText(fm_workbook_t* wb, std::uint32_t row, std::uint32_t col, const std::string& expected) {
  const fm_value_t value = SheetValue(wb, row, col);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_TEXT) << "row=" << row << " col=" << col;
  EXPECT_EQ(std::string(value.u.text), expected) << "row=" << row << " col=" << col;
}

void BuildScratch(WorkbookGuard& wb, std::uint32_t* cache_id = nullptr) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, cache_id, nullptr), 0) << fm_get_last_error_message();
  FillScratchSource(wb.handle);
}

std::int32_t Pending(const fm_workbook_t* wb) {
  std::int32_t pending = -1;
  EXPECT_EQ(fm_sheet_is_pivot_projection_pending(wb, 0U, 0U, &pending), 0) << fm_get_last_error_message();
  return pending;
}

bool LayoutHasNumber(fm_workbook_t* wb, double number) {
  PivotCellsGuard projected;
  EXPECT_EQ(fm_sheet_get_pivot_layout(wb, 0U, 0U, &projected.handle), 0) << fm_get_last_error_message();
  for (const fm_pivot_cell_info_t& cell : CollectCells(projected.handle)) {
    if (cell.value.kind == FM_VALUE_KIND_NUMBER && cell.value.u.number == number) {
      return true;
    }
  }
  return false;
}

void Save(fm_workbook_t* wb) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0) << fm_get_last_error_message();
}

}  // namespace

TEST(FormulonCApiPivotSaveShape, ApiPivotReportIsInTheCellsAndFormulasReadIt) {
  WorkbookGuard wb;
  BuildScratch(wb);
  ExpectSheetText(wb.handle, 1U, kLabelCol, "North");
  ExpectSheetNumber(wb.handle, 1U, kValueCol, 400.0);
  ExpectSheetNumber(wb.handle, 2U, kValueCol, 600.0);
  ExpectSheetNumber(wb.handle, 3U, kValueCol, 1000.0);

  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, 9U, 0U, "=SUM(E1:E10)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ExpectSheetNumber(wb.handle, 9U, 0U, 2000.0);

  // Hiding South re-projects the report one row shorter; the old last row is cleared.
  ASSERT_EQ(SetPivotFieldItemVisible(wb.handle, 0U, 0U, 0U, 1U, 0), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ExpectSheetNumber(wb.handle, 2U, kValueCol, 400.0);
  EXPECT_EQ(SheetValue(wb.handle, 3U, kValueCol).kind, FM_VALUE_KIND_BLANK);
  ExpectSheetNumber(wb.handle, 9U, 0U, 800.0);

  // The saved file holds the cells.
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &loaded.handle), 0) << fm_get_last_error_message();
  ExpectSheetText(loaded.handle, 1U, kLabelCol, "North");
  ExpectSheetNumber(loaded.handle, 1U, kValueCol, 400.0);
  ExpectSheetNumber(loaded.handle, 2U, kValueCol, 400.0);
  EXPECT_EQ(SheetValue(loaded.handle, 3U, kValueCol).kind, FM_VALUE_KIND_BLANK);
}

TEST(FormulonCApiPivotSaveShape, EditGrowingOntoAMergeBlocksTheReport) {
  WorkbookGuard wb;
  BuildScratch(wb);
  const formulon::pivot::PivotTable* table = wb.handle->wb->sheet(0U).pivot_tables()[0].get();
  fm_cell_rect_t merged{};
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0U, record_arg(fm_cell_rect_t{0U, 5U, 1U, 5U}), &merged), 0)
      << fm_get_last_error_message();

  // A second data field widens the report into column F.
  fm_pivot_data_field_t spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  spec.name = "Count of Amount";
  spec.field_index = 1U;
  spec.aggregation = FM_PIVOT_AGGREGATION_COUNT;
  spec.number_format = "";
  spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  spec.has_show_as_base_field = 0;
  spec.has_show_as_base_item = 0;
  std::size_t index = 99;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0U, 0U, &spec, &index), 0) << fm_get_last_error_message();
  EXPECT_EQ(index, 1U);

  // The definition is applied in place; the old report is cleared, the anchor shows #SPILL! and the merge stays.
  EXPECT_EQ(wb.handle->wb->sheet(0U).pivot_tables()[0].get(), table);
  std::size_t data_fields = 0;
  ASSERT_EQ(fm_sheet_get_pivot_data_field_count(wb.handle, 0U, 0U, &data_fields), 0);
  EXPECT_EQ(data_fields, 2U);
  const fm_value_t anchor = SheetValue(wb.handle, 0U, kLabelCol);
  ASSERT_EQ(anchor.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(anchor.u.error_code, static_cast<std::int32_t>(formulon::ErrorCode::Spill));
  EXPECT_EQ(SheetValue(wb.handle, 3U, kValueCol).kind, FM_VALUE_KIND_BLANK);
  EXPECT_EQ(SheetValue(wb.handle, 0U, 5U).kind, FM_VALUE_KIND_BLANK);
  std::size_t merges = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb.handle, 0U, &merges), 0);
  EXPECT_EQ(merges, 1U);
  EXPECT_EQ(Pending(wb.handle), 0);
  std::int32_t present = 0;
  fm_spill_info_t spill = FM_STRUCT_INIT(fm_spill_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_spill(wb.handle, 0U, 0U, &present, &spill), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
  EXPECT_EQ(spill.rect.first_row, 0U);
  EXPECT_EQ(spill.rect.first_col, kLabelCol);
  EXPECT_EQ(spill.rect.last_row, 3U);
  EXPECT_EQ(spill.rect.last_col, 5U);

  // An options edit of the blocked report is applied too.
  fm_pivot_options_patch_t grid = FM_STRUCT_INIT(fm_pivot_options_patch_t);
  grid.grid_drop_zones_op = FM_PATCH_OP_SET;
  grid.grid_drop_zones = 1;
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0U, record_arg(fm_cell_rect_t{4U, 3U, 4U, 4U}), &merged), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_update_pivot_options(wb.handle, 0U, 0U, &grid), 0) << fm_get_last_error_message();
  fm_pivot_options_t options = FM_STRUCT_INIT(fm_pivot_options_t);
  ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, 0U, &options), 0);
  EXPECT_EQ(options.grid_drop_zones, 1);
  ASSERT_EQ(fm_sheet_get_pivot_spill(wb.handle, 0U, 0U, &present, &spill), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
}

TEST(FormulonCApiPivotSaveShape, CreatingAPivotOverAnotherReportCreatesNothing) {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  BuildScratch(wb, &cache_id);
  std::size_t created = 99;
  EXPECT_EQ(AddPivot(wb.handle, 0U, "PT2", cache_id, 1U, kValueCol, &created), kPreconditionFailed);
  EXPECT_EQ(created, 99U);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_count(wb.handle, 0U, &count), 0);
  EXPECT_EQ(count, 1U);
  ExpectSheetNumber(wb.handle, 1U, kValueCol, 400.0);

  // Beside the report there is room.
  ASSERT_EQ(AddPivot(wb.handle, 0U, "PT2", cache_id, 0U, 7U, &created), 0) << fm_get_last_error_message();
  EXPECT_EQ(created, 1U);
  fm_pivot_options_t options = FM_STRUCT_INIT(fm_pivot_options_t);
  ASSERT_EQ(fm_sheet_get_pivot_options(wb.handle, 0U, created, &options), 0);
  EXPECT_EQ(options.item_print_titles, 1);
  EXPECT_EQ(options.indent, 0U);
  EXPECT_EQ(options.multiple_field_filters, 0);
}

TEST(FormulonCApiPivotSaveShape, RemovingAPivotClearsItsReport) {
  WorkbookGuard wb;
  BuildScratch(wb);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, 9U, 0U, "=SUM(E1:E10)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ExpectSheetNumber(wb.handle, 9U, 0U, 2000.0);

  ASSERT_EQ(fm_sheet_remove_pivot(wb.handle, 0U, 0U), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  for (std::uint32_t row = 0; row < 4U; ++row) {
    EXPECT_EQ(SheetValue(wb.handle, row, kLabelCol).kind, FM_VALUE_KIND_BLANK) << row;
    EXPECT_EQ(SheetValue(wb.handle, row, kValueCol).kind, FM_VALUE_KIND_BLANK) << row;
  }
  ExpectSheetNumber(wb.handle, 9U, 0U, 0.0);
}

TEST(FormulonCApiPivotSaveShape, CacheEditLeavesTheReportPendingUntilProjected) {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  BuildScratch(wb, &cache_id);
  EXPECT_EQ(Pending(wb.handle), 0);

  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 0U, 1U, 1100.0), 0);
  EXPECT_EQ(Pending(wb.handle), 1);
  // The cells are stale; the layout API is not.
  ExpectSheetNumber(wb.handle, 1U, kValueCol, 400.0);
  EXPECT_TRUE(LayoutHasNumber(wb.handle, 1400.0));

  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, 0U), 0) << fm_get_last_error_message();
  EXPECT_EQ(Pending(wb.handle), 0);
  ExpectSheetNumber(wb.handle, 1U, kValueCol, 1400.0);

  // A save projects what is pending first.
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 1U, 1U, 900.0), 0);
  EXPECT_EQ(Pending(wb.handle), 1);
  Save(wb.handle);
  EXPECT_EQ(Pending(wb.handle), 0);
  ExpectSheetNumber(wb.handle, 2U, kValueCol, 1300.0);

  // So does the next table edit.
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, 1U, 1U, 200.0), 0);
  EXPECT_EQ(Pending(wb.handle), 1);
  ASSERT_EQ(fm_sheet_rename_pivot(wb.handle, 0U, 0U, "Renamed"), 0) << fm_get_last_error_message();
  EXPECT_EQ(Pending(wb.handle), 0);
  ExpectSheetNumber(wb.handle, 2U, kValueCol, 600.0);
}

TEST(FormulonCApiPivotSaveShape, AProjectionOntoAMergeIsBlockedUntilAValueWrite) {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  BuildScratch(wb, &cache_id);
  // A new item grows the report to D1:E5, onto a merge.
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache_id, 0U, "West", ignored_index()), 0);
  std::size_t record = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache_id, &record), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, record, 0U, 2.0), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, record, 1U, 500.0), 0);
  fm_cell_rect_t merged{};
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0U, record_arg(fm_cell_rect_t{4U, 3U, 4U, 4U}), &merged), 0)
      << fm_get_last_error_message();

  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, 0U), 0) << fm_get_last_error_message();
  // Blocked, not pending: the old report is cleared and the anchor shows #SPILL!.
  EXPECT_EQ(Pending(wb.handle), 0);
  EXPECT_EQ(SheetValue(wb.handle, 3U, kValueCol).kind, FM_VALUE_KIND_BLANK);
  EXPECT_EQ(SheetValue(wb.handle, 0U, kLabelCol).kind, FM_VALUE_KIND_ERROR);
  std::int32_t present = 0;
  fm_spill_info_t spill = FM_STRUCT_INIT(fm_spill_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_spill(wb.handle, 0U, 0U, &present, &spill), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
  EXPECT_EQ(spill.rect.last_row, 4U);
  EXPECT_EQ(spill.rect.last_col, kValueCol);
  // The save goes on and the report stays blocked.
  Save(wb.handle);
  ASSERT_EQ(fm_sheet_get_pivot_spill(wb.handle, 0U, 0U, &present, &spill), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);

  // Without the merge, the next value write anywhere renders the report.
  ASSERT_EQ(fm_sheet_remove_merge(wb.handle, 0U, record_arg(fm_cell_rect_t{4U, 3U, 4U, 4U})), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0U, 20U, 0U, 1.0), 0);
  ExpectSheetNumber(wb.handle, 4U, kValueCol, 1500.0);
  ASSERT_EQ(fm_sheet_get_pivot_spill(wb.handle, 0U, 0U, &present, &spill), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  EXPECT_EQ(spill.blocked, 0);
}

TEST(FormulonCApiPivotSaveShape, ProjectingAPivotWithoutRecordsWritesNothing) {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  BuildScratch(wb, &cache_id);
  fm_pivot_cache_options_patch_t options = FM_STRUCT_INIT(fm_pivot_cache_options_patch_t);
  options.save_data_op = FM_PATCH_OP_SET;
  options.save_data = 0;
  ASSERT_EQ(fm_workbook_update_pivot_cache_options(wb.handle, cache_id, &options), 0) << fm_get_last_error_message();
  // A source that does not resolve refuses the reload's refresh, so the cache stays without records.
  ASSERT_EQ(SetPivotCacheWorksheetSource(wb.handle, cache_id, /*present=*/1, "A1:B5", "NoSuchSheet", nullptr), 0);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();

  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &loaded.handle), 0) << fm_get_last_error_message();
  // The cells read with the file are the report.
  ExpectSheetNumber(loaded.handle, 1U, kValueCol, 400.0);
  ASSERT_EQ(fm_sheet_set_text(loaded.handle, 0U, 1U, kValueCol, "kept"), 0);
  EXPECT_EQ(fm_sheet_project_pivot(loaded.handle, 0U, 0U), 0) << fm_get_last_error_message();
  ExpectSheetText(loaded.handle, 1U, kValueCol, "kept");
}

TEST(FormulonCApiPivotSaveShape, SavingDoesNotMemoiseAReportWithoutItsCalculatedCells) {
  WorkbookGuard wb;
  std::string dir = __FILE__;
  dir = dir.substr(0, dir.find_last_of("/\\"));
  std::ifstream in(dir + "/../fixtures/excel/pivot_calc/cf1_price_amt_per_qty.xlsx", std::ios::binary);
  ASSERT_TRUE(in.good());
  const std::vector<std::uint8_t> bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), nullptr, &wb.handle), 0) << fm_get_last_error_message();
  // The save evaluates the report for its span without the formula host.
  Save(wb.handle);
  // The layout API still shows the calculated field: 630 / 20 at the grand total.
  EXPECT_TRUE(LayoutHasNumber(wb.handle, 31.5));
  EXPECT_TRUE(LayoutHasNumber(wb.handle, 50.0));
}

TEST(FormulonCApiPivotSaveShape, ProjectionEntryPointsRejectBadArguments) {
  WorkbookGuard wb;
  BuildScratch(wb);
  std::int32_t pending = -1;
  EXPECT_EQ(fm_sheet_project_pivot(nullptr, 0U, 0U), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_is_pivot_projection_pending(nullptr, 0U, 0U, &pending), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_is_pivot_projection_pending(wb.handle, 0U, 0U, nullptr), kBindingNullPointer);
  for (const std::size_t index : {std::size_t{1}, std::size_t{0x80000000U}, static_cast<std::size_t>(0x100000000ULL)}) {
    EXPECT_EQ(fm_sheet_project_pivot(wb.handle, index, 0U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE) << index;
    EXPECT_EQ(fm_sheet_project_pivot(wb.handle, 0U, index), FM_STATUS_INDEX_OUT_OF_RANGE) << index;
    EXPECT_EQ(fm_sheet_is_pivot_projection_pending(wb.handle, index, 0U, &pending), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE)
        << index;
    EXPECT_EQ(fm_sheet_is_pivot_projection_pending(wb.handle, 0U, index, &pending), FM_STATUS_INDEX_OUT_OF_RANGE)
        << index;
  }
}

namespace {

std::vector<std::uint8_t> ReadFixture(const std::string& relative) {
  std::ifstream in(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + relative, std::ios::binary);
  return std::vector<std::uint8_t>(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
}

std::vector<std::uint8_t> Save(fm_workbook_t* wb, int32_t format) {
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  EXPECT_EQ(fm_workbook_save(wb, format, &bytes, &len), 0) << fm_get_last_error_message();
  std::vector<std::uint8_t> out;
  if (bytes != nullptr) {
    out.assign(bytes, bytes + len);
    fm_buffer_destroy(bytes);
  }
  return out;
}

/// The record types of an `.xlsb` part, in order, with each record's first payload byte (-1 when empty).
std::vector<std::pair<std::uint32_t, int>> RecordTypes(const std::string& part) {
  std::vector<std::pair<std::uint32_t, int>> out;
  std::size_t i = 0;
  const auto varint = [&part, &i](int bytes) {
    std::uint32_t value = 0;
    for (int k = 0; k < bytes && i < part.size(); ++k) {
      const auto c = static_cast<std::uint8_t>(part[i++]);
      value |= static_cast<std::uint32_t>(c & 0x7FU) << (7 * k);
      if ((c & 0x80U) == 0U) {
        break;
      }
    }
    return value;
  };
  while (i < part.size()) {
    const std::uint32_t type = varint(2);
    const std::uint32_t length = varint(4);
    out.emplace_back(type, length > 0U && i < part.size() ? static_cast<std::uint8_t>(part[i]) : -1);
    i += length;
  }
  return out;
}

/// The first field's BrtBeginPCDFAtbl flag byte in an `.xlsb` cache definition part.
int FirstFieldItemTableFlags(const std::string& part) {
  constexpr std::uint32_t kBeginPcdfAtbl = 189;
  for (const auto& [type, first] : RecordTypes(part)) {
    if (type == kBeginPcdfAtbl) {
      return first;
    }
  }
  return -1;
}

constexpr const char* kCacheXml = "xl/pivotCache/pivotCacheDefinition1.xml";
constexpr const char* kCacheBin = "xl/pivotCache/pivotCacheDefinition1.bin";
constexpr const char* kDataSourceClean = "<xlpdsc:dataSourceClean>1</xlpdsc:dataSourceClean>";

/// Whether an `.xlsb` cache definition part holds the dataSourceClean block (FRT 5153).
bool XlsbDataSourceClean(const std::string& part) {
  constexpr std::uint32_t kDataSourceCleanRecord = 5153;
  for (const auto& [type, first] : RecordTypes(part)) {
    if (type == kDataSourceCleanRecord) {
      return true;
    }
  }
  return false;
}

/// A pivot on PV over Sheet1!K1:L3 (Item / Amt; A 1, B 2), refreshed and projected, as Excel's no-save probe built.
/// With `source_name` the cache names its source (a defined name or a table) instead of the range.
void BuildStaleProbePivot(fm_workbook_t* wb, const char* source_name = nullptr) {
  ASSERT_EQ(fm_sheet_set_text(wb, 0U, 0U, 10U, "Item"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb, 0U, 0U, 11U, "Amt"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb, 0U, 1U, 10U, "A"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb, 0U, 1U, 11U, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_text(wb, 0U, 2U, 10U, "B"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb, 0U, 2U, 11U, 2.0), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb, 0U, &cache_id), 0) << fm_get_last_error_message();
  if (source_name == nullptr) {
    ASSERT_EQ(SetPivotCacheWorksheetSource(wb, cache_id, 1, "K1:L3", "Sheet1", nullptr), 0);
  } else {
    ASSERT_EQ(SetPivotCacheWorksheetSource(wb, cache_id, 1, nullptr, nullptr, source_name), 0)
        << fm_get_last_error_message();
  }
  ASSERT_EQ(fm_workbook_refresh_pivot_cache(wb, cache_id), 0) << fm_get_last_error_message();
  std::size_t sheet = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb, "PV", 1U, &sheet), 0) << fm_get_last_error_message();
  fm_pivot_spec_t spec{sizeof(fm_pivot_spec_t), "PivotS", cache_id, 2U, 0U};
  std::size_t pivot = 0;
  ASSERT_EQ(fm_sheet_add_pivot(wb, sheet, &spec, &pivot), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb, sheet, pivot, 0U, FM_PIVOT_AXIS_ROW), 0) << fm_get_last_error_message();
  fm_pivot_data_field_t data{};
  data.struct_size = sizeof(data);
  data.field_index = 1U;
  data.aggregation = FM_PIVOT_AGGREGATION_SUM;
  std::size_t index = 0;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb, sheet, pivot, &data, &index), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_project_pivot(wb, sheet, pivot), 0) << fm_get_last_error_message();
}

/// Column A of sheet 1 (the pivot sheet) from row 3, as the report shows it.
std::vector<std::string> RowLabels(const std::vector<std::uint8_t>& package) {
  WorkbookGuard wb;
  EXPECT_EQ(fm_workbook_load(package.data(), package.size(), nullptr, &wb.handle), 0) << fm_get_last_error_message();
  std::vector<std::string> out;
  for (std::uint32_t row = 3U; row < 8U && wb.handle != nullptr; ++row) {
    const char* text = nullptr;
    std::int32_t status = 0;
    EXPECT_EQ(fm_sheet_get_display_text(wb.handle, 1U, row, 0U, &text, &status), 0);
    out.emplace_back(text != nullptr ? text : "");
  }
  return out;
}

}  // namespace

// A Region field keeping an unused blank item (<m u="1"/>) besides its text: Excel writes containsBlank for it and
// refuses a save without. Both books refresh on load, so a plain save and one after a refresh derive it again.
TEST(FormulonCApiPivotSaveShape, AKeptUnusedBlankItemKeepsContainsBlank) {
  for (const std::string name : {"P_ne_Bescstar", "P_ne_esctildeB"}) {
    SCOPED_TRACE(name);
    const std::string stem = "xlsb_parity/pivot/caption_equal/" + name;
    const std::vector<std::uint8_t> source = ReadFixture(stem + ".xlsx");
    const std::vector<std::uint8_t> excel_xlsb = ReadFixture(stem + ".xlsb");
    ASSERT_FALSE(source.empty());
    const int excel_flags = FirstFieldItemTableFlags(ExtractZipEntry(excel_xlsb, kCacheBin));
    ASSERT_NE(excel_flags, -1);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_load(source.data(), source.size(), nullptr, &wb.handle), 0) << fm_get_last_error_message();
    for (const bool refresh : {false, true}) {
      SCOPED_TRACE(refresh ? "after refresh" : "plain save");
      if (refresh) {
        std::uint32_t cache_id = 0;
        ASSERT_EQ(fm_workbook_get_pivot_cache_id_at(wb.handle, 0U, &cache_id), 0);
        ASSERT_EQ(fm_workbook_refresh_pivot_cache(wb.handle, cache_id), 0) << fm_get_last_error_message();
      }
      const std::string xml = ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSX), kCacheXml);
      EXPECT_NE(xml.find("<cacheField name=\"Region\" numFmtId=\"0\"><sharedItems containsBlank=\"1\" count=\"15\">"),
                std::string::npos)
          << xml.substr(0, 600);
      EXPECT_EQ(FirstFieldItemTableFlags(ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSB), kCacheBin)),
                excel_flags);
    }
  }
}

// A row inserted inside the source widens its reference and marks it edited, as Excel does: the save keeps the
// records and the report, writes the wider reference, and leaves dataSourceClean out until a refresh.
TEST(FormulonCApiPivotSaveShape, AnInsertInsideTheSourceKeepsTheRecordsAndClearsDataSourceClean) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  BuildStaleProbePivot(wb.handle);
  const std::vector<std::string> report = {"A", "B", "Grand Total", "", ""};
  std::string xml = ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSX), kCacheXml);
  EXPECT_NE(xml.find(kDataSourceClean), std::string::npos) << xml;

  ASSERT_EQ(fm_sheet_insert_rows(wb.handle, 0U, 1U, 1U), 0) << fm_get_last_error_message();
  const std::vector<std::uint8_t> xlsx = Save(wb.handle, FM_WORKBOOK_FORMAT_XLSX);
  const std::vector<std::uint8_t> xlsb = Save(wb.handle, FM_WORKBOOK_FORMAT_XLSB);
  xml = ExtractZipEntry(xlsx, kCacheXml);
  EXPECT_NE(xml.find("ref=\"K1:L4\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("recordCount=\"2\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("refreshOnLoad"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("dataSourceClean"), std::string::npos) << xml;
  EXPECT_FALSE(XlsbDataSourceClean(ExtractZipEntry(xlsb, kCacheBin)));
  EXPECT_EQ(RowLabels(xlsx), report);
  EXPECT_EQ(RowLabels(xlsb), report);

  // The inserted row filled, a refresh reads three records and marks the source clean again.
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 1U, 10U, "C"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0U, 1U, 11U, 3.0), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_id_at(wb.handle, 0U, &cache_id), 0);
  ASSERT_EQ(fm_workbook_refresh_pivot_cache(wb.handle, cache_id), 0) << fm_get_last_error_message();
  xml = ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSX), kCacheXml);
  EXPECT_NE(xml.find(kDataSourceClean), std::string::npos) << xml;
  EXPECT_TRUE(XlsbDataSourceClean(ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSB), kCacheBin)));
}

// Excel marks the source edited for any row or column inserted or deleted on its sheet, wherever it goes, and keeps
// it clean for a structural edit on another sheet.
TEST(FormulonCApiPivotSaveShape, AnyStructuralEditOnTheSourceSheetClearsDataSourceClean) {
  const struct {
    const char* label;
    std::size_t sheet;
    bool rows;
    bool remove;
    std::uint32_t at;
    const char* ref;
    bool clean;
  } kCases[] = {
      {"row above the source", 0U, true, false, 0U, "K2:L4", false},
      {"column to the left", 0U, false, false, 0U, "L1:M3", false},
      {"row below the source", 0U, true, false, 10U, "K1:L3", false},
      {"column to the right", 0U, false, false, 20U, "K1:L3", false},
      {"empty row below deleted", 0U, true, true, 10U, "K1:L3", false},
      {"row on the pivot sheet", 1U, true, false, 0U, "K1:L3", true},
  };
  for (const auto& c : kCases) {
    SCOPED_TRACE(c.label);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    BuildStaleProbePivot(wb.handle);
    if (c.rows) {
      ASSERT_EQ(c.remove ? fm_sheet_delete_rows(wb.handle, c.sheet, c.at, c.at)
                         : fm_sheet_insert_rows(wb.handle, c.sheet, c.at, c.at),
                0)
          << fm_get_last_error_message();
    } else {
      ASSERT_EQ(fm_sheet_insert_columns(wb.handle, c.sheet, c.at, c.at), 0) << fm_get_last_error_message();
    }
    const std::string xml = ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSX), kCacheXml);
    EXPECT_NE(xml.find(std::string("ref=\"") + c.ref + "\""), std::string::npos) << xml;
    EXPECT_EQ(xml.find(kDataSourceClean) != std::string::npos, c.clean) << xml;
    EXPECT_EQ(XlsbDataSourceClean(ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSB), kCacheBin)), c.clean);
  }
}

// A value typed into a range or defined-name source, or an AutoFilter set on it, marks the source edited in both
// containers; a refresh marks it clean again.
TEST(FormulonCApiPivotSaveShape, AValueEditOrAutoFilterOnTheSourceClearsDataSourceCleanUntilRefresh) {
  enum class Source { kRange, kName };
  enum class Edit { kText, kNumber, kBlank, kAutoFilter, kAutoFilterXml };
  for (const Source source : {Source::kRange, Source::kName}) {
    for (const Edit edit : {Edit::kText, Edit::kNumber, Edit::kBlank, Edit::kAutoFilter, Edit::kAutoFilterXml}) {
      SCOPED_TRACE(static_cast<int>(source) * 10 + static_cast<int>(edit));
      WorkbookGuard wb;
      ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
      const char* source_name = nullptr;
      if (source == Source::kName) {
        fm_defined_name_t name{};
        name.struct_size = sizeof(name);
        name.name = "Src";
        name.formula = "=Sheet1!$K$1:$L$3";
        name.local_sheet_id = FM_DEFINED_NAME_SCOPE_WORKBOOK;
        ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &name), 0) << fm_get_last_error_message();
        source_name = "Src";
      }
      BuildStaleProbePivot(wb.handle, source_name);
      ASSERT_NE(ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSX), kCacheXml).find(kDataSourceClean),
                std::string::npos);
      switch (edit) {
        case Edit::kText:
          ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 1U, 10U, "Z"), 0);
          break;
        case Edit::kNumber:
          ASSERT_EQ(fm_sheet_set_number(wb.handle, 0U, 1U, 11U, 7.0), 0);
          break;
        case Edit::kBlank:
          ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0U, 2U, 11U), 0);
          break;
        case Edit::kAutoFilter: {
          fm_auto_filter_t filter{};
          filter.struct_size = sizeof(filter);
          filter.range = {0U, 10U, 2U, 11U};
          ASSERT_EQ(fm_sheet_set_auto_filter(wb.handle, 0U, &filter), 0) << fm_get_last_error_message();
          break;
        }
        case Edit::kAutoFilterXml:
          ASSERT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0U, "<autoFilter ref=\"K1:L3\"/>"), 0)
              << fm_get_last_error_message();
          break;
      }
      EXPECT_EQ(ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSX), kCacheXml).find(kDataSourceClean),
                std::string::npos);
      EXPECT_FALSE(XlsbDataSourceClean(ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSB), kCacheBin)));

      std::uint32_t cache_id = 0;
      ASSERT_EQ(fm_workbook_get_pivot_cache_id_at(wb.handle, 0U, &cache_id), 0);
      ASSERT_EQ(fm_workbook_refresh_pivot_cache(wb.handle, cache_id), 0) << fm_get_last_error_message();
      EXPECT_NE(ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSX), kCacheXml).find(kDataSourceClean),
                std::string::npos);
      EXPECT_TRUE(XlsbDataSourceClean(ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSB), kCacheBin)));
    }
  }
}

// Every load starts clean: a cache saved without the extension and one saved with it are each written with it once.
TEST(FormulonCApiPivotSaveShape, ALoadedCacheIsWrittenClean) {
  for (const char* relative : {"pivot_aggregation/C9__sum.xlsx", "xlsb_parity/pivot/caption_equal/P_ne_B.xlsx"}) {
    SCOPED_TRACE(relative);
    const std::vector<std::uint8_t> source = ReadFixture(relative);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_load(source.data(), source.size(), nullptr, &wb.handle), 0) << fm_get_last_error_message();
    const std::string xml = ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSX), kCacheXml);
    EXPECT_EQ(CountOccurrences(xml, kDataSourceClean), 1U) << xml;
    EXPECT_TRUE(XlsbDataSourceClean(ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSB), kCacheBin)));
  }
}

// Insert Cells on the source's sheet marks a range source edited, inside the range or not. A defined-name source is
// judged against the name's range: a row inserted inside it marks it, one elsewhere does not. A table source is
// never marked.
TEST(FormulonCApiPivotSaveShape, DataSourceCleanFollowsHowTheSourceIsGiven) {
  enum class Source { kRange, kName, kTable };
  const struct {
    const char* label;
    Source source;
    bool band;
    fm_cell_rect_t rect;  ///< The band for Insert Cells, or rows first_row..last_row for a row insert.
    bool clean;
  } kCases[] = {
      {"cells inserted inside the range", Source::kRange, true, {1U, 10U, 1U, 11U}, false},
      {"cells inserted at H2:I2", Source::kRange, true, {1U, 7U, 1U, 8U}, false},
      {"row inserted inside the name", Source::kName, false, {1U, 0U, 1U, 0U}, false},
      {"row inserted below the name", Source::kName, false, {10U, 0U, 10U, 0U}, true},
      {"row inserted inside the table", Source::kTable, false, {1U, 0U, 1U, 0U}, true},
  };
  for (const auto& c : kCases) {
    SCOPED_TRACE(c.label);
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const char* source_name = nullptr;
    if (c.source == Source::kName) {
      fm_defined_name_t name{};
      name.struct_size = sizeof(name);
      name.name = "Src";
      name.formula = "=Sheet1!$K$1:$L$3";
      name.local_sheet_id = FM_DEFINED_NAME_SCOPE_WORKBOOK;
      ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &name), 0) << fm_get_last_error_message();
      source_name = "Src";
    }
    if (c.source == Source::kTable) {
      ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 0U, 10U, "Item"), 0);
      ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 0U, 11U, "Amt"), 0);
      fm_table_spec_t spec{};
      spec.struct_size = sizeof(spec);
      spec.range = {0U, 10U, 2U, 11U};
      spec.name = "Tbl";
      std::size_t table = 0;
      ASSERT_EQ(fm_sheet_add_table(wb.handle, 0U, &spec, &table), 0) << fm_get_last_error_message();
      source_name = "Tbl";
    }
    BuildStaleProbePivot(wb.handle, source_name);
    if (c.band) {
      ASSERT_EQ(fm_sheet_insert_cells(wb.handle, 0U, &c.rect, FM_CELL_SHIFT_DOWN), 0) << fm_get_last_error_message();
    } else {
      ASSERT_EQ(fm_sheet_insert_rows(wb.handle, 0U, c.rect.first_row, c.rect.last_row), 0)
          << fm_get_last_error_message();
    }
    const std::string xml = ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSX), kCacheXml);
    EXPECT_EQ(xml.find(kDataSourceClean) != std::string::npos, c.clean) << xml;
    EXPECT_EQ(XlsbDataSourceClean(ExtractZipEntry(Save(wb.handle, FM_WORKBOOK_FORMAT_XLSB), kCacheBin)), c.clean);
  }
}
