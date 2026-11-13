// What keeps a spill or a PivotTable report from being drawn, through the
// `cause` of `fm_spill_info_t`: `fm_sheet_get_spill` and
// `fm_sheet_get_pivot_spill` over formulas and reports that are drawn,
// occupied, merged, past the sheet edge or absent, and over blocked states
// read back from a saved file.

#include <cstddef>
#include <cstdint>

#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"
#include "sheet.h"
#include "value.h"

namespace {

constexpr std::uint32_t kLastRow = formulon::Sheet::kMaxRows - 1U;
constexpr std::uint32_t kLastCol = formulon::Sheet::kMaxCols - 1U;

// `BuildScratchPivot` draws its report over D1:E4; one more Region item grows it into D5.
constexpr std::uint32_t kPivotRow = 0U;
constexpr std::uint32_t kPivotCol = 3U;
constexpr fm_cell_rect_t kDrawn{0U, 3U, 3U, 4U};
constexpr fm_cell_rect_t kGrown{0U, 3U, 4U, 4U};

struct SpillRead {
  std::int32_t present = -1;
  fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
};

SpillRead CellSpill(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  SpillRead read;
  EXPECT_EQ(fm_sheet_get_spill(wb, 0U, row, col, &read.present, &read.info), 0) << fm_get_last_error_message();
  return read;
}

SpillRead PivotSpill(const fm_workbook_t* wb) {
  SpillRead read;
  EXPECT_EQ(fm_sheet_get_pivot_spill(wb, 0U, 0U, &read.present, &read.info), 0) << fm_get_last_error_message();
  return read;
}

// The anchor is the region's top-left unless `anchor` says otherwise (a report with page fields).
void ExpectSpill(const SpillRead& read, std::int32_t blocked, std::int32_t cause, const fm_cell_rect_t& rect,
                 const fm_cell_address_t* anchor = nullptr) {
  EXPECT_EQ(read.present, 1);
  EXPECT_EQ(read.info.anchor.row, anchor != nullptr ? anchor->row : rect.first_row);
  EXPECT_EQ(read.info.anchor.col, anchor != nullptr ? anchor->col : rect.first_col);
  EXPECT_EQ(read.info.blocked, blocked);
  EXPECT_EQ(read.info.cause, cause);
  EXPECT_EQ(read.info.rect.first_row, rect.first_row);
  EXPECT_EQ(read.info.rect.first_col, rect.first_col);
  EXPECT_EQ(read.info.rect.last_row, rect.last_row);
  EXPECT_EQ(read.info.rect.last_col, rect.last_col);
}

void ExpectNothing(const SpillRead& read) {
  EXPECT_EQ(read.present, 0);
  EXPECT_EQ(read.info.blocked, 0);
  EXPECT_EQ(read.info.cause, FM_SPILL_CAUSE_ABSENT);
}

void ExpectBlockedPivot(const fm_workbook_t* wb, std::int32_t cause, const fm_cell_rect_t& rect) {
  ExpectSpill(PivotSpill(wb), 1, cause, rect);
  ExpectSpill(CellSpill(wb, rect.first_row, rect.first_col), 1, cause, rect);
}

void Build(WorkbookGuard& wb, std::uint32_t* cache_id = nullptr) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(BuildScratchPivot(wb.handle, cache_id, nullptr), 0) << fm_get_last_error_message();
}

// Adds a third Region item with a record, which leaves the report pending until it is projected.
void AddItem(fm_workbook_t* wb, std::uint32_t cache_id) {
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb, cache_id, 0U, "West", ignored_index()), 0);
  std::size_t record = 99;
  ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb, cache_id, &record), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb, cache_id, record, 0U, 2.0), 0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb, cache_id, record, 1U, 500.0), 0);
  // The source carries the record too, so a reload's refresh reads the same cache.
  ASSERT_EQ(fm_sheet_set_text(wb, 0U, 5U, 0U, "West"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb, 0U, 5U, 1U, 500.0), 0);
  ASSERT_EQ(SetPivotCacheWorksheetSource(wb, cache_id, /*present=*/1, "A1:B6", "Sheet1", nullptr), 0);
}

void Grow(fm_workbook_t* wb, std::uint32_t cache_id) {
  AddItem(wb, cache_id);
  ASSERT_EQ(fm_sheet_project_pivot(wb, 0U, 0U), 0) << fm_get_last_error_message();
}

void SaveAndLoad(fm_workbook_t* wb, WorkbookGuard& loaded) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0) << fm_get_last_error_message();
}

bool IsSpillError(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  fm_value_t value{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0U, row, col, &value), 0) << fm_get_last_error_message();
  return value.kind == FM_VALUE_KIND_ERROR &&
         value.u.error_code == static_cast<std::int32_t>(formulon::ErrorCode::Spill);
}

}  // namespace

// ---------------------------------------------------------------------------
// Formulas
// ---------------------------------------------------------------------------

TEST(FormulonCApiSpillCause, DrawnFormulaSpillHasNoCause) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, 0U, 0U, "=SEQUENCE(3)"), 0);
  const fm_cell_rect_t drawn{0U, 0U, 2U, 0U};
  ExpectSpill(CellSpill(wb.handle, 0U, 0U), 0, FM_SPILL_CAUSE_ABSENT, drawn);
  ExpectSpill(CellSpill(wb.handle, 2U, 0U), 0, FM_SPILL_CAUSE_ABSENT, drawn);
}

TEST(FormulonCApiSpillCause, AnchorIsTheFormulaCellAtAnchorAndPhantom) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, 1U, 2U, "=SEQUENCE(3,2)"), 0);  // C2:D4
  const fm_cell_rect_t drawn{1U, 2U, 3U, 3U};
  const fm_cell_address_t anchor{1U, 2U};
  ExpectSpill(CellSpill(wb.handle, 1U, 2U), 0, FM_SPILL_CAUSE_ABSENT, drawn, &anchor);
  ExpectSpill(CellSpill(wb.handle, 3U, 3U), 0, FM_SPILL_CAUSE_ABSENT, drawn, &anchor);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, 6U, 4U, "=SEQUENCE(3)"), 0);  // E7 meets E9
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 8U, 4U, "x"), 0);
  const fm_cell_address_t blocked_anchor{6U, 4U};
  ExpectSpill(CellSpill(wb.handle, 6U, 4U), 1, FM_SPILL_CAUSE_OCCUPIED, fm_cell_rect_t{6U, 4U, 8U, 4U},
              &blocked_anchor);
}

TEST(FormulonCApiSpillCause, OccupiedAndMergedFormulaSpills) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, 0U, 2U, "=SEQUENCE(3)"), 0);  // C1 meets C3
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 2U, 2U, "x"), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0U, record_arg(fm_cell_rect_t{1U, 3U, 1U, 4U}), nullptr), 0);  // D2:E2
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, 0U, 3U, "=SEQUENCE(3)"), 0);                             // D1 meets it
  ExpectSpill(CellSpill(wb.handle, 0U, 2U), 1, FM_SPILL_CAUSE_OCCUPIED, fm_cell_rect_t{0U, 2U, 2U, 2U});
  ExpectSpill(CellSpill(wb.handle, 0U, 3U), 1, FM_SPILL_CAUSE_MERGED, fm_cell_rect_t{0U, 3U, 2U, 3U});
  // Only the anchor reports the attempt.
  ExpectNothing(CellSpill(wb.handle, 1U, 2U));
  ExpectNothing(CellSpill(wb.handle, 2U, 2U));
}

TEST(FormulonCApiSpillCause, BeyondEdgeIsReportedClipped) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, kLastRow - 1U, 0U, "=SEQUENCE(5)"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, 0U, kLastCol - 1U, "=SEQUENCE(1,5)"), 0);
  // An occupant inside the grid does not change the cause of a range that runs past the edge.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, kLastRow - 3U, 1U, "=SEQUENCE(5)"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, kLastRow - 2U, 1U, "x"), 0);
  ExpectSpill(CellSpill(wb.handle, kLastRow - 1U, 0U), 1, FM_SPILL_CAUSE_BEYOND_EDGE,
              fm_cell_rect_t{kLastRow - 1U, 0U, kLastRow, 0U});
  ExpectSpill(CellSpill(wb.handle, 0U, kLastCol - 1U), 1, FM_SPILL_CAUSE_BEYOND_EDGE,
              fm_cell_rect_t{0U, kLastCol - 1U, 0U, kLastCol});
  ExpectSpill(CellSpill(wb.handle, kLastRow - 3U, 1U), 1, FM_SPILL_CAUSE_BEYOND_EDGE,
              fm_cell_rect_t{kLastRow - 3U, 1U, kLastRow, 1U});
  ExpectNothing(CellSpill(wb.handle, kLastRow, 0U));
}

TEST(FormulonCApiSpillCause, CellInNoSpillReportsNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, 0U, 0U, "=1+1"), 0);
  ExpectNothing(CellSpill(wb.handle, 0U, 0U));
  ExpectNothing(CellSpill(wb.handle, 5U, 5U));
}

// ---------------------------------------------------------------------------
// PivotTable reports
// ---------------------------------------------------------------------------

TEST(FormulonCApiSpillCause, DrawnPivotAnchorReportsItsExtent) {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  Build(wb, &cache_id);
  ExpectSpill(PivotSpill(wb.handle), 0, FM_SPILL_CAUSE_ABSENT, kDrawn);
  ExpectSpill(CellSpill(wb.handle, kPivotRow, kPivotCol), 0, FM_SPILL_CAUSE_ABSENT, kDrawn);
  // The report's other cells are ordinary cells.
  ExpectNothing(CellSpill(wb.handle, kPivotRow + 1U, kPivotCol));
  ExpectNothing(CellSpill(wb.handle, kDrawn.last_row, kDrawn.last_col));

  // A cache change leaves the report pending: it still reads the range it was drawn over.
  AddItem(wb.handle, cache_id);
  std::int32_t pending = 0;
  ASSERT_EQ(fm_sheet_is_pivot_projection_pending(wb.handle, 0U, 0U, &pending), 0);
  ASSERT_EQ(pending, 1);
  ExpectSpill(PivotSpill(wb.handle), 0, FM_SPILL_CAUSE_ABSENT, kDrawn);
  ASSERT_EQ(fm_sheet_project_pivot(wb.handle, 0U, 0U), 0) << fm_get_last_error_message();
  ExpectSpill(PivotSpill(wb.handle), 0, FM_SPILL_CAUSE_ABSENT, kGrown);
}

// A report with a page field draws its page rows above the anchor; `rect` includes them, `anchor` does not.
TEST(FormulonCApiSpillCause, PivotWithPageFieldAnchorsBelowThePageRows) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0, &cache), 0);
  ASSERT_EQ(SetPivotCacheWorksheetSource(wb.handle, cache, 1, "A1:C3", "Sheet1", nullptr), 0);
  std::size_t fields[3] = {99, 99, 99};
  const char* const names[3] = {"Region", "Kind", "Amount"};
  for (std::size_t i = 0; i < 3U; ++i) {
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache, names[i], &fields[i]), 0);
  }
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache, fields[0], "North", ignored_index()),
            0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache, fields[1], "A", ignored_index()), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache, fields[1], "B", ignored_index()), 0);
  for (std::size_t i = 0; i < 2U; ++i) {
    std::size_t record = 99;
    ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache, &record), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache, record, fields[0], 0.0), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache, record, fields[1], static_cast<double>(i)),
              0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache, record, fields[2], 10.0), 0);
  }
  std::size_t pivot = 99;
  ASSERT_EQ(AddPivot(wb.handle, 0, "PT", cache, /*anchor_row=*/4U, /*anchor_col=*/3U, &pivot), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0U, pivot, fields[1], FM_PIVOT_AXIS_ROW), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0U, pivot, fields[2], FM_PIVOT_AXIS_VALUE), 0);
  fm_pivot_data_field_t sum = FM_STRUCT_INIT(fm_pivot_data_field_t);
  sum.name = "Sum of Amount";
  sum.field_index = static_cast<std::uint32_t>(fields[2]);
  sum.aggregation = FM_PIVOT_AGGREGATION_SUM;
  sum.number_format = "";
  sum.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0U, pivot, &sum, ignored_index()), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0U, pivot, fields[0], FM_PIVOT_AXIS_PAGE), 0)
      << fm_get_last_error_message();

  const SpillRead read = PivotSpill(wb.handle);
  EXPECT_EQ(read.present, 1);
  EXPECT_EQ(read.info.anchor.row, 4U);
  EXPECT_EQ(read.info.anchor.col, 3U);
  EXPECT_LT(read.info.rect.first_row, 4U);
  const SpillRead at_anchor = CellSpill(wb.handle, 4U, 3U);
  EXPECT_EQ(at_anchor.present, 1);
  EXPECT_EQ(at_anchor.info.anchor.row, 4U);
  EXPECT_EQ(at_anchor.info.rect.first_row, read.info.rect.first_row);
}

TEST(FormulonCApiSpillCause, OccupiedPivot) {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  Build(wb, &cache_id);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, kGrown.last_row, kGrown.last_col, "x"), 0);
  Grow(wb.handle, cache_id);
  ExpectBlockedPivot(wb.handle, FM_SPILL_CAUSE_OCCUPIED, kGrown);
  ExpectNothing(CellSpill(wb.handle, kPivotRow + 1U, kPivotCol));
}

TEST(FormulonCApiSpillCause, MergedPivot) {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  Build(wb, &cache_id);
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0U, record_arg(fm_cell_rect_t{4U, 3U, 4U, 4U}), nullptr), 0);
  Grow(wb.handle, cache_id);
  ExpectBlockedPivot(wb.handle, FM_SPILL_CAUSE_MERGED, kGrown);
}

TEST(FormulonCApiSpillCause, PivotPastTheEdgeIsReportedClipped) {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  Build(wb, &cache_id);
  const std::uint32_t row = kLastRow - (kDrawn.last_row - kDrawn.first_row);
  ASSERT_EQ(fm_sheet_set_pivot_anchor(wb.handle, 0U, 0U, row, kPivotCol), 0) << fm_get_last_error_message();
  Grow(wb.handle, cache_id);
  EXPECT_TRUE(IsSpillError(wb.handle, row, kPivotCol));
  ExpectBlockedPivot(wb.handle, FM_SPILL_CAUSE_BEYOND_EDGE, fm_cell_rect_t{row, kPivotCol, kLastRow, kDrawn.last_col});
}

TEST(FormulonCApiSpillCause, LoadedDrawnPivotReportsItsLocation) {
  WorkbookGuard wb;
  Build(wb);
  FillScratchSource(wb.handle);
  WorkbookGuard loaded;
  SaveAndLoad(wb.handle, loaded);
  // Not laid out since the load, the report covers the location it was read with.
  ExpectSpill(PivotSpill(loaded.handle), 0, FM_SPILL_CAUSE_ABSENT, kDrawn);
  ExpectSpill(CellSpill(loaded.handle, kPivotRow, kPivotCol), 0, FM_SPILL_CAUSE_ABSENT, kDrawn);
  ExpectNothing(CellSpill(loaded.handle, kPivotRow + 1U, kPivotCol));
}

// ---------------------------------------------------------------------------
// Blocked states read from a file
// ---------------------------------------------------------------------------

TEST(FormulonCApiSpillCause, ManualCalcReportsTheLoadedRange) {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  Build(wb, &cache_id);
  FillScratchSource(wb.handle);
  // A11 meets A13; a formula near the last row runs past it; the report meets a merge.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, 10U, 0U, "=SEQUENCE(3)"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 12U, 0U, "x"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0U, kLastRow - 1U, 0U, "=SEQUENCE(5)"), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0U, record_arg(fm_cell_rect_t{4U, 3U, 4U, 4U}), nullptr), 0);
  Grow(wb.handle, cache_id);
  // Refreshed from its source, the cache is saved as Excel saves it, with nothing to refresh on load.
  ASSERT_EQ(fm_workbook_refresh_pivot_cache(wb.handle, cache_id), 0) << fm_get_last_error_message();
  ExpectBlockedPivot(wb.handle, FM_SPILL_CAUSE_MERGED, kGrown);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_MANUAL), 0);

  WorkbookGuard loaded;
  SaveAndLoad(wb.handle, loaded);
  std::int32_t mode = -1;
  ASSERT_EQ(fm_workbook_get_calc_mode(loaded.handle, &mode), 0);
  ASSERT_EQ(mode, FM_CALC_MODE_MANUAL);
  // Before a recalc, each blocked state is what the file recorded. The file's reason for the report does not tell a
  // merge from another occupant, and a formula past the edge records its anchor alone.
  EXPECT_TRUE(IsSpillError(loaded.handle, 10U, 0U));
  EXPECT_TRUE(IsSpillError(loaded.handle, kPivotRow, kPivotCol));
  ExpectSpill(CellSpill(loaded.handle, 10U, 0U), 1, FM_SPILL_CAUSE_OCCUPIED, fm_cell_rect_t{10U, 0U, 12U, 0U});
  ExpectSpill(CellSpill(loaded.handle, kLastRow - 1U, 0U), 1, FM_SPILL_CAUSE_BEYOND_EDGE,
              fm_cell_rect_t{kLastRow - 1U, 0U, kLastRow - 1U, 0U});
  ExpectBlockedPivot(loaded.handle, FM_SPILL_CAUSE_OCCUPIED, kGrown);

  // A recalc lays everything out again.
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0) << fm_get_last_error_message();
  ExpectSpill(CellSpill(loaded.handle, 10U, 0U), 1, FM_SPILL_CAUSE_OCCUPIED, fm_cell_rect_t{10U, 0U, 12U, 0U});
  ExpectSpill(CellSpill(loaded.handle, kLastRow - 1U, 0U), 1, FM_SPILL_CAUSE_BEYOND_EDGE,
              fm_cell_rect_t{kLastRow - 1U, 0U, kLastRow, 0U});
  ExpectBlockedPivot(loaded.handle, FM_SPILL_CAUSE_MERGED, kGrown);
}

TEST(FormulonCApiSpillCause, ReportWhoseCacheHasNoRecordsKeepsTheLoadedRange) {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  Build(wb, &cache_id);
  FillScratchSource(wb.handle);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, kGrown.last_row, kGrown.last_col, "x"), 0);
  Grow(wb.handle, cache_id);
  fm_pivot_cache_options_patch_t options = FM_STRUCT_INIT(fm_pivot_cache_options_patch_t);
  options.save_data_op = FM_PATCH_OP_SET;
  options.save_data = 0;
  ASSERT_EQ(fm_workbook_update_pivot_cache_options(wb.handle, cache_id, &options), 0) << fm_get_last_error_message();

  WorkbookGuard loaded;
  SaveAndLoad(wb.handle, loaded);
  ExpectBlockedPivot(loaded.handle, FM_SPILL_CAUSE_OCCUPIED, kGrown);
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), 0) << fm_get_last_error_message();
  ExpectBlockedPivot(loaded.handle, FM_SPILL_CAUSE_OCCUPIED, kGrown);
}

TEST(FormulonCApiSpillCause, OlderRecordSizeIsRejected) {
  WorkbookGuard wb;
  Build(wb);
  std::int32_t present = -1;
  fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
  info.struct_size = static_cast<std::uint32_t>(offsetof(fm_spill_info_t, cause));
  EXPECT_EQ(fm_sheet_get_spill(wb.handle, 0U, kPivotRow, kPivotCol, &present, &info), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_get_pivot_spill(wb.handle, 0U, 0U, &present, &info), FM_STATUS_INVALID_ARGUMENT);
}
