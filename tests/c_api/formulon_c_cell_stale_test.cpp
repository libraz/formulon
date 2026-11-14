//
// fm_sheet_is_cell_stale: a formula cell whose value has not been computed
// since its inputs changed, under manual (or automatic-except-tables)
// calculation; always 0 under automatic. These pin the engine's rule; which
// cells Excel draws with stale-value formatting is not measured.

#include <cstdint>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_destroy(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

bool Stale(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col, std::size_t sheet = 0) {
  int32_t stale = -1;
  EXPECT_EQ(fm_sheet_is_cell_stale(wb, sheet, row, col, &stale), FM_STATUS_OK);
  return stale == 1;
}

// A1 = 1, B1 = A1, C1 = B1, D1 = 1+1, recalculated, in manual mode.
void ManualChain(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_MANUAL), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0U, 0U, 1.0), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0U, 1U, "=A1"), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0U, 2U, "=B1"), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0U, 3U, "=1+1"), FM_STATUS_OK);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
}

}  // namespace

// An edit leaves every formula that reads it, directly or through others, stale.
TEST(FormulonCApiCellStale, AnEditMarksTheFormulasThatReadIt) {
  WorkbookGuard wb;
  ManualChain(wb);
  EXPECT_FALSE(Stale(wb.handle, 0U, 1U));
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0U, 0U, 2.0), FM_STATUS_OK);
  EXPECT_FALSE(Stale(wb.handle, 0U, 0U));  // the edited value is not a formula
  EXPECT_TRUE(Stale(wb.handle, 0U, 1U));
  EXPECT_TRUE(Stale(wb.handle, 0U, 2U));
  EXPECT_FALSE(Stale(wb.handle, 0U, 3U));  // reads nothing that changed
}

// A volatile formula is not stale on its own; one that reads an edited cell is.
TEST(FormulonCApiCellStale, AVolatileIsStaleOnlyThroughWhatItReads) {
  WorkbookGuard wb;
  ManualChain(wb);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1U, 0U, "=NOW()"), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1U, 1U, "=RAND()+A1"), FM_STATUS_OK);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 5U, 5U, 9.0), FM_STATUS_OK);  // unrelated edit
  EXPECT_FALSE(Stale(wb.handle, 1U, 0U));
  EXPECT_FALSE(Stale(wb.handle, 1U, 1U));
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0U, 0U, 3.0), FM_STATUS_OK);
  EXPECT_FALSE(Stale(wb.handle, 1U, 0U));
  EXPECT_TRUE(Stale(wb.handle, 1U, 1U));
}

// Every formula of a loaded manual-mode workbook is stale until its first recalc.
TEST(FormulonCApiCellStale, ALoadedManualBookIsStaleUntilItsFirstRecalc) {
  WorkbookGuard wb;
  ManualChain(wb);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), FM_STATUS_OK);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), FM_STATUS_OK);
  EXPECT_TRUE(Stale(loaded.handle, 0U, 1U));
  EXPECT_TRUE(Stale(loaded.handle, 0U, 3U));
  ASSERT_EQ(fm_workbook_recalc(loaded.handle), FM_STATUS_OK);
  EXPECT_FALSE(Stale(loaded.handle, 0U, 1U));
  EXPECT_FALSE(Stale(loaded.handle, 0U, 3U));
}

// A recalc clears what it computed: a viewport recalc its closure, a full recalc everything.
TEST(FormulonCApiCellStale, ARecalcClearsWhatItComputed) {
  WorkbookGuard wb;
  ManualChain(wb);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 9U, 0U, "=A1*10"), FM_STATUS_OK);  // A10, far from the viewport
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0U, 0U, 2.0), FM_STATUS_OK);
  const fm_cell_rect_t viewport{0U, 0U, 0U, 2U};  // A1:C1
  ASSERT_EQ(fm_sheet_recalc_viewport(wb.handle, 0, &viewport, nullptr), FM_STATUS_OK);
  EXPECT_FALSE(Stale(wb.handle, 0U, 1U));
  EXPECT_FALSE(Stale(wb.handle, 0U, 2U));
  EXPECT_TRUE(Stale(wb.handle, 9U, 0U));
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_FALSE(Stale(wb.handle, 9U, 0U));
}

// A formula entered in manual mode is not computed on entry, so it is stale.
TEST(FormulonCApiCellStale, AFormulaEnteredInManualModeIsStale) {
  WorkbookGuard wb;
  ManualChain(wb);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2U, 0U, "=A1+1"), FM_STATUS_OK);
  EXPECT_TRUE(Stale(wb.handle, 2U, 0U));
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_FALSE(Stale(wb.handle, 2U, 0U));
}

TEST(FormulonCApiCellStale, AutomaticCalculationIsNeverStale) {
  WorkbookGuard wb;
  ManualChain(wb);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0U, 0U, 2.0), FM_STATUS_OK);
  ASSERT_TRUE(Stale(wb.handle, 0U, 1U));
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_AUTO), FM_STATUS_OK);
  EXPECT_FALSE(Stale(wb.handle, 0U, 1U));
}

// A spill member is not a formula cell; its anchor carries the state.
TEST(FormulonCApiCellStale, ASpillMemberReadsNotStale) {
  WorkbookGuard wb;
  ManualChain(wb);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 4U, 0U, "=SEQUENCE(3)+A1"), FM_STATUS_OK);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0U, 0U, 2.0), FM_STATUS_OK);
  EXPECT_TRUE(Stale(wb.handle, 4U, 0U));
  EXPECT_FALSE(Stale(wb.handle, 5U, 0U));
}

TEST(FormulonCApiCellStale, BoundaryArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  int32_t stale = -1;
  EXPECT_EQ(fm_sheet_is_cell_stale(nullptr, 0, 0U, 0U, &stale), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_is_cell_stale(wb.handle, 0, 0U, 0U, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_is_cell_stale(wb.handle, 4, 0U, 0U, &stale), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_is_cell_stale(wb.handle, 0, 1048576U, 0U, &stale), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_sheet_is_cell_stale(wb.handle, 0, 0U, 16384U, &stale), FM_STATUS_INVALID_ARGUMENT);
  ASSERT_EQ(fm_sheet_is_cell_stale(wb.handle, 0, 7U, 7U, &stale), FM_STATUS_OK);
  EXPECT_EQ(stale, 0);  // an empty cell
}
