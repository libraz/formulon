//
// fm_sheet_get_circular_reference_count / _at: the cells of a sheet in an
// unresolved cycle, row-major, against Mac Excel 16.115's
// Worksheet.CircularReference. Excel names one member per sheet and which one
// is not determined by the book, so every member is listed; a cell that only
// reads a cycle is never listed, a cross-sheet cycle is listed on each sheet
// for that sheet's members, and iterative calculation lists nothing.

#include <cstdint>
#include <utility>
#include <vector>

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

using Cells = std::vector<std::pair<std::uint32_t, std::uint32_t>>;

Cells Circular(const fm_workbook_t* wb, std::size_t sheet) {
  std::size_t count = 99U;
  EXPECT_EQ(fm_sheet_get_circular_reference_count(wb, sheet, &count), FM_STATUS_OK);
  Cells out;
  for (std::size_t i = 0; i < count; ++i) {
    fm_cell_address_t cell{};
    EXPECT_EQ(fm_sheet_get_circular_reference_at(wb, sheet, i, &cell), FM_STATUS_OK);
    out.emplace_back(cell.row, cell.col);
  }
  return out;
}

void Formula(fm_workbook_t* wb, std::size_t sheet, std::uint32_t row, std::uint32_t col, const char* text) {
  ASSERT_EQ(fm_sheet_set_formula(wb, sheet, row, col, text), FM_STATUS_OK) << text;
}

}  // namespace

TEST(FormulonCApiCircularReference, APairListsBothMembersRowMajor) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  Formula(wb.handle, 0, 1U, 3U, "=A1");  // D2
  Formula(wb.handle, 0, 0U, 0U, "=D2");  // A1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Circular(wb.handle, 0), (Cells{{0U, 0U}, {1U, 3U}}));
}

TEST(FormulonCApiCircularReference, ASelfReferenceIsItsOwnCycle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  Formula(wb.handle, 0, 1U, 1U, "=B2+1");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Circular(wb.handle, 0), (Cells{{1U, 1U}}));
}

// A cell that only reads a cycle (c7 C1, d9 A1) is never reported.
TEST(FormulonCApiCircularReference, ACellThatOnlyReadsACycleIsNotListed) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  Formula(wb.handle, 0, 0U, 0U, "=B1");    // A1 reads the cycle
  Formula(wb.handle, 0, 0U, 1U, "=C1");    // B1
  Formula(wb.handle, 0, 0U, 2U, "=B1");    // C1
  Formula(wb.handle, 0, 0U, 3U, "=B1+1");  // D1 reads the cycle
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Circular(wb.handle, 0), (Cells{{0U, 1U}, {0U, 2U}}));
}

TEST(FormulonCApiCircularReference, TwoCyclesOnOneSheetAreBothListed) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  Formula(wb.handle, 0, 0U, 0U, "=B1");
  Formula(wb.handle, 0, 0U, 1U, "=A1");
  Formula(wb.handle, 0, 3U, 3U, "=E4");
  Formula(wb.handle, 0, 3U, 4U, "=D4");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Circular(wb.handle, 0), (Cells{{0U, 0U}, {0U, 1U}, {3U, 3U}, {3U, 4U}}));
}

// Each sheet lists its own members of a cycle that crosses sheets.
TEST(FormulonCApiCircularReference, ACrossSheetCycleIsListedOnEachSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  std::size_t second = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Sheet2", 1U, &second), FM_STATUS_OK);
  Formula(wb.handle, 0, 0U, 0U, "=Sheet2!B3");
  Formula(wb.handle, second, 2U, 1U, "=Sheet1!A1");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Circular(wb.handle, 0), (Cells{{0U, 0U}}));
  EXPECT_EQ(Circular(wb.handle, second), (Cells{{2U, 1U}}));
}

TEST(FormulonCApiCircularReference, NoCycleListsNothingAndBreakingACycleClearsIt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Circular(wb.handle, 0), Cells{});
  Formula(wb.handle, 0, 0U, 0U, "=B1");
  Formula(wb.handle, 0, 0U, 1U, "=A1");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  ASSERT_EQ(Circular(wb.handle, 0).size(), 2U);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0U, 1U, 1.0), FM_STATUS_OK);
  // Automatic calculation: the read recalculates first.
  EXPECT_EQ(Circular(wb.handle, 0), Cells{});
}

// Manual calculation reports the last recalc until the next one.
TEST(FormulonCApiCircularReference, ManualCalculationReportsTheLastRecalc) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_MANUAL), FM_STATUS_OK);
  Formula(wb.handle, 0, 0U, 0U, "=B1");
  Formula(wb.handle, 0, 0U, 1U, "=A1");
  EXPECT_EQ(Circular(wb.handle, 0), Cells{});
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Circular(wb.handle, 0).size(), 2U);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0U, 1U, 1.0), FM_STATUS_OK);
  EXPECT_EQ(Circular(wb.handle, 0).size(), 2U);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Circular(wb.handle, 0), Cells{});
}

// Excel reports no circular reference while iterative calculation is on (e1).
TEST(FormulonCApiCircularReference, IterativeCalculationListsNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  Formula(wb.handle, 0, 0U, 0U, "=B1+1");
  Formula(wb.handle, 0, 0U, 1U, "=A1/2");
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  ASSERT_EQ(Circular(wb.handle, 0).size(), 2U);
  fm_iterative_settings_patch_t patch = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
  patch.enabled_op = FM_PATCH_OP_SET;
  patch.enabled = 1;
  ASSERT_EQ(fm_workbook_update_iterative(wb.handle, &patch), FM_STATUS_OK);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Circular(wb.handle, 0), Cells{});
}

TEST(FormulonCApiCircularReference, BoundaryArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  std::size_t count = 0;
  fm_cell_address_t cell{};
  EXPECT_EQ(fm_sheet_get_circular_reference_count(nullptr, 0, &count), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_circular_reference_count(wb.handle, 0, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_circular_reference_at(wb.handle, 0, 0, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_sheet_get_circular_reference_count(wb.handle, 5, &count), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_circular_reference_at(wb.handle, 5, 0, &cell), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_get_circular_reference_at(wb.handle, 0, 0, &cell), FM_STATUS_INDEX_OUT_OF_RANGE);
}
