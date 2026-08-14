//
// Stable C ABI dynamic-array spill payload regression tests.

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
};

}  // namespace

TEST(FormulonCApiSpill, NotSpilledLiteralReturnsNotEngaged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=42"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_spill_info_t info{};
  ASSERT_EQ(fm_workbook_spill_info(wb.handle, 0, 0, 0, &info), 0);
  EXPECT_EQ(info.engaged, 0);
  EXPECT_EQ(info.rows, 0U);
  EXPECT_EQ(info.cols, 0U);
}

TEST(FormulonCApiSpill, SequenceSpillReportsAnchorAndShape) {
  // A1 = =SEQUENCE(2,3) spills 2 rows × 3 cols starting at A1.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(2,3)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  // Anchor lookup
  fm_spill_info_t at_anchor{};
  ASSERT_EQ(fm_workbook_spill_info(wb.handle, 0, 0, 0, &at_anchor), 0);
  EXPECT_EQ(at_anchor.engaged, 1);
  EXPECT_EQ(at_anchor.anchor_row, 0U);
  EXPECT_EQ(at_anchor.anchor_col, 0U);
  EXPECT_EQ(at_anchor.rows, 2U);
  EXPECT_EQ(at_anchor.cols, 3U);

  // Phantom lookup (any cell within the region, e.g. B2 = row=1, col=1)
  fm_spill_info_t at_phantom{};
  ASSERT_EQ(fm_workbook_spill_info(wb.handle, 0, 1, 1, &at_phantom), 0);
  EXPECT_EQ(at_phantom.engaged, 1);
  EXPECT_EQ(at_phantom.anchor_row, 0U);
  EXPECT_EQ(at_phantom.anchor_col, 0U);
  EXPECT_EQ(at_phantom.rows, 2U);
  EXPECT_EQ(at_phantom.cols, 3U);

  // Outside region (D1, beyond cols=3 from anchor col=0 → col=3 is just past)
  fm_spill_info_t outside{};
  ASSERT_EQ(fm_workbook_spill_info(wb.handle, 0, 0, 3, &outside), 0);
  EXPECT_EQ(outside.engaged, 0);
}

TEST(FormulonCApiSpill, ExpandBlankPadSpillsZerosAndCountReadsTwelve) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=EXPAND({1,2;3,4},3,4,)"), 0);  // A1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_spill_info_t info{};
  ASSERT_EQ(fm_workbook_spill_info(wb.handle, 0, 0, 0, &info), 0);
  ASSERT_EQ(info.engaged, 1);
  EXPECT_EQ(info.rows, 3U);
  EXPECT_EQ(info.cols, 4U);

  // C1 and D3 are generated pad cells, and the committed spill stores their
  // projected numeric-zero values for later grid reads.
  fm_value_t c1{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 2, &c1), 0);
  ASSERT_EQ(c1.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(c1.u.number, 0.0);
  fm_value_t d3{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 2, 3, &d3), 0);
  ASSERT_EQ(d3.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(d3.u.number, 0.0);

  // A dependent spill reference reads the committed grid values, including
  // all eight zero pads, so COUNT reports all twelve cells.
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 5, "=COUNT(A1#)"), 0);  // F1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t count{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 5, &count), 0);
  ASSERT_EQ(count.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(count.u.number, 12.0);
}

TEST(FormulonCApiSpill, RawRangeSortSpillsBlankAsZeroAndCountReadsThree) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  // A1:A3 is {2, blank, 1}; the absent A2 is the interior raw-grid blank.
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 0, "=2"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 2, 0, "=1"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 3, "=SORT(A1:A3)"), 0);  // D1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t d3{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 2, 3, &d3), 0);
  ASSERT_EQ(d3.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(d3.u.number, 0.0);

  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 5, "=COUNT(D1#)"), 0);  // F1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t count{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 5, &count), 0);
  ASSERT_EQ(count.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(count.u.number, 3.0);
}

TEST(FormulonCApiSpill, CellEnumerationIncludesSpillPhantoms) {
  // D1 (row 0, col 3) = =SEQUENCE(3,1) spills to D1:D3. The phantoms D2 and
  // D3 must appear in the flat enumeration with their spilled values and no
  // formula text, and the cell count must bound the iterable index range.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_excel_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 0, 3, "=SEQUENCE(3,1)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  size_t count = 0;
  ASSERT_EQ(fm_workbook_cell_count(wb.handle, 0, &count), 0);
  ASSERT_GT(count, 0U);

  bool saw_d2 = false;
  bool saw_d3 = false;
  for (size_t i = 0; i < count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    const char* formula = nullptr;
    fm_value_t value{};
    ASSERT_EQ(fm_workbook_cell_at(wb.handle, 0, i, &row, &col, &formula, &value), 0);
    if (row == 1 && col == 3) {
      saw_d2 = true;
      EXPECT_EQ(formula, nullptr);
      ASSERT_EQ(value.kind, FM_VAL_NUMBER);
      EXPECT_DOUBLE_EQ(value.u.number, 2.0);
    }
    if (row == 2 && col == 3) {
      saw_d3 = true;
      EXPECT_EQ(formula, nullptr);
      ASSERT_EQ(value.kind, FM_VAL_NUMBER);
      EXPECT_DOUBLE_EQ(value.u.number, 3.0);
    }
  }
  EXPECT_TRUE(saw_d2);
  EXPECT_TRUE(saw_d3);
}

TEST(FormulonCApiSpill, CellEnumerationCacheInvalidatesAfterMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);

  // Starts a cached enumeration pass for the original one-cell sheet.
  uint32_t row = 0;
  uint32_t col = 0;
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(row, 0U);
  EXPECT_EQ(col, 0U);

  // A new stored cell must invalidate the cached coordinate list, even
  // though this is not a fresh handle or a fresh process.
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 5, 7, 2.0), 0);
  size_t count = 0;
  ASSERT_EQ(fm_workbook_cell_count(wb.handle, 0, &count), 0);
  // A row's slots start at its first populated column, so writing column 7
  // materialises one slot in row 5 alongside the original A1 slot.
  ASSERT_EQ(count, 2U);
  ASSERT_EQ(fm_workbook_cell_at(wb.handle, 0, count - 1U, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(row, 5U);
  EXPECT_EQ(col, 7U);
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 2.0);
}

TEST(FormulonCApiSpill, CellEnumerationCacheInvalidatesAfterSheetRemoval) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second"), 0);
  // Give both sheets the same enumeration revision while making their
  // coordinate lists observably different.
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 1, 1, 1, 2.0), 0);

  uint32_t row = 0;
  uint32_t col = 0;
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);

  // Sheet 1 becomes sheet 0. Without an explicit invalidation this can
  // incorrectly reuse the old sheet-0 cache because both revisions are 1.
  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, 0), 0);
  size_t count = 0;
  ASSERT_EQ(fm_workbook_cell_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  ASSERT_EQ(fm_workbook_cell_at(wb.handle, 0, count - 1U, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(row, 1U);
  EXPECT_EQ(col, 1U);
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 2.0);
}

TEST(FormulonCApiSpill, CellEnumerationCacheInvalidatesAfterSheetMove) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 1, 1, 1, 2.0), 0);

  uint32_t row = 0;
  uint32_t col = 0;
  fm_value_t value{};
  ASSERT_EQ(fm_workbook_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);

  // As with removal, the incoming sheet has the same revision as the
  // previously cached one, so index movement itself must discard the cache.
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 1, 0), 0);
  size_t count = 0;
  ASSERT_EQ(fm_workbook_cell_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  ASSERT_EQ(fm_workbook_cell_at(wb.handle, 0, count - 1U, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(row, 1U);
  EXPECT_EQ(col, 1U);
  ASSERT_EQ(value.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 2.0);
}

TEST(FormulonCApiSpill, OutOfRangeSheetReturnsInvalidArgument) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_spill_info_t info{};
  fm_status_t rc = fm_workbook_spill_info(wb.handle, 99, 0, 0, &info);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(FormulonCApiSpill, NullArgsReturnBindingNullPointer) {
  fm_spill_info_t info{};
  EXPECT_EQ(fm_workbook_spill_info(nullptr, 0, 0, 0, &info),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_spill_info(wb.handle, 0, 0, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}
