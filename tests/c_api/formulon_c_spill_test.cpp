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
