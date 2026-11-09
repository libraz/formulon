//
// Stable C ABI dynamic-array spill payload regression tests.

#include <limits>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_out_index.h"
#include "formulon_c_record_arg.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
};

}  // namespace

TEST(FormulonCApiSpill, NotSpilledLiteralReportsAbsent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=42"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
  int32_t info_present = -1;
  ASSERT_EQ(fm_sheet_get_spill(wb.handle, 0, 0, 0, &info_present, &info), 0);
  EXPECT_EQ(info_present, 0);
  EXPECT_EQ(info.rect.last_row, 0U);
  EXPECT_EQ(info.rect.last_col, 0U);
}

TEST(FormulonCApiSpill, SequenceSpillReportsAnchorAndShape) {
  // A1 = =SEQUENCE(2,3) spills 2 rows × 3 cols starting at A1.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=SEQUENCE(2,3)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  // Anchor lookup
  fm_spill_info_t at_anchor = FM_STRUCT_INIT(fm_spill_info_t);
  int32_t at_anchor_present = -1;
  ASSERT_EQ(fm_sheet_get_spill(wb.handle, 0, 0, 0, &at_anchor_present, &at_anchor), 0);
  EXPECT_EQ(at_anchor_present, 1);
  EXPECT_EQ(at_anchor.rect.first_row, 0U);
  EXPECT_EQ(at_anchor.rect.first_col, 0U);
  EXPECT_EQ(at_anchor.rect.last_row - at_anchor.rect.first_row + 1U, 2U);
  EXPECT_EQ(at_anchor.rect.last_col - at_anchor.rect.first_col + 1U, 3U);

  // Phantom lookup (any cell within the region, e.g. B2 = row=1, col=1)
  fm_spill_info_t at_phantom = FM_STRUCT_INIT(fm_spill_info_t);
  int32_t at_phantom_present = -1;
  ASSERT_EQ(fm_sheet_get_spill(wb.handle, 0, 1, 1, &at_phantom_present, &at_phantom), 0);
  EXPECT_EQ(at_phantom_present, 1);
  EXPECT_EQ(at_phantom.rect.first_row, 0U);
  EXPECT_EQ(at_phantom.rect.first_col, 0U);
  EXPECT_EQ(at_phantom.rect.last_row - at_phantom.rect.first_row + 1U, 2U);
  EXPECT_EQ(at_phantom.rect.last_col - at_phantom.rect.first_col + 1U, 3U);

  // Outside region (D1, beyond cols=3 from anchor col=0 → col=3 is just past)
  fm_spill_info_t outside = FM_STRUCT_INIT(fm_spill_info_t);
  int32_t outside_present = -1;
  ASSERT_EQ(fm_sheet_get_spill(wb.handle, 0, 0, 3, &outside_present, &outside), 0);
  EXPECT_EQ(outside_present, 0);
}

TEST(FormulonCApiSpill, ExpandBlankPadSpillsZerosAndCountReadsTwelve) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=EXPAND({1,2;3,4},3,4,)"), 0);  // A1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
  int32_t info_present = -1;
  ASSERT_EQ(fm_sheet_get_spill(wb.handle, 0, 0, 0, &info_present, &info), 0);
  ASSERT_EQ(info_present, 1);
  EXPECT_EQ(info.rect.last_row - info.rect.first_row + 1U, 3U);
  EXPECT_EQ(info.rect.last_col - info.rect.first_col + 1U, 4U);

  // C1 and D3 are generated pad cells, and the committed spill stores their
  // projected numeric-zero values for later grid reads.
  fm_value_t c1{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 2, &c1), 0);
  ASSERT_EQ(c1.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(c1.u.number, 0.0);
  fm_value_t d3{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 2, 3, &d3), 0);
  ASSERT_EQ(d3.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(d3.u.number, 0.0);

  // A dependent spill reference reads the committed grid values, including
  // all eight zero pads, so COUNT reports all twelve cells.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 5, "=COUNT(A1#)"), 0);  // F1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t count{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 5, &count), 0);
  ASSERT_EQ(count.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(count.u.number, 12.0);
}

TEST(FormulonCApiSpill, RawRangeSortSpillsBlankAsZeroAndCountReadsThree) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  // A1:A3 is {2, blank, 1}; the absent A2 is the interior raw-grid blank.
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=2"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 2, 0, "=1"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 3, "=SORT(A1:A3)"), 0);  // D1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t d3{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 2, 3, &d3), 0);
  ASSERT_EQ(d3.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(d3.u.number, 0.0);

  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 5, "=COUNT(D1#)"), 0);  // F1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t count{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 5, &count), 0);
  ASSERT_EQ(count.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(count.u.number, 3.0);
}

TEST(FormulonCApiSpill, CellEnumerationIncludesSpillPhantoms) {
  // D1 (row 0, col 3) = =SEQUENCE(3,1) spills to D1:D3. The phantoms D2 and
  // D3 must appear in the flat enumeration with their spilled values and no
  // formula text, and the cell count must bound the iterable index range.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 3, "=SEQUENCE(3,1)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cell_count(wb.handle, 0, &count), 0);
  ASSERT_GT(count, 0U);

  bool saw_d2 = false;
  bool saw_d3 = false;
  for (size_t i = 0; i < count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    const char* formula = nullptr;
    fm_value_t value{};
    ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, i, &row, &col, &formula, &value), 0);
    if (row == 1 && col == 3) {
      saw_d2 = true;
      EXPECT_EQ(formula, nullptr);
      ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
      EXPECT_DOUBLE_EQ(value.u.number, 2.0);
    }
    if (row == 2 && col == 3) {
      saw_d3 = true;
      EXPECT_EQ(formula, nullptr);
      ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
      EXPECT_DOUBLE_EQ(value.u.number, 3.0);
    }
  }
  EXPECT_TRUE(saw_d2);
  EXPECT_TRUE(saw_d3);
}

TEST(FormulonCApiSpill, CellEnumerationCacheInvalidatesAfterMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);

  // Starts a cached enumeration pass for the original one-cell sheet.
  uint32_t row = 0;
  uint32_t col = 0;
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(row, 0U);
  EXPECT_EQ(col, 0U);

  // A new stored cell must invalidate the cached coordinate list, even
  // though this is not a fresh handle or a fresh process.
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 5, 7, 2.0), 0);
  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cell_count(wb.handle, 0, &count), 0);
  // A row's slots start at its first populated column, so writing column 7
  // materialises one slot in row 5 alongside the original A1 slot.
  ASSERT_EQ(count, 2U);
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, count - 1U, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(row, 5U);
  EXPECT_EQ(col, 7U);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 2.0);
}

TEST(FormulonCApiSpill, CellEnumerationCacheInvalidatesAfterSpillMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-ja_JP"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);

  // Warm the one-cell coordinate list before the spill exists.
  uint32_t row = 0;
  uint32_t col = 0;
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);

  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 3, "=SEQUENCE(2,1)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cell_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 3U);
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, count - 1U, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(row, 1U);
  EXPECT_EQ(col, 3U);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 2.0);
}

TEST(FormulonCApiSpill, CellEnumerationCacheKeepsTwoAlternatingSheetsWarm) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second", append_position(wb.handle), ignored_index()), 0);

  // Both sheets reach the same revision, but their coordinate lists differ.
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2, 2, 2.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 1, 1, 1, 3.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 1, 3, 3, 4.0), 0);

  uint32_t row = 0;
  uint32_t col = 0;
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);
  ASSERT_EQ(row, 0U);
  ASSERT_EQ(col, 0U);
  const auto first_sheet_data = wb.handle->cell_enumeration_cache[0].addresses.data();
  ASSERT_EQ(wb.handle->cell_enumeration_cache[0].sheet_index, 0U);
  ASSERT_EQ(wb.handle->cell_enumeration_cache[0].addresses.size(), 2U);

  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 1, 0, &row, &col, nullptr, &value), 0);
  ASSERT_EQ(row, 1U);
  ASSERT_EQ(col, 1U);
  const auto second_sheet_data = wb.handle->cell_enumeration_cache[0].addresses.data();
  ASSERT_EQ(wb.handle->cell_enumeration_cache[0].sheet_index, 1U);
  ASSERT_EQ(wb.handle->cell_enumeration_cache[1].sheet_index, 0U);
  ASSERT_EQ(wb.handle->cell_enumeration_cache[1].addresses.data(), first_sheet_data);

  // Alternating hits promote the existing entries. Their buffers stay put,
  // proving that a two-sheet pass does not repeatedly rebuild either list.
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 1, &row, &col, nullptr, &value), 0);
  ASSERT_EQ(row, 2U);
  ASSERT_EQ(col, 2U);
  ASSERT_EQ(wb.handle->cell_enumeration_cache[0].addresses.data(), first_sheet_data);
  ASSERT_EQ(wb.handle->cell_enumeration_cache[1].addresses.data(), second_sheet_data);

  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 1, 1, &row, &col, nullptr, &value), 0);
  ASSERT_EQ(row, 3U);
  ASSERT_EQ(col, 3U);
  ASSERT_EQ(wb.handle->cell_enumeration_cache[0].addresses.data(), second_sheet_data);
  ASSERT_EQ(wb.handle->cell_enumeration_cache[1].addresses.data(), first_sheet_data);
}

TEST(FormulonCApiSpill, CellEnumerationCacheSurvivesSheetAddition) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);

  uint32_t row = 0;
  uint32_t col = 0;
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);
  ASSERT_EQ(wb.handle->cell_enumeration_cache[0].sheet_index, 0U);
  const auto* addresses = wb.handle->cell_enumeration_cache[0].addresses.data();
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second", append_position(wb.handle), ignored_index()), 0);

  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(wb.handle->cell_enumeration_cache[0].sheet_index, 0U);
  EXPECT_EQ(wb.handle->cell_enumeration_cache[0].addresses.data(), addresses);
  EXPECT_EQ(row, 0U);
  EXPECT_EQ(col, 0U);
}

TEST(FormulonCApiSpill, CellEnumerationCacheSurvivesFailedSheetMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second", append_position(wb.handle), ignored_index()), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);

  uint32_t row = 0;
  uint32_t col = 0;
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);
  const auto data = wb.handle->cell_enumeration_cache[0].addresses.data();
  const auto revision = wb.handle->cell_enumeration_cache[0].revision;

  EXPECT_NE(fm_workbook_add_sheet(wb.handle, "SHEET1", append_position(wb.handle), ignored_index()), 0);
  EXPECT_NE(fm_workbook_move_sheet(wb.handle, 99U, 0U), 0);
  EXPECT_NE(fm_workbook_remove_sheet(wb.handle, 99U), 0);
  EXPECT_EQ(wb.handle->cell_enumeration_cache[0].sheet_index, 0U);
  EXPECT_EQ(wb.handle->cell_enumeration_cache[0].revision, revision);
  EXPECT_EQ(wb.handle->cell_enumeration_cache[0].addresses.data(), data);
}

TEST(FormulonCApiSpill, CellEnumerationCacheInvalidatesAfterSheetRemoval) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second", append_position(wb.handle), ignored_index()), 0);
  // Give both sheets the same enumeration revision while making their
  // coordinate lists observably different.
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 1, 1, 1, 2.0), 0);

  uint32_t row = 0;
  uint32_t col = 0;
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);

  // Sheet 1 becomes sheet 0. Without an explicit invalidation this can
  // incorrectly reuse the old sheet-0 cache because both revisions are 1.
  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, 0), 0);
  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cell_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, count - 1U, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(row, 1U);
  EXPECT_EQ(col, 1U);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 2.0);
}

TEST(FormulonCApiSpill, CellEnumerationCacheInvalidatesAfterSheetMove) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second", append_position(wb.handle), ignored_index()), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 1, 1, 1, 2.0), 0);

  uint32_t row = 0;
  uint32_t col = 0;
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);

  // As with removal, the incoming sheet has the same revision as the
  // previously cached one, so index movement itself must discard the cache.
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 1, 0), 0);
  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_cell_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, count - 1U, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(row, 1U);
  EXPECT_EQ(col, 1U);
  ASSERT_EQ(value.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_DOUBLE_EQ(value.u.number, 2.0);
}

TEST(FormulonCApiSpill, CellEnumerationCacheInvalidatesBothSlotsAfterSheetMove) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Second", append_position(wb.handle), ignored_index()), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 1, 1, 1, 2.0), 0);

  // Warm sheet 1 first so it sits in the second slot after sheet 0 is read.
  uint32_t row = 0;
  uint32_t col = 0;
  fm_value_t value{};
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 1, 0, &row, &col, nullptr, &value), 0);
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);

  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 1, 0), 0);
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 1, 0, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(row, 0U);
  EXPECT_EQ(col, 0U);
  ASSERT_EQ(fm_sheet_get_cell_at(wb.handle, 0, 0, &row, &col, nullptr, &value), 0);
  EXPECT_EQ(row, 1U);
  EXPECT_EQ(col, 1U);
}

TEST(FormulonCApiSpill, OutOfRangeSheetReturnsSheetIndexOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
  int32_t present = -1;
  fm_status_t rc = fm_sheet_get_spill(wb.handle, 99, 0, 0, &present, &info);
  EXPECT_EQ(rc, static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
}

TEST(FormulonCApiSpill, NullArgsReturnBindingNullPointer) {
  fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
  int32_t present = -1;
  EXPECT_EQ(fm_sheet_get_spill(nullptr, 0, 0, 0, &present, &info),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_get_spill(wb.handle, 0, 0, 0, &present, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_get_spill(wb.handle, 0, 0, 0, nullptr, &info),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

namespace {

struct SpillRead {
  int32_t present = -1;
  fm_spill_info_t info = FM_STRUCT_INIT(fm_spill_info_t);
};

SpillRead SpillAt(fm_workbook_t* wb, uint32_t row, uint32_t col) {
  SpillRead read;
  EXPECT_EQ(fm_sheet_get_spill(wb, 0, row, col, &read.present, &read.info), 0);
  return read;
}

/// Expects `info` to report a blocked anchor at `(row, col)` that attempted `rows` x `cols`.
void ExpectBlocked(const SpillRead& read, uint32_t row, uint32_t col, uint32_t rows, uint32_t cols) {
  const fm_spill_info_t& info = read.info;
  EXPECT_EQ(read.present, 1);
  EXPECT_EQ(info.blocked, 1);
  EXPECT_EQ(info.rect.first_row, row);
  EXPECT_EQ(info.rect.first_col, col);
  EXPECT_EQ(info.rect.last_row - info.rect.first_row + 1U, rows);
  EXPECT_EQ(info.rect.last_col - info.rect.first_col + 1U, cols);
}

void ExpectNothing(const SpillRead& read) {
  const fm_spill_info_t& info = read.info;
  EXPECT_EQ(read.present, 0);
  EXPECT_EQ(info.blocked, 0);
  EXPECT_EQ(info.rect.first_row, 0U);
  EXPECT_EQ(info.rect.first_col, 0U);
  EXPECT_EQ(info.rect.last_row, 0U);
  EXPECT_EQ(info.rect.last_col, 0U);
}

}  // namespace

TEST(FormulonCApiSpill, OccupiedCellBlocksAnchorUntilCleared) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 3, "=SEQUENCE(5)"), 0);  // D2
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 3, 3, "x"), 0);                // D4
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  ExpectBlocked(SpillAt(wb.handle, 1, 3), 1, 3, 5, 1);
  // Only the anchor reports it: an empty cell inside the attempted range and
  // the blocker itself read as nothing.
  ExpectNothing(SpillAt(wb.handle, 2, 3));
  ExpectNothing(SpillAt(wb.handle, 3, 3));
  ExpectNothing(SpillAt(wb.handle, 5, 3));

  ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0, 3, 3), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  const SpillRead read = SpillAt(wb.handle, 1, 3);
  const fm_spill_info_t& info = read.info;
  EXPECT_EQ(read.present, 1);
  EXPECT_EQ(info.blocked, 0);
  EXPECT_EQ(info.rect.first_row, 1U);
  EXPECT_EQ(info.rect.first_col, 3U);
  EXPECT_EQ(info.rect.last_row - info.rect.first_row + 1U, 5U);
  EXPECT_EQ(info.rect.last_col - info.rect.first_col + 1U, 1U);
}

TEST(FormulonCApiSpill, MergedCellBlocksAnchor) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_cell_rect_t merged{};
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0, record_arg(fm_cell_rect_t{3, 3, 3, 4}), &merged), 0);  // D4:E4
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 3, "=SEQUENCE(5)"), 0);                           // D2
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  ExpectBlocked(SpillAt(wb.handle, 1, 3), 1, 3, 5, 1);
  ExpectNothing(SpillAt(wb.handle, 2, 3));
}

TEST(FormulonCApiSpill, OtherSpillBlocksAnchor) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 3, 1, "=SEQUENCE(1,5)"), 0);  // B4 spills B4:F4
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 1, 3, "=SEQUENCE(5)"), 0);  // D2 meets D4
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  ExpectBlocked(SpillAt(wb.handle, 1, 3), 1, 3, 5, 1);
  // D4 belongs to the committed spill, not to the blocked one.
  const SpillRead d4_read = SpillAt(wb.handle, 3, 3);
  const fm_spill_info_t& d4 = d4_read.info;
  EXPECT_EQ(d4_read.present, 1);
  EXPECT_EQ(d4.blocked, 0);
  EXPECT_EQ(d4.rect.first_row, 3U);
  EXPECT_EQ(d4.rect.first_col, 1U);
  ExpectNothing(SpillAt(wb.handle, 2, 3));
}
