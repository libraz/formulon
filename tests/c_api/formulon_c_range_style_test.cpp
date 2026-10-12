// Attribute-only formatting of a selection (fm_sheet_range_style_bases /
// fm_sheet_remap_range_style) and its undo capture. The expectations mirror
// how Mac Excel 365 propagates a font-size change over a column, row, range
// and whole-sheet selection when other rows / columns carry styles.

#include <cstdint>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

constexpr std::uint32_t kLastRow = formulon::Sheet::kMaxRows - 1U;
constexpr std::uint32_t kLastCol = formulon::Sheet::kMaxCols - 1U;
constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);
constexpr std::int32_t kSourceCell = 0;
constexpr std::int32_t kSourceRow = 1;
constexpr std::int32_t kSourceColumn = 2;

// Synthetic patch: Normal 0 -> 5, row style 1 -> 6, column style 2 -> 7.
const std::vector<std::uint32_t> kFrom = {0U, 1U, 2U};
const std::vector<std::uint32_t> kTo = {5U, 6U, 7U};

void MakeWorkbook(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(10U);
}

std::pair<std::int32_t, std::uint32_t> Effective(const WorkbookGuard& wb, std::uint32_t row, std::uint32_t col) {
  fm_effective_style style{};
  EXPECT_EQ(fm_sheet_get_effective_style(wb.handle, 0, row, col, &style), 0);
  return {style.source, style.xf_index};
}

fm_status_t Remap(WorkbookGuard& wb, std::uint32_t r1, std::uint32_t c1, std::uint32_t r2, std::uint32_t c2) {
  return fm_sheet_remap_range_style(wb.handle, 0, r1, c1, r2, c2, kFrom.data(), kTo.data(), kFrom.size());
}

std::vector<fm_column_layout_t> Columns(const WorkbookGuard& wb) {
  size_t count = 0;
  EXPECT_EQ(fm_sheet_get_column_count(wb.handle, 0, &count), 0);
  std::vector<fm_column_layout_t> out(count);
  for (size_t i = 0; i < count; ++i) {
    EXPECT_EQ(fm_sheet_get_column(wb.handle, 0, i, &out[i]), 0);
  }
  return out;
}

const formulon::Cell* CellAt(const WorkbookGuard& wb, std::uint32_t row, std::uint32_t col) {
  return wb.handle->workbook().sheet(0).cell_at(row, col);
}

using Eff = std::pair<std::int32_t, std::uint32_t>;

}  // namespace

TEST(FormulonCApiRangeStyle, ColumnSelectionCrossesStyledRows) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_row_style(wb.handle, 0, 2U, 2U, 1U), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 9U, 1U, 5.0), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 9U, 2U, 7.0), 0);

  ASSERT_EQ(Remap(wb, 0U, 1U, kLastRow, 1U), 0);

  const auto cols = Columns(wb);
  ASSERT_EQ(cols.size(), 1U);
  EXPECT_EQ(cols[0].first, 1U);
  EXPECT_EQ(cols[0].last, 1U);
  EXPECT_EQ(cols[0].style_xf, 5U);
  EXPECT_EQ(Effective(wb, 2U, 1U), Eff(kSourceCell, 6U));  // B3 keeps the row's other attributes
  EXPECT_EQ(Effective(wb, 2U, 0U), Eff(kSourceRow, 1U));   // A3
  EXPECT_EQ(Effective(wb, 2U, 2U), Eff(kSourceRow, 1U));   // C3
  EXPECT_EQ(Effective(wb, 9U, 1U), Eff(kSourceCell, 5U));  // existing B10
  EXPECT_EQ(Effective(wb, 9U, 2U), Eff(kSourceCell, 0U));  // C10 untouched
  EXPECT_EQ(Effective(wb, 50000U, 1U), Eff(kSourceColumn, 5U));
  EXPECT_EQ(CellAt(wb, 50000U, 1U), nullptr);

  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 60000U, 1U, 1.0), 0);
  EXPECT_EQ(Effective(wb, 60000U, 1U), Eff(kSourceCell, 5U));  // a later entry takes the column style
}

TEST(FormulonCApiRangeStyle, RowSelectionCrossesStyledColumns) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_column_style(wb.handle, 0, 2U, 2U, 2U), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 3U, 2U, 1.0), 0);  // C4, enters with the column style

  ASSERT_EQ(Remap(wb, 2U, 0U, 4U, kLastCol), 0);

  for (std::uint32_t row = 2U; row <= 4U; ++row) {
    EXPECT_EQ(Effective(wb, row, 0U), Eff(kSourceRow, 5U)) << row;
    EXPECT_EQ(Effective(wb, row, 2U), Eff(kSourceCell, 7U)) << row;
  }
  EXPECT_EQ(Effective(wb, 5U, 2U), Eff(kSourceColumn, 2U));
  EXPECT_DOUBLE_EQ(CellAt(wb, 3U, 2U)->cached_value.as_number(), 1.0);
  size_t rows = 0;
  ASSERT_EQ(fm_sheet_get_row_override_count(wb.handle, 0, &rows), 0);
  EXPECT_EQ(rows, 3U);
}

TEST(FormulonCApiRangeStyle, RangeSelectionMaterialisesEveryCell) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_row_style(wb.handle, 0, 2U, 2U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_column_style(wb.handle, 0, 2U, 2U, 2U), 0);

  ASSERT_EQ(Remap(wb, 1U, 1U, 3U, 3U), 0);

  EXPECT_EQ(Effective(wb, 1U, 1U), Eff(kSourceCell, 5U));
  EXPECT_EQ(Effective(wb, 1U, 2U), Eff(kSourceCell, 7U));
  EXPECT_EQ(Effective(wb, 1U, 3U), Eff(kSourceCell, 5U));
  EXPECT_EQ(Effective(wb, 2U, 1U), Eff(kSourceCell, 6U));
  EXPECT_EQ(Effective(wb, 2U, 2U), Eff(kSourceCell, 6U));
  EXPECT_EQ(Effective(wb, 2U, 3U), Eff(kSourceCell, 6U));
  EXPECT_EQ(Effective(wb, 3U, 2U), Eff(kSourceCell, 7U));
  EXPECT_EQ(Effective(wb, 2U, 4U), Eff(kSourceRow, 1U));  // outside the range
  ASSERT_EQ(Columns(wb).size(), 1U);
  EXPECT_EQ(Columns(wb)[0].style_xf, 2U);
}

TEST(FormulonCApiRangeStyle, ColumnSelectionSplitsByColumnStyle) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_column_style(wb.handle, 0, 2U, 2U, 2U), 0);

  ASSERT_EQ(Remap(wb, 0U, 1U, kLastRow, 3U), 0);

  const auto cols = Columns(wb);
  ASSERT_EQ(cols.size(), 3U);
  EXPECT_EQ(cols[0].first, 1U);
  EXPECT_EQ(cols[0].style_xf, 5U);
  EXPECT_EQ(cols[1].first, 2U);
  EXPECT_EQ(cols[1].style_xf, 7U);
  EXPECT_EQ(cols[2].first, 3U);
  EXPECT_EQ(cols[2].style_xf, 5U);
}

TEST(FormulonCApiRangeStyle, WholeSheetNeedsNoIntersectionCells) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_row_style(wb.handle, 0, 2U, 2U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_column_style(wb.handle, 0, 2U, 2U, 2U), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 6U, 4U, 3.0), 0);

  ASSERT_EQ(Remap(wb, 0U, 0U, kLastRow, kLastCol), 0);

  const auto cols = Columns(wb);
  ASSERT_EQ(cols.size(), 3U);
  EXPECT_EQ(cols[0].first, 0U);
  EXPECT_EQ(cols[0].last, 1U);
  EXPECT_EQ(cols[0].style_xf, 5U);
  EXPECT_EQ(cols[1].style_xf, 7U);
  EXPECT_EQ(cols[2].first, 3U);
  EXPECT_EQ(cols[2].last, kLastCol);
  EXPECT_EQ(cols[2].style_xf, 5U);
  EXPECT_EQ(Effective(wb, 2U, 2U), Eff(kSourceRow, 6U));
  EXPECT_EQ(Effective(wb, 2U, 9000U), Eff(kSourceRow, 6U));
  EXPECT_EQ(Effective(wb, 6U, 4U), Eff(kSourceCell, 5U));
  EXPECT_EQ(Effective(wb, 7U, 2U), Eff(kSourceColumn, 7U));
  size_t rows = 0;
  ASSERT_EQ(fm_sheet_get_row_override_count(wb.handle, 0, &rows), 0);
  EXPECT_EQ(rows, 1U);
  size_t cells = 0;
  ASSERT_EQ(fm_workbook_cell_count(wb.handle, 0, &cells), 0);
  EXPECT_EQ(cells, 1U);
}

TEST(FormulonCApiRangeStyle, BasesListEveryXfTheSelectionResolvesTo) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_row_style(wb.handle, 0, 2U, 2U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_column_style(wb.handle, 0, 2U, 2U, 2U), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 9U, 1U, 4U), 0);

  size_t len = 99;
  ASSERT_EQ(fm_sheet_range_style_bases(wb.handle, 0, 0U, 1U, kLastRow, 1U, nullptr, 0U, &len), 0);
  ASSERT_EQ(len, 3U);
  std::vector<std::uint32_t> xfs(len);
  ASSERT_EQ(fm_sheet_range_style_bases(wb.handle, 0, 0U, 1U, kLastRow, 1U, xfs.data(), xfs.size(), &len), 0);
  EXPECT_EQ(xfs, (std::vector<std::uint32_t>{0U, 1U, 4U}));
}

std::vector<std::uint32_t> Bases(const WorkbookGuard& wb, std::uint32_t first_row, std::uint32_t first_col,
                                 std::uint32_t last_row, std::uint32_t last_col) {
  std::vector<std::uint32_t> xfs(16U);
  size_t len = 0;
  EXPECT_EQ(
      fm_sheet_range_style_bases(wb.handle, 0, first_row, first_col, last_row, last_col, xfs.data(), xfs.size(), &len),
      0);
  xfs.resize(len);
  return xfs;
}

TEST(FormulonCApiRangeStyle, BasesOmitStylesNoCellResolvesTo) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 2U, 2U, 1U), 0);
  EXPECT_EQ(Bases(wb, 2U, 2U, 2U, 2U), (std::vector<std::uint32_t>{1U}));
  EXPECT_EQ(Bases(wb, 2U, 2U, 2U, 3U), (std::vector<std::uint32_t>{0U, 1U}));

  // A row style shadows the column style wherever both cover a cell.
  ASSERT_EQ(fm_sheet_set_row_style(wb.handle, 0, 5U, 6U, 3U), 0);
  ASSERT_EQ(fm_sheet_set_column_style(wb.handle, 0, 4U, 4U, 2U), 0);
  EXPECT_EQ(Bases(wb, 5U, 4U, 6U, 4U), (std::vector<std::uint32_t>{3U}));
  EXPECT_EQ(Bases(wb, 5U, 4U, 7U, 4U), (std::vector<std::uint32_t>{2U, 3U}));

  // Owning cells covering a whole styled row leave the row style unused.
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 5U, 4U, 4U), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 6U, 4U, 4U), 0);
  EXPECT_EQ(Bases(wb, 5U, 4U, 6U, 4U), (std::vector<std::uint32_t>{4U}));
}

TEST(FormulonCApiRangeStyle, SnapshotRestoresPriorInheritance) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_row_style(wb.handle, 0, 2U, 2U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_column_style(wb.handle, 0, 2U, 2U, 2U), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 9U, 1U, "=1+2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  for (const auto& rect : std::vector<std::vector<std::uint32_t>>{
           {0U, 1U, kLastRow, 1U}, {2U, 0U, 4U, kLastCol}, {1U, 1U, 9U, 3U}, {0U, 0U, kLastRow, kLastCol}}) {
    BufferGuard capture;
    ASSERT_EQ(
        fm_sheet_snapshot_range_style(wb.handle, 0, rect[0], rect[1], rect[2], rect[3], &capture.data, &capture.len),
        0);
    ASSERT_EQ(Remap(wb, rect[0], rect[1], rect[2], rect[3]), 0);
    ASSERT_EQ(fm_sheet_restore_range_style(wb.handle, 0, capture.data, capture.len), 0);

    EXPECT_EQ(Effective(wb, 2U, 1U), Eff(kSourceRow, 1U));
    EXPECT_EQ(Effective(wb, 2U, 2U), Eff(kSourceRow, 1U));
    EXPECT_EQ(Effective(wb, 5U, 2U), Eff(kSourceColumn, 2U));
    EXPECT_EQ(Effective(wb, 5U, 1U).first, 3);
    EXPECT_EQ(Effective(wb, 3U, 0U).first, 3);
    const auto cols = Columns(wb);
    ASSERT_EQ(cols.size(), 1U);
    EXPECT_EQ(cols[0].first, 2U);
    EXPECT_EQ(cols[0].last, 2U);
    EXPECT_EQ(cols[0].style_xf, 2U);
    size_t rows = 0;
    ASSERT_EQ(fm_sheet_get_row_override_count(wb.handle, 0, &rows), 0);
    EXPECT_EQ(rows, 1U);
    EXPECT_EQ(CellAt(wb, 9U, 1U)->formula_text, "=1+2");
    EXPECT_DOUBLE_EQ(CellAt(wb, 9U, 1U)->cached_value.as_number(), 3.0);
  }
}

TEST(FormulonCApiRangeStyle, ResultSurvivesSaveLoad) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_row_style(wb.handle, 0, 2U, 2U, 1U), 0);
  ASSERT_EQ(Remap(wb, 0U, 1U, kLastRow, 1U), 0);

  for (const std::int32_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    BufferGuard saved;
    ASSERT_EQ(fm_workbook_save_as(wb.handle, format, &saved.data, &saved.len), 0);
    WorkbookGuard reloaded;
    ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
    EXPECT_EQ(Effective(reloaded, 2U, 1U), Eff(kSourceCell, 6U)) << format;
    EXPECT_EQ(Effective(reloaded, 2U, 2U), Eff(kSourceRow, 1U)) << format;
    EXPECT_EQ(Effective(reloaded, 700U, 1U), Eff(kSourceColumn, 5U)) << format;
  }
}

TEST(FormulonCApiRangeStyle, RejectsBadPairsAndOversizedWritesWithoutMutation) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_column_style(wb.handle, 0, 2U, 3U, 2U), 0);

  const std::uint32_t repeat_from[] = {0U, 0U};
  const std::uint32_t repeat_to[] = {5U, 6U};
  EXPECT_EQ(fm_sheet_remap_range_style(wb.handle, 0, 0U, 0U, 3U, 3U, repeat_from, repeat_to, 2U), kInvalidArgument);
  const std::uint32_t from[] = {0U};
  const std::uint32_t past_table[] = {10U};
  EXPECT_EQ(fm_sheet_remap_range_style(wb.handle, 0, 0U, 0U, 3U, 3U, from, past_table, 1U), kInvalidArgument);
  EXPECT_EQ(fm_sheet_remap_range_style(wb.handle, 0, 0U, 0U, kLastRow + 1U, 3U, from, repeat_to, 1U), kInvalidArgument);
  // Two styled columns crossed by almost every row: over a million writes.
  EXPECT_EQ(Remap(wb, 0U, 0U, kLastRow - 1U, kLastCol), kPreconditionFailed);
  EXPECT_EQ(Remap(wb, 0U, 0U, 2048U, 1023U), kPreconditionFailed);

  size_t cells = 0;
  ASSERT_EQ(fm_workbook_cell_count(wb.handle, 0, &cells), 0);
  EXPECT_EQ(cells, 0U);
  size_t rows = 0;
  ASSERT_EQ(fm_sheet_get_row_override_count(wb.handle, 0, &rows), 0);
  EXPECT_EQ(rows, 0U);
  EXPECT_EQ(Columns(wb).size(), 1U);
}

TEST(FormulonCApiRangeStyle, RestoreRejectsMalformedCaptures) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0U, 0U, 9U), 0);
  BufferGuard capture;
  ASSERT_EQ(fm_sheet_snapshot_range_style(wb.handle, 0, 0U, 0U, 1U, 1U, &capture.data, &capture.len), 0);

  EXPECT_EQ(fm_sheet_restore_range_style(wb.handle, 0, capture.data, capture.len - 1U), kInvalidArgument);
  const std::uint8_t junk[] = {'X', 'Y', 'Z', 'W', 1U};
  EXPECT_EQ(fm_sheet_restore_range_style(wb.handle, 0, junk, sizeof(junk)), kInvalidArgument);

  WorkbookGuard smaller;
  ASSERT_EQ(fm_workbook_create(&smaller.handle), 0);
  EXPECT_EQ(fm_sheet_restore_range_style(smaller.handle, 0, capture.data, capture.len), kInvalidArgument);
  EXPECT_EQ(smaller.handle->workbook().sheet(0).cell_at(0U, 0U), nullptr);
}

TEST(FormulonCApiRangeStyle, RowStyleSpanSetsAndClearsManyRows) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 10U, 30.0), 0);
  ASSERT_EQ(fm_sheet_set_row_style(wb.handle, 0, 0U, 99999U, 1U), 0);
  size_t rows = 0;
  ASSERT_EQ(fm_sheet_get_row_override_count(wb.handle, 0, &rows), 0);
  EXPECT_EQ(rows, 100000U);
  EXPECT_EQ(Effective(wb, 99999U, 5U), Eff(kSourceRow, 1U));

  ASSERT_EQ(fm_sheet_clear_row_style(wb.handle, 0, 0U, 99999U), 0);
  ASSERT_EQ(fm_sheet_get_row_override_count(wb.handle, 0, &rows), 0);
  EXPECT_EQ(rows, 1U);  // the height override stays
  EXPECT_EQ(fm_sheet_set_row_style(wb.handle, 0, 5U, 4U, 1U), kInvalidArgument);
  EXPECT_EQ(fm_sheet_clear_row_style(wb.handle, 0, 0U, kLastRow + 1U), kInvalidArgument);
}
