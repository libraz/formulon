//
// Stable C ABI tests for sheet geometry in points, range enumeration with a
// resumable cursor, formula text in A1 / R1C1, merges by rectangle, and the
// pagination detail accessors.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <set>
#include <string>
#include <tuple>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

static_assert(sizeof(fm_rect_pt) == 32U);
static_assert(sizeof(fm_sheet_format_defaults) == 32U);
static_assert(sizeof(fm_page_layout) == 56U);
static_assert(offsetof(fm_page_layout, origin_x_pt) == 24U);
static_assert(sizeof(fm_print_titles) == 24U);
static_assert(sizeof(fm_paper_info) == 24U);
static_assert(sizeof(fm_margins_pt) == 48U);
static_assert(offsetof(fm_width_model, normal_font_name) == 32U);

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr std::uint64_t kCursorEnd = std::numeric_limits<std::uint64_t>::max();
constexpr std::uint64_t kCursorStride = 16384U;

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct RangeGuard {
  fm_cell_range_t* handle = nullptr;
  ~RangeGuard() { fm_cell_range_destroy(handle); }
  RangeGuard() = default;
  RangeGuard(const RangeGuard&) = delete;
  RangeGuard& operator=(const RangeGuard&) = delete;
};

struct PaginationGuard {
  fm_pagination_t* handle = nullptr;
  ~PaginationGuard() { fm_pagination_destroy(handle); }
  PaginationGuard() = default;
  PaginationGuard(const PaginationGuard&) = delete;
  PaginationGuard& operator=(const PaginationGuard&) = delete;
};

/// One enumerated cell, copied out of its page handle.
struct SeenCell {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::string formula;  // empty when the cell holds no formula
  fm_value_kind_t kind = FM_VAL_BLANK;
  double number = 0.0;
  std::string text;

  friend bool operator==(const SeenCell& a, const SeenCell& b) {
    return std::tie(a.row, a.col, a.formula, a.kind, a.number, a.text) ==
           std::tie(b.row, b.col, b.formula, b.kind, b.number, b.text);
  }
};

/// Fetches one page and appends its cells; returns the next cursor.
std::uint64_t FetchPage(fm_workbook_t* wb, std::uint32_t first_row, std::uint32_t first_col, std::uint32_t last_row,
                        std::uint32_t last_col, std::uint64_t cursor, std::uint32_t limit, std::vector<SeenCell>* out) {
  RangeGuard page;
  EXPECT_EQ(fm_sheet_cells_in_range(wb, 0, first_row, first_col, last_row, last_col, cursor, limit, &page.handle), 0);
  if (page.handle == nullptr) {
    return kCursorEnd;
  }
  size_t count = 0;
  EXPECT_EQ(fm_cell_range_count(page.handle, &count), 0);
  for (size_t i = 0; i < count; ++i) {
    SeenCell cell;
    const char* formula = nullptr;
    fm_value_t value{};
    EXPECT_EQ(fm_cell_range_at(page.handle, i, &cell.row, &cell.col, &formula, &value), 0);
    cell.formula = formula == nullptr ? "" : formula;
    cell.kind = value.kind;
    if (value.kind == FM_VAL_NUMBER) {
      cell.number = value.u.number;
    } else if (value.kind == FM_VAL_TEXT) {
      cell.text = value.u.text;
    }
    out->push_back(cell);
  }
  std::uint64_t next = 0;
  EXPECT_EQ(fm_cell_range_next_cursor(page.handle, &next), 0);
  return next;
}

/// A sparse sheet: literals, text, a formula, and a vertical spill.
void SeedSparseSheet(fm_workbook_t* wb) {
  ASSERT_EQ(fm_workbook_set_number(wb, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_workbook_set_text(wb, 0, 0, 3, "head"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb, 0, 2, 1, 2.0), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb, 0, 2, 2, "=A1+B3"), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb, 0, 4, 4, "=SEQUENCE(3)"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb, 0, 900, 7, 9.0), 0);
  ASSERT_EQ(fm_workbook_set_text(wb, 0, 50000, 100, "far"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
}

}  // namespace

/* -------------------------------------------------------------------------- */
/* Geometry                                                                   */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiCellGeometry, WidthModelReportsCalibratedCalibri) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_width_model display{};
  ASSERT_EQ(fm_sheet_width_model(wb.handle, 0, FM_GEOMETRY_DISPLAY, &display), 0);
  EXPECT_DOUBLE_EQ(display.points_per_char, 5.25);
  EXPECT_DOUBLE_EQ(display.padding_pt, 3.75);
  EXPECT_EQ(display.calibrated, 1);
  EXPECT_STREQ(display.normal_font_name, "Calibri");
  EXPECT_DOUBLE_EQ(display.normal_font_size, 11.0);
  EXPECT_STREQ(display.platform, "win");

  fm_width_model print{};
  ASSERT_EQ(fm_sheet_width_model(wb.handle, 0, FM_GEOMETRY_PRINT, &print), 0);
  EXPECT_DOUBLE_EQ(print.points_per_char, 39.0 / 7.0);
  EXPECT_DOUBLE_EQ(print.padding_pt, 27.0 / 7.0);

  EXPECT_EQ(fm_sheet_width_model(wb.handle, 0, 2, &print), kInvalidArgument);
  EXPECT_EQ(fm_sheet_width_model(wb.handle, 0, -1, &print), kInvalidArgument);
}

TEST(FormulonCApiCellGeometry, CharsAndPointsConvertBothWays) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  double pt = 0.0;
  ASSERT_EQ(fm_sheet_column_chars_to_pt(wb.handle, 0, FM_GEOMETRY_DISPLAY, 10.0, &pt), 0);
  EXPECT_DOUBLE_EQ(pt, 10.0 * 5.25 + 3.75);
  double chars = 0.0;
  ASSERT_EQ(fm_sheet_column_pt_to_chars(wb.handle, 0, FM_GEOMETRY_DISPLAY, pt, &chars), 0);
  EXPECT_DOUBLE_EQ(chars, 10.0);
  ASSERT_EQ(fm_sheet_column_chars_to_pt(wb.handle, 0, FM_GEOMETRY_PRINT, 0.0, &pt), 0);
  EXPECT_EQ(pt, 0.0);

  EXPECT_EQ(fm_sheet_column_chars_to_pt(wb.handle, 0, FM_GEOMETRY_DISPLAY, -1.0, &pt), kInvalidArgument);
  EXPECT_EQ(
      fm_sheet_column_pt_to_chars(wb.handle, 0, FM_GEOMETRY_DISPLAY, std::numeric_limits<double>::infinity(), &chars),
      kInvalidArgument);
  EXPECT_NE(fm_sheet_column_chars_to_pt(wb.handle, 0, FM_GEOMETRY_DISPLAY, 1.0, nullptr), 0);
}

TEST(FormulonCApiCellGeometry, TrackSizesFollowOverridesAndHiddenState) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 1U, 20.0), 0);
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 2U, 2U, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 1U, 30.0), 0);
  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 2U, 1), 0);

  double width = 0.0;
  ASSERT_EQ(fm_sheet_column_width_pt(wb.handle, 0, 1U, FM_GEOMETRY_DISPLAY, &width), 0);
  EXPECT_DOUBLE_EQ(width, 20.0 * 5.25 + 3.75);
  ASSERT_EQ(fm_sheet_column_width_pt(wb.handle, 0, 2U, FM_GEOMETRY_DISPLAY, &width), 0);
  EXPECT_EQ(width, 0.0);
  double height = 0.0;
  ASSERT_EQ(fm_sheet_row_height_pt(wb.handle, 0, 1U, &height), 0);
  EXPECT_EQ(height, 30.0);
  ASSERT_EQ(fm_sheet_row_height_pt(wb.handle, 0, 2U, &height), 0);
  EXPECT_EQ(height, 0.0);

  EXPECT_EQ(fm_sheet_column_width_pt(wb.handle, 0, 16384U, FM_GEOMETRY_DISPLAY, &width), kInvalidArgument);
  EXPECT_EQ(fm_sheet_row_height_pt(wb.handle, 0, 1048576U, &height), kInvalidArgument);
}

TEST(FormulonCApiCellGeometry, CellRectSpansRangeAndSkipsHiddenTracks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 1U, 20.0), 0);
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 2U, 2U, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 1U, 30.0), 0);

  double col_a = 0.0;
  double col_b = 0.0;
  double row_1 = 0.0;
  double row_3 = 0.0;
  ASSERT_EQ(fm_sheet_column_width_pt(wb.handle, 0, 0U, FM_GEOMETRY_DISPLAY, &col_a), 0);
  ASSERT_EQ(fm_sheet_column_width_pt(wb.handle, 0, 1U, FM_GEOMETRY_DISPLAY, &col_b), 0);
  ASSERT_EQ(fm_sheet_row_height_pt(wb.handle, 0, 0U, &row_1), 0);
  ASSERT_EQ(fm_sheet_row_height_pt(wb.handle, 0, 2U, &row_3), 0);

  // B2:C3 -- column C is hidden and contributes nothing.
  fm_rect_pt rect{};
  ASSERT_EQ(fm_sheet_cell_rect_pt(wb.handle, 0, 1U, 1U, 2U, 2U, FM_GEOMETRY_DISPLAY, &rect), 0);
  EXPECT_DOUBLE_EQ(rect.x, col_a);
  EXPECT_DOUBLE_EQ(rect.y, row_1);
  EXPECT_DOUBLE_EQ(rect.width, col_b);
  EXPECT_DOUBLE_EQ(rect.height, 30.0 + row_3);

  // A single cell agrees with the per-track getters exactly.
  ASSERT_EQ(fm_sheet_cell_rect_pt(wb.handle, 0, 1U, 1U, 1U, 1U, FM_GEOMETRY_PRINT, &rect), 0);
  double col_b_print = 0.0;
  ASSERT_EQ(fm_sheet_column_width_pt(wb.handle, 0, 1U, FM_GEOMETRY_PRINT, &col_b_print), 0);
  EXPECT_EQ(rect.width, col_b_print);
  EXPECT_EQ(rect.height, 30.0);

  EXPECT_EQ(fm_sheet_cell_rect_pt(wb.handle, 0, 2U, 0U, 1U, 0U, FM_GEOMETRY_DISPLAY, &rect), kInvalidArgument);
  EXPECT_EQ(fm_sheet_cell_rect_pt(wb.handle, 0, 0U, 0U, 0U, 0U, 7, &rect), kInvalidArgument);
  EXPECT_NE(fm_sheet_cell_rect_pt(wb.handle, 0, 0U, 0U, 0U, 0U, FM_GEOMETRY_DISPLAY, nullptr), 0);
}

TEST(FormulonCApiCellGeometry, DefaultsDriveUnoverriddenTracks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_sheet_format_defaults defaults{};
  defaults.default_col_width = 12.0;
  defaults.default_row_height = 21.0;
  defaults.base_col_width = 8.0;
  defaults.has_default_col_width = 1;
  defaults.has_default_row_height = 1;
  ASSERT_EQ(fm_sheet_set_format_defaults(wb.handle, 0, &defaults), 0);
  double width = 0.0;
  ASSERT_EQ(fm_sheet_column_width_pt(wb.handle, 0, 5U, FM_GEOMETRY_DISPLAY, &width), 0);
  EXPECT_DOUBLE_EQ(width, 12.0 * 5.25 + 3.75);
  double height = 0.0;
  ASSERT_EQ(fm_sheet_row_height_pt(wb.handle, 0, 5U, &height), 0);
  EXPECT_EQ(height, 21.0);
}

/* -------------------------------------------------------------------------- */
/* Range enumeration and formula text                                         */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiCellRange, EnumeratesPopulatedCellsInRowMajorOrder) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedSparseSheet(wb.handle);

  std::vector<SeenCell> cells;
  EXPECT_EQ(FetchPage(wb.handle, 0U, 0U, 1048575U, 16383U, 0U, 0U, &cells), kCursorEnd);
  std::vector<std::pair<std::uint32_t, std::uint32_t>> coords;
  for (const SeenCell& cell : cells) {
    coords.emplace_back(cell.row, cell.col);
  }
  const std::vector<std::pair<std::uint32_t, std::uint32_t>> expected = {
      {0U, 0U}, {0U, 3U}, {2U, 1U}, {2U, 2U}, {4U, 4U}, {5U, 4U}, {6U, 4U}, {900U, 7U}, {50000U, 100U}};
  EXPECT_EQ(coords, expected);

  EXPECT_EQ(cells[1].kind, FM_VAL_TEXT);
  EXPECT_EQ(cells[1].text, "head");
  EXPECT_EQ(cells[3].formula, "=A1+B3");
  EXPECT_EQ(cells[3].number, 3.0);
  // The spill anchor carries the formula; its phantoms carry only values.
  EXPECT_EQ(cells[4].formula, "=SEQUENCE(3)");
  EXPECT_EQ(cells[5].formula, "");
  EXPECT_EQ(cells[5].number, 2.0);
  EXPECT_EQ(cells[6].number, 3.0);
}

TEST(FormulonCApiCellRange, CursorResumesWithoutGaps) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedSparseSheet(wb.handle);

  std::vector<SeenCell> whole;
  ASSERT_EQ(FetchPage(wb.handle, 0U, 0U, 1048575U, 16383U, 0U, 0U, &whole), kCursorEnd);

  for (const std::uint32_t limit : {1U, 2U, 4U}) {
    std::vector<SeenCell> paged;
    std::uint64_t cursor = 0;
    std::size_t pages = 0;
    while (cursor != kCursorEnd) {
      const std::size_t before = paged.size();
      cursor = FetchPage(wb.handle, 0U, 0U, 1048575U, 16383U, cursor, limit, &paged);
      ASSERT_LE(paged.size() - before, limit);
      ASSERT_LT(++pages, 100U);
      if (cursor != kCursorEnd) {
        // The cursor names the next populated cell itself.
        EXPECT_EQ(paged.size() - before, limit);
      }
    }
    EXPECT_EQ(paged, whole) << "limit " << limit;
    std::set<std::pair<std::uint32_t, std::uint32_t>> unique;
    for (const SeenCell& cell : paged) {
      unique.emplace(cell.row, cell.col);
    }
    EXPECT_EQ(unique.size(), paged.size());
  }
}

TEST(FormulonCApiCellRange, ClipsToRectangleAndCursorMayStartMidRow) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  SeedSparseSheet(wb.handle);

  // B1:E6 holds D1, B3, C3, E5 and the first phantom E6.
  std::vector<SeenCell> cells;
  EXPECT_EQ(FetchPage(wb.handle, 0U, 1U, 5U, 4U, 0U, 0U, &cells), kCursorEnd);
  ASSERT_EQ(cells.size(), 5U);
  EXPECT_EQ(cells.front().col, 3U);
  EXPECT_EQ(cells.back().row, 5U);

  // A cursor at C3 skips everything before it.
  cells.clear();
  EXPECT_EQ(FetchPage(wb.handle, 0U, 1U, 5U, 4U, 2U * kCursorStride + 2U, 0U, &cells), kCursorEnd);
  ASSERT_EQ(cells.size(), 3U);
  EXPECT_EQ(cells.front().row, 2U);
  EXPECT_EQ(cells.front().col, 2U);

  // An empty rectangle ends immediately.
  cells.clear();
  EXPECT_EQ(FetchPage(wb.handle, 0U, 10U, 20U, 20U, 10U, 0U, &cells), kCursorEnd);
  EXPECT_TRUE(cells.empty());
}

TEST(FormulonCApiCellRange, RejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  RangeGuard page;
  EXPECT_EQ(fm_sheet_cells_in_range(wb.handle, 0, 5U, 0U, 1U, 0U, 0U, 0U, &page.handle), kInvalidArgument);
  EXPECT_EQ(page.handle, nullptr);
  EXPECT_EQ(fm_sheet_cells_in_range(wb.handle, 0, 0U, 0U, 1U, 1U, 1048576U * kCursorStride, 0U, &page.handle),
            kInvalidArgument);
  EXPECT_EQ(fm_sheet_cells_in_range(wb.handle, 0, 0U, 0U, 1U, 1U, kCursorEnd, 0U, &page.handle), kInvalidArgument);
  EXPECT_EQ(fm_sheet_cells_in_range(wb.handle, 3, 0U, 0U, 1U, 1U, 0U, 0U, &page.handle), kInvalidArgument);
  EXPECT_NE(fm_sheet_cells_in_range(wb.handle, 0, 0U, 0U, 1U, 1U, 0U, 0U, nullptr), 0);

  ASSERT_EQ(fm_sheet_cells_in_range(wb.handle, 0, 0U, 0U, 1U, 1U, 0U, 0U, &page.handle), 0);
  size_t count = 99;
  ASSERT_EQ(fm_cell_range_count(page.handle, &count), 0);
  EXPECT_EQ(count, 0U);
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  fm_value_t value{};
  EXPECT_EQ(fm_cell_range_at(page.handle, 0, &row, &col, nullptr, &value), kInvalidArgument);
  EXPECT_NE(fm_cell_range_count(nullptr, &count), 0);
  std::uint64_t next = 0;
  EXPECT_NE(fm_cell_range_next_cursor(nullptr, &next), 0);
  fm_cell_range_destroy(nullptr);
}

TEST(FormulonCApiCellRange, FormulaGettersReadA1AndR1C1) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_formula(wb.handle, 0, 4, 2, "=SUM($A$1:B4)+A5"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);

  const char* text = nullptr;
  ASSERT_EQ(fm_workbook_get_formula(wb.handle, 0, 4, 2, &text), 0);
  EXPECT_STREQ(text, "=SUM($A$1:B4)+A5");
  ASSERT_EQ(fm_workbook_get_formula_r1c1(wb.handle, 0, 4, 2, &text), 0);
  EXPECT_STREQ(text, "SUM(R1C1:R[-1]C[-1])+RC[-2]");

  ASSERT_EQ(fm_workbook_get_formula(wb.handle, 0, 0, 0, &text), 0);
  EXPECT_STREQ(text, "");
  ASSERT_EQ(fm_workbook_get_formula_r1c1(wb.handle, 0, 9, 9, &text), 0);
  EXPECT_STREQ(text, "");

  EXPECT_EQ(fm_workbook_get_formula(wb.handle, 0, 1048576U, 0, &text), kInvalidArgument);
  EXPECT_EQ(fm_workbook_get_formula_r1c1(wb.handle, 4, 0, 0, &text), kInvalidArgument);
  EXPECT_NE(fm_workbook_get_formula(wb.handle, 0, 0, 0, nullptr), 0);
}

TEST(FormulonCApiCellRange, MergesInRangeReportsIntersectingMerges) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0, fm_merge_range{0U, 0U, 1U, 1U}), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0, fm_merge_range{5U, 5U, 6U, 8U}), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0, fm_merge_range{20U, 0U, 20U, 3U}), 0);

  std::uint32_t total = 0;
  // Corners given transposed: rows 1..6, columns 1..6.
  ASSERT_EQ(fm_sheet_merges_in_range(wb.handle, 0, fm_merge_range{6U, 6U, 1U, 1U}, nullptr, 0U, &total), 0);
  EXPECT_EQ(total, 2U);
  fm_merge_range out[1] = {};
  ASSERT_EQ(fm_sheet_merges_in_range(wb.handle, 0, fm_merge_range{1U, 1U, 6U, 6U}, out, 1U, &total), 0);
  EXPECT_EQ(total, 2U);
  EXPECT_EQ(out[0].first_row, 0U);
  EXPECT_EQ(out[0].last_col, 1U);

  EXPECT_NE(fm_sheet_merges_in_range(wb.handle, 0, fm_merge_range{0U, 0U, 1U, 1U}, nullptr, 1U, &total), 0);
  EXPECT_EQ(fm_sheet_merges_in_range(wb.handle, 9, fm_merge_range{0U, 0U, 1U, 1U}, nullptr, 0U, &total),
            kInvalidArgument);
}

/* -------------------------------------------------------------------------- */
/* Pagination detail                                                          */
/* -------------------------------------------------------------------------- */

namespace {

/// Fills `A1:T200` so the sheet spans several pages on both axes.
void FillGrid(fm_workbook_t* wb) {
  for (std::uint32_t row = 0; row < 200U; ++row) {
    for (std::uint32_t col = 0; col < 20U; ++col) {
      ASSERT_EQ(fm_workbook_set_number(wb, 0, row, col, 1.0), 0);
    }
  }
}

}  // namespace

TEST(FormulonCApiPaginationDetail, ReportsPaperMarginsScaleAndTitles) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup paperSize=\"9\" orientation=\"landscape\"/>"), 0);
  ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, "1:1", ""), 0);

  PaginationGuard p;
  ASSERT_EQ(fm_workbook_paginate(wb.handle, 0, &p.handle), 0);
  fm_paper_info paper{};
  ASSERT_EQ(fm_pagination_paper(p.handle, &paper), 0);
  EXPECT_EQ(paper.landscape, 1);
  EXPECT_EQ(paper.known, 1);
  EXPECT_GT(paper.width_pt, paper.height_pt);

  fm_margins_pt margins{};
  ASSERT_EQ(fm_pagination_margins(p.handle, &margins), 0);
  EXPECT_DOUBLE_EQ(margins.left, 0.7 * 72.0);
  EXPECT_DOUBLE_EQ(margins.top, 0.75 * 72.0);

  fm_rect_pt printable{};
  ASSERT_EQ(fm_pagination_printable(p.handle, &printable), 0);
  EXPECT_DOUBLE_EQ(printable.x, margins.left);
  EXPECT_GT(printable.width, 0.0);

  double scale = 0.0;
  ASSERT_EQ(fm_pagination_scale(p.handle, &scale), 0);
  EXPECT_EQ(scale, 1.0);

  fm_print_titles titles{};
  ASSERT_EQ(fm_pagination_print_titles(p.handle, &titles), 0);
  EXPECT_EQ(titles.has_rows, 1);
  EXPECT_EQ(titles.first_row, 0U);
  EXPECT_EQ(titles.last_row, 0U);
  EXPECT_EQ(titles.has_cols, 0);

  EXPECT_NE(fm_pagination_paper(nullptr, &paper), 0);
  EXPECT_NE(fm_pagination_margins(p.handle, nullptr), 0);
  EXPECT_NE(fm_pagination_scale(p.handle, nullptr), 0);
  EXPECT_NE(fm_pagination_print_titles(nullptr, &titles), 0);
  EXPECT_NE(fm_pagination_printable(nullptr, &printable), 0);
}

TEST(FormulonCApiPaginationDetail, UnknownPaperFallsBackAndSaysSo) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup paperSize=\"999\"/>"), 0);
  PaginationGuard p;
  ASSERT_EQ(fm_workbook_paginate(wb.handle, 0, &p.handle), 0);
  fm_paper_info paper{};
  ASSERT_EQ(fm_pagination_paper(p.handle, &paper), 0);
  EXPECT_EQ(paper.known, 0);
}

TEST(FormulonCApiPaginationDetail, PagesFollowPageOrderAndCoverEveryPage) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup pageOrder=\"overThenDown\"/>"), 0);

  PaginationGuard p;
  ASSERT_EQ(fm_workbook_paginate(wb.handle, 0, &p.handle), 0);
  int32_t order = -1;
  ASSERT_EQ(fm_pagination_page_order(p.handle, &order), 0);
  EXPECT_EQ(order, 1);

  const std::uint32_t page_count = fm_pagination_page_count(p.handle);
  ASSERT_GT(fm_pagination_vertical_break_count(p.handle), 0U);
  ASSERT_GT(page_count, 2U);
  std::vector<fm_page_layout> pages(page_count);
  for (std::uint32_t i = 0; i < page_count; ++i) {
    ASSERT_EQ(fm_pagination_page_at(p.handle, i, &pages[i]), 0);
    EXPECT_GT(pages[i].width_pt, 0.0);
    EXPECT_GT(pages[i].height_pt, 0.0);
  }
  // Over then down: the second page sits right of the first, same row band.
  EXPECT_EQ(pages[1].first_row, pages[0].first_row);
  EXPECT_GT(pages[1].first_col, pages[0].last_col);
  fm_page_layout extra{};
  EXPECT_EQ(fm_pagination_page_at(p.handle, page_count, &extra), kInvalidArgument);
  EXPECT_NE(fm_pagination_page_order(nullptr, &order), 0);
}

TEST(FormulonCApiPaginationDetail, ManualBreaksAreFlagged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillGrid(wb.handle);
  ASSERT_EQ(fm_sheet_add_row_break(wb.handle, 0, 10U, 1), 0);

  PaginationGuard p;
  ASSERT_EQ(fm_workbook_paginate(wb.handle, 0, &p.handle), 0);
  const size_t h_count = fm_pagination_horizontal_break_count(p.handle);
  ASSERT_GT(h_count, 1U);
  std::size_t manual_breaks = 0;
  for (size_t i = 0; i < h_count; ++i) {
    std::uint32_t row = 0;
    int32_t manual = -1;
    ASSERT_EQ(fm_pagination_horizontal_break_at(p.handle, i, &row), 0);
    ASSERT_EQ(fm_pagination_horizontal_break_is_manual(p.handle, i, &manual), 0);
    EXPECT_EQ(manual, row == 10U ? 1 : 0) << "break before row " << row;
    manual_breaks += static_cast<std::size_t>(manual);
  }
  EXPECT_EQ(manual_breaks, 1U);

  const size_t v_count = fm_pagination_vertical_break_count(p.handle);
  for (size_t i = 0; i < v_count; ++i) {
    int32_t manual = -1;
    ASSERT_EQ(fm_pagination_vertical_break_is_manual(p.handle, i, &manual), 0);
    EXPECT_EQ(manual, 0);
  }
  int32_t manual = 0;
  EXPECT_EQ(fm_pagination_horizontal_break_is_manual(p.handle, h_count, &manual), kInvalidArgument);
  EXPECT_EQ(fm_pagination_vertical_break_is_manual(p.handle, v_count, &manual), kInvalidArgument);
  EXPECT_NE(fm_pagination_vertical_break_is_manual(nullptr, 0, &manual), 0);
}
