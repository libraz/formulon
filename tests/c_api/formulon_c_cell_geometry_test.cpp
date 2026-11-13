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
#include "formulon_c_count_of.h"
#include "formulon_c_record_arg.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

static_assert(sizeof(fm_rect_pt_info_t) == 40U);
static_assert(sizeof(fm_sheet_format_defaults_t) == 64U);
static_assert(offsetof(fm_sheet_format_defaults_t, custom_height) == 40U);
static_assert(offsetof(fm_sheet_format_defaults_t, zero_height) == 44U);
static_assert(sizeof(fm_page_layout_info_t) == 56U);
static_assert(offsetof(fm_page_layout_info_t, origin_x_pt) == 24U);
static_assert(sizeof(fm_print_titles_t) == 28U);
static_assert(sizeof(fm_paper_info_t) == 32U);
static_assert(sizeof(fm_margins_pt_info_t) == 56U);
static_assert(offsetof(fm_width_model_info_t, normal_font_name) == 40U);

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

fm_status_t SetWholeRowBreak(fm_workbook_t* wb, uint32_t row) {
  fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
  brk.manual = FM_TRISTATE_TRUE;
  return fm_sheet_set_row_break(wb, 0, row, &brk);
}

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
  int32_t kind = FM_VALUE_KIND_BLANK;
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
  EXPECT_EQ(fm_sheet_get_cells_in_range(wb, 0, record_arg(fm_cell_rect_t{first_row, first_col, last_row, last_col}),
                                        cursor, limit, &page.handle),
            0);
  if (page.handle == nullptr) {
    return kCursorEnd;
  }
  size_t count = 0;
  EXPECT_EQ(fm_cell_range_get_cell_count(page.handle, &count), 0);
  for (size_t i = 0; i < count; ++i) {
    SeenCell cell;
    const char* formula = nullptr;
    fm_value_t value{};
    EXPECT_EQ(fm_cell_range_get_cell_at(page.handle, i, &cell.row, &cell.col, &formula, &value), 0);
    cell.formula = formula == nullptr ? "" : formula;
    cell.kind = value.kind;
    if (value.kind == FM_VALUE_KIND_NUMBER) {
      cell.number = value.u.number;
    } else if (value.kind == FM_VALUE_KIND_TEXT) {
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
  ASSERT_EQ(fm_sheet_set_number(wb, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 0, 3, "head"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb, 0, 2, 1, 2.0), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb, 0, 2, 2, "=A1+B3"), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb, 0, 4, 4, "=SEQUENCE(3)"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb, 0, 900, 7, 9.0), 0);
  ASSERT_EQ(fm_sheet_set_text(wb, 0, 50000, 100, "far"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
}

}  // namespace

/* -------------------------------------------------------------------------- */
/* Geometry                                                                   */
/* -------------------------------------------------------------------------- */

// Zero is Windows, the platform of the runtime default profile.
TEST(FormulonCApiCellGeometry, WidthPlatformOrdinals) {
  EXPECT_EQ(FM_WIDTH_PLATFORM_WIN, 0);
  EXPECT_EQ(FM_WIDTH_PLATFORM_MAC, 1);
}

TEST(FormulonCApiCellGeometry, WidthModelReportsCalibratedCalibri) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_width_model_info_t display = FM_STRUCT_INIT(fm_width_model_info_t);
  ASSERT_EQ(fm_sheet_get_width_model(wb.handle, 0, FM_GEOMETRY_MODE_DISPLAY, &display), 0);
  EXPECT_DOUBLE_EQ(display.points_per_char, 5.25);
  EXPECT_DOUBLE_EQ(display.padding_pt, 3.75);
  EXPECT_EQ(display.calibrated, 1);
  EXPECT_STREQ(display.normal_font_name, "Calibri");
  EXPECT_DOUBLE_EQ(display.normal_font_size_pt, 11.0);
  EXPECT_EQ(display.platform, FM_WIDTH_PLATFORM_WIN);

  fm_width_model_info_t print = FM_STRUCT_INIT(fm_width_model_info_t);
  ASSERT_EQ(fm_sheet_get_width_model(wb.handle, 0, FM_GEOMETRY_MODE_PRINT, &print), 0);
  EXPECT_DOUBLE_EQ(print.points_per_char, 39.0 / 7.0);
  EXPECT_DOUBLE_EQ(print.padding_pt, 27.0 / 7.0);

  EXPECT_EQ(fm_sheet_get_width_model(wb.handle, 0, 2, &print), kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_width_model(wb.handle, 0, -1, &print), kInvalidArgument);
}

TEST(FormulonCApiCellGeometry, MacProfileWidthModelUsesWholePoints) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "mac-365-en_US"), 0);
  fm_width_model_info_t display = FM_STRUCT_INIT(fm_width_model_info_t);
  ASSERT_EQ(fm_sheet_get_width_model(wb.handle, 0, FM_GEOMETRY_MODE_DISPLAY, &display), 0);
  EXPECT_DOUBLE_EQ(display.points_per_char, 6.0);
  EXPECT_DOUBLE_EQ(display.padding_pt, 5.0);
  EXPECT_EQ(display.calibrated, 1);
  EXPECT_STREQ(display.normal_font_name, "Calibri");
  EXPECT_EQ(display.platform, FM_WIDTH_PLATFORM_MAC);

  // Print is the Windows figure on every host.
  fm_width_model_info_t print = FM_STRUCT_INIT(fm_width_model_info_t);
  ASSERT_EQ(fm_sheet_get_width_model(wb.handle, 0, FM_GEOMETRY_MODE_PRINT, &print), 0);
  EXPECT_EQ(print.platform, FM_WIDTH_PLATFORM_WIN);
  EXPECT_DOUBLE_EQ(print.points_per_char, 39.0 / 7.0);

  // A sheet with no <sheetFormatPr> derives its default from base 10 although the base still reads 8.
  fm_sheet_format_defaults_t defaults = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &defaults), 0);
  EXPECT_DOUBLE_EQ(defaults.base_column_width_chars, 8.0);
  EXPECT_DOUBLE_EQ(defaults.effective_default_column_width_chars, 10.83203125);

  double pt = 0.0;
  ASSERT_EQ(fm_sheet_convert_column_chars_to_pt(wb.handle, 0, 10.83203125, FM_GEOMETRY_MODE_DISPLAY, &pt), 0);
  EXPECT_DOUBLE_EQ(pt, 65.0);
  ASSERT_EQ(fm_sheet_convert_column_chars_to_pt(wb.handle, 0, 10.0, FM_GEOMETRY_MODE_DISPLAY, &pt), 0);
  EXPECT_DOUBLE_EQ(pt, 60.0);
  double chars = 0.0;
  ASSERT_EQ(fm_sheet_convert_column_pt_to_chars(wb.handle, 0, 60.0, FM_GEOMETRY_MODE_DISPLAY, &chars), 0);
  EXPECT_DOUBLE_EQ(chars, 10.0);

  // Back to Windows: the same workbook reads the 96-DPI figures again.
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-en_US"), 0);
  ASSERT_EQ(fm_sheet_get_width_model(wb.handle, 0, FM_GEOMETRY_MODE_DISPLAY, &display), 0);
  EXPECT_DOUBLE_EQ(display.points_per_char, 5.25);
  EXPECT_EQ(display.platform, FM_WIDTH_PLATFORM_WIN);
}

// The body font of a new Mac workbook follows the theme; its face is the locale's theme face, not the stored name.
TEST(FormulonCApiCellGeometry, MacNewWorkbookReportsTheThemeFace) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_for_profile("mac-365-ja_JP", &wb.handle), 0);
  fm_width_model_info_t display = FM_STRUCT_INIT(fm_width_model_info_t);
  ASSERT_EQ(fm_sheet_get_width_model(wb.handle, 0, FM_GEOMETRY_MODE_DISPLAY, &display), 0);
  EXPECT_STREQ(display.normal_font_name, "游ゴシック");
  EXPECT_DOUBLE_EQ(display.normal_font_size_pt, 12.0);
  EXPECT_DOUBLE_EQ(display.points_per_char, 7.0);
  EXPECT_DOUBLE_EQ(display.padding_pt, 5.0);
  EXPECT_EQ(display.calibrated, 1);

  fm_sheet_format_defaults_t defaults = FM_STRUCT_INIT(fm_sheet_format_defaults_t);
  ASSERT_EQ(fm_sheet_get_format_defaults(wb.handle, 0, &defaults), 0);
  EXPECT_DOUBLE_EQ(defaults.effective_default_column_width_chars, 10.0 + 182.0 / 256.0);
  double pt = 0.0;
  ASSERT_EQ(fm_sheet_convert_column_chars_to_pt(wb.handle, 0, defaults.effective_default_column_width_chars,
                                                FM_GEOMETRY_MODE_DISPLAY, &pt),
            0);
  EXPECT_DOUBLE_EQ(pt, 75.0);
  ASSERT_EQ(fm_sheet_convert_column_chars_to_pt(wb.handle, 0, 10.0, FM_GEOMETRY_MODE_DISPLAY, &pt), 0);
  EXPECT_DOUBLE_EQ(pt, 70.0);
}

TEST(FormulonCApiCellGeometry, CharsAndPointsConvertBothWays) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  double pt = 0.0;
  // Stored width 10 is Truncate((2560 + 18) / 256 * 7) = 70 px.
  ASSERT_EQ(fm_sheet_convert_column_chars_to_pt(wb.handle, 0, 10.0, FM_GEOMETRY_MODE_DISPLAY, &pt), 0);
  EXPECT_DOUBLE_EQ(pt, 70.0 * 0.75);
  double chars = 0.0;
  ASSERT_EQ(fm_sheet_convert_column_pt_to_chars(wb.handle, 0, pt, FM_GEOMETRY_MODE_DISPLAY, &chars), 0);
  EXPECT_DOUBLE_EQ(chars, 10.0);
  ASSERT_EQ(fm_sheet_convert_column_chars_to_pt(wb.handle, 0, 0.0, FM_GEOMETRY_MODE_PRINT, &pt), 0);
  EXPECT_EQ(pt, 0.0);

  EXPECT_EQ(fm_sheet_convert_column_chars_to_pt(wb.handle, 0, -1.0, FM_GEOMETRY_MODE_DISPLAY, &pt), kInvalidArgument);
  EXPECT_EQ(fm_sheet_convert_column_pt_to_chars(wb.handle, 0, std::numeric_limits<double>::infinity(),
                                                FM_GEOMETRY_MODE_DISPLAY, &chars),
            kInvalidArgument);
  EXPECT_NE(fm_sheet_convert_column_chars_to_pt(wb.handle, 0, 1.0, FM_GEOMETRY_MODE_DISPLAY, nullptr), 0);
  // The value comes before the mode: an unknown mode is refused, not read as a width.
  EXPECT_EQ(fm_sheet_convert_column_chars_to_pt(wb.handle, 0, 10.0, 7, &pt), kInvalidArgument);
}

TEST(FormulonCApiCellGeometry, TrackSizesFollowOverridesAndHiddenState) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 1U, 20.0), 0);
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 2U, 2U, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 1U, 1U, 30.0), 0);
  ASSERT_EQ(fm_sheet_set_row_hidden(wb.handle, 0, 2U, 2U, 1), 0);

  double width = 0.0;
  ASSERT_EQ(fm_sheet_get_column_width_pt(wb.handle, 0, 1U, FM_GEOMETRY_MODE_DISPLAY, &width), 0);
  EXPECT_DOUBLE_EQ(width, 140.0 * 0.75);  // Stored width 20 is 140 px.
  ASSERT_EQ(fm_sheet_get_column_width_pt(wb.handle, 0, 2U, FM_GEOMETRY_MODE_DISPLAY, &width), 0);
  EXPECT_EQ(width, 0.0);
  double height = 0.0;
  ASSERT_EQ(fm_sheet_get_row_height_pt(wb.handle, 0, 1U, &height), 0);
  EXPECT_EQ(height, 30.0);
  ASSERT_EQ(fm_sheet_get_row_height_pt(wb.handle, 0, 2U, &height), 0);
  EXPECT_EQ(height, 0.0);

  EXPECT_EQ(fm_sheet_get_column_width_pt(wb.handle, 0, 16384U, FM_GEOMETRY_MODE_DISPLAY, &width), kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_row_height_pt(wb.handle, 0, 1048576U, &height), kInvalidArgument);
}

TEST(FormulonCApiCellGeometry, CellRectSpansRangeAndSkipsHiddenTracks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 1U, 1U, 20.0), 0);
  ASSERT_EQ(fm_sheet_set_column_hidden(wb.handle, 0, 2U, 2U, 1), 0);
  ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 1U, 1U, 30.0), 0);

  double col_a = 0.0;
  double col_b = 0.0;
  double row_1 = 0.0;
  double row_3 = 0.0;
  ASSERT_EQ(fm_sheet_get_column_width_pt(wb.handle, 0, 0U, FM_GEOMETRY_MODE_DISPLAY, &col_a), 0);
  ASSERT_EQ(fm_sheet_get_column_width_pt(wb.handle, 0, 1U, FM_GEOMETRY_MODE_DISPLAY, &col_b), 0);
  ASSERT_EQ(fm_sheet_get_row_height_pt(wb.handle, 0, 0U, &row_1), 0);
  ASSERT_EQ(fm_sheet_get_row_height_pt(wb.handle, 0, 2U, &row_3), 0);

  // B2:C3 -- column C is hidden and contributes nothing.
  fm_rect_pt_info_t rect = FM_STRUCT_INIT(fm_rect_pt_info_t);
  ASSERT_EQ(fm_sheet_get_cell_rect_pt(wb.handle, 0, record_arg(fm_cell_rect_t{1U, 1U, 2U, 2U}),
                                      FM_GEOMETRY_MODE_DISPLAY, &rect),
            0);
  EXPECT_DOUBLE_EQ(rect.x, col_a);
  EXPECT_DOUBLE_EQ(rect.y, row_1);
  EXPECT_DOUBLE_EQ(rect.width_pt, col_b);
  EXPECT_DOUBLE_EQ(rect.height_pt, 30.0 + row_3);

  // A single cell agrees with the per-track getters exactly.
  ASSERT_EQ(fm_sheet_get_cell_rect_pt(wb.handle, 0, record_arg(fm_cell_rect_t{1U, 1U, 1U, 1U}), FM_GEOMETRY_MODE_PRINT,
                                      &rect),
            0);
  double col_b_print = 0.0;
  ASSERT_EQ(fm_sheet_get_column_width_pt(wb.handle, 0, 1U, FM_GEOMETRY_MODE_PRINT, &col_b_print), 0);
  EXPECT_EQ(rect.width_pt, col_b_print);
  EXPECT_EQ(rect.height_pt, 30.0);

  EXPECT_EQ(fm_sheet_get_cell_rect_pt(wb.handle, 0, record_arg(fm_cell_rect_t{2U, 0U, 1U, 0U}),
                                      FM_GEOMETRY_MODE_DISPLAY, &rect),
            kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_cell_rect_pt(wb.handle, 0, record_arg(fm_cell_rect_t{0U, 0U, 0U, 0U}), 7, &rect),
            kInvalidArgument);
  EXPECT_NE(fm_sheet_get_cell_rect_pt(wb.handle, 0, record_arg(fm_cell_rect_t{0U, 0U, 0U, 0U}),
                                      FM_GEOMETRY_MODE_DISPLAY, nullptr),
            0);
}

TEST(FormulonCApiCellGeometry, DefaultsDriveUnoverriddenTracks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_sheet_format_defaults_patch_t defaults = FM_STRUCT_INIT(fm_sheet_format_defaults_patch_t);
  defaults.default_column_width_chars_op = FM_PATCH_OP_SET;
  defaults.default_column_width_chars = 12.0;
  defaults.default_row_height_pt_op = FM_PATCH_OP_SET;
  defaults.default_row_height_pt = 21.0;
  ASSERT_EQ(fm_sheet_update_format_defaults(wb.handle, 0, &defaults), 0);
  double width = 0.0;
  ASSERT_EQ(fm_sheet_get_column_width_pt(wb.handle, 0, 5U, FM_GEOMETRY_MODE_DISPLAY, &width), 0);
  EXPECT_DOUBLE_EQ(width, 84.0 * 0.75);  // Stored width 12 is 84 px.
  double height = 0.0;
  ASSERT_EQ(fm_sheet_get_row_height_pt(wb.handle, 0, 5U, &height), 0);
  EXPECT_EQ(height, 21.0);
}

/* -------------------------------------------------------------------------- */
/* Range enumeration and formula text                                         */
/* -------------------------------------------------------------------------- */

TEST(FormulonCApiCellRange, EnumeratesPopulatedCellsInRowMajorOrder) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
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

  EXPECT_EQ(cells[1].kind, FM_VALUE_KIND_TEXT);
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
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
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
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
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
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  RangeGuard page;
  EXPECT_EQ(fm_sheet_get_cells_in_range(wb.handle, 0, record_arg(fm_cell_rect_t{5U, 0U, 1U, 0U}), 0U, 0U, &page.handle),
            kInvalidArgument);
  EXPECT_EQ(page.handle, nullptr);
  EXPECT_EQ(fm_sheet_get_cells_in_range(wb.handle, 0, record_arg(fm_cell_rect_t{0U, 0U, 1U, 1U}),
                                        1048576U * kCursorStride, 0U, &page.handle),
            kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_cells_in_range(wb.handle, 0, record_arg(fm_cell_rect_t{0U, 0U, 1U, 1U}), kCursorEnd, 0U,
                                        &page.handle),
            kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_cells_in_range(wb.handle, 3, record_arg(fm_cell_rect_t{0U, 0U, 1U, 1U}), 0U, 0U, &page.handle),
            FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_NE(fm_sheet_get_cells_in_range(wb.handle, 0, record_arg(fm_cell_rect_t{0U, 0U, 1U, 1U}), 0U, 0U, nullptr), 0);

  ASSERT_EQ(fm_sheet_get_cells_in_range(wb.handle, 0, record_arg(fm_cell_rect_t{0U, 0U, 1U, 1U}), 0U, 0U, &page.handle),
            0);
  size_t count = 99;
  ASSERT_EQ(fm_cell_range_get_cell_count(page.handle, &count), 0);
  EXPECT_EQ(count, 0U);
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  fm_value_t value{};
  EXPECT_EQ(fm_cell_range_get_cell_at(page.handle, 0, &row, &col, nullptr, &value), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_NE(fm_cell_range_get_cell_count(nullptr, &count), 0);
  std::uint64_t next = 0;
  EXPECT_NE(fm_cell_range_next_cursor(nullptr, &next), 0);
  fm_cell_range_destroy(nullptr);
}

TEST(FormulonCApiCellRange, FormulaGettersReadA1AndR1C1) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 4, 2, "=SUM($A$1:B4)+A5"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);

  const char* text = nullptr;
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_formula(wb.handle, 0, 4, 2, &present, &text), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(text, "=SUM($A$1:B4)+A5");
  ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 4, 2, &present, &text), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(text, "=SUM(R1C1:R[-1]C[-1])+RC[-2]");

  // A cell without a formula reads absent with a NULL text.
  ASSERT_EQ(fm_sheet_get_formula(wb.handle, 0, 0, 0, &present, &text), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(text, nullptr);
  ASSERT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 0, 9, 9, &present, &text), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(text, nullptr);

  EXPECT_EQ(fm_sheet_get_formula(wb.handle, 0, 1048576U, 0, &present, &text), kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_formula_r1c1(wb.handle, 4, 0, 0, &present, &text), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_NE(fm_sheet_get_formula(wb.handle, 0, 0, 0, &present, nullptr), 0);
  EXPECT_NE(fm_sheet_get_formula(wb.handle, 0, 0, 0, nullptr, &text), 0);
}

TEST(FormulonCApiCellRange, MergesInRangeReportsIntersectingMerges) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0, record_arg(fm_cell_rect_t{0U, 0U, 1U, 1U}), nullptr), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0, record_arg(fm_cell_rect_t{5U, 5U, 6U, 8U}), nullptr), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0, record_arg(fm_cell_rect_t{20U, 0U, 20U, 3U}), nullptr), 0);

  std::size_t total = 0;
  ASSERT_EQ(fm_sheet_get_merges_in_range(wb.handle, 0, record_arg(fm_cell_rect_t{1U, 1U, 6U, 6U}), nullptr, 0U, &total),
            0);
  EXPECT_EQ(total, 2U);
  // Transposed corners are an inverted rectangle, not another spelling of it.
  EXPECT_EQ(fm_sheet_get_merges_in_range(wb.handle, 0, record_arg(fm_cell_rect_t{6U, 6U, 1U, 1U}), nullptr, 0U, &total),
            FM_STATUS_INVALID_ARGUMENT);
  fm_cell_rect_t out[1] = {};
  ASSERT_EQ(fm_sheet_get_merges_in_range(wb.handle, 0, record_arg(fm_cell_rect_t{1U, 1U, 6U, 6U}), out, 1U, &total), 0);
  EXPECT_EQ(total, 2U);
  EXPECT_EQ(out[0].first_row, 0U);
  EXPECT_EQ(out[0].last_col, 1U);

  EXPECT_NE(fm_sheet_get_merges_in_range(wb.handle, 0, record_arg(fm_cell_rect_t{0U, 0U, 1U, 1U}), nullptr, 1U, &total),
            0);
  EXPECT_EQ(fm_sheet_get_merges_in_range(wb.handle, 9, record_arg(fm_cell_rect_t{0U, 0U, 1U, 1U}), nullptr, 0U, &total),
            FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

/* -------------------------------------------------------------------------- */
/* Pagination detail                                                          */
/* -------------------------------------------------------------------------- */

namespace {

/// Fills `A1:T200` so the sheet spans several pages on both axes.
void FillGrid(fm_workbook_t* wb) {
  for (std::uint32_t row = 0; row < 200U; ++row) {
    for (std::uint32_t col = 0; col < 20U; ++col) {
      ASSERT_EQ(fm_sheet_set_number(wb, 0, row, col, 1.0), 0);
    }
  }
}

}  // namespace

TEST(FormulonCApiPaginationDetail, ReportsPaperMarginsScaleAndTitles) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup paperSize=\"9\" orientation=\"landscape\"/>"), 0);
  fm_print_titles_t repeat = FM_STRUCT_INIT(fm_print_titles_t);
  repeat.has_rows = 1;
  ASSERT_EQ(fm_sheet_set_print_titles(wb.handle, 0, &repeat), 0);

  PaginationGuard p;
  ASSERT_EQ(fm_sheet_paginate(wb.handle, 0, &p.handle), 0);
  fm_paper_info_t paper = FM_STRUCT_INIT(fm_paper_info_t);
  ASSERT_EQ(fm_pagination_get_paper(p.handle, &paper), 0);
  EXPECT_EQ(paper.landscape, 1);
  EXPECT_EQ(paper.known, 1);
  EXPECT_GT(paper.width_pt, paper.height_pt);

  fm_margins_pt_info_t margins = FM_STRUCT_INIT(fm_margins_pt_info_t);
  ASSERT_EQ(fm_pagination_get_margins(p.handle, &margins), 0);
  // Margins are reported in inches, as `<pageMargins>` states them.
  EXPECT_DOUBLE_EQ(margins.left_in, 0.7);
  EXPECT_DOUBLE_EQ(margins.top_in, 0.75);

  fm_rect_pt_info_t printable = FM_STRUCT_INIT(fm_rect_pt_info_t);
  ASSERT_EQ(fm_pagination_get_printable(p.handle, &printable), 0);
  EXPECT_DOUBLE_EQ(printable.x, margins.left_in * 72.0);
  EXPECT_GT(printable.width_pt, 0.0);

  double scale_pct = 0.0;
  ASSERT_EQ(fm_pagination_get_scale(p.handle, &scale_pct), 0);
  EXPECT_EQ(scale_pct, 100.0);

  fm_print_titles_t titles = FM_STRUCT_INIT(fm_print_titles_t);
  ASSERT_EQ(fm_pagination_get_print_titles(p.handle, &titles), 0);
  EXPECT_EQ(titles.has_rows, 1);
  EXPECT_EQ(titles.first_row, 0U);
  EXPECT_EQ(titles.last_row, 0U);
  EXPECT_EQ(titles.has_columns, 0);

  EXPECT_NE(fm_pagination_get_paper(nullptr, &paper), 0);
  EXPECT_NE(fm_pagination_get_margins(p.handle, nullptr), 0);
  EXPECT_NE(fm_pagination_get_scale(p.handle, nullptr), 0);
  EXPECT_NE(fm_pagination_get_print_titles(nullptr, &titles), 0);
  EXPECT_NE(fm_pagination_get_printable(nullptr, &printable), 0);
}

TEST(FormulonCApiPaginationDetail, UnknownPaperFallsBackAndSaysSo) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup paperSize=\"999\"/>"), 0);
  PaginationGuard p;
  ASSERT_EQ(fm_sheet_paginate(wb.handle, 0, &p.handle), 0);
  fm_paper_info_t paper = FM_STRUCT_INIT(fm_paper_info_t);
  ASSERT_EQ(fm_pagination_get_paper(p.handle, &paper), 0);
  EXPECT_EQ(paper.known, 0);
}

// Zero is the OOXML default `downThenOver`.
TEST(FormulonCApiPaginationDetail, PageOrderOrdinalsStartAtTheSchemaDefault) {
  EXPECT_EQ(FM_PAGE_ORDER_DOWN_THEN_OVER, 0);
  EXPECT_EQ(FM_PAGE_ORDER_OVER_THEN_DOWN, 1);
}

TEST(FormulonCApiPaginationDetail, PagesFollowPageOrderAndCoverEveryPage) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  ASSERT_EQ(fm_sheet_set_page_setup_xml(wb.handle, 0, "<pageSetup pageOrder=\"overThenDown\"/>"), 0);

  PaginationGuard p;
  ASSERT_EQ(fm_sheet_paginate(wb.handle, 0, &p.handle), 0);
  int32_t order = -1;
  ASSERT_EQ(fm_pagination_get_page_order(p.handle, &order), 0);
  EXPECT_EQ(order, FM_PAGE_ORDER_OVER_THEN_DOWN);

  const std::size_t page_count = count_of(fm_pagination_get_page_count, p.handle);
  ASSERT_GT(count_of(fm_pagination_get_column_break_count, p.handle), 0U);
  ASSERT_GT(page_count, 2U);
  std::vector<fm_page_layout_info_t> pages(page_count, FM_STRUCT_INIT(fm_page_layout_info_t));
  for (std::uint32_t i = 0; i < page_count; ++i) {
    ASSERT_EQ(fm_pagination_get_page_at(p.handle, i, &pages[i]), 0);
    EXPECT_GT(pages[i].width_pt, 0.0);
    EXPECT_GT(pages[i].height_pt, 0.0);
  }
  // Over then down: the second page sits right of the first, same row band.
  EXPECT_EQ(pages[1].rect.first_row, pages[0].rect.first_row);
  EXPECT_GT(pages[1].rect.first_col, pages[0].rect.last_col);
  fm_page_layout_info_t extra = FM_STRUCT_INIT(fm_page_layout_info_t);
  EXPECT_EQ(fm_pagination_get_page_at(p.handle, page_count, &extra), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_NE(fm_pagination_get_page_order(nullptr, &order), 0);
}

TEST(FormulonCApiPaginationDetail, ManualBreaksAreFlagged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  ASSERT_EQ(SetWholeRowBreak(wb.handle, 10U), 0);

  PaginationGuard p;
  ASSERT_EQ(fm_sheet_paginate(wb.handle, 0, &p.handle), 0);
  const size_t h_count = count_of(fm_pagination_get_row_break_count, p.handle);
  ASSERT_GT(h_count, 1U);
  std::size_t manual_breaks = 0;
  for (size_t i = 0; i < h_count; ++i) {
    fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
    ASSERT_EQ(fm_pagination_get_row_break_at(p.handle, i, &brk), 0);
    EXPECT_EQ(brk.manual, brk.row == 10U ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE) << "break before row " << brk.row;
    manual_breaks += brk.manual == FM_TRISTATE_TRUE ? 1U : 0U;
  }
  EXPECT_EQ(manual_breaks, 1U);

  const size_t v_count = count_of(fm_pagination_get_column_break_count, p.handle);
  for (size_t i = 0; i < v_count; ++i) {
    fm_column_break_t brk = FM_STRUCT_INIT(fm_column_break_t);
    ASSERT_EQ(fm_pagination_get_column_break_at(p.handle, i, &brk), 0);
    EXPECT_EQ(brk.manual, FM_TRISTATE_FALSE);
  }
  fm_row_break_t row_out = FM_STRUCT_INIT(fm_row_break_t);
  fm_column_break_t column_out = FM_STRUCT_INIT(fm_column_break_t);
  EXPECT_EQ(fm_pagination_get_row_break_at(p.handle, h_count, &row_out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_pagination_get_column_break_at(p.handle, v_count, &column_out), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_NE(fm_pagination_get_column_break_at(nullptr, 0, &column_out), 0);
  EXPECT_EQ(fm_pagination_get_row_break_at(p.handle, 0, nullptr), FM_STATUS_BINDING_NULL_POINTER);
}

TEST(FormulonCApiPaginationDetail, ManualBreakReadsItsDeclaredSpanAndAutomaticOnesTheWholeSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillGrid(wb.handle);
  fm_row_break_t declared = FM_STRUCT_INIT(fm_row_break_t);
  declared.first = 2U;
  declared.last = 5U;
  declared.manual = FM_TRISTATE_TRUE;
  ASSERT_EQ(fm_sheet_set_row_break(wb.handle, 0, 10U, &declared), 0);

  PaginationGuard p;
  ASSERT_EQ(fm_sheet_paginate(wb.handle, 0, &p.handle), 0);
  const size_t count = count_of(fm_pagination_get_row_break_count, p.handle);
  std::size_t manual_seen = 0;
  std::size_t automatic_seen = 0;
  for (size_t i = 0; i < count; ++i) {
    fm_row_break_t brk = FM_STRUCT_INIT(fm_row_break_t);
    ASSERT_EQ(fm_pagination_get_row_break_at(p.handle, i, &brk), 0);
    if (brk.manual == FM_TRISTATE_TRUE) {
      EXPECT_EQ(brk.row, 10U);
      EXPECT_EQ(brk.first, 2U);
      EXPECT_EQ(brk.last, 5U);
      ++manual_seen;
    } else {
      EXPECT_EQ(brk.first, 0U) << "automatic break before row " << brk.row;
      EXPECT_EQ(brk.last, 0U) << "automatic break before row " << brk.row;
      ++automatic_seen;
    }
  }
  EXPECT_EQ(manual_seen, 1U);
  EXPECT_GT(automatic_seen, 0U);
}

TEST(FormulonCApiPaginationDetail, ScaleIsAPercentageAndMarginsAreInches) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1.0), 0);
  fm_page_setup_patch_t setup = FM_STRUCT_INIT(fm_page_setup_patch_t);
  setup.scale_pct_op = FM_PATCH_OP_SET;
  setup.scale_pct = 50U;
  ASSERT_EQ(fm_sheet_update_page_setup(wb.handle, 0, &setup), 0);
  fm_page_margins_patch_t margins_patch = FM_STRUCT_INIT(fm_page_margins_patch_t);
  margins_patch.left_in_op = FM_PATCH_OP_SET;
  margins_patch.left_in = 0.5;
  margins_patch.header_in_op = FM_PATCH_OP_SET;
  margins_patch.header_in = 0.25;
  ASSERT_EQ(fm_sheet_update_page_margins(wb.handle, 0, &margins_patch), 0);

  PaginationGuard p;
  ASSERT_EQ(fm_sheet_paginate(wb.handle, 0, &p.handle), 0);
  double scale_pct = 0.0;
  ASSERT_EQ(fm_pagination_get_scale(p.handle, &scale_pct), 0);
  EXPECT_DOUBLE_EQ(scale_pct, 50.0);

  fm_margins_pt_info_t margins = FM_STRUCT_INIT(fm_margins_pt_info_t);
  ASSERT_EQ(fm_pagination_get_margins(p.handle, &margins), 0);
  EXPECT_DOUBLE_EQ(margins.left_in, 0.5);
  EXPECT_DOUBLE_EQ(margins.header_in, 0.25);
  // An untouched margin keeps the OOXML default, still in inches.
  EXPECT_DOUBLE_EQ(margins.right_in, 0.7);

  fm_rect_pt_info_t printable = FM_STRUCT_INIT(fm_rect_pt_info_t);
  ASSERT_EQ(fm_pagination_get_printable(p.handle, &printable), 0);
  EXPECT_DOUBLE_EQ(printable.x, 0.5 * 72.0);
}
