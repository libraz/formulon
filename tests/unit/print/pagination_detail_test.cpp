#include <algorithm>
#include <cstdint>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "gtest/gtest.h"
#include "io/ooxml/print_settings_parse.h"
#include "print/page_setup.h"
#include "print/pagination.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace print {
namespace {

constexpr double kPointsPerMm = 72.0 / 25.4;

DefinedName SheetName(std::string name, std::string formula) {
  DefinedName dn;
  dn.name = std::move(name);
  dn.formula = std::move(formula);
  dn.local_sheet_id = 0;
  return dn;
}

// A4 portrait with 60-character columns and 300 pt rows: one column and two
// rows fit per page, so an A1:D4 area is a 4-wide by 2-tall page grid.
Workbook MakeGridWorkbook() {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  ColumnLayout span;
  span.first = 0;
  span.last = 3;
  span.width = 60.0;
  span.has_width = true;
  sheet.mutable_layout().columns.push_back(span);
  sheet.mutable_format_defaults().has_default_row_height = true;
  sheet.mutable_format_defaults().default_row_height = 300.0;
  wb.set_defined_names({SheetName("_xlnm.Print_Area", "Sheet1!$A$1:$D$4")});
  return wb;
}

TEST(PaginationDetail, OverThenDownOrdersPages) {
  Workbook wb = MakeGridWorkbook();
  wb.sheet(0).mutable_print_settings().page_setup.page_order = PageOrder::kOverThenDown;
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const PaginationResult& r = result.value();
  EXPECT_EQ(r.page_order, PageOrder::kOverThenDown);
  ASSERT_EQ(r.pages.size(), 8U);

  // Rows 0-1 left to right, then rows 2-3 left to right.
  for (std::uint32_t i = 0; i < 8U; ++i) {
    EXPECT_EQ(r.pages[i].first_row, (i / 4U) * 2U) << i;
    EXPECT_EQ(r.pages[i].first_col, i % 4U) << i;
  }
}

TEST(PaginationDetail, DownThenOverIsTheDefaultOrder) {
  Workbook wb = MakeGridWorkbook();
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const PaginationResult& r = result.value();
  EXPECT_EQ(r.page_order, PageOrder::kDownThenOver);
  ASSERT_EQ(r.pages.size(), 8U);
  for (std::uint32_t i = 0; i < 8U; ++i) {
    EXPECT_EQ(r.pages[i].first_col, i / 2U) << i;
    EXPECT_EQ(r.pages[i].first_row, (i % 2U) * 2U) << i;
  }
}

TEST(PaginationDetail, PagesAgreeWithBreaks) {
  Workbook wb = MakeGridWorkbook();
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const PaginationResult& r = result.value();
  ASSERT_EQ(r.pages.size(), r.page_count);

  std::set<std::uint32_t> row_starts;
  std::set<std::uint32_t> col_starts;
  for (const PageLayout& page : r.pages) {
    EXPECT_LE(page.first_row, page.last_row);
    EXPECT_LE(page.first_col, page.last_col);
    row_starts.insert(page.first_row);
    col_starts.insert(page.first_col);
  }
  // Every page boundary after the first is a reported break, and vice versa.
  row_starts.erase(row_starts.begin());
  col_starts.erase(col_starts.begin());
  EXPECT_EQ(std::vector<std::uint32_t>(row_starts.begin(), row_starts.end()), r.h_breaks);
  EXPECT_EQ(std::vector<std::uint32_t>(col_starts.begin(), col_starts.end()), r.v_breaks);

  // The pages tile the print area exactly.
  std::uint32_t covered_rows = 0;
  for (const PageLayout& page : r.pages) {
    if (page.first_col == 0U) {
      covered_rows += page.last_row - page.first_row + 1U;
    }
  }
  EXPECT_EQ(covered_rows, 4U);
  EXPECT_EQ(r.pages.back().last_row, 3U);
  EXPECT_EQ(r.pages.back().last_col, 3U);
}

TEST(PaginationDetail, PageRectanglesStartAtTheMargins) {
  Workbook wb = MakeGridWorkbook();
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const PaginationResult& r = result.value();
  const PageLayout& first = r.pages.front();
  // Default margins: 0.7 in left, 0.75 in top.
  EXPECT_DOUBLE_EQ(first.origin_x_pt, 0.7 * 72.0);
  EXPECT_DOUBLE_EQ(first.origin_y_pt, 0.75 * 72.0);
  EXPECT_DOUBLE_EQ(first.height_pt, 600.0);
  EXPECT_DOUBLE_EQ(first.width_pt, 60.0 * 39.0 / 7.0 + 27.0 / 7.0);
}

TEST(PaginationDetail, ReportsPaperMarginsPrintableAndScale) {
  Workbook wb = MakeGridWorkbook();
  PageSetup& setup = wb.sheet(0).mutable_print_settings().page_setup;
  setup.orientation = Orientation::kLandscape;
  setup.scale = 50;
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const PaginationResult& r = result.value();

  EXPECT_TRUE(r.paper.landscape);
  EXPECT_TRUE(r.paper.known);
  EXPECT_NEAR(r.paper.width_pt, 297.0 * kPointsPerMm, 1e-9);
  EXPECT_NEAR(r.paper.height_pt, 210.0 * kPointsPerMm, 1e-9);
  EXPECT_DOUBLE_EQ(r.margins.left, 0.7 * 72.0);
  EXPECT_DOUBLE_EQ(r.margins.top, 0.75 * 72.0);
  EXPECT_DOUBLE_EQ(r.margins.header, 0.3 * 72.0);
  EXPECT_DOUBLE_EQ(r.printable.x, r.margins.left);
  EXPECT_DOUBLE_EQ(r.printable.y, r.margins.top);
  EXPECT_NEAR(r.printable.width, 297.0 * kPointsPerMm - 2.0 * 0.7 * 72.0, 1e-9);
  EXPECT_DOUBLE_EQ(r.scale, 0.5);
}

TEST(PaginationDetail, UnknownPaperCodeIsFlagged) {
  Workbook wb = MakeGridWorkbook();
  wb.sheet(0).mutable_print_settings().page_setup.paper_size = 9999;
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_FALSE(result.value().paper.known);
  EXPECT_NEAR(result.value().paper.width_pt, 210.0 * kPointsPerMm, 1e-9);
}

TEST(PaginationDetail, ManualBreaksAreFlaggedPerBreak) {
  Workbook wb = MakeGridWorkbook();
  // Row 3 starts a new page even though it would fit after row 2's page.
  ManualBreak brk;
  brk.id = 1;
  brk.manual = true;
  wb.sheet(0).mutable_print_settings().manual_row_breaks.push_back(brk);
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const PaginationResult& r = result.value();
  ASSERT_EQ(r.h_break_manual.size(), r.h_breaks.size());
  ASSERT_EQ(r.v_break_manual.size(), r.v_breaks.size());
  for (std::size_t i = 0; i < r.h_breaks.size(); ++i) {
    EXPECT_EQ(r.h_break_manual[i], r.h_breaks[i] == 1U) << i;
  }
  for (const bool manual : r.v_break_manual) {
    EXPECT_FALSE(manual);
  }
  EXPECT_NE(std::find(r.h_breaks.begin(), r.h_breaks.end(), 1U), r.h_breaks.end());
}

TEST(PaginationDetail, PrintTitlesAndRepeatedOffsetsAreReported) {
  Workbook wb = MakeGridWorkbook();
  wb.set_defined_names({SheetName("_xlnm.Print_Area", "Sheet1!$A$1:$D$4"),
                        SheetName("_xlnm.Print_Titles", "Sheet1!$A:$A,Sheet1!$1:$1")});
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const PaginationResult& r = result.value();
  EXPECT_TRUE(r.print_titles.has_rows);
  EXPECT_EQ(r.print_titles.first_row, 0U);
  EXPECT_EQ(r.print_titles.last_row, 0U);
  EXPECT_TRUE(r.print_titles.has_cols);
  EXPECT_EQ(r.print_titles.first_col, 0U);
  EXPECT_EQ(r.print_titles.last_col, 0U);

  // A page past the title row / column sits behind the reprinted titles;
  // the first page holds them itself.
  const PageLayout& first = r.pages.front();
  EXPECT_DOUBLE_EQ(first.origin_x_pt, r.printable.x);
  EXPECT_DOUBLE_EQ(first.origin_y_pt, r.printable.y);
  const PageLayout& last = r.pages.back();
  EXPECT_GT(last.origin_x_pt, r.printable.x);
  EXPECT_GT(last.origin_y_pt, r.printable.y);
}

TEST(PaginationDetail, PageOrderParsesFromPageSetup) {
  SheetPrintSettings settings;
  EXPECT_EQ(settings.page_setup.page_order, PageOrder::kDownThenOver);
  io::ooxml::refresh_structured_views("pageSetup", "<pageSetup orientation=\"landscape\" pageOrder=\"overThenDown\"/>",
                                      settings);
  EXPECT_EQ(settings.page_setup.page_order, PageOrder::kOverThenDown);
  // A fragment that omits the attribute reverts to the schema default.
  io::ooxml::refresh_structured_views("pageSetup", "<pageSetup paperSize=\"9\"/>", settings);
  EXPECT_EQ(settings.page_setup.page_order, PageOrder::kDownThenOver);
}

}  // namespace
}  // namespace print
}  // namespace formulon
