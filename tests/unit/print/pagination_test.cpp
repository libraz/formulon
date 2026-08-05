//
// Unit tests for the page-break engine (`src/print/pagination`).
//
// The tests assert structural correctness — page count, break ordering,
// and which track each break precedes — rather than exact 1-bit parity
// with Excel's font-metric-dependent rounding.

#include "print/pagination.h"

#include <cstdint>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "print/print_area.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace print {
namespace {

io::DefinedName PrintArea(std::string formula, std::int32_t sheet_id) {
  io::DefinedName dn;
  dn.name = "_xlnm.Print_Area";
  dn.formula = std::move(formula);
  dn.local_sheet_id = sheet_id;
  return dn;
}

// Sets a uniform column width (in OOXML character units) for [first, last].
void SetColumnWidth(Sheet* sheet, std::uint32_t first, std::uint32_t last, double width) {
  ColumnLayout span;
  span.first = first;
  span.last = last;
  span.width = width;
  sheet->mutable_layout().columns.push_back(span);
}

// Sets the height (in points) of a single row.
void SetRowHeight(Sheet* sheet, std::uint32_t row, double height) {
  RowLayout layout;
  layout.row = row;
  layout.height = height;
  sheet->mutable_layout().row_overrides.push_back(layout);
}

TEST(PaginationTest, EmptySheetWithNoPrintAreaProducesNoPages) {
  Workbook wb = Workbook::create();
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 0U);
  EXPECT_TRUE(result.value().print_area.empty());
}

TEST(PaginationTest, OutOfRangeSheetIndexIsRejected) {
  Workbook wb = Workbook::create();
  auto result = paginate(wb, 99);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(PaginationTest, SingleSmallPrintAreaIsOnePage) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$C$3", 0)});
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 1U);
  EXPECT_TRUE(result.value().h_breaks.empty());
  EXPECT_TRUE(result.value().v_breaks.empty());
}

TEST(PaginationTest, WideTableSuppressesAutoVerticalBreaks) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // 20 very wide columns (60 char units ~= 315 pt each) would overrun
  // the ~494 pt A4 body width several times. Excel's PageBreakPreview
  // does not auto-break columns at explicit print scale -- the overflow
  // clips at the right margin -- and Formulon's pagination mirrors
  // that. Only manually-inserted column breaks contribute.
  SetColumnWidth(&sheet, 0, 19, 60.0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$T$5", 0)});

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_TRUE(result.value().v_breaks.empty());
  EXPECT_EQ(result.value().page_count, 1U);
}

TEST(PaginationTest, TallTableForcesHorizontalBreak) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // Inflate the first 30 rows to 100 pt each so a 30-row print area
  // (3000 pt) far exceeds the ~734 pt A4 body height.
  for (std::uint32_t row = 0; row < 30; ++row) {
    SetRowHeight(&sheet, row, 100.0);
  }
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$B$30", 0)});

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_FALSE(result.value().h_breaks.empty());
  for (std::size_t i = 1; i < result.value().h_breaks.size(); ++i) {
    EXPECT_LT(result.value().h_breaks[i - 1], result.value().h_breaks[i]);
  }
  EXPECT_GE(result.value().page_count, 2U);
}

TEST(PaginationTest, EmptySheetWithFullColumnPrintAreaProducesNoPages) {
  // A whole-column print area on an empty sheet must short-circuit rather
  // than walk all 1,048,576 rows. The result mirrors the empty-sheet /
  // no-content spec: zero pages, returned immediately.
  Workbook wb = Workbook::create();
  wb.set_defined_names({PrintArea("Sheet1!$A:$A", 0)});
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 0U);
  EXPECT_TRUE(result.value().h_breaks.empty());
  EXPECT_TRUE(result.value().v_breaks.empty());
}

TEST(PaginationTest, EmptySheetWithFullRowPrintAreaProducesNoPages) {
  Workbook wb = Workbook::create();
  wb.set_defined_names({PrintArea("Sheet1!$1:$1", 0)});
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 0U);
}

TEST(PaginationTest, HiddenRowsAreExcludedFromPaginationExtent) {
  // 30 rows at 100 pt each force several horizontal breaks; hiding all but
  // the first five collapses the printed height to a single page.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  for (std::uint32_t row = 0; row < 30; ++row) {
    RowLayout layout;
    layout.row = row;
    layout.height = 100.0;
    layout.hidden = row >= 5;  // rows 6..30 hidden.
    sheet.mutable_layout().row_overrides.push_back(layout);
  }
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$B$30", 0)});
  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  // Five visible rows (500 pt) fit in one A4 body; hidden rows add nothing.
  EXPECT_TRUE(result.value().h_breaks.empty());
  EXPECT_EQ(result.value().page_count, 1U);
}

TEST(PaginationTest, OutlineOnlyRowUsesDefaultHeight) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(48, 0, Value::number(2.0));
  // 49 default rows are just beyond the A4 body. If this outline-only row
  // were incorrectly treated as 0pt the range would fit on one page.
  sheet.mutable_layout().row_overrides.push_back(RowLayout{24U, 0.0, false, 1U, false});
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$A$49", 0)});

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  // The outline-only row has no `ht`; it retains the default 15pt height
  // rather than silently collapsing to zero.
  EXPECT_EQ(result.value().page_count, 2U);
  ASSERT_EQ(result.value().h_breaks.size(), 1U);
  EXPECT_EQ(result.value().h_breaks[0], 44U);
}

TEST(PaginationTest, ExplicitVisibleRowWithoutHeightUsesDefaultHeight) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(48, 0, Value::number(2.0));
  // This models `<row hidden="0">`: it is an explicit row override but
  // has no `ht`, so it must not collapse during pagination.
  sheet.mutable_layout().row_overrides.push_back(RowLayout{24U, 0.0, false, 0U, false});
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$A$49", 0)});

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 2U);
  ASSERT_EQ(result.value().h_breaks.size(), 1U);
  EXPECT_EQ(result.value().h_breaks[0], 44U);
}

TEST(PaginationTest, FiftyThousandMetadataOnlyRowsRemainPractical) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  constexpr std::uint32_t kRows = 50'000U;
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(kRows - 1U, 0, Value::number(2.0));
  sheet.mutable_layout().row_overrides.reserve(kRows);
  for (std::uint32_t row = 0; row < kRows; ++row) {
    sheet.mutable_layout().row_overrides.push_back(RowLayout{row, 0.0, false, 1U, false});
  }
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$A$50000", 0)});

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_GT(result.value().page_count, 1U);
}

TEST(PaginationTest, ManualColumnBreakIsHonored) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$F$3", 0)});
  // Force a vertical break before column index 3 (column D).
  ManualBreak brk;
  brk.id = 3;
  brk.min = 0;
  brk.max = 0;
  brk.manual = true;
  sheet.mutable_print_settings().manual_col_breaks.push_back(brk);

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_FALSE(result.value().v_breaks.empty());
  EXPECT_EQ(result.value().v_breaks.front(), 3U);
  EXPECT_EQ(result.value().page_count, 2U);
}

TEST(PaginationTest, ManualRowBreakIsHonored) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$B$10", 0)});
  ManualBreak brk;
  brk.id = 4;         // Break before row index 4 (row 5).
  brk.manual = true;  // `man` defaults to false; a honored break is manual.
  sheet.mutable_print_settings().manual_row_breaks.push_back(brk);

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_FALSE(result.value().h_breaks.empty());
  EXPECT_EQ(result.value().h_breaks.front(), 4U);
  EXPECT_EQ(result.value().page_count, 2U);
}

TEST(PaginationTest, FitToWidthCollapsesToSingleColumnPage) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // Without fit-to-page this wide print area would need multiple
  // column-pages; fitToWidth=1 must shrink it to exactly one.
  SetColumnWidth(&sheet, 0, 19, 60.0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$T$5", 0)});
  sheet.mutable_print_settings().page_setup.fit_to_page = true;
  sheet.mutable_print_settings().page_setup.fit_to_width = 1;
  sheet.mutable_print_settings().page_setup.fit_to_height = 0;  // Unconstrained.

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_TRUE(result.value().v_breaks.empty());
}

TEST(PaginationTest, ScalePercentChangesPageCount) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  // A tall single-column print area that fits on a single page at 100%
  // scale (40 rows * 15 pt = 600 pt < 663 pt A4 body). Column auto-
  // breaks are intentionally suppressed (see WideTableSuppressesAuto*),
  // so this test exercises the ROW axis where scale changes the page
  // count: at 400% scale each row becomes 60 pt, overflowing the body
  // and forcing horizontal breaks.
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$A$40", 0)});

  sheet.mutable_print_settings().page_setup.scale = 100;
  auto at_full = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(at_full)) << at_full.error().message;

  sheet.mutable_print_settings().page_setup.scale = 400;
  auto enlarged = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(enlarged)) << enlarged.error().message;
  EXPECT_GT(enlarged.value().page_count, at_full.value().page_count);
}

TEST(PaginationTest, UsedRangeIsPaginatedWhenPrintAreaAbsent) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(2, 4, Value::number(2.0));

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  // `result.print_area` mirrors Excel's `PageSetup.PrintArea` exactly:
  // empty when no `_xlnm.Print_Area` defined name is set, even if the
  // sheet has populated cells. The used range is only consulted
  // internally to size the page grid.
  EXPECT_TRUE(result.value().print_area.empty());
  EXPECT_EQ(result.value().page_count, 1U);
}

TEST(PaginationTest, SpillPhantomsExtendUsedRangeForPagination) {
  // With no print area, pagination falls back to the used range, which must
  // include a dynamic-array spill's phantoms, not just its anchor. A tall
  // single-column spill A1:A60 (60 rows * 15 pt = 900 pt) exceeds the A4 body
  // height and forces at least one horizontal break; counting only the anchor
  // A1 would leave a single page.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  std::vector<Value> cells;
  cells.reserve(60);
  for (std::uint32_t i = 0; i < 60; ++i) {
    cells.push_back(Value::number(static_cast<double>(i + 1)));
  }
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 60U, 1U, std::move(cells)));

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_GE(result.value().page_count, 2U);
  EXPECT_FALSE(result.value().h_breaks.empty());
}

TEST(PaginationTest, AutomaticColumnBreakDoesNotForceAnExtraPage) {
  // Excel persists automatic breaks (man="0") once a sheet is previewed.
  // An auto column break must not be treated as a forced page boundary.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$F$3", 0)});
  ManualBreak brk;
  brk.id = 3;
  brk.manual = false;  // Automatic break.
  sheet.mutable_print_settings().manual_col_breaks.push_back(brk);

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_TRUE(result.value().v_breaks.empty());
  EXPECT_EQ(result.value().page_count, 1U);
}

TEST(PaginationTest, AutomaticRowBreakDoesNotForceAnExtraPage) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$B$10", 0)});
  ManualBreak brk;
  brk.id = 4;
  brk.manual = false;  // Automatic break.
  sheet.mutable_print_settings().manual_row_breaks.push_back(brk);

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_TRUE(result.value().h_breaks.empty());
  EXPECT_EQ(result.value().page_count, 1U);
}

TEST(PaginationTest, MultiAreaColumnBreakCountedPerArea) {
  // Two print rectangles, each spanning columns A..F, share a manual column
  // break at column index 3. Each print area is an independent page grid
  // (area boundaries break the page), so the break splits BOTH areas' column
  // span — symmetric with how a shared manual ROW break splits both of two
  // side-by-side areas. Each area is a single row-band, so
  // page_count = 2 + 2 = 4. The break still appears once in the
  // de-duplicated v_breaks position collection.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  wb.set_defined_names({PrintArea("Sheet1!$A$1:$F$3,Sheet1!$A$5:$F$7", 0)});
  // Populate both rectangles so the used-range intersection keeps them.
  sheet.set_cell_value(0, 0, Value::number(1.0));
  sheet.set_cell_value(0, 5, Value::number(1.0));
  sheet.set_cell_value(2, 5, Value::number(1.0));
  sheet.set_cell_value(4, 0, Value::number(1.0));
  sheet.set_cell_value(6, 5, Value::number(1.0));
  ManualBreak brk;
  brk.id = 3;  // Column D, inside both rectangles' column span.
  brk.manual = true;
  sheet.mutable_print_settings().manual_col_breaks.push_back(brk);

  auto result = paginate(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().page_count, 4U);
  // The break appears once in the de-duplicated v_breaks collection.
  ASSERT_EQ(result.value().v_breaks.size(), 1U);
  EXPECT_EQ(result.value().v_breaks.front(), 3U);
}

TEST(PaginationTest, MultiAreaRowAndColumnBreaksAreSymmetric) {
  // Symmetry check: two side-by-side areas (columns A..C and E..G) sharing a
  // manual ROW break, versus two stacked areas (rows 1..3 and 5..7) sharing
  // a manual COLUMN break, must produce the same per-area page count.
  //
  // Side-by-side areas + shared row break: each area splits into 2 row-bands
  // over 1 column-page -> 2 + 2 = 4.
  Workbook rows_wb = Workbook::create();
  Sheet& rows_sheet = rows_wb.sheet(0);
  rows_wb.set_defined_names({PrintArea("Sheet1!$A$1:$C$6,Sheet1!$E$1:$G$6", 0)});
  rows_sheet.set_cell_value(0, 0, Value::number(1.0));
  rows_sheet.set_cell_value(5, 0, Value::number(1.0));
  rows_sheet.set_cell_value(0, 4, Value::number(1.0));
  rows_sheet.set_cell_value(5, 6, Value::number(1.0));
  ManualBreak row_brk;
  row_brk.id = 3;  // Row 4, inside both areas' row span.
  row_brk.manual = true;
  rows_sheet.mutable_print_settings().manual_row_breaks.push_back(row_brk);
  auto rows_result = paginate(rows_wb, 0);
  ASSERT_TRUE(static_cast<bool>(rows_result)) << rows_result.error().message;
  EXPECT_EQ(rows_result.value().page_count, 4U);

  // Stacked areas + shared column break: mirror shape -> also 4.
  Workbook cols_wb = Workbook::create();
  Sheet& cols_sheet = cols_wb.sheet(0);
  cols_wb.set_defined_names({PrintArea("Sheet1!$A$1:$F$3,Sheet1!$A$5:$F$7", 0)});
  cols_sheet.set_cell_value(0, 0, Value::number(1.0));
  cols_sheet.set_cell_value(0, 5, Value::number(1.0));
  cols_sheet.set_cell_value(4, 0, Value::number(1.0));
  cols_sheet.set_cell_value(6, 5, Value::number(1.0));
  ManualBreak col_brk;
  col_brk.id = 3;  // Column D, inside both areas' column span.
  col_brk.manual = true;
  cols_sheet.mutable_print_settings().manual_col_breaks.push_back(col_brk);
  auto cols_result = paginate(cols_wb, 0);
  ASSERT_TRUE(static_cast<bool>(cols_result)) << cols_result.error().message;
  EXPECT_EQ(cols_result.value().page_count, rows_result.value().page_count);
}

}  // namespace
}  // namespace print
}  // namespace formulon
