// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
  brk.id = 4;  // Break before row index 4 (row 5).
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

}  // namespace
}  // namespace print
}  // namespace formulon
