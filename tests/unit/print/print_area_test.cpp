// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for Print_Area / Print_Titles resolution
// (`src/print/print_area`).

#include "print/print_area.h"

#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace print {
namespace {

// Builds a workbook with a single sheet and the given sheet-scoped
// defined names installed.
Workbook MakeWorkbook(std::vector<io::DefinedName> names) {
  Workbook wb = Workbook::create();
  wb.set_defined_names(std::move(names));
  return wb;
}

io::DefinedName SheetScoped(std::string name, std::string formula, std::int32_t sheet_id) {
  io::DefinedName dn;
  dn.name = std::move(name);
  dn.formula = std::move(formula);
  dn.local_sheet_id = sheet_id;
  return dn;
}

TEST(PrintAreaTest, AbsentDefinedNameYieldsEmptyVector) {
  Workbook wb = Workbook::create();
  auto result = resolve_print_area(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_TRUE(result.value().empty());
}

TEST(PrintAreaTest, SingleRangeResolution) {
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Area", "Sheet1!$A$1:$H$80", 0)});
  auto result = resolve_print_area(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_EQ(result.value().size(), 1U);
  const CellRange& r = result.value().front();
  EXPECT_EQ(r.first_row, 0U);
  EXPECT_EQ(r.first_col, 0U);
  EXPECT_EQ(r.last_row, 79U);
  EXPECT_EQ(r.last_col, 7U);
}

TEST(PrintAreaTest, MultiAreaCommaSeparated) {
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Area", "Sheet1!$A$1:$D$20,Sheet1!$F$1:$H$20", 0)});
  auto result = resolve_print_area(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_EQ(result.value().size(), 2U);
  EXPECT_EQ(result.value()[0].last_col, 3U);
  EXPECT_EQ(result.value()[1].first_col, 5U);
  EXPECT_EQ(result.value()[1].last_col, 7U);
}

TEST(PrintAreaTest, QuotedSheetNameWithEmbeddedCommaSplitsCorrectly) {
  // The sheet name itself contains a comma (`'Sheet,1'`); a naive
  // top-level `,` split would cut this into 4 pieces instead of the 2
  // actual print areas.
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Area", "'Sheet,1'!$A$1:$B$2,'Sheet,1'!$C$3:$D$4", 0)});
  auto result = resolve_print_area(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_EQ(result.value().size(), 2U);
  EXPECT_EQ(result.value()[0].first_col, 0U);
  EXPECT_EQ(result.value()[0].last_col, 1U);
  EXPECT_EQ(result.value()[1].first_col, 2U);
  EXPECT_EQ(result.value()[1].last_col, 3U);
}

TEST(PrintAreaTest, AnchorFreeAndUnqualifiedRangeAccepted) {
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Area", "A1:C3", 0)});
  auto result = resolve_print_area(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_EQ(result.value().size(), 1U);
  EXPECT_EQ(result.value().front().last_row, 2U);
  EXPECT_EQ(result.value().front().last_col, 2U);
}

TEST(PrintAreaTest, WorkbookScopedNameIsIgnored) {
  // local_sheet_id == -1 means workbook scope; it must not match a
  // sheet-scoped print-area lookup.
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Area", "Sheet1!$A$1:$B$2", -1)});
  auto result = resolve_print_area(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_TRUE(result.value().empty());
}

TEST(PrintAreaTest, WholeColumnSpanResolvesToFullRowExtent) {
  // A whole-column print area ($A:$C) is valid: it covers every column in
  // [A, C] across the entire row extent. It must not be rejected as
  // malformed.
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Area", "Sheet1!$A:$C", 0)});
  auto result = resolve_print_area(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_EQ(result.value().size(), 1U);
  const CellRange& r = result.value().front();
  EXPECT_EQ(r.first_col, 0U);
  EXPECT_EQ(r.last_col, 2U);
  EXPECT_EQ(r.first_row, 0U);
  EXPECT_EQ(r.last_row, Sheet::kMaxRows - 1U);
}

TEST(PrintAreaTest, WholeRowSpanResolvesToFullColumnExtent) {
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Area", "Sheet1!$1:$50", 0)});
  auto result = resolve_print_area(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_EQ(result.value().size(), 1U);
  const CellRange& r = result.value().front();
  EXPECT_EQ(r.first_row, 0U);
  EXPECT_EQ(r.last_row, 49U);
  EXPECT_EQ(r.first_col, 0U);
  EXPECT_EQ(r.last_col, Sheet::kMaxCols - 1U);
}

TEST(PrintAreaTest, OverLargeRowIsClampedToGridCeiling) {
  // A row endpoint past Excel's grid ceiling must be clamped, not honored
  // verbatim (an unclamped value would let pagination reserve a runaway
  // track vector).
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Area", "Sheet1!$1:$99999999", 0)});
  auto result = resolve_print_area(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_EQ(result.value().size(), 1U);
  EXPECT_EQ(result.value().front().last_row, Sheet::kMaxRows - 1U);
}

TEST(PrintAreaTest, MalformedFormulaReturnsInvalidArea) {
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Area", "Sheet1!$A$1:not-a-cell", 0)});
  auto result = resolve_print_area(wb, 0);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kPrintInvalidArea);
}

TEST(PrintAreaTest, EmptyFormulaReturnsInvalidArea) {
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Area", "", 0)});
  auto result = resolve_print_area(wb, 0);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kPrintInvalidArea);
}

TEST(PrintTitlesTest, AbsentDefinedNameYieldsEmptyTitles) {
  Workbook wb = Workbook::create();
  auto result = resolve_print_titles(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_FALSE(result.value().repeat_rows.has_value());
  EXPECT_FALSE(result.value().repeat_cols.has_value());
}

TEST(PrintTitlesTest, RepeatRowsResolution) {
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Titles", "Sheet1!$1:$2", 0)});
  auto result = resolve_print_titles(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_TRUE(result.value().repeat_rows.has_value());
  EXPECT_EQ(result.value().repeat_rows->first, 0U);
  EXPECT_EQ(result.value().repeat_rows->second, 1U);
  EXPECT_FALSE(result.value().repeat_cols.has_value());
}

TEST(PrintTitlesTest, RepeatColsResolution) {
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Titles", "Sheet1!$A:$A", 0)});
  auto result = resolve_print_titles(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_TRUE(result.value().repeat_cols.has_value());
  EXPECT_EQ(result.value().repeat_cols->first, 0U);
  EXPECT_EQ(result.value().repeat_cols->second, 0U);
  EXPECT_FALSE(result.value().repeat_rows.has_value());
}

TEST(PrintTitlesTest, BothRowsAndColsResolution) {
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Titles", "Sheet1!$A:$B,Sheet1!$1:$1", 0)});
  auto result = resolve_print_titles(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_TRUE(result.value().repeat_cols.has_value());
  ASSERT_TRUE(result.value().repeat_rows.has_value());
  EXPECT_EQ(result.value().repeat_cols->second, 1U);
  EXPECT_EQ(result.value().repeat_rows->first, 0U);
}

TEST(PrintTitlesTest, QuotedSheetNameWithEmbeddedCommaSplitsCorrectly) {
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Titles", "'Sheet,1'!$A:$B,'Sheet,1'!$1:$1", 0)});
  auto result = resolve_print_titles(wb, 0);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  ASSERT_TRUE(result.value().repeat_cols.has_value());
  ASSERT_TRUE(result.value().repeat_rows.has_value());
  EXPECT_EQ(result.value().repeat_cols->second, 1U);
  EXPECT_EQ(result.value().repeat_rows->first, 0U);
}

TEST(PrintTitlesTest, MalformedTokenReturnsInvalidArea) {
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Titles", "Sheet1!$1$1", 0)});
  auto result = resolve_print_titles(wb, 0);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kPrintInvalidArea);
}

TEST(PrintTitlesTest, OutOfGridRepeatRowSpanReturnsInvalidArea) {
  // `1:4294967295` would otherwise drive pagination through a
  // multi-billion-row loop; the span must be rejected up front.
  Workbook wb = MakeWorkbook({SheetScoped("_xlnm.Print_Titles", "Sheet1!$1:$4294967295", 0)});
  auto result = resolve_print_titles(wb, 0);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kPrintInvalidArea);
}

}  // namespace
}  // namespace print
}  // namespace formulon
