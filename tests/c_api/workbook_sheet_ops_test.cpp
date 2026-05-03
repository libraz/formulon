// Copyright 2026 libraz. Licensed under the MIT License.
//
// C ABI smoke tests for the workbook structural mutation surface added
// in the sheet-rename / move / remove + defined-name editing bundle.
// Each test drives the public C entry points (`fm_workbook_*`) only;
// no C++ engine type leaks across the boundary.

#include <cstdint>
#include <cstring>
#include <string>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"

namespace {

// Sentinel index past `sheet_count()` for any test workbook in this
// file, used to drive the out-of-range rejection paths.
constexpr std::uint32_t kOutOfRangeIndex = 99U;

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

// Builds a workbook with three sheets — `Alpha`, `Beta`, `Gamma`.
void MakeThreeSheets(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create_empty(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Alpha"), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Beta"), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Gamma"), 0);
}

const char* SheetName(fm_workbook_t* wb, std::size_t idx) {
  const char* name = nullptr;
  EXPECT_EQ(fm_workbook_sheet_name(wb, idx, &name), 0);
  return name != nullptr ? name : "";
}

}  // namespace

TEST(WorkbookSheetOpsCApi, RenameUpdatesName) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  ASSERT_EQ(fm_workbook_rename_sheet(wb.handle, 1, "Charlie"), 0);
  EXPECT_STREQ(SheetName(wb.handle, 1), "Charlie");
}

TEST(WorkbookSheetOpsCApi, RenameRejectsBadName) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  // Empty name.
  EXPECT_NE(fm_workbook_rename_sheet(wb.handle, 0, ""), 0);
  // Forbidden character.
  EXPECT_NE(fm_workbook_rename_sheet(wb.handle, 0, "Bad/Name"), 0);
  // Collision (case-insensitive).
  EXPECT_NE(fm_workbook_rename_sheet(wb.handle, 0, "BETA"), 0);
  // Original name must still be intact.
  EXPECT_STREQ(SheetName(wb.handle, 0), "Alpha");
}

TEST(WorkbookSheetOpsCApi, RenameRejectsOutOfRange) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  fm_status_t rc = fm_workbook_rename_sheet(wb.handle, kOutOfRangeIndex, "Whatever");
  EXPECT_NE(rc, 0);
}

TEST(WorkbookSheetOpsCApi, RenameNullArgumentsRejected) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  EXPECT_NE(fm_workbook_rename_sheet(nullptr, 0, "X"), 0);
  EXPECT_NE(fm_workbook_rename_sheet(wb.handle, 0, nullptr), 0);
}

TEST(WorkbookSheetOpsCApi, RemoveDropsSheet) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  ASSERT_EQ(fm_workbook_remove_sheet(wb.handle, 1), 0);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 2U);
  EXPECT_STREQ(SheetName(wb.handle, 0), "Alpha");
  EXPECT_STREQ(SheetName(wb.handle, 1), "Gamma");
}

TEST(WorkbookSheetOpsCApi, RemoveRejectsLastSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);  // single Sheet1
  fm_status_t rc = fm_workbook_remove_sheet(wb.handle, 0);
  EXPECT_NE(rc, 0);
  EXPECT_EQ(fm_workbook_sheet_count(wb.handle), 1U);
}

TEST(WorkbookSheetOpsCApi, RemoveRejectsOutOfRange) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  fm_status_t rc = fm_workbook_remove_sheet(wb.handle, kOutOfRangeIndex);
  EXPECT_NE(rc, 0);
}

TEST(WorkbookSheetOpsCApi, MoveForward) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  // Alpha -> end (post-removal index 2).
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 0, 2), 0);
  EXPECT_STREQ(SheetName(wb.handle, 0), "Beta");
  EXPECT_STREQ(SheetName(wb.handle, 1), "Gamma");
  EXPECT_STREQ(SheetName(wb.handle, 2), "Alpha");
}

TEST(WorkbookSheetOpsCApi, MoveBackward) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 2, 0), 0);
  EXPECT_STREQ(SheetName(wb.handle, 0), "Gamma");
  EXPECT_STREQ(SheetName(wb.handle, 1), "Alpha");
  EXPECT_STREQ(SheetName(wb.handle, 2), "Beta");
}

TEST(WorkbookSheetOpsCApi, MoveNoOp) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  ASSERT_EQ(fm_workbook_move_sheet(wb.handle, 1, 1), 0);
  EXPECT_STREQ(SheetName(wb.handle, 1), "Beta");
}

TEST(WorkbookSheetOpsCApi, MoveRejectsOutOfRange) {
  WorkbookGuard wb;
  MakeThreeSheets(wb);
  EXPECT_NE(fm_workbook_move_sheet(wb.handle, kOutOfRangeIndex, 0), 0);
  EXPECT_NE(fm_workbook_move_sheet(wb.handle, 0, kOutOfRangeIndex), 0);
}

TEST(WorkbookSheetOpsCApi, SetDefinedNameAddsAndUpdates) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_workbook_defined_name_count(wb.handle), 0U);

  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, "MyName", "=42"), 0);
  EXPECT_EQ(fm_workbook_defined_name_count(wb.handle), 1U);

  const char* nm = nullptr;
  const char* fr = nullptr;
  ASSERT_EQ(fm_workbook_defined_name_at(wb.handle, 0, &nm, &fr), 0);
  EXPECT_STREQ(nm, "MyName");
  EXPECT_STREQ(fr, "=42");

  // Update the formula text via case-insensitive match.
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, "MYNAME", "=99"), 0);
  EXPECT_EQ(fm_workbook_defined_name_count(wb.handle), 1U);
  ASSERT_EQ(fm_workbook_defined_name_at(wb.handle, 0, &nm, &fr), 0);
  EXPECT_STREQ(nm, "MyName");  // authored case preserved
  EXPECT_STREQ(fr, "=99");
}

TEST(WorkbookSheetOpsCApi, SetDefinedNameEmptyFormulaRemoves) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, "Tmp", "=1"), 0);
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, "Tmp", ""), 0);
  EXPECT_EQ(fm_workbook_defined_name_count(wb.handle), 0U);
}

TEST(WorkbookSheetOpsCApi, SetDefinedNameRejectsEmptyName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_status_t rc = fm_workbook_set_defined_name(wb.handle, "", "=1");
  EXPECT_NE(rc, 0);
}

TEST(WorkbookSheetOpsCApi, SetDefinedNameNullArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_workbook_set_defined_name(nullptr, "n", "=1"), 0);
  EXPECT_NE(fm_workbook_set_defined_name(wb.handle, nullptr, "=1"), 0);
  EXPECT_NE(fm_workbook_set_defined_name(wb.handle, "n", nullptr), 0);
}
