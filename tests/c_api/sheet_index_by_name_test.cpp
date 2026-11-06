// `fm_workbook_get_sheet_index`: a case-insensitive lookup of a sheet by name
// whose miss is a successful call with `*out_present == 0`.

#include <cstddef>
#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

namespace {

struct Found {
  fm_status_t rc;
  size_t index;
  int32_t present;
};

Found Lookup(const fm_workbook_t* wb, const char* name) {
  Found found{0, 77U, -1};
  found.rc = fm_workbook_get_sheet_index(wb, name, &found.index, &found.present);
  return found;
}

}  // namespace

TEST(FormulonCApiSheetIndexByName, FindsASheetByItsExactName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  size_t added = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Budget", &added), 0);
  ASSERT_EQ(added, 1U);

  const Found first = Lookup(wb.handle, "Sheet1");
  EXPECT_EQ(first.rc, 0);
  EXPECT_EQ(first.present, 1);
  EXPECT_EQ(first.index, 0U);
  const Found second = Lookup(wb.handle, "Budget");
  EXPECT_EQ(second.rc, 0);
  EXPECT_EQ(second.present, 1);
  EXPECT_EQ(second.index, 1U);
}

TEST(FormulonCApiSheetIndexByName, IgnoresCase) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  size_t added = 0;
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Budget", &added), 0);

  for (const char* spelling : {"budget", "BUDGET", "bUdGeT"}) {
    const Found found = Lookup(wb.handle, spelling);
    EXPECT_EQ(found.rc, 0) << spelling;
    EXPECT_EQ(found.present, 1) << spelling;
    EXPECT_EQ(found.index, 1U) << spelling;
  }
  EXPECT_EQ(Lookup(wb.handle, "sHEET1").index, 0U);
}

TEST(FormulonCApiSheetIndexByName, AMissIsASuccessfulCallWithoutAPresentFlag) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (const char* name : {"", "Missing", "Sheet1 ", "Sheet"}) {
    const Found found = Lookup(wb.handle, name);
    EXPECT_EQ(found.rc, 0) << "[" << name << "]";
    EXPECT_EQ(found.present, 0) << "[" << name << "]";
    EXPECT_EQ(found.index, 0U) << "[" << name << "]";
    EXPECT_STREQ(fm_get_last_error_message(), "");
  }
}

TEST(FormulonCApiSheetIndexByName, FollowsARename) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_rename_sheet(wb.handle, 0, "Renamed"), 0);
  EXPECT_EQ(Lookup(wb.handle, "Sheet1").present, 0);
  const Found found = Lookup(wb.handle, "renamed");
  EXPECT_EQ(found.present, 1);
  EXPECT_EQ(found.index, 0U);
}

TEST(FormulonCApiSheetIndexByName, RejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  size_t index = 0;
  int32_t present = 0;
  EXPECT_EQ(fm_workbook_get_sheet_index(nullptr, "Sheet1", &index, &present), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_sheet_index(wb.handle, "Sheet1", nullptr, &present), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_sheet_index(wb.handle, "Sheet1", &index, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_sheet_index(wb.handle, nullptr, &index, &present), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_workbook_get_sheet_index(wb.handle, "\xFF\xFE", &index, &present), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
}
