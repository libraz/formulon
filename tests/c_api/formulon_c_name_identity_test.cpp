//
// C ABI: name identity folds case, width and kana type the same way on every
// profile, for defined-name lookups and for sheet names being added.

#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kAlreadyExists = static_cast<fm_status_t>(formulon::FormulonErrorCode::kAlreadyExists);

constexpr const char* kFullWidthName = "\xEF\xBD\x8E\xEF\xBD\x81\xEF\xBD\x8D\xEF\xBD\x85";  // ｎａｍｅ
constexpr const char* kFullWidthSheet1 =
    "\xEF\xBD\x93\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x91";  // ｓｈｅｅｔ１

// Builds the workbook under test: the default one when `profile_id` is null.
void Create(WorkbookGuard& wb, const char* profile_id) {
  if (profile_id == nullptr) {
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  } else {
    ASSERT_EQ(fm_workbook_create_for_profile(profile_id, &wb.handle), 0);
  }
}

constexpr const char* kProfiles[] = {nullptr, "win-365-en_US", "mac-365-en_US", "mac-365-ja_JP"};

TEST(FormulonCApiNameIdentity, DefinedNameLookupFoldsWidth) {
  for (const char* profile : kProfiles) {
    SCOPED_TRACE(profile == nullptr ? "default" : profile);
    WorkbookGuard wb;
    Create(wb, profile);
    const fm_defined_name_t entry{sizeof(fm_defined_name_t),      "Name",  "Sheet1!$A$1",
                                  FM_DEFINED_NAME_SCOPE_WORKBOOK, nullptr, 0};
    ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, &entry), 0);

    int32_t present = 0;
    fm_defined_name_t found = FM_STRUCT_INIT(fm_defined_name_t);
    ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, kFullWidthName, FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &found),
              0);
    ASSERT_EQ(present, 1);
    EXPECT_STREQ(found.name, "Name");
    EXPECT_STREQ(found.formula, "=Sheet1!$A$1");

    // A distinct name stays absent.
    ASSERT_EQ(fm_workbook_get_defined_name(wb.handle, "Names", FM_DEFINED_NAME_SCOPE_WORKBOOK, &present, &found), 0);
    EXPECT_EQ(present, 0);
  }
}

TEST(FormulonCApiNameIdentity, AddedSheetCollidingByWidthIsRefused) {
  for (const char* profile : kProfiles) {
    SCOPED_TRACE(profile == nullptr ? "default" : profile);
    WorkbookGuard wb;
    Create(wb, profile);
    size_t index = 0;
    EXPECT_EQ(fm_workbook_add_sheet(wb.handle, kFullWidthSheet1, 1, &index), kAlreadyExists);
    size_t count = 0;
    ASSERT_EQ(fm_workbook_get_sheet_count(wb.handle, &count), 0);
    EXPECT_EQ(count, 1U);

    // A width variant of a different number is a new name.
    ASSERT_EQ(fm_workbook_add_sheet(
                  wb.handle, "\xEF\xBD\x93\xEF\xBD\x88\xEF\xBD\x85\xEF\xBD\x85\xEF\xBD\x94\xEF\xBC\x92", 1, &index),
              0);
    EXPECT_EQ(index, 1U);
  }
}

}  // namespace
