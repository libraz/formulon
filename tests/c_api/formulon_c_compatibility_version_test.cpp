//
// fm_workbook_get/set_compatibility_version: the workbook's Compatibility
// Version, 2 for a new workbook and 1 for a loaded one without a stamp; a later
// Excel's stamp reads as stored and evaluates as 2; 1, 2 and 3 are settable (3 evaluates as 2)
// and anything else is INVALID_ARGUMENT, leaving the value. A change between
// 1 and 2 or 3 re-evaluates LEN, MID, FIND, SEARCH and REPLACE in the new unit;
// a change between 2 and 3 leaves every formula computed.

#include <cstdint>
#include <fstream>
#include <iterator>
#include <limits>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

int32_t Version(const fm_workbook_t* wb) {
  int32_t version = -1;
  EXPECT_EQ(fm_workbook_get_compatibility_version(wb, &version), FM_STATUS_OK);
  return version;
}

bool Stale(const fm_workbook_t* wb) {
  int32_t stale = -1;
  EXPECT_EQ(fm_sheet_is_cell_stale(wb, 0, 0U, 0U, &stale), FM_STATUS_OK);
  return stale == 1;
}

double Number(fm_workbook_t* wb, uint32_t row, uint32_t col) {
  fm_value_t value{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, row, col, &value), FM_STATUS_OK);
  return value.kind == FM_VALUE_KIND_NUMBER ? value.u.number : -1.0;
}

// Loads tests/fixtures/excel/compat_version/<name>: A1 "a😀b", B1 =LEN(A1), C1 =FIND("b",A1), no cached values.
void LoadFixture(const std::string& name, WorkbookGuard* out) {
  std::string dir = __FILE__;
  dir = dir.substr(0, dir.find_last_of("/\\"));
  std::ifstream in(dir + "/../fixtures/excel/compat_version/" + name, std::ios::binary);
  ASSERT_TRUE(in.good()) << name;
  const std::vector<uint8_t> bytes((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out->handle), FM_STATUS_OK);
  ASSERT_EQ(fm_workbook_recalc(out->handle), FM_STATUS_OK);
}

}  // namespace

TEST(FormulonCApiCompatibilityVersion, ANewWorkbookIsVersion2) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Version(wb.handle), FM_COMPATIBILITY_VERSION_2);
}

// A change re-evaluates: LEN("a😀") is 2 in version 2 and 3 in version 1.
TEST(FormulonCApiCompatibilityVersion, SettingAVersionReEvaluatesTheTextFunctions) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0U, 0U, "=LEN(\"a\xF0\x9F\x98\x80\")"), FM_STATUS_OK);
  EXPECT_EQ(Number(wb.handle, 0U, 0U), 2.0);
  ASSERT_EQ(fm_workbook_set_compatibility_version(wb.handle, FM_COMPATIBILITY_VERSION_1), FM_STATUS_OK);
  EXPECT_EQ(Version(wb.handle), FM_COMPATIBILITY_VERSION_1);
  EXPECT_EQ(Number(wb.handle, 0U, 0U), 3.0);
  ASSERT_EQ(fm_workbook_set_compatibility_version(wb.handle, FM_COMPATIBILITY_VERSION_1), FM_STATUS_OK);  // no-op
  ASSERT_EQ(fm_workbook_set_compatibility_version(wb.handle, FM_COMPATIBILITY_VERSION_2), FM_STATUS_OK);
  EXPECT_EQ(Number(wb.handle, 0U, 0U), 2.0);
}

TEST(FormulonCApiCompatibilityVersion, RefusalsLeaveTheValue) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  EXPECT_EQ(fm_workbook_set_compatibility_version(wb.handle, 0), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_workbook_set_compatibility_version(wb.handle, 4), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(fm_workbook_set_compatibility_version(wb.handle, -1), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(Version(wb.handle), FM_COMPATIBILITY_VERSION_2);
  int32_t version = 0;
  EXPECT_EQ(fm_workbook_get_compatibility_version(nullptr, &version), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_get_compatibility_version(wb.handle, nullptr), FM_STATUS_BINDING_NULL_POINTER);
  EXPECT_EQ(fm_workbook_set_compatibility_version(nullptr, 2), FM_STATUS_BINDING_NULL_POINTER);
}

// 3 is stored as given and evaluates as 2, so 2 <-> 3 leaves formulas computed
// and 3 -> 1 -> 3 marks them both times; an undo back to 3 restores the value.
TEST(FormulonCApiCompatibilityVersion, Version3IsSettableAndDirtiesOnlyWhenTheUnitChanges) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  ASSERT_EQ(fm_workbook_set_calc_mode(wb.handle, FM_CALC_MODE_MANUAL), FM_STATUS_OK);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0U, 0U, "=LEN(\"a\xF0\x9F\x98\x80\")"), FM_STATUS_OK);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  ASSERT_FALSE(Stale(wb.handle));

  ASSERT_EQ(fm_workbook_set_compatibility_version(wb.handle, FM_COMPATIBILITY_VERSION_3), FM_STATUS_OK);
  EXPECT_EQ(Version(wb.handle), FM_COMPATIBILITY_VERSION_3);
  EXPECT_FALSE(Stale(wb.handle));
  ASSERT_EQ(fm_workbook_set_compatibility_version(wb.handle, FM_COMPATIBILITY_VERSION_3), FM_STATUS_OK);  // no-op
  EXPECT_FALSE(Stale(wb.handle));

  ASSERT_EQ(fm_workbook_set_compatibility_version(wb.handle, FM_COMPATIBILITY_VERSION_1), FM_STATUS_OK);
  EXPECT_TRUE(Stale(wb.handle));
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Number(wb.handle, 0U, 0U), 3.0);

  ASSERT_EQ(fm_workbook_set_compatibility_version(wb.handle, FM_COMPATIBILITY_VERSION_3), FM_STATUS_OK);
  EXPECT_EQ(Version(wb.handle), FM_COMPATIBILITY_VERSION_3);
  EXPECT_TRUE(Stale(wb.handle));
  ASSERT_EQ(fm_workbook_recalc(wb.handle), FM_STATUS_OK);
  EXPECT_EQ(Number(wb.handle, 0U, 0U), 2.0);

  ASSERT_EQ(fm_workbook_set_compatibility_version(wb.handle, FM_COMPATIBILITY_VERSION_2), FM_STATUS_OK);
  EXPECT_FALSE(Stale(wb.handle));
}

// Excel computes a book without a stamp as version 1: LEN("a😀b") is 4 and FIND("b") is 4.
TEST(FormulonCApiCompatibilityVersion, ALoadedBookWithoutAStampIsVersion1) {
  WorkbookGuard wb;
  LoadFixture("stamp_none.xlsx", &wb);
  EXPECT_EQ(Version(wb.handle), FM_COMPATIBILITY_VERSION_1);
  EXPECT_EQ(Number(wb.handle, 0U, 1U), 4.0);
  EXPECT_EQ(Number(wb.handle, 0U, 2U), 4.0);
}

// A later Excel's stamp reads as stored (saturating at INT32_MAX) and evaluates as version 2.
TEST(FormulonCApiCompatibilityVersion, AStampAboveTheKnownVersionsReadsAsStored) {
  WorkbookGuard four;
  LoadFixture("stamp_4.xlsx", &four);
  EXPECT_EQ(Version(four.handle), 4);
  EXPECT_EQ(Number(four.handle, 0U, 1U), 3.0);
  EXPECT_EQ(Number(four.handle, 0U, 2U), 3.0);
  EXPECT_EQ(fm_workbook_set_compatibility_version(four.handle, 4), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(Version(four.handle), 4);

  WorkbookGuard top;
  LoadFixture("stamp_4294967295.xlsx", &top);
  EXPECT_EQ(Version(top.handle), std::numeric_limits<int32_t>::max());
  EXPECT_EQ(Number(top.handle, 0U, 1U), 3.0);
}
