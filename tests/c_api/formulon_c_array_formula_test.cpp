//
// Stable C ABI legacy array (CSE) formula authoring and reading tests.

#include <cstring>
#include <string>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kPreconditionFailed = static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed);

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
};

void FillColumnA(fm_workbook_t* wb) {
  for (uint32_t r = 0; r < 3; ++r) {
    ASSERT_EQ(fm_sheet_set_number(wb, 0, r, 0, static_cast<double>(r + 1)), 0);
  }
}

}  // namespace

TEST(FormulonCApiArrayFormula, NotInABlockReportsNotEngaged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=42"), 0);
  fm_array_formula_t info{};
  info.engaged = 7;
  ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 0, 0, &info), 0);
  EXPECT_EQ(info.engaged, 0);
  EXPECT_EQ(info.rows, 0U);
  EXPECT_EQ(info.formula, nullptr);
}

TEST(FormulonCApiArrayFormula, SetThenGetFromEveryCellAndRecalcDistributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillColumnA(wb.handle);
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, 0, 2, 3, 1, "=A1:A3*2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  const char* stored = nullptr;
  ASSERT_EQ(fm_sheet_get_formula(wb.handle, 0, 0, 2, &stored), 0);
  const std::string expected = stored;
  for (uint32_t r = 0; r < 3; ++r) {
    fm_array_formula_t info{};
    ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, r, 2, &info), 0);
    EXPECT_EQ(info.engaged, 1);
    EXPECT_EQ(info.anchor_row, 0U);
    EXPECT_EQ(info.anchor_col, 2U);
    EXPECT_EQ(info.rows, 3U);
    EXPECT_EQ(info.cols, 1U);
    ASSERT_NE(info.formula, nullptr);
    EXPECT_EQ(expected, info.formula);
    fm_value_t v{};
    ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, r, 2, &v), 0);
    ASSERT_EQ(v.kind, FM_VAL_NUMBER);
    EXPECT_DOUBLE_EQ(v.u.number, 2.0 * (r + 1));
  }
}

TEST(FormulonCApiArrayFormula, AnchorBlankClearsTheBlockAndNonAnchorWriteIsRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillColumnA(wb.handle);
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, 0, 2, 3, 1, "=A1:A3*2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 2, 5.0), kPreconditionFailed);

  ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0, 0, 2), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  for (uint32_t r = 0; r < 3; ++r) {
    fm_array_formula_t info{};
    ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, r, 2, &info), 0);
    EXPECT_EQ(info.engaged, 0) << r;
    fm_value_t v{};
    ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, r, 2, &v), 0);
    EXPECT_EQ(v.kind, FM_VAL_BLANK) << r;
  }
}

TEST(FormulonCApiArrayFormula, ReplacesAContainedBlockAndRefusesAPartialOverlap) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillColumnA(wb.handle);
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, 1, 3, 2, 1, "=A1:A2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 0, 2, 3, 3, 1, "=A1:A3"), kPreconditionFailed);
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, 0, 3, 4, 2, "=A1:A3"), 0);
  fm_array_formula_t info{};
  ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 2, 4, &info), 0);
  EXPECT_EQ(info.engaged, 1);
  EXPECT_EQ(info.anchor_row, 0U);
  EXPECT_EQ(info.rows, 4U);
  EXPECT_EQ(info.cols, 2U);
}

TEST(FormulonCApiArrayFormula, ArgumentChecks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_array_formula_t info{};
  EXPECT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 0, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_get_array_formula(wb.handle, 9, 0, 0, &info), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 0, 0, 0, 1, 1, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 9, 0, 0, 1, 1, "=1"), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 0, 0, 0, 0, 1, "=1"), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 0, 0, 0, 1, 0, "=1"), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 0, 4294967295U, 0, 1, 1, "=1"), kInvalidArgument);
  EXPECT_NE(fm_sheet_set_array_formula(wb.handle, 0, 0, 0, 1, 1, "=1+"), 0);
}
