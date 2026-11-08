//
// Stable C ABI legacy array (CSE) formula authoring and reading tests.

#include <cstring>
#include <string>

#include "c_api/formulon_c.h"
#include "formulon_c_record_arg.h"
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
  fm_array_formula_t info = FM_STRUCT_INIT(fm_array_formula_t);
  info.engaged = 7;
  ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 0, 0, &info), 0);
  EXPECT_EQ(info.engaged, 0);
  EXPECT_EQ(info.rect.last_row, 0U);
  EXPECT_EQ(info.formula, nullptr);
}

TEST(FormulonCApiArrayFormula, SetThenGetFromEveryCellAndRecalcDistributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillColumnA(wb.handle);
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, record_arg(fm_cell_rect_t{0, 2, 2U, 2U}), "=A1:A3*2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  const char* stored = nullptr;
  ASSERT_EQ(fm_sheet_get_formula(wb.handle, 0, 0, 2, &stored), 0);
  const std::string expected = stored;
  for (uint32_t r = 0; r < 3; ++r) {
    fm_array_formula_t info = FM_STRUCT_INIT(fm_array_formula_t);
    ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, r, 2, &info), 0);
    EXPECT_EQ(info.engaged, 1);
    EXPECT_EQ(info.rect.first_row, 0U);
    EXPECT_EQ(info.rect.first_col, 2U);
    EXPECT_EQ(info.rect.last_row - info.rect.first_row + 1U, 3U);
    EXPECT_EQ(info.rect.last_col - info.rect.first_col + 1U, 1U);
    ASSERT_NE(info.formula, nullptr);
    EXPECT_EQ(expected, info.formula);
    fm_value_t v{};
    ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, r, 2, &v), 0);
    ASSERT_EQ(v.kind, FM_VALUE_KIND_NUMBER);
    EXPECT_DOUBLE_EQ(v.u.number, 2.0 * (r + 1));
  }
}

TEST(FormulonCApiArrayFormula, AnchorBlankClearsTheBlockAndNonAnchorWriteIsRefused) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillColumnA(wb.handle);
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, record_arg(fm_cell_rect_t{0, 2, 2U, 2U}), "=A1:A3*2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 2, 5.0), kPreconditionFailed);

  ASSERT_EQ(fm_sheet_set_blank(wb.handle, 0, 0, 2), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  for (uint32_t r = 0; r < 3; ++r) {
    fm_array_formula_t info = FM_STRUCT_INIT(fm_array_formula_t);
    ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, r, 2, &info), 0);
    EXPECT_EQ(info.engaged, 0) << r;
    fm_value_t v{};
    ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, r, 2, &v), 0);
    EXPECT_EQ(v.kind, FM_VALUE_KIND_BLANK) << r;
  }
}

TEST(FormulonCApiArrayFormula, ReplacesAContainedBlockAndRefusesAPartialOverlap) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  FillColumnA(wb.handle);
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, record_arg(fm_cell_rect_t{1, 3, 2U, 3U}), "=A1:A2"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 0, record_arg(fm_cell_rect_t{2, 3, 4U, 3U}), "=A1:A3"),
            kPreconditionFailed);
  ASSERT_EQ(fm_sheet_set_array_formula(wb.handle, 0, record_arg(fm_cell_rect_t{0, 3, 3U, 4U}), "=A1:A3"), 0);
  fm_array_formula_t info = FM_STRUCT_INIT(fm_array_formula_t);
  ASSERT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 2, 4, &info), 0);
  EXPECT_EQ(info.engaged, 1);
  EXPECT_EQ(info.rect.first_row, 0U);
  EXPECT_EQ(info.rect.last_row - info.rect.first_row + 1U, 4U);
  EXPECT_EQ(info.rect.last_col - info.rect.first_col + 1U, 2U);
}

TEST(FormulonCApiArrayFormula, ArgumentChecks) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_array_formula_t info = FM_STRUCT_INIT(fm_array_formula_t);
  EXPECT_EQ(fm_sheet_get_array_formula(wb.handle, 0, 0, 0, nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_get_array_formula(wb.handle, 9, 0, 0, &info), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 0U, 0U}), nullptr), kNullPointer);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 9, record_arg(fm_cell_rect_t{0, 0, 0U, 0U}), "=1"),
            FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 0, record_arg(fm_cell_rect_t{1, 0, 0, 0}), "=1"), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 0, record_arg(fm_cell_rect_t{0, 1, 0, 0}), "=1"), kInvalidArgument);
  EXPECT_EQ(fm_sheet_set_array_formula(wb.handle, 0, record_arg(fm_cell_rect_t{4294967295U, 0, 4294967295U, 0U}), "=1"),
            kInvalidArgument);
  EXPECT_NE(fm_sheet_set_array_formula(wb.handle, 0, record_arg(fm_cell_rect_t{0, 0, 0U, 0U}), "=1+"), 0);
}
