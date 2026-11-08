//
// Stable C ABI tests for `fm_sheet_get_lambda_text`. The rendering
// logic itself is pinned in `tests/unit/eval/lambda_format_test.cpp`;
// these tests exercise the boundary's NULL / range guards and the
// "cell is not a lambda" absent result.

#include <cstdint>
#include <cstring>
#include <string>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

}  // namespace

TEST(FormulonCApiLambdaText, FreshCellIsNotALambda) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  // No formula written; the cell does not exist, so no lambda is present.
  const char* text = nullptr;
  int32_t present = 1;
  ASSERT_EQ(fm_sheet_get_lambda_text(wb.handle, 0, 0, 0, &present, &text), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(text, nullptr);
}

TEST(FormulonCApiLambdaText, NumericCellIsNotALambda) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 42.0), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  const char* text = nullptr;
  int32_t present = 1;
  ASSERT_EQ(fm_sheet_get_lambda_text(wb.handle, 0, 0, 0, &present, &text), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(text, nullptr);
}

TEST(FormulonCApiLambdaText, BareLambdaProjectsToCalcError) {
  // Mac Excel 365 displays a bare top-level LAMBDA as #CALC! rather
  // than retaining the closure as the cell value. Formulon mirrors
  // that surface contract; this test pins it through the C ABI so
  // future divergences are caught at the boundary.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0, 0, "=LAMBDA(x, x*2)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  fm_value_t v{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0, 0, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VALUE_KIND_ERROR);
  // A #CALC! cell has no lambda surface to render.
  const char* text = nullptr;
  int32_t present = 1;
  ASSERT_EQ(fm_sheet_get_lambda_text(wb.handle, 0, 0, 0, &present, &text), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(text, nullptr);
}

TEST(FormulonCApiLambdaText, OutOfRangeSheetReturnsSheetIndexOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const char* text = nullptr;
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_lambda_text(wb.handle, 99, 0, 0, &present, &text),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
}

TEST(FormulonCApiLambdaText, NullArgsReturnBindingNullPointer) {
  const char* text = nullptr;
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_lambda_text(nullptr, 0, 0, 0, &present, &text),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  EXPECT_EQ(fm_sheet_get_lambda_text(wb.handle, 0, 0, 0, &present, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_sheet_get_lambda_text(wb.handle, 0, 0, 0, nullptr, &text),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}
